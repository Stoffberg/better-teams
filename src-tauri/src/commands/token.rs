use aes::cipher::{BlockDecryptMut, KeyIvInit};
use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine};
use pbkdf2::pbkdf2_hmac;
use regex::Regex;
use rusqlite::{Connection, OpenFlags};
#[cfg(target_os = "macos")]
use security_framework::passwords::get_generic_password;
use serde::{Deserialize, Serialize};
use sha1::Sha1;
use std::fs;
#[cfg(unix)]
use std::os::unix::fs::PermissionsExt;
use std::path::PathBuf;
use std::time::{Instant, SystemTime, UNIX_EPOCH};

type Aes128CbcDec = cbc::Decryptor<aes::Aes128>;

static PRESENCE_CACHE: std::sync::Mutex<Option<(Instant, Vec<CachedPresenceEntry>)>> =
    std::sync::Mutex::new(None);

const PBKDF2_SALT: &[u8] = b"saltysalt";
const PBKDF2_ITERATIONS: u32 = 1003;
const PBKDF2_KEY_LENGTH: usize = 16;
const ENCRYPTED_PREFIX: &[u8] = b"v10";
const CDL_WORKER_AGGREGATED_USER_DATA_PREFIX: &str =
    "https://cdl-worker/cdl-worker-cache-manager/cdl-worker-aggregated-user-data/";

// ---------- public structs ----------

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ExtractedToken {
    pub host: String,
    pub name: String,
    pub token: String,
    pub audience: Option<String>,
    pub upn: Option<String>,
    pub tenant_id: Option<String>,
    pub skype_id: Option<String>,
    pub expires_at: String,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct AccountOption {
    pub upn: Option<String>,
    pub tenant_id: Option<String>,
}

#[derive(Debug, Serialize, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CachedPresenceEntry {
    pub mri: String,
    pub presence: CachedPresenceInfo,
}

#[derive(Debug, Serialize, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CachedPresenceInfo {
    pub availability: Option<String>,
    pub activity: Option<String>,
}

// ---------- internal types ----------

struct CookieRow {
    host_key: String,
    name: String,
    encrypted_value: Vec<u8>,
}

struct TokenExtractionOutcome {
    tokens: Vec<ExtractedToken>,
    decrypt_failures: usize,
}

#[derive(Debug, Deserialize)]
struct AggregatedUserDataEnvelope {
    presence: Option<PresenceEnvelope>,
}

#[derive(Debug, Deserialize)]
struct PresenceEnvelope {
    mri: Option<String>,
    presence: Option<CachedPresenceInfo>,
}

// ---------- helpers ----------

#[cfg(target_os = "macos")]
fn cookies_db_path() -> PathBuf {
    let home = std::env::var("HOME")
        .map(PathBuf::from)
        .unwrap_or_else(|_| PathBuf::from("~"));
    home.join("Library/Containers/com.microsoft.teams2/Data/Library/Application Support/Microsoft/MSTeams/EBWebView/WV2Profile_tfw/Cookies")
}

#[cfg(target_os = "macos")]
fn read_cookie_rows(db_path: &PathBuf) -> Result<Vec<CookieRow>, String> {
    let conn = Connection::open_with_flags(db_path, OpenFlags::SQLITE_OPEN_READ_ONLY)
        .map_err(|e| format!("Failed to open cookies DB: {e}"))?;

    let mut stmt = conn
        .prepare(
            "SELECT host_key, name, encrypted_value
             FROM cookies
             WHERE (host_key LIKE '%teams%' OR host_key LIKE '%skype%')
               AND (name = 'authtoken' OR name = 'skypetoken_asm')
             ORDER BY expires_utc DESC",
        )
        .map_err(|e| format!("Failed to prepare query: {e}"))?;

    let rows = stmt
        .query_map([], |row| {
            Ok(CookieRow {
                host_key: row.get(0)?,
                name: row.get(1)?,
                encrypted_value: row.get(2)?,
            })
        })
        .map_err(|e| format!("Failed to execute query: {e}"))?;

    let mut result = Vec::new();
    for row in rows {
        match row {
            Ok(r) => result.push(r),
            Err(e) => eprintln!("Skipping cookie row: {e}"),
        }
    }
    Ok(result)
}

#[cfg(target_os = "macos")]
fn get_safe_storage_key() -> Result<Vec<u8>, String> {
    get_generic_password("Microsoft Teams Safe Storage", "Microsoft Teams")
        .map_err(|e| format!("Failed to read keychain: {e}"))
}

#[cfg(target_os = "macos")]
fn cached_decryption_key_path() -> Result<PathBuf, String> {
    let base_dir = dirs::data_local_dir()
        .ok_or_else(|| "Failed to resolve local data directory".to_string())?;
    Ok(base_dir
        .join("Better Teams")
        .join("teams-safe-storage-key.bin"))
}

#[cfg(target_os = "macos")]
fn read_cached_decryption_key() -> Result<Option<[u8; PBKDF2_KEY_LENGTH]>, String> {
    let path = cached_decryption_key_path()?;
    if !path.exists() {
        return Ok(None);
    }

    let bytes =
        fs::read(&path).map_err(|e| format!("Failed to read cached Teams decryption key: {e}"))?;
    if bytes.len() != PBKDF2_KEY_LENGTH {
        return Ok(None);
    }

    let mut key = [0u8; PBKDF2_KEY_LENGTH];
    key.copy_from_slice(&bytes);
    Ok(Some(key))
}

#[cfg(target_os = "macos")]
fn write_cached_decryption_key(key: &[u8; PBKDF2_KEY_LENGTH]) -> Result<(), String> {
    let path = cached_decryption_key_path()?;
    let parent = path
        .parent()
        .ok_or_else(|| "Failed to resolve cache directory".to_string())?;

    fs::create_dir_all(parent).map_err(|e| format!("Failed to create cache directory: {e}"))?;
    fs::write(&path, key).map_err(|e| format!("Failed to cache Teams decryption key: {e}"))?;

    #[cfg(unix)]
    fs::set_permissions(&path, fs::Permissions::from_mode(0o600))
        .map_err(|e| format!("Failed to secure cached Teams decryption key: {e}"))?;

    Ok(())
}

fn derive_decryption_key(safe_storage_key: &[u8]) -> [u8; PBKDF2_KEY_LENGTH] {
    let mut key = [0u8; PBKDF2_KEY_LENGTH];
    pbkdf2_hmac::<Sha1>(safe_storage_key, PBKDF2_SALT, PBKDF2_ITERATIONS, &mut key);
    key
}

fn decrypt_value(encrypted: &[u8], key: &[u8; PBKDF2_KEY_LENGTH]) -> Result<String, String> {
    if encrypted.is_empty() {
        return Ok(String::new());
    }

    // Check for "v10" prefix
    if encrypted.len() < 3 || &encrypted[..3] != ENCRYPTED_PREFIX {
        // Not encrypted – return as-is
        return Ok(String::from_utf8_lossy(encrypted).to_string());
    }

    let ciphertext = &encrypted[3..];
    if ciphertext.is_empty() {
        return Ok(String::new());
    }

    // IV = 16 space (0x20) bytes
    let iv: [u8; 16] = [0x20u8; 16];

    // We need a mutable copy because the decryptor works in-place
    let mut buf = ciphertext.to_vec();

    let decryptor = Aes128CbcDec::new(key.into(), &iv.into());
    let decrypted = decryptor
        .decrypt_padded_mut::<aes::cipher::block_padding::Pkcs7>(&mut buf)
        .map_err(|e| format!("Decryption failed: {e}"))?;

    // Use lossy conversion — the decrypted cookie may contain non-UTF-8 padding
    // bytes before the actual JWT. The JWT is extracted via regex afterwards.
    Ok(String::from_utf8_lossy(decrypted).into_owned())
}

fn extract_jwt(raw: &str) -> String {
    let jwt_part = r"eyJ[A-Za-z0-9_-]+\.eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+";

    // Try Bearer%3D first
    let pattern = format!(r"Bearer%3D({jwt_part})");
    if let Ok(re) = Regex::new(&pattern) {
        if let Some(caps) = re.captures(raw) {
            if let Some(m) = caps.get(1) {
                return m.as_str().to_string();
            }
        }
    }

    // Try Bearer%20
    let pattern = format!(r"Bearer%20({jwt_part})");
    if let Ok(re) = Regex::new(&pattern) {
        if let Some(caps) = re.captures(raw) {
            if let Some(m) = caps.get(1) {
                return m.as_str().to_string();
            }
        }
    }

    // Try raw JWT
    let pattern = format!(r"({jwt_part})");
    if let Ok(re) = Regex::new(&pattern) {
        if let Some(caps) = re.captures(raw) {
            if let Some(m) = caps.get(1) {
                return m.as_str().to_string();
            }
        }
    }

    String::new()
}

fn decode_jwt_payload(token: &str) -> Option<serde_json::Value> {
    let parts: Vec<&str> = token.split('.').collect();
    if parts.len() < 2 {
        return None;
    }
    let payload_bytes = URL_SAFE_NO_PAD.decode(parts[1]).ok()?;
    serde_json::from_slice(&payload_bytes).ok()
}

fn unix_to_iso8601(ts: i64) -> String {
    // Convert Unix timestamp (seconds) to ISO 8601 string
    let secs = ts;
    let days_since_epoch = secs / 86400;
    let time_of_day = secs % 86400;

    // Calculate date from days since 1970-01-01
    let mut remaining_days = days_since_epoch;
    let mut year: i64 = 1970;

    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if remaining_days < days_in_year {
            break;
        }
        remaining_days -= days_in_year;
        year += 1;
    }

    let days_in_months: [i64; 12] = if is_leap_year(year) {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut month: usize = 0;
    for (i, &dim) in days_in_months.iter().enumerate() {
        if remaining_days < dim {
            month = i;
            break;
        }
        remaining_days -= dim;
    }

    let day = remaining_days + 1;
    let hours = time_of_day / 3600;
    let minutes = (time_of_day % 3600) / 60;
    let seconds = time_of_day % 60;

    format!(
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}.000Z",
        year,
        month + 1,
        day,
        hours,
        minutes,
        seconds
    )
}

fn is_leap_year(y: i64) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

fn now_unix() -> i64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs() as i64
}

#[cfg(target_os = "macos")]
fn service_worker_cache_path() -> PathBuf {
    let home = std::env::var("HOME")
        .map(PathBuf::from)
        .unwrap_or_else(|_| PathBuf::from("~"));
    home.join("Library/Containers/com.microsoft.teams2/Data/Library/Application Support/Microsoft/MSTeams/EBWebView/WV2Profile_tfw/Service Worker/CacheStorage")
}

fn collect_files_recursive(path: &PathBuf, files: &mut Vec<PathBuf>) -> Result<(), String> {
    let entries = fs::read_dir(path).map_err(|e| format!("Failed to read directory: {e}"))?;
    for entry in entries {
        let entry = entry.map_err(|e| format!("Failed to read directory entry: {e}"))?;
        let entry_path = entry.path();
        let metadata =
            entry.metadata()
                .map_err(|e| format!("Failed to read metadata for {:?}: {e}", entry_path))?;
        if metadata.is_dir() {
            collect_files_recursive(&entry_path, files)?;
            continue;
        }
        if metadata.is_file() {
            files.push(entry_path);
        }
    }
    Ok(())
}

fn extract_json_object(bytes: &[u8], start: usize) -> Option<&[u8]> {
    if bytes.get(start).copied()? != b'{' {
        return None;
    }

    let mut depth = 0usize;
    let mut in_string = false;
    let mut escaped = false;

    for (index, byte) in bytes.iter().enumerate().skip(start) {
        if in_string {
            if escaped {
                escaped = false;
                continue;
            }
            match *byte {
                b'\\' => escaped = true,
                b'"' => in_string = false,
                _ => {}
            }
            continue;
        }

        match *byte {
            b'"' => in_string = true,
            b'{' => depth += 1,
            b'}' => {
                depth = depth.checked_sub(1)?;
                if depth == 0 {
                    return Some(&bytes[start..=index]);
                }
            }
            _ => {}
        }
    }

    None
}

fn parse_cached_presence_entry(bytes: &[u8]) -> Option<CachedPresenceEntry> {
    let marker_index = bytes
        .windows(CDL_WORKER_AGGREGATED_USER_DATA_PREFIX.len())
        .position(|window| window == CDL_WORKER_AGGREGATED_USER_DATA_PREFIX.as_bytes())?;
    let json_start = bytes
        .iter()
        .enumerate()
        .skip(marker_index + CDL_WORKER_AGGREGATED_USER_DATA_PREFIX.len())
        .find_map(|(index, byte)| (*byte == b'{').then_some(index))?;
    let json_bytes = extract_json_object(bytes, json_start)?;
    let envelope: AggregatedUserDataEnvelope = serde_json::from_slice(json_bytes).ok()?;
    let presence_envelope = envelope.presence?;
    let mri = presence_envelope.mri?;
    let presence = presence_envelope.presence?;
    let availability = presence.availability.as_deref().unwrap_or_default().trim();
    let activity = presence.activity.as_deref().unwrap_or_default().trim();

    if mri.trim().is_empty() || (availability.is_empty() && activity.is_empty()) {
        return None;
    }

    Some(CachedPresenceEntry { mri, presence })
}

#[cfg(target_os = "macos")]
fn scan_all_presence() -> Result<Vec<CachedPresenceEntry>, String> {
    let cache_path = service_worker_cache_path();
    if !cache_path.exists() {
        return Ok(Vec::new());
    }

    let mut files = Vec::new();
    collect_files_recursive(&cache_path, &mut files)?;

    let mut by_mri = std::collections::HashMap::<String, CachedPresenceEntry>::new();

    for file in files {
        let bytes = match fs::read(&file) {
            Ok(bytes) => bytes,
            Err(_) => continue,
        };
        let Some(entry) = parse_cached_presence_entry(&bytes) else {
            continue;
        };
        let key = entry.mri.trim().to_ascii_lowercase();
        by_mri.insert(key, entry);
    }

    Ok(by_mri.into_values().collect())
}

#[cfg(target_os = "macos")]
fn filter_presence(
    all_entries: &[CachedPresenceEntry],
    user_mris: &[String],
) -> Vec<CachedPresenceEntry> {
    let entry_by_mri: std::collections::HashMap<String, &CachedPresenceEntry> = all_entries
        .iter()
        .map(|e| (e.mri.trim().to_ascii_lowercase(), e))
        .collect();

    let mut result = Vec::new();
    for user_mri in user_mris {
        let key = user_mri.trim().to_ascii_lowercase();
        if let Some(entry) = entry_by_mri.get(&key) {
            result.push((*entry).clone());
        }
    }
    result
}

const PRESENCE_CACHE_TTL_SECS: u64 = 30;

#[cfg(target_os = "macos")]
fn extract_cached_presence(user_mris: &[String]) -> Result<Vec<CachedPresenceEntry>, String> {
    let mut cache = PRESENCE_CACHE
        .lock()
        .map_err(|e| format!("Failed to lock presence cache: {e}"))?;

    let all_entries = if let Some((timestamp, ref entries)) = *cache {
        if timestamp.elapsed().as_secs() < PRESENCE_CACHE_TTL_SECS {
            return Ok(filter_presence(entries, user_mris));
        }
        // Cache expired, re-scan
        let entries = scan_all_presence()?;
        let result = filter_presence(&entries, user_mris);
        *cache = Some((Instant::now(), entries));
        return Ok(result);
    } else {
        scan_all_presence()?
    };

    let result = filter_presence(&all_entries, user_mris);
    *cache = Some((Instant::now(), all_entries));
    Ok(result)
}

#[cfg(not(target_os = "macos"))]
fn extract_cached_presence(_user_mris: &[String]) -> Result<Vec<CachedPresenceEntry>, String> {
    Ok(Vec::new())
}

#[cfg(target_os = "macos")]
fn extract_tokens_with_key(
    rows: &[CookieRow],
    decryption_key: &[u8; PBKDF2_KEY_LENGTH],
) -> TokenExtractionOutcome {
    let now = now_unix();
    let mut tokens: Vec<ExtractedToken> = Vec::new();
    let mut decrypt_failures = 0usize;

    for row in rows {
        let decrypted = match decrypt_value(&row.encrypted_value, decryption_key) {
            Ok(v) => v,
            Err(e) => {
                eprintln!("Failed to decrypt cookie {}: {e}", row.name);
                decrypt_failures += 1;
                continue;
            }
        };

        let jwt = extract_jwt(&decrypted);
        if jwt.is_empty() {
            continue;
        }

        let payload = decode_jwt_payload(&jwt);

        let exp = payload
            .as_ref()
            .and_then(|p| p.get("exp"))
            .and_then(|v| v.as_i64())
            .unwrap_or(0);

        // Filter out expired tokens
        if exp < now {
            continue;
        }

        let upn = payload
            .as_ref()
            .and_then(|p| p.get("upn"))
            .and_then(|v| v.as_str())
            .map(|s| s.to_string());

        let tenant_id = payload
            .as_ref()
            .and_then(|p| p.get("tid"))
            .and_then(|v| v.as_str())
            .map(|s| s.to_string());

        let skype_id = payload
            .as_ref()
            .and_then(|p| p.get("skypeid"))
            .and_then(|v| v.as_str())
            .map(|s| s.to_string());

        let audience = payload
            .as_ref()
            .and_then(|p| p.get("aud"))
            .and_then(|v| v.as_str())
            .map(|s| s.to_string());

        let expires_at = unix_to_iso8601(exp);

        tokens.push(ExtractedToken {
            host: row.host_key.clone(),
            name: row.name.clone(),
            token: jwt,
            audience,
            upn,
            tenant_id,
            skype_id,
            expires_at,
        });
    }

    // Sort by expiry descending (the DB already orders by expires_utc DESC,
    // but we re-sort by JWT exp to be safe)
    tokens.sort_by(|a, b| b.expires_at.cmp(&a.expires_at));

    TokenExtractionOutcome {
        tokens,
        decrypt_failures,
    }
}

#[cfg(target_os = "macos")]
fn extract_all_tokens() -> Result<Vec<ExtractedToken>, String> {
    let db_path = cookies_db_path();
    if !db_path.exists() {
        return Err("Teams cookies database not found".to_string());
    }

    let rows = read_cookie_rows(&db_path)?;
    let cached_key = read_cached_decryption_key()?;

    if let Some(key) = cached_key {
        let outcome = extract_tokens_with_key(&rows, &key);
        if !outcome.tokens.is_empty() || outcome.decrypt_failures == 0 {
            return Ok(outcome.tokens);
        }
    }

    let safe_key = get_safe_storage_key()?;
    let decryption_key = derive_decryption_key(&safe_key);
    let outcome = extract_tokens_with_key(&rows, &decryption_key);

    write_cached_decryption_key(&decryption_key)?;

    Ok(outcome.tokens)
}

#[cfg(not(target_os = "macos"))]
fn extract_all_tokens() -> Result<Vec<ExtractedToken>, String> {
    Err("Better Teams token extraction currently supports macOS Teams 2 only".to_string())
}

// ---------- Tauri commands ----------

#[tauri::command]
pub fn extract_tokens() -> Result<Vec<ExtractedToken>, String> {
    extract_all_tokens()
}

#[tauri::command]
pub fn get_auth_token(tenant_id: Option<String>) -> Result<Option<ExtractedToken>, String> {
    let tokens = extract_all_tokens()?;

    let best = tokens.into_iter().find(|t| {
        if t.name != "authtoken" {
            return false;
        }
        match &tenant_id {
            Some(tid) => t.tenant_id.as_deref() == Some(tid.as_str()),
            None => true,
        }
    });

    Ok(best)
}

#[tauri::command]
pub fn get_available_accounts() -> Result<Vec<AccountOption>, String> {
    let tokens = extract_all_tokens()?;

    let mut seen = std::collections::HashSet::new();
    let mut accounts: Vec<AccountOption> = Vec::new();

    for t in tokens {
        if t.name != "authtoken" {
            continue;
        }
        // Deduplicate by (upn, tenant_id)
        let key = (t.upn.clone(), t.tenant_id.clone());
        if seen.contains(&key) {
            continue;
        }
        seen.insert(key);
        accounts.push(AccountOption {
            upn: t.upn,
            tenant_id: t.tenant_id,
        });
    }

    Ok(accounts)
}

#[tauri::command]
pub fn get_cached_presence(user_mris: Vec<String>) -> Result<Vec<CachedPresenceEntry>, String> {
    extract_cached_presence(&user_mris)
}

#[cfg(test)]
mod tests {
    use super::{extract_json_object, parse_cached_presence_entry};

    #[test]
    fn extracts_balanced_json_with_strings() {
        let bytes = br#"xxx{"a":"brace } text","nested":{"b":true}}yyy"#;
        let json = extract_json_object(bytes, 3).expect("json");
        assert_eq!(
            std::str::from_utf8(json).expect("utf8"),
            r#"{"a":"brace } text","nested":{"b":true}}"#
        );
    }

    #[test]
    fn parses_cached_presence_entry_from_binary_blob() {
        let bytes = br#"junkhttps://cdl-worker/cdl-worker-cache-manager/cdl-worker-aggregated-user-data/user-1{"presence":{"mri":"8:orgid:user-1","presence":{"availability":"Available","activity":"InAMeeting"}}}tail"#;
        let entry = parse_cached_presence_entry(bytes).expect("presence entry");
        assert_eq!(entry.mri, "8:orgid:user-1");
        assert_eq!(entry.presence.availability.as_deref(), Some("Available"));
        assert_eq!(entry.presence.activity.as_deref(), Some("InAMeeting"));
    }
}
