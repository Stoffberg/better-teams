use sha1::{Digest, Sha1};
use std::fs;
use std::path::{Path, PathBuf};
use tauri::{AppHandle, Manager};

const IMAGE_CACHE_DIR: &str = "images";

fn image_cache_dir(app: &AppHandle) -> Result<PathBuf, String> {
    let dir = app
        .path()
        .app_cache_dir()
        .map_err(|e| format!("Failed to resolve app cache dir: {e}"))?
        .join(IMAGE_CACHE_DIR);
    fs::create_dir_all(&dir).map_err(|e| format!("Failed to create image cache dir: {e}"))?;
    Ok(dir)
}

fn normalized_extension(extension: Option<String>) -> String {
    match extension
        .as_deref()
        .map(|value| value.trim().to_ascii_lowercase())
        .as_deref()
    {
        Some("jpg") | Some("jpeg") => "jpg".to_string(),
        Some("png") => "png".to_string(),
        Some("gif") => "gif".to_string(),
        Some("webp") => "webp".to_string(),
        Some("avif") => "avif".to_string(),
        _ => "img".to_string(),
    }
}

fn hashed_filename(cache_key: &str, extension: Option<String>) -> String {
    let digest = Sha1::digest(cache_key.as_bytes());
    let hex = format!("{digest:x}");
    format!("{hex}.{}", normalized_extension(extension))
}

fn is_within_dir(path: &Path, dir: &Path) -> bool {
    path.starts_with(dir)
}

#[tauri::command]
pub fn cache_image_file(
    app: AppHandle,
    cache_key: String,
    bytes: Vec<u8>,
    extension: Option<String>,
) -> Result<String, String> {
    let dir = image_cache_dir(&app)?;
    let path = dir.join(hashed_filename(&cache_key, extension));
    fs::write(&path, bytes).map_err(|e| format!("Failed to write cached image file: {e}"))?;
    Ok(path.to_string_lossy().into_owned())
}

#[tauri::command]
pub fn remove_cached_image_files(app: AppHandle, paths: Vec<String>) -> Result<(), String> {
    let dir = image_cache_dir(&app)?;
    for path in paths {
        let candidate = PathBuf::from(path);
        if !is_within_dir(&candidate, &dir) {
            continue;
        }
        match fs::remove_file(&candidate) {
            Ok(()) => {}
            Err(err) if err.kind() == std::io::ErrorKind::NotFound => {}
            Err(err) => {
                return Err(format!("Failed to remove cached image file: {err}"));
            }
        }
    }
    Ok(())
}
