# Security Policy

## Supported Versions

Security fixes target the default branch. Tagged releases are snapshots, not long-term support lines.

## Reporting a Vulnerability

Report vulnerabilities privately through GitHub Security Advisories for this repository.

Do not open public issues for:

- Token extraction bugs.
- Cookie database handling.
- Electron IPC or preload bridge escapes.
- Local file or image cache access.
- Anything that exposes chat content, account identifiers, tokens, or Teams session data.

Include:

- A short impact summary.
- Reproduction steps.
- Affected platform and Better Teams version or commit.
- Any logs with tokens, cookies, tenant IDs, user IDs, and private chat content removed.

Expected handling:

- Acknowledgement as soon as practical.
- Triage on the default branch first.
- Patch, release notes, and advisory publication once the fix is available.
