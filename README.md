# Better Teams

[![CI](https://github.com/Stoffberg/better-teams/actions/workflows/ci.yml/badge.svg)](https://github.com/Stoffberg/better-teams/actions/workflows/ci.yml)
[![CodeQL](https://github.com/Stoffberg/better-teams/actions/workflows/codeql.yml/badge.svg)](https://github.com/Stoffberg/better-teams/actions/workflows/codeql.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A calmer, faster Microsoft Teams desktop app with a focused custom UI.

Better Teams is an Electron desktop client for people who want Teams chat without the weight of the full Teams interface. It reads the local macOS Teams 2 session, renders conversations in a lean React UI, and keeps the app shaped around chat, people, presence, and fast switching.

This project is independent and is not affiliated with, endorsed by, or supported by Microsoft.

## Status

Better Teams is currently macOS-first and expects Microsoft Teams 2 to be installed and signed in on the same machine. Token extraction, presence cache reads, and packaged-app validation are built around that environment.

## Features

- Focused chat workspace with fast conversation switching.
- Local account discovery from the signed-in Teams 2 profile.
- Rich message rendering for links, mentions, reactions, images, and grouped messages.
- Presence and profile cards from Teams data already available on the machine.
- File-backed image caching through a constrained Electron asset protocol.
- Tray behavior for a desktop-native app loop.

## Requirements

- macOS.
- Microsoft Teams 2 installed and signed in.
- Bun 1.1 or newer.
- Xcode command line tools for native Electron dependencies.

## Development

Install dependencies:

```sh
bun install
```

Start the app:

```sh
bun run dev
```

Run the normal verification stack:

```sh
bun run verify
```

Useful targeted commands:

```sh
bun run typecheck
bun run check
bun run test
bun run test:coverage
bun run package
bun run test:e2e
```

The packaged smoke test expects `out/` to exist, so run `bun run package` before `bun run test:e2e`.

## Release Builds

Local macOS packaging:

```sh
bun run make
```

Tagged releases are handled by GitHub Actions. Push a `v*` tag, let the release workflow build the DMG/ZIP assets, and publish the generated draft release after checking the artifacts.

## Security

Better Teams touches local Teams session data. Do not paste tokens, cookie database contents, screenshots with private chat data, or workspace-specific credentials into public issues.

Report security problems privately through GitHub Security Advisories. See [SECURITY.md](SECURITY.md).

## Contributing

Keep changes tight and behavior-backed. Use Bun or PNPM, not npm. Before opening a PR, run:

```sh
bun run verify
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for the project workflow.

## License

MIT. See [LICENSE](LICENSE).
