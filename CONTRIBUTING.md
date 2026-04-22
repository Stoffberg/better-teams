# Contributing

## Workflow

Use Bun for project commands:

```sh
bun install
bun run dev
```

Before opening a PR:

```sh
bun run verify
```

Targeted checks:

```sh
bun run typecheck
bun run check
bun run test
bun run test:coverage
bun run package
bun run test:e2e
```

## Pull Requests

Keep PRs small enough to review properly. Explain the behavior change, why it exists, and what you verified.

For UI changes, include screenshots or a short recording when the diff affects layout, rendering, or interaction.

For Electron, auth, cache, IPC, and Teams API changes, describe the blast radius. These areas can leak data or break startup in ways tests will not fully catch.

## Project Rules

- Use Bun or PNPM, not npm.
- Do not edit generated output.
- Do not commit tokens, Teams cookie data, local account IDs, or private chat content.
- Do not lower validation just to make a PR pass.
- Prefer behavior tests over coverage theater.
