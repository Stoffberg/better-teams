## Architecture
- Dependency direction is `desktop-electron -> app -> core/ui`.
- `core` imports no renderer, React, DOM, Electron, preload, or app code.
- `ui` contains reusable primitives only; it imports no app feature code.
- `app` owns React renderer state, feature controllers, routes, providers, and React Query orchestration.
- `desktop-electron` owns Electron main/preload/runtime adapters, durable token/image/presence storage, HTTP bridge, and packaging entry points.
- IPC contracts live in `packages/desktop-electron/src/preload/contracts.ts` and must validate payloads at the boundary.

## Boundaries
- Components render.
- Feature hooks/controllers orchestrate.
- Services talk to Teams.
- Durable desktop state stays in `desktop-electron`; React Query stays tenant-scoped in `app`.
- New `src/lib` junk drawers are not allowed. Put code in the owning package and feature/domain module.

## Imports
- Use `@better-teams/core/*` for Teams domain, parsing, schemas, clients, and query helpers.
- Use `@better-teams/ui/*` for shared UI primitives.
- Use `@better-teams/app/*` only inside the renderer package.
- Use `@better-teams/desktop-electron/*` only inside the desktop package or renderer type augmentation.

## Tooling
- Use `bun` for package-manager operations.
- Do not edit dependency lists by hand.
- Keep Electron Forge for desktop packaging unless a separate build-system migration is explicitly planned.
