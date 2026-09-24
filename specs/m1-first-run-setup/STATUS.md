# Status

This page tracks the work packages in [10-work-plan.md](10-work-plan.md). The coordinating session updates it at each check-in.

Agents: the coordinator can't receive messages from you. Tell it where you are by pushing branches and opening PRs whose titles start with the package ID. Put questions in the PR body under **Spec questions**.

**Last updated:** 24 Sep 2026, 15:00 UTC

## Packages with work

| Package | Repo | Branch / PR | State | Notes |
|---|---|---|---|---|
| OR-1 | OrcaSlicer | [Muon-3D/OrcaSlicer#3](https://github.com/Muon-3D/OrcaSlicer/pull/3) | In review | Matches the spec. The red "Check profiles" isn't caused by this PR: it's red on `main` (run 36002278074) and has been on every run since January. The PR's second commit fixes the system-profile step. The custom-preset step still fails because upstream vendors that this fork lacks are missing. |
| OR-2 | OrcaSlicer | [Muon-3D/OrcaSlicer#4](https://github.com/Muon-3D/OrcaSlicer/pull/4) | In review | Matches the spec. It uses `/#/setup`, which works before OS-1's redirect lands. Linux build passed; macOS still running. The Windows build failed on a deps cache miss (`fail-on-cache-miss`) before any code compiled, and "Build all" is also red on `main`. |
| MR-6 | Moonraker | [Muon-3D/Moonraker#20](https://github.com/Muon-3D/Moonraker/pull/20) (opened before this spec) | In review | Already implements `GET /server/muon/link`, `POST …/start` and `POST …/cancel`, bridging to muon-link at `127.0.0.1:7131`. Confirm is deliberately **not** in Moonraker (ADR 0018, LINK-3). Build on this PR; don't write a second `muon_link`. |

## In-flight work that overlaps the spec

Build on these PRs rather than around them.

| PR | Overlaps |
|---|---|
| [Muon-3D/muon-link#24](https://github.com/Muon-3D/muon-link/pull/24): account link and orchestrator connection | ML-1, MR-5, and UI-3 P12 |
| [Muon-3D/MuonUI#31](https://github.com/Muon-3D/MuonUI/pull/31) (draft): region prompt before joining | UI-2 P5b, KAN-324 |
| [Muon-3D/MuonUI#39](https://github.com/Muon-3D/MuonUI/pull/39): Aux Wi-Fi overhaul (KAN-339) | UI-2 |
| [Muon-3D/MuonUI#47](https://github.com/Muon-3D/MuonUI/pull/47) and [Muon-3D/Moonraker#21](https://github.com/Muon-3D/Moonraker/pull/21): SEC-8 protection levels | 07 §3, where setup writes join the protected set |
| [Muon-3D/MuonOS#300](https://github.com/Muon-3D/MuonOS/pull/300): hotspot `security_enabled` (KAN-376) | OS-5, FL-8 |
| [Muon-3D/MuonOS#305](https://github.com/Muon-3D/MuonOS/pull/305): GATE-1 listener declarations | OS-1 |

## Not started yet

No branches seen yet for MR-1–MR-5, MR-7–MR-9, OS-*, ML-1, UI-*, FL-*, CON-1, or QA-*.
