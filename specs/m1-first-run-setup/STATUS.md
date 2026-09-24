# Status

This page tracks the work packages in [10-work-plan.md](10-work-plan.md). The coordinating session updates it at each check-in.

Agents: the coordinator can't receive messages from you. Tell it where you are by pushing branches and opening PRs whose titles start with the package ID. Put questions in the PR body under **Spec questions**.

**Last updated:** 24 Sep 2026, 16:00 UTC

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
| [Muon-3D/MuonUI#31](https://github.com/Muon-3D/MuonUI/pull/31) (draft): setup route, region picker and pre-join region prompt | UI-1 builds on its branch; UI-2 P4 and P7a; KAN-324 |
| [Muon-3D/MuonUI#39](https://github.com/Muon-3D/MuonUI/pull/39): Aux Wi-Fi overhaul (KAN-339) | UI-2 |
| [Muon-3D/MuonUI#47](https://github.com/Muon-3D/MuonUI/pull/47) and [Muon-3D/Moonraker#21](https://github.com/Muon-3D/Moonraker/pull/21): SEC-8 protection levels | 07 §3, where setup writes join the protected set |
| [Muon-3D/MuonOS#300](https://github.com/Muon-3D/MuonOS/pull/300): hotspot `security_enabled` (KAN-376) | OS-5, FL-8 |
| [Muon-3D/MuonOS#305](https://github.com/Muon-3D/MuonOS/pull/305): GATE-1 listener declarations | OS-1 |

## Spec changes since v1

- **15:20: the account link corrected against the real code.** It's muon-link PR #24 (`/link/*`), not `/pairing/*`, which is client pairing. Moonraker PR #20 is MR-6; build on it. Confirm and unlink are called by the panel directly and never through Moonraker. The orchestrator makes the code, and `remote.link` mirrors muon-link's `LinkPhase` as-is (`url`, integer `expires_at`). `self_hosted` is out of phase 1. New packages: ML-2 (proof of a knob press on confirm) and OS-10 (nginx `/muon-link/` on `:100`, and the orchestrator env vars). Affects 02 §5.9 and §9, 03 §6, 04 P12, 05 S6, 07 S3–S4, 10, and the fixtures.

- **16:00: MuonOS and MuonUI checked against the real code.** The main changes:
  - **Region:** confirmed **after** the join, on draft MuonOS#174's `/region` shapes and MuonUI#31's `regionPromptFor` and `RegionPicker`. `/region/suggest` is gone.
  - **Marker:** #174's `setup.json`, read and written with `GET` and `POST /setup`.
  - **Hotspot rules:** these match the real `muon-ap-lifecycle.sh`, with a new `ap-hotspot-kept-on` file and a transient auto-off timer.
  - **Captive portal:** the nginx snippet now passes `nginx -t`; D11 makes the hotspot reach the printer only.
  - **Wi-Fi join:** a `status: "restored"` response now counts as a failure.
  - **Internet check:** reuses the Nexigon check.
  - **Time:** new sudoers entries and a time-zone file.
  - **Panel:** follows MuonUI's real conventions: Back is a left-rail item, the theme is light, the existing keyboard and `UpdatingOverlay` are reused, and i18n is added as UI-0.
  - **New packages:** OS-11 (**urgent:** Wi-Fi PSKs are logged to the persisted journal by sudo) and UI-0.
  - **Also recorded:** no unit can declare a region yet, because there are no signing keys.

## Not started yet

No branches seen yet for MR-1–MR-5, MR-7–MR-9, OS-*, ML-1, UI-*, FL-*, CON-1, or QA-*.
