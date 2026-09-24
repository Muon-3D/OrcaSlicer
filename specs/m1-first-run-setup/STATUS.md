# Status

This page tracks the work packages in [10-work-plan.md](10-work-plan.md). The coordinating session updates it at each check-in.

Agents: the coordinator can't receive messages from you. Tell it where you are by pushing branches and opening PRs whose titles start with the package ID. Put questions in the PR body under **Spec questions**.

**Last updated:** 24 Sep 2026, 17:25 UTC. MR-1, OS-7 and ML-1 have no new commits since 16:36; their fixes are still pending.

## Packages with work

| Package | Repo | Branch / PR | State | Notes |
|---|---|---|---|---|
| OR-1 | OrcaSlicer | [Muon-3D/OrcaSlicer#3](https://github.com/Muon-3D/OrcaSlicer/pull/3) | In review | Matches the spec. The red "Check profiles" isn't caused by this PR: it's red on `main` (run 36002278074) and has been on every run since January. The PR's second commit fixes the system-profile step. The custom-preset step still fails because upstream vendors that this fork lacks are missing. |
| OR-2 | OrcaSlicer | [Muon-3D/OrcaSlicer#4](https://github.com/Muon-3D/OrcaSlicer/pull/4) | In review | Matches the spec. It uses `/#/setup`, which works before OS-1's redirect lands. Linux build passed; macOS still running. The Windows build failed on a deps cache miss (`fail-on-cache-miss`) before any code compiled, and "Build all" is also red on `main`. |
| MR-6 | Moonraker | [Muon-3D/Moonraker#20](https://github.com/Muon-3D/Moonraker/pull/20) (opened before this spec) | In review | Already implements `GET /server/muon/link`, `POST …/start` and `POST …/cancel`, bridging to muon-link at `127.0.0.1:7131`. Confirm is deliberately **not** in Moonraker (ADR 0018, LINK-3). Build on this PR; don't write a second `muon_link`. |
| MR-1 | Moonraker | branch `claude/m1-first-run-setup` @ `0c11719` (no PR yet) | **Changes needed** | A solid core. It passes request headers through for the CSRF and Host checks, floors `reset`, and has fakes and tests. It was written against an older spec, so it needs these changes: **(1)** marker calls become `GET /setup/complete`, and `POST /setup/complete {"by":"muon_setup"}` at `finish`; `reset` calls `DELETE /setup/complete`. **(2)** Add `/server/aux/setup/complete` to `FLOOR_PREFIXES`, with a test. **(3)** Region: `state.region` must be Aux `GET /region` passed through as-is, plus the derived `market`, and `options.region` must be `GET /region/options` as-is. Drop the `for_language`, `all`, `support_code` and `source` reshaping in `region.py`, and carry `network.region_confirmed` (02 §5.2, §5.6, §6). **(4)** Read the station count from `POST /wifi/ap/count` only; `/wifi/ap/stations` was dropped, and polling it would 404 every 2 s. |
| OS-7 | MuonOS | branch `feat/KAN-413-setup-marker` @ `9bcb4cc` (no PR yet) | **Adopted as the marker contract** | `GET`, `POST` and `DELETE /setup/complete`, persisted by `muon3d-setup.toml`, with the ID-9 and data-inventory rows. The spec now follows it (03 §7). Still to add: `/server/aux/setup/complete` in `fluidd.nginx.template`'s 403 list and in `test_floor.py`, paired with MR-1's floor change. Coordinate with #174, whose own `setup_routes.py`/`setup.toml` under `/setup` would conflict. |
| ML-1 | muon-link | branch `feat/ML-1-link-contract` @ `b9f9678`, stacked on #24 (no PR yet) | **Matches the spec** | `docs/ACCOUNT_LINK_ADMIN.md` and `tests/link_admin_contract.rs` pin the same `LinkPhase` shapes and 409s as 02 §5.9. |
| OS-1 | MuonOS | branch `feat/KAN-410-captive-portal` @ `c3e06c5` (no PR yet) | **Changes needed** | Good: the dnsmasq wildcard drop-in, the nginx maps, `/setup` → `/#/setup`, the `tcp/443` reset on `ap0`, and tests. Missing from the spec: **(1)** D11 forwarding isolation. `10-ap-isolate` must reject **all** forwarding from `ap0`, including to `eth0` subnets, and run on `eth0` events. Today hotspot clients are NATed out and can reach an Ethernet LAN. The new dnsmasq comment claims "the hotspot never routes anywhere", which isn't true until this lands. **(2)** Pin `ipv4.addresses=10.42.0.1/24` in `ap0-con`, with the persisted-profile migration. Both the nginx map and the QR codes depend on that address. |
| FL-7 | Fluidd | branch `fix/KAN-408-app-host-self-probe` @ `d6d7970` (no PR yet) | **Matches the spec** | Adds `app.muon3d.com` to both blacklists. It also returns an empty `ApiConfig` at once when no endpoint is left, which removes the leftover 5 s sleep. Tests added. |

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

- **15:10: the account link corrected against the real code.** It's muon-link PR #24 (`/link/*`), not `/pairing/*`, which is client pairing. Moonraker PR #20 is MR-6; build on it. Confirm and unlink are called by the panel directly and never through Moonraker. The orchestrator makes the code, and `remote.link` mirrors muon-link's `LinkPhase` as-is (`url`, integer `expires_at`). `self_hosted` is out of phase 1. New packages: ML-2 (proof of a knob press on confirm) and OS-10 (nginx `/muon-link/` on `:100`, and the orchestrator env vars). Affects 02 §5.9 and §9, 03 §6, 04 P12, 05 S6, 07 S3–S4, 10, and the fixtures.

- **15:40: MuonOS and MuonUI checked against the real code.** The main changes:
  - **Region:** confirmed **after** the join, on draft MuonOS#174's `/region` shapes and MuonUI#31's `regionPromptFor` and `RegionPicker`. `/region/suggest` is gone.
  - **Marker:** #174's `setup.json` via `GET` and `POST /setup`. *Superseded at 16:40 by KAN-413's `/setup/complete`, below.*
  - **Hotspot rules:** these match the real `muon-ap-lifecycle.sh`, with a new `ap-hotspot-kept-on` file and a transient auto-off timer.
  - **Captive portal:** the nginx snippet now passes `nginx -t`; D11 makes the hotspot reach the printer only.
  - **Wi-Fi join:** a `status: "restored"` response now counts as a failure.
  - **Internet check:** reuses the Nexigon check.
  - **Time:** new sudoers entries and a time-zone file.
  - **Panel:** follows MuonUI's real conventions: Back is a left-rail item, the theme is light, the existing keyboard and `UpdatingOverlay` are reused, and i18n is added as UI-0.
  - **New packages:** OS-11 (**urgent:** Wi-Fi PSKs are logged to the persisted journal by sudo) and UI-0.
  - **Also recorded:** no unit can declare a region yet, because there are no signing keys.

- **16:40: the marker contract follows MuonOS KAN-413.** It's `GET`, `POST` and `DELETE /setup/complete` over `/var/lib/muon3d/setup/complete`, not #174's `/setup`. The whole `/server/aux/setup/complete` prefix is floored in both Moonraker and the MuonOS nginx list. Affects 01 §7, 02 §1, §4 and §5.11, 03 §1 and §7, 10 (OS-7, MR-1), and the README.

## Not started yet

No branches seen yet for MR-2–MR-5, MR-7–MR-9, OS-2–OS-6, OS-8–OS-11, UI-*, FL-1–FL-6, FL-8, FL-9, CON-1, or QA-*. **OS-11 (Wi-Fi passwords in the journal) is urgent and still unassigned.**
