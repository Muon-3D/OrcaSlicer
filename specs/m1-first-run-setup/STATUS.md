# Status

This page tracks the work packages in [10-work-plan.md](10-work-plan.md). The coordinating session updates it at each check-in.

Agents: the coordinator can't receive messages from you. Tell it where you are by pushing branches and opening PRs whose titles start with the package ID. Put questions in the PR body under **Spec questions**. **Link the spec commit you built against** (10 §1 rule 7). Several PRs were built on the first version (`42e8d51`) and have to catch up: check the log at the end of this page and [fixtures/README.md](fixtures/README.md).

**Last updated:** 25 Sep 2026, 13:35 UTC. Review comments went out on 16 PRs at about 04:40; since then no package agent has pushed or replied. Moonraker#27 merged, MuonOS#316 re-recorded its contract, and the WS-13 decisions (OrcaSlicer#6) are in the spec.

## Decisions needed from the owner

1. **Remote Wi-Fi writes.** Once FL-9 puts Fluidd's Aux client on Iroh, a remote caller can change the printer's Wi-Fi and strand it offline. Should remote writes to `/server/aux/wifi/*` be refused? (07 §3)
2. **KAN-341's AP-4.** OS-5 (MuonOS#314) makes the hotspot go off once a set-up printer is connected, which reverses KAN-341's "on unless disabled". It needs Jack's agreement.
3. **Bench runs (C3).** MuonOS#312, #313, #314 and #315 all fail the PR-policy check until someone runs them on a unit and links the record in `Evidence:`.
4. **OS-11 is still unassigned.** Wi-Fi passwords are written to the persisted journal. Urgent, and independent of setup.
5. **FL-9's hardware check** needs someone with a linked unit.

## Merge order

- **Moonraker:** #27 (merged 25 Sep) → MR-1 #22 → MR-9 #23 → MR-2 #24 → MR-4 #26 (needs MR-2's clock). MR-7 #28 goes after #22. **#25 goes before any MuonOS PR that adds the routes it floors.**
- **MuonOS pin bump:** one PR bumps the Moonraker pin past #25 and #22 and, in the same change, adds every new floor entry to `EXPECTED_FLOOR` in `test_trusted_clients.py`, the `:80` vhost's 403 list and `test_floor.py` (02 §1). Nobody has opened it yet.
- **MuonOS:** #313 (OS-7) waits for that pin bump. #314 (OS-5) also waits for a pinned Moonraker that writes the marker on migration, or every field unit's hotspot is forced on. #312 (OS-1) waits until the pinned Fluidd serves `/setup`. #316 is independent.
- **MuonUI:** UI-1 #48 moves onto #31's branch, so it lands after #39 and #31.

## Packages with work

### Moonraker (`Muon-3D/Moonraker`)

| Package | PR | State | Notes |
|---|---|---|---|
| MR-1 | [#22](https://github.com/Muon-3D/Moonraker/pull/22) @ `d0cd6f2` | **Changes needed** | The marker change is done. **Still open from the first review:** region pass-through plus `region_confirmed` (blocking). **New:** migration must write the marker with `by: migrated`, a reset can be undone by the marker retry, drop the #174 fallback, take `hotspot` from `GET /wifi/ap/stations`, and the websocket Host check. The floor entry moved to #25. Lint passes; CI runs no pytest (144 local tests pass). |
| MR-9 | [#23](https://github.com/Muon-3D/Moonraker/pull/23) @ `fd79d45` | Nearly there | Matches 02 §7. Map only "not answering" to 503, not a real Aux 500. Resolve the test-file conflict with #27. |
| MR-2 | [#24](https://github.com/Muon-3D/Moonraker/pull/24) @ `97d4d99`, on #23 | **Changes needed** | Clock only from the hotspot; zones from `zone.tab`; apply `tz` even when the clock is refused; map Aux 422 to `invalid_timezone`; Keep must not store a null name. Its Aux time assumptions match OS-6. |
| — | [#25](https://github.com/Muon-3D/Moonraker/pull/25) @ `254fc71` | **Changes needed** | The Moonraker half of OS-5/6/7's floor. Owns the floor entries. Change `/server/aux/setup/complete` to `/server/aux/setup`, add `/server/aux/time`, and add `check_floor` tests. |
| MR-4 | [#26](https://github.com/Muon-3D/Moonraker/pull/26) @ `3b88cbe` | **Changes needed** | **Blocking:** every successful update is recorded as `update_failed`, because it judges from update_manager's cached version. Also map `printer_busy`, keep following Aux after `OtaDeploy.update()` returns, and make the update check wait. |
| — | [#27](https://github.com/Muon-3D/Moonraker/pull/27) (KAN-403) | **Merged** into `master` (`1a4be94`) | EndpointId in identity. #23 now resolves its test-file conflict and pins `endpoint_id`. |
| MR-7 | [#28](https://github.com/Muon-3D/Moonraker/pull/28) @ `cbc641c` | **Changes needed** | **Blocking (motion safety):** a paused print isn't busy, and the self-test can run before the clips are confirmed. The "Finish setup" card loses `ready`. Re-check the macro at `start`. |
| MR-6 | [#20](https://github.com/Muon-3D/Moonraker/pull/20) (before this spec) | In review | `muon_link` GET, start and cancel. Confirm is never in Moonraker. |

### MuonOS (`Muon-3D/MuonOS`)

| Package | PR | State | Notes |
|---|---|---|---|
| OS-1 | [#312](https://github.com/Muon-3D/MuonOS/pull/312) @ `6be52c5` | **Changes needed** | `main` merged; the build is green. **Neither code request is done.** Isolation (D11) is now stated as properties in 03 §2: the printer routes nothing, IPv4 and IPv6, loaded before NM, backend-independent, never blocking `10.42.0.1`. Also pin `10.42.0.1` (a new test asserts the opposite), `ipv6.method=disabled` and AP client isolation. PR policy needs C3. |
| OS-7 | [#313](https://github.com/Muon-3D/MuonOS/pull/313) @ `a54cc43` | Waiting on the pin bump | The marker matches 03 §7, and `main` is merged. The floor pairing correctly moved to the pin bump. #174 also persists `/var/lib/muon3d/setup`, under another file name; rename to `setup.toml` or have #174 drop its file. Tell #174. PR policy needs C3. |
| OS-5 | [#314](https://github.com/Muon-3D/MuonOS/pull/314) @ `3d30868` | **Changes needed** | **Blocking:** Ethernet isn't an uplink, and H4 uses `ap-hotspot-requested`, a file older firmware left behind. Also clear the deadline under H1, and re-check the inputs at the end of each run. Its auto-off design and `GET /wifi/ap/stations` were adopted into 03 §1. |
| OS-6 | [#315](https://github.com/Muon-3D/MuonOS/pull/315) @ `02163a0` | **Changes needed** | **Blocking:** the clock accepts dates up to 2286. The NTP check fails open. Drop `fake-hwclock save` unless B8 shows the file is persisted. `date -s` and `/var/lib/muon3d/time` were adopted into 03 §3. |
| — | [#316](https://github.com/Muon-3D/MuonOS/pull/316) @ `a42c845` (KAN-403) | CI green | The Aux contract is re-recorded. Ready apart from the C3 run the PR itself promises. |

### Other repos

| Package | PR | State | Notes |
|---|---|---|---|
| ML-1 | [muon-link#25](https://github.com/Muon-3D/muon-link/pull/25) @ `b9f9678` | Matches the spec | Stacked on #24. |
| UI-1 | [MuonUI#48](https://github.com/Muon-3D/MuonUI/pull/48) @ `6f1f7de`, draft | **Changes needed** | Built on the first-version spec: wrong base (not #31), old fixtures and types, press-and-hold Back, and it raises the hotspot. Tests 222/222 pass, and typecheck is unchanged. |
| FL-1 | [Fluidd#17](https://github.com/Muon-3D/Muon3D_Fluidd/pull/17) @ `98eb768` | **Changes needed** | First-version fixtures and types, and the region line is unreachable in `screenFor`. The client rules are good. |
| FL-7 | [Fluidd#14](https://github.com/Muon-3D/Muon3D_Fluidd/pull/14) @ `d6d7970` | One addition | Matched the spec until KAN-404: the blacklist now needs `control.muon3d.com` as well as `app.muon3d.com`. |
| FL-8 | [Fluidd#15](https://github.com/Muon-3D/Muon3D_Fluidd/pull/15) @ `97e60f4` | One fix | Never put the hotspot key in a Fluidd QR code (07 S1). Ships with MuonOS#300. |
| FL-9 | [Fluidd#16](https://github.com/Muon-3D/Muon3D_Fluidd/pull/16) @ `edb5b29` | One fix | Over Iroh, non-2xx answers resolve as successes; fix `validateStatus`. The hardware check is still owed. |
| OR-1 | [OrcaSlicer#3](https://github.com/Muon-3D/OrcaSlicer/pull/3) | In review | Matches the spec. "Check profiles" is red on `main` too. |
| OR-2 | [OrcaSlicer#4](https://github.com/Muon-3D/OrcaSlicer/pull/4) | In review | Matches the spec. Linux and macOS pass. Windows fails on a deps cache miss, and Flatpak on a wxWidgets patch that no longer applies; neither is caused by this PR. |
| Spec | [OrcaSlicer#5](https://github.com/Muon-3D/OrcaSlicer/pull/5) | Taken into the spec | KAN-399 `app` driver kind and KAN-400 rule S12. |
| Spec | [OrcaSlicer#6](https://github.com/Muon-3D/OrcaSlicer/pull/6) | Taken into the spec | WS-13: `control.muon3d.com` (KAN-404), the link code (KAN-405), Iroh_BLE for phase 2 (KAN-402). |

## In-flight work that overlaps the spec

Build on these PRs rather than around them.

| PR | Overlaps |
|---|---|
| [Muon-3D/muon-link#24](https://github.com/Muon-3D/muon-link/pull/24): account link and orchestrator connection | ML-1, MR-5, and UI-3 P12 |
| [Muon-3D/MuonUI#31](https://github.com/Muon-3D/MuonUI/pull/31) (draft): setup route, region picker and pre-join region prompt | UI-1 builds on its branch; UI-2 P4 and P7a; KAN-324 |
| [Muon-3D/MuonUI#39](https://github.com/Muon-3D/MuonUI/pull/39): Aux Wi-Fi overhaul (KAN-339) | UI-2 |
| [Muon-3D/MuonUI#47](https://github.com/Muon-3D/MuonUI/pull/47) and [Muon-3D/Moonraker#21](https://github.com/Muon-3D/Moonraker/pull/21): SEC-8 protection levels | 07 §3, where setup writes join the protected set |
| [Muon-3D/MuonOS#174](https://github.com/Muon-3D/MuonOS/pull/174) (draft): region routes | OS-2; its `/setup` routes and `setup.toml` are replaced by OS-7 |
| [Muon-3D/MuonOS#300](https://github.com/Muon-3D/MuonOS/pull/300): hotspot `security_enabled` (KAN-376) | OS-5, FL-8 |
| [Muon-3D/MuonOS#305](https://github.com/Muon-3D/MuonOS/pull/305): GATE-1 listener declarations | OS-1 |

## Not started yet

No branches yet for MR-3, MR-5, MR-8, OS-2, OS-3, OS-4, OS-8, OS-9, OS-10, OS-11, ML-2, UI-0, UI-2 to UI-5, FL-2 to FL-6, CON-1, OR-3, QA-1 or QA-2, nor for the MuonOS pin bump. **OS-11 (Wi-Fi passwords in the journal) is urgent and still unassigned.**

## Spec changes since v1

- **15:10: the account link corrected against the real code.** It's muon-link PR #24 (`/link/*`), not `/pairing/*`, which is client pairing. Moonraker PR #20 is MR-6; build on it. Confirm and unlink are called by the panel directly and never through Moonraker. The orchestrator makes the code, and `remote.link` mirrors muon-link's `LinkPhase` as-is (`url`, integer `expires_at`). `self_hosted` is out of phase 1. New packages: ML-2 (proof of a knob press on confirm) and OS-10 (nginx `/muon-link/` on `:100`, and the orchestrator env vars).

- **15:40: MuonOS and MuonUI checked against the real code.** Region is confirmed **after** the join, on draft MuonOS#174's `/region` shapes and MuonUI#31's helpers (`state.03-region-apply` became `04b` and `04c`). Hotspot rules follow the real `muon-ap-lifecycle.sh`. The captive-portal nginx snippet passes `nginx -t`. A `status: "restored"` join counts as a failure. The internet check reuses the Nexigon check. The panel follows MuonUI's conventions. New packages: OS-11 and UI-0.

- **16:40: the marker contract follows MuonOS KAN-413:** `GET`, `POST` and `DELETE /setup/complete` over `/var/lib/muon3d/setup/complete`, not #174's `/setup`.

- **25 Sep 02:39: the Muon3D app (OrcaSlicer#5, KAN-399 and KAN-400).** `driver.kind` may be `app`, which only changes P8's title. New rule S12: an app that reads the hotspot key from the panel's QR code keeps it in memory only. Bench run R17.

- **25 Sep 04:29: the overnight reviews folded in.** Every PR author should read the parts that touch their package:
  - **Marker:** migration writes it with `by: migrated`; retries stop when `state` leaves `complete`, and `reset` cancels them; no fallback to #174's `/setup` (01 §7, 02 §4, 03 §7).
  - **`rev`:** `reset` keeps it increasing; a `GET` replaces the held state whatever its `rev`; a lapse is announced without a `rev` change (01 §5, 02 §5.11, §6).
  - **Region:** `region` carries every Aux field, including `explanation` and `enforcement`. `network.region_error` records a failed apply (new fixture `04d`) (02 §5.6a, §6).
  - **Update:** judge after a boot from Aux `/update/status`, never from update_manager's cache; Aux's 409 codes are mapped; the update check waits (02 §5.6, §5.8).
  - **Time:** zones from `zone.tab`; `clock` is hotspot-only; `tz` applies even when the clock is refused (02 §3, §5.2, §5.4).
  - **Ready:** the order, paused-print and macro-defined guards come before anything moves; item statuses and `op.item` are listed; a skipped `ready` stays on the card (02 §5.10, fixture 07).
  - **Identity:** `endpoint_id` and `dev_mode`; 503 only when Aux isn't answering (02 §7, 06 §1, 07 §4).
  - **Floor:** Moonraker#25 owns `/server/aux/setup`, `/server/aux/wifi/ap/auto_off` and `/server/aux/time`. The MuonOS pairing is `EXPECTED_FLOOR` in `test_trusted_clients.py`, landed with the pin bump (02 §1, 03).
  - **Websocket setup writes** check the upgrade request's `Host` (02 §3, 07 S6).
  - **D11** is stated as properties: the printer routes nothing, with AP client isolation, `ipv6.method=disabled` and bench B7 (03 §2, 07 S7).
  - **Panel:** the Vuetify light theme colours (the `--m3d-*` tokens resolve dark); any bundled QR encoder; exceptions for printing and the Klippy modal; P2 while the hotspot starts; P8 through a lapse (04).
  - **Phone page:** the join-result rule keyed on the `rev` at Connect; "finished" means addresses or an error; mount without `appInit` or `initCloud`; `getRandomValues` for the client ID (05 §2, §4, §5).
  - **Work plan:** FL-1, FL-8, FL-9, MR-3, MR-7, MR-9, OS-1 and OS-8 rows updated. PR titles may use `type(ID): …`.

- **25 Sep 04:31: hotspot and clock follow OS-5 and OS-6 as built.**
  - **Hotspot:** an uplink is `wlan0` or `eth0`. H1 clears the deadline. There's a 15-minute grace after each boot, which `ap-hotspot-disabled` skips. The deadline is boot-relative in `/run/muon3d/ap/auto-off`, with a path unit and a timer, and no sudo helper. The lifecycle re-checks its inputs at the end of each run. `GET /wifi/ap/stations` feeds `state.hotspot`. A join after `complete` re-arms the auto-off (03 §1, 02 §6, §5.11).
  - **Clock:** `date -u -s`, not `timedatectl set-time`. Times more than 20 years after the build are refused, and so is a request when the sync state can't be read. The zone lives in `/var/lib/muon3d/time`. Bench B8 checks whether `fake-hwclock` is persisted. The clock rule is KAN-270 prerequisite 2, not KAN-198 (03 §3, 01 §2.2, 07 S9).

- **25 Sep 13:35: the WS-13 decisions (OrcaSlicer#6, ADRs 0025–0027).**
  - **KAN-404:** the console host becomes `control.muon3d.com`. It has no DNS record until the KAN-414 cutover, so FL-2 defaults to `app.muon3d.com` until then, and FL-7 blacklists both hosts. The panel (P12) and phone page (S6) show the host of `remote.link.url` rather than a hard-coded name.
  - **KAN-405:** the console mints a 6-digit code valid for 600 s, limits failed claims and reissues a code after a decline (S4, 03 §6). LINK-3 needs the claiming client's own key, which the console doesn't send yet (KAN-415). Until it does, the authority fingerprint shown at `offer` doesn't satisfy LINK-3 (07 S3, 04 P12).
  - **KAN-402:** phase 2 Bluetooth uses Iroh_BLE inside muon-link instead of a separate GATT daemon. A BLE setup peer needs a new `bluetooth` caller class, reported by muon-link (BT-2) (03 §8, 10, 09).

**Corrections to earlier notes here:**
- MR-1's `/wifi/ap/stations` poll made one 404 per start, not one every 2 s.
- #313 was never going to carry the nginx 403 entry; the pin bump does.
