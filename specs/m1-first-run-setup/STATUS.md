# Status

This page tracks the work packages in [10-work-plan.md](10-work-plan.md). The coordinating session updates it at each check-in.

Agents: the coordinator can't receive messages from you. Tell it where you are by pushing branches and opening PRs whose titles start with the package ID. Put questions in the PR body under **Spec questions**. **Link the spec commit you built against** (10 §1 rule 7), and read the log at the end of this page and [fixtures/README.md](fixtures/README.md) for anything newer.

**Last updated:** 26 Sep 2026, 12:59 UTC. Most of the first wave merged overnight:

- **Moonraker `master` (`935c72c`):** MR-1, MR-2, MR-4, MR-7, MR-9 and the floor PR.
- **Fluidd `develop` (`d16bca0`):** FL-1, FL-7, FL-8 and FL-9.
- **MuonOS `main`:** OS-5, OS-6 and OS-7, plus the pin bumps for Moonraker (with the floor pairing) and Fluidd.
- **muon-link:** #24 and ML-1.
- **MuonUI:** UI-1, into #31's branch.

An audit of the merged code found gaps. Issues are turned off in these repos, so the lists are comments on the merged PRs: [Moonraker](https://github.com/Muon-3D/Moonraker/pull/22#issuecomment-5846417945), [Fluidd](https://github.com/Muon-3D/Muon3D_Fluidd/pull/17#issuecomment-5846418769) and MuonOS (on #314). The spec took the fixes at `aa0fe2f` and in today's later commit.

## Decisions needed from the owner

1. **SEC-8 Level 1 isn't enforced in setup yet (security).** On Moonraker `master`, after `finish`, a LAN or hotspot browser can still rename a Protected printer and skip `ready`. It's item 1 of the Moonraker follow-ups. It needs an agent, and it should go in before the next Moonraker pin bump.
2. **Don't promote an image with OS-5 until it can finish setup (new).** Since #314 and #320 merged, a unit that `muon_setup` stores as `new` keeps its hotspot up, and the owner's "off" is ignored, until setup finishes. But no image can finish it yet: UI-1 isn't pinned and FL-3 isn't built. That hits fresh and factory-reset dev units on `latest` now (workaround: `POST /server/muon/setup/language`, then `finish`, from the LAN). It would also hit any field unit that migration misses, such as an Ethernet-only unit driven from OrcaSlicer and never linked. The spec now adds job history and uploaded G-code as migration signals (01 §7), and OS-9 must carry UI-1 or FL-3 before promotion.
3. **C3 runs are owed on merged PRs.** #308, #313, #314, #315 and #320 merged on a waiver from Jack that is only relayed in the PR bodies. The runs must happen before beta or stable, and Jack should record the waiver himself. #312 is waiting on the same thing.
4. **KAN-341's AP-4.** OS-5 (MuonOS#314) merged, so a set-up printer's hotspot now turns off once the printer is on a network. That reverses KAN-341's "on unless disabled". If Jack hasn't agreed, record his answer on KAN-341.
5. **The region and Wi-Fi chain is waiting on device evidence.** The panel's network screens (UI-2) come after MuonUI#39 and #31, and #39 needs MuonOS#210. OS-2 builds on #174. Each of #174 and #210 waits on a C3 run that only a person can do, and #174 also needs the human KAN-351 check. Both now conflict with `main`.
6. **MuonUI#31's C3 gate.** The old record covered a `SetupView` that has since been replaced. Is an unkeyed run of the region prompt and the REGION card enough to merge it?
7. **OS-11 is still unassigned.** Wi-Fi passwords are written to the persisted journal. It's urgent and independent of setup.
8. **FL-9's hardware check** needs someone with a linked unit (Fluidd follow-up item 1).
9. **Second wave.** MR-3, MR-8, OS-10, ML-2, UI-4, FL-4 and FL-5 can start now; see "Ready to start" below. Do you want handoff prompts for them?

Closed since the last update:
- **Remote Wi-Fi writes.** muon-link's `ABSOLUTE_DENY` refuses `/server/aux/` to every remote session, so a remote caller can't change the printer's Wi-Fi (07 §3).
- **Flooring `/server/aux/region/country`.** Not floored. Like the Wi-Fi routes, it's open to the LAN at Level 0 (SEC-1 trusts the LAN) and protected at Level 1. The token limits it to assessed configurations, so a LAN caller can't choose an illegal one.

## Merge order from here

- **Moonraker:** the SEC-8 follow-up first, then the other follow-ups in any order. MR-6 #20 goes in once it's fixed; MR-5 needs it. MR-3 is the next package on the critical path.
- **MuonOS pin bumps:** the next Moonraker bump carries the SEC-8 follow-up and #20. If #20 adds `/server/muon/link/start` to `PROTECTED_PREFIXES` and MuonOS mirrors that list, update the mirror in the same bump. The next Fluidd bump carries #19 and the follow-ups.
- **MuonOS:** #312 (OS-1) was waiting for a Fluidd that serves `/setup`; #323 pinned one. #321 (OS-2) and #322 (OS-3) sit on #174 and #210.
- **MuonUI:** #39 → #31 → UI-2. UI-4 can be built on #31's branch now.

## Packages

### Moonraker (`Muon-3D/Moonraker`)

| Package | PR | State | Notes |
|---|---|---|---|
| MR-1 | [#22](https://github.com/Muon-3D/Moonraker/pull/22) | **Merged** 26 Sep | Follow-ups 1 (SEC-8), 3, 5, 6 and 7 are in the [list](https://github.com/Muon-3D/Moonraker/pull/22#issuecomment-5846417945). |
| MR-9 | [#23](https://github.com/Muon-3D/Moonraker/pull/23) | **Merged** into #22 | Follow-up 2: an Aux timeout answers 500 rather than 503. |
| MR-2 | [#24](https://github.com/Muon-3D/Moonraker/pull/24) | **Merged** 26 Sep | |
| MR-4 | [#26](https://github.com/Muon-3D/Moonraker/pull/26) | **Merged** 26 Sep | The blocking bug is fixed: it now judges an update from Aux. |
| MR-7 | [#28](https://github.com/Muon-3D/Moonraker/pull/28) | **Merged** 26 Sep | Paused prints and manifest order are fixed. Follow-up 4: an unreadable `print_stats` still lets the self-test run. |
| — | [#25](https://github.com/Muon-3D/Moonraker/pull/25), [#27](https://github.com/Muon-3D/Moonraker/pull/27) | **Merged** | The floor entries, and `endpoint_id` in identity. |
| MR-6 | [#20](https://github.com/Muon-3D/Moonraker/pull/20) | **Changes needed** ([review](https://github.com/Muon-3D/Moonraker/pull/20#issuecomment-5846402897)) | Three blockers: `link/start` isn't in `PROTECTED_PREFIXES`; a 409 on `start` should return the current phase; `start` and `cancel` need the Host/Origin checks. Also the 429 limit, polling, and public methods. |

### MuonOS (`Muon-3D/MuonOS`)

| Package | PR | State | Notes |
|---|---|---|---|
| OS-1 | [#312](https://github.com/Muon-3D/MuonOS/pull/312) @ `d7e48c5` | **Code done**; waiting on C3 | Every review item is fixed: an `inet` forward chain with policy drop that loads before NetworkManager, the new `10-ap-isolate` match, `10.42.0.1/24`, `ipv6.method=disabled`, `ap-isolation=1`, and the migration of persisted profiles. The Fluidd gate is met by #323. Tests and a real-nginx run pass. PR policy is red only for C3 (decision 3). Three small non-blocking notes. |
| OS-2 | [#321](https://github.com/Muon-3D/MuonOS/pull/321) @ `f83879b`, on #174 | Two fixes | All four items are there. Hold the guard until the agent exits, not until the timeout; parse the reason from more than the last stderr line. CI is red only for C3. |
| OS-3 | [#322](https://github.com/Muon-3D/MuonOS/pull/322) @ `9eee473`, on #210 | **Changes needed** | gitleaks flags two fake test values, so CI never ran the tests. **The join reason is read after NetworkManager resets it to 0**, so a wrong password would come back as `timeout`. Map it from nmcli's own failure line instead, and check on a unit. Also: retry a hidden join once, bound the uplink check to 5 s, and read the hub from the agent config. |
| OS-5 | [#314](https://github.com/Muon-3D/MuonOS/pull/314) | **Merged** 26 Sep 10:55 | Follow-ups: the end-of-run re-check snapshots after deciding, so a change made while it decides is lost (reproduced). See also decisions 2 and 4. |
| OS-6 | [#315](https://github.com/Muon-3D/MuonOS/pull/315) | **Merged** 25 Sep 23:59 | |
| OS-7 | [#313](https://github.com/Muon-3D/MuonOS/pull/313) | **Merged** 26 Sep 01:53 | |
| Pins | [#308](https://github.com/Muon-3D/MuonOS/pull/308), [#320](https://github.com/Muon-3D/MuonOS/pull/320), [#323](https://github.com/Muon-3D/MuonOS/pull/323) | **Merged** | SEC-8; Moonraker `935c72c` with the setup floor in `EXPECTED_FLOOR`; Fluidd with #14–#17. |
| — | [#316](https://github.com/Muon-3D/MuonOS/pull/316) (KAN-403) | Conflicts with `main` | Aux's half of `endpoint_id`. Merge `main` in; its own C3 run is still owed. Moonraker's half (#27) is already pinned. |

### Other repos

| Package | PR | State | Notes |
|---|---|---|---|
| ML-1 | [muon-link#25](https://github.com/Muon-3D/muon-link/pull/25) | **Merged** 26 Sep | muon-link#24 merged too. |
| UI-1 | [MuonUI#48](https://github.com/Muon-3D/MuonUI/pull/48) | **Merged** into #31's branch (`f9defc4`) | Its follow-ups are listed on #31. |
| — | [MuonUI#31](https://github.com/Muon-3D/MuonUI/pull/31), draft | Not ready ([status](https://github.com/Muon-3D/MuonUI/pull/31#issuecomment-5846404048)) | Needs #39 first, a merge of `main` (the `SettingsView` conflict), a new description, the C3 decision (decision 4), and the UI-1 follow-ups. |
| FL-1 | [Fluidd#17](https://github.com/Muon-3D/Muon3D_Fluidd/pull/17) | **Merged** 26 Sep | Follow-ups 2, 4 and 6 are in the [list](https://github.com/Muon-3D/Muon3D_Fluidd/pull/17#issuecomment-5846418769). |
| FL-7 | [Fluidd#14](https://github.com/Muon-3D/Muon3D_Fluidd/pull/14) | **Merged** 26 Sep | |
| FL-8 | [Fluidd#15](https://github.com/Muon-3D/Muon3D_Fluidd/pull/15) | **Merged** 26 Sep | Follow-up 3: the QR code's escaping. |
| FL-9 | [Fluidd#16](https://github.com/Muon-3D/Muon3D_Fluidd/pull/16) | **Merged** 26 Sep | Follow-ups 1 (the hardware check) and 5. |
| — | [Fluidd#19](https://github.com/Muon-3D/Muon3D_Fluidd/pull/19) | **One change** ([review](https://github.com/Muon-3D/Muon3D_Fluidd/pull/19#issuecomment-5846403307)) | Opens a found printer at an address that answers. Before counting an address, check that it answers as the same printer (06 §2.3 item 6). |
| — | [Fluidd#18](https://github.com/Muon-3D/Muon3D_Fluidd/pull/18) (KAN-401) | Open, no conflict with the spec | The `/link?code=` path form. Now that #14 has merged, it can add `control.muon3d.com` if #14 didn't. |
| OR-1, OR-2 | [OrcaSlicer#3](https://github.com/Muon-3D/OrcaSlicer/pull/3), [#4](https://github.com/Muon-3D/OrcaSlicer/pull/4) | **Merged** 25 Sep | OR-2's status column fills in once MR-8 publishes the TXT keys. |

## Ready to start

Every dependency listed in 10 for these packages is met:

| Package | Why now |
|---|---|
| **MR-3** (network and join) | MR-1 is merged. OS-2 and OS-3 are still open, so use fakes for their codes (10). Take the join reason from Aux's `code`, never `state_reason` (03 §4). It's on the critical path: UI-2 and FL-3 wait for it. |
| **MR-8** (Zeroconf TXT) | MR-1 is merged. |
| **OS-10** (`/muon-link/` on `:100`) | ML-1 is merged. |
| **OS-11** (PSKs in the journal) | No dependencies. Urgent. |
| **ML-2** (knob proof on confirm) | ML-1 is merged. |
| **UI-4** (P13, P14, the Finish card) | UI-1 and MR-7 are merged. Build it on #31's branch. |
| **FL-4** (Add printer) | FL-1 and MR-9 are merged. Take #19 first. |
| **FL-5** (the Dashboard banner) | FL-1 is merged. |
| **FL-3** (the phone screens) | FL-1 is merged. It can build against the fixtures now and finish once MR-3 lands. |

Still blocked: MR-5 (on #20), UI-2 (on MR-3, #39 and #31), UI-3 (on MR-5 and OS-10), FL-2 (on CON-1), OS-4 (on OS-3), OS-8 (on the hardware team), OS-9 (on the UI and FL packages), and QA-1 (on OS-1 and a dev build with MR-3).

## In-flight work that overlaps the spec

| PR | Overlaps |
|---|---|
| [Muon-3D/MuonUI#39](https://github.com/Muon-3D/MuonUI/pull/39): Aux Wi-Fi overhaul (KAN-339) | UI-2. It needs MuonOS#210. |
| [Muon-3D/MuonOS#210](https://github.com/Muon-3D/MuonOS/pull/210): `/wifi/saved` and correcting a saved password | OS-3 and MR-1's migration signal. Conflicts with `main`; C3 owed. |
| [Muon-3D/MuonOS#174](https://github.com/Muon-3D/MuonOS/pull/174) (draft): region routes | OS-2 and MR-3's region step. Its `/setup` routes and `setup.toml` are replaced by OS-7. Needs a keyed C3 run and the human KAN-351 check. |
| [Muon-3D/MuonOS#300](https://github.com/Muon-3D/MuonOS/pull/300): hotspot `security_enabled` (KAN-376) | OS-5 and FL-8. Its C3 run passed, but it conflicts with `main`. FL-8 was meant to ship with it and is already pinned (#323), so merge `main` into #300 and land it next. |
| [Muon-3D/MuonOS#305](https://github.com/Muon-3D/MuonOS/pull/305): GATE-1 listener declarations | OS-1. |
| [Muon-3D/muon-link#27](https://github.com/Muon-3D/muon-link/pull/27) (sealed alerts), [#26](https://github.com/Muon-3D/muon-link/pull/26) (mobile binding) | Neither touches `/link/*`, as far as the titles show. Not checked against the spec. |

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

- **25 Sep 13:31: the WS-13 decisions (OrcaSlicer#6, ADRs 0025–0027).**
  - **KAN-404:** the console host becomes `control.muon3d.com`. It has no DNS record until the KAN-414 cutover, so FL-2 defaults to `app.muon3d.com` until then, and FL-7 blacklists both hosts. The panel (P12) and phone page (S6) show the host of `remote.link.url` rather than a hard-coded name.
  - **KAN-405:** the console mints a 6-digit code valid for 600 s, limits failed claims and reissues a code after a decline (S4, 03 §6). LINK-3 needs the claiming client's own key, which the console doesn't send yet (KAN-415). Until it does, the authority fingerprint shown at `offer` doesn't satisfy LINK-3 (07 S3, 04 P12).
  - **KAN-402:** phase 2 Bluetooth uses Iroh_BLE inside muon-link instead of a separate GATT daemon. A BLE setup peer needs a new `bluetooth` caller class, reported by muon-link (BT-2) (03 §8, 10, 09).

- **25 Sep 16:39: SEC-8 Level 1 and setup (Moonraker#21 merged).** When the printer is Protected and setup is `complete`, `muon_setup` refuses step writes from `lan` and `hotspot` callers with 403 `protected`. Before `complete`, setup stays open, and reads stay open. `muon_setup` checks this in its own caller check rather than through `PROTECTED_PREFIXES`. `/server/muon/link/start` does join `PROTECTED_PREFIXES` (MR-6). The phone page's S10 shows "Change Wi-Fi on its screen" (02 §3, 05 S10, 07 §3, 08, 10 MR-1 and MR-6).

- **25 Sep 17:44: the link confirm contract (muon-link#24).** `POST /link/confirm` carries `{account, fingerprint}` from the offer the panel showed: 400 without them, 409 if a different offer waits. `start` is refused while an offer waits (03 §6, 04 P12, 08 `link_offer_changed`).

- **26 Sep 12:48 (`aa0fe2f`): the post-merge audits of Moonraker, Fluidd and MuonUI.**
  - **SEC-8 Level 1** covers every post-setup write except `driver` and `card/dismiss`, and `/server/muon/identity/name` joins `PROTECTED_PREFIXES` (02 §3).
  - **Identity 503:** a 500 with no HTTP response behind it counts as "not answering" (02 §7).
  - **`muon_link`:** it turns a 409 on `start` into the current phase, applies the Host/Origin checks, and answers 429 past the limit (02 §9).
  - **Setup behaviour:** an unreadable `print_stats` refuses the self-test. Only Aux blocks migration. The panel's claim never lapses (02 §5.10, §6, 01 §7).
  - **Phone page:** a lost write is judged against the state at `sentRev` (05 §4).
  - **Panel:** no Back on P8, P7a chosen from the state, and redirects suspended while printing (04).
  - **Errors:** a failed join carries `error.detail.ssid` (fixture 05).
  - **Remote Aux writes** are closed by `ABSOLUTE_DENY`.

- **26 Sep 12:59: the MuonOS audit and the OS-2 and OS-3 reviews.**
  - **Migration** also counts job history and uploaded G-code, which catches Ethernet-only OrcaSlicer units (01 §7, 02 §8).
  - **Hotspot:** H1 needs a setup surface in the same image before promotion. The lifecycle snapshots its inputs *before* deciding (03 §1, 10 OS-5).
  - **Floor:** a floored route never reaches `main` before its floor (03 intro). `setup.toml` is the persist file (03 §7).
  - **D11:** the INPUT half should meet the same properties. tcp/22, udp/5353 and udp/7127 stay open on `10.42.0.1` (03 §2, 07).
  - **Region:** `configurations` is added to `/region/options` (fixtures `options.*`). Three new build-fault codes map to `region_apply_failed`. The guard is held until the agent exits. A 504 is checked against `GET /region` before its code is read (03 §3, 02 §5.6a, 08).
  - **Wi-Fi:** the join reason comes from Aux's `code`, which Aux maps from nmcli's failure line; `state_reason` has been reset to 0 by then. Hidden joins use `hidden yes` with one retry. The internet check is a TLS handshake to the configured Nexigon hub, bounded to 5 s. `uplink_unavailable` reads as `internet: null` (03 §4, 02 §5.6, 07 S8, 10 MR-3 and OS-3).

**Corrections to earlier notes here:**
- MR-1's `/wifi/ap/stations` poll made one 404 per start, not one every 2 s.
- #313 was never going to carry the nginx 403 entry; the pin bump does.
