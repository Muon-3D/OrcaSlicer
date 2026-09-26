# 10 · Work plan

These work packages are sized for one agent or one PR each. Every package names its repo, what it depends on, and when it's done. The shared contract is [02-setup-api.md](02-setup-api.md), and every package can be built against the fixtures in [`fixtures/`](fixtures/) before its dependencies land.

## 1. Rules for implementing agents

1. **Read first.** Read [README.md](README.md), then the files for your repo. The contract in 02, and the error codes in 08, are authoritative across repos.
2. **Match the repo.** Follow its conventions (lint, tests, style). Each file lists them for its repo.
3. **Don't invent external payloads.** Two interfaces are known by name only:
   - muon-link's `/link/*` (PR #24; `/pairing/*` is client pairing and is not used here);
   - the exact shapes of Aux's `/region/*`.
   Read their source. If a field you need isn't there, add it in that repo (it's listed as a package here) or stop and ask. Never guess a shape and hard-code it.
4. **No secrets anywhere.** Passwords and the hotspot key never appear in logs, state, events, fixtures or test snapshots.
5. **Never factory-reset a real unit** (KAN-351). Use `POST /server/muon/setup/reset` on a development unit, which resets setup state only.
6. **Keep the spec in step.** If an implementation needs the contract to change, change this spec in the same PR series and say so in the PR description. The spec lives at `specs/m1-first-run-setup/` in `Muon-3D/OrcaSlicer` until the team moves it.
7. **One PR per package.** Title it `<ID>: <summary> (KAN-xxx)`, or `type(<ID>): <summary> (KAN-xxx)` where the repo checks semantic PR titles, and link the spec section in the description **at the spec commit you built against**. When the spec moves on, check `STATUS.md`'s "Spec changes" log and update the PR.

## 2. Packages

### Moonraker fork (`Muon-3D/Moonraker`)

| ID | Package | Depends on | Done when |
|---|---|---|---|
| MR-1 | `muon_setup` core: the `[muon_setup]` config and template pin, the state document, persistence (namespace `muon_setup`), `rev`/`driver`/`op` handling, caller classification and write hygiene (§3), migration (01 §7), `GET` state and options, `goto`, `skip`, `finish`, `card/dismiss`, `reset`, `notify_muon_setup_changed`, the `muon_setup:complete` event, the floor entry for `reset`, and the SEC-8 Level 1 check on post-setup writes (02 §3). `finish` calls OS-5 `auto_off` and Aux `POST /setup/complete` (OS-7) through `aux_api_proxy`, with a fake in tests. The floor entry for `reset` also goes into MuonOS `fluidd.nginx.template`. | – | 02 §8 tests 1–4, 7, 8 and 13 pass; flake8 and mypy are clean |
| MR-2 | Language (including Fluidd's default locale), clock, time zone, name (refactor `aux_api_proxy` to expose `set_friendly_name` and `get_identity`), and the tzdata lookup for options | MR-1; OS-6 for real clock and time-zone writes | Unit tests for each; the identity endpoint behaves as before |
| MR-3 | Networks and join: normalising the scan with `channel_permitted`, the join orchestration (an optional region switch → join → phases; `status: "restored"` and 400 count as failures; the reason is Aux's join `code`, never `state_reason`), cancel, CA upload, Ethernet, the internet result, update visibility, and **§5.6a region confirmation**. Adds a timeout parameter to `aux_api_proxy.post()`, and makes its error mapping keep Aux's `detail.code`. | MR-1; OS-2 (region codes), OS-3 (W5 uplink). Uses fakes until then. | 02 §8 tests 5, 6, 9 and 10 pass; `/server/muon/setup/network` is covered by Moonraker's verbose-log redaction (07 S2) |
| MR-4 | The update step through `update_manager`, including the version check after reboot | MR-1 | Test 11 passes |
| MR-5 | The remote step: `local`, `cloud` (mirroring `LinkPhase`, renewing past `expires_at`, never during `offer`), and `later`. `self_hosted` isn't offered in phase 1. | MR-1, MR-6 | Test 12 passes |
| MR-6 | `muon_link`: build on [Muon-3D/Moonraker#20](https://github.com/Muon-3D/Moonraker/pull/20), which already has `GET /server/muon/link`, `POST …/start` and `POST …/cancel`. Add: polling `GET /link` with the `muon_link:link_changed` event, public `status()`/`start()`/`cancel()`, a limit of 5 `start` calls per minute per IP (429), a 409 on `start` returned as the current phase, §3's Host/Origin/content-type checks on `start` and `cancel`, and `/server/muon/link/start` in SEC-8's `PROTECTED_PREFIXES`. MuonOS has no test that mirrors that list (26 Sep), so nothing changes there. **Never add confirm or unlink.** | PR #20 merged (or stacked on it); ML-1 | `tests/test_muon_link.py` extended and passing; Fluidd's `startLanLink()` works against a dev unit |
| MR-7 | The ready step: the manifest (loaded from a file, with a default and a schema check), confirm, macro start through `klippy_apis`, hiding undefined macros, and the busy, order and not-ready guards (02 §5.10) | MR-1 | Unit tests pass, including a paused print, the clips-before-self-test order, and a "Finish setup" card that stops halfway |
| MR-8 | Zeroconf: the `_octoprint._tcp` record advertises **port 80** (nginx), and TXT gains `name=<display>` and `setup=<state>`; the record is re-registered when either changes. Coordinate with KAN-364's NET-3 record so it also carries `setup`. | MR-1 | OrcaSlicer's Bonjour dialog lists a dev unit, and `dns-sd -L` shows port 80 and the TXT keys |
| MR-9 | `aux_api_proxy`: register the identity and `dev_mode` endpoints in `__init__` (503 while Aux isn't answering), add `setup` to `/server/muon/identity`, and fix `tests/test_aux_api_proxy.py`, whose `FakeServer` has no `database` (42 of 44 tests fail today) | MR-1 | That test file is green |

### MuonOS (`Muon-3D/MuonOS`: Aux API, image, services)

| ID | Package | Depends on | Done when |
|---|---|---|---|
| OS-1 | Captive portal and hotspot isolation (D11): the printer routes nothing, with the properties in 03 §2 item 1 (IPv4 and IPv6, interface-based, loaded before NM, backend-independent, never blocking `10.42.0.1`); `tcp/443` reject with tcp reset on `ap0`; pin `10.42.0.1`, `ipv6.method=disabled` and `wifi.ap-isolation=1` in `ap0-con`, migrating persisted profiles in `muon-ap-provision.sh`; the dnsmasq drop-in; the nginx `map`/redirect and `location = /setup` (must pass `check-nginx-templates.sh`); update `test_firewall_ruleset.py`, the `10-ap-isolate` tests and AP-7. Don't merge before the pinned Fluidd serves `/setup`. | #305 (listeners) | B5 and B7 pass; iOS and Android open `/setup` by themselves; from a phone on the hotspot, a LAN host and a public IP are unreachable with and without Ethernet, including after `ap0-con` goes down and up |
| OS-2 | Region, on draft #174: stable `detail.code` values from the agent's `OUTCOMES`; `SET_COUNTRY_TIMEOUT_S` under 60 s; a concurrency guard returning `busy`, held until the agent exits; `configurations` in `/region/options`. [MuonOS#321](https://github.com/Muon-3D/MuonOS/pull/321), on #174's branch. **Blocked for real devices until signing keys and tokens exist** (KAN-321, KAN-132). | #174 merged | Aux tests pass; R2 and R9 pass on the bench |
| OS-3 | Wi-Fi: W1 (replacing a saved secret is #210; `hidden yes` under the existing grant, with one retry on exit 10), W2 (stable error codes from nmcli's own failure line, not `state_reason`), W5 (`/wifi/uplink`, with a bounded TLS handshake to the configured Nexigon hub), W6 (`/wifi/saved` is #210). [MuonOS#322](https://github.com/Muon-3D/MuonOS/pull/322). | #210 | Aux tests pass |
| OS-4 | Enterprise: W3 (an `eap` connect through a new privilege path, D-Bus/polkit or `nmcli connection add`) and W4 (`/wifi/ca_cert`, plus a persist declaration and inventory rows for `/etc/NetworkManager/certs`). The largest OS item; it can ship after phase 1 without blocking anything. | OS-3 | R6 passes against N6 |
| OS-5 | Hotspot rules H1–H4 in `muon-ap-lifecycle.sh` (03 §1): an uplink is `wlan0` or `eth0`, the `ap-hotspot-kept-on` marker, H1 clearing the deadline, the boot grace that `ap-hotspot-disabled` skips, and the end-of-run re-check; `POST /wifi/ap/auto_off` writing a boot-relative deadline to `/run/muon3d/ap/auto-off`, with a path unit and a transient timer; `GET /wifi/ap/stations`. Floor entries: Moonraker#25, then the MuonOS pin bump. **Don't merge until the pinned Moonraker writes the marker on migration** (02 §4), or every field unit's hotspot is forced on. *Merged 26 Sep.* Follow-ups: the end-of-run re-check must snapshot before deciding (03 §1), and **don't promote an image with OS-5 until it also carries UI-1 or FL-3** (03 §1 "H1 needs a way to finish setup"). | OS-7 (H1 reads the marker); MR-1 migration marker; a setup surface in the image before promotion | H1–H4 verified on a dev unit, including Ethernet only; R8 passes |
| OS-6 | Time: `GET/POST /time` and `POST /time/zone` with pinned sudoers entries; the time zone kept in `/var/lib/muon3d/time/timezone` (`muon3d-time.toml`), re-applied by `muon3d-timezone.service`; a data-inventory row. `/server/aux/time` floored (02 §1). | – | Aux tests pass; the time zone survives a reboot |
| OS-7 | The setup marker: **MuonOS `feat/KAN-413-setup-marker`**. `GET`, `POST` and `DELETE /setup/complete` over `/var/lib/muon3d/setup/complete`; `setup.toml` in `rugix-ctrl-config`; rows in ID-9 and the data inventory. The floor pairing lands with the Moonraker pin bump past Moonraker#25 (02 §1): `EXPECTED_FLOOR` in `test_trusted_clients.py`, the 403 location in `fluidd.nginx.template`, and `FLOOR_CASES` in `test_floor.py`. Accepts `by: migrated`. #174 drops its own `setup_routes.py`/`setup.toml`, or rebases onto this. | – | `test_setup_marker.py`, `test_persisted_paths.py` and the ID-9 test pass; the marker survives an OTA update |
| OS-8 | Ship `/usr/share/muon/setup/ready.json`, with the hardware team, and the `MUON_SELF_TEST` macro if they want one | Hardware team decision | `ready.json` validates in MuonOS CI with Moonraker's manifest validator (an invalid file silently falls back to the default) |
| OS-9 | Pin the new MuonUI and Fluidd builds in the image, updating the Fluidd zip checksum guard (it blocked the KAN-321 C3 run) | The UI and FL packages | The image builds and boots to setup on a clean flash |
| OS-10 | muon-link wiring for the account link: an nginx `location /muon-link/` on the **`:100` vhost only**, `MUON_LINK_ORCH_ID` and `MUON_LINK_RELAY_URL` in the muon-link unit, and **never** `MUON_LINK_DISCOVERABLE=1` (03 §6) | ML-1 | The panel can read `/muon-link/link`; `:80` can't reach it |
| OS-11 | **Security, urgent and independent of setup:** stop Wi-Fi PSKs reaching the persisted journal. Pass secrets through nmcli `passwd-file` (0600, tmpfs) or D-Bus instead of argv; add sudoers `!syslog` for the connect command; vacuum existing entries on update (07 §3). | – | `journalctl` shows no PSK after a connect; a test covers the argv |

### muon-link

| ID | Package | Depends on | Done when |
|---|---|---|---|
| ML-1 | Land [Muon-3D/muon-link#24](https://github.com/Muon-3D/muon-link/pull/24) (the account link, `/link/*`), keeping the `LinkPhase` shapes in 02 §5.9 | – | PR merged; contract test against `GET /link` |
| ML-2 | **Security:** proof of a knob press on `POST /link/confirm` (and `/link/unlink`). Today any loopback process can confirm. Use the same mechanism as Aux's `KnobConfirmationBackend` (`dev_mode_consent.py`), or a one-time nonce that MuonUI gets from sk_daemon on a press. | ML-1 | A loopback `curl` can't confirm without a press |

### MuonUI (`Muon-3D/MuonUI`)

| ID | Package | Depends on | Done when |
|---|---|---|---|
| UI-0 | Add i18n to MuonUI: `vue-i18n` and `src/locales/{en,de,fr,es,it}.json`, all loaded at start | – | Locale switching works; the P1 title renders in each language |
| UI-1 | The setup shell on #31's branch: `setupStore`, the `useMoonraker` helpers, the router check, `screenFor`, `ArcProgress`, `setupQr` (the `qrcode` dependency), the mock backend routes; P1 Language, P2 Here or phone, P8 Following | UI-0; MR-1 (fixtures until then); #31 author agreement | 04 §6 tests for these screens pass; boots into P1 on a clean unit |
| UI-2 | Network: P4 list (with `regionPromptFor`), P5 password (a new `KeyboardSet` with a "use phone" key; Show and last-character reveal), P6 Joining, P7 Connected and errors, P7a Region (`RegionPicker`), P7b Time zone | UI-1, MR-3, **after #39 and #31** | Tests pass; R1 and R2 pass on the bench |
| UI-3 | P9 Name, P10 Update, P11 Remote, P12 Link code and the confirm through muon-link directly (NET-10(d)) | UI-1, MR-4, MR-5, OS-10 | R13 passes |
| UI-4 | P13 Ready, P14 Done, the home "Finish setup" card, and the Settings entries "Add a phone or computer" (the AP-3 QR) and "Link to account" | UI-1, MR-7 | Tests pass |
| UI-5 | Setup strings translated into de, fr, es and it | UI-1–UI-4 | No missing keys (a CI check) |

### Fluidd fork (`Muon-3D/Muon3D_Fluidd`)

| ID | Package | Depends on | Done when |
|---|---|---|---|
| FL-1 | The `/setup` shell: the route (lazy, chrome-less, `printerIndependent`), `main.ts` mounting at once on `/setup` without `appInit` or `initCloud`, `client.ts`, `state.ts`, `screen.ts`, the `notifyMuonSetupChanged` socket action, the page-language rule and header switcher (05 §9), and the bundle budget check | MR-1 (fixtures until then) | 05 §11 client and screen tests pass, including `screenFor` over fixtures 04b, 04c and 04d; `/setup` ≤ 900 KB gzipped with the shared vendor chunk |
| FL-2 | Cloud base URL: read `VUE_APP_MUON_CLOUD_URL` (matching `envPrefix: 'VUE_'`, declared in `env.d.ts`), default to the console host on any other origin (`https://control.muon3d.com` after the KAN-414 cutover; `https://app.muon3d.com` until then), and call `setCloudBaseUrl` on init | CON-1 | A printer-served Fluidd can sign in and claim a code |
| FL-3 | The `/setup` screens S0–S10, the region line, the error mapping, and the `en` keys | FL-1, MR-3 | `Setup.spec.ts` and E2E scenarios 1–8 pass |
| FL-4 | `AddPrinterDialog`: the entry points, the `LinkClaim` refactor, discovery merging the mDNS feed with the sweep, and the `setup` badges | FL-1, MR-9 | 06 §2.6 tests pass; R14 and R15 pass |
| FL-5 | The Dashboard "Finish setup" banner | FL-1 | A unit test over the state variants |
| FL-6 | i18n: de, fr, es and it for `app.muon.setup.*` and `app.muon.add_printer.*`; convert the Muon screens touched here to `$t`; add `app.general.confirm.enter_password` and `app.general.btn.connect` | FL-3, FL-4 | `npm run i18n-extract` reports no missing keys for these namespaces |
| FL-7 | Add the console hosts, `app.muon3d.com` and `control.muon3d.com`, to the host blacklist in `public/config.json` and `server/config.json`, so the console doesn't probe itself for 5 s at startup | – | Loading the console shows no 5 s wait |
| FL-8 | Make `useHotspotCheck` a pure origin read with no lifecycle hook (`isHotspotOrigin()`: only `10.42.0.1`), and stop the hotspot card's QR from saying `T:nopass` when the key is redacted. Draw the join QR only for an open hotspot (`security_enabled === false`); never put a key in a Fluidd QR (07 S1). Ships with MuonOS#300 (KAN-376). | – | `HotspotManagerCard.spec.ts` extended; passes |
| FL-9 | Rebind `auxAxios`'s adapter and `Authorization` on each request from `Vue.$httpClient`. (Routing it through `$httpClient` doesn't work: the generated client drops `basePath` when the instance has a `baseURL`.) muon-link refuses `/server/aux/` to every remote session, so over Iroh the Aux cards say "Wi-Fi can only be changed on the printer's own network" and send nothing. Stop `AddInstanceDialog` verification going over Iroh. Make the Iroh adapter reject non-2xx responses, as axios's `settle()` does. | – | On a linked unit opened from the console: the Wi-Fi and hotspot cards show that line and the network log has no `/server/aux/` request; the same cards load on the printer's own network; a 400 over Iroh rejects |

### Console (`muon-console`, not reviewed for this spec)

| ID | Package | Done when |
|---|---|---|
| CON-1 | `/v1/*` answers CORS preflight for any origin with bearer auth and no credentials mode, so printer-served Fluidd pages can sign in and claim codes | FL-2 works from `http://<printer>` |

### OrcaSlicer (`Muon-3D/OrcaSlicer`)

| ID | Package | Depends on | Done when |
|---|---|---|---|
| OR-1 | Remove the placeholder `print_host` and bump the Muon3D profile version | – | 06 OR-1 check |
| OR-2 | Bonjour dialog: `muon_hints`, the `name` and `setup` TXT keys, the status column, the needs-setup prompt, empty-state help, Search again, and `.pot` updated | MR-8 for full testing | 06 OR-2 manual check; builds on all three platforms |
| OR-3 | *Phase 2.* Find new M1 hotspots in the computer's Wi-Fi scan | – | – |

### QA

| ID | Package | Depends on |
|---|---|---|
| QA-1 | Bench measurements B1–B8 (03 §9), **early**: they may change the phone copy and timeouts | OS-1, OS-5, a dev build with MR-1/MR-3 |
| QA-2 | The full bench matrix R1–R16 and the phase 1 acceptance checklist (09 §4) | Everything in phase 1 |

## 3. Order and parallel work

```
Week-ish 1   MR-1 ─────────┐   OS-11 (urgent)  OS-7  OS-1  OS-6  ML-1  UI-0   FL-7 FL-8 FL-9  OR-1
             (contract)    │
Week-ish 2   MR-2 MR-4 MR-7│  OS-5  OS-3(#210) OS-2(#174) ─▶ MR-3   MR-6(#20) ─▶ MR-5   MR-8 MR-9
             UI-1  FL-1  (against fixtures)            QA-1 (B1–B8 on a dev build)
Week-ish 3   UI-2  FL-3  (against the real API)   UI-3  FL-4  FL-2(+CON-1)  OS-4  OR-2
Week-ish 4   UI-4  FL-5  UI-5  FL-6   OS-8  OS-9
Then         QA-2 → phase 1 done
```

- **Critical path:** MR-1 → (OS-3 + #210, OS-2 + #174) → MR-3 → UI-2 / FL-3 → QA-2. Declaring a region on real devices also waits on signing keys and tokens (KAN-321, KAN-132). Until then, every path runs as market `none`.
- **Start OS-11 now.** It fixes a live security bug and needs nothing else.
- **Can start at once**, since they don't touch the new contract: OS-1, OS-6, OS-7, ML-1, UI-0, FL-7, FL-8, FL-9 and OR-1.
- **Can build on fixtures** before MR-1 lands: UI-1 and FL-1.
- **Release together.** MuonOS pins MuonUI and the Fluidd zip, so MR, OS, UI and FL ship in one image (OS-9). OrcaSlicer ships on its own schedule. OR-2's needs-setup handling only lights up once MR-8 is on printers.
- **Rollout safety.** MuonUI redirects into setup only when `muon_setup` reports `new` or `in_progress`, and migration (01 §7) marks existing units `complete`. An image with the new components therefore never pushes a field unit into setup.

## 4. Phase 2 (Bluetooth), after phase 1 acceptance

| ID | Repo | Package |
|---|---|---|
| BT-1 | muon-link, MuonOS | Iroh_BLE's `blelink-bluer` peripheral inside muon-link, the advertising windows, unclaimed gating, the setup ALPN (`provision/1`), the GATE-1 declaration, and coexistence tests (03 §8) |
| BT-2 | muon-link, Moonraker | muon-link reports the transport of each connection; `muon_setup` maps a BLE setup peer to caller kind `bluetooth` (03 §8) |
| BT-3 | Fluidd | "Find nearby" on control.muon3d.com (Web Bluetooth through `blelink-wasm`, Chrome and Edge), running the same screens over Iroh |
| BT-4 | muon3d-app | "Find nearby" in the Muon3D app through `blelink-ffi` |
| OR-3 | OrcaSlicer | Hotspot spotting |

## 5. Jira mapping

| Jira | Relationship to this spec |
|---|---|
| KAN-203 First-run setup flow | **The umbrella.** MR-1–MR-7, UI-1–UI-5, FL-1, FL-3 and FL-5. The completion marker is OS-7. |
| KAN-190 Pairing | **Client** pairing (`/pairing/*`) isn't used by setup phase 1. The account link is NET-10(d) and ADR 0018: muon-link#24 (ML-1), Moonraker#20 (MR-6), the panel's link screen (UI-3, P12) and OS-10. |
| KAN-324 Language and country | **Rescope** to Rev 11 and #31: the region is confirmed after the join (UI-2 P7a, FL-3 §3.1, OS-2), not asked as its own step |
| KAN-311 Wi-Fi onboarding is panel-only | **Close as superseded.** KAN-350 made the hotspot trusted, and the hotspot plus a phone is now the primary path |
| KAN-346 / AP-3 | The hotspot QR code: UI-1 P2 and UI-4 "Add a phone or computer" |
| KAN-364 / NET-3 | MR-8, and the discovery merge in FL-4 |
| KAN-376 Fluidd hotspot card | Overlaps FL-8 |
| KAN-326 Hotspot band pin | A prerequisite of OS-5 |
| KAN-329 Bench measurements | QA-1 and QA-2 (B1–B8, R1–R17) |
| KAN-339 Wi-Fi refactor | OS-3 (W6 `/wifi/saved`, MuonOS#210) |
| KAN-330 Installed base | The migration rule (01 §7) keeps fielded units out of setup; KAN-330's region prompt stays separate |
| KAN-270 prerequisite 2, clock before TLS | The clock rules (01 §2.2) and OS-6 |
| KAN-378 Diagnostic bundle | The "Save a diagnostic bundle" action after repeated region failures |
| SEC-8 Level 1 | `muon_setup` refuses post-setup writes from `lan` and `hotspot` callers in its own caller check (02 §3); `/server/muon/link/start` and `/server/muon/identity/name` go into `PROTECTED_PREFIXES`. `/server/muon/setup` itself is **not** added there. |
| New tickets needed | OS-11 (security), OS-1, OS-4, OS-6, OS-7, OS-10, ML-2, UI-0, MR-8, MR-9, FL-2, FL-4, FL-6–FL-9, CON-1; a follow-up for `aux_api_proxy` to retry its spec fetch in the background |
