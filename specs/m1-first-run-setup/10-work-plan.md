# 10 · Work plan

These work packages are sized for one agent or one PR each. Every package names its repo, what it depends on, and when it's done. The shared contract is [02-setup-api.md](02-setup-api.md), and every package can be built against the fixtures in [`fixtures/`](fixtures/) before its dependencies land.

## 1. Rules for implementing agents

1. **Read first.** Read [README.md](README.md), then the files for your repo. The contract in 02, and the error codes in 08, are authoritative across repos.
2. **Match the repo.** Follow its conventions (lint, tests, style). Each file lists them for its repo.
3. **Don't invent external payloads.** Two interfaces are known by name only:
   - muon-link's `/pairing/*`;
   - the exact shapes of Aux's `/region/*`.
   Read their source. If a field you need isn't there, add it in that repo (it's listed as a package here) or stop and ask. Never guess a shape and hard-code it.
4. **No secrets anywhere.** Passwords and the hotspot key never appear in logs, state, events, fixtures or test snapshots.
5. **Never factory-reset a real unit** (KAN-351). Use `POST /server/muon/setup/reset` on a development unit, which resets setup state only.
6. **Keep the spec in step.** If an implementation needs the contract to change, change this spec in the same PR series and say so in the PR description. The spec lives at `specs/m1-first-run-setup/` in `Muon-3D/OrcaSlicer` until the team moves it.
7. **One PR per package.** Title it `<ID>: <summary> (KAN-xxx)`, and link the spec section in the description.

## 2. Packages

### Moonraker fork (`Muon-3D/Moonraker`)

| ID | Package | Depends on | Done when |
|---|---|---|---|
| MR-1 | `muon_setup` core: the `[muon_setup]` config and template pin, the state document, persistence (namespace `muon_setup`), `rev`/`driver`/`op` handling, caller classification and write hygiene (§3), migration (01 §7), `GET` state and options, `goto`, `skip`, `finish`, `card/dismiss`, `reset`, `notify_muon_setup_changed`, the `muon_setup:complete` event, and the floor entry for `reset`. `finish` calls OS-5 `auto_off` and OS-7 `setup/complete` through `aux_api_proxy`, with a fake in tests. | – | 02 §8 tests 1–4, 7, 8 and 13 pass; flake8 and mypy are clean |
| MR-2 | Language (including Fluidd's default locale), clock, time zone, name (refactor `aux_api_proxy` to expose `set_friendly_name` and `get_identity`), and the tzdata lookup for options | MR-1; OS-6 for real clock and time-zone writes | Unit tests for each; the identity endpoint behaves as before |
| MR-3 | Networks and join: normalising the scan, the region suggestion, the join orchestration (region apply → join → phases), mapping failures, cancel, CA upload, Ethernet, the internet result, and computing update visibility | MR-1; OS-2, OS-3 and OS-4 for the real Aux (fakes until then) | 02 §8 tests 5, 6, 9 and 10 pass |
| MR-4 | The update step through `update_manager`, including the version check after reboot | MR-1 | Test 11 passes |
| MR-5 | The remote step: `local`, `cloud` (with code renewal), `self_hosted` behind a capability flag, and `later` | MR-1, MR-6 | Test 12 passes |
| MR-6 | The `muon_link` component: `GET /server/muon/link`, `POST …/link/start`, `…/confirm` (panel-only and floored) and `…/cancel`, bridging to `127.0.0.1:7131`, plus the `muon_link:link_changed` event | ML-1 | `tests/test_muon_link.py` passes; Fluidd's `startLanLink()` works against a dev unit |
| MR-7 | The ready step: the manifest (loaded from a file, with a default and a schema check), confirm, macro start through `klippy_apis`, hiding undefined macros, and the busy and not-ready guards | MR-1 | Unit tests pass |
| MR-8 | Zeroconf: the `_octoprint._tcp` record advertises **port 80** (nginx), and TXT gains `name=<display>` and `setup=<state>`; the record is re-registered when either changes. Coordinate with KAN-364's NET-3 record so it also carries `setup`. | MR-1 | OrcaSlicer's Bonjour dialog lists a dev unit, and `dns-sd -L` shows port 80 and the TXT keys |
| MR-9 | `aux_api_proxy`: register the identity endpoints in `__init__` (503 while Aux is down), add `setup` to `/server/muon/identity`, and fix `tests/test_aux_api_proxy.py`, whose `FakeServer` has no `database` (42 of 44 tests fail today) | MR-1 | That test file is green |

### MuonOS (`Muon-3D/MuonOS`: Aux API, image, services)

| ID | Package | Depends on | Done when |
|---|---|---|---|
| OS-1 | Captive portal: the dnsmasq drop-in, the nginx `map`/redirect and `location = /setup`, and the image checks | – | B5 passes; iOS and Android open `/setup` by themselves on a dev unit; `nginx -t` passes; GATE-1 check green |
| OS-2 | Region additions (03 §3): `GET /region/suggest` (or the country element in the scan), the `language` parameter, per-configuration channels, the support code, and stable `detail.code` values | – | Aux tests pass; R2 and R9 pass on the bench |
| OS-3 | Wi-Fi: W1 (`hidden`, replacing a saved secret), W2 (error codes, no secret logging), W5 (`/wifi/uplink` plus the internet check on the OTA host, recorded in the privacy inventory), W6 (`/wifi/saved`) | – | Aux tests pass, including a secret-logging test |
| OS-4 | Enterprise: W3 (the `eap` connect) and W4 (`/wifi/ca_cert`) | OS-3 | R6 passes against N6 |
| OS-5 | Hotspot lifecycle H1–H4, `POST /wifi/ap/auto_off`, `GET /wifi/ap/stations`, and floor entries | KAN-326 band pin delivered | H1–H4 verified on a dev unit; R8 passes |
| OS-6 | Time: `GET/POST /time`, `POST /time/zone`, and persisting the time zone | – | Aux tests pass; the time zone survives a reboot |
| OS-7 | The setup marker: `GET/POST/DELETE /setup/complete`, the persist declaration, the ID-9 inventory, and floor entries | – | The marker survives an OTA update; KAN-351's human check shows it cleared by reset |
| OS-8 | Ship `/usr/share/muon/setup/ready.json`, with the hardware team, and the `MUON_SELF_TEST` macro if they want one | Hardware team decision | The manifest validates |
| OS-9 | Pin the new MuonUI and Fluidd builds in the image, updating the Fluidd zip checksum guard (it blocked the KAN-321 C3 run) | The UI and FL packages | The image builds and boots to setup on a clean flash |

### muon-link

| ID | Package | Depends on | Done when |
|---|---|---|---|
| ML-1 | Confirm or add the pairing contract in 03 §6: a status route, the claimant's label at `offer`, 6 digits, a 120 s TTL, single use, and 3 attempts | – | A contract test against the admin endpoint |

### MuonUI (`Muon-3D/MuonUI`)

| ID | Package | Depends on | Done when |
|---|---|---|---|
| UI-1 | The setup shell: client, store, router guard, `screenFor`, `RingProgress`; P1 Language, P2 Here or phone (with `SetupQr` and the network-details view) and P8 Following | MR-1 (fixtures until then) | 04 §7 tests for these screens pass; boots into P1 on a clean unit |
| UI-2 | Network: P4 list, P5 `RingKeyboard`, P5b region line and picker, P6 Joining, P7 Connected and errors, P7b Time zone | UI-1, MR-3 | Tests pass; R1 and R2 pass on the bench |
| UI-3 | P9 Name, P10 Update, P11 Remote, P12 Link code and the knob confirm (the KAN-190 panel half) | UI-1, MR-4, MR-5 | R13 passes |
| UI-4 | P13 Ready, P14 Done, the home "Finish setup" card, and the Settings entries "Add a phone or computer" (the AP-3 QR) and "Link to account" | UI-1, MR-7 | Tests pass |
| UI-5 | Setup strings translated into de, fr, es and it | UI-1–UI-4 | No missing keys (a CI check) |

### Fluidd fork (`Muon-3D/Muon3D_Fluidd`)

| ID | Package | Depends on | Done when |
|---|---|---|---|
| FL-1 | The `/setup` shell: the route (lazy, chrome-less, `printerIndependent`), `client.ts`, `state.ts`, `screen.ts`, the `notifyMuonSetupChanged` socket action, and the bundle budget check | MR-1 (fixtures until then) | 05 §11 client and screen tests pass; `/setup` chunk ≤ 900 KB gzipped |
| FL-2 | Cloud base URL: read `VUE_APP_MUON_CLOUD_URL` (matching `envPrefix: 'VUE_'`, declared in `env.d.ts`), default to `https://app.muon3d.com` on any other origin, and call `setCloudBaseUrl` on init | CON-1 | A printer-served Fluidd can sign in and claim a code |
| FL-3 | The `/setup` screens S0–S10, the region line, the error mapping, and the `en` keys | FL-1, MR-3 | `Setup.spec.ts` and E2E scenarios 1–8 pass |
| FL-4 | `AddPrinterDialog`: the entry points, the `LinkClaim` refactor, discovery merging the mDNS feed with the sweep, and the `setup` badges | FL-1, MR-9 | 06 §2.6 tests pass; R14 and R15 pass |
| FL-5 | The Dashboard "Finish setup" banner | FL-1 | A unit test over the state variants |
| FL-6 | i18n: de, fr, es and it for `app.muon.setup.*` and `app.muon.add_printer.*`; convert the Muon screens touched here to `$t`; add `app.general.confirm.enter_password` and `app.general.btn.connect` | FL-3, FL-4 | `npm run i18n-extract` reports no missing keys for these namespaces |
| FL-7 | Add `app.muon3d.com` to the host blacklist in `public/config.json` and `server/config.json`, so app.muon3d.com doesn't probe itself for 5 s at startup | – | Loading app.muon3d.com shows no 5 s wait |
| FL-8 | Fix `useHotspotCheck` (call it inside `setup()`, and treat only the same origin as the hotspot), and stop the hotspot card's QR from saying `T:nopass` when the key is redacted (overlaps KAN-376) | – | `HotspotManagerCard.spec.ts` extended; passes |
| FL-9 | Route the aux API client through `Vue.$httpClient`, or rebind `auxAxios`'s adapter, so it works over Iroh and carries auth. Stop `AddInstanceDialog` verification going over Iroh. | – | A cloud-activated printer's Wi-Fi card loads |

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
| QA-1 | Bench measurements B1–B6 (03 §9), **early**: they may change the phone copy and timeouts | OS-1, OS-5, a dev build with MR-1/MR-3 |
| QA-2 | The full bench matrix R1–R16 and the phase 1 acceptance checklist (09 §4) | Everything in phase 1 |

## 3. Order and parallel work

```
Week-ish 1   MR-1 ─────────┐          OS-1  OS-5  OS-6  OS-7  ML-1   FL-7 FL-8 FL-9  OR-1
             (contract)    │
Week-ish 2   MR-2 MR-4 MR-7│  OS-2 OS-3 ─▶ MR-3        MR-6 ─▶ MR-5    MR-8 MR-9
             UI-1  FL-1  (against fixtures)            QA-1 (B1–B6 on a dev build)
Week-ish 3   UI-2  FL-3  (against the real API)   UI-3  FL-4  FL-2(+CON-1)  OS-4  OR-2
Week-ish 4   UI-4  FL-5  UI-5  FL-6   OS-8  OS-9
Then         QA-2 → phase 1 done
```

- **Critical path:** MR-1 → (OS-2 + OS-3) → MR-3 → UI-2 / FL-3 → QA-2.
- **Can start at once**, since they don't touch the new contract: OS-1, OS-5, OS-6, OS-7, ML-1, FL-7, FL-8, FL-9 and OR-1.
- **Can build on fixtures** before MR-1 lands: UI-1 and FL-1.
- **Release together.** MuonOS pins MuonUI and the Fluidd zip, so MR, OS, UI and FL ship in one image (OS-9). OrcaSlicer ships on its own schedule. OR-2's needs-setup handling only lights up once MR-8 is on printers.
- **Rollout safety.** MuonUI redirects into setup only when `muon_setup` reports `new` or `in_progress`, and migration (01 §7) marks existing units `complete`. An image with the new components therefore never pushes a field unit into setup.

## 4. Phase 2 (Bluetooth), after phase 1 acceptance

| ID | Repo | Package |
|---|---|---|
| BT-1 | MuonOS | The `muon-setup-ble` daemon, GATT service, SPAKE2 and knob confirm, advertising windows, the GATE-1 declaration, coexistence tests (03 §8) |
| BT-2 | Moonraker | The `bluetooth` caller kind through the daemon's Unix socket (SO_PEERCRED) |
| BT-3 | Fluidd | "Find nearby" on app.muon3d.com (Web Bluetooth, Chrome and Edge), running the same screens over the BLE transport |
| OR-3 | OrcaSlicer | Hotspot spotting |

## 5. Jira mapping

| Jira | Relationship to this spec |
|---|---|
| KAN-203 First-run setup flow | **The umbrella.** MR-1–MR-7, UI-1–UI-5, FL-1, FL-3 and FL-5. The completion marker is OS-7. |
| KAN-190 Pairing | The panel half is UI-3 (P12, the missing MuonUI pairing screen); MR-6 and ML-1 |
| KAN-324 Language and country | **Rescope** to Rev 11: the country is confirmed at join (UI-2 P5b, FL-3 §3.1, OS-2), not asked as its own step |
| KAN-311 Wi-Fi onboarding is panel-only | **Close as superseded.** KAN-350 made the hotspot trusted, and the hotspot plus a phone is now the primary path |
| KAN-346 / AP-3 | The hotspot QR code: UI-1 P2 and UI-4 "Add a phone or computer" |
| KAN-364 / NET-3 | MR-8, and the discovery merge in FL-4 |
| KAN-376 Fluidd hotspot card | Overlaps FL-8 |
| KAN-326 Hotspot band pin | A prerequisite of OS-5 |
| KAN-329 Bench measurements | QA-1 and QA-2 (B1–B6, R1–R16) |
| KAN-339 Wi-Fi refactor | OS-3 (W6 `/wifi/saved`, MuonOS#210) |
| KAN-330 Installed base | The migration rule (01 §7) keeps fielded units out of setup; KAN-330's region prompt stays separate |
| KAN-270 No RTC | The clock rules (01 §2.2) and OS-6 |
| KAN-378 Diagnostic bundle | The "Save a diagnostic bundle" action after repeated region failures |
| SEC-8 Level 1 | Add `/server/muon/setup/*` writes and `/server/muon/link/start` to the protected set (07 §3) |
| New tickets needed | OS-1, OS-4, OS-6, OS-7, MR-6, MR-8, MR-9, FL-2, FL-4, FL-6–FL-9, CON-1, OR-1, OR-2 |
