# 09 · Testing and acceptance

Each work package lists its own unit tests. This file collects them, adds the end-to-end and bench tests, and states when each phase is done.

## 1. Automated tests, by repo

| Repo | Suite | Run with | Detail |
|---|---|---|---|
| Moonraker | `tests/test_muon_setup.py`, `tests/test_muon_link.py`, additions to `tests/test_muon_floor.py`, and a `[muon_setup]` pin next to `tests/test_m1_config_template.py` | `pytest tests/test_muon_setup.py` (install `pytest`, `pytest-asyncio`, `pytest-timeout`) plus `flake8 --max-line-length=88` and `mypy` over `moonraker`, which is what CI runs | [02-setup-api.md §8](02-setup-api.md#8-tests-minimum) |
| Fluidd | `src/services/muon-setup/__tests__/*`, `Setup.spec.ts`, `AddPrinterDialog.spec.ts`, `discovery.spec.ts`, and the router tests | `npm run lint -- --no-fix && npx vitest run && npm run type-check && npm run circular-check && npm run build` | [05 §11](05-phone-setup-page.md#11-tests), [06 §2.6](06-add-printer.md#26-tests) |
| Fluidd E2E | `tests/e2e/setup.spec.ts` (Playwright) against `tests/e2e/mock-setup-server.mjs` | `npx playwright test` with `executablePath: '/opt/pw-browsers/chromium'` | §2 |
| MuonUI | `setupScreen.spec.ts`, `setupStore.spec.ts`, `setupQr.spec.ts` and a keyboard-set spec | `npm test` and `npm run build`; CI only builds | [04 §6](04-panel.md#6-tests-and-checks) |
| MuonOS / Aux | Tests for the new routes (`/wifi/ap/auto_off`, `/wifi/uplink`, `/wifi/saved`, `/wifi/ca_cert`, `/time*`, `/setup/complete`; the `/region/*` routes are tested in MuonOS#174), W2 error-code mapping, and a secret-logging test over `nmcli_gate.py` | The repo's pytest | [03](03-printer-os.md) |
| MuonOS image | Config checks: the dnsmasq drop-in is present; the nginx config passes `nginx -t`; `/var/lib/muon3d/setup` is in Rugix persist; no new listener shows up in GATE-1's `network.listeners-declared` | The image CI (`scripts/check-nginx-templates.sh`, the recipe pytest suites in `Checks`, and DevTools checks) | [03 §2](03-printer-os.md#2-captive-portal) |
| OrcaSlicer | Build on all three platforms, plus the manual checks | `./build_linux.sh -s` etc. | [06 §3](06-add-printer.md#3-orcaslicer) |
| Muon3D app (`Muon-3D/muon3d-app`, KAN-390) | Tests next to the Wi-Fi QR parser and the hotspot join module. They give a QR code with a known key, then check that the key isn't in any log call, stored value, network request or rendered text, and that the join call is the only place it goes (07 S12). They also check that a driver claim sends `kind: "app"`. | `npm run check` | [07 S12](07-security.md#1-rules), R17 |

## 2. End-to-end test with a mock printer (Fluidd)

`tests/e2e/mock-setup-server.mjs` is a small Node server. It serves the built Fluidd, `GET`/`POST /server/muon/setup*` and a `/websocket` that sends `notify_muon_setup_changed`. It is driven by a script of states, so no real printer is needed.

Scenarios:

1. **Happy path.** Start → pick a network → the region line shows GB (`ap`) → Connect → the phases advance → Connected → Name and Remote (local) → at the printer → Done.
2. **Wrong password.** The error appears, the password field returns, and a retry succeeds.
3. **The phone drops during the join.** The server closes the WebSocket and refuses HTTP for 8 s while it advances the state. The page shows "Reconnecting…", then the result, never an error.
4. **Captive-window reload.** The page reloads mid-flow. The SSID draft is restored and the password field is empty.
5. **`stale_rev`.** The server changes the state under the page. The page re-renders and shows the note.
6. **Region variants.**
   - Before joining, `regionPromptFor` gives `join`, `offer-switch` (Change region, then join) or `locked`.
   - After joining, the region line comes from `joined-network`, or falls back to `preselect`.
   - Confirm drops the connection for 8 s; the page must restore.
   - The `locked` and `none` markets show no line.
7. **Link.** Choose Link a Muon account → a code appears → the code renews after 120 s → `offer` → `linked`.
8. **Following the panel.** `driver.kind = panel` shows S9, and Continue here takes over.

## 3. Bench matrix (hardware)

Record the results in KAN-329. Each row is run from a factory-fresh unit. KAN-351 applies: only a person may factory-reset an M1.

**Phones:**

- iPhone on the current and previous iOS;
- Pixel on the current and previous Android;
- a Samsung Galaxy on the current One UI.

**Computers:** a MacBook on current macOS, and a Windows 11 laptop.

**Networks (access points):**

| AP | Band / channel | Security | Why |
|---|---|---|---|
| N1 | 2.4 GHz ch 6 | WPA2 | Baseline |
| N2 | 2.4 GHz ch 13, UK country IE | WPA2 | KAN-329 criterion 2: must be found, the region resolves to GB, and the join works, from factory state, with no second device on the panel path |
| N3 | 5 GHz ch 36 | WPA3-SAE | The AP follows to 5 GHz |
| N4 | 5 GHz ch 100 (DFS) | WPA2/WPA3 mixed | Does the AP come back? (B2) |
| N5 | 2.4 GHz ch 1, hidden SSID | WPA2 | "Other network…" |
| N6 | hostapd with 802.1X (FreeRADIUS, PEAP-MSCHAPv2, test CA) | Enterprise | eduroam stand-in |
| N7 | 2.4 GHz with a captive portal (e.g. a guest network) | Open | `portal_required` |
| N8 | 2.4 GHz with no upstream internet | WPA2 | `no_internet` |
| N9 | Ethernet with DHCP | – | "Connected by cable" |

**Runs:**

| # | Run | Pass when |
|---|---|---|
| R1 | Panel only, N1 | Reaches Ready. The ring keyboard enters a 12-character password in ≤ 60 s (first-time user). |
| R2 | Panel only, N2, EU-tokened unit. Needs a signing key and token (OS-2). | The network is listed. If the fallback region doesn't allow ch 13, "Change your printer's region?" appears. The join works, the region line shows the United Kingdom, and the apply succeeds (KAN-329 criterion 2). |
| R3 | iPhone path, N1 | The page opens by itself within 10 s of joining. Connected in ≤ 3 min. The phone drops and the page restores. The panel shows the result. |
| R4 | Android path (Pixel and Samsung), N1 | The page opens by itself, or via a notification, or via the URL QR code. Record which (B4). |
| R5 | iPhone path, N3 and N4 | Measure B1/B2 and record what the owner sees. On N4, if the AP doesn't return, the panel shows the address and the phone page shows "Check the printer's screen". |
| R6 | Laptop path via the Wi-Fi details (turn the knob on P2), N6 | Enterprise joins with the CA certificate uploaded, and with "Don't check". |
| R7 | N7 and N8 | Correct warnings. Setup completes. |
| R8 | N9 | Ethernet is accepted. The hotspot turns off 15 minutes after finish. |
| R9 | A US-locked unit on a network on ch 13 | "This printer is set for the United States. Contact support." |
| R10 | A unit with no token (today every unit is `no-signing-key`) | The re-registration notice appears. Joins on ch 1–11 work, and setup completes without a region. |
| R11 | Pull the power during the join, during the region apply, and during the update | It resumes at the same step with `interrupted`, or for the update, the OTA result |
| R12 | A field unit updated from pre-setup firmware | It does **not** enter setup (`migrated`) |
| R13 | Link from the phone path | A code appears on the panel and the phone. It's claimed on control.muon3d.com from another device, confirmed on the knob, and remote is done. |
| R14 | Add printer from control.muon3d.com (https) | Opens at "What does your printer's screen show?". Each answer leads to the right next step. |
| R15 | Add printer from the printer's own Fluidd (http) with a new M1 and a set-up M1 on the LAN | Badges show "Needs setup" and "Ready", and routing is correct |
| R16 | OrcaSlicer Browse with the same two printers, after MR-8 | Both are listed with their status. The empty-state help appears when both are off. |
| R17 | The Muon3D app path on an iPhone and on a Pixel, N1, once the app has setup screens (KAN-390) | The app joins the hotspot from the panel's Wi-Fi QR code, and the panel shows "Setting up from the Muon3D app". Setup reaches Connected and the phone returns to its usual Wi-Fi. The hotspot key isn't in the device log (Console.app on iOS, `adb logcat` on Android), the app's stored data, its crash reports, or its network traffic to any host (07 S12). |

## 4. Acceptance by phase

**Phase 1 is done when all of these are true:**

- [ ] R1–R4 and R7–R16 pass. R5 and R6 are measured and their results recorded.
- [ ] The median time for the iPhone and Android phone paths from new to Connected is ≤ 3 minutes over 10 runs each.
- [ ] No password or hotspot key appears in any Moonraker, Aux or nginx log, or in `journalctl` (including sudo's `COMMAND=` lines; OS-11), after R1–R16. Check both the raw journal and the diagnostic bundle.
- [ ] No request from the phone page leaves for any host other than the printer. Check with the browser's network log on R3.
- [ ] GATE-1's listener check passes with no new listeners.
- [ ] Every automated suite in §1 is green in CI.
- [ ] The `/setup` page and the panel are translated into en, de, fr, es and it. The Muon Fluidd screens touched by this work (Welcome, AddPrinterDialog, LinkPrinterDialog steps) use i18n keys.

**Phase 2 (Bluetooth) is done when:**

- [ ] The Muon3D app on an iPhone and on a Pixel, and Web Bluetooth in Chrome on Android and desktop, each find a new M1 over Iroh_BLE, confirm on the knob, and complete R3's flow with no hotspot join.
- [ ] A BLE setup peer gets caller kind `bluetooth`, and a remote peer still gets `remote` (read-only).
- [ ] Coexistence: an AP + STA + BLE session shows no hotspot drops beyond B1/B2.
- [ ] muon-link advertises over BLE only in the allowed windows.
