# 04 · Panel screens (MuonUI)

**Repo:** `Muon-3D/MuonUI` (checked at `main` `f5ffa7c`). MuonOS pins it as a submodule and in `recipes/muon-ui/files/MuonUI.pin`.

**Stack:**

- Vue 3.5, Vuetify 3.13, Pinia 2.3, vue-router 4 in hash mode (`src/router/index.ts`), Vite 5.
- Vitest 2 with happy-dom.
- It runs as a Chromium kiosk under `cage`, served by the loopback-only nginx vhost on `:100`.
- The vhost's CSP is `script-src 'self'; img-src 'self' data:; connect-src 'self' …`. Anything new must be bundled, and images must be inline SVG or `data:`.
- The shell is a white 480×480 circle using the Vuetify **light** theme, like every other screen. Use the Vuetify theme colours (`rgb(var(--v-theme-accent))`, `rgb(var(--v-theme-secondary))`, `rgb(var(--v-theme-background))`), as `RadialMenuScreen.vue` and `VerticalDotScrollbar.vue` do. **Don't use the `--m3d-*` tokens** in `src/styles/m3d-tokens.css`: `index.html` sets `data-theme="dark"`, so they resolve to the dark palette on the panel.

**Read `AGENTS.md` first.** It is the repo's own conventions file. Note that its mention of `WifiSubMenuView.vue` is stale.

## 0. Open PRs this work builds on

| PR | State | What to do |
|---|---|---|
| [Muon-3D/MuonUI#31](https://github.com/Muon-3D/MuonUI/pull/31) | Draft | It adds the `/setup` route (`src/views/screens/SetupView.vue`), `src/components/RegionPicker.vue`, `src/helpers/regionCountries.ts` (`SPOKEN_IN`, `Intl.DisplayNames`), `src/helpers/regionPrompt.ts` (`regionPromptFor`, `regionPromptMessage`), the connect-time region prompt in `WifiManagerView`, a REGION Settings card and `src/aux_api/regionApi.ts`. **Build the setup flow on #31's branch, with its author's agreement.** Keep the route, `RegionPicker`, `regionCountries`, `regionPrompt` and `regionApi`. Replace `SetupView.vue`'s contents. Remove #31's `aux.setup.get()` and `aux.setup.complete()` calls (in `App.vue`'s `openSetupIfThisPrinterHasNeverBeenSetUp` and in `finish()`), because `muon_setup` owns the marker. |
| [Muon-3D/MuonUI#39](https://github.com/Muon-3D/MuonUI/pull/39) | Open | The KAN-339 Wi-Fi overhaul: `src/aux_api/wifiService.ts`, and `GET /wifi/saved`, which needs MuonOS#210. #31, #39 and this work all edit `WifiManagerView`'s `connectToSelected`, so land them in that order: **#39, then #31, then UI-2.** |
| [Muon-3D/MuonUI#47](https://github.com/Muon-3D/MuonUI/pull/47) | Open | Adds the PROTECTION Settings card (SEC-8) at order 2. |
| [Muon-3D/MuonUI#46](https://github.com/Muon-3D/MuonUI/pull/46) | Draft | Adds `CircularScrollView.vue`, `SideBackButton.vue` and an `html.reduced-motion` class (on by default). Use them once merged. |

## 1. Architecture

| File | What |
|---|---|
| `src/moonraker/useMoonraker.ts` | Keep the transport here, as AGENTS.md asks. Add typed helpers that call `muon_setup` over the existing WebSocket JSON-RPC: `useMoonraker().client.call('server.muon.setup.network', {...})`. JSON-RPC keeps the error messages, which `client.post` drops on non-2xx, and it skips the HTTP Host and Content-Type checks. |
| `src/stores/setupStore.ts` | A **setup-style** Pinia store, like `updateStore.ts`. It holds the latest `muon_setup` state plus panel-only UI state. It subscribes with `client.onNotification('notify_muon_setup_changed', ([s]) => apply(s))`, re-fetches `server.muon.setup` in `client.onOpen(...)`, and starts from `moonrakerStore.init()`. The client already reconnects every 1 s. |
| `src/helpers/setupScreen.ts` | A pure function, `screenFor(state, local) → ScreenId`, with unit tests. It is the only thing that decides which screen shows. |
| `src/views/screens/SetupView.vue` | #31's route view, rewritten to render whatever `screenFor` returns. |
| `src/components/setup/*.vue` | One component per screen below. |
| `src/helpers/setupQr.ts` | Render inline SVG with error correction M and a 4-module quiet zone, using a bundled encoder: `qrcode` (`QRCode.toString(text, {type: 'svg', errorCorrectionLevel: 'M', margin: 4})`) or `uqr` (MIT, no dependencies, the same encoder Fluidd uses). The CSP blocks `blob:`. It must build on Node 18, which is in the CI matrix. |
| `src/i18n/` and `src/locales/{en,de,fr,es,it}.json` | **UI-0.** MuonUI has no i18n today. Add `vue-i18n` and load all five catalogues at startup: P1 shows each language's own title as the knob passes it, and the ready-manifest `title_key`s need them. Use plain `en`, not #31's `en-GB`. |
| `src/dev/mockBackend.ts` and `mockServer.ts` | Add `server.muon.setup*`, and push `notify_muon_setup_changed` built from `specs/m1-first-run-setup/fixtures/`, so `npm run dev:mock` shows every screen. |

**Routing:**

- Keep #31's check at mount, but read `setupStore.state` instead of Aux `/setup`.
- Also watch the store. Whenever the state is `new` or `in_progress` and the route isn't `setup`, call `router.replace({name: 'setup'})`. This also covers a reset. **Exception:** never redirect away from `printing` while a print runs.
- While on `/setup`, suppress `App.vue`'s global update prompt (lines 493–501) and its rollback notice; P10 covers updates.
- While on `/setup`, also suppress the Klippy fault modal. A new unit's Klipper may not be ready, and setup doesn't need it before P13. At P13, a `macro` item that can't run because Klipper isn't ready shows the item's error instead.

**Talking to other services:**

- **During setup the panel doesn't call Aux Wi-Fi routes.** `muon_setup` does.
- **Exceptions:**
  - reading the hotspot credentials for the QR code (§3 P2);
  - muon-link link confirm and cancel, through `/muon-link/` (OS-10).

## 2. Knob and display

**Knob input:**

- Screens use `KnobMenu.vue` with `v-knob-item` (`src/directives/knobItem.ts`).
- **Actions fire on release.**
- Use `holdKnob()` / `holdDuring()` (`src/helpers/knobHold.ts`) during async work.

**Back:**

- Back is the existing **left-rail Back item at order 0**: `SideBackButton.vue` from #46, or the pattern in `SettingsView.vue`.
- There is **no** press-and-hold Back. MuonUI has no long-press primitive, and a hold would clash with the global 5 s power-menu hold (`App.vue:342–347`, `622–634`).
- P1 has no Back item.

**Haptics** (`src/haptics/profiles.ts`):

- Lists use `menu()`, which has soft end stops.
- The language wheel uses `menuWrapping()`.

**Confirmations.** Actions that commit something use `Modal.vue` with `weightedCommit`: the link confirm, skipping Wi-Fi, stopping linking, and unlinking.

**Layout and colour:**

- Centre content with `padding: 24px 56px`, as #31 does, since there is no safe-area helper.
- Use the type scale from `SettingsView.vue` and `Modal.vue`.
- Use `rgb(var(--v-theme-accent))` for highlight and progress and `rgb(var(--v-theme-secondary))` for tracks.

**Progress** uses `ArcProgress.vue`, filled to `index(cursor) / count(visible steps)`.

**Motion.** Honour both `prefers-reduced-motion` and #46's `html.reduced-motion`.

## 3. Screens

Screen letters A–J match the mockups on the design page. There is no P3: the separate country screen was removed (README D7).

### P1 · Language

- **Content:**
  - A `menuWrapping` wheel of the five languages, each in its own language, sorted by endonym: Deutsch, English, Español, Français, Italiano.
  - Focus starts on English.
  - The title ("Welcome") and the hint ("Turn to choose · Press to select") show **in the focused language**.
- **Press:** posts `language`.

### P2 · Here or on a phone (A)

- **Shown when:**
  - the cursor is at `network`;
  - the driver is `null`;
  - this screen hasn't been passed before.
- **Hotspot SSID and key.** Use the same call the HOTSPOT card uses: `aux.ap.apShowCredentialsWifiApShowGet()`. On `:100`, nginx rewrites `/server/aux/wifi/ap/show` straight to Aux with `X-Muon-Local-UI: 1`. That header is the only way to get the key (07 S1). Never log the key or store it.
- **The QR code depends on `setupStore.state.hotspot.clients`,** which `muon_setup` reads from Aux `GET /wifi/ap/stations` (02 §6):

  | State | QR | Text |
  |---|---|---|
  | No station | Wi-Fi join (§5) | "Set up with your phone" / "Scan with the camera" |
  | A station, no phone driver | `http://10.42.0.1/setup` | "Phone connected" / "If the page didn't open, scan this" |
  | A phone, a computer or the Muon3D app claims the driver | – | Go to P8 |

- **Items:**
  - **Set up here** (focused) claims `driver=panel`.
  - **Show network details** shows the SSID and key as text.
- **The hotspot is up during setup** because of H1 ([03-printer-os.md §1](03-printer-os.md#1-hotspot-lifecycle)). The panel doesn't raise it. If `hotspot.up` is `false` (for example just after boot), show "Starting the hotspot…" in place of the QR code and wait for the state to change. **Set up here** stays available.

### P4 · Wi-Fi list (C)

- **Data** comes from `server.muon.setup.networks`.
- **Reuse** `WifiStrengthIcon.vue` or `getWifiIconName()`, `VerticalDotScrollbar.vue`, and the Wi-Fi row styles.
- **Items, in order:**
  1. Back
  2. Scan again
  3. The networks, strongest first
  4. Other network…
  5. Skip for now, confirmed by a `Modal`: "Walnut will stay offline. Its hotspot stays on so you can reach it."
- **Choosing a network runs #31's `regionPromptFor(state.region, network.channel)`:**
  - **`join`:** go to P5, or straight to joining if the network is open or saved. If a saved network fails with `wrong_password`, open P5. #39's `shouldOfferReauthentication()` does the same.
  - **`offer-switch`:** #31's `Modal`, "Change your printer's region?", with **Change region** and **Not now**. **Change region** joins with `region` set to `state.region.detected_country`, and `muon_setup` applies it before joining.
  - **`locked`:** a `Modal` with "This printer is set for the United States. Contact support." and **OK**.
- **Enterprise networks** show the "Use your phone" QR screen.
- **Ethernet with an address** shows "Connected by cable · 192.168.1.37" with **Continue** (posts `network {kind: ethernet}`) and **Use Wi-Fi instead**.
- **No valid token** (`region.market == "none"`): the first time this screen opens, show "This printer needs re-registering". Networks on channels 1–11 still work.

### P5 · Password

- **Keyboard.** Use `KeyboardOverlay.vue` with `KeyboardRing.vue`. Add a `KeyboardSet` to `src/components/keyboard/KeyboardSets.ts`: `AlphaBoard` plus a **"use phone"** key, which opens the P2 URL QR code and keeps the buffer. Shift stays a toggle.
- **Password field.** Add a **Show** toggle and a 1 s reveal of the last character typed. The current field (`WifiManagerView.vue:92–101`) has neither.
- **Submit** posts `network`.

### P6 · Joining

- **A checklist driven by `op`:**
  - *Setting the region* (only for a switch before joining)
  - *Password accepted*
  - *Got an address*
  - *Checking internet*
  - *Checking for updates*
- **Back** asks through a `Modal`, then calls `network/cancel`.

### P7 · Connected (E)

- **Content:** "Connected", the SSID, `hostname_local`, and the IPv4 address.
- **Variants:**
  - `internet: false` adds "No internet · Printing over the network still works."
  - `internet: "portal"` shows the warning variant, with **Choose another network** and **Continue anyway**.
- **Errors** follow [08-errors.md](08-errors.md).

### P7a · Region (B)

This follows KAN-321 Rev 11 and #31: the region comes from the network the printer joined, and the owner confirms it after joining.

- **Shown when:**
  - P7 has finished;
  - `network.region_confirmed` is `false` ([02-setup-api.md §5.6](02-setup-api.md#56-joining-a-network));
  - the market isn't `locked` or `none`.
- **Content:** "Region / **United Kingdom**". The country is `state.region.detected_country` when `basis` is `joined-network`, otherwise `options.region.preselect`. Names come from `regionCountries.ts`.
- **Items:**
  - **Confirm** (focused) posts `region {country}`.
  - **Change** opens #31's `RegionPicker.vue` (detected, then `SPOKEN_IN` for the language, then the rest by name) and posts `region` with the choice.
- **While `op.kind == "region_apply"`:** "Applying… (about 8 seconds)". Wi-Fi and the hotspot drop and come back.
- **On failure:** the message for its code ([08-errors.md §2](08-errors.md#2-region-time-zone-and-clock)).

### P7b · Time zone

- **Shown only** when no phone has set the time zone and the declared country has more than one zone.
- **Content:** the zones in `options.timezones` order (principal zone first), each with its local time.
- **Press:** posts `timezone`.

### P8 · Following a phone (F)

- **Shown while** `driver.kind` is `phone`, `web` or `app`, whether or not the claim has lapsed.
- **Content:** the title for the driver kind, then "<Step> · step n of N", with `ArcProgress`:

  | `driver.kind` | Title |
  |---|---|
  | `phone` | "Setting up from a phone" |
  | `web` | "Setting up from a computer" |
  | `app` | "Setting up from the Muon3D app" |

- **Operations:** an `op`'s progress shows here too, and a join result shows for 5 s.
- **Continue here** claims `driver=panel`.
- **Lapsed** (after 30 s, announced by `muon_setup` without a `rev` change, 02 §6): add "The phone went quiet" for `phone` and `app`, or "The computer went quiet" for `web`. **Continue here** stays focused.

### P9 · Name

- "This printer is called / **Walnut**", with **Keep** (focused) and **Rename**.
- **Rename** opens `KeyboardOverlay` with `AlphaBoard`, for 1–32 characters.
- Either choice posts `name`.

### P10 · Update

- **Shown only if** `update.status == "pending"`.
- **Content:**
  - "Update available / **1.4**";
  - "About 1 minute" if the update is staged (KAN-358);
  - **Install now** (focused) and **Later**.
- **Install now** posts `update {action: install}`. `muon_setup` starts the install through Aux `POST /update/install`.
- **UI-3 must make `updateStore` raise the existing `UpdatingOverlay.vue` (KAN-215)** whenever Aux reports an install in progress, not only when the panel started it.
- The existing rollback `Modal` (`App.vue:87–104`) covers `update_failed`.

### P11 · Remote access (G)

- **Title:** "Use Walnut away from home?"
- **Items:**
  - **Keep it on my network** (focused): "Walnut never contacts Muon. You print from this network."
  - **Link a Muon account**: "Print and watch from anywhere. Muon never sees your files." It's disabled, with "Needs internet", unless `network.internet` is `true` and the clock is synced.
  - **Do this later**.
- "My own server" isn't offered in phase 1.

### P12 · Link code (H)

P12 renders `remote.link`, which is muon-link's `LinkPhase` ([02-setup-api.md §5.9](02-setup-api.md#59-remote-access)).

- **`connecting`:** "Getting a code…"
- **`code`:**
  - The code in large mono digits, grouped in threes. Handle any length.
  - A QR of `remote.link.url`, exactly as given.
  - "Scan, or enter it at / **app.muon3d.com**".
  - `ArcProgress` counts down to `expires_at`. `muon_setup` renews the code after that.
- **`offer`:** a `Modal` with `weightedCommit`: "Link Walnut to / **jed@example.com**?", with the authority key's short form (`fingerprint`, shown as `9f3c 1a7b e2d0 4c11`) underneath (LINK-3).
  - Focus starts on **Cancel**.
  - **Confirm** calls `POST /muon-link/link/confirm`, and **Cancel** calls `POST /muon-link/link/cancel`. Both go through the `:100` location added by OS-10; that's same-origin, so the CSP allows it.
- **`linked`:** "Linked to jed@example.com".
- **`failed`:** the `link_failed` error screen.
- **Back** asks "Stop linking?" through a `Modal`, then calls `remote/cancel`.

### P13 · Ready to print (I)

- **Content:** the checklist from `options.ready_manifest`.
- **Item kinds:**
  - **`confirm`:** an illustration from `public/setup/`, then **Done**, which posts `ready {item, action: confirm}`.
  - **`macro`:** **Start** posts `ready {item, action: start}`. Progress shows during `op.kind == "ready_item"`. A failure shows its message with **Try again** and **Skip**.
  - **`panel_flow`:** the existing MuonUI flow runs, then the panel posts `confirm`.
- **The last item** is **Finish** once the required items are done, and **Skip for now** before that. Either posts `finish`.

### P14 · Ready (J)

- **Content:**
  - "All set / **Walnut is ready**";
  - "Send a print from OrcaSlicer, or open / **Muon-walnut-8987.local**";
  - a URL QR code for `http://<ipv4>/`;
  - the skipped steps.
- **Done** goes to `home`.

### Home: "Finish setup"

- **Where:** Home is the 7-slot `RadialMenuScreen` in `HomeState.vue`, and #46 uses 6 of the 7 slots. Add a **Finish setup** radial item in the free slot.
- **Shown only when:**
  - the state is `complete`;
  - one of `network`, `remote` and `ready` is `skipped`;
  - `!card_dismissed`.
- **It opens** a list of those steps plus **Dismiss**, which posts `card/dismiss`.

## 4. Settings entries

Settings runs inside `CircularScrollView` (#46). Renumber the knob orders:

| Order | Card | Source |
|---|---|---|
| 0 | Back | existing |
| 1 | HOTSPOT (pressing it toggles) | existing, KAN-346 |
| 2 | **Add a phone or computer** | new. P2 outside setup. Raises the hotspot on request (H4). |
| 3 | PROTECTION | #47 |
| 4 | **Link to account** | new. Always shows the state, including "Not linked" (LINK-8). When unlinked it runs P11 → P12. When linked it shows the account, the key's short form, and **Unlink** (a `Modal` with `weightedCommit`, then `POST /muon-link/link/unlink`). |
| 5 | SOFTWARE check | existing |
| 6 | DOWNLOAD & UPDATE / RESTART TO UPDATE | existing |
| 7 | REGION | #31. It calls Aux `/region/*` directly, because the region is a device setting. |

The app fallback copy points here: "Press the knob, then open Settings › Link to account." Settings is reached from the Home radial.

## 5. QR codes

| QR | Payload | Notes |
|---|---|---|
| Wi-Fi join (P2) | `WIFI:T:WPA;S:<ssid>;P:<psk>;;` | Backslash-escape `\`, `;`, `,`, `:` and `"`. Always `T:WPA`. |
| Setup page | `http://10.42.0.1/setup` | nginx redirects it to `/#/setup` (OS-1) |
| Link (P12) | `remote.link.url` as given | The orchestrator sets it |
| Printer address (P14) | `http://<ipv4>/` | Use the IP address, because Android doesn't always resolve `.local` |

**Rendering:** inline SVG, error correction M, a 4-module quiet zone on a white square, and modules of at least 4 px.

## 6. Tests and checks

| Test | Covers |
|---|---|
| `src/helpers/setupScreen.spec.ts` | Every screen, including P7a and P7b. The region variants: `locked`, `none`, and `regionPromptFor` returning `join`, `offer-switch` or `locked`. Also a lapsed driver, an `op` during P8, P8's title for each of `phone`, `web` and `app`, and the Home item. |
| `src/stores/setupStore.spec.ts` | With `src/test/fakeWebSocket.ts`: notifications are applied, `rev` ordering holds, and the store re-fetches on reconnect. |
| `src/helpers/setupQr.spec.ts` | Wi-Fi escaping, and that the PSK is never logged. |
| Keyboard set spec | The "use phone" key keeps the buffer. |

**Before pushing:**

- Run `npm test` and `npm run build`. CI only builds, so the tests won't run unless you run them.
- `npm run typecheck` has 24 errors on `main`, so the rule is **no new errors**.
- There is no lint.
