# 05 · The `/setup` page (Fluidd fork)

**Repo:** `Muon-3D/Muon3D_Fluidd` (branch `develop`). The stack is Vue 2.7, Vuetify 2.7, vue-router 3 in hash mode, Vuex 3 and vue-i18n 8, built with Vite 5. Tests use Vitest with jsdom, in `__tests__/*.spec.ts` next to the code.

The `/setup` page is the phone and computer renderer of the setup state machine. It is served by the printer, both on its hotspot (`http://10.42.0.1/setup`) and on the LAN (`http://Muon-walnut-8987.local/setup`). It must work inside the captive-portal windows that iOS and Android open by themselves. It is small, it has no Fluidd chrome, and it talks only to `muon_setup` ([02-setup-api.md](02-setup-api.md)).

## 1. Files

| File | What |
|---|---|
| `src/views/Setup.vue` | The route view. Lazy-loaded (`component: () => import('@/views/Setup.vue')`) so the rest of Fluidd doesn't load before it paints. |
| `src/components/muon-setup/*.vue` | One component per screen (S0–S9 below) plus `SetupShell.vue`, which holds the header, the progress bar, the reconnect banner and the language switcher. They are **not** auto-registered (`src/components/muon-setup` isn't in the unplugin dirs), so import them explicitly. |
| `src/services/muon-setup/client.ts` | A standalone client: `fetch` for requests and its **own** `WebSocket` for `notify_muon_setup_changed`. It is independent of `Vue.$socket` and Vuex (§4). |
| `src/services/muon-setup/state.ts` | `Vue.observable` store holding the latest state, connection status and local draft, the same pattern as `src/services/muon-cloud/state.ts`. |
| `src/services/muon-setup/errors.ts` | Maps `error.code` to i18n keys ([08-errors.md](08-errors.md)). |
| `src/services/muon-setup/__tests__/*.spec.ts` | Unit tests (§11). |
| `src/router/index.ts` | Adds the route (§2). |
| `src/router/managedPath.ts` | Adds `/setup` to the chrome-less set (§2). |
| `src/store/socket/actions.ts` | Adds `notifyMuonSetupChanged`, so Fluidd's main socket doesn't log an unknown-action error and the dashboard card (§10) updates. |
| `src/locales/{en,de,fr,es,it}.yaml` | New keys under `app.muon.setup.*` (§9). |

## 2. Route, chrome and boot

**Route:**

```ts
{ path: '/setup', name: 'setup', component: () => import('@/views/Setup.vue'),
  meta: { printerIndependent: true } }
```

**Chrome and boot changes:**

- **Chrome.** Add `'/setup'` to `MANAGED_CONSOLE_PATHS` in `src/router/managedPath.ts`. `App.vue` then hides:
  - the app bar, the drawers and the footer;
  - the socket-disconnected overlay, the klippy card and the global dialogs;
  - the flash messages.

  The page must therefore show every error itself. The loading shell is also skipped, because `v-app v-if="loading && !managedConsoleRoute"`. Update `src/router/__tests__/managed-path.spec.ts` to assert `/setup` is chrome-less and that the paths already listed keep their chrome.
- **Route name.** Don't name the route `/onboarding`, because `managed-routes.spec.ts` asserts that name does not resolve.
- **Redirects.** `printerIndependent: true` stops `main.ts:100-107` from redirecting to `/welcome`, and lets the view render before the socket connects.
- **`init.ts:228-229`** sends every later `appInit()` to `/`. The setup page must never call `appInit`, `activateLocalPrinter` or `activateCloudPrinter`. "Open Walnut" (S8) is a plain link to `http://<host>/`, which loads the full app fresh.
- **Hotspot boot.** On the hotspot, `init.ts` probes `ws://10.42.0.1/websocket` and `:7125`. The first one wins quickly, so leave that as it is. The page doesn't wait for it: `Setup.vue` starts its own client in `created()`.
- **Pretty URL.** nginx maps `GET /setup` to a `302` to `/#/setup`, which gives a short QR payload and portal redirect ([03-printer-os.md §2](03-printer-os.md#2-captive-portal)). The page accepts both forms.
- **Page weight.** Setup must not pull in Monaco, three.js (`PrinterModel3d`), the G-code viewer or the Iroh WASM.
  - Check with `vite build` and the bundle report.
  - The `/setup` entry, with the shared vendor chunk, should be ≤ 900 KB gzipped.
  - First paint on the hotspot should be ≤ 3 s on a mid-range Android phone.

## 3. Screens

The screens below are the **phone** layout: one column, a max width of 480 px, centred on desktop. Tokens come from `src/scss/muon3d.scss` (`--m3d-*`). The page uses the Muon wordmark (`.muon-wordmark`), IBM Plex Sans and the dark palette by default. The design page mockups 1–6 are the visual reference.

The header on every screen shows the `MUON3D` wordmark, the printer's `display` ("Walnut · 8987") and a language switcher, which lists only the configured setup languages. The progress text reads "n of N". Here N counts the steps that aren't `hidden`, and n is the cursor's index among them. The same formula drives the panel ring.

Which screen shows is a pure function of the state (§5). The page keeps no step order of its own. There is no S2: the separate country screen was removed when the region moved into the Wi-Fi step (README D7).

| ID | Shown when | Content | Actions |
|---|---|---|---|
| **S0 Connecting** | No state has arrived yet | Spinner. "Connecting to your printer…" After 10 s: "Make sure this phone is still on the Wi-Fi network *Muon-…*." | – |
| **S1 Start** | `state` is `new` or `in_progress` and the page hasn't claimed the driver | "Set up Walnut · 8987". "About three minutes. Your phone stays connected to the printer until the end, then goes back to your usual Wi-Fi." | **Start** posts `clock` with `Date.now()` and `Intl.DateTimeFormat().resolvedOptions().timeZone` (hotspot origin only, §6), then posts `language` if that step is still pending, then claims `driver=phone`. |
| **S3 Wi-Fi** | cursor = `network`, no `op` | "Choose Wi-Fi". Networks come from `GET …/networks`, and the page rescans on entry and on pull-down or the "Scan again" button. Each row shows the SSID, a lock, signal bars and "Saved" if saved. Tapping a network expands the row inline. Secured networks get a password field (`type=password`, `autocomplete="off"`, `autocapitalize="off"`, `spellcheck="false"`) with Show/Hide, and then **Connect**. A network with `channel_permitted: false` runs the pre-join region check first (§3.1). Enterprise rows open S3b. WEP rows are disabled with "Not supported". Below the list: **Other network…** (S3b), and **Skip for now** with the confirm "Walnut will stay offline. You can connect it later from its screen." Ethernet with an address shows a top card, "Connected by cable · 192.168.1.37", with **Use this connection**. | Connect posts `network`. |
| **S3b Other network** | From S3 | Fields: SSID, and Security (None / WPA2/WPA3 Personal / WPA2/WPA3 Enterprise). Personal adds the password. Enterprise adds EAP method (PEAP, TTLS), inner method (MSCHAPv2, PAP), identity, password, anonymous identity (optional), CA certificate (upload a `.pem`/`.crt`/`.der` → `network/ca_cert`) or "Don't check the certificate (not recommended)", and domain (optional, shown when a CA is given). | **Connect** |
| **S4 Joining** | `op.kind` is `region_apply` or `join` | "Joining HomeWiFi". A checklist following `op`: *Setting the region* (only when an apply runs), *Password accepted* (after `authenticating`), *Got an address · 192.168.1.37*, *Checking internet*, *Checking for updates*. It warns **before** Connect and repeats here: "Your phone may drop off Walnut's Wi-Fi for a few seconds while the printer changes channel. Keep this page open; if it closes, rejoin *Muon-walnut-8987* and it picks up where it left off." | **Cancel** (`network/cancel`), only during `join` |
| **S4r Result** | The join finished (`network.addresses` set), or `network.error` is set | **Success:** "Walnut is on HomeWiFi", plus the address and hostname, then the **region line** (§3.1) while `network.region_confirmed` is `false`. `internet: false` adds the note "No internet. Printing over the network works; updates and remote access are off." `portal_required` gives a warning card with **Choose another network**. **Failure:** the error copy from [08-errors.md](08-errors.md), the password field again with the text visible for `wrong_password`, and **Try again** / **Choose another network**. | Continue |
| **S5 Finish** | cursor is `name`, `update` or `remote` | One page, as in mockup 6. **Name**: a text field pre-filled with `name.value`, max 32. **Update**, only if `update.status == pending`: "Update available · 1.4" with **Install now** / **Later**. **Remote access**: "Keep it on my network" (default) or "Link a Muon account", each with its one-line explanation, ("My own server" isn't offered in phase 1). | **Continue** posts `name`, then `update: later` if the update wasn't installed, then `remote`, in that order. Each post uses the `rev` returned by the one before. |
| **S5u Updating** | `op.kind == "update_install"` | Progress bar from `op.progress`, and "Walnut will restart. This page reconnects by itself." | – |
| **S6 Link** | `remote.mode == "cloud"` and the step isn't done | Renders `remote.link` (muon-link's `LinkPhase`). **`connecting`:** "Getting a code…". **`code`:** the code in big mono digits, updating live as codes renew, with "Enter this code at **app.muon3d.com** on any phone or computer that's online, then confirm on Walnut's screen." A plain link opens `remote.link.url` (no `target`, because of the captive-portal window). **`offer`:** "Confirm on Walnut's screen" and the account. **`linked`:** ✓ "Linked to {account}". **`failed`:** the `link_failed` copy. | **Do this later** posts `remote: later`. |
| **S7 At the printer** | cursor = `ready` | "Last steps happen at the printer", with the manifest items and their live status. | **Skip for now** posts `skip ready`. |
| **S8 Done** | `state == complete` and the page drove setup this session | "Walnut is ready". **Open Walnut** is a link to `http://<hostname>.local/` with the IP link under it. Also "How to send your first print: add Walnut in OrcaSlicer (Printer › Connection › Browse)", and the skipped steps as a list. "Walnut's hotspot turns off in 15 minutes and your phone will go back to its usual Wi-Fi." | – |
| **S9 Following the panel** | `driver.kind == "panel"` and not lapsed | "Walnut's screen is in charge" with the current step name. | **Continue here** claims the driver. |
| **S10 Set up already** | `state == complete` and the page didn't drive setup (e.g. the hotspot fallback, E5) | "Walnut is set up". The network status shows `network.ssid`, or "Not connected" in recovery mode. | **Change Wi-Fi** runs S3/S4 against the complete state. **Open Walnut** is a link. |

### 3.1 Region

This follows KAN-321 Rev 11 and mirrors MuonUI#31 and panel screens P4 and P7a ([04-panel.md](04-panel.md#p7a--region-b)). Port #31's helpers into `src/services/muon-setup/region.ts`:

- `regionPromptFor(region, channel)` → `join | offer-switch | locked`;
- `SPOKEN_IN`;
- country names from `Intl.DisplayNames`.

**Before joining.** If a network's `channel_permitted` is `false`, run `regionPromptFor(state.region, network.channel)`:

| Result | What the page does |
|---|---|
| `join` | Nothing extra |
| `offer-switch` | An inline card, "Change your printer's region?", with **Change region** (join with `region` set to `state.region.detected_country`) and **Not now** |
| `locked` | "This printer is set for the United States. Contact support." |

**After joining** (S4r, while `network.region_confirmed` is `false`):

- **Content:** "Region: **United Kingdom**", taken from `state.region.detected_country` (basis `joined-network`) or else `options.region.preselect`. Then **Confirm** and **Change**.
- **Change** opens a bottom sheet with three sections:
  1. the detected country;
  2. "Where <language> is spoken" (`SPOKEN_IN`);
  3. "All countries", sorted by name, with a search box that filters `options.region.countries`. There's no free text.
- **Confirm, or picking a country,** posts `region {country}`. The page then shows "Applying… Walnut's Wi-Fi will drop for about 8 seconds" and waits for the state, using the reconnect rules in §4.
- **Skipped entirely** in `locked` and `none` markets.

**Market `none`.** A one-time notice at the top of S3: "This printer needs re-registering." Joining still works on channels 1–11.

**Every screen:**

- Shows the reconnect banner (§4) when disconnected.
- Handles `stale_rev` by re-rendering from the returned state, with an inline note: "This step changed on Walnut's screen."
- Handles `busy` with "Walnut is busy with another step. One moment…"

## 4. Client and reconnect rules

`src/services/muon-setup/client.ts`

```ts
export interface SetupClient {
  start (): void                                   // opens WS, GETs state
  stop (): void
  get (): Promise<SetupState>
  post<T = SetupResult> (path: string, body: object, opts?: { timeoutMs?: number }): Promise<T>
  networks (rescan: boolean): Promise<NetworksResult>
  options (country?: string): Promise<SetupOptions>
  uploadCaCert (file: File): Promise<{ ok: boolean, ca_cert_id?: string, error?: SetupError }>
}
```

**Connection:**

- **Base URL.** `window.location.origin`. The page never talks to another host.
- **WebSocket.** Open `ws(s)://<host>/websocket`. On open, send `server.connection.identify` with `{client_name: "muon-setup", version, type: "web", url: location.href}`.
  - For trusted clients no token is needed. If the socket closes with 401, fetch `/access/oneshot_token` and retry with `?token=`.
  - Listen for `notify_muon_setup_changed`, and replace the stored state with `params[0]` if its `rev` is ≥ the current `rev`.
- **Reconnect.** On close or error, retry every 1 s, forever, while the page is visible. After every reconnect, `GET /server/muon/setup`.
- **Polling fallback.** While the socket is down, `GET` every 2 s. The first answer wins.

**Requests:**

- **Writes.**
  - `fetch(POST, JSON)` with `Content-Type: application/json`.
  - The body carries the current `rev`.
  - The timeout is 10 s for every write. `network` answers at once and reports the region apply and the join by event.
- **When a write gets no answer**, because of a network error or a timeout:
  - The page does **not** show a failure. It shows "Reconnecting to Walnut…", keeps the screen, and waits for the next state.
  - If the state then shows the step done or an `op` running, it moves on.
  - If the state shows the step still `pending` with no `op`, it shows "That didn't reach Walnut. Try again."
- **Driver renewal.** While `driver.client_id` is this page's ID and `document.visibilityState == "visible"`, post `driver` every 10 s. The `client_id` is a random UUID made at page load and kept in `sessionStorage` when available.

**The reconnect banner:**

- It appears after 2 s disconnected: "Reconnecting to Walnut…"
- After 30 s it adds: "Check the printer's screen. If this phone left Walnut's Wi-Fi, join *Muon-walnut-8987* again."

**Drafts.**

- Keep the SSID, the security choice and the Enterprise identity in `sessionStorage`, guarded by try/catch, so a reload restores them.
- **Never** store passwords.
- Clear the draft on success.

## 5. Choosing the screen

Selection is a pure function, `screenFor(state, local) -> ScreenId`, in `src/services/muon-setup/screen.ts`. It's unit-tested over fixture states, and it's the only place that decides what shows.

1. No state yet → S0.
2. `state == complete`:
   - `local.droveSetup` is true → S8;
   - `local.changingWifi` is true → S3/S4/S4r;
   - otherwise → S10.
3. `op?.kind` is `region_apply` or `join` → S4. `op?.kind == "update_install"` → S5u.
4. `driver?.kind == "panel" && !driver.lapsed && driver.client_id != local.clientId` → S9.
5. The page hasn't claimed the driver yet → S1.
6. Otherwise, by `cursor`:
   - `language` → S1 (unreachable after Start, which posts the language);
   - `network` → S4r while the join has finished but `network.region_confirmed` is `false` (the region line), or while a result is still unacknowledged; otherwise S3;
   - `name`, `update` or `remote` → S5, or S6 if `remote.mode == "cloud"`;
   - `ready` → S7;
   - `finish` → S8.

## 6. Clock and time zone

- **Only from the hotspot.** The page posts `clock` with `epoch_ms` and `tz` when **Start** is tapped, and only when the origin is `10.42.0.1`. The M1 has no RTC (KAN-270), and the phone is the only trustworthy clock before NTP. A computer on the LAN doesn't post the clock; NTP handles it there.
- **Time zone.** `tz` is always the phone's zone. The panel only shows its time-zone screen when no phone has set one.

## 7. Captive-portal window constraints

| Constraint | Rule |
|---|---|
| The iOS CNA closes when the Wi-Fi network changes, and doesn't share storage with Safari. | Losing the page loses nothing. The result and the link code stay on the panel ([01-flow.md §5](01-flow.md#5-rules-for-surviving-a-dropped-connection)). |
| The CNA ignores `target=_blank` and `window.open`. | Use plain links. Never make a step depend on opening a new window. |
| The CNA can reload the page after the Wi-Fi join. | Rebuild the screen from the state and the `sessionStorage` draft. |
| Service workers may fail in the CNA. | `<register-service-worker />` in `App.vue` must fail silently there. The page must not depend on the service worker. |
| Some Android versions show a "Sign in to network" notification instead of opening by itself. | The panel's second QR (the URL) covers this ([04-panel.md §P2](04-panel.md#p2--here-or-on-a-phone-a)). |
| The phone may prefer mobile data while the Wi-Fi has "no internet". | All requests go to the page's own origin, `10.42.0.1`, which a phone routes over Wi-Fi. Don't add any request to an internet host during setup. |

## 8. Accessibility

- Every control is reachable by keyboard and labelled.
- Status changes are announced in an `aria-live="polite"` region: the joining phases, the reconnect banner and the link code.
- Colour contrast meets WCAG AA in both palettes.
- Touch targets are ≥ 44 px.
- `prefers-reduced-motion` turns off the spinners' rotation.

## 9. Language

- **Page language.**
  - If `steps.language` is done, the page language is `steps.language.value`.
  - Otherwise it's the first setup language that matches `navigator.languages`, or else `en`.
- **Changing language.** The header switcher changes the page language at once. If `steps.language` is still pending, it also posts `language`.
- **Load only what's needed.** Load just the one locale file, through the existing `loadLocaleMessagesAsync`, and don't run `config/onLocaleChange`, which blanks `App.vue` while it waits.
- **Printer-side default.** `muon_setup` sets Fluidd's default locale (`fluidd` namespace, `uiSettings.general.locale`) when the language step completes, if that key is unset. The full Fluidd app then opens in the owner's language.
- **New keys** go under `app.muon.setup.<screen>.<key>` in `en.yaml`, with translations in `de`, `fr`, `es` and `it`. Error keys are `app.muon.setup.error.<code>`. Don't concatenate strings; use named placeholders (`{name}`, `{ssid}`).
- **Missing base keys.** Add `app.general.confirm.enter_password` and `app.general.btn.connect`, which the Wi-Fi card already uses but `en.yaml` lacks.

## 10. The "Finish setup" card in full Fluidd

In `src/views/Dashboard.vue`, add a `v-alert`-style banner above the cards. It reads the `muon-setup` service's state, which the main socket feeds through `notifyMuonSetupChanged` plus one `GET` on connect.

- `state != complete`: "Walnut isn't set up yet." with **Continue setup**, which goes to `/setup`.
- `state == complete`, with skipped steps and `!card_dismissed`: "Finish setting up Walnut" and one link per skipped step. The links go to `/setup?step=network`, `…remote` and `…ready`; for `ready`, the link says "Do this on the printer's screen". A dismiss ✕ posts `card/dismiss`.
- On a printer reached over Iroh (a cloud printer), hide the banner. Setup is local-only.

## 11. Tests

| Test | Covers |
|---|---|
| `screen.spec.ts` | `screenFor` over fixture states, one per row of §3, plus `stale_rev`, a lapsed driver and a complete state with the Wi-Fi change flag. |
| `client.spec.ts` | Uses `vi.useFakeTimers()`, a fake `WebSocket` and a stubbed `fetch`. Covers: reconnect every 1 s; a `GET` after reconnect; polling while down; a lost write that is then resolved by the state; driver renewal only while visible; `rev` ordering, where an older notification is ignored. |
| `Setup.spec.ts` | `shallowMount` with `$t: k => k`. Covers: `regionPromptFor` results before joining (`join`, `offer-switch`, `locked`); the region line after joining, from `joined-network` and from `preselect`; the `locked` and `none` markets; Start posting clock and time zone only on `10.42.0.1`; S3 never persisting a password; S5 posting in order and passing each returned `rev` on. |
| `managed-path.spec.ts` | `/setup` is chrome-less. |
| Router test | `router.resolve('/setup')` resolves to `name: 'setup'` with `meta.printerIndependent`. |
| E2E (Playwright, `/opt/pw-browsers/chromium`) | Against a mock `muon_setup` server (a small Node script in `tests/e2e/mock-setup-server.mjs` that serves the fixture states and scripted `op` progress). Covers the happy path, a wrong password, a dropped connection mid-join (the server drops the WebSocket for 8 s), `stale_rev`, and the link code renewing. |

Run `npm run lint -- --no-fix`, `npx vitest run`, `npm run type-check`, `npm run circular-check` and `npm run build` before pushing. CI runs lint, unit tests, circular-check and build.
