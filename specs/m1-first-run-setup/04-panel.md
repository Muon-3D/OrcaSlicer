# 04 · Panel screens (MuonUI)

**Repo:** `Muon-3D/MuonUI`. MuonOS pins it as a submodule and in `recipes/muon-ui/files/MuonUI.pin`. It is Vue single-file components with Vuetify, Vite, TypeScript and Vitest. It runs as a Chromium kiosk under `cage`, served by nginx on `127.0.0.1:100`, and its `^~ /server/` block goes to Moonraker.

The knob arrives through `SKControllerStore`, which talks to the sk_daemon WebSocket. There is an on-screen keyboard in `src/components/keyboard/KeyboardKeys.ts` and a Wi-Fi manager in `WifiManagerView.vue`. MuonUI#31 (`03bdbe3`) added a setup route with a Region screen. KAN-346 added a Settings HOTSPOT card that shows the SSID and key as text.

> The MuonUI repo wasn't available when this spec was written, so the file paths below are proposals. The implementing agent should first read `src/router/index.ts`, the MuonUI#31 setup route, `SKControllerStore`, `KeyboardKeys.ts`, `KnobMenu` and `hotspotGuard.ts`. The setup here then **extends** the MuonUI#31 route rather than adding a second one.

## 1. Architecture

| Proposed file | What |
|---|---|
| `src/services/setup/client.ts` | HTTP to `/server/muon/setup*` and a subscription to `notify_muon_setup_changed` over MuonUI's existing Moonraker WebSocket. It follows the reconnect rules in [05-phone-setup-page.md §4](05-phone-setup-page.md#4-client-and-reconnect-rules), with a 1 s retry and a `GET` after each reconnect. |
| `src/stores/setup.ts` | The latest state, plus panel-only UI state such as the focused item, the keyboard buffer and a dismissed "here or phone" flag. Use whatever store pattern `SKControllerStore` uses. |
| `src/views/setup/SetupView.vue` | The route view. It picks the screen with `screenFor(state, local)`, a pure function with unit tests, like the phone page's. |
| `src/views/setup/screens/*.vue` | One component per screen below. |
| `src/components/setup/RingProgress.vue` | The progress arc around the rim. |
| `src/components/setup/RingKeyboard.vue` | The rim keyboard (§4). |
| `src/components/setup/SetupQr.vue` | QR rendering for a round display (§5). |
| Router guard | At boot and on every navigation, it redirects to `/setup` while `state` is `new` or `in_progress`. The only exception is when the owner opened a Settings screen from inside setup (§6). |

- **Rendering rule.** The panel renders `muon_setup` state and sends it intents. It never calls `/server/aux/wifi/*` directly during setup; `muon_setup` does that.
- **Existing screens.** Reuse the existing `WifiManagerView.vue` pieces (list rows, signal icons) and the MuonUI#31 region screen's pieces wherever they fit.

## 2. Knob and display model

**The display:**

- It is a 480×480 round DSI panel. Keep content inside the inscribed square (339×339, centred) unless it is deliberately on the rim: the progress ring, the ring keyboard and the list wheel's fade.
- Minimum text size is 18 px for body text and 28 px for titles. The display is read at arm's length.

**The knob:**

| Input | Meaning |
|---|---|
| Turn | Move focus one item per haptic detent. At the ends of a list there is a hard end stop, with no wrap-around. On the language wheel and the ring keyboard, focus wraps. |
| Press | Choose the focused item. |
| Press and hold (≥ 600 ms) | Back: `goto` the previous non-hidden step, or leave a sub-screen. On P1 (language) it does nothing. |

**Progress.** The rim ring fills to `index(cursor) / count(visible steps)`. It uses the device accent `#80d6d1` on a `#1d2427` track.

**Motion.** Screens slide horizontally: forward goes left, back goes right. The motion runs 180 ms with `cubic-bezier(0.2,0,0,1)` and is off when reduced motion is set.

## 3. Screens

Screen letters A–J match the mockups on the design page. There is no P3: the separate country screen was removed when the region moved into the join (P5b, README D7).

### P1 · Language

- **Content:**
  - A wheel of the configured languages written in their own language and sorted by endonym (KAN-324): Deutsch, English, Español, Français, Italiano.
  - Focus starts on English.
  - The title ("Welcome") and the hint ("Turn to choose · Press to select") are shown **in the focused language**, so a person who can't read English still knows what to do.
- **Press** posts `language` and goes to P2.
- **Nothing else** is on this screen. No QR code yet, because every later screen depends on the language.

### P2 · Here or on a phone (A)

- **Shown when** the cursor is at `network`, the driver is `null`, and the owner hasn't pressed through this screen before.
- **Content**, based on `hotspot.clients` and the driver:

  | State | QR code | Text |
  |---|---|---|
  | No station on `ap0` | Wi-Fi join QR (§5) | "Set up with your phone" / "Scan with the camera" / "or press to set up here" |
  | At least one station, no phone driver | URL QR `http://10.42.0.1/setup` | "Phone connected" / "If the page didn't open, scan this" / "or press to set up here" |
  | A phone claims the driver | – | Go to P8 |

- **Turn** reveals the network details, the same values as the HOTSPOT card (KAN-346): "Network **Muon-walnut-8987**" and "Password **<12-char key>**". These are for laptops and anyone who can't scan. Turning back returns to the QR code.
- **Press** claims `driver=panel` and continues to P4.
- **The hotspot must be up** on this screen. If it's down, the panel asks Aux to bring it up (AP-4: up while the station isn't connected).

### P4 · Wi-Fi list (C)

- **Rows, in order:**
  1. **Scan again**.
  2. The networks, strongest first. Each shows the SSID, a lock icon and signal bars.
  3. **Other network…**
  4. **Skip for now**.
- **Rescan.** The list rescans on entry. A small phone icon in the header re-opens the P2 QR screen.
- **Choosing a network:**
  - A WPA2 or WPA3 Personal network opens P5.
  - An open network goes straight to P5b.
  - A saved network joins with its stored secret. If the join fails with `wrong_password`, P5 opens.
  - An **Enterprise** network opens the "Use your phone" screen: "Enterprise networks need a phone or computer", the URL QR and "Press to go back".
  - **Skip for now** asks for confirmation: "Walnut will stay offline. Its hotspot stays on so you can reach it." (Confirm / Back).
  - A network with `channel_permitted: false` is listed with the muted note "Not available in this region". Choosing it shows "That network is on a channel your printer is not set for." On a US-locked unit it adds "This printer is set for the United States. Contact support." The action is **Back**.
- **No valid token** (`region.market == "none"`): the first time this screen opens, show "This printer needs re-registering" with the support code, then the list. Networks on channels 1–11 still work.
- **Ethernet.** If `ethernet.address` is set when this screen opens, the screen shows "Connected by cable · 192.168.1.37" with **Continue**, which posts `network {kind: ethernet}`, and **Use Wi-Fi instead**.

### P5 · Password, ring keyboard (D)

- **Centre:** the SSID, and a password field that shows dots with the last character visible for 1 s. **Show** toggles plain text.
- **Rim:** the ring keyboard (§4).
- **Done** goes to P5b.
- **Back:** press and hold leaves to P4, after a confirmation if the buffer isn't empty.

### P5b · Join and region (B)

This screen follows KAN-321 Rev 11. The region is confirmed as a line at the moment of joining, not asked up front.

- **Content:** "Join / **HomeWiFi**?", then the line **Region: United Kingdom**, and the rows **Join** (focused) and **Change region**.
  - The country comes from the chosen network's `region_suggestion`.
  - When the suggestion's source is `default` (the token's default country), the line reads "Region: Germany · Is this right?" so the owner looks at it.
  - When there's no suggestion at all, **Join** is disabled and focus starts on **Choose region**.
- **Skipped entirely** when `region.market == "locked"`, or when the country is already declared and matches the suggestion (for example, a later change of network). P5 goes straight to P6.
- **Join** posts `network` with `region` and goes to P6. The phases start with "Setting the region…" when an apply is needed.
- **Change region** opens the picker, in three tiers:
  1. The detected country, if any, with **Confirm**.
  2. "Where <language> is spoken": `region.for_language`.
  3. **All countries…**: first `region.all`'s continents, then that continent's countries. If only one continent is offered, go straight to the country list.

  The picker is only ever a list from `options.region`. There is no free text. Choosing a country returns here with the new line.
- **Failures:**
  - `region_apply_failed`: "We could not apply that region." with **Try again** and **Choose another region**. After two failures, a third option appears: "Save a diagnostic bundle" (KAN-378).
  - `needs_reregistration`: "This printer needs re-registering" and the support code.
  - `region_not_offered`: "This printer is not registered for that country."

### P6 · Joining

- **Content:** a checklist driven by `op`: *Setting the region* (only while `op.kind == "region_apply"`), *Password accepted*, *Got an address*, *Checking internet*, *Checking for updates*. Then P7 or an error screen.
- **Press** does nothing while joining.
- **Press and hold** asks "Cancel joining?" and then calls `network/cancel`.

### P7 · Connected (E)

- **Content:** "Connected", then the SSID, then `hostname_local` (`Muon-walnut-8987.local`) with the IPv4 address under it, then "Press to continue".
- **`internet: false`** adds "No internet · Printing over the network still works."
- **`portal_required`** shows a warning variant: "HomeWiFi needs a sign-in page, which a printer can't complete." with **Choose another network** and **Continue anyway**.
- **Error screens** for the other failures follow [08-errors.md](08-errors.md). Each has one primary action (press) and one secondary action (turn to it).

### P7b · Time zone

- **Shown only on the panel path**, after P7, when `clock.tz` isn't set by a phone and the declared country has more than one zone ([01-flow.md §2.2](01-flow.md#22-time-zone-and-clock)).
- **Content:** "Which time zone?" with the zones for the country, most populous first. Each row shows the zone's city name and the current local time, for example "Chicago · 14:05".
- **Press** posts `timezone`. There is no skip: the first row is a good default, and pressing through accepts it.

### P8 · Following a phone (F)

- **Shown while** `driver.kind` is `phone` or `web` and the claim hasn't lapsed.
- **Content:** "Setting up from a phone" (or "from a computer" for `web`), then "<Step name> · step n of N". The ring fills as usual.
- **During operations:** while an `op` runs, the matching progress (setting the region, joining) is shown *here too*, so the owner can watch from either screen.
- **Results:** when a join finishes, the result (P7 or the error) is shown for 5 s and then returns to P8. The panel must show the result even if the phone has dropped off.
- **Press** claims `driver=panel`, and the panel continues at the cursor.
- **Lapsed claim** (more than 30 s without renewal): the screen reads "The phone went quiet" with **Continue here** (press). It doesn't take over automatically.

### P9 · Name

- "This printer is called / **Walnut**" with **Keep** (focused) and **Rename**.
- Rename opens the ring keyboard with the current name selected. The name is 1–32 characters and letters are allowed in any case. It is then posted with `name`.

### P10 · Update

- **Shown only if** `update.status == "pending"`.
- **Content:** "Update available / **1.4**", "About 1 minute" if staged (KAN-358), otherwise "A few minutes". **Install now** (focused) and **Later**.
- **Installing:** the full-screen updating state (KAN-215 if it exists, otherwise a progress ring from `op.progress`), then the reboot.
- **After the boot,** the panel resumes at P11. If the update rolled back: "The update didn't finish. Walnut is still on 1.3.2." with **Continue**.

### P11 · Remote access (G)

- **Title:** "Use Walnut away from home?"
- **Options,** each with one sentence of copy:

  | Option | Copy | Focus |
  |---|---|---|
  | **Keep it on my network** | "Walnut never contacts Muon. You print from this network." | **Focused by default** (Tier 1) |
  | **Link a Muon account** | "Print and watch from anywhere. Muon never sees your files." | |
  | **My own server** | "Use your organisation's server." Only if `capabilities.self_hosted`. On the panel it shows the URL of the Fluidd page where the owner enters the server address. | |

- **Link a Muon account** posts `remote {mode: cloud}` and goes to P12.
- **Unavailable:** if `network.internet` isn't `true` or the clock isn't synced, *Link a Muon account* is disabled with "Needs internet", and pressing it explains why.
- **Do this later** is the last row. It posts `remote {mode: later}`.

### P12 · Link code (H)

This is the MuonUI pairing screen that KAN-190 still lacks.

- **Content:**
  - The code in large mono digits, grouped 3+3 (`482 913`).
  - The link QR (§5).
  - "Scan, or enter it at / **app.muon3d.com**".
- **Countdown.** The ring counts down the code's 120 s TTL. `muon_setup` renews it automatically while this screen is open ([02-setup-api.md §5.9](02-setup-api.md#59-remote-access)).
- **After the claim** (`phase: offer`), the confirm screen shows: "Link Walnut to / **jed@example.com**?" with **Confirm** and **Cancel**. Focus starts on **Cancel**, so a stray press can't link. Turning to Confirm and pressing links the printer. This is the owner's physical-presence proof ([07-security.md](07-security.md)).
- **Linked:** "Linked to jed@example.com", then continue.
- **Back:** press and hold asks "Stop linking?" and then calls `remote/cancel`.

### P13 · Ready to print (I)

- **Content:** the checklist from `options.ready_manifest`. Each item row shows its state.
- **Choosing an item** opens its screen:
  - **`confirm`** items show an illustration and instructions, then "Done" (press), which posts `ready {item, action: confirm}`.
  - **`macro`** items show "Start" (press), which posts `ready {item, action: start}`. Progress shows while `op.kind == "ready_item"`, then the result. A failure shows the gcode error message and **Try again** / **Skip**.
  - **`panel_flow`** items open the existing MuonUI flow (e.g. load filament). When it returns successfully, the panel posts `ready {item, action: confirm}`.
- **Finishing.** The last row is **Finish** once every required item is done, or **Skip for now** before that. Either one posts `finish`.

### P14 · Ready (J)

- "All set / **Walnut is ready**"
- "Send a print from OrcaSlicer, or open / **Muon-walnut-8987.local**"
- A URL QR `http://<ipv4>/` for phones on the same network. It isn't shown when the network was skipped.
- If any steps were skipped: "Still to do: Connect to Wi-Fi · Link a Muon account"
- **Press** goes to home.

### Home: "Finish setup" card

- **Shown** on the home screen when `state == complete`, a step among `network`, `remote` and `ready` is `skipped`, and `!card_dismissed`.
- **Opening it** lists those steps. Choosing one runs that step's screens only and then returns home.
- **Dismiss** is its own row and posts `card/dismiss`.

### Settings entries

Add these, or extend the existing ones:

| Entry | What |
|---|---|
| **Add a phone or computer** | The P2 screen outside setup. It brings the hotspot up if it's down (AP-5) and shows the Wi-Fi QR code, the details and the URL QR code. This builds the missing AP-3 QR. |
| **Link to account** | P11 → P12 outside setup. The app fallback copy tells owners to use it: "Press the knob, then open Settings › Link to account." |
| **Region** | The existing MuonUI#31 region screen, unchanged. It calls Aux `/region/*` directly, because the region is a device setting, not a setup step. After an apply, `muon_setup` sees the change the next time it reads `GET /region`. |
| **Wi-Fi** | The existing `WifiManagerView.vue`. Choosing a network while setup is complete also goes through `muon_setup` (`network`), so the phone page and the panel stay in step. |

## 4. The ring keyboard

**Rim slots.** 32 slots sit around the rim, 11.25° apart. Slot 0 is at 12 o'clock and they run clockwise:

| Slots | Lower-case layer | Upper-case layer | Symbols layer |
|---|---|---|---|
| 0–25 | `a`–`z` | `A`–`Z` | `0`–`9` then ``! @ # $ % & * - _ . , ; : ' " ?`` |
| 26 | space | space | `/ \ + = ( ) [ ] { } < > ~ ` ^ \|` (extra symbols on a second page, reached with `26`) |
| 27 | ⌫ delete | ⌫ | ⌫ |
| 28 | ⇧ caps (one-shot; double press locks) | ⇧ | – |
| 29 | `123` / `abc` layer switch | | |
| 30 | 📱 "use phone" (opens the P2 QR screen, keeping the buffer) | | |
| 31 | ✓ Done | | |

**Using it:**

- **Controls.** The focused slot is enlarged and filled with the accent. One detent moves one slot, and focus wraps. Press types the focused character or activates the control. Press and hold is Back.
- **Reuse.** Take `KeyboardKeys.ts`'s key sets as the source of symbols so the existing keyboard and this one agree.
- **Acceleration.** Turning fast (more than 6 detents in 300 ms) skips 2 slots per detent. This keeps a 26-letter rim usable.
- **Masking.** The typed buffer is shown in the centre. Passwords are masked except for the last character, which is visible for 1 s, and "Show" reveals the whole password.
- **Target.** Typing a 12-character password should take no more than 60 s for a first-time user. Test it on the bench ([09-testing.md](09-testing.md)).

## 5. QR codes

| QR | Payload | Notes |
|---|---|---|
| Wi-Fi join (P2) | `WIFI:T:WPA;S:<ssid>;P:<psk>;;` | Escape `\`, `;`, `,`, `:` and `"` in the SSID and PSK with a backslash. Always `T:WPA`, never `nopass`; unlike the Fluidd card bug, the panel has the real key. |
| Setup page (P2, "Use your phone") | `http://10.42.0.1/setup` | nginx redirects it to `/#/setup`. |
| Link (P12) | `https://app.muon3d.com/link?code=<6 digits>` | Opens Fluidd's existing `LinkLanding` route. |
| Printer address (P14) | `http://<ipv4>/` | Use the IP, not `.local`, because Android can't always resolve mDNS. |

**Where the panel gets the Wi-Fi key.** `muon_setup` never puts the hotspot key in its state. The panel reads it through the same panel-only path the HOTSPOT card uses today (KAN-346). If that path is an Aux route, it must stay loopback-only, and it must be added to `muon_floor.FLOOR_PREFIXES` if it goes through Moonraker.

**Rendering:**

- Error correction level M.
- Dark modules on a white rounded square with a 4-module quiet zone.
- Module size ≥ 4 px. The Wi-Fi QR, about 41×41 modules, is then about 200 px, which fits inside the inscribed square with room for a line of text.

## 6. Leaving and re-entering setup

- **Settings during setup.** From any setup screen, the panel's Settings shortcut (if MuonUI has one) is allowed for Wi-Fi, Region, Hotspot and About. Returning goes back to `/setup`. Printing, moving the axes and the file browser are not reachable until `state == complete`. The one exception is a `ready` item's own flow.
- **Power loss.** The panel resumes at `cursor`. A screen that was mid-operation shows the result of that operation (`interrupted`), not a spinner.

## 7. Tests

| Test | Covers |
|---|---|
| `screenFor.spec.ts` | Every screen P1–P14 including P5b and P7b; the region variants (`locked`, `none`, `picker` with an `ap`, `neighbours`, `default` or missing suggestion, and a channel that isn't permitted); a lapsed driver; `op` during P8; and the post-setup card. |
| `RingKeyboard.spec.ts` | Slot mapping for each layer, wrap-around, acceleration, one-shot and locked caps, masking and the last-character reveal, and that the phone slot keeps the buffer. |
| `SetupQr.spec.ts` | Wi-Fi escaping, with SSIDs and PSKs containing `;`, `:` and `\`. The PSK must never be logged. |
| Router guard | Redirects to `/setup` while setup is incomplete, and stops redirecting after `complete`. |
| On the bench | [09-testing.md](09-testing.md) lists the knob and hardware checks. |
