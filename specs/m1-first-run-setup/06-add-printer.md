# 06 · "Add printer" in every app

Every app has exactly one entry point, **Add printer**. The apps are the printer's own Fluidd, app.muon3d.com and OrcaSlicer, and later the Muon app.

It detects what it can, then routes the owner to the right next step. When detection finds nothing, it asks "What does your printer's screen show?", because the panel always shows the next step. No app ever stops at "No printers found".

## 1. Routing table

This table covers every app. The first matching row wins.

| What the app knows | Where the owner goes |
|---|---|
| The page is served from `10.42.0.1` (the phone or laptop is on the printer's hotspot) | "You're connected to **Walnut** directly" → **Set up Walnut** (`/setup`). If `setup == complete`: **Open Walnut** / **Change Wi-Fi**. |
| A found printer has `identity.setup` = `new` or `in_progress` | **Set up** → `http://<host>/#/setup` (the LAN path, `driver=web`) |
| A found printer has `setup == complete`, and the owner isn't signed in | **Add** → a local instance |
| A found printer has `setup == complete`, the owner is signed in, and the link phase is `unlinked`, `code` or `failed` | **Add**, and **Link to my account** (the existing LAN shortcut, `startLanLink`) |
| Nothing found, or the page is https | "What does your printer's screen show?" (§3) |

## 2. Fluidd: `AddPrinterDialog`

**Repo:** `Muon-3D/Muon3D_Fluidd`. **New file:** `src/components/muon-cloud/AddPrinterDialog.vue`. This folder isn't auto-registered, so import the component explicitly. Use `app-dialog`, which goes fullscreen on mobile.

### 2.1 Entry points to replace

| Where | Today | After |
|---|---|---|
| `src/views/Welcome.vue` | "Search again" / "Enter an address" in the LAN panel; "Link a printer" in the account panel | One primary **Add printer** button in the header. The LAN list stays, now showing status badges (§2.3). The account panel keeps sign-in, and its "Link a printer" button opens **Add printer**. |
| `src/components/muon-cloud/PrinterSwitcher.vue` | "Add a local printer" (AddInstanceDialog) + "Link a printer" | One **Add printer** entry |
| `src/components/muon-cloud/CloudAccountMenu.vue` | "Link a printer" | **Add printer** |
| `src/views/LinkLanding.vue` (`/link?code=`) | The account dialog, then `LinkPrinterDialog` with `initial-code` | The account dialog if needed, then `AddPrinterDialog` opened at the `code` step with the code filled in |
| `src/views/Fleet.vue` | The link button | **Add printer** |

`AddInstanceDialog.vue` stays as the "Enter an address" sub-step. `LinkPrinterDialog.vue`'s claim, confirm and done steps move into a shared piece, either `useLinkClaim()` or a `LinkClaimSteps.vue` component, that both dialogs use. That piece keeps the existing API calls:

- `POST /v1/links/claim`;
- polling `GET /v1/links/:id` every 1.5 s;
- handling for the `linked`, `declined` and `expired` results.

### 2.2 Steps inside the dialog

```
find ──(pick a printer)──▶ setup redirect | add instance | link (code)
  │
  └──"Don't see it?"──▶ screen ──▶ new-guide | code | home-guide | off-guide
                                     │          │
                                     │          └─▶ sign-in (if needed) ─▶ claim ─▶ confirm ─▶ done
                                     └─▶ "Search again" (http) / "I have a code" (https)
```

| Step | Content |
|---|---|
| **find** (http origins only) | Starts `discoverPrinters()` (§2.3). While it searches: "Looking on this network…" with a spinner, then the list. There is always a footer link: **Don't see it? Tell us what its screen shows**. |
| **screen** | "What does your printer's screen show?" Four large choices, each with a small drawing of the panel screen: **A list of languages, or a QR code** · **Six digits** · **Its home screen** · **Nothing, it's dark**. |
| **new-guide** | "It's new. Let's set it up." 1) "On the printer, pick a language." 2) "Scan the QR code on its screen with your phone's camera. Your phone joins the printer and the setup page opens." 3) "About three minutes." A secondary block: "On a computer? Join the Wi-Fi network shown on the printer (turn the knob on the QR screen to see its name and password), then open **http://10.42.0.1/setup**." Button: **Search again** (http), or **I have a code** (https), which goes to **code**. |
| **code** | Six-digit input (existing `v-otp-input`). If the owner is signed out, the dialog first opens `CloudAccountDialog` and comes back. Then claim → confirm ("Turn the knob to **Confirm** and press it") → done. |
| **home-guide** | "Press the knob, then open **Settings › Link to account**. A code appears." Buttons: **I have the code** (→ code), **Enter its address instead** (→ AddInstanceDialog). |
| **off-guide** | "Check that it's plugged in and the switch on the back is on. If the screen stays dark, contact support." |

**Copy.** All copy is i18n'd under `app.muon.add_printer.*` ([05-phone-setup-page.md §9](05-phone-setup-page.md#9-language)).

### 2.3 Discovery on http origins

1. **The mDNS feed.** Read `${BASE_URL}muon/lan-printers.json` first. `config/discoverLanInstances` already fetches it (KAN-364). Every entry in it is a set-up printer on the LAN.
2. **The IP sweep.** Then run the existing IP sweep (`discovery.ts`), unchanged except that `probe()` now reads `identity.setup`.
3. **Merge.** Merge the two lists by `identity.fingerprint`, falling back to the host.
4. **Rows.** Each row shows the name (`display`), the host, and a badge:
   - `setup` is `new` or `in_progress`: **New · needs setup** (accent).
   - `complete` and linked to this account: **In your account**.
   - `complete`: **Ready**.
   - The link status text stays as it is today.
5. **Hotspot banner.** If `location.hostname === '10.42.0.1'`, the banner in §1 sits above the list. Detect this by hostname, not with `useHotspotCheck`, which is broken; fixing it is FL-8 in the work plan.

### 2.4 app.muon3d.com (https)

- **No search.** The browser blocks an https page from reaching `http://` LAN hosts, so the dialog opens at **screen**, with the line "This page can't search your network, so tell us what the printer's screen shows."
- **Phase 2 (Chrome and Edge only).** A **Find nearby** button uses Web Bluetooth ([03-printer-os.md §8](03-printer-os.md#8-phase-2-bluetooth)). Hide it when `navigator.bluetooth` is missing.

### 2.5 Things this dialog depends on

The fixes themselves are in [10-work-plan.md](10-work-plan.md).

- **Cloud base URL (FL-2).**
  - Account calls currently go to `window.location.origin`, because `envPrefix: 'VUE_'` drops `VITE_MUON_CLOUD_URL` and `setCloudBaseUrl` is never called.
  - On any origin other than app.muon3d.com, the default must become `https://app.muon3d.com`.
  - The console API must then accept CORS from any origin for `/v1/*`, with bearer-token auth and no cookies (CON-1).
- **The link endpoints Fluidd already calls.**
  - These are `GET /server/muon/link` and `POST /server/muon/link/start`.
  - They don't exist in the Moonraker fork yet. MR-6 adds them ([03-printer-os.md §6](03-printer-os.md#6-link-service-contract)).

### 2.6 Tests

- `AddPrinterDialog.spec.ts`: covers the routing table in §1 over stubbed discovery results, the https path starting at **screen**, the signed-out code path opening sign-in first, and `/link?code=` starting at **code**.
- `discovery.spec.ts`: covers merging the feed with the sweep, the `setup` badge mapping, and dropping `.invalid` hosts. Stub `fetch` as in `lan-instances.spec.ts`.

## 3. OrcaSlicer

**Repo:** `Muon-3D/OrcaSlicer` (this repo). OrcaSlicer finds printers with its Bonjour dialog, which opens from the Browse button in `PhysicalPrinterDialog` (`src/slic3r/GUI/PhysicalPrinterDialog.cpp:151`). It looks for `_octoprint._tcp`, which Moonraker publishes because `octoprint_discovery: True` is set. It reads the TXT keys `version`, `model` and `addr_pref` (`src/slic3r/GUI/BonjourDialog.cpp:127`), and `addr_pref=hostname` makes it put the `.local` name in the address field.

### OR-1 · Remove the placeholder host

- `resources/profiles/Muon3D/machine/fdm_common_muon_m1.json:471` sets `"print_host": "Muon-M1-SERIALNUMBERHERE.local"`. That name hasn't existed since KAN-357. Set it to `""`.
- Bump `"version"` in `resources/profiles/Muon3D.json` from `02.03.00.11` to `02.03.00.12`, so installed profiles update.
- **Check:** with an M1 preset and no physical printer, the Device tab shows OrcaSlicer's normal "no connection" state and makes no mDNS lookup of the placeholder.
- If the release process uses `MUON-PROFILE-UPDATE-SCRIPT-GENERATOR.ps1` for testers, re-run it.

### OR-2 · Better Bonjour results for M1s, and help when none are found

Changes in `src/slic3r/GUI/BonjourDialog.{hpp,cpp}` and `PhysicalPrinterDialog.cpp`:

1. **Constructor.** `BonjourDialog(wxWindow *parent, PrinterTechnology tech, bool muon_hints = false)`.
   - `PhysicalPrinterDialog` passes `true` when any preset selected in the dialog is a Muon3D system preset or inherits from one. Check the vendor in the preset bundle: `preset.vendor && preset.vendor->id == "Muon3D"`.
2. **TXT keys.** Add `name` and `setup` to `txt_keys`. They are published by MR-8.
   - When `name` is present, show it in the "Service name" column instead of the raw service name.
   - Add a "Status" column when `muon_hints` is true: **Ready** / **Needs setup**.
3. **Choosing a printer that needs setup** (`setup` isn't `complete`) doesn't block. It shows a `MessageDialog`: "Walnut isn't set up yet. Finish setup on its screen, or open its setup page in your browser." with the buttons **Open setup page** (`wxLaunchDefaultBrowser("http://<host>/setup")`) and **Use it anyway**.
4. **Empty result.**
   - When `EVT_BONJOUR_COMPLETE` fires and the list is empty and `muon_hints` is true, replace the label with a help panel (a `wxStaticText`, wrapped at `50 * em`).
   - **Copy:**

     > **No printers found on this network.**
     > Look at the printer's screen:
     > • **A list of languages or a QR code** — it's new. Scan the code with your phone to set it up, then search again.
     > • **Its home screen** — make sure this computer is on the same network, or type its address (Settings › Network on the printer).
     > • **Nothing** — check that it's switched on.

   - Add a **Search again** button that re-runs `show_and_lookup`'s lookup without closing the dialog.
   - Wrap every string in `_L()`, then run `scripts/run_gettext.sh` to update `localization/i18n/OrcaSlicer.pot`.
5. **Dark mode and DPI.** Keep `UpdateDlgDarkUI`, and size everything in `em` units.

**Manual check** (the Bonjour dialog has no unit tests):

- On a LAN with one M1 that is set up and one that isn't, both appear with the right status once MR-8 lands.
- With no printers, the help panel appears, and **Search again** works.
- A non-Muon preset shows the old dialog unchanged.

### OR-3 · Phase 2: spot a new M1's hotspot

- Read the computer's Wi-Fi scan through the platform API: Windows `WlanGetNetworkBssList`, macOS CoreWLAN `CWInterface.scanForNetworks`, Linux NetworkManager D-Bus.
- Match `^Muon-[a-z]{4,7}-[0-9a-f]{4}$`.
- If a match is found, add a row "New M1 nearby: Muon-walnut-8987 · scan the code on its screen with your phone to set it up".
- Only read the scan. Never join a network from the slicer.
- **Parked until phase 2.**

## 4. The future app (phase 3, parked)

- Bluetooth-first discovery, with a "New M1 nearby" card.
- It runs the same `muon_setup` API over BLE ([03-printer-os.md §8](03-printer-os.md#8-phase-2-bluetooth)).
- During setup it pairs over Iroh with the same knob confirmation (KAN-190), so there's no second code.

Nothing in phases 1–2 should block this. The only requirement is that the transports all carry one API.
