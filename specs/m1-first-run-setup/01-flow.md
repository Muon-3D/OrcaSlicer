# 01 · Flow and state machine

This file defines what setup is: the ways into it, the steps, and the rules that move a printer from "new" to "complete". Every surface renders this state machine: the panel, the phone page and the apps. None of them keeps its own copy of progress.

> **Change from the design page (24 Sep).** The design page asked for the country as step 2, before the Wi-Fi scan, following KAN-324. The newer KAN-321 regulatory design (Rev 11, 12 Sep; ADR 0005) replaces that. Setup doesn't ask for a country up front. When the owner picks a network, the region is resolved from that network and shown as a line to confirm: "Region: United Kingdom · Change". The Aux routes for this are already built. This spec follows Rev 11, so there is one screen fewer for most owners.

## 1. Ways in

| ID | Way in | What happens |
|---|---|---|
| E1 | **At the printer.** Switch it on. | If `setup.state` isn't `complete`, MuonUI opens setup at the cursor instead of the home screen. |
| E2 | **With a phone.** Scan the QR code on the panel. | The QR code joins the phone to the printer's hotspot. The phone's captive-portal check then opens `http://10.42.0.1/setup`, which is the `/setup` route of the printer's own Fluidd. |
| E3 | **From an app.** Choose "Add printer" in app.muon3d.com, the printer's Fluidd, or OrcaSlicer. | The app detects what it can and routes the owner ([06-add-printer.md](06-add-printer.md)). A printer that isn't set up is always routed into this flow. |
| E4 | **Finish setup.** The card on the panel home screen or the Fluidd dashboard. | It runs only the chosen skipped step, then returns. |
| E5 | **Recovery.** A set-up printer loses its network, and the hotspot comes back (AP-4, KAN-347). | A phone that joins the hotspot gets the same `/setup` page, in recovery mode, with **Change Wi-Fi**. |

## 2. Steps

Steps always run in this order. The panel's progress ring and the phone's "n of N" count only the steps that aren't `hidden`.

| # | ID | Required | Can be skipped | Hidden when | Done when | Skipping means |
|---|---|---|---|---|---|---|
| 1 | `language` | Yes | No | Never | A supported language code is saved | – |
| 2 | `network` | No | Yes | Never | Wi-Fi is joined and has an IPv4 address, with the region declared (§2.1). Or: Ethernet has an address and the owner accepts it. Internet access is **not** required. | The printer stays offline, and the hotspot stays on as the way in. A "Finish setup" card appears. |
| 3 | `name` | No | "Keep" counts as done | Never | Keep or Rename | – (it can't end up skipped) |
| 4 | `update` | No | Yes ("Later") | When this step becomes current: no internet, no update available, or the clock not synced (KAN-270) | The update is installed and the printer has rebooted into it | No card appears. The normal update notice covers it. |
| 5 | `remote` | No | Yes ("Do this later") | Never | One of: `local` chosen; or `cloud` chosen and the link confirmed on the panel (muon-link phase `linked`). `self_hosted` isn't offered in phase 1. | Behaves as `local`. A "Link a Muon account" card appears. |
| 6 | `ready` | No | Yes | Never | Every required item in the ready manifest is done ([04-panel.md §P13](04-panel.md#p13--ready-to-print-i)) | A "Before your first print" card appears. |
| – | *finish* | – | – | – | – | `finish` marks setup `complete`. This is the "Done" screen, not a step. |

A **time zone** screen can also appear on the panel, after the join, but only when §2.2 needs it. It isn't a counted step.

### 2.1 Region (inside `network`)

The region follows KAN-321 Rev 11, and Aux implements it. `muon_setup` only reads it and asks for it to be applied.

- **Where the regions come from.** The market token on config partition `p1` lists the configurations the unit may use (`configs`). It can also carry a `default_country`.
- **At boot.** A tokened unit with no declared country gets the token's fallback before NetworkManager starts. A unit without a valid token stays at the world domain `00`, which allows 2.4 GHz channels 1–11 at 20 dBm.
- **Choosing a network gives a suggested country.** The sources, in order:
  1. that access point's country element;
  2. the neighbouring networks, where at least 3 must name a country and that country must lead by at least 2;
  3. the token's `default_country`;
  4. nothing: the owner must choose.

  Every source is filtered through the token.
- **Confirming.**
  - The phone and the panel show the suggestion as a line next to Connect: **Region: United Kingdom · Change**.
  - Pressing Connect with that line on screen is the owner's confirmation.
  - **Change** opens the three-tier picker: the detected country, then the countries where the chosen language is spoken, then every country, reached by continent or initial letter.
- **Applying.** If the confirmed country differs from the one applied, Aux applies it live before the join.
  - Aux takes every Wi-Fi link down, applies the country, verifies it and brings the links back. This takes about 8 s, **and the hotspot drops with it**.
  - The declaration is recorded only when the apply succeeds.
- **Channels the region doesn't allow.** A network on such a channel (for example channel 13 on a US unit) is listed but marked. Choosing it shows "That network is on a channel your printer is not set for." On a US-locked unit it adds "This printer is set for the United States. Contact support."
- **No valid token.** The region line reads "Region: not set". Networks on channels 1–11 still work. The owner sees "This printer needs re-registering" once, with the support code from Aux.

### 2.2 Time zone and clock

The M1 has no RTC. It restores the time from `fake-hwclock` at boot and corrects it only once NTP runs (KAN-270).

- **Phone path.** The page posts the phone's clock and IANA time zone as soon as the owner taps **Start**. The time zone is then set whatever the region.
- **Panel path.** The clock comes from NTP after the join. The time zone depends on the declared country:
  - one zone: that zone;
  - several zones: the panel shows a **time zone** list after "Connected", with the most populous zone first.
  - no country declared: UTC, until the owner sets one in Settings.
- **Before `update` and `remote: cloud`.** Both need a synced clock for TLS. If NTP hasn't synced within 15 s of the join, `update` is hidden and `remote: cloud` shows `clock_unsynced` ([08-errors.md](08-errors.md)).

### 2.3 Other step rules

1. **Language comes first and is global.** It sets the panel language and Fluidd's default language ([02-setup-api.md §5.3](02-setup-api.md#53-language)). The phone page follows its browser language until the owner changes it on the page; if `language` is still pending, the change is posted.
2. **Ready is panel-only.** Its items need a person at the printer. The phone shows progress and offers "Skip for now".
3. **Enterprise Wi-Fi needs a phone or computer.** The panel lists Enterprise networks. Choosing one shows the "Use your phone" QR code rather than three ring-keyboard fields. The phone and computer page supports PEAP and TTLS ([03-printer-os.md §4](03-printer-os.md#4-wi-fi-join)).

## 3. State model

```
setup.state:   new ──(first write)──▶ in_progress ──(finish)──▶ complete
                 ▲                                                  │
                 └────────── reset (factory reset / development) ───┘

step.status:   pending ──▶ done ──(value changed later)──▶ done
                  │
                  ▼
               skipped            hidden (computed when the step becomes current)
```

- **`cursor`** is the step the owner is on: a step ID or `"finish"`.
  - When a step becomes `done` or `skipped`, the cursor moves to the next step that is `pending` and not `hidden`. If there isn't one, it moves to `"finish"`.
  - `goto` can move back to any earlier step that's `done` or `skipped`. It can move forward only as far as the first `pending` step.
  - Going back doesn't reset later steps. Submitting a step again overwrites its value.
- **`op`** is the long-running operation in flight: `region_apply`, `join`, `update_install`, `link` or `ready_item`.
  - While an `op` runs, every write except that `op`'s own cancel returns `busy`.
  - An `op` cut short by a power loss is marked failed with `interrupted` at the next boot. `update_install` is the exception: it follows the OTA rules.
- **`driver`** is who is in charge: `panel`, `phone`, `web` or `bluetooth` (phase 2).
  - It's advisory. Any allowed writer may write, and a write from a surface that isn't the driver makes that surface the driver.
  - Every write carries the state's `rev`. A write with an old `rev` gets `stale_rev` and the current state, so two screens can't overwrite each other without knowing.
  - The phone renews its claim every 10 s. The panel treats a claim older than 30 s as lapsed.
- **Persistence.** Every change is saved before the event is sent ([02-setup-api.md §4](02-setup-api.md#4-persistence)). After a power loss, setup resumes at the same `cursor` on whichever screen picks it up.

## 4. Happy paths

### 4.1 Panel only (Wi-Fi with internet, EU-tokened unit)

```
power on → P1 Language (turn, press)
         → P2 Here or phone (press = set up here; driver=panel)
         → P4 Wi-Fi list (scan) → P5 ring-keyboard password → Done
         → P5b "Join HomeWiFi? · Region: United Kingdom · Join / Change region"
             op=region_apply (~8 s, only if GB differs from the fallback applied at boot)
             op=join: associating → authenticating → dhcp → internet_check → update_check
         → P7 Connected (Muon-walnut-8987.local · 192.168.1.37)
         → [P7b Time zone, only if the country has several]
         → P9 Name: Keep
         → P10 Update (only if one exists): Install now | Later
         → P11 Remote access: Keep it on my network (focused) → done
         → P13 Ready to print: the manifest items
         → finish → P14 Ready → home
```

### 4.2 Phone path (iPhone or Android, Wi-Fi with internet)

```
panel P1 Language (owner picks) → panel P2 shows the Wi-Fi QR code
phone camera scans it → joins "Muon-walnut-8987" (WPA2, per-device key)
  → the OS captive check opens http://10.42.0.1/setup
    (panel sees a station on ap0 → its QR changes to the page URL)
phone S1 Start → posts clock + time zone, claims driver=phone → panel P8 Following
phone S3 Wi-Fi list → tap HomeWiFi → password + "Region: United Kingdom · Change" → Connect
phone S4 Joining: warns that the phone will drop off for a few seconds, twice
    region apply (hotspot down ~8 s) → join → ap0 follows the uplink's channel
    the captive window may close; the phone rejoins the hotspot and the page restores from state
    the panel shows the result regardless
phone S5 Name + Update (if any) + Remote access → Continue
phone S7 "Last steps at the printer" → panel P13 Ready to print → finish
phone S8 Done: address, "Open Walnut", "the hotspot turns off in 15 minutes"
```

### 4.3 From app.muon3d.com

```
Add printer → "What does your printer's screen show?"
  a language list or a QR code → "It's new. Scan the code on its screen with your phone…"
                                 → after setup, come back and enter the link code
  six digits                  → code entry → existing claim → knob confirm on the panel
  its home screen             → "Press the knob, open Settings › Link to account"
```

## 5. Rules for surviving a dropped connection

These apply to every surface.

1. The printer is the only source of truth. A surface that reconnects fetches `GET /server/muon/setup` and never replays local state.
2. A dropped connection is never treated as a failed operation. The surface shows "Reconnecting to Walnut…" and waits for the state. After 30 s the phone adds "Check the printer's screen". The panel always shows the result.
3. Anything the owner may need later also stays on the panel: the join result, the address and the link code. Losing the phone page loses nothing.
4. The phone page keeps a draft of what the owner typed in `sessionStorage`, guarded by try/catch: the SSID, the security type and the Enterprise identity, **never passwords**. A captive-portal window that closes and reopens can then restore it.

## 6. After setup

- **`finish` does four things:**
  - sets `state: complete`;
  - writes the KAN-203 completion marker;
  - asks the OS to switch the hotspot off in 15 minutes, if an uplink has an address;
  - switches the captive portal to landing mode ([03-printer-os.md §2](03-printer-os.md#2-captive-portal)).
- **The "Finish setup" card.**
  - It lists the steps among `network`, `remote` and `ready` whose status is `skipped`.
  - It appears on the panel home screen and as a Fluidd dashboard banner.
  - It goes away when there are none, or when the owner dismisses it. A dismissal is stored as `card_dismissed: true`, and a new skip shows the card again.
- **Writes after `complete`.** Only the step endpoints for `network`, `name`, `remote` and `ready` still accept writes, under the access rules in [07-security.md](07-security.md).

## 7. Printers already in the field

A printer updated from firmware that had no setup flow MUST NOT be sent into setup. The first time `muon_setup` starts with no stored state, it looks for any of these:

- the KAN-203 marker;
- a saved NetworkManager Wi-Fi profile that isn't the hotspot;
- an existing link;
- a Moonraker database that already holds Fluidd UI settings.

If it finds one, it writes `state: complete`, with every step `done` and `source: "migrated"`, and shows no card. Getting those units a region is KAN-330's non-blocking prompt, not this flow.

## 8. Factory reset

- **ADR 0005 and KAN-351.** A factory reset clears the Moonraker database, which holds the setup state and the name override. It also clears saved networks, the declared country, the hostname override and the Iroh identity (and with it every link).
- **The market token on `p1` survives.**
- **After a reset**, setup starts at `new`. The region line appears again at the next join ("Setup re-runs and asks again… costs one screen").

## 9. Timing targets

| Measure | Target |
|---|---|
| Power on to the language screen | Same as today's boot to the home screen. Setup adds no work before first paint. |
| QR scan to page visible (iOS) | ≤ 10 s median |
| QR scan to page visible (Android) | ≤ 15 s median. When it doesn't open by itself, the URL QR code takes one more scan. |
| New printer to "Connected" on Wi-Fi, phone path | ≤ 3 min median over 10 runs |
| Region apply | ≤ 10 s, including the hotspot's return (Rev 11 measured 7.9 s) |
| Wi-Fi join with the right password | ≤ 20 s to an address, plus ≤ 5 s for the internet check |
