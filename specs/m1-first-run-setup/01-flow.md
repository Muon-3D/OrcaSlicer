# 01 · Flow and state machine

This file defines what setup is: the ways into it, the steps, and the rules that move a printer from "new" to "complete". Every surface renders this state machine: the panel, the phone page and the apps. None of them keeps its own copy of progress.

> **Change from the design page (24 Sep).** The design page asked for the country as step 2, before the Wi-Fi scan, following KAN-324. The newer KAN-321 regulatory design (Rev 11, 12 Sep; ADR 0005) replaces that: setup doesn't ask for a country up front. After the printer joins a network, the region is taken from that network and shown as a line to confirm: "Region: United Kingdom · Change". The draft MuonOS#174 and MuonUI#31 already implement this. This spec follows them, so there is one screen fewer for most owners.

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
| 2 | `network` | No | Yes | Never | Wi-Fi is joined and has an IPv4 address, and the region is confirmed (§2.1). Or: Ethernet has an address and the owner accepts it. Internet access is **not** required. | The printer stays offline, and the hotspot stays on as the way in. A "Finish setup" card appears. |
| 3 | `name` | No | "Keep" counts as done | Never | Keep or Rename | – (it can't end up skipped) |
| 4 | `update` | No | Yes ("Later") | When this step becomes current: no internet, no update available, or the clock not synced (KAN-270) | The update is installed and the printer has rebooted into it | No card appears. The normal update notice covers it. |
| 5 | `remote` | No | Yes ("Do this later") | Never | One of: `local` chosen; or `cloud` chosen and the link confirmed on the panel (muon-link phase `linked`). `self_hosted` isn't offered in phase 1. | Behaves as `local`. A "Link a Muon account" card appears. |
| 6 | `ready` | No | Yes | Never | Every required item in the ready manifest is done ([04-panel.md §P13](04-panel.md#p13--ready-to-print-i)) | A "Before your first print" card appears. |
| – | *finish* | – | – | – | – | `finish` marks setup `complete`. This is the "Done" screen, not a step. |

A **time zone** screen can also appear on the panel, after the join, but only when §2.2 needs it. It isn't a counted step.

### 2.1 Region (inside `network`)

The region follows KAN-321 Rev 11, draft MuonOS#174 and draft MuonUI#31. Aux implements it; `muon_setup` reads it and asks for it to be applied.

- **Where the choices come from.** The market token on config partition `p1` lists the configurations the unit may use.
  - A tokened unit with no declared country gets the token's fallback configuration at boot.
  - A unit without a valid token stays on the world domain `00`, which allows 2.4 GHz channels 1–11 at 20 dBm.
  - **Today no unit has a valid token**, because no signing key is installed. Every unit is therefore on `00`, and setup must work that way.
- **The market** is derived from Aux:
  - `none`: `GET /region/options` returns `countries == []`;
  - `locked`: `locked == true` (a US unit);
  - otherwise `picker`.
- **Before joining.** Nothing can suggest a country for one particular network. Aux doesn't expose each network's country element.
  - If the chosen network's channel isn't in `GET /region`'s `channels`, the surfaces run MuonUI#31's `regionPromptFor`.
  - `offer-switch`: "Change your printer's region?" with **Change region** / **Not now**. **Change region** applies the detected country before joining.
  - `locked`: "This printer is set for the United States. Contact support."
- **After joining.** Aux now reports `detected_country` with `basis: "joined-network"`, the associated network's own country element.
  - The panel and the phone show **Region: United Kingdom · Change**. When there's no detection, they show `preselect`.
  - **Confirm** declares it. **Change** opens the picker:
    1. the detected country;
    2. the countries where the chosen language is spoken;
    3. the rest, by name.
  - The step is `done` only once the region is confirmed. The exceptions:
    - `locked` and `none` markets skip the confirmation;
    - so does a unit whose declared country already equals the detected one.
- **Applying** (`POST /region/country`) is live and takes about 8 s. Aux takes every Wi-Fi profile down, **the hotspot included**, applies the country, verifies it and brings the profiles back. It records the declaration only on success.
- **No valid token.** The panel shows "This printer needs re-registering" once, and networks on channels 1–11 still work.

### 2.2 Time zone and clock

The M1 has no RTC. Until NTP (`systemd-timesyncd`) corrects it, a boot starts from `fake-hwclock`'s saved time, or from the image's baked time if `/etc/fake-hwclock.data` isn't persisted (bench B8 checks which). Either way it can be weeks or years behind. The clock-before-TLS rule is KAN-270, prerequisite 2.

- **Phone path.** When the owner taps **Start**, the page posts the phone's clock and IANA time zone. The time zone is then set, whatever the region.
- **Panel path.** The clock comes from NTP after the join. The time zone depends on the declared country:
  - one zone: that zone;
  - several zones: the panel shows a **time zone** list after the region is confirmed, principal zone first (tzdata's `zone.tab` order);
  - no country declared: UTC, until the owner sets one.
- **`update` and `remote: cloud` need a synced clock** for TLS. If NTP hasn't synced within 15 s of the join, `update` is hidden and `remote: cloud` shows `clock_unsynced` ([08-errors.md](08-errors.md)).

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
- **`driver`** is who is in charge: `panel`, `phone`, `web`, `app` (the Muon3D phone app) or `bluetooth` (phase 2).
  - It's advisory. Any allowed writer may write, and a write from a surface that isn't the driver makes that surface the driver.
  - Every write carries the state's `rev`. A write with an old `rev` gets `stale_rev` and the current state, so two screens can't overwrite each other without knowing.
  - A phone, a computer or the app renews its claim every 10 s. The panel treats a claim older than 30 s as lapsed.
- **Persistence.** Every change is saved before the event is sent ([02-setup-api.md §4](02-setup-api.md#4-persistence)). After a power loss, setup resumes at the same `cursor` on whichever screen picks it up.

## 4. Happy paths

### 4.1 Panel only (Wi-Fi with internet, EU-tokened unit)

```
power on → P1 Language (turn, press)
         → P2 Here or phone (Set up here; driver=panel)
         → P4 Wi-Fi list (scan) → P5 password on the ring keyboard → submit
             (a network on a channel the region doesn't allow → "Change your printer's region?")
             op=join: associating → authenticating → dhcp → internet_check → update_check
         → P7 Connected (Muon-walnut-8987.local · 192.168.1.37)
         → P7a "Region: United Kingdom · Confirm / Change"
             op=region_apply (~8 s: Wi-Fi and the hotspot drop and come back)
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
  → the OS captive check is redirected → http://10.42.0.1/setup opens by itself
    (the panel sees a station on ap0 → its QR changes to the page URL)
phone S1 Start → posts the clock and time zone, claims driver=phone → panel P8 Following
phone S3 Wi-Fi list → tap HomeWiFi → password → Connect
phone S4 Joining: warns that the phone will drop off for a few seconds
    join → ap0 follows the uplink's channel → the phone reconnects (the captive window may reopen)
phone S4r Connected → region line "Region: United Kingdom · Confirm / Change"
    Confirm → region apply (Wi-Fi and the hotspot drop ~8 s) → the page restores from state
    (the panel shows every result either way)
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

1. The printer is the only source of truth. A surface that reconnects fetches `GET /server/muon/setup` and never replays local state. A `GET` result replaces the held state whatever its `rev`. `rev` ordering applies only between notifications and write results ([02 §6](02-setup-api.md#6-state-document)).
2. A dropped connection is never treated as a failed operation. The surface shows "Reconnecting to Walnut…" and waits for the state. After 30 s the phone adds "Check the printer's screen". The panel always shows the result.
3. Anything the owner may need later also stays on the panel: the join result, the address and the link code. Losing the phone page loses nothing.
4. The phone page keeps a draft of what the owner typed in `sessionStorage`, guarded by try/catch: the SSID, the security type and the Enterprise identity, **never passwords**. A captive-portal window that closes and reopens can then restore it.

## 6. After setup

- **`finish` does four things:**
  - sets `state: complete`;
  - writes the KAN-203 completion marker;
  - asks the OS to switch the hotspot off in 15 minutes, if an uplink has an address;
  - leaves the captive portal as it is. The redirect stays on permanently, and the page itself switches to its "set up already" view ([03-printer-os.md §2](03-printer-os.md#2-captive-portal)).
- **The "Finish setup" card.**
  - It lists the steps among `network`, `remote` and `ready` whose status is `skipped`.
  - It appears on the panel home screen and as a Fluidd dashboard banner.
  - It goes away when there are none, or when the owner dismisses it. A dismissal is stored as `card_dismissed: true`, and a new skip shows the card again.
- **Writes after `complete`.** Only the step endpoints for `network`, `name`, `remote` and `ready` still accept writes, under the access rules in [07-security.md](07-security.md).

## 7. Printers already in the field

Printers updated from firmware that had no setup flow MUST NOT be sent into setup. The first time `muon_setup` starts with no stored state, it treats the printer as already set up if any of these exist:

- a marker, i.e. Aux `GET /setup/complete` returns `complete: true` (MuonOS KAN-413);
- a saved NetworkManager Wi-Fi profile other than `ap0-con` and the development image's baked `Muon3D_Dev` profile. Reading saved profiles needs `GET /wifi/saved` (MuonOS#210); until then, use `GET /wifi/show?ssid=` on the scan results;
- a link (muon-link `GET /link` returns `linked`);
- a Moonraker database that already holds Fluidd UI settings.

If so, it writes `state: complete`, with every step `done` and `source: "migrated"`, and shows no card. It also writes the marker with Aux `POST /setup/complete {"by": "migrated"}`, retried every 30 s until Aux answers ([02 §4](02-setup-api.md#4-persistence)). Without the marker, H1 would keep the unit's hotspot up for good. Getting those units a region is KAN-330's separate, non-blocking prompt, not this flow.

- **While the check can't finish** (Aux or muon-link isn't answering yet), `muon_setup` reports `complete` provisionally. It persists nothing and writes no marker, and it decides again as soon as they answer. A new printer then moves to `new`, and the panel's store watch sends it into setup ([04 §1](04-panel.md#1-architecture)). A field printer is never sent into setup by a slow boot.
- **Factory QA must end with `muon3d-factory-reset`.** Otherwise a saved Wi-Fi network or Fluidd settings left from QA make a new unit look like one already in the field, and it skips setup.

## 8. Factory reset

Factory reset is `/usr/sbin/muon3d-factory-reset`: root only, at the console or over SSH, running `rugix-ctrl state reset` (KAN-172). It clears everything persisted:

- the Moonraker database (setup state, name override);
- saved networks;
- the Iroh identity, and with it every link;
- the owner's hotspot choice.

It also clears, once the PR that persists each one has merged:

- the setup marker (OS-7, MuonOS#313);
- the declared country (#174; ADR 0005, KAN-351).

**What survives or comes back:**

- The market token on config partition `p1` survives.
- The hostname and SSID are re-derived from the serial, so **the printer keeps its name**.
- The hotspot key is regenerated, so **the Wi-Fi QR code changes**.

After a reset, setup starts again at `new`, and the region line appears again at the next join.

## 9. Timing targets

| Measure | Target |
|---|---|
| Power on to the language screen | Same as today's boot to the home screen. Setup adds no work before first paint. |
| QR scan to page visible (iOS) | ≤ 10 s median |
| QR scan to page visible (Android) | ≤ 15 s median. When it doesn't open by itself, the URL QR code takes one more scan. |
| New printer to "Connected" on Wi-Fi, phone path | ≤ 3 min median over 10 runs |
| Region apply | ≤ 10 s, including the hotspot's return (Rev 11 measured 7.9 s) |
| Wi-Fi join with the right password | ≤ 20 s to an address, plus ≤ 5 s for the internet check |
