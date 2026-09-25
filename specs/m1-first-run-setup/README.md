# M1 first-run setup: implementation spec

**Status:** v1 for implementation · 24 September 2026
**Design companion:** the "M1 First-Run Setup" design page (panel and phone mockups), at https://claude.ai/artifact/4Vx7JULUcdeXfpZhnoxFU2
**Jira umbrella:** KAN-203. Related: KAN-190, KAN-321/324 (region), KAN-346, KAN-364, KAN-376, KAN-329.

This folder specifies how a new Muon3D M1 gets from the box to its first print. It is written for the agents and engineers who will build it across six repos. Each file is self-contained for its audience. [10-work-plan.md](10-work-plan.md) splits the work into packages sized for one PR each.

## What we're building

- **One flow, one entry point.** Setting up, pairing and linking are steps of one flow, which lives on the printer. Every app offers a single **Add printer** that works out which steps are left. Nobody is asked whether their printer is set up.
- **It starts at the printer.** Switch it on and the round knob panel starts setup: **language → network → name → update → remote access → ready to print**. The panel can finish every step on its own, offline.
- **The knob picks, the phone types.** From the second screen, the panel shows a QR code. Scanning it joins the phone to the printer's hotspot, and the setup page opens by itself as a captive portal. No app or internet is needed. The page is a `/setup` route in the Fluidd fork, served by the printer. The panel and the phone show the same state, and either one can take over at any time.
- **The region is confirmed after joining.** Following KAN-321 Rev 11, draft MuonOS#174 and draft MuonUI#31, the region comes from the network the printer joined. It shows as a line, "Region: United Kingdom · Change", which the owner confirms.
- **Linking is last and optional.** At the remote-access step, **Keep it on my network** is focused. **Link a Muon account** reuses today's 6-digit code and knob confirmation. The printer contacts Muon only after the owner chooses it (Tier 1).
- **Nothing dead-ends.** When an app can't find the printer, it asks "What does your printer's screen show?", because the screen always shows the next step.

## Decisions (locked for v1)

| # | Decision | Why |
|---|---|---|
| D1 | Setup starts at the printer, and every app routes printers that aren't set up into the same flow. | That's where the owner looks after unboxing, and it works with no phone, account or internet. |
| D2 | **Keep it on my network** is focused at remote access, and linking is an equal choice. | The Tier 1 promise (KAN-187). Changing the focus later changes nothing else. |
| D3 | Bluetooth is phase 2, as a second transport for the same API. | iPhones can't use Web Bluetooth, so the hotspot path must work first. |
| D4 | The setup page opens as a captive portal, using DNS and nginx on the hotspot. It's permanent, with no runtime toggling. | The alternative is a second QR scan for everyone. The hotspot never routes anywhere (AP-7), so hijacking DNS is harmless. |
| D5 | The hotspot turns off 15 minutes after a successful setup when an uplink is connected (H3). AP-4 still brings it back whenever the uplink is lost. | It sends the phone back to its usual Wi-Fi, and it shrinks the attack surface. |
| D6 | The "Ready to print" content is **provisional** and data-driven (a manifest). The hardware team owns the items. | The owner's decision was pending. The flow ships with a default manifest the team can edit without code changes. |
| D7 | The region follows KAN-321 Rev 11 and MuonUI#31: it's confirmed **after** the join, not asked up front. This **changes the design page**, which had the country as step 2. | Rev 11 (12 Sep) and ADR 0005 supersede KAN-324's order. Aux can't suggest a country for a network before joining it. This removes a screen for most owners. |
| D8 | A factory reset clears the declared country, and setup asks again. The market token survives. | ADR 0005 and KAN-351 supersede KAN-325's "keep the country". |
| D9 | `muon_setup`, a Moonraker component, owns the state machine. MuonUI and the `/setup` page are thin renderers. | One source of truth. It survives phones dropping off and power loss, and it lets Bluetooth and the future app reuse the same API. |
| D10 | Enterprise Wi-Fi (PEAP and TTLS) is supported from a phone or computer, not from the panel. | Three ring-keyboard fields plus a certificate is too much for a knob. Labs and universities need it. |
| D11 | The hotspot reaches **only the printer**: no internet, no LAN and no other hotspot client (AP-7 made true). The printer routes nothing. Its DNS answers every name with `10.42.0.1`. | That wildcard is what makes phones open the setup page by themselves. Today the hotspot NATs clients to the internet and leaves an Ethernet LAN reachable, which is a security gap (07 §3). |

## System map

```
                         ┌──────────────── the printer (CM4, MuonOS) ────────────────────────┐
 phone / laptop          │                                                                   │
 on hotspot  ─ WPA2 ─▶  ap0 10.42.0.1 ─▶ dnsmasq (every name → 10.42.0.1)                     │
 (captive portal)        │                 nginx :80 ─┬─ Fluidd  (/setup route)  ◀─────┐      │
                         │                            └─ /server ─▶ Moonraker :7125 (lo)  │    │
 laptop on LAN ─────────▶ wlan0/eth0 ───▶ nginx :80 ──┘              │                   │    │
                         │                                  muon_setup ◀─ muon_link ─▶ muon-link :7131
 panel (Chromium kiosk)  │  nginx :100 (lo) ─▶ MuonUI ─▶ /server ─▶ │                   │    │
 + knob (sk_daemon)      │                                          ▼                   │    │
                         │                              aux_api_proxy ─▶ Aux API :6789 (lo)   │
                         │                                   /wifi/* /region/* /time* /setup/complete
                         │                                   ─▶ NetworkManager, regdomain, timedatectl
                         └───────────────────────────────────────────────────────────────────┘
 control.muon3d.com (https) ── can't reach the LAN ──▶ "What does the screen show?" → code → console /v1/links/*
 OrcaSlicer ── DNS-SD _octoprint._tcp (port 80, TXT name/setup) ──▶ Browse dialog
```

## Files

| File | For | What it covers |
|---|---|---|
| [01-flow.md](01-flow.md) | Everyone | The ways in, the steps and their rules, region and clock, the state model, the happy paths, surviving drops, after setup, migration, factory reset, timing targets |
| [02-setup-api.md](02-setup-api.md) | Moonraker, MuonUI, Fluidd | The `muon_setup` component: conventions, config, caller rules, persistence, every endpoint, the state document, the ready manifest, the `muon_link` bridge, tests |
| [03-printer-os.md](03-printer-os.md) | MuonOS, muon-link | Hotspot lifecycle, captive portal, Aux region, time, Wi-Fi and Enterprise changes, the link contract, the completion marker, Bluetooth (phase 2), bench measurements |
| [04-panel.md](04-panel.md) | MuonUI | The knob model, screens P1–P14, the ring keyboard, QR codes, tests |
| [05-phone-setup-page.md](05-phone-setup-page.md) | Fluidd | The `/setup` route, screens S0–S10, the region line, the client and its reconnect rules, captive-window constraints, i18n, the dashboard card, tests |
| [06-add-printer.md](06-add-printer.md) | Fluidd, OrcaSlicer | The routing table, `AddPrinterDialog`, control.muon3d.com, OrcaSlicer Bonjour and profile changes, the future app |
| [07-security.md](07-security.md) | Everyone | Rules S1–S12, the access summary, known gaps, privacy notes |
| [08-errors.md](08-errors.md) | UI authors | Every error code with its copy and actions |
| [09-testing.md](09-testing.md) | Everyone, QA | Automated suites by repo, the E2E mock, the bench matrix, acceptance checklists |
| [10-work-plan.md](10-work-plan.md) | Leads, agents | Rules for agents, packages MR/OS/ML/UI/FL/CON/OR/QA, order, phase 2, Jira mapping |
| [fixtures/](fixtures/) | UI authors | Example state, options and network payloads |

**Reading order by repo:**

- **Moonraker:** README → 01 → 02 → 07 → 08 → 10
- **MuonOS / Aux:** README → 01 → 03 → 07 → 10
- **muon-link:** 03 §6 → 02 §9
- **MuonUI:** README → 01 → 04 → 02 §5–6 → 08 → fixtures
- **Fluidd:** README → 01 → 05 → 06 §1–2 → 02 §5–6 → 08 → fixtures
- **OrcaSlicer:** 06 §3 → 10 (OR-*)

## Glossary

| Term | Meaning |
|---|---|
| Panel | The 480×480 round DSI display inside the haptic knob. It runs MuonUI. |
| Hotspot, `ap0` | The printer's own Wi-Fi network: `Muon-<word>-<hex4>`, per-device WPA2 key, `10.42.0.1/24`. It shares the radio with `wlan0`. |
| Uplink | The printer's connection to the owner's network (`wlan0` or `eth0`). |
| Aux API | MuonOS's privileged FastAPI service on `127.0.0.1:6789`, proxied by Moonraker at `/server/aux/*`. |
| `muon_setup` | The new Moonraker component that owns setup state (02). |
| Cursor | The step the owner is on. |
| Driver | The surface currently in charge: `panel`, `phone`, `web`, `app` or `bluetooth`. |
| `op` | A long-running operation: `region_apply`, `join`, `update_install`, `link` or `ready_item`. |
| `rev` | The state revision, used for optimistic concurrency between screens. |
| Market token | A signed file on config partition `p1` that lists the region configurations a unit may use (KAN-321 Rev 11). |
| Region line | "Region: United Kingdom · Change", confirmed by connecting. |
| CNA | The captive-portal window iOS and Android open by themselves. |
| Level 0 Open | SEC-1: the LAN and the hotspot are trusted, with no sign-in. |
| The floor | `muon_floor.FLOOR_PREFIXES`: endpoints refused to anything but loopback. |
| Link | Binding the printer to a Muon account: a code on the panel, a claim on control.muon3d.com, then a knob confirm. |
| Marker | `/var/lib/muon3d/setup/complete`, behind Aux `/setup/complete` (MuonOS KAN-413). It records that setup finished (KAN-203). |

## Verify before building

Each repo was checked against the code on 24 Sep (Moonraker `dc76b59`, Fluidd `9694dec`, OrcaSlicer, MuonOS `4d9f6e3`, MuonUI `f5ffa7c`, muon-link `bfcd43f`). What's still open:

1. **Draft PRs this spec builds on.** Several pieces exist only in unmerged PRs:
   - MuonOS#174 (region routes; its own setup routes are superseded by KAN-413);
   - MuonOS#210 (`/wifi/saved`);
   - MuonOS#305 (listeners);
   - MuonUI#31 (setup route, region picker);
   - MuonUI#39 (Wi-Fi overhaul);
   - Moonraker#20 (`muon_link`);
   - muon-link#24 (account link).

   Agents must coordinate with those PRs' authors, not fork them.
2. **No unit can declare a region yet.** There are no signing keys or tokens (KAN-321, KAN-132), so setup runs as market `none` until they exist.
3. **The account link has open questions:**
   - ~~the LINK-2 conflict over who mints the code~~ and ~~the code's limits in muon-console~~: settled 25 Sep (ADR 0026). The console mints it and limits failed claims;
   - the LINK-3 client-key comparison on the panel (KAN-415);
   - no proof of a knob press on `/link/confirm` (ML-2).

   See 03 §6.
4. **Ready to print.** The manifest content and the `MUON_SELF_TEST` macro belong to the hardware team (D6).
5. **Bluetooth (phase 2)** uses Iroh_BLE (decided 25 Sep, ADR 0027; AP-10 amended in MuonOS#317). It needs a `bluetooth` caller class through muon-link (03 §8), and Bluetooth SIG qualification.
6. **Security bugs found on the way** (07 §3):
   - Wi-Fi passwords in the persisted journal (OS-11, urgent);
   - hotspot clients reaching the internet and an Ethernet LAN (OS-1).
