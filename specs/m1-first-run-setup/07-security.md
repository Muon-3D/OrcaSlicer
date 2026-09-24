# 07 · Security and privacy

Setup runs before the owner has an identity on the printer, so the trust anchor is **physical presence**. Two things happen only at the panel: the hotspot key is shown there, and linking an account is confirmed there with the knob. Everything below follows from that.

This builds on what ships today:

- **SEC-1 Level 0 Open.** The LAN and the hotspot are trusted, with no sign-in.
- **SEC-2 the floor.** Loopback-only surfaces are listed in `muon_floor.FLOOR_PREFIXES`.
- **GATE-2.** muon-link callers arrive as the sentinel `192.0.2.1` with one-shot tokens.
- **KAN-103.** The hotspot has a per-device WPA2 key that is never returned over the network.
- **AP-7 and `10-ap-isolate`.** Hotspot clients can reach only the printer, on DHCP, DNS, `:80` and ICMP.

## 1. Rules

| # | Rule | Where it's enforced |
|---|---|---|
| S1 | The hotspot key and the Wi-Fi join QR appear **only on the panel**. `muon_setup` state, events and logs never contain them. `GET /wifi/ap/show` keeps returning `security_enabled` and never the key (KAN-376). | MuonUI; `muon_setup` tests |
| S2 | Wi-Fi passwords and Enterprise passwords are sent once, straight to Aux. They are never stored by `muon_setup`, never echoed in a response, a state or an event, and never logged at any level by Moonraker, Aux or `nmcli_gate.py`. Only NetworkManager's root-only keyfiles hold them. | `muon_setup` + Aux W2 tests |
| S3 | Linking is confirmed **only at the panel**. MuonUI calls muon-link's `/link/confirm` directly, through the loopback-only `:100` vhost. Moonraker never forwards confirm or unlink, and PR #20's tests assert it. The panel shows the authority key's short form (LINK-3). **Gap:** muon-link accepts `/link/confirm` from any loopback process, with no proof of a knob press. ML-2 adds one. | MuonUI, muon-link, Moonraker PR #20 |
| S4 | The orchestrator makes link codes, and muon-link enforces no lifetime or attempt cap (PR #24). The limits must be confirmed in muon-console, and the conflict with LINK-2 (printer-made code, 3 attempts) resolved ([03-printer-os.md §6](03-printer-os.md#6-account-link-muon-link)). Moonraker's `link/start` is limited to 5 per minute per caller IP (MR-6). MuonOS must never set `MUON_LINK_DISCOVERABLE=1`. | muon-console, muon-link, `muon_link`, MuonOS |
| S5 | Remote callers (the Iroh gateway) can read setup state but can't write to it. Setup is local-only. | `muon_setup` caller check |
| S6 | Setup HTTP writes need `Content-Type: application/json`, a `Host` that is one of the printer's names or addresses, and, if an `Origin` is present, a matching one. This blocks form-based CSRF and DNS rebinding from websites visited on the owner's LAN. | `muon_setup` §3 |
| S7 | Hotspot clients never get routed anywhere: no LAN, no internet. The captive DNS answers everything with `10.42.0.1`, and nothing is forwarded. Confirm AP-7 on hardware (it's fixed in code, not yet verified on a unit). | `10-ap-isolate`, dnsmasq drop-in |
| S8 | Until the owner opts in at `remote`, the printer makes **no** connection to Muon account or link services (Tier 1, KAN-187). The internet check and the update check use only the OTA host already listed in the privacy inventory. Region detection is local. | Aux W5, `muon_setup` |
| S9 | `POST /time` works only while NTP hasn't synced, and only accepts times later than the image build. At worst, a hotspot client can set a wrong clock until NTP corrects it, which breaks TLS but opens nothing. | Aux |
| S10 | Enterprise networks without a CA certificate (`no_ca_check`) need an explicit choice labelled "not recommended". A CA certificate plus a domain is the recommended path. | `/setup` page S3b |
| S11 | `POST /server/muon/setup/reset` is panel-only and floored. It resets setup state only. Factory reset stays `/usr/sbin/muon3d-factory-reset` (console or SSH only, KAN-172). | floor + tests |

## 2. Who can do what

The matrix per caller kind is in [02-setup-api.md §3](02-setup-api.md#3-who-is-calling). In short:

- The panel can do everything.
- The hotspot and the LAN can do everything except the ready-item actions, the link confirmation and `reset`.
- Remote callers can only read.

## 3. Known gaps, and what to do about them

| Gap | Effect | Action |
|---|---|---|
| **Loopback is not physical presence** (SEC-2/3). Some remote paths still end on `127.0.0.1` and look like the panel. | A remote caller on such a path could act as the panel. | GATE-2 moves muon-link to the sentinel. S3's knob backend closes the gap for linking. Track any other loopback-terminating path under SEC-2. |
| **The LAN path sends the Wi-Fi password over plain HTTP.** This happens when a computer on the LAN runs `/setup`, e.g. to change Wi-Fi. | Anyone sniffing the owner's LAN could read it. This is the same exposure as today's Fluidd Wi-Fi card under Level 0. | Accept for phase 1. The hotspot path is encrypted by WPA2. Revisit with HTTPS on the LAN if SEC-8 or KAN-187 add it. |
| **Level 0 lets any LAN device change the Wi-Fi after setup** through `/server/muon/setup/network` or `/server/aux/wifi/*`. | A LAN attacker could move the printer to another network. | Add `/server/muon/setup/*` writes and `/server/muon/link/start` to SEC-8 Level 1's protected set, alongside `/server/aux/*`. |
| **The hotspot key is 12 characters**, and the character set isn't specified (KAN-326). | If it isn't drawn from a large alphabet, it may be brute-forceable offline from a captured handshake. | Confirm `muon-ap-provision` uses at least 60 bits of entropy (for example, 12 characters from a 32+ character alphabet). |
| **A PSK QR code visible on the panel** can be photographed by anyone in the room. | They join the hotspot and can drive setup. | Accepted. Being in the room is the trust model. After setup, the hotspot goes off (H3), and linking still needs a knob press. |

## 4. Privacy notes for the published privacy posture (KAN-187)

- Setup adds no telemetry.
- The phone page makes no request to any host other than the printer.
- The printer sends nothing about setup anywhere. The only outbound call during setup is the internet and update check to the OTA host, and only after a network is joined.
- **Muon sees** the printer only after the owner picks **Link a Muon account** and confirms on the knob.
