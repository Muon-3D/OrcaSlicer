# 07 · Security and privacy

Setup runs before the owner has an identity on the printer, so the trust anchor is **physical presence**. Two things happen only at the panel: the hotspot key is shown there, and linking an account is confirmed there with the knob. Everything below follows from that.

This builds on what ships today:

- **SEC-1 Level 0 Open.** The LAN and the hotspot are trusted, with no sign-in.
- **SEC-2 the floor.** Loopback-only surfaces are listed in `muon_floor.FLOOR_PREFIXES`.
- **GATE-2.** muon-link callers arrive as the sentinel `192.0.2.1` with one-shot tokens.
- **KAN-103.** The hotspot has a per-device WPA2 key that is never returned over the network.
- **AP-7 and `10-ap-isolate`** say hotspot clients reach only the printer. **That wasn't true before OS-1:** NetworkManager's shared mode NATs them out, and `10-ap-isolate` only covers the `wlan0` subnet. OS-1 makes it true (S7). `ap0` still admits tcp/22, udp/5353 and udp/7127 on `10.42.0.1`; those are the printer's own services, which D11 allows (03 §2).

## 1. Rules

| # | Rule | Where it's enforced |
|---|---|---|
| S1 | The hotspot key and the Wi-Fi join QR appear **only on the panel**, and `muon_setup` state, events and logs never contain them. Aux returns the key from `GET /wifi/ap/show` only when the request carries `X-Muon-Local-UI: 1`. Only the loopback `:100` vhost sets that header, rewriting `/server/aux/wifi/ap/show` straight to Aux. Keep it that way: the `:80` server must never set or pass the header, and `aux_api_proxy` builds its own headers, so a LAN client can't smuggle it in. A phone app that reads the key from the panel's QR code follows S12. | MuonOS nginx templates, Aux `ap_routes.py`; `muon_setup` tests |
| S2 | Wi-Fi passwords and Enterprise passwords are sent once, straight to Aux. They are never stored by `muon_setup`, never echoed in a response, a state or an event, and never logged at any level by Moonraker, Aux or `nmcli_gate.py`. That includes Moonraker's verbose request logging: `/server/muon/setup/network` must be covered by its redaction. Only NetworkManager's root-only keyfiles hold them. | `muon_setup` + Aux W2 tests; MR-3 |
| S3 | Linking is confirmed **only at the panel**. MuonUI calls muon-link's `/link/confirm` directly, through the loopback-only `:100` vhost. Moonraker never forwards confirm or unlink, and PR #20's tests assert it. The panel must show the short form of the claiming client's own key, and the client shows the same short form (LINK-3). Today the console offers only the account email and the authority key, which doesn't satisfy LINK-3. KAN-415 builds it (ADR 0026). muon-link binds a confirm to the offer shown (`{account, fingerprint}`, muon-link#24), so a second claim can't slip another account under the owner's press. **Gap:** muon-link still accepts `/link/confirm` from any loopback process, with no proof of a knob press. ML-2 adds one. | MuonUI, muon-link, Moonraker PR #20 |
| S4 | The orchestrator makes link codes: 6 digits, valid for 600 s, and muon-link enforces no lifetime or attempt cap (PR #24). muon-console limits failed claims per account and per IP, and issues a fresh code when the panel declines a claim, so a stranger's claim can't lock the owner out. LINK-2 now says this (decided 25 Sep, Muon_Internal_Documentation ADR 0026, KAN-405). Moonraker's `link/start` is limited to 5 per minute per caller IP (MR-6). MuonOS must never set `MUON_LINK_DISCOVERABLE=1`. | muon-console, muon-link, `muon_link`, MuonOS |
| S5 | Remote callers (the Iroh gateway) can read setup state but can't write to it. Setup is local-only. | `muon_setup` caller check |
| S6 | Setup HTTP writes need `Content-Type: application/json`, a `Host` that is one of the printer's names or addresses, and, if an `Origin` is present, a matching one. Setup writes over websocket JSON-RPC check the upgrade request's `Host` against the same set, because Moonraker's websocket origin check passes a rebound name. This blocks form-based CSRF and DNS rebinding from websites visited on the owner's LAN. | `muon_setup` §3 |
| S7 | Hotspot clients reach **only the printer** (README D11): not the internet, not the LAN, and not each other. Today NetworkManager's shared mode NATs them to the internet, and `10-ap-isolate` only rejects traffic from `ap0` to the `wlan0` subnet, which leaves an Ethernet LAN reachable. OS-1 makes the printer route nothing (03 §2) and turns on AP client isolation. The captive DNS answers every name with `10.42.0.1`. | `muon3d-firewall.nft`, `ap0-con`, dnsmasq drop-in (OS-1) |
| S8 | Until the owner opts in at `remote`, the printer makes **no** connection to Muon account or link services (Tier 1, KAN-187). The internet check is a TLS handshake with the Nexigon hub the update check already uses (`eu.nexigon.cloud` in production, read from the image's agent config), so it adds no new destination. Region detection is local. | Aux W5, `muon_setup` |
| S9 | `POST /time` works only while NTP hasn't synced (refusing if it can't tell), and only accepts times between the image build and 20 years after it. `/server/aux/time` is floored, so only `muon_setup` reaches it, and `muon_setup` accepts `clock` only from a hotspot caller. At worst, a hotspot client can set a wrong clock until NTP corrects it, which breaks TLS but opens nothing. | Aux, the floor, `muon_setup` §3 |
| S10 | Enterprise networks without a CA certificate (`no_ca_check`) need an explicit choice labelled "not recommended". A CA certificate plus a domain is the recommended path. | `/setup` page S3b |
| S11 | `POST /server/muon/setup/reset` is panel-only and floored. It resets setup state only. Factory reset stays `/usr/sbin/muon3d-factory-reset` (console or SSH only, KAN-172). | floor + tests |
| S12 | This rule extends S1 to a phone app that joins the hotspot by scanning the panel's Wi-Fi QR code. The app keeps the key in memory only, and gives it only to the operating system's join call: `NEHotspotConfiguration` with `joinOnce` on iOS, and `WifiNetworkSpecifier` on Android. The app never stores the key (not in the Keychain, the Keystore, preferences, a cache or a crash report). It never logs the key, never sends it to the printer, the console or any other host, and never shows it. It drops its copy when the join call returns. If the join fails, the app tells the owner to join from the phone's Wi-Fi settings with the key from the panel (P2 **Show network details**). The Muon3D app states this as `APP-32`. | The Muon3D app: unit tests and the R17 bench run ([09-testing.md](09-testing.md)) |

## 2. Who can do what

The matrix per caller kind is in [02-setup-api.md §3](02-setup-api.md#3-who-is-calling). In short:

- The panel can do everything.
- The hotspot and the LAN can do everything except the ready-item actions, the link confirmation and `reset`.
- Remote callers can only read.

## 3. Known gaps, and what to do about them

| Gap | Effect | Action |
|---|---|---|
| **Wi-Fi passwords are written to the persisted journal (existing bug).** `POST /wifi/connect` runs `sudo nmcli … password <psk>`, and sudo logs the command line to `/var/log/journal`, which is persisted. The PSK is also visible in `/proc/<pid>/cmdline` while the command runs. | Anyone who can read the journal, or a diagnostic bundle that escapes `redaction.py`, sees the owner's Wi-Fi password. Units in the field already have these entries. | **OS-11, urgent and independent of setup.** Pass the secret through nmcli's `passwd-file` (0600, on tmpfs) or over D-Bus. Add sudoers `!syslog` for the command. Vacuum the existing entries on update. |
| **Hotspot clients are NATed to the internet and can reach an Ethernet LAN (existing).** | Anyone given the hotspot key gets onto the owner's wired network. | OS-1 (S7). |
| **Loopback is not physical presence** (SEC-2/3). Some remote paths still end on `127.0.0.1` and look like the panel. | A remote caller on such a path could act as the panel. That includes the ready step's `start`, which moves the printer. | GATE-2 moves muon-link to the sentinel. S3's knob backend closes the gap for linking. Track any other loopback-terminating path under SEC-2. |
| **The LAN path sends the Wi-Fi password over plain HTTP.** This happens when a computer on the LAN runs `/setup`, e.g. to change Wi-Fi. | Anyone sniffing the owner's LAN could read it. This is the same exposure as today's Fluidd Wi-Fi card under Level 0. | Accept for phase 1. The hotspot path is encrypted by WPA2. Revisit with HTTPS on the LAN if SEC-8 or KAN-187 add it. |
| **Level 0 lets any LAN device change the Wi-Fi after setup** through `/server/muon/setup/network` or `/server/aux/wifi/*`. | A LAN attacker could move the printer to another network. | Accepted at Level 0 (SEC-1). At Level 1 (SEC-8, Moonraker#21), `/server/aux/*` is protected, and `muon_setup` refuses post-setup writes from `lan` and `hotspot` callers (02 §3). `/server/muon/link/start` joins `PROTECTED_PREFIXES` (MR-6). |
| ~~**Remote (Iroh) callers can write `/server/aux/wifi/*`.**~~ **Closed:** muon-link's `ABSOLUTE_DENY` refuses every `/server/aux/` path to every remote session, whatever its role (`crates/muon-link-device/src/policy.rs`). FL-9 shows "Wi-Fi can only be changed on the printer's own network" over Iroh instead. | – | None. Keep it that way until the Aux API has authentication of its own. |
| **The hotspot key is 12 characters** from a 31-symbol alphabet (`muon-ap-provision.sh`), about 59.5 bits (KAN-326). | Offline brute force from a captured handshake is impractical at that strength, but it's just under the 60-bit target. | Accepted for phase 1. Go to 13 characters if the key format changes for another reason. |
| **A PSK QR code visible on the panel** can be photographed by anyone in the room. | They join the hotspot and can drive setup. | Accepted. Being in the room is the trust model. After setup, the hotspot goes off (H3), and linking still needs a knob press. |

## 4. Privacy notes for the published privacy posture (KAN-187)

- Setup adds no telemetry.
- The phone page makes no request to any host other than the printer.
- The printer sends nothing about setup anywhere. The only outbound call during setup is the internet and update check to the OTA host, and only after a network is joined.
- **Muon sees** the printer only after the owner picks **Link a Muon account** and confirms on the knob.
- **Any LAN or hotspot client can read `/server/muon/identity`:** the name, SSID, SSH fingerprint, setup state and Iroh EndpointId (KAN-403). That is the set MuonOS's `PRIV-6` already discloses. The EndpointId is a match key only: holding it must not locate the printer (`NET-4`), and no service, including the console's link flow, may accept it as proof of ownership or presence.
