# 03 · Printer OS work (MuonOS, Aux API, muon-link)

**Repos:** `Muon-3D/MuonOS` (the Aux API lives at `recipes/aux_api/files/aux_api/`: FastAPI on `127.0.0.1:6789`, token in `/run/aux_api/token`), and muon-link (pairing admin endpoint on `127.0.0.1:7131`).

Neither repo was available when this spec was written. Routes marked **exists** come from Jira and the Moonraker proxy tests. Routes marked **new** are what this spec needs. Before adding a new route, check whether MuonOS already has an equivalent, and prefer extending it.

The Aux API is published to Moonraker automatically: `aux_api_proxy` rebuilds `/server/aux/*` from `/openapi.json` at startup. Any new route therefore appears there as well. Each new route must say whether it may be reached from the network. Anything that mustn't be goes into `muon_floor.FLOOR_PREFIXES` as `/server/aux/<path>`.

## 1. Hotspot lifecycle

**What exists today:**

- `ap0` is created by the udev rule `00-ap0.rules`.
- The profile is `/etc/NetworkManager/system-connections/ap0-con.nmconnection`, Rugix-persisted.
- The per-device 12-character WPA2 key comes from `muon-ap-provision` (KAN-103). The SSID `Muon-<word>-<hex4>` matches the hostname (KAN-357).
- `muon-ap-lifecycle.sh` enforces **AP-4**: the AP is up whenever the station isn't connected. A timer re-checks it (KAN-235).
- The default is on (`AP_DEFAULT_UP`, KAN-341). The owner's choice is stored in `ap-hotspot-requested` and `ap-hotspot-disabled` under `/home`.
- The panel refuses to switch the AP off while it's the only way in (KAN-347).
- Firewall `10-ap-isolate` allows only DHCP, DNS, `:80` and ICMP from `ap0`, and drops forwarding. Hotspot clients can't reach the LAN or the internet (AP-7).

**New policy.** Update `muon-ap-lifecycle.sh` and the dispatcher:

| Rule | When | AP |
|---|---|---|
| H1 | The setup marker (§7) is absent | **Up**, whatever the owner-choice files say. On a fresh or reset unit those files don't exist anyway. |
| H2 | The marker is present and the station isn't connected | **Up** (AP-4, unchanged) |
| H3 | The marker is present, the station is connected, and `ap-hotspot-requested` is absent | **Down**, once the auto-off deadline has passed |
| H4 | The owner turns it on (Settings › Add a phone or computer, or the Fluidd hotspot card) | **Up**. This writes `ap-hotspot-requested`. The owner turns it off the same way. |

- **H3 makes the default after setup "off when connected".** That is the direction MuonOS#193 proposes. Before setup, the default stays on.
- **New route `POST /wifi/ap/auto_off`**, body `{ "after_s": 900 }`:
  - It writes a deadline to `/run/muon3d/ap-auto-off` (tmpfs) and wakes the lifecycle check.
  - At the deadline the check applies H3. If the station isn't connected by then, H2 keeps the AP up.
  - Loopback-only: add `/server/aux/wifi/ap/auto_off` to the floor.
- **New route `GET /wifi/ap/stations`** returns `{ "up": true, "count": 1 }`.
  - The count comes from `iw dev ap0 station dump`.
  - Don't return MAC addresses.
  - This replaces `muon_setup`'s use of the ambiguous `POST /wifi/ap/count`.
- **Band pin (KAN-326).** The hotspot must be pinned to 2.4 GHz channels 1–11 (`band=bg`) *before* the station associates. That keeps it visible under the world domain `00`. Delivering the pin to fielded units (KAN-326) is a prerequisite.
- **Channel following.** `ap0` and `wlan0` share one radio, so after association the AP moves to the uplink's channel. If the uplink is on a 5 GHz DFS channel, the AP may not be able to beacon. Measure this (§9). If the AP can't come back, `GET /wifi/ap/stations` reports `up: false`, and the panel shows the result and the printer's LAN address. The phone page has already warned about this.

## 2. Captive portal

The design goal is that a phone joining the hotspot opens `/setup` by itself, and nothing needs to be switched on and off at runtime.

**DNS.** NetworkManager's shared-mode dnsmasq already serves `ap0` on `10.42.0.1`, and it is declared in GATE-1 (`/etc/muon3d/listeners.d`). Add the drop-in `/etc/NetworkManager/dnsmasq-shared.d/muon-captive.conf`:

```conf
# Every name resolves to the printer. The hotspot never routes to the internet (AP-7),
# so this is safe to leave on permanently.
address=/#/10.42.0.1
# No AAAA answers, so dual-stack phones fall back to IPv4.
address=/#/::
```

Check the AAAA behaviour on the image with `dig AAAA captive.apple.com @10.42.0.1`. The expected answer is NODATA or `::`, and never a routable address.

**HTTP.** Add this to the nginx server on `:80`, the one that serves Fluidd:

```nginx
# Requests that reached the printer through the hotspot, for a host name that isn't the printer's own,
# are OS captive-portal probes or stray traffic. Send them to the setup page.
map $host $muon_own_host {
    default                                 0;
    10.42.0.1                               1;
    muon3d.local                            1;
    ~*^muon-[a-z]{4,7}-[0-9a-f]{4}(\.local)?$  1;
}

server {
    # ... existing Fluidd server on :80 ...
    set $muon_portal "";
    if ($server_addr = 10.42.0.1) { set $muon_portal "hotspot"; }
    if ($muon_own_host = 0)       { set $muon_portal "${muon_portal}-foreign"; }
    if ($muon_portal = "hotspot-foreign") {
        add_header Cache-Control "no-store" always;
        return 302 http://10.42.0.1/setup;
    }

    location = /setup { return 302 /#/setup; }
}
```

- **Why this triggers the portal.** Every OS's probe then fails in a way that shows a portal:
  - Apple (`captive.apple.com/hotspot-detect.html`)
  - Android and ChromeOS (`connectivitycheck.gstatic.com/generate_204` and similar)
  - Windows (`www.msftconnecttest.com/connecttest.txt`)
  - Firefox (`detectportal.firefox.com/success.txt`)
  - NetworkManager (`nmcheck.gnome.org`)
- **Port 443 stays closed** on `ap0` (`10-ap-isolate`). HTTPS probes therefore fail fast and the plain-HTTP result decides.
- **No new listeners** for GATE-1. **No runtime toggling:**
  - Before setup, the redirect lands on the setup flow.
  - After setup, it lands on the same page in "set up already" or recovery mode ([05-phone-setup-page.md](05-phone-setup-page.md) S10).
- **RFC 8910 / DHCP option 114 is rejected.** The captive-portal API it points to must be served over HTTPS with a publicly valid certificate, which an offline printer can't have.

## 3. Region, time zone and clock

**Region: exists (KAN-321 Rev 11, "Built").** The routes are `GET /region`, `GET /region/options` and `POST /region/country`.

- `muon_setup` needs these fields from them. Add any that are missing:

| Need | Field | Where |
|---|---|---|
| Market | `picker` / `locked` / `none`, derived from the token's `configs` (`["us"]` → locked; no valid token → none) | `GET /region/options` |
| Offered countries | Grouped by continent, from the token plus `regions.json` | `GET /region/options` |
| Tier 2 order | Countries for a language, by speakers (German → DE, AT, CH) | `GET /region/options?language=de` (**new param**) |
| Token default | `default_country` | `GET /region/options` |
| Applied state | `country`, `config`, `declared`, `source`, `applied_at` | `GET /region` |
| Permitted channels | Channels the applied configuration permits | `GET /region` (Rev 11 says it already returns these) |
| Per-configuration channels | To check a network's channel *before* applying | `GET /region/options` (**new**: `channels` per config) |
| Support code | Shown when the token is missing or invalid | `GET /region` (**new** if absent) |

- **New route `GET /region/suggest?ssid=<ssid>`** returns `{ "country": "GB" | null, "source": "ap" | "neighbours" | "default" | null }`.
  - It uses the latest scan's country elements, following Rev 11's order: the AP's own element first, then the neighbours' plurality, then the token default. It's filtered through the token.
  - If the scan model can carry each BSS's country element, extend `GET /wifi/scan` instead, and let `muon_setup` apply the same rule.
- **Apply (`POST /region/country`).**
  - It is live: it takes every Wi-Fi link down, **`ap0` included**, applies the country, verifies it against the channel-map fingerprint, brings back what it took down, and only then records the declaration. It takes about 8 s.
  - `muon_setup` needs **stable error codes** in `detail.code`: `busy`, `not_offered`, `apply_failed`, `no_token`.
  - Loopback-only? No. Level 0 lets the LAN and the hotspot change the region, the same as today's Wi-Fi card. Revisit under SEC-8.

**Time: new.** Privileged, behind the `aux_api` sudo boundary.

| Route | Does |
|---|---|
| `GET /time` | `{ "epoch_ms", "ntp_synced", "tz" }` |
| `POST /time` `{ "epoch_ms" }` | Refuses with `409 ntp_synced` when `timedatectl show -p NTPSynchronized --value` is `yes`. Otherwise it runs `timedatectl set-time @<s>` (or `date -s`), then `fake-hwclock save`. |
| `POST /time/zone` `{ "tz" }` | Validates `tz` against `/usr/share/zoneinfo`, then runs `timedatectl set-timezone`. |

- **Persisting the time zone.** The root overlay is volatile, so the zone won't survive a reboot on its own.
  - Either add `/etc/localtime` and `/etc/timezone` to Rugix `[[persist]]`, or keep the zone in `/var/lib/muon3d/setup/timezone` and have a boot oneshot apply it.
  - Pick whichever MuonOS already does for similar files (see KAN-95), and add it to the ID-9 state inventory.
- **Reachable from the hotspot and LAN.** The setup page posts the time from the phone. `POST /time` is harmless once NTP has synced, because it refuses then.

## 4. Wi-Fi join

**Exists:**

- `GET /wifi/scan?rescan=1` returns `DeviceWifi {in_use, ssid, bssid, mode, chan, freq, rate, signal, security}`.
- `POST /wifi/connect` takes `{ssid, password}`. It waits up to 45 s, with a 10 s restore on failure (KAN-339).
- `GET /wifi/device/status` returns `{device, device_type, state, connection, state_reason?}`.
- Also `/wifi/current`, `/wifi/show`, `/wifi/forget` and `/wifi/up`.

**How `muon_setup` drives a join:**

1. Start `POST /wifi/connect` as a background task, with a 60 s timeout to match the proxy.
2. At the same time, poll `GET /wifi/device/status` every 500 ms and map NetworkManager device states to `op.phase`:

   | NM state | `op.phase` |
   |---|---|
   | `prepare`, `config` | `associating` |
   | `need-auth` | `authenticating` |
   | `ip-config`, `ip-check` | `dhcp` |
   | `activated` | address obtained → `internet_check` |

3. On failure, map `detail.code` (W2 below) or the NM `state_reason`:

   | NM `state_reason` | `code` |
   |---|---|
   | `no-secrets`, `supplicant-disconnect`, `supplicant-failed`, `supplicant-timeout` on PSK | `wrong_password` |
   | The same on 802.1X | `eap_failed` |
   | `ssid-not-found` | `ssid_not_found` |
   | `ip-config-unavailable`, `dhcp-failed`, `dhcp-error` | `no_address` |
   | Anything else after 45 s | `timeout` |

**Aux changes:**

| ID | Change |
|---|---|
| W1 | `POST /wifi/connect` accepts `hidden: bool` and `security`. It *replaces* the secret of an existing profile with the same SSID, fixing MuonOS#210 (no way to re-enter a password). |
| W2 | Failures return `detail: {"code": "<one of the §4 codes>", "message": "…"}`. On failure the new profile is deleted and the previous connection restored (as today). Secrets are never logged: audit `nmcli` argument logging and `nmcli_gate.py`. |
| W3 | **Enterprise.** `POST /wifi/connect` accepts `eap: {method: "peap"\|"ttls", phase2: "mschapv2"\|"pap", identity, anonymous_identity?, password, ca_cert_id?, domain_suffix_match?, no_ca_check?}`. It writes `802-1x.*` into a root-only keyfile, with the password stored in the keyfile (`password-flags=0`) so it works headless. |
| W4 | `POST /wifi/ca_cert` takes a multipart file (≤ 16 KiB, PEM or DER). It is validated with `openssl x509`, stored as `/etc/NetworkManager/certs/<id>.pem` (0600 root, a persisted path), and returns `{ "id" }`. |
| W5 | `GET /wifi/uplink` returns `{ "kind": "wifi"\|"ethernet"\|"none", "ssid", "addresses": [...], "gateway", "internet": true\|false\|"portal", "checked_at" }`. The internet check runs on every transition to connected and on demand (`?recheck=1`). |
| W6 | `GET /wifi/saved` returns the saved SSIDs in one call (proposed in KAN-339). It replaces the N+1 `GET /wifi/show/<ssid>`. |

**Internet check (W5).**

- **No new destinations.**
  - Use the host that the OTA update check already contacts.
  - Add a `generate_204` path to it, or use an equivalent path that returns 204.
  - Record it in the privacy inventory (MuonOS#174).
- **Results:**
  - `204` → `internet: true`.
  - Any redirect, or a `200` with a body → `"portal"` (a guest network that needs a web sign-in).
  - DNS failure or a timeout (5 s) → `false`.
- **Never contact a Muon account or link service before the owner opts in** (the Tier 1 promise, KAN-187).

## 5. Ethernet

`GET /wifi/uplink` (W5) covers `eth0`. If Ethernet has an address when `network` becomes current, `muon_setup` offers "Connected by cable". No Aux change is needed beyond W5.

## 6. Link service contract

**muon-link exists** (muon-link#14, `86809b3`). Its admin endpoint `127.0.0.1:7131` serves `/pairing/open`, `/pairing/confirm` and `/pairing/cancel` (KAN-190).

**What the setup flow needs.** Confirm each item in muon-link, and add what's missing:

1. **Open** returns a 6-digit code with a 120 s TTL. The code is single-use and allows 3 attempts before a knob press re-arms it (KAN-190).
2. **Status** can be read, with the phases `unlinked`, `code`, `offer`, `linked` and `failed`. In `offer` it includes the claimant's display label (account email or name), which the panel shows at the confirm. If there's no status route, add `GET /pairing/status`.
3. **Confirm** is accepted only from Moonraker's `muon_link` component, which is itself panel-only and floored ([02-setup-api.md §9](02-setup-api.md#9-link-endpoints-fluidd-already-calls-new-component-muon_link)).
4. **Cancel** closes the window and invalidates the code.

The console side (`/v1/links/claim`, `/v1/links/:id`) is unchanged, except for CON-1 ([10-work-plan.md](10-work-plan.md)).

## 7. Completion marker and factory reset

- **Marker:** `/var/lib/muon3d/setup/complete`, a JSON file: `{ "version": 1, "completed_at": "<ISO 8601>", "by": "muon_setup" }`.
  - MuonOS#174 recorded `/var/lib/muon3d/setup` as a persisted path. Confirm it is in Rugix `[[persist]]` and in the ID-9 state inventory.
  - **New routes:**
    - `GET /setup/complete` returns `{ "complete": bool, "completed_at" }`.
    - `POST /setup/complete` writes the marker. It is loopback-only, so add `/server/aux/setup/complete` to the floor.
    - `DELETE /setup/complete` is also loopback-only and floored. It's for `reset`.
- **Readers:** the hotspot lifecycle (H1/H3) and `muon_setup`'s migration check ([01-flow.md §7](01-flow.md#7-printers-already-in-the-field)).
- **Factory reset** (`/usr/sbin/muon3d-factory-reset` → `rugix-ctrl state reset`, KAN-172):
  - It clears the marker, the Moonraker database (the setup state and the name), saved networks and the certs, the declared country (ADR 0005), the Iroh identity, and the owner-choice files.
  - The market token on `p1` survives.
  - No change is needed.
  - KAN-351's human-only reset check should also confirm that the printer boots into setup at `language` afterwards.

## 8. Phase 2: Bluetooth

This isn't built in phase 1. It is written down so phase 1 doesn't block it.

- **Radio.** The CM4's CYW43455 carries Wi-Fi and Bluetooth on one radio, managed by BlueZ.
- **Daemon.** A new `muon-setup-ble` daemon **advertises only** while the marker is absent, or while the panel's "Add a phone or computer" screen is open. The rest of the time it isn't listening, which keeps it out of the GATE-1 census; declare it in `/etc/muon3d/listeners.d` for the times it is.
- **GATT.** One service with three characteristics:
  - `request` (write): JSON-RPC, `{ "id", "method": "server.muon.setup.network", "params": {…} }`, chunked;
  - `response` (notify);
  - `state` (notify): `notify_muon_setup_changed` bodies.
  The methods are exactly the `muon_setup` RPC names. There is no second API.
- **Talking to Moonraker.** The daemon calls Moonraker over a Unix socket, checked with SO_PEERCRED like `muon_gateway`. `muon_setup` classifies it as caller kind `bluetooth`, with the same rights as `hotspot`.
- **Security.**
  - A SPAKE2 exchange keyed with a 6-digit code shown on the panel.
  - AES-GCM session encryption for every message after the key exchange.
  - A knob confirmation on the panel before the first write.
  - Wi-Fi secrets are accepted only on an encrypted session.
  - Improv Wi-Fi's "press to authorise" is the prior art. Its message set can't carry a region, so adapt it rather than adopt it.
- **Clients.** Web Bluetooth in Chrome and Edge (desktop and Android) and the future app. There is no iOS Safari or Firefox support, which is why the hotspot path comes first.
- **Bench.** Test Bluetooth and Wi-Fi coexistence (AP + STA + BLE) before enabling it.

## 9. Bench measurements (add to KAN-329)

| # | Measure | Why |
|---|---|---|
| B1 | How long `ap0` is gone during `POST /region/country`, and whether iOS and Android rejoin by themselves afterwards | The phone path's first drop |
| B2 | How long `ap0` is gone when `wlan0` joins on channels 1, 6, 11, 13, 36 and 100 (DFS) | The second drop. Does the AP come back on DFS channels? |
| B3 | Whether the iOS captive-portal window closes and reopens on B1 and B2, and whether the reopened page restores from state | The phone path's feel |
| B4 | Android captive behaviour after a Wi-Fi QR join on Pixel and Samsung (latest two OS versions): opens by itself, or a notification? | Whether the URL QR code is needed |
| B5 | The `address=/#/::` AAAA answer on the shipped dnsmasq | So dual-stack phones don't stall |
| B6 | Time from QR scan to page visible, and from Start to Connected (10 runs, iOS and Android) | [01-flow.md §9](01-flow.md#9-timing-targets) targets |
