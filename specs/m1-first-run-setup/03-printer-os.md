# 03 · Printer OS work (MuonOS, Aux API, muon-link)

**Repos:**

- **`Muon-3D/MuonOS`** (checked at `main` `4d9f6e3`). The Aux API lives in `recipes/aux_api/files/aux_api/` and runs as FastAPI on `127.0.0.1:6789`, with its token in `/run/aux_api/token`. Its contract is `contracts/muon.aux-api/v2/openapi.json`.
- **`Muon-3D/muon-link`.** Its admin endpoint is on `127.0.0.1:7131`.

Aux is published to Moonraker automatically: `aux_api_proxy` rebuilds `/server/aux/*` from `/openapi.json`. Any new route therefore appears there too.

**The floor has two lists that must match.** The `:80` Fluidd server returns 403 for the floored Aux paths (`recipes/klipper_moonraker_fluidd/files/fluidd/fluidd.nginx.template:122-136`), and a test requires that list to equal Moonraker's `FLOOR_PREFIXES`. Every path added to the floor must be added in **both** repos.

**Before you start:**

- The region and setup routes exist only in **draft [Muon-3D/MuonOS#174](https://github.com/Muon-3D/MuonOS/pull/174)** ("DO NOT MERGE (WIP): KAN-321"). They are not on `main`.
- No signing key is installed yet, so **every unit reports `no-signing-key`, and no unit can declare a country today.** Setup must work in that state ([01-flow.md §2.1](01-flow.md#21-region-inside-network)).
- Other open PRs this work depends on:
  - [#210](https://github.com/Muon-3D/MuonOS/pull/210): `GET /wifi/saved`, and replacing a saved password;
  - [#300](https://github.com/Muon-3D/MuonOS/pull/300): `security_enabled`;
  - [#305](https://github.com/Muon-3D/MuonOS/pull/305): `/etc/muon3d/listeners.d`;
  - [#249](https://github.com/Muon-3D/MuonOS/pull/249): a network-authority check on the Wi-Fi routes.

## 1. Hotspot lifecycle

**Today:**

- `ap0` comes from `00-ap0.rules`. Its profile is `ap0-con.nmconnection` (`autoconnect=false`, `band=bg`, channel 6, WPA-PSK, `ipv4 method=shared`). The profile directory is persisted, so booted units keep their own older copy; #174 migrates them.
- `muon-ap-provision` writes the 12-character key.
- `muon-ap-lifecycle.sh` (since KAN-341): the hotspot is **up unless `/home/printer_admin/ap-hotspot-disabled` exists**, and it's **forced up while `wlan0` isn't connected**. `ap-hotspot-requested` is a legacy file and is ignored. There is no `AP_DEFAULT_UP`.
- What triggers it: the dispatcher `20-ap-lifecycle`, and `muon-sta-watchdog.timer`, which does nothing while `wlan0` is connected.
- `POST /wifi/ap/count` returns the number of stations as a bare int, via pyroute2 `get_stations`.
- Measured: the hotspot follows the uplink, **including onto 5 GHz DFS channel 124**.

**The new policy.** Change `muon-ap-lifecycle.sh`:

| Rule | When | AP |
|---|---|---|
| H1 | Setup isn't complete: the marker `/var/lib/muon3d/setup/complete` (§7) is missing | **Up**, whatever the owner-choice file says |
| H2 | Setup is complete and `wlan0` isn't connected | **Up** (unchanged) |
| H3 | Setup is complete, an uplink (`wlan0` or `eth0`) is connected, the auto-off deadline has passed, and `/home/printer_admin/ap-hotspot-kept-on` is absent | **Down** |
| H4 | The owner turns it on (Settings › Add a phone or computer, or the hotspot card) | **Up**. This writes `ap-hotspot-kept-on`, a **new** file name, because pre-KAN-341 units may still have `ap-hotspot-requested`. Turning it off removes the file and writes `ap-hotspot-disabled`, as today. |

- **H3 changes KAN-341's default once setup is done.** It becomes "off while connected", which is the direction of MuonOS#193. Before setup, and after a factory reset, the default is on.
- **New route `POST /wifi/ap/auto_off`** with `{ "after_s": 900 }`.
  - Aux is unprivileged, so give it a sudo grant for a new helper, `/usr/libexec/muon3d/muon-ap-auto-off <seconds>`.
  - The helper starts a transient timer, `systemd-run --on-active=<s> --unit=muon-ap-auto-off …`, which re-runs the lifecycle when it fires.
  - It must be loopback-only. Add `/server/aux/wifi/ap/auto_off` to `FLOOR_PREFIXES` and to the `:80` nginx 403 list.
- **Station count.** `muon_setup` reads it from `POST /wifi/ap/count`. Keep that route; MuonUI's contract uses it.
- **Channel following.** Once the uplink associates, the hotspot moves to the uplink's channel, which can be 5 GHz or DFS. Phones that support 5 GHz follow it. Record in QA-1 how long the drop lasts.

## 2. Captive portal

**Goal:** a phone that joins the hotspot opens `/setup` by itself, with nothing to switch on or off at runtime.

**Today:**

- The hotspot **does** route: NetworkManager's shared mode NATs hotspot clients to the internet.
- `10-ap-isolate` rejects only traffic from `ap0` to the **`wlan0` subnet**, so a LAN reached over Ethernet isn't isolated.
- `muon3d-firewall.nft` is default-deny on input. It accepts tcp/80, tcp/22, udp/7127, udp/5353 and ICMP on every interface, plus udp/53, tcp/53 and udp/67 on `ap0`.
- Nothing is installed under `dnsmasq-shared.d`.

**Decision (README D11): the hotspot talks to the printer only.** It's a way to reach the printer, not a way to the internet or the LAN. That makes AP-7 true, and it lets the DNS wildcard below stay on permanently.

**OS-1 changes:**

1. **Isolation.** In `dispatcher.d/10-ap-isolate`, reject all forwarding from `ap0`: to the internet, to the `wlan0` subnet and to any `eth0` subnet. Run it on `eth0` events too.
   - Keep this out of `muon3d-firewall.nft`: `tests/test_firewall_ruleset.py` forbids forward hooks there, and its comment about "the hotspot's internet access" needs updating.
   - Hotspot clients lose internet access, which today is incidental.
2. **HTTPS fails fast.** Add `iifname "ap0" tcp dport 443 reject with tcp reset` to `muon3d-firewall.nft`, so HTTPS captive probes fail at once instead of timing out. A reject isn't a listener; check it against #305's rule that declared ports must equal allowed ports.
3. **Pin `10.42.0.1`** in `ap0-con` (`ipv4.addresses=10.42.0.1/24`). It's NetworkManager's shared-mode default today, not pinned. Pinning it needs the same migration of persisted profiles that #174 uses.
4. **DNS.** Add `/etc/NetworkManager/dnsmasq-shared.d/muon-captive.conf`:

   ```conf
   # Every name resolves to the printer; the hotspot routes nowhere else (D11).
   address=/#/10.42.0.1
   # No routable AAAA answers, so dual-stack phones fall back to IPv4.
   address=/#/::
   ```

5. **HTTP.** Add this to `fluidd.nginx.template`. It must pass `scripts/check-nginx-templates.sh` (`nginx -t`) in CI. Fluidd uses hash routing, so `/#/setup` is correct.

   ```nginx
   # http context
   map $host $muon_own_host {
       default                                        0;
       10.42.0.1                                      1;
       muon3d.local                                   1;
       "~*^muon-[a-z]{4,7}-[0-9a-f]{4}(\.local)?$"    1;
   }
   map "$server_addr:$muon_own_host" $muon_portal {
       default          0;
       "10.42.0.1:0"    1;
   }

   # inside the existing :80 server
   if ($muon_portal) { return 302 http://10.42.0.1/setup; }
   location = /setup { return 302 /#/setup; }
   ```

   **Why this opens the portal.** A request for a foreign host that arrives over the hotspot is an OS captive-portal probe, so it gets redirected: Apple `captive.apple.com`, Android `connectivitycheck.gstatic.com`, Windows `msftconnecttest.com`, Firefox `detectportal.firefox.com` and NetworkManager `nmcheck.gnome.org`. Every OS then opens the portal.

**No runtime toggling:**

- Before setup completes, the page opens the setup flow.
- After it completes, the same redirect opens the page in "set up already" or recovery mode ([05-phone-setup-page.md](05-phone-setup-page.md) S10).

**GATE-1.** The dnsmasq listeners on `10.42.0.1` are declared by #305. OS-1 adds no new listener.

**Rejected:** RFC 8910 / DHCP option 114. It needs an HTTPS API with a publicly valid certificate, which an offline printer can't have.

## 3. Region, time zone and clock

### Region: the draft #174 contract

| Route | Response |
|---|---|
| `GET /region` | `{reason, explanation, domain, declared_country, configuration, surroundings, detected_country, basis, enforcement, locked, channels}` |
| `GET /region/options` | `{countries, preselect, basis, locked}` |
| `POST /region/country` `{country}` | `{country, configuration, domain, verified}` |

**`GET /region` fields:**

- `surroundings` is one of `settled`, `offer`, `not-registered`, `locked` or `unknown`.
- `basis` is `joined-network` or `plurality`.
- `channels` lists the channels on which the applied configuration may start a transmission. It's empty when no configuration is applied.

**`GET /region/options` fields:**

- `countries` is a flat, sorted list of ISO codes.
- `preselect` is the detected country if the token allows it, otherwise the token's `default_country`.

**How `muon_setup` reads them:**

- **Market.** `countries == []` means `none` (no valid token). `locked == true` means `locked` (a US unit). Anything else is `picker`.
- **Suggestion.** `detected_country` with `basis`. `joined-network` works only **after** the printer has joined, because it reads the associated BSSID. `plurality` needs at least 3 access points naming a country, with that country ahead by at least 2.
- **No per-network suggestion before joining.** Nothing exposes each network's country element, so the flow confirms the region after the join (01 §2.1), as MuonUI#31 does. No `/region/suggest` route is needed.
- **Apply.** It takes down **every** active wireless profile, `ap0-con` included. It then applies the country, verifies the channel map, brings the profiles back, and only then records the declaration. It takes about 8 s.

**Where things live:**

| What | Path |
|---|---|
| Token | `/run/rugix/mounts/config/muon3d/region-token.json` (config partition; survives a reset) |
| Declared country | `/var/lib/muon3d/region/declared-country` |
| Status | `/run/muon3d/region-status.json` |
| Table | `/usr/share/muon3d/region/regions.json` |

**OS-2, changes needed on #174 before `muon_setup` can rely on it:**

1. **Stable error codes.** Return `detail: {code, message}`, as `/update/*` already does, using the region agent's `OUTCOMES` vocabulary:
   - `no-token`, `unreadable-token`, `bad-token-format`, `bad-signature`, `unknown-serial`, `serial-mismatch`, `no-signing-key`;
   - `country-not-in-token`;
   - `apply-failed`, `intersected`, `readback-mismatch`;
   - `busy`.

   Today a 409 carries the agent's last stderr line as free text.
2. **Timeouts and concurrency.** Cut `SET_COUNTRY_TIMEOUT_S` from 90 s to under 60 s, the Moonraker proxy's limit. Add a concurrency guard that returns `busy`.
3. **Channels per configuration.** Add each configuration's channels to `GET /region/options`, taken from `regions.json`'s `initiable_2g4` and `initiable_5g`. The surfaces can then say whether a network is reachable after a switch.
4. **Signing keys and tokens** (KAN-321 / KAN-132). Until they exist, every unit is `no-signing-key` and setup treats it as market `none`.

**Tier-2 picker data.** "Countries where this language is spoken" and the country names stay on the UI side: `SPOKEN_IN` and `Intl.DisplayNames` in MuonUI#31's `regionCountries.ts`, ported to Fluidd. `regions.json` has no language or continent data.

### Time (OS-6, new)

**Today:**

- There is no timezone handling.
- `systemd-timesyncd` runs with Debian defaults.
- `/etc/fake-hwclock.data` **isn't persisted**, so every boot starts from the image's baked time.
- There are no sudo grants for `timedatectl`, `date` or `fake-hwclock`.
- The clock-before-TLS precondition is **KAN-198**.

**Routes to add:**

| Route | Does |
|---|---|
| `GET /time` | `{ "epoch_ms", "ntp_synced", "tz" }` |
| `POST /time` `{ "epoch_ms" }` | Refuses with `409 ntp_synced` if `timedatectl show -p NTPSynchronized --value` is `yes`. Otherwise it runs `sudo -n timedatectl set-time @<s>`. |
| `POST /time/zone` `{ "tz" }` | Validates against `/usr/share/zoneinfo` and runs `sudo -n timedatectl set-timezone`. |

**Persisting the time zone.**

- Write it to `/var/lib/muon3d/setup/timezone`, and add a boot oneshot that applies it before `muon_setup` starts.
- Don't bind-mount `/etc/localtime`: `timedatectl` replaces that symlink with an atomic rename, which a bind mount breaks.
- Add sudoers entries for the exact commands, and a row in `docs/privacy/data-inventory.md`.

## 4. Wi-Fi join

**Today (`wifi_routes.py`):**

- `GET /wifi/scan?rescan=` returns `DeviceWifi {in_use, ssid, bssid, mode, chan, freq, rate, signal, security}`.
- `GET /wifi/device/status` returns `{device, device_type, state, connection, state_reason?, user_disconnected}`. `state_reason` is nmcli's raw `"N (text)"`.
- `GET /wifi/show?ssid=` is a query parameter, not a path segment.
- `POST /wifi/connect` takes `{ssid, password?}`. It blocks for up to 45 s, plus 10 s for rollback. Its results:

  | Response | Meaning |
  |---|---|
  | `200 {"status":"connecting","ssid"}` | Success. It's sent after the activation wait. |
  | `200 {"status":"restored","ssid":<previous>,"warning"}` | **The join failed and the previous connection was restored.** |
  | `400` | Free text `Could not connect to 'X': …`. Nothing was restored. |
  | `409` | More than 4 concurrent Wi-Fi operations |
  | `500` | The restore failed |

**How `muon_setup` drives a join (MR-3):**

1. Run `POST /wifi/connect` as a task, with a 60 s timeout. `aux_api_proxy`'s internal `post()` is fixed at 15 s, so add a timeout parameter to it.
2. At the same time, poll `GET /wifi/device/status` every 500 ms and map NetworkManager's state onto `op.phase`:
   - preparing or configuring → `associating`;
   - need-auth → `authenticating`;
   - ip-config or ip-check → `dhcp`;
   - activated → `internet_check`.
3. Treat **`status: "restored"`** and **400** as failures, and read the reason from `state_reason`'s numeric prefix. Check these values on a device against NetworkManager's `NMDeviceStateReason`:

   | `state_reason` | `code` |
   |---|---|
   | 7 (no secrets), or 8–11 (supplicant) on PSK | `wrong_password` |
   | 8–11 on 802.1X | `eap_failed` |
   | 53 (SSID not found) | `ssid_not_found` |
   | 5, 15–17 (IP config / DHCP) | `no_address` |
   | Anything else, or 45 s passing | `timeout` |

**Changes needed:**

| ID | Change |
|---|---|
| W1 | Replacing a saved network's password is **#210**. Hidden networks and a security hint are new work. They need `nmcli connection add`, which Aux has no sudo grant for today. |
| W2 | Return `detail: {code, message}` with the codes above instead of free text. |
| W3 | **Enterprise** (PEAP/TTLS). This needs a new privilege path: `nmcli connection add` or D-Bus/polkit. The sudoers header already says widening means D-Bus/polkit. It's the largest OS item, and it can ship later without blocking anything else (D10). |
| W4 | `POST /wifi/ca_cert` stores certificates in `/etc/NetworkManager/certs`. That directory **isn't persisted**, so it needs a `[[persist]]` declaration, an ID-9 row and a data-inventory row. |
| W5 | **The internet check.** The update check already talks to the Nexigon hub (`eu.nexigon.cloud`, via `nexigon-agent`, before any opt-in). Muon doesn't run that host, so it can't add a `generate_204` endpoint. Use the existing `check_connectivity()` (`update/ota_nexigon_client.py:392-414`) to decide `internet: true/false`. Report `internet: "portal"` when that check fails with a TLS or redirect error on an otherwise working uplink. `GET /wifi/uplink` returns `{kind, ssid, addresses, gateway, internet, checked_at}`, and adds no new destination. |
| W6 | `GET /wifi/saved` is **#210**. |

**Security bug, OS-11, independent of setup and urgent.**

- `POST /wifi/connect` runs `sudo nmcli … device wifi connect <ssid> password <psk>`.
- sudo logs `COMMAND=…` with the PSK to the **persisted** journal (`/var/log/journal`).
- The PSK is also visible in `/proc/<pid>/cmdline` while the command runs.
- `redaction.py` only cleans diagnostic bundles.

**Fix:**

- Stop passing the secret on the command line. Use nmcli's `passwd-file` with a 0600 file on tmpfs, or D-Bus.
- Add a sudoers `Defaults!<cmnd> !syslog` for the connect command.
- On update, vacuum the journal entries on units that already have them.

## 5. Ethernet

`GET /wifi/uplink` (W5) covers `eth0`. If Ethernet has an address when `network` becomes current, `muon_setup` offers "Connected by cable".

## 6. Account link (muon-link)

The account link is **muon-link PR #24** ([Muon-3D/muon-link#24](https://github.com/Muon-3D/muon-link/pull/24), open; NET-10(d), ADR 0018). It is not the `/pairing/*` routes, which are **client pairing** (8-digit code, SAS comparison, burn limits) and aren't used in setup phase 1.

| Admin route on `127.0.0.1:7131` | Caller | Does |
|---|---|---|
| `GET /link` | Moonraker `muon_link`; panel | Returns the current `LinkPhase` ([02-setup-api.md §5.9](02-setup-api.md#59-remote-access)) |
| `POST /link/start` | Moonraker `muon_link` | Sends `LinkStart` and moves to `connecting`. The orchestrator's `LinkCode` then moves it to `code`. |
| `POST /link/confirm` | **Panel only**, directly | Accepts the pending offer, which moves to `linked` |
| `POST /link/cancel` | Moonraker `muon_link`; panel | Declines a pending offer and drops the connection |
| `POST /link/unlink` | **Panel only** | Removes the link and every grant (LINK-8) |

Errors come back as `409 {"error": "<sentence>"}`. Nothing is pushed, so callers poll.

**OS-10 (new):**

1. Add `location /muon-link/ { proxy_pass http://127.0.0.1:7131/; }` to the **`:100` vhost only** (`recipes/muon-ui/files/muon-ui.nginx.template`). Never add it on `:80`.
2. Set `MUON_LINK_ORCH_ID` and `MUON_LINK_RELAY_URL` in `recipes/muon-link/files/muon-link.service`, and `MUON_LINK_ORCH_ADDR` if needed.
3. **Never set `MUON_LINK_DISCOVERABLE=1`** (NET-2, Tier 1).

**Open questions for the connectivity owner:**

- **LINK-2 vs PR #24.** LINK-2 has the printer mint an 8-character Crockford code with a lifetime and a 3-attempt cap. PR #24 has the orchestrator mint a digits-only code, and muon-link enforces nothing.
- **LINK-4 can't work with orchestrator-made codes.** There's no code without internet.
- **`/link/confirm` has no proof of a knob press** (ML-2).

## 7. Completion marker and factory reset

**Contract (decided 16:40): MuonOS branch `feat/KAN-413-setup-marker` (OS-7).** It replaces #174's `GET/POST /setup` for the marker.

| Route | Does |
|---|---|
| `GET /setup/complete` | Returns `{ "complete", "completed_at", "by" }` |
| `POST /setup/complete` `{ "by": "muon_setup" }` | Writes the marker. Idempotent: the first record is kept. |
| `DELETE /setup/complete` | Clears the marker. Used only by `muon_setup`'s development `reset`. |

- **The marker** is `/var/lib/muon3d/setup/complete`. Its existence is the fact, so a damaged file still reads as complete; the hotspot lifecycle checks it with `[[ -e ]]`.
- **Persisted** by `muon3d-setup.toml`. It survives OTA updates and rollback, and a factory reset removes it. KAN-413 adds the ID-9 and data-inventory rows.
- **`muon_setup` writes it** with `POST /setup/complete` at `finish`. It reads `GET /setup/complete` only during the migration check ([01-flow.md §7](01-flow.md#7-printers-already-in-the-field)). Once `muon_setup` has its own state, that state is authoritative.
- **The whole prefix `/server/aux/setup/complete` must be floored**, in **both** lists:
  - Moonraker's `FLOOR_PREFIXES`;
  - MuonOS `fluidd.nginx.template`'s 403 list, and `recipes/klipper_moonraker_fluidd/tests/test_floor.py`.

  Without it, a LAN client could `DELETE` the marker through `aux_api_proxy`.
- **#174 and MuonUI#31 overlap with this.**
  - #174 has its own `setup_routes.py` and `setup.toml` under the same `/setup` prefix. It must drop them, or rebase onto KAN-413.
  - MuonUI#31's `aux.setup.get()` and `aux.setup.complete()` calls go away (UI-1).
- **Time zone.** `/var/lib/muon3d/setup/timezone` (OS-6) sits in the same persisted directory.

**Factory reset** is `/usr/sbin/muon3d-factory-reset`, run as root from the console or SSH. It runs `rugix-ctrl state reset`, which clears everything persisted.

| What | After a reset |
|---|---|
| Market token (config partition) | Survives |
| Hostname and SSID | Re-derived from the serial, so the printer keeps the same name |
| Hotspot key | Regenerated, so the Wi-Fi QR code changes |
| Declared country and setup marker | Cleared (the marker once KAN-413 merges, the country once #174 merges) |

## 8. Phase 2: Bluetooth

**This contradicts MuonOS `docs/connectivity/SPEC.md:1077-1080`** (AP-10: Bluetooth "explicitly not built"). Phase 2 needs that SPEC amended first. The design stays as written, so phase 1 doesn't block it:

- **When it listens.** A `muon-setup-ble` daemon advertises only while setup is incomplete, or while "Add a phone or computer" is open.
- **What it offers.** One GATT service carrying `muon_setup`'s JSON-RPC methods.
- **Security.** SPAKE2, keyed with a 6-digit code shown on the panel, AES-GCM, a knob confirmation before the first write, and Wi-Fi secrets only over an encrypted session.
- **Plumbing.** It reaches Moonraker over a Unix socket checked with SO_PEERCRED. It's a GATE-1 listener, declared in `listeners.d`.
- **Prerequisite.** Coexistence tests: AP, station and BLE together.

## 9. Bench measurements (add to KAN-329)

| # | Measure | Why |
|---|---|---|
| B1 | How long `ap0` is gone during `POST /region/country`, and whether iOS and Android rejoin on their own | The drop when the region is confirmed |
| B2 | How long `ap0` is gone when `wlan0` joins on channels 1, 6, 11, 13, 36 and 124 (DFS), and whether each phone follows onto 5 GHz | The drop at join |
| B3 | Whether the iOS captive-portal window closes and reopens on B1/B2, and whether the page restores from state | How the phone path feels |
| B4 | Android captive behaviour after a Wi-Fi QR join, on Pixel and Samsung | Whether the URL QR code is needed |
| B5 | The AAAA answer from `address=/#/::` on the shipped dnsmasq | Dual-stack phones |
| B6 | QR scan to page visible, and Start to Connected, over 10 runs each on iOS and Android | [01-flow.md §9](01-flow.md#9-timing-targets) |
