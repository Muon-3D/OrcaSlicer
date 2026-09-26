# 03 · Printer OS work (MuonOS, Aux API, muon-link)

**Repos:**

- **`Muon-3D/MuonOS`** (checked at `main` `4d9f6e3`; OS-5, OS-6 and OS-7 merged on 25–26 Sep, and `main` was re-audited at `972dc702`, so the **Today** notes in §1, §3 and §7 describe the tree before those packages). The Aux API lives in `recipes/aux_api/files/aux_api/` and runs as FastAPI on `127.0.0.1:6789`, with its token in `/run/aux_api/token`. Its contract is `contracts/muon.aux-api/v2/openapi.json`.
- **`Muon-3D/muon-link`.** Its admin endpoint is on `127.0.0.1:7131`.

Aux is published to Moonraker automatically: `aux_api_proxy` rebuilds `/server/aux/*` from `/openapi.json`. Any new route therefore appears there too.

**The floor has two lists that must match.** The `:80` Fluidd server returns 403 for the floored Aux paths (`recipes/klipper_moonraker_fluidd/files/fluidd/fluidd.nginx.template:122-136`). `test_trusted_clients.py` requires `EXPECTED_FLOOR` to equal the pinned Moonraker's `FLOOR_PREFIXES` exactly and in order, and requires a `location ^~ … { return 403; }` for each entry. Every path added to the floor is added in **both** repos. The MuonOS half lands in the PR that bumps the Moonraker pin past the new entries, together with rows in `test_floor.py`'s `FLOOR_CASES` ([02 §1](02-setup-api.md#1-conventions-this-component-follows)).

**A floored route never reaches `main` before its floor.** A MuonOS PR that adds an Aux route which must be floored merges only after the pin bump that floors it, or in the same PR. This applies to every such route, not only OS-7's. OS-6 merged 1 h 44 min before `/server/aux/time` was floored (26 Sep), and in that window `latest` builds let any LAN client set the clock.

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

An **uplink** is `wlan0` **or** `eth0` connected. The **deadline** is the auto-off time in `/run/muon3d/ap/auto-off`.

| Rule | When | AP |
|---|---|---|
| H1 | Setup isn't complete: the marker `/var/lib/muon3d/setup/complete` (§7) is missing | **Up**, whatever the owner-choice files say. Also remove any deadline, so a stale one can't take the hotspot down when setup finishes again in the same boot. |
| H2 | Setup is complete and there is no uplink | **Up** |
| H3 | Setup is complete, there is an uplink, no deadline is pending, and `/home/printer_admin/ap-hotspot-kept-on` is absent | **Down** |
| H4 | The owner turns it on (Settings › Add a phone or computer, or the hotspot card) | **Up**. This writes `ap-hotspot-kept-on` and removes `ap-hotspot-disabled`. Turning it off removes `ap-hotspot-kept-on`, writes `ap-hotspot-disabled` and ends any pending deadline. |

- **Use `ap-hotspot-kept-on`, never `ap-hotspot-requested`.** Firmware from MuonOS#37 to #212 (9–14 Sep) wrote `ap-hotspot-requested` from `POST /wifi/ap/up`, and KAN-341 left those files behind. Reading that name would keep those printers' hotspots on for good. It stays ignored.
- **H1 needs a way to finish setup in the same image.** In an image that has H1 but neither the panel's setup (UI-1) nor the phone page's screens (FL-3), a unit stored as `new` keeps its hotspot up for good, and the owner's "off" is ignored. That covers fresh and factory-reset units, and any field unit that migration misses. Don't promote OS-5 to beta or stable without one of them (OS-9). On a `latest` dev unit, `POST /server/muon/setup/language` followed by `POST /server/muon/setup/finish` from the LAN completes setup.
- **H3 changes KAN-341's default once setup is done.** It becomes "off while connected", which is the direction of MuonOS#193. Before setup, and after a factory reset, the default is on. Reversing KAN-341's AP-4 needs its owner's agreement (Jack).
- **Deadlines.** A deadline is pending after `finish` (`auto_off`, 900 s) and for 15 minutes after each boot, so an owner can always reach a freshly booted printer. **`ap-hotspot-disabled` skips the boot grace**, so an owner's "off" survives a reboot while there's an uplink.
- **New route `POST /wifi/ap/auto_off`** with `{ "after_s": 900 }` returns `{ "after_s", "auto_off_at" }`.
  - Aux writes the deadline, as seconds since boot, to `/run/muon3d/ap/auto-off` (tmpfs, `0750 aux_api` through tmpfiles). No sudo grant is needed. Seconds since boot, not wall-clock time, because the clock may be stepped by NTP or by the phone.
  - A path unit re-runs the lifecycle when the file changes, and the lifecycle arms a transient timer (`systemd-run --on-active=<s>`) for the deadline.
  - It must be loopback-only. Moonraker#25 adds `/server/aux/wifi/ap/auto_off` to `FLOOR_PREFIXES`; the MuonOS pin bump adds it to the `:80` nginx 403 list and `EXPECTED_FLOOR`.
- **What runs the lifecycle:** the dispatcher `20-ap-lifecycle` on `wlan0` **and `eth0`** events, the path unit, the deadline timer, and `muon-sta-watchdog.timer`.
- **No lost triggers.** The lifecycle service is a oneshot, and systemd folds a path trigger that arrives during a run into that run. So each pass snapshots its inputs (the marker, the deadline file, the owner-choice files and the uplink state) **before** it decides, and at the end compares them with what it reads then. If any differ, it runs again. A snapshot taken after deciding misses a change that lands while the script decides. The deadline the script writes itself doesn't count as a change (or costs one extra quiet pass). Test it by changing an input from inside the fake `nmcli device` call and from inside the fake `systemd-run`. The result must not depend on the order in which the marker and the deadline are written. (`muon_setup` writes the marker first anyway: 02 §5.11.)
- **Changing Wi-Fi after setup** (E5, S10): after any successful join once `state` is `complete`, `muon_setup` calls `auto_off` again, so the hotspot stays up long enough for the phone to see the result.
- **Station count.** `GET /wifi/ap/stations` returns `{ "up", "count", "auto_off_at" }`, where `auto_off_at` is a Unix time computed from the time left, or `null`. `muon_setup` reads it for `state.hotspot`, including `hotspot.auto_off_at`, which only the OS knows after a reboot or an owner's "off". `POST /wifi/ap/count` stays for MuonUI's existing contract.
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

1. **Isolation: the printer routes nothing.** The rule is defined by its properties, not a mechanism. It must:
   - drop all forwarded traffic, IPv4 and IPv6. Nothing on the printer needs forwarding once the hotspot stops sharing the uplink, so this also stops the printer routing between `wlan0` and `eth0` (NetworkManager's shared mode turns IPv4 forwarding on for every interface);
   - match on interfaces, not subnets, so it needs no `wlan0` or `eth0` events;
   - be loaded before NetworkManager raises `ap0` at boot, and never be removed by a `down` event;
   - not depend on NetworkManager's firewall backend. With the iptables backend, NM puts its own `-i ap0 -s 10.42.0.0/24 -j ACCEPT` at the top of FORWARD on every activation, ahead of a jump added with `-C || -I 1`;
   - never block the hotspot's own access to `10.42.0.1` (DHCP, DNS, `:80`). The INPUT rule in `10-ap-isolate` matches the uplink subnet, so on a `10.0.0.0/8` uplink it rejects DNS, DHCP renewals and `:80` on the hotspot. Use `-m addrtype --dst-type LOCAL ! -d 10.42.0.1` (or the nft equivalent) instead.

   **The properties above are required of the forward half.** The INPUT half (the hotspot reaches only `10.42.0.1` among the printer's own addresses) should meet them too. An nft `input` rule such as `iifname "ap0" fib daddr type local ip daddr != 10.42.0.1 reject` does, and the `10-ap-isolate` dispatcher rule is acceptable for phase 1. On `10.42.0.1`, tcp/22, udp/5353 and udp/7127 stay open to the hotspot: they are the printer's own services, and D11 is about what the hotspot reaches beyond the printer.

   The simplest way to get all of these is a `forward` chain with policy `drop` in `muon3d-firewall.nft`, which loads before NetworkManager. `tests/test_firewall_ruleset.py` forbids forward hooks there today to protect "the hotspot's internet access", which D11 removes, so update that test and the comments in `muon3d-firewall.nft` and `docs/connectivity/SPEC.md` AP-7. Update the `10-ap-isolate` tests that pin today's behaviour. Hotspot clients lose internet access, which today is incidental.
   - **Client to client.** Traffic between two hotspot clients is bridged by the Wi-Fi firmware and never reaches FORWARD. Set `wifi.ap-isolation=1` in `ap0-con` (bench B7 checks brcmfmac honours it).
2. **HTTPS fails fast.** Add `iifname "ap0" tcp dport 443 reject with tcp reset` to `muon3d-firewall.nft`, so HTTPS captive probes fail at once instead of timing out. A reject isn't a listener; check it against #305's rule that declared ports must equal allowed ports.
3. **Pin `10.42.0.1`** in `ap0-con` (`address1=10.42.0.1/24` under `[ipv4]`). It's NetworkManager's shared-mode default today, not pinned. Persisted profiles need migrating: `muon-ap-provision.sh` already edits the persisted profile before NetworkManager starts, so do it there (#174's `pin_hotspot()` is in a draft and can't be reused). In the same migration, set `ipv6.method=disabled` (with `auto`, the printer may accept router advertisements from a hotspot client) and `wifi.ap-isolation=1`. Say in the PR what happens on rollback.
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
       "~*^muon-[a-z]+-[0-9a-f]{4}(\.local)?$"         1;
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
| `GET /region/options` | `{countries, preselect, basis, locked, configurations}` |
| `POST /region/country` `{country}` | `{country, configuration, domain, verified}` |

**`GET /region` fields:**

- `surroundings` is one of `settled`, `offer`, `not-registered`, `locked` or `unknown`.
- `basis` is `joined-network` or `plurality`.
- `channels` lists the channels on which the applied configuration may start a transmission. It's empty when no configuration is applied.

**`GET /region/options` fields:**

- `countries` is a flat, sorted list of ISO codes.
- `preselect` is the detected country if the token allows it, otherwise the token's `default_country`.
- `configurations` (OS-2) is `{<config id>: {countries, channels}}`, one entry per configuration the token offers: the offered countries it covers, and the channels it may start a transmission on (`initiable_2g4` plus `initiable_5g`). It's `{}` without a usable token.

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
   - `no-table`, `bad-table`, `no-fingerprint` (build or data faults);
   - `busy`.

   Read the reason from the agent's whole output, not only its last stderr line: openssl's and iw's messages can span several lines, and a `bad-signature` read as `apply-failed` tells the owner the wrong thing.

   Today a 409 carries the agent's last stderr line as free text.
2. **Timeouts and concurrency.** Cut `SET_COUNTRY_TIMEOUT_S` from 90 s to under 60 s, the Moonraker proxy's limit; a timeout answers 504 with `code: "busy"`. Add a concurrency guard that answers 409 `busy`. **The guard is held until the agent exits, including after a timeout.** Killing `sudo` doesn't stop the root agent, so releasing the guard at the timeout lets a retry start a second agent on the same radio.
3. **Channels per configuration.** Add `configurations` to `GET /region/options`, taken from `regions.json`'s `initiable_2g4` and `initiable_5g`. The surfaces can then say whether a network is reachable after a switch.
4. **Signing keys and tokens** (KAN-321 / KAN-132). Until they exist, every unit is `no-signing-key` and setup treats it as market `none`.

**Tier-2 picker data.** "Countries where this language is spoken" and the country names stay on the UI side: `SPOKEN_IN` and `Intl.DisplayNames` in MuonUI#31's `regionCountries.ts`, ported to Fluidd. `regions.json` has no language or continent data.

### Time (OS-6, new)

**Today:**

- There is no timezone handling.
- `systemd-timesyncd` runs with Debian defaults.
- `/etc/fake-hwclock.data` **may not be persisted.** MuonOS's docs disagree (IN-DEVELOPMENT.md says a vendor recipe persists it; the Rugix audit says nothing does), so bench B8 checks it. If it isn't, every boot starts from the image's baked time.
- There are no sudo grants for `timedatectl`, `date` or `fake-hwclock`.
- The clock-before-TLS precondition is **KAN-270, prerequisite 2** ("Time sync has to become a precondition of an install"). KAN-198, the release-side signing, is done.

**Routes (built in MuonOS#315, branch `feat/KAN-412-clock-and-time-zone`):**

| Route | Does |
|---|---|
| `GET /time` | `{ "epoch_ms", "ntp_synced": true \| false \| null, "tz" }`. `tz` is the zone in effect (where `/etc/localtime` points), falling back to `UTC`. `null` means timedated didn't answer; treat it as not synced. |
| `POST /time` `{ "epoch_ms" }` | Refuses with `409 ntp_synced` if NTP has synced. Ask `timedatectl show -p NTPSynchronized --value`; if timedated can't answer, fall back to `/run/systemd/timesync/synchronized`, and if neither can be read, **refuse** (fail closed). Refuses with `422 invalid_clock` for a time before the image build or more than 20 years after it. The build time is `/etc/muon3d/build.json`'s `created_at`, or 2026-01-01 if that can't be read (the same floor `muon_setup` uses). Otherwise it runs `sudo -n date -u -s @<s>`. `503 clock_unavailable` if `date` fails. (`timedatectl set-time` isn't used: timedated refuses it while NTP is enabled.) Drop `fake-hwclock save` and its sudo grant unless B8 shows `/etc/fake-hwclock.data` is persisted. |
| `POST /time/zone` `{ "tz" }` | `422 invalid_timezone` unless the name is in the image's tzdata. Runs `sudo -n timedatectl set-timezone`. `503 timezone_unavailable` if that fails, or `503 timezone_not_kept` if the zone was applied but couldn't be saved. |

Errors carry `detail: {code, message}`.

**Persisting the time zone.**

- It's kept in `/var/lib/muon3d/time/timezone`, persisted by its own `muon3d-time.toml`, and `muon3d-timezone.service` re-applies it early in every boot, before `muon_setup` starts.
- Don't bind-mount `/etc/localtime`: `timedatectl` replaces that symlink with an atomic rename, which a bind mount breaks.
- The sudoers entries are pinned to their argument shapes (`date -u -s @<10 digits>`, `timedatectl set-timezone <zone>`). There's a row in `docs/privacy/data-inventory.md`.

**The floor.** Aux's routes are mirrored at `/server/aux/time` and `/server/aux/time/zone`, which any LAN or remote client could otherwise call directly. `/server/aux/time` is floored: in Moonraker's `FLOOR_PREFIXES` (Moonraker#25, with the other OS routes only `muon_setup` may write) and in MuonOS's pin bump (02 §1). `muon_setup` reaches Aux through the in-process helper, so it's unaffected, and it accepts `clock` only from a hotspot caller (02 §3).

## 4. Wi-Fi join

**Today (`wifi_routes.py`):**

- `GET /wifi/scan?rescan=` returns `DeviceWifi {in_use, ssid, bssid, mode, chan, freq, rate, signal, security}`.
- `GET /wifi/device/status` returns `{device, device_type, state, connection, state_reason?, user_disconnected}`. `state_reason` is nmcli's raw `"N (text)"`. **It can't explain a failed join:** NetworkManager 1.42 (bookworm) moves a failed device straight to `disconnected` with reason 0, so a read after nmcli exits almost always gives `0 (No reason given)`.
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
3. Treat **`status: "restored"`** and **400** as failures, and take the code from Aux (W2): `detail.code` on a 400 or 500, or `code` on a `restored` answer. Don't derive it from `state_reason` (see above).

   **How Aux maps a failure (W2).** nmcli exit 10 is `ssid_not_found` and exit 3 is `timeout`. Otherwise Aux maps the reason nmcli prints at the moment of failure ("Error: Connection activation failed: <reason>", in English because nmcli runs under `LANG=C`), falling back to `GENERAL.REASON` only when that isn't 0. Check a wrong-password join on a device before merging.

   | Reason | `code` |
   |---|---|
   | 7 (no secrets), or 8–11 (supplicant) on PSK | `wrong_password` |
   | 8–11 on 802.1X | `eap_failed` |
   | 53 (SSID not found) | `ssid_not_found` |
   | 5, 15–17 (IP config / DHCP) | `no_address` |
   | Anything else, or 45 s passing | `timeout` |

**Changes needed:**

| ID | Change |
|---|---|
| W1 | Replacing a saved network's password is **#210**. Hidden networks (OS-3, MuonOS#322) use `nmcli device wifi connect … hidden yes` under the existing `--wait 45 device wifi connect *` grant: nmcli does a directed scan and reads the security from the probe response, so no `connection add` and no new grant are needed. nmcli looks for the AP straight after requesting that scan, so **retry once after about 3 s on exit 10** when `hidden` is set. The security hint validates the request and refuses `wep` and `enterprise` (`unsupported_security`) before anything changes. N5 must pass on a fresh unit. |
| W2 | Return `detail: {code, message}` with the codes above instead of free text, and add `code` beside `warning` on a `restored` answer. The 409 for too many Wi-Fi operations keeps a string `detail`. Raise refusals inside an `except` with `from None`, so no exception chain carries the PSK argv. |
| W3 | **Enterprise** (PEAP/TTLS). This needs a new privilege path: `nmcli connection add` or D-Bus/polkit. The sudoers header already says widening means D-Bus/polkit. It's the largest OS item, and it can ship later without blocking anything else (D10). |
| W4 | `POST /wifi/ca_cert` stores certificates in `/etc/NetworkManager/certs`. That directory **isn't persisted**, so it needs a `[[persist]]` declaration, an ID-9 row and a data-inventory row. |
| W5 | **The internet check.** The update check already talks to the Nexigon hub before any opt-in. Muon doesn't run that host, so it can't add a `generate_204` endpoint. Decide `internet` with a **TLS handshake on port 443 to the hub the image's agent is configured for** (read from the installed agent config; `eu.nexigon.cloud` in production). The chain and hostname are verified, the time isn't (there's no RTC before NTP), and no HTTP request is sent. `true` if it verifies; `"portal"` if a connection opens but TLS or the certificate fails; `false` if nothing connects; `null` with no uplink. **Bound the whole check, DNS included, to about 5 s**, so a broken resolver can't hold the Wi-Fi operation slots. `check_connectivity()` isn't used: it runs `nexigon-agent`, returns only error text, and fails before the clock is set. `GET /wifi/uplink` returns `{kind, ssid, addresses, gateway, internet, checked_at}` and adds no new destination. A 503 with `code: "uplink_unavailable"` means netlink couldn't be read; `muon_setup` treats it as `internet: null`. |
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
| `POST /link/start` | Moonraker `muon_link` | Sends `LinkStart` and moves to `connecting`. The orchestrator's `LinkCode` then moves it to `code`. **Refused (409) while an offer waits**; a new code never replaces a pending offer. |
| `POST /link/confirm` `{ "account", "fingerprint" }` | **Panel only**, directly | Accepts the pending offer, which moves to `linked`. The body must echo the `offer` the panel showed, read from `GET /link`: 400 without it, 409 if the waiting offer is a different one. A different second offer is declined, and the attempt fails. (muon-link#24 `d26c5b5`.) |
| `POST /link/cancel` | Moonraker `muon_link`; panel | Declines a pending offer and drops the connection |
| `POST /link/unlink` | **Panel only** | Removes the link and every grant (LINK-8) |

Errors come back as `409 {"error": "<sentence>"}`. Nothing is pushed, so callers poll.

**OS-10 (new):**

1. Add `location /muon-link/ { proxy_pass http://127.0.0.1:7131/; }` to the **`:100` vhost only** (`recipes/muon-ui/files/muon-ui.nginx.template`). Never add it on `:80`.
2. Set `MUON_LINK_ORCH_ID` and `MUON_LINK_RELAY_URL` in `recipes/muon-link/files/muon-link.service`, and `MUON_LINK_ORCH_ADDR` if needed.
3. **Never set `MUON_LINK_DISCOVERABLE=1`** (NET-2, Tier 1).

**Open questions for the connectivity owner:**

- ~~**LINK-2 vs PR #24.**~~ **Settled 25 Sep (ADR 0026, KAN-405):** LINK-2 now follows the console. The orchestrator mints a 6-digit code valid for 600 s, and muon-console limits failed claims and reissues a code after a decline. muon-link still enforces nothing, and doesn't need to. LINK-3 (the client-key comparison on the panel) is unchanged and is KAN-415.
- **LINK-4 can't work with orchestrator-made codes.** There's no code without internet.
- **`/link/confirm` has no proof of a knob press** (ML-2).

## 7. Completion marker and factory reset

**Contract (decided 16:40): MuonOS branch `feat/KAN-413-setup-marker` (OS-7).** It replaces #174's `GET/POST /setup` for the marker.

| Route | Does |
|---|---|
| `GET /setup/complete` | Returns `{ "complete", "completed_at", "by" }` |
| `POST /setup/complete` `{ "by": "muon_setup" \| "migrated" }` | Writes the marker. Idempotent: the first record is kept. `migrated` is for printers the migration check marks complete. |
| `DELETE /setup/complete` | Clears the marker and returns `{ "complete": false, "completed_at": null, "by": null }`. Used only by `muon_setup`'s `reset`. |

- **The marker** is `/var/lib/muon3d/setup/complete`. Its existence is the fact, so a damaged file still reads as complete; the hotspot lifecycle checks it with `[[ -e ]]`.
- **Persisted** by `recipes/rugix-ctrl-config/files/state/setup.toml`. It survives OTA updates and rollback, and a factory reset removes it. KAN-413 adds the ID-9 and data-inventory rows.
- **`muon_setup` writes it** with `POST /setup/complete` at `finish` (`by: muon_setup`) and when migration marks a printer complete (`by: migrated`). It reads `GET /setup/complete` only during the migration check ([01-flow.md §7](01-flow.md#7-printers-already-in-the-field)). Once `muon_setup` has its own state, that state is authoritative.
- **The whole prefix `/server/aux/setup` must be floored**, in **both** lists. Flooring `/server/aux/setup` rather than `/server/aux/setup/complete` also covers #174's `POST /setup` if it lands first, and anything added under it later:
  - Moonraker's `FLOOR_PREFIXES`, added by Moonraker#25;
  - MuonOS's `EXPECTED_FLOOR` in `test_trusted_clients.py`, a `return 403` location in `fluidd.nginx.template`, and a `FLOOR_CASES` row in `test_floor.py`, all in the MuonOS PR that bumps the Moonraker pin past #25. #313 must not merge before that.

  Without it, a LAN client could `DELETE` the marker through `aux_api_proxy`.
- **#174 and MuonUI#31 overlap with this.**
  - #174 has its own `setup_routes.py` under the same `/setup` prefix, and a `setup.toml` at the same path as OS-7's, so the two conflict on purpose. #174 must drop both, or rebase onto KAN-413.
  - MuonUI#31's `aux.setup.get()` and `aux.setup.complete()` calls go away (UI-1).
- **Time zone.** OS-6 keeps it separately, in `/var/lib/muon3d/time/timezone` with its own `muon3d-time.toml`.

**Factory reset** is `/usr/sbin/muon3d-factory-reset`, run as root from the console or SSH. It runs `rugix-ctrl state reset`, which clears everything persisted.

| What | After a reset |
|---|---|
| Market token (config partition) | Survives |
| Hostname and SSID | Re-derived from the serial, so the printer keeps the same name |
| Hotspot key | Regenerated, so the Wi-Fi QR code changes |
| Declared country and setup marker | Cleared (the marker once KAN-413 merges, the country once #174 merges) |

## 8. Phase 2: Bluetooth

**Decided 25 Sep: phase 2 uses Iroh_BLE** ([Muon-3D/Iroh_BLE](https://github.com/Muon-3D/Iroh_BLE), Harry's clean-room BLE transport for Iroh; Muon_Internal_Documentation ADR 0027, KAN-402). AP-10 in MuonOS `docs/connectivity/SPEC.md` now says so (MuonOS#317). The earlier `muon-setup-ble` design (a GATT service carrying JSON-RPC, SPAKE2, AES-GCM) is dropped: Iroh already authenticates both ends with the endpoint keys, so a second key exchange adds nothing.

- **Transport.** muon-link on the printer runs Iroh_BLE's BlueZ peripheral (`blelink-bluer`) beside its IP transports. A phone (the Muon3D app, through `blelink-ffi`) or a Chromium browser (`blelink-wasm`) opens an ordinary Iroh connection over BLE. The connection moves to Wi-Fi or the relay later without dropping.
- **When it listens.** It advertises only while setup is incomplete, or while "Add a phone or computer" is open. Unclaimed printers accept only the setup protocol over BLE (Iroh_BLE's `EndpointHooks` gating).
- **What it carries.** `muon_setup`'s methods, over an ALPN of their own (Iroh_BLE's planned `provision/1`, not started). The phone learns the printer's EndpointId from the panel QR code before it trusts the connection.
- **Caller class (new, blocking).** A Bluetooth peer arrives through muon-link, so today S5 would class it as `remote` and make it read-only. muon-link must tell Moonraker which transport a connection used, and `muon_setup` must map a BLE setup peer to caller kind `bluetooth` with hotspot rights while setup isn't complete. That fact must come from muon-link, never from a claim by the client.
- **Knob confirmation** before the first write, as on the other paths.
- **Still open in Iroh_BLE:** `HELLO` authentication (a stranger in range can squat a known EndpointId's route, which is denial of service, not access), a first run on real radios, and L2CAP in the phone drivers.
- **Bluetooth SIG qualification** is still owed for the product, about $12k per design (Iroh_BLE PLAN §7). The clean-room licence doesn't remove it.
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
| B7 | Two phones on the hotspot with `wifi.ap-isolation=1`: can one reach the other? Also `nft list tables` (NM's firewall backend) and `ip -6 addr` on `ap0` after a client sends a router advertisement | S7, OS-1 |
| B8 | `findmnt /etc/fake-hwclock.data`, and the clock after a power cycle with no network | 01 §2.2, OS-6 |
