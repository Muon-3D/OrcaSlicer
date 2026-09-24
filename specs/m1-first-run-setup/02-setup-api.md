# 02 · The `muon_setup` Moonraker component

**Repo:** `Muon-3D/Moonraker`. **New file:** `moonraker/components/muon_setup.py`. **Tests:** `tests/test_muon_setup.py`.

`muon_setup` owns the setup state machine in [01-flow.md](01-flow.md). MuonUI, the Fluidd `/setup` page, the Fluidd dashboard card, and later the Bluetooth bridge all render its state and send it intents. None of them talks to the Aux Wi-Fi routes directly during setup.

## 1. Conventions this component follows

The implementing agent must match these repo conventions. They were verified in the fork at `dc76b59`.

| Topic | Convention |
|---|---|
| Loading | A component loads only if its config section exists (`server.py:267-287`). Add `[muon_setup]` to `core/M1/moonraker.core.conf.template`. Load order is arbitrary, so look up `aux_api_proxy`, `database`, `klippy_apis`, `update_manager` and `machine` lazily at call time, or call `self.server.load_component(config, "aux_api_proxy")` in `__init__`. |
| Endpoints | `self.server.register_endpoint(path, ["GET"] or ["POST"], handler)`. Use one verb per path so each RPC name is simply the path with dots: `/server/muon/setup/network` becomes `server.muon.setup.network`. Leave `auth_required` at its default, because trusted clients already pass (SEC-1 Level 0). |
| Notifications | `self.server.register_notification("muon_setup:muon_setup_changed")` in `__init__`, then `self.server.send_event("muon_setup:muon_setup_changed", state)`. Clients receive `{"method":"notify_muon_setup_changed","params":[state]}`. |
| In-process events | `send_event("muon_setup:complete", state)` on finish, for other components through `register_event_handler`. |
| Database | Moonraker SQLite at `/home/printer_data/database`. It survives reboots and OTA updates and is wiped by a factory reset (ID-2/ID-3). Register a **new** namespace with `database.register_local_namespace("muon_setup", forbidden=True)`. Don't register `"muon"` again, because `aux_api_proxy` already owns it and a second registration raises. |
| Talking to Aux | Use the helpers on `aux_api_proxy` (`get()` / `post()`), which add `X-Aux-Api-Key` and map Aux errors. Don't make raw HTTP calls to `127.0.0.1:6789`. |
| Startup | `component_init` must **not** raise if Aux is down. Register every endpoint in `__init__`. Endpoints that need Aux return `{"ok": false, "error": {"code": "aux_unavailable"}}` until Aux answers. `aux_api_proxy` currently fails to register identity when Aux is late; don't repeat that mistake ([10-work-plan.md MR-9](10-work-plan.md)). |
| Floor | `muon_floor.py` only knows *loopback* and *network*. Add `/server/muon/setup/reset` to `FLOOR_PREFIXES` and extend `tests/test_muon_floor.py`. Every other rule in §3 is enforced inside this component. |
| Tests | Plain pytest unit tests with hand-written fakes (`FakeServer`, `FakeConfig`, `FakeWebRequest`, a fake `database` with async `get_item`/`insert_item`/`delete_item`/`register_local_namespace`, and a fake `aux_api_proxy`). Run coroutines with `asyncio.run(...)`, as in `tests/test_aux_api_proxy.py`. CI runs flake8 (max line 88) and mypy over `moonraker`, so both must pass. |

## 2. Config

```ini
# core/M1/moonraker.core.conf.template
[muon_setup]
# Languages offered on the first screen, in BCP 47, in the order shown.
languages: en, de, fr, es, it
# Seconds the hotspot stays up after `finish` when an uplink has an address.
hotspot_off_delay: 900
# Seconds before the panel treats a phone's driver claim as lapsed.
driver_lease: 30
# Seconds allowed for a Wi-Fi join, from submit to an address.
join_timeout: 45
# Ready-to-print manifest shipped in the image; the built-in default is used if missing.
ready_manifest: /usr/share/muon/setup/ready.json
```

Add a test next to `tests/test_m1_config_template.py` that pins the `[muon_setup]` section and its keys.

## 3. Who is calling

Classify each request once, at the start of the handler. Use `webreq.get_ip_address()`, which honours nginx's `X-Real-IP`, and `webreq.get_current_user()`.

```python
def caller_kind(webreq) -> str:
    user = webreq.get_current_user()
    ip = webreq.get_ip_address()
    if isinstance(webreq.transport, InternalTransport):  # another component
        return "internal"
    if user is not None and user.source == "muon_gateway":
        return "remote"                                   # Iroh via muon-link (sentinel 192.0.2.1)
    if ip is None:
        return "other"                                    # MQTT etc.: treat as network, deny writes
    if ip == ip_address("192.0.2.1"):
        return "remote"
    if muon_floor.local_address(ip):
        return "panel"                                    # MuonUI on :100 → loopback
    if ip in ip_network("10.42.0.0/24"):
        return "hotspot"
    if user is not None and "network" in user.groups:
        return "lan"                                      # a Level 0 trusted client
    return "other"
```

| Action | panel | hotspot | lan | remote | other |
|---|---|---|---|---|---|
| `GET` state, options, networks | ✓ | ✓ | ✓ | ✓ (state only) | ✗ |
| Step writes while `state != complete` | ✓ | ✓ | ✓ | ✗ | ✗ |
| Step writes after `complete` (`network`, `name`, `remote`) | ✓ | ✓ | ✓ | ✗ | ✗ |
| `ready` item actions (`start`, `confirm`) | ✓ | ✗ | ✗ | ✗ | ✗ |
| `ready` skip | ✓ | ✓ | ✓ | ✗ | ✗ |
| `reset` | ✓ (and floored) | ✗ | ✗ | ✗ | ✗ |

Any refused action returns HTTP 403 with the `ServerError` message `"muon_setup: not allowed from <kind>"`.

**HTTP write hygiene (CSRF / DNS rebinding).** Refuse any HTTP `POST` in this component with 415 unless it has `Content-Type: application/json`. Refuse it with 403 if the `Host` header isn't one of these:

- `10.42.0.1`
- the printer's hostname, bare or with `.local`
- `muon3d.local`
- `localhost:100`
- one of the printer's current IP addresses

If an `Origin` header is present, it must match the same set. Websocket JSON-RPC calls skip these checks, because Moonraker already checks the websocket origin. If `WebRequest` doesn't expose request headers, add a small `# MUON` accessor in `application.py` and cover it with a test.

## 4. Persistence

- **Namespace** `muon_setup`, **key** `state`: the whole state document in §6, minus the fields computed on read (`hotspot`, `clock`, `capabilities`).
- **Write-through.** Persist before sending the change notification. Use `insert_item` and await it.
- **Completion marker.** On `finish`, also write the KAN-203 marker so older MuonUI builds and MuonOS services agree ([03-printer-os.md §7](03-printer-os.md#7-completion-marker-and-factory-reset)).
- **First start with no stored state.** Run the migration check in [01-flow.md §7](01-flow.md#7-printers-already-in-the-field) before creating a `new` state.
- **Schema version.** `"version": 1`. Load an unknown version as read-only, log it, and treat it as `complete`, so a downgrade never re-runs setup.

## 5. Endpoints

**Every write:**

- takes a JSON body containing `rev`, the `rev` the client last saw;
- returns HTTP 200 with the body below, whatever the domain outcome.

```json
{ "ok": true, "error": null, "state": { … } }
```

- **Domain failures** are `ok: false` with `error: {"code": "<code>", "message": "<English, for logs>", "detail": {…}}`. The codes are listed in [08-errors.md](08-errors.md). Clients show their own translated copy based on `code`, never `message`.
- **Transport, permission and validation failures** use `ServerError`: 400 for a malformed body, 403 for a refused caller, 415 for the wrong content type.
- **`stale_rev`.** A write whose `rev` is not the current `rev` gets `ok: false, error.code: "stale_rev"` together with the current state. The client re-renders and doesn't retry by itself.
- **`busy`.** A write while `op` is set gets `busy`. The one exception is the matching `…/cancel`.

### 5.1 `GET /server/muon/setup` · `server.muon.setup`

Returns the state document (§6). Every change also goes out as `notify_muon_setup_changed`.

### 5.2 `GET /server/muon/setup/options` · `server.muon.setup.options`

Query `country` (optional) adds that country's time zones. Query `language` (optional, defaults to `steps.language.value`) orders the picker's second tier.

```jsonc
{
  "languages": [ { "code": "en", "endonym": "English" }, { "code": "de", "endonym": "Deutsch" } ],
  "region": {
    "market": "picker",                     // picker | locked (token configs == ["us"]) | none (no valid token)
    "applied": { "country": "DE", "config": "de", "declared": false },   // from Aux GET /region
    "default_country": "DE",                // from the token; may be null
    "permitted_channels": [1,2,3,4,5,6,7,8,9,10,11,12,13,36,40,44,48],
    "for_language": ["DE", "AT", "CH"],     // picker tier 2, ordered by speaker count
    "all": { "Europe": ["AT","BE","…"], "Americas": ["CA","…"] },     // picker tier 3, offered countries only
    "support_code": null                    // set when market == none ("needs re-registering")
  },
  "timezones": ["Europe/London"],           // only with ?country=
  "ready_manifest": { "version": 1, "items": [ … ] }
}
```

- **Region data** comes from Aux `GET /region` and `GET /region/options` (KAN-321 Rev 11). `muon_setup` never keeps its own country list.
- **Time zones** come from the image's tzdata (`/usr/share/zoneinfo/zone1970.tab`). Order them so the most populous zone is first; `zone1970.tab` already lists a country's principal zone first.

### 5.3 Language

`POST /server/muon/setup/language` with `{ "rev": 3, "code": "de" }`

- The `code` must be in the configured `languages`, otherwise `unsupported_language`.
- Sets `steps.language = {status: "done", value, source}` and moves `new` to `in_progress`.
- Writes Fluidd's default locale if the owner hasn't chosen one: `database.insert_item("fluidd", "uiSettings.general.locale", code)` when `get_item(...)` returns nothing. See [05-phone-setup-page.md §9](05-phone-setup-page.md#9-language).

### 5.4 Clock and time zone

`POST /server/muon/setup/clock` with `{ "epoch_ms": 1790251200000, "tz": "Europe/London" }`

- No `rev` is needed, and the call doesn't change `rev`.
- **The clock** is set through Aux only when the system clock isn't NTP-synchronised **and** it differs by more than 2 s. `epoch_ms` must be later than the image's build time, otherwise the result is `invalid_clock`.
- **The time zone**, if given and valid in tzdata, is set through Aux. It sets `clock.tz_source = "phone"`.

`POST /server/muon/setup/timezone` with `{ "rev": 9, "tz": "America/Chicago" }`

- `tz` must be one of the declared country's zones. Otherwise the result is `invalid_timezone`.
- It sets `clock.tz_source = "owner"`.

### 5.5 Networks

`GET /server/muon/setup/networks?rescan=true` · `server.muon.setup.networks`

- Scans synchronously, with a 15 s cap, through Aux `GET /wifi/scan?rescan=1`. Under an undeclared domain, Aux widens to `00` and scans passively (Rev 11).
- Normalises the result and adds a region suggestion to each network:

```json
{
  "ok": true,
  "scanned_at": 1790251262.4,
  "ethernet": { "present": true, "carrier": false, "address": null },
  "networks": [
    { "ssid": "HomeWiFi", "security": "wpa2", "signal": 78, "band": "2.4", "channel": 6,
      "bssids": 2, "saved": false, "in_use": false, "supported": true,
      "channel_permitted": true,
      "region_suggestion": { "country": "GB", "source": "ap" } }   // ap | neighbours | default | null
  ]
}
```

- **Normalising:**
  - Drop empty SSIDs; hidden networks go through "Other network…".
  - Drop the printer's own hotspot SSID.
  - Merge entries with the same SSID. Keep the strongest signal, count the BSSIDs, and report the band and channel of the strongest.
  - Sort by signal, strongest first.
- **Security values:** `open`, `owe`, `wep`, `wpa2`, `wpa3`, `wpa2_wpa3`, `enterprise`. `wep` has `supported: false`.
- **`channel_permitted`** checks the network's channel against `region.permitted_channels` for the *suggested* country's configuration, or the applied one if there's no suggestion.
- **`region_suggestion`** follows the Rev 11 order:
  1. the network's own country element (`ap`);
  2. the neighbours' plurality, where at least 3 must name a country and it must lead by at least 2 (`neighbours`);
  3. the token's `default_country` (`default`).

  Every source is filtered through the token.
  - Use Aux's own resolver, the one behind the connect-time prompt in `WifiManagerView`.
  - If Aux doesn't expose it per network, add `GET /region/suggest?ssid=` to Aux ([03-printer-os.md §3](03-printer-os.md#3-region-time-zone-and-clock)) rather than parsing country elements in Moonraker.

### 5.6 Joining a network

`POST /server/muon/setup/network`

```jsonc
{ "rev": 7, "kind": "wifi",
  "ssid": "HomeWiFi", "hidden": false, "security": "wpa2",
  "psk": "correct horse battery staple",
  "region": "GB",                 // the country the owner confirmed on the region line; omitted on locked units
  "eap": null }
// Enterprise:
{ "rev": 7, "kind": "wifi", "ssid": "eduroam", "hidden": false, "security": "enterprise", "region": "GB",
  "eap": { "method": "peap", "phase2": "mschapv2", "identity": "ab123@uni.ac.uk",
           "anonymous_identity": "anonymous@uni.ac.uk", "password": "…",
           "ca_cert_id": "c_1f2e", "domain_suffix_match": "uni.ac.uk", "no_ca_check": false } }
// Ethernet already has an address, and the owner confirms it:
{ "rev": 7, "kind": "ethernet" }
```

**Behaviour:**

1. **Validate the request.**
   - `ssid` must be 1–32 bytes UTF-8.
   - `region` is required when `region.market == "picker"` and no country is declared yet (`region_required`). It must be offered by the token (`region_not_offered`).
   - A `psk` must be 8–63 characters, or exactly 64 hex characters. `open` and `owe` take no psk.
   - For `enterprise`, `method` must be `peap` or `ttls`, `phase2` must be `mschapv2` or `pap`, and `identity` and `password` are required. Either `ca_cert_id` or `no_ca_check: true` is required.
   - Any violation gives `invalid_network`, with `detail.field` naming the field.
2. **Start the operation.** The HTTP response comes back immediately with `ok: true` and the state showing `op`.
3. **Apply the region if needed.** If `region` differs from `region.applied.country`, or no country has been declared yet:
   - Set `op = {kind: "region_apply"}` and call Aux `POST /region/country`.
   - This is live and takes about 8 s. All Wi-Fi goes down, the **hotspot** included, and comes back.
   - Aux failures map as follows:
     - `busy` → `region_busy`;
     - a read-back mismatch → `region_apply_failed`;
     - a country not in the token → `region_not_offered`;
     - no valid token → `needs_reregistration`.

     Stop there on failure. Nothing is declared.
   - If, after the apply, `permitted_channels` doesn't include the network's channel, stop with `channel_not_permitted`.
4. **Join through Aux** ([03-printer-os.md §4](03-printer-os.md#4-wi-fi-join)). Set `op = {kind: "join", phase: "saving"}` and move `op.phase` through `saving` → `associating` → `authenticating` → `dhcp` → `internet_check` → `update_check`. Notify on each change.
5. **Success means an IPv4 address on the uplink.** Set:
   - `network = {status: "done", kind, ssid, addresses: [...], hostname_local: "<hostname>.local", internet: true|false|"portal", error: null}`;
   - `update.status` to `pending` only if the internet check passed, an update exists and the clock is synced. Otherwise it is `hidden`.
6. **Failure.** `network.error = {code, at_phase}`, the step stays `pending`, `op` is cleared, and nothing is kept: the Aux connection profile is deleted. The codes are `wrong_password`, `ssid_not_found`, `no_address`, `timeout`, `eap_failed`, `cert_invalid` and `unsupported_security`.
7. **`internet: "portal"`** (a guest network with a web sign-in page) is still `done`, because the address works on the LAN. It carries `error: {code: "portal_required"}` so the surfaces can warn and offer "Choose another network".
8. **Never store or echo secrets.** `psk` and `eap.password` are never persisted by `muon_setup`, never logged (redact them in any debug dump), and never put in state or events. Only NetworkManager keeps them.

`POST /server/muon/setup/network/cancel` with `{}`: cancels a running join. The partial profile is removed.

`POST /server/muon/setup/network/ca_cert` takes a multipart upload of one PEM or DER file of at most 16 KiB.

- It returns `{ok: true, ca_cert_id}`.
- The certificate is handed to Aux to store root-only. `muon_setup` only keeps the ID.
- Allowed callers are panel, hotspot and lan.

### 5.7 Name

`POST /server/muon/setup/name` with `{ "rev": 9, "name": "Walnut" }`. An empty string or leaving `name` out means "Keep".

- Refactor `aux_api_proxy._set_identity_name_handler` into a public `async set_friendly_name(name) -> dict` and `async get_identity() -> dict`. Both the existing endpoint and `muon_setup` call these.
- The rules are unchanged: the name is stripped, at most 32 characters, and stored in the `muon` namespace under `friendly_name`.
- The hostname and hotspot SSID keep the derived name (ID-2).
- Marks the step `done`, with `value` set to the effective name.

### 5.8 Update

`POST /server/muon/setup/update` with `{ "rev": 10, "action": "install" | "later" }`

- **`later`** marks the step `skipped`.
- **`install`** runs these steps:
  1. Store `op = {kind: "update_install", target: "<version>"}`.
  2. Call `update_manager`'s upgrade path for `MuonOS`, the same as `POST /machine/update/upgrade?name=MuonOS`.
  3. Mirror its `update_manager:update_response` progress into `op.progress` (0–1).
  4. The printer reboots into the new slot.
- **After the reboot**, `muon_setup` compares `update_manager`'s `version` with `op.target`:
  - equal: the step is `done`;
  - otherwise it rolled back: the step is `pending` with `error: {code: "update_failed"}`.
- `muon_setup` doesn't commit the update itself. It follows the existing OTA commit policy (KAN-358).
- `update_manager` refuses while a print is running (503). Pass that on as `printer_busy`.

### 5.9 Remote access

`POST /server/muon/setup/remote`

```jsonc
{ "rev": 11, "mode": "local" }                       // done at once; no outbound connection is made
{ "rev": 11, "mode": "cloud" }                       // starts the link flow
{ "rev": 11, "mode": "self_hosted", "url": "https://orch.example.org" }   // only if capabilities.self_hosted
{ "rev": 11, "mode": "later" }                       // skipped; the Finish setup card appears
```

**`cloud`** runs the link flow:

1. It calls the link service's start. This is the backend behind `POST /server/muon/link/start` / `GET /server/muon/link`, whose phases Fluidd already expects: `unavailable | unlinked | connecting | code | offer | linked | failed`. See [03-printer-os.md §6](03-printer-os.md#6-link-service-contract).
2. It mirrors that service into the state:

   ```json
   { "phase": "code", "code": "482913", "expires_at": 1790251390.0,
     "claim_url": "https://app.muon3d.com/link?code=482913", "account": null }
   ```

3. **Code renewal.** While `cursor == "remote"`, the mode is `cloud` and the phase is `code`, `muon_setup` starts a new code whenever the old one expires. A code lives 120 s (KAN-190).
4. **Linked.** When the phase reaches `linked`, which happens after the claim and the knob confirmation in the existing flow, the step is `done`, with `account` set to the account's display email or name.
5. **Failure.** `unavailable` or `failed` gives `error.code = "link_unavailable"` or `"link_failed"`. The step stays `pending`, so the owner can pick another option.
6. **Blocked.** If the clock isn't synced, or `network.internet` isn't `true`, `cloud` returns `ok: false` with `clock_unsynced` or `no_internet`.

`POST /server/muon/setup/remote/cancel` with `{}`: stops code renewal. The link service's pending code is cancelled.

`capabilities.cloud_link` is `false` when the link service isn't installed. `capabilities.self_hosted` is `false` until the Tier 3 orchestrator configuration exists. Surfaces hide options whose capability is `false`.

### 5.10 Ready to print

`POST /server/muon/setup/ready` with `{ "rev": 14, "item": "self_test", "action": "start" | "confirm" | "skip" }`

- `confirm` is for `kind: "confirm"` items and for `kind: "panel_flow"` items once MuonUI has finished that flow.
- `start` is for `kind: "macro"` items. It runs `klippy_apis.run_gcode(item.macro)` under `op = {kind: "ready_item"}`.
  - The item is `done` if the macro returns without error.
  - Otherwise it is `failed`, with `error: {code: "self_test_failed", message: <gcode error>}`.
- An item can't start while a print is running (`printer_busy`) or while Klippy isn't ready (`printer_not_ready`).
- `start` and `confirm` are panel-only (§3).
- When every `required: true` item is `done`, the step is `done`.
- `POST /server/muon/setup/skip` with `{"step": "ready"}` skips the whole step.

#### Ready manifest

MuonOS ships the manifest at `ready_manifest` (`/usr/share/muon/setup/ready.json`, read-only in the image). If the file is missing or invalid, `muon_setup` uses the built-in default below and logs a warning.

```jsonc
{
  "version": 1,
  "items": [
    { "id": "transport_clips", "kind": "confirm",    "required": true,
      "title_key": "setup.ready.transport_clips.title", "body_key": "setup.ready.transport_clips.body",
      "image": "transport_clips.webp" },
    { "id": "self_test",       "kind": "macro",      "required": true,
      "macro": "MUON_SELF_TEST", "est_seconds": 120,
      "title_key": "setup.ready.self_test.title", "body_key": "setup.ready.self_test.body" },
    { "id": "load_filament",   "kind": "panel_flow", "required": false,
      "flow": "load_filament",
      "title_key": "setup.ready.load_filament.title", "body_key": "setup.ready.load_filament.body" }
  ]
}
```

- **Fields:**
  - `kind` is one of `confirm`, `macro` or `panel_flow`.
  - `title_key` and `body_key` are i18n keys resolved by each surface. The manifest contains no copy.
  - `image` is a file name under MuonUI's `public/setup/`.
- **Provisional content.** The items and the `MUON_SELF_TEST` macro name are **provisional**. The hardware team owns them: what to remove before the first move, what the self-test checks, and whether calibration runs here or on the first print.
  - Until the macro exists in the Klipper config, `muon_setup` hides `macro` items whose macro isn't defined. Check with `klippy_apis` for `gcode_macro <name>`.
  - Validate the manifest against the schema in `tests/test_muon_setup.py`.

### 5.11 Navigation and control

| Endpoint | Body | Effect |
|---|---|---|
| `POST /server/muon/setup/driver` | `{ "rev": n, "kind": "panel"\|"phone"\|"web", "client_id": "<uuid>" }` | Claims the driver. Calling it again with the same `client_id` renews the claim. It does **not** change `rev`, so the phone can renew every 10 s without causing `stale_rev`. |
| `POST /server/muon/setup/goto` | `{ "rev": n, "step": "network" }` | Moves the cursor ([01-flow.md §3](01-flow.md#3-state-model)). An invalid target gives `invalid_step`. |
| `POST /server/muon/setup/skip` | `{ "rev": n, "step": "network"\|"update"\|"remote"\|"ready" }` | Marks the step `skipped`. `language` gives `not_skippable`. |
| `POST /server/muon/setup/finish` | `{ "rev": n }` | Requires `language` to be `done`, otherwise `required_steps_pending`. Remaining `pending` optional steps become `skipped`. Then `state = complete`, the marker is written, `muon_setup:complete` fires, and the hotspot auto-off is scheduled ([03-printer-os.md §1](03-printer-os.md#1-hotspot-lifecycle)). |
| `POST /server/muon/setup/card/dismiss` | `{ "rev": n }` | `card_dismissed = true` |
| `POST /server/muon/setup/reset` | `{}` | Panel only and floored. Clears the namespace and the marker, and starts over as `new`. For development and support; a real factory reset clears the database anyway (ADR 0005). |

## 6. State document

```jsonc
{
  "version": 1,
  "rev": 42,                                // +1 on every change except driver renewals
  "state": "in_progress",                   // new | in_progress | complete
  "cursor": "network",                      // step id | "finish"
  "driver": { "kind": "phone", "client_id": "b7f3…", "since": 1790251200.1,
              "renewed": 1790251260.4, "lapsed": false } ,          // or null
  "op": { "kind": "join", "id": "op_7", "started": 1790251262.0,
          "phase": "dhcp", "progress": null },                        // or null
  "printer": { "name": "Walnut", "display": "Walnut · 8987", "hostname": "Muon-walnut-8987",
               "fingerprint": "…" },
  "hotspot": { "up": true, "ssid": "Muon-walnut-8987", "clients": 1,
               "auto_off_at": null, "address": "10.42.0.1" },         // computed on read
  "clock": { "synced": false, "source": "fake_hwclock",              // fake_hwclock | phone | ntp; computed
             "tz": "Europe/London", "tz_source": "phone" },         // phone | owner | region | default
  "region": { "market": "picker", "country": "GB", "declared": true, "config": "gb",
              "source": "ap" },                                     // from Aux GET /region; computed
  "capabilities": { "ethernet": true, "enterprise": true, "cloud_link": true,
                    "self_hosted": false, "bluetooth": false },       // computed
  "card_dismissed": false,
  "steps": {
    "language": { "status": "done",    "value": "en", "source": "panel" },
    "network":  { "status": "pending", "kind": null, "ssid": null, "addresses": [],
                  "hostname_local": null, "internet": null, "error": null },
    "name":     { "status": "pending", "value": "Walnut", "derived": "walnut" },
    "update":   { "status": "hidden",  "current": "1.3.2", "available": null, "error": null },
    "remote":   { "status": "pending", "mode": null, "link": null, "error": null },
    "ready":    { "status": "pending",
                  "items": [ { "id": "transport_clips", "status": "pending", "error": null },
                             { "id": "self_test",       "status": "pending", "error": null },
                             { "id": "load_filament",   "status": "pending", "error": null } ] }
  }
}
```

- `source` is the caller kind (§3), or `migrated`.
- `hotspot.clients` is the number of associated stations on `ap0`. The panel uses it to switch its QR code ([04-panel.md §P2](04-panel.md#p2--here-or-on-a-phone-a)). Poll Aux every 2 s while `state != complete` and the hotspot is up, and notify only when the count changes.
- `driver.lapsed` is computed on read as `now - renewed > driver_lease`.

## 7. Identity and discovery changes in `aux_api_proxy`

- `GET /server/muon/identity` gains `"setup": "new" | "in_progress" | "complete"`, taken from `muon_setup` if it is loaded, else `null`. Apps use this to route a found printer ([06-add-printer.md](06-add-printer.md)).
- Register the identity endpoints in `__init__` rather than after `_fetch_spec()`. Return 503 while Aux is unreachable, instead of leaving the route missing.

## 8. Tests (minimum)

Add `tests/test_muon_setup.py` with fakes for `database`, `aux_api_proxy`, `klippy_apis` and `update_manager`, covering:

1. A fresh start with no marker gives `new` and cursor `language`. A fresh start where the marker exists, or a saved Wi-Fi exists, gives `complete` with `source: "migrated"`.
2. The order rules:
   - `goto` forward past the first pending step gives `invalid_step`.
   - Skipping `language` gives `not_skippable`.
3. `rev`:
   - A write with an old `rev` gives `stale_rev` plus the current state.
   - A driver renewal doesn't change `rev`.
4. `op`: a second write during a `join` gives `busy`, and `network/cancel` clears it.
5. Join mapping: every Aux failure reason in [03-printer-os.md §4](03-printer-os.md#4-wi-fi-join) maps to the right `code`, and on failure the profile-delete call is made.
6. Secrets: after a join, no `psk` or password appears in the stored state, the events or the log records (use `caplog`).
7. Access: build the §3 table as a parametrised test over caller kinds × endpoints.
8. `finish` marks pending optional steps as `skipped`, writes the marker, fires `muon_setup:complete`, and asks Aux for the hotspot auto-off.
9. Region:
   - A join with a `region` that differs from the applied country calls `POST /region/country` before `/wifi/connect`.
   - Aux `busy`, a read-back mismatch, a country not offered and a missing token each map to their codes, and none of them declares anything.
   - A channel not permitted after the apply gives `channel_not_permitted`.
   - `region` is required only in `picker` markets with nothing declared yet.
10. Markets: a `locked` market needs no `region` in the join, and `none` passes through `needs_reregistration`. Neither blocks a join on a permitted channel.
11. The update step: a version that matches the target after the reboot is `done`, and a rollback gives `update_failed`.
12. Remote `cloud`: codes renew while the cursor is on `remote`, and they stop after `goto` away or `remote/cancel`.
13. `aux_unavailable`: the component loads and every Aux-backed call returns the code. The GET still works.

## 9. Link endpoints Fluidd already calls (new component `muon_link`)

Fluidd's `discovery.ts` and `LinkPrinterDialog.vue` already call `GET /server/muon/link` and `POST /server/muon/link/start`, but nothing in the Moonraker fork serves them. Add them in a small component, `moonraker/components/muon_link.py` with a `[muon_link]` section, which bridges to muon-link's admin endpoint on `127.0.0.1:7131` (KAN-190, muon-link#14). `muon_setup` uses its Python methods; it doesn't make HTTP calls to itself.

| Endpoint | Callers | Does |
|---|---|---|
| `GET /server/muon/link` | panel, hotspot, lan | `{ phase, code?, expires_at?, account?, message? }`. The phases are `unavailable \| unlinked \| connecting \| code \| offer \| linked \| failed`, the same shape Fluidd expects. `unavailable` means muon-link isn't reachable. |
| `POST /server/muon/link/start` | panel, hotspot, lan | Opens a pairing window (`/pairing/open`) and returns once the phase is `code`. It is rate-limited to 5 per minute per caller IP. |
| `POST /server/muon/link/confirm` | **panel only; add to `FLOOR_PREFIXES`** | The knob Confirm on the offer (`/pairing/confirm`). When Aux's `KnobConfirmationBackend` (the one `dev_mode_consent` uses) works, route through it, so the proof is a physical press and not just a loopback caller. |
| `POST /server/muon/link/cancel` | panel, hotspot, lan | `/pairing/cancel` |

- **Payloads.** KAN-190 doesn't specify the request or response bodies for muon-link's `/pairing/*`. Read them from muon-link at `86809b3` and map them to the phases above. If a phase can't be derived, stop and ask.
- **Events.** Emit `muon_link:link_changed` (clients receive `notify_link_changed`) on every phase change. `muon_setup` listens with `register_event_handler`.
- **Tests.** `tests/test_muon_link.py` with a fake admin endpoint, covering: the phase mapping, the rate limit, confirm refused from every caller kind except panel, and the floor membership test for `/server/muon/link/confirm`.
