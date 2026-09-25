# 02 · The `muon_setup` Moonraker component

**Repo:** `Muon-3D/Moonraker`. **New package:** `moonraker/components/muon_setup/` (`__init__.py` plus modules such as `model.py` and `region.py`). **Tests:** `tests/test_muon_setup.py`.

`muon_setup` owns the setup state machine in [01-flow.md](01-flow.md). MuonUI, the Fluidd `/setup` page, the Fluidd dashboard card, and later the Bluetooth bridge all render its state and send it intents. None of them talks to the Aux Wi-Fi routes directly during setup.

## 1. Conventions this component follows

The implementing agent must match these repo conventions. They were verified in the fork at `dc76b59`.

| Topic | Convention |
|---|---|
| Loading | A component loads only if its config section exists (`server.py:267-287`). Add `[muon_setup]` to `core/M1/moonraker.core.conf.template`. Load order is arbitrary, so look up `aux_api_proxy`, `database`, `klippy_apis`, `muon_link` and `machine` lazily at call time, or call `self.server.load_component(config, "aux_api_proxy")` in `__init__`. |
| Endpoints | `self.server.register_endpoint(path, ["GET"] or ["POST"], handler)`. Use one verb per path so each RPC name is simply the path with dots: `/server/muon/setup/network` becomes `server.muon.setup.network`. Leave `auth_required` at its default, because trusted clients already pass (SEC-1 Level 0). |
| Notifications | `self.server.register_notification("muon_setup:muon_setup_changed")` in `__init__`, then `self.server.send_event("muon_setup:muon_setup_changed", state)`. Clients receive `{"method":"notify_muon_setup_changed","params":[state]}`. |
| In-process events | `send_event("muon_setup:complete", state)` on finish, for other components through `register_event_handler`. |
| Database | Moonraker SQLite at `/home/printer_data/database`. It survives reboots and OTA updates and is wiped by a factory reset (ID-2/ID-3). Register a **new** namespace with `database.register_local_namespace("muon_setup", forbidden=True)`. Don't register `"muon"` again, because `aux_api_proxy` already owns it and a second registration raises. |
| Talking to Aux | Use the helpers on `aux_api_proxy` (`get()` / `post()` / `delete()`), which add `X-Aux-Api-Key` and map Aux errors. Don't make raw HTTP calls to `127.0.0.1:6789`. The error mapping must keep Aux's `detail.code` as well as the message: §5.6a and the join mapping need it (MR-3). |
| Startup | `component_init` must **not** raise if Aux is down. Register every endpoint in `__init__`. Endpoints that need Aux return `{"ok": false, "error": {"code": "aux_unavailable"}}` until Aux answers. `aux_api_proxy` registers its Muon endpoints at load (MR-9). Its mirrored `/server/aux/*` routes still need Aux at load; retrying the spec fetch in the background is a follow-up ticket. |
| Floor | `muon_floor.py` only knows *loopback* and *network*. Four new `FLOOR_PREFIXES` entries, each with `is_floor_endpoint` and `check_floor` tests in `tests/test_muon_floor.py`:<br>• `/server/muon/setup/reset`, added by MR-1;<br>• `/server/aux/setup` (the whole prefix, so it also covers #174's `/setup` if that lands first), `/server/aux/wifi/ap/auto_off` and `/server/aux/time`, added by Moonraker#25, the Moonraker half of OS-5, OS-6 and OS-7. That PR owns them because the routes reach the LAN through `aux_api_proxy` on any image that carries MuonOS#313, #314 or #315, whether or not MR-1 has merged.<br>**MuonOS must match.** The MuonOS PR that bumps the Moonraker pin past these entries adds, in the same change: each path to `EXPECTED_FLOOR` in `test_trusted_clients.py` (an exact, ordered comparison with the pinned tuple), a `location ^~ … { return 403; }` for each in `fluidd.nginx.template`, and rows in `test_floor.py`'s `FLOOR_CASES`. Every other rule in §3 is enforced inside this component. |
| Tests | Plain pytest unit tests with hand-written fakes (`FakeServer`, `FakeConfig`, `FakeWebRequest`, a fake `database` with async `get_item`/`insert_item`/`delete_item`/`register_local_namespace`, and a fake `aux_api_proxy`). Run coroutines with `asyncio.run(...)`, as in `tests/test_aux_api_proxy.py`. CI runs flake8 (max line 88) and mypy over `moonraker`, so both must pass. |

## 2. Config

```ini
# core/M1/moonraker.core.conf.template
[muon_setup]
# Languages offered on the first screen, in BCP 47. Surfaces sort them by endonym.
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
| `POST …/clock` (02 §5.4) | ✗ | ✓ | ✗ | ✗ | ✗ |
| `ready` item actions (`start`, `confirm`) | ✓ | ✗ | ✗ | ✗ | ✗ |
| `ready` skip | ✓ | ✓ | ✓ | ✗ | ✗ |
| `reset` | ✓ (and floored) | ✗ | ✗ | ✗ | ✗ |

Any refused action returns HTTP 403 with the `ServerError` message `"muon_setup: not allowed from <kind>"`.

`internal` callers (other components) may do everything except `reset`.

**HTTP write hygiene (CSRF / DNS rebinding).** Refuse any HTTP `POST` in this component with 415 unless it has `Content-Type: application/json`. The one exception is `network/ca_cert`, which requires `multipart/form-data` instead. Refuse it with 403 if the `Host` header isn't one of these:

- `10.42.0.1`
- the printer's hostname, bare or with `.local`
- `muon3d.local`
- `localhost` (MuonUI on `:100`: nginx sends `Host $host`, which drops the port; its `Origin` is `http://localhost:100`)
- one of the printer's current IP addresses

If an `Origin` header is present, it must match the same set. If `WebRequest` doesn't expose request headers, add a small `# MUON` accessor in `application.py` and cover it with a test.

**Websocket writes.** Moonraker's websocket `check_origin` runs Tornado's same-origin check first, and under DNS rebinding the `Origin` and the `Host` are both the attacker's name, so it passes. So a setup write over websocket JSON-RPC checks the **upgrade request's `Host`** against the same set. The content-type rule doesn't apply to websocket calls. Only `muon_setup` writes use this check, so nothing else in Moonraker changes.

## 4. Persistence

- **Namespace** `muon_setup`, **key** `state`: the whole state document in §6, minus the fields computed on read (`hotspot`, `clock`, `capabilities`).
- **Write-through.** Persist before sending the change notification. Use `insert_item` and await it.
- **Completion marker.** The OS hotspot rules read it (MuonOS KAN-413, [03-printer-os.md §7](03-printer-os.md#7-completion-marker-and-factory-reset)), so it must follow `state` exactly:
  - On `finish`, call Aux `POST /setup/complete {"by": "muon_setup"}`.
  - When the migration check marks a printer `complete`, call `POST /setup/complete {"by": "migrated"}`. Without it, H1 keeps a field unit's hotspot up for good.
  - If Aux is down, retry every 30 s. Before each attempt, check that `state` is still `complete`, and cancel the retry in `reset`. Otherwise a reset during an Aux outage is undone by the next retry.
  - `reset` calls `DELETE /setup/complete`.
  - `muon_setup` reads `GET /setup/complete` only in the migration check. Once `muon_setup` has state of its own, that state is authoritative.
  - Use only these routes. There is no fallback to #174's `GET`/`POST /setup`: OS-7 replaces it, and a marker written there is one H1 never reads and `reset` can't clear.
- **First start with no stored state.** Run the migration check in [01-flow.md §7](01-flow.md#7-printers-already-in-the-field) before creating a `new` state.
- **Schema version.** `"version": 1`. Load an unknown version as read-only, log it, and treat it as `complete`, so a downgrade never re-runs setup. Writes to a read-only document return HTTP 409.

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
- **Order.** Before `complete`, a write for a step after the first `pending` step gets `invalid_step` (01 §3). After `complete`, only `network`, `name`, `remote` and a `ready` skip accept writes.
- **Wrapping.** Over HTTP, Moonraker wraps every answer as `{"result": {…}}`. Over websocket JSON-RPC, the answer is the JSON-RPC `result`. A notification's `params[0]` is the bare state.

### 5.1 `GET /server/muon/setup` · `server.muon.setup`

Returns the state document (§6). Every change also goes out as `notify_muon_setup_changed`.

### 5.2 `GET /server/muon/setup/options` · `server.muon.setup.options`

The optional `country` query adds that country's time zones.

```jsonc
{
  "languages": [ { "code": "en", "endonym": "English" }, { "code": "de", "endonym": "Deutsch" } ],
  "region": { "countries": ["AT","BE","CH","DE","FR","GB","IE","…"],   // Aux GET /region/options, verbatim
              "preselect": "GB", "basis": "plurality", "locked": false },
  "timezones": ["Europe/London"],                                       // only with ?country=
  "ready_manifest": { "version": 1, "items": [ … ] }
}
```

- **`region`** is Aux `GET /region/options` (MuonOS#174), passed through unchanged. `muon_setup` never keeps its own country list.
- **Picker tier 2 is built by each surface,** from MuonUI#31's `SPOKEN_IN` and `Intl.DisplayNames` (`regionCountries.ts`, ported to Fluidd). That's "countries where this language is spoken", ordered by speakers.
- **Time zones** come from the image's tzdata, `/usr/share/zoneinfo/zone.tab`: one row per country and zone, in tzdata's order, so the principal zone comes first (DE gives `Europe/Berlin`, `Europe/Busingen`). Don't use `zone1970.tab`: its rows cover several countries, so DE would start with `Europe/Zurich` and NO, SE and DK would get `Europe/Berlin`.

### 5.3 Language

`POST /server/muon/setup/language` with `{ "rev": 3, "code": "de" }`

- The `code` must be in the configured `languages`, otherwise `unsupported_language`.
- Sets `steps.language = {status: "done", value, source}` and moves `new` to `in_progress`.
- Writes Fluidd's default locale if the owner hasn't chosen one: `database.insert_item("fluidd", "uiSettings.general.locale", code)` when `get_item(...)` returns nothing. See [05-phone-setup-page.md §9](05-phone-setup-page.md#9-language).

### 5.4 Clock and time zone

`POST /server/muon/setup/clock` with `{ "epoch_ms": 1790251200000, "tz": "Europe/London" }`

- No `rev` is needed, and the call doesn't change `rev`.
- **The clock** is set through Aux only when the system clock isn't NTP-synchronised **and** it differs by more than 2 s. `epoch_ms` must be later than the image's build time, otherwise the result is `invalid_clock`.
- **The time zone**, if given and valid in tzdata, is set through Aux, even when `epoch_ms` is refused. It sets `clock.tz_source = "phone"`. Aux's 422 maps to `invalid_timezone`.
- **Only a hotspot caller** may post `clock` (§3). The phone page on `10.42.0.1` is the only surface that knows the owner's local time; everyone else gets 403.

`POST /server/muon/setup/timezone` with `{ "rev": 9, "tz": "America/Chicago" }`

- `tz` must be one of the declared country's zones. Otherwise the result is `invalid_timezone`.
- It sets `clock.tz_source = "owner"`.

### 5.5 Networks

`GET /server/muon/setup/networks?rescan=true` · `server.muon.setup.networks`

**Scanning.** It scans synchronously, with a 15 s cap, through Aux `GET /wifi/scan?rescan=true`. It then normalises the result:

```json
{
  "ok": true,
  "scanned_at": 1790251262.4,
  "ethernet": { "present": true, "carrier": false, "address": null },
  "networks": [
    { "ssid": "HomeWiFi", "security": "wpa2", "signal": 78, "band": "2.4", "channel": 6,
      "bssids": 2, "saved": false, "in_use": false, "supported": true, "channel_permitted": true }
  ]
}
```

**Normalising:**

- Drop empty SSIDs; hidden networks go through "Other network…".
- Drop the printer's own hotspot SSID.
- Merge entries with the same SSID. Keep the strongest signal, count the BSSIDs, and report the band and channel of the strongest.
- Sort by signal, strongest first.

**Fields:**

- **`security`** is one of `open`, `owe`, `wep`, `wpa2`, `wpa3`, `wpa2_wpa3`, `enterprise`. `wep` has `supported: false`.
- **`saved`** comes from Aux `GET /wifi/saved` (MuonOS#210). Until that lands, it's `false`.
- **`channel_permitted`** is `true` when `state.region.channels` is empty (no configuration applied yet) or contains the network's channel.
  - Surfaces decide what to do with a network that isn't permitted: they run MuonUI#31's `regionPromptFor(state.region, channel)`, which gives `join`, `offer-switch` or `locked`.
  - **No network carries a country suggestion before joining.** Aux doesn't expose each network's country element, so the region is confirmed after the join (§5.6a).

### 5.6 Joining a network

`POST /server/muon/setup/network`

```jsonc
{ "rev": 7, "kind": "wifi",
  "ssid": "HomeWiFi", "hidden": false, "security": "wpa2",
  "psk": "correct horse battery staple",
  "region": null,                 // set only when the owner accepted "Change your printer's region?" before joining
  "eap": null }
// Enterprise:
{ "rev": 7, "kind": "wifi", "ssid": "eduroam", "hidden": false, "security": "enterprise",
  "eap": { "method": "peap", "phase2": "mschapv2", "identity": "ab123@uni.ac.uk",
           "anonymous_identity": "anonymous@uni.ac.uk", "password": "…",
           "ca_cert_id": "c_1f2e", "domain_suffix_match": "uni.ac.uk", "no_ca_check": false } }
// Ethernet already has an address, and the owner confirms it:
{ "rev": 7, "kind": "ethernet" }
```

**Behaviour:**

1. **Validate the request.**
   - `ssid` must be 1–32 bytes UTF-8.
   - `region`, if given, must be in `options.region.countries` (`region_not_offered`).
   - `hidden` and `security` hints need new Aux work (03 §4, W1). Until that lands, `hidden: true` returns `unsupported_security`.
   - A `psk` must be 8–63 characters, or exactly 64 hex characters. `open` and `owe` take no psk.
   - For `enterprise`, `method` must be `peap` or `ttls`, `phase2` must be `mschapv2` or `pap`, and `identity` and `password` are required. Either `ca_cert_id` or `no_ca_check: true` is required.
   - Any violation gives `invalid_network`, with `detail.field` naming the field.
2. **Start the operation.** The HTTP response comes back immediately with `ok: true` and the state showing `op`.
3. **Switch the region first, if asked.** When `region` is set, run the region apply in §5.6a before joining. Stop there if it fails.
4. **Join through Aux** ([03-printer-os.md §4](03-printer-os.md#4-wi-fi-join)). Set `op = {kind: "join", phase: "saving"}` and move `op.phase` through `saving` → `associating` → `authenticating` → `dhcp` → `internet_check` → `update_check`, notifying on each change. Treat Aux's `200 {"status": "restored"}` and any 400 as a failed join, and take the reason from `state_reason`'s numeric prefix. At `update_check`, refresh the clock state first, then call Aux `POST /update/check {"wait": true}` with a time limit of about 20 s (or poll `GET /update/status` until the check finishes), and read `update_available` and `target_version`. `{"wait": false}` only schedules a check, and the answer read straight after it is the previous one.
5. **Success means an IPv4 address on the uplink.** Set:
   - `network = {kind, ssid, addresses: [...], hostname_local: "<hostname>.local", internet: true|false|"portal", error: null, region_confirmed}`;
   - `region_confirmed` is `true` when the market is `locked` or `none`, or when `declared_country` is already set and equals `detected_country`. Otherwise it's `false`.
   - `network.status` becomes `done` only when `region_confirmed` is `true`. Until then the cursor stays on `network`, and surfaces show the region line (§5.6a).
   - `update.status` to `pending` only if the internet check passed, an update exists and the clock is synced. Otherwise it is `hidden`.
6. **Failure.** `network.error = {code, at_phase}`, the step stays `pending`, `op` is cleared, and nothing is kept: the Aux connection profile is deleted. The codes are `wrong_password`, `ssid_not_found`, `no_address`, `timeout`, `eap_failed`, `cert_invalid` and `unsupported_security`.
7. **`internet: "portal"`** (a guest network with a web sign-in page) is still `done`, because the address works on the LAN. It carries `error: {code: "portal_required"}` so the surfaces can warn and offer "Choose another network".
8. **Never store or echo secrets.** `psk` and `eap.password` are never persisted by `muon_setup`, never logged (redact them in any debug dump), and never put in state or events. Only NetworkManager keeps them.

#### 5.6a Region confirmation

This follows KAN-321 Rev 11 and MuonUI#31: the region comes from the network the printer has joined.

`POST /server/muon/setup/region` with `{ "rev": 8, "country": "GB" }`.

**What happens:**

1. **Validate.** `country` must be in `options.region.countries`, otherwise `region_not_offered`.
2. **Apply.** Set `op = {kind: "region_apply"}` and call Aux `POST /region/country {country}`.
   - Use a 60 s timeout; `aux_api_proxy`'s `post()` needs a timeout parameter for this.
   - The apply is live and takes about 8 s. Every Wi-Fi profile goes down and comes back, **the hotspot included**.
   - Aux records the declaration only when the apply succeeds.
3. **Map failures** from the `detail.code` values OS-2 adds (the region agent's `OUTCOMES` names):

   | Aux code | `muon_setup` code |
   |---|---|
   | `country-not-in-token` | `region_not_offered` |
   | `no-token`, `unreadable-token`, `bad-token-format`, `bad-signature`, `unknown-serial`, `serial-mismatch`, `no-signing-key` | `needs_reregistration` |
   | `apply-failed`, `intersected`, `readback-mismatch` | `region_apply_failed` |
   | `busy` | `region_busy` |
   | 504 | Re-read `GET /region` before deciding |

   On failure nothing is declared, `network.region_confirmed` stays `false`, and `network.region_error = {code, message}` records the failure. The apply drops the hotspot for about 8 s, so a phone usually misses the write's HTTP answer and learns the result from the state. The next region apply or join clears `region_error`.
4. **Success.** Re-read `GET /region`, set `network.region_confirmed = true`, `network.region_error = null` and `network.status = done`. If the declared country has exactly one time zone and `clock.tz_source` isn't `phone` or `owner`, set that zone (`tz_source: "region"`). Then move the cursor on.

**Also:**

- **Called while `network` is still pending**, the endpoint only applies the region; `region_confirmed` is set by the join. This is the "switch before joining" case.
- **Until OS-2 lands** stable codes, Aux returns free text. Map it by substring, and fall back to `region_apply_failed`.

`POST /server/muon/setup/network/cancel` with `{}`: cancels a running join. The partial profile is removed.

`POST /server/muon/setup/network/ca_cert` takes a multipart upload of one PEM or DER file of at most 16 KiB.

- It returns `{ok: true, ca_cert_id}`.
- The certificate is handed to Aux to store root-only. `muon_setup` only keeps the ID.
- Allowed callers are panel, hotspot and lan.

### 5.7 Name

`POST /server/muon/setup/name` with `{ "rev": 9, "name": "Walnut" }`. An empty string or leaving `name` out means "Keep".

- Refactor `aux_api_proxy._set_identity_name_handler` into a public `async set_friendly_name(name) -> dict` and `async get_identity() -> dict`. Both the existing endpoint and `muon_setup` call these.
- The rules are unchanged: the name is stripped, at most 32 characters (counted in code points), and stored in the `muon` namespace under `friendly_name`. Also reject control characters (C0 and C1, newlines included). The surfaces cap input at 32, so the 400 only follows a client bug.
- It's stored in Moonraker's database, so it works while Aux is down.
- The hostname and hotspot SSID keep the derived name (ID-2).
- Marks the step `done`, with `value` set to the effective name. **Keep** never clears an earlier rename: its value is the stored name, or the derived one if there is none. Going back to the derived name is a Settings action (`/server/muon/identity/name` with `""`), not a setup one.

### 5.8 Update

`POST /server/muon/setup/update` with `{ "rev": 10, "action": "install" | "later" }`

- **`later`** marks the step `skipped`.
- **`install`** runs these steps:
  1. Store `op = {kind: "update_install", target: "<version>"}`.
  2. Start the install through Aux `POST /update/install` (`aux_api_proxy.ota_start()`), or through update_manager's `MuonOS` updater, which calls the same route. The panel's `UpdatingOverlay` (KAN-215) follows Aux `/update/status` whoever starts the install.
  3. Mirror Aux `GET /update/status` progress into `op.progress` (0–1) until Aux's `state` is `rebooting` or `failed`. `OtaDeploy.update()` can return early (after 300 s without progress, or on losing contact) while the install is still running, so keep following Aux after it returns.
  4. The printer reboots into the new slot.
- **After a boot**, read Aux `GET /update/status`. While its `state` is `installing` or `rebooting`, keep the `op` and keep mirroring. Judge only once it's `idle`, `commit_pending` or `failed`, by comparing `current_version` with `op.target`:
  - equal: the step is `done`;
  - otherwise it rolled back: the step is `pending` with `error: {code: "update_failed"}`.
- **Never judge from update_manager's cached version.** It reads `?` until a refresh, and the M1 refreshes weekly.
- `muon_setup` doesn't commit the update itself. It follows the existing OTA commit policy (KAN-358).
- **Aux refusals** are 409s with `detail.code`: `printer_busy` (printing, paused or busy; KAN-75) maps to `printer_busy`; `busy` and `invalid_state` (the running system needs a commit first) map to `update_failed`.

### 5.9 Remote access

`POST /server/muon/setup/remote`

```jsonc
{ "rev": 11, "mode": "local" }    // done at once; no outbound connection is made
{ "rev": 11, "mode": "cloud" }    // starts the account link
{ "rev": 11, "mode": "later" }    // skipped; the Finish setup card appears
```

**`self_hosted` isn't accepted in phase 1.** muon-link configures an orchestrator only through environment variables and a restart: `MUON_LINK_ORCH_ID` and `MUON_LINK_RELAY_URL`, optionally `MUON_LINK_ORCH_ADDR`. No route can set them at runtime, and the self-hosted orchestrator in the muon-link repo doesn't speak `muon/orch/1` yet. `capabilities.self_hosted` therefore stays `false`, and the surfaces don't offer "My own server".

**`cloud` runs the account link** (NET-10(d), ADR 0018, muon-link#24) through the `muon_link` component (§9):

1. **Start.** Call `muon_link.start()`, which forwards to muon-link's `POST /link/start`. It returns `{"phase":"connecting"}` at once. The code arrives later, from the orchestrator, not from muon-link.
2. **Mirror.** Copy `muon_link`'s phase object into `remote.link` **unchanged** on every `muon_link:link_changed` event. The shapes are muon-link's `LinkPhase`:

   ```jsonc
   {"phase":"unavailable"}                                        // no orchestrator configured
   {"phase":"unlinked"}
   {"phase":"connecting"}
   {"phase":"code","code":"482913","expires_at":1790251620,"url":"https://app.muon3d.com/link?code=482913"}
   {"phase":"offer","account":"jed@example.com","authority":"Muon3D","fingerprint":"9f3c1a7be2d04c11"}
   {"phase":"linked","account":"jed@example.com","connected":true}
   {"phase":"failed","message":"could not reach the Muon3D service: …"}
   ```

   - `expires_at` is an integer in Unix seconds.
   - `url` is the QR content. The orchestrator sets it; the surfaces don't build it.
   - The code's length is the orchestrator's choice. It is digits only.
3. **Code renewal.** muon-link never expires a code itself.
   - While `cursor == "remote"`, the mode is `cloud` and the phase is `code`, `muon_setup` calls `start()` again once the synced clock passes `expires_at`.
   - It also retries once on `failed`.
   - It **never** calls `start()` during `offer`, because that silently drops the pending offer.
4. **Linked.** The phase becomes `linked` the moment the owner confirms on the panel. The panel calls muon-link directly (§9); Moonraker never forwards confirm. The step is then `done`, with `account = linked.account`. `connected` may briefly be `false`, and grants only work once signed time arrives. The surfaces don't wait for it.
5. **Already linked.** If `start()` returns 409 "already linked", read `GET /link` and treat the step as `done` with that account.
6. **Failure.**
   - A 503 from `muon_link` (muon-link isn't answering) or `phase: "unavailable"` gives `link_unavailable`.
   - `phase: "failed"` after the one retry gives `link_failed`, with `detail.message`.
   - In both cases the step stays `pending`.
7. **Blocked.** If the clock isn't synced, or `network.internet` isn't `true`, `cloud` returns `ok: false` with `clock_unsynced` or `no_internet`. The orchestrator makes the code, so no code can be issued on the hotspot alone.

`POST /server/muon/setup/remote/cancel` with `{}`: stops renewal and calls `muon_link.cancel()`, which forwards to `/link/cancel`. That declines a pending offer and drops the orchestrator connection. It does nothing once the printer is linked.

`capabilities.cloud_link` comes from `GET /server/muon/link`: it's `false` on a 503 or on `phase: "unavailable"`. Don't use muon-link's `/status.orchestrator_configured`, which only reflects the legacy `MUON_LINK_ORCHESTRATOR_URL`.

### 5.10 Ready to print

`POST /server/muon/setup/ready` with `{ "rev": 14, "item": "self_test", "action": "start" | "confirm" | "skip" }`

- `confirm` is for `kind: "confirm"` items and for `kind: "panel_flow"` items once MuonUI has finished that flow.
- `start` is for `kind: "macro"` items. It runs `klippy_apis.run_gcode(item.macro)` under `op = {kind: "ready_item", item: "<id>", phase: "running", progress: null}`.
  - **It moves the printer, so it refuses unless all of these hold:**
    - every earlier `required` item in manifest order is `done` (otherwise `invalid_step`), so the self-test can't run with the transport clips on;
    - Klipper defines `gcode_macro <macro>` at that moment (otherwise `invalid_step`, or `printer_not_ready` if Klipper can't say);
    - print_stats `state` isn't `printing` or `paused` (`printer_busy`);
    - Klippy is ready (`printer_not_ready`).
  - The item is `done` if the macro returns without error.
  - Otherwise it is `failed`, with `error: {code: "self_test_failed", message, detail: {gcode_error: "<Klipper's text>"}}`.
  - There's no cancel. The panel's emergency stop fails the item.
- `start` and `confirm` are panel-only (§3). Check the caller first, before `busy`, `stale_rev` or any Klipper query, so a phone always gets 403.
- `skip` marks one item `skipped`. That doesn't count as done, so a skipped required item keeps the step pending.
- When every `required: true` item is `done`, the step is `done`. A re-run never takes a `done` step back to `pending`.
- `POST /server/muon/setup/skip` with `{"step": "ready"}` skips the whole step.
- **After `complete`,** a skipped `ready` stays `skipped` while the owner works through it from the "Finish setup" card, and becomes `done` once its required items are. It never becomes `pending`, which would drop it off the card.

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
  - Until the macro exists in the Klipper config, `muon_setup` hides **pending** `macro` items whose macro isn't defined, and doesn't count them as required. Check with `klippy_apis` for `gcode_macro <name>` when Klippy becomes ready, when the state loads, and before each ready action. With the default manifest and no `MUON_SELF_TEST`, `ready` is done once the clips are confirmed. That is intended until OS-8 ships the macro, and OS-8 and QA-2 check the release image either defines it or someone decided it shouldn't. A macro that appears later makes its item `pending` again, but doesn't reopen a `done` step.
  - Validate the manifest against the schema in `tests/test_muon_setup.py`.

### 5.11 Navigation and control

| Endpoint | Body | Effect |
|---|---|---|
| `POST /server/muon/setup/driver` | `{ "rev": n, "kind": "panel"\|"phone"\|"web"\|"app", "client_id": "<uuid>" }` | Claims the driver. Calling it again with the same `client_id` renews the claim. It does **not** change `rev`, so the phone can renew every 10 s without causing `stale_rev`. `app` is the Muon3D phone app, on the hotspot or the LAN. The kind only chooses the panel's words (04 P8). Permissions come from the caller kind (§3), never from the driver kind. |
| `POST /server/muon/setup/goto` | `{ "rev": n, "step": "network" }` | Moves the cursor ([01-flow.md §3](01-flow.md#3-state-model)). An invalid target gives `invalid_step`. |
| `POST /server/muon/setup/skip` | `{ "rev": n, "step": "network"\|"update"\|"remote"\|"ready" }` | Marks the step `skipped`. `language` gives `not_skippable`. |
| `POST /server/muon/setup/finish` | `{ "rev": n }` | Requires `language` to be `done`, otherwise `required_steps_pending`. Remaining `pending` optional steps become `skipped`. Then `state = complete`, the marker is written, `muon_setup:complete` fires, and the hotspot auto-off is scheduled, in that order ([03-printer-os.md §1](03-printer-os.md#1-hotspot-lifecycle)). After `complete`, a successful `network` join schedules the auto-off again (E5). |
| `POST /server/muon/setup/card/dismiss` | `{ "rev": n }` | `card_dismissed = true` |
| `POST /server/muon/setup/reset` | `{}` | Panel only and floored. Cancels any pending marker retry (§4), clears the `muon_setup` namespace, calls Aux `DELETE /setup/complete`, and starts over as `new`. **`rev` keeps increasing:** the new document's `rev` is the old one + 1, never 1, so screens that hold the old state accept it. If the `DELETE` fails, the reset still happens, and the `DELETE` is retried every 30 s for as long as `state` isn't `complete`. For development and support; a real factory reset clears everything (ADR 0005). |

## 6. State document

```jsonc
{
  "version": 1,
  "rev": 42,                                // +1 on every change except driver renewals and lapses
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
  "region": { "market": "picker",                                   // derived: none | locked | picker
              "reason": "…", "explanation": "…", "declared_country": "GB", "configuration": "gb",
              "domain": "GB", "surroundings": "settled", "detected_country": "GB",
              "basis": "joined-network", "enforcement": "unknown",
              "locked": false, "channels": [1,2,3,4,5,6,7,8,9,10,11,12,13,36,40,44,48] },
                                                                    // Aux GET /region verbatim + market; computed
  "capabilities": { "ethernet": true, "enterprise": true, "cloud_link": true,
                    "self_hosted": false, "bluetooth": false },       // computed
  "card_dismissed": false,
  "steps": {
    "language": { "status": "done",    "value": "en", "source": "panel" },
    "network":  { "status": "pending", "kind": null, "ssid": null, "addresses": [],
                  "hostname_local": null, "internet": null, "error": null,
                  "region_confirmed": false, "region_error": null },
    "name":     { "status": "pending", "value": "Walnut", "derived": "walnut" },
    "update":   { "status": "hidden",  "current": "1.3.2", "available": null, "error": null },
    "remote":   { "status": "pending", "mode": null, "link": null, "error": null },  // mode: local | cloud | later | null
    "ready":    { "status": "pending",
                  "items": [ { "id": "transport_clips", "status": "pending", "error": null },
                             { "id": "self_test",       "status": "pending", "error": null },
                             { "id": "load_filament",   "status": "pending", "error": null } ] }
  }
}
```

- `source` is the caller kind (§3), or `migrated`.
- `hotspot` comes from Aux `GET /wifi/ap/stations` (`{up, count, auto_off_at}`; 03 §1). `hotspot.clients` is its `count`, the number of associated stations on `ap0`, and `hotspot.auto_off_at` is its `auto_off_at`: only the OS knows the deadline after a reboot or an owner's "off", so don't store your own. The panel uses `clients` to switch its QR code ([04-panel.md §P2](04-panel.md#p2--here-or-on-a-phone-a)). Poll Aux every 2 s while `state != complete` and the hotspot is up, and notify only when something changes. Until OS-5 is on the image, fall back to `POST /wifi/ap/count`.
- `driver.kind` is `panel`, `phone`, `web` or `app`, and `bluetooth` in phase 2.
- `steps.ready.items[].status` is `pending | done | failed | skipped | hidden`. `error` is `null`, `self_test_failed` or `interrupted` (`{code, at_phase}`).
- For `op.kind == "ready_item"`, `op.item` is the item's `id`.
- `driver.lapsed` is computed on read as `now - renewed > driver_lease`.
- **Renewals aren't announced.** When a claim lapses, `muon_setup` sends `notify_muon_setup_changed` once, without changing `rev`, so the panel can show P8's lapse line.
- **`region`** carries every field of Aux `GET /region`, including `explanation` (free English text for logs; never parse or show it) and `enforcement`.
- **Clients and `rev`.** A `GET` result, including the one after a reconnect, replaces the held state whatever its `rev`. `rev` ordering applies only between notifications and write results: ignore one whose `rev` is lower than the held state's.

## 7. Identity and discovery changes in `aux_api_proxy`

- `GET /server/muon/identity` returns `name, source, derived_name, suffix, display, ssid, fingerprint, endpoint_id, setup`.
  - `endpoint_id` is the Iroh EndpointId, or `null` (KAN-403, Moonraker#27 and MuonOS#316).
  - `setup` is `new | in_progress | complete` from `muon_setup`, or `null` when `muon_setup` isn't configured. It reads `complete` while `muon_setup`'s boot check runs (01 §7). Apps use it to route a found printer ([06-add-printer.md](06-add-printer.md)).
- `identity`, `identity/name` and `dev_mode` are registered in `__init__`, not after `_fetch_spec()` (MR-9). They answer 503 while Aux isn't answering (connection refused or timeout: HTTP 599, or 500 caused by a connection error). Aux's own errors keep their status.

## 8. Tests (minimum)

Add `tests/test_muon_setup.py` with fakes for `database`, `aux_api_proxy` and `klippy_apis`, covering:

1. A fresh start with no marker gives `new` and cursor `language`. A fresh start gives `complete` with `source: "migrated"` when the marker exists, or when a saved Wi-Fi profile exists other than the hotspot and the dev image's baked `Muon3D_Dev` profile.
2. The order rules:
   - `goto` forward past the first pending step gives `invalid_step`.
   - Skipping `language` gives `not_skippable`.
3. `rev`:
   - A write with an old `rev` gives `stale_rev` plus the current state.
   - A driver renewal doesn't change `rev`.
   - A driver claim with `kind: "app"` from a hotspot caller and from a LAN caller is stored as `app`, and gets the same access as `phone` and `web` from that caller.
4. `op`: a second write during a `join` gives `busy`, and `network/cancel` clears it.
5. Join mapping: every Aux failure reason in [03-printer-os.md §4](03-printer-os.md#4-wi-fi-join) maps to the right `code`, and on failure the profile-delete call is made.
6. Secrets: after a join, no `psk` or password appears in the stored state, the events or the log records (use `caplog`).
7. Access: build the §3 table as a parametrised test over caller kinds × endpoints.
8. `finish` marks pending optional steps as `skipped`, writes the marker, fires `muon_setup:complete`, and asks Aux for the hotspot auto-off.
9. Region:
   - After a join in a `picker` market, `network` stays `pending` with `region_confirmed: false`. `POST region` then declares the country and marks the step done.
   - A join with `region` set calls `POST /region/country` before `/wifi/connect`.
   - Aux `status: "restored"` counts as a failed join.
   - Aux `busy`, a read-back mismatch, a country not offered and a missing token each map to their codes, and none of them declares anything.
   - A channel not permitted after the apply gives `channel_not_permitted`.
10. Markets: `locked` and `none` set `region_confirmed: true` on join. `none` is derived from `countries == []`, and `locked` from `locked == true`.
11. The update step: a version that matches the target after the reboot is `done`, and a rollback gives `update_failed`.
12. Remote `cloud`:
    - codes renew once the clock passes `expires_at`, and stop after `goto` away or `remote/cancel`;
    - `start()` is never called during `offer`;
    - a 503 or `unavailable` gives `link_unavailable`;
    - a 409 "already linked" gives `done`.
13. `aux_unavailable`: the component loads and every Aux-backed call returns the code. The GET still works.

## 9. The `muon_link` component (Moonraker PR #20)

[Muon-3D/Moonraker#20](https://github.com/Muon-3D/Moonraker/pull/20) (open) already adds `moonraker/components/muon_link.py` with `[muon_link] admin_address: 127.0.0.1:7131`. It forwards exactly three routes to muon-link's loopback admin endpoint. **Build on it. Don't write a second component.**

| Endpoint | Forwards to | Notes |
|---|---|---|
| `GET /server/muon/link` | `GET /link` | Returns the `LinkPhase` (§5.9). A 503 means muon-link isn't answering, and a 409 carries muon-link's `{"error"}` sentence. |
| `POST /server/muon/link/start` | `POST /link/start` | Returns `connecting` at once. Fluidd's `startLanLink()` already polls until it sees `code`. |
| `POST /server/muon/link/cancel` | `POST /link/cancel` | |

**Confirm and unlink are never forwarded.** PR #20's tests assert this, following ADR 0018 and LINK-3. The panel calls muon-link's `POST /link/confirm`, `/link/cancel` and `/link/unlink` directly, through an nginx `/muon-link/` location on the loopback-only `:100` vhost (OS-10). Nothing is added to `FLOOR_PREFIXES` for linking.

**MR-6 adds to PR #20**, after it merges or as a follow-up PR:

1. **Polling and events.** muon-link pushes nothing, so `muon_link` polls `GET /link`: every 1 s while the phase is `connecting`, `code` or `offer`, and every 30 s otherwise, stopping when nothing is subscribed. It emits `muon_link:link_changed` (clients receive `notify_link_changed`) when the phase object changes.
2. **Python methods.** Add public async `status()`, `start()` and `cancel()` for `muon_setup` to call, so it doesn't make HTTP calls to itself.
3. **Rate limit.** Allow at most 5 `start` calls per minute per caller IP. muon-link has no limit of its own.
4. **Tests.** Extend `tests/test_muon_link.py` to cover polling cadence, change detection and the rate limit.

