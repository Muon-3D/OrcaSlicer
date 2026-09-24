# 08 · Error catalogue

Every domain error from `muon_setup` is `{ ok: false, error: { code, message, detail } }` ([02-setup-api.md §5](02-setup-api.md#5-endpoints)). Surfaces show **their own translated copy** keyed by `code`, never `message`:

- Fluidd: `app.muon.setup.error.<code>`;
- MuonUI: `setup.error.<code>`.

The copy below is the English source text. `{ssid}`, `{name}` and the other braces are named placeholders; don't concatenate strings.

Every error screen has **one primary action**. On the panel that is the press; on the phone it's the filled button. Secondary actions are optional.

## 1. Network

| Code | When | Panel copy | Phone copy | Primary / secondary |
|---|---|---|---|---|
| `wrong_password` | The join failed at authentication (PSK) | "{ssid} didn't accept that password." | Same | Try again (the password field opens again with the text visible on the phone) / Choose another network |
| `ssid_not_found` | The network isn't visible at join time | "{name} can't see {ssid} from here." | Same | Scan again / Choose another network |
| `no_address` | Associated, but DHCP failed | "{name} joined {ssid}, but the router didn't give it an address." | Same, plus "Restarting the router often fixes this." | Try again / Choose another network |
| `timeout` | No result within 45 s | "Joining {ssid} took too long." | Same | Try again / Choose another network |
| `eap_failed` | 802.1X authentication failed | "{ssid} didn't accept that username or password." | Same | Try again / Choose another network |
| `cert_invalid` | The CA certificate is unreadable, or the server certificate failed validation | "{ssid}'s certificate couldn't be checked." | Same, plus "Check the certificate file and domain." | Try again / Choose another network |
| `unsupported_security` | WEP, or a mode Aux can't do | "{name} can't join this kind of network." | Same | Choose another network |
| `invalid_network` | Validation failed (`detail.field`) | – (the panel can't send invalid input) | Inline under the field: "Check this field." | – |
| `portal_required` *(warning; the step is done)* | The internet check was redirected | "{ssid} needs a sign-in page, which a printer can't complete." | Same, plus "Printing over this network works, but updates and remote access won't." | Choose another network / Continue anyway |
| `no_internet` *(warning; the step is done)* | The internet check failed | "Connected to {ssid}. It has no internet, so updates and remote access are off." | Same | Continue |

## 2. Region, time zone and clock

| Code | When | Copy (both surfaces) | Primary / secondary |
|---|---|---|---|
| `region_required` | Joining in a `picker` market with no country declared or chosen | "Choose the printer's region first." | Choose region |
| `region_not_offered` | The country isn't in the token | "This printer is not registered for that country." (Rev 11) | Choose another region |
| `region_apply_failed` | The Aux read-back gate failed | "We could not apply that region." (Rev 11) | Try again / Choose another region. After 2 failures: Save a diagnostic bundle (KAN-378) |
| `region_busy` | Aux reports `busy` | "The Wi-Fi radio is busy. Trying again…" (retried automatically once after 3 s) | Try again |
| `needs_reregistration` | No valid market token | "This printer needs re-registering. Support code: {code}" (Rev 11) | Continue (channels 1–11 still work) |
| `channel_not_permitted` | The chosen network's channel isn't allowed | "That network is on a channel your printer is not set for." A locked unit adds "This printer is set for the United States. Contact support." | Choose another network |
| `invalid_timezone` | The zone isn't valid for the country | – (the lists only offer valid ones) | – |
| `invalid_clock` | `epoch_ms` is before the image build | – (silent; logged) | – |
| `clock_unsynced` | `update` or `remote: cloud` without a synced clock | "{name} hasn't got the time from the internet yet." | Try again (after 5 s) / Continue without |

## 3. Update and remote access

| Code | When | Copy | Primary / secondary |
|---|---|---|---|
| `update_failed` | Rolled back after the reboot | "The update didn't finish. {name} is still on {version}." | Continue |
| `printer_busy` | A print is running | "{name} is printing. Try again when it finishes." | OK |
| `link_unavailable` | muon-link isn't reachable | "Linking isn't available right now." | Keep it on my network / Try again |
| `link_failed` | muon-link reports `failed` | "Linking didn't work." plus `detail.message` if present | Try again / Keep it on my network |
| `link_expired` | The code expired and wasn't renewed, e.g. because the owner left the screen | "That code has expired." | Show a new code |
| `link_declined` *(cloud side)* | Cancel was chosen at the knob | "Linking was cancelled on the printer." | Try again / Keep it on my network |

## 4. Ready to print

| Code | When | Copy | Primary / secondary |
|---|---|---|---|
| `self_test_failed` | The macro returned an error | "The self-test found a problem: {detail}" | Try again / Skip |
| `printer_not_ready` | Klippy isn't ready | "The printer isn't ready yet. Try again in a moment." | Try again |

## 5. Flow and transport

| Code | When | What the surface does |
|---|---|---|
| `stale_rev` | Another screen changed the state | Re-render from the returned state. Inline note: "This step changed on {name}'s screen." (phone) or "Changed from a phone" (panel). No retry. |
| `busy` | An `op` is running | "{name} is busy with another step. One moment…" Wait for the state. |
| `invalid_step` / `not_skippable` / `required_steps_pending` | Client bug | Log it, re-render from the state, and show nothing to the owner. |
| `unsupported_language` | Client bug | Fall back to `en`. |
| `interrupted` | An `op` was cut short by a power loss | "That step was interrupted. Try again." |
| `aux_unavailable` | The Aux API isn't answering | "{name} is still starting up…" Retry every 3 s. After 60 s: "Something's wrong. Restart the printer." |
| *(no response)* | A network error or timeout on a write | **Not an error.** "Reconnecting to {name}…" Wait for the state ([05-phone-setup-page.md §4](05-phone-setup-page.md#4-client-and-reconnect-rules)). After 30 s on the phone: "Check the printer's screen." |
