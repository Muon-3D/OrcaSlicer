# Fixtures

These are example payloads that follow [02-setup-api.md](../02-setup-api.md). UI work (UI-*, FL-*) can build and test against them before `muon_setup` lands. The Fluidd E2E mock server (`tests/e2e/mock-setup-server.mjs`) replays them.

They are **illustrative**:

- The country lists are samples, not the certified list. The real data comes from Aux `/region` and `/region/options` (MuonOS#174).
- The fingerprint is a placeholder.
- No fixture may ever contain a real password or hotspot key.

| File | Scenario |
|---|---|
| `state.01-new.json` | Factory-fresh. The cursor is at `language`, and the EU token's fallback region (DE) is applied but not declared. |
| `state.02-network-phone-driver.json` | The language is done, a phone has claimed the driver, and the page has posted the clock and time zone. The cursor is at `network`. |
| `state.04-joining.json` | `op.kind = join`, `phase = dhcp`. |
| `state.04b-region-confirm.json` | Joined HomeWiFi. `region.detected_country = GB` (`basis: joined-network`) and `network.region_confirmed = false`, so the region line shows. |
| `state.04c-region-apply.json` | Confirm was pressed. `op.kind = region_apply`, and Wi-Fi and the hotspot drop for about 8 s. |
| `state.05-network-wrong-password.json` | The join failed with `wrong_password` at `authenticating`. |
| `state.06-remote-link-code.json` | The network is done with internet, the update is skipped, and remote is `cloud` with a live code. |
| `state.07-ready-self-test.json` | Linked. The panel drives `ready`, and `self_test` is running. |
| `state.08-complete-with-skips.json` | Complete, with `remote` and `ready` skipped, so the "Finish setup" card shows. The hotspot auto-off is scheduled. |
| `result.stale-rev.json` | A write response that was rejected with `stale_rev`, carrying the current state. |
| `options.picker-eu.json` | Aux `GET /region/options` for a European-tokened unit (`preselect: GB`, `basis: plurality`). |
| `options.locked-us.json` | A US unit (token `configs: ["us"]`). There is no picker. |
| `options.none.json` | No valid token (`countries: []`), which is every unit today. |
| `networks.eu-unit.json` | Scan results: a network on ch 13, an Enterprise network, and an unsupported WEP network. |
| `networks.us-locked-unit.json` | The same air as seen by a US unit. The ch 13 network has `channel_permitted: false`. |
