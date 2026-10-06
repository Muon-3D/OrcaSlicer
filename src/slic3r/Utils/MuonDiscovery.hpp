#ifndef slic3r_MuonDiscovery_hpp_
#define slic3r_MuonDiscovery_hpp_

// How the printer search recognises Muon M1 printers and what it shows for them.
// Plain strings in and out, no wxWidgets, so the rules can be unit tested.

#include <optional>
#include <string>
#include <unordered_map>

namespace Slic3r {
namespace MuonDiscovery {

using TxtData = std::unordered_map<std::string, std::string>;

// True when a Bonjour (_octoprint._tcp) reply comes from a Muon printer: Moonraker
// on an M1 publishes model "Muon M1" (octoprint_model_name), and newer images add
// the "setup" key. Other services are left alone.
bool is_muon_service(const TxtData &txt);

// The address to fill in for a Muon printer. The service record names Moonraker's
// own port (7125), which listens on loopback only; the network reaches the printer
// through nginx on port 80, so the address carries no port. `host` is the host name
// or the IP address the reply came from; an IPv6 address is bracketed.
std::string muon_address(const std::string &host, bool host_is_ipv6, const std::string &path);

// The short name an M1 is known by on every surface (host name, hotspot SSID,
// Bluetooth local name), lower case: "Muon-walnut-8987.local", "Muon-walnut-8987"
// and "walnut-8987" all give "walnut-8987". Empty when `name` is not an M1's.
std::string short_name(const std::string &name);

// How the printer's screen shows a short name: "walnut-8987" is "Walnut · 8987".
// Empty when `name` is not an M1's.
std::string display_name(const std::string &name);

// The address an M1 has on its own Wi-Fi hotspot (Muon-<name>), where an
// unset printer waits to be set up.
constexpr const char *HOTSPOT_ADDRESS = "10.42.0.1";
// Moonraker's identity route. Readable without signing in, so a printer that
// cannot tell the network its name can still be listed (ID-4).
constexpr const char *IDENTITY_PATH = "/server/muon/identity";

// What GET /server/muon/identity says about a printer.
struct Identity
{
    std::string short_name;   // "walnut-8987", for matching the same printer found another way
    std::string display;      // "Walnut · 8987" (the owner's rename, if any, is in `name`)
    std::string name;         // what the owner calls it
    std::string setup;        // "new", "in_progress", "complete", or empty when not reported
    std::string endpoint_id;  // the Iroh EndpointId, or empty
};

// Parses Moonraker's answer ({"result": {...}}). Empty unless it answers like an M1:
// it has to name the printer.
std::optional<Identity> parse_identity(const std::string &body);

// muon_setup's state, as the "setup" TXT key and the identity route report it.
inline bool setup_pending(const std::string &state) { return state == "new" || state == "in_progress"; }
inline bool setup_complete(const std::string &state) { return state == "complete"; }

} // namespace MuonDiscovery
} // namespace Slic3r

#endif
