#ifndef slic3r_MuonBleScan_hpp_
#define slic3r_MuonBleScan_hpp_

// Hears Muon M1 printers advertising over Bluetooth LE (Iroh_BLE, KAN-432).
// Detection only: nothing here connects to a printer.
//
// An M1 advertises its GATT service UUID in the advertising packet and puts its
// local name ("walnut-8987") and ten bytes of manufacturer data in the scan
// response, so the platform scanners scan actively (they ask for the scan
// response) but never connect.
//
// This header has no platform or wx dependency. Each platform implements
// create_scanner() in its own file; elsewhere it returns nullptr.

#include <cstdint>
#include <functional>
#include <map>
#include <memory>
#include <optional>
#include <string>
#include <vector>

namespace Slic3r {
namespace MuonBle {

// Iroh_BLE blelink-proto consts.rs: SERVICE_UUID and MANUFACTURER_ID_DEV (the
// SIG's test id, used until Muon3D has its own).
constexpr const char *SERVICE_UUID = "843ce1e6-c64c-4ac1-b06b-5afd8329b7f9";
constexpr uint16_t    COMPANY_ID   = 0xffff;

// The manufacturer data under COMPANY_ID (blelink-proto adv.rs AdvData):
// byte 0 version, byte 1 flags (bit 0 unclaimed, bit 1 busy), bytes 2..10 the
// first 8 bytes of the printer's Iroh EndpointId.
struct Advert
{
    bool        unclaimed { false };  // not set up yet
    bool        busy { false };       // another device holds its one Bluetooth slot
    std::string endpoint_prefix;      // 16 lower case hex digits
};

// Empty for a short value or a version this code does not know.
std::optional<Advert> decode_manufacturer_data(const std::vector<uint8_t> &data);

// A printer heard nearby.
struct Printer
{
    std::string id;               // the platform's id for the device
    std::string short_name;       // "walnut-8987", or empty when it did not say
    std::string display;          // "Walnut · 8987", else the advertised name, else empty
    bool        state_known { false };  // the manufacturer data was heard
    bool        unclaimed { false };
    bool        busy { false };
    std::string endpoint_prefix;

    bool operator==(const Printer &other) const;
    bool operator!=(const Printer &other) const { return !(*this == other); }
};

// A device is an M1 when it advertises the service, or when its manufacturer data
// decodes and its name is an M1's (company id 0xffff alone proves nothing).
std::optional<Printer> printer_from(const std::string &id, const std::string &local_name, bool has_service,
                                    const std::vector<uint8_t> *manufacturer_data);

// Merges what one scan hears about each device. Platforms report the advertising
// packet and the scan response separately, and again every time they are heard;
// observe() returns a printer only when that adds or changes one.
class Tracker
{
public:
    // `local_name` and `manufacturer_data` are null when this report did not carry them.
    std::optional<Printer> observe(const std::string &id, const std::string *local_name, bool has_service,
                                   const std::vector<uint8_t> *manufacturer_data);

private:
    struct Heard
    {
        std::string          name;
        bool                 service { false };
        bool                 has_data { false };
        std::vector<uint8_t> data;
    };
    std::map<std::string, Heard>   heard;
    std::map<std::string, Printer> reported;
};

enum class Radio
{
    On,
    Off,          // the adapter is there and switched off
    Denied,       // the system did not let this app use Bluetooth
    Unavailable,  // no adapter, no Bluetooth service, or no scanner on this platform
};

class Scanner
{
public:
    using PrinterFn = std::function<void(const Printer &)>;
    using RadioFn   = std::function<void(Radio)>;

    virtual ~Scanner() = default;

    // Starts listening and returns at once. Both callbacks run on a background
    // thread. on_radio is called once when the scan starts or fails to.
    virtual void start(PrinterFn on_printer, RadioFn on_radio) = 0;
    // Stops listening; calling it twice is fine. A callback already running may
    // still finish, so the target of the callbacks must guard its own lifetime.
    virtual void stop() = 0;
};

// This platform's scanner, or nullptr where there is none.
std::unique_ptr<Scanner> create_scanner();

} // namespace MuonBle
} // namespace Slic3r

#endif
