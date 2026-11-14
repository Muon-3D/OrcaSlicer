#include "MuonBleScan.hpp"

#include "MuonDiscovery.hpp"

namespace Slic3r {
namespace MuonBle {

std::optional<Advert> decode_manufacturer_data(const std::vector<uint8_t> &data)
{
    // Version 1 is the only one there is. Like blelink-proto's AdvData::decode, accept later
    // versions with the same layout (unknown flag bits and trailing bytes are ignored), so a
    // newer printer is still listed rather than dropped.
    if (data.size() < 10 || data[0] < 1)
        return std::nullopt;

    static const char *hex = "0123456789abcdef";
    Advert advert;
    advert.unclaimed = (data[1] & 0x01) != 0;
    advert.busy      = (data[1] & 0x02) != 0;
    for (size_t i = 2; i < 10; ++i) {
        advert.endpoint_prefix += hex[data[i] >> 4];
        advert.endpoint_prefix += hex[data[i] & 0x0f];
    }
    return advert;
}

bool Printer::operator==(const Printer &other) const
{
    return id == other.id && short_name == other.short_name && display == other.display &&
           state_known == other.state_known && unclaimed == other.unclaimed && busy == other.busy &&
           endpoint_prefix == other.endpoint_prefix;
}

std::optional<Printer> printer_from(const std::string &id, const std::string &local_name, bool has_service,
                                    const std::vector<uint8_t> *manufacturer_data)
{
    const std::optional<Advert> advert = manufacturer_data != nullptr ? decode_manufacturer_data(*manufacturer_data) : std::nullopt;
    const std::string short_name = MuonDiscovery::short_name(local_name);
    if (!has_service && !(advert && !short_name.empty()))
        return std::nullopt;

    Printer printer;
    printer.id         = id;
    printer.short_name = short_name;
    printer.display    = short_name.empty() ? local_name : MuonDiscovery::display_name(short_name);
    if (advert) {
        printer.state_known     = true;
        printer.unclaimed       = advert->unclaimed;
        printer.busy            = advert->busy;
        printer.endpoint_prefix = advert->endpoint_prefix;
    }
    return printer;
}

std::optional<Printer> Tracker::observe(const std::string &id, const std::string *local_name, bool has_service,
                                        const std::vector<uint8_t> *manufacturer_data)
{
    Heard &h = heard[id];
    if (local_name != nullptr && !local_name->empty())
        h.name = *local_name;
    h.service |= has_service;
    if (manufacturer_data != nullptr) {
        h.has_data = true;
        h.data     = *manufacturer_data;
    }

    std::optional<Printer> printer = printer_from(id, h.name, h.service, h.has_data ? &h.data : nullptr);
    if (!printer)
        return std::nullopt;
    const auto it = reported.find(id);
    if (it != reported.end() && it->second == *printer)
        return std::nullopt;
    reported[id] = *printer;
    return printer;
}

#if !defined(_WIN32) && !defined(__APPLE__) && !defined(__linux__)
std::unique_ptr<Scanner> create_scanner() { return nullptr; }
#endif

} // namespace MuonBle
} // namespace Slic3r
