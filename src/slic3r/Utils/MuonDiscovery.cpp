#include "MuonDiscovery.hpp"

#include <algorithm>
#include <cctype>
#include <regex>

#include <nlohmann/json.hpp>

namespace Slic3r {
namespace MuonDiscovery {

static std::string to_lower(std::string s)
{
    std::transform(s.begin(), s.end(), s.begin(), [](unsigned char c) { return char(std::tolower(c)); });
    return s;
}

bool is_muon_service(const TxtData &txt)
{
    if (txt.count("setup") != 0)
        return true;
    const auto model = txt.find("model");
    return model != txt.end() && to_lower(model->second).rfind("muon", 0) == 0;
}

std::string muon_address(const std::string &host, bool host_is_ipv6, const std::string &path)
{
    std::string address = host_is_ipv6 ? "[" + host + "]" : host;
    if (!path.empty() && path != "/")
        address += path.front() == '/' ? path : "/" + path;
    return address;
}

std::string short_name(const std::string &name)
{
    // The word list gives 4-7 lower case letters; the suffix is 4 hex digits.
    static const std::regex re("^(?:muon-)?([a-z]+)-([0-9a-f]{4})$", std::regex::icase);
    const std::string host = name.substr(0, name.find('.'));
    std::smatch m;
    if (!std::regex_match(host, m, re))
        return {};
    return to_lower(m[1].str()) + "-" + to_lower(m[2].str());
}

std::string display_name(const std::string &name)
{
    const std::string s = short_name(name);
    if (s.empty())
        return {};
    const size_t dash = s.find('-');
    std::string word = s.substr(0, dash);
    std::string suffix = s.substr(dash + 1);
    word[0] = char(std::toupper((unsigned char) word[0]));
    std::transform(suffix.begin(), suffix.end(), suffix.begin(), [](unsigned char c) { return char(std::toupper(c)); });
    // U+00B7 MIDDLE DOT, as the printer's screen writes it (ID-4).
    return word + " \xC2\xB7 " + suffix;
}

std::optional<Identity> parse_identity(const std::string &body)
{
    nlohmann::json j = nlohmann::json::parse(body, nullptr, /*allow_exceptions=*/false);
    if (!j.is_object())
        return std::nullopt;
    const auto result = j.find("result");
    if (result == j.end() || !result->is_object())
        return std::nullopt;

    const auto str = [&result](const char *key) -> std::string {
        const auto it = result->find(key);
        return it != result->end() && it->is_string() ? it->get<std::string>() : std::string();
    };

    Identity id;
    id.name        = str("name");
    id.display     = str("display");
    id.setup       = str("setup");
    id.endpoint_id = str("endpoint_id");

    // The derived name is what every other surface carries; the owner's rename is not.
    id.short_name = short_name(str("ssid"));
    if (id.short_name.empty())
        id.short_name = short_name(str("derived_name") + "-" + str("suffix"));
    if (id.short_name.empty())
        id.short_name = short_name(str("derived_name"));

    // It has to name the printer: an answer that does not is not an M1's.
    if (id.name.empty() && id.short_name.empty())
        return std::nullopt;
    if (id.display.empty())
        id.display = !id.short_name.empty() ? display_name(id.short_name) : id.name;
    return id;
}

} // namespace MuonDiscovery
} // namespace Slic3r
