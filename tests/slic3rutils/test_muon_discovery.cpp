#include <catch2/catch_all.hpp>

#include "slic3r/Utils/MuonDiscovery.hpp"

using namespace Slic3r::MuonDiscovery;

TEST_CASE("Muon printers are recognised from their Bonjour TXT record", "[MuonDiscovery]")
{
    SECTION("an M1 today: Moonraker's octoprint_model_name") {
        // Moonraker's core/M1 [zeroconf]: octoprint_model_name: Muon M1, octoprint_addr_pref: hostname
        CHECK(is_muon_service({{"model", "Muon M1"}, {"version", "1.1.0"}, {"addr_pref", "hostname"}, {"path", "/"}}));
    }
    SECTION("a later M1 that publishes its setup state") {
        CHECK(is_muon_service({{"model", "Moonraker"}, {"setup", "new"}}));
    }
    SECTION("other services are left alone") {
        CHECK_FALSE(is_muon_service({{"model", "Moonraker"}, {"version", "1.1.0"}}));
        CHECK_FALSE(is_muon_service({{"model", "SL1"}}));
        CHECK_FALSE(is_muon_service({{"path", "/"}}));
        CHECK_FALSE(is_muon_service({}));
    }
}

TEST_CASE("A Muon printer's address drops the advertised port", "[MuonDiscovery]")
{
    // The service record says 7125, which Moonraker only binds on loopback.
    CHECK(muon_address("192.168.1.40", false, "/") == "192.168.1.40");
    CHECK(muon_address("192.168.1.40", false, "") == "192.168.1.40");
    CHECK(muon_address("192.168.1.40", false, "/moonraker") == "192.168.1.40/moonraker");
    CHECK(muon_address("fe80::1", true, "/") == "[fe80::1]");
}

TEST_CASE("M1 names match across host name, hotspot SSID and Bluetooth name", "[MuonDiscovery]")
{
    CHECK(short_name("Muon-walnut-8987.local") == "walnut-8987");
    CHECK(short_name("Muon-walnut-8987") == "walnut-8987");
    CHECK(short_name("walnut-8987") == "walnut-8987");
    CHECK(short_name("BOXWOOD-367A") == "boxwood-367a");

    CHECK(short_name("Muon-M1-10000000abcdef12.local").empty()); // an un-migrated host name
    CHECK(short_name("octopi").empty());
    CHECK(short_name("walnut-898").empty());
    CHECK(short_name("walnut-89zz").empty());
    CHECK(short_name("").empty());

    CHECK(display_name("walnut-8987") == "Walnut \xC2\xB7 8987");
    CHECK(display_name("Muon-boxwood-367a.local") == "Boxwood \xC2\xB7 367A");
    CHECK(display_name("octopi").empty());
}

TEST_CASE("The identity route says whether 10.42.0.1 is an M1", "[MuonDiscovery]")
{
    SECTION("an unset printer on its hotspot") {
        // The shape of GET /server/muon/identity (aux_api_proxy.get_identity).
        const auto id = parse_identity(R"({"result": {
            "name": "walnut", "source": "derived", "derived_name": "walnut", "suffix": "8987",
            "display": "Walnut · 8987", "ssid": "Muon-walnut-8987",
            "endpoint_id": "4f1c0a9e2b7d6c5a4f1c0a9e2b7d6c5a4f1c0a9e2b7d6c5a4f1c0a9e2b7d6c5a",
            "fingerprint": "ab:cd", "setup": "new"}})");
        REQUIRE(id);
        CHECK(id->short_name == "walnut-8987");
        CHECK(id->display == "Walnut \xC2\xB7 8987");
        CHECK(id->setup == "new");
        CHECK(setup_pending(id->setup));
        CHECK(id->endpoint_id.size() == 64);
    }
    SECTION("a renamed printer still matches by its derived name") {
        const auto id = parse_identity(R"({"result": {"name": "Kitchen", "source": "owner",
            "derived_name": "walnut", "suffix": "8987", "display": "Kitchen · 8987", "ssid": null, "setup": "complete"}})");
        REQUIRE(id);
        CHECK(id->short_name == "walnut-8987");
        CHECK(id->display == "Kitchen \xC2\xB7 8987");
        CHECK(setup_complete(id->setup));
    }
    SECTION("an older image that reports no setup state") {
        const auto id = parse_identity(R"({"result": {"name": "walnut", "derived_name": "walnut", "suffix": "8987"}})");
        REQUIRE(id);
        CHECK(id->display == "Walnut \xC2\xB7 8987");
        CHECK(id->setup.empty());
        CHECK_FALSE(setup_pending(id->setup));
        CHECK_FALSE(setup_complete(id->setup));
    }
    SECTION("anything else at that address is not listed") {
        CHECK_FALSE(parse_identity(""));
        CHECK_FALSE(parse_identity("<html>router login</html>"));
        CHECK_FALSE(parse_identity(R"({"result": {"klippy_state": "ready"}})"));  // /server/info-like
        CHECK_FALSE(parse_identity(R"({"error": {"code": 404, "message": "Not Found"}})"));
        CHECK_FALSE(parse_identity(R"({"result": "walnut"})"));
        CHECK_FALSE(parse_identity("[]"));
    }
}
