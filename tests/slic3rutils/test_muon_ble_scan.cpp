#include <catch2/catch_all.hpp>

#include "slic3r/Utils/MuonBleScan.hpp"

using namespace Slic3r::MuonBle;

namespace {
// blelink-proto AdvData::encode(): version 1, flags, the EndpointId's first 8 bytes.
const std::vector<uint8_t> UNCLAIMED = {1, 0x01, 0xde, 0xad, 0xbe, 0xef, 0x00, 0x01, 0x02, 0x03};
const std::vector<uint8_t> SET_UP    = {1, 0x00, 0xde, 0xad, 0xbe, 0xef, 0x00, 0x01, 0x02, 0x03};
const std::vector<uint8_t> BUSY      = {1, 0x03, 0xde, 0xad, 0xbe, 0xef, 0x00, 0x01, 0x02, 0x03, 0xff}; // trailing bytes are allowed
} // namespace

TEST_CASE("Muon manufacturer data decodes", "[MuonBle]")
{
    const auto unclaimed = decode_manufacturer_data(UNCLAIMED);
    REQUIRE(unclaimed);
    CHECK(unclaimed->unclaimed);
    CHECK_FALSE(unclaimed->busy);
    CHECK(unclaimed->endpoint_prefix == "deadbeef00010203");

    const auto set_up = decode_manufacturer_data(SET_UP);
    REQUIRE(set_up);
    CHECK_FALSE(set_up->unclaimed);

    const auto busy = decode_manufacturer_data(BUSY);
    REQUIRE(busy);
    CHECK(busy->unclaimed);
    CHECK(busy->busy);

    CHECK_FALSE(decode_manufacturer_data({}));
    CHECK_FALSE(decode_manufacturer_data({1, 1, 0xde, 0xad, 0xbe, 0xef, 0, 1, 2})); // 9 bytes
    CHECK_FALSE(decode_manufacturer_data({0, 1, 0xde, 0xad, 0xbe, 0xef, 0, 1, 2, 3})); // version 0

    // A later version keeps the layout: still decoded, unknown flag bits and extra bytes ignored.
    const auto later = decode_manufacturer_data({2, 0x05, 0xde, 0xad, 0xbe, 0xef, 0, 1, 2, 3, 0x99});
    REQUIRE(later);
    CHECK(later->unclaimed);
    CHECK_FALSE(later->busy);
    CHECK(later->endpoint_prefix == "deadbeef00010203");
}

TEST_CASE("An advertisement becomes a printer row only when it is an M1's", "[MuonBle]")
{
    SECTION("an unset printer, name and data from the scan response") {
        const auto printer = printer_from("aa", "walnut-8987", true, &UNCLAIMED);
        REQUIRE(printer);
        CHECK(printer->short_name == "walnut-8987");
        CHECK(printer->display == "Walnut \xC2\xB7 8987");
        CHECK(printer->state_known);
        CHECK(printer->unclaimed);
        CHECK(printer->endpoint_prefix == "deadbeef00010203");
    }
    SECTION("the service alone, before the scan response is heard") {
        const auto printer = printer_from("aa", "", true, nullptr);
        REQUIRE(printer);
        CHECK(printer->display.empty());
        CHECK_FALSE(printer->state_known);
        CHECK_FALSE(printer->unclaimed);
    }
    SECTION("the scan response alone, with an M1 name") {
        const auto printer = printer_from("aa", "walnut-8987", false, &SET_UP);
        REQUIRE(printer);
        CHECK_FALSE(printer->unclaimed);
    }
    SECTION("company id 0xffff is the SIG's test id, so it is not enough") {
        CHECK_FALSE(printer_from("bb", "headphones", false, &UNCLAIMED));
        CHECK_FALSE(printer_from("bb", "", false, &UNCLAIMED));
    }
    SECTION("an M1 name without Muon's data or service") {
        CHECK_FALSE(printer_from("cc", "walnut-8987", false, nullptr));
    }
}

TEST_CASE("The tracker merges advertisement and scan response", "[MuonBle]")
{
    Tracker tracker;
    const std::string name = "walnut-8987";

    // Windows reports the two packets separately: the service first, nameless.
    auto first = tracker.observe("aa", nullptr, true, nullptr);
    REQUIRE(first);
    CHECK(first->display.empty());

    // Then the scan response, which names it and carries the flags.
    auto second = tracker.observe("aa", &name, false, &UNCLAIMED);
    REQUIRE(second);
    CHECK(second->display == "Walnut \xC2\xB7 8987");
    CHECK(second->unclaimed);

    // Heard again unchanged: no update.
    CHECK_FALSE(tracker.observe("aa", nullptr, true, nullptr));
    CHECK_FALSE(tracker.observe("aa", &name, false, &UNCLAIMED));

    // Set up since: an update.
    auto third = tracker.observe("aa", &name, false, &SET_UP);
    REQUIRE(third);
    CHECK_FALSE(third->unclaimed);

    // Another device that is not an M1 never reports.
    const std::string other = "headphones";
    CHECK_FALSE(tracker.observe("bb", &other, false, &UNCLAIMED));
}
