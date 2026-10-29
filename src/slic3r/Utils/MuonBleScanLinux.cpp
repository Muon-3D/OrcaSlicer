// Linux: BlueZ over D-Bus, through GDBus (GIO, which GTK already brings).
//
// A worker thread sets a discovery filter for Muon's service UUID, starts
// discovery on the first adapter, and reads the Device1 objects BlueZ keeps
// (Name, UUIDs, ManufacturerData) once a second. BlueZ scans actively, so the
// scan response's name and manufacturer data land on the same Device1.

#include "MuonBleScan.hpp"

#include <algorithm>
#include <atomic>
#include <cctype>
#include <chrono>
#include <condition_variable>
#include <mutex>
#include <thread>

#include <gio/gio.h>

namespace Slic3r {
namespace MuonBle {

namespace {

constexpr int CALL_TIMEOUT_MS = 2000;

struct Shared
{
    std::mutex              mutex;
    std::condition_variable cv;
    bool                    stop { false };
};

GVariant *call(GDBusConnection *bus, const char *path, const char *iface, const char *method, GVariant *args,
               const GVariantType *reply_type, std::string *error_name = nullptr)
{
    GError   *error = nullptr;
    GVariant *reply = g_dbus_connection_call_sync(bus, "org.bluez", path, iface, method, args, reply_type,
                                                  G_DBUS_CALL_FLAGS_NONE, CALL_TIMEOUT_MS, nullptr, &error);
    if (error != nullptr) {
        if (error_name != nullptr) {
            gchar *remote = g_dbus_error_get_remote_error(error);
            *error_name   = remote != nullptr ? remote : "";
            g_free(remote);
        }
        g_error_free(error);
    }
    return reply;
}

GVariant *managed_objects(GDBusConnection *bus)
{
    return call(bus, "/", "org.freedesktop.DBus.ObjectManager", "GetManagedObjects", nullptr,
                G_VARIANT_TYPE("(a{oa{sa{sv}}})"));
}

bool lookup_bool(GVariant *props, const char *key)
{
    gboolean value = FALSE;
    return g_variant_lookup(props, key, "b", &value) && value;
}

std::string lookup_string(GVariant *props, const char *key)
{
    const gchar *value = nullptr;
    return g_variant_lookup(props, key, "&s", &value) && value != nullptr ? std::string(value) : std::string();
}

bool has_service_uuid(GVariant *props)
{
    GVariant *uuids = g_variant_lookup_value(props, "UUIDs", G_VARIANT_TYPE("as"));
    if (uuids == nullptr)
        return false;
    bool        found = false;
    GVariantIter it;
    const gchar *uuid = nullptr;
    g_variant_iter_init(&it, uuids);
    while (!found && g_variant_iter_next(&it, "&s", &uuid)) {
        std::string s(uuid);
        std::transform(s.begin(), s.end(), s.begin(), [](unsigned char c) { return char(std::tolower(c)); });
        found = s == SERVICE_UUID;
    }
    g_variant_unref(uuids);
    return found;
}

bool manufacturer_data(GVariant *props, std::vector<uint8_t> &out)
{
    GVariant *data = g_variant_lookup_value(props, "ManufacturerData", G_VARIANT_TYPE("a{qv}"));
    if (data == nullptr)
        return false;
    bool         found = false;
    GVariantIter it;
    guint16      company = 0;
    GVariant    *value   = nullptr;
    g_variant_iter_init(&it, data);
    while (g_variant_iter_next(&it, "{qv}", &company, &value)) {
        if (!found && company == COMPANY_ID && g_variant_is_of_type(value, G_VARIANT_TYPE_BYTESTRING)) {
            gsize         n     = 0;
            const guint8 *bytes = static_cast<const guint8 *>(g_variant_get_fixed_array(value, &n, 1));
            out.assign(bytes, bytes + n);
            found = true;
        }
        g_variant_unref(value);
    }
    g_variant_unref(data);
    return found;
}

// The first adapter's object path, or empty. `powered` says whether it is switched on.
std::string find_adapter(GDBusConnection *bus, bool &powered)
{
    std::string adapter;
    GVariant   *objects = managed_objects(bus);
    if (objects == nullptr)
        return adapter;
    GVariantIter *it         = nullptr;
    const gchar  *path       = nullptr;
    GVariant     *interfaces = nullptr;
    g_variant_get(objects, "(a{oa{sa{sv}}})", &it);
    while (adapter.empty() && g_variant_iter_next(it, "{&o@a{sa{sv}}}", &path, &interfaces)) {
        GVariant *props = g_variant_lookup_value(interfaces, "org.bluez.Adapter1", G_VARIANT_TYPE("a{sv}"));
        if (props != nullptr) {
            adapter = path;
            powered = lookup_bool(props, "Powered");
            g_variant_unref(props);
        }
        g_variant_unref(interfaces);
    }
    g_variant_iter_free(it);
    g_variant_unref(objects);
    return adapter;
}

void report_devices(GDBusConnection *bus, Tracker &tracker, const Scanner::PrinterFn &on_printer)
{
    GVariant *objects = managed_objects(bus);
    if (objects == nullptr)
        return;
    GVariantIter *it         = nullptr;
    const gchar  *path       = nullptr;
    GVariant     *interfaces = nullptr;
    g_variant_get(objects, "(a{oa{sa{sv}}})", &it);
    while (g_variant_iter_next(it, "{&o@a{sa{sv}}}", &path, &interfaces)) {
        GVariant *props = g_variant_lookup_value(interfaces, "org.bluez.Device1", G_VARIANT_TYPE("a{sv}"));
        g_variant_unref(interfaces);
        if (props == nullptr)
            continue;
        // BlueZ drops RSSI when a device has not been heard in this discovery: skip what it only remembers.
        gint16 rssi = 0;
        if (g_variant_lookup(props, "RSSI", "n", &rssi)) {
            std::vector<uint8_t> data;
            const bool           has_data = manufacturer_data(props, data);
            const std::string    name     = lookup_string(props, "Name");
            if (auto printer = tracker.observe(path, &name, has_service_uuid(props), has_data ? &data : nullptr))
                on_printer(*printer);
        }
        g_variant_unref(props);
    }
    g_variant_iter_free(it);
    g_variant_unref(objects);
}

void run(std::shared_ptr<Shared> shared, Scanner::PrinterFn on_printer, Scanner::RadioFn on_radio)
{
    GError          *error = nullptr;
    GDBusConnection *bus   = g_bus_get_sync(G_BUS_TYPE_SYSTEM, nullptr, &error);
    if (bus == nullptr) {
        g_clear_error(&error);
        on_radio(Radio::Unavailable);
        return;
    }

    bool              powered = false;
    const std::string adapter = find_adapter(bus, powered);
    if (adapter.empty() || !powered) {
        on_radio(adapter.empty() ? Radio::Unavailable : Radio::Off);
        g_object_unref(bus);
        return;
    }

    // Only devices that advertise Muon's service, over LE, reported every time they are heard.
    GVariantBuilder filter;
    g_variant_builder_init(&filter, G_VARIANT_TYPE("a{sv}"));
    const gchar *uuids[] = { SERVICE_UUID, nullptr };
    g_variant_builder_add(&filter, "{sv}", "UUIDs", g_variant_new_strv(uuids, -1));
    g_variant_builder_add(&filter, "{sv}", "Transport", g_variant_new_string("le"));
    g_variant_builder_add(&filter, "{sv}", "DuplicateData", g_variant_new_boolean(TRUE));
    if (GVariant *r = call(bus, adapter.c_str(), "org.bluez.Adapter1", "SetDiscoveryFilter", g_variant_new("(a{sv})", &filter), nullptr))
        g_variant_unref(r);

    std::string error_name;
    GVariant   *started = call(bus, adapter.c_str(), "org.bluez.Adapter1", "StartDiscovery", nullptr, nullptr, &error_name);
    // InProgress: another client is already discovering, and BlueZ still reports to this one.
    if (started == nullptr && error_name != "org.bluez.Error.InProgress") {
        on_radio(error_name == "org.bluez.Error.NotReady" ? Radio::Off :
                 error_name == "org.bluez.Error.NotAuthorized" || error_name == "org.freedesktop.DBus.Error.AccessDenied" ? Radio::Denied :
                                                                   Radio::Unavailable);
        g_object_unref(bus);
        return;
    }
    // Stop only a discovery this client started; with InProgress it belongs to someone else.
    const bool own_discovery = started != nullptr;
    if (started != nullptr)
        g_variant_unref(started);
    on_radio(Radio::On);

    Tracker tracker;
    for (;;) {
        report_devices(bus, tracker, on_printer);
        std::unique_lock<std::mutex> lock(shared->mutex);
        if (shared->cv.wait_for(lock, std::chrono::seconds(1), [&shared] { return shared->stop; }))
            break;
    }

    if (own_discovery) {
        if (GVariant *r = call(bus, adapter.c_str(), "org.bluez.Adapter1", "StopDiscovery", nullptr, nullptr))
            g_variant_unref(r);
    }
    g_object_unref(bus);
}

class LinuxScanner : public Scanner
{
public:
    ~LinuxScanner() override
    {
        stop();
        // A D-Bus call can take CALL_TIMEOUT_MS; don't hold the UI for it. The thread owns what it uses.
        if (m_thread.joinable())
            m_thread.detach();
    }

    void start(PrinterFn on_printer, RadioFn on_radio) override
    {
        if (m_thread.joinable())
            return;
        m_thread = std::thread(run, m_shared, std::move(on_printer), std::move(on_radio));
    }

    void stop() override
    {
        {
            std::lock_guard<std::mutex> lock(m_shared->mutex);
            m_shared->stop = true;
        }
        m_shared->cv.notify_all();
    }

private:
    std::shared_ptr<Shared> m_shared { std::make_shared<Shared>() };
    std::thread             m_thread;
};

} // namespace

std::unique_ptr<Scanner> create_scanner() { return std::make_unique<LinuxScanner>(); }

} // namespace MuonBle
} // namespace Slic3r
