// Windows: Windows.Devices.Bluetooth.Advertisement.BluetoothLEAdvertisementWatcher,
// through the WinRT ABI and WRL (no C++/WinRT, whose headers want coroutine
// support this C++17 build does not switch on).
//
// The watcher scans actively, so each advertising packet and each scan response
// arrives as its own Received event; Tracker merges them per Bluetooth address.
// It runs on a worker thread in the multithreaded apartment; events arrive on
// the thread pool.
//
// This file is built without the precompiled header (see src/slic3r/CMakeLists.txt),
// so that wx and boost headers stay out of the WinRT ones.

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include "MuonBleScan.hpp"

#include <chrono>
#include <condition_variable>
#include <cstdio>
#include <mutex>
#include <thread>

#include <windows.h>
#include <roapi.h>
#include <robuffer.h>
#include <wrl/client.h>
#include <wrl/event.h>
#include <wrl/wrappers/corewrappers.h>
#include <windows.devices.bluetooth.h>
#include <windows.devices.bluetooth.advertisement.h>
#include <windows.foundation.collections.h>
#include <windows.storage.streams.h>

#pragma comment(lib, "runtimeobject.lib")

namespace Slic3r {
namespace MuonBle {

namespace {

using namespace ABI::Windows::Devices::Bluetooth;
using namespace ABI::Windows::Devices::Bluetooth::Advertisement;
using namespace ABI::Windows::Foundation;
using namespace ABI::Windows::Foundation::Collections;
using ABI::Windows::Storage::Streams::IBuffer;
using Microsoft::WRL::Callback;
using Microsoft::WRL::ComPtr;
using Microsoft::WRL::Wrappers::HStringReference;

using ReceivedHandler = ITypedEventHandler<BluetoothLEAdvertisementWatcher *, BluetoothLEAdvertisementReceivedEventArgs *>;
using StoppedHandler  = ITypedEventHandler<BluetoothLEAdvertisementWatcher *, BluetoothLEAdvertisementWatcherStoppedEventArgs *>;

// 843ce1e6-c64c-4ac1-b06b-5afd8329b7f9, MuonBle::SERVICE_UUID.
constexpr GUID SERVICE_GUID = { 0x843ce1e6, 0xc64c, 0x4ac1, { 0xb0, 0x6b, 0x5a, 0xfd, 0x83, 0x29, 0xb7, 0xf9 } };

struct Shared
{
    std::mutex              mutex;
    std::condition_variable cv;
    bool                    stop { false };
    bool                    stopped_by_system { false };
    Radio                   stopped_radio { Radio::Unavailable };
    Tracker                 tracker;  // guarded by mutex: Received runs on several pool threads
};

std::string to_utf8(HSTRING hstring)
{
    UINT32         length = 0;
    const wchar_t *text   = WindowsGetStringRawBuffer(hstring, &length);
    if (text == nullptr || length == 0)
        return {};
    const int size = WideCharToMultiByte(CP_UTF8, 0, text, int(length), nullptr, 0, nullptr, nullptr);
    std::string out(size_t(size > 0 ? size : 0), '\0');
    if (size > 0)
        WideCharToMultiByte(CP_UTF8, 0, text, int(length), &out[0], size, nullptr, nullptr);
    return out;
}

bool buffer_bytes(IBuffer *buffer, std::vector<uint8_t> &out)
{
    ComPtr<::Windows::Storage::Streams::IBufferByteAccess> access;
    UINT32 length = 0;
    ::byte *bytes = nullptr;
    if (FAILED(buffer->get_Length(&length)) || FAILED(buffer->QueryInterface(IID_PPV_ARGS(&access))) || FAILED(access->Buffer(&bytes)))
        return false;
    out.assign(bytes, bytes + length);
    return true;
}

bool has_service(IBluetoothLEAdvertisement *advertisement)
{
    ComPtr<IVector<GUID>> uuids;
    UINT32 count = 0;
    if (FAILED(advertisement->get_ServiceUuids(&uuids)) || !uuids || FAILED(uuids->get_Size(&count)))
        return false;
    for (UINT32 i = 0; i < count; ++i) {
        GUID uuid;
        if (SUCCEEDED(uuids->GetAt(i, &uuid)) && IsEqualGUID(uuid, SERVICE_GUID))
            return true;
    }
    return false;
}

bool manufacturer_data(IBluetoothLEAdvertisement *advertisement, std::vector<uint8_t> &out)
{
    ComPtr<IVector<BluetoothLEManufacturerData *>> list;
    UINT32 count = 0;
    if (FAILED(advertisement->get_ManufacturerData(&list)) || !list || FAILED(list->get_Size(&count)))
        return false;
    for (UINT32 i = 0; i < count; ++i) {
        ComPtr<IBluetoothLEManufacturerData> data;
        UINT16 company = 0;
        ComPtr<IBuffer> buffer;
        if (SUCCEEDED(list->GetAt(i, &data)) && SUCCEEDED(data->get_CompanyId(&company)) && company == COMPANY_ID &&
            SUCCEEDED(data->get_Data(&buffer)) && buffer && buffer_bytes(buffer.Get(), out))
            return true;
    }
    return false;
}

void on_received(Shared &shared, IBluetoothLEAdvertisementReceivedEventArgs *args, const Scanner::PrinterFn &on_printer)
{
    UINT64 address = 0;
    ComPtr<IBluetoothLEAdvertisement> advertisement;
    if (FAILED(args->get_BluetoothAddress(&address)) || FAILED(args->get_Advertisement(&advertisement)) || !advertisement)
        return;

    std::string name;
    HSTRING     local_name = nullptr;
    if (SUCCEEDED(advertisement->get_LocalName(&local_name))) {
        name = to_utf8(local_name);
        WindowsDeleteString(local_name);
    }
    std::vector<uint8_t> data;
    const bool           has_data = manufacturer_data(advertisement.Get(), data);
    const bool           service  = has_service(advertisement.Get());

    char id[17];
    std::snprintf(id, sizeof(id), "%012llx", static_cast<unsigned long long>(address));

    std::optional<Printer> printer;
    {
        std::lock_guard<std::mutex> lock(shared.mutex);
        if (shared.stop)
            return;
        printer = shared.tracker.observe(id, &name, service, has_data ? &data : nullptr);
    }
    if (printer)
        on_printer(*printer);
}

Radio radio_for(BluetoothError error)
{
    switch (error) {
    case BluetoothError_RadioNotAvailable:
    case BluetoothError_DisabledByUser: return Radio::Off;
    case BluetoothError_DisabledByPolicy:
    case BluetoothError_ConsentRequired: return Radio::Denied;
    default: return Radio::Unavailable;
    }
}

void run(std::shared_ptr<Shared> shared, Scanner::PrinterFn on_printer, Scanner::RadioFn on_radio)
{
    const HRESULT init = RoInitialize(RO_INIT_MULTITHREADED);
    {
        ComPtr<IInspectable>                       inspectable;
        ComPtr<IBluetoothLEAdvertisementWatcher>   watcher;
        if (FAILED(RoActivateInstance(HStringReference(RuntimeClass_Windows_Devices_Bluetooth_Advertisement_BluetoothLEAdvertisementWatcher).Get(), &inspectable)) ||
            FAILED(inspectable.As(&watcher))) {
            on_radio(Radio::Unavailable);
        } else {
            // Active: ask for the scan response, which carries the name and the manufacturer data.
            watcher->put_ScanningMode(BluetoothLEScanningMode_Active);

            EventRegistrationToken received_token {};
            EventRegistrationToken stopped_token {};
            watcher->add_Received(Callback<ReceivedHandler>([shared, on_printer](IBluetoothLEAdvertisementWatcher *,
                                                                                 IBluetoothLEAdvertisementReceivedEventArgs *args) {
                                      on_received(*shared, args, on_printer);
                                      return S_OK;
                                  }).Get(),
                                  &received_token);
            // Stopped fires when the radio is off or the system refuses, as well as after Stop().
            watcher->add_Stopped(Callback<StoppedHandler>([shared](IBluetoothLEAdvertisementWatcher *,
                                                                                  IBluetoothLEAdvertisementWatcherStoppedEventArgs *args) {
                                     BluetoothError error = BluetoothError_Success;
                                     args->get_Error(&error);
                                     {
                                         std::lock_guard<std::mutex> lock(shared->mutex);
                                         if (!shared->stop) {
                                             shared->stopped_by_system = true;
                                             shared->stopped_radio     = radio_for(error);
                                         }
                                     }
                                     shared->cv.notify_all();
                                     return S_OK;
                                 }).Get(),
                                 &stopped_token);

            if (FAILED(watcher->Start())) {
                on_radio(Radio::Unavailable);
            } else {
                // Start() returns before the radio is checked; a refusal arrives as Stopped shortly after.
                std::unique_lock<std::mutex> lock(shared->mutex);
                if (!shared->cv.wait_for(lock, std::chrono::milliseconds(500), [&shared] { return shared->stop || shared->stopped_by_system; })) {
                    lock.unlock();
                    on_radio(Radio::On);
                    lock.lock();
                }
                shared->cv.wait(lock, [&shared] { return shared->stop || shared->stopped_by_system; });
                const bool by_system = shared->stopped_by_system && !shared->stop;
                const Radio radio    = shared->stopped_radio;
                lock.unlock();
                if (by_system)
                    on_radio(radio);
                watcher->Stop();
            }
            watcher->remove_Received(received_token);
            watcher->remove_Stopped(stopped_token);
        }
    }
    if (SUCCEEDED(init))
        RoUninitialize();
}

class WinScanner : public Scanner
{
public:
    ~WinScanner() override
    {
        stop();
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

std::unique_ptr<Scanner> create_scanner() { return std::make_unique<WinScanner>(); }

} // namespace MuonBle
} // namespace Slic3r
