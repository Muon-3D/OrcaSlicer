// macOS: CoreBluetooth. A CBCentralManager on its own serial queue scans for
// Muon's service UUID (scanning in the foreground is active, so the scan
// response's name and manufacturer data arrive too) and never connects.
//
// Using Bluetooth needs NSBluetoothAlwaysUsageDescription in Info.plist, or
// macOS ends the process; without the key this reports Unavailable and does
// nothing. The first scan makes macOS ask the user for permission.
//
// Manual retain/release, like the other .mm files here.

#import <Foundation/Foundation.h>
#import <CoreBluetooth/CoreBluetooth.h>

#include "MuonBleScan.hpp"

using Slic3r::MuonBle::Printer;
using Slic3r::MuonBle::Radio;
using Slic3r::MuonBle::Scanner;
using Slic3r::MuonBle::Tracker;

@interface MuonBleScanDelegate : NSObject <CBCentralManagerDelegate>
{
@public
    Scanner::PrinterFn on_printer;
    Scanner::RadioFn   on_radio;
    Tracker            tracker;
    BOOL               stopped;
    BOOL               radio_reported;
}
@end

@implementation MuonBleScanDelegate

- (void)reportRadio:(Radio)radio
{
    if (!radio_reported) {
        radio_reported = YES;
        on_radio(radio);
    }
}

- (void)centralManagerDidUpdateState:(CBCentralManager *)central
{
    if (stopped)
        return;
    switch (central.state) {
    case CBManagerStatePoweredOn: {
        NSArray<CBUUID *> *services = @[ [CBUUID UUIDWithString:@(Slic3r::MuonBle::SERVICE_UUID)] ];
        // Duplicates, so the scan response is reported after the advertising packet.
        [central scanForPeripheralsWithServices:services options:@{ CBCentralManagerScanOptionAllowDuplicatesKey : @YES }];
        [self reportRadio:Radio::On];
        break;
    }
    case CBManagerStatePoweredOff: [self reportRadio:Radio::Off]; break;
    case CBManagerStateUnauthorized: [self reportRadio:Radio::Denied]; break;
    case CBManagerStateUnsupported: [self reportRadio:Radio::Unavailable]; break;
    default:
        // Unknown or resetting: another update follows.
        break;
    }
}

- (void)centralManager:(CBCentralManager *)central
    didDiscoverPeripheral:(CBPeripheral *)peripheral
        advertisementData:(NSDictionary<NSString *, id> *)advertisementData
                     RSSI:(NSNumber *)RSSI
{
    if (stopped)
        return;

    std::string name;
    NSString *local_name = advertisementData[CBAdvertisementDataLocalNameKey];
    if (local_name == nil)
        local_name = peripheral.name;
    if (local_name != nil && local_name.UTF8String != nullptr)
        name = local_name.UTF8String;

    // CoreBluetooth hands over the whole value: a little-endian company id, then the data.
    std::vector<uint8_t> data;
    bool has_data = false;
    NSData *manufacturer = advertisementData[CBAdvertisementDataManufacturerDataKey];
    if (manufacturer != nil && manufacturer.length >= 2) {
        const uint8_t *bytes = static_cast<const uint8_t *>(manufacturer.bytes);
        if (uint16_t(bytes[0] | (bytes[1] << 8)) == Slic3r::MuonBle::COMPANY_ID) {
            data.assign(bytes + 2, bytes + manufacturer.length);
            has_data = true;
        }
    }

    // The scan is filtered on Muon's service, so every device reported advertises it.
    const std::string id = peripheral.identifier.UUIDString.UTF8String;
    if (auto printer = tracker.observe(id, &name, true, has_data ? &data : nullptr))
        on_printer(*printer);
}

@end

namespace Slic3r {
namespace MuonBle {

namespace {

class MacScanner : public Scanner
{
public:
    MacScanner() : m_queue(dispatch_queue_create("com.muon3d.orca.ble-scan", DISPATCH_QUEUE_SERIAL)) {}

    ~MacScanner() override
    {
        stop();
        dispatch_release(m_queue);
    }

    void start(PrinterFn on_printer, RadioFn on_radio) override
    {
        if (m_central != nil)
            return;
        if ([[NSBundle mainBundle] objectForInfoDictionaryKey:@"NSBluetoothAlwaysUsageDescription"] == nil) {
            on_radio(Radio::Unavailable);
            return;
        }

        MuonBleScanDelegate *delegate = [[MuonBleScanDelegate alloc] init];
        delegate->on_printer = std::move(on_printer);
        delegate->on_radio   = std::move(on_radio);
        m_delegate = delegate;
        // Calls centralManagerDidUpdateState: on m_queue once the state is known.
        m_central = [[CBCentralManager alloc] initWithDelegate:delegate
                                                         queue:m_queue
                                                       options:@{ CBCentralManagerOptionShowPowerAlertKey : @NO }];
    }

    void stop() override
    {
        if (m_central == nil)
            return;
        CBCentralManager    *central  = m_central;
        MuonBleScanDelegate *delegate = m_delegate;
        m_central  = nil;
        m_delegate = nil;
        // On the delegate's queue, so no callback runs after `stopped` is set.
        dispatch_async(m_queue, ^{
            delegate->stopped = YES;
            if (central.state == CBManagerStatePoweredOn)
                [central stopScan];
            central.delegate = nil;
            [central release];
            [delegate release];
        });
    }

private:
    dispatch_queue_t     m_queue;
    CBCentralManager    *m_central { nil };
    MuonBleScanDelegate *m_delegate { nil };
};

} // namespace

std::unique_ptr<Scanner> create_scanner() { return std::make_unique<MacScanner>(); }

} // namespace MuonBle
} // namespace Slic3r
