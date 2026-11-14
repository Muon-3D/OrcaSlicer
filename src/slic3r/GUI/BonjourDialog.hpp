#ifndef slic3r_BonjourDialog_hpp_
#define slic3r_BonjourDialog_hpp_

#include <cstddef>
#include <memory>
#include <string>
#include <vector>

#include <boost/asio/ip/address.hpp>

#include <wx/dialog.h>
#include <wx/string.h>

#include "libslic3r/PrintConfig.hpp"

class wxBoxSizer;
class wxCommandEvent;
class wxListView;
class wxStaticText;
class wxTimer;
class wxTimerEvent;
class address;

namespace Slic3r {

class Bonjour;
struct BonjourReply;
class BonjourReplyEvent;
class ReplySet;
struct LifetimeGuard;
class MuonBleEvent;
namespace MuonBle { class Scanner; }


class BonjourDialog: public wxDialog
{
public:
	// muon_hints: the printer preset is a Muon3D one, so show the name and setup state
	// that Muon printers publish, and help the owner when nothing is found.
	BonjourDialog(wxWindow *parent, Slic3r::PrinterTechnology, bool muon_hints = false);
	BonjourDialog(BonjourDialog &&) = delete;
	BonjourDialog(const BonjourDialog &) = delete;
	BonjourDialog &operator=(BonjourDialog &&) = delete;
	BonjourDialog &operator=(const BonjourDialog &) = delete;
	~BonjourDialog();

	bool show_and_lookup();
	wxString get_selected() const;
private:
	// A Muon printer found other than by Bonjour: on its own hotspot, or nearby over Bluetooth.
	struct MuonFound
	{
		enum class Via { Hotspot, Bluetooth } via { Via::Hotspot };
		std::string short_name;       // "walnut-8987", to match the same printer found another way
		std::string endpoint_prefix;  // the first 8 bytes of its EndpointId in hex, or empty
		wxString    display;          // "Walnut · 8987"
		wxString    address;          // what get_selected() returns; empty over Bluetooth
		std::string setup;            // muon_setup's state, or empty when it was not readable
		std::string bluetooth_id;     // the scanner's id for the device
		bool        state_known { false };  // over Bluetooth: the advertisement said whether it is set up
	};

	// What a list row shows: a Bonjour reply, or a printer found another way.
	struct Row
	{
		const BonjourReply *reply { nullptr };
		const MuonFound    *found { nullptr };
	};

	wxListView *list;
	std::unique_ptr<ReplySet> replies;
	std::vector<MuonFound> found;
	// What each list row shows, indexed by the row's item data.
	std::vector<Row> rows;
	wxStaticText *label;
	wxBoxSizer *help_sizer { nullptr };
	// "Bluetooth is off", when the Bluetooth scan could not run for a reason the owner can fix.
	wxStaticText *bluetooth_note { nullptr };
	std::unique_ptr<MuonBle::Scanner> bluetooth;
	std::shared_ptr<Bonjour> bonjour;
	std::shared_ptr<LifetimeGuard> guard;
	std::unique_ptr<wxTimer> timer;
	unsigned timer_state;
	// Tags the events of each lookup, so a lookup that "Search again" replaced is ignored.
	unsigned lookup_id { 0 };
	Slic3r::PrinterTechnology tech;
	bool muon_hints;
	int status_column { -1 };

	void lookup();
	void probe_hotspot(unsigned id);
	void refresh_list();
	const Row *selected_row() const;
	void show_help(bool show);

	virtual void on_reply(BonjourReplyEvent &);
	void on_hotspot(wxCommandEvent &);
	void on_bluetooth(MuonBleEvent &);
	void on_bluetooth_radio(wxCommandEvent &);
	void stop_bluetooth();
	void on_complete(wxCommandEvent &);
	void on_ok(wxCommandEvent &);
	void on_timer(wxTimerEvent &);
    void on_timer_process();
};

class IPListDialog : public wxDialog
{
public:
	IPListDialog(wxWindow* parent, const wxString& hostname, const std::vector<boost::asio::ip::address>& ips, size_t& selected_index);
	IPListDialog(IPListDialog&&) = delete;
	IPListDialog(const IPListDialog&) = delete;
	IPListDialog& operator=(IPListDialog&&) = delete;
	IPListDialog& operator=(const IPListDialog&) = delete;
	~IPListDialog();

	virtual void EndModal(int retCode) wxOVERRIDE;
private:
	wxListView*		m_list;
	size_t&			m_selected_index;
};

}

#endif
