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
	// A Muon printer found other than by Bonjour: on its own hotspot.
	struct MuonFound
	{
		std::string short_name;  // "walnut-8987", to match the same printer found by Bonjour
		wxString    display;     // "Walnut · 8987"
		wxString    address;     // what get_selected() returns
		std::string setup;       // muon_setup's state, or empty when it was not readable
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
