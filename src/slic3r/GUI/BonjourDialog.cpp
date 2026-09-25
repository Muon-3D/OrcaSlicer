#include "slic3r/Utils/Bonjour.hpp"   // On Windows, boost needs to be included before wxWidgets headers

#include "BonjourDialog.hpp"

#include <set>
#include <mutex>

#include <boost/nowide/convert.hpp>

#include <wx/sizer.h>
#include <wx/button.h>
#include <wx/listctrl.h>
#include <wx/stattext.h>
#include <wx/timer.h>
#include <wx/utils.h>
#include <wx/wupdlock.h>

#include "slic3r/GUI/GUI.hpp"
#include "slic3r/GUI/GUI_App.hpp"
#include "slic3r/GUI/I18N.hpp"
#include "slic3r/GUI/MsgDialog.hpp"
#include "slic3r/GUI/format.hpp"
#include "slic3r/Utils/Bonjour.hpp"

namespace Slic3r {


class BonjourReplyEvent : public wxEvent
{
public:
	BonjourReply reply;
	unsigned lookup_id;

	BonjourReplyEvent(wxEventType eventType, int winid, BonjourReply &&reply, unsigned lookup_id) :
		wxEvent(winid, eventType),
		reply(std::move(reply)),
		lookup_id(lookup_id)
	{}

	virtual wxEvent *Clone() const
	{
		return new BonjourReplyEvent(*this);
	}
};

wxDEFINE_EVENT(EVT_BONJOUR_REPLY, BonjourReplyEvent);

wxDECLARE_EVENT(EVT_BONJOUR_COMPLETE, wxCommandEvent);
wxDEFINE_EVENT(EVT_BONJOUR_COMPLETE, wxCommandEvent);

class ReplySet: public std::set<BonjourReply> {};

struct LifetimeGuard
{
	std::mutex mutex;
	BonjourDialog *dialog;

	LifetimeGuard(BonjourDialog *dialog) : dialog(dialog) {}
};

// Muon printers publish muon_setup's state in the "setup" TXT key: "new", "in_progress" or "complete".
// Anything else, or no key at all (a Muon printer from before the key existed), gets no status.
static bool muon_needs_setup(const BonjourReply &reply)
{
	const auto it = reply.txt_data.find("setup");
	return it != reply.txt_data.end() && (it->second == "new" || it->second == "in_progress");
}

static bool muon_setup_complete(const BonjourReply &reply)
{
	const auto it = reply.txt_data.find("setup");
	return it != reply.txt_data.end() && it->second == "complete";
}

// The name the owner knows the printer by: the "name" TXT key, else its host name.
static wxString muon_display_name(const BonjourReply &reply)
{
	const auto it = reply.txt_data.find("name");
	if (it != reply.txt_data.end())
		return GUI::from_u8(it->second);
	if (!reply.hostname.empty())
		return GUI::from_u8(reply.hostname.substr(0, reply.hostname.find('.')));
	return GUI::from_u8(reply.service_name);
}

// The printer serves its setup page on port 80, whatever port the service record names.
static wxString muon_setup_url(const BonjourReply &reply)
{
	std::string host;
	if (reply.ip.is_v4())
		host = reply.ip.to_string();
	else if (!reply.hostname.empty())
		host = reply.hostname;
	else
		host = "[" + reply.ip.to_string() + "]";
	return GUI::from_u8("http://" + host + "/#/setup");
}

BonjourDialog::BonjourDialog(wxWindow *parent, Slic3r::PrinterTechnology tech, bool muon_hints)
	: wxDialog(parent, wxID_ANY, _(L("Network lookup")), wxDefaultPosition, wxDefaultSize, wxDEFAULT_DIALOG_STYLE|wxRESIZE_BORDER)
	, list(new wxListView(this, wxID_ANY, wxDefaultPosition, wxDefaultSize, wxLC_REPORT|wxSIMPLE_BORDER))
	, replies(new ReplySet)
	, label(new wxStaticText(this, wxID_ANY, ""))
	, timer(new wxTimer())
	, timer_state(0)
	, tech(tech)
	, muon_hints(muon_hints)
{
	const int em = GUI::wxGetApp().em_unit();
	list->SetMinSize(wxSize(80 * em, 30 * em));

	wxBoxSizer *vsizer = new wxBoxSizer(wxVERTICAL);

	vsizer->Add(label, 0, wxEXPAND | wxTOP | wxLEFT | wxRIGHT, em);

	if (muon_hints) {
		// Shown in place of the label when a lookup finds nothing.
		auto *help_title = new wxStaticText(this, wxID_ANY, _L("No printers found on this network."));
		help_title->SetFont(help_title->GetFont().Bold());
		help_title->Wrap(50 * em);

		const wxString bullet = wxString::FromUTF8("\xE2\x80\xA2 "); // U+2022 BULLET
		auto *help_text = new wxStaticText(this, wxID_ANY,
			_L("Look at the printer's screen:") + "\n" +
			bullet + _L("A list of languages or a QR code — it's new. Scan the code with your phone to set it up, then search again.") + "\n" +
			bullet + _L("Its home screen — make sure this computer is on the same network, or type its address (Settings › Network on the printer).") + "\n" +
			bullet + _L("Nothing — check that it's switched on."));
		help_text->Wrap(50 * em);

		help_sizer = new wxBoxSizer(wxVERTICAL);
		help_sizer->Add(help_title, 0, wxBOTTOM, em / 2);
		help_sizer->Add(help_text, 0);
		vsizer->Add(help_sizer, 0, wxEXPAND | wxTOP | wxLEFT | wxRIGHT, em);
		vsizer->Hide(help_sizer);
	}

	list->SetSingleStyle(wxLC_SINGLE_SEL);
	list->SetSingleStyle(wxLC_SORT_DESCENDING);
	list->AppendColumn(_(L("Address")), wxLIST_FORMAT_LEFT, 5 * em);
	list->AppendColumn(_(L("Hostname")), wxLIST_FORMAT_LEFT, 10 * em);
	list->AppendColumn(_(L("Service name")), wxLIST_FORMAT_LEFT, 20 * em);
	if (tech == ptFFF) {
		list->AppendColumn(_(L("OctoPrint version")), wxLIST_FORMAT_LEFT, 5 * em);
	}
	if (muon_hints) {
		status_column = list->GetColumnCount();
		list->AppendColumn(_L("Status"), wxLIST_FORMAT_LEFT, 10 * em);
	}

	vsizer->Add(list, 1, wxEXPAND | wxALL, em);

	wxBoxSizer *button_sizer = new wxBoxSizer(wxHORIZONTAL);
	if (muon_hints) {
		auto *search_again = new wxButton(this, wxID_ANY, _L("Search again"));
		search_again->Bind(wxEVT_BUTTON, [this](wxCommandEvent &) { lookup(); });
		button_sizer->Add(search_again, 0, wxALL, em);
	}
	button_sizer->Add(new wxButton(this, wxID_OK, _L("OK")), 0, wxALL, em);
	button_sizer->Add(new wxButton(this, wxID_CANCEL, _L("Cancel")), 0, wxALL, em);
	// ^ Note: The Ok/Cancel labels are translated by wxWidgets

	vsizer->Add(button_sizer, 0, wxALIGN_CENTER);
	SetSizerAndFit(vsizer);

	Bind(EVT_BONJOUR_REPLY, &BonjourDialog::on_reply, this);
	Bind(EVT_BONJOUR_COMPLETE, &BonjourDialog::on_complete, this);
	Bind(wxEVT_BUTTON, &BonjourDialog::on_ok, this, wxID_OK);

	Bind(wxEVT_TIMER, &BonjourDialog::on_timer, this);
	GUI::wxGetApp().UpdateDlgDarkUI(this);
}

BonjourDialog::~BonjourDialog()
{
	// Needed bacuse of forward defs
}

bool BonjourDialog::show_and_lookup()
{
	Show();   // Because we need GetId() to work before ShowModal()

	lookup();

	bool res = ShowModal() == wxID_OK && list->GetFirstSelected() >= 0;
	{
		// Tell the background thread the dialog is going away...
		std::lock_guard<std::mutex> lock_guard(guard->mutex);
		guard->dialog = nullptr;
	}
	return res;
}

wxString BonjourDialog::get_selected() const
{
	auto sel = list->GetFirstSelected();
	return sel >= 0 ? list->GetItemText(sel) : wxString();
}


// Private

// Starts a lookup, replacing any lookup still running and the results shown so far.
void BonjourDialog::lookup()
{
	if (guard) {
		// Stop the lookup this one replaces from queueing any more events.
		std::lock_guard<std::mutex> lock_guard(guard->mutex);
		guard->dialog = nullptr;
	}
	const unsigned id = ++lookup_id;

	replies->clear();
	rows.clear();
	list->DeleteAllItems();
	if (muon_hints) {
		show_help(false);
	}

	timer->Stop();
	timer->SetOwner(this);
	timer_state = 1;
	timer->Start(1000);
    on_timer_process();

	// The background thread needs to queue messages for this dialog
	// and for that it needs a valid pointer to it (mandated by the wxWidgets API).
	// Here we put the pointer under a shared_ptr and protect it by a mutex,
	// so that both threads can access it safely.
	guard = std::make_shared<LifetimeGuard>(this);
	auto dguard = guard;

	// Ask for the TXT keys we care about
	// - "version" for OctoPrint version column
	// - "model" for FFF/SLA filtering
	// - "addr_pref" optional hint (e.g. "hostname" or "ip") from the service
	// - "name" and "setup", published by Muon printers, for the name and status columns
	Bonjour::TxtKeys txt_keys { "version", "model", "addr_pref" };
	if (muon_hints) {
		txt_keys.insert("name");
		txt_keys.insert("setup");
	}

    bonjour = Bonjour("octoprint")
		.set_txt_keys(std::move(txt_keys))
		.set_retries(3)
		.set_timeout(4)
		.on_reply([dguard, id](BonjourReply &&reply) {
			std::lock_guard<std::mutex> lock_guard(dguard->mutex);
			auto dialog = dguard->dialog;
			if (dialog != nullptr) {
				auto evt = new BonjourReplyEvent(EVT_BONJOUR_REPLY, dialog->GetId(), std::move(reply), id);
				wxQueueEvent(dialog, evt);
			}
		})
		.on_complete([dguard, id]() {
			std::lock_guard<std::mutex> lock_guard(dguard->mutex);
			auto dialog = dguard->dialog;
			if (dialog != nullptr) {
				auto evt = new wxCommandEvent(EVT_BONJOUR_COMPLETE, dialog->GetId());
				evt->SetInt(int(id));
				wxQueueEvent(dialog, evt);
			}
		})
		.lookup();
}

const BonjourReply *BonjourDialog::selected_reply() const
{
	const long sel = list->GetFirstSelected();
	if (sel < 0) {
		return nullptr;
	}
	const size_t row = size_t(list->GetItemData(sel));
	return row < rows.size() ? rows[row] : nullptr;
}

void BonjourDialog::show_help(bool show)
{
	if (help_sizer == nullptr || GetSizer()->IsShown(help_sizer) == show) {
		return;
	}
	GetSizer()->Show(label, !show);
	GetSizer()->Show(help_sizer, show);
	// Grow to fit the help, but never shrink a dialog the user has resized.
	wxSize size = GetSize();
	size.IncTo(GetSizer()->ComputeFittingWindowSize(this));
	SetSize(size);
	Layout();
}

void BonjourDialog::on_reply(BonjourReplyEvent &e)
{
	if (e.lookup_id != lookup_id) {
		// From a lookup that "Search again" replaced
		return;
	}

	if (replies->find(e.reply) != replies->end()) {
		// We already have this reply
		return;
	}

	// Filter replies based on selected technology
	const auto model = e.reply.txt_data.find("model");
	const bool sl1 = model != e.reply.txt_data.end() && model->second == "SL1";
	if ((tech == ptFFF && sl1) || (tech == ptSLA && !sl1)) {
		return;
	}

	replies->insert(std::move(e.reply));

	auto selected = get_selected();

	wxWindowUpdateLocker freeze_guard(this);
	(void)freeze_guard;

	list->DeleteAllItems();
	rows.clear();

	// The whole list is recreated so that we benefit from it already being sorted in the set.
	// (And also because wxListView's sorting API is bananas.)
	for (const auto &reply : *replies) {
		// Check for optional preference provided by the service
		bool prefer_hostname = false;
		{
			auto it_pref = reply.txt_data.find("addr_pref");
			if (it_pref != reply.txt_data.end() && it_pref->second == "hostname") {
				prefer_hostname = true;
			}
		}

		// Column 0 ("Address") is what get_selected() will return.
		// If addr_pref=hostname, put hostname there and IP in column 1.
		// Otherwise keep current behavior: IP in col0, hostname in col1.
		wxString primary   = prefer_hostname ? reply.hostname    : reply.full_address;
		wxString secondary = prefer_hostname ? reply.full_address : reply.hostname;

		auto item = list->InsertItem(0, primary);
		list->SetItemData(item, long(rows.size()));
		rows.push_back(&reply);
		list->SetItem(item, 1, secondary);

		const auto it_name = muon_hints ? reply.txt_data.find("name") : reply.txt_data.end();
		list->SetItem(item, 2, it_name != reply.txt_data.end() ? GUI::from_u8(it_name->second) : wxString(reply.service_name));

		if (tech == ptFFF) {
			const auto it = reply.txt_data.find("version");
			if (it != reply.txt_data.end()) {
				list->SetItem(item, 3, GUI::from_u8(it->second));
			}
		}

		if (muon_hints) {
			if (muon_needs_setup(reply)) {
				list->SetItem(item, status_column, _L("Needs setup"));
			} else if (muon_setup_complete(reply)) {
				list->SetItem(item, status_column, _L("Ready"));
			}
		}
	}


	const int em = GUI::wxGetApp().em_unit();

	for (int i = 0; i < list->GetColumnCount(); i++) {
		list->SetColumnWidth(i, wxLIST_AUTOSIZE);
		if (list->GetColumnWidth(i) < 10 * em) { list->SetColumnWidth(i, 10 * em); }
	}

	if (!selected.IsEmpty()) {
		// Attempt to preserve selection
		auto hit = list->FindItem(-1, selected);
		if (hit >= 0) { list->SetItemState(hit, wxLIST_STATE_SELECTED, wxLIST_STATE_SELECTED); }
	}
}

void BonjourDialog::on_complete(wxCommandEvent &e)
{
	if (unsigned(e.GetInt()) != lookup_id) {
		// From a lookup that "Search again" replaced
		return;
	}

	timer_state = 0;
	if (muon_hints) {
		show_help(list->GetItemCount() == 0);
	}
}

void BonjourDialog::on_ok(wxCommandEvent &e)
{
	const BonjourReply *reply = muon_hints ? selected_reply() : nullptr;
	if (reply != nullptr && muon_needs_setup(*reply)) {
		// Don't block: the owner can still use a printer that isn't set up.
		GUI::MessageDialog dialog(this,
			GUI::format_wxstr(_L("%1% isn't set up yet. Finish setup on its screen, or open its setup page in your browser."), muon_display_name(*reply)),
			wxEmptyString, wxYES | wxNO | wxICON_QUESTION);
		dialog.SetButtonLabel(wxID_YES, _L("Open setup page"));
		dialog.SetButtonLabel(wxID_NO, _L("Use it anyway"));
		const int res = dialog.ShowModal();
		if (res == wxID_YES) {
			// Stay in the list, so the owner can search again once setup is done.
			wxLaunchDefaultBrowser(muon_setup_url(*reply));
			return;
		}
		if (res != wxID_NO) {
			return;
		}
	}
	e.Skip();
}

void BonjourDialog::on_timer(wxTimerEvent &)
{
    on_timer_process();
}

// This is here so the function can be bound to wxEVT_TIMER and also called
// explicitly (wxTimerEvent should not be created by user code).
void BonjourDialog::on_timer_process()
{
    const auto search_str = _L("Searching for devices");

    if (timer_state > 0) {
        const std::string dots(timer_state, '.');
        label->SetLabel(search_str + dots);
        timer_state = (timer_state) % 3 + 1;
    } else {
        label->SetLabel(search_str + ": " + _L("Finished") + ".");
        timer->Stop();
    }
}

IPListDialog::IPListDialog(wxWindow* parent, const wxString& hostname, const std::vector<boost::asio::ip::address>& ips, size_t& selected_index)
	: wxDialog(parent, wxID_ANY, _(L("Multiple resolved IP addresses")), wxDefaultPosition, wxDefaultSize, wxDEFAULT_DIALOG_STYLE | wxRESIZE_BORDER)
	, m_list(new wxListView(this, wxID_ANY, wxDefaultPosition, wxDefaultSize, wxLC_REPORT | wxSIMPLE_BORDER))
	, m_selected_index (selected_index)
{
	const int em = GUI::wxGetApp().em_unit();
	m_list->SetMinSize(wxSize(40 * em, 30 * em));

	wxBoxSizer* vsizer = new wxBoxSizer(wxVERTICAL);

	auto* label = new wxStaticText(this, wxID_ANY, GUI::format_wxstr(_L("There are several IP addresses resolving to hostname %1%.\nPlease select one that should be used."), hostname));
	vsizer->Add(label, 0, wxEXPAND | wxTOP | wxLEFT | wxRIGHT, em);

	m_list->SetSingleStyle(wxLC_SINGLE_SEL);
	m_list->AppendColumn(_(L("Address")), wxLIST_FORMAT_LEFT, 40 * em);

	for (size_t i = 0; i < ips.size(); i++) 
		m_list->InsertItem(i, boost::nowide::widen(ips[i].to_string()));

	m_list->Select(0);

	vsizer->Add(m_list, 1, wxEXPAND | wxALL, em);

	wxBoxSizer* button_sizer = new wxBoxSizer(wxHORIZONTAL);
	button_sizer->Add(new wxButton(this, wxID_OK, _L("OK")), 0, wxALL, em);
	button_sizer->Add(new wxButton(this, wxID_CANCEL, _L("Cancel")), 0, wxALL, em);

	vsizer->Add(button_sizer, 0, wxALIGN_CENTER);
	SetSizerAndFit(vsizer);

	GUI::wxGetApp().UpdateDlgDarkUI(this);
}

IPListDialog::~IPListDialog()
{
}

void IPListDialog::EndModal(int retCode)
{
	if (retCode == wxID_OK) {
		m_selected_index = (size_t)m_list->GetFirstSelected();
	}
	wxDialog::EndModal(retCode);
}

}
