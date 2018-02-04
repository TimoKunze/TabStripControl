// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:E7BB2F30-C5DD-4370-B7E2-19A7EDF169EE> version("1.10") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_TABSTRIPU, CMainDlg, &__uuidof(_ITabStripEvents), &LIBID_TabStripCtlLibU, 1, 10>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> tabStripUWnd;

	CMainDlg() : tabStripUWnd(this, 1)
	{
	}

	struct Controls
	{
		CImageList imageList;
		CComPtr<ITabStrip> tabstripU;

		~Controls()
		{
			if(!imageList.IsNull()) {
				imageList.Destroy();
			}
		}
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 4, ClickTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 36, RecreatedControlWindowTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 42, TabGetInfoTipTextTabstripu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_TABSTRIPU, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void InsertTabs(void);

	void __stdcall ClickTabstripu(LPDISPATCH tsTab, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails);
	void __stdcall RecreatedControlWindowTabstripu(long /*hWnd*/);
	void __stdcall TabGetInfoTipTextTabstripu(LPDISPATCH /*tsTab*/, long /*maxInfoTipLength*/, BSTR* infoTipText);
};
