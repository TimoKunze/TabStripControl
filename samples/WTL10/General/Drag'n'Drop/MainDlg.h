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
		CF_TARGETCLSID = static_cast<CLIPFORMAT>(RegisterClipboardFormat(CFSTR_TARGETCLSID));
	}

	BOOL bRightDrag;
	LONG lastDropTarget;

	struct Controls
	{
		CButton oleDDCheck;
		CImageList imageList;
		CComPtr<ITabStrip> tabstripU;

		~Controls()
		{
			if(!imageList.IsNull()) {
				imageList.Destroy();
			}
		}
	} controls;

	CLIPFORMAT CF_TARGETCLSID;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		COMMAND_ID_HANDLER(ID_OPTIONS_AERODRAGIMAGES, OnOptionsAeroDragImages)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 1, AbortedDragTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 8, DragMouseMoveTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 9, DropTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 13, KeyDownTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 23, MouseUpTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 24, OLECompleteDragTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 25, OLEDragDropTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 26, OLEDragEnterTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 27, OLEDragLeaveTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 28, OLEDragMouseMoveTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 31, OLESetDataTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 32, OLEStartDragTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 36, RecreatedControlWindowTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 40, TabBeginDragTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(_ITabStripEvents), 41, TabBeginRDragTabstripu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_TABSTRIPU, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_OLEDDCHECK, DLSZ_MOVE_X)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnOptionsAeroDragImages(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void InsertTabs(void);
	void UpdateMenu(void);

	void __stdcall AbortedDragTabstripu();
	void __stdcall DragMouseMoveTabstripu(LPDISPATCH* dropTarget, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/, VARIANT_BOOL* autoActivateTab, long* /*autoScrollVelocity*/);
	void __stdcall DropTabstripu(LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/);
	void __stdcall KeyDownTabstripu(short* keyCode, short /*shift*/);
	void __stdcall MouseUpTabstripu(LPDISPATCH /*tsTab*/, short button, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/);
	void __stdcall OLECompleteDragTabstripu(LPDISPATCH data, long /*performedEffect*/);
	void __stdcall OLEDragDropTabstripu(LPDISPATCH data, long* effect, LPDISPATCH /*dropTarget*/, short /*button*/, short shift, long /*x*/, long /*y*/, long /*hitTestDetails*/);
	void __stdcall OLEDragEnterTabstripu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long x, long y, long /*hitTestDetails*/, VARIANT_BOOL* autoActivateTab, long* /*autoScrollVelocity*/);
	void __stdcall OLEDragLeaveTabstripu(LPDISPATCH /*data*/, LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/);
	void __stdcall OLEDragMouseMoveTabstripu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long x, long y, long /*hitTestDetails*/, VARIANT_BOOL* autoActivateTab, long* /*autoScrollVelocity*/);
	void __stdcall OLESetDataTabstripu(LPDISPATCH data, long formatID, long /*index*/, long /*dataOrViewAspect*/);
	void __stdcall OLEStartDragTabstripu(LPDISPATCH data);
	void __stdcall RecreatedControlWindowTabstripu(long /*hWnd*/);
	void __stdcall TabBeginDragTabstripu(LPDISPATCH tsTab, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/);
	void __stdcall TabBeginRDragTabstripu(LPDISPATCH tsTab, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/);
};
