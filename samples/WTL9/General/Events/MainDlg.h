// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////
#pragma once
#include <initguid.h>

#import <libid:E7BB2F30-C5DD-4370-B7E2-19A7EDF169EE> version("1.10") raw_dispinterfaces
#import <libid:3DF28271-B5ED-46F4-9EFF-E60A4CDD0D02> version("1.10") raw_dispinterfaces

DEFINE_GUID(LIBID_TabStripCtlLibU, 0xE7BB2F30, 0xC5DD, 0x4370, 0xB7, 0xE2, 0x19, 0xA7, 0xED, 0xF1, 0x69, 0xEE);
DEFINE_GUID(LIBID_TabStripCtlLibA, 0x3DF28271, 0xB5ED, 0x46F4, 0x9E, 0xFF, 0xE6, 0x0A, 0x4C, 0xDD, 0x0D, 0x02);

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_TABSTRIPU, CMainDlg, &__uuidof(TabStripCtlLibU::_ITabStripEvents), &LIBID_TabStripCtlLibU, 1, 10>,
    public IDispEventImpl<IDC_TABSTRIPA, CMainDlg, &__uuidof(TabStripCtlLibA::_ITabStripEvents), &LIBID_TabStripCtlLibA, 1, 10>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> tabstripUContainerWnd;
	CContainedWindowT<CWindow> tabstripUWnd;
	CContainedWindowT<CAxWindow> tabstripAContainerWnd;
	CContainedWindowT<CWindow> tabstripAWnd;

	BOOL tabstripAIsFocused;

	CMainDlg() :
	    tabstripUContainerWnd(this, 1),
	    tabstripAContainerWnd(this, 2),
	    tabstripUWnd(this, 3),
	    tabstripAWnd(this, 4)
	{
	}

	struct Controls
	{
		CEdit logEdit;
		CButton aboutButton;
		CImageList imageList;
		CComPtr<TabStripCtlLibU::ITabStrip> tabstripU;
		CComPtr<TabStripCtlLibA::ITabStrip> tabstripA;

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
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusU)

		ALT_MSG_MAP(4)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusA)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 1, AbortedDragTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 0, ActiveTabChangedTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 2, ActiveTabChangingTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 3, CaretChangedTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 4, ClickTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 5, ContextMenuTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 6, DblClickTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 7, DestroyedControlWindowTabstripu)
		//SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 8, DragMouseMoveTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 9, DropTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 10, FreeTabDataTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 11, InsertedTabTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 12, InsertingTabTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 13, KeyDownTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 14, KeyPressTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 15, KeyUpTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 16, MClickTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 17, MDblClickTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 18, MouseDownTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 19, MouseEnterTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 20, MouseHoverTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 21, MouseLeaveTabstripu)
		//SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 22, MouseMoveTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 23, MouseUpTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 24, OLECompleteDragTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 25, OLEDragDropTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 26, OLEDragEnterTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 27, OLEDragLeaveTabstripu)
		//SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 28, OLEDragMouseMoveTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 29, OLEGiveFeedbackTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 30, OLEQueryContinueDragTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 31, OLESetDataTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 32, OLEStartDragTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 33, OwnerDrawTabTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 34, RClickTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 35, RDblClickTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 36, RecreatedControlWindowTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 37, RemovedTabTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 38, RemovingTabTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 39, ResizedControlWindowTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 40, TabBeginDragTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 41, TabBeginRDragTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 42, TabGetInfoTipTextTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 43, TabMouseEnterTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 44, TabMouseLeaveTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 45, TabSelectionChangedTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 46, OLEDragEnterPotentialTargetTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 47, OLEDragLeavePotentialTargetTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 48, OLEReceivedNewDataTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 49, XClickTabstripu)
		SINK_ENTRY_EX(IDC_TABSTRIPU, __uuidof(TabStripCtlLibU::_ITabStripEvents), 50, XDblClickTabstripu)

		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 1, AbortedDragTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 0, ActiveTabChangedTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 2, ActiveTabChangingTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 3, CaretChangedTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 4, ClickTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 5, ContextMenuTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 6, DblClickTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 7, DestroyedControlWindowTabstripa)
		//SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 8, DragMouseMoveTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 9, DropTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 10, FreeTabDataTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 11, InsertedTabTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 12, InsertingTabTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 13, KeyDownTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 14, KeyPressTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 15, KeyUpTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 16, MClickTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 17, MDblClickTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 18, MouseDownTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 19, MouseEnterTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 20, MouseHoverTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 21, MouseLeaveTabstripa)
		//SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 22, MouseMoveTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 23, MouseUpTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 24, OLECompleteDragTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 25, OLEDragDropTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 26, OLEDragEnterTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 27, OLEDragLeaveTabstripa)
		//SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 28, OLEDragMouseMoveTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 29, OLEGiveFeedbackTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 30, OLEQueryContinueDragTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 31, OLESetDataTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 32, OLEStartDragTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 33, OwnerDrawTabTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 34, RClickTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 35, RDblClickTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 36, RecreatedControlWindowTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 37, RemovedTabTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 38, RemovingTabTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 39, ResizedControlWindowTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 40, TabBeginDragTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 41, TabBeginRDragTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 42, TabGetInfoTipTextTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 43, TabMouseEnterTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 44, TabMouseLeaveTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 45, TabSelectionChangedTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 46, OLEDragEnterPotentialTargetTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 47, OLEDragLeavePotentialTargetTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 48, OLEReceivedNewDataTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 49, XClickTabstripa)
		SINK_ENTRY_EX(IDC_TABSTRIPA, __uuidof(TabStripCtlLibA::_ITabStripEvents), 50, XDblClickTabstripa)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusU(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusA(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);

	void AddLogEntry(CAtlString text);
	void CloseDialog(int nVal);
	void InsertTabsA(void);
	void InsertTabsU(void);

	void __stdcall AbortedDragTabstripu();
	void __stdcall ActiveTabChangedTabstripu(LPDISPATCH previousActiveTab, LPDISPATCH newActiveTab);
	void __stdcall ActiveTabChangingTabstripu(LPDISPATCH previousActiveTab, VARIANT_BOOL* cancelChange);
	void __stdcall CaretChangedTabstripu(LPDISPATCH previousCaretTab, LPDISPATCH newCaretTab);
	void __stdcall ClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ContextMenuTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall DblClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall DestroyedControlWindowTabstripu(long hWnd);
	void __stdcall DragMouseMoveTabstripu(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity);
	void __stdcall DropTabstripu(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall FreeTabDataTabstripu(LPDISPATCH tsTab);
	void __stdcall InsertedTabTabstripu(LPDISPATCH tsTab);
	void __stdcall InsertingTabTabstripu(LPDISPATCH tsTab, VARIANT_BOOL* cancelInsertion);
	void __stdcall KeyDownTabstripu(short* keyCode, short shift);
	void __stdcall KeyPressTabstripu(short* keyAscii);
	void __stdcall KeyUpTabstripu(short* keyCode, short shift);
	void __stdcall MClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MDblClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseDownTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseEnterTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseHoverTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseLeaveTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseMoveTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseUpTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLECompleteDragTabstripu(LPDISPATCH data, long performedEffect);
	void __stdcall OLEDragDropTabstripu(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragEnterTabstripu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetTabstripu(long hWndPotentialTarget);
	void __stdcall OLEDragLeaveTabstripu(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetTabstripu(void);
	void __stdcall OLEDragMouseMoveTabstripu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity);
	void __stdcall OLEGiveFeedbackTabstripu(long effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragTabstripu(BOOL pressedEscape, short button, short shift, long* actionToContinueWith);
	void __stdcall OLEReceivedNewDataTabstripu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLESetDataTabstripu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLEStartDragTabstripu(LPDISPATCH data);
	void __stdcall OwnerDrawTabTabstripu(LPDISPATCH tsTab, TabStripCtlLibU::OwnerDrawTabStateConstants tabState, long hDC, TabStripCtlLibU::RECTANGLE* drawingRectangle);
	void __stdcall RClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RDblClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RecreatedControlWindowTabstripu(long hWnd);
	void __stdcall RemovedTabTabstripu(LPDISPATCH tsTab);
	void __stdcall RemovingTabTabstripu(LPDISPATCH tsTab, VARIANT_BOOL* cancelDeletion);
	void __stdcall ResizedControlWindowTabstripu();
	void __stdcall TabBeginDragTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall TabBeginRDragTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall TabGetInfoTipTextTabstripu(LPDISPATCH tsTab, long maxInfoTipLength, BSTR* infoTipText);
	void __stdcall TabMouseEnterTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall TabMouseLeaveTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall TabSelectionChangedTabstripu(LPDISPATCH tsTab);
	void __stdcall XClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall XDblClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);

	void __stdcall AbortedDragTabstripa();
	void __stdcall ActiveTabChangedTabstripa(LPDISPATCH previousActiveTab, LPDISPATCH newActiveTab);
	void __stdcall ActiveTabChangingTabstripa(LPDISPATCH previousActiveTab, VARIANT_BOOL* cancelChange);
	void __stdcall CaretChangedTabstripa(LPDISPATCH previousCaretTab, LPDISPATCH newCaretTab);
	void __stdcall ClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ContextMenuTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall DblClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall DestroyedControlWindowTabstripa(long hWnd);
	void __stdcall DragMouseMoveTabstripa(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity);
	void __stdcall DropTabstripa(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall FreeTabDataTabstripa(LPDISPATCH tsTab);
	void __stdcall InsertedTabTabstripa(LPDISPATCH tsTab);
	void __stdcall InsertingTabTabstripa(LPDISPATCH tsTab, VARIANT_BOOL* cancelInsertion);
	void __stdcall KeyDownTabstripa(short* keyCode, short shift);
	void __stdcall KeyPressTabstripa(short* keyAscii);
	void __stdcall KeyUpTabstripa(short* keyCode, short shift);
	void __stdcall MClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MDblClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseDownTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseEnterTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseHoverTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseLeaveTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseMoveTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseUpTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLECompleteDragTabstripa(LPDISPATCH data, long performedEffect);
	void __stdcall OLEDragDropTabstripa(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragEnterTabstripa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetTabstripa(long hWndPotentialTarget);
	void __stdcall OLEDragLeaveTabstripa(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetTabstripa(void);
	void __stdcall OLEDragMouseMoveTabstripa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity);
	void __stdcall OLEGiveFeedbackTabstripa(long effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragTabstripa(BOOL pressedEscape, short button, short shift, long* actionToContinueWith);
	void __stdcall OLEReceivedNewDataTabstripa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLESetDataTabstripa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLEStartDragTabstripa(LPDISPATCH data);
	void __stdcall OwnerDrawTabTabstripa(LPDISPATCH tsTab, TabStripCtlLibA::OwnerDrawTabStateConstants tabState, long hDC, TabStripCtlLibA::RECTANGLE* drawingRectangle);
	void __stdcall RClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RDblClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RecreatedControlWindowTabstripa(long hWnd);
	void __stdcall RemovedTabTabstripa(LPDISPATCH tsTab);
	void __stdcall RemovingTabTabstripa(LPDISPATCH tsTab, VARIANT_BOOL* cancelDeletion);
	void __stdcall ResizedControlWindowTabstripa();
	void __stdcall TabBeginDragTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall TabBeginRDragTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall TabGetInfoTipTextTabstripa(LPDISPATCH tsTab, long maxInfoTipLength, BSTR* infoTipText);
	void __stdcall TabMouseEnterTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall TabMouseLeaveTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall TabSelectionChangedTabstripa(LPDISPATCH tsTab);
	void __stdcall XClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall XDblClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails);
};
