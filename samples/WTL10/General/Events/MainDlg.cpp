// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	return CWindow::IsDialogMessage(pMsg);
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	if(controls.tabstripU) {
		IDispEventImpl<IDC_TABSTRIPU, CMainDlg, &__uuidof(TabStripCtlLibU::_ITabStripEvents), &LIBID_TabStripCtlLibU, 1, 10>::DispEventUnadvise(controls.tabstripU);
		controls.tabstripU.Release();
	}
	if(controls.tabstripA) {
		IDispEventImpl<IDC_TABSTRIPA, CMainDlg, &__uuidof(TabStripCtlLibA::_ITabStripEvents), &LIBID_TabStripCtlLibA, 1, 10>::DispEventUnadvise(controls.tabstripA);
		controls.tabstripA.Release();
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// Init resizing
	DlgResize_Init(false, false);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	controls.imageList.CreateFromImage(IDB_ICONS, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.logEdit = GetDlgItem(IDC_EDITLOG);
	controls.aboutButton = GetDlgItem(ID_APP_ABOUT);

	tabstripUContainerWnd.SubclassWindow(GetDlgItem(IDC_TABSTRIPU));
	tabstripUContainerWnd.QueryControl(__uuidof(TabStripCtlLibU::ITabStrip), reinterpret_cast<LPVOID*>(&controls.tabstripU));
	if(controls.tabstripU) {
		IDispEventImpl<IDC_TABSTRIPU, CMainDlg, &__uuidof(TabStripCtlLibU::_ITabStripEvents), &LIBID_TabStripCtlLibU, 1, 10>::DispEventAdvise(controls.tabstripU);
		tabstripUWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.tabstripU->GethWnd())));
		InsertTabsU();
	}

	tabstripAContainerWnd.SubclassWindow(GetDlgItem(IDC_TABSTRIPA));
	tabstripAContainerWnd.QueryControl(__uuidof(TabStripCtlLibA::ITabStrip), reinterpret_cast<LPVOID*>(&controls.tabstripA));
	if(controls.tabstripA) {
		IDispEventImpl<IDC_TABSTRIPA, CMainDlg, &__uuidof(TabStripCtlLibA::_ITabStripEvents), &LIBID_TabStripCtlLibA, 1, 10>::DispEventAdvise(controls.tabstripA);
		tabstripAWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.tabstripA->GethWnd())));
		InsertTabsA();
	}

	// force control resize
	WINDOWPOS dummy = {0};
	BOOL b = FALSE;
	OnWindowPosChanged(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&dummy), b);

	return TRUE;
}

LRESULT CMainDlg::OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.logEdit && controls.aboutButton) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			CRect rc;
			GetClientRect(&rc);
			int cx = static_cast<int>(0.4 * static_cast<double>(rc.Width()));
			controls.logEdit.SetWindowPos(NULL, rc.Width() - cx, 0, cx, rc.Height() - 32, 0);
			controls.aboutButton.SetWindowPos(NULL, rc.Width() - cx, rc.Height() - 27, 0, 0, SWP_NOSIZE);
			tabstripUContainerWnd.SetWindowPos(NULL, 0, 0, rc.Width() - cx - 5, (rc.Height() - 5) / 2, SWP_NOMOVE);
			tabstripAContainerWnd.SetWindowPos(NULL, 0, (rc.Height() - 5) / 2 + 5, rc.Width() - cx - 5, (rc.Height() - 5) / 2, 0);
		}
	}

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(tabstripAIsFocused) {
		if(controls.tabstripA) {
			controls.tabstripA->About();
		}
	} else {
		if(controls.tabstripU) {
			controls.tabstripU->About();
		}
	}
	return 0;
}

LRESULT CMainDlg::OnSetFocusU(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	tabstripAIsFocused = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusA(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	tabstripAIsFocused = TRUE;
	return 0;
}

void CMainDlg::AddLogEntry(CAtlString text)
{
	static int cLines = 0;
	static CAtlString oldText;

	cLines++;
	if(cLines > 50) {
		// delete the first line
		int pos = oldText.Find(TEXT("\r\n"));
		oldText = oldText.Mid(pos + lstrlen(TEXT("\r\n")), oldText.GetLength());
		cLines--;
	}
	oldText += text;
	oldText += TEXT("\r\n");

	controls.logEdit.SetWindowText(oldText);
	int l = oldText.GetLength();
	controls.logEdit.SetSel(l, l);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::InsertTabsA(void)
{
	controls.tabstripA->PuthImageList(HandleToLong(controls.imageList.m_hImageList));

	CComPtr<TabStripCtlLibA::ITabStripTabs> pTabs = controls.tabstripA->GetTabs();
	pTabs->Add(TEXT("Tab 1"), -1, 0, 1);
	pTabs->Add(TEXT("Tab 2"), -1, 1, 2);
	pTabs->Add(TEXT("Tab 3"), -1, 2, 3);
	pTabs->Add(TEXT("Tab 4"), -1, 3, 4);
	pTabs->Add(TEXT("Tab 5"), -1, 4, 5);
	pTabs->Add(TEXT("Tab 6"), -1, 5, 6);
	pTabs->Add(TEXT("Tab 7"), -1, 6, 7);
	pTabs->Add(TEXT("Tab 8"), -1, 7, 8);
	pTabs->Add(TEXT("Tab 9"), -1, 8, 9);
	pTabs->Add(TEXT("Tab 10"), -1, 9, 10);
}

void CMainDlg::InsertTabsU(void)
{
	controls.tabstripU->PuthImageList(HandleToLong(controls.imageList.m_hImageList));

	CComPtr<TabStripCtlLibU::ITabStripTabs> pTabs = controls.tabstripU->GetTabs();
	pTabs->Add(TEXT("Tab 1"), -1, 0, 1);
	pTabs->Add(TEXT("Tab 2"), -1, 1, 2);
	pTabs->Add(TEXT("Tab 3"), -1, 2, 3);
	pTabs->Add(TEXT("Tab 4"), -1, 3, 4);
	pTabs->Add(TEXT("Tab 5"), -1, 4, 5);
	pTabs->Add(TEXT("Tab 6"), -1, 5, 6);
	pTabs->Add(TEXT("Tab 7"), -1, 6, 7);
	pTabs->Add(TEXT("Tab 8"), -1, 7, 8);
	pTabs->Add(TEXT("Tab 9"), -1, 8, 9);
	pTabs->Add(TEXT("Tab 10"), -1, 9, 10);
}

void __stdcall CMainDlg::AbortedDragTabstripu()
{
	AddLogEntry(CAtlString(TEXT("TabStripU_AbortedDrag")));

	controls.tabstripU->PutRefDropHilitedTab(NULL);
}

void __stdcall CMainDlg::ActiveTabChangedTabstripu(LPDISPATCH previousActiveTab, LPDISPATCH newActiveTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pPrevTab = previousActiveTab;
	if(pPrevTab) {
		BSTR text = pPrevTab->GetText();
		str += TEXT("TabStripU_ActiveTabChanged: previousActiveTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_ActiveTabChanged: previousActiveTab=NULL");
	}
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pNewTab = newActiveTab;
	if(pNewTab) {
		BSTR text = pNewTab->GetText();
		str += TEXT(", newActiveTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newActiveTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ActiveTabChangingTabstripu(LPDISPATCH previousActiveTab, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pPrevTab = previousActiveTab;
	if(pPrevTab) {
		BSTR text = pPrevTab->GetText();
		str += TEXT("TabStripU_ActiveTabChanging: previousActiveTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_ActiveTabChanging: previousActiveTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelChange=%i"), *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CaretChangedTabstripu(LPDISPATCH previousCaretTab, LPDISPATCH newCaretTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pPrevTab = previousCaretTab;
	if(pPrevTab) {
		BSTR text = pPrevTab->GetText();
		str += TEXT("TabStripU_CaretChanged: previousCaretTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_CaretChanged: previousCaretTab=NULL");
	}
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pNewTab = newCaretTab;
	if(pNewTab) {
		BSTR text = pNewTab->GetText();
		str += TEXT(", newCaretTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_Click: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_Click: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_ContextMenu: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_ContextMenu: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_DblClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_DblClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowTabstripu(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TabStripU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveTabstripu(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = *dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_DragMouseMove: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_DragMouseMove: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoActivateTab=%i, autoScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoActivateTab, *autoScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropTabstripu(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_Drop: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_Drop: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeTabDataTabstripu(LPDISPATCH tsTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_FreeTabData: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_FreeTabData: tsTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedTabTabstripu(LPDISPATCH tsTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_InsertedTab: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_InsertedTab: tsTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingTabTabstripu(LPDISPATCH tsTab, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_InsertingTab: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_InsertingTab: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownTabstripu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TabStripU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);

	if(*keyCode == VK_ESCAPE) {
		if(controls.tabstripU->GetDraggedTabs()) {
			//controls.tabstripU->EndDrag(VARIANT_TRUE);
		}
	}
}

void __stdcall CMainDlg::KeyPressTabstripu(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("TabStripU_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpTabstripu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TabStripU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_MClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_MClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_MDblClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_MDblClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_MouseDown: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_MouseDown: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_MouseEnter: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_MouseEnter: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_MouseHover: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_MouseHover: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_MouseLeave: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_MouseLeave: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_MouseMove: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_MouseMove: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_MouseUp: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_MouseUp: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragTabstripu(LPDISPATCH data, long performedEffect)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripU_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripU_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropTabstripu(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.tabstripU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterTabstripu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = *dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoActivateTab=%i, autoScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoActivateTab, *autoScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetTabstripu(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("TabStripU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveTabstripu(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripU_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetTabstripu(void)
{
	AddLogEntry(TEXT("TabStripU_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveTabstripu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = *dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoActivateTab=%i, autoScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoActivateTab, *autoScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackTabstripu(long effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("TabStripU_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragTabstripu(BOOL pressedEscape, short button, short shift, long* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("TabStripU_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataTabstripu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripU_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripU_OLEReceivedNewData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataTabstripu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripU_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripU_OLESetData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragTabstripu(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripU_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripU_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::OwnerDrawTabTabstripu(LPDISPATCH tsTab, TabStripCtlLibU::OwnerDrawTabStateConstants tabState, long hDC, TabStripCtlLibU::RECTANGLE* drawingRectangle)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_OwnerDrawTab: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_OwnerDrawTab: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", tabState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i)"), tabState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_RClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_RClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_RDblClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_RDblClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowTabstripu(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TabStripU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
	InsertTabsU();
}

void __stdcall CMainDlg::RemovedTabTabstripu(LPDISPATCH tsTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_RemovedTab: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_RemovedTab: tsTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingTabTabstripu(LPDISPATCH tsTab, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_RemovingTab: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_RemovingTab: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowTabstripu()
{
	AddLogEntry(CAtlString(TEXT("TabStripU_ResizedControlWindow")));
}

void __stdcall CMainDlg::TabBeginDragTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_TabBeginDrag: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_TabBeginDrag: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(controls.tabstripU->GetMultiSelect() == VARIANT_FALSE) {
		controls.tabstripU->OLEDrag(NULL, TabStripCtlLibU::odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.tabstripU->CreateTabContainer(_variant_t(tsTab)), -1);
	} else {
		CComPtr<TabStripCtlLibU::ITabStripTabs> pTabs = controls.tabstripU->GetTabs();
		pTabs->PutFilterType(TabStripCtlLibU::fpSelected, TabStripCtlLibU::ftIncluding);
		_variant_t filter;
		filter.Clear();
		CComSafeArray<VARIANT> arr;
		arr.Create(1, 1);
		arr.SetAt(1, _variant_t(true));
		filter.parray = arr.Detach();
		filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: TabStrip expects an array of VARIANTs!
		pTabs->PutFilter(TabStripCtlLibU::fpSelected, filter);

		controls.tabstripU->OLEDrag(NULL, TabStripCtlLibU::odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.tabstripU->CreateTabContainer(_variant_t(pTabs)), -1);
	}
}

void __stdcall CMainDlg::TabBeginRDragTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_TabBeginRDrag: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_TabBeginRDrag: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::TabGetInfoTipTextTabstripu(LPDISPATCH tsTab, long maxInfoTipLength, BSTR* infoTipText)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_TabGetInfoTipText: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_TabGetInfoTipText: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", maxInfoTipLength=%i, infoTipText=%s"), maxInfoTipLength, OLE2W(*infoTipText));
	str += tmp;

	AddLogEntry(str);

	tmp.Format(TEXT("ID: %i, TabData: 0x%X"), pTab->GetID(), pTab->GetTabData());
	*infoTipText = _bstr_t(tmp).Detach();
}

void __stdcall CMainDlg::TabMouseEnterTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_TabMouseEnter: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_TabMouseEnter: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::TabMouseLeaveTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_TabMouseLeave: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_TabMouseLeave: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::TabSelectionChangedTabstripu(LPDISPATCH tsTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_TabSelectionChanged: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_TabSelectionChanged: tsTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_XClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_XClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickTabstripu(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibU::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripU_XDblClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripU_XDblClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}


void __stdcall CMainDlg::AbortedDragTabstripa()
{
	AddLogEntry(CAtlString(TEXT("TabStripA_AbortedDrag")));

	//controls.tabstripA->PutRefDropHilitedTab(NULL);
}

void __stdcall CMainDlg::ActiveTabChangedTabstripa(LPDISPATCH previousActiveTab, LPDISPATCH newActiveTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pPrevTab = previousActiveTab;
	if(pPrevTab) {
		BSTR text = pPrevTab->GetText();
		str += TEXT("TabStripA_ActiveTabChanged: previousActiveTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_ActiveTabChanged: previousActiveTab=NULL");
	}
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pNewTab = newActiveTab;
	if(pNewTab) {
		BSTR text = pNewTab->GetText();
		str += TEXT(", newActiveTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newActiveTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ActiveTabChangingTabstripa(LPDISPATCH previousActiveTab, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pPrevTab = previousActiveTab;
	if(pPrevTab) {
		BSTR text = pPrevTab->GetText();
		str += TEXT("TabStripA_ActiveTabChanging: previousActiveTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_ActiveTabChanging: previousActiveTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelChange=%i"), *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CaretChangedTabstripa(LPDISPATCH previousCaretTab, LPDISPATCH newCaretTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pPrevTab = previousCaretTab;
	if(pPrevTab) {
		BSTR text = pPrevTab->GetText();
		str += TEXT("TabStripA_CaretChanged: previousCaretTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_CaretChanged: previousCaretTab=NULL");
	}
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pNewTab = newCaretTab;
	if(pNewTab) {
		BSTR text = pNewTab->GetText();
		str += TEXT(", newCaretTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_Click: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_Click: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_ContextMenu: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_ContextMenu: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_DblClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_DblClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowTabstripa(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TabStripA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveTabstripa(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = *dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_DragMouseMove: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_DragMouseMove: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoActivateTab=%i, autoScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoActivateTab, *autoScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropTabstripa(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_Drop: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_Drop: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeTabDataTabstripa(LPDISPATCH tsTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_FreeTabData: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_FreeTabData: tsTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedTabTabstripa(LPDISPATCH tsTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_InsertedTab: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_InsertedTab: tsTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingTabTabstripa(LPDISPATCH tsTab, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_InsertingTab: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_InsertingTab: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownTabstripa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TabStripA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);

	if(controls.tabstripA->GetDraggedTabs()) {
		//controls.tabstripA->EndDrag(VARIANT_TRUE);
	}
}

void __stdcall CMainDlg::KeyPressTabstripa(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("TabStripA_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpTabstripa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TabStripA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_MClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_MClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_MDblClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_MDblClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_MouseDown: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_MouseDown: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_MouseEnter: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_MouseEnter: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_MouseHover: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_MouseHover: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_MouseLeave: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_MouseLeave: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_MouseMove: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_MouseMove: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_MouseUp: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_MouseUp: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragTabstripa(LPDISPATCH data, long performedEffect)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripA_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripA_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropTabstripa(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.tabstripA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterTabstripa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = *dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoActivateTab=%i, autoScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoActivateTab, *autoScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetTabstripa(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("TabStripA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveTabstripa(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripA_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetTabstripa(void)
{
	AddLogEntry(TEXT("TabStripA_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveTabstripa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* autoActivateTab, long* autoScrollVelocity)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = *dropTarget;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoActivateTab=%i, autoScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoActivateTab, *autoScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackTabstripa(long effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("TabStripA_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragTabstripa(BOOL pressedEscape, short button, short shift, long* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("TabStripA_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataTabstripa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripA_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripA_OLEReceivedNewData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataTabstripa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripA_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripA_OLESetData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragTabstripa(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TabStripA_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TabStripA_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::OwnerDrawTabTabstripa(LPDISPATCH tsTab, TabStripCtlLibA::OwnerDrawTabStateConstants tabState, long hDC, TabStripCtlLibA::RECTANGLE* drawingRectangle)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_OwnerDrawTab: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_OwnerDrawTab: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", tabState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i)"), tabState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_RClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_RClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_RDblClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_RDblClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowTabstripa(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TabStripA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
	InsertTabsA();
}

void __stdcall CMainDlg::RemovedTabTabstripa(LPDISPATCH tsTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_RemovedTab: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_RemovedTab: tsTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingTabTabstripa(LPDISPATCH tsTab, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_RemovingTab: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_RemovingTab: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowTabstripa()
{
	AddLogEntry(CAtlString(TEXT("TabStripA_ResizedControlWindow")));
}

void __stdcall CMainDlg::TabBeginDragTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_TabBeginDrag: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_TabBeginDrag: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(controls.tabstripA->GetMultiSelect() == VARIANT_FALSE) {
		controls.tabstripA->OLEDrag(NULL, TabStripCtlLibA::odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.tabstripA->CreateTabContainer(_variant_t(tsTab)), -1);
	} else {
		CComPtr< TabStripCtlLibA::ITabStripTabs > pTabs = controls.tabstripA->GetTabs();
		pTabs->PutFilterType(TabStripCtlLibA::fpSelected, TabStripCtlLibA::ftIncluding);
		_variant_t filter;
		filter.Clear();
		CComSafeArray<VARIANT> arr;
		arr.Create(1, 1);
		arr.SetAt(1, _variant_t(true));
		filter.parray = arr.Detach();
		filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: TabStrip expects an array of VARIANTs!
		pTabs->PutFilter(TabStripCtlLibA::fpSelected, filter);

		controls.tabstripA->OLEDrag(NULL, TabStripCtlLibA::odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.tabstripA->CreateTabContainer(_variant_t(pTabs)), -1);
	}
}

void __stdcall CMainDlg::TabBeginRDragTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_TabBeginRDrag: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_TabBeginRDrag: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::TabGetInfoTipTextTabstripa(LPDISPATCH tsTab, long maxInfoTipLength, BSTR* infoTipText)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_TabGetInfoTipText: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_TabGetInfoTipText: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", maxInfoTipLength=%i, infoTipText=%s"), maxInfoTipLength, OLE2W(*infoTipText));
	str += tmp;

	AddLogEntry(str);

	tmp.Format(TEXT("ID: %i, TabData: 0x%X"), pTab->GetID(), pTab->GetTabData());
	*infoTipText = _bstr_t(tmp).Detach();
}

void __stdcall CMainDlg::TabMouseEnterTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_TabMouseEnter: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_TabMouseEnter: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::TabMouseLeaveTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_TabMouseLeave: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_TabMouseLeave: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::TabSelectionChangedTabstripa(LPDISPATCH tsTab)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_TabSelectionChanged: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_TabSelectionChanged: tsTab=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_XClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_XClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickTabstripa(LPDISPATCH tsTab, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<TabStripCtlLibA::ITabStripTab> pTab = tsTab;
	if(pTab) {
		BSTR text = pTab->GetText();
		str += TEXT("TabStripA_XDblClick: tsTab=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("TabStripA_XDblClick: tsTab=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}
