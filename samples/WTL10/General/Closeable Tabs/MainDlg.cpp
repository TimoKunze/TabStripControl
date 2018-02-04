// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include <olectl.h>
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
		DispEventUnadvise(controls.tabstripU);
		controls.tabstripU.Release();
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

	tabStripUWnd.SubclassWindow(GetDlgItem(IDC_TABSTRIPU));
	tabStripUWnd.QueryControl(__uuidof(ITabStrip), reinterpret_cast<LPVOID*>(&controls.tabstripU));
	if(controls.tabstripU) {
		DispEventAdvise(controls.tabstripU);
		InsertTabs();
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.tabstripU) {
		controls.tabstripU->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::InsertTabs(void)
{
	controls.tabstripU->PuthImageList(HandleToLong(controls.imageList.m_hImageList));

	CComPtr<ITabStripTabs> pTabs = controls.tabstripU->GetTabs();
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

void __stdcall CMainDlg::ClickTabstripu(LPDISPATCH tsTab, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails)
{
	CComQIPtr<ITabStripTab> pTab = tsTab;
	if(pTab) {
		if(hitTestDetails & htCloseButton) {
			if(pTab->GetActive() != VARIANT_FALSE) {
				long newActiveTab = pTab->GetIndex() + 1;
				if(newActiveTab == controls.tabstripU->GetTabs()->Count()) {
					newActiveTab -= 2;
				}
				if(newActiveTab != -1) {
					controls.tabstripU->PutRefActiveTab(controls.tabstripU->GetTabs()->GetItem(newActiveTab, titIndex));
				}
			}

			// close the tab
			controls.tabstripU->GetTabs()->Remove(pTab->GetIndex(), titIndex);
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowTabstripu(long /*hWnd*/)
{
	InsertTabs();
}

void __stdcall CMainDlg::TabGetInfoTipTextTabstripu(LPDISPATCH /*tsTab*/, long /*maxInfoTipLength*/, BSTR* infoTipText)
{
	HitTestConstants hitTestDetails;
	POINT pt;
	GetCursorPos(&pt);
	::ScreenToClient(static_cast<HWND>(LongToHandle(controls.tabstripU->GethWnd())), &pt);

	controls.tabstripU->HitTest(pt.x, pt.y, &hitTestDetails);
	if(hitTestDetails & htCloseButton) {
		*infoTipText = _bstr_t(TEXT("Close this tab")).Detach();
	}
}
