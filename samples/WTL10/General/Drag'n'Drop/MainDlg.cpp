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

	controls.oleDDCheck = GetDlgItem(IDC_OLEDDCHECK);
	controls.oleDDCheck.SetCheck(BST_CHECKED);
	tabStripUWnd.SubclassWindow(GetDlgItem(IDC_TABSTRIPU));
	tabStripUWnd.QueryControl(__uuidof(ITabStrip), reinterpret_cast<LPVOID*>(&controls.tabstripU));
	if(controls.tabstripU) {
		DispEventAdvise(controls.tabstripU);
		InsertTabs();

		UpdateMenu();
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

LRESULT CMainDlg::OnOptionsAeroDragImages(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(GetMenuState(hMenu, ID_OPTIONS_AERODRAGIMAGES, MF_BYCOMMAND) & MF_CHECKED) {
		controls.tabstripU->PutOLEDragImageStyle(odistClassic);
	} else {
		controls.tabstripU->PutOLEDragImageStyle(odistAeroIfAvailable);
	}
	UpdateMenu();
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
}

void CMainDlg::UpdateMenu(void)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(controls.tabstripU->GetOLEDragImageStyle() == odistAeroIfAvailable) {
		CheckMenuItem(hMenu, ID_OPTIONS_AERODRAGIMAGES, MF_BYCOMMAND | MF_CHECKED);
	} else {
		CheckMenuItem(hMenu, ID_OPTIONS_AERODRAGIMAGES, MF_BYCOMMAND | MF_UNCHECKED);
	}
}

void __stdcall CMainDlg::AbortedDragTabstripu()
{
	//controls.tabstripU->PutRefDropHilitedTab(NULL);
	controls.tabstripU->SetInsertMarkPosition(impNowhere, NULL);
	lastDropTarget = -1;
}

void __stdcall CMainDlg::DragMouseMoveTabstripu(LPDISPATCH* /*dropTarget*/, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/, VARIANT_BOOL* autoActivateTab, long* /*autoScrollVelocity*/)
{
	*autoActivateTab = VARIANT_FALSE;

	CComPtr<ITabStripTab> pTab = NULL;
	InsertMarkPositionConstants newInsertMarkRelativePosition = impNowhere;
	controls.tabstripU->GetClosestInsertMarkPosition(x, y, &newInsertMarkRelativePosition, &pTab);
	LONG newDropTarget = -1;
	if(pTab) {
		newDropTarget = pTab->GetIndex();
	}
	if(newDropTarget != lastDropTarget) {
		//controls.tabstripU->putref_DropHilitedTab(&pTab);
    lastDropTarget = newDropTarget;
	}
	controls.tabstripU->SetInsertMarkPosition(newInsertMarkRelativePosition, pTab);
}

void __stdcall CMainDlg::DropTabstripu(LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/)
{
	if(bRightDrag) {
		HMENU hMenu = LoadMenu(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_DRAGMENU));
		HMENU hPopupMenu = GetSubMenu(hMenu, 0);

		POINT pt = {x, y};
		::ClientToScreen(static_cast<HWND>(LongToHandle(controls.tabstripU->GethWnd())), &pt);
		int selectedCmd = TrackPopupMenuEx(hPopupMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, static_cast<HWND>(LongToHandle(controls.tabstripU->GethWnd())), NULL);
		DestroyMenu(hPopupMenu);
		DestroyMenu(hMenu);

		switch(selectedCmd) {
			case ID_ACTIONS_CANCEL:
				AbortedDragTabstripu();
				return;
				break;
			case ID_ACTIONS_COPYTAB:
				MessageBox(TEXT("TODO: Copy the tab."), TEXT("Command not implemented"), MB_ICONINFORMATION);
				AbortedDragTabstripu();
				return;
				break;
			case ID_ACTIONS_MOVETAB:
				// fall through
				break;
		}
	}

	// move the tab
	InsertMarkPositionConstants insertionMark;
	CComPtr<ITabStripTab> pTab = NULL;
	controls.tabstripU->GetInsertMarkPosition(&insertionMark, &pTab);

	LONG insertAt = -1;
	switch(insertionMark) {
		case impAfter:
			insertAt = pTab->GetIndex() + 1;
			break;
		case impBefore:
			insertAt = pTab->GetIndex();
			break;
		case impNowhere:
			insertAt = -1;
			break;
	}

	if(insertAt != -1) {
		CComQIPtr<IEnumVARIANT> pEnum = controls.tabstripU->GetDraggedTabs();
		pEnum->Reset();     // NOTE: It's important to call Reset()!
		ULONG ul = 0;
		_variant_t v;
		v.Clear();
		while(pEnum->Next(1, &v, &ul) == S_OK) {
			CComQIPtr<ITabStripTab> pTab2 = v.pdispVal;
			if(pTab2->GetIndex() < insertAt) {
				pTab2->PutIndex(insertAt - 1);
			} else {
				pTab2->PutIndex(insertAt);
				insertAt++;
			}
		}
	}

	//controls.tabstripU->PutRefDropHilitedTab(NULL);
	controls.tabstripU->SetInsertMarkPosition(impNowhere, NULL);
	lastDropTarget = -1;
}

void __stdcall CMainDlg::KeyDownTabstripu(short* keyCode, short /*shift*/)
{
	if(*keyCode == VK_ESCAPE) {
		if(controls.tabstripU->GetDraggedTabs()) {
			controls.tabstripU->EndDrag(VARIANT_TRUE);
		}
	}
}

void __stdcall CMainDlg::MouseUpTabstripu(LPDISPATCH /*tsTab*/, short button, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/)
{
	if(controls.oleDDCheck.GetCheck() == BST_UNCHECKED) {
		if(((button == 1) && !bRightDrag) || ((button == 2) && bRightDrag)) {
			if(controls.tabstripU->GetDraggedTabs()) {
				// Are we within the (extended) client area?
				if(lastDropTarget != -1) {
					// yes
					controls.tabstripU->EndDrag(VARIANT_FALSE);
				} else {
					// no
					controls.tabstripU->EndDrag(VARIANT_TRUE);
				}
			}
		}
	}
}

void __stdcall CMainDlg::OLECompleteDragTabstripu(LPDISPATCH data, long performedEffect)
{
	if(performedEffect == odeMove) {
		// remove the dragged tabs
		controls.tabstripU->GetDraggedTabs()->RemoveAll(VARIANT_TRUE);
	} else {
		CComQIPtr<IDataObject> pData = data;
		if(pData) {
			FORMATETC format = {CF_TARGETCLSID, NULL, 1, -1, TYMED_HGLOBAL};
			if(pData->QueryGetData(&format) == S_OK) {
				STGMEDIUM storageMedium = {0};
				if(pData->GetData(&format, &storageMedium) == S_OK) {
					GUID* pCLSIDOfTarget = reinterpret_cast<GUID*>(GlobalLock(storageMedium.hGlobal));
					if(*pCLSIDOfTarget == CLSID_RecycleBin) {
						// remove the dragged tabs
						controls.tabstripU->GetDraggedTabs()->RemoveAll(VARIANT_TRUE);
					}
					GlobalUnlock(storageMedium.hGlobal);
					ReleaseStgMedium(&storageMedium);
				}
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragDropTabstripu(LPDISPATCH data, long* effect, LPDISPATCH /*dropTarget*/, short /*button*/, short shift, long /*x*/, long /*y*/, long /*hitTestDetails*/)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		CAtlString str;
		if(pData->GetFormat(CF_UNICODETEXT, -1, 1) != VARIANT_FALSE) {
			str = pData->GetData(CF_UNICODETEXT, -1, 1).bstrVal;
		} else if(pData->GetFormat(CF_TEXT, -1, 1) != VARIANT_FALSE) {
			str = pData->GetData(CF_TEXT, -1, 1).bstrVal;
		} else if(pData->GetFormat(CF_OEMTEXT, -1, 1) != VARIANT_FALSE) {
			str = pData->GetData(CF_OEMTEXT, -1, 1).bstrVal;
		}
		str.Replace(TEXT("\r\n"), TEXT(""));

		if(str != TEXT("")) {
			// insert a new tab
			InsertMarkPositionConstants insertionMark;
			CComPtr<ITabStripTab> pTab = NULL;
			controls.tabstripU->GetInsertMarkPosition(&insertionMark, &pTab);

			LONG insertAt = -1;
			switch(insertionMark) {
				case impAfter:
					insertAt = pTab->GetIndex() + 1;
					break;
				case impBefore:
					insertAt = pTab->GetIndex();
					break;
				case impNowhere:
					insertAt = -1;
					break;
			}
			controls.tabstripU->PutRefCaretTab(controls.tabstripU->GetTabs()->Add(_bstr_t(str).Detach(), insertAt, 1, 0));

		} else if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			// insert a new tab for each file/folder
			InsertMarkPositionConstants insertionMark;
			CComPtr<ITabStripTab> pTab = NULL;
			controls.tabstripU->GetInsertMarkPosition(&insertionMark, &pTab);

			LONG insertAt = -1;
			switch(insertionMark) {
				case impAfter:
					insertAt = pTab->GetIndex() + 1;
					break;
				case impBefore:
					insertAt = pTab->GetIndex();
					break;
				case impNowhere:
					insertAt = -1;
					break;
			}

			_variant_t v = pData->GetData(CF_HDROP, -1, 1);
			CComSafeArray<BSTR> arr;
			arr.Attach(v.parray);
			for(LONG i = arr.GetLowerBound(); i <= arr.GetUpperBound(); i++) {
				pTab = controls.tabstripU->GetTabs()->Add(_bstr_t(arr.GetAt(i)), insertAt, 1, 0);
				controls.tabstripU->PutRefCaretTab(pTab);
				insertAt = pTab->GetIndex() + 1;
			}
			arr.Detach();
		}
	}

	//controls.tabstripU->PutRefDropHilitedTab(NULL);
	controls.tabstripU->SetInsertMarkPosition(impNowhere, NULL);
	lastDropTarget = -1;

	if(!controls.tabstripU->GetDraggedTabs()) {
		*effect = odeCopy;
	} else {
		*effect = odeMove;
	}
	if(shift & 1/*vbShiftMask*/) {
		*effect = odeMove;
	}
	if(shift & 2/*vbCtrlMask*/) {
		*effect = odeCopy;
	}
	if(shift & 4/*vbAltMask*/) {
		*effect = odeLink;
	}
}

void __stdcall CMainDlg::OLEDragEnterTabstripu(LPDISPATCH data, long* effect, LPDISPATCH* /*dropTarget*/, short /*button*/, short shift, long x, long y, long /*hitTestDetails*/, VARIANT_BOOL* autoActivateTab, long* /*autoScrollVelocity*/)
{
	*autoActivateTab = VARIANT_FALSE;
	CComQIPtr<IOLEDataObject> pData = data;

	_bstr_t dropActionDescription;
	DropDescriptionIconConstants dropDescriptionIcon = ddiNoDrop;
	_bstr_t dropTargetDescription;

	CComPtr<ITabStripTab> pTab = NULL;
	InsertMarkPositionConstants newInsertMarkRelativePosition = impNowhere;
	controls.tabstripU->GetClosestInsertMarkPosition(x, y, &newInsertMarkRelativePosition, &pTab);
	LONG newDropTarget = -1;
	if(pTab) {
    dropTargetDescription = pTab->GetText();
		newDropTarget = pTab->GetIndex();
	}
	if(newDropTarget != lastDropTarget) {
		//controls.tabstripU->putref_DropHilitedTab(&pTab);
		lastDropTarget = newDropTarget;
	}
	controls.tabstripU->SetInsertMarkPosition(newInsertMarkRelativePosition, pTab);

	if(!controls.tabstripU->GetDraggedTabs()) {
		*effect = odeCopy;
	} else {
		*effect = odeMove;
	}
	if(shift & 1/*vbShiftMask*/) {
		*effect = odeMove;
	}
	if(shift & 2/*vbCtrlMask*/) {
		*effect = odeCopy;
	}
	if(shift & 4/*vbAltMask*/) {
		*effect = odeLink;
	}

	if(!pTab) {
		dropDescriptionIcon = ddiNoDrop;
		dropActionDescription = TEXT("Cannot place here");
	} else {
		switch(*effect) {
			case odeMove:
				switch(newInsertMarkRelativePosition) {
					case impAfter:
						dropActionDescription = TEXT("Move after %1");
						break;
					case impBefore:
						dropActionDescription = TEXT("Move before %1");
						break;
					default:
						dropActionDescription = TEXT("Move to %1");
						break;
				}
				dropDescriptionIcon = ddiMove;
				break;
			case odeLink:
				switch(newInsertMarkRelativePosition) {
					case impAfter:
						dropActionDescription = TEXT("Create link after %1");
						break;
					case impBefore:
						dropActionDescription = TEXT("Create link before %1");
						break;
					default:
						dropActionDescription = TEXT("Create link in %1");
						break;
				}
				dropDescriptionIcon = ddiLink;
				break;
			default:
				switch(newInsertMarkRelativePosition) {
					case impAfter:
						dropActionDescription = TEXT("Insert copy after %1");
						break;
					case impBefore:
						dropActionDescription = TEXT("Insert copy before %1");
						break;
					default:
						dropActionDescription = TEXT("Copy to %1");
						break;
				}
				dropDescriptionIcon = ddiCopy;
				break;
		}
	}
	pData->SetDropDescription(dropTargetDescription, dropActionDescription, dropDescriptionIcon);
}

void __stdcall CMainDlg::OLEDragLeaveTabstripu(LPDISPATCH /*data*/, LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/)
{
	//controls.tabstripU->PutRefDropHilitedTab(NULL);
	controls.tabstripU->SetInsertMarkPosition(impNowhere, NULL);
	lastDropTarget = -1;
}

void __stdcall CMainDlg::OLEDragMouseMoveTabstripu(LPDISPATCH data, long* effect, LPDISPATCH* /*dropTarget*/, short /*button*/, short shift, long x, long y, long /*hitTestDetails*/, VARIANT_BOOL* autoActivateTab, long* /*autoScrollVelocity*/)
{
	*autoActivateTab = VARIANT_FALSE;
	CComQIPtr<IOLEDataObject> pData = data;

	_bstr_t dropActionDescription;
	DropDescriptionIconConstants dropDescriptionIcon = ddiNoDrop;
	_bstr_t dropTargetDescription;

	CComPtr<ITabStripTab> pTab = NULL;
	InsertMarkPositionConstants newInsertMarkRelativePosition = impNowhere;
	controls.tabstripU->GetClosestInsertMarkPosition(x, y, &newInsertMarkRelativePosition, &pTab);
	LONG newDropTarget = -1;
	if(pTab) {
    dropTargetDescription = pTab->GetText();
		newDropTarget = pTab->GetIndex();
	}
	if(newDropTarget != lastDropTarget) {
		//controls.tabstripU->putref_DropHilitedTab(&pTab);
		lastDropTarget = newDropTarget;
	}
	controls.tabstripU->SetInsertMarkPosition(newInsertMarkRelativePosition, pTab);

	if(!controls.tabstripU->GetDraggedTabs()) {
		*effect = odeCopy;
	} else {
		*effect = odeMove;
	}
	if(shift & 1/*vbShiftMask*/) {
		*effect = odeMove;
	}
	if(shift & 2/*vbCtrlMask*/) {
		*effect = odeCopy;
	}
	if(shift & 4/*vbAltMask*/) {
		*effect = odeLink;
	}

	if(!pTab) {
		dropDescriptionIcon = ddiNoDrop;
		dropActionDescription = TEXT("Cannot place here");
	} else {
		switch(*effect) {
			case odeMove:
				switch(newInsertMarkRelativePosition) {
					case impAfter:
						dropActionDescription = TEXT("Move after %1");
						break;
					case impBefore:
						dropActionDescription = TEXT("Move before %1");
						break;
					default:
						dropActionDescription = TEXT("Move to %1");
						break;
				}
				dropDescriptionIcon = ddiMove;
				break;
			case odeLink:
				switch(newInsertMarkRelativePosition) {
					case impAfter:
						dropActionDescription = TEXT("Create link after %1");
						break;
					case impBefore:
						dropActionDescription = TEXT("Create link before %1");
						break;
					default:
						dropActionDescription = TEXT("Create link in %1");
						break;
				}
				dropDescriptionIcon = ddiLink;
			default:
				switch(newInsertMarkRelativePosition) {
					case impAfter:
						dropActionDescription = TEXT("Insert copy after %1");
						break;
					case impBefore:
						dropActionDescription = TEXT("Insert copy before %1");
						break;
					default:
						dropActionDescription = TEXT("Copy to %1");
						break;
				}
				dropDescriptionIcon = ddiCopy;
		}
	}
	pData->SetDropDescription(dropTargetDescription, dropActionDescription, dropDescriptionIcon);
}

void __stdcall CMainDlg::OLESetDataTabstripu(LPDISPATCH data, long formatID, long /*index*/, long /*dataOrViewAspect*/)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		switch(formatID) {
			case CF_TEXT:
			case CF_OEMTEXT:
			case CF_UNICODETEXT: {
				CComQIPtr<IEnumVARIANT> pEnum = controls.tabstripU->GetDraggedTabs();
				pEnum->Reset();     // NOTE: It's important to call Reset()!
				ULONG ul = 0;
				_variant_t v;
				v.Clear();
				CComBSTR str = L"";
				while(pEnum->Next(1, &v, &ul) == S_OK) {
					CComQIPtr<ITabStripTab> pTab = v.pdispVal;
					str.AppendBSTR(pTab->GetText());
					str.Append(TEXT("\r\n"));
				}
				pData->SetData(formatID, _variant_t(str), -1, 1);
				break;
			}
		}
	}
}

void __stdcall CMainDlg::OLEStartDragTabstripu(LPDISPATCH data)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		_variant_t invalidVariant;
		invalidVariant.vt = VT_ERROR;

		pData->SetData(CF_TEXT, invalidVariant, -1, 1);
		pData->SetData(CF_OEMTEXT, invalidVariant, -1, 1);
		pData->SetData(CF_UNICODETEXT, invalidVariant, -1, 1);
		pData->SetData(CF_HDROP, invalidVariant, -1, 1);
	}
}

void __stdcall CMainDlg::RecreatedControlWindowTabstripu(long /*hWnd*/)
{
	InsertTabs();
}

void __stdcall CMainDlg::TabBeginDragTabstripu(LPDISPATCH tsTab, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/)
{
	bRightDrag = FALSE;
	lastDropTarget = -1;
	if(controls.tabstripU->GetMultiSelect() == VARIANT_FALSE) {
		if(controls.oleDDCheck.GetCheck() == BST_CHECKED) {
			IOLEDataObject* p = NULL;
			controls.tabstripU->OLEDrag(reinterpret_cast<long*>(&p), odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.tabstripU->CreateTabContainer(_variant_t(tsTab)), -1);
		} else {
			controls.tabstripU->BeginDrag(controls.tabstripU->CreateTabContainer(_variant_t(tsTab)), static_cast<OLE_HANDLE>(-1), NULL, NULL);
		}
	} else {
		CComPtr<ITabStripTabs> pTabs = controls.tabstripU->GetTabs();
		pTabs->PutFilterType(fpSelected, ftIncluding);
		_variant_t filter;
		filter.Clear();
		CComSafeArray<VARIANT> arr;
		arr.Create(1, 1);
		arr.SetAt(1, _variant_t(true));
		filter.parray = arr.Detach();
		filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: TabStrip expects an array of VARIANTs!
		pTabs->PutFilter(fpSelected, filter);

		if(controls.oleDDCheck.GetCheck() == BST_CHECKED) {
			IDataObject* p = NULL;
			controls.tabstripU->OLEDrag(reinterpret_cast<long*>(&p), odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.tabstripU->CreateTabContainer(_variant_t(pTabs)), -1);
		} else {
			controls.tabstripU->BeginDrag(controls.tabstripU->CreateTabContainer(_variant_t(pTabs)), static_cast<OLE_HANDLE>(-1), NULL, NULL);
		}
	}
}

void __stdcall CMainDlg::TabBeginRDragTabstripu(LPDISPATCH tsTab, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/)
{
	bRightDrag = TRUE;
	lastDropTarget = -1;
	if(controls.tabstripU->GetMultiSelect() == VARIANT_FALSE) {
		if(controls.oleDDCheck.GetCheck() == BST_CHECKED) {
			IOLEDataObject* p = NULL;
			controls.tabstripU->OLEDrag(reinterpret_cast<long*>(&p), odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.tabstripU->CreateTabContainer(_variant_t(tsTab)), -1);
		} else {
			controls.tabstripU->BeginDrag(controls.tabstripU->CreateTabContainer(_variant_t(tsTab)), static_cast<OLE_HANDLE>(-1), NULL, NULL);
		}
	} else {
		CComPtr<ITabStripTabs> pTabs = controls.tabstripU->GetTabs();
		pTabs->PutFilterType(fpSelected, ftIncluding);
		_variant_t filter;
		filter.Clear();
		CComSafeArray< VARIANT > arr;
		arr.Create(1, 1);
		arr.SetAt(1, _variant_t(true));
		filter.parray = arr.Detach();
		filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: TabStrip expects an array of VARIANTs!
		pTabs->PutFilter(fpSelected, filter);

		if(controls.oleDDCheck.GetCheck() == BST_CHECKED) {
			IDataObject* p = NULL;
			controls.tabstripU->OLEDrag(reinterpret_cast<long*>(&p), odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.tabstripU->CreateTabContainer(_variant_t(pTabs)), -1);
		} else {
			controls.tabstripU->BeginDrag(controls.tabstripU->CreateTabContainer(_variant_t(pTabs)), static_cast<OLE_HANDLE>(-1), NULL, NULL);
		}
	}
}
