// CommonProperties.cpp: The control's "Common" property page

#include "stdafx.h"
#include "CommonProperties.h"


CommonProperties::CommonProperties()
{
	m_dwTitleID = IDS_TITLECOMMONPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGCOMMONPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP CommonProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<CommonProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.disabledEventsList.SubclassWindow(GetDlgItem(IDC_DISABLEDEVENTSBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.disabledEventsList.GetImageList(LVSIL_STATE));
	controls.disabledEventsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.disabledEventsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.disabledEventsList.AddColumn(TEXT(""), 0);
	controls.disabledEventsList.GetToolTips().SetTitle(TTI_INFO, TEXT("Affected events"));

	controls.tabBoundingBoxList.SubclassWindow(GetDlgItem(IDC_TABBOUNDINGBOX));
	hStateImageList = SetupStateImageList(controls.tabBoundingBoxList.GetImageList(LVSIL_STATE));
	controls.tabBoundingBoxList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.tabBoundingBoxList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER, LVS_EX_DOUBLEBUFFER | LVS_EX_FULLROWSELECT);
	controls.tabBoundingBoxList.AddColumn(TEXT(""), 0);

	// setup the toolbar
	CRect toolbarRect;
	GetClientRect(&toolbarRect);
	toolbarRect.OffsetRect(0, 2);
	toolbarRect.left += toolbarRect.right - 46;
	toolbarRect.bottom = toolbarRect.top + 22;
	controls.toolbar.Create(*this, toolbarRect, NULL, WS_CHILDWINDOW | WS_VISIBLE | TBSTYLE_TRANSPARENT | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE, 0);
	controls.toolbar.SetButtonStructSize();
	controls.imagelistEnabled.CreateFromImage(IDB_TOOLBARENABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetImageList(controls.imagelistEnabled);
	controls.imagelistDisabled.CreateFromImage(IDB_TOOLBARDISABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetDisabledImageList(controls.imagelistDisabled);

	// insert the buttons
	TBBUTTON buttons[2];
	ZeroMemory(buttons, sizeof(buttons));
	buttons[0].iBitmap = 0;
	buttons[0].idCommand = ID_LOADSETTINGS;
	buttons[0].fsState = TBSTATE_ENABLED;
	buttons[0].fsStyle = BTNS_BUTTON;
	buttons[1].iBitmap = 1;
	buttons[1].idCommand = ID_SAVESETTINGS;
	buttons[1].fsStyle = BTNS_BUTTON;
	buttons[1].fsState = TBSTATE_ENABLED;
	controls.toolbar.AddButtons(2, buttons);

	LoadSettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<CommonProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT CommonProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(pNotificationDetails->hwndFrom != controls.disabledEventsList) {
		return 0;
	}

	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	switch(pDetails->iItem) {
		case 0:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MClick, MouseEnter, MouseHover, MouseLeave, TabMouseEnter, TabMouseLeave, MouseMove"))));
			break;
		case 1:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Click, DblClick, MClick, MDblClick, RClick, RDblClick, XClick, XDblClick"))));
			break;
		case 2:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("KeyDown, KeyUp, KeyPress"))));
			break;
		case 3:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("InsertingTab, InsertedTab"))));
			break;
		case 4:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("RemovingTab, RemovedTab"))));
			break;
		case 5:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("FreeTabData"))));
			break;
		case 6:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("TabSelectionChanged"))));
			break;
	}
	ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));

	delete[] pBuffer;
	return 0;
}

LRESULT CommonProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	if(pDetails->uChanged & LVIF_STATE) {
		if((pDetails->uNewState & LVIS_STATEIMAGEMASK) != (pDetails->uOldState & LVIS_STATEIMAGEMASK)) {
			if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 3) {
				if(pNotificationDetails->hwndFrom != properties.hWndCheckMarksAreSetFor) {
					LVITEM item = {0};
					item.state = INDEXTOSTATEIMAGEMASK(1);
					item.stateMask = LVIS_STATEIMAGEMASK;
					::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem, reinterpret_cast<LPARAM>(&item));
				}
			}
			SetDirty(TRUE);
		}
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	CComQIPtr<ITabStrip, &IID_ITabStrip> pControl(m_ppUnk[0]);
	if(pControl) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			pControl->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
	}
	return 0;
}

LRESULT CommonProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	CComQIPtr<ITabStrip, &IID_ITabStrip> pControl(m_ppUnk[0]);
	if(pControl) {
		CFileDialog dlg(FALSE, NULL, TEXT("TabStrip Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			pControl->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
	}
	return 0;
}


void CommonProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		CComQIPtr<ITabStrip, &IID_ITabStrip> pControl(m_ppUnk[object]);
		if(pControl) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			pControl->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, deTabInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(4, LVIS_STATEIMAGEMASK), l, deTabDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(5, LVIS_STATEIMAGEMASK), l, deFreeTabData);
			SetBit(controls.disabledEventsList.GetItemState(6, LVIS_STATEIMAGEMASK), l, deTabSelectionChanged);
			pControl->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			TabBoundingBoxDefinitionConstants boundingBox = static_cast<TabBoundingBoxDefinitionConstants>(0);
			pControl->get_TabBoundingBoxDefinition(&boundingBox);
			l = static_cast<LONG>(boundingBox);
			SetBit(controls.tabBoundingBoxList.GetItemState(0, LVIS_STATEIMAGEMASK), l, tbbdTabIcon);
			SetBit(controls.tabBoundingBoxList.GetItemState(1, LVIS_STATEIMAGEMASK), l, tbbdTabLabel);
			SetBit(controls.tabBoundingBoxList.GetItemState(2, LVIS_STATEIMAGEMASK), l, tbbdTabCloseButton);
			pControl->put_TabBoundingBoxDefinition(static_cast<TabBoundingBoxDefinitionConstants>(l));
		}
	}

	SetDirty(FALSE);
}

void CommonProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	DisabledEventsConstants* pDisabledEvents = static_cast<DisabledEventsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(DisabledEventsConstants)));
	if(pDisabledEvents) {
		TabBoundingBoxDefinitionConstants* pBoundingBox = static_cast<TabBoundingBoxDefinitionConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(TabBoundingBoxDefinitionConstants)));
		if(pBoundingBox) {
			for(UINT object = 0; object < m_nObjects; ++object) {
				CComQIPtr<ITabStrip, &IID_ITabStrip> pControl(m_ppUnk[object]);
				if(pControl) {
					pControl->get_DisabledEvents(&pDisabledEvents[object]);
					pControl->get_TabBoundingBoxDefinition(&pBoundingBox[object]);
				}
			}

			// fill the listboxes
			LONG* pl = reinterpret_cast<LONG*>(pDisabledEvents);
			properties.hWndCheckMarksAreSetFor = controls.disabledEventsList;
			controls.disabledEventsList.DeleteAllItems();
			controls.disabledEventsList.AddItem(0, 0, TEXT("Mouse events"));
			controls.disabledEventsList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, deMouseEvents), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.AddItem(1, 0, TEXT("Click events"));
			controls.disabledEventsList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, deClickEvents), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.AddItem(2, 0, TEXT("Keyboard events"));
			controls.disabledEventsList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, deKeyboardEvents), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.AddItem(3, 0, TEXT("Tab insertions"));
			controls.disabledEventsList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, deTabInsertionEvents), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.AddItem(4, 0, TEXT("Tab deletions"));
			controls.disabledEventsList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, deTabDeletionEvents), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.AddItem(5, 0, TEXT("FreeTabData event"));
			controls.disabledEventsList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, deFreeTabData), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.AddItem(6, 0, TEXT("TabSelectionChanged event"));
			controls.disabledEventsList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, deTabSelectionChanged), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

			pl = reinterpret_cast<LONG*>(pBoundingBox);
			properties.hWndCheckMarksAreSetFor = controls.tabBoundingBoxList;
			controls.tabBoundingBoxList.DeleteAllItems();
			controls.tabBoundingBoxList.AddItem(0, 0, TEXT("Icon"));
			controls.tabBoundingBoxList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, tbbdTabIcon), LVIS_STATEIMAGEMASK);
			controls.tabBoundingBoxList.AddItem(1, 0, TEXT("Text label"));
			controls.tabBoundingBoxList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, tbbdTabLabel), LVIS_STATEIMAGEMASK);
			controls.tabBoundingBoxList.AddItem(2, 0, TEXT("Close button"));
			controls.tabBoundingBoxList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, tbbdTabCloseButton), LVIS_STATEIMAGEMASK);
			controls.tabBoundingBoxList.SetColumnWidth(0, LVSCW_AUTOSIZE);

			properties.hWndCheckMarksAreSetFor = NULL;

			HeapFree(GetProcessHeap(), 0, pBoundingBox);
		}
		HeapFree(GetProcessHeap(), 0, pDisabledEvents);
	}

	SetDirty(FALSE);
}

int CommonProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
{
	int stateImageIndex = 1;
	for(UINT object = 0; object < arraysize; ++object) {
		if(pArray[object] & bitsToCheckFor) {
			if(stateImageIndex == 1) {
				stateImageIndex = (object == 0 ? 2 : 3);
			}
		} else {
			if(stateImageIndex == 2) {
				stateImageIndex = (object == 0 ? 1 : 3);
			}
		}
	}

	return INDEXTOSTATEIMAGEMASK(stateImageIndex);
}

void CommonProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
{
	stateImageMask >>= 12;
	switch(stateImageMask) {
		case 1:
			value &= ~bitToSet;
			break;
		case 2:
			value |= bitToSet;
			break;
	}
}