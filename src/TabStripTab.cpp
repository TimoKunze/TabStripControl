// TabStripTab.cpp: A wrapper for existing tabstrip tabs.

#include "stdafx.h"
#include "TabStripTab.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP TabStripTab::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ITabStripTab, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


TabStripTab::Properties::~Properties()
{
	if(pOwnerTabStrip) {
		pOwnerTabStrip->Release();
	}
}

HWND TabStripTab::Properties::GetTabStripHWnd(void)
{
	ATLASSUME(pOwnerTabStrip);

	OLE_HANDLE handle = NULL;
	pOwnerTabStrip->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT TabStripTab::SaveState(int tabIndex, LPTCITEM pTarget, HWND hWndTabStrip/* = NULL*/)
{
	if(!hWndTabStrip) {
		hWndTabStrip = properties.GetTabStripHWnd();
	}
	ATLASSERT(IsWindow(hWndTabStrip));
	if((tabIndex < 0) || (tabIndex >= static_cast<int>(SendMessage(hWndTabStrip, TCM_GETITEMCOUNT, 0, 0)))) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pTarget, TCITEM);
	if(!pTarget) {
		return E_POINTER;
	}

	ZeroMemory(pTarget, sizeof(TCITEM));
	pTarget->cchTextMax = MAX_TABTEXTLENGTH;
	pTarget->pszText = reinterpret_cast<LPTSTR>(malloc((pTarget->cchTextMax + 1) * sizeof(TCHAR)));
	pTarget->dwStateMask = TCIS_BUTTONPRESSED | TCIS_HIGHLIGHTED;
	// TODO: What about TCIF_RTLREADING?
	pTarget->mask = TCIF_IMAGE | TCIF_PARAM | TCIF_STATE | TCIF_TEXT;

	if(SendMessage(hWndTabStrip, TCM_GETITEM, properties.tabIndex, reinterpret_cast<LPARAM>(pTarget))) {
		return S_OK;
	}
	return E_FAIL;
}


void TabStripTab::Attach(int tabIndex)
{
	properties.tabIndex = tabIndex;
}

void TabStripTab::Detach(void)
{
	properties.tabIndex = -1;
}

HRESULT TabStripTab::LoadState(LPTCITEM pSource)
{
	ATLASSERT_POINTER(pSource, TCITEM);
	if(!pSource) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	if(pSource->mask & TCIF_STATE) {
		if(pSource->dwState & TCIS_HIGHLIGHTED) {
			SendMessage(hWndTabStrip, TCM_HIGHLIGHTITEM, properties.tabIndex, MAKELPARAM(TRUE, 0));
		} else {
			SendMessage(hWndTabStrip, TCM_HIGHLIGHTITEM, properties.tabIndex, MAKELPARAM(FALSE, 0));
		}
	}

	VARIANT_BOOL b = VARIANT_FALSE;
	if(pSource->mask & TCIF_STATE) {
		b = BOOL2VARIANTBOOL(pSource->dwState & TCIS_BUTTONPRESSED);
		put_Selected(b);
	}

	LONG l = 0;
	if(pSource->mask & TCIF_IMAGE) {
		l = pSource->iImage;
		put_IconIndex(l);
	}
	l = 0;
	if(pSource->mask & TCIF_PARAM) {
		l = pSource->lParam;
		put_TabData(l);
	}

	if(pSource->mask & TCIF_TEXT) {
		BSTR bstr = _bstr_t(pSource->pszText).Detach();
		put_Text(bstr);
	}
	return S_OK;
}

HRESULT TabStripTab::LoadState(VirtualTabStripTab* pSource)
{
	ATLASSUME(pSource);
	if(!pSource) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	VARIANT_BOOL b = VARIANT_FALSE;
	pSource->get_DropHilited(&b);
	if(b != VARIANT_FALSE) {
		SendMessage(hWndTabStrip, TCM_HIGHLIGHTITEM, properties.tabIndex, MAKELPARAM(TRUE, 0));
	} else {
		SendMessage(hWndTabStrip, TCM_HIGHLIGHTITEM, properties.tabIndex, MAKELPARAM(FALSE, 0));
	}

	b = VARIANT_FALSE;
	pSource->get_Selected(&b);
	put_Selected(b);

	LONG l = 0;
	pSource->get_IconIndex(&l);
	put_IconIndex(l);
	l = 0;
	pSource->get_Index(&l);
	if(properties.tabIndex != l) {
		// TODO: Move the tab?!
	}
	l = 0;
	pSource->get_TabData(&l);
	put_TabData(l);

	BSTR bstr;
	pSource->get_Text(&bstr);
	put_Text(bstr);
	SysFreeString(bstr);
	return S_OK;
}

HRESULT TabStripTab::SaveState(LPTCITEM pTarget, HWND hWndTabStrip/* = NULL*/)
{
	return SaveState(properties.tabIndex, pTarget, hWndTabStrip);
}

HRESULT TabStripTab::SaveState(VirtualTabStripTab* pTarget)
{
	ATLASSUME(pTarget);
	if(!pTarget) {
		return E_POINTER;
	}

	pTarget->SetOwner(properties.pOwnerTabStrip);
	TCITEM tab = {0};
	HRESULT hr = SaveState(&tab);
	pTarget->LoadState(&tab, properties.tabIndex);
	SECUREFREE(tab.pszText);

	return hr;
}

void TabStripTab::SetOwner(TabStrip* pOwner)
{
	if(properties.pOwnerTabStrip) {
		properties.pOwnerTabStrip->Release();
	}
	properties.pOwnerTabStrip = pOwner;
	if(properties.pOwnerTabStrip) {
		properties.pOwnerTabStrip->AddRef();
	}
}


STDMETHODIMP TabStripTab::get_Active(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	int activeTab = SendMessage(hWndTabStrip, TCM_GETCURSEL, 0, 0);
	*pValue = BOOL2VARIANTBOOL(activeTab == properties.tabIndex);
	return S_OK;
}

STDMETHODIMP TabStripTab::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	int caretTab = SendMessage(hWndTabStrip, TCM_GETCURFOCUS, 0, 0);
	*pValue = BOOL2VARIANTBOOL(caretTab == properties.tabIndex);
	return S_OK;
}

STDMETHODIMP TabStripTab::get_DropHilited(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	TCITEM tab = {0};
	tab.dwStateMask = TCIS_HIGHLIGHTED;
	tab.mask = TCIF_STATE;
	if(SendMessage(hWndTabStrip, TCM_GETITEM, properties.tabIndex, reinterpret_cast<LPARAM>(&tab))) {
		*pValue = BOOL2VARIANTBOOL((tab.dwState & TCIS_HIGHLIGHTED) == TCIS_HIGHLIGHTED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::get_HasCloseButton(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.pOwnerTabStrip->IsCloseableTab(properties.tabIndex));
	return S_OK;
}

STDMETHODIMP TabStripTab::put_HasCloseButton(VARIANT_BOOL newValue)
{
	return properties.pOwnerTabStrip->SetCloseableTab(properties.tabIndex, VARIANTBOOL2BOOL(newValue));
}

STDMETHODIMP TabStripTab::get_hAssociatedWindow(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(properties.pOwnerTabStrip->GetAssociatedWindow(properties.tabIndex));
	return S_OK;
}

STDMETHODIMP TabStripTab::put_hAssociatedWindow(OLE_HANDLE newValue)
{
	HWND hAssociatedWindow = static_cast<HWND>(LongToHandle(newValue));
	properties.pOwnerTabStrip->SetAssociatedWindow(properties.tabIndex, hAssociatedWindow);

	if(hAssociatedWindow && IsWindow(hAssociatedWindow)) {
		HWND hWndTabStrip = properties.GetTabStripHWnd();
		ATLASSERT(IsWindow(hWndTabStrip));

		int activeTab = SendMessage(hWndTabStrip, TCM_GETCURSEL, 0, 0);
		if(activeTab == properties.tabIndex) {
			ShowWindow(hAssociatedWindow, SW_SHOW);
		} else {
			ShowWindow(hAssociatedWindow, SW_HIDE);
		}
	}
	return S_OK;
}

STDMETHODIMP TabStripTab::get_Height(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	CRect rc;
	if(SendMessage(hWndTabStrip, TCM_GETITEMRECT, properties.tabIndex, reinterpret_cast<LPARAM>(&rc))) {
		*pValue = rc.Height();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	TCITEM tab = {0};
	tab.mask = TCIF_IMAGE;
	if(SendMessage(hWndTabStrip, TCM_GETITEM, properties.tabIndex, reinterpret_cast<LPARAM>(&tab))) {
		*pValue = tab.iImage;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::put_IconIndex(LONG newValue)
{
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	TCITEM tab = {0};
	tab.iImage = newValue;
	tab.mask = TCIF_IMAGE;
	if(SendMessage(hWndTabStrip, TCM_SETITEM, properties.tabIndex, reinterpret_cast<LPARAM>(&tab))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerTabStrip) {
		if((*pValue = properties.pOwnerTabStrip->TabIndexToID(properties.tabIndex)) != -1) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.tabIndex;
	return S_OK;
}

STDMETHODIMP TabStripTab::put_Index(LONG newValue)
{
	if(newValue == properties.tabIndex) {
		return S_OK;
	}
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	int numberOfTabs = SendMessage(hWndTabStrip, TCM_GETITEMCOUNT, 0, 0);
	if(newValue >= numberOfTabs) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	TCITEM tabData = {0};
	if(SUCCEEDED(SaveState(&tabData, hWndTabStrip))) {
		BOOL isActive = (properties.tabIndex == SendMessage(hWndTabStrip, TCM_GETCURSEL, 0, 0));
		BOOL isCaret = (properties.tabIndex == SendMessage(hWndTabStrip, TCM_GETCURFOCUS, 0, 0));

		// make tab ID management easier
		// NOTE: Should we extend SaveState so that it may save either the ID or the lParam?
		tabData.lParam = properties.pOwnerTabStrip->TabIndexToID(properties.tabIndex);
		tabData.mask |= TCIF_NOINTERCEPTION;

		properties.pOwnerTabStrip->EnterSilentTabInsertionSection();
		BOOL insertBefore = (newValue < properties.tabIndex);
		int newTabIndex = SendMessage(hWndTabStrip, TCM_INSERTITEM, (insertBefore ? newValue : newValue + 1), reinterpret_cast<LPARAM>(&tabData));
		// free the caption string
		SECUREFREE(tabData.pszText);

		// sync our internal tab map
		properties.pOwnerTabStrip->MoveTabIndex(properties.tabIndex, (insertBefore ? newTabIndex : newTabIndex - 1));
		properties.pOwnerTabStrip->EnterSilentActiveTabChangeSection();
		properties.pOwnerTabStrip->EnterSilentCaretChangeSection();
		properties.pOwnerTabStrip->EnterSilentSelectionChangesSection();

		/* TODO: doesn't really work
		// collect all selection states, so that we can restore them later
		#ifdef USE_STL
			std::vector<DWORD> selectionStatesBefore;
		#else
			CAtlArray<DWORD> selectionStatesBefore;
		#endif
		TCITEM tab = {0};
		tab.dwStateMask = TCIS_BUTTONPRESSED;
		tab.mask = TCIF_STATE;
		int numberOfTabs = SendMessage(hWndTabStrip, TCM_GETITEMCOUNT, 0, 0);
		for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
			SendMessage(hWndTabStrip, TCM_GETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab));
			#ifdef USE_STL
				selectionStatesBefore.push_back(tab.dwState & TCIS_BUTTONPRESSED);
			#else
				selectionStatesBefore.Add(tab.dwState & TCIS_BUTTONPRESSED);
			#endif
		}*/

		if(isActive) {
			SendMessage(hWndTabStrip, TCM_SETCURSEL, newTabIndex, 0);
		}

		/* TODO: doesn't really work
		// restore selection states
		for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
			if(!isActive) {
				tab.dwState = selectionStatesBefore[tabIndex];
				SendMessage(hWndTabStrip, TCM_SETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab));
			}
		}*/

		if(isCaret) {
			SendMessage(hWndTabStrip, TCM_SETCURFOCUS, newTabIndex, 0);
		}

		properties.pOwnerTabStrip->LeaveSilentSelectionChangesSection();
		properties.pOwnerTabStrip->LeaveSilentCaretChangeSection();
		properties.pOwnerTabStrip->LeaveSilentActiveTabChangeSection();

		if(properties.pOwnerTabStrip->cachedSettings.dropHilitedTab == (insertBefore ? properties.tabIndex + 1 : properties.tabIndex)) {
			properties.pOwnerTabStrip->cachedSettings.dropHilitedTab = (insertBefore ? newTabIndex : newTabIndex - 1);
		}
		if(properties.pOwnerTabStrip->insertMark.tabIndex == (insertBefore ? properties.tabIndex + 1 : properties.tabIndex)) {
			properties.pOwnerTabStrip->insertMark.tabIndex = (insertBefore ? newTabIndex : newTabIndex - 1);
		}

		// restore flag
		properties.pOwnerTabStrip->LeaveSilentTabInsertionSection();

		// prepare the tab's deletion
		properties.pOwnerTabStrip->EnterSilentTabDeletionSection();
		// finally remove the old tab
		SendMessage(hWndTabStrip, TCM_DELETEITEM, (insertBefore ? properties.tabIndex + 1 : properties.tabIndex), 0);
		properties.pOwnerTabStrip->LeaveSilentTabDeletionSection();
		if(properties.pOwnerTabStrip->cachedSettings.dropHilitedTab > (insertBefore ? properties.tabIndex + 1 : properties.tabIndex)) {
			--(properties.pOwnerTabStrip->cachedSettings.dropHilitedTab);
		}
		if(properties.pOwnerTabStrip->insertMark.tabIndex > (insertBefore ? properties.tabIndex + 1 : properties.tabIndex)) {
			--(properties.pOwnerTabStrip->insertMark.tabIndex);
		}

		properties.tabIndex = (insertBefore ? newTabIndex : newTabIndex - 1);
		return S_OK;
	} else {     // SUCCEEDED(SaveState(&tabData, hWndTabStrip))
		return E_FAIL;
	}
}

STDMETHODIMP TabStripTab::get_Left(OLE_XPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	RECT rc = {0};
	if(SendMessage(hWndTabStrip, TCM_GETITEMRECT, properties.tabIndex, reinterpret_cast<LPARAM>(&rc))) {
		*pValue = rc.left;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	TCITEM tab = {0};
	tab.dwStateMask = TCIS_BUTTONPRESSED;
	tab.mask = TCIF_STATE;
	if(SendMessage(hWndTabStrip, TCM_GETITEM, properties.tabIndex, reinterpret_cast<LPARAM>(&tab))) {
		*pValue = BOOL2VARIANTBOOL((tab.dwState & TCIS_BUTTONPRESSED) == TCIS_BUTTONPRESSED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::put_Selected(VARIANT_BOOL newValue)
{
	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	if(CWindow(hWndTabStrip).GetStyle() & TCS_MULTISELECT) {
		TCITEM tab = {0};
		if(newValue != VARIANT_FALSE) {
			tab.dwState = TCIS_BUTTONPRESSED;
		}
		tab.dwStateMask = TCIS_BUTTONPRESSED;
		tab.mask = TCIF_STATE;
		if(SendMessage(hWndTabStrip, TCM_SETITEM, properties.tabIndex, reinterpret_cast<LPARAM>(&tab))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::get_TabData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	TCITEM tab = {0};
	tab.mask = TCIF_PARAM;
	if(SendMessage(hWndTabStrip, TCM_GETITEM, properties.tabIndex, reinterpret_cast<LPARAM>(&tab))) {
		*pValue = tab.lParam;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::put_TabData(LONG newValue)
{
	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	TCITEM tab = {0};
	tab.lParam = newValue;
	tab.mask = TCIF_PARAM;
	if(SendMessage(hWndTabStrip, TCM_SETITEM, properties.tabIndex, reinterpret_cast<LPARAM>(&tab))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	TCITEM tab = {0};
	tab.mask = TCIF_TEXT;
	tab.cchTextMax = MAX_TABTEXTLENGTH;
	tab.pszText = reinterpret_cast<LPTSTR>(malloc((tab.cchTextMax + 1) * sizeof(TCHAR)));
	if(SendMessage(hWndTabStrip, TCM_GETITEM, properties.tabIndex, reinterpret_cast<LPARAM>(&tab))) {
		*pValue = _bstr_t(tab.pszText).Detach();
		SECUREFREE(tab.pszText);
		return S_OK;
	}
	SECUREFREE(tab.pszText);
	return E_FAIL;
}

STDMETHODIMP TabStripTab::put_Text(BSTR newValue)
{
	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	TCITEM tab = {0};
	tab.mask = TCIF_TEXT;
	COLE2T converter(newValue);
	tab.pszText = converter;
	tab.cchTextMax = lstrlen(tab.pszText);
	if(SendMessage(hWndTabStrip, TCM_SETITEM, properties.tabIndex, reinterpret_cast<LPARAM>(&tab))) {
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP TabStripTab::get_Top(OLE_YPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	RECT rc = {0};
	if(SendMessage(hWndTabStrip, TCM_GETITEMRECT, properties.tabIndex, reinterpret_cast<LPARAM>(&rc))) {
		*pValue = rc.top;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStripTab::get_Width(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	CRect rc;
	if(SendMessage(hWndTabStrip, TCM_GETITEMRECT, properties.tabIndex, reinterpret_cast<LPARAM>(&rc))) {
		*pValue = rc.Width();
		return S_OK;
	}
	return E_FAIL;
}


STDMETHODIMP TabStripTab::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(!phImageList) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerTabStrip);

	// ask the control for a drag image
	POINT upperLeftPoint = {0};
	*phImageList = HandleToLong(properties.pOwnerTabStrip->CreateLegacyDragImage(properties.tabIndex, &upperLeftPoint, NULL));

	if(*phImageList) {
		if(pXUpperLeft) {
			*pXUpperLeft = upperLeftPoint.x;
		}
		if(pYUpperLeft) {
			*pYUpperLeft = upperLeftPoint.y;
		}
		return S_OK;
	}
	return E_FAIL;
}