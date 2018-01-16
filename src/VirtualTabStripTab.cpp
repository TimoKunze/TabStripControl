// VirtualTabStripTab.cpp: A wrapper for non-existing tabstrip tabs.

#include "stdafx.h"
#include "VirtualTabStripTab.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP VirtualTabStripTab::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualTabStripTab, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualTabStripTab::Properties::~Properties()
{
	if(settings.pszText != LPSTR_TEXTCALLBACK) {
		SECUREFREE(settings.pszText);
	}
	if(pOwnerTabStrip) {
		pOwnerTabStrip->Release();
	}
}

HWND VirtualTabStripTab::Properties::GetTabStripHWnd(void)
{
	ATLASSUME(pOwnerTabStrip);

	OLE_HANDLE handle = NULL;
	pOwnerTabStrip->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT VirtualTabStripTab::LoadState(LPTCITEM pSource, int tabIndex)
{
	ATLASSERT_POINTER(pSource, TCITEM);

	SECUREFREE(properties.settings.pszText);
	properties.settings = *pSource;
	if(properties.settings.mask & TCIF_TEXT) {
		// duplicate the tab's text
		if(properties.settings.pszText != LPSTR_TEXTCALLBACK) {
			properties.settings.cchTextMax = lstrlen(pSource->pszText);
			properties.settings.pszText = reinterpret_cast<LPTSTR>(malloc((properties.settings.cchTextMax + 1) * sizeof(TCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopy(properties.settings.pszText, properties.settings.cchTextMax + 1, pSource->pszText)));
		}
	}
	properties.tabIndex = tabIndex;

	return S_OK;
}

HRESULT VirtualTabStripTab::SaveState(LPTCITEM pTarget, int& tabIndex)
{
	ATLASSERT_POINTER(pTarget, TCITEM);

	SECUREFREE(pTarget->pszText);
	*pTarget = properties.settings;
	if(pTarget->mask & TCIF_TEXT) {
		// duplicate the tab's text
		if(pTarget->pszText != LPSTR_TEXTCALLBACK) {
			pTarget->pszText = reinterpret_cast<LPTSTR>(malloc((pTarget->cchTextMax + 1) * sizeof(TCHAR)));
			ATLASSERT(pTarget->pszText);
			if(pTarget->pszText) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pTarget->pszText, pTarget->cchTextMax + 1, properties.settings.pszText)));
			}
		}
	}
	tabIndex = properties.tabIndex;

	return S_OK;
}

void VirtualTabStripTab::SetOwner(TabStrip* pOwner)
{
	if(properties.pOwnerTabStrip) {
		properties.pOwnerTabStrip->Release();
	}
	properties.pOwnerTabStrip = pOwner;
	if(properties.pOwnerTabStrip) {
		properties.pOwnerTabStrip->AddRef();
	}
}


STDMETHODIMP VirtualTabStripTab::get_DropHilited(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & TCIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.dwState & TCIS_HIGHLIGHTED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualTabStripTab::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & TCIF_IMAGE) {
		*pValue = properties.settings.iImage;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualTabStripTab::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.tabIndex;
	return S_OK;
}

STDMETHODIMP VirtualTabStripTab::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & TCIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.dwState & TCIS_BUTTONPRESSED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualTabStripTab::get_TabData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & TCIF_PARAM) {
		*pValue = static_cast<LONG>(properties.settings.lParam);
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualTabStripTab::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & TCIF_TEXT) {
		*pValue = _bstr_t(properties.settings.pszText).Detach();
	} else {
		*pValue = SysAllocString(L"");
	}

	return S_OK;
}