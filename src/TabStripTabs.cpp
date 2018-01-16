// TabStripTabs.cpp: Manages a collection of TabStripTab objects

#include "stdafx.h"
#include "TabStripTabs.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP TabStripTabs::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ITabStripTabs, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP TabStripTabs::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<TabStripTabs>* pTabStripTabsObj = NULL;
	CComObject<TabStripTabs>::CreateInstance(&pTabStripTabsObj);
	pTabStripTabsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pTabStripTabsObj->properties);

	pTabStripTabsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pTabStripTabsObj->Release();
	return S_OK;
}

STDMETHODIMP TabStripTabs::Next(ULONG numberOfMaxTabs, VARIANT* pTabs, ULONG* pNumberOfTabsReturned)
{
	ATLASSERT_POINTER(pTabs, VARIANT);
	if(!pTabs) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	// check each item in the tabstrip
	int numberOfTabs = SendMessage(hWndTabStrip, TCM_GETITEMCOUNT, 0, 0);
	ULONG i = 0;
	for(i = 0; i < numberOfMaxTabs; ++i) {
		VariantInit(&pTabs[i]);

		do {
			if(properties.lastEnumeratedTab >= 0) {
				if(properties.lastEnumeratedTab < numberOfTabs) {
					properties.lastEnumeratedTab = GetNextTabToProcess(properties.lastEnumeratedTab, numberOfTabs);
				}
			} else {
				properties.lastEnumeratedTab = GetFirstTabToProcess(numberOfTabs);
			}
			if(properties.lastEnumeratedTab >= numberOfTabs) {
				properties.lastEnumeratedTab = -1;
			}
		} while((properties.lastEnumeratedTab != -1) && (!IsPartOfCollection(properties.lastEnumeratedTab)));

		if(properties.lastEnumeratedTab != -1) {
			ClassFactory::InitTabStripTab(properties.lastEnumeratedTab, properties.pOwnerTabStrip, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pTabs[i].pdispVal));
			pTabs[i].vt = VT_DISPATCH;
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfTabsReturned) {
		*pNumberOfTabsReturned = i;
	}

	return (i == numberOfMaxTabs ? S_OK : S_FALSE);
}

STDMETHODIMP TabStripTabs::Reset(void)
{
	properties.lastEnumeratedTab = -1;
	return S_OK;
}

STDMETHODIMP TabStripTabs::Skip(ULONG numberOfTabsToSkip)
{
	VARIANT dummy;
	ULONG numTabsReturned = 0;
	// we could skip all tabs at once, but it's easier to skip them one after the other
	for(ULONG i = 1; i <= numberOfTabsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numTabsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numTabsReturned == 0) {
			// there're no more tabs to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


BOOL TabStripTabs::IsValidBooleanFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BOOL);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL TabStripTabs::IsValidIntegerFilter(VARIANT& filter, int minValue, int maxValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT)) && element.intVal >= minValue && element.intVal <= maxValue;
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL TabStripTabs::IsValidIntegerFilter(VARIANT& filter, int minValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT)) && element.intVal >= minValue;
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL TabStripTabs::IsValidIntegerFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_UI4));
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL TabStripTabs::IsValidStringFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BSTR);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

int TabStripTabs::GetFirstTabToProcess(int numberOfTabs)
{
	if(numberOfTabs == 0) {
		return -1;
	}
	return 0;
}

int TabStripTabs::GetNextTabToProcess(int previousTab, int numberOfTabs)
{
	if(previousTab < numberOfTabs - 1) {
		return previousTab + 1;
	} else {
		return -1;
	}
}

BOOL TabStripTabs::IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(VARIANT_BOOL, VARIANT_BOOL);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.boolVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(element.boolVal == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL TabStripTabs::IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		int v = 0;
		if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT))) {
			v = element.intVal;
		}
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(LONG, LONG);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, v)) {
				foundMatch = TRUE;
			}
		} else {
			if(v == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL TabStripTabs::IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(BSTR, BSTR);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.bstrVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(properties.caseSensitiveFilters) {
				if(lstrcmpW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			} else {
				if(lstrcmpiW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL TabStripTabs::IsPartOfCollection(int tabIndex, HWND hWndTabStrip/* = NULL*/)
{
	if(!hWndTabStrip) {
		hWndTabStrip = properties.GetTabStripHWnd();
	}
	ATLASSERT(IsWindow(hWndTabStrip));

	if(!IsValidTabIndex(tabIndex, hWndTabStrip)) {
		return FALSE;
	}

	BOOL isPart = FALSE;
	// we declare this one here to avoid compiler warnings
	TCITEM tab = {0};

	if(properties.filterType[fpIndex] != ftDeactivated) {
		if(IsIntegerInSafeArray(properties.filter[fpIndex].parray, tabIndex, properties.comparisonFunction[fpIndex])) {
			if(properties.filterType[fpIndex] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpIndex] == ftIncluding) {
				goto Exit;
			}
		}
	}

	BOOL mustRetrieveTabData = FALSE;
	if(properties.filterType[fpIconIndex] != ftDeactivated) {
		tab.mask |= TCIF_IMAGE;
		mustRetrieveTabData = TRUE;
	}
	if(properties.filterType[fpTabData] != ftDeactivated) {
		tab.mask |= TCIF_PARAM;
		mustRetrieveTabData = TRUE;
	}
	if(properties.filterType[fpSelected] != ftDeactivated) {
		tab.mask |= TCIF_STATE;
		tab.dwStateMask |= TCIS_BUTTONPRESSED;
		mustRetrieveTabData = TRUE;
	}
	if(properties.filterType[fpText] != ftDeactivated) {
		tab.mask |= TCIF_TEXT;
		tab.cchTextMax = MAX_TABTEXTLENGTH;
		tab.pszText = reinterpret_cast<LPTSTR>(malloc((tab.cchTextMax + 1) * sizeof(TCHAR)));
		mustRetrieveTabData = TRUE;
	}

	if(mustRetrieveTabData) {
		if(!SendMessage(hWndTabStrip, TCM_GETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab))) {
			goto Exit;
		}

		// apply the filters

		if(properties.filterType[fpSelected] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpSelected].parray, BOOL2VARIANTBOOL((tab.dwState & TCIS_BUTTONPRESSED) == TCIS_BUTTONPRESSED), properties.comparisonFunction[fpSelected])) {
				if(properties.filterType[fpSelected] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpSelected] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpTabData] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpTabData].parray, static_cast<int>(tab.lParam), properties.comparisonFunction[fpTabData])) {
				if(properties.filterType[fpTabData] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpTabData] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpIconIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpIconIndex].parray, tab.iImage, properties.comparisonFunction[fpIconIndex])) {
				if(properties.filterType[fpIconIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpIconIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpText] != ftDeactivated) {
			if(IsStringInSafeArray(properties.filter[fpText].parray, CComBSTR(tab.pszText), properties.comparisonFunction[fpText])) {
				if(properties.filterType[fpText] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpText] == ftIncluding) {
					goto Exit;
				}
			}
		}
	}
	isPart = TRUE;

Exit:
	if(tab.pszText) {
		SECUREFREE(tab.pszText);
	}
	return isPart;
}

void TabStripTabs::OptimizeFilter(FilteredPropertyConstants filteredProperty)
{
	if(filteredProperty != fpSelected) {
		// currently we optimize boolean filters only
		return;
	}

	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(properties.filter[filteredProperty].parray, 1, &lBound);
	SafeArrayGetUBound(properties.filter[filteredProperty].parray, 1, &uBound);

	// now that we have the bounds, iterate the array
	VARIANT element;
	UINT numberOfTrues = 0;
	UINT numberOfFalses = 0;
	for(LONG i = lBound; i <= uBound; ++i) {
		SafeArrayGetElement(properties.filter[filteredProperty].parray, &i, &element);
		if(element.boolVal == VARIANT_FALSE) {
			++numberOfFalses;
		} else {
			++numberOfTrues;
		}

		VariantClear(&element);
	}

	if(numberOfTrues > 0 && numberOfFalses > 0) {
		// we've something like True Or False Or True - we can deactivate this filter
		properties.filterType[filteredProperty] = ftDeactivated;
		VariantClear(&properties.filter[filteredProperty]);
	} else if(numberOfTrues == 0 && numberOfFalses > 1) {
		// False Or False Or False... is still False, so we need just one False
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_FALSE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	} else if(numberOfFalses == 0 && numberOfTrues > 1) {
		// True Or True Or True... is still True, so we need just one True
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_TRUE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	}
}

#ifdef USE_STL
	HRESULT TabStripTabs::RemoveTabs(std::list<int>& tabsToRemove, HWND hWndTabStrip)
#else
	HRESULT TabStripTabs::RemoveTabs(CAtlList<int>& tabsToRemove, HWND hWndTabStrip)
#endif
{
	ATLASSERT(IsWindow(hWndTabStrip));

	//CWindowEx(hWndTabStrip).InternalSetRedraw(FALSE);
	// sort in reverse order
	#ifdef USE_STL
		tabsToRemove.sort(std::greater<int>());
		for(std::list<int>::const_iterator iter = tabsToRemove.begin(); iter != tabsToRemove.end(); ++iter) {
			SendMessage(hWndTabStrip, TCM_DELETEITEM, *iter, 0);
		}
	#else
		// perform a crude bubble sort
		for(size_t j = 0; j < tabsToRemove.GetCount(); ++j) {
			for(size_t i = 0; i < tabsToRemove.GetCount() - 1; ++i) {
				if(tabsToRemove.GetAt(tabsToRemove.FindIndex(i)) < tabsToRemove.GetAt(tabsToRemove.FindIndex(i + 1))) {
					tabsToRemove.SwapElements(tabsToRemove.FindIndex(i), tabsToRemove.FindIndex(i + 1));
				}
			}
		}

		for(size_t i = 0; i < tabsToRemove.GetCount(); ++i) {
			SendMessage(hWndTabStrip, TCM_DELETEITEM, tabsToRemove.GetAt(tabsToRemove.FindIndex(i)), 0);
		}
	#endif
	//CWindowEx(hWndTabStrip).InternalSetRedraw(TRUE);

	return S_OK;
}


TabStripTabs::Properties::~Properties()
{
	for(int i = 0; i < NUMBEROFFILTERS; ++i) {
		VariantClear(&filter[i]);
	}
	if(pOwnerTabStrip) {
		pOwnerTabStrip->Release();
	}
}

void TabStripTabs::Properties::CopyTo(TabStripTabs::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerTabStrip = this->pOwnerTabStrip;
		if(pTarget->pOwnerTabStrip) {
			pTarget->pOwnerTabStrip->AddRef();
		}
		pTarget->lastEnumeratedTab = this->lastEnumeratedTab;
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;

		for(int i = 0; i < NUMBEROFFILTERS; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}

HWND TabStripTabs::Properties::GetTabStripHWnd(void)
{
	ATLASSUME(pOwnerTabStrip);

	OLE_HANDLE handle = NULL;
	pOwnerTabStrip->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void TabStripTabs::SetOwner(TabStrip* pOwner)
{
	if(properties.pOwnerTabStrip) {
		properties.pOwnerTabStrip->Release();
	}
	properties.pOwnerTabStrip = pOwner;
	if(properties.pOwnerTabStrip) {
		properties.pOwnerTabStrip->AddRef();
	}
}


STDMETHODIMP TabStripTabs::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP TabStripTabs::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP TabStripTabs::get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndex:
		case fpTabData:
		case fpSelected:
		case fpText:
			*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TabStripTabs::put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue)
{
	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndex:
		case fpTabData:
		case fpSelected:
		case fpText:
			properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TabStripTabs::get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndex:
		case fpTabData:
		case fpSelected:
		case fpText:
			VariantClear(pValue);
			VariantCopy(pValue, &properties.filter[filteredProperty]);
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TabStripTabs::put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue)
{
	// check 'newValue'
	switch(filteredProperty) {
		case fpIconIndex:
			if(!IsValidIntegerFilter(newValue, -1)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpIndex:
			if(!IsValidIntegerFilter(newValue, 0)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpSelected:
			if(!IsValidBooleanFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpTabData:
			if(!IsValidIntegerFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpText:
			if(!IsValidStringFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		default:
			return E_INVALIDARG;
			break;
	}

	VariantClear(&properties.filter[filteredProperty]);
	VariantCopy(&properties.filter[filteredProperty], &newValue);
	OptimizeFilter(filteredProperty);
	return S_OK;
}

STDMETHODIMP TabStripTabs::get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FilterTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndex:
		case fpTabData:
		case fpSelected:
		case fpText:
			*pValue = properties.filterType[filteredProperty];
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TabStripTabs::put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
{
	if((newValue < 0) || (newValue > 2)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndex:
		case fpTabData:
		case fpSelected:
		case fpText:
			properties.filterType[filteredProperty] = newValue;
			if(newValue != ftDeactivated) {
				properties.usingFilters = TRUE;
			} else {
				properties.usingFilters = FALSE;
				for(int i = 0; i < NUMBEROFFILTERS; ++i) {
					if(properties.filterType[i] != ftDeactivated) {
						properties.usingFilters = TRUE;
						break;
					}
				}
			}
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TabStripTabs::get_Item(LONG tabIdentifier, TabIdentifierTypeConstants tabIdentifierType/* = titIndex*/, ITabStripTab** ppTab/* = NULL*/)
{
	ATLASSERT_POINTER(ppTab, ITabStripTab*);
	if(!ppTab) {
		return E_POINTER;
	}

	// retrieve the tab's index
	int tabIndex = -1;
	switch(tabIdentifierType) {
		case titID:
			if(properties.pOwnerTabStrip) {
				tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabIdentifier);
			}
			break;
		case titIndex:
			tabIndex = tabIdentifier;
			break;
	}

	if(tabIndex != -1) {
		if(IsPartOfCollection(tabIndex)) {
			ClassFactory::InitTabStripTab(tabIndex, properties.pOwnerTabStrip, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(ppTab));
			if(*ppTab) {
				return S_OK;
			}
		}
	}

	// tab not found
	if(tabIdentifierType == titIndex) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP TabStripTabs::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP TabStripTabs::Add(BSTR tabText, LONG insertAt/* = -1*/, LONG iconIndex/* = -1*/, LONG tabData/* = 0*/, ITabStripTab** ppAddedTab/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedTab, ITabStripTab*);
	if(!ppAddedTab) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	if(insertAt == -1) {
		insertAt = 0x7FFFFFFF;
	}
	ATLASSERT(insertAt >= 0);
	if(insertAt < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	TCITEM insertionData = {0};
	insertionData.iImage = iconIndex;
	insertionData.lParam = tabData;
	insertionData.mask = TCIF_IMAGE | TCIF_PARAM | TCIF_TEXT;

	COLE2T converter(tabText);
	insertionData.pszText = converter;
	insertionData.cchTextMax = lstrlen(insertionData.pszText);

	// finally insert the tab
	*ppAddedTab = NULL;
	int tabIndex = SendMessage(hWndTabStrip, TCM_INSERTITEM, insertAt, reinterpret_cast<LPARAM>(&insertionData));
	if(tabIndex != -1) {
		ClassFactory::InitTabStripTab(tabIndex, properties.pOwnerTabStrip, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(ppAddedTab), FALSE);
		hr = S_OK;
	}

	return hr;
}

STDMETHODIMP TabStripTabs::Contains(LONG tabIdentifier, TabIdentifierTypeConstants tabIdentifierType/* = titIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	// retrieve the tab's index
	int tabIndex = -1;
	switch(tabIdentifierType) {
		case titID:
			if(properties.pOwnerTabStrip) {
				tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabIdentifier);
			}
			break;
		case titIndex:
			tabIndex = tabIdentifier;
			break;
	}

	*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(tabIndex));
	return S_OK;
}

STDMETHODIMP TabStripTabs::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	if(!properties.usingFilters) {
		*pValue = SendMessage(hWndTabStrip, TCM_GETITEMCOUNT, 0, 0);
		return S_OK;
	}

	// count the tabs "by hand"
	*pValue = 0;
	int numberOfTabs = SendMessage(hWndTabStrip, TCM_GETITEMCOUNT, 0, 0);
	int tabIndex = GetFirstTabToProcess(numberOfTabs);
	while(tabIndex != -1) {
		if(IsPartOfCollection(tabIndex, hWndTabStrip)) {
			++(*pValue);
		}
		tabIndex = GetNextTabToProcess(tabIndex, numberOfTabs);
	}
	return S_OK;
}

STDMETHODIMP TabStripTabs::DeselectAll(VARIANT_BOOL keepActiveTabSelected/* = VARIANT_FALSE*/)
{
	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	if(!properties.usingFilters) {
		SendMessage(hWndTabStrip, TCM_DESELECTALL, VARIANTBOOL2BOOL(keepActiveTabSelected), 0);
		return S_OK;
	}

	// find the tabs to deselect "by hand"
	TCITEM tab = {0};
	tab.dwStateMask = TCIS_BUTTONPRESSED;
	tab.mask = TCIF_STATE;
	int numberOfTabs = SendMessage(hWndTabStrip, TCM_GETITEMCOUNT, 0, 0);
	int tabIndex = GetFirstTabToProcess(numberOfTabs);
	while(tabIndex != -1) {
		if(IsPartOfCollection(tabIndex, hWndTabStrip)) {
			if(!SendMessage(hWndTabStrip, TCM_SETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab))) {
				return E_FAIL;
			}
		}
		tabIndex = GetNextTabToProcess(tabIndex, numberOfTabs);
	}
	return S_OK;
}

STDMETHODIMP TabStripTabs::Remove(LONG tabIdentifier, TabIdentifierTypeConstants tabIdentifierType/* = titIndex*/)
{
	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	// retrieve the tab's index
	int tabIndex = -1;
	switch(tabIdentifierType) {
		case titID:
			if(properties.pOwnerTabStrip) {
				tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabIdentifier);
			}
			break;
		case titIndex:
			tabIndex = tabIdentifier;
			break;
	}

	if(tabIndex != -1) {
		if(IsPartOfCollection(tabIndex)) {
			if(SendMessage(hWndTabStrip, TCM_DELETEITEM, tabIndex, 0)) {
				return S_OK;
			}
		}
	} else {
		// tab not found
		if(tabIdentifierType == titIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}

	return E_FAIL;
}

STDMETHODIMP TabStripTabs::RemoveAll(void)
{
	HWND hWndTabStrip = properties.GetTabStripHWnd();
	ATLASSERT(IsWindow(hWndTabStrip));

	if(!properties.usingFilters) {
		if(SendMessage(hWndTabStrip, TCM_DELETEALLITEMS, 0, 0)) {
			return S_OK;
		} else {
			return E_FAIL;
		}
	}

	// find the tabs to remove "by hand"
	#ifdef USE_STL
		std::list<int> tabsToRemove;
	#else
		CAtlList<int> tabsToRemove;
	#endif
	int numberOfTabs = SendMessage(hWndTabStrip, TCM_GETITEMCOUNT, 0, 0);
	int tabIndex = GetFirstTabToProcess(numberOfTabs);
	while(tabIndex != -1) {
		if(IsPartOfCollection(tabIndex, hWndTabStrip)) {
			#ifdef USE_STL
				tabsToRemove.push_back(tabIndex);
			#else
				tabsToRemove.AddTail(tabIndex);
			#endif
		}
		tabIndex = GetNextTabToProcess(tabIndex, numberOfTabs);
	}
	return RemoveTabs(tabsToRemove, hWndTabStrip);
}