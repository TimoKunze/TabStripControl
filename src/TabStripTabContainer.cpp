// TabStripTabContainer.cpp: Manages a collection of TabStripTab objects

#include "stdafx.h"
#include "TabStripTabContainer.h"
#include "ClassFactory.h"


DWORD TabStripTabContainer::nextID = 0;


TabStripTabContainer::TabStripTabContainer()
{
	containerID = ++nextID;
}

TabStripTabContainer::~TabStripTabContainer()
{
	properties.pOwnerTabStrip->DeregisterTabContainer(containerID);
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP TabStripTabContainer::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ITabStripTabContainer, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP TabStripTabContainer::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<TabStripTabContainer>* pTSTabContainerObj = NULL;
	CComObject<TabStripTabContainer>::CreateInstance(&pTSTabContainerObj);
	pTSTabContainerObj->AddRef();

	// clone all settings
	pTSTabContainerObj->properties = properties;
	if(pTSTabContainerObj->properties.pOwnerTabStrip) {
		pTSTabContainerObj->properties.pOwnerTabStrip->AddRef();
	}
	#ifdef USE_STL
		pTSTabContainerObj->properties.tabs.resize(properties.tabs.size());
		std::copy(properties.tabs.begin(), properties.tabs.end(), pTSTabContainerObj->properties.tabs.begin());
	#else
		//pTSTabContainerObj->properties.tabs.Copy(properties.tabs);
		pTSTabContainerObj->tabs.Copy(tabs);
	#endif

	pTSTabContainerObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pTSTabContainerObj->Release();

	if(*ppEnumerator) {
		properties.pOwnerTabStrip->RegisterTabContainer(static_cast<ITabContainer*>(pTSTabContainerObj));
	}
	return S_OK;
}

STDMETHODIMP TabStripTabContainer::Next(ULONG numberOfMaxTabs, VARIANT* pTabs, ULONG* pNumberOfTabsReturned)
{
	ATLASSERT_POINTER(pTabs, VARIANT);
	if(!pTabs) {
		return E_POINTER;
	}

	ULONG i = 0;
	for(i = 0; i < numberOfMaxTabs; ++i) {
		VariantInit(&pTabs[i]);

		#ifdef USE_STL
			if(properties.nextEnumeratedTab >= static_cast<int>(properties.tabs.size())) {
				properties.nextEnumeratedTab = 0;
				// there's nothing more to iterate
				break;
			}
			int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(properties.tabs[properties.nextEnumeratedTab]);
		#else
			//if(properties.nextEnumeratedTab >= static_cast<int>(properties.tabs.GetCount())) {
			if(properties.nextEnumeratedTab >= static_cast<int>(tabs.GetCount())) {
				properties.nextEnumeratedTab = 0;
				// there's nothing more to iterate
				break;
			}
			//int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(properties.tabs[properties.nextEnumeratedTab]);
			int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabs[properties.nextEnumeratedTab]);
		#endif

		ClassFactory::InitTabStripTab(tabIndex, properties.pOwnerTabStrip, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pTabs[i].pdispVal));
		pTabs[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedTab;
	}
	if(pNumberOfTabsReturned) {
		*pNumberOfTabsReturned = i;
	}

	return (i == numberOfMaxTabs ? S_OK : S_FALSE);
}

STDMETHODIMP TabStripTabContainer::Reset(void)
{
	properties.nextEnumeratedTab = 0;
	return S_OK;
}

STDMETHODIMP TabStripTabContainer::Skip(ULONG numberOfTabsToSkip)
{
	properties.nextEnumeratedTab += numberOfTabsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ITabContainer
void TabStripTabContainer::RemovedTab(LONG tabID)
{
	if(tabID == -1) {
		RemoveAll();
	} else {
		Remove(tabID);
	}
}

DWORD TabStripTabContainer::GetID(void)
{
	return containerID;
}
// implementation of ITabContainer
//////////////////////////////////////////////////////////////////////


TabStripTabContainer::Properties::~Properties()
{
	if(pOwnerTabStrip) {
		pOwnerTabStrip->Release();
	}
}

HWND TabStripTabContainer::Properties::GetTabStripHWnd(void)
{
	ATLASSUME(pOwnerTabStrip);

	OLE_HANDLE handle = NULL;
	pOwnerTabStrip->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void TabStripTabContainer::SetOwner(TabStrip* pOwner)
{
	if(properties.pOwnerTabStrip) {
		properties.pOwnerTabStrip->Release();
	}
	properties.pOwnerTabStrip = pOwner;
	if(properties.pOwnerTabStrip) {
		properties.pOwnerTabStrip->AddRef();
	}
}


STDMETHODIMP TabStripTabContainer::get_Item(LONG tabID, ITabStripTab** ppTab)
{
	ATLASSERT_POINTER(ppTab, ITabStripTab*);
	if(!ppTab) {
		return E_POINTER;
	}

	#ifdef USE_STL
		std::vector<LONG>::iterator iter = std::find(properties.tabs.begin(), properties.tabs.end(), tabID);
		if(iter != properties.tabs.end()) {
			int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabID);
			if(tabIndex != -1) {
				ClassFactory::InitTabStripTab(tabIndex, properties.pOwnerTabStrip, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(ppTab));
				return S_OK;
			}
		}
	#else
		//for(size_t i = 0; i < properties.tabs.GetCount(); ++i) {
		for(size_t i = 0; i < tabs.GetCount(); ++i) {
			//if(properties.tabs[i] == tabID) {
			if(tabs[i] == tabID) {
				int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabID);
				if(tabIndex != -1) {
					ClassFactory::InitTabStripTab(tabIndex, properties.pOwnerTabStrip, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(ppTab));
					return S_OK;
				}
				break;
			}
		}
	#endif

	return E_INVALIDARG;
}

STDMETHODIMP TabStripTabContainer::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP TabStripTabContainer::Add(VARIANT tabs)
{
	HRESULT hr = E_FAIL;
	LONG id = -1;
	switch(tabs.vt) {
		case VT_DISPATCH:
			if(tabs.pdispVal) {
				CComQIPtr<ITabStripTab, &IID_ITabStripTab> pTSTab = tabs.pdispVal;
				if(pTSTab) {
					// add a single TabStripTab object
					hr = pTSTab->get_ID(&id);
				} else {
					CComQIPtr<ITabStripTabs, &IID_ITabStripTabs> pTSTabs(tabs.pdispVal);
					if(pTSTabs) {
						// add a TabStripTabs collection
						CComQIPtr<IEnumVARIANT, &IID_IEnumVARIANT> pEnumerator(pTSTabs);
						if(pEnumerator) {
							hr = S_OK;
							VARIANT tab;
							VariantInit(&tab);
							ULONG dummy = 0;
							while(pEnumerator->Next(1, &tab, &dummy) == S_OK) {
								if((tab.vt == VT_DISPATCH) && tab.pdispVal) {
									pTSTab = tab.pdispVal;
									if(pTSTab) {
										id = -1;
										pTSTab->get_ID(&id);
										#ifdef USE_STL
											std::vector<LONG>::iterator iter = std::find(properties.tabs.begin(), properties.tabs.end(), id);
											if(iter == properties.tabs.end()) {
												properties.tabs.push_back(id);
											}
										#else
											BOOL alreadyThere = FALSE;
											//for(size_t i = 0; i < properties.tabs.GetCount(); ++i) {
											for(size_t i = 0; i < this->tabs.GetCount(); ++i) {
												//if(properties.tabs[i] == id) {
												if(this->tabs[i] == id) {
													alreadyThere = TRUE;
													break;
												}
											}
											if(!alreadyThere) {
												//properties.tabs.Add(id);
												this->tabs.Add(id);
											}
										#endif
									}
								}
								VariantClear(&tab);
							}
							return S_OK;
						}
					}
				}
			}
			break;
		default:
			VARIANT v;
			VariantInit(&v);
			hr = VariantChangeType(&v, &tabs, 0, VT_UI4);
			id = v.ulVal;
			break;
	}
	if(FAILED(hr)) {
		// invalid arg - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	#ifdef USE_STL
		std::vector<LONG>::iterator iter = std::find(properties.tabs.begin(), properties.tabs.end(), id);
		if(iter == properties.tabs.end()) {
			properties.tabs.push_back(id);
		}
	#else
		BOOL alreadyThere = FALSE;
		//for(size_t i = 0; i < properties.tabs.GetCount(); ++i) {
		for(size_t i = 0; i < this->tabs.GetCount(); ++i) {
			//if(properties.tabs[i] == id) {
			if(this->tabs[i] == id) {
				alreadyThere = TRUE;
				break;
			}
		}
		if(!alreadyThere) {
			//properties.tabs.Add(id);
			this->tabs.Add(id);
		}
	#endif
	return S_OK;
}

STDMETHODIMP TabStripTabContainer::Clone(ITabStripTabContainer** ppClone)
{
	ATLASSERT_POINTER(ppClone, ITabStripTabContainer*);
	if(!ppClone) {
		return E_POINTER;
	}

	*ppClone = NULL;
	CComObject<TabStripTabContainer>* pTSTabContainerObj = NULL;
	CComObject<TabStripTabContainer>::CreateInstance(&pTSTabContainerObj);
	pTSTabContainerObj->AddRef();

	// clone all settings
	pTSTabContainerObj->properties = properties;
	if(pTSTabContainerObj->properties.pOwnerTabStrip) {
		pTSTabContainerObj->properties.pOwnerTabStrip->AddRef();
	}
	#ifdef USE_STL
		pTSTabContainerObj->properties.tabs.resize(properties.tabs.size());
		std::copy(properties.tabs.begin(), properties.tabs.end(), pTSTabContainerObj->properties.tabs.begin());
	#else
		//pTSTabContainerObj->properties.tabs.Copy(properties.tabs);
		pTSTabContainerObj->tabs.Copy(tabs);
	#endif

	pTSTabContainerObj->QueryInterface(__uuidof(ITabStripTabContainer), reinterpret_cast<LPVOID*>(ppClone));
	pTSTabContainerObj->Release();

	if(*ppClone) {
		properties.pOwnerTabStrip->RegisterTabContainer(static_cast<ITabContainer*>(pTSTabContainerObj));
	}
	return S_OK;
}

STDMETHODIMP TabStripTabContainer::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef USE_STL
		*pValue = static_cast<LONG>(properties.tabs.size());
	#else
		//*pValue = static_cast<LONG>(properties.tabs.GetCount());
		*pValue = static_cast<LONG>(tabs.GetCount());
	#endif
	return S_OK;
}

STDMETHODIMP TabStripTabContainer::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(!phImageList) {
		return E_POINTER;
	}

	*phImageList = NULL;
	#ifdef USE_STL
		switch(properties.tabs.size()) {
	#else
		//switch(properties.tabs.GetCount()) {
		switch(tabs.GetCount()) {
	#endif
		case 0:
			return S_OK;
			break;
		case 1: {
			ATLASSUME(properties.pOwnerTabStrip);
			//int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(properties.tabs[0]);
			#ifdef USE_STL
				int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(properties.tabs[0]);
			#else
				int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabs[0]);
			#endif
			if(tabIndex != -1) {
				POINT upperLeftPoint = {0};
				*phImageList = HandleToLong(properties.pOwnerTabStrip->CreateLegacyDragImage(tabIndex, &upperLeftPoint, NULL));
				if(*phImageList) {
					if(pXUpperLeft) {
						*pXUpperLeft = upperLeftPoint.x;
					}
					if(pYUpperLeft) {
						*pYUpperLeft = upperLeftPoint.y;
					}
					return S_OK;
				}
			}
			break;
		}
		default: {
			// create a large drag image out of small drag images
			ATLASSUME(properties.pOwnerTabStrip);

			BOOL use32BPPImage = RunTimeHelper::IsCommCtrl6();

			// calculate the bitmap's required size and collect each tab's imagelist
			#ifdef USE_STL
				std::vector<HIMAGELIST> imageLists;
				std::vector<RECT> tabBoundingRects;
			#else
				CAtlArray<HIMAGELIST> imageLists;
				CAtlArray<RECT> tabBoundingRects;
			#endif
			POINT upperLeftPoint = {0};
			CRect boundingRect;
			#ifdef USE_STL
				for(std::vector<LONG>::iterator iter = properties.tabs.begin(); iter != properties.tabs.end(); ++iter) {
					int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(*iter);
			#else
				//for(size_t i = 0; i < properties.tabs.GetCount(); ++i) {
				for(size_t i = 0; i < tabs.GetCount(); ++i) {
					//int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(properties.tabs[i]);
					int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabs[i]);
			#endif
				if(tabIndex != -1) {
					// NOTE: Windows skips tabs outside the client area to improve performance. We don't.
					POINT pt = {0};
					RECT tabBoundingRect = {0};
					HIMAGELIST hImageList = properties.pOwnerTabStrip->CreateLegacyDragImage(tabIndex, &pt, &tabBoundingRect);
					boundingRect.UnionRect(&boundingRect, &tabBoundingRect);

					#ifdef USE_STL
						if(imageLists.size() == 0) {
					#else
						if(imageLists.GetCount() == 0) {
					#endif
						upperLeftPoint = pt;
					} else {
						upperLeftPoint.x = min(upperLeftPoint.x, pt.x);
						upperLeftPoint.y = min(upperLeftPoint.y, pt.y);
					}
					#ifdef USE_STL
						imageLists.push_back(hImageList);
						tabBoundingRects.push_back(tabBoundingRect);
					#else
						imageLists.Add(hImageList);
						tabBoundingRects.Add(tabBoundingRect);
					#endif
				}
			}
			CRect dragImageRect(0, 0, boundingRect.Width(), boundingRect.Height());

			// setup the DCs we'll draw into
			HDC hCompatibleDC = GetDC(HWND_DESKTOP);
			CDC memoryDC;
			memoryDC.CreateCompatibleDC(hCompatibleDC);
			CDC maskMemoryDC;
			maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

			// create the bitmap and its mask
			CBitmap dragImage;
			LPRGBQUAD pDragImageBits = NULL;
			if(use32BPPImage) {
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = dragImageRect.Width();
				bitmapInfo.bmiHeader.biHeight = -dragImageRect.Height();
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;

				dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
			} else {
				dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageRect.Width(), dragImageRect.Height());
			}
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
			memoryDC.FillRect(&dragImageRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
			CBitmap dragImageMask;
			dragImageMask.CreateBitmap(dragImageRect.Width(), dragImageRect.Height(), 1, 1, NULL);
			HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);
			maskMemoryDC.FillRect(&dragImageRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

			// draw each single drag image into our bitmap
			BOOL rightToLeft = FALSE;
			HWND hWndTabStrip = properties.GetTabStripHWnd();
			if(IsWindow(hWndTabStrip)) {
				rightToLeft = ((CWindow(hWndTabStrip).GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
			}
			#ifdef USE_STL
				for(size_t i = 0; i < imageLists.size(); ++i) {
			#else
				for(size_t i = 0; i < imageLists.GetCount(); ++i) {
			#endif
				if(rightToLeft) {
					ImageList_Draw(imageLists[i], 0, memoryDC, boundingRect.right - tabBoundingRects[i].right, tabBoundingRects[i].top - boundingRect.top, ILD_NORMAL);
					ImageList_Draw(imageLists[i], 0, maskMemoryDC, boundingRect.right - tabBoundingRects[i].right, tabBoundingRects[i].top - boundingRect.top, ILD_MASK);
				} else {
					ImageList_Draw(imageLists[i], 0, memoryDC, tabBoundingRects[i].left - boundingRect.left, tabBoundingRects[i].top - boundingRect.top, ILD_NORMAL);
					ImageList_Draw(imageLists[i], 0, maskMemoryDC, tabBoundingRects[i].left - boundingRect.left, tabBoundingRects[i].top - boundingRect.top, ILD_MASK);
				}

				ImageList_Destroy(imageLists[i]);
			}

			// clean up
			#ifdef USE_STL
				imageLists.clear();
				tabBoundingRects.clear();
			#else
				imageLists.RemoveAll();
				tabBoundingRects.RemoveAll();
			#endif
			memoryDC.SelectBitmap(hPreviousBitmap);
			maskMemoryDC.SelectBitmap(hPreviousBitmapMask);
			ReleaseDC(HWND_DESKTOP, hCompatibleDC);

			// create the imagelist
			HIMAGELIST hImageList = ImageList_Create(dragImageRect.Width(), dragImageRect.Height(), (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
			ImageList_SetBkColor(hImageList, CLR_NONE);
			ImageList_Add(hImageList, dragImage, dragImageMask);
			*phImageList = HandleToLong(hImageList);

			if(*phImageList) {
				if(pXUpperLeft) {
					*pXUpperLeft = upperLeftPoint.x;
				}
				if(pYUpperLeft) {
					*pYUpperLeft = upperLeftPoint.y;
				}
				return S_OK;
			}
			break;
		}
	}

	return E_FAIL;
}

STDMETHODIMP TabStripTabContainer::Remove(LONG tabID, VARIANT_BOOL removePhysically/* = VARIANT_FALSE*/)
{
	#ifdef USE_STL
		for(size_t i = 0; i < properties.tabs.size(); ++i) {
			if(properties.tabs[i] == tabID) {
				if(removePhysically == VARIANT_FALSE) {
					properties.tabs.erase(std::find(properties.tabs.begin(), properties.tabs.end(), tabID));
					if(i < (size_t) properties.nextEnumeratedTab) {
						--properties.nextEnumeratedTab;
					}
				} else {
					HWND hWndTabStrip = properties.GetTabStripHWnd();
					if(IsWindow(hWndTabStrip)) {
						int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabID);
						ATLASSERT(tabIndex >= 0);
						SendMessage(hWndTabStrip, TCM_DELETEITEM, tabIndex, 0);
					}

					// our owner will notify us about the deletion, so we don't need to remove the tab explicitly
				}

				return S_OK;
			}
		}
	#else
		//for(size_t i = 0; i < properties.tabs.GetCount(); ++i) {
		for(size_t i = 0; i < tabs.GetCount(); ++i) {
			//if(properties.tabs[i] == tabID) {
			if(tabs[i] == tabID) {
				if(removePhysically == VARIANT_FALSE) {
					//properties.tabs.RemoveAt(i);
					tabs.RemoveAt(i);
					if(i < (size_t) properties.nextEnumeratedTab) {
						--properties.nextEnumeratedTab;
					}
				} else {
					HWND hWndTabStrip = properties.GetTabStripHWnd();
					if(IsWindow(hWndTabStrip)) {
						int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabID);
						ATLASSERT(tabIndex >= 0);
						SendMessage(hWndTabStrip, TCM_DELETEITEM, tabIndex, 0);
					}

					// our owner will notify us about the deletion, so we don't need to remove the tab explicitly
				}

				return S_OK;
			}
		}
	#endif
	return E_FAIL;
}

STDMETHODIMP TabStripTabContainer::RemoveAll(VARIANT_BOOL removePhysically/* = VARIANT_FALSE*/)
{
	if(removePhysically != VARIANT_FALSE) {
		HWND hWndTabStrip = properties.GetTabStripHWnd();
		if(IsWindow(hWndTabStrip)) {
			#ifdef USE_STL
				while(properties.tabs.size() > 0) {
					int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(*properties.tabs.begin());
					ATLASSERT(tabIndex >= 0);
					SendMessage(hWndTabStrip, TCM_DELETEITEM, tabIndex, 0);
				}
			#else
				//while(properties.tabs.GetCount() > 0) {
				while(tabs.GetCount() > 0) {
					//int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(properties.tabs[0]);
					int tabIndex = properties.pOwnerTabStrip->IDToTabIndex(tabs[0]);
					ATLASSERT(tabIndex >= 0);
					SendMessage(hWndTabStrip, TCM_DELETEITEM, tabIndex, 0);
				}
			#endif
		}
	} else {
		#ifdef USE_STL
			properties.tabs.clear();
		#else
			//properties.tabs.RemoveAll();
			tabs.RemoveAll();
		#endif
	}
	Reset();
	return S_OK;
}