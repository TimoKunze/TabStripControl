// TabStrip.cpp: Superclasses SysTabControl32.

#include "stdafx.h"
#include "TabStrip.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
FONTDESC TabStrip::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


TabStrip::TabStrip()
{
	pSimpleFrameSite = NULL;

	properties.font.InitializePropertyWatcher(this, DISPID_TABSTRIPCTL_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_TABSTRIPCTL_MOUSEICON);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// initialize
	tabUnderMouse = -1;
	pToolTipBuffer = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, (MAX_PATH + 1) * sizeof(WCHAR));

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	if(RunTimeHelper::IsVista()) {
		CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
		ATLASSUME(pWICImagingFactory);
	}
}

TabStrip::~TabStrip()
{
	if(pToolTipBuffer) {
		HeapFree(GetProcessHeap(), 0, pToolTipBuffer);
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP TabStrip::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ITabStrip, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP TabStrip::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 21 VT_I4 properties...
	pSize->QuadPart += 21 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 16 VT_BOOL properties...
	pSize->QuadPart += 16 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));

	// ...and 2 VT_DISPATCH properties
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP TabStrip::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0;
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x0A0A0A0A/*4x AppID*/) {
		// might be a legacy stream, that starts with the ATL version
		if(signature == 0x0700 || signature == 0x0710 || signature == 0x0800 || signature == 0x0900) {
			version = 0x0099;
		} else {
			return E_FAIL;
		}
	}
	//LONG version = 0;
	if(version != 0x0099) {
		if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
			return hr;
		}
		if(version > 0x0102) {
			return E_FAIL;
		}
	}

	DWORD atlVersion;
	if(version == 0x0099) {
		atlVersion = 0x0900;
	} else {
		if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
			return hr;
		}
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(version != 0x0100) {
		if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
			return hr;
		}
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = NULL;
	if(version == 0x0100) {
		pfnReadVariantFromStream = ReadVariantFromStream_Legacy;
	} else {
		pfnReadVariantFromStream = ReadVariantFromStream;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL closeableTabs = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	/*if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;*/
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragActivateTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragScrollTimeBase = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS fixedTabWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL focusOnButtonDown = propertyValue.boolVal;

	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS horizontalPadding = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL hotTracking = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR insertMarkColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS minTabWidth = propertyValue.lVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL multiRow = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL multiSelect = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL ownerDrawn = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL raggedTabRows = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RegisterForOLEDragDropConstants registerForOLEDragDrop = static_cast<RegisterForOLEDragDropConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL scrollToOpposite = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showButtonSeparators = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showToolTips = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	StyleConstants style = static_cast<StyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	TabBoundingBoxDefinitionConstants tabBoundingBoxDefinition = static_cast<TabBoundingBoxDefinitionConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	TabCaptionAlignmentConstants tabCaptionAlignment = static_cast<TabCaptionAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS tabHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	TabPlacementConstants tabPlacement = static_cast<TabPlacementConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useFixedTabWidth = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS verticalPadding = propertyValue.lVal;

	CloseableTabsModeConstants closeableTabsMode = ctmDisplayOnAllTabs;
	OLEDragImageStyleConstants oleDragImageStyle = odistClassic;
	if(version >= 0x0101) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		oleDragImageStyle = static_cast<OLEDragImageStyleConstants>(propertyValue.lVal);
		if(version >= 0x0102) {
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
				return hr;
			}
			closeableTabsMode = static_cast<CloseableTabsModeConstants>(propertyValue.lVal);
		}
	}


	hr = put_AllowDragDrop(allowDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CloseableTabs(closeableTabs);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CloseableTabsMode(closeableTabsMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	/*hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}*/
	hr = put_DragActivateTime(dragActivateTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragScrollTimeBase(dragScrollTimeBase);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FixedTabWidth(fixedTabWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FocusOnButtonDown(focusOnButtonDown);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HorizontalPadding(horizontalPadding);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HotTracking(hotTracking);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_InsertMarkColor(insertMarkColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MinTabWidth(minTabWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MultiRow(multiRow);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MultiSelect(multiSelect);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OLEDragImageStyle(oleDragImageStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OwnerDrawn(ownerDrawn);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RaggedTabRows(raggedTabRows);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ScrollToOpposite(scrollToOpposite);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowButtonSeparators(showButtonSeparators);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowToolTips(showToolTips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Style(style);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TabBoundingBoxDefinition(tabBoundingBoxDefinition);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TabCaptionAlignment(tabCaptionAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TabHeight(tabHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TabPlacement(tabPlacement);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseFixedTabWidth(useFixedTabWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_VerticalPadding(verticalPadding);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP TabStrip::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x0A0A0A0A/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0102;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.closeableTabs);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	/*propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;*/
	propertyValue.lVal = properties.dragActivateTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.dragScrollTimeBase;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.fixedTabWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.focusOnButtonDown);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	VARTYPE vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.horizontalPadding;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.hotTracking);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.insertMarkColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.minTabWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.multiRow);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.multiSelect);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.ownerDrawn);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.raggedTabRows);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.registerForOLEDragDrop;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.scrollToOpposite);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showButtonSeparators);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showToolTips);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.style;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.tabBoundingBoxDefinition;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.tabCaptionAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.tabHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.tabPlacement;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useFixedTabWidth);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.verticalPadding;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0101 starts here
	propertyValue.lVal = properties.oleDragImageStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0102 starts here
	propertyValue.lVal = properties.closeableTabsMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND TabStrip::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_TAB_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<TabStrip>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT TabStrip::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("TabStrip ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void TabStrip::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP TabStrip::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	pSimpleFrameSite = NULL;
	if(pClientSite) {
		pClientSite->QueryInterface(IID_ISimpleFrameSite, reinterpret_cast<LPVOID*>(&pSimpleFrameSite));
	}

	HRESULT hr = CComControl<TabStrip>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	return hr;
}

STDMETHODIMP TabStrip::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL TabStrip::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		if(dialogCode & DLGC_WANTTAB) {
			if(pMessage->wParam == VK_TAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
		switch(pMessage->wParam) {
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				if(dialogCode & DLGC_WANTARROWS) {
					if(!(GetKeyState(VK_MENU) & 0x8000)) {
						SendMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam);
						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
			case VK_TAB:
				if(pMessage->message == WM_KEYDOWN) {
					if(GetAsyncKeyState(VK_CONTROL) & 0x8000) {
						int activeTab = SendMessage(TCM_GETCURSEL, 0, 0);
						activeTab += (GetAsyncKeyState(VK_SHIFT) & 0x8000 ? -1 : 1);
						int tabCount = SendMessage(TCM_GETITEMCOUNT, 0, 0);
						if(activeTab < 0) {
							activeTab = tabCount - 1;
						} else if(activeTab >= tabCount) {
							activeTab = 0;
						}
						SendMessage(TCM_SETCURSEL, activeTab, 0);

						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
		}
	}
	return CComControl<TabStrip>::PreTranslateAccelerator(pMessage, hReturnValue);
}

HIMAGELIST TabStrip::CreateLegacyDragImage(int tabIndex, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle)
{
	/********************************************************************************************************
	 * Known problems:                                                                                      *
	 * - We use hardcoded margins.                                                                          *
	 ********************************************************************************************************/

	// retrieve tab details
	BOOL tabIsActive = (tabIndex == SendMessage(TCM_GETCURSEL, 0, 0));

	// retrieve window details
	BOOL layoutRTL = ((GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);

	// create the DCs we'll draw into
	HDC hCompatibleDC = GetDC();
	CDC memoryDC;
	memoryDC.CreateCompatibleDC(hCompatibleDC);
	CDC maskMemoryDC;
	maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

	// calculate the bounding rectangle
	CRect tabBoundingRect;
	SendMessage(TCM_GETITEMRECT, tabIndex, reinterpret_cast<LPARAM>(&tabBoundingRect));
	if(tabIsActive) {
		tabBoundingRect.InflateRect(2, 2);
	} else {
		tabBoundingRect.InflateRect(-1, 0);
	}
	if(pBoundingRectangle) {
		*pBoundingRectangle = tabBoundingRect;
	}

	// calculate drag image size and upper-left corner
	SIZE dragImageSize = {0};
	if(pUpperLeftPoint) {
		pUpperLeftPoint->x = tabBoundingRect.left;
		pUpperLeftPoint->y = tabBoundingRect.top;
	}
	dragImageSize.cx = tabBoundingRect.Width();
	dragImageSize.cy = tabBoundingRect.Height();

	// offset RECTs
	SIZE offset = {0};
	offset.cx = tabBoundingRect.left;
	offset.cy = tabBoundingRect.top;
	tabBoundingRect.OffsetRect(-offset.cx, -offset.cy);

	// setup the DCs we'll draw into
	memoryDC.SetBkColor(GetSysColor(COLOR_WINDOW));
	memoryDC.SetTextColor(GetSysColor(COLOR_WINDOWTEXT));
	memoryDC.SetBkMode(TRANSPARENT);

	// create drag image bitmap
	/* NOTE: We prefer creating 32bpp drag images, because this improves performance of
	         TabStripTabContainer::CreateDragImage(). */
	BOOL doAlphaChannelProcessing = RunTimeHelper::IsCommCtrl6();
	BITMAPINFO bitmapInfo = {0};
	if(doAlphaChannelProcessing) {
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = dragImageSize.cx;
		bitmapInfo.bmiHeader.biHeight = -dragImageSize.cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = BI_RGB;
	}
	CBitmap dragImage;
	LPRGBQUAD pDragImageBits = NULL;
	if(doAlphaChannelProcessing) {
		dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
	} else {
		dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageSize.cx, dragImageSize.cy);
	}
	HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
	CBitmap dragImageMask;
	dragImageMask.CreateBitmap(dragImageSize.cx, dragImageSize.cy, 1, 1, NULL);
	HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);

	// initialize the bitmap
	RECT rc = tabBoundingRect;
	if(doAlphaChannelProcessing && pDragImageBits) {
		// we need a transparent background
		LPRGBQUAD pPixel = pDragImageBits;
		for(int y = 0; y < dragImageSize.cy; ++y) {
			for(int x = 0; x < dragImageSize.cx; ++x, ++pPixel) {
				pPixel->rgbRed = 0xFF;
				pPixel->rgbGreen = 0xFF;
				pPixel->rgbBlue = 0xFF;
				pPixel->rgbReserved = 0x00;
			}
		}
	} else {
		memoryDC.FillRect(&rc, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
	}
	maskMemoryDC.FillRect(&rc, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));

	// draw the tab
	memoryDC.SetViewportOrg(-offset.cx, -offset.cy);
	SendMessage(WM_PRINTCLIENT, reinterpret_cast<WPARAM>(static_cast<HDC>(memoryDC)));
	memoryDC.SetViewportOrg(0, 0);

	if(doAlphaChannelProcessing && pDragImageBits) {
		// correct the alpha channel
		LPRGBQUAD pPixel = pDragImageBits;
		POINT pt;
		for(pt.y = 0; pt.y < dragImageSize.cy; ++pt.y) {
			for(pt.x = 0; pt.x < dragImageSize.cx; ++pt.x, ++pPixel) {
				if(layoutRTL) {
					// we're working on raw data, so we've to handle WS_EX_LAYOUTRTL ourselves
					POINT pt2 = pt;
					pt2.x = dragImageSize.cx - pt.x - 1;
					if(maskMemoryDC.GetPixel(pt2.x, pt2.y) == 0x00000000) {
						if(pPixel->rgbReserved == 0x00) {
							pPixel->rgbReserved = 0xFF;
						}
					}
				} else {
					// layout is left to right
					if(maskMemoryDC.GetPixel(pt.x, pt.y) == 0x00000000) {
						if(pPixel->rgbReserved == 0x00) {
							pPixel->rgbReserved = 0xFF;
						}
					}
				}
			}
		}
	}

	memoryDC.SelectBitmap(hPreviousBitmap);
	maskMemoryDC.SelectBitmap(hPreviousBitmapMask);

	// create the imagelist
	HIMAGELIST hDragImageList = ImageList_Create(dragImageSize.cx, dragImageSize.cy, (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
	ImageList_SetBkColor(hDragImageList, CLR_NONE);
	ImageList_Add(hDragImageList, dragImage, dragImageMask);

	ReleaseDC(hCompatibleDC);

	return hDragImageList;
}

BOOL TabStrip::CreateLegacyOLEDragImage(ITabStripTabContainer* pTabs, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pTabs);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	// use a normal legacy drag image as base
	OLE_HANDLE h = NULL;
	OLE_XPOS_PIXELS xUpperLeft = 0;
	OLE_YPOS_PIXELS yUpperLeft = 0;
	pTabs->CreateDragImage(&xUpperLeft, &yUpperLeft, &h);
	if(h) {
		HIMAGELIST hImageList = static_cast<HIMAGELIST>(LongToHandle(h));

		// retrieve the drag image's size
		int bitmapHeight;
		int bitmapWidth;
		ImageList_GetIconSize(hImageList, &bitmapWidth, &bitmapHeight);
		pDragImage->sizeDragImage.cx = bitmapWidth;
		pDragImage->sizeDragImage.cy = bitmapHeight;

		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		pDragImage->hbmpDragImage = NULL;

		if(RunTimeHelper::IsCommCtrl6()) {
			// handle alpha channel
			IImageList* pImgLst = NULL;
			HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
			if(hMod) {
				typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
				HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
				if(pfnHIMAGELIST_QueryInterface) {
					pfnHIMAGELIST_QueryInterface(hImageList, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst));
				}
				FreeLibrary(hMod);
			}
			if(!pImgLst) {
				pImgLst = reinterpret_cast<IImageList*>(hImageList);
				pImgLst->AddRef();
			}
			ATLASSUME(pImgLst);

			DWORD imageFlags = 0;
			pImgLst->GetItemFlags(0, &imageFlags);
			if(imageFlags & ILIF_ALPHA) {
				// the drag image makes use of the alpha channel
				IMAGEINFO imageInfo = {0};
				ImageList_GetImageInfo(hImageList, 0, &imageInfo);

				// fetch raw data
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
				bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;
				LPRGBQUAD pSourceBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * sizeof(RGBQUAD)));
				GetDIBits(memoryDC, imageInfo.hbmImage, 0, pDragImage->sizeDragImage.cy, pSourceBits, &bitmapInfo, DIB_RGB_COLORS);
				// create target bitmap
				LPRGBQUAD pDragImageBits = NULL;
				pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
				ATLASSERT(pDragImageBits);
				pDragImage->crColorKey = 0xFFFFFFFF;

				if(pDragImageBits) {
					// transfer raw data
					CopyMemory(pDragImageBits, pSourceBits, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * 4);
				}

				// clean up
				HeapFree(GetProcessHeap(), 0, pSourceBits);
				DeleteObject(imageInfo.hbmImage);
				DeleteObject(imageInfo.hbmMask);
			}
			pImgLst->Release();
		}

		if(!pDragImage->hbmpDragImage) {
			// fallback mode
			memoryDC.SetBkMode(TRANSPARENT);

			// create target bitmap
			HDC hCompatibleDC = ::GetDC(HWND_DESKTOP);
			pDragImage->hbmpDragImage = CreateCompatibleBitmap(hCompatibleDC, bitmapWidth, bitmapHeight);
			::ReleaseDC(HWND_DESKTOP, hCompatibleDC);
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(pDragImage->hbmpDragImage);

			// draw target bitmap
			pDragImage->crColorKey = RGB(0xF4, 0x00, 0x00);
			CBrush backroundBrush;
			backroundBrush.CreateSolidBrush(pDragImage->crColorKey);
			memoryDC.FillRect(CRect(0, 0, bitmapWidth, bitmapHeight), backroundBrush);
			ImageList_Draw(hImageList, 0, memoryDC, 0, 0, ILD_NORMAL);

			// clean up
			memoryDC.SelectBitmap(hPreviousBitmap);
		}

		ImageList_Destroy(hImageList);

		if(pDragImage->hbmpDragImage) {
			// retrieve the offset
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			if(GetExStyle() & WS_EX_LAYOUTRTL) {
				pDragImage->ptOffset.x = xUpperLeft + pDragImage->sizeDragImage.cx - mousePosition.x;
			} else {
				pDragImage->ptOffset.x = mousePosition.x - xUpperLeft;
			}
			pDragImage->ptOffset.y = mousePosition.y - yUpperLeft;

			succeeded = TRUE;
		}
	}

	return succeeded;
}

BOOL TabStrip::CreateVistaOLEDragImage(ITabStripTabContainer* pTabs, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pTabs);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	CTheme themingEngine;
	themingEngine.OpenThemeData(NULL, VSCLASS_DRAGDROP);
	if(themingEngine.IsThemeNull()) {
		// FIXME: What should we do here?
		ATLASSERT(FALSE && "Current theme does not define the \"DragDrop\" class.");
	} else {
		// retrieve the drag image's size
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();

		themingEngine.GetThemePartSize(memoryDC, DD_IMAGEBG, 1, NULL, TS_TRUE, &pDragImage->sizeDragImage);
		MARGINS margins = {0};
		themingEngine.GetThemeMargins(memoryDC, DD_IMAGEBG, 1, TMT_CONTENTMARGINS, NULL, &margins);
		pDragImage->sizeDragImage.cx -= margins.cxLeftWidth + margins.cxRightWidth;
		pDragImage->sizeDragImage.cy -= margins.cyTopHeight + margins.cyBottomHeight;
	}

	ATLASSERT(pDragImage->sizeDragImage.cx > 0);
	ATLASSERT(pDragImage->sizeDragImage.cy > 0);

	// create target bitmap
	BITMAPINFO bitmapInfo = {0};
	bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
	bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
	bitmapInfo.bmiHeader.biPlanes = 1;
	bitmapInfo.bmiHeader.biBitCount = 32;
	bitmapInfo.bmiHeader.biCompression = BI_RGB;
	LPRGBQUAD pDragImageBits = NULL;
	pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);

	HIMAGELIST hSourceImageList = (cachedSettings.hHighResImageList ? cachedSettings.hHighResImageList : cachedSettings.hImageList);
	if(!hSourceImageList) {
		// report success, although we've an empty drag image
		return TRUE;
	}

	IImageList2* pImgLst = NULL;
	HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
	if(hMod) {
		typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
		HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
		if(pfnHIMAGELIST_QueryInterface) {
			pfnHIMAGELIST_QueryInterface(hSourceImageList, IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
		}
		FreeLibrary(hMod);
	}
	if(!pImgLst) {
		IImageList* p = reinterpret_cast<IImageList*>(hSourceImageList);
		p->QueryInterface(IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
	}
	ATLASSUME(pImgLst);

	if(pImgLst) {
		LONG numberOfTabs = 0;
		pTabs->Count(&numberOfTabs);
		ATLASSERT(numberOfTabs > 0);
		// don't display more than 5 (10) thumbnails
		numberOfTabs = min(numberOfTabs, (hSourceImageList == cachedSettings.hHighResImageList ? 5 : 10));

		CComPtr<IUnknown> pUnknownEnum = NULL;
		pTabs->get__NewEnum(&pUnknownEnum);
		CComQIPtr<IEnumVARIANT> pEnum = pUnknownEnum;
		ATLASSUME(pEnum);
		if(pEnum) {
			int cx = 0;
			int cy = 0;
			pImgLst->GetIconSize(&cx, &cy);
			SIZE thumbnailSize;
			thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfTabs - 1);
			if(thumbnailSize.cy < 8) {
				// don't get smaller than 8x8 thumbnails
				numberOfTabs = (pDragImage->sizeDragImage.cy - 8) / 3 + 1;
				thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfTabs - 1);
			}
			thumbnailSize.cx = thumbnailSize.cy;
			int thumbnailBufferSize = thumbnailSize.cx * thumbnailSize.cy * sizeof(RGBQUAD);
			LPRGBQUAD pThumbnailBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, thumbnailBufferSize));
			ATLASSERT(pThumbnailBits);
			if(pThumbnailBits) {
				// iterate over the dragged tabs
				VARIANT v;
				int i = 0;
				CComPtr<ITabStripTab> pTab = NULL;
				while(pEnum->Next(1, &v, NULL) == S_OK) {
					if(v.vt == VT_DISPATCH) {
						v.pdispVal->QueryInterface(IID_ITabStripTab, reinterpret_cast<LPVOID*>(&pTab));
						ATLASSUME(pTab);
						if(pTab) {
							// get the tab's icon
							LONG icon = 0;
							pTab->get_IconIndex(&icon);

							pImgLst->ForceImagePresent(icon, ILFIP_ALWAYS);
							HICON hIcon = NULL;
							pImgLst->GetIcon(icon, ILD_TRANSPARENT, &hIcon);
							ATLASSERT(hIcon);
							if(hIcon) {
								// finally create the thumbnail
								ZeroMemory(pThumbnailBits, thumbnailBufferSize);
								HRESULT hr = CreateThumbnail(hIcon, thumbnailSize, pThumbnailBits, TRUE);
								DestroyIcon(hIcon);
								if(FAILED(hr)) {
									pTab = NULL;
									VariantClear(&v);
									break;
								}

								// add the thumbail to the drag image keeping the alpha channel intact
								if(i == 0) {
									LPRGBQUAD pDragImagePixel = pDragImageBits;
									LPRGBQUAD pThumbnailPixel = pThumbnailBits;
									for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx, pThumbnailPixel += thumbnailSize.cx) {
										CopyMemory(pDragImagePixel, pThumbnailPixel, thumbnailSize.cx * sizeof(RGBQUAD));
									}
								} else {
									LPRGBQUAD pDragImagePixel = pDragImageBits;
									LPRGBQUAD pThumbnailPixel = pThumbnailBits;
									pDragImagePixel += 3 * i * pDragImage->sizeDragImage.cx;
									for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx) {
										LPRGBQUAD p = pDragImagePixel + 2 * i;
										for(int x = 0; x < thumbnailSize.cx; ++x, ++p, ++pThumbnailPixel) {
											// merge the pixels
											p->rgbRed = pThumbnailPixel->rgbRed * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbRed / 0xFF;
											p->rgbGreen = pThumbnailPixel->rgbGreen * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbGreen / 0xFF;
											p->rgbBlue = pThumbnailPixel->rgbBlue * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbBlue / 0xFF;
											p->rgbReserved = pThumbnailPixel->rgbReserved + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbReserved / 0xFF;
										}
									}
								}
							}

							++i;
							pTab = NULL;
							if(i == numberOfTabs) {
								VariantClear(&v);
								break;
							}
						}
					}
					VariantClear(&v);
				}
				HeapFree(GetProcessHeap(), 0, pThumbnailBits);
				succeeded = TRUE;
			}
		}

		pImgLst->Release();
	}

	return succeeded;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP TabStrip::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);

	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
		if(dragDropStatus.useItemCountLabelHack) {
			dragDropStatus.pDropTargetHelper->DragLeave();
			dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
			dragDropStatus.useItemCountLabelHack = FALSE;
		}
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP TabStrip::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSource
STDMETHODIMP TabStrip::GiveFeedback(DWORD effect)
{
	VARIANT_BOOL useDefaultCursors = VARIANT_TRUE;
	//if(flags.usingThemes && RunTimeHelper::IsVista()) {
		ATLASSUME(dragDropStatus.pSourceDataObject);

		BOOL isShowingLayered = FALSE;
		FORMATETC format = {0};
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("IsShowingLayered")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		STGMEDIUM medium = {0};
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				isShowingLayered = *static_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}
		BOOL useDropDescriptionHack = FALSE;
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				useDropDescriptionHack = *static_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}

		if(isShowingLayered && properties.oleDragImageStyle != odistClassic) {
			SetCursor(static_cast<HCURSOR>(LoadImage(NULL, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED)));
			useDefaultCursors = VARIANT_FALSE;
		}
		if(useDropDescriptionHack) {
			// this will make drop descriptions work
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("DragWindow")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
				if(medium.hGlobal) {
					// WM_USER + 1 (with wParam = 0 and lParam = 0) hides the drag image
					#define WM_SETDROPEFFECT				WM_USER + 2     // (wParam = DCID_*, lParam = 0)
					#define DDWM_UPDATEWINDOW				WM_USER + 3     // (wParam = 0, lParam = 0)
					typedef enum DROPEFFECTS
					{
						DCID_NULL = 0,
						DCID_NO = 1,
						DCID_MOVE = 2,
						DCID_COPY = 3,
						DCID_LINK = 4,
						DCID_MAX = 5
					} DROPEFFECTS;

					HWND hWndDragWindow = *static_cast<HWND*>(GlobalLock(medium.hGlobal));
					GlobalUnlock(medium.hGlobal);

					DROPEFFECTS dropEffect = DCID_NULL;
					switch(effect) {
						case DROPEFFECT_NONE:
							dropEffect = DCID_NO;
							break;
						case DROPEFFECT_COPY:
							dropEffect = DCID_COPY;
							break;
						case DROPEFFECT_MOVE:
							dropEffect = DCID_MOVE;
							break;
						case DROPEFFECT_LINK:
							dropEffect = DCID_LINK;
							break;
					}
					if(::IsWindow(hWndDragWindow)) {
						::PostMessage(hWndDragWindow, WM_SETDROPEFFECT, dropEffect, 0);
					}
				}
				ReleaseStgMedium(&medium);
			}
		}
	//}

	Raise_OLEGiveFeedback(effect, &useDefaultCursors);
	return (useDefaultCursors == VARIANT_FALSE ? S_OK : DRAGDROP_S_USEDEFAULTCURSORS);
}

STDMETHODIMP TabStrip::QueryContinueDrag(BOOL pressedEscape, DWORD keyState)
{
	HRESULT actionToContinueWith = S_OK;
	if(pressedEscape) {
		actionToContinueWith = DRAGDROP_S_CANCEL;
	} else if(!(keyState & dragDropStatus.draggingMouseButton)) {
		actionToContinueWith = DRAGDROP_S_DROP;
	}
	Raise_OLEQueryContinueDrag(pressedEscape, keyState, &actionToContinueWith);
	return actionToContinueWith;
}
// implementation of IDropSource
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSourceNotify
STDMETHODIMP TabStrip::DragEnterTarget(HWND hWndTarget)
{
	Raise_OLEDragEnterPotentialTarget(HandleToLong(hWndTarget));
	return S_OK;
}

STDMETHODIMP TabStrip::DragLeaveTarget(void)
{
	Raise_OLEDragLeavePotentialTarget();
	return S_OK;
}
// implementation of IDropSourceNotify
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP TabStrip::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP TabStrip::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_TABSTRIPCTL_APPEARANCE:
		case DISPID_TABSTRIPCTL_BORDERSTYLE:
		case DISPID_TABSTRIPCTL_CLOSEABLETABS:
		case DISPID_TABSTRIPCTL_CLOSEABLETABSMODE:
		case DISPID_TABSTRIPCTL_DISPLAYAREAHEIGHT:
		case DISPID_TABSTRIPCTL_DISPLAYAREALEFT:
		case DISPID_TABSTRIPCTL_DISPLAYAREATOP:
		case DISPID_TABSTRIPCTL_DISPLAYAREAWIDTH:
		case DISPID_TABSTRIPCTL_FIXEDTABWIDTH:
		case DISPID_TABSTRIPCTL_HORIZONTALPADDING:
		case DISPID_TABSTRIPCTL_MINTABWIDTH:
		case DISPID_TABSTRIPCTL_MOUSEICON:
		case DISPID_TABSTRIPCTL_MOUSEPOINTER:
		case DISPID_TABSTRIPCTL_SHOWBUTTONSEPARATORS:
		case DISPID_TABSTRIPCTL_STYLE:
		case DISPID_TABSTRIPCTL_TABCAPTIONALIGNMENT:
		case DISPID_TABSTRIPCTL_TABHEIGHT:
		case DISPID_TABSTRIPCTL_TABPLACEMENT:
		case DISPID_TABSTRIPCTL_USEFIXEDTABWIDTH:
		case DISPID_TABSTRIPCTL_VERTICALPADDING:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_TABSTRIPCTL_ACTIVETAB:
		case DISPID_TABSTRIPCTL_CARETTAB:
		case DISPID_TABSTRIPCTL_DISABLEDEVENTS:
		//case DISPID_TABSTRIPCTL_DONTREDRAW:
		case DISPID_TABSTRIPCTL_FOCUSONBUTTONDOWN:
		case DISPID_TABSTRIPCTL_HOTTRACKING:
		case DISPID_TABSTRIPCTL_HOVERTIME:
		case DISPID_TABSTRIPCTL_MULTIROW:
		case DISPID_TABSTRIPCTL_MULTISELECT:
		case DISPID_TABSTRIPCTL_OWNERDRAWN:
		case DISPID_TABSTRIPCTL_PROCESSCONTEXTMENUKEYS:
		case DISPID_TABSTRIPCTL_RAGGEDTABROWS:
		//case DISPID_TABSTRIPCTL_REFLECTCONTEXTMENUMESSAGES:
		case DISPID_TABSTRIPCTL_RIGHTTOLEFT:
		case DISPID_TABSTRIPCTL_SCROLLTOOPPOSITE:
		case DISPID_TABSTRIPCTL_SHOWTOOLTIPS:
		case DISPID_TABSTRIPCTL_TABBOUNDINGBOXDEFINITION:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_TABSTRIPCTL_INSERTMARKCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_TABSTRIPCTL_APPID:
		case DISPID_TABSTRIPCTL_APPNAME:
		case DISPID_TABSTRIPCTL_APPSHORTNAME:
		case DISPID_TABSTRIPCTL_BUILD:
		case DISPID_TABSTRIPCTL_CHARSET:
		case DISPID_TABSTRIPCTL_HDRAGIMAGELIST:
		case DISPID_TABSTRIPCTL_HHIGHRESIMAGELIST:
		case DISPID_TABSTRIPCTL_HIMAGELIST:
		case DISPID_TABSTRIPCTL_HWND:
		case DISPID_TABSTRIPCTL_HWNDARROWBUTTONS:
		case DISPID_TABSTRIPCTL_HWNDTOOLTIP:
		case DISPID_TABSTRIPCTL_ISRELEASE:
		case DISPID_TABSTRIPCTL_PROGRAMMER:
		case DISPID_TABSTRIPCTL_TESTER:
		case DISPID_TABSTRIPCTL_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_TABSTRIPCTL_ALLOWDRAGDROP:
		case DISPID_TABSTRIPCTL_DRAGACTIVATETIME:
		case DISPID_TABSTRIPCTL_DRAGGEDTABS:
		case DISPID_TABSTRIPCTL_DRAGSCROLLTIMEBASE:
		case DISPID_TABSTRIPCTL_DROPHILITEDTAB:
		case DISPID_TABSTRIPCTL_OLEDRAGIMAGESTYLE:
		case DISPID_TABSTRIPCTL_REGISTERFOROLEDRAGDROP:
		case DISPID_TABSTRIPCTL_SHOWDRAGIMAGE:
		case DISPID_TABSTRIPCTL_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_TABSTRIPCTL_FONT:
		case DISPID_TABSTRIPCTL_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_TABSTRIPCTL_TABS:
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_TABSTRIPCTL_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString TabStrip::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString TabStrip::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString TabStrip::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=TabStrip&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString TabStrip::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString TabStrip::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString TabStrip::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP TabStrip::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_TABSTRIPCTL_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_TABSTRIPCTL_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_TABSTRIPCTL_CLOSEABLETABSMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.closeableTabsMode), description);
			break;
		case DISPID_TABSTRIPCTL_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_TABSTRIPCTL_OLEDRAGIMAGESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.oleDragImageStyle), description);
			break;
		case DISPID_TABSTRIPCTL_REGISTERFOROLEDRAGDROP:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.registerForOLEDragDrop), description);
			break;
		case DISPID_TABSTRIPCTL_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_TABSTRIPCTL_STYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.style), description);
			break;
		case DISPID_TABSTRIPCTL_TABCAPTIONALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.tabCaptionAlignment), description);
			break;
		case DISPID_TABSTRIPCTL_TABPLACEMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.tabPlacement), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<TabStrip>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP TabStrip::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_TABSTRIPCTL_BORDERSTYLE:
		case DISPID_TABSTRIPCTL_CLOSEABLETABSMODE:
		case DISPID_TABSTRIPCTL_OLEDRAGIMAGESTYLE:
			c = 2;
			break;
		case DISPID_TABSTRIPCTL_APPEARANCE:
		case DISPID_TABSTRIPCTL_REGISTERFOROLEDRAGDROP:
		case DISPID_TABSTRIPCTL_STYLE:
		case DISPID_TABSTRIPCTL_TABCAPTIONALIGNMENT:
			c = 3;
			break;
		case DISPID_TABSTRIPCTL_RIGHTTOLEFT:
		case DISPID_TABSTRIPCTL_TABPLACEMENT:
			c = 4;
			break;
		case DISPID_TABSTRIPCTL_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<TabStrip>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if(property == DISPID_TABSTRIPCTL_CLOSEABLETABSMODE) {
			propertyValue++;
		} else if((property == DISPID_TABSTRIPCTL_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP TabStrip::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_TABSTRIPCTL_APPEARANCE:
		case DISPID_TABSTRIPCTL_BORDERSTYLE:
		case DISPID_TABSTRIPCTL_CLOSEABLETABSMODE:
		case DISPID_TABSTRIPCTL_MOUSEPOINTER:
		case DISPID_TABSTRIPCTL_OLEDRAGIMAGESTYLE:
		case DISPID_TABSTRIPCTL_REGISTERFOROLEDRAGDROP:
		case DISPID_TABSTRIPCTL_RIGHTTOLEFT:
		case DISPID_TABSTRIPCTL_STYLE:
		case DISPID_TABSTRIPCTL_TABCAPTIONALIGNMENT:
		case DISPID_TABSTRIPCTL_TABPLACEMENT:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<TabStrip>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP TabStrip::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	return IPerPropertyBrowsingImpl<TabStrip>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT TabStrip::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_TABSTRIPCTL_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TABSTRIPCTL_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TABSTRIPCTL_CLOSEABLETABSMODE:
			switch(cookie) {
				case ctmDisplayOnAllTabs:
					description = GetResStringWithNumber(IDP_CLOSEABLETABSMODEDISPLAYONALLTABS, ctmDisplayOnAllTabs);
					break;
				case ctmDisplayOnActiveTab:
					description = GetResStringWithNumber(IDP_CLOSEABLETABSMODEDISPLAYONACTIVETAB, ctmDisplayOnActiveTab);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TABSTRIPCTL_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TABSTRIPCTL_OLEDRAGIMAGESTYLE:
			switch(cookie) {
				case odistClassic:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLECLASSIC, odistClassic);
					break;
				case odistAeroIfAvailable:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLEAERO, odistAeroIfAvailable);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TABSTRIPCTL_REGISTERFOROLEDRAGDROP:
			switch(cookie) {
				case rfoddNoDragDrop:
					description = GetResStringWithNumber(IDP_REGISTERFOROLEDRAGDROPNONE, rfoddNoDragDrop);
					break;
				case rfoddNativeDragDrop:
					description = GetResStringWithNumber(IDP_REGISTERFOROLEDRAGDROPNATIVE, rfoddNativeDragDrop);
					break;
				case rfoddAdvancedDragDrop:
					description = GetResStringWithNumber(IDP_REGISTERFOROLEDRAGDROPADVANCED, rfoddAdvancedDragDrop);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TABSTRIPCTL_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TABSTRIPCTL_STYLE:
			switch(cookie) {
				case sTabs:
					description = GetResStringWithNumber(IDP_STYLETABS, sTabs);
					break;
				case sButtons:
					description = GetResStringWithNumber(IDP_STYLEBUTTONS, sButtons);
					break;
				case sFlatButtons:
					description = GetResStringWithNumber(IDP_STYLEFLATBUTTONS, sFlatButtons);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TABSTRIPCTL_TABCAPTIONALIGNMENT:
			switch(cookie) {
				case tcaNormal:
					description = GetResStringWithNumber(IDP_TABCAPTIONALIGNMENTNORMAL, tcaNormal);
					break;
				case tcaForceIconLeft:
					description = GetResStringWithNumber(IDP_TABCAPTIONALIGNMENTFORCEICONLEFT, tcaForceIconLeft);
					break;
				case tcaForceCaptionLeft:
					description = GetResStringWithNumber(IDP_TABCAPTIONALIGNMENTFORCECAPTIONLEFT, tcaForceCaptionLeft);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TABSTRIPCTL_TABPLACEMENT:
			switch(cookie) {
				case tpTop:
					description = GetResStringWithNumber(IDP_TABPLACEMENTTOP, tpTop);
					break;
				case tpBottom:
					description = GetResStringWithNumber(IDP_TABPLACEMENTBOTTOM, tpBottom);
					break;
				case tpLeft:
					description = GetResStringWithNumber(IDP_TABPLACEMENTLEFT, tpLeft);
					break;
				case tpRight:
					description = GetResStringWithNumber(IDP_TABPLACEMENTRIGHT, tpRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP TabStrip::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 4;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StockColorPage;
		pPropertyPages->pElems[2] = CLSID_StockFontPage;
		pPropertyPages->pElems[3] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP TabStrip::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<TabStrip>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP TabStrip::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP TabStrip::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = properties.hAcceleratorTable;
	pControlInfo->cAccel = static_cast<USHORT>(properties.hAcceleratorTable ? CopyAcceleratorTable(properties.hAcceleratorTable, NULL, 0) : 0);
	pControlInfo->dwFlags = 0;
	return S_OK;
}

STDMETHODIMP TabStrip::OnMnemonic(LPMSG pMessage)
{
	if(GetStyle() & WS_DISABLED) {
		return S_OK;
	}

	ATLASSERT(pMessage->message == WM_SYSKEYDOWN);

	SHORT pressedKeyCode = static_cast<SHORT>(pMessage->wParam);
	int tabToSelect = -1;

	TCITEM tab = {0};
	tab.mask = TCIF_TEXT;
	tab.cchTextMax = MAX_TABTEXTLENGTH;
	tab.pszText = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (tab.cchTextMax + 1) * sizeof(TCHAR)));
	if(tab.pszText) {
		int numberOfTabs = SendMessage(TCM_GETITEMCOUNT, 0, 0);
		for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
			SendMessage(TCM_GETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab));

			if(tab.pszText) {
				for(int i = lstrlen(tab.pszText) - 1; i > 0; --i) {
					if((tab.pszText[i - 1] == TEXT('&')) && (tab.pszText[i] != TEXT('&'))) {
						// TODO: Does this work with MFC?
						if((VkKeyScan(tab.pszText[i]) == pressedKeyCode) || (VkKeyScan(static_cast<TCHAR>(tolower(tab.pszText[i]))) == pressedKeyCode)) {
							tabToSelect = tabIndex;
							break;
						}
					}
				}
			}
		}
		HeapFree(GetProcessHeap(), 0, tab.pszText);
	}
	if(tabToSelect != -1) {
		if(tabToSelect != SendMessage(TCM_GETCURSEL, 0, 0)) {
			SendMessage(TCM_SETCURSEL, tabToSelect, 0);
		}
	}
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT TabStrip::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT TabStrip::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_TABSTRIPCTL_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT TabStrip::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP TabStrip::get_ActiveTab(ITabStripTab** ppActiveTab)
{
	ATLASSERT_POINTER(ppActiveTab, ITabStripTab*);
	if(!ppActiveTab) {
		return E_POINTER;
	}

	if(IsWindow()) {
		int activeTab = SendMessage(TCM_GETCURSEL, 0, 0);
		ClassFactory::InitTabStripTab(activeTab, this, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(ppActiveTab));
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP TabStrip::putref_ActiveTab(ITabStripTab* pNewActiveTab)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_ACTIVETAB);

	HRESULT hr = E_FAIL;

	int newActiveTab = -1;
	if(pNewActiveTab) {
		LONG l = -1;
		pNewActiveTab->get_Index(&l);
		newActiveTab = l;
	}

	if(IsWindow()) {
		SendMessage(TCM_SETCURSEL, newActiveTab, 0);
		hr = S_OK;
	}
	FireOnChanged(DISPID_TABSTRIPCTL_ACTIVETAB);
	return hr;
}

STDMETHODIMP TabStrip::get_AllowDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowDragDrop = ((SendMessage(TCM_GETEXTENDEDSTYLE, 0, 0) & TCS_EX_DETECTDRAGDROP) == TCS_EX_DETECTDRAGDROP);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowDragDrop);
	return S_OK;
}

STDMETHODIMP TabStrip::put_AllowDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_ALLOWDRAGDROP);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowDragDrop != b) {
		/* We set the style before adjusting properties.allowDragDrop because we need the old value of
		   properties.allowDragDrop in OnSetExtendedStyle. */
		if(IsWindow()) {
			SendMessage(TCM_SETEXTENDEDSTYLE, TCS_EX_DETECTDRAGDROP, (b ? TCS_EX_DETECTDRAGDROP : 0));
		}
		properties.allowDragDrop = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_ALLOWDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP TabStrip::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_APPEARANCE);

	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TABSTRIPCTL_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 10;
	return S_OK;
}

STDMETHODIMP TabStrip::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"TabStrip");
	return S_OK;
}

STDMETHODIMP TabStrip::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"TabStripCtl");
	return S_OK;
}

STDMETHODIMP TabStrip::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP TabStrip::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_BORDERSTYLE);

	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TABSTRIPCTL_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP TabStrip::get_CaretTab(ITabStripTab** ppCaretTab)
{
	ATLASSERT_POINTER(ppCaretTab, ITabStripTab*);
	if(!ppCaretTab) {
		return E_POINTER;
	}

	if(IsWindow()) {
		int caretTab = SendMessage(TCM_GETCURFOCUS, 0, 0);
		ClassFactory::InitTabStripTab(caretTab, this, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(ppCaretTab));
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP TabStrip::putref_CaretTab(ITabStripTab* pNewCaretTab)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_CARETTAB);

	HRESULT hr = E_FAIL;

	int newCaretTab = -1;
	if(pNewCaretTab) {
		LONG l = -1;
		pNewCaretTab->get_Index(&l);
		newCaretTab = l;
	}

	if(IsWindow()) {
		SendMessage(TCM_SETCURFOCUS, newCaretTab, 0);
		hr = S_OK;
	}
	FireOnChanged(DISPID_TABSTRIPCTL_CARETTAB);
	return hr;
}

STDMETHODIMP TabStrip::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP TabStrip::get_CloseableTabs(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.closeableTabs);
	return S_OK;
}

STDMETHODIMP TabStrip::put_CloseableTabs(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_CLOSEABLETABS);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.closeableTabs != b) {
		properties.closeableTabs = b;
		SetDirty(TRUE);

		if(properties.closeableTabs) {
			mouseStatus.overCloseButton = -1;
			mouseStatus.overCloseButtonOnMouseDown = -1;
		}
		FireViewChange();
		FireOnChanged(DISPID_TABSTRIPCTL_CLOSEABLETABS);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_CloseableTabsMode(CloseableTabsModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, CloseableTabsModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.closeableTabsMode;
	return S_OK;
}

STDMETHODIMP TabStrip::put_CloseableTabsMode(CloseableTabsModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_CLOSEABLETABSMODE);

	if(properties.closeableTabsMode != newValue) {
		properties.closeableTabsMode = newValue;
		SetDirty(TRUE);

		if(properties.closeableTabs) {
			mouseStatus.overCloseButton = -1;
			mouseStatus.overCloseButtonOnMouseDown = -1;
		}
		FireViewChange();
		FireOnChanged(DISPID_TABSTRIPCTL_CLOSEABLETABSMODE);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP TabStrip::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_DISABLEDEVENTS);

	if(properties.disabledEvents != newValue) {
		if((properties.disabledEvents & deMouseEvents) != (newValue & deMouseEvents)) {
			if(IsWindow()) {
				if(!properties.closeableTabs && ((newValue & deMouseEvents) == deMouseEvents)) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
					tabUnderMouse = -1;
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_DisplayAreaHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		CRect rc;
		GetWindowRect(&rc);
		ScreenToClient(&rc);
		SendMessage(TCM_ADJUSTRECT, FALSE, reinterpret_cast<LPARAM>(&rc));
		*pValue = rc.Height();
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_DisplayAreaLeft(OLE_XPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		RECT rc = {0};
		GetWindowRect(&rc);
		ScreenToClient(&rc);
		SendMessage(TCM_ADJUSTRECT, FALSE, reinterpret_cast<LPARAM>(&rc));
		*pValue = rc.left;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_DisplayAreaTop(OLE_YPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		RECT rc = {0};
		GetWindowRect(&rc);
		ScreenToClient(&rc);
		SendMessage(TCM_ADJUSTRECT, FALSE, reinterpret_cast<LPARAM>(&rc));
		*pValue = rc.top;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_DisplayAreaWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		CRect rc;
		GetWindowRect(&rc);
		ScreenToClient(&rc);
		SendMessage(TCM_ADJUSTRECT, FALSE, reinterpret_cast<LPARAM>(&rc));
		*pValue = rc.Width();
	} else {
		*pValue = 0;
	}
	return S_OK;
}
//
//STDMETHODIMP TabStrip::get_DontRedraw(VARIANT_BOOL* pValue)
//{
//	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
//	if(!pValue) {
//		return E_POINTER;
//	}
//
//	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
//	return S_OK;
//}
//
//STDMETHODIMP TabStrip::put_DontRedraw(VARIANT_BOOL newValue)
//{
//	PUTPROPPROLOG(DISPID_TABSTRIPCTL_DONTREDRAW);
//
//	UINT b = VARIANTBOOL2BOOL(newValue);
//	if(properties.dontRedraw != b) {
//		properties.dontRedraw = b;
//		SetDirty(TRUE);
//		if(IsWindow()) {
//			SetRedraw(!b);
//		}
//		FireOnChanged(DISPID_TABSTRIPCTL_DONTREDRAW);
//	}
//	return S_OK;
//}

STDMETHODIMP TabStrip::get_DragActivateTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragActivateTime;
	return S_OK;
}

STDMETHODIMP TabStrip::put_DragActivateTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_DRAGACTIVATETIME);

	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragActivateTime != newValue) {
		properties.dragActivateTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_DRAGACTIVATETIME);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_DraggedTabs(ITabStripTabContainer** ppTabs)
{
	ATLASSERT_POINTER(ppTabs, ITabStripTabContainer*);
	if(!ppTabs) {
		return E_POINTER;
	}

	*ppTabs = NULL;
	if(dragDropStatus.pDraggedTabs) {
		return dragDropStatus.pDraggedTabs->Clone(ppTabs);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_DragScrollTimeBase(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragScrollTimeBase;
	return S_OK;
}

STDMETHODIMP TabStrip::put_DragScrollTimeBase(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_DRAGSCROLLTIMEBASE);

	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragScrollTimeBase != newValue) {
		properties.dragScrollTimeBase = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_DRAGSCROLLTIMEBASE);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_DropHilitedTab(ITabStripTab** ppDropHilitedTab)
{
	ATLASSERT_POINTER(ppDropHilitedTab, ITabStripTab*);
	if(!ppDropHilitedTab) {
		return E_POINTER;
	}

	ClassFactory::InitTabStripTab(cachedSettings.dropHilitedTab, this, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(ppDropHilitedTab));
	return S_OK;
}

STDMETHODIMP TabStrip::putref_DropHilitedTab(ITabStripTab* pNewDropHilitedTab)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_DROPHILITEDTAB);

	HRESULT hr = E_FAIL;

	int newDropHilitedTab = -1;
	if(pNewDropHilitedTab) {
		LONG l = -1;
		pNewDropHilitedTab->get_Index(&l);
		newDropHilitedTab = l;
	}
	if(cachedSettings.dropHilitedTab == newDropHilitedTab) {
		return S_OK;
	}

	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		if(newDropHilitedTab == -1) {
			if(SendMessage(TCM_HIGHLIGHTITEM, cachedSettings.dropHilitedTab, MAKELPARAM(FALSE, 0))) {
				hr = S_OK;
			}
		} else {
			if(SendMessage(TCM_HIGHLIGHTITEM, newDropHilitedTab, MAKELPARAM(TRUE, 0))) {
				hr = S_OK;
			}
		}
		dragDropStatus.ShowDragImage(TRUE);
	}

	FireOnChanged(DISPID_TABSTRIPCTL_DROPHILITEDTAB);
	return hr;
}

STDMETHODIMP TabStrip::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP TabStrip::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_ENABLED);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_TABSTRIPCTL_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_FixedTabWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.fixedTabWidth;
	return S_OK;
}

STDMETHODIMP TabStrip::put_FixedTabWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_FIXEDTABWIDTH);

	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.fixedTabWidth != newValue) {
		properties.fixedTabWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TCM_SETITEMSIZE, 0, MAKELPARAM(properties.fixedTabWidth, properties.tabHeight));
		}
		FireOnChanged(DISPID_TABSTRIPCTL_FIXEDTABWIDTH);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_FocusOnButtonDown(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.focusOnButtonDown = ((GetStyle() & TCS_FOCUSONBUTTONDOWN) == TCS_FOCUSONBUTTONDOWN);
	}

	*pValue = BOOL2VARIANTBOOL(properties.focusOnButtonDown);
	return S_OK;
}

STDMETHODIMP TabStrip::put_FocusOnButtonDown(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_FOCUSONBUTTONDOWN);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.focusOnButtonDown != b) {
		properties.focusOnButtonDown = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.focusOnButtonDown) {
				ModifyStyle(0, TCS_FOCUSONBUTTONDOWN);
			} else {
				ModifyStyle(TCS_FOCUSONBUTTONDOWN, 0);
			}
		}
		FireOnChanged(DISPID_TABSTRIPCTL_FOCUSONBUTTONDOWN);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP TabStrip::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_FONT);

	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TABSTRIPCTL_FONT);
	return S_OK;
}

STDMETHODIMP TabStrip::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_FONT);

	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TABSTRIPCTL_FONT);
	return S_OK;
}

STDMETHODIMP TabStrip::get_hDragImageList(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(dragDropStatus.IsDragging()) {
		*pValue = HandleToLong(dragDropStatus.hDragImageList);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStrip::get_hHighResImageList(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(cachedSettings.hHighResImageList);
	return S_OK;
}

STDMETHODIMP TabStrip::put_hHighResImageList(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_HHIGHRESIMAGELIST);

	if(cachedSettings.hHighResImageList != reinterpret_cast<HIMAGELIST>(LongToHandle(newValue))) {
		cachedSettings.hHighResImageList = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
		FireOnChanged(DISPID_TABSTRIPCTL_HHIGHRESIMAGELIST);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_hImageList(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(cachedSettings.hImageList);
	} else {
		*pValue = NULL;
	}
	return S_OK;
}

STDMETHODIMP TabStrip::put_hImageList(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_HIMAGELIST);

	if(IsWindow()) {
		SendMessage(TCM_SETIMAGELIST, 0, newValue);
		FireViewChange();
	}

	FireOnChanged(DISPID_TABSTRIPCTL_HIMAGELIST);
	return S_OK;
}

STDMETHODIMP TabStrip::get_HorizontalPadding(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.horizontalPadding;
	return S_OK;
}

STDMETHODIMP TabStrip::put_HorizontalPadding(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_HORIZONTALPADDING);

	if(properties.horizontalPadding != newValue) {
		properties.horizontalPadding = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TCM_SETPADDING, 0, MAKELPARAM(properties.horizontalPadding, properties.verticalPadding));

			// temporarily toggle the TCS_FIXEDWIDTH style, so that SysTabControl32 recalculates the tab metrics
			if(GetStyle() & TCS_FIXEDWIDTH) {
				ModifyStyle(TCS_FIXEDWIDTH, 0);
				ModifyStyle(0, TCS_FIXEDWIDTH);
			} else {
				ModifyStyle(0, TCS_FIXEDWIDTH);
				ModifyStyle(TCS_FIXEDWIDTH, 0);
			}
		}
		FireOnChanged(DISPID_TABSTRIPCTL_HORIZONTALPADDING);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_HotTracking(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.hotTracking = ((GetStyle() & TCS_HOTTRACK) == TCS_HOTTRACK);
	}

	*pValue = BOOL2VARIANTBOOL(properties.hotTracking);
	return S_OK;
}

STDMETHODIMP TabStrip::put_HotTracking(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_HOTTRACKING);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.hotTracking != b) {
		properties.hotTracking = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.hotTracking) {
				ModifyStyle(0, TCS_HOTTRACK);
			} else {
				ModifyStyle(TCS_HOTTRACK, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TABSTRIPCTL_HOTTRACKING);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP TabStrip::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_HOVERTIME);

	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP TabStrip::get_hWndArrowButtons(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(static_cast<HWND>(GetDlgItem(1)));
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_hWndToolTip(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(TCM_GETTOOLTIPS, 0, 0)));
	}
	return S_OK;
}

STDMETHODIMP TabStrip::put_hWndToolTip(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_HWNDTOOLTIP);

	if(IsWindow()) {
		SendMessage(TCM_SETTOOLTIPS, reinterpret_cast<WPARAM>(LongToHandle(newValue)), 0);
	}

	FireOnChanged(DISPID_TABSTRIPCTL_HWNDTOOLTIP);
	return S_OK;
}

STDMETHODIMP TabStrip::get_InsertMarkColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(TCM_GETINSERTMARKCOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.insertMarkColor = 0;
		} else if(color != OLECOLOR2COLORREF(properties.insertMarkColor)) {
			properties.insertMarkColor = color;
		}
	}

	*pValue = properties.insertMarkColor;
	return S_OK;
}

STDMETHODIMP TabStrip::put_InsertMarkColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_INSERTMARKCOLOR);

	if(properties.insertMarkColor != newValue) {
		properties.insertMarkColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TCM_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_TABSTRIPCTL_INSERTMARKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP TabStrip::get_MinTabWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.minTabWidth;
	return S_OK;
}

STDMETHODIMP TabStrip::put_MinTabWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_MINTABWIDTH);

	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.minTabWidth != newValue) {
		properties.minTabWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TCM_SETMINTABWIDTH, 0, properties.minTabWidth);
		}
		FireOnChanged(DISPID_TABSTRIPCTL_MINTABWIDTH);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP TabStrip::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_MOUSEICON);

	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TABSTRIPCTL_MOUSEICON);
	return S_OK;
}

STDMETHODIMP TabStrip::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_MOUSEICON);

	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TABSTRIPCTL_MOUSEICON);
	return S_OK;
}

STDMETHODIMP TabStrip::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP TabStrip::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_MOUSEPOINTER);

	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_MultiRow(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.multiRow = ((GetStyle() & TCS_MULTILINE) == TCS_MULTILINE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.multiRow);
	return S_OK;
}

STDMETHODIMP TabStrip::put_MultiRow(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_MULTIROW);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.multiRow != b) {
		properties.multiRow = b;
		SetDirty(TRUE);

		if(!properties.multiRow) {
			// single line and tpLeft/tpRight are incompatible
			TabPlacementConstants tabPlacement = static_cast<TabPlacementConstants>(0);
			get_TabPlacement(&tabPlacement);
			switch(tabPlacement) {
				case tpLeft:
				case tpRight:
					put_TabPlacement(tpTop);
					break;
			}

			// single line and scroll to opposite are incompatible
			put_ScrollToOpposite(VARIANT_FALSE);
		}
		if(IsWindow()) {
			if(properties.multiRow) {
				ModifyStyle(TCS_SINGLELINE, TCS_MULTILINE);
			} else {
				ModifyStyle(TCS_MULTILINE, TCS_SINGLELINE);
			}
		}
		FireOnChanged(DISPID_TABSTRIPCTL_MULTIROW);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_MultiSelect(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.multiSelect = ((GetStyle() & TCS_MULTISELECT) == TCS_MULTISELECT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.multiSelect);
	return S_OK;
}

STDMETHODIMP TabStrip::put_MultiSelect(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_MULTISELECT);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.multiSelect != b) {
		properties.multiSelect = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.multiSelect) {
				ModifyStyle(0, TCS_MULTISELECT);
			} else {
				ModifyStyle(TCS_MULTISELECT, 0);
			}
		}
		FireOnChanged(DISPID_TABSTRIPCTL_MULTISELECT);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OLEDragImageStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.oleDragImageStyle;
	return S_OK;
}

STDMETHODIMP TabStrip::put_OLEDragImageStyle(OLEDragImageStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_OLEDRAGIMAGESTYLE);
	if(properties.oleDragImageStyle != newValue) {
		properties.oleDragImageStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_OLEDRAGIMAGESTYLE);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_OwnerDrawn(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.ownerDrawn = ((GetStyle() & TCS_OWNERDRAWFIXED) == TCS_OWNERDRAWFIXED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.ownerDrawn);
	return S_OK;
}

STDMETHODIMP TabStrip::put_OwnerDrawn(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_OWNERDRAWN);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.ownerDrawn != b) {
		properties.ownerDrawn = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.ownerDrawn) {
				ModifyStyle(0, TCS_OWNERDRAWFIXED);
			} else {
				ModifyStyle(TCS_OWNERDRAWFIXED, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TABSTRIPCTL_OWNERDRAWN);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP TabStrip::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_PROCESSCONTEXTMENUKEYS);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP TabStrip::get_RaggedTabRows(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.raggedTabRows = ((GetStyle() & TCS_RAGGEDRIGHT) == TCS_RAGGEDRIGHT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.raggedTabRows);
	return S_OK;
}

STDMETHODIMP TabStrip::put_RaggedTabRows(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_RAGGEDTABROWS);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.raggedTabRows != b) {
		properties.raggedTabRows = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.raggedTabRows) {
				ModifyStyle(0, TCS_RAGGEDRIGHT);
			} else {
				ModifyStyle(TCS_RAGGEDRIGHT, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TABSTRIPCTL_RAGGEDTABROWS);
	}
	return S_OK;
}

/*STDMETHODIMP TabStrip::get_ReflectContextMenuMessages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.reflectContextMenuMessages);
	return S_OK;
}

STDMETHODIMP TabStrip::put_ReflectContextMenuMessages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_REFLECTCONTEXTMENUMESSAGES);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.reflectContextMenuMessages != b) {
		properties.reflectContextMenuMessages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_REFLECTCONTEXTMENUMESSAGES);
	}
	return S_OK;
}*/

STDMETHODIMP TabStrip::get_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RegisterForOLEDragDropConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if((SendMessage(TCM_GETEXTENDEDSTYLE, 0, 0) & TCS_EX_REGISTERDROP) == TCS_EX_REGISTERDROP) {
			properties.registerForOLEDragDrop = rfoddNativeDragDrop;
		}
	}

	*pValue = properties.registerForOLEDragDrop;
	return S_OK;
}

STDMETHODIMP TabStrip::put_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_REGISTERFOROLEDRAGDROP);

	if(properties.registerForOLEDragDrop != newValue) {
		properties.registerForOLEDragDrop = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TCM_SETEXTENDEDSTYLE, TCS_EX_REGISTERDROP, 0);
			RevokeDragDrop(*this);
			switch(properties.registerForOLEDragDrop) {
				case rfoddNativeDragDrop:
					SendMessage(TCM_SETEXTENDEDSTYLE, TCS_EX_REGISTERDROP, TCS_EX_REGISTERDROP);
					break;
				case rfoddAdvancedDragDrop: {
					ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
					break;
				}
			}
		}
		FireOnChanged(DISPID_TABSTRIPCTL_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		DWORD style = GetExStyle();
		if(style & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(style & WS_EX_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP TabStrip::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_RIGHTTOLEFT);

	if(properties.rightToLeft != newValue) {
		BOOL changingLayout = ((properties.rightToLeft & rtlLayout) != (newValue & rtlLayout));
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.rightToLeft & rtlLayout) {
				if(properties.rightToLeft & rtlText) {
					ModifyStyleEx(WS_EX_LTRREADING, WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT | WS_EX_RTLREADING);
				} else {
					ModifyStyleEx(WS_EX_RTLREADING, WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT | WS_EX_LTRREADING);
				}
			} else {
				if(properties.rightToLeft & rtlText) {
					ModifyStyleEx(WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT | WS_EX_LTRREADING, WS_EX_RTLREADING);
				} else {
					ModifyStyleEx(WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT | WS_EX_RTLREADING, WS_EX_LTRREADING);
				}
			}
			if(changingLayout) {
				// this will force an update of the up-down control's position
				CRect clientRectangle;
				GetClientRect(&clientRectangle);
				SendMessage(WM_SIZE, SIZE_RESTORED, MAKELPARAM(clientRectangle.Width(), clientRectangle.Height()));
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TABSTRIPCTL_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_ScrollToOpposite(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.scrollToOpposite = ((GetStyle() & TCS_SCROLLOPPOSITE) == TCS_SCROLLOPPOSITE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.scrollToOpposite);
	return S_OK;
}

STDMETHODIMP TabStrip::put_ScrollToOpposite(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_SCROLLTOOPPOSITE);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.scrollToOpposite != b) {
		if(IsWindow()) {
			if(IsInDesignMode()) {
				//RecreateControlWindow();
				MessageBox(TEXT("The control window won't apply this setting until the Form it is placed on is closed and reloaded."), TEXT("TabStrip - ScrollToOpposite"), MB_ICONINFORMATION);
			} else {
				// Set not supported at runtime - raise VB runtime error 382
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 382);
			}
		}
		properties.scrollToOpposite = b;
		SetDirty(TRUE);

		if(properties.scrollToOpposite) {
			// single line and scroll to opposite are incompatible
			put_MultiRow(VARIANT_TRUE);
			// buttons and scroll to opposite are incompatible
			put_Style(sTabs);
		}
		FireOnChanged(DISPID_TABSTRIPCTL_SCROLLTOOPPOSITE);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_ShowButtonSeparators(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showButtonSeparators = ((SendMessage(TCM_GETEXTENDEDSTYLE, 0, 0) & TCS_EX_FLATSEPARATORS) == TCS_EX_FLATSEPARATORS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showButtonSeparators);
	return S_OK;
}

STDMETHODIMP TabStrip::put_ShowButtonSeparators(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_SHOWBUTTONSEPARATORS);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showButtonSeparators != b) {
		properties.showButtonSeparators = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.showButtonSeparators) {
				SendMessage(TCM_SETEXTENDEDSTYLE, TCS_EX_FLATSEPARATORS, TCS_EX_FLATSEPARATORS);
			} else {
				SendMessage(TCM_SETEXTENDEDSTYLE, TCS_EX_FLATSEPARATORS, 0);
			}
		}
		FireOnChanged(DISPID_TABSTRIPCTL_SHOWBUTTONSEPARATORS);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_ShowDragImage(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!dragDropStatus.IsDragging()) {
		return E_FAIL;
	}

	*pValue =  BOOL2VARIANTBOOL(dragDropStatus.IsDragImageVisible());
	return S_OK;
}

STDMETHODIMP TabStrip::put_ShowDragImage(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_SHOWDRAGIMAGE);

	if(!dragDropStatus.hDragImageList && !dragDropStatus.pDropTargetHelper) {
		return E_FAIL;
	}

	if(newValue == VARIANT_FALSE) {
		dragDropStatus.HideDragImage(FALSE);
	} else {
		dragDropStatus.ShowDragImage(FALSE);
	}

	FireOnChanged(DISPID_TABSTRIPCTL_SHOWDRAGIMAGE);
	return S_OK;
}

STDMETHODIMP TabStrip::get_ShowToolTips(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showToolTips = ((GetStyle() & TCS_TOOLTIPS) == TCS_TOOLTIPS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showToolTips);
	return S_OK;
}

STDMETHODIMP TabStrip::put_ShowToolTips(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_SHOWTOOLTIPS);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showToolTips != b) {
		if(IsWindow()) {
			if(IsInDesignMode()) {
				//RecreateControlWindow();
			} else {
				// Set not supported at runtime - raise VB runtime error 382
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 382);
			}
		}
		properties.showToolTips = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_SHOWTOOLTIPS);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_Style(StyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, StyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetStyle();
		if(style & TCS_FLATBUTTONS) {
			properties.style = sFlatButtons;
		} else if(style & TCS_BUTTONS) {
			properties.style = sButtons;
		} else {
			properties.style = sTabs;
		}
	}

	*pValue = properties.style;
	return S_OK;
}

STDMETHODIMP TabStrip::put_Style(StyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_STYLE);

	if(properties.style != newValue) {
		properties.style = newValue;
		SetDirty(TRUE);

		switch(properties.style) {
			case sButtons:
			case sFlatButtons:
				// buttons and scroll to opposite are incompatible
				put_ScrollToOpposite(VARIANT_FALSE);
				break;
		}
		if(IsWindow()) {
			switch(properties.style) {
				case sTabs:
					ModifyStyle(TCS_BUTTONS | TCS_FLATBUTTONS, TCS_TABS);
					break;
				case sButtons:
					ModifyStyle(TCS_FLATBUTTONS | TCS_TABS, TCS_BUTTONS);
					break;
				case sFlatButtons:
					ModifyStyle(TCS_TABS, TCS_BUTTONS | TCS_FLATBUTTONS);
					break;
			}
		}
		FireOnChanged(DISPID_TABSTRIPCTL_STYLE);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP TabStrip::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_SUPPORTOLEDRAGIMAGES);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_TabBoundingBoxDefinition(TabBoundingBoxDefinitionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TabBoundingBoxDefinitionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.tabBoundingBoxDefinition;
	return S_OK;
}

STDMETHODIMP TabStrip::put_TabBoundingBoxDefinition(TabBoundingBoxDefinitionConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_TABBOUNDINGBOXDEFINITION);

	if(properties.tabBoundingBoxDefinition != newValue) {
		properties.tabBoundingBoxDefinition = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_TABBOUNDINGBOXDEFINITION);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_TabCaptionAlignment(TabCaptionAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TabCaptionAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetStyle();
		if(style & TCS_FORCELABELLEFT) {
			properties.tabCaptionAlignment = tcaForceCaptionLeft;
		} else if(style & TCS_FORCEICONLEFT) {
			properties.tabCaptionAlignment = tcaForceIconLeft;
		} else {
			properties.tabCaptionAlignment = tcaNormal;
		}
	}

	*pValue = properties.tabCaptionAlignment;
	return S_OK;
}

STDMETHODIMP TabStrip::put_TabCaptionAlignment(TabCaptionAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_TABCAPTIONALIGNMENT);

	if(properties.tabCaptionAlignment != newValue) {
		properties.tabCaptionAlignment = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.tabCaptionAlignment) {
				case tcaNormal:
					ModifyStyle(TCS_FORCEICONLEFT | TCS_FORCELABELLEFT, 0);
					break;
				case tcaForceIconLeft:
					ModifyStyle(TCS_FORCELABELLEFT, TCS_FORCEICONLEFT);
					break;
				case tcaForceCaptionLeft:
					ModifyStyle(TCS_FORCEICONLEFT, TCS_FORCELABELLEFT);
					break;
			}
		}
		FireOnChanged(DISPID_TABSTRIPCTL_TABCAPTIONALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_TabHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.tabHeight;
	return S_OK;
}

STDMETHODIMP TabStrip::put_TabHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_TABHEIGHT);

	if(properties.tabHeight != newValue) {
		properties.tabHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TCM_SETITEMSIZE, 0, MAKELPARAM(properties.fixedTabWidth, properties.tabHeight));
		}
		FireOnChanged(DISPID_TABSTRIPCTL_TABHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_TabPlacement(TabPlacementConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TabPlacementConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetStyle();
		if(style & TCS_VERTICAL) {
			if(style & TCS_RIGHT) {
				properties.tabPlacement = tpRight;
			} else {
				properties.tabPlacement = tpLeft;
			}
		} else {
			if(style & TCS_BOTTOM) {
				properties.tabPlacement = tpBottom;
			} else {
				properties.tabPlacement = tpTop;
			}
		}
	}

	*pValue = properties.tabPlacement;
	return S_OK;
}

STDMETHODIMP TabStrip::put_TabPlacement(TabPlacementConstants newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_TABPLACEMENT);

	if(properties.tabPlacement != newValue) {
		if(IsWindow()) {
			switch(newValue) {
				case tpTop:
					if(GetStyle() & TCS_VERTICAL) {
						if(IsInDesignMode()) {
							//RecreateControlWindow();
							MessageBox(TEXT("The control window won't apply this setting until the Form it is placed on is closed and reloaded."), TEXT("TabStrip - TabPlacement"), MB_ICONINFORMATION);
						} else {
							// Set not supported at runtime - raise VB runtime error 382
							return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 382);
						}
					} else {
						ModifyStyle(TCS_BOTTOM | TCS_RIGHT | TCS_VERTICAL, 0);
					}
					break;
				case tpBottom:
					if(GetStyle() & TCS_VERTICAL) {
						if(IsInDesignMode()) {
							//RecreateControlWindow();
							MessageBox(TEXT("The control window won't apply this setting until the Form it is placed on is closed and reloaded."), TEXT("TabStrip - TabPlacement"), MB_ICONINFORMATION);
						} else {
							// Set not supported at runtime - raise VB runtime error 382
							return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 382);
						}
					} else {
						ModifyStyle(TCS_RIGHT | TCS_VERTICAL, TCS_BOTTOM);
					}
					break;
				case tpLeft:
				case tpRight:
					if(IsInDesignMode()) {
						//RecreateControlWindow();
						MessageBox(TEXT("The control window won't apply this setting until the Form it is placed on is closed and reloaded."), TEXT("TabStrip - TabPlacement"), MB_ICONINFORMATION);
					} else {
						// Set not supported at runtime - raise VB runtime error 382
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 382);
					}
					break;
			}
		}
		properties.tabPlacement = newValue;
		SetDirty(TRUE);

		switch(properties.tabPlacement) {
			case tpLeft:
			case tpRight:
				// single line and tpLeft/tpRight are incompatible
				put_MultiRow(VARIANT_TRUE);
				break;
		}
		FireOnChanged(DISPID_TABSTRIPCTL_TABPLACEMENT);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_Tabs(ITabStripTabs** ppTabs)
{
	ATLASSERT_POINTER(ppTabs, ITabStripTabs*);
	if(!ppTabs) {
		return E_POINTER;
	}

	ClassFactory::InitTabStripTabs(this, IID_ITabStripTabs, reinterpret_cast<LPUNKNOWN*>(ppTabs));
	return S_OK;
}

STDMETHODIMP TabStrip::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP TabStrip::get_UseFixedTabWidth(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.useFixedTabWidth = ((GetStyle() & TCS_FIXEDWIDTH) == TCS_FIXEDWIDTH);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useFixedTabWidth);
	return S_OK;
}

STDMETHODIMP TabStrip::put_UseFixedTabWidth(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_USEFIXEDTABWIDTH);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useFixedTabWidth != b) {
		properties.useFixedTabWidth = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useFixedTabWidth) {
				ModifyStyle(0, TCS_FIXEDWIDTH);
			} else {
				ModifyStyle(TCS_FIXEDWIDTH, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TABSTRIPCTL_USEFIXEDTABWIDTH);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP TabStrip::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_USESYSTEMFONT);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_TABSTRIPCTL_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP TabStrip::get_VerticalPadding(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.verticalPadding;
	return S_OK;
}

STDMETHODIMP TabStrip::put_VerticalPadding(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TABSTRIPCTL_VERTICALPADDING);

	if(properties.verticalPadding != newValue) {
		properties.verticalPadding = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TCM_SETPADDING, 0, MAKELPARAM(properties.horizontalPadding, properties.verticalPadding));

			// temporarily toggle the TCS_FIXEDWIDTH style, so that SysTabControl32 recalculates the tab metrics
			if(GetStyle() & TCS_FIXEDWIDTH) {
				ModifyStyle(TCS_FIXEDWIDTH, 0);
				ModifyStyle(0, TCS_FIXEDWIDTH);
			} else {
				ModifyStyle(0, TCS_FIXEDWIDTH);
				ModifyStyle(TCS_FIXEDWIDTH, 0);
			}
		}
		FireOnChanged(DISPID_TABSTRIPCTL_VERTICALPADDING);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP TabStrip::BeginDrag(ITabStripTabContainer* pDraggedTabs, OLE_HANDLE hDragImageList/* = NULL*/, OLE_XPOS_PIXELS* pXHotSpot/* = NULL*/, OLE_YPOS_PIXELS* pYHotSpot/* = NULL*/)
{
	ATLASSUME(pDraggedTabs);
	if(!pDraggedTabs) {
		return E_POINTER;
	}

	int xHotSpot = 0;
	if(pXHotSpot) {
		xHotSpot = *pXHotSpot;
	}
	int yHotSpot = 0;
	if(pYHotSpot) {
		yHotSpot = *pYHotSpot;
	}
	HRESULT hr = dragDropStatus.BeginDrag(*this, pDraggedTabs, static_cast<HIMAGELIST>(LongToHandle(hDragImageList)), &xHotSpot, &yHotSpot);
	SetCapture();
	if(pXHotSpot) {
		*pXHotSpot = xHotSpot;
	}
	if(pYHotSpot) {
		*pYHotSpot = yHotSpot;
	}

	if(dragDropStatus.hDragImageList) {
		ImageList_BeginDrag(dragDropStatus.hDragImageList, 0, xHotSpot, yHotSpot);
		dragDropStatus.dragImageIsHidden = 0;
		ImageList_DragEnter(0, 0, 0);

		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ImageList_DragMove(mousePosition.x, mousePosition.y);
	}
	return hr;
}

STDMETHODIMP TabStrip::CalculateDisplayArea(RECTANGLE* pWindowRectangle, RECTANGLE* pDisplayArea)
{
	ATLASSERT_POINTER(pWindowRectangle, RECTANGLE);
	if(!pWindowRectangle) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pDisplayArea, RECTANGLE);
	if(!pDisplayArea) {
		return E_POINTER;
	}

	if(IsWindow()) {
		RECT rc = {pWindowRectangle->Left, pWindowRectangle->Top, pWindowRectangle->Right, pWindowRectangle->Bottom};
		SendMessage(TCM_ADJUSTRECT, FALSE, reinterpret_cast<LPARAM>(&rc));
		pDisplayArea->Left = rc.left;
		pDisplayArea->Top = rc.top;
		pDisplayArea->Right = rc.right;
		pDisplayArea->Bottom = rc.bottom;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStrip::CalculateWindowRectangle(RECTANGLE* pDisplayArea, RECTANGLE* pWindowRectangle)
{
	ATLASSERT_POINTER(pDisplayArea, RECTANGLE);
	if(!pDisplayArea) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pWindowRectangle, RECTANGLE);
	if(!pWindowRectangle) {
		return E_POINTER;
	}

	if(IsWindow()) {
		RECT rc = {pDisplayArea->Left, pDisplayArea->Top, pDisplayArea->Right, pDisplayArea->Bottom};
		SendMessage(TCM_ADJUSTRECT, TRUE, reinterpret_cast<LPARAM>(&rc));
		pWindowRectangle->Left = rc.left;
		pWindowRectangle->Top = rc.top;
		pWindowRectangle->Right = rc.right;
		pWindowRectangle->Bottom = rc.bottom;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStrip::CountTabRows(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = SendMessage(TCM_GETROWCOUNT, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStrip::CreateTabContainer(VARIANT tabs/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, ITabStripTabContainer** ppContainer/* = NULL*/)
{
	ATLASSERT_POINTER(ppContainer, ITabStripTabContainer*);
	if(!ppContainer) {
		return E_POINTER;
	}

	*ppContainer = NULL;
	CComObject<TabStripTabContainer>* pTSTabContainerObj = NULL;
	CComObject<TabStripTabContainer>::CreateInstance(&pTSTabContainerObj);
	pTSTabContainerObj->AddRef();

	// clone all settings
	pTSTabContainerObj->SetOwner(this);

	pTSTabContainerObj->QueryInterface(__uuidof(ITabStripTabContainer), reinterpret_cast<LPVOID*>(ppContainer));
	pTSTabContainerObj->Release();

	if(*ppContainer) {
		(*ppContainer)->Add(tabs);
		RegisterTabContainer(static_cast<ITabContainer*>(pTSTabContainerObj));
	}
	return S_OK;
}

STDMETHODIMP TabStrip::EndDrag(VARIANT_BOOL abort)
{
	if(!dragDropStatus.IsDragging()) {
		return E_FAIL;
	}

	KillTimer(timers.ID_DRAGSCROLL);
	KillTimer(timers.ID_DRAGACTIVATE);
	ReleaseCapture();
	if(dragDropStatus.hDragImageList) {
		dragDropStatus.HideDragImage(TRUE);
		ImageList_EndDrag();
	}

	HRESULT hr = S_OK;
	if(abort) {
		hr = Raise_AbortedDrag();
	} else {
		hr = Raise_Drop();
	}

	dragDropStatus.EndDrag();
	Invalidate();

	return hr;
}

STDMETHODIMP TabStrip::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP TabStrip::GetClosestInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, ITabStripTab** ppTSTab)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppTSTab, ITabStripTab*);
	if(!ppTSTab) {
		return E_POINTER;
	}

	BOOL vertical = ((GetStyle() & TCS_VERTICAL) == TCS_VERTICAL);
	int proposedTabIndex = -1;
	InsertMarkPositionConstants proposedRelativePosition = impNowhere;

	POINT pt = {x, y};
	int numberOfTabs = SendMessage(TCM_GETITEMCOUNT, 0, 0);
	int previousCenter = -100000;
	for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
		CRect tabBoundingRectangle;
		SendMessage(TCM_GETITEMRECT, tabIndex, reinterpret_cast<LPARAM>(&tabBoundingRectangle));
		if(vertical) {
			BOOL isFirstInRow = !((tabBoundingRectangle.left <= previousCenter) && (previousCenter <= tabBoundingRectangle.right));
			previousCenter = tabBoundingRectangle.CenterPoint().x;
			if((pt.x >= tabBoundingRectangle.left) && (pt.x <= tabBoundingRectangle.right)) {
				if(isFirstInRow && (pt.y < tabBoundingRectangle.top)) {
					proposedTabIndex = tabIndex;
					proposedRelativePosition = impBefore;
				} else if(pt.y >= tabBoundingRectangle.top) {
					proposedTabIndex = tabIndex;
					if(pt.y < tabBoundingRectangle.CenterPoint().y) {
						proposedRelativePosition = impBefore;
					} else {
						proposedRelativePosition = impAfter;
					}
				}
			}
		} else {
			BOOL isFirstInRow = !((tabBoundingRectangle.top <= previousCenter) && (previousCenter <= tabBoundingRectangle.bottom));
			previousCenter = tabBoundingRectangle.CenterPoint().y;
			if((pt.y >= tabBoundingRectangle.top) && (pt.y <= tabBoundingRectangle.bottom)) {
				if(isFirstInRow && (pt.x < tabBoundingRectangle.left)) {
					proposedTabIndex = tabIndex;
					proposedRelativePosition = impBefore;
				} else if(pt.x >= tabBoundingRectangle.left) {
					proposedTabIndex = tabIndex;
					if(pt.x < tabBoundingRectangle.CenterPoint().x) {
						proposedRelativePosition = impBefore;
					} else {
						proposedRelativePosition = impAfter;
					}
				}
			}
		}
	}

	if(proposedTabIndex == -1) {
		*ppTSTab = NULL;
		*pRelativePosition = impNowhere;
	} else {
		*pRelativePosition = proposedRelativePosition;
		ClassFactory::InitTabStripTab(proposedTabIndex, this, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(ppTSTab), FALSE);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, ITabStripTab** ppTSTab)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppTSTab, ITabStripTab*);
	if(!ppTSTab) {
		return E_POINTER;
	}

	if(insertMark.tabIndex != -1) {
		if(insertMark.afterTab) {
			*pRelativePosition = impAfter;
		} else {
			*pRelativePosition = impBefore;
		}
		ClassFactory::InitTabStripTab(insertMark.tabIndex, this, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(ppTSTab));
	} else {
		*ppTSTab = NULL;
		*pRelativePosition = impNowhere;
	}

	return S_OK;
}

STDMETHODIMP TabStrip::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, ITabStripTab** ppHitTab)
{
	ATLASSERT_POINTER(ppHitTab, ITabStripTab*);
	if(!ppHitTab) {
		return E_POINTER;
	}

	if(IsWindow()) {
		UINT hitTestFlags = static_cast<UINT>(*pHitTestDetails);
		int tabIndex = HitTest(x, y, &hitTestFlags, TRUE);

		if(pHitTestDetails) {
			*pHitTestDetails = static_cast<HitTestConstants>(hitTestFlags);
		}
		ClassFactory::InitTabStripTab(tabIndex, this, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(ppHitTab));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TabStrip::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP TabStrip::OLEDrag(LONG* pDataObject/* = NULL*/, OLEDropEffectConstants supportedEffects/* = odeCopyOrMove*/, OLE_HANDLE hWndToAskForDragImage/* = -1*/, ITabStripTabContainer* pDraggedTabs/* = NULL*/, LONG itemCountToDisplay/* = -1*/, OLEDropEffectConstants* pPerformedEffects/* = NULL*/)
{
	if(supportedEffects == odeNone) {
		// don't waste time
		return S_OK;
	}
	if(hWndToAskForDragImage == -1) {
		ATLASSUME(pDraggedTabs);
		if(!pDraggedTabs) {
			return E_INVALIDARG;
		}
	}
	ATLASSERT_POINTER(pDataObject, LONG);
	LONG dummy = NULL;
	if(!pDataObject) {
		pDataObject = &dummy;
	}
	ATLASSERT_POINTER(pPerformedEffects, OLEDropEffectConstants);
	OLEDropEffectConstants performedEffect = odeNone;
	if(!pPerformedEffects) {
		pPerformedEffects = &performedEffect;
	}

	HWND hWnd = NULL;
	if(hWndToAskForDragImage == -1) {
		hWnd = *this;
	} else {
		hWnd = static_cast<HWND>(LongToHandle(hWndToAskForDragImage));
	}

	CComPtr<IOLEDataObject> pOLEDataObject = NULL;
	CComPtr<IDataObject> pDataObjectToUse = NULL;
	if(*pDataObject) {
		pDataObjectToUse = *reinterpret_cast<LPDATAOBJECT*>(pDataObject);

		CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->Attach(pDataObjectToUse);
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->Release();
	} else {
		CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->SetOwner(this);
		if(itemCountToDisplay == -1) {
			if(pDraggedTabs) {
				if(flags.usingThemes && RunTimeHelper::IsVista()) {
					pDraggedTabs->Count(&pOLEDataObjectObj->properties.numberOfItemsToDisplay);
				}
			}
		} else if(itemCountToDisplay >= 0) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				pOLEDataObjectObj->properties.numberOfItemsToDisplay = itemCountToDisplay;
			}
		}
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&pDataObjectToUse));
		pOLEDataObjectObj->Release();
	}
	ATLASSUME(pDataObjectToUse);
	pDataObjectToUse->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));
	CComQIPtr<IDropSource, &IID_IDropSource> pDragSource(this);

	if(pDraggedTabs) {
		pDraggedTabs->Clone(&dragDropStatus.pDraggedTabs);
	}
	POINT mousePosition = {0};
	GetCursorPos(&mousePosition);
	ScreenToClient(&mousePosition);

	if(properties.supportOLEDragImages) {
		IDragSourceHelper* pDragSourceHelper = NULL;
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pDragSourceHelper));
		if(pDragSourceHelper) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				IDragSourceHelper2* pDragSourceHelper2 = NULL;
				pDragSourceHelper->QueryInterface(IID_IDragSourceHelper2, reinterpret_cast<LPVOID*>(&pDragSourceHelper2));
				if(pDragSourceHelper2) {
					pDragSourceHelper2->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT);
					// this was the only place we actually use IDragSourceHelper2
					pDragSourceHelper->Release();
					pDragSourceHelper = static_cast<IDragSourceHelper*>(pDragSourceHelper2);
				}
			}

			HRESULT hr = pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
			if(FAILED(hr)) {
				/* This happens if full window dragging is deactivated. Actually, InitializeFromWindow() contains a
				   fallback mechanism for this case. This mechanism retrieves the passed window's class name and
				   builds the drag image using TVM_CREATEDRAGIMAGE if it's SysTreeView32, LVM_CREATEDRAGIMAGE if
				   it's SysListView32 and so on. Our class name is TabStrip[U|A], so we're doomed.
				   So how can we have drag images anyway? Well, we use a very ugly hack: We temporarily activate
				   full window dragging. */
				BOOL fullWindowDragging;
				SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &fullWindowDragging, 0);
				if(!fullWindowDragging) {
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, TRUE, NULL, 0);
					pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, FALSE, NULL, 0);
				}
			}

			if(pDragSourceHelper) {
				pDragSourceHelper->Release();
			}
		}
	}

	if(IsLeftMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_LBUTTON;
	} else if(IsRightMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_RBUTTON;
	}
	if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
		dragDropStatus.useItemCountLabelHack = TRUE;
	}

	if(pOLEDataObject) {
		Raise_OLEStartDrag(pOLEDataObject);
	}
	HRESULT hr = DoDragDrop(pDataObjectToUse, pDragSource, supportedEffects, reinterpret_cast<LPDWORD>(pPerformedEffects));
	KillTimer(timers.ID_DRAGSCROLL);
	KillTimer(timers.ID_DRAGACTIVATE);
	if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
		Raise_OLECompleteDrag(pOLEDataObject, *pPerformedEffects);
	}

	dragDropStatus.Reset();
	return S_OK;
}

STDMETHODIMP TabStrip::Refresh(void)
{
	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		Invalidate();
		UpdateWindow();
		dragDropStatus.ShowDragImage(TRUE);
	}
	return S_OK;
}

STDMETHODIMP TabStrip::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP TabStrip::SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, ITabStripTab* pTSTab)
{
	int tabIndex = -1;
	if(pTSTab) {
		LONG l = NULL;
		pTSTab->get_Index(&l);
		tabIndex = l;
	}

	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		switch(relativePosition) {
			case impNowhere:
				if(SendMessage(TCM_SETINSERTMARK, FALSE, -1)) {
					hr = S_OK;
				}
				break;
			case impBefore:
				if(SendMessage(TCM_SETINSERTMARK, FALSE, tabIndex)) {
					hr = S_OK;
				}
				break;
			case impAfter:
				if(SendMessage(TCM_SETINSERTMARK, TRUE, tabIndex)) {
					hr = S_OK;
				}
				break;
			case impDontChange:
				hr = S_OK;
				break;
		}
		dragDropStatus.ShowDragImage(TRUE);
	}

	return hr;
}


LRESULT TabStrip::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT TabStrip::OnContextMenu(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(reinterpret_cast<HWND>(wParam) == *this) {
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if((mousePosition.x != -1) && (mousePosition.y != -1)) {
			ScreenToClient(&mousePosition);
		}

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);
		return 0;
	/*} else if(properties.reflectContextMenuMessages) {
		return ::SendMessage(reinterpret_cast<HWND>(wParam), OCM__BASE + message, wParam, lParam);*/
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT TabStrip::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		/* SysTabControl32 already sent a WM_NOTIFYFORMAT/NF_QUERY message to the parent window, but
		   unfortunately our reflection object did not yet know where to reflect this message to. So we ask
		   SysTabControl32 to send the message again. */
		SendMessage(WM_NOTIFYFORMAT, reinterpret_cast<WPARAM>(reinterpret_cast<LPCREATESTRUCT>(lParam)->hwndParent), NF_REQUERY);

		Raise_RecreatedControlWindow(HandleToLong(static_cast<HWND>(*this)));
	}
	return lr;
}

LRESULT TabStrip::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT TabStrip::OnHScroll(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(properties.closeableTabs) {
		// redraw around the arrow buttons
		CWindow arrowButtons = GetDlgItem(1);
		CRect rc;
		arrowButtons.GetWindowRect(&rc);
		rc.InflateRect(5, 5);
		ScreenToClient(&rc);
		InvalidateRect(&rc);
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT TabStrip::OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT TabStrip::OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT TabStrip::OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	// SysTabControl32 sends NM_CLICK on each WM_LBUTTONUP
	mouseStatus.ignoreNM_CLICK = TRUE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_DblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TabStrip::OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!dragDropStatus.IsDragging() && properties.closeableTabs) {
		UINT hitTestDetails = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		int tabIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
		if(hitTestDetails & TCHT_CLOSEBUTTON) {
			mouseStatus.overCloseButton = tabIndex;
			mouseStatus.overCloseButtonOnMouseDown = tabIndex;
		} else {
			// signal the HitTest method that we didn't hit a close button - will be reset to -1 on Click event
			mouseStatus.overCloseButtonOnMouseDown = -2;
		}
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(flags.silentSelectionChanges == 0) {
		if(!(properties.disabledEvents & deTabSelectionChanged)) {
			if(GetAsyncKeyState(VK_CONTROL) & 0x8000) {
				DWORD style = GetStyle();
				if(((style & (TCS_FLATBUTTONS | TCS_BUTTONS)) != 0) && ((style & TCS_MULTISELECT) == TCS_MULTISELECT)) {
					UINT dummy = 0;
					int tabIndex = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &dummy);
					if(tabIndex != -1) {
						Raise_TabSelectionChanged(tabIndex);
					}
				}
			}
		}
	}

	if(properties.allowDragDrop) {
		UINT dummy = 0;
		int tabIndex = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &dummy);
		if(tabIndex != -1) {
			dragDropStatus.candidate.tabIndex = tabIndex;
			dragDropStatus.candidate.button = MK_LBUTTON;
			dragDropStatus.candidate.position.x = GET_X_LPARAM(lParam);
			dragDropStatus.candidate.position.y = GET_Y_LPARAM(lParam);
			/* NOTE: We should call SetCapture() here, but this won't work, because SysTabControl32 already
			         captures the mouse in buttons mode. */
		}
	}

	if(!dragDropStatus.IsDragging() && properties.closeableTabs) {
		UINT hitTestDetails = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		int tabIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
		if(hitTestDetails & TCHT_CLOSEBUTTON) {
			mouseStatus.overCloseButton = tabIndex;
			RECT closeButtonRectangle = {0};
			if(CalculateCloseButtonRectangle(tabIndex, tabIndex == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
				InvalidateRect(&closeButtonRectangle);
			}
		}
	}

	return lr;
}

LRESULT TabStrip::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.allowDragDrop) {
		dragDropStatus.candidate.tabIndex = -1;
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	if(!dragDropStatus.IsDragging() && properties.closeableTabs) {
		UINT hitTestDetails = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		int tabIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
		if(hitTestDetails & TCHT_CLOSEBUTTON) {
			mouseStatus.overCloseButton = tabIndex;
			RECT closeButtonRectangle = {0};
			if(CalculateCloseButtonRectangle(tabIndex, tabIndex == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
				InvalidateRect(&closeButtonRectangle);
			}
		}
	}

	return 0;
}

LRESULT TabStrip::OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TabStrip::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT TabStrip::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TabStrip::OnMouseActivate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	/* Hide the message from CComControl, so that the control doesn't receive the focus if the user activates
	   another tab. */
	return DefWindowProc(message, wParam, lParam);
}

LRESULT TabStrip::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TabStrip::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	/* NOTE: We check for mouseStatus.enteredControl because SysTabControl32 sends a WM_MOUSELEAVE after
	         each WM_MOUSEWHEEL message if it has the focus. */
	if(!(properties.disabledEvents & deMouseEvents) && mouseStatus.enteredControl) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);
	}

	if(!dragDropStatus.IsDragging() && properties.closeableTabs) {
		if(mouseStatus.overCloseButton != -1) {
			RECT closeButtonRectangle = {0};
			if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
				InvalidateRect(&closeButtonRectangle);
			}
			mouseStatus.overCloseButton = -1;
		}
	}

	return 0;
}

LRESULT TabStrip::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	if(properties.allowDragDrop) {
		if(dragDropStatus.candidate.tabIndex != -1) {
			// calculate the rectangle, that the mouse cursor must leave to start dragging
			int clickRectWidth = GetSystemMetrics(SM_CXDRAG);
			if(clickRectWidth < 4) {
				clickRectWidth = 4;
			}
			int clickRectHeight = GetSystemMetrics(SM_CYDRAG);
			if(clickRectHeight < 4) {
				clickRectHeight = 4;
			}
			CRect rc(dragDropStatus.candidate.position.x - clickRectWidth, dragDropStatus.candidate.position.y - clickRectHeight, dragDropStatus.candidate.position.x + clickRectWidth, dragDropStatus.candidate.position.y + clickRectHeight);

			if(!rc.PtInRect(CPoint(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)))) {
				NMTCBEGINDRAG details = {0};
				details.hdr.idFrom = GetDlgCtrlID();
				details.hdr.hwndFrom = *this;
				details.iItem = dragDropStatus.candidate.tabIndex;
				details.ptDrag = dragDropStatus.candidate.position;
				switch(dragDropStatus.candidate.button) {
					case MK_LBUTTON:
						details.hdr.code = TCN_BEGINDRAG;
						GetParent().SendMessage(WM_NOTIFY, details.hdr.idFrom, reinterpret_cast<LPARAM>(&details));
						break;
					case MK_RBUTTON:
						details.hdr.code = TCN_BEGINRDRAG;
						GetParent().SendMessage(WM_NOTIFY, details.hdr.idFrom, reinterpret_cast<LPARAM>(&details));
						break;
				}
				dragDropStatus.candidate.tabIndex = -1;
			}
		}
	}

	if(dragDropStatus.IsDragging()) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_DragMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(properties.closeableTabs) {
		UINT hitTestDetails = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		int tabIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
		if(hitTestDetails & TCHT_CLOSEBUTTON) {
			if(mouseStatus.overCloseButton == -1) {
				mouseStatus.overCloseButton = tabIndex;
				RECT closeButtonRectangle = {0};
				if(CalculateCloseButtonRectangle(tabIndex, tabIndex == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
					InvalidateRect(&closeButtonRectangle);
				}
			}
		} else if(mouseStatus.overCloseButton != -1) {
			RECT closeButtonRectangle = {0};
			if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
				InvalidateRect(&closeButtonRectangle);
			}
			mouseStatus.overCloseButton = -1;
		}
	}

	return 0;
}

LRESULT TabStrip::OnNCHitTest(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	return HTCLIENT;
}

LRESULT TabStrip::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(wParam && (!RunTimeHelper::IsCommCtrl6() || !flags.usingThemes || (RunTimeHelper::IsCommCtrl6() && (GetStyle() & TCS_OWNERDRAWFIXED) == TCS_OWNERDRAWFIXED))) {
		// Without this code, we get a black background as soon as our WindowlessLabel control is one of the contained controls.
		RECT clientRectangle = {0};
		GetClientRect(&clientRectangle);
		FillRect(reinterpret_cast<HDC>(wParam), &clientRectangle, GetSysColorBrush(COLOR_3DFACE));
	}
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(message == WM_PAINT) {
		HDC hDC = GetDC();
		DrawCloseButtonsAndInsertionMarks(hDC);
		ReleaseDC(hDC);
	} else {
		DrawCloseButtonsAndInsertionMarks(reinterpret_cast<HDC>(wParam));
	}
	return lr;
}

LRESULT TabStrip::OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	// SysTabControl32 sends NM_RCLICK on each WM_RBUTTONUP
	mouseStatus.ignoreNM_RCLICK = TRUE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TabStrip::OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	if(!(properties.disabledEvents & deTabSelectionChanged)) {
		if(!(GetAsyncKeyState(VK_CONTROL) & 0x8000)) {
			DWORD style = GetStyle();
			if(((style & (TCS_FLATBUTTONS | TCS_BUTTONS)) != 0) && ((style & TCS_MULTISELECT) == TCS_MULTISELECT)) {
				UINT dummy = 0;
				int tabIndex = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &dummy);
				if(tabIndex != -1) {
					// HACK: TCM_GETCURSEL returns -1 after the active tab was changed using the right mouse button
					SendMessage(TCM_SETCURSEL, tabIndex, 0);
				}
			}
		}
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(flags.silentSelectionChanges == 0) {
		if(!(properties.disabledEvents & deTabSelectionChanged)) {
			if(GetAsyncKeyState(VK_CONTROL) & 0x8000) {
				DWORD style = GetStyle();
				if(((style & (TCS_FLATBUTTONS | TCS_BUTTONS)) != 0) && ((style & TCS_MULTISELECT) == TCS_MULTISELECT)) {
					UINT dummy = 0;
					int tabIndex = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &dummy);
					if(tabIndex != -1) {
						Raise_TabSelectionChanged(tabIndex);
					}
				}
			}
		}
	}

	if(properties.allowDragDrop) {
		UINT dummy = 0;
		int tabIndex = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &dummy);
		if(tabIndex != -1) {
			dragDropStatus.candidate.tabIndex = tabIndex;
			dragDropStatus.candidate.button = MK_RBUTTON;
			dragDropStatus.candidate.position.x = GET_X_LPARAM(lParam);
			dragDropStatus.candidate.position.y = GET_Y_LPARAM(lParam);
			/* NOTE: We should call SetCapture() here, but this won't work, because SysTabControl32 already
			         captures the mouse in buttons mode. */
		}
	}

	return lr;
}

LRESULT TabStrip::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.allowDragDrop) {
		dragDropStatus.candidate.tabIndex = -1;
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TabStrip::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
		}
	}

	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT TabStrip::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TABSTRIPCTL_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_TABSTRIPCTL_FONT);
	}

	return lr;
}
//
//LRESULT TabStrip::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
//{
//	if(lParam == 71216) {
//		// the message was sent by ourselves
//		lParam = 0;
//		if(wParam) {
//			// We're gonna activate redrawing - does the client allow this?
//			if(properties.dontRedraw) {
//				// no, so eat this message
//				return 0;
//			}
//		}
//	} else {
//		// TODO: Should we really do this?
//		properties.dontRedraw = !wParam;
//	}
//
//	return DefWindowProc(message, wParam, lParam);
//}

LRESULT TabStrip::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETNONCLIENTMETRICS) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT TabStrip::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT TabStrip::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		//case timers.ID_REDRAW:
		//	if(IsWindowVisible()) {
		//		KillTimer(timers.ID_REDRAW);
		//		SetRedraw(!properties.dontRedraw);
		//	} else {
		//		// wait... (this fixes visibility problems if another control displays a nag screen)
		//	}
		//	break;

		case timers.ID_DRAGSCROLL:
			AutoScroll();
			break;

		case timers.ID_DRAGACTIVATE:
			KillTimer(timers.ID_DRAGACTIVATE);
			if((dragDropStatus.lastDropTarget != -1) && (dragDropStatus.lastDropTarget != SendMessage(TCM_GETCURSEL, 0, 0))) {
				SendMessage(TCM_SETCURSEL, dragDropStatus.lastDropTarget, 0);
			}
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT TabStrip::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT TabStrip::OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_XDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TabStrip::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT TabStrip::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TabStrip::OnReflectedDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPDRAWITEMSTRUCT pDetails = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);

	ATLASSERT(pDetails->CtlType == ODT_TAB);
	ATLASSERT(pDetails->itemAction == ODA_DRAWENTIRE);
	ATLASSERT((pDetails->itemState & (ODS_GRAYED | ODS_DISABLED | ODS_CHECKED | ODS_FOCUS | ODS_DEFAULT | ODS_NOACCEL | ODS_NOFOCUSRECT | ODS_COMBOBOXEDIT | ODS_HOTLIGHT | ODS_INACTIVE)) == 0);

	if(pDetails->hwndItem == *this) {
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(pDetails->itemID, this, FALSE);

		Raise_OwnerDrawTab(pTSTab, static_cast<OwnerDrawTabStateConstants>(pDetails->itemState), HandleToLong(pDetails->hDC), reinterpret_cast<RECTANGLE*>(&pDetails->rcItem));

		return TRUE;
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT TabStrip::OnReflectedNotifyFormat(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == NF_QUERY) {
		#ifdef UNICODE
			return NFR_UNICODE;
		#else
			return NFR_ANSI;
		#endif
	}
	return 0;
}

LRESULT TabStrip::OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	BOOL succeeded = FALSE;
	BOOL useVistaDragImage = FALSE;
	if(dragDropStatus.pDraggedTabs) {
		if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
			succeeded = CreateVistaOLEDragImage(dragDropStatus.pDraggedTabs, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
			useVistaDragImage = succeeded;
		}
		if(!succeeded) {
			// use a legacy drag image as fallback
			succeeded = CreateLegacyOLEDragImage(dragDropStatus.pDraggedTabs, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
		}

		if(succeeded && RunTimeHelper::IsVista()) {
			FORMATETC format = {0};
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			STGMEDIUM medium = {0};
			medium.tymed = TYMED_HGLOBAL;
			medium.hGlobal = GlobalAlloc(GPTR, sizeof(BOOL));
			if(medium.hGlobal) {
				LPBOOL pUseVistaDragImage = static_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				*pUseVistaDragImage = useVistaDragImage;
				GlobalUnlock(medium.hGlobal);

				dragDropStatus.pSourceDataObject->SetData(&format, &medium, TRUE);
			}
		}
	}

	wasHandled = succeeded;
	// TODO: Why do we have to return FALSE to have the set offset not ignored if a Vista drag image is used?
	return succeeded && !useVistaDragImage;
}

LRESULT TabStrip::OnCreateDragImage(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return reinterpret_cast<LRESULT>(CreateLegacyDragImage(lParam, reinterpret_cast<LPPOINT>(wParam), NULL));
}

LRESULT TabStrip::OnDeleteAllItems(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!(properties.disabledEvents & deTabDeletionEvents)) {
		if(flags.silentTabDeletions == 0) {
			Raise_RemovingTab(NULL, &cancel);
		}
	}
	if(cancel == VARIANT_FALSE) {
		if(!(properties.disabledEvents & deMouseEvents)) {
			if(tabUnderMouse != -1) {
				// we're removing the tab below the mouse cursor
				CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabUnderMouse, this);
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);

				UINT hitTestDetails = 0;
				HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				Raise_TabMouseLeave(pTSTab, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				tabUnderMouse = -1;
			}
		}

		if((GetStyle() & (TCS_FLATBUTTONS | TCS_BUTTONS)) != 0) {
			Raise_CaretChanged(caretChangedStatus.previousCaretTab, -1);
		}
		Raise_ActiveTabChanged(activeTabChangedStatus.previousActiveTab, -1);

		if(flags.silentTabDeletions == 0) {
			if(!(properties.disabledEvents & deFreeTabData)) {
				Raise_FreeTabData(NULL);
			}
		}

		// finally pass the message to the tabstrip
		lr = DefWindowProc(message, wParam, lParam);
		if(lr) {
			mouseStatus.overCloseButton = -1;
			mouseStatus.overCloseButtonOnMouseDown = -1;
			cachedSettings.dropHilitedTab = -1;
			dragDropStatus.lastDropTarget = -1;
			insertMark.tabIndex = -1;
			RebuildAcceleratorTable();
			if(flags.silentTabDeletions == 0) {
				if(!(properties.disabledEvents & deTabDeletionEvents)) {
					Raise_RemovedTab(NULL);
				}
			}

			RemoveTabFromTabContainers(-1);
			#ifdef USE_STL
				tabIDs.clear();
				tabParams.clear();
			#else
				tabIDs.RemoveAll();
				tabParams.RemoveAll();
			#endif
		}
	}

	return lr;
}

LRESULT TabStrip::OnDeleteItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<ITabStripTab> pTabStripTabItf = NULL;
	CComObject<TabStripTab>* pTabStripTabObj = NULL;
	if(!(properties.disabledEvents & deTabDeletionEvents) && (flags.silentTabDeletions == 0)) {
		CComObject<TabStripTab>::CreateInstance(&pTabStripTabObj);
		pTabStripTabObj->AddRef();
		pTabStripTabObj->SetOwner(this);
		pTabStripTabObj->Attach(wParam);
		pTabStripTabObj->QueryInterface(IID_ITabStripTab, reinterpret_cast<LPVOID*>(&pTabStripTabItf));
		pTabStripTabObj->Release();
		Raise_RemovingTab(pTabStripTabItf, &cancel);
	}

	if(cancel == VARIANT_FALSE) {
		CComPtr<IVirtualTabStripTab> pVTabStripTabItf = NULL;
		if(!(properties.disabledEvents & deTabDeletionEvents) && (flags.silentTabDeletions == 0)) {
			CComObject<VirtualTabStripTab>* pVTabStripTabObj = NULL;
			CComObject<VirtualTabStripTab>::CreateInstance(&pVTabStripTabObj);
			pVTabStripTabObj->AddRef();
			pVTabStripTabObj->SetOwner(this);
			if(pTabStripTabObj) {
				pTabStripTabObj->SaveState(pVTabStripTabObj);
			}

			pVTabStripTabObj->QueryInterface(IID_IVirtualTabStripTab, reinterpret_cast<LPVOID*>(&pVTabStripTabItf));
			pVTabStripTabObj->Release();
		}

		if(!(properties.disabledEvents & deMouseEvents)) {
			if(static_cast<int>(wParam) == tabUnderMouse) {
				// we're removing the tab below the mouse cursor
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);

				UINT hitTestDetails = 0;
				HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				Raise_TabMouseLeave(pTabStripTabItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				tabUnderMouse = -1;
			}
		}

		if((GetStyle() & (TCS_FLATBUTTONS | TCS_BUTTONS)) != 0) {
			if(caretChangedStatus.previousCaretTab == static_cast<int>(wParam)) {
				Raise_CaretChanged(caretChangedStatus.previousCaretTab, -1);
			}
		}
		if(activeTabChangedStatus.previousActiveTab == static_cast<int>(wParam)) {
			Raise_ActiveTabChanged(activeTabChangedStatus.previousActiveTab, -1);
		}

		if(flags.silentTabDeletions == 0) {
			if(!(properties.disabledEvents & deFreeTabData)) {
				if(pTabStripTabItf) {
					Raise_FreeTabData(pTabStripTabItf);
				} else {
					CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(wParam, this, FALSE);
					Raise_FreeTabData(pTSTab);
				}
			}
		}

		// finally pass the message to the tabstrip
		LONG tabIDBeingRemoved = TabIndexToID(wParam);
		lr = DefWindowProc(message, wParam, lParam);
		if(lr) {
			if(!dragDropStatus.IsDragging() && properties.closeableTabs) {
				if(mouseStatus.overCloseButtonOnMouseDown == static_cast<int>(wParam)) {
					mouseStatus.overCloseButtonOnMouseDown = -1;
				}
				UINT hitTestDetails = 0;
				POINT mousePosition = {0};
				GetCursorPos(&mousePosition);
				ScreenToClient(&mousePosition);
				int tabIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
				if(hitTestDetails & TCHT_CLOSEBUTTON) {
					if(mouseStatus.overCloseButton != -1) {
						RECT closeButtonRectangle = {0};
						if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
							InvalidateRect(&closeButtonRectangle);
						}
					}
					mouseStatus.overCloseButton = tabIndex;
					if(mouseStatus.overCloseButton != -1) {
						RECT closeButtonRectangle = {0};
						if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
							InvalidateRect(&closeButtonRectangle);
						}
					}
				} else if(mouseStatus.overCloseButton != -1) {
					RECT closeButtonRectangle = {0};
					if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
						InvalidateRect(&closeButtonRectangle);
					}
					mouseStatus.overCloseButton = -1;
				}
			}

			if(flags.silentTabDeletions == 0) {
				if(cachedSettings.dropHilitedTab == static_cast<int>(wParam)) {
					cachedSettings.dropHilitedTab = -1;
				} else if(cachedSettings.dropHilitedTab > static_cast<int>(wParam)) {
					--cachedSettings.dropHilitedTab;
				}
				if(dragDropStatus.lastDropTarget == static_cast<int>(wParam)) {
					dragDropStatus.lastDropTarget = -1;
				}
				if(insertMark.tabIndex == static_cast<int>(wParam)) {
					insertMark.tabIndex = -1;
				} else if(insertMark.tabIndex > static_cast<int>(wParam)) {
					--insertMark.tabIndex;
				}

				RemoveTabFromTabContainers(tabIDBeingRemoved);
				#ifdef USE_STL
					tabIDs.erase(tabIDs.begin() + static_cast<int>(wParam));
					std::list<TabData>::iterator iter2 = tabParams.begin();
					if(iter2 != tabParams.end()) {
						std::advance(iter2, wParam);
						if(iter2 != tabParams.end()) {
							tabParams.erase(iter2);
						}
					}
				#else
					tabIDs.RemoveAt(static_cast<int>(wParam));
					POSITION p = tabParams.FindIndex(wParam);
					if(p) {
						tabParams.RemoveAt(p);
					}
				#endif
			}

			RebuildAcceleratorTable();
			if(flags.silentTabDeletions == 0) {
				if(!(properties.disabledEvents & deTabDeletionEvents)) {
					Raise_RemovedTab(pVTabStripTabItf);
				}
			}
		}

		if(!(properties.disabledEvents & deMouseEvents)) {
			if(lr) {
				// maybe we have a new tab below the mouse cursor now
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);

				UINT hitTestDetails = 0;
				int newTabUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				if(newTabUnderMouse != tabUnderMouse) {
					SHORT button = 0;
					SHORT shift = 0;
					WPARAM2BUTTONANDSHIFT(-1, button, shift);
					if(tabUnderMouse != -1) {
						pTabStripTabItf = ClassFactory::InitTabStripTab(tabUnderMouse, this);
						Raise_TabMouseLeave(pTabStripTabItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}

					tabUnderMouse = newTabUnderMouse;

					if(tabUnderMouse != -1) {
						pTabStripTabItf = ClassFactory::InitTabStripTab(tabUnderMouse, this);
						Raise_TabMouseEnter(pTabStripTabItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
				}
			}
		}
	}

	return lr;
}

LRESULT TabStrip::OnDeselectAll(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(properties.disabledEvents & deTabSelectionChanged) {
		return DefWindowProc(message, wParam, lParam);
	} else {
		// store the selection states before processing the message
		#ifdef USE_STL
			std::vector<DWORD> selectionStatesBefore;
		#else
			CAtlArray<DWORD> selectionStatesBefore;
		#endif
		TCITEM tab = {0};
		tab.dwStateMask = TCIS_BUTTONPRESSED;
		tab.mask = TCIF_STATE;
		int numberOfTabs = SendMessage(TCM_GETITEMCOUNT, 0, 0);
		if(flags.silentSelectionChanges == 0) {
			for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
				SendMessage(TCM_GETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab));
				#ifdef USE_STL
					selectionStatesBefore.push_back(tab.dwState & TCIS_BUTTONPRESSED);
				#else
					selectionStatesBefore.Add(tab.dwState & TCIS_BUTTONPRESSED);
				#endif
			}
		}

		// let SysTabControl32 handle the message
		LRESULT lr = DefWindowProc(message, wParam, lParam);

		if(flags.silentSelectionChanges == 0) {
			// now retrieve the selection states again and raise the event if necessary
			for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
				SendMessage(TCM_GETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab));
				if((tab.dwState & TCIS_BUTTONPRESSED) != selectionStatesBefore[tabIndex]) {
					Raise_TabSelectionChanged(tabIndex);
				}
			}
		}

		return lr;
	}
}

LRESULT TabStrip::OnGetExtendedStyle(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(properties.allowDragDrop) {
		lr |= TCS_EX_DETECTDRAGDROP;
	}
	return lr;
}

LRESULT TabStrip::OnGetInsertMark(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	ATLASSERT(lParam);
	if(lParam == 0) {
		return FALSE;
	}

	LPTCINSERTMARK pDetails = reinterpret_cast<LPTCINSERTMARK>(lParam);
	if(pDetails->size == sizeof(TCINSERTMARK)) {
		pDetails->tabIndex = insertMark.tabIndex;
		pDetails->afterTab = insertMark.afterTab;
		return TRUE;
	}
	return FALSE;
}

LRESULT TabStrip::OnGetInsertMarkColor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	return insertMark.color;
}

LRESULT TabStrip::OnGetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL overwriteLParam = FALSE;

	LPTCITEM pTabData = reinterpret_cast<LPTCITEM>(lParam);
	#ifdef DEBUG
		if(message == TCM_GETITEMA) {
			ATLASSERT_POINTER(pTabData, TCITEMA);
		} else {
			ATLASSERT_POINTER(pTabData, TCITEMW);
		}
	#endif
	if(!pTabData) {
		return DefWindowProc(message, wParam, lParam);
	}

	if(pTabData->mask & TCIF_NOINTERCEPTION) {
		// SysTabControl32 shouldn't see this flag
		pTabData->mask &= ~TCIF_NOINTERCEPTION;
	} else if(pTabData->mask == TCIF_PARAM) {
		// just look up in the 'tabParams' list
		#ifdef USE_STL
			std::list<TabData>::iterator iter = tabParams.begin();
			if(iter != tabParams.end()) {
				std::advance(iter, wParam);
				if(iter != tabParams.end()) {
					pTabData->lParam = iter->lParam;
					return TRUE;
				}
			}
		#else
			POSITION p = tabParams.FindIndex(wParam);
			if(p) {
				pTabData->lParam = tabParams.GetAt(p).lParam;
				return TRUE;
			}
		#endif
		// tab not found
		return FALSE;
	} else if(pTabData->mask & TCIF_PARAM) {
		overwriteLParam = TRUE;
	}

	// forward the message
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(overwriteLParam) {
		#ifdef USE_STL
			std::list<TabData>::iterator iter = tabParams.begin();
			if(iter != tabParams.end()) {
				std::advance(iter, wParam);
				if(iter != tabParams.end()) {
					pTabData->lParam = iter->lParam;
				}
			}
		#else
			POSITION p = tabParams.FindIndex(wParam);
			if(p) {
				pTabData->lParam = tabParams.GetAt(p).lParam;
			}
		#endif
	}

	return lr;
}

LRESULT TabStrip::OnHighlightItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		if(LOWORD(lParam)) {
			// a tab got highlighted
			if(static_cast<int>(wParam) != cachedSettings.dropHilitedTab) {
				DefWindowProc(TCM_HIGHLIGHTITEM, cachedSettings.dropHilitedTab, MAKELPARAM(FALSE, 0));
				cachedSettings.dropHilitedTab = static_cast<int>(wParam);
			}
		} else {
			// a tab got un-highlighted
			if(static_cast<int>(wParam) == cachedSettings.dropHilitedTab) {
				cachedSettings.dropHilitedTab = -1;
			}
		}
	}
	return lr;
}

LRESULT TabStrip::OnInsertItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	int insertedTab = -1;

	if(!(properties.disabledEvents & deTabInsertionEvents)) {
		CComObject<VirtualTabStripTab>* pVTabStripTabObj = NULL;
		CComPtr<IVirtualTabStripTab> pVTabStripTabItf = NULL;
		VARIANT_BOOL cancel = VARIANT_FALSE;
		if(flags.silentTabInsertions == 0) {
			CComObject<VirtualTabStripTab>::CreateInstance(&pVTabStripTabObj);
			pVTabStripTabObj->AddRef();
			pVTabStripTabObj->SetOwner(this);

			#ifdef UNICODE
				BOOL mustConvert = (message == TCM_INSERTITEMA);
			#else
				BOOL mustConvert = (message == TCM_INSERTITEMW);
			#endif
			if(mustConvert) {
				#ifdef UNICODE
					LPTCITEMA pTabData = reinterpret_cast<LPTCITEMA>(lParam);
					TCITEMW convertedTabData = {0};
					CA2W converter(pTabData->pszText);
					convertedTabData.pszText = converter;
				#else
					LPTCITEMW pTabData = reinterpret_cast<LPTCITEMW>(lParam);
					TCITEMA convertedTabData = {0};
					CW2A converter(pTabData->pszText);
					convertedTabData.pszText = converter;
				#endif
				convertedTabData.cchTextMax = pTabData->cchTextMax;
				convertedTabData.mask = pTabData->mask;
				convertedTabData.dwState = pTabData->dwState;
				convertedTabData.dwStateMask = pTabData->dwStateMask;
				convertedTabData.iImage = pTabData->iImage;
				convertedTabData.lParam = pTabData->lParam;
				pVTabStripTabObj->LoadState(&convertedTabData, wParam);
			} else {
				pVTabStripTabObj->LoadState(reinterpret_cast<LPTCITEM>(lParam), wParam);
			}

			pVTabStripTabObj->QueryInterface(IID_IVirtualTabStripTab, reinterpret_cast<LPVOID*>(&pVTabStripTabItf));
			pVTabStripTabObj->Release();

			Raise_InsertingTab(pVTabStripTabItf, &cancel);
			pVTabStripTabObj = NULL;
		}
		if(cancel == VARIANT_FALSE) {
			// finally pass the message to the tabstrip
			LPTCITEM pDetails = reinterpret_cast<LPTCITEM>(lParam);
			TabData tabData = {0};
			tabData.isCloseable = TRUE;
			pDetails->mask &= ~TCIF_NOINTERCEPTION;
			if(flags.silentTabInsertions == 0) {
				tabData.lParam = pDetails->lParam;
				pDetails->lParam = GetNewTabID();
				pDetails->mask |= TCIF_PARAM;
			}
			insertedTab = DefWindowProc(message, wParam, lParam);
			if(insertedTab != -1) {
				if(!dragDropStatus.IsDragging() && properties.closeableTabs) {
					mouseStatus.overCloseButtonOnMouseDown = -1;
					UINT hitTestDetails = 0;
					POINT mousePosition = {0};
					GetCursorPos(&mousePosition);
					ScreenToClient(&mousePosition);
					int tabIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
					if(hitTestDetails & TCHT_CLOSEBUTTON) {
						if(mouseStatus.overCloseButton != -1) {
							RECT closeButtonRectangle = {0};
							if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
								InvalidateRect(&closeButtonRectangle);
							}
						}
						mouseStatus.overCloseButton = tabIndex;
						if(mouseStatus.overCloseButton != -1) {
							RECT closeButtonRectangle = {0};
							if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
								InvalidateRect(&closeButtonRectangle);
							}
						}
					} else if(mouseStatus.overCloseButton != -1) {
						RECT closeButtonRectangle = {0};
						if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
							InvalidateRect(&closeButtonRectangle);
						}
						mouseStatus.overCloseButton = -1;
					}
				}

				if(flags.silentTabInsertions == 0) {
					/* NOTE: Activating silent insertions implies that the 'tabIDs' vector and the 'tabParams' list
					         have to be updated explicitly. */
					#ifdef USE_STL
						if(insertedTab >= static_cast<int>(tabIDs.size())) {
							tabIDs.push_back(static_cast<LONG>(pDetails->lParam));
						} else {
							tabIDs.insert(tabIDs.begin() + insertedTab, static_cast<LONG>(pDetails->lParam));
						}
						std::list<TabData>::iterator iter2 = tabParams.begin();
						if(iter2 != tabParams.end()) {
							std::advance(iter2, insertedTab);
						}
						tabParams.insert(iter2, tabData);
					#else
						if(insertedTab >= static_cast<int>(tabIDs.GetCount())) {
							tabIDs.Add(static_cast<LONG>(pDetails->lParam));
						} else {
							tabIDs.InsertAt(insertedTab, static_cast<LONG>(pDetails->lParam));
						}
						POSITION p = tabParams.FindIndex(insertedTab);
						if(p) {
							tabParams.InsertBefore(p, tabData);
						} else {
							tabParams.AddTail(tabData);
						}
					#endif
				}

				if(insertedTab <= cachedSettings.dropHilitedTab) {
					++cachedSettings.dropHilitedTab;
				}
				if(insertedTab <= insertMark.tabIndex) {
					++insertMark.tabIndex;
				}
				RebuildAcceleratorTable();

				if(flags.silentTabInsertions == 0) {
					CComPtr<ITabStripTab> pTabStripTabItf = ClassFactory::InitTabStripTab(insertedTab, this, FALSE);
					Raise_InsertedTab(pTabStripTabItf);
				}
			}
		}
	} else {
		// finally pass the message to the tabstrip
		LPTCITEM pDetails = reinterpret_cast<LPTCITEM>(lParam);
		TabData tabData = {0};
		tabData.isCloseable = TRUE;
		pDetails->mask &= ~TCIF_NOINTERCEPTION;
		if(flags.silentTabInsertions == 0) {
			tabData.lParam = pDetails->lParam;
			pDetails->lParam = GetNewTabID();
			pDetails->mask |= TCIF_PARAM;
		}
		insertedTab = DefWindowProc(message, wParam, lParam);
		if(insertedTab != -1) {
			if(!dragDropStatus.IsDragging() && properties.closeableTabs) {
				mouseStatus.overCloseButtonOnMouseDown = -1;
				UINT hitTestDetails = 0;
				POINT mousePosition = {0};
				GetCursorPos(&mousePosition);
				ScreenToClient(&mousePosition);
				int tabIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
				if(hitTestDetails & TCHT_CLOSEBUTTON) {
					if(mouseStatus.overCloseButton != -1) {
						RECT closeButtonRectangle = {0};
						if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
							InvalidateRect(&closeButtonRectangle);
						}
					}
					mouseStatus.overCloseButton = tabIndex;
					if(mouseStatus.overCloseButton != -1) {
						RECT closeButtonRectangle = {0};
						if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
							InvalidateRect(&closeButtonRectangle);
						}
					}
				} else if(mouseStatus.overCloseButton != -1) {
					RECT closeButtonRectangle = {0};
					if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, mouseStatus.overCloseButton == SendMessage(TCM_GETCURSEL, 0, 0), &closeButtonRectangle)) {
						InvalidateRect(&closeButtonRectangle);
					}
					mouseStatus.overCloseButton = -1;
				}
			}

			if(flags.silentTabInsertions == 0) {
				/* NOTE: Activating silent insertions implies that the 'tabIDs' vector and the 'tabParams' list
					        have to be updated explicitly. */
				#ifdef USE_STL
					if(insertedTab >= static_cast<int>(tabIDs.size())) {
						tabIDs.push_back(static_cast<LONG>(pDetails->lParam));
					} else {
						tabIDs.insert(tabIDs.begin() + insertedTab, static_cast<LONG>(pDetails->lParam));
					}
					std::list<TabData>::iterator iter2 = tabParams.begin();
					if(iter2 != tabParams.end()) {
						std::advance(iter2, insertedTab);
					}
					tabParams.insert(iter2, tabData);
				#else
					if(insertedTab >= static_cast<int>(tabIDs.GetCount())) {
						tabIDs.Add(static_cast<LONG>(pDetails->lParam));
					} else {
						tabIDs.InsertAt(insertedTab, static_cast<LONG>(pDetails->lParam));
					}
					POSITION p = tabParams.FindIndex(insertedTab);
					if(p) {
						tabParams.InsertBefore(p, tabData);
					} else {
						tabParams.AddTail(tabData);
					}
				#endif
			}

			if(insertedTab <= cachedSettings.dropHilitedTab) {
				++cachedSettings.dropHilitedTab;
			}
			if(insertedTab <= insertMark.tabIndex) {
				++insertMark.tabIndex;
			}
			RebuildAcceleratorTable();
		}
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		if(insertedTab != -1) {
			// maybe we have a new tab below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);

			UINT hitTestDetails = 0;
			int newTabUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			if(newTabUnderMouse != tabUnderMouse) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(tabUnderMouse != -1) {
					CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabUnderMouse, this);
					Raise_TabMouseLeave(pTSTab, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				tabUnderMouse = newTabUnderMouse;

				if(tabUnderMouse != -1) {
					CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabUnderMouse, this);
					Raise_TabMouseEnter(pTSTab, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}
			}
		}
	}

	// HACK: we've somehow to initialize caretChangedStatus and activeTabChangedStatus
	if((GetStyle() & (TCS_FLATBUTTONS | TCS_BUTTONS)) != 0) {
		int caretTab = SendMessage(TCM_GETCURFOCUS, 0, 0);
		if(caretTab != caretChangedStatus.previousCaretTab) {
			Raise_CaretChanged(caretChangedStatus.previousCaretTab, caretTab);
		}
	}
	int activeTab = SendMessage(TCM_GETCURSEL, 0, 0);
	if(activeTab != activeTabChangedStatus.previousActiveTab) {
		Raise_ActiveTabChanged(activeTabChangedStatus.previousActiveTab, activeTab);
	}

	return insertedTab;
}

LRESULT TabStrip::OnSetCurFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if((GetStyle() & (TCS_FLATBUTTONS | TCS_BUTTONS)) != 0) {
		Raise_CaretChanged(caretChangedStatus.previousCaretTab, wParam);
	}
	return lr;
}

LRESULT TabStrip::OnSetCurSel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	VARIANT_BOOL cancel = VARIANT_FALSE;
	Raise_ActiveTabChanging(activeTabChangedStatus.previousActiveTab, &cancel);
	if(cancel == VARIANT_FALSE) {
		LRESULT lr = DefWindowProc(message, wParam, lParam);
		if(lr != -1) {
			Raise_ActiveTabChanged(activeTabChangedStatus.previousActiveTab, wParam);
		}
		return lr;
	}
	return -1;
}

LRESULT TabStrip::OnSetExtendedStyle(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam & ~TCS_EX_DETECTDRAGDROP, lParam & ~TCS_EX_DETECTDRAGDROP);
	if(properties.allowDragDrop) {
		lr |= TCS_EX_DETECTDRAGDROP;
	}
	if((wParam == 0) || ((wParam & TCS_EX_DETECTDRAGDROP) == TCS_EX_DETECTDRAGDROP)) {
		properties.allowDragDrop = ((lParam & TCS_EX_DETECTDRAGDROP) == TCS_EX_DETECTDRAGDROP);
	}
	return lr;
}

LRESULT TabStrip::OnSetImageList(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TABSTRIPCTL_HIMAGELIST) == S_FALSE) {
		return 0;
	}

	/* We must store the new settings before the call to DefWindowProc, because we may need them to
	   handle the WM_PAINT message this message probably will cause. */
	cachedSettings.hImageList = reinterpret_cast<HIMAGELIST>(lParam);

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	// TODO: Isn't there a better way for correcting the values if an error occured?
	cachedSettings.hImageList = reinterpret_cast<HIMAGELIST>(SendMessage(TCM_GETIMAGELIST, 0, 0));
	FireOnChanged(DISPID_TABSTRIPCTL_HIMAGELIST);
	return lr;
}

LRESULT TabStrip::OnSetInsertMark(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	RECT oldInsertMarkRect = {0};
	RECT newInsertMarkRect = {0};

	if(insertMark.tabIndex != -1) {
		// calculate the current insertion mark's rectangle
		RECT tabBoundingRectangle = {0};
		SendMessage(TCM_GETITEMRECT, insertMark.tabIndex, reinterpret_cast<LPARAM>(&tabBoundingRectangle));
		if(GetStyle() & TCS_VERTICAL) {
			if(insertMark.afterTab) {
				oldInsertMarkRect.top = tabBoundingRectangle.bottom - 4;
				oldInsertMarkRect.bottom = tabBoundingRectangle.bottom + 4;
			} else {
				oldInsertMarkRect.top = tabBoundingRectangle.top - 4;
				oldInsertMarkRect.bottom = tabBoundingRectangle.top + 4;
			}
			oldInsertMarkRect.left = tabBoundingRectangle.left;
			oldInsertMarkRect.right = tabBoundingRectangle.right;
		} else {
			if(insertMark.afterTab) {
				oldInsertMarkRect.left = tabBoundingRectangle.right - 4;
				oldInsertMarkRect.right = tabBoundingRectangle.right + 4;
			} else {
				oldInsertMarkRect.left = tabBoundingRectangle.left - 4;
				oldInsertMarkRect.right = tabBoundingRectangle.left + 4;
			}
			oldInsertMarkRect.top = tabBoundingRectangle.top;
			oldInsertMarkRect.bottom = tabBoundingRectangle.bottom;
		}
	}

	insertMark.ProcessSetInsertMark(wParam, lParam);

	if(insertMark.tabIndex != -1) {
		// calculate the new insertion mark's rectangle
		RECT tabBoundingRectangle = {0};
		SendMessage(TCM_GETITEMRECT, insertMark.tabIndex, reinterpret_cast<LPARAM>(&tabBoundingRectangle));
		if(GetStyle() & TCS_VERTICAL) {
			if(insertMark.afterTab) {
				newInsertMarkRect.top = tabBoundingRectangle.bottom - 4;
				newInsertMarkRect.bottom = tabBoundingRectangle.bottom + 4;
			} else {
				newInsertMarkRect.top = tabBoundingRectangle.top - 4;
				newInsertMarkRect.bottom = tabBoundingRectangle.top + 4;
			}
			newInsertMarkRect.left = tabBoundingRectangle.left;
			newInsertMarkRect.right = tabBoundingRectangle.right;
		} else {
			if(insertMark.afterTab) {
				newInsertMarkRect.left = tabBoundingRectangle.right - 4;
				newInsertMarkRect.right = tabBoundingRectangle.right + 4;
			} else {
				newInsertMarkRect.left = tabBoundingRectangle.left - 4;
				newInsertMarkRect.right = tabBoundingRectangle.left + 4;
			}
			newInsertMarkRect.top = tabBoundingRectangle.top;
			newInsertMarkRect.bottom = tabBoundingRectangle.bottom;
		}
	}

	// redraw
	if(oldInsertMarkRect.right - oldInsertMarkRect.left > 0) {
		InvalidateRect(&oldInsertMarkRect);
	}
	if(newInsertMarkRect.right - newInsertMarkRect.left > 0) {
		InvalidateRect(&newInsertMarkRect);
	}

	// always succeed
	return TRUE;
}

LRESULT TabStrip::OnSetInsertMarkColor(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TABSTRIPCTL_INSERTMARKCOLOR) == S_FALSE) {
		return 0;
	}

	COLORREF previousColor = insertMark.color;
	insertMark.color = static_cast<COLORREF>(lParam);

	// calculate insertion mark rectangle
	RECT tabBoundingRectangle = {0};
	SendMessage(TCM_GETITEMRECT, insertMark.tabIndex, reinterpret_cast<LPARAM>(&tabBoundingRectangle));
	RECT insertMarkRect = {0};
	if(GetStyle() & TCS_VERTICAL) {
		if(insertMark.afterTab) {
			insertMarkRect.top = tabBoundingRectangle.bottom - 4;
			insertMarkRect.bottom = tabBoundingRectangle.bottom + 4;
		} else {
			insertMarkRect.top = tabBoundingRectangle.top - 4;
			insertMarkRect.bottom = tabBoundingRectangle.top + 4;
		}
		insertMarkRect.left = tabBoundingRectangle.left;
		insertMarkRect.right = tabBoundingRectangle.right;
	} else {
		if(insertMark.afterTab) {
			insertMarkRect.left = tabBoundingRectangle.right - 4;
			insertMarkRect.right = tabBoundingRectangle.right + 4;
		} else {
			insertMarkRect.left = tabBoundingRectangle.left - 4;
			insertMarkRect.right = tabBoundingRectangle.left + 4;
		}
		insertMarkRect.top = tabBoundingRectangle.top;
		insertMarkRect.bottom = tabBoundingRectangle.bottom;
	}

	// redraw this rectangle
	InvalidateRect(&insertMarkRect);

	SetDirty(TRUE);
	FireOnChanged(DISPID_TABSTRIPCTL_INSERTMARKCOLOR);
	return previousColor;
}

LRESULT TabStrip::OnSetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPTCITEM pTabData = reinterpret_cast<LPTCITEM>(lParam);
	#ifdef DEBUG
		if(message == TCM_SETITEMA) {
			ATLASSERT_POINTER(pTabData, TCITEMA);
		} else {
			ATLASSERT_POINTER(pTabData, TCITEMW);
		}
	#endif
	if(!pTabData) {
		return DefWindowProc(message, wParam, lParam);
	}

	if(pTabData->mask & TCIF_NOINTERCEPTION) {
		// SysTabControl32 shouldn't see this flag
		pTabData->mask &= ~TCIF_NOINTERCEPTION;
	} else if(pTabData->mask == LVIF_PARAM) {
		// simply update the 'tabParams' list
		#ifdef USE_STL
			std::list<TabData>::iterator iter = tabParams.begin();
			if(iter != tabParams.end()) {
				std::advance(iter, wParam);
				if(iter != tabParams.end()) {
					iter->lParam = pTabData->lParam;
					return TRUE;
				}
			}
		#else
			POSITION p = tabParams.FindIndex(wParam);
			if(p) {
				tabParams.GetAt(p).lParam = pTabData->lParam;
				return TRUE;
			}
		#endif
		// tab not found
		return FALSE;
	} else if(pTabData->mask & TCIF_PARAM) {
		#ifdef USE_STL
			std::list<TabData>::iterator iter = tabParams.begin();
			if(iter != tabParams.end()) {
				std::advance(iter, wParam);
				if(iter != tabParams.end()) {
					iter->lParam = pTabData->lParam;
					pTabData->mask &= ~TCIF_PARAM;
				}
			}
		#else
			POSITION p = tabParams.FindIndex(wParam);
			if(p) {
				tabParams.GetAt(p).lParam = pTabData->lParam;
				pTabData->mask &= ~TCIF_PARAM;
			}
		#endif
	}

	if(properties.disabledEvents & deTabSelectionChanged) {
		LRESULT lr = DefWindowProc(message, wParam, lParam);
		if(lr) {
			if(pTabData->mask & TCIF_STATE) {
				if(pTabData->dwStateMask & TCIS_HIGHLIGHTED) {
					if(static_cast<int>(wParam) == cachedSettings.dropHilitedTab) {
						if((pTabData->dwState & TCIS_HIGHLIGHTED) == 0) {
							// TCIS_HIGHLIGHTED was removed
							cachedSettings.dropHilitedTab = -1;
						}
					}
				}
			}
			if(pTabData->mask & TCIF_TEXT) {
				RebuildAcceleratorTable();
			}
		}
		return lr;
	} else {
		// retrieve the current selection state
		TCITEM tab = {0};
		if(flags.silentSelectionChanges == 0) {
			tab.dwStateMask = TCIS_BUTTONPRESSED;
			tab.mask = TCIF_STATE;
			SendMessage(TCM_GETITEM, wParam, reinterpret_cast<LPARAM>(&tab));
		}
		DWORD previousState = (tab.dwState & TCIS_BUTTONPRESSED);

		// let SysTabControl32 handle the message
		LRESULT lr = DefWindowProc(message, wParam, lParam);
		if(lr) {
			if(pTabData->mask & TCIF_STATE) {
				if(pTabData->dwStateMask & TCIS_HIGHLIGHTED) {
					if(static_cast<int>(wParam) == cachedSettings.dropHilitedTab) {
						if((pTabData->dwState & TCIS_HIGHLIGHTED) == 0) {
							// TCIS_HIGHLIGHTED was removed
							cachedSettings.dropHilitedTab = -1;
						}
					}
				}
			}
			if(pTabData->mask & TCIF_TEXT) {
				RebuildAcceleratorTable();
			}
		}

		if(flags.silentSelectionChanges == 0) {
			// check for selection change
			SendMessage(TCM_GETITEM, wParam, reinterpret_cast<LPARAM>(&tab));
			if((tab.dwState & TCIS_BUTTONPRESSED) != previousState) {
				Raise_TabSelectionChanged(wParam);
			}
		}
		return lr;
	}
}

LRESULT TabStrip::OnSetItemSize(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TABSTRIPCTL_FIXEDTABWIDTH) == S_FALSE || FireOnRequestEdit(DISPID_TABSTRIPCTL_TABHEIGHT) == S_FALSE) {
		return 0;
	}

	properties.fixedTabWidth = GET_X_LPARAM(lParam);
	properties.tabHeight = GET_Y_LPARAM(lParam);
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	FireOnChanged(DISPID_TABSTRIPCTL_FIXEDTABWIDTH);
	FireOnChanged(DISPID_TABSTRIPCTL_TABHEIGHT);
	return lr;
}

LRESULT TabStrip::OnSetMinTabWidth(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TABSTRIPCTL_MINTABWIDTH) == S_FALSE) {
		return 0;
	}

	properties.minTabWidth = lParam;
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	FireOnChanged(DISPID_TABSTRIPCTL_MINTABWIDTH);
	return lr;
}

LRESULT TabStrip::OnSetPadding(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TABSTRIPCTL_HORIZONTALPADDING) == S_FALSE || FireOnRequestEdit(DISPID_TABSTRIPCTL_VERTICALPADDING) == S_FALSE) {
		return 0;
	}

	properties.horizontalPadding = GET_X_LPARAM(lParam);
	properties.verticalPadding = GET_Y_LPARAM(lParam);
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	FireOnChanged(DISPID_TABSTRIPCTL_HORIZONTALPADDING);
	FireOnChanged(DISPID_TABSTRIPCTL_VERTICALPADDING);
	return lr;
}


LRESULT TabStrip::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTTDISPINFOA);
	if(!pDetails || pDetails->hdr.hwndFrom != reinterpret_cast<HWND>(SendMessage(TCM_GETTOOLTIPS, 0, 0))) {
		return 0;
	}

	ZeroMemory(pToolTipBuffer, MAX_PATH * sizeof(CHAR));

	BSTR tooltip = NULL;
	CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(pDetails->hdr.idFrom, this);
	Raise_TabGetInfoTipText(pTSTab, MAX_PATH, &tooltip);
	if(lstrlenA(CW2A(tooltip)) > 0) {
		lstrcpynA(reinterpret_cast<LPSTR>(pToolTipBuffer), CW2A(tooltip), MAX_PATH);
	}
	pDetails->lpszText = reinterpret_cast<LPSTR>(pToolTipBuffer);

	return 0;
}

LRESULT TabStrip::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTTDISPINFOW);
	if(!pDetails || pDetails->hdr.hwndFrom != reinterpret_cast<HWND>(SendMessage(TCM_GETTOOLTIPS, 0, 0))) {
		return 0;
	}

	ZeroMemory(pToolTipBuffer, MAX_PATH * sizeof(WCHAR));

	BSTR tooltip = NULL;
	CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(pDetails->hdr.idFrom, this);
	Raise_TabGetInfoTipText(pTSTab, MAX_PATH, &tooltip);
	//#ifdef UNICODE
		if(lstrlenW(CW2CW(tooltip)) > 0) {
			lstrcpynW(reinterpret_cast<LPWSTR>(pToolTipBuffer), CW2CW(tooltip), MAX_PATH);
		}
	//#else
		//OSVERSIONINFO ovi = {sizeof(OSVERSIONINFO)};
		//if(GetVersionEx(&ovi) && (ovi.dwPlatformId == VER_PLATFORM_WIN32_WINDOWS)) {
		//	/* HACK: On Windows 9x strange things happen. Instead of the ANSI notification, the Unicode
		//	         notification is sent. This isn't really a problem, but lstrcpynW() seems to be broken. */
		//	LPCWSTR pToolTip = CW2CW(tooltip);
		//	if(lstrlenW(pToolTip) > 0) {
		//		// as lstrcpynW() seems to be broken, copy the string manually
		//		LPWSTR tmp = reinterpret_cast<LPWSTR>(pToolTipBuffer);
		//		for(int i = 0; (*pToolTip != 0) && (i < MAX_PATH); ++i, ++pToolTip, ++tmp) {
		//			*tmp = *pToolTip;
		//		}
		//		*tmp = 0;
		//	}
		//} else {
		//	if(lstrlenW(CW2CW(tooltip)) > 0) {
		//		lstrcpynW(reinterpret_cast<LPWSTR>(pToolTipBuffer), CW2CW(tooltip), MAX_PATH);
		//	}
		//}
	//#endif
	pDetails->lpszText = reinterpret_cast<LPWSTR>(pToolTipBuffer);

	return 0;
}

LRESULT TabStrip::OnClickNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deClickEvents)) {
		// SysTabControl32 sends NM_CLICK on each WM_LBUTTONUP
		if(mouseStatus.ignoreNM_CLICK) {
			mouseStatus.ignoreNM_CLICK = FALSE;
		} else {
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			button |= 1/*MouseButtonConstants.vbLeftButton*/;
			Raise_Click(button, shift, mousePosition.x, mousePosition.y);
		}
	}
	if(properties.closeableTabs && properties.closeableTabsMode == ctmDisplayOnActiveTab && mouseStatus.overCloseButtonOnMouseDown == -2 && !dragDropStatus.IsDragging()) {
		mouseStatus.overCloseButtonOnMouseDown = -1;

		UINT hitTestDetails = 0;
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);
		int tabIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
		if(hitTestDetails & TCHT_CLOSEBUTTON) {
			mouseStatus.overCloseButton = tabIndex;
			RECT closeButtonRectangle = {0};
			if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, TRUE, &closeButtonRectangle)) {
				InvalidateRect(&closeButtonRectangle);
			}
		}
	} else {
		mouseStatus.overCloseButtonOnMouseDown = -1;
	}
	return 0;
}

LRESULT TabStrip::OnRClickNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deClickEvents)) {
		// SysTabControl32 sends NM_RCLICK on each WM_RBUTTONUP
		if(mouseStatus.ignoreNM_RCLICK) {
			mouseStatus.ignoreNM_RCLICK = FALSE;
		} else {
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			button |= 2/*MouseButtonConstants.vbRightButton*/;
			Raise_RClick(button, shift, mousePosition.x, mousePosition.y);
		}
	}
	return 0;
}

LRESULT TabStrip::OnBeginDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTCBEGINDRAG pDetails = reinterpret_cast<LPNMTCBEGINDRAG>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTCBEGINDRAG);
	if(!pDetails) {
		return 0;
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button |= 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = pDetails->ptDrag;
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(pDetails->iItem, this);
	Raise_TabBeginDrag(pTSTab, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));

	return 0;
}

LRESULT TabStrip::OnBeginRDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTCBEGINDRAG pDetails = reinterpret_cast<LPNMTCBEGINDRAG>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTCBEGINDRAG);
	if(!pDetails) {
		return 0;
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button |= 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = pDetails->ptDrag;
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(pDetails->iItem, this);
	Raise_TabBeginRDrag(pTSTab, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));

	return 0;
}

LRESULT TabStrip::OnFocusChangeNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	int tabIndex = SendMessage(TCM_GETCURFOCUS, 0, 0);
	Raise_CaretChanged(caretChangedStatus.previousCaretTab, tabIndex);
	return 0;
}

LRESULT TabStrip::OnGetObjectNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMOBJECTNOTIFY pDetails = reinterpret_cast<LPNMOBJECTNOTIFY>(pNotificationDetails);
	ATLASSERT(*(pDetails->piid) == IID_IDropTarget);
	pDetails->hResult = QueryInterface(*(pDetails->piid), &pDetails->pObject);
	return 0;
}

LRESULT TabStrip::OnSelChangeNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(properties.closeableTabs) {
		// redraw around the arrow buttons
		CWindow arrowButtons = GetDlgItem(1);
		CRect rc;
		arrowButtons.GetWindowRect(&rc);
		rc.InflateRect(5, 5);
		ScreenToClient(&rc);
		InvalidateRect(&rc);
	}

	int tabIndex = SendMessage(TCM_GETCURSEL, 0, 0);
	Raise_ActiveTabChanged(activeTabChangedStatus.previousActiveTab, tabIndex);
	return 0;
}

LRESULT TabStrip::OnSelChangingNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	activeTabChangedStatus.previousActiveTab = SendMessage(TCM_GETCURSEL, 0, 0);
	VARIANT_BOOL cancel = VARIANT_FALSE;
	Raise_ActiveTabChanging(activeTabChangedStatus.previousActiveTab, &cancel);
	return (cancel != VARIANT_FALSE);
}


void TabStrip::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// use the system font
			properties.font.currentFont.Attach(static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
			properties.font.owningFontResource = FALSE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update our font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT TabStrip::Raise_AbortedDrag(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_AbortedDrag();
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_ActiveTabChanged(int previousActiveTab, int newActiveTab)
{
	HWND hWndToShow = GetAssociatedWindow(newActiveTab);
	if(hWndToShow && ::IsWindow(hWndToShow)) {
		::ShowWindow(hWndToShow, SW_SHOW);
	}

	if(newActiveTab == previousActiveTab) {
		return S_OK;
	}

	if(!(GetStyle() & (TCS_FLATBUTTONS | TCS_BUTTONS))) {
		Raise_CaretChanged(previousActiveTab, newActiveTab);
	}

	activeTabChangedStatus.previousActiveTab = newActiveTab;
	HRESULT hr = S_OK;
	if(flags.silentActiveTabChanges == 0) {
		//if(m_nFreezeEvents == 0) {
			CComPtr<ITabStripTab> pPreviousActiveTab = ClassFactory::InitTabStripTab(previousActiveTab, this);
			CComPtr<ITabStripTab> pNewActiveTab = ClassFactory::InitTabStripTab(newActiveTab, this);
			hr = Fire_ActiveTabChanged(pPreviousActiveTab, pNewActiveTab);
		//}
	}

	if(!(properties.disabledEvents & deTabSelectionChanged)) {
		if(flags.silentSelectionChanges == 0) {
			// now retrieve the selection states again and raise the TabSelectionChanged event if necessary
			TCITEM tab = {0};
			tab.dwStateMask = TCIS_BUTTONPRESSED;
			tab.mask = TCIF_STATE;
			int numberOfTabs = SendMessage(TCM_GETITEMCOUNT, 0, 0);
			#ifdef USE_STL
				if(numberOfTabs == static_cast<int>(activeTabChangedStatus.selectionStatesBeforeChange.size())) {
			#else
				if(numberOfTabs == static_cast<int>(activeTabChangedStatus.selectionStatesBeforeChange.GetCount())) {
			#endif
				for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
					if(tabIndex != newActiveTab) {
						SendMessage(TCM_GETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab));
						if((tab.dwState & TCIS_BUTTONPRESSED) != activeTabChangedStatus.selectionStatesBeforeChange[tabIndex]) {
							Raise_TabSelectionChanged(tabIndex);
						}
					}
				}
			}
		}
		#ifdef USE_STL
			activeTabChangedStatus.selectionStatesBeforeChange.clear();
		#else
			activeTabChangedStatus.selectionStatesBeforeChange.RemoveAll();
		#endif

		if(flags.silentSelectionChanges == 0) {
			if(newActiveTab != -1) {
				Raise_TabSelectionChanged(newActiveTab);
			}
		}
	}
	return hr;
}

inline HRESULT TabStrip::Raise_ActiveTabChanging(int previousActiveTab, VARIANT_BOOL* pCancel)
{
	HRESULT hr = S_OK;
	if(flags.silentActiveTabChanges == 0) {
		//if(m_nFreezeEvents == 0) {
			CComPtr<ITabStripTab> pPreviousActiveTab = ClassFactory::InitTabStripTab(previousActiveTab, this);
			hr = Fire_ActiveTabChanging(pPreviousActiveTab, pCancel);
		//}
	}

	if(*pCancel == VARIANT_FALSE) {
		HWND hWndToHide = GetAssociatedWindow(previousActiveTab);
		if(hWndToHide && ::IsWindow(hWndToHide)) {
			::ShowWindow(hWndToHide, SW_HIDE);
		}
		if(!(properties.disabledEvents & deTabSelectionChanged)) {
			// store the selection states before the change
			TCITEM tab = {0};
			tab.dwStateMask = TCIS_BUTTONPRESSED;
			tab.mask = TCIF_STATE;
			int numberOfTabs = SendMessage(TCM_GETITEMCOUNT, 0, 0);
			for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
				SendMessage(TCM_GETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab));
				#ifdef USE_STL
					activeTabChangedStatus.selectionStatesBeforeChange.push_back(tab.dwState & TCIS_BUTTONPRESSED);
				#else
					activeTabChangedStatus.selectionStatesBeforeChange.Add(tab.dwState & TCIS_BUTTONPRESSED);
				#endif
			}
		}
	}
	return hr;
}

inline HRESULT TabStrip::Raise_CaretChanged(int previousCaretTab, int newCaretTab)
{
	if(newCaretTab == previousCaretTab) {
		return S_OK;
	}

	caretChangedStatus.previousCaretTab = newCaretTab;
	if(flags.silentCaretChanges == 0) {
		//if(m_nFreezeEvents == 0) {
			CComPtr<ITabStripTab> pPreviousCaretTab = ClassFactory::InitTabStripTab(previousCaretTab, this);
			CComPtr<ITabStripTab> pNewCaretTab = ClassFactory::InitTabStripTab(newCaretTab, this);
			return Fire_CaretChanged(pPreviousCaretTab, pNewCaretTab);
		//}
	}
	return S_OK;
}

inline HRESULT TabStrip::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedTab = HitTest(x, y, &hitTestDetails);

	if(properties.closeableTabs && properties.closeableTabsMode == ctmDisplayOnActiveTab && mouseStatus.overCloseButtonOnMouseDown == -2 && !dragDropStatus.IsDragging()) {
		mouseStatus.overCloseButtonOnMouseDown = -1;

		UINT hitTestDetails2 = 0;
		int tabIndex = HitTest(x, y, &hitTestDetails2, TRUE);
		if(hitTestDetails2 & TCHT_CLOSEBUTTON) {
			mouseStatus.overCloseButton = tabIndex;
			RECT closeButtonRectangle = {0};
			if(CalculateCloseButtonRectangle(mouseStatus.overCloseButton, TRUE, &closeButtonRectangle)) {
				InvalidateRect(&closeButtonRectangle);
			}
		}
	} else {
		mouseStatus.overCloseButtonOnMouseDown = -1;
	}

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(mouseStatus.lastClickedTab, this);
		return Fire_Click(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
			if(properties.processContextMenuKeys) {
				// retrieve the focused tab and propose its rectangle's middle as the menu's position
				int tabIndex = SendMessage(TCM_GETCURFOCUS, 0, 0);
				if(tabIndex != -1) {
					CRect tabRectangle;
					if(SendMessage(TCM_GETITEMRECT, tabIndex, reinterpret_cast<LPARAM>(&tabRectangle))) {
						CPoint centerPoint = tabRectangle.CenterPoint();
						x = centerPoint.x;
						y = centerPoint.y;
					}
				}
			} else {
				return S_OK;
			}
		}

		UINT hitTestDetails = 0;
		int tabIndex = HitTest(x, y, &hitTestDetails);
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
		return Fire_ContextMenu(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int tabIndex = HitTest(x, y, &hitTestDetails);
	if(tabIndex != mouseStatus.lastClickedTab) {
		tabIndex = -1;
	}
	mouseStatus.lastClickedTab = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
		return Fire_DblClick(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_DestroyedControlWindow(LONG hWnd)
{
	//KillTimer(timers.ID_REDRAW);
	RemoveTabFromTabContainers(-1);
	#ifdef USE_STL
		tabIDs.clear();
		tabParams.clear();
	#else
		tabIDs.RemoveAll();
		tabParams.RemoveAll();
	#endif

	if(properties.registerForOLEDragDrop == rfoddAdvancedDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_DragMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(dragDropStatus.hDragImageList) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ImageList_DragMove(mousePosition.x, mousePosition.y);
	}

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(x, y, &hitTestDetails, TRUE);
	ITabStripTab* pDropTarget = NULL;
	ClassFactory::InitTabStripTab(dropTarget, this, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	VARIANT_BOOL autoActivateTab = VARIANT_TRUE;
	if(dropTarget == -1) {
		autoActivateTab = VARIANT_FALSE;
	}

	LONG autoScrollVelocity = 0;
	BOOL handleAutoScroll = FALSE;
	if(((GetStyle() & TCS_MULTILINE) == 0) && ((SendMessage(TCM_GETEXTENDEDSTYLE, 0, 0) & TCS_EX_REGISTERDROP) == 0) && (properties.dragScrollTimeBase != 0)) {
		handleAutoScroll = TRUE;
		/* Use a 16 pixels wide border at both ends of the tab bar as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(x, y);
		CRect noScrollZone(0, 0, 0, 0);
		GetWindowRect(&noScrollZone);
		ScreenToClient(&noScrollZone);

		BOOL rightToLeft = ((GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
		if(rightToLeft) {
			// noScrollZone is right to left, too, and PtInRect doesn't like that
			LONG buffer = noScrollZone.left;
			noScrollZone.left = noScrollZone.right;
			noScrollZone.right = buffer;
		}

		BOOL isInScrollZone = noScrollZone.PtInRect(mousePos);
		if(isInScrollZone) {
			// we're within the window rectangle, so do further checks
			if(GetStyle() & TCS_BOTTOM) {
				noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, 0, properties.DRAGSCROLLZONEWIDTH, 0);
				isInScrollZone = !(noScrollZone.PtInRect(mousePos) || (mousePos.y < noScrollZone.bottom - properties.tabHeight));
			} else {
				noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, 0, properties.DRAGSCROLLZONEWIDTH, 0);
				isInScrollZone = !(noScrollZone.PtInRect(mousePos) || (mousePos.y > noScrollZone.top + properties.tabHeight));
			}
		}
		if(rightToLeft) {
			// mousePos is right to left, so make noScrollZone right to left again
			LONG buffer = noScrollZone.left;
			noScrollZone.left = noScrollZone.right;
			noScrollZone.right = buffer;
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose a velocity
			if(mousePos.x < noScrollZone.left) {
				autoScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_DragMouseMove(&pDropTarget, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails), &autoActivateTab, &autoScrollVelocity);
	//}

	if(pDropTarget) {
		LONG l = -1;
		pDropTarget->get_Index(&l);
		dropTarget = l;
		// we don't want to produce a mem-leak...
		pDropTarget->Release();
	} else {
		dropTarget = -1;
	}

	if(autoActivateTab == VARIANT_FALSE) {
		// cancel auto-activation
		KillTimer(timers.ID_DRAGACTIVATE);
	} else {
		if(dropTarget != dragDropStatus.lastDropTarget) {
			// cancel auto-activation of previous target
			KillTimer(timers.ID_DRAGACTIVATE);
			if(properties.dragActivateTime != 0) {
				// start timered auto-activation of new target
				SetTimer(timers.ID_DRAGACTIVATE, (properties.dragActivateTime == -1 ? GetDoubleClickTime() : properties.dragActivateTime));
			}
		}
	}
	dragDropStatus.lastDropTarget = dropTarget;

	if(handleAutoScroll) {
		dragDropStatus.autoScrolling.currentScrollVelocity = autoScrollVelocity;

		LONG smallestInterval = abs(autoScrollVelocity);
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT TabStrip::Raise_Drop(void)
{
	//if(m_nFreezeEvents == 0) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);

		UINT hitTestDetails = 0;
		int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
		CComPtr<ITabStripTab> pDropTarget = ClassFactory::InitTabStripTab(dropTarget, this);

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		return Fire_Drop(pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_FreeTabData(ITabStripTab* pTSTab)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FreeTabData(pTSTab);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_InsertedTab(ITabStripTab* pTSTab)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedTab(pTSTab);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_InsertingTab(IVirtualTabStripTab* pTSTab, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingTab(pTSTab, pCancel);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedTab = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(mouseStatus.lastClickedTab, this);
		return Fire_MClick(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int tabIndex = HitTest(x, y, &hitTestDetails);
	if(tabIndex != mouseStatus.lastClickedTab) {
		tabIndex = -1;
	}
	mouseStatus.lastClickedTab = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
		return Fire_MDblClick(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	if(!mouseStatus.hoveredControl) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = *this;
		trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		Raise_MouseHover(button, shift, x, y);
	}
	mouseStatus.StoreClickCandidate(button);

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		int tabIndex = HitTest(x, y, &hitTestDetails);

		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
		return Fire_MouseDown(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	UINT hitTestDetails = 0;
	int tabIndex = HitTest(x, y, &hitTestDetails);
	tabUnderMouse = tabIndex;

	CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		Fire_MouseEnter(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	if(pTSTab) {
		Raise_TabMouseEnter(pTSTab, button, shift, x, y, hitTestDetails);
	}
	return hr;
}

inline HRESULT TabStrip::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.hoveredControl) {
		mouseStatus.HoverControl();

		//if(m_nFreezeEvents == 0) {
			UINT hitTestDetails = 0;
			int tabIndex = HitTest(x, y, &hitTestDetails);

			CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
			return Fire_MouseHover(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		//}
	}
	return S_OK;
}

inline HRESULT TabStrip::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int tabIndex = HitTest(x, y, &hitTestDetails);

	CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabUnderMouse, this);
	if(pTSTab) {
		Raise_TabMouseLeave(pTSTab, button, shift, x, y, hitTestDetails);
	}
	tabUnderMouse = -1;

	mouseStatus.LeaveControl();

	//if(m_nFreezeEvents == 0) {
		pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
		return Fire_MouseLeave(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	UINT hitTestDetails = 0;
	int tabIndex = HitTest(x, y, &hitTestDetails);

	CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
	// Do we move over another tab than before?
	if(tabIndex != tabUnderMouse) {
		CComPtr<ITabStripTab> pPrevTSTab = ClassFactory::InitTabStripTab(tabUnderMouse, this);
		if(pPrevTSTab) {
			Raise_TabMouseLeave(pPrevTSTab, button, shift, x, y, hitTestDetails);
		}
		tabUnderMouse = tabIndex;
		if(pTSTab) {
			Raise_TabMouseEnter(pTSTab, button, shift, x, y, hitTestDetails);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea;
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					/*if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y);
					}*/
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					/*if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}*/
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		int tabIndex = HitTest(x, y, &hitTestDetails);
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
		return Fire_MouseUp(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLECompleteDrag(pData, performedEffect);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);
	KillTimer(timers.ID_DRAGACTIVATE);

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
	CComPtr<ITabStripTab> pDropTarget = ClassFactory::InitTabStripTab(dropTarget, this);

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
		   always be the same as the data object that was passed to Raise_OLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT TabStrip::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	ITabStripTab* pDropTarget = NULL;
	ClassFactory::InitTabStripTab(dropTarget, this, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	VARIANT_BOOL autoActivateTab = VARIANT_TRUE;
	if(dropTarget == -1) {
		autoActivateTab = VARIANT_FALSE;
	}

	LONG autoScrollVelocity = 0;
	BOOL handleAutoScroll = FALSE;
	if(((GetStyle() & TCS_MULTILINE) == 0) && ((SendMessage(TCM_GETEXTENDEDSTYLE, 0, 0) & TCS_EX_REGISTERDROP) == 0) && (properties.dragScrollTimeBase != 0)) {
		handleAutoScroll = TRUE;
		/* Use a 16 pixels wide border at both ends of the tab bar as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(mousePosition.x, mousePosition.y);
		CRect noScrollZone(0, 0, 0, 0);
		GetWindowRect(&noScrollZone);
		ScreenToClient(&noScrollZone);
		BOOL isInScrollZone = noScrollZone.PtInRect(mousePos);
		if(isInScrollZone) {
			// we're within the window rectangle, so do further checks
			if(GetStyle() & TCS_BOTTOM) {
				noScrollZone.DeflateRect(0, properties.DRAGSCROLLZONEWIDTH, 0, properties.DRAGSCROLLZONEWIDTH);
				isInScrollZone = !(noScrollZone.PtInRect(mousePos) || (mousePos.y < noScrollZone.bottom - properties.tabHeight));
			} else {
				noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, 0, properties.DRAGSCROLLZONEWIDTH, 0);
				isInScrollZone = !(noScrollZone.PtInRect(mousePos) || (mousePos.y > noScrollZone.top + properties.tabHeight));
			}
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose a velocity
			if(mousePos.x < noScrollZone.left) {
				autoScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoScrollVelocity = 1;
			}
		}
	}

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails), &autoActivateTab, &autoScrollVelocity);
		}
	//}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}

	if(autoActivateTab == VARIANT_FALSE) {
		// cancel auto-activation
		KillTimer(timers.ID_DRAGACTIVATE);
	} else {
		if(dropTarget != dragDropStatus.lastDropTarget) {
			// cancel auto-activation of previous target
			KillTimer(timers.ID_DRAGACTIVATE);
			if(properties.dragActivateTime != 0) {
				// start timered auto-activation of new target
				SetTimer(timers.ID_DRAGACTIVATE, (properties.dragActivateTime == -1 ? GetDoubleClickTime() : properties.dragActivateTime));
			}
		}
	}
	dragDropStatus.lastDropTarget = dropTarget;

	if(handleAutoScroll) {
		dragDropStatus.autoScrolling.currentScrollVelocity = autoScrollVelocity;

		LONG smallestInterval = abs(autoScrollVelocity);
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT TabStrip::Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragEnterPotentialTarget(hWndPotentialTarget);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_OLEDragLeave(void)
{
	KillTimer(timers.ID_DRAGSCROLL);
	KillTimer(timers.ID_DRAGACTIVATE);

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails);
	CComPtr<ITabStripTab> pDropTarget = ClassFactory::InitTabStripTab(dropTarget, this);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT TabStrip::Raise_OLEDragLeavePotentialTarget(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragLeavePotentialTarget();
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	ITabStripTab* pDropTarget = NULL;
	ClassFactory::InitTabStripTab(dropTarget, this, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	VARIANT_BOOL autoActivateTab = VARIANT_TRUE;
	if(dropTarget == -1) {
		autoActivateTab = VARIANT_FALSE;
	}

	LONG autoScrollVelocity = 0;
	BOOL handleAutoScroll = FALSE;
	if(((GetStyle() & TCS_MULTILINE) == 0) && ((SendMessage(TCM_GETEXTENDEDSTYLE, 0, 0) & TCS_EX_REGISTERDROP) == 0) && (properties.dragScrollTimeBase != 0)) {
		handleAutoScroll = TRUE;
		/* Use a 16 pixels wide border at both ends of the tab bar as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(mousePosition.x, mousePosition.y);
		CRect noScrollZone(0, 0, 0, 0);
		GetWindowRect(&noScrollZone);
		ScreenToClient(&noScrollZone);

		BOOL rightToLeft = ((GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
		if(rightToLeft) {
			// noScrollZone is right to left, too, and PtInRect doesn't like that
			LONG buffer = noScrollZone.left;
			noScrollZone.left = noScrollZone.right;
			noScrollZone.right = buffer;
		}

		BOOL isInScrollZone = noScrollZone.PtInRect(mousePos);
		if(isInScrollZone) {
			// we're within the window rectangle, so do further checks
			if(GetStyle() & TCS_BOTTOM) {
				noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, 0, properties.DRAGSCROLLZONEWIDTH, 0);
				isInScrollZone = !(noScrollZone.PtInRect(mousePos) || (mousePos.y < noScrollZone.bottom - properties.tabHeight));
			} else {
				noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, 0, properties.DRAGSCROLLZONEWIDTH, 0);
				isInScrollZone = !(noScrollZone.PtInRect(mousePos) || (mousePos.y > noScrollZone.top + properties.tabHeight));
			}
		}
		if(rightToLeft) {
			// mousePos is right to left, so make noScrollZone right to left again
			LONG buffer = noScrollZone.left;
			noScrollZone.left = noScrollZone.right;
			noScrollZone.right = buffer;
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose a velocity
			if(mousePos.x < noScrollZone.left) {
				autoScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails), &autoActivateTab, &autoScrollVelocity);
		}
	//}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}

	if(autoActivateTab == VARIANT_FALSE) {
		// cancel auto-activation
		KillTimer(timers.ID_DRAGACTIVATE);
	} else {
		if(dropTarget != dragDropStatus.lastDropTarget) {
			// cancel auto-activation of previous target
			KillTimer(timers.ID_DRAGACTIVATE);
			if(properties.dragActivateTime != 0) {
				// start timered auto-activation of new target
				SetTimer(timers.ID_DRAGACTIVATE, (properties.dragActivateTime == -1 ? GetDoubleClickTime() : properties.dragActivateTime));
			}
		}
	}
	dragDropStatus.lastDropTarget = dropTarget;

	if(handleAutoScroll) {
		dragDropStatus.autoScrolling.currentScrollVelocity = autoScrollVelocity;

		LONG smallestInterval = abs(autoScrollVelocity);
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT TabStrip::Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEGiveFeedback(static_cast<OLEDropEffectConstants>(effect), pUseDefaultCursors);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith)
{
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);
		return Fire_OLEQueryContinueDrag(BOOL2VARIANTBOOL(pressedEscape), button, shift, reinterpret_cast<OLEActionToContinueWithConstants*>(pActionToContinueWith));
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT TabStrip::Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEReceivedNewData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT TabStrip::Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLESetData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_OLEStartDrag(IOLEDataObject* pData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEStartDrag(pData);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_OwnerDrawTab(ITabStripTab* pTSTab, OwnerDrawTabStateConstants tabState, LONG hDC, RECTANGLE* pDrawingRectangle)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OwnerDrawTab(pTSTab, tabState, hDC, pDrawingRectangle);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedTab = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(mouseStatus.lastClickedTab, this);
		return Fire_RClick(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int tabIndex = HitTest(x, y, &hitTestDetails);
	if(tabIndex != mouseStatus.lastClickedTab) {
		tabIndex = -1;
	}
	mouseStatus.lastClickedTab = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
		return Fire_RDblClick(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_RecreatedControlWindow(LONG hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop == rfoddAdvancedDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	/*if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}*/

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_RemovedTab(IVirtualTabStripTab* pTSTab)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovedTab(pTSTab);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_RemovingTab(ITabStripTab* pTSTab, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingTab(pTSTab, pCancel);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_TabBeginDrag(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TabBeginDrag(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_TabBeginRDrag(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TabBeginRDrag(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_TabGetInfoTipText(ITabStripTab* pTSTab, LONG maxInfoTipLength, BSTR* pInfoTipText)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TabGetInfoTipText(pTSTab, maxInfoTipLength, pInfoTipText);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_TabMouseEnter(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_TabMouseEnter(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT TabStrip::Raise_TabMouseLeave(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_TabMouseLeave(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT TabStrip::Raise_TabSelectionChanged(int tabIndex)
{
	//if(m_nFreezeEvents == 0) {
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
		return Fire_TabSelectionChanged(pTSTab);
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedTab = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(mouseStatus.lastClickedTab, this);
		return Fire_XClick(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT TabStrip::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int tabIndex = HitTest(x, y, &hitTestDetails);
	if(tabIndex != mouseStatus.lastClickedTab) {
		tabIndex = -1;
	}
	mouseStatus.lastClickedTab = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITabStripTab> pTSTab = ClassFactory::InitTabStripTab(tabIndex, this);
		return Fire_XDblClick(pTSTab, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}


void TabStrip::EnterSilentTabInsertionSection(void)
{
	++flags.silentTabInsertions;
}

void TabStrip::LeaveSilentTabInsertionSection(void)
{
	--flags.silentTabInsertions;
}

void TabStrip::EnterSilentTabDeletionSection(void)
{
	++flags.silentTabDeletions;
}

void TabStrip::LeaveSilentTabDeletionSection(void)
{
	--flags.silentTabDeletions;
}

void TabStrip::EnterSilentActiveTabChangeSection(void)
{
	++flags.silentActiveTabChanges;
}

void TabStrip::LeaveSilentActiveTabChangeSection(void)
{
	--flags.silentActiveTabChanges;
}

void TabStrip::EnterSilentCaretChangeSection(void)
{
	++flags.silentCaretChanges;
}

void TabStrip::LeaveSilentCaretChangeSection(void)
{
	--flags.silentCaretChanges;
}

void TabStrip::EnterSilentSelectionChangesSection(void)
{
	++flags.silentSelectionChanges;
}

void TabStrip::LeaveSilentSelectionChangesSection(void)
{
	--flags.silentSelectionChanges;
}


void TabStrip::RecreateControlWindow(void)
{
	// This method shouldn't be used, because it will destroy all contained controls.
	ATLASSERT(FALSE);
	/*if(m_bInPlaceActive) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}*/
}

DWORD TabStrip::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD TabStrip::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	if(properties.focusOnButtonDown) {
		style |= TCS_FOCUSONBUTTONDOWN;
	}
	if(properties.hotTracking) {
		style |= TCS_HOTTRACK;
	}
	if(properties.multiRow) {
		style |= TCS_MULTILINE;
	}
	if(properties.multiSelect) {
		style |= TCS_MULTISELECT;
	}
	if(properties.ownerDrawn) {
		style |= TCS_OWNERDRAWFIXED;
	}
	if(properties.raggedTabRows) {
		style |= TCS_RAGGEDRIGHT;
	}
	if(properties.scrollToOpposite) {
		style |= TCS_SCROLLOPPOSITE;
	}
	if(properties.showToolTips) {
		style |= TCS_TOOLTIPS;
	}
	switch(properties.style) {
		case sButtons:
			style |= TCS_BUTTONS;
			break;
		case sFlatButtons:
			style |= TCS_BUTTONS | TCS_FLATBUTTONS;
			break;
	}
	switch(properties.tabCaptionAlignment) {
		case tcaForceIconLeft:
			style |= TCS_FORCEICONLEFT;
			break;
		case tcaForceCaptionLeft:
			style |= TCS_FORCELABELLEFT;
			break;
	}
	switch(properties.tabPlacement) {
		case tpBottom:
			style |= TCS_BOTTOM;
			break;
		case tpLeft:
			style |= TCS_VERTICAL;
			break;
		case tpRight:
			style |= TCS_VERTICAL | TCS_RIGHT;
			break;
	}
	if(properties.useFixedTabWidth) {
		style |= TCS_FIXEDWIDTH;
	}
	return style;
}

void TabStrip::SendConfigurationMessages(void)
{
	DWORD extendedStyle = 0;
	if(properties.allowDragDrop) {
		extendedStyle |= TCS_EX_DETECTDRAGDROP;
	}
	if(properties.registerForOLEDragDrop == rfoddNativeDragDrop) {
		extendedStyle |= TCS_EX_REGISTERDROP;
	}
	if(properties.showButtonSeparators) {
		extendedStyle |= TCS_EX_FLATSEPARATORS;
	}

	SendMessage(TCM_SETEXTENDEDSTYLE, 0, extendedStyle);

	SendMessage(TCM_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));
	SendMessage(TCM_SETITEMSIZE, 0, MAKELPARAM(properties.fixedTabWidth, properties.tabHeight));
	SendMessage(TCM_SETMINTABWIDTH, 0, properties.minTabWidth);
	SendMessage(TCM_SETPADDING, 0, MAKELPARAM(properties.horizontalPadding, properties.verticalPadding));

	ApplyFont();

	if(IsInDesignMode()) {
		// insert some dummy tabs
		TCITEM tab = {0};
		tab.mask = TCIF_TEXT;
		tab.pszText = TEXT("Dummy Tab 1");
		SendMessage(TCM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&tab));
		tab.pszText = TEXT("Dummy Tab 2");
		SendMessage(TCM_INSERTITEM, 1, reinterpret_cast<LPARAM>(&tab));
		tab.pszText = TEXT("Dummy Tab 3");
		SendMessage(TCM_INSERTITEM, 2, reinterpret_cast<LPARAM>(&tab));
		tab.pszText = TEXT("Dummy Tab 4");
		SendMessage(TCM_INSERTITEM, 3, reinterpret_cast<LPARAM>(&tab));
		tab.pszText = TEXT("Dummy Tab 5");
		SendMessage(TCM_INSERTITEM, 4, reinterpret_cast<LPARAM>(&tab));
	}
}

HCURSOR TabStrip::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

int TabStrip::IDToTabIndex(LONG ID)
{
	#ifdef USE_STL
		std::vector<LONG>::iterator iter = std::find(tabIDs.begin(), tabIDs.end(), ID);
		if(iter != tabIDs.end()) {
			return std::distance(tabIDs.begin(), iter);
		}
	#else
		for(size_t i = 0; i < tabIDs.GetCount(); ++i) {
			if(tabIDs[i] == ID) {
				return i;
			}
		}
	#endif
	return -1;
}

LONG TabStrip::TabIndexToID(int tabIndex)
{
	ATLASSERT(IsWindow());

	TCITEM tab = {0};
	tab.mask = TCIF_PARAM | TCIF_NOINTERCEPTION;
	if(SendMessage(TCM_GETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab))) {
		return tab.lParam;
	}
	return -1;
}

LONG TabStrip::GetNewTabID(void)
{
	static LONG lastID = 0;

	return ++lastID;
}

void TabStrip::RegisterTabContainer(ITabContainer* pContainer)
{
	ATLASSUME(pContainer);
	#ifdef _DEBUG
		#ifdef USE_STL
			std::unordered_map<DWORD, ITabContainer*>::iterator iter = tabContainers.find(pContainer->GetID());
			ATLASSERT(iter == tabContainers.end());
		#else
			CAtlMap<DWORD, ITabContainer*>::CPair* pEntry = tabContainers.Lookup(pContainer->GetID());
			ATLASSERT(!pEntry);
		#endif
	#endif
	tabContainers[pContainer->GetID()] = pContainer;
}

void TabStrip::DeregisterTabContainer(DWORD containerID)
{
	#ifdef USE_STL
		std::unordered_map<DWORD, ITabContainer*>::iterator iter = tabContainers.find(containerID);
		ATLASSERT(iter != tabContainers.end());
		if(iter != tabContainers.end()) {
			tabContainers.erase(iter);
		}
	#else
		tabContainers.RemoveKey(containerID);
	#endif
}

void TabStrip::RemoveTabFromTabContainers(LONG tabID)
{
	#ifdef USE_STL
		for(std::unordered_map<DWORD, ITabContainer*>::const_iterator iter = tabContainers.begin(); iter != tabContainers.end(); ++iter) {
			iter->second->RemovedTab(tabID);
		}
	#else
		POSITION p = tabContainers.GetStartPosition();
		while(p) {
			tabContainers.GetValueAt(p)->RemovedTab(tabID);
			tabContainers.GetNextValue(p);
		}
	#endif
}

int TabStrip::HitTest(LONG x, LONG y, UINT* pFlags, BOOL ignoreBoundingBoxDefinition/* = FALSE*/)
{
	ATLASSERT(IsWindow());

	UINT hitTestFlags = 0;
	if(pFlags) {
		hitTestFlags = *pFlags;
	}
	TCHITTESTINFO hitTestInfo = {{x, y}, hitTestFlags };
	int tabIndex = SendMessage(TCM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));

	if(tabIndex == -1) {
		// Are we outside the tab headers?
		CRect windowRectangle;
		GetWindowRect(&windowRectangle);
		ScreenToClient(&windowRectangle);
		if(windowRectangle.PtInRect(hitTestInfo.pt)) {
			// Are we inside the client area?
			CRect clientRectangle(&windowRectangle);
			SendMessage(TCM_ADJUSTRECT, FALSE, reinterpret_cast<LPARAM>(&clientRectangle));
			if(clientRectangle.PtInRect(hitTestInfo.pt)) {
				// bingo
				hitTestInfo.flags |= TCHT_CLIENTAREA;
			} else {
				// we're probably in the blank area next to a tab
			}
		} else {
			if(x < windowRectangle.left) {
				hitTestInfo.flags |= TCHT_TOLEFT;
			} else if(x > windowRectangle.right) {
				hitTestInfo.flags |= TCHT_TORIGHT;
			}
			if(y < windowRectangle.top) {
				hitTestInfo.flags |= TCHT_ABOVE;
			} else if(y > windowRectangle.bottom) {
				hitTestInfo.flags |= TCHT_BELOW;
			}
		}
	} else if(properties.closeableTabs) {
		/* When the Click event is raised, the clicked tab already is the new caret tab. So if the tab didn't
		 * have a close button during MouseDown, we shouldn't return TCHT_CLOSEBUTTON until after the Click
		 * event. Otherwise the tab will close on activation if the close buttons materializes below the mouse
		 * cursor.
		 */
		if(properties.closeableTabsMode != ctmDisplayOnActiveTab || mouseStatus.overCloseButtonOnMouseDown >= -1) {
			CRect buttonRectangle;
			if(CalculateCloseButtonRectangle(tabIndex, tabIndex == SendMessage(TCM_GETCURSEL, 0, 0), &buttonRectangle)) {
				if(buttonRectangle.PtInRect(hitTestInfo.pt)) {
					hitTestInfo.flags = TCHT_CLOSEBUTTON;
				}
			}
		}
	}

	hitTestFlags = hitTestInfo.flags;
	if(pFlags) {
		*pFlags = hitTestFlags;
	}
	if(!ignoreBoundingBoxDefinition && ((properties.tabBoundingBoxDefinition & hitTestFlags) != hitTestFlags)) {
		tabIndex = -1;
	}
	return tabIndex;
}

BOOL TabStrip::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

void TabStrip::MoveTabIndex(int oldTabIndex, int newTabIndex)
{
	LONG tabID = tabIDs[oldTabIndex];
	#ifdef USE_STL
		tabIDs.erase(tabIDs.begin() + oldTabIndex);
		if(newTabIndex >= static_cast<int>(tabIDs.size())) {
			tabIDs.push_back(tabID);
		} else {
			tabIDs.insert(tabIDs.begin() + newTabIndex, tabID);
		}
		std::list<TabData>::iterator iter2 = tabParams.begin();
		if(iter2 != tabParams.end()) {
			std::advance(iter2, oldTabIndex);
			TabData tmp = *iter2;
			tabParams.erase(iter2);

			iter2 = tabParams.begin();
			if(iter2 != tabParams.end()) {
				std::advance(iter2, newTabIndex);
				tabParams.insert(iter2, tmp);
			}
		}
	#else
		tabIDs.RemoveAt(oldTabIndex);
		if(newTabIndex >= static_cast<int>(tabIDs.GetCount())) {
			tabIDs.Add(tabID);
		} else {
			tabIDs.InsertAt(newTabIndex, tabID);
		}
		POSITION pOldEntry = tabParams.FindIndex(oldTabIndex);
		if(pOldEntry) {
			TabData tmp = tabParams.GetAt(pOldEntry);
			tabParams.RemoveAt(pOldEntry);

			POSITION pNewEntry = tabParams.FindIndex(newTabIndex);
			if(pNewEntry) {
				tabParams.InsertBefore(pNewEntry, tmp);
			} else {
				tabParams.AddTail(tmp);
			}
		}
	#endif
}

HWND TabStrip::GetAssociatedWindow(int tabIndex)
{
	#ifdef USE_STL
		std::list<TabData>::iterator iter = tabParams.begin();
		if(iter != tabParams.end()) {
			std::advance(iter, tabIndex);
			if(iter != tabParams.end()) {
				return iter->hAssociatedWindow;
			}
		}
	#else
		POSITION p = tabParams.FindIndex(tabIndex);
		if(p) {
			return tabParams.GetAt(p).hAssociatedWindow;
		}
	#endif
	return NULL;
}

void TabStrip::SetAssociatedWindow(int tabIndex, HWND hAssociatedWindow)
{
	#ifdef USE_STL
		std::list<TabData>::iterator iter = tabParams.begin();
		if(iter != tabParams.end()) {
			std::advance(iter, tabIndex);
			if(iter != tabParams.end()) {
				iter->hAssociatedWindow = hAssociatedWindow;
			}
		}
	#else
		POSITION p = tabParams.FindIndex(tabIndex);
		if(p) {
			tabParams.GetAt(p).hAssociatedWindow = hAssociatedWindow;
		}
	#endif
	ATLASSERT(GetAssociatedWindow(tabIndex) == hAssociatedWindow);
}

BOOL TabStrip::IsCloseableTab(int tabIndex)
{
	if(properties.closeableTabs) {
		#ifdef USE_STL
			std::list<TabData>::iterator iter = tabParams.begin();
			if(iter != tabParams.end()) {
				std::advance(iter, tabIndex);
				if(iter != tabParams.end()) {
					return iter->isCloseable;
				}
			}
		#else
			POSITION p = tabParams.FindIndex(tabIndex);
			if(p) {
				return tabParams.GetAt(p).isCloseable;
			}
		#endif
	}
	return FALSE;
}

HRESULT TabStrip::SetCloseableTab(int tabIndex, BOOL isCloseable)
{
	if(properties.closeableTabs) {
		#ifdef USE_STL
			std::list<TabData>::iterator iter = tabParams.begin();
			if(iter != tabParams.end()) {
				std::advance(iter, tabIndex);
				if(iter != tabParams.end()) {
					iter->isCloseable = isCloseable;
					return S_OK;
				}
			}
		#else
			POSITION p = tabParams.FindIndex(tabIndex);
			if(p) {
				tabParams.GetAt(p).isCloseable = isCloseable;
				return S_OK;
			}
		#endif
	}
	return E_INVALIDARG;
}

void TabStrip::AutoScroll(void)
{
	CUpDownCtrl upDownControl = GetDlgItem(1);
	if(!upDownControl.IsWindow()) {
		return;
	}

	LONG realScrollTimeBase = properties.dragScrollTimeBase;
	if(realScrollTimeBase == -1) {
		realScrollTimeBase = GetDoubleClickTime() / 4;
	}

	if(dragDropStatus.autoScrolling.currentScrollVelocity < 0) {
		// Have we been waiting long enough since the last scroll to the left?
		if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Left) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentScrollVelocity))) {
			int lowerBound = 0;
			int upperBound = 0;
			upDownControl.GetRange32(lowerBound, upperBound);
			int currentPosition = upDownControl.GetPos32();
			if(currentPosition > lowerBound) {
				// scroll left
				dragDropStatus.autoScrolling.lastScroll_Left = GetTickCount();
				dragDropStatus.HideDragImage(TRUE);
				SendMessage(WM_HSCROLL, MAKEWPARAM(SB_THUMBPOSITION, --currentPosition), 0);
				dragDropStatus.ShowDragImage(TRUE);
			}
		}
	} else if(dragDropStatus.autoScrolling.currentScrollVelocity > 0) {
		// Have we been waiting long enough since the last scroll to the right?
		if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Right) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentScrollVelocity))) {
			int lowerBound = 0;
			int upperBound = 0;
			upDownControl.GetRange32(lowerBound, upperBound);
			int currentPosition = upDownControl.GetPos32();
			if(currentPosition < upperBound) {
				// scroll right
				dragDropStatus.autoScrolling.lastScroll_Right = GetTickCount();
				dragDropStatus.HideDragImage(TRUE);
				SendMessage(WM_HSCROLL, MAKEWPARAM(SB_THUMBPOSITION, ++currentPosition), 0);
				dragDropStatus.ShowDragImage(TRUE);
			}
		}
	}
}

void TabStrip::RebuildAcceleratorTable(void)
{
	if(properties.hAcceleratorTable) {
		DestroyAcceleratorTable(properties.hAcceleratorTable);
		properties.hAcceleratorTable = NULL;
	}

	// create a new accelerator table
	TCITEM tab = {0};
	tab.mask = TCIF_TEXT;
	tab.cchTextMax = MAX_TABTEXTLENGTH;
	tab.pszText = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (tab.cchTextMax + 1) * sizeof(TCHAR)));
	if(tab.pszText) {
		int numberOfTabs = SendMessage(TCM_GETITEMCOUNT, 0, 0);
		TCHAR* pAcceleratorChars = static_cast<TCHAR*>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, numberOfTabs * sizeof(TCHAR)));
		if(pAcceleratorChars) {
			int numberOfTabsWithAccelerator = 0;
			for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
				SendMessage(TCM_GETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab));

				pAcceleratorChars[tabIndex] = TEXT('\0');
				if(tab.pszText) {
					for(int i = lstrlen(tab.pszText) - 1; i > 0; --i) {
						if((tab.pszText[i - 1] == TEXT('&')) && (tab.pszText[i] != TEXT('&'))) {
							++numberOfTabsWithAccelerator;
							pAcceleratorChars[tabIndex] = tab.pszText[i];
							break;
						}
					}
				}
			}

			if(numberOfTabsWithAccelerator) {
				LPACCEL pAccelerators = static_cast<LPACCEL>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, (numberOfTabsWithAccelerator * 4) * sizeof(ACCEL)));
				if(pAccelerators) {
					int i = 0;
					for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
						if(pAcceleratorChars[tabIndex] != TEXT('\0')) {
							pAccelerators[i * 4].cmd = static_cast<WORD>(tabIndex);
							pAccelerators[i * 4].fVirt = FALT;
							pAccelerators[i * 4].key = static_cast<WORD>(tolower(pAcceleratorChars[tabIndex]));

							pAccelerators[i * 4 + 1].cmd = static_cast<WORD>(tabIndex);
							pAccelerators[i * 4 + 1].fVirt = 0;
							pAccelerators[i * 4 + 1].key = static_cast<WORD>(tolower(pAcceleratorChars[tabIndex]));

							pAccelerators[i * 4 + 2].cmd = static_cast<WORD>(tabIndex);
							pAccelerators[i * 4 + 2].fVirt = FVIRTKEY | FALT;
							pAccelerators[i * 4 + 2].key = LOBYTE(VkKeyScan(pAcceleratorChars[tabIndex]));

							pAccelerators[i * 4 + 3].cmd = static_cast<WORD>(tabIndex);
							pAccelerators[i * 4 + 3].fVirt = FVIRTKEY | FALT | FSHIFT;
							pAccelerators[i * 4 + 3].key = LOBYTE(VkKeyScan(pAcceleratorChars[tabIndex]));

							++i;
						}
					}
					properties.hAcceleratorTable = CreateAcceleratorTable(pAccelerators, numberOfTabsWithAccelerator * 4);
					HeapFree(GetProcessHeap(), 0, pAccelerators);
				}
			}

			HeapFree(GetProcessHeap(), 0, pAcceleratorChars);
		}
		HeapFree(GetProcessHeap(), 0, tab.pszText);
	}

	// report the new accelerator table to the container
	CComQIPtr<IOleControlSite, &IID_IOleControlSite> pSite(m_spClientSite);
	if(pSite) {
		pSite->OnControlInfoChanged();
	}
}

BOOL TabStrip::CalculateCloseButtonRectangle(int tabIndex, BOOL tabIsActive, LPRECT pButtonRectangle)
{
	if(!pButtonRectangle) {
		return FALSE;
	}
	if(properties.closeableTabsMode == ctmDisplayOnActiveTab && !tabIsActive) {
		return FALSE;
	}
	if(!IsCloseableTab(tabIndex)) {
		return FALSE;
	}

	CRect tabBoundingRectangle;
	if(SendMessage(TCM_GETITEMRECT, tabIndex, reinterpret_cast<LPARAM>(&tabBoundingRectangle))) {
		*pButtonRectangle = tabBoundingRectangle;

		DWORD style = GetStyle();
		BOOL buttons = ((style & (TCS_FLATBUTTONS | TCS_BUTTONS)) != 0);
		if(style & TCS_VERTICAL) {
			pButtonRectangle->bottom -= 3;
			pButtonRectangle->top = pButtonRectangle->bottom - 13;
			if(style & TCS_RIGHT) {
				if(tabIsActive) {
					pButtonRectangle->right -= (buttons ? 3 : 2);
				} else {
					pButtonRectangle->right -= 3;
				}
				pButtonRectangle->left = pButtonRectangle->right - 13;
			} else {
				if(tabIsActive) {
					pButtonRectangle->left += (buttons ? 2 : 1);
				} else {
					pButtonRectangle->left += 2;
				}
				pButtonRectangle->right = pButtonRectangle->left + 13;
			}
		} else {
			pButtonRectangle->right -= 3;
			pButtonRectangle->left = pButtonRectangle->right - 14;
			if(style & TCS_BOTTOM) {
				if(tabIsActive) {
					pButtonRectangle->top += (buttons ? 2 : 3);
				} else {
					pButtonRectangle->top += 2;
				}
				pButtonRectangle->bottom = pButtonRectangle->top + 14;
			} else {
				if(tabIsActive) {
					pButtonRectangle->top += 2;
				} else {
					pButtonRectangle->top += (buttons ? 2 : 3);
				}
				pButtonRectangle->bottom = pButtonRectangle->top + 14;
			}
		}

		return TRUE;
	}
	return FALSE;
}

void TabStrip::DrawCloseButtonsAndInsertionMarks(HDC hTargetDC)
{
	CDCHandle targetDC = hTargetDC;

	if(properties.closeableTabs) {
		CTheme themingEngine;
		if(flags.usingThemes) {
			themingEngine.OpenThemeData(*this, VSCLASS_WINDOW);
		}
		DWORD style = GetStyle();
		BOOL enabled = ((style & WS_DISABLED) == 0);
		BOOL buttons = ((style & (TCS_FLATBUTTONS | TCS_BUTTONS)) != 0);

		POINT mousePosition;
		GetCursorPos(&mousePosition);
		ScreenToClient(&mousePosition);

		int numberOfTabs = SendMessage(TCM_GETITEMCOUNT, 0, 0);
		int activeTab = SendMessage(TCM_GETCURSEL, 0, 0);
		for(int tabIndex = (properties.closeableTabsMode == ctmDisplayOnActiveTab ? activeTab : 0); tabIndex < (properties.closeableTabsMode == ctmDisplayOnActiveTab ? activeTab + 1 : numberOfTabs); ++tabIndex) {
			RECT buttonRectangle = {0};
			if(CalculateCloseButtonRectangle(tabIndex, tabIndex == activeTab, &buttonRectangle)) {
				if(buttons || themingEngine.IsThemeNull()) {
					if(mouseStatus.overCloseButton == tabIndex) {
						if(IsLeftMouseButtonDown()) {
							// pushed
							targetDC.DrawFrameControl(&buttonRectangle, DFC_CAPTION, DFCS_CAPTIONCLOSE | DFCS_FLAT | (enabled ? 0 : DFCS_INACTIVE) | DFCS_PUSHED);
						} else {
							// hot
							targetDC.DrawFrameControl(&buttonRectangle, DFC_CAPTION, DFCS_CAPTIONCLOSE | DFCS_FLAT | (enabled ? 0 : DFCS_INACTIVE) | DFCS_HOT);
						}
					} else {
						// normal
						targetDC.DrawFrameControl(&buttonRectangle, DFC_CAPTION, DFCS_CAPTIONCLOSE | DFCS_FLAT | (enabled ? 0 : DFCS_INACTIVE));
					}
				} else {
					if(mouseStatus.overCloseButton == tabIndex) {
						if(IsLeftMouseButtonDown()) {
							// pushed
							themingEngine.DrawThemeBackground(targetDC, WP_SMALLCLOSEBUTTON, (enabled ? CBS_PUSHED : CBS_DISABLED), &buttonRectangle, NULL);
						} else {
							// hot
							themingEngine.DrawThemeBackground(targetDC, WP_SMALLCLOSEBUTTON, (enabled ? CBS_HOT : CBS_DISABLED), &buttonRectangle, NULL);
						}
					} else {
						// normal
						themingEngine.DrawThemeBackground(targetDC, WP_SMALLCLOSEBUTTON, (enabled ? CBS_NORMAL : CBS_DISABLED), &buttonRectangle, NULL);
					}
				}
			}
		}
	}

	if(insertMark.tabIndex != -1) {
		CPen pen;
		pen.CreatePen(PS_SOLID, 1, insertMark.color);
		HPEN hPreviousPen = targetDC.SelectPen(pen);

		RECT tabBoundingRectangle = {0};
		SendMessage(TCM_GETITEMRECT, insertMark.tabIndex, reinterpret_cast<LPARAM>(&tabBoundingRectangle));
		RECT insertMarkRect = {0};
		if(GetStyle() & TCS_VERTICAL) {
			if(insertMark.afterTab) {
				insertMarkRect.top = tabBoundingRectangle.bottom - 3;
				insertMarkRect.bottom = tabBoundingRectangle.bottom + 3;
			} else {
				insertMarkRect.top = tabBoundingRectangle.top - 3;
				insertMarkRect.bottom = tabBoundingRectangle.top + 3;
			}
			insertMarkRect.left = tabBoundingRectangle.left + 1;
			insertMarkRect.right = tabBoundingRectangle.right - 1;

			// draw the main line
			targetDC.MoveTo(insertMarkRect.left, insertMarkRect.top + 2);
			targetDC.LineTo(insertMarkRect.right, insertMarkRect.top + 2);
			targetDC.MoveTo(insertMarkRect.left, insertMarkRect.top + 3);
			targetDC.LineTo(insertMarkRect.right, insertMarkRect.top + 3);

			// draw the ends
			targetDC.MoveTo(insertMarkRect.left, insertMarkRect.top);
			targetDC.LineTo(insertMarkRect.left, insertMarkRect.bottom);
			targetDC.MoveTo(insertMarkRect.left + 1, insertMarkRect.top + 1);
			targetDC.LineTo(insertMarkRect.left + 1, insertMarkRect.bottom - 1);
			targetDC.MoveTo(insertMarkRect.right - 1, insertMarkRect.top + 1);
			targetDC.LineTo(insertMarkRect.right - 1, insertMarkRect.bottom - 1);
			targetDC.MoveTo(insertMarkRect.right, insertMarkRect.top);
			targetDC.LineTo(insertMarkRect.right, insertMarkRect.bottom);
		} else {
			if(insertMark.afterTab) {
				insertMarkRect.left = tabBoundingRectangle.right - 3;
				insertMarkRect.right = tabBoundingRectangle.right + 3;
			} else {
				insertMarkRect.left = tabBoundingRectangle.left - 3;
				insertMarkRect.right = tabBoundingRectangle.left + 3;
			}
			insertMarkRect.top = tabBoundingRectangle.top + 1;
			insertMarkRect.bottom = tabBoundingRectangle.bottom - 1;

			// draw the main line
			targetDC.MoveTo(insertMarkRect.left + 2, insertMarkRect.top);
			targetDC.LineTo(insertMarkRect.left + 2, insertMarkRect.bottom);
			targetDC.MoveTo(insertMarkRect.left + 3, insertMarkRect.top);
			targetDC.LineTo(insertMarkRect.left + 3, insertMarkRect.bottom);

			// draw the ends
			targetDC.MoveTo(insertMarkRect.left, insertMarkRect.top);
			targetDC.LineTo(insertMarkRect.right, insertMarkRect.top);
			targetDC.MoveTo(insertMarkRect.left + 1, insertMarkRect.top + 1);
			targetDC.LineTo(insertMarkRect.right - 1, insertMarkRect.top + 1);
			targetDC.MoveTo(insertMarkRect.left + 1, insertMarkRect.bottom - 1);
			targetDC.LineTo(insertMarkRect.right - 1, insertMarkRect.bottom - 1);
			targetDC.MoveTo(insertMarkRect.left, insertMarkRect.bottom);
			targetDC.LineTo(insertMarkRect.right, insertMarkRect.bottom);
		}

		targetDC.SelectPen(hPreviousPen);
	}
}

BOOL TabStrip::IsLeftMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	}
}

BOOL TabStrip::IsRightMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	}
}


HRESULT TabStrip::CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing)
{
	if(!hIcon || !pBits || !pWICImagingFactory) {
		return E_FAIL;
	}

	ICONINFO iconInfo;
	GetIconInfo(hIcon, &iconInfo);
	ATLASSERT(iconInfo.hbmColor);
	BITMAP bitmapInfo = {0};
	if(iconInfo.hbmColor) {
		GetObject(iconInfo.hbmColor, sizeof(BITMAP), &bitmapInfo);
	} else if(iconInfo.hbmMask) {
		GetObject(iconInfo.hbmMask, sizeof(BITMAP), &bitmapInfo);
	}
	bitmapInfo.bmHeight = abs(bitmapInfo.bmHeight);
	BOOL needsFrame = ((bitmapInfo.bmWidth < size.cx) || (bitmapInfo.bmHeight < size.cy));
	if(iconInfo.hbmColor) {
		DeleteObject(iconInfo.hbmColor);
	}
	if(iconInfo.hbmMask) {
		DeleteObject(iconInfo.hbmMask);
	}

	HRESULT hr = E_FAIL;

	CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
	if(!needsFrame) {
		hr = pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapScaler);
	}
	if(needsFrame || SUCCEEDED(hr)) {
		CComPtr<IWICBitmap> pWICBitmapSource = NULL;
		hr = pWICImagingFactory->CreateBitmapFromHICON(hIcon, &pWICBitmapSource);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapSource);
		if(SUCCEEDED(hr)) {
			if(!needsFrame) {
				hr = pWICBitmapScaler->Initialize(pWICBitmapSource, size.cx, size.cy, WICBitmapInterpolationModeFant);
			}
			if(SUCCEEDED(hr)) {
				WICRect rc = {0};
				if(needsFrame) {
					rc.Height = bitmapInfo.bmHeight;
					rc.Width = bitmapInfo.bmWidth;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					LPRGBQUAD pIconBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, rc.Width * rc.Height * sizeof(RGBQUAD)));
					hr = pWICBitmapSource->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pIconBits));
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						// center the icon
						int xIconStart = (size.cx - bitmapInfo.bmWidth) / 2;
						int yIconStart = (size.cy - bitmapInfo.bmHeight) / 2;
						LPRGBQUAD pIconPixel = pIconBits;
						LPRGBQUAD pPixel = pBits;
						pPixel += yIconStart * size.cx;
						for(int y = yIconStart; y < yIconStart + bitmapInfo.bmHeight; ++y, pPixel += size.cx, pIconPixel += bitmapInfo.bmWidth) {
							CopyMemory(pPixel + xIconStart, pIconPixel, bitmapInfo.bmWidth * sizeof(RGBQUAD));
						}
						HeapFree(GetProcessHeap(), 0, pIconBits);

						rc.Height = size.cy;
						rc.Width = size.cx;

						// TODO: now draw a frame around it
					}
				} else {
					rc.Height = size.cy;
					rc.Width = size.cx;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					hr = pWICBitmapScaler->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pBits));
					ATLASSERT(SUCCEEDED(hr));

					if(SUCCEEDED(hr) && doAlphaChannelPostProcessing) {
						for(int i = 0; i < rc.Width * rc.Height; ++i, ++pBits) {
							if(pBits->rgbReserved == 0x00) {
								ZeroMemory(pBits, sizeof(RGBQUAD));
							}
						}
					}
				}
			} else {
				ATLASSERT(FALSE && "Bitmap scaler failed");
			}
		}
	}
	return hr;
}