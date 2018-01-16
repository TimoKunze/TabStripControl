//////////////////////////////////////////////////////////////////////
/// \class TabStrip
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c SysTabControl32</em>
///
/// This class superclasses \c SysTabControl32 and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo Verify documentation of \c PreTranslateAccelerator.
///
/// \if UNICODE
///   \sa TabStripCtlLibU::ITabStrip, TabStripTab, TabStripTabs, TabStripTabContainer, VirtualTabStripTab
/// \else
///   \sa TabStripCtlLibA::ITabStrip, TabStripTab, TabStripTabs, TabStripTabContainer, VirtualTabStripTab
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TabStripCtlU.h"
#else
	#include "TabStripCtlA.h"
#endif
#include "_ITabStripEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "CommonProperties.h"
#include "TabStripTab.h"
#include "TabStripTabs.h"
#include "TabStripTabContainer.h"
#include "VirtualTabStripTab.h"
#include "TargetOLEDataObject.h"
#include "SourceOLEDataObject.h"


class ATL_NO_VTABLE TabStrip : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<ITabStrip, &IID_ITabStrip, &LIBID_TabStripCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<ITabStrip, &IID_ITabStrip, &LIBID_TabStripCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<TabStrip>,
    public IOleControlImpl<TabStrip>,
    public IOleObjectImpl<TabStrip>,
    public IOleInPlaceActiveObjectImpl<TabStrip>,
    public IViewObjectExImpl<TabStrip>,
    public IOleInPlaceObjectWindowlessImpl<TabStrip>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TabStrip>,
    public Proxy_ITabStripEvents<TabStrip>,
    public IPersistStorageImpl<TabStrip>,
    public IPersistPropertyBagImpl<TabStrip>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<TabStrip>,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_TabStrip, &__uuidof(_ITabStripEvents), &LIBID_TabStripCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_TabStrip, &__uuidof(_ITabStripEvents), &LIBID_TabStripCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<TabStrip>,
    public CComCoClass<TabStrip, &CLSID_TabStrip>,
    public CComControl<TabStrip>,
    public IPerPropertyBrowsingImpl<TabStrip>,
    public IDropTarget,
    public IDropSource,
    public IDropSourceNotify,
    public ICategorizeProperties,
    public ICreditsProvider
{
	friend class TabStripTab;
	friend class TabStripTabs;
	friend class TabStripTabContainer;
	friend class SourceOLEDataObject;

protected:
	/// \brief <em>A custom extended tabstrip style which enables drag detection</em>
	#define TCS_EX_DETECTDRAGDROP 0x80000000
	/// \brief <em>A custom tabstrip message which retrieves the position of the insertion mark</em>
	///
	/// \sa TCM_SETINSERTMARK, TCINSERTMARK, TCM_GETINSERTMARKCOLOR
	#define TCM_GETINSERTMARK (WM_USER + 2)
	/// \brief <em>A custom tabstrip message which sets the position of the insertion mark</em>
	///
	/// \sa TCM_GETINSERTMARK, TCM_SETINSERTMARKCOLOR
	#define TCM_SETINSERTMARK (WM_USER + 3)
	/// \brief <em>A custom tabstrip message which retrieves the color of the insertion mark</em>
	///
	/// \sa TCM_SETINSERTMARKCOLOR, TCM_GETINSERTMARK
	#define TCM_GETINSERTMARKCOLOR (WM_USER + 4)
	/// \brief <em>A custom tabstrip message which sets the color of the insertion mark</em>
	///
	/// \sa TCM_GETINSERTMARKCOLOR, TCM_SETINSERTMARK
	#define TCM_SETINSERTMARKCOLOR (WM_USER + 5)
	/// \brief <em>A custom tabstrip notification which notifies of a detected drag action</em>
	///
	/// \sa TCN_BEGINRDRAG, tagNMTCBEGINDRAG
	#define TCN_BEGINDRAG (TCN_FIRST - 50)
	/// \brief <em>A custom tabstrip notification which notifies of a detected drag action</em>
	///
	/// \sa TCN_BEGINDRAG, tagNMTCBEGINDRAG
	#define TCN_BEGINRDRAG (TCN_FIRST - 51)

	/// \brief <em>A struct used with \c TCM_GETINSERTMARK</em>
	///
	/// \sa TCM_GETINSERTMARK
	typedef struct
	{
		/// \brief <em>The struct's size in bytes</em>
		UINT size;
		/// \brief <em>\c TRUE, if the insertion mark is displayed after the specified tab; otherwise \c FALSE</em>
		BOOL afterTab;
		/// \brief <em>The zero-based index of the tab at which the insertion mark is displayed</em>
		int tabIndex;
	} TCINSERTMARK, *LPTCINSERTMARK;

	/// \brief <em>A struct used with \c TCN_BEGINDRAG and \c TCN_BEGINRDRAG</em>
	///
	/// \sa TCN_BEGINDRAG, TCN_BEGINRDRAG
	typedef struct tagNMTCBEGINDRAG
	{
		/// \brief <em>The notification header</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms672586.aspx">NMHDR</a>
		NMHDR hdr;
		/// \brief <em>The dragged tab's zero-based index</em>
		int iItem;
		/// \brief <em>The position at which the drag was detected</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms536119.aspx">POINT</a>
		POINT ptDrag;
	} NMTCBEGINDRAG, *LPNMTCBEGINDRAG;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa ~TabStrip
	TabStrip();
	/// \brief <em>The destructor of this class</em>
	///
	/// Used for cleaning up.
	///
	/// \sa TabStrip
	~TabStrip();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST | OLEMISC_SIMPLEFRAME)
		DECLARE_REGISTRY_RESOURCEID(IDR_TABSTRIP)

		#ifdef UNICODE
			DECLARE_WND_SUPERCLASS(TEXT("TabStripU"), WC_TABCONTROLW)
		#else
			DECLARE_WND_SUPERCLASS(TEXT("TabStripA"), WC_TABCONTROLA)
		#endif

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(TabStrip)
			COM_INTERFACE_ENTRY(ITabStrip)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IViewObjectEx)
			COM_INTERFACE_ENTRY(IViewObject2)
			COM_INTERFACE_ENTRY(IViewObject)
			COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceObject)
			COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
			COM_INTERFACE_ENTRY(IOleControl)
			COM_INTERFACE_ENTRY(IOleObject)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IQuickActivate)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY(IProvideClassInfo)
			COM_INTERFACE_ENTRY(IProvideClassInfo2)
			COM_INTERFACE_ENTRY_IID(IID_ICategorizeProperties, ICategorizeProperties)
			COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
			COM_INTERFACE_ENTRY(IPerPropertyBrowsing)
			COM_INTERFACE_ENTRY(IDropTarget)
			COM_INTERFACE_ENTRY(IDropSource)
			COM_INTERFACE_ENTRY(IDropSourceNotify)
		END_COM_MAP()

		BEGIN_PROP_MAP(TabStrip)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("AllowDragDrop", DISPID_TABSTRIPCTL_ALLOWDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Appearance", DISPID_TABSTRIPCTL_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_TABSTRIPCTL_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("CloseableTabs", DISPID_TABSTRIPCTL_CLOSEABLETABS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("CloseableTabsMode", DISPID_TABSTRIPCTL_CLOSEABLETABSMODE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_TABSTRIPCTL_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			//PROP_ENTRY_TYPE("DontRedraw", DISPID_TABSTRIPCTL_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DragActivateTime", DISPID_TABSTRIPCTL_DRAGACTIVATETIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DragScrollTimeBase", DISPID_TABSTRIPCTL_DRAGSCROLLTIMEBASE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Enabled", DISPID_TABSTRIPCTL_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("FixedTabWidth", DISPID_TABSTRIPCTL_FIXEDTABWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("FocusOnButtonDown", DISPID_TABSTRIPCTL_FOCUSONBUTTONDOWN, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_TABSTRIPCTL_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("HorizontalPadding", DISPID_TABSTRIPCTL_HORIZONTALPADDING, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HotTracking", DISPID_TABSTRIPCTL_HOTTRACKING, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("HoverTime", DISPID_TABSTRIPCTL_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("InsertMarkColor", DISPID_TABSTRIPCTL_INSERTMARKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("MinTabWidth", DISPID_TABSTRIPCTL_MINTABWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_TABSTRIPCTL_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_TABSTRIPCTL_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MultiRow", DISPID_TABSTRIPCTL_MULTIROW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("MultiSelect", DISPID_TABSTRIPCTL_MULTISELECT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("OLEDragImageStyle", DISPID_TABSTRIPCTL_OLEDRAGIMAGESTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("OwnerDrawn", DISPID_TABSTRIPCTL_OWNERDRAWN, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_TABSTRIPCTL_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RaggedTabRows", DISPID_TABSTRIPCTL_RAGGEDTABROWS, CLSID_NULL, VT_BOOL)
			//PROP_ENTRY_TYPE("ReflectContextMenuMessages", DISPID_TABSTRIPCTL_REFLECTCONTEXTMENUMESSAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_TABSTRIPCTL_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_TABSTRIPCTL_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ScrollToOpposite", DISPID_TABSTRIPCTL_SCROLLTOOPPOSITE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ShowButtonSeparators", DISPID_TABSTRIPCTL_SHOWBUTTONSEPARATORS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ShowToolTips", DISPID_TABSTRIPCTL_SHOWTOOLTIPS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Style", DISPID_TABSTRIPCTL_STYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_TABSTRIPCTL_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("TabBoundingBoxDefinition", DISPID_TABSTRIPCTL_TABBOUNDINGBOXDEFINITION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("TabCaptionAlignment", DISPID_TABSTRIPCTL_TABCAPTIONALIGNMENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("TabHeight", DISPID_TABSTRIPCTL_TABHEIGHT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("TabPlacement", DISPID_TABSTRIPCTL_TABPLACEMENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("UseFixedTabWidth", DISPID_TABSTRIPCTL_USEFIXEDTABWIDTH, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_TABSTRIPCTL_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("VerticalPadding", DISPID_TABSTRIPCTL_VERTICALPADDING, CLSID_NULL, VT_I4)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(TabStrip)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_ITabStripEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP_SIMPLE_FRAME(TabStrip)
			BEGIN_MSG_FILTERING(pSimpleFrameSite)
				FILTERED_MESSAGE_HANDLER(WM_CHAR, OnChar)
				FILTERED_MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
				FILTERED_MESSAGE_HANDLER(WM_CREATE, OnCreate)
				FILTERED_MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
				FILTERED_MESSAGE_HANDLER(WM_HSCROLL, OnHScroll)
				FILTERED_MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
				FILTERED_MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
				FILTERED_MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
				FILTERED_MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
				FILTERED_MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
				FILTERED_MESSAGE_HANDLER(WM_NCHITTEST, OnNCHitTest)
				FILTERED_MESSAGE_HANDLER(WM_PAINT, OnPaint)
				FILTERED_MESSAGE_HANDLER(WM_PRINTCLIENT, OnPaint)
				FILTERED_MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
				FILTERED_MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)     // TODO: Should we use NM_SETCURSOR instead?
				FILTERED_MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
				//FILTERED_MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
				FILTERED_MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
				FILTERED_MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
				FILTERED_MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
				FILTERED_MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
				FILTERED_MESSAGE_HANDLER(WM_TIMER, OnTimer)
				FILTERED_MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
				FILTERED_MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

				FILTERED_MESSAGE_HANDLER(OCM_DRAWITEM, OnReflectedDrawItem)
				FILTERED_MESSAGE_HANDLER(OCM__BASE + WM_NOTIFYFORMAT, OnReflectedNotifyFormat)

				FILTERED_MESSAGE_HANDLER(GetDragImageMessage(), OnGetDragImage)

				FILTERED_MESSAGE_HANDLER(TCM_CREATEDRAGIMAGE, OnCreateDragImage)
				FILTERED_MESSAGE_HANDLER(TCM_DELETEALLITEMS, OnDeleteAllItems)
				FILTERED_MESSAGE_HANDLER(TCM_DELETEITEM, OnDeleteItem)
				FILTERED_MESSAGE_HANDLER(TCM_DESELECTALL, OnDeselectAll)
				FILTERED_MESSAGE_HANDLER(TCM_GETEXTENDEDSTYLE, OnGetExtendedStyle)
				FILTERED_MESSAGE_HANDLER(TCM_GETINSERTMARK, OnGetInsertMark)
				FILTERED_MESSAGE_HANDLER(TCM_GETINSERTMARKCOLOR, OnGetInsertMarkColor)
				FILTERED_MESSAGE_HANDLER(TCM_GETITEMA, OnGetItem)
				FILTERED_MESSAGE_HANDLER(TCM_GETITEMW, OnGetItem)
				FILTERED_MESSAGE_HANDLER(TCM_HIGHLIGHTITEM, OnHighlightItem)
				FILTERED_MESSAGE_HANDLER(TCM_INSERTITEMA, OnInsertItem)
				FILTERED_MESSAGE_HANDLER(TCM_INSERTITEMW, OnInsertItem)
				FILTERED_MESSAGE_HANDLER(TCM_SETCURFOCUS, OnSetCurFocus)
				FILTERED_MESSAGE_HANDLER(TCM_SETCURSEL, OnSetCurSel)
				FILTERED_MESSAGE_HANDLER(TCM_SETEXTENDEDSTYLE, OnSetExtendedStyle)
				FILTERED_MESSAGE_HANDLER(TCM_SETIMAGELIST, OnSetImageList)
				FILTERED_MESSAGE_HANDLER(TCM_SETINSERTMARK, OnSetInsertMark)
				FILTERED_MESSAGE_HANDLER(TCM_SETINSERTMARKCOLOR, OnSetInsertMarkColor)
				FILTERED_MESSAGE_HANDLER(TCM_SETITEMA, OnSetItem)
				FILTERED_MESSAGE_HANDLER(TCM_SETITEMW, OnSetItem)
				FILTERED_MESSAGE_HANDLER(TCM_SETITEMSIZE, OnSetItemSize)
				FILTERED_MESSAGE_HANDLER(TCM_SETMINTABWIDTH, OnSetMinTabWidth)
				FILTERED_MESSAGE_HANDLER(TCM_SETPADDING, OnSetPadding)

				FILTERED_NOTIFY_CODE_HANDLER(TTN_GETDISPINFOA, OnToolTipGetDispInfoNotificationA)
				FILTERED_NOTIFY_CODE_HANDLER(TTN_GETDISPINFOW, OnToolTipGetDispInfoNotificationW)

				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(NM_CLICK, OnClickNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(NM_RCLICK, OnRClickNotification)

				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TCN_BEGINDRAG, OnBeginDragNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TCN_BEGINRDRAG, OnBeginRDragNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TCN_FOCUSCHANGE, OnFocusChangeNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TCN_GETOBJECT, OnGetObjectNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TCN_SELCHANGE, OnSelChangeNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TCN_SELCHANGING, OnSelChangingNotification)

				FILTERED_CHAIN_MSG_MAP(CComControl<TabStrip>)
				FILTERED_DEFAULT_REFLECTION_HANDLER()
				FILTERED_REFLECT_NOTIFICATIONS()
			END_MSG_FILTERING(pSimpleFrameSite)
		END_MSG_MAP_SIMPLE_FRAME()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportsErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of persistance
	///
	//@{
	/// \brief <em>Overrides \c IPersistStreamInitImpl::GetSizeMax to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we communicate directly with the stream. This requires we override
	/// \c IPersistStreamInitImpl::GetSizeMax.
	///
	/// \param[in] pSize The maximum number of bytes that persistence of the control's properties will
	///            consume.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687287.aspx">IPersistStreamInit::GetSizeMax</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pSize);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Load to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Load and read directly from
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680730.aspx">IPersistStreamInit::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPSTREAM pStream);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Save to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Save and write directly into
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694439.aspx">IPersistStreamInit::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPSTREAM pStream, BOOL clearDirtyFlag);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ITabStrip
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c ActiveTab property</em>
	///
	/// Retrieves the control's active tab. The active tab is the tab whose content is currently displayed.
	///
	/// \param[out] ppActiveTab Receives the active tab's \c ITabStripTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_ActiveTab, get_CaretTab, TabStripTab::get_Active, get_MultiSelect,
	///     Raise_ActiveTabChanging, Raise_ActiveTabChanged
	virtual HRESULT STDMETHODCALLTYPE get_ActiveTab(ITabStripTab** ppActiveTab);
	/// \brief <em>Sets the \c ActiveTab property</em>
	///
	/// Sets the control's active tab. The active tab is the tab whose content is currently displayed.
	///
	/// \param[in] pNewActiveTab The new active tab's \c ITabStripTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ActiveTab, putref_CaretTab, put_MultiSelect, Raise_ActiveTabChanging, Raise_ActiveTabChanged
	virtual HRESULT STDMETHODCALLTYPE putref_ActiveTab(ITabStripTab* pNewActiveTab);
	/// \brief <em>Retrieves the current setting of the \c AllowDragDrop property</em>
	///
	/// Retrieves whether drag'n'drop mode can be entered. If set to \c VARIANT_TRUE, drag'n'drop mode
	/// can be entered by pressing the left or right mouse button over a tab and then moving the mouse
	/// with the button still pressed. If set to \c VARIANT_FALSE, drag'n'drop mode is not available -
	/// this also means the \c TCN_BEGINDRAG and \c TCN_BEGINRDRAG notifications are not sent.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowDragDrop, get_RegisterForOLEDragDrop, get_DragActivateTime, get_DragScrollTimeBase,
	///     SetInsertMarkPosition, Raise_TabBeginDrag, Raise_TabBeginRDrag
	virtual HRESULT STDMETHODCALLTYPE get_AllowDragDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowDragDrop property</em>
	///
	/// Sets whether drag'n'drop mode can be entered. If set to \c VARIANT_TRUE, drag'n'drop mode
	/// can be entered by pressing the left or right mouse button over a tab and then moving the mouse
	/// with the button still pressed. If set to \c VARIANT_FALSE, drag'n'drop mode is not available -
	/// this also means the \c TCN_BEGINDRAG and \c TCN_BEGINRDRAG notifications are not sent.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowDragDrop, put_RegisterForOLEDragDrop, put_DragActivateTime, put_DragScrollTimeBase,
	///     SetInsertMarkPosition, Raise_TabBeginDrag, Raise_TabBeginRDrag
	virtual HRESULT STDMETHODCALLTYPE put_AllowDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Appearance property</em>
	///
	/// Retrieves the kind of border that is drawn around the control. Any of the values defined by
	/// the \c AppearanceConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Appearance, get_BorderStyle, TabStripCtlLibU::AppearanceConstants
	/// \else
	///   \sa put_Appearance, get_BorderStyle, TabStripCtlLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Appearance(AppearanceConstants* pValue);
	/// \brief <em>Sets the \c Appearance property</em>
	///
	/// Sets the kind of border that is drawn around the control. Any of the values defined by the
	/// \c AppearanceConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Appearance, put_BorderStyle, TabStripCtlLibU::AppearanceConstants
	/// \else
	///   \sa get_Appearance, put_BorderStyle, TabStripCtlLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Appearance(AppearanceConstants newValue);
	/// \brief <em>Retrieves the control's application ID</em>
	///
	/// Retrieves the control's application ID. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppID(SHORT* pValue);
	/// \brief <em>Retrieves the control's application name</em>
	///
	/// Retrieves the control's application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppName(BSTR* pValue);
	/// \brief <em>Retrieves the control's short application name</em>
	///
	/// Retrieves the control's short application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The short application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_Build, get_CharSet, get_IsRelease, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppShortName(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c BorderStyle property</em>
	///
	/// Retrieves the kind of inner border that is drawn around the control. Any of the values defined
	/// by the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_BorderStyle, get_Appearance, TabStripCtlLibU::BorderStyleConstants
	/// \else
	///   \sa put_BorderStyle, get_Appearance, TabStripCtlLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BorderStyle(BorderStyleConstants* pValue);
	/// \brief <em>Sets the \c BorderStyle property</em>
	///
	/// Sets the kind of inner border that is drawn around the control. Any of the values defined by
	/// the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_BorderStyle, put_Appearance, TabStripCtlLibU::BorderStyleConstants
	/// \else
	///   \sa get_BorderStyle, put_Appearance, TabStripCtlLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BorderStyle(BorderStyleConstants newValue);
	/// \brief <em>Retrieves the control's build number</em>
	///
	/// Retrieves the control's build number. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The build number.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Build(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c CaretTab property</em>
	///
	/// Retrieves the control's caret tab. The caret tab is the tab that has the keyboard focus.
	///
	/// \param[out] ppCaretTab Receives the caret tab's \c ITabStripTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_CaretTab, get_ActiveTab, TabStripTab::get_Caret, get_MultiSelect, Raise_CaretChanged
	virtual HRESULT STDMETHODCALLTYPE get_CaretTab(ITabStripTab** ppCaretTab);
	/// \brief <em>Sets the \c CaretTab property</em>
	///
	/// Sets the control's caret tab. The caret tab is the tab that has the keyboard focus.
	///
	/// \param[in] pNewCaretTab The new caret tab's \c ITabStripTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CaretTab, putref_ActiveTab, put_MultiSelect, Raise_CaretChanged
	virtual HRESULT STDMETHODCALLTYPE putref_CaretTab(ITabStripTab* pNewCaretTab);
	/// \brief <em>Retrieves the control's character set</em>
	///
	/// Retrieves the control's character set (Unicode/ANSI). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The character set.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_CharSet(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c CloseableTabs property</em>
	///
	/// Retrieves whether close buttons are drawn for each tab. If set to \c VARIANT_TRUE, close buttons
	/// are drawn; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CloseableTabs, get_CloseableTabsMode, TabStripTab::get_HasCloseButton
	virtual HRESULT STDMETHODCALLTYPE get_CloseableTabs(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CloseableTabs property</em>
	///
	/// Sets whether close buttons are drawn for each tab. If set to \c VARIANT_TRUE, close buttons
	/// are drawn; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CloseableTabs, put_CloseableTabsMode, TabStripTab::put_HasCloseButton
	virtual HRESULT STDMETHODCALLTYPE put_CloseableTabs(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledEvents property</em>
	///
	/// Retrieves which tabs are drawn with close buttons. Any of the values defined by the
	/// \c CloseableTabsModeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_CloseableTabsMode, get_CloseableTabs, TabStripCtlLibU::CloseableTabsModeConstants
	/// \else
	///   \sa put_CloseableTabsMode, get_CloseableTabs, TabStripCtlLibA::CloseableTabsModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CloseableTabsMode(CloseableTabsModeConstants* pValue);
	/// \brief <em>Sets the \c DisabledEvents property</em>
	///
	/// Sets which tabs are drawn with close buttons. Any of the values defined by the
	/// \c CloseableTabsModeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_CloseableTabsMode, put_CloseableTabs, TabStripCtlLibU::CloseableTabsModeConstants
	/// \else
	///   \sa get_CloseableTabsMode, put_CloseableTabs, TabStripCtlLibA::CloseableTabsModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_CloseableTabsMode(CloseableTabsModeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledEvents property</em>
	///
	/// Retrieves the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisabledEvents, TabStripCtlLibU::DisabledEventsConstants
	/// \else
	///   \sa put_DisabledEvents, TabStripCtlLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisabledEvents(DisabledEventsConstants* pValue);
	/// \brief <em>Sets the \c DisabledEvents property</em>
	///
	/// Sets the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisabledEvents, TabStripCtlLibU::DisabledEventsConstants
	/// \else
	///   \sa get_DisabledEvents, TabStripCtlLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisabledEvents(DisabledEventsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayAreaHeight property</em>
	///
	/// Retrieves the height of the control's display area in pixels. The display area is the area
	/// available for the tabs' content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_TabHeight, TabStripTab::get_Height, get_DisplayAreaLeft, get_DisplayAreaTop,
	///     get_DisplayAreaWidth, CalculateDisplayArea
	virtual HRESULT STDMETHODCALLTYPE get_DisplayAreaHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayAreaLeft property</em>
	///
	/// Retrieves the distance between the control's left border and the display area's left border in
	/// pixels. The display area is the area available for the tabs' content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TabStripTab::get_Left, get_DisplayAreaHeight, get_DisplayAreaTop, get_DisplayAreaWidth,
	///     CalculateDisplayArea
	virtual HRESULT STDMETHODCALLTYPE get_DisplayAreaLeft(OLE_XPOS_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayAreaTop property</em>
	///
	/// Retrieves the distance between the control's upper border and the display area's upper border in
	/// pixels. The display area is the area available for the tabs' content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TabStripTab::get_Top, get_DisplayAreaHeight, get_DisplayAreaLeft, get_DisplayAreaWidth,
	///     CalculateDisplayArea
	virtual HRESULT STDMETHODCALLTYPE get_DisplayAreaTop(OLE_YPOS_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayAreaWidth property</em>
	///
	/// Retrieves the width of the control's display area in pixels. The display area is the area
	/// available for the tabs' content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TabStripTab::get_Width, get_DisplayAreaHeight, get_DisplayAreaLeft, get_DisplayAreaTop,
	///     CalculateDisplayArea
	virtual HRESULT STDMETHODCALLTYPE get_DisplayAreaWidth(OLE_XSIZE_PIXELS* pValue);
	// \brief <em>Retrieves the current setting of the \c DontRedraw property</em>
	//
	// Retrieves whether automatic redrawing of the control is enabled or disabled. If set to
	// \c VARIANT_FALSE, the control will redraw itself automatically; otherwise not.
	//
	// \param[out] pValue The property's current setting.
	//
	// \return An \c HRESULT error code.
	//
	// \sa put_DontRedraw
	//virtual HRESULT STDMETHODCALLTYPE get_DontRedraw(VARIANT_BOOL* pValue);
	// \brief <em>Sets the \c DontRedraw property</em>
	//
	// Enables or disables automatic redrawing of the control. Disabling redraw while doing large
	// changes on the control (like adding many panels) may increase performance.
	// If set to \c VARIANT_FALSE, the control will redraw itself automatically; otherwise not.
	//
	// \param[in] newValue The setting to apply.
	//
	// \return An \c HRESULT error code.
	//
	// \sa get_DontRedraw
	//virtual HRESULT STDMETHODCALLTYPE put_DontRedraw(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DragActivateTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be placed over a tab during a
	/// drag'n'drop operation before this tab will be activated automatically. If set to 0,
	/// auto-activation is disabled. If set to -1, the system's double-click time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DragActivateTime, get_AllowDragDrop, get_RegisterForOLEDragDrop, get_DragScrollTimeBase,
	///     Raise_DragMouseMove, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragActivateTime(LONG* pValue);
	/// \brief <em>Sets the \c DragActivateTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be placed over a tab during a
	/// drag'n'drop operation before this tab will be activated automatically. If set to 0,
	/// auto-activation is disabled. If set to -1, the system's double-click time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DragActivateTime, put_AllowDragDrop, put_RegisterForOLEDragDrop, put_DragScrollTimeBase,
	///     Raise_DragMouseMove, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragActivateTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c DraggedTabs property</em>
	///
	/// Retrieves a collection object wrapping the control's tabs that are currently dragged. These
	/// are the same tabs that were passed to the \c BeginDrag or \c OLEDrag method.
	///
	/// \param[out] ppTabs Receives the collection object's \c ITabStripTabContainer implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa BeginDrag, OLEDrag, TabStripTabContainer, get_Tabs
	virtual HRESULT STDMETHODCALLTYPE get_DraggedTabs(ITabStripTabContainer** ppTabs);
	/// \brief <em>Retrieves the current setting of the \c DragScrollTimeBase property</em>
	///
	/// Retrieves the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time, divided by 4, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DragScrollTimeBase, get_AllowDragDrop, get_RegisterForOLEDragDrop, get_DragActivateTime,
	///     Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragScrollTimeBase(LONG* pValue);
	/// \brief <em>Sets the \c DragScrollTimeBase property</em>
	///
	/// Sets the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time divided by 4 is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DragScrollTimeBase, put_AllowDragDrop, put_RegisterForOLEDragDrop, put_DragActivateTime,
	///     Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragScrollTimeBase(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c DropHilitedTab property</em>
	///
	/// Retrieves the tab, that is the target of a drag'n'drop operation. Its background is drawn
	/// highlighted. If set to \c NULL, no tab is a drop target.
	///
	/// \param[out] ppDropHilitedTab Receives the drop target's \c ITabStripTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_DropHilitedTab, TabStripTab::get_DropHilited, get_AllowDragDrop,
	///     get_RegisterForOLEDragDrop, get_CaretTab, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DropHilitedTab(ITabStripTab** ppDropHilitedTab);
	/// \brief <em>Sets the \c DropHilitedTab property</em>
	///
	/// Sets the tab, that is the target of a drag'n'drop operation. It's background is drawn
	/// highlighted. If set to \c NULL, no tab is a drop target.
	///
	/// \param[in] pNewDropHilitedTab The new drop target's \c ITabStripTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DropHilitedTab, put_AllowDragDrop, put_RegisterForOLEDragDrop, putref_CaretTab,
	///     Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE putref_DropHilitedTab(ITabStripTab* pNewDropHilitedTab);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the control is enabled or disabled for user input. If set to \c VARIANT_TRUE,
	/// it reacts to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Enables or disables the control for user input. If set to \c VARIANT_TRUE, the control reacts
	/// to user input; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FixedTabWidth property</em>
	///
	/// Retrieves the tabs' fixed width in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FixedTabWidth, get_UseFixedTabWidth, get_MinTabWidth, get_TabHeight, get_DisplayAreaWidth
	virtual HRESULT STDMETHODCALLTYPE get_FixedTabWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c FixedTabWidth property</em>
	///
	/// Sets the tabs' fixed width in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FixedTabWidth, put_UseFixedTabWidth, put_MinTabWidth, put_TabHeight
	virtual HRESULT STDMETHODCALLTYPE put_FixedTabWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c FocusOnButtonDown property</em>
	///
	/// Retrieves whether the control receives the keyboard focus if a tab is clicked. If set to
	/// \c VARIANT_TRUE, the keyboard focus is set to the control; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FocusOnButtonDown
	virtual HRESULT STDMETHODCALLTYPE get_FocusOnButtonDown(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c FocusOnButtonDown property</em>
	///
	/// Sets whether the control receives the keyboard focus if a tab is clicked. If set to \c VARIANT_TRUE,
	/// the keyboard focus is set to the control; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FocusOnButtonDown
	virtual HRESULT STDMETHODCALLTYPE put_FocusOnButtonDown(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the tabs' text.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Font, putref_Font, get_UseSystemFont, TabStripTab::get_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_Font(IFontDisp** ppFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the tabs' text.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewFont is cloned.
	///
	/// \sa get_Font, putref_Font, put_UseSystemFont, TabStripTab::put_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_Font(IFontDisp* pNewFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the tabs' text.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Font, put_Font, put_UseSystemFont, TabStripTab::put_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_Font(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c hDragImageList property</em>
	///
	/// Retrieves the handle to the imagelist containing the drag image that is used during a
	/// drag'n'drop operation to visualize the dragged tabs.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, BeginDrag, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_hDragImageList(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hHighResImageList property</em>
	///
	/// Retrieves the handle to the imagelist containing the control's icons that are used when icons with a
	/// high resolution are required. Currently the only usage of this imagelist is the creation of Aero OLE
	/// drag images.\n
	/// If set to \c NULL, the imagelist specified by the \c hImageList property is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \sa put_hHighResImageList, TabStripTab::get_IconIndex, get_hImageList, get_SupportOLEDragImages,
	///     OLEDrag
	virtual HRESULT STDMETHODCALLTYPE get_hHighResImageList(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hHighResImageList property</em>
	///
	/// Sets the handle to the imagelist containing the control's icons that are used when icons with a
	/// high resolution are required. Currently the only usage of this imagelist is the creation of Aero OLE
	/// drag images.\n
	/// If set to \c NULL, the imagelist specified by the \c hImageList property is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \sa get_hHighResImageList, TabStripTab::put_IconIndex, put_hImageList, put_SupportOLEDragImages,
	///     OLEDrag
	virtual HRESULT STDMETHODCALLTYPE put_hHighResImageList(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c hImageList property</em>
	///
	/// Retrieves the handle to the imagelist containing the tabs' icons. If set to NULL, no icons are drawn.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \sa put_hImageList, TabStripTab::get_IconIndex, get_hHighResImageList
	virtual HRESULT STDMETHODCALLTYPE get_hImageList(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hImageList property</em>
	///
	/// Sets the handle to the imagelist containing the tabs' icons. If set to NULL, no icons are drawn.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \sa get_hImageList, TabStripTab::put_IconIndex, put_hHighResImageList
	virtual HRESULT STDMETHODCALLTYPE put_hImageList(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c HorizontalPadding property</em>
	///
	/// Retrieves the amount of space (padding) to the left and right of each tab's icon and label in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HorizontalPadding, get_VerticalPadding, TabStripTab::get_IconIndex, TabStripTab::get_Text
	virtual HRESULT STDMETHODCALLTYPE get_HorizontalPadding(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c HorizontalPadding property</em>
	///
	/// Sets the amount of space (padding) to the left and right of each tab's icon and label in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HorizontalPadding, put_VerticalPadding, TabStripTab::put_IconIndex, TabStripTab::put_Text
	virtual HRESULT STDMETHODCALLTYPE put_HorizontalPadding(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c HotTracking property</em>
	///
	/// Retrieves whether the tab underneath the mouse cursor becomes highlighted. If set to \c VARIANT_TRUE,
	/// the tab becomes highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed tabstrips.
	///
	/// \sa put_HotTracking
	virtual HRESULT STDMETHODCALLTYPE get_HotTracking(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c HotTracking property</em>
	///
	/// Sets whether the tab underneath the mouse cursor becomes highlighted. If set to \c VARIANT_TRUE,
	/// the tab becomes highlighted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed tabstrips.
	///
	/// \sa get_HotTracking
	virtual HRESULT STDMETHODCALLTYPE put_HotTracking(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c HoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE get_HoverTime(LONG* pValue);
	/// \brief <em>Sets the \c HoverTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE put_HoverTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c hWnd property</em>
	///
	/// Retrieves the control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWndArrowButtons, Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndArrowButtons property</em>
	///
	/// Retrieves the window handle of the up-down control that is displayed to let the user scroll the tabs.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWnd, get_hWndToolTip
	virtual HRESULT STDMETHODCALLTYPE get_hWndArrowButtons(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndToolTip property</em>
	///
	/// Retrieves the tooltip control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hWndToolTip, get_hWndArrowButtons, get_ShowToolTips
	virtual HRESULT STDMETHODCALLTYPE get_hWndToolTip(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hWndToolTip property</em>
	///
	/// Sets the tooltip control to use.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_hWndToolTip, put_ShowToolTips
	virtual HRESULT STDMETHODCALLTYPE put_hWndToolTip(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c InsertMarkColor property</em>
	///
	/// Retrieves the color that is the control's insertion mark is drawn in.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_InsertMarkColor, GetInsertMarkPosition
	virtual HRESULT STDMETHODCALLTYPE get_InsertMarkColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c InsertMarkColor property</em>
	///
	/// Sets the color that is the control's insertion mark is drawn in.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_InsertMarkColor, SetInsertMarkPosition
	virtual HRESULT STDMETHODCALLTYPE put_InsertMarkColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the control's release type</em>
	///
	/// Retrieves the control's release type. This property is part of the fingerprint that uniquely
	/// identifies each software written by Timo "TimoSoft" Kunze. If set to \c VARIANT_TRUE, the
	/// control was compiled for release; otherwise it was compiled for debugging.
	///
	/// \param[out] pValue The release type.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_IsRelease(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c MinTabWidth property</em>
	///
	/// Retrieves the tabs' minimum width in pixels. If set to -1, the system's default tab width is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MinTabWidth, get_FixedTabWidth, get_TabHeight, get_DisplayAreaWidth
	virtual HRESULT STDMETHODCALLTYPE get_MinTabWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c MinTabWidth property</em>
	///
	/// Sets the tabs' minimum width in pixels. If set to -1, the system's default tab width is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MinTabWidth, put_FixedTabWidth, put_TabHeight
	virtual HRESULT STDMETHODCALLTYPE put_MinTabWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c MouseIcon property</em>
	///
	/// Retrieves a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[out] ppMouseIcon Receives the picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, TabStripCtlLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, TabStripCtlLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MouseIcon(IPictureDisp** ppMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewMouseIcon is cloned.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, TabStripCtlLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, TabStripCtlLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, TabStripCtlLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, TabStripCtlLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE putref_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Retrieves the current setting of the \c MousePointer property</em>
	///
	/// Retrieves the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MousePointer, get_MouseIcon, TabStripCtlLibU::MousePointerConstants
	/// \else
	///   \sa put_MousePointer, get_MouseIcon, TabStripCtlLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MousePointer(MousePointerConstants* pValue);
	/// \brief <em>Sets the \c MousePointer property</em>
	///
	/// Sets the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MousePointer, put_MouseIcon, TabStripCtlLibU::MousePointerConstants
	/// \else
	///   \sa get_MousePointer, put_MouseIcon, TabStripCtlLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c MultiRow property</em>
	///
	/// Retrieves whether the tabs are displayed in multiple rows if they don't fit next to each other.
	/// If set to \c VARIANT_TRUE, multiple rows are used; otherwise two arrow buttons are displayed that
	/// let the user scroll the tabs.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MultiRow, get_Style, get_TabPlacement, get_UseFixedTabWidth, get_RaggedTabRows, CountTabRows
	virtual HRESULT STDMETHODCALLTYPE get_MultiRow(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c MultiRow property</em>
	///
	/// Sets whether the tabs are displayed in multiple rows if they don't fit next to each other.
	/// If set to \c VARIANT_TRUE, multiple rows are used; otherwise two arrow buttons are displayed that
	/// let the user scroll the tabs.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c VARIANT_FALSE while the \c ScrollToOpposite property is set to
	///          \c VARIANT_TRUE, the \c ScrollToOpposite property will be changed to \c VARIANT_FALSE.\n
	///          If this property is set to \c VARIANT_FALSE while the \c TabPlacement property is set to
	///          \c tpLeft or \c tpRight, the \c TabPlacement property will be changed to \c tpTop.
	///
	/// \sa get_MultiRow, put_ScrollToOpposite, put_Style, put_TabPlacement, put_UseFixedTabWidth,
	///     put_RaggedTabRows, CountTabRows
	virtual HRESULT STDMETHODCALLTYPE put_MultiRow(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c MultiSelect property</em>
	///
	/// Retrieves whether multiple tabs can be selected at the same time. If set to \c VARIANT_TRUE, the
	/// user may select multiple tabs; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MultiSelect, get_Style, get_ActiveTab, get_CaretTab, TabStripTab::get_Selected
	virtual HRESULT STDMETHODCALLTYPE get_MultiSelect(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c MultiSelect property</em>
	///
	/// Sets whether multiple tabs can be selected at the same time. If set to \c VARIANT_TRUE, the
	/// user may select multiple tabs; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Multiselection won't work if the \c Style property is set to \c sTabs.
	///
	/// \sa get_MultiSelect, put_Style, putref_ActiveTab, putref_CaretTab, TabStripTab::put_Selected
	virtual HRESULT STDMETHODCALLTYPE put_MultiSelect(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c OLEDragImageStyle property</em>
	///
	/// Retrieves the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag,
	///       TabStripCtlLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag,
	///       TabStripCtlLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue);
	/// \brief <em>Sets the \c OLEDragImageStyle property</em>
	///
	/// Sets the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag,
	///       TabStripCtlLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag,
	///       TabStripCtlLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OLEDragImageStyle(OLEDragImageStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c OwnerDrawn property</em>
	///
	/// Retrieves whether the client application draws the tabs itself. If set to \c VARIANT_TRUE, the
	/// control will fire the \c OwnerDrawTab event each time a tab must be drawn. If set to
	/// \c VARIANT_FALSE, the control will draw the tabs itself.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_OwnerDrawn, Raise_OwnerDrawTab
	virtual HRESULT STDMETHODCALLTYPE get_OwnerDrawn(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c OwnerDrawn property</em>
	///
	/// Sets whether the client application draws the tabs itself. If set to \c VARIANT_TRUE, the
	/// control will fire the \c OwnerDrawTab event each time a tab must be drawn. If set to
	/// \c VARIANT_FALSE, the control will draw the tabs itself.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_OwnerDrawn, Raise_OwnerDrawTab
	virtual HRESULT STDMETHODCALLTYPE put_OwnerDrawn(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessContextMenuKeys property</em>
	///
	/// Retrieves whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10] or
	/// [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the event is fired; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE get_ProcessContextMenuKeys(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessContextMenuKeys property</em>
	///
	/// Sets whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10] or
	/// [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the event is fired; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE put_ProcessContextMenuKeys(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's programmer(s)</em>
	///
	/// Retrieves the name(s) of the control's programmer(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The programmer.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Programmer(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c RaggedTabRows property</em>
	///
	/// Retrieves whether tabs are stretched so that each tab row in a multi-row control fills up the
	/// control's whole width. If set to \c VARIANT_FALSE, the tabs are stretched; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RaggedTabRows, get_MultiRow, get_UseFixedTabWidth
	virtual HRESULT STDMETHODCALLTYPE get_RaggedTabRows(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RaggedTabRows property</em>
	///
	/// Sets whether tabs are stretched so that each tab row in a multi-row control fills up the
	/// control's whole width. If set to \c VARIANT_FALSE, the tabs are stretched; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RaggedTabRows, put_MultiRow, put_UseFixedTabWidth
	virtual HRESULT STDMETHODCALLTYPE put_RaggedTabRows(VARIANT_BOOL newValue);
	// \brief <em>Retrieves the current setting of the \c ReflectContextMenuMessages property</em>
	//
	// Retrieves whether whether \c WM_CONTEXTMENU messages are reflected back to the child windows they
	// come from.\n
	// Usually controls send the \c WM_CONTEXTMENU message to the parent window, so to trap this message
	// the parent window of the control must be subclassed. This design is great for C++ applications,
	// because in a C++ app, you watch the main window's messages anyway, but usually don't watch a single
	// control's messages. In a VB6 app, things are different. You don't watch the main window's messages
	// by default, but the controls watch their own messages. So in a VB6 app it would be better if the
	// controls would send \c WM_CONTEXTMENU to themselves and not to the parent window.\n
	// TimoSoft controls have a \c ContextMenu event, but this event won't be raised by default, because
	// the \c WM_CONTEXTMENU goes to the wrong window. That's the reason why all samples provided with
	// those controls subclass the Form and reflect \c WM_CONTEXTMENU back to the control. \c TabStrip can
	// make things a bit easier and do the message reflection for any child controls.\n
	// If this property is set to \c VARIANT_TRUE, the \c WM_CONTEXTMENU message gets reflected; otherwise
	// not.
	//
	// \param[out] pValue The property's current setting.
	//
	// \return An \c HRESULT error code.
	//
	// \sa put_ReflectContextMenuMessages, Raise_ContextMenu
	//virtual HRESULT STDMETHODCALLTYPE get_ReflectContextMenuMessages(VARIANT_BOOL* pValue);
	// \brief <em>Sets the \c ReflectContextMenuMessages property</em>
	//
	// Sets whether whether \c WM_CONTEXTMENU messages are reflected back to the child windows they
	// come from.\n
	// Usually controls send the \c WM_CONTEXTMENU message to the parent window, so to trap this message
	// the parent window of the control must be subclassed. This design is great for C++ applications,
	// because in a C++ app, you watch the main window's messages anyway, but usually don't watch a single
	// control's messages. In a VB6 app, things are different. You don't watch the main window's messages
	// by default, but the controls watch their own messages. So in a VB6 app it would be better if the
	// controls would send \c WM_CONTEXTMENU to themselves and not to the parent window.\n
	// TimoSoft controls have a \c ContextMenu event, but this event won't be raised by default, because
	// the \c WM_CONTEXTMENU goes to the wrong window. That's the reason why all samples provided with
	// those controls subclass the Form and reflect \c WM_CONTEXTMENU back to the control. \c TabStrip can
	// make things a bit easier and do the message reflection for any child controls.\n
	// If this property is set to \c VARIANT_TRUE, the \c WM_CONTEXTMENU message gets reflected; otherwise
	// not.
	//
	// \param[in] newValue The setting to apply.
	//
	// \return An \c HRESULT error code.
	//
	// \sa get_ReflectContextMenuMessages, Raise_ContextMenu
	//virtual HRESULT STDMETHODCALLTYPE put_ReflectContextMenuMessages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RegisterForOLEDragDrop property</em>
	///
	/// Retrieves whether the control is registered as a target for OLE drag'n'drop. Any of the
	/// values defined by the \c RegisterForOLEDragDropConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_RegisterForOLEDragDrop, get_AllowDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter,
	///       TabStripCtlLibU::RegisterForOLEDragDropConstants
	/// \else
	///   \sa put_RegisterForOLEDragDrop, get_AllowDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter,
	///       TabStripCtlLibA::RegisterForOLEDragDropConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants* pValue);
	/// \brief <em>Sets the \c RegisterForOLEDragDrop property</em>
	///
	/// Sets whether the control is registered as a target for OLE drag'n'drop. Any of the
	/// values defined by the \c RegisterForOLEDragDropConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_RegisterForOLEDragDrop, put_AllowDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter,
	///       TabStripCtlLibU::RegisterForOLEDragDropConstants
	/// \else
	///   \sa get_RegisterForOLEDragDrop, put_AllowDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter,
	///       TabStripCtlLibA::RegisterForOLEDragDropConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeft property</em>
	///
	/// Retrieves whether bidirectional features are enabled or disabled. Any combination of the values
	/// defined by the \c RightToLeftConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_RightToLeft, TabStripCtlLibU::RightToLeftConstants
	/// \else
	///   \sa put_RightToLeft, TabStripCtlLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeft(RightToLeftConstants* pValue);
	/// \brief <em>Sets the \c RightToLeft property</em>
	///
	/// Enables or disables bidirectional features. Any combination of the values defined by the
	/// \c RightToLeftConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Setting or clearing the \c rtlLayout flag at runtime leads to drawing glitches.
	///
	/// \if UNICODE
	///   \sa get_RightToLeft, TabStripCtlLibU::RightToLeftConstants
	/// \else
	///   \sa get_RightToLeft, TabStripCtlLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(RightToLeftConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ScrollToOpposite property</em>
	///
	/// Retrieves whether unneeded tabs scroll to the opposite side of the control when a tab is selected.
	/// If set to \c VARIANT_TRUE, tabs scroll to the opposite side; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ScrollToOpposite, get_MultiRow, get_TabPlacement
	virtual HRESULT STDMETHODCALLTYPE get_ScrollToOpposite(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ScrollOpposite property</em>
	///
	/// Sets whether unneeded tabs scroll to the opposite side of the control when a tab is selected.
	/// If set to \c VARIANT_TRUE, tabs scroll to the opposite side; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c VARIANT_TRUE while the \c MultiRow property is set to
	///          \c VARIANT_FALSE, the \c MultiRow property will be changed to \c VARIANT_TRUE.\n
	///          If this property is set to \c VARIANT_TRUE while the \c Style property is set to \c sButtons
	///          or \c sFlatButtons, the \c Style property will be changed to \c sTabs.
	///
	/// \attention This property can't be changed at runtime.
	///
	/// \sa get_ScrollToOpposite, put_MultiRow, put_Style, put_TabPlacement
	virtual HRESULT STDMETHODCALLTYPE put_ScrollToOpposite(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowButtonSeparators property</em>
	///
	/// Retrieves whether the control draws separator lines between the tabs if the \c Style property is
	/// set to \c sFlatButtons. If set to \c VARIANT_TRUE, separators are drawn; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowButtonSeparators, get_Style
	virtual HRESULT STDMETHODCALLTYPE get_ShowButtonSeparators(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowButtonSeparators property</em>
	///
	/// Sets whether the control draws separator lines between the tabs if the \c Style property is
	/// set to \c sFlatButtons. If set to \c VARIANT_TRUE, separators are drawn; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowButtonSeparators, put_Style
	virtual HRESULT STDMETHODCALLTYPE put_ShowButtonSeparators(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowDragImage property</em>
	///
	/// Retrieves whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible;
	/// otherwise hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowDragImage, get_hDragImageList, get_SupportOLEDragImages, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_ShowDragImage(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowDragImage property</em>
	///
	/// Sets whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible; otherwise
	/// hidden.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, get_hDragImageList, put_SupportOLEDragImages, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_ShowDragImage(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowToolTips property</em>
	///
	/// Retrieves whether the control displays any tooltips. If set to \c VARIANT_TRUE, the control displays
	/// tab-specific tooltips; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowToolTips, get_hWndToolTip, Raise_TabGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE get_ShowToolTips(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowToolTips property</em>
	///
	/// Enables or disables tooltips. If set to \c VARIANT_TRUE, the control displays tab-specific
	/// tooltips; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control.
	///
	/// \sa get_ShowToolTips, put_hWndToolTip, Raise_TabGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE put_ShowToolTips(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Style property</em>
	///
	/// Retrieves the tabs' style. Any of the values defined by the \c StyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Style, get_Appearance, get_BorderStyle, get_MultiRow, get_TabPlacement,
	///       TabStripCtlLibU::StyleConstants
	/// \else
	///   \sa put_Style, get_Appearance, get_BorderStyle, get_MultiRow, get_TabPlacement,
	///       TabStripCtlLibA::StyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Style(StyleConstants* pValue);
	/// \brief <em>Sets the \c Style property</em>
	///
	/// Sets the tabs' style. Any of the values defined by the \c StyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c sButtons or \c sFlatButtons while the \c ScrollToOpposite
	///          property is set to \c VARIANT_TRUE, the \c ScrollToOpposite property will be changed to
	///          \c VARIANT_FALSE.\n
	///          Changing this property at runtime leads to drawing glitches if the control is themed.
	///
	/// \if UNICODE
	///   \sa get_Style, put_Appearance, put_BorderStyle, put_MultiRow, put_ScrollToOpposite,
	///       put_TabPlacement, TabStripCtlLibU::StyleConstants
	/// \else
	///   \sa get_Style, put_Appearance, put_BorderStyle, put_MultiRow, put_ScrollToOpposite,
	///       put_TabPlacement, TabStripCtlLibA::StyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Style(StyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c SupportOLEDragImages property</em>
	///
	/// Retrieves whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, get_hHighResImageList, get_ShowDragImage,
	///     get_OLEDragImageStyle, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646441.aspx">IDragSourceHelper</a>
	virtual HRESULT STDMETHODCALLTYPE get_SupportOLEDragImages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SupportOLEDragImages property</em>
	///
	/// Sets whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, put_hHighResImageList, put_ShowDragImage,
	///     put_OLEDragImageStyle, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646441.aspx">IDragSourceHelper</a>
	virtual HRESULT STDMETHODCALLTYPE put_SupportOLEDragImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c TabBoundingBoxDefinition property</em>
	///
	/// Retrieves the parts of a tab that get handled as such if firing any kind of mouse event. Any
	/// combination of the values defined by the \c TabBoundingBoxDefinitionConstants enumeration is
	/// valid. E. g. if set to \c tbbdTabIcon, the \c tsTab parameter of the \c MouseMove event will
	/// identify the tab only if the mouse cursor is located over the tab's icon; otherwise (e. g. if the
	/// cursor is located over the tab's text) it will be \c NULL.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_TabBoundingBoxDefinition, HitTest, TabStripCtlLibU::TabBoundingBoxDefinitionConstants
	/// \else
	///   \sa put_TabBoundingBoxDefinition, HitTest, TabStripCtlLibA::TabBoundingBoxDefinitionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TabBoundingBoxDefinition(TabBoundingBoxDefinitionConstants* pValue);
	/// \brief <em>Sets the \c TabBoundingBoxDefinition property</em>
	///
	/// Sets the parts of a tab that get handled as such if firing any kind of mouse event. Any combination
	/// of the values defined by the \c TabBoundingBoxDefinitionConstants enumeration is valid. E. g. if set
	/// to \c tbbdTabIcon, the \c tsTab parameter of the \c MouseMove event will identify the tab only if the
	/// mouse cursor is located over the tab's icon; otherwise (e. g. if the cursor is located over the tab's
	/// text) it will be \c NULL.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_TabBoundingBoxDefinition, HitTest, TabStripCtlLibU::TabBoundingBoxDefinitionConstants
	/// \else
	///   \sa get_TabBoundingBoxDefinition, HitTest, TabStripCtlLibA::TabBoundingBoxDefinitionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TabBoundingBoxDefinition(TabBoundingBoxDefinitionConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c TabCaptionAlignment property</em>
	///
	/// Retrieves the alignment of the tabs' captions. Any of the values defined by the
	/// \c TabCaptionAlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_TabCaptionAlignment, get_UseFixedTabWidth, TabStripTab::get_IconIndex,
	///       TabStripCtlLibU::TabCaptionAlignmentConstants
	/// \else
	///   \sa put_TabCaptionAlignment, get_UseFixedTabWidth, TabStripTab::get_IconIndex,
	///       TabStripCtlLibA::TabCaptionAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TabCaptionAlignment(TabCaptionAlignmentConstants* pValue);
	/// \brief <em>Sets the \c TabCaptionAlignment property</em>
	///
	/// Sets the alignment of the tabs' captions. Any of the values defined by the
	/// \c TabCaptionAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c UseFixedTabWidth property is set to \c VARIANT_FALSE.
	///
	/// \if UNICODE
	///   \sa get_TabCaptionAlignment, put_UseFixedTabWidth, TabStripTab::put_IconIndex,
	///       TabStripCtlLibU::TabCaptionAlignmentConstants
	/// \else
	///   \sa get_TabCaptionAlignment, put_UseFixedTabWidth, TabStripTab::put_IconIndex,
	///       TabStripCtlLibA::TabCaptionAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TabCaptionAlignment(TabCaptionAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c TabHeight property</em>
	///
	/// Retrieves the tabs' height in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_TabHeight, get_FixedTabWidth, get_MinTabWidth, get_DisplayAreaHeight
	virtual HRESULT STDMETHODCALLTYPE get_TabHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c TabHeight property</em>
	///
	/// Sets the tabs' height in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_TabHeight, put_FixedTabWidth, put_MinTabWidth
	virtual HRESULT STDMETHODCALLTYPE put_TabHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c TabPlacement property</em>
	///
	/// Retrieves the side of the control at which the tabs are displayed. Any of the values defined by the
	/// \c TabPlacementConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_TabPlacement, get_MultiRow, get_Style, TabStripCtlLibU::TabPlacementConstants
	/// \else
	///   \sa put_TabPlacement, get_MultiRow, get_Style, TabStripCtlLibA::TabPlacementConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TabPlacement(TabPlacementConstants* pValue);
	/// \brief <em>Sets the \c TabPlacement property</em>
	///
	/// Sets the side of the control at which the tabs are displayed. Any of the values defined by the
	/// \c TabPlacementConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c tpLeft or \c tpRight while the \c MultiRow property is set
	///          to \c VARIANT_FALSE, the \c MultiRow property will be changed to \c VARIANT_TRUE.
	///
	/// \attention This property can't be changed from or to \c tpLeft or \c tpRight at runtime.
	///
	/// \if UNICODE
	///   \sa get_TabPlacement, put_MultiRow, put_Style, TabStripCtlLibU::TabPlacementConstants
	/// \else
	///   \sa get_TabPlacement, put_MultiRow, put_Style, TabStripCtlLibA::TabPlacementConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TabPlacement(TabPlacementConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Tabs property</em>
	///
	/// Retrieves a collection object wrapping the control's tabs.
	///
	/// \param[out] ppTabs Receives the collection object's \c ITabStripTabs implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TabStripTabs
	virtual HRESULT STDMETHODCALLTYPE get_Tabs(ITabStripTabs** ppTabs);
	/// \brief <em>Retrieves the name(s) of the control's tester(s)</em>
	///
	/// Retrieves the name(s) of the control's tester(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The name(s) of the tester(s).
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease,
	///     get_Programmer
	virtual HRESULT STDMETHODCALLTYPE get_Tester(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c UseFixedTabWidth property</em>
	///
	/// Retrieves whether all tabs are the same width. If set to \c VARIANT_TRUE, the tabs are the same
	/// width; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseFixedTabWidth, get_FixedTabWidth, get_RaggedTabRows, get_TabCaptionAlignment
	virtual HRESULT STDMETHODCALLTYPE get_UseFixedTabWidth(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseFixedTabWidth property</em>
	///
	/// Sets whether all tabs are the same width. If set to \c VARIANT_TRUE, the tabs are the same
	/// width; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseFixedTabWidth, put_FixedTabWidth, put_RaggedTabRows, put_TabCaptionAlignment
	virtual HRESULT STDMETHODCALLTYPE put_UseFixedTabWidth(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseSystemFont property</em>
	///
	/// Retrieves whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseSystemFont, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_UseSystemFont(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseSystemFont property</em>
	///
	/// Sets whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseSystemFont, put_Font, putref_Font
	virtual HRESULT STDMETHODCALLTYPE put_UseSystemFont(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Version(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c VerticalPadding property</em>
	///
	/// Retrieves the amount of space (padding) above and below each tab's icon and label in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_VerticalPadding, get_HorizontalPadding, TabStripTab::get_IconIndex, TabStripTab::get_Text
	virtual HRESULT STDMETHODCALLTYPE get_VerticalPadding(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c VerticalPadding property</em>
	///
	/// Sets the amount of space (padding) above and below each tab's icon and label in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_VerticalPadding, put_HorizontalPadding, TabStripTab::put_IconIndex, TabStripTab::put_Text
	virtual HRESULT STDMETHODCALLTYPE put_VerticalPadding(OLE_YSIZE_PIXELS newValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Enters drag'n'drop mode</em>
	///
	/// \param[in] pDraggedTabs The dragged tabs collection object's \c ITabStripTabContainer
	///            implementation.
	/// \param[in] hDragImageList The imagelist containing the drag image that shall be used to
	///            visualize the drag'n'drop operation. If -1, the method creates the drag image itself;
	///            if \c NULL, no drag image is used.
	/// \param[in,out] pXHotSpot The x-coordinate (in pixels) of the drag image's hotspot relative to the
	///                drag image's upper-left corner. If the \c hDragImageList parameter is set to -1 or
	///                \c NULL, this parameter is ignored. This parameter will be changed to the value that
	///                finally was used by the method.
	/// \param[in,out] pYHotSpot The y-coordinate (in pixels) of the drag image's hotspot relative to the
	///                drag image's upper-left corner. If the \c hDragImageList parameter is set to -1 or
	///                \c NULL, this parameter is ignored. This parameter will be changed to the value that
	///                finally was used by the method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OLEDrag, get_DraggedTabs, EndDrag, get_hDragImageList, Raise_TabBeginDrag,
	///     Raise_TabBeginRDrag, TabStripTab::CreateDragImage, TabStripTabContainer::CreateDragImage
	virtual HRESULT STDMETHODCALLTYPE BeginDrag(ITabStripTabContainer* pDraggedTabs, OLE_HANDLE hDragImageList = NULL, OLE_XPOS_PIXELS* pXHotSpot = NULL, OLE_YPOS_PIXELS* pYHotSpot = NULL);
	/// \brief <em>Calculates a display area based on a window rectangle</em>
	///
	/// Calculates the display area, that a tabstrip control with the specified window rectangle would have.
	///
	/// \param[in] pWindowRectangle The window rectangle (in pixels) for which the display area shall be
	///            calculated.
	/// \param[in,out] pDisplayArea The calculated display area rectangle (in pixels).
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CalculateWindowRectangle, get_DisplayAreaHeight, get_DisplayAreaLeft, get_DisplayAreaTop,
	///       get_DisplayAreaWidth, TabStripCtlLibU::RECTANGLE
	/// \else
	///   \sa CalculateWindowRectangle, get_DisplayAreaHeight, get_DisplayAreaLeft, get_DisplayAreaTop,
	///       get_DisplayAreaWidth, TabStripCtlLibA::RECTANGLE
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE CalculateDisplayArea(RECTANGLE* pWindowRectangle, RECTANGLE* pDisplayArea);
	/// \brief <em>Calculates a window rectangle based on a display area</em>
	///
	/// Calculates the window rectangle, that a tabstrip control with the specified display area would have.
	///
	/// \param[in] pDisplayArea The display area rectangle (in pixels) for which the window rectangle shall
	///            be calculated.
	/// \param[in,out] pWindowRectangle The calculated window rectangle (in pixels).
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CalculateDisplayArea, get_DisplayAreaHeight, get_DisplayAreaLeft, get_DisplayAreaTop,
	///       get_DisplayAreaWidth, TabStripCtlLibU::RECTANGLE
	/// \else
	///   \sa CalculateDisplayArea, get_DisplayAreaHeight, get_DisplayAreaLeft, get_DisplayAreaTop,
	///       get_DisplayAreaWidth, TabStripCtlLibA::RECTANGLE
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE CalculateWindowRectangle(RECTANGLE* pDisplayArea, RECTANGLE* pWindowRectangle);
	/// \brief <em>Retrieves the current number of tab rows in the control</em>
	///
	/// \param[out] pValue The current number of tab rows.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MultiRow
	virtual HRESULT STDMETHODCALLTYPE CountTabRows(LONG* pValue);
	/// \brief <em>Creates a new \c TabStripTabContainer object</em>
	///
	/// Retrieves a new \c TabStripTabContainer object and fills it with the specified tabs.
	///
	/// \param[in] tabs The tab(s) to add to the collection. May be either \c Empty, a tab ID, a
	///            \c TabStripTab object or a \c TabStripTabs collection.
	/// \param[out] ppContainer Receives the new collection object's \c ITabStripTabContainer
	///             implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa TabStripTabContainer::Clone, TabStripTabContainer::Add
	virtual HRESULT STDMETHODCALLTYPE CreateTabContainer(VARIANT tabs = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), ITabStripTabContainer** ppContainer = NULL);
	/// \brief <em>Exits drag'n'drop mode</em>
	///
	/// \param[in] abort If \c VARIANT_TRUE, the drag'n'drop operation will be handled as aborted;
	///            otherwise it will be handled as a drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DraggedTabs, BeginDrag, Raise_AbortedDrag, Raise_Drop
	virtual HRESULT STDMETHODCALLTYPE EndDrag(VARIANT_BOOL abort);
	/// \brief <em>Finishes a pending drop operation</em>
	///
	/// During a drag'n'drop operation the drag image is displayed until the \c OLEDragDrop event has been
	/// handled. This order is intended by Microsoft Windows. However, if a message box is displayed from
	/// within the \c OLEDragDrop event, or the drop operation cannot be performed asynchronously and takes
	/// a long time, it may be desirable to remove the drag image earlier.\n
	/// This method will break the intended order and finish the drag'n'drop operation (including removal
	/// of the drag image) immediately.
	///
	/// \remarks This method will fail if not called from the \c OLEDragDrop event handler or if no drag
	///          images are used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_OLEDragDrop, get_SupportOLEDragImages
	virtual HRESULT STDMETHODCALLTYPE FinishOLEDragDrop(void);
	/// \brief <em>Proposes a position for the control's insertion mark</em>
	///
	/// Retrieves the insertion mark position that is closest to the specified point.
	///
	/// \param[in] x The x-coordinate (in pixels) of the point for which to retrieve the closest
	///            insertion mark position. It must be relative to the control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point for which to retrieve the closest
	///            insertion mark position. It must be relative to the control's upper-left corner.
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified tab. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] ppTSTab Receives the \c ITabStripTab implementation of the tab, at which the
	///             insertion mark should be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetInsertMarkPosition, GetInsertMarkPosition, TabStripCtlLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetInsertMarkPosition, GetInsertMarkPosition, TabStripCtlLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetClosestInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, ITabStripTab** ppTSTab);
	/// \brief <em>Retrieves the position of the control's insertion mark</em>
	///
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified tab. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] ppTSTab Receives the \c ITabStripTab implementation of the tab, at which the insertion
	///             mark is displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetInsertMarkPosition, GetClosestInsertMarkPosition,
	///       TabStripCtlLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetInsertMarkPosition, GetClosestInsertMarkPosition,
	///       TabStripCtlLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, ITabStripTab** ppTSTab);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in,out] pHitTestDetails Receives a value specifying the exact part of the control the
	///                specified point lies in. Any of the values defined by the \c HitTestConstants
	///                enumeration is valid.
	/// \param[out] ppHitTab Receives the "hit" tab's \c ITabStripTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_TabBoundingBoxDefinition, TabStripCtlLibU::HitTestConstants, HitTest
	/// \else
	///   \sa get_TabBoundingBoxDefinition, TabStripCtlLibA::HitTestConstants, HitTest
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, ITabStripTab** ppHitTab);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Enters OLE drag'n'drop mode</em>
	///
	/// \param[in] pDataObject A pointer to the \c IDataObject implementation to use during OLE
	///            drag'n'drop. If not specified, the control's own implementation is used.
	/// \param[in] supportedEffects A bit field defining all drop effects the client wants to support.
	///            Any combination of the values defined by the \c OLEDropEffectConstants enumeration
	///            (except \c odeScroll) is valid.
	/// \param[in] hWndToAskForDragImage The handle of the window, that is awaiting the
	///            \c DI_GETDRAGIMAGE message to specify the drag image to use. If -1, the method
	///            creates the drag image itself. If \c SupportOLEDragImages is set to \c VARIANT_FALSE,
	///            no drag image is used.
	/// \param[in] pDraggedTabs The dragged tabs collection object's \c ITabStripTabContainer
	///            implementation. It is used to generate the drag image and is ignored if
	///            \c hWndToAskForDragImage is not -1.
	/// \param[in] itemCountToDisplay The number to display in the item count label of Aero drag images.
	///            If set to 0 or 1, no item count label is displayed. If set to -1, the number of tabs
	///            contained in the \c pDraggedTabs collection is displayed in the item count label. If
	///            set to any value larger than 1, this value is displayed in the item count label.
	/// \param[out] pPerformedEffects The performed drop effect. Any of the values defined by the
	///             \c OLEDropEffectConstants enumeration (except \c odeScroll) is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa BeginDrag, Raise_TabBeginDrag, Raise_TabBeginRDrag, Raise_OLEStartDrag,
	///       get_SupportOLEDragImages, get_OLEDragImageStyle, TabStripCtlLibU::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \else
	///   \sa BeginDrag, Raise_TabBeginDrag, Raise_TabBeginRDrag, Raise_OLEStartDrag,
	///       get_SupportOLEDragImages, get_OLEDragImageStyle, TabStripCtlLibA::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE OLEDrag(LONG* pDataObject = NULL, OLEDropEffectConstants supportedEffects = odeCopyOrMove, OLE_HANDLE hWndToAskForDragImage = -1, ITabStripTabContainer* pDraggedTabs = NULL, LONG itemCountToDisplay = -1, OLEDropEffectConstants* pPerformedEffects = NULL);
	/// \brief <em>Advises the control to redraw itself</em>
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Refresh(void);
	/// \brief <em>Saves the control's settings to the specified file</em>
	///
	/// \param[in] file The file to write to.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadSettingsFromFile
	virtual HRESULT STDMETHODCALLTYPE SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Sets the position of the control's insertion mark</em>
	///
	/// \param[in] relativePosition The insertion mark's position relative to the specified tab. Any of
	///            the values defined by the \c InsertMarkPositionConstants enumeration is valid.
	/// \param[in] pTSTab The \c ITabStripTab implementation of the tab, at which the insertion
	///            mark will be displayed. If set to \c NULL, the insertion mark will be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa GetInsertMarkPosition, GetClosestInsertMarkPosition, get_InsertMarkColor, get_AllowDragDrop,
	///       get_RegisterForOLEDragDrop, TabStripCtlLibU::InsertMarkPositionConstants
	/// \else
	///   \sa GetInsertMarkPosition, GetClosestInsertMarkPosition, get_InsertMarkColor, get_AllowDragDrop,
	///       get_RegisterForOLEDragDrop, TabStripCtlLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, ITabStripTab* pTSTab);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Property object changes
	///
	//@{
	/// \brief <em>Will be called after a property object was changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that was changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnChanged
	HRESULT OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/);
	/// \brief <em>Will be called before a property object is changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that is about to be changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnRequestEdit
	HRESULT OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Called to create the control window</em>
	///
	/// Called to create the control window. This method overrides \c CWindowImpl::Create() and is
	/// used to customize the window styles.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] rect The control's bounding rectangle.
	/// \param[in] szWindowName The control's window name.
	/// \param[in] dwStyle The control's window style. Will be ignored.
	/// \param[in] dwExStyle The control's extended window style. Will be ignored.
	/// \param[in] MenuOrID The control's ID.
	/// \param[in] lpCreateParam The window creation data. Will be passed to the created window.
	///
	/// \return The created window's handle.
	///
	/// \sa OnCreate, GetStyleBits, GetExStyleBits
	HWND Create(HWND hWndParent, ATL::_U_RECT rect = NULL, LPCTSTR szWindowName = NULL, DWORD dwStyle = 0, DWORD dwExStyle = 0, ATL::_U_MENUorID MenuOrID = 0U, LPVOID lpCreateParam = NULL);
	/// \brief <em>Called to draw the control</em>
	///
	/// Called to draw the control. This method overrides \c CComControlBase::OnDraw() and is used to prevent
	/// the "ATL 9.0" drawing in user mode and replace it in design mode.
	///
	/// \param[in] drawInfo Contains any details like the device context required for drawing.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT OnDraw(ATL_DRAWINFO& drawInfo);
	/// \brief <em>Called after receiving the last message (typically \c WM_NCDESTROY)</em>
	///
	/// \param[in] hWnd The window being destroyed.
	///
	/// \sa OnCreate, OnDestroy
	void OnFinalMessage(HWND /*hWnd*/);
	/// \brief <em>Informs an embedded object of its display location within its container</em>
	///
	/// \param[in] pClientSite The \c IOleClientSite implementation of the container application's
	///            client side.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms684013.aspx">IOleObject::SetClientSite</a>
	virtual HRESULT STDMETHODCALLTYPE IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite);
	/// \brief <em>Notifies the control when the container's document window is activated or deactivated</em>
	///
	/// ATL's implementation of \c OnDocWindowActivate calls \c IOleInPlaceObject_UIDeactivate if the control
	/// is deactivated. This causes a bug in MDI apps. If the control sits on a \c MDI child window and has
	/// the focus and the title bar of another top-level window (not the MDI parent window) of the same
	/// process is clicked, the focus is moved from the ATL based ActiveX control to the next control on the
	/// MDI child before it is moved to the other top-level window that was clicked. If the focus is set back
	/// to the MDI child, the ATL based control no longer has the focus.
	///
	/// \param[in] fActivate If \c TRUE, the document window is activated; otherwise deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/0kz79wfc.aspx">IOleInPlaceActiveObjectImpl::OnDocWindowActivate</a>
	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL /*fActivate*/);

	/// \brief <em>A keyboard or mouse message needs to be translated</em>
	///
	/// The control's container calls this method if it receives a keyboard or mouse message. It gives
	/// us the chance to customize keystroke translation (i. e. to react to them in a non-default way).
	/// This method overrides \c CComControlBase::PreTranslateAccelerator.
	///
	/// \param[in] pMessage A \c MSG structure containing details about the received window message.
	/// \param[out] hReturnValue A reference parameter of type \c HRESULT which will be set to \c S_OK,
	///             if the message was translated, and to \c S_FALSE otherwise.
	///
	/// \return \c FALSE if the object's container should translate the message; otherwise \c TRUE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/hxa56938.aspx">CComControlBase::PreTranslateAccelerator</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646373.aspx">TranslateAccelerator</a>
	BOOL PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue);

	//////////////////////////////////////////////////////////////////////
	/// \name Drag image creation
	///
	//@{
	/// \brief <em>Creates a legacy drag image for the specified tab</em>
	///
	/// Creates a drag image for the specified tab in the style of Windows versions prior to Vista. The
	/// drag image is added to an imagelist which is returned.
	///
	/// \param[in] tabIndex The tab for which to create the drag image.
	/// \param[out] pUpperLeftPoint Receives the coordinates (in pixels) of the drag image's upper-left
	///             corner relative to the control's upper-left corner.
	/// \param[out] pBoundingRectangle Receives the drag image's bounding rectangle (in pixels) relative to
	///             the control's upper-left corner.
	///
	/// \return An imagelist containing the drag image.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa OnCreateDragImage, CreateLegacyOLEDragImage, TabStripTabContainer::CreateDragImage
	HIMAGELIST CreateLegacyDragImage(int tabIndex, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle);
	/// \brief <em>Creates a legacy OLE drag image for the specified tabs</em>
	///
	/// Creates an OLE drag image for the specified tabs in the style of Windows versions prior to Vista.
	///
	/// \param[in] pTabs A \c ITabStripTabContainer object wrapping the tabs for which to create the drag
	///            image.
	/// \param[out] pDragImage Receives the drag image including transparency information and the coordinates
	///             (in pixels) of the drag image's upper-left corner relative to the control's upper-left
	///             corner.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa OnGetDragImage, CreateVistaOLEDragImage, CreateLegacyDragImage, TabStripTabContainer,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759778.aspx">SHDRAGIMAGE</a>
	BOOL CreateLegacyOLEDragImage(ITabStripTabContainer* pTabs, __in LPSHDRAGIMAGE pDragImage);
	/// \brief <em>Creates a Vista-like OLE drag image for the specified tabs</em>
	///
	/// Creates an OLE drag image for the specified tabs in the style of Windows Vista and newer.
	///
	/// \param[in] pTabs A \c ITabStripTabContainer object wrapping the tabs for which to create the drag
	///            image.
	/// \param[out] pDragImage Receives the drag image including transparency information and the coordinates
	///             (in pixels) of the drag image's upper-left corner relative to the control's upper-left
	///             corner.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa OnGetDragImage, CreateLegacyOLEDragImage, CreateLegacyDragImage, TabStripTabContainer,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759778.aspx">SHDRAGIMAGE</a>
	BOOL CreateVistaOLEDragImage(ITabStripTabContainer* pTabs, __in LPSHDRAGIMAGE pDragImage);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropTarget
	///
	//@{
	/// \brief <em>Indicates whether a drop can be accepted, and, if so, the effect of the drop</em>
	///
	/// This method is called by the \c DoDragDrop function to determine the target's preferred drop
	/// effect the first time the user moves the mouse into the control during OLE drag'n'Drop. The
	/// target communicates the \c DoDragDrop function the drop effect it wants to be used on drop.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the dragged
	///            data.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragOver, DragLeave, Drop, Raise_OLEDragEnter,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680106.aspx">IDropTarget::DragEnter</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Notifies the target that it no longer is the target of the current OLE drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse out of the
	/// control during OLE drag'n'Drop or if the user canceled the operation. The target must release
	/// any references it holds to the data object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, Drop, Raise_OLEDragLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680110.aspx">IDropTarget::DragLeave</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeave(void);
	/// \brief <em>Communicates the current drop effect to the \c DoDragDrop function</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse over the
	/// control during OLE drag'n'Drop. The target communicates the \c DoDragDrop function the drop
	/// effect it wants to be used on drop.
	///
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, the current drop effect (defined by the \c DROPEFFECT
	///                enumeration). On return, this paramter must be set to the drop effect that the
	///                target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragLeave, Drop, Raise_OLEDragMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680129.aspx">IDropTarget::DragOver</a>
	virtual HRESULT STDMETHODCALLTYPE DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Incorporates the source data into the target and completes the drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user completes the drag'n'drop
	/// operation. The target must incorporate the dragged data into itself and pass the used drop
	/// effect back to the \c DoDragDrop function. The target must release any references it holds to
	/// the data object.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the data
	///            to transfer.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target finally executed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, DragLeave, Raise_OLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
	virtual HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSource
	///
	//@{
	/// \brief <em>Notifies the source of the current drop effect during OLE drag'n'drop</em>
	///
	/// This method is called frequently by the \c DoDragDrop function to notify the source of the
	/// last drop effect that the target has chosen. The source should set an appropriate mouse cursor.
	///
	/// \param[in] effect The drop effect chosen by the target. Any of the values defined by the
	///            \c DROPEFFECT enumeration is valid.
	///
	/// \return \c S_OK if the method has set a custom mouse cursor; \c DRAGDROP_S_USEDEFAULTCURSORS to
	///         use default mouse cursors; or an error code otherwise.
	///
	/// \sa QueryContinueDrag, Raise_OLEGiveFeedback,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693723.aspx">IDropSource::GiveFeedback</a>
	virtual HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD effect);
	/// \brief <em>Determines whether a drag'n'drop operation should be continued, canceled or completed</em>
	///
	/// This method is called by the \c DoDragDrop function to determine whether a drag'n'drop
	/// operation should be continued, canceled or completed.
	///
	/// \param[in] pressedEscape Indicates whether the user has pressed the \c ESC key since the
	///            previous call of this method. If \c TRUE, the key has been pressed; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	///
	/// \return \c S_OK if the drag'n'drop operation should continue; \c DRAGDROP_S_DROP if it should
	///         be completed; \c DRAGDROP_S_CANCEL if it should be canceled; or an error code otherwise.
	///
	/// \sa GiveFeedback, Raise_OLEQueryContinueDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690076.aspx">IDropSource::QueryContinueDrag</a>
	virtual HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL pressedEscape, DWORD keyState);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSourceNotify
	///
	//@{
	/// \brief <em>Notifies the source that the user drags the mouse cursor into a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor into a potential drop target window.
	///
	/// \param[in] hWndTarget The potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragLeaveTarget, Raise_OLEDragEnterPotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragEnterTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnterTarget(HWND hWndTarget);
	/// \brief <em>Notifies the source that the user drags the mouse cursor out of a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor out of a potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnterTarget, Raise_OLEDragLeavePotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragLeaveTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeaveTarget(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICategorizeProperties
	///
	//@{
	/// \brief <em>Retrieves a category's name</em>
	///
	/// \param[in] category The ID of the category whose name is requested.
	// \param[in] languageID The locale identifier identifying the language in which name should be
	//            provided.
	/// \param[out] pName The category's name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::GetCategoryName
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName);
	/// \brief <em>Maps a property to a category</em>
	///
	/// \param[in] property The ID of the property whose category is requested.
	/// \param[out] pCategory The category's ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::MapPropertyToCategory
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory(DISPID property, PROPCAT* pCategory);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICreditsProvider
	///
	//@{
	/// \brief <em>Retrieves the name of the control's authors</em>
	///
	/// \return A string containing the names of all authors.
	CAtlString GetAuthors(void);
	/// \brief <em>Retrieves the URL of the website that has information about the control</em>
	///
	/// \return A string containing the URL.
	CAtlString GetHomepage(void);
	/// \brief <em>Retrieves the URL of the website where users can donate via Paypal</em>
	///
	/// \return A string containing the URL.
	CAtlString GetPaypalLink(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank especially</em>
	///
	/// \return A string containing the special thanks.
	CAtlString GetSpecialThanks(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank</em>
	///
	/// \return A string containing the thanks.
	CAtlString GetThanks(void);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \return A string containing the version.
	CAtlString GetVersion(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPerPropertyBrowsing
	///
	//@{
	/// \brief <em>A displayable string for a property's current value is required</em>
	///
	/// This method is called if the caller's user interface requests a user-friendly description of the
	/// specified property's current value that may be displayed instead of the value itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display name is requested.
	/// \param[out] pDescription The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688734.aspx">IPerPropertyBrowsing::GetDisplayString</a>
	virtual HRESULT STDMETHODCALLTYPE GetDisplayString(DISPID property, BSTR* pDescription);
	/// \brief <em>Displayable strings for a property's predefined values are required</em>
	///
	/// This method is called if the caller's user interface requests user-friendly descriptions of the
	/// specified property's predefined values that may be displayed instead of the values itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display names are requested.
	/// \param[in,out] pDescriptions A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of strings. This array will be
	///                filled with the display name strings.
	/// \param[in,out] pCookies A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of \c DWORD values. Each \c DWORD
	///                value identifies a predefined value of the property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedValue, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms679724.aspx">IPerPropertyBrowsing::GetPredefinedStrings</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies);
	/// \brief <em>A property's predefined value identified by a token is required</em>
	///
	/// This method is called if the caller's user interface requires a property's predefined value that
	/// it has the token of. The token was returned by the \c GetPredefinedStrings method.
	/// We use this method for enumeration-type properties to transform strings like "1 - At Root"
	/// back to the underlying enumeration value (here: \c lsLinesAtRoot).
	///
	/// \param[in] property The ID of the property for which a predefined value is requested.
	/// \param[in] cookie Token identifying which value to return. The token was previously returned
	///            in the \c pCookies array filled by \c IPerPropertyBrowsing::GetPredefinedStrings.
	/// \param[out] pPropertyValue A \c VARIANT that will receive the predefined value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690401.aspx">IPerPropertyBrowsing::GetPredefinedValue</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue);
	/// \brief <em>A property's property page is required</em>
	///
	/// This method is called to request the \c CLSID of the property page used to edit the specified
	/// property.
	///
	/// \param[in] property The ID of the property whose property page is requested.
	/// \param[out] pPropertyPage The property page's \c CLSID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms694476.aspx">IPerPropertyBrowsing::MapPropertyToPage</a>
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToPage(DISPID property, CLSID* pPropertyPage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves a displayable string for a specified setting of a specified property</em>
	///
	/// Retrieves a user-friendly description of the specified property's specified setting. This
	/// description may be displayed by the caller's user interface instead of the setting itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property for which to retrieve the display name.
	/// \param[in] cookie Token identifying the setting for which to retrieve a description.
	/// \param[out] description The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedStrings, GetResStringWithNumber
	HRESULT GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISpecifyPropertyPages
	///
	//@{
	/// \brief <em>The property pages to show are required</em>
	///
	/// This method is called if the property pages, that may be displayed for this object, are required.
	///
	/// \param[out] pPropertyPages A caller-allocated, counted array structure containing the element
	///             count and address of a callee-allocated array of \c GUID structures. Each \c GUID
	///             structure identifies a property page to display.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CommonProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687276.aspx">ISpecifyPropertyPages::GetPages</a>
	virtual HRESULT STDMETHODCALLTYPE GetPages(CAUUID* pPropertyPages);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c KeyPress event.
	///
	/// \sa OnKeyDown, Raise_KeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnRButtonDown, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnContextMenu(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CREATE handler</em>
	///
	/// Will be called right after the control was created.
	/// We use this handler to subclass the parent window, to configure the control window and to raise the
	/// \c RecreatedControlWindow event.
	///
	/// \sa OnDestroy, OnFinalMessage, Raise_RecreatedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632619.aspx">WM_CREATE</a>
	LRESULT OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control is being destroyed.
	/// We use this handler to tidy up and to raise the \c DestroyedControlWindow event.
	///
	/// \sa OnCreate, OnFinalMessage, Raise_DestroyedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_HSCROLL handler</em>
	///
	/// Will be called if the tabs are scrolled.
	/// We use this handler to force a redraw in case the \c CloseableTabs property is set to
	/// \c VARIANT_TRUE.
	///
	/// \sa get_CloseableTabs,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787575.aspx">WM_HSCROLL</a>
	LRESULT OnHScroll(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyDown event.
	///
	/// \sa OnKeyUp, OnChar, Raise_KeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnKeyDown, OnChar, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the left mouse
	/// button.
	/// We use this handler to raise the \c DblClick event.
	///
	/// \sa OnMButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_DblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnClickNotification, OnMButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnLButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEACTIVATE handler</em>
	///
	/// Will be called if the control is inactive and the user clicked in its client area.
	/// We use this handler to hide the message from \c CComControl, so that activating another tab doesn't
	/// set the focus to the control.
	///
	/// \sa get_FocusOnButtonDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645612.aspx">WM_MOUSEACTIVATE</a>
	LRESULT OnMouseActivate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the control's client area for a previously
	/// specified number of milliseconds.
	/// We use this handler to raise the \c MouseHover event.
	///
	/// \sa Raise_MouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NCHITTEST handler</em>
	///
	/// Will be called if the Windows window engine needs to know what the specified point belongs to.
	/// We use this handler to support mouse events and for proper behavior in the VB6 IDE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms645618.aspx">WM_NCHITTEST</a>
	LRESULT OnNCHitTest(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_PAINT and \c WM_PRINTCLIENT handler</em>
	///
	/// Will be called if the control needs to be drawn.
	/// We use this handler to avoid the control being drawn by \c CComControl. This makes Vista's graphic
	/// effects work.
	///
	/// \sa DrawCloseButtonsAndInsertionMarks,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534913.aspx">WM_PRINTCLIENT</a>
	LRESULT OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the right mouse
	/// button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnXButtonDblClk, Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnRClickNotification, OnLButtonDown, OnMButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFONT handler</em>
	///
	/// Will be called if the control's font shall be set.
	/// We use this handler to synchronize our font settings with the new font.
	///
	/// \sa get_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632642.aspx">WM_SETFONT</a>
	LRESULT OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	// \brief <em>\c WM_SETREDRAW handler</em>
	//
	// Will be called if the control's redraw state shall be changed.
	// We use this handler for proper handling of the \c DontRedraw property.
	//
	// \sa get_DontRedraw,
	//     <a href="https://msdn.microsoft.com/en-us/library/ms534853.aspx">WM_SETREDRAW</a>
	//LRESULT OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETTINGCHANGE handler</em>
	///
	/// Will be called if a system setting was changed.
	/// We use this handler to update our appearance.
	///
	/// \attention This message is posted to top-level windows only, so actually we'll never receive it.
	///
	/// \sa OnThemeChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms725497.aspx">WM_SETTINGCHANGE</a>
	LRESULT OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_THEMECHANGED handler</em>
	///
	/// Will be called on themable systems if the theme was changed.
	/// We use this handler to set the \c usingThemes, which is used during drag-image creation.
	///
	/// \sa OnSettingChange, Flags::usingThemes,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632650.aspx">WM_THEMECHANGED</a>
	LRESULT OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_TIMER handler</em>
	///
	/// Will be called when a timer expires that's associated with the control.
	/// We use this handler mainly for the auto-scroll feature.
	///
	/// \sa Raise_OLEDragMouseEnter, Raise_OLEDragMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644902.aspx">WM_TIMER</a>
	LRESULT OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_WINDOWPOSCHANGED handler</em>
	///
	/// Will be called if the control was moved.
	/// We use this handler to resize the control on COM level and to raise the \c ResizedControlWindow
	/// event.
	///
	/// \sa Raise_ResizedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGED</a>
	LRESULT OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using one of the extended
	/// mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnRButtonDblClk, Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_DRAWITEM handler</em>
	///
	/// Will be called if the control's parent window is asked to draw a tabstrip tab.
	/// We use this handler to raise the \c OwnerDrawTab event.
	///
	/// \sa get_OwnerDrawn, Raise_OwnerDrawTab,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms673316.aspx">WM_DRAWITEM</a>
	LRESULT OnReflectedDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NOTIFYFORMAT handler</em>
	///
	/// Will be called if the control asks its parent window which format (Unicode/ANSI) the \c WM_NOTIFY
	/// notifications should have.
	/// We use this handler for proper Unicode support.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms672615.aspx">WM_NOTIFYFORMAT</a>
	LRESULT OnReflectedNotifyFormat(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c DI_GETDRAGIMAGE handler</em>
	///
	/// Will be called during OLE drag'n'drop if the control is queried for a drag image.
	///
	/// \sa OLEDrag, CreateLegacyOLEDragImage, CreateVistaOLEDragImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	LRESULT OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c TCM_CREATEDRAGIMAGE handler</em>
	///
	/// Will be called if the control shall create a drag image.
	/// We use this handler to provide drag images.
	///
	/// \sa CreateLegacyDragImage, TCM_CREATEDRAGIMAGE
	LRESULT OnCreateDragImage(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_DELETEALLITEMS handler</em>
	///
	/// Will be called if all tabstrip tabs shall be removed from the control.
	/// We use this handler mainly to raise the \c RemovingTab and \c RemovedTab events.
	///
	/// \sa OnDeleteItem, OnInsertItem, Raise_FreeTabData, Raise_RemovingTab, Raise_RemovedTab,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650965.aspx">TCM_DELETEALLITEMS</a>
	LRESULT OnDeleteAllItems(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_DELETEITEM handler</em>
	///
	/// Will be called if a tabstrip tab shall be removed from the control.
	/// We use this handler mainly to raise the \c RemovingTab and \c RemovedTab events.
	///
	/// \sa OnDeleteAllItems, OnInsertItem, Raise_FreeTabData, Raise_RemovingTab, Raise_RemovedTab,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650968.aspx">TCM_DELETEITEM</a>
	LRESULT OnDeleteItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_DESELECTALL handler</em>
	///
	/// Will be called if all the control's tabs shall be deselected.
	/// We use this handler to raise the \c TabSelectionChanged event.
	///
	/// \sa Raise_TabSelectionChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650971.aspx">TCM_DESELECTALL</a>
	LRESULT OnDeselectAll(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_GETEXTENDEDSTYLE handler</em>
	///
	/// Will be called if the control's extended styles (\c TCS_EX_*) are requested.
	/// We use this handler to handle \c TCS_EX_DETECTDRAGDROP.
	///
	/// \sa OnSetExtendedStyle, TCS_EX_DETECTDRAGDROP,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650980.aspx">TCM_GETEXTENDEDSTYLE</a>
	LRESULT OnGetExtendedStyle(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_GETINSERTMARK handler</em>
	///
	/// Will be called if the position of the control's insertion mark shall be retrieved.
	/// We use this handler for insertion mark support.
	///
	/// \sa OnSetInsertMark, insertMark, GetInsertMarkPosition, OnGetInsertMarkColor, TCM_GETINSERTMARK
	LRESULT OnGetInsertMark(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_GETINSERTMARKCOLOR handler</em>
	///
	/// Will be called if the color of the control's insertion mark shall be retrieved.
	/// We use this handler for insertion mark support.
	///
	/// \sa OnSetInsertMarkColor, insertMark, OnGetInsertMark, TCM_GETINSERTMARKCOLOR
	LRESULT OnGetInsertMarkColor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_GETITEM handler</em>
	///
	/// Will be called if some or all of a tabstrip tab's attributes are requested from the control.
	/// We use this handler to hide the tab's unique ID from the caller and return the tab's
	/// associated data out of the \c tabParams list instead.
	///
	/// \sa tabParams, OnSetItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650986.aspx">TCM_GETITEM</a>
	LRESULT OnGetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_HIGHLIGHTITEM handler</em>
	///
	/// Will be called if a tabstrip tab's highlight state shall be changed.
	/// We use this handler to ensure that only one tab is highlighted.
	///
	/// \sa putref_DropHilitedTab,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651061.aspx">TCM_HIGHLIGHTITEM</a>
	LRESULT OnHighlightItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_INSERTITEM handler</em>
	///
	/// Will be called if a new tabstrip tab shall be inserted into the control.
	/// We use this handler mainly to raise the \c InsertingTab and \c InsertedTab events.
	///
	/// \sa OnDeleteItem, Raise_InsertingTab, Raise_InsertedTab,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651063.aspx">TCM_INSERTITEM</a>
	LRESULT OnInsertItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_SETCURFOCUS handler</em>
	///
	/// Will be called if the control's caret tab shall be changed.
	/// We use this handler to raise the \c CaretChanged event.
	///
	/// \sa OnFocusChangeNotification, Raise_CaretChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651065.aspx">TCM_SETCURFOCUS</a>
	LRESULT OnSetCurFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_SETCURSEL handler</em>
	///
	/// Will be called if the control's active tab shall be changed.
	/// We use this handler to raise the \c ActiveTabChanging and \c ActiveTabChanged events.
	///
	/// \sa OnSelChangingNotification, OnSelChangeNotification, Raise_ActiveTabChanging,
	///     Raise_ActiveTabChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651066.aspx">TCM_SETCURSEL</a>
	LRESULT OnSetCurSel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_SETEXTENDEDSTYLE handler</em>
	///
	/// Will be called if the control's extended styles (\c TCS_EX_*) shall be set.
	/// We use this handler to handle \c TCS_EX_DETECTDRAGDROP.
	///
	/// \sa OnGetExtendedStyle, TCS_EX_DETECTDRAGDROP,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651067.aspx">TCM_SETEXTENDEDSTYLE</a>
	LRESULT OnSetExtendedStyle(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_SETIMAGELIST handler</em>
	///
	/// Will be called if the control's imagelist shall be changed.
	/// We use this handler to synchronize our cached settings.
	///
	/// \sa cachedSettings, get_hImageList,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651068.aspx">TCM_SETIMAGELIST</a>
	LRESULT OnSetImageList(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_SETINSERTMARK handler</em>
	///
	/// Will be called if the control's insertion mark shall be repositioned.
	/// We use this handler for insertion mark support.
	///
	/// \sa OnGetInsertMark, insertMark, SetInsertMarkPosition, TCM_SETINSERTMARK
	LRESULT OnSetInsertMark(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_SETINSERTMARKCOLOR handler</em>
	///
	/// Will be called if the color of the control's insertion mark shall be changed.
	/// We use this handler for insertion mark support.
	///
	/// \sa OnGetInsertMarkColor, insertMark, OnSetInsertMark, TCM_SETINSERTMARKCOLOR
	LRESULT OnSetInsertMarkColor(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_SETITEM handler</em>
	///
	/// Will be called if the control is requested to update some or all of a tabstrip tab's attributes.
	/// We use this handler to avoid the tab's unique ID is overwritten and store the tab's
	/// associated data in the \c tabParams list instead.
	///
	/// \sa tabParams, OnGetItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651069.aspx">TCM_SETITEM</a>
	LRESULT OnSetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_SETITEMSIZE handler</em>
	///
	/// Will be called if the control's tab height and fixed tab width shall be changed.
	/// We use this handler to synchronize the \c FixedTabWidth and \c TabHeight properties.
	///
	/// \sa OnSetMinTabWidth, get_FixedTabWidth, get_TabHeight,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651071.aspx">TCM_SETITEMSIZE</a>
	LRESULT OnSetItemSize(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_SETMINTABWIDTH handler</em>
	///
	/// Will be called if the control's minimum tab width shall be changed.
	/// We use this handler to synchronize the \c MinTabWidth property.
	///
	/// \sa OnSetItemSize, get_MinTabWidth,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651072.aspx">TCM_SETMINTABWIDTH</a>
	LRESULT OnSetMinTabWidth(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCM_SETPADDING handler</em>
	///
	/// Will be called if the amount of space (padding) around each tab's icon and label in pixels shall
	/// be changed.
	/// We use this handler to synchronize the \c HorizontalPadding and \c VerticalPadding properties.
	///
	/// \sa get_HorizontalPadding, get_VerticalPadding,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651073.aspx">TCM_SETPADDING</a>
	LRESULT OnSetPadding(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c TTN_GETDISPINFOA handler</em>
	///
	/// Will be called if the tooltip control requests the text to display from the control.
	/// We use this handler for the \c ShowToolTips property.
	///
	/// \sa OnToolTipGetDispInfoNotificationW, get_ShowToolTips,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650474.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOW handler</em>
	///
	/// Will be called if the tooltip control requests the text to display from the control.
	/// We use this handler for the \c ShowToolTips property.
	///
	/// \sa OnToolTipGetDispInfoNotificationA, get_ShowToolTips,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650474.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c NM_CLICK handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user clicked into the control's
	/// client area using the left mouse button.
	/// We use this handler to raise the \c Click event.
	///
	/// \sa OnLButtonDown, OnRClickNotification, Raise_Click,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650944.aspx">NM_CLICK (tab)</a>
	LRESULT OnClickNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c NM_RCLICK handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user clicked into the control's
	/// client area using the right mouse button.
	/// We use this handler to raise the \c RClick event.
	///
	/// \sa OnRButtonDown, OnClickNotification, Raise_RClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650946.aspx">NM_RCLICK (tab)</a>
	LRESULT OnRClickNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCN_BEGINDRAG handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user wants to drag a tabstrip
	/// tab using the left mouse button.
	/// We use this handler to raise the \c TabBeginDrag event.
	///
	/// \sa OnBeginRDragNotification, Raise_TabBeginDrag, TCN_BEGINDRAG
	LRESULT OnBeginDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCN_BEGINRDRAG handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user wants to drag a tabstrip
	/// tab using the right mouse button.
	/// We use this handler to raise the \c TabBeginRDrag event.
	///
	/// \sa OnBeginDragNotification, Raise_TabBeginRDrag, TCN_BEGINRDRAG
	LRESULT OnBeginRDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCN_FOCUSCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the keyboard focus has changed to
	/// another tab.
	/// We use this handler to raise the \c CaretChanged event.
	///
	/// \sa OnSelChangeNotification, Raise_CaretChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650950.aspx">TCN_FOCUSCHANGE</a>
	LRESULT OnFocusChangeNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCN_GETOBJECT handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user is dragging data over a tab.
	/// We use this handler to make \c TCS_EX_REGISTERDROP work together with our implementation of
	/// \c IDropTargt.
	///
	/// \sa put_RegisterForOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650954.aspx">TCN_GETOBJECT</a>
	LRESULT OnGetObjectNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCN_SELCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that another tab has become active.
	/// We use this handler to raise the \c ActiveTabChanged event.
	///
	/// \sa OnSelChangingNotification, OnFocusChangeNotification, Raise_ActiveTabChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650958.aspx">TCN_SELCHANGE</a>
	LRESULT OnSelChangeNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c TCN_SELCHANGING handler</em>
	///
	/// Will be called if the control's parent window is notified, that another tab is about to become
	/// active.
	/// We use this handler to raise the \c ActiveTabChanging event.
	///
	/// \sa OnSelChangeNotification, OnFocusChangeNotification, Raise_ActiveTabChanging,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650960.aspx">TCN_SELCHANGING</a>
	LRESULT OnSelChangingNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c AbortedDrag event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_AbortedDrag, TabStripCtlLibU::_ITabStripEvents::AbortedDrag,
	///       Raise_Drop, EndDrag
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_AbortedDrag, TabStripCtlLibA::_ITabStripEvents::AbortedDrag,
	///       Raise_Drop, EndDrag
	/// \endif
	inline HRESULT Raise_AbortedDrag(void);
	/// \brief <em>Raises the \c ActiveTabChanged event</em>
	///
	/// \param[in] previousActiveTab The previous active tab.
	/// \param[in] newActiveTab The new active tab.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_ActiveTabChanged,
	///       TabStripCtlLibU::_ITabStripEvents::ActiveTabChanged, Raise_ActiveTabChanging,
	///       TabStripTab::get_Active, get_ActiveTab, Raise_CaretChanged
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_ActiveTabChanged,
	///       TabStripCtlLibA::_ITabStripEvents::ActiveTabChanged, Raise_ActiveTabChanging,
	///       TabStripTab::get_Active, get_ActiveTab, Raise_CaretChanged
	/// \endif
	inline HRESULT Raise_ActiveTabChanged(int previousActiveTab, int newActiveTab);
	/// \brief <em>Raises the \c ActiveTabChanging event</em>
	///
	/// \param[in] previousActiveTab The previous active tab.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort the active tab change; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_ActiveTabChanging,
	///       TabStripCtlLibU::_ITabStripEvents::ActiveTabChanging, Raise_ActiveTabChanged,
	///       TabStripTab::get_Active, get_ActiveTab, Raise_CaretChanged
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_ActiveTabChanging,
	///       TabStripCtlLibA::_ITabStripEvents::ActiveTabChanging, Raise_ActiveTabChanged,
	///       TabStripTab::get_Active, get_ActiveTab, Raise_CaretChanged
	/// \endif
	inline HRESULT Raise_ActiveTabChanging(int previousActiveTab, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c CaretChanged event</em>
	///
	/// \param[in] previousCaretTab The previous caret tab.
	/// \param[in] newCaretTab The new caret tab.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_CaretChanged, TabStripCtlLibU::_ITabStripEvents::CaretChanged,
	///       TabStripTab::get_Caret, get_CaretTab, Raise_ActiveTabChanged
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_CaretChanged, TabStripCtlLibA::_ITabStripEvents::CaretChanged,
	///       TabStripTab::get_Caret, get_CaretTab, Raise_ActiveTabChanged
	/// \endif
	inline HRESULT Raise_CaretChanged(int previousCaretTab, int newCaretTab);
	/// \brief <em>Raises the \c Click event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_Click, TabStripCtlLibU::_ITabStripEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_Click, TabStripCtlLibA::_ITabStripEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c ContextMenu event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_ContextMenu, TabStripCtlLibU::_ITabStripEvents::ContextMenu,
	///       Raise_RClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_ContextMenu, TabStripCtlLibA::_ITabStripEvents::ContextMenu,
	///       Raise_RClick
	/// \endif
	inline HRESULT Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_DblClick, TabStripCtlLibU::_ITabStripEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_DblClick, TabStripCtlLibA::_ITabStripEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_DestroyedControlWindow,
	///       TabStripCtlLibU::_ITabStripEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_DestroyedControlWindow,
	///       TabStripCtlLibA::_ITabStripEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c DragMouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_DragMouseMove, TabStripCtlLibU::_ITabStripEvents::DragMouseMove,
	///       Raise_MouseMove, Raise_OLEDragMouseMove, BeginDrag
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_DragMouseMove, TabStripCtlLibA::_ITabStripEvents::DragMouseMove,
	///       Raise_MouseMove, Raise_OLEDragMouseMove, BeginDrag
	/// \endif
	inline HRESULT Raise_DragMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c Drop event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_Drop, TabStripCtlLibU::_ITabStripEvents::Drop, Raise_AbortedDrag,
	///       EndDrag
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_Drop, TabStripCtlLibA::_ITabStripEvents::Drop, Raise_AbortedDrag,
	///       EndDrag
	/// \endif
	inline HRESULT Raise_Drop(void);
	/// \brief <em>Raises the \c FreeTabData event</em>
	///
	/// \param[in] pTSTab The tab whose associated data shall be freed. If \c NULL, all tabs' associated data
	///            shall be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_FreeTabData, TabStripCtlLibU::_ITabStripEvents::FreeTabData,
	///       Raise_RemovingTab, Raise_RemovedTab, TabStripTab::put_TabData
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_FreeTabData, TabStripCtlLibA::_ITabStripEvents::FreeTabData,
	///       Raise_RemovingTab, Raise_RemovedTab, TabStripTab::put_TabData
	/// \endif
	inline HRESULT Raise_FreeTabData(ITabStripTab* pTSTab);
	/// \brief <em>Raises the \c InsertedTab event</em>
	///
	/// \param[in] pTSTab The inserted tab.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_InsertedTab, TabStripCtlLibU::_ITabStripEvents::InsertedTab,
	///       Raise_InsertingTab, Raise_RemovedTab
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_InsertedTab, TabStripCtlLibA::_ITabStripEvents::InsertedTab,
	///       Raise_InsertingTab, Raise_RemovedTab
	/// \endif
	inline HRESULT Raise_InsertedTab(ITabStripTab* pTSTab);
	/// \brief <em>Raises the \c InsertingTab event</em>
	///
	/// \param[in] pTSTab The tab being inserted.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort insertion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_InsertingTab, TabStripCtlLibU::_ITabStripEvents::InsertingTab,
	///       Raise_InsertedTab, Raise_RemovingTab
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_InsertingTab, TabStripCtlLibA::_ITabStripEvents::InsertingTab,
	///       Raise_InsertedTab, Raise_RemovingTab
	/// \endif
	inline HRESULT Raise_InsertingTab(IVirtualTabStripTab* pTSTab, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_KeyDown, TabStripCtlLibU::_ITabStripEvents::KeyDown, Raise_KeyUp,
	///       Raise_KeyPress
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_KeyDown, TabStripCtlLibA::_ITabStripEvents::KeyDown, Raise_KeyUp,
	///       Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyDown(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_KeyPress, TabStripCtlLibU::_ITabStripEvents::KeyPress,
	///       Raise_KeyDown, Raise_KeyUp
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_KeyPress, TabStripCtlLibA::_ITabStripEvents::KeyPress,
	///       Raise_KeyDown, Raise_KeyUp
	/// \endif
	inline HRESULT Raise_KeyPress(SHORT* pKeyAscii);
	/// \brief <em>Raises the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_KeyUp, TabStripCtlLibU::_ITabStripEvents::KeyUp, Raise_KeyDown,
	///       Raise_KeyPress
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_KeyUp, TabStripCtlLibA::_ITabStripEvents::KeyUp, Raise_KeyDown,
	///       Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c MClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_MClick, TabStripCtlLibU::_ITabStripEvents::MClick,
	///       Raise_MDblClick, Raise_Click, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_MClick, TabStripCtlLibA::_ITabStripEvents::MClick,
	///       Raise_MDblClick, Raise_Click, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_MDblClick, TabStripCtlLibU::_ITabStripEvents::MDblClick,
	///       Raise_MClick, Raise_DblClick, Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_MDblClick, TabStripCtlLibA::_ITabStripEvents::MDblClick,
	///       Raise_MClick, Raise_DblClick, Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_MouseDown, TabStripCtlLibU::_ITabStripEvents::MouseDown,
	///       Raise_MouseUp, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_MouseDown, TabStripCtlLibA::_ITabStripEvents::MouseDown,
	///       Raise_MouseUp, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseEnter event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_MouseEnter, TabStripCtlLibU::_ITabStripEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_TabMouseEnter, Raise_MouseHover, Raise_MouseMove
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_MouseEnter, TabStripCtlLibA::_ITabStripEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_TabMouseEnter, Raise_MouseHover, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseHover event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_MouseHover, TabStripCtlLibU::_ITabStripEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_MouseHover, TabStripCtlLibA::_ITabStripEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	/// \endif
	inline HRESULT Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseLeave event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_MouseLeave, TabStripCtlLibU::_ITabStripEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_TabMouseLeave, Raise_MouseHover, Raise_MouseMove
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_MouseLeave, TabStripCtlLibA::_ITabStripEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_TabMouseLeave, Raise_MouseHover, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_MouseMove, TabStripCtlLibU::_ITabStripEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_MouseMove, TabStripCtlLibA::_ITabStripEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp
	/// \endif
	inline HRESULT Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseUp event</em>
	///
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_MouseUp, TabStripCtlLibU::_ITabStripEvents::MouseUp,
	///       Raise_MouseDown, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_MouseUp, TabStripCtlLibA::_ITabStripEvents::MouseUp,
	///       Raise_MouseDown, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLECompleteDrag,
	///       TabStripCtlLibU::_ITabStripEvents::OLECompleteDrag, Raise_OLEStartDrag,
	///       TabStripCtlLibU::IOLEDataObject::GetData, OLEDrag
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLECompleteDrag,
	///       TabStripCtlLibA::_ITabStripEvents::OLECompleteDrag, Raise_OLEStartDrag,
	///       TabStripCtlLibA::IOLEDataObject::GetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect);
	/// \brief <em>Raises the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client finally executed.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragDrop, TabStripCtlLibU::_ITabStripEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragDrop, TabStripCtlLibA::_ITabStripEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data. If \c NULL, the cached data object is used. We use this when
	///            we call this method from other places than \c DragEnter.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragEnter, TabStripCtlLibU::_ITabStripEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragEnter, TabStripCtlLibA::_ITabStripEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The potential drop target window's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragEnterPotentialTarget,
	///       TabStripCtlLibU::_ITabStripEvents::OLEDragEnterPotentialTarget,
	///       Raise_OLEDragLeavePotentialTarget, OLEDrag
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragEnterPotentialTarget,
	///       TabStripCtlLibA::_ITabStripEvents::OLEDragEnterPotentialTarget,
	///       Raise_OLEDragLeavePotentialTarget, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragLeave, TabStripCtlLibU::_ITabStripEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragLeave, TabStripCtlLibA::_ITabStripEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(void);
	/// \brief <em>Raises the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragLeavePotentialTarget,
	///       TabStripCtlLibU::_ITabStripEvents::OLEDragLeavePotentialTarget,
	///       Raise_OLEDragEnterPotentialTarget, OLEDrag
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragLeavePotentialTarget,
	///       TabStripCtlLibA::_ITabStripEvents::OLEDragLeavePotentialTarget,
	///       Raise_OLEDragEnterPotentialTarget, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragLeavePotentialTarget(void);
	/// \brief <em>Raises the \c OLEDragMouseMove event</em>
	///
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragMouseMove,
	///       TabStripCtlLibU::_ITabStripEvents::OLEDragMouseMove, Raise_OLEDragEnter,
	///       Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLEDragMouseMove,
	///       TabStripCtlLibA::_ITabStripEvents::OLEDragMouseMove, Raise_OLEDragEnter,
	///       Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c DROPEFFECT enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLEGiveFeedback,
	///       TabStripCtlLibU::_ITabStripEvents::OLEGiveFeedback, Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLEGiveFeedback,
	///       TabStripCtlLibA::_ITabStripEvents::OLEGiveFeedback, Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors);
	/// \brief <em>Raises the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c TRUE, the user has pressed the \c ESC key since the last time
	///            this event was raised; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue (\c S_OK), cancel
	///                (\c DRAGDROP_S_CANCEL) or complete (\c DRAGDROP_S_DROP) the drag'n'drop operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLEQueryContinueDrag,
	///       TabStripCtlLibU::_ITabStripEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLEQueryContinueDrag,
	///       TabStripCtlLibA::_ITabStripEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \endif
	inline HRESULT Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith);
	/// \brief <em>Raises the \c OLEReceivedNewData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the data object has received data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLEReceivedNewData,
	///       TabStripCtlLibU::_ITabStripEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLEReceivedNewData,
	///       TabStripCtlLibA::_ITabStripEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLESetData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the drop target is requesting data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLESetData, TabStripCtlLibU::_ITabStripEvents::OLESetData,
	///       Raise_OLEStartDrag, SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLESetData, TabStripCtlLibA::_ITabStripEvents::OLESetData,
	///       Raise_OLEStartDrag, SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OLEStartDrag, TabStripCtlLibU::_ITabStripEvents::OLEStartDrag,
	///       Raise_OLESetData, Raise_OLECompleteDrag, SourceOLEDataObject::SetData, OLEDrag
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OLEStartDrag, TabStripCtlLibA::_ITabStripEvents::OLEStartDrag,
	///       Raise_OLESetData, Raise_OLECompleteDrag, SourceOLEDataObject::SetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEStartDrag(IOLEDataObject* pData);
	/// \brief <em>Raises the \c OwnerDrawTab event</em>
	///
	/// \param[in] pTSTab The tab to draw.
	/// \param[in] tabState The tab's current state (e. g. selected). Any of the values defined by the
	///            \c OwnerDrawTabStateConstants enumeration is valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_OwnerDrawTab, TabStripCtlLibU::_ITabStripEvents::OwnerDrawTab,
	///       get_OwnerDrawn, TabStripCtlLibU::RECTANGLE, TabStripCtlLibU::OwnerDrawTabStateConstants
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_OwnerDrawTab, TabStripCtlLibA::_ITabStripEvents::OwnerDrawTab,
	///       get_OwnerDrawn, TabStripCtlLibA::RECTANGLE, TabStripCtlLibA::OwnerDrawTabStateConstants
	/// \endif
	inline HRESULT Raise_OwnerDrawTab(ITabStripTab* pTSTab, OwnerDrawTabStateConstants tabState, LONG hDC, RECTANGLE* pDrawingRectangle);
	/// \brief <em>Raises the \c RClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_RClick, TabStripCtlLibU::_ITabStripEvents::RClick,
	///       Raise_ContextMenu, Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_RClick, TabStripCtlLibA::_ITabStripEvents::RClick,
	///       Raise_ContextMenu, Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_RDblClick, TabStripCtlLibU::_ITabStripEvents::RDblClick,
	///       Raise_RClick, Raise_DblClick, Raise_MDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_RDblClick, TabStripCtlLibA::_ITabStripEvents::RDblClick,
	///       Raise_RClick, Raise_DblClick, Raise_MDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_RecreatedControlWindow,
	///       TabStripCtlLibU::_ITabStripEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_RecreatedControlWindow,
	///       TabStripCtlLibA::_ITabStripEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c RemovedTab event</em>
	///
	/// \param[in] pTSTab The removed tab. If \c NULL, all tabs were removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_RemovedTab,
	///       TabStripCtlLibU::_ITabStripEvents::RemovedTab, Raise_RemovingTab, Raise_InsertedTab
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_RemovedTab,
	///       TabStripCtlLibA::_ITabStripEvents::RemovedTab, Raise_RemovingTab, Raise_InsertedTab
	/// \endif
	inline HRESULT Raise_RemovedTab(IVirtualTabStripTab* pTSTab);
	/// \brief <em>Raises the \c RemovingTab event</em>
	///
	/// \param[in] pTSTab The tab being removed. If \c NULL, all tabs are removed.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort deletion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_RemovingTab,
	///       TabStripCtlLibU::_ITabStripEvents::RemovingTab, Raise_RemovedTab, Raise_InsertingTab
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_RemovingTab,
	///       TabStripCtlLibA::_ITabStripEvents::RemovingTab, Raise_RemovedTab, Raise_InsertingTab
	/// \endif
	inline HRESULT Raise_RemovingTab(ITabStripTab* pTSTab, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_ResizedControlWindow,
	///       TabStripCtlLibU::_ITabStripEvents::ResizedControlWindow
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_ResizedControlWindow,
	///       TabStripCtlLibA::_ITabStripEvents::ResizedControlWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c TabBeginDrag event</em>
	///
	/// \param[in] pTSTab The tab that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbLeftButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_TabBeginDrag, TabStripCtlLibU::_ITabStripEvents::TabBeginDrag,
	///       BeginDrag, OLEDrag, get_AllowDragDrop, Raise_TabBeginRDrag, TabStripCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_TabBeginDrag, TabStripCtlLibA::_ITabStripEvents::TabBeginDrag,
	///       BeginDrag, OLEDrag, get_AllowDragDrop, Raise_TabBeginRDrag, TabStripCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_TabBeginDrag(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c TabBeginRDrag event</em>
	///
	/// \param[in] pTSTab The tab that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbRightButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_TabBeginRDrag, TabStripCtlLibU::_ITabStripEvents::TabBeginRDrag,
	///       BeginDrag, OLEDrag, get_AllowDragDrop, Raise_TabBeginDrag, TabStripCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_TabBeginRDrag, TabStripCtlLibA::_ITabStripEvents::TabBeginRDrag,
	///       BeginDrag, OLEDrag, get_AllowDragDrop, Raise_TabBeginDrag, TabStripCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_TabBeginRDrag(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c TabGetInfoTipText event</em>
	///
	/// \param[in] pTSTab The tab whose info tip text is requested.
	/// \param[in] maxInfoTipLength The maximum number of characters the info tip text may consist of.
	/// \param[out] pInfoTipText The tab's info tip text.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_TabGetInfoTipText,
	///       TabStripCtlLibU::_ITabStripEvents::TabGetInfoTipText, put_ShowToolTips
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_TabGetInfoTipText,
	///       TabStripCtlLibA::_ITabStripEvents::TabGetInfoTipText, put_ShowToolTips
	/// \endif
	inline HRESULT Raise_TabGetInfoTipText(ITabStripTab* pTSTab, LONG maxInfoTipLength, BSTR* pInfoTipText);
	/// \brief <em>Raises the \c TabMouseEnter event</em>
	///
	/// \param[in] pTSTab The tab that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_TabMouseEnter, TabStripCtlLibU::_ITabStripEvents::TabMouseEnter,
	///       Raise_TabMouseLeave, Raise_MouseMove, TabStripCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_TabMouseEnter, TabStripCtlLibA::_ITabStripEvents::TabMouseEnter,
	///       Raise_TabMouseLeave, Raise_MouseMove, TabStripCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_TabMouseEnter(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c TabMouseLeave event</em>
	///
	/// \param[in] pTSTab The tab that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_TabMouseLeave, TabStripCtlLibU::_ITabStripEvents::TabMouseLeave,
	///       Raise_TabMouseEnter, Raise_MouseMove, TabStripCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_TabMouseLeave, TabStripCtlLibA::_ITabStripEvents::TabMouseLeave,
	///       Raise_TabMouseEnter, Raise_MouseMove, TabStripCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_TabMouseLeave(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c TabSelectionChanged event</em>
	///
	/// \param[in] tabIndex The tab that was selected/deselected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_TabSelectionChanged,
	///       TabStripCtlLibU::_ITabStripEvents::TabSelectionChanged, Raise_ActiveTabChanged,
	///       Raise_CaretChanged, TabStripTab::get_Selected, get_MultiSelect,
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_TabSelectionChanged,
	///       TabStripCtlLibA::_ITabStripEvents::TabSelectionChanged, Raise_ActiveTabChanged,
	///       Raise_CaretChanged, TabStripTab::get_Selected, get_MultiSelect,
	/// \endif
	inline HRESULT Raise_TabSelectionChanged(int tabIndex);
	/// \brief <em>Raises the \c XClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_XClick, TabStripCtlLibU::_ITabStripEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_XClick, TabStripCtlLibA::_ITabStripEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick
	/// \endif
	inline HRESULT Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c XDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITabStripEvents::Fire_XDblClick, TabStripCtlLibU::_ITabStripEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick
	/// \else
	///   \sa Proxy_ITabStripEvents::Fire_XDblClick, TabStripCtlLibA::_ITabStripEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick
	/// \endif
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies the font to ourselves</em>
	///
	/// This method sets our font to the font specified by the \c Font property.
	///
	/// \sa get_Font
	void ApplyFont(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleObject
	///
	//@{
	/// \brief <em>Implementation of \c IOleObject::DoVerb</em>
	///
	/// Will be called if one of the control's registered actions (verbs) shall be executed; e. g.
	/// executing the "About" verb will display the About dialog.
	///
	/// \param[in] verbID The requested verb's ID.
	/// \param[in] pMessage A \c MSG structure describing the event (such as a double-click) that
	///            invoked the verb.
	/// \param[in] pActiveSite The \c IOleClientSite implementation of the control's active client site
	///            where the event occurred that invoked the verb.
	/// \param[in] reserved Reserved; must be zero.
	/// \param[in] hWndParent The handle of the document window containing the control.
	/// \param[in] pBoundingRectangle A \c RECT structure containing the coordinates and size in pixels
	///            of the control within the window specified by \c hWndParent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnumVerbs,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694508.aspx">IOleObject::DoVerb</a>
	virtual HRESULT STDMETHODCALLTYPE DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle);
	/// \brief <em>Implementation of \c IOleObject::EnumVerbs</em>
	///
	/// Will be called if the control's container requests the control's registered actions (verbs); e. g.
	/// we provide a verb "About" that will display the About dialog on execution.
	///
	/// \param[out] ppEnumerator Receives the \c IEnumOLEVERB implementation of the verbs' enumerator.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692781.aspx">IOleObject::EnumVerbs</a>
	virtual HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** ppEnumerator);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleControl
	///
	//@{
	/// \brief <em>Implementation of \c IOleControl::GetControlInfo</em>
	///
	/// Will be called if the container requests details about the control's keyboard mnemonics and keyboard
	/// behavior.
	///
	/// \param[in, out] pControlInfo The requested details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnMnemonic, Properties::hAcceleratorTable,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693730.aspx">IOleControl::GetControlInfo</a>
	virtual HRESULT STDMETHODCALLTYPE GetControlInfo(LPCONTROLINFO pControlInfo);
	/// \brief <em>Implementation of \c IOleControl::OnMnemonic</em>
	///
	/// Will be called if the user has pressed a keystroke that matches one of the entries in the control's
	/// accelerator table.
	///
	/// \param[in] pMessage The \c MSG structure describing the keystroke.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetControlInfo, Properties::hAcceleratorTable,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680699.aspx">IOleControl::OnMnemonic</a>
	virtual HRESULT STDMETHODCALLTYPE OnMnemonic(LPMSG pMessage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Executes the About verb</em>
	///
	/// Will be called if the control's registered actions (verbs) "About" shall be executed. We'll
	/// display the About dialog.
	///
	/// \param[in] hWndParent The window to use as parent for any user interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb, About
	HRESULT DoVerbAbout(HWND hWndParent);

	//////////////////////////////////////////////////////////////////////
	/// \name Tab insertion and deletion
	///
	//@{
	/// \brief <em>Increments the \c silentTabInsertions flag</em>
	///
	/// \sa LeaveSilentTabInsertionSection, Flags::silentTabInsertions, EnterSilentTabDeletionSection,
	///     TabStripTab::put_Index, Raise_InsertingTab, Raise_InsertedTab
	void EnterSilentTabInsertionSection(void);
	/// \brief <em>Decrements the \c silentTabInsertions flag</em>
	///
	/// \sa EnterSilentTabInsertionSection, Flags::silentTabInsertions, LeaveSilentTabDeletionSection,
	///     TabStripTab::put_Index, Raise_InsertingTab, Raise_InsertedTab
	void LeaveSilentTabInsertionSection(void);
	/// \brief <em>Increments the \c silentTabDeletions flag</em>
	///
	/// \sa LeaveSilentTabDeletionSection, Flags::silentTabDeletions, EnterSilentTabInsertionSection,
	///     TabStripTab::put_Index, Raise_RemovingTab, Raise_RemovedTab
	void EnterSilentTabDeletionSection(void);
	/// \brief <em>Decrements the \c silentTabDeletions flag</em>
	///
	/// \sa EnterSilentTabDeletionSection, Flags::silentTabDeletions, LeaveSilentTabInsertionSection,
	///     TabStripTab::put_Index, Raise_RemovingTab, Raise_RemovedTab
	void LeaveSilentTabDeletionSection(void);
	/// \brief <em>Increments the \c silentActiveTabChanges flag</em>
	///
	/// \sa LeaveSilentActiveTabChangeSection, Flags::silentActiveTabChanges, EnterSilentCaretChangeSection,
	///     TabStripTab::put_Index, Raise_ActiveTabChanging, Raise_ActiveTabChanged
	void EnterSilentActiveTabChangeSection(void);
	/// \brief <em>Decrements the \c silentActiveTabChanges flag</em>
	///
	/// \sa EnterSilentActiveTabChangeSection, Flags::silentActiveTabChanges, LeaveSilentCaretChangeSection,
	///     TabStripTab::put_Index, Raise_ActiveTabChanging, Raise_ActiveTabChanged
	void LeaveSilentActiveTabChangeSection(void);
	/// \brief <em>Increments the \c silentCaretChanges flag</em>
	///
	/// \sa LeaveSilentCaretChangeSection, Flags::silentCaretChanges, EnterSilentActiveTabChangeSection,
	///     TabStripTab::put_Index, Raise_CaretChanged
	void EnterSilentCaretChangeSection(void);
	/// \brief <em>Decrements the \c silentCaretChanges flag</em>
	///
	/// \sa EnterSilentCaretChangeSection, Flags::silentCaretChanges, LeaveSilentActiveTabChangeSection,
	///     TabStripTab::put_Index, Raise_CaretChanged
	void LeaveSilentCaretChangeSection(void);
	/// \brief <em>Increments the \c silentSelectionChanges flag</em>
	///
	/// \sa LeaveSilentSelectionChangesSection, Flags::silentSelectionChanges, EnterSilentCaretChangeSection,
	///     TabStripTab::put_Index, Raise_TabSelectionChanged
	void EnterSilentSelectionChangesSection(void);
	/// \brief <em>Decrements the \c silentSelectionChanges flag</em>
	///
	/// \sa EnterSilentSelectionChangesSection, Flags::silentSelectionChanges, LeaveSilentCaretChangeSection,
	///     TabStripTab::put_Index, Raise_TabSelectionChanged
	void LeaveSilentSelectionChangesSection(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name MFC clones
	///
	//@{
	/// \brief <em>A rewrite of MFC's \c COleControl::RecreateControlWindow</em>
	///
	/// Destroys and re-creates the control window.
	///
	/// \remarks This rewrite probably isn't complete.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/5tx8h2fd.aspx">COleControl::RecreateControlWindow</a>
	void RecreateControlWindow(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Control window configuration
	///
	//@{
	/// \brief <em>Calculates the extended window style bits</em>
	///
	/// Calculates the extended window style bits that the control must have set to match the current
	/// property settings.
	///
	/// \return A bit field of \c WS_EX_* constants specifying the required extended window styles.
	///
	/// \sa GetStyleBits, SendConfigurationMessages, Create
	DWORD GetExStyleBits(void);
	/// \brief <em>Calculates the window style bits</em>
	///
	/// Calculates the window style bits that the control must have set to match the current property
	/// settings.
	///
	/// \return A bit field of \c WS_* and \c TCS_* constants specifying the required window styles.
	///
	/// \sa GetExStyleBits, SendConfigurationMessages, Create
	DWORD GetStyleBits(void);
	/// \brief <em>Configures the control by sending messages</em>
	///
	/// Sends \c WM_* and \c TCM_* messages to the control window to make it match the current property
	/// settings. Will be called out of \c Raise_RecreatedControlWindow.
	///
	/// \sa GetExStyleBits, GetStyleBits, Raise_RecreatedControlWindow
	void SendConfigurationMessages(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Value translation
	///
	//@{
	/// \brief <em>Translates a \c MousePointerConstants type mouse cursor into a \c HCURSOR type mouse cursor</em>
	///
	/// \param[in] mousePointer The \c MousePointerConstants type mouse cursor to translate.
	///
	/// \return The translated \c HCURSOR type mouse cursor.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \else
	///   \sa TabStripCtlLibA::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \endif
	HCURSOR MousePointerConst2hCursor(MousePointerConstants mousePointer);
	/// \brief <em>Translates a tab's unique ID to its index</em>
	///
	/// \param[in] ID The unique ID of the tab whose index is requested.
	///
	/// \return The requested tab's index if successful; -1 otherwise.
	///
	/// \sa tabIDs, TabIndexToID, TabStripTabs::get_Item, TabStripTabs::Remove
	int IDToTabIndex(LONG ID);
	/// \brief <em>Translates a tab's index to its unique ID</em>
	///
	/// \param[in] tabIndex The index of the tab whose unique ID is requested.
	///
	/// \return The requested tab's unique ID if successful; -1 otherwise.
	///
	/// \sa tabIDs, IDToTabIndex, TabStripTab::get_ID
	LONG TabIndexToID(int tabIndex);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the tab that lies below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pFlags A bit field of TCHT_* flags, that holds further details about the control's
	///             part below the specified point.
	/// \param[in] ignoreBoundingBoxDefinition If \c TRUE the setting of the \c TabBoundingBoxDefinition
	///            property is ignored; otherwise the returned tab index is set to -1, if the
	///            \c TabBoundingBoxDefinition property's setting and the \c pFlags parameter don't match.
	///
	/// \return The "hit" tab's index.
	///
	/// \sa HitTest, get_TabBoundingBoxDefinition
	int HitTest(LONG x, LONG y, UINT* pFlags, BOOL ignoreBoundingBoxDefinition = FALSE);
	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);
	/// \brief <em>Moves a tab within the \c tabIDs vector and \c tabParams hash map to a new position</em>
	///
	/// \param[in] oldTabIndex The tab's current zero-based index.
	/// \param[in] newTabIndex The tab's new zero-based index.
	///
	/// \sa tabIDs, tabParams, TabStripTab::put_Index, IDToTabIndex, TabIndexToID
	void MoveTabIndex(int oldTabIndex, int newTabIndex);
	/// \brief <em>Retrieves the handle of the window that is associated with the specified tab</em>
	///
	/// The client application can associate a window with a tab. If the tab is selected, the window is
	/// made visible. If the tab is deselected, the window is hidden.
	///
	/// \param[in] tabIndex The zero-based index of the tab for which to retrieve the associated window.
	///
	/// \return The handle of the window that is associated with the specified tab.
	///
	/// \sa SetAssociatedWindow, TabStripTab::get_hAssociatedWindow
	HWND GetAssociatedWindow(int tabIndex);
	/// \brief <em>Sets the handle of the window that is associated with the specified tab</em>
	///
	/// The client application can associate a window with a tab. If the tab is selected, the window is
	/// made visible. If the tab is deselected, the window is hidden.
	///
	/// \param[in] tabIndex The zero-based index of the tab for which to set the associated window.
	/// \param[in] hAssociatedWindow The handle of the window to associate with the specified tab.
	///
	/// \sa GetAssociatedWindow, TabStripTab::put_hAssociatedWindow
	void SetAssociatedWindow(int tabIndex, __in_opt HWND hAssociatedWindow);
	/// \brief <em>Retrieves whether the specified tab has a close button</em>
	///
	/// \param[in] tabIndex The zero-based index of the tab for which to check the closeable state.
	///
	/// \return \c TRUE, if the tab is closeable; otherwise \c FALSE.
	///
	/// \sa SetCloseableTab, TabStripTab::get_HasCloseButton
	BOOL IsCloseableTab(int tabIndex);
	/// \brief <em>Sets whether the specified tab has a close button</em>
	///
	/// \param[in] tabIndex The zero-based index of the tab for which to set the closeable state.
	/// \param[in] isCloseable \c TRUE to make the specified tab closeable; \c FALSE to make it not
	///            closeable.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsCloseableTab, TabStripTab::put_HasCloseButton
	HRESULT SetCloseableTab(int tabIndex, BOOL isCloseable);
	/// \brief <em>Auto-scrolls the control</em>
	///
	/// \sa OnTimer, Raise_OLEDragMouseMove, DragDropStatus::AutoScrolling
	void AutoScroll(void);
	/// \brief <em>Recreates the control's accelerator table</em>
	///
	/// \sa Properties::hAcceleratorTable, OnInsertItem, OnDeleteItem, OnDeleteAllItems, OnSetItem,
	///     GetControlInfo, OnMnemonic
	void RebuildAcceleratorTable(void);

	/// \brief <em>Calculates the bounding rectangle of a tab's close button</em>
	///
	/// \param[in] tabIndex The tab whose close button's bounding rectangle is calculated.
	/// \param[in] tabIsActive If \c TRUE, the tab is the control's active tab; otherwise not.
	/// \param[out] pButtonRectangle Receives the bounding rectangle in pixels.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa DrawCloseButtonsAndInsertionMarks, get_ActiveTab
	BOOL CalculateCloseButtonRectangle(int tabIndex, BOOL tabIsActive, LPRECT pButtonRectangle);
	/// \brief <em>Draws any close buttons and insertion marks</em>
	///
	/// This method is called by \c OnPaint to draw the control's close buttons and insertion mark.
	///
	/// \param[in] hTargetDC The device context to be used for drawing.
	///
	/// \sa OnPaint
	void DrawCloseButtonsAndInsertionMarks(HDC hTargetDC);
	/// \brief <em>Retrieves whether the logical left mouse button is held down</em>
	///
	/// \return \c TRUE if the logical left mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsRightMouseButtonDown
	BOOL IsLeftMouseButtonDown(void);
	/// \brief <em>Retrieves whether the logical right mouse button is held down</em>
	///
	/// \return \c TRUE if the logical right mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsLeftMouseButtonDown
	BOOL IsRightMouseButtonDown(void);


	/// \brief <em>Holds a control instance's properties' settings</em>
	typedef struct Properties
	{
		/// \brief <em>Holds a font property's settings</em>
		typedef struct FontProperty
		{
		protected:
			/// \brief <em>Holds the control's default font</em>
			///
			/// \sa GetDefaultFont
			static FONTDESC defaultFont;

		public:
			/// \brief <em>Holds whether we're listening for events fired by the font object</em>
			///
			/// If greater than 0, we're advised to the \c IFontDisp object identified by \c pFontDisp. I. e.
			/// we will be notified if a property of the font object changes. If 0, we won't receive any events
			/// fired by the \c IFontDisp object.
			///
			/// \sa pFontDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>Flag telling \c OnSetFont not to retrieve the current font if set to \c TRUE</em>
			///
			/// \sa OnSetFont
			UINT dontGetFontObject : 1;
			/// \brief <em>The control's current font</em>
			///
			/// \sa ApplyFont, owningFontResource
			CFont currentFont;
			/// \brief <em>If \c TRUE, \c currentFont may destroy the font resource; otherwise not</em>
			///
			/// \sa currentFont
			UINT owningFontResource : 1;
			/// \brief <em>A pointer to the font object's implementation of \c IFontDisp</em>
			IFontDisp* pFontDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<TabStrip> >* pPropertyNotifySink;

			FontProperty()
			{
				watching = 0;
				dontGetFontObject = FALSE;
				owningFontResource = TRUE;
				pFontDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~FontProperty()
			{
				Release();
			}

			FontProperty& operator =(const FontProperty& source)
			{
				Release();

				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				pFontDisp = source.pFontDisp;
				if(pFontDisp) {
					pFontDisp->AddRef();
				}
				owningFontResource = source.owningFontResource;
				if(!owningFontResource) {
					currentFont.Attach(source.currentFont.m_hFont);
				}
				dontGetFontObject = source.dontGetFontObject;

				if(source.watching > 0) {
					StartWatching();
				}

				return *this;
			}

			/// \brief <em>Retrieves a default font that may be used</em>
			///
			/// \return A \c FONTDESC structure containing the default font.
			///
			/// \sa defaultFont
			static FONTDESC GetDefaultFont(void)
			{
				return defaultFont;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(TabStrip* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<TabStrip> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(owningFontResource) {
					if(!currentFont.IsNull()) {
						currentFont.DeleteObject();
					}
				} else {
					currentFont.Detach();
				}
				if(pFontDisp) {
					pFontDisp->Release();
					pFontDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pFontDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} FontProperty;

		/// \brief <em>Holds a picture property's settings</em>
		typedef struct PictureProperty
		{
			/// \brief <em>Holds whether we're listening for events fired by the picture object</em>
			///
			/// If greater than 0, we're advised to the \c IPictureDisp object identified by \c pPictureDisp.
			/// I. e. we will be notified if a property of the picture object changes. If 0, we won't receive any
			/// events fired by the \c IPictureDisp object.
			///
			/// \sa pPictureDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>A pointer to the picture object's implementation of \c IPictureDisp</em>
			IPictureDisp* pPictureDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<TabStrip> >* pPropertyNotifySink;

			PictureProperty()
			{
				watching = 0;
				pPictureDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~PictureProperty()
			{
				Release();
			}

			PictureProperty& operator =(const PictureProperty& source)
			{
				Release();

				pPictureDisp = source.pPictureDisp;
				if(pPictureDisp) {
					pPictureDisp->AddRef();
				}
				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				if(source.watching > 0) {
					StartWatching();
				}
				return *this;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(TabStrip* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<TabStrip> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(pPictureDisp) {
					pPictureDisp->Release();
					pPictureDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pPictureDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} PictureProperty;

		/// \brief <em>The autoscroll-zone's default width</em>
		///
		/// The default width (in pixels) of the border around the control's client area, that's
		/// sensitive for auto-scrolling during a drag'n'drop operation. If the mouse cursor's position
		/// lies within this area during a drag'n'drop operation, the control will be auto-scrolled.
		///
		/// \sa dragScrollTimeBase, Raise_OLEDragMouseMove
		static const int DRAGSCROLLZONEWIDTH = 16;

		/// \brief <em>Holds the \c AllowDragDrop property's setting</em>
		///
		/// \sa get_AllowDragDrop, put_AllowDragDrop
		UINT allowDragDrop : 1;
		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c CloseableTabs property's setting</em>
		///
		/// \sa get_CloseableTabs, put_CloseableTabs
		UINT closeableTabs : 1;
		/// \brief <em>Holds the \c CloseableTabsMode property's setting</em>
		///
		/// \sa get_CloseableTabsMode, put_CloseableTabsMode
		CloseableTabsModeConstants closeableTabsMode;
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		// \brief <em>Holds the \c DontRedraw property's setting</em>
		//
		// \sa get_DontRedraw, put_DontRedraw
		//UINT dontRedraw : 1;
		/// \brief <em>Holds the \c DragActivateTime property's setting</em>
		///
		/// \sa get_DragActivateTime, put_DragActivateTime
		long dragActivateTime;
		/// \brief <em>Holds the \c DragScrollTimeBase property's setting</em>
		///
		/// \sa get_DragScrollTimeBase, put_DragScrollTimeBase
		long dragScrollTimeBase;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c FixedTabWidth property's setting</em>
		///
		/// \sa get_FixedTabWidth, put_FixedTabWidth
		OLE_XSIZE_PIXELS fixedTabWidth;
		/// \brief <em>Holds the \c FocusOnButtonDown property's setting</em>
		///
		/// \sa get_FocusOnButtonDown, put_FocusOnButtonDown
		UINT focusOnButtonDown : 1;
		/// \brief <em>Holds the \c Font property's settings</em>
		///
		/// \sa get_Font, put_Font, putref_Font
		FontProperty font;
		/// \brief <em>Holds the control's accelerator table</em>
		///
		/// \sa RebuildAcceleratorTable, GetControlInfo, OnMnemonic
		HACCEL hAcceleratorTable;
		/// \brief <em>Holds the \c HorizontalPadding property's setting</em>
		///
		/// \sa get_HorizontalPadding, put_HorizontalPadding
		OLE_XSIZE_PIXELS horizontalPadding;
		/// \brief <em>Holds the \c HotTracking property's setting</em>
		///
		/// \sa get_HotTracking, put_HotTracking
		UINT hotTracking : 1;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c InsertMarkColor property's setting</em>
		///
		/// \sa get_InsertMarkColor, put_InsertMarkColor
		OLE_COLOR insertMarkColor;
		/// \brief <em>Holds the \c MinTabWidth property's setting</em>
		///
		/// \sa get_MinTabWidth, put_MinTabWidth
		OLE_XSIZE_PIXELS minTabWidth;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
		/// \brief <em>Holds the \c MultiRow property's setting</em>
		///
		/// \sa get_MultiRow, put_MultiRow
		UINT multiRow : 1;
		/// \brief <em>Holds the \c MultiSelect property's setting</em>
		///
		/// \sa get_MultiSelect, put_MultiSelect
		UINT multiSelect : 1;
		/// \brief <em>Holds the \c OLEDragImageStyle property's setting</em>
		///
		/// \sa get_OLEDragImageStyle, put_OLEDragImageStyle
		OLEDragImageStyleConstants oleDragImageStyle;
		/// \brief <em>Holds the \c OwnerDrawn property's setting</em>
		///
		/// \sa get_OwnerDrawn, put_OwnerDrawn
		UINT ownerDrawn : 1;
		/// \brief <em>Holds the \c ProcessContextMenuKeys property's setting</em>
		///
		/// \sa get_ProcessContextMenuKeys, put_ProcessContextMenuKeys
		UINT processContextMenuKeys : 1;
		/// \brief <em>Holds the \c RaggedTabRows property's setting</em>
		///
		/// \sa get_RaggedTabRows, put_RaggedTabRows
		UINT raggedTabRows : 1;
		// \brief <em>Holds the \c ReflectContextMenuMessages property's setting</em>
		//
		// \sa get_ReflectContextMenuMessages, put_ReflectContextMenuMessages
		//UINT reflectContextMenuMessages : 1;
		/// \brief <em>Holds the \c RegisterForOLEDragDrop property's setting</em>
		///
		/// \sa get_RegisterForOLEDragDrop, put_RegisterForOLEDragDrop
		RegisterForOLEDragDropConstants registerForOLEDragDrop;
		/// \brief <em>Holds the \c RightToLeft property's setting</em>
		///
		/// \sa get_RightToLeft, put_RightToLeft
		RightToLeftConstants rightToLeft;
		/// \brief <em>Holds the \c ScrollToOpposite property's setting</em>
		///
		/// \sa get_ScrollToOpposite, put_ScrollToOpposite
		UINT scrollToOpposite : 1;
		/// \brief <em>Holds the \c ShowButtonSeparators property's setting</em>
		///
		/// \sa get_ShowButtonSeparators, put_ShowButtonSeparators
		UINT showButtonSeparators : 1;
		/// \brief <em>Holds the \c ShowToolTips property's setting</em>
		///
		/// \sa get_ShowToolTips, put_ShowToolTips
		UINT showToolTips : 1;
		/// \brief <em>Holds the \c Style property's setting</em>
		///
		/// \sa get_Style, put_Style
		StyleConstants style;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		/// \brief <em>Holds the \c TabBoundingBoxDefinition property's setting</em>
		///
		/// \sa get_TabBoundingBoxDefinition, put_TabBoundingBoxDefinition
		TabBoundingBoxDefinitionConstants tabBoundingBoxDefinition;
		/// \brief <em>Holds the \c TabCaptionAlignment property's setting</em>
		///
		/// \sa get_TabCaptionAlignment, put_TabCaptionAlignment
		TabCaptionAlignmentConstants tabCaptionAlignment;
		/// \brief <em>Holds the \c TabHeight property's setting</em>
		///
		/// \sa get_TabHeight, put_TabHeight
		OLE_YSIZE_PIXELS tabHeight;
		/// \brief <em>Holds the \c TabPlacement property's setting</em>
		///
		/// \sa get_TabPlacement, put_TabPlacement
		TabPlacementConstants tabPlacement;
		/// \brief <em>Holds the \c UseFixedTabWidth property's setting</em>
		///
		/// \sa get_UseFixedTabWidth, put_UseFixedTabWidth
		UINT useFixedTabWidth : 1;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;
		/// \brief <em>Holds the \c VerticalPadding property's setting</em>
		///
		/// \sa get_VerticalPadding, put_VerticalPadding
		OLE_YSIZE_PIXELS verticalPadding;

		Properties()
		{
			ResetToDefaults();
		}

		~Properties()
		{
			if(hAcceleratorTable) {
				DestroyAcceleratorTable(hAcceleratorTable);
			}
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			allowDragDrop = TRUE;
			appearance = a2D;
			borderStyle = bsNone;
			closeableTabs = FALSE;
			closeableTabsMode = ctmDisplayOnAllTabs;
			disabledEvents = static_cast<DisabledEventsConstants>(deMouseEvents | deClickEvents | deKeyboardEvents | deTabInsertionEvents | deTabDeletionEvents | deFreeTabData | deTabSelectionChanged);
			//dontRedraw = FALSE;
			dragActivateTime = -1;
			dragScrollTimeBase = -1;
			enabled = TRUE;
			fixedTabWidth = 96;
			focusOnButtonDown = FALSE;
			hAcceleratorTable = NULL;
			horizontalPadding = 6;
			hotTracking = FALSE;
			hoverTime = -1;
			insertMarkColor = 0;
			minTabWidth = -1;
			mousePointer = mpDefault;
			multiRow = FALSE;
			multiSelect = FALSE;
			oleDragImageStyle = odistClassic;
			ownerDrawn = FALSE;
			processContextMenuKeys = TRUE;
			raggedTabRows = TRUE;
			//reflectContextMenuMessages = TRUE;
			registerForOLEDragDrop = rfoddNoDragDrop;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			scrollToOpposite = FALSE;
			showButtonSeparators = TRUE;
			showToolTips = TRUE;
			style = sTabs;
			supportOLEDragImages = TRUE;
			tabBoundingBoxDefinition = static_cast<TabBoundingBoxDefinitionConstants>(tbbdTabIcon | tbbdTabLabel | tbbdTabCloseButton);
			tabCaptionAlignment = tcaNormal;
			tabHeight = 18;
			tabPlacement = tpTop;
			useFixedTabWidth = FALSE;
			useSystemFont = TRUE;
			verticalPadding = 3;
		}
	} Properties;
	/// \brief <em>Holds the control's properties' settings</em>
	Properties properties;

	/// \brief <em>Holds some frequently used settings</em>
	///
	/// Holds some settings that otherwise would be requested from the control window very often.
	/// The cached settings are updated whenever the corresponding control window's settings change.
	struct CachedSettings
	{
		/// \brief <em>Holds the index of the tab, that currently is highlighted as the control's drop target</em>
		///
		/// \sa get_DropHilitedTab, putref_DropHilitedTab, OnHighlightItem
		int dropHilitedTab;
		/// \brief <em>Holds the \c hHighResImageList property's setting</em>
		///
		/// \sa get_hHighResImageList, put_hHighResImageList
		HIMAGELIST hHighResImageList;
		/// \brief <em>Holds the \c hImageList property's setting</em>
		///
		/// \sa get_hImageList, put_hImageList, OnSetImageList
		HIMAGELIST hImageList;

		CachedSettings()
		{
			dropHilitedTab = -1;
			hHighResImageList = NULL;
			hImageList = NULL;
		}

		/// \brief <em>Fills the caches with the settings of the specified control window</em>
		///
		/// \param[in] hWndTabStrip The tabstrip window whose settings will be cached.
		void CacheSettings(HWND hWndTabStrip)
		{
			hHighResImageList = NULL;
			dropHilitedTab = -1;
			TCITEM tab = {0};
			tab.dwStateMask = TCIS_HIGHLIGHTED;
			tab.mask = TCIF_STATE;
			int numberOfTabs = ::SendMessage(hWndTabStrip, TCM_GETITEMCOUNT, 0, 0);
			for(int tabIndex = 0; tabIndex < numberOfTabs; ++tabIndex) {
				if(::SendMessage(hWndTabStrip, TCM_GETITEM, tabIndex, reinterpret_cast<LPARAM>(&tab))) {
					if(tab.dwState & TCIS_HIGHLIGHTED) {
						dropHilitedTab = tabIndex;
						break;
					}
				}
			}

			hImageList = reinterpret_cast<HIMAGELIST>(::SendMessage(hWndTabStrip, TCM_GETIMAGELIST, 0, 0));
		}
	} cachedSettings;

	/// \brief <em>Holds the control's flags</em>
	struct Flags
	{
		/// \brief <em>If not 0, we won't raise \c Insert*Tab events</em>
		///
		/// \sa silentTabDeletions, EnterSilentTabInsertionSection, LeaveSilentTabInsertionSection,
		///     OnInsertItem, Raise_InsertingTab, Raise_InsertedTab
		int silentTabInsertions;
		/// \brief <em>If not 0, we won't raise \c Remov*Tab events</em>
		///
		/// \sa silentTabInsertions, EnterSilentTabDeletionSection, LeaveSilentTabDeletionSection,
		///     OnDeleteItem, OnDeleteAllItems, Raise_RemovingTab, Raise_RemovedTab
		int silentTabDeletions;
		/// \brief <em>If not 0, we won't raise \c ActiveTabChang* events</em>
		///
		/// \sa silentCaretChanges, EnterSilentActiveTabChangeSection, LeaveSilentActiveTabChangeSection,
		///     Raise_ActiveTabChanging, Raise_ActiveTabChanged
		int silentActiveTabChanges;
		/// \brief <em>If not 0, we won't raise \c CaretChanged event</em>
		///
		/// \sa silentActiveTabChanges, silentSelectionChanges, EnterSilentCaretChangeSection,
		///     LeaveSilentCaretChangeSection, Raise_CaretChanged
		int silentCaretChanges;
		/// \brief <em>If not 0, we won't raise \c TabSelectionChanged event</em>
		///
		/// \sa silentCaretChanges, EnterSilentSelectionChangesSection, LeaveSilentSelectionChangesSection,
		///     Raise_TabSelectionChanged
		int silentSelectionChanges;
		/// \brief <em>If \c TRUE, we're using themes</em>
		///
		/// \sa OnThemeChanged
		UINT usingThemes : 1;

		Flags()
		{
			silentTabInsertions = 0;
			silentTabDeletions = 0;
			silentActiveTabChanges = 0;
			silentCaretChanges = 0;
			silentSelectionChanges = 0;
			usingThemes = FALSE;
		}
	} flags;

	//////////////////////////////////////////////////////////////////////
	/// \name Tab management
	///
	//@{
	/// \brief <em>Holds additional tab data</em>
	typedef struct TabData
	{
		/// \brief <em>Specifies an user-defined integer value associated with the tab</em>
		LPARAM lParam;
		/// \brief <em>Holds the handle of a window that is associated with the tab</em>
		HWND hAssociatedWindow;
		/// \brief <em>Specifies whether the tab can be closed</em>
		BOOL isCloseable;
	} TabData;
	#ifdef USE_STL
		/// \brief <em>A list of all tab IDs</em>
		///
		/// Holds the unique IDs of all tab strip tabs in the control. The tab's index in the control equals
		/// its index in the vector.
		///
		/// \sa OnInsertItem, OnDeleteAllItems, OnDeleteItem, OnGetItem, OnSetItem, TabStripTab::get_ID,
		///     MoveTabIndex
		std::vector<LONG> tabIDs;
		/// \brief <em>A list of all items</em>
		///
		/// Holds additional data for each tabstrip tab in the control.
		///
		/// \sa OnInsertItem, OnDeleteAllItems, OnDeleteItem, OnGetItem, OnSetItem, TabStripTab::get_ID,
		///     MoveTabIndex
		std::list<TabData> tabParams;
	#else
		/// \brief <em>A list of all tab IDs</em>
		///
		/// Holds the unique IDs of all tab strip tabs in the control. The tab's index in the control equals
		/// its index in the vector.
		///
		/// \sa OnInsertItem, OnDeleteAllItems, OnDeleteItem, OnGetItem, OnSetItem, TabStripTab::get_ID,
		///     MoveTabIndex
		CAtlArray<LONG> tabIDs;
		/// \brief <em>A list of all tabs</em>
		///
		/// Holds additional data for each tabstrip tab in the control.
		///
		/// \sa OnInsertItem, OnDeleteAllItems, OnDeleteItem, OnGetItem, OnSetItem, TabStripTab::get_ID,
		///     MoveTabIndex
		CAtlList<TabData> tabParams;
	#endif
	/// \brief <em>Retrieves a new unique tab ID at each call</em>
	///
	/// \return A new unique tab ID.
	///
	/// \sa tabIDs, TabStripTab::get_ID
	LONG GetNewTabID(void);
	#ifdef USE_STL
		/// \brief <em>A map of all \c TabStripTabContainer objects that we've created</em>
		///
		/// Holds pointers to all \c TabStripTabContainer objects that we've created. We use this map to
		/// inform the containers of removed tabs. The container's ID is stored as key; the container's
		/// \c ITabContainer implementation is stored as value.
		///
		/// \sa CreateTabContainer, RegisterTabContainer, TabStripTabContainer
		std::unordered_map<DWORD, ITabContainer*> tabContainers;
	#else
		/// \brief <em>A map of all \c TabStripTabContainer objects that we've created</em>
		///
		/// Holds pointers to all \c TabStripTabContainer objects that we've created. We use this map to
		/// inform the containers of removed tabs. The container's ID is stored as key; the container's
		/// \c ITabContainer implementation is stored as value.
		///
		/// \sa CreateTabContainer, RegisterTabContainer, TabStripTabContainer
		CAtlMap<DWORD, ITabContainer*> tabContainers;
	#endif
	/// \brief <em>Registers the specified \c TabStripTabContainer collection</em>
	///
	/// Registers the specified \c TabStripTabContainer collection so that it is informed of tab deletions.
	///
	/// \param[in] pContainer The container's \c ITabContainer implementation.
	///
	/// \sa DeregisterTabContainer, tabContainers, RemoveTabFromTabContainers
	void RegisterTabContainer(ITabContainer* pContainer);
	/// \brief <em>De-registers the specified \c TabStripTabContainer collection</em>
	///
	/// De-registers the specified \c TabStripTabContainer collection so that it no longer is informed of
	/// tab deletions.
	///
	/// \param[in] containerID The container's ID.
	///
	/// \sa RegisterTabContainer, tabContainers
	void DeregisterTabContainer(DWORD containerID);
	/// \brief <em>Removes the specified tab from all registered \c TabStripTabContainer collections</em>
	///
	/// \param[in] tabID The unique ID of the tab to remove. If -1, all tabs are removed.
	///
	/// \sa tabContainers, RegisterTabContainer, OnDeleteItem, Raise_DestroyedControlWindow
	void RemoveTabFromTabContainers(LONG tabID);
	///@}
	//////////////////////////////////////////////////////////////////////


	/// \brief <em>Holds the index of the tab below the mouse cursor</em>
	///
	/// \attention This member is not reliable with \c deMouseEvents being set.
	int tabUnderMouse;
	/// \brief <em>A buffer that is used in \c OnToolTipNeedTextNotification</em>
	///
	/// In \c OnToolTipNeedTextNotification we've the problem that we must allocate memory which can't
	/// be freed when leaving the method. To avoid memory leaks, we use a class-wide buffer which is
	/// allocated on object instantiation and freed on object destruction.
	///
	/// \sa OnToolTipGetDispInfoNotificationA, OnToolTipGetDispInfoNotificationW
	LPVOID pToolTipBuffer;

	/// \brief <em>Holds mouse status variables</em>
	typedef struct MouseStatus
	{
	protected:
		/// \brief <em>Holds all mouse buttons that may cause a click event in the immediate future</em>
		///
		/// A bit field of \c SHORT values representing those mouse buttons that are currently pressed and
		/// may cause a click event in the immediate future.
		///
		/// \sa StoreClickCandidate, IsClickCandidate, RemoveClickCandidate, Raise_Click, Raise_MClick,
		///     Raise_RClick, Raise_XClick
		SHORT clickCandidates;

	public:
		/// \brief <em>If \c TRUE, the \c MouseEnter event already was raised</em>
		///
		/// \sa Raise_MouseEnter
		UINT enteredControl : 1;
		/// \brief <em>If \c TRUE, the \c MouseHover event already was raised</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa Raise_MouseHover
		UINT hoveredControl : 1;
		/// \brief <em>Holds the mouse cursor's last position</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		POINT lastPosition;
		/// \brief <em>Holds the index of the last clicked tab</em>
		///
		/// Holds the index of the last clicked tab. We use this to ensure that the \c *DblClick events
		/// are not raised accidently.
		///
		/// \attention This member is not reliable with \c deClickEvents being set.
		///
		/// \sa Raise_Click, Raise_DblClick, Raise_MClick, Raise_MDblClick, Raise_RClick, Raise_RDblClick
		int lastClickedTab;
		/// \brief <em>If \c TRUE, the next \c NM_CLICK notification won't trigger a \c Click event</em>
		///
		/// \sa OnClickNotification
		UINT ignoreNM_CLICK : 1;
		/// \brief <em>If \c TRUE, the next \c NM_RCLICK notification won't trigger a \c RClick event</em>
		///
		/// \sa OnRClickNotification
		UINT ignoreNM_RCLICK : 1;
		/// \brief <em>The zero-based index of the tab over whose close button's bounding rectangle the mouse cursor is located</em>
		///
		/// \sa OnMouseMove, OnLButtonDown, OnLButtonUp, DrawCloseButtonsAndInsertionMarks
		int overCloseButton;
		/// \brief <em>The zero-based index of the tab over whose close button's bounding rectangle the mouse cursor has been located during last mouse down</em>
		///
		/// \sa OnLButtonDown, Raise_Click, HitTest, OnClickNotification
		int overCloseButtonOnMouseDown;

		MouseStatus()
		{
			clickCandidates = 0;
			enteredControl = FALSE;
			hoveredControl = FALSE;
			lastClickedTab = -1;
			ignoreNM_CLICK = FALSE;
			ignoreNM_RCLICK = FALSE;
			overCloseButton = -1;
			overCloseButtonOnMouseDown = -1;
		}

		/// \brief <em>Changes flags to indicate the \c MouseEnter event was just raised</em>
		///
		/// \sa enteredControl, HoverControl, LeaveControl
		void EnterControl(void)
		{
			RemoveAllClickCandidates();
			enteredControl = TRUE;
			lastClickedTab = -1;
			ignoreNM_CLICK = FALSE;
			ignoreNM_RCLICK = FALSE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseHover event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl, LeaveControl
		void HoverControl(void)
		{
			enteredControl = TRUE;
			hoveredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseLeave event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl
		void LeaveControl(void)
		{
			enteredControl = FALSE;
			hoveredControl = FALSE;
			lastClickedTab = -1;
		}

		/// \brief <em>Stores a mouse button as click candidate</em>
		///
		/// param[in] button The mouse button to store.
		///
		/// \sa clickCandidates, IsClickCandidate, RemoveClickCandidate
		void StoreClickCandidate(SHORT button)
		{
			// avoid combined click events
			if(clickCandidates == 0) {
				clickCandidates |= button;
			}
		}

		/// \brief <em>Retrieves whether a mouse button is a click candidate</em>
		///
		/// \param[in] button The mouse button to check.
		///
		/// \return \c TRUE if the button is stored as a click candidate; otherwise \c FALSE.
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa clickCandidates, StoreClickCandidate, RemoveClickCandidate
		BOOL IsClickCandidate(SHORT button)
		{
			return (clickCandidates & button);
		}

		/// \brief <em>Removes a mouse button from the list of click candidates</em>
		///
		/// \param[in] button The mouse button to remove.
		///
		/// \sa clickCandidates, RemoveAllClickCandidates, StoreClickCandidate, IsClickCandidate
		void RemoveClickCandidate(SHORT button)
		{
			clickCandidates &= ~button;
		}

		/// \brief <em>Clears the list of click candidates</em>
		///
		/// \sa clickCandidates, RemoveClickCandidate, StoreClickCandidate, IsClickCandidate
		void RemoveAllClickCandidates(void)
		{
			clickCandidates = 0;
		}
	} MouseStatus;

	/// \brief <em>Holds the control's mouse status</em>
	MouseStatus mouseStatus;

	/// \brief <em>Holds the position of the control's insertion mark</em>
	struct InsertMark
	{
		/// \brief <em>The zero-based index of the tab at which the insertion mark is placed</em>
		int tabIndex;
		/// \brief <em>The insertion mark's position relative to the tab</em>
		BOOL afterTab;
		/// \brief <em>The insertion mark's color</em>
		COLORREF color;

		InsertMark()
		{
			Reset();
		}

		/// \brief <em>Resets all member variables</em>
		void Reset(void)
		{
			tabIndex = -1;
			afterTab = FALSE;
			color = RGB(0, 0, 0);
		}

		/// \brief <em>Processes the parameters of a \c TCM_SETINSERTMARK message to store the new insertion mark</em>
		///
		/// \sa OnSetInsertMark
		void ProcessSetInsertMark(WPARAM wParam, LPARAM lParam)
		{
			afterTab = static_cast<BOOL>(wParam);
			tabIndex = static_cast<int>(lParam);
		}
	} insertMark;

	//////////////////////////////////////////////////////////////////////
	/// \name Drag'n'Drop
	///
	//@{
	/// \brief <em>The \c CLSID_WICImagingFactory object used to create WIC objects that are required during drag image creation</em>
	///
	/// \sa OnGetDragImage, CreateThumbnail
	CComPtr<IWICImagingFactory> pWICImagingFactory;
	/// \brief <em>Creates a thumbnail of the specified icon in the specified size</em>
	///
	/// \param[in] hIcon The icon to create the thumbnail for.
	/// \param[in] size The thumbnail's size in pixels.
	/// \param[in,out] pBits The thumbnail's DIB bits.
	/// \param[in] doAlphaChannelPostProcessing WIC has problems to handle the alpha channel of the icon
	///            specified by \c hIcon. If this parameter is set to \c TRUE, some post-processing is done
	///            to correct the pixel failures. Otherwise the failures are not corrected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnGetDragImage, pWICImagingFactory
	HRESULT CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing);

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		/// \brief <em>The \c ITabStripTabContainer implementation of the collection of the dragged tabs</em>
		ITabStripTabContainer* pDraggedTabs;
		/// \brief <em>The handle of the imagelist containing the drag image</em>
		///
		/// \sa get_hDragImageList
		HIMAGELIST hDragImageList;
		/// \brief <em>Enables or disables auto-destruction of \c hDragImageList</em>
		///
		/// Controls whether the imagelist defined by \c hDragImageList is destroyed automatically. If set to
		/// \c TRUE, it is destroyed in \c EndDrag; otherwise not.
		///
		/// \sa hDragImageList, EndDrag
		UINT autoDestroyImgLst : 1;
		/// \brief <em>Indicates whether the drag image is visible or hidden</em>
		///
		/// If this value is 0, the drag image is visible; otherwise not.
		///
		/// \sa get_hDragImageList, get_ShowDragImage, put_ShowDragImage, ShowDragImage, HideDragImage,
		///     IsDragImageVisible
		int dragImageIsHidden;
		/// \brief <em>The zero-based index of the last drop target</em>
		int lastDropTarget;

		//////////////////////////////////////////////////////////////////////
		/// \name OLE Drag'n'Drop
		///
		//@{
		/// \brief <em>The currently dragged data</em>
		CComPtr<IOLEDataObject> pActiveDataObject;
		/// \brief <em>The currently dragged data for the case that the we're the drag source</em>
		CComPtr<IDataObject> pSourceDataObject;
		/// \brief <em>Holds the mouse cursors last position (in screen coordinates)</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
		/// \brief <em>Holds the mouse button (as \c MK_* constant) that the drag'n'drop operation is performed with</em>
		DWORD draggingMouseButton;
		/// \brief <em>If \c TRUE, we'll hide and re-show the drag image in \c IDropTarget::DragEnter so that the item count label is displayed</em>
		///
		/// \sa DragEnter, OLEDrag
		UINT useItemCountLabelHack : 1;
		/// \brief <em>Holds the \c IDataObject to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		IDataObject* drop_pDataObject;
		/// \brief <em>Holds the mouse position to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		POINT drop_mousePosition;
		/// \brief <em>Holds the drop effect to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		DWORD drop_effect;
		//@}
		//////////////////////////////////////////////////////////////////////

		/// \brief <em>Holds data and flags related to auto-scrolling</em>
		///
		/// \sa AutoScroll
		struct AutoScrolling
		{
			/// \brief <em>Holds the current speed multiplier used for auto-scrolling</em>
			LONG currentScrollVelocity;
			/// \brief <em>Holds the current interval of the auto-scroll timer</em>
			LONG currentTimerInterval;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the left</em>
			DWORD lastScroll_Left;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the right</em>
			DWORD lastScroll_Right;

			AutoScrolling()
			{
				Reset();
			}

			/// \brief <em>Resets all member variables to their defaults</em>
			void Reset(void)
			{
				currentScrollVelocity = 0;
				currentTimerInterval = 0;
				lastScroll_Left = 0;
				lastScroll_Right = 0;
			}
		} autoScrolling;

		/// \brief <em>Holds data required for drag-detection</em>
		struct Candidate
		{
			/// \brief <em>The zero-based index of the tab, that might be dragged</em>
			int tabIndex;
			/// \brief <em>The position in pixels at which dragging might start</em>
			POINT position;
			/// \brief <em>The mouse button (as \c MK_* constant) that the drag'n'drop operation might be performed with</em>
			int button;

			Candidate()
			{
				tabIndex = -1;
			}
		} candidate;

		DragDropStatus()
		{
			pActiveDataObject = NULL;
			pSourceDataObject = NULL;
			pDropTargetHelper = NULL;
			draggingMouseButton = 0;

			pDraggedTabs = NULL;
			hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;
			lastDropTarget = -1;
			drop_pDataObject = NULL;
		}

		~DragDropStatus()
		{
			if(pDropTargetHelper) {
				pDropTargetHelper->Release();
			}
			ATLASSERT(!pDraggedTabs);
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			if(this->pDraggedTabs) {
				this->pDraggedTabs->Release();
				this->pDraggedTabs = NULL;
			}
			if(hDragImageList && autoDestroyImgLst) {
				ImageList_Destroy(hDragImageList);
			}
			hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;
			lastDropTarget = -1;

			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			if(this->pSourceDataObject) {
				this->pSourceDataObject = NULL;
			}
			draggingMouseButton = 0;
			useItemCountLabelHack = FALSE;
			drop_pDataObject = NULL;
		}

		/// \brief <em>Decrements the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, HideDragImage, IsDragImageVisible
		void ShowDragImage(BOOL commonDragDropOnly)
		{
			if(hDragImageList) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					ImageList_DragShowNolock(TRUE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					pDropTargetHelper->Show(TRUE);
				}
			}
		}

		/// \brief <em>Increments the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, ShowDragImage, IsDragImageVisible
		void HideDragImage(BOOL commonDragDropOnly)
		{
			if(hDragImageList) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					ImageList_DragShowNolock(FALSE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					pDropTargetHelper->Show(FALSE);
				}
			}
		}

		/// \brief <em>Retrieves whether we're currently displaying a drag image</em>
		///
		/// \return \c TRUE, if we're displaying a drag image; otherwise \c FALSE.
		///
		/// \sa dragImageIsHidden, ShowDragImage, HideDragImage
		BOOL IsDragImageVisible(void)
		{
			return (dragImageIsHidden == 0);
		}

		/// \brief <em>Performs any tasks that must be done after a drag'n'drop operation started</em>
		///
		/// \param[in] hWndTabStrip The tabstrip window, that the method will work on to calculate the position
		///            of the drag image's hotspot.
		/// \param[in] pDraggedTabs The \c ITabStripTabContainer implementation of the collection of
		///            the dragged tabs.
		/// \param[in] hDragImageList The imagelist containing the drag image that shall be used to
		///            visualize the drag'n'drop operation. If -1, the method will create the drag image
		///            itself; if \c NULL, no drag image will be displayed.
		/// \param[in,out] pXHotSpot The x-coordinate (in pixels) of the drag image's hotspot relative to the
		///                drag image's upper-left corner. If the \c hDragImageList parameter is set to
		///                \c NULL, this parameter is ignored. If the \c hDragImageList parameter is set to
		///                -1, this parameter is set to the hotspot calculated by the method.
		/// \param[in,out] pYHotSpot The y-coordinate (in pixels) of the drag image's hotspot relative to the
		///                drag image's upper-left corner. If the \c hDragImageList parameter is set to
		///                \c NULL, this parameter is ignored. If the \c hDragImageList parameter is set to
		///                -1, this parameter is set to the hotspot calculated by the method.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa EndDrag
		HRESULT BeginDrag(HWND hWndTabStrip, ITabStripTabContainer* pDraggedTabs, HIMAGELIST hDragImageList, PINT pXHotSpot, PINT pYHotSpot)
		{
			ATLASSUME(pDraggedTabs);
			if(!pDraggedTabs) {
				return E_INVALIDARG;
			}

			UINT b = FALSE;
			if(hDragImageList == static_cast<HIMAGELIST>(LongToHandle(-1))) {
				OLE_HANDLE h = NULL;
				OLE_XPOS_PIXELS xUpperLeft = 0;
				OLE_YPOS_PIXELS yUpperLeft = 0;
				if(FAILED(pDraggedTabs->CreateDragImage(&xUpperLeft, &yUpperLeft, &h))) {
					return E_FAIL;
				}
				hDragImageList = static_cast<HIMAGELIST>(LongToHandle(h));
				b = TRUE;

				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				::ScreenToClient(hWndTabStrip, &mousePosition);
				if(CWindow(hWndTabStrip).GetExStyle() & WS_EX_LAYOUTRTL) {
					SIZE dragImageSize = {0};
					ImageList_GetIconSize(hDragImageList, reinterpret_cast<PINT>(&dragImageSize.cx), reinterpret_cast<PINT>(&dragImageSize.cy));
					*pXHotSpot = xUpperLeft + dragImageSize.cx - mousePosition.x;
				} else {
					*pXHotSpot = mousePosition.x - xUpperLeft;
				}
				*pYHotSpot = mousePosition.y - yUpperLeft;
			}

			if(this->hDragImageList && this->autoDestroyImgLst) {
				ImageList_Destroy(this->hDragImageList);
			}

			this->autoDestroyImgLst = b;
			this->hDragImageList = hDragImageList;
			if(this->pDraggedTabs) {
				this->pDraggedTabs->Release();
				this->pDraggedTabs = NULL;
			}
			pDraggedTabs->Clone(&this->pDraggedTabs);
			ATLASSUME(this->pDraggedTabs);
			this->lastDropTarget = -1;

			dragImageIsHidden = 1;
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done after a drag'n'drop operation ended</em>
		///
		/// \sa BeginDrag
		void EndDrag(void)
		{
			if(pDraggedTabs) {
				pDraggedTabs->Release();
				pDraggedTabs = NULL;
			}
			if(autoDestroyImgLst && hDragImageList) {
				ImageList_Destroy(hDragImageList);
			}
			hDragImageList = NULL;
			dragImageIsHidden = 1;
			lastDropTarget = -1;
			autoScrolling.Reset();
		}

		/// \brief <em>Retrieves whether we're in drag'n'drop mode</em>
		///
		/// \return \c TRUE if we're in drag'n'drop mode; otherwise \c FALSE.
		///
		/// \sa BeginDrag, EndDrag
		BOOL IsDragging(void)
		{
			return (pDraggedTabs != NULL);
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop
		HRESULT OLEDragEnter(void)
		{
			this->lastDropTarget = -1;
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter
		void OLEDragLeaveOrDrop(void)
		{
			this->lastDropTarget = -1;
			autoScrolling.Reset();
		}
	} dragDropStatus;
	///@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds data and flags related to the \c CaretChanged event</em>
	///
	/// \sa Raise_CaretChanged
	struct CaretChangedStatus
	{
		/// \brief <em>The previous caret tab</em>
		int previousCaretTab;

		CaretChangedStatus()
		{
			previousCaretTab = -1;
		}
	} caretChangedStatus;

	/// \brief <em>Holds data and flags related to the \c ActiveTabChanged event</em>
	///
	/// \sa Raise_ActiveTabChanged
	struct ActiveTabChangedStatus
	{
		/// \brief <em>The previous active tab</em>
		int previousActiveTab;
		#ifdef USE_STL
			/// \brief <em>Holds the tabs' selection states as they're set before the active tab change</em>
			std::vector<DWORD> selectionStatesBeforeChange;
		#else
			/// \brief <em>Holds the tabs' selection states as they're set before the active tab change</em>
			CAtlArray<DWORD> selectionStatesBeforeChange;
		#endif

		ActiveTabChangedStatus()
		{
			previousActiveTab = -1;
		}
	} activeTabChangedStatus;

	/// \brief <em>The container's implementation of \c ISimpleFrameSite</em>
	CComPtr<ISimpleFrameSite> pSimpleFrameSite;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		// \brief <em>The ID of the timer that is used to redraw the control after recreation</em>
		//static const UINT_PTR ID_REDRAW = 12;
		/// \brief <em>The ID of the timer that is used to auto-scroll the control window during drag'n'drop</em>
		static const UINT_PTR ID_DRAGSCROLL = 13;
		/// \brief <em>The ID of the timer that is used to auto-activate tabs during drag'n'drop</em>
		static const UINT_PTR ID_DRAGACTIVATE = 14;

		// \brief <em>The interval of the timer that is used to redraw the control after recreation</em>
		//static const UINT INT_REDRAW = 10;
	} timers;

private:
};     // TabStrip

OBJECT_ENTRY_AUTO(__uuidof(TabStrip), TabStrip)