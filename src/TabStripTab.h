//////////////////////////////////////////////////////////////////////
/// \class TabStripTab
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing tabstrip tab</em>
///
/// This class is a wrapper around a tabstrip tab that - unlike a tab wrapped by
/// \c VirtualTabStripTab - really exists in the control.
///
/// \if UNICODE
///   \sa TabStripCtlLibU::ITabStripTab, VirtualTabStripTab, TabStripTabs, TabStrip
/// \else
///   \sa TabStripCtlLibA::ITabStripTab, VirtualTabStripTab, TabStripTabs, TabStrip
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TabStripCtlU.h"
#else
	#include "TabStripCtlA.h"
#endif
#include "_ITabStripTabEvents_CP.h"
#include "helpers.h"
#include "TabStrip.h"


class ATL_NO_VTABLE TabStripTab : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TabStripTab, &CLSID_TabStripTab>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TabStripTab>,
    public Proxy_ITabStripTabEvents<TabStripTab>, 
    #ifdef UNICODE
    	public IDispatchImpl<ITabStripTab, &IID_ITabStripTab, &LIBID_TabStripCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<ITabStripTab, &IID_ITabStripTab, &LIBID_TabStripCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class TabStrip;
	friend class TabStripTabs;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TABSTRIPTAB)

		BEGIN_COM_MAP(TabStripTab)
			COM_INTERFACE_ENTRY(ITabStripTab)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TabStripTab)
			CONNECTION_POINT_ENTRY(__uuidof(_ITabStripTabEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
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
	/// \name Implementation of ITabStripTab
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Active property</em>
	///
	/// Retrieves whether the tab is the control's active tab, i. e. it's the tab whose content is currently
	/// displayed. If it is the active tab, this property is set to \c VARIANT_TRUE; otherwise it's set to
	/// \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Caret, get_Selected, TabStrip::get_ActiveTab
	virtual HRESULT STDMETHODCALLTYPE get_Active(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the tab is the control's caret tab, i. e. it has the keyboard focus. If it is the
	/// caret tab, this property is set to \c VARIANT_TRUE; otherwise it's set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Active, get_Selected, TabStrip::get_CaretTab
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c DropHilited property</em>
	///
	/// Retrieves whether the tab is drawn as the target of a drag'n'drop operation, i. e. whether its
	/// background is highlighted. If this property is set to \c VARIANT_TRUE, the tab is highlighted;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TabStrip::get_DropHilitedTab, get_Selected
	virtual HRESULT STDMETHODCALLTYPE get_DropHilited(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c HasCloseButton property</em>
	///
	/// Retrieves whether a close button is drawn on the tab. If set to \c VARIANT_TRUE, the tab has a close
	/// button; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HasCloseButton, TabStrip::get_CloseableTabs
	virtual HRESULT STDMETHODCALLTYPE get_HasCloseButton(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c HasCloseButton property</em>
	///
	/// Sets whether a close button is drawn on the tab. If set to \c VARIANT_TRUE, the tab has a close
	/// button; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HasCloseButton, TabStrip::put_CloseableTabs
	virtual HRESULT STDMETHODCALLTYPE put_HasCloseButton(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c hAssociatedWindow property</em>
	///
	/// Retrieves the handle of the window associated with the tab. The window will be made visible
	/// automatically if the tab is selected and it will be hidden automatically if the tab is deselected.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hAssociatedWindow, TabStrip::GetAssociatedWindow
	virtual HRESULT STDMETHODCALLTYPE get_hAssociatedWindow(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hAssociatedWindow property</em>
	///
	/// Sets the handle of the window associated with the tab. The window will be made visible automatically
	/// if the tab is selected and it will be hidden automatically if the tab is deselected.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_hAssociatedWindow, TabStrip::SetAssociatedWindow
	virtual HRESULT STDMETHODCALLTYPE put_hAssociatedWindow(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c Height property</em>
	///
	/// Retrieves the height (in pixels) of the tab's bounding rectangle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Left, get_Top, get_Width, TabStrip::get_DisplayAreaHeight
	virtual HRESULT STDMETHODCALLTYPE get_Height(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the tab's icon in the control's imagelist identified by
	/// \c TabStrip::get_hImageList. If set to -1, no icon is displayed for this tab.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_IconIndex, TabStrip::get_hImageList
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Sets the \c IconIndex property</em>
	///
	/// Sets the zero-based index of the tab's icon in the control's imagelist identified by
	/// \c TabStrip::get_hImageList. If set to -1, no icon is displayed for this tab.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_IconIndex, TabStrip::put_hImageList
	virtual HRESULT STDMETHODCALLTYPE put_IconIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this tab.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks A tab's ID will never change.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Index, TabStripCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa get_Index, TabStripCtlLibA::TabIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index identifying this tab.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although adding or removing tabs changes other tabs' indexes, the index is the best
	///          (and fastest) option to identify a tab.
	///
	/// \if UNICODE
	///   \sa put_Index, get_ID, TabStripCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa put_Index, get_ID, TabStripCtlLibA::TabIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Sets the \c Index property</em>
	///
	/// Sets the zero-based index identifying this tab. Setting this property moves the tab.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although adding or removing tabs changes other tabs' indexes, the index is the best
	///          (and fastest) option to identify a tab.
	///
	/// \if UNICODE
	///   \sa get_Index, get_ID, TabStripCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa get_Index, get_ID, TabStripCtlLibA::TabIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Index(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Left property</em>
	///
	/// Retrieves the distance (in pixels) between the control's left border and the left border of the
	/// tab's bounding rectangle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Top, get_Height, get_Width
	virtual HRESULT STDMETHODCALLTYPE get_Left(OLE_XPOS_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the tab is drawn as a selected tab. If this property is set to \c VARIANT_TRUE,
	/// the tab is selected; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Selected, get_Active, get_Caret, get_DropHilited, TabStrip::get_MultiSelect,
	///     TabStrip::get_ActiveTab, TabStrip::get_CaretTab
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Selected property</em>
	///
	/// Sets whether the tab is drawn as a selected tab. If this property is set to \c VARIANT_TRUE,
	/// the tab is selected; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Selected, TabStrip::get_MultiSelect, TabStrip::putref_ActiveTab, TabStrip::putref_CaretTab
	virtual HRESULT STDMETHODCALLTYPE put_Selected(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c TabData property</em>
	///
	/// Retrieves the \c LONG value associated with the tab. Use this property to associate any data
	/// with the tab.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_TabData, TabStrip::Raise_FreeTabData
	virtual HRESULT STDMETHODCALLTYPE get_TabData(LONG* pValue);
	/// \brief <em>Sets the \c TabData property</em>
	///
	/// Sets the \c LONG value associated with the tab. Use this property to associate any data
	/// with the tab.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_TabData, TabStrip::Raise_FreeTabData
	virtual HRESULT STDMETHODCALLTYPE put_TabData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the tab's text. The maximum number of characters in this text is specified by
	/// \c MAX_TABTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, TabStripTabs::Add, MAX_TABTEXTLENGTH
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the tab's text. The maximum number of characters in this text is specified by
	/// \c MAX_TABTEXTLENGTH.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, TabStripTabs::Add, MAX_TABTEXTLENGTH
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c Top property</em>
	///
	/// Retrieves the distance (in pixels) between the control's top border and the top border of the
	/// tab's bounding rectangle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Left, get_Height, get_Width
	virtual HRESULT STDMETHODCALLTYPE get_Top(OLE_YPOS_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c Width property</em>
	///
	/// Retrieves the width (in pixels) of the tab's bounding rectangle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Left, get_Top, get_Height, TabStrip::get_DisplayAreaWidth
	virtual HRESULT STDMETHODCALLTYPE get_Width(OLE_XSIZE_PIXELS* pValue);

	/// \brief <em>Retrieves an imagelist containing the tab's drag image</em>
	///
	/// Retrieves the handle to an imagelist containing a bitmap that can be used to visualize
	/// dragging of this tab.
	///
	/// \param[out] pXUpperLeft The x-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] pYUpperLeft The y-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] phImageList The imagelist containing the drag image.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa TabStrip::CreateLegacyDragImage
	virtual HRESULT STDMETHODCALLTYPE CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft = NULL, OLE_YPOS_PIXELS* pYUpperLeft = NULL, OLE_HANDLE* phImageList = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given tab</em>
	///
	/// Attaches this object to a given tab, so that the tab's properties can be retrieved and set
	/// using this object's methods.
	///
	/// \param[in] tabIndex The tab to attach to.
	///
	/// \sa Detach
	void Attach(int tabIndex);
	/// \brief <em>Detaches this object from a tab</em>
	///
	/// Detaches this object from the tab it currently wraps, so that it doesn't wrap any tab anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// Applies the settings from a given source to the tab wrapped by this object.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(LPTCITEM pSource);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// \overload
	HRESULT LoadState(VirtualTabStripTab* pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndTabStrip The tabstrip window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(LPTCITEM pTarget, HWND hWndTabStrip = NULL);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \overload
	HRESULT SaveState(VirtualTabStripTab* pTarget);
	/// \brief <em>Sets the owner of this item</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerTabStrip
	void SetOwner(__in_opt TabStrip* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c TabStrip object that owns this item</em>
		///
		/// \sa SetOwner
		TabStrip* pOwnerTabStrip;
		/// \brief <em>The index of the tab wrapped by this object</em>
		int tabIndex;

		Properties()
		{
			pOwnerTabStrip = NULL;
			tabIndex = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning tabstrip's window handle</em>
		///
		/// \return The window handle of the tabstrip that contains this tab.
		///
		/// \sa pOwnerTabStrip
		HWND GetTabStripHWnd(void);
	} properties;

	/// \brief <em>Writes a given object's settings to a given target</em>
	///
	/// \param[in] tabIndex The tab for which to save the settings.
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndTabStrip The tabstrip window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(int tabIndex, LPTCITEM pTarget, HWND hWndTabStrip = NULL);
};     // TabStripTab

OBJECT_ENTRY_AUTO(__uuidof(TabStripTab), TabStripTab)