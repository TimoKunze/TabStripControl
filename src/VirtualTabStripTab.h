//////////////////////////////////////////////////////////////////////
/// \class VirtualTabStripTab
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing tabstrip tab</em>
///
/// This class is a wrapper around a tabstrip tab that does not yet or not anymore exist in the control.
///
/// \if UNICODE
///   \sa TabStripCtlLibU::IVirtualTabStripTab, TabStripTab, TabStrip
/// \else
///   \sa TabStripCtlLibA::IVirtualTabStripTab, TabStripTab, TabStrip
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TabStripCtlU.h"
#else
	#include "TabStripCtlA.h"
#endif
#include "_IVirtualTabStripTabEvents_CP.h"
#include "helpers.h"
#include "TabStrip.h"


class ATL_NO_VTABLE VirtualTabStripTab : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualTabStripTab, &CLSID_VirtualTabStripTab>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualTabStripTab>,
    public Proxy_IVirtualTabStripTabEvents<VirtualTabStripTab>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualTabStripTab, &IID_IVirtualTabStripTab, &LIBID_TabStripCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualTabStripTab, &IID_IVirtualTabStripTab, &LIBID_TabStripCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class TabStrip;
	friend class TabStripTab;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALTABSTRIPTAB)

		BEGIN_COM_MAP(VirtualTabStripTab)
			COM_INTERFACE_ENTRY(IVirtualTabStripTab)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualTabStripTab)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualTabStripTabEvents))
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
	/// \name Implementation of IVirtualTabStripTab
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c DropHilited property</em>
	///
	/// Retrieves whether the tab will be or was drawn as the target of a drag'n'drop operation, i. e.
	/// whether its background will be or was highlighted. If this property is set to \c VARIANT_TRUE,
	/// the tab will be or was highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TabStrip::get_DropHilitedTab, get_Selected
	virtual HRESULT STDMETHODCALLTYPE get_DropHilited(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the tab's icon in the control's imagelist identified by
	/// \c TabStrip::get_hImageList. If set to -1, no icon is displayed for this tab.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TabStrip::get_hImageList
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index that will identify or has identified the tab.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TabStripTabs::Add
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the tab will be or was drawn as a selected tab. If this property is set to
	/// \c VARIANT_TRUE, the tab will be or was selected; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_DropHilited, TabStrip::get_MultiSelect, TabStrip::get_ActiveTab, TabStrip::get_CaretTab
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c TabData property</em>
	///
	/// Retrieves the \c LONG value that will be or was associated with the tab.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TabStrip::Raise_FreeTabData
	virtual HRESULT STDMETHODCALLTYPE get_TabData(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the tab's text. The maximum number of characters in this text is specified by
	/// \c MAX_TABTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IVirtualTabStripTab interface.\n
	///          This property is read-only.
	///
	/// \sa TabStripTabs::Add, MAX_TABTEXTLENGTH
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Initializes this object with given data</em>
	///
	/// Initializes this object with the settings from a given source.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	/// \param[in] tabIndex The tab's zero-based index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(__in LPTCITEM pSource, int tabIndex);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[out] tabIndex The tab's zero-based index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(__in LPTCITEM pTarget, int& tabIndex);
	/// \brief <em>Sets the owner of this tab</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerTabStrip
	void SetOwner(__in_opt TabStrip* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c TabStrip object that owns this tab</em>
		///
		/// \sa SetOwner
		TabStrip* pOwnerTabStrip;
		/// \brief <em>A structure holding this tab's settings</em>
		TCITEM settings;
		/// \brief <em>The index of the tab wrapped by this object</em>
		int tabIndex;

		Properties()
		{
			pOwnerTabStrip = NULL;
			ZeroMemory(&settings, sizeof(settings));
			tabIndex = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning tabstrip's window handle</em>
		///
		/// \return The window handle of the tabstrip that will contain or has contained this tab.
		///
		/// \sa pOwnerTabStrip
		HWND GetTabStripHWnd(void);
	} properties;
};     // VirtualTabStripTab

OBJECT_ENTRY_AUTO(__uuidof(VirtualTabStripTab), VirtualTabStripTab)