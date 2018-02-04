//////////////////////////////////////////////////////////////////////
/// \class TabStripTabContainer
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c TabStripTab objects</em>
///
/// This class provides easy access to collections of \c TabStripTab objects. While a
/// \c TabStripTabs object is used to group tabs that have certain properties in common, a
/// \c TabStripTabContainer object is used to group any tabs and acts more like a clipboard.
///
/// \if UNICODE
///   \sa TabStripCtlLibU::ITabStripTabContainer, TabStripTab, TabStripTabs, TabStrip
/// \else
///   \sa TabStripCtlLibA::ITabStripTabContainer, TabStripTab, TabStripTabs, TabStrip
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TabStripCtlU.h"
#else
	#include "TabStripCtlA.h"
#endif
#include "_ITabStripTabContainerEvents_CP.h"
#include "ITabContainer.h"
#include "helpers.h"
#include "TabStrip.h"
#include "TabStripTab.h"


class ATL_NO_VTABLE TabStripTabContainer : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TabStripTabContainer, &CLSID_TabStripTabContainer>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TabStripTabContainer>,
    public Proxy_ITabStripTabContainerEvents<TabStripTabContainer>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<ITabStripTabContainer, &IID_ITabStripTabContainer, &LIBID_TabStripCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<ITabStripTabContainer, &IID_ITabStripTabContainer, &LIBID_TabStripCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public ITabContainer
{
	friend class TabStrip;
	friend class TabStripTabContainer;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used to generate and set this object's ID.
	///
	/// \sa ~TabStripTabContainer
	TabStripTabContainer();
	/// \brief <em>The destructor of this class</em>
	///
	/// Used to deregister the container.
	///
	/// \sa TabStripTabContainer
	~TabStripTabContainer();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TABSTRIPTABCONTAINER)

		BEGIN_COM_MAP(TabStripTabContainer)
			COM_INTERFACE_ENTRY(ITabStripTabContainer)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TabStripTabContainer)
			CONNECTION_POINT_ENTRY(__uuidof(_ITabStripTabContainerEvents))
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
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the tabs</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c TabStripTab objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x tabs</em>
	///
	/// Retrieves the next \c numberOfMaxTabs tabs from the iterator.
	///
	/// \param[in] numberOfMaxTabs The maximum number of tabs the array identified by \c pTabs can
	///            contain.
	/// \param[in,out] pTabs An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a tab's \c ITabStripTab implementation.
	/// \param[out] pNumberOfTabsReturned The number of tabs that actually were copied to the array
	///             identified by \c pTabs.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, TabStripTab,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxTabs, VARIANT* pTabs, ULONG* pNumberOfTabsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// tab in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x tabs</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfTabsToSkip tabs.
	///
	/// \param[in] numberOfTabsToSkip The number of tabs to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfTabsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ITabStripTabContainer
	///
	//@{
	/// \brief <em>Retrieves a \c TabStripTab object from the collection</em>
	///
	/// Retrieves a \c TabStripTab object from the collection that wraps the tabstrip tab identified
	/// by \c tabID.
	///
	/// \param[in] tabID The unique ID of the tabstrip tab to retrieve.
	/// \param[out] ppTab Receives the item's \c ITabStripTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TabStripTab::get_ID, Add, Remove
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG tabID, ITabStripTab** ppTab);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c TabStripTab objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds the specified tab(s) to the collection</em>
	///
	/// \param[in] tabsToAdd The tab(s) to add. May be either a tab ID, a \c TabStripTab object or a
	///            \c TabStripTabs collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa TabStripTab::get_ID, Count, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Add(VARIANT tabsToAdd);
	/// \brief <em>Clones the collection object</em>
	///
	/// Retrieves an exact copy of the collection.
	///
	/// \param[out] ppClone Receives the clone's \c ITabStripTabContainer implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa TabStrip::CreateTabContainer
	virtual HRESULT STDMETHODCALLTYPE Clone(ITabStripTabContainer** ppClone);
	/// \brief <em>Counts the tabs in the collection</em>
	///
	/// Retrieves the number of \c TabStripTab objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Retrieves an imagelist containing the tabs' common drag image</em>
	///
	/// Retrieves the handle to an imagelist containing a bitmap that can be used to visualize
	/// dragging of the tabs of this collection.
	///
	/// \param[out] pXUpperLeft The x-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] pYUpperLeft The y-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] phImageList The handle to the imagelist.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa TabStrip::CreateLegacyDragImage
	virtual HRESULT STDMETHODCALLTYPE CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft = NULL, OLE_YPOS_PIXELS* pYUpperLeft = NULL, OLE_HANDLE* phImageList = NULL);
	/// \brief <em>Removes the specified tab from the collection</em>
	///
	/// \param[in] tabID The unique ID of the tabstrip tab to remove.
	/// \param[in] removePhysically If \c VARIANT_TRUE, the tab is removed from the control, too.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa TabStripTab::get_ID, Add, Count, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG tabID, VARIANT_BOOL removePhysically = VARIANT_FALSE);
	/// \brief <em>Removes all tabs from the collection</em>
	///
	/// \param[in] removePhysically If \c VARIANT_TRUE, the tabs get removed from the control, too.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(VARIANT_BOOL removePhysically = VARIANT_FALSE);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerTabStrip
	void SetOwner(__in_opt TabStrip* pOwner);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ITabContainer
	///
	//@{
	/// \brief <em>A tab was removed and the tab container should check its content</em>
	///
	/// \param[in] tabID The unique ID of the removed tab. If -1, all tabs were removed.
	///
	/// \sa TabStrip::RemoveTabFromItemContainers
	void RemovedTab(LONG tabID);
	/// \brief <em>Retrieves the container's ID used to identify it</em>
	///
	/// \return The container's ID.
	///
	/// \sa TabStrip::DeregisterTabContainer, containerID
	DWORD GetID(void);
	/// \brief <em>Holds the container's ID</em>
	DWORD containerID;
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c TabStrip object that owns this collection</em>
		///
		/// \sa SetOwner
		TabStrip* pOwnerTabStrip;
		#ifdef USE_STL
			/// \brief <em>Holds the tabs that this object contains</em>
			std::vector<LONG> tabs;
		#else
			/// \brief <em>Holds the tabs that this object contains</em>
			//CAtlArray<LONG> tabs;
		#endif
		/// \brief <em>Points to the next enumerated tab</em>
		int nextEnumeratedTab;

		Properties()
		{
			pOwnerTabStrip = NULL;
			nextEnumeratedTab = 0;
		}

		~Properties();

		/// \brief <em>Retrieves the owning tabstrip's window handle</em>
		///
		/// \return The window handle of the tabstrip that contains the tabs in this collection.
		///
		/// \sa pOwnerTabStrip
		HWND GetTabStripHWnd(void);
	} properties;

	/* TODO: If we move this one into the Properties struct, the compiler complains that the private
	         = operator of CAtlArray cannot be accessed?! */
	#ifndef USE_STL
		/// \brief <em>Holds the tabs that this object contains</em>
		CAtlArray<LONG> tabs;
	#endif

	/// \brief <em>Incremented by the constructor and used as the constructed object's ID</em>
	///
	/// \sa TabStripTabContainer, containerID
	static DWORD nextID;
};     // TabStripTabContainer

OBJECT_ENTRY_AUTO(__uuidof(TabStripTabContainer), TabStripTabContainer)