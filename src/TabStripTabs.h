//////////////////////////////////////////////////////////////////////
/// \class TabStripTabs
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c TabStripTab objects</em>
///
/// This class provides easy access (including filtering) to collections of \c TabStripTab objects.
/// \c TabStripTabs objects are used to group tabs that have certain properties in common.
///
/// \if UNICODE
///   \sa TabStripCtlLibU::ITabStripTabs, TabStripTab, TabStrip
/// \else
///   \sa TabStripCtlLibA::ITabStripTabs, TabStripTab, TabStrip
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TabStripCtlU.h"
#else
	#include "TabStripCtlA.h"
#endif
#include "_ITabStripTabsEvents_CP.h"
#include "helpers.h"
#include "TabStrip.h"
#include "TabStripTab.h"


class ATL_NO_VTABLE TabStripTabs : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TabStripTabs, &CLSID_TabStripTabs>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TabStripTabs>,
    public Proxy_ITabStripTabsEvents<TabStripTabs>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<ITabStripTabs, &IID_ITabStripTabs, &LIBID_TabStripCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<ITabStripTabs, &IID_ITabStripTabs, &LIBID_TabStripCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class TabStrip;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TABSTRIPTABS)

		BEGIN_COM_MAP(TabStripTabs)
			COM_INTERFACE_ENTRY(ITabStripTabs)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TabStripTabs)
			CONNECTION_POINT_ENTRY(__uuidof(_ITabStripTabsEvents))
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
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfTabsToSkip items.
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
	/// \name Implementation of ITabStripTabs
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c CaseSensitiveFilters property</em>
	///
	/// Retrieves whether string comparisons, that are done when applying the filters on a tab, are case
	/// sensitive. If this property is set to \c VARIANT_TRUE, string comparisons are case sensitive;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CaseSensitiveFilters, get_Filter, get_ComparisonFunction
	virtual HRESULT STDMETHODCALLTYPE get_CaseSensitiveFilters(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CaseSensitiveFilters property</em>
	///
	/// Sets whether string comparisons, that are done when applying the filters on a tab, are case
	/// sensitive. If this property is set to \c VARIANT_TRUE, string comparisons are case sensitive;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CaseSensitiveFilters, put_Filter, put_ComparisonFunction
	virtual HRESULT STDMETHODCALLTYPE put_CaseSensitiveFilters(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ComparisonFunction property</em>
	///
	/// Retrieves a tab filter's comparison function. This property takes the address of a function
	/// having the following signature:
	/// \code
	///   BOOL IsEqual(T tabProperty, T pattern);
	/// \endcode
	/// where T stands for the filtered property's type (\c VARIANT_BOOL, \c LONG or \c BSTR). This function
	/// must compare its arguments and return a non-zero value if the arguments are equal and zero
	/// otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters,
	///       TabStripCtlLibU::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters,
	///       TabStripCtlLibA::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue);
	/// \brief <em>Sets the \c ComparisonFunction property</em>
	///
	/// Sets a tab filter's comparison function. This property takes the address of a function
	/// having the following signature:\n
	/// \code
	///   BOOL IsEqual(T tabProperty, T pattern);
	/// \endcode
	/// where T stands for the filtered property's type (\c VARIANT_BOOL, \c LONG or \c BSTR). This function
	/// must compare its arguments and return a non-zero value if the arguments are equal and zero
	/// otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters,
	///       TabStripCtlLibU::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters,
	///       TabStripCtlLibA::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Filter property</em>
	///
	/// Retrieves a tab filter.\n
	/// An \c ITabStripTabs collection can be filtered by any of \c ITabStripTab's properties, that
	/// the \c FilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction, TabStripCtlLibU::FilteredPropertyConstants
	/// \else
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction, TabStripCtlLibA::FilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue);
	/// \brief <em>Sets the \c Filter property</em>
	///
	/// Sets a tab filter.\n
	/// An \c ITabStripTabs collection can be filtered by any of \c ITabStripTab's properties, that
	/// the \c FilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       TabStripCtlLibU::FilteredPropertyConstants
	/// \else
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       TabStripCtlLibA::FilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c FilterType property</em>
	///
	/// Retrieves a tab filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_FilterType, get_Filter, TabStripCtlLibU::FilteredPropertyConstants,
	///       TabStripCtlLibU::FilterTypeConstants
	/// \else
	///   \sa put_FilterType, get_Filter, TabStripCtlLibA::FilteredPropertyConstants,
	///       TabStripCtlLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue);
	/// \brief <em>Sets the \c FilterType property</em>
	///
	/// Sets a tab filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_FilterType, put_Filter, IsPartOfCollection, TabStripCtlLibU::FilteredPropertyConstants,
	///       TabStripCtlLibU::FilterTypeConstants
	/// \else
	///   \sa get_FilterType, put_Filter, IsPartOfCollection, TabStripCtlLibA::FilteredPropertyConstants,
	///       TabStripCtlLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue);
	/// \brief <em>Retrieves a \c TabStripTab object from the collection</em>
	///
	/// Retrieves a \c TabStripTab object from the collection that wraps the tabstrip tab identified
	/// by \c tabIdentifier.
	///
	/// \param[in] tabIdentifier A value that identifies the tabstrip tab to be retrieved.
	/// \param[in] tabIdentifierType A value specifying the meaning of \c tabIdentifier. Any of the
	///            values defined by the \c TabIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppTab Receives the tab's \c ITabStripTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa TabStripTab, Add, Remove, Contains, TabStripCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa TabStripTab, Add, Remove, Contains, TabStripCtlLibA::TabIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG tabIdentifier, TabIdentifierTypeConstants tabIdentifierType = titIndex, ITabStripTab** ppTab = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c TabStripTab objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator A pointer to the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds a tab to the tabstrip</em>
	///
	/// Adds a tab with the specified properties at the specified position in the control and returns a
	/// \c TabStripTab object wrapping the inserted tab.
	///
	/// \param[in] tabText The new tab's caption text. The maximum number of characters in this text is
	///            specified by \c MAX_TABTEXTLENGTH.
	/// \param[in] insertAt The new tab's zero-based index. If set to -1, the tab will be inserted as the
	///            last tab.
	/// \param[in] iconIndex The zero-based index of the tab's icon in the control's imagelist identified
	///            by \c TabStrip::get_hImageList. If set to -1, no icon will be displayed for the new tab.
	/// \param[in] tabData A \c LONG value that will be associated with the tab.
	/// \param[out] ppAddedTab Receives the added tab's \c ITabStripTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Count, Remove, RemoveAll, TabStripTab::get_Text, TabStripTab::get_IconIndex,
	///     TabStripTab::get_TabData, TabStrip::get_hImageList, MAX_TABTEXTLENGTH
	virtual HRESULT STDMETHODCALLTYPE Add(BSTR tabText, LONG insertAt = -1, LONG iconIndex = -1, LONG tabData = 0, ITabStripTab** ppAddedTab = NULL);
	/// \brief <em>Retrieves whether the specified tab is part of the tab collection</em>
	///
	/// \param[in] tabIdentifier A value that identifies the tab to be checked.
	/// \param[in] tabIdentifierType A value specifying the meaning of \c tabIdentifier. Any of the
	///            values defined by the \c TabIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the tab is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, Add, Remove, TabStripCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa get_Filter, Add, Remove, TabStripCtlLibA::TabIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(LONG tabIdentifier, TabIdentifierTypeConstants tabIdentifierType = titIndex, VARIANT_BOOL* pValue = NULL);
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
	/// \brief <em>Removes the selection from all tabs in the collection</em>
	///
	/// Sets the \c Selected property of all \c TabStripTab objects in the collection to \c VARIANT_FALSE.
	///
	/// \param[in] keepActiveTabSelected If \c VARIANT_TRUE, the control's active tab remains selected;
	///            otherwise it is deselected as well.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa TabStripTab::get_Selected, TabStrip::get_MultiSelect, TabStrip::get_ActiveTab
	virtual HRESULT STDMETHODCALLTYPE DeselectAll(VARIANT_BOOL keepActiveTabSelected = VARIANT_FALSE);
	/// \brief <em>Removes the specified tab in the collection from the tabstrip</em>
	///
	/// \param[in] tabIdentifier A value that identifies the tabstrip tab to be removed.
	/// \param[in] tabIdentifierType A value specifying the meaning of \c tabIdentifier. Any of the
	///            values defined by the \c TabIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, TabStripCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, TabStripCtlLibA::TabIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG tabIdentifier, TabIdentifierTypeConstants tabIdentifierType = titIndex);
	/// \brief <em>Removes all tabs in the collection from the tabstrip</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
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
	/// \name Filter validation
	///
	//@{
	/// \brief <em>Validates a filter for a \c VARIANT_BOOL type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type \c VARIANT_BOOL.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidIntegerFilter, IsValidStringFilter, put_Filter
	BOOL IsValidBooleanFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type
	/// \c LONG or compatible.
	///
	/// \param[in] filter The \c VARIANT to check.
	/// \param[in] minValue The minimum value that the corresponding property would accept.
	/// \param[in] maxValue The maximum value that the corresponding property would accept.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidStringFilter, put_Filter
	BOOL IsValidIntegerFilter(VARIANT& filter, int minValue, int maxValue);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// \overload
	BOOL IsValidIntegerFilter(VARIANT& filter, int minValue);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// \overload
	BOOL IsValidIntegerFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c BSTR type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type \c BSTR.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidIntegerFilter, put_Filter
	BOOL IsValidStringFilter(VARIANT& filter);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Filter appliance
	///
	//@{
	/// \brief <em>Retrieves the control's first tab that might be in the collection</em>
	///
	/// \param[in] numberOfTabs The number of tabstrip tabs in the control.
	///
	/// \return The tab being the collection's first candidate. -1 if no tab was found.
	///
	/// \sa GetNextTabToProcess, Count, RemoveAll, Next
	int GetFirstTabToProcess(int numberOfTabs);
	/// \brief <em>Retrieves the control's next tab that might be in the collection</em>
	///
	/// \param[in] previousTab The tab at which to start the search for a new collection candidate.
	/// \param[in] numberOfTabs The number of tabs in the control.
	///
	/// \return The next tab being a candidate for the collection. -1 if no tab was found.
	///
	/// \sa GetFirstTabToProcess, Count, RemoveAll, Next
	int GetNextTabToProcess(int previousTab, int numberOfTabs);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified boolean value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsIntegerInSafeArray, IsStringInSafeArray, get_ComparisonFunction
	BOOL IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified integer value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsBooleanInSafeArray, IsStringInSafeArray, get_ComparisonFunction
	BOOL IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified \c BSTR value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsBooleanInSafeArray, IsIntegerInSafeArray, get_ComparisonFunction
	BOOL IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether a tab is part of the collection (applying the filters)</em>
	///
	/// \param[in] tabIndex The tab to check.
	/// \param[in] hWndTabStrip The tabstrip window the method will work on.
	///
	/// \return \c TRUE, if the tab is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(int tabIndex, HWND hWndTabStrip = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Shortens a filter as much as possible</em>
	///
	/// Optimizes a filter by detecting redundancies, tautologies and so on.
	///
	/// \param[in] filteredProperty The filter to optimize. Any of the values defined by the
	///            \c FilteredPropertyConstants enumeration is valid.
	///
	/// \sa put_Filter, put_FilterType
	void OptimizeFilter(FilteredPropertyConstants filteredProperty);
	#ifdef USE_STL
		/// \brief <em>Removes the specified tabs</em>
		///
		/// \param[in] tabsToRemove A vector containing all tabs to remove.
		/// \param[in] hWndTabStrip The tabstrip window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveTabs(std::list<int>& tabsToRemove, HWND hWndTabStrip);
	#else
		/// \brief <em>Removes the specified tabs</em>
		///
		/// \param[in] tabsToRemove A vector containing all tabs to remove.
		/// \param[in] hWndTabStrip The tabstrip window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveTabs(CAtlList<int>& tabsToRemove, HWND hWndTabStrip);
	#endif

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		#define NUMBEROFFILTERS 12
		/// \brief <em>Holds the \c CaseSensitiveFilters property's setting</em>
		///
		/// \sa get_CaseSensitiveFilters, put_CaseSensitiveFilters
		UINT caseSensitiveFilters : 1;
		/// \brief <em>Holds the \c ComparisonFunction property's setting</em>
		///
		/// \sa get_ComparisonFunction, put_ComparisonFunction
		LPVOID comparisonFunction[NUMBEROFFILTERS];
		/// \brief <em>Holds the \c Filter property's setting</em>
		///
		/// \sa get_Filter, put_Filter
		VARIANT filter[NUMBEROFFILTERS];
		/// \brief <em>Holds the \c FilterType property's setting</em>
		///
		/// \sa get_FilterType, put_FilterType
		FilterTypeConstants filterType[NUMBEROFFILTERS];

		/// \brief <em>The \c TabStrip object that owns this collection</em>
		///
		/// \sa SetOwner
		TabStrip* pOwnerTabStrip;
		/// \brief <em>Holds the last enumerated tab</em>
		int lastEnumeratedTab;
		/// \brief <em>If \c TRUE, we must filter the items</em>
		///
		/// \sa put_Filter, put_FilterType
		UINT usingFilters : 1;

		Properties()
		{
			caseSensitiveFilters = FALSE;
			pOwnerTabStrip = NULL;
			lastEnumeratedTab = -1;

			for(int i = 0; i < NUMBEROFFILTERS; ++i) {
				VariantInit(&filter[i]);
				filterType[i] = ftDeactivated;
				comparisonFunction[i] = NULL;
			}
			usingFilters = FALSE;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);

		/// \brief <em>Retrieves the owning tabstrip's window handle</em>
		///
		/// \return The window handle of the tabstrip that contains the tabs in this collection.
		///
		/// \sa pOwnerTabStrip
		HWND GetTabStripHWnd(void);
	} properties;
};     // TabStripTabs

OBJECT_ENTRY_AUTO(__uuidof(TabStripTabs), TabStripTabs)