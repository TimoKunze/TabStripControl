//////////////////////////////////////////////////////////////////////
/// \class ClassFactory
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A helper class for creating special objects</em>
///
/// This class provides methods to create objects and initialize them with given values.
///
/// \todo Improve documentation.
///
/// \sa TabStrip
//////////////////////////////////////////////////////////////////////


#pragma once

#include "TabStripTab.h"
#include "TabStripTabs.h"
#include "TargetOLEDataObject.h"


class ClassFactory
{
public:
	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's \c IOLEDataObject implementation as a smart pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	///
	/// \return A smart pointer to the created object's \c IOLEDataObject implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static CComPtr<IOLEDataObject> InitOLEDataObject(IDataObject* pDataObject)
	{
		CComPtr<IOLEDataObject> pOLEDataObj = NULL;
		InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<LPUNKNOWN*>(&pOLEDataObj));
		return pOLEDataObj;
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// \overload
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's implementation of a given interface as a raw pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppDataObject Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static void InitOLEDataObject(IDataObject* pDataObject, REFIID requiredInterface, LPUNKNOWN* ppDataObject)
	{
		ATLASSERT_POINTER(ppDataObject, LPUNKNOWN);
		ATLASSUME(ppDataObject);

		*ppDataObject = NULL;

		// create an OLEDataObject object and initialize it with the given parameters
		CComObject<TargetOLEDataObject>* pOLEDataObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObj);
		pOLEDataObj->AddRef();
		pOLEDataObj->Attach(pDataObject);
		pOLEDataObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppDataObject));
		pOLEDataObj->Release();
	};

	/// \brief <em>Creates a \c TabStripTab object</em>
	///
	/// Creates a \c TabStripTab object that represents a given tabstrip tab and returns its
	/// \c ITabStripTab implementation as a smart pointer.
	///
	/// \param[in] tabIndex The index of the tab for which to create the object.
	/// \param[in] pTabStrip The \c TabStrip object the \c TabStripTab object will work on.
	/// \param[in] validateTabIndex If \c TRUE, the method will check whether the \c tabIndex parameter
	///            identifies an existing tab; otherwise not.
	///
	/// \return A smart pointer to the created object's \c ITabStripTab implementation.
	static inline CComPtr<ITabStripTab> InitTabStripTab(int tabIndex, TabStrip* pTabStrip, BOOL validateTabIndex = TRUE)
	{
		CComPtr<ITabStripTab> pTab = NULL;
		InitTabStripTab(tabIndex, pTabStrip, IID_ITabStripTab, reinterpret_cast<LPUNKNOWN*>(&pTab), validateTabIndex);
		return pTab;
	};

	/// \brief <em>Creates a \c TabStripTab object</em>
	///
	/// \overload
	///
	/// Creates a \c TabStripTab object that represents a given tabstrip tab and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] tabIndex The index of the tab for which to create the object.
	/// \param[in] pTabStrip The \c TabStrip object the \c TabStripTab object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTab Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateTabIndex If \c TRUE, the method will check whether the \c tabIndex parameter
	///            identifies an existing tab; otherwise not.
	static inline void InitTabStripTab(int tabIndex, TabStrip* pTabStrip, REFIID requiredInterface, LPUNKNOWN* ppTab, BOOL validateTabIndex = TRUE)
	{
		ATLASSERT_POINTER(ppTab, LPUNKNOWN);
		ATLASSUME(ppTab);

		*ppTab = NULL;
		if(validateTabIndex && !IsValidTabIndex(tabIndex, pTabStrip)) {
			// there's nothing useful we could return
			return;
		}

		// create a TabStripTab object and initialize it with the given parameters
		CComObject<TabStripTab>* pTabStripTabObj = NULL;
		CComObject<TabStripTab>::CreateInstance(&pTabStripTabObj);
		pTabStripTabObj->AddRef();
		pTabStripTabObj->SetOwner(pTabStrip);
		pTabStripTabObj->Attach(tabIndex);
		pTabStripTabObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTab));
		pTabStripTabObj->Release();
	};

	/// \brief <em>Creates a \c TabStripTabs object</em>
	///
	/// Creates a \c TabStripTabs object that represents a collection of tabstrip tabs and returns its
	/// \c ITabStripTabs implementation as a smart pointer.
	///
	/// \param[in] pTabStrip The \c TabStrip object the \c TabStripTabs object will work on.
	///
	/// \return A smart pointer to the created object's \c ITabStripTabs implementation.
	static inline CComPtr<ITabStripTabs> InitTabStripTabs(TabStrip* pTabStrip)
	{
		CComPtr<ITabStripTabs> pTabs = NULL;
		InitTabStripTabs(pTabStrip, IID_ITabStripTabs, reinterpret_cast<LPUNKNOWN*>(&pTabs));
		return pTabs;
	};

	/// \brief <em>Creates a \c TabStripTabs object</em>
	///
	/// \overload
	///
	/// Creates a \c TabStripTabs object that represents a collection of tabstrip tabs and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] pTabStrip The \c TabStrip object the \c TabStripTabs object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTabs Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitTabStripTabs(TabStrip* pTabStrip, REFIID requiredInterface, LPUNKNOWN* ppTabs)
	{
		ATLASSERT_POINTER(ppTabs, LPUNKNOWN);
		ATLASSUME(ppTabs);

		*ppTabs = NULL;
		// create a TabStripTabs object and initialize it with the given parameters
		CComObject<TabStripTabs>* pTabStripTabsObj = NULL;
		CComObject<TabStripTabs>::CreateInstance(&pTabStripTabsObj);
		pTabStripTabsObj->AddRef();

		pTabStripTabsObj->SetOwner(pTabStrip);

		pTabStripTabsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTabs));
		pTabStripTabsObj->Release();
	};
};     // ClassFactory