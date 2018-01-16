//////////////////////////////////////////////////////////////////////
/// \class ITabContainer
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between a \c TabStripTabContainer object and its creator object</em>
///
/// This interface allows an \c TabStrip object to inform a \c TabStripTabContainer object
/// about tab deletions.
///
/// \sa TabStrip::RegisterTabContainer, TabStripTabContainer
//////////////////////////////////////////////////////////////////////


#pragma once


class ITabContainer
{
public:
	/// \brief <em>A tab was removed and the tab container should check its content</em>
	///
	/// \param[in] tabID The unique ID of the removed tab. If 0, all tabs were removed.
	///
	/// \sa TabStrip::RemoveTabFromTabContainers
	virtual void RemovedTab(LONG tabID) = 0;
	/// \brief <em>Retrieves the container's ID used to identify it</em>
	///
	/// \return The container's ID.
	///
	/// \sa TabStrip::DeregisterTabContainer, containerID
	virtual DWORD GetID(void) = 0;

	/// \brief <em>Holds the container's ID</em>
	DWORD containerID;
};     // ITabContainer