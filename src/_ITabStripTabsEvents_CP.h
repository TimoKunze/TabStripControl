//////////////////////////////////////////////////////////////////////
/// \class Proxy_ITabStripTabsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _ITabStripTabsEvents interface</em>
///
/// \if UNICODE
///   \sa TabStripTabs, TabStripCtlLibU::_ITabStripTabsEvents
/// \else
///   \sa TabStripTabs, TabStripCtlLibA::_ITabStripTabsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_ITabStripTabsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ITabStripTabsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_ITabStripTabsEvents