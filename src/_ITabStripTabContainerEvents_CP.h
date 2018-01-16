//////////////////////////////////////////////////////////////////////
/// \class Proxy_ITabStripTabContainerEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _ITabStripTabContainerEvents interface</em>
///
/// \if UNICODE
///   \sa TabStripTabContainer, TabStripCtlLibU::_ITabStripTabContainerEvents
/// \else
///   \sa TabStripTabContainer, TabStripCtlLibA::_ITabStripTabContainerEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_ITabStripTabContainerEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ITabStripTabContainerEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_ITabStripTabContainerEvents