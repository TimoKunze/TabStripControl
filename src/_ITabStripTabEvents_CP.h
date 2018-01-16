//////////////////////////////////////////////////////////////////////
/// \class Proxy_ITabStripTabEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _ITabStripTabEvents interface</em>
///
/// \if UNICODE
///   \sa TabStripTab, TabStripCtlLibU::_ITabStripTabEvents
/// \else
///   \sa TabStripTab, TabStripCtlLibA::_ITabStripTabEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_ITabStripTabEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ITabStripTabEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_ITabStripTabEvents