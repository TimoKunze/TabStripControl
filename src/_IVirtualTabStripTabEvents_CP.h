//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualTabStripTabEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualTabStripTabEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualTabStripTab, TabStripCtlLibU::_IVirtualTabStripTabEvents
/// \else
///   \sa VirtualTabStripTab, TabStripCtlLibA::_IVirtualTabStripTabEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualTabStripTabEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualTabStripTabEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualTabStripTabEvents