//////////////////////////////////////////////////////////////////////
/// \class Proxy_IOLEDataObjectEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IOLEDataObjectEvents interface</em>
///
/// \if UNICODE
///   \sa SourceOLEDataObject, TargetOLEDataObject, TabStripCtlLibU::_IOLEDataObjectEvents
/// \else
///   \sa SourceOLEDataObject, TargetOLEDataObject, TabStripCtlLibA::_IOLEDataObjectEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IOLEDataObjectEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IOLEDataObjectEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IOLEDataObjectEvents