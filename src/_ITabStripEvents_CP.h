//////////////////////////////////////////////////////////////////////
/// \class Proxy_ITabStripEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _ITabStripEvents interface</em>
///
/// \if UNICODE
///   \sa TabStrip, TabStripCtlLibU::_ITabStripEvents
/// \else
///   \sa TabStrip, TabStripCtlLibA::_ITabStripEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_ITabStripEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ITabStripEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c AbortedDrag event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::AbortedDrag, TabStrip::Raise_AbortedDrag, Fire_Drop
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::AbortedDrag, TabStrip::Raise_AbortedDrag, Fire_Drop
	/// \endif
	HRESULT Fire_AbortedDrag(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_ABORTEDDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ActiveTabChanged event</em>
	///
	/// \param[in] pPreviousActiveTab The previous active tab.
	/// \param[in] pNewActiveTab The new active tab.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::ActiveTabChanged, TabStrip::Raise_ActiveTabChanged,
	///       Fire_ActiveTabChanging, Fire_CaretChanged
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::ActiveTabChanged, TabStrip::Raise_ActiveTabChanged,
	///       Fire_ActiveTabChanging, Fire_CaretChanged
	/// \endif
	HRESULT Fire_ActiveTabChanged(ITabStripTab* pPreviousActiveTab, ITabStripTab* pNewActiveTab)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pPreviousActiveTab;
				p[0] = pNewActiveTab;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_ACTIVETABCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ActiveTabChanging event</em>
	///
	/// \param[in] pPreviousActiveTab The previous active tab.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should abort the active tab change;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::ActiveTabChanging, TabStrip::Raise_ActiveTabChanging,
	///       Fire_ActiveTabChanged, Fire_CaretChanged
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::ActiveTabChanging, TabStrip::Raise_ActiveTabChanging,
	///       Fire_ActiveTabChanged, Fire_CaretChanged
	/// \endif
	HRESULT Fire_ActiveTabChanging(ITabStripTab* pPreviousActiveTab, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pPreviousActiveTab;
				p[0].pboolVal = pCancelChange;			p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_ACTIVETABCHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CaretChanged event</em>
	///
	/// \param[in] pPreviousCaretTab The previous caret tab.
	/// \param[in] pNewCaretTab The new caret tab.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::CaretChanged, TabStrip::Raise_CaretChanged,
	///       Fire_ActiveTabChanged
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::CaretChanged, TabStrip::Raise_CaretChanged,
	///       Fire_ActiveTabChanged
	/// \endif
	HRESULT Fire_CaretChanged(ITabStripTab* pPreviousCaretTab, ITabStripTab* pNewCaretTab)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pPreviousCaretTab;
				p[0] = pNewCaretTab;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_CARETCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Click event</em>
	///
	/// \param[in] pTSTab The clicked tab. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::Click, TabStrip::Raise_Click, Fire_DblClick, Fire_MClick,
	///       Fire_RClick, Fire_XClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::Click, TabStrip::Raise_Click, Fire_DblClick, Fire_MClick,
	///       Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_Click(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_CLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ContextMenu event</em>
	///
	/// \param[in] pTSTab The tab that the context menu refers to. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the menu's proposed position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::ContextMenu, TabStrip::Raise_ContextMenu, Fire_RClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::ContextMenu, TabStrip::Raise_ContextMenu, Fire_RClick
	/// \endif
	HRESULT Fire_ContextMenu(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_CONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DblClick event</em>
	///
	/// \param[in] pTSTab The double-clicked tab. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::DblClick, TabStrip::Raise_DblClick, Fire_Click,
	///       Fire_MDblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::DblClick, TabStrip::Raise_DblClick, Fire_Click,
	///       Fire_MDblClick, Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_DblClick(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_DBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::DestroyedControlWindow,
	///       TabStrip::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::DestroyedControlWindow,
	///       TabStrip::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \endif
	HRESULT Fire_DestroyedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_DESTROYEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DragMouseMove event</em>
	///
	/// \param[in,out] ppDropTarget The tab that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another tab.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoActivateTab If set to \c VARIANT_TRUE, the caller should auto-activate the tab
	///                specified by \c ppDropTarget; otherwise it should cancel auto-activation. See the
	///                following <strong>remarks</strong> section for details.
	/// \param[in,out] pAutoScrollVelocity The speed multiplier for auto-scrolling. If set to 0, the caller
	///                should disable auto-scrolling; if set to a value less than 0, it should scroll the
	///                control to the left or top (depending on the setting of the \c TabPlacement property);
	///                if set to a value greater than 0, it should scroll the control to the right or bottom.
	///                The higher/lower the value is, the faster the caller should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Auto-activation is timered, i. e. the caller should start the timer after this event, if
	///          the \c ppDropTarget parameter specifies another tab than the last time this event was
	///          fired. If the timer isn't canceled, the caller should activate the tab when the timer
	///          expires.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::DragMouseMove, TabStrip::Raise_DragMouseMove,
	///       Fire_MouseMove, Fire_OLEDragMouseMove
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::DragMouseMove, TabStrip::Raise_DragMouseMove,
	///       Fire_MouseMove, Fire_OLEDragMouseMove
	/// \endif
	HRESULT Fire_DragMouseMove(ITabStripTab** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoActivateTab, LONG* pAutoScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[7].vt = VT_DISPATCH | VT_BYREF;
				p[6] = button;																									p[6].vt = VT_I2;
				p[5] = shift;																										p[5].vt = VT_I2;
				p[4] = x;																												p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																												p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].pboolVal = pAutoActivateTab;																p[1].vt = VT_BOOL | VT_BYREF;
				p[0].plVal = pAutoScrollVelocity;																p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_DRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Drop event</em>
	///
	/// \param[in] pDropTarget The tab object that is the nearest one from the mouse cursor's position.
	///            If the mouse cursor's position lies outside the control's client area, this parameter
	///            should be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::Drop, TabStrip::Raise_Drop, Fire_AbortedDrag
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::Drop, TabStrip::Raise_Drop, Fire_AbortedDrag
	/// \endif
	HRESULT Fire_Drop(ITabStripTab* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pDropTarget;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_DROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FreeTabData event</em>
	///
	/// \param[in] pTSTab The tab whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::FreeTabData, TabStrip::Raise_FreeTabData, Fire_RemovingTab,
	///       Fire_RemovedTab
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::FreeTabData, TabStrip::Raise_FreeTabData, Fire_RemovingTab,
	///       Fire_RemovedTab
	/// \endif
	HRESULT Fire_FreeTabData(ITabStripTab* pTSTab)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pTSTab;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_FREETABDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedTab event</em>
	///
	/// \param[in] pTSTab The inserted tab.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::InsertedTab, TabStrip::Raise_InsertedTab, Fire_InsertingTab,
	///       Fire_RemovedTab
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::InsertedTab, TabStrip::Raise_InsertedTab, Fire_InsertingTab,
	///       Fire_RemovedTab
	/// \endif
	HRESULT Fire_InsertedTab(ITabStripTab* pTSTab)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pTSTab;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_INSERTEDTAB, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingTab event</em>
	///
	/// \param[in] pTSTab The tab being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::InsertingTab, TabStrip::Raise_InsertingTab,
	///       Fire_InsertedTab, Fire_RemovingTab
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::InsertingTab, TabStrip::Raise_InsertingTab,
	///       Fire_InsertedTab, Fire_RemovingTab
	/// \endif
	HRESULT Fire_InsertingTab(IVirtualTabStripTab* pTSTab, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pTSTab;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_INSERTINGTAB, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::KeyDown, TabStrip::Raise_KeyDown, Fire_KeyUp, Fire_KeyPress
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::KeyDown, TabStrip::Raise_KeyDown, Fire_KeyUp, Fire_KeyPress
	/// \endif
	HRESULT Fire_KeyDown(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_KEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::KeyPress, TabStrip::Raise_KeyPress, Fire_KeyDown, Fire_KeyUp
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::KeyPress, TabStrip::Raise_KeyPress, Fire_KeyDown, Fire_KeyUp
	/// \endif
	HRESULT Fire_KeyPress(SHORT* pKeyAscii)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].piVal = pKeyAscii;		p[0].vt = VT_I2 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_KEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::KeyUp, TabStrip::Raise_KeyUp, Fire_KeyDown, Fire_KeyPress
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::KeyUp, TabStrip::Raise_KeyUp, Fire_KeyDown, Fire_KeyPress
	/// \endif
	HRESULT Fire_KeyUp(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_KEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MClick event</em>
	///
	/// \param[in] pTSTab The clicked tab. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::MClick, TabStrip::Raise_MClick, Fire_MDblClick, Fire_Click,
	///       Fire_RClick, Fire_XClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::MClick, TabStrip::Raise_MClick, Fire_MDblClick, Fire_Click,
	///       Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MClick(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_MCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MDblClick event</em>
	///
	/// \param[in] pTSTab The double-clicked tab. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::MDblClick, TabStrip::Raise_MDblClick, Fire_MClick,
	///       Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::MDblClick, TabStrip::Raise_MDblClick, Fire_MClick,
	///       Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_MDblClick(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_MDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseDown event</em>
	///
	/// \param[in] pTSTab The tab that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::MouseDown, TabStrip::Raise_MouseDown, Fire_MouseUp,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::MouseDown, TabStrip::Raise_MouseDown, Fire_MouseUp,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseDown(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_MOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseEnter event</em>
	///
	/// \param[in] pTSTab The tab that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::MouseEnter, TabStrip::Raise_MouseEnter, Fire_MouseLeave,
	///       Fire_TabMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::MouseEnter, TabStrip::Raise_MouseEnter, Fire_MouseLeave,
	///       Fire_TabMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseEnter(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_MOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseHover event</em>
	///
	/// \param[in] pTSTab The tab that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::MouseHover, TabStrip::Raise_MouseHover, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseMove
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::MouseHover, TabStrip::Raise_MouseHover, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseHover(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_MOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseLeave event</em>
	///
	/// \param[in] pTSTab The tab that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::MouseLeave, TabStrip::Raise_MouseLeave, Fire_MouseEnter,
	///       Fire_TabMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::MouseLeave, TabStrip::Raise_MouseLeave, Fire_MouseEnter,
	///       Fire_TabMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseLeave(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_MOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseMove event</em>
	///
	/// \param[in] pTSTab The tab that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::MouseMove, TabStrip::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::MouseMove, TabStrip::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp
	/// \endif
	HRESULT Fire_MouseMove(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_MOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseUp event</em>
	///
	/// \param[in] pTSTab The tab that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::MouseUp, TabStrip::Raise_MouseUp, Fire_MouseDown,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::MouseUp, TabStrip::Raise_MouseUp, Fire_MouseDown,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseUp(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_MOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLECompleteDrag, TabStrip::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLECompleteDrag, TabStrip::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \endif
	HRESULT Fire_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pData;
				p[0].lVal = static_cast<LONG>(performedEffect);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLECOMPLETEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client finally executed.
	/// \param[in] pDropTarget The tab object that is the nearest one from the mouse cursor's position.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLEDragDrop, TabStrip::Raise_OLEDragDrop, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLEDragDrop, TabStrip::Raise_OLEDragDrop, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp
	/// \endif
	HRESULT Fire_OLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, ITabStripTab* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pData;
				p[6].plVal = reinterpret_cast<PLONG>(pEffect);		p[6].vt = VT_I4 | VT_BYREF;
				p[5] = pDropTarget;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The tab that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another tab.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoActivateTab If set to \c VARIANT_TRUE, the caller should auto-activate the tab
	///                specified by \c ppDropTarget; otherwise it should cancel auto-activation. See the
	///                following <strong>remarks</strong> section for details.
	/// \param[in,out] pAutoScrollVelocity The speed multiplier for auto-scrolling. If set to 0, the caller
	///                should disable auto-scrolling; if set to a value less than 0, it should scroll the
	///                control to the left or top (depending on the setting of the \c TabPlacement property);
	///                if set to a value greater than 0, it should scroll the control to the right or bottom.
	///                The higher/lower the value is, the faster the caller should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Auto-activation is timered, i. e. the caller should start the timer after this event, if
	///          the \c ppDropTarget parameter specifies another tab than the last time this event was
	///          fired. If the timer isn't canceled, the caller should activate the tab when the timer
	///          expires.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLEDragEnter, TabStrip::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLEDragEnter, TabStrip::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter
	/// \endif
	HRESULT Fire_OLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, ITabStripTab** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoActivateTab, LONG* pAutoScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[10];
				p[9] = pData;
				p[8].plVal = reinterpret_cast<PLONG>(pEffect);									p[8].vt = VT_I4 | VT_BYREF;
				p[7].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[7].vt = VT_DISPATCH | VT_BYREF;
				p[6] = button;																									p[6].vt = VT_I2;
				p[5] = shift;																										p[5].vt = VT_I2;
				p[4] = x;																												p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																												p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].pboolVal = pAutoActivateTab;																p[1].vt = VT_BOOL | VT_BYREF;
				p[0].plVal = pAutoScrollVelocity;																p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 10, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The handle of the potential drag'n'drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLEDragEnterPotentialTarget,
	///       TabStrip::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLEDragEnterPotentialTarget,
	///       TabStrip::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \endif
	HRESULT Fire_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndPotentialTarget;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLEDRAGENTERPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeave event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in] pDropTarget The tab that is the current target of the drag'n'drop operation.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLEDragLeave, TabStrip::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLEDragLeave, TabStrip::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave
	/// \endif
	HRESULT Fire_OLEDragLeave(IOLEDataObject* pData, ITabStripTab* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pData;
				p[5] = pDropTarget;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLEDragLeavePotentialTarget,
	///       TabStrip::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLEDragLeavePotentialTarget,
	///       TabStrip::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \endif
	HRESULT Fire_OLEDragLeavePotentialTarget(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLEDRAGLEAVEPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragMouseMove event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The tab that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another tab.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoActivateTab If set to \c VARIANT_TRUE, the caller should auto-activate the tab
	///                specified by \c ppDropTarget; otherwise it should cancel auto-activation. See the
	///                following <strong>remarks</strong> section for details.
	/// \param[in,out] pAutoScrollVelocity The speed multiplier for auto-scrolling. If set to 0, the caller
	///                should disable auto-scrolling; if set to a value less than 0, it should scroll the
	///                control to the left or top (depending on the setting of the \c TabPlacement property);
	///                if set to a value greater than 0, it should scroll the control to the right or bottom.
	///                The higher/lower the value is, the faster the caller should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Auto-activation is timered, i. e. the caller should start the timer after this event, if
	///          the \c ppDropTarget parameter specifies another tab than the last time this event was
	///          fired. If the timer isn't canceled, the caller should activate the tab when the timer
	///          expires.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLEDragMouseMove, TabStrip::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLEDragMouseMove, TabStrip::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove
	/// \endif
	HRESULT Fire_OLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, ITabStripTab** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoActivateTab, LONG* pAutoScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[10];
				p[9] = pData;
				p[8].plVal = reinterpret_cast<PLONG>(pEffect);									p[8].vt = VT_I4 | VT_BYREF;
				p[7].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[7].vt = VT_DISPATCH | VT_BYREF;
				p[6] = button;																									p[6].vt = VT_I2;
				p[5] = shift;																										p[5].vt = VT_I2;
				p[4] = x;																												p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																												p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].pboolVal = pAutoActivateTab;																p[1].vt = VT_BOOL | VT_BYREF;
				p[0].plVal = pAutoScrollVelocity;																p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 10, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c OLEDropEffectConstants enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLEGiveFeedback, TabStrip::Raise_OLEGiveFeedback,
	///       Fire_OLEQueryContinueDrag
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLEGiveFeedback, TabStrip::Raise_OLEGiveFeedback,
	///       Fire_OLEQueryContinueDrag
	/// \endif
	HRESULT Fire_OLEGiveFeedback(OLEDropEffectConstants effect, VARIANT_BOOL* pUseDefaultCursors)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = static_cast<LONG>(effect);			p[1].vt = VT_I4;
				p[0].pboolVal = pUseDefaultCursors;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLEGIVEFEEDBACK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c VARIANT_TRUE, the user has pressed the \c ESC key since the last
	///            time this event was fired; otherwise not.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue, cancel or complete the
	///                drag'n'drop operation. Any of the values defined by the
	///                \c OLEActionToContinueWithConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLEQueryContinueDrag, TabStrip::Raise_OLEQueryContinueDrag,
	///       Fire_OLEGiveFeedback
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLEQueryContinueDrag, TabStrip::Raise_OLEQueryContinueDrag,
	///       Fire_OLEGiveFeedback
	/// \endif
	HRESULT Fire_OLEQueryContinueDrag(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, OLEActionToContinueWithConstants* pActionToContinueWith)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pressedEscape;
				p[2] = button;																									p[2].vt = VT_I2;
				p[1] = shift;																										p[1].vt = VT_I2;
				p[0].plVal = reinterpret_cast<PLONG>(pActionToContinueWith);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLEQUERYCONTINUEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEReceivedNewData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the data object has received data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLEReceivedNewData, TabStrip::Raise_OLEReceivedNewData,
	///       Fire_OLESetData
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLEReceivedNewData, TabStrip::Raise_OLEReceivedNewData,
	///       Fire_OLESetData
	/// \endif
	HRESULT Fire_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLERECEIVEDNEWDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLESetData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the drop target is requesting data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLESetData, TabStrip::Raise_OLESetData, Fire_OLEStartDrag
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLESetData, TabStrip::Raise_OLESetData, Fire_OLEStartDrag
	/// \endif
	HRESULT Fire_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLESETDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OLEStartDrag, TabStrip::Raise_OLEStartDrag, Fire_OLESetData,
	///       Fire_OLECompleteDrag
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OLEStartDrag, TabStrip::Raise_OLEStartDrag, Fire_OLESetData,
	///       Fire_OLECompleteDrag
	/// \endif
	HRESULT Fire_OLEStartDrag(IOLEDataObject* pData)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pData;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OLESTARTDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OwnerDrawTab event</em>
	///
	/// \param[in] pTSTab The tab to draw.
	/// \param[in] tabState The tab's current state (e. g. selected). Any of the values defined by the
	///            \c OwnerDrawTabStateConstants enumeration is valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::OwnerDrawTab, TabStrip::Raise_OwnerDrawTab
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::OwnerDrawTab, TabStrip::Raise_OwnerDrawTab
	/// \endif
	HRESULT Fire_OwnerDrawTab(ITabStripTab* pTSTab, OwnerDrawTabStateConstants tabState, LONG hDC, RECTANGLE* pDrawingRectangle)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pTSTab;
				p[2].lVal = static_cast<LONG>(tabState);		p[2].vt = VT_I4;
				p[1] = hDC;

				// pack the pDrawingRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{4AD0DE37-E9D7-4e36-A127-CFFDA357DC63}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TabStripCtlLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{90A884B3-C35D-44db-9710-5F694CBBCDA1}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TabStripCtlLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#endif
				VariantInit(&p[0]);
				p[0].vt = VT_RECORD | VT_BYREF;
				p[0].pRecInfo = pRecordInfo;
				p[0].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Bottom = pDrawingRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Left = pDrawingRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Right = pDrawingRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Top = pDrawingRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_OWNERDRAWTAB, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[0].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RClick event</em>
	///
	/// \param[in] pTSTab The clicked tab. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::RClick, TabStrip::Raise_RClick, Fire_ContextMenu,
	///       Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::RClick, TabStrip::Raise_RClick, Fire_ContextMenu,
	///       Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick
	/// \endif
	HRESULT Fire_RClick(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_RCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RDblClick event</em>
	///
	/// \param[in] pTSTab The double-clicked tab. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::RDblClick, TabStrip::Raise_RDblClick, Fire_RClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::RDblClick, TabStrip::Raise_RDblClick, Fire_RClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_RDblClick(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_RDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::RecreatedControlWindow,
	///       TabStrip::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::RecreatedControlWindow,
	///       TabStrip::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \endif
	HRESULT Fire_RecreatedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_RECREATEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovedTab event</em>
	///
	/// \param[in] pTSTab The removed tab.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::RemovedTab, TabStrip::Raise_RemovedTab, Fire_RemovingTab,
	///       Fire_InsertedTab
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::RemovedTab, TabStrip::Raise_RemovedTab, Fire_RemovingTab,
	///       Fire_InsertedTab
	/// \endif
	HRESULT Fire_RemovedTab(IVirtualTabStripTab* pTSTab)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pTSTab;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_REMOVEDTAB, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingTab event</em>
	///
	/// \param[in] pTSTab The tab being removed.
	/// \param[in,out] pCancelDeletion If \c VARIANT_TRUE, the caller should abort deletion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::RemovingTab, TabStrip::Raise_RemovingTab, Fire_RemovedTab,
	///       Fire_InsertingTab
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::RemovingTab, TabStrip::Raise_RemovingTab, Fire_RemovedTab,
	///       Fire_InsertingTab
	/// \endif
	HRESULT Fire_RemovingTab(ITabStripTab* pTSTab, VARIANT_BOOL* pCancelDeletion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pTSTab;
				p[0].pboolVal = pCancelDeletion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_REMOVINGTAB, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::ResizedControlWindow, TabStrip::Raise_ResizedControlWindow
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::ResizedControlWindow, TabStrip::Raise_ResizedControlWindow
	/// \endif
	HRESULT Fire_ResizedControlWindow(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_RESIZEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c TabBeginDrag event</em>
	///
	/// \param[in] pTSTab The tab that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbLeftButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::TabBeginDrag, TabStrip::Raise_TabBeginDrag,
	///       Fire_TabBeginRDrag
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::TabBeginDrag, TabStrip::Raise_TabBeginDrag,
	///       Fire_TabBeginRDrag
	/// \endif
	HRESULT Fire_TabBeginDrag(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_TABBEGINDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c TabBeginRDrag event</em>
	///
	/// \param[in] pTSTab The tab that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbRightButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::TabBeginRDrag, TabStrip::Raise_TabBeginRDrag,
	///       Fire_TabBeginDrag
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::TabBeginRDrag, TabStrip::Raise_TabBeginRDrag,
	///       Fire_TabBeginDrag
	/// \endif
	HRESULT Fire_TabBeginRDrag(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_TABBEGINRDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c TabGetInfoTipText event</em>
	///
	/// \param[in] pTSTab The tab whose info tip text is requested.
	/// \param[in] maxInfoTipLength The maximum number of characters the info tip text may consist of.
	/// \param[out] pInfoTipText The tab's info tip text.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::TabGetInfoTipText, TabStrip::Raise_TabGetInfoTipText
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::TabGetInfoTipText, TabStrip::Raise_TabGetInfoTipText
	/// \endif
	HRESULT Fire_TabGetInfoTipText(ITabStripTab* pTSTab, LONG maxInfoTipLength, BSTR* pInfoTipText)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pTSTab;
				p[1] = maxInfoTipLength;
				p[0].pbstrVal = pInfoTipText;		p[0].vt = VT_BSTR | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_TABGETINFOTIPTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c TabMouseEnter event</em>
	///
	/// \param[in] pTSTab The tab that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::TabMouseEnter, TabStrip::Raise_TabMouseEnter,
	///       Fire_TabMouseLeave, Fire_MouseMove
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::TabMouseEnter, TabStrip::Raise_TabMouseEnter,
	///       Fire_TabMouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_TabMouseEnter(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_TABMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c TabMouseLeave event</em>
	///
	/// \param[in] pTSTab The tab that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::TabMouseLeave, TabStrip::Raise_TabMouseLeave,
	///       Fire_TabMouseEnter, Fire_MouseMove
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::TabMouseLeave, TabStrip::Raise_TabMouseLeave,
	///       Fire_TabMouseEnter, Fire_MouseMove
	/// \endif
	HRESULT Fire_TabMouseLeave(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_TABMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c TabSelectionChanged event</em>
	///
	/// \param[in] pTSTab The item that was selected/deselected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::TabSelectionChanged, TabStrip::Raise_TabSelectionChanged,
	///       Fire_ActiveTabChanged, Fire_CaretChanged
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::TabSelectionChanged, TabStrip::Raise_TabSelectionChanged,
	///       Fire_ActiveTabChanged, Fire_CaretChanged
	/// \endif
	HRESULT Fire_TabSelectionChanged(ITabStripTab* pTSTab)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pTSTab;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_TABSELECTIONCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XClick event</em>
	///
	/// \param[in] pTSTab The clicked tab. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::XClick, TabStrip::Raise_XClick, Fire_XDblClick,
	///       Fire_Click, Fire_MClick, Fire_RClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::XClick, TabStrip::Raise_XClick, Fire_XDblClick,
	///       Fire_Click, Fire_MClick, Fire_RClick
	/// \endif
	HRESULT Fire_XClick(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_XCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XDblClick event</em>
	///
	/// \param[in] pTSTab The double-clicked tab. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TabStripCtlLibU::_ITabStripEvents::XDblClick, TabStrip::Raise_XDblClick, Fire_XClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \else
	///   \sa TabStripCtlLibA::_ITabStripEvents::XDblClick, TabStrip::Raise_XDblClick, Fire_XClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \endif
	HRESULT Fire_XDblClick(ITabStripTab* pTSTab, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTSTab;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TABSTRIPCTLE_XDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_ITabStripEvents