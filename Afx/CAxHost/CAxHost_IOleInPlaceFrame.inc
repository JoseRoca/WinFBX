' ########################################################################################
' Microsoft Windows
' File: CAxHost_IOleIOleInPlaceFrame.inc
' Contents: OLE Container class
' Implementation of the IOleInPlaceFrame interface.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' this CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
'                            *** CAxHost_IOleWindow class ***
' ########################################################################################

' ========================================================================================
' IOleWindow interface implemented
' The IOleWindow interface provides methods that allow an application to obtain the handle
' to the various windows that participate in in-place activation, and also to enter and
' exit context-sensitive help mode.
' ========================================================================================
TYPE CAxHost_IOleWindow EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pAxHost AS CAxHost PTR)
   DECLARE DESTRUCTOR

Private:
   cRef AS ULONG = 1
   m_pAxHost AS CAxHost PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CAxHost_IOleWindow
   CAXH_DP("CAxHost_IOleWindow Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_IOleWindow (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_IOleWindow constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_IOleWindow
   CAXH_DP("### CAxHost_IOleWindow destructor")
   m_pAxHost->m_pIOleWindowImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_IOleWindow.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IOleWindow.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_IOleWindow AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAxHost_IOleWindow.Release () AS DWORD
   this.cRef -= 1
   CAXH_DP("CAxHost_IOleWindow Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle to one of the windows participating in in-place activation
' (frame, document, parent, or in-place object window).
' ========================================================================================
FUNCTION CAxHost_IOleWindow.GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   CAXH_DP("CAxHost_IOleWindow GetWindow")
   IF phwnd = NULL THEN RETURN E_INVALIDARG
   *phwnd = m_pAxHost->m_hwnd
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether context-sensitive help mode should be entered during an in-place
' activation session.
' ========================================================================================
FUNCTION CAxHost_IOleWindow.ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
   CAXH_DP("CAxHost_IOleWindow ContextSensitiveHelp")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ########################################################################################
'                       *** CAxHost_IOleInPlaceUIWindow class ***
' ########################################################################################

' ========================================================================================
' IOleInPlaceUIWindow interface implemented
' Implemented by container applications and used by object applications to negotiate border
' space on the document or frame window. The container provides a RECT structure in which
' the object can place toolbars and other similar controls, determines if tools can in fact
' be installed around the object's window frame, allocates space for the border, and
' establishes a communication channel between the object and each frame and document window.
' ========================================================================================
TYPE CAxHost_IOleInPlaceUIWindow EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetBorder (BYVAL lprectBorder AS LPRECT) AS HRESULT
   DECLARE VIRTUAL FUNCTION RequestBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetActiveObject (BYVAL pActiveObject AS LPOLEINPLACEACTIVEOBJECT, BYVAL pszObjName AS LPCOLESTR) AS HRESULT

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pAxHost AS CAxHost PTR)
   DECLARE DESTRUCTOR

Private:
   cRef AS ULONG = 1
   m_pAxHost AS CAxHost PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CAxHost_IOleInPlaceUIWindow
   CAXH_DP("CAxHost_IOleInPlaceUIWindow Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_IOleInPlaceUIWindow (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_IOleInPlaceUIWindow constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_IOleInPlaceUIWindow
   CAXH_DP("### CAxHost_IOleInPlaceUIWindow destructor")
   m_pAxHost->m_pIOleInPlaceUIWindowImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceUIWindow.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceUIWindow.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_IOleInPlaceUIWindow AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceUIWindow.Release () AS DWORD
   this.cRef -= 1
   CAXH_DP("CAxHost_IOleInPlaceUIWindow Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle to one of the windows participating in in-place activation
' (frame, document, parent, or in-place object window).
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceUIWindow.GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceUIWindow GetWindow")
   IF phwnd = NULL THEN RETURN E_INVALIDARG
   *phwnd = m_pAxHost->m_hwnd
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether context-sensitive help mode should be entered during an in-place
' activation session.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceUIWindow.ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceUIWindow ContextSensitiveHelp")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies a RECT structure for toolbars and controls.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceUIWindow.GetBorder (BYVAL lprectBorder AS LPRECT) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceUIWindow GetBorder")
   IF lprectBorder = NULL THEN RETURN E_INVALIDARG ELSE RETURN INPLACE_E_NOTOOLSPACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if tools can be installed around object's window frame.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceUIWindow.RequestBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceUIWindow RequestBorderSpace")
   IF pborderwidths = NULL THEN RETURN E_INVALIDARG ELSE RETURN INPLACE_E_NOTOOLSPACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allocates space for the border.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceUIWindow.SetBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceUIWindow SetBorderSpace")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides for direct communication between the object and each document and frame window.
' SetActiveObject will be called with both NULL and non-NULL pointers.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceUIWindow.SetActiveObject (BYVAL pActiveObject AS LPOLEINPLACEACTIVEOBJECT, BYVAL pszObjName AS LPCOLESTR) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceUIWindow SetActiveObject")
   IF m_pAxHost->m_pIOleInPlaceActiveObject <> NULL THEN
      AfxSafeRelease(m_pAxHost->m_pIOleInPlaceActiveObject)
      m_pAxHost->m_pIOleInPlaceActiveObject = NULL
   END IF
   IF pActiveObject THEN
      AfxSafeAddRef(pActiveObject)
      m_pAxHost->m_pIOleInPlaceActiveObject = CAST(ANY PTR, pActiveObject)
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================


' ########################################################################################
'                         *** CAxHost_IOleInPlaceFrame class ***
' ########################################################################################

' ========================================================================================
' IOleInPlaceFrame interface implemented
' Controls the container's top-level frame window. This control involves allowing the
' container to insert its menu group into the composite menu, install the composite menu
' into the appropriate window frame, and remove the container's menu elements from the
' composite menu. It sets and displays status text relevant to the in-place object. It
' also enables or disables the frame's modeless dialog boxes, and translates accelerator
' keystrokes intended for the container's frame.
' ========================================================================================
TYPE CAxHost_IOleInPlaceFrame EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetBorder (BYVAL lprectBorder AS LPRECT) AS HRESULT
   DECLARE VIRTUAL FUNCTION RequestBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetActiveObject (BYVAL pActiveObject AS LPOLEINPLACEACTIVEOBJECT, BYVAL pszObjName AS LPCOLESTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION InsertMenus (BYVAL hmenuShared AS HMENU, BYVAL lpMenuWidths AS LPOLEMENUGROUPWIDTHS) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetMenu (BYVAL hmenuShared AS HMENU, BYVAL holemenu AS HOLEMENU, BYVAL hwndActiveObject AS HWND) AS HRESULT
   DECLARE VIRTUAL FUNCTION RemoveMenus (BYVAL hmenuShared AS HMENU) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetStatusText (BYVAL pszStatusText AS LPCOLESTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION EnableModeless (BYVAL fEnable AS BOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION TranslateAccelerator (BYVAL lpmsg AS LPMSG, BYVAL wID AS WORD) AS HRESULT

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pAxHost AS CAxHost PTR)
   DECLARE DESTRUCTOR

Private:
   cRef AS ULONG = 1
   m_pAxHost AS CAxHost PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CAxHost_IOleInPlaceFrame
   CAXH_DP("CAxHost_IOleInPlaceFrame Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_IOleInPlaceFrame (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_IOleInPlaceFrame constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_IOleInPlaceFrame
   CAXH_DP("### CAxHost_IOleInPlaceFrame destructor")
   m_pAxHost->m_pIOleInPlaceFrameImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_IOleInPlaceFrame AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.Release () AS DWORD
   this.cRef -= 1
   CAXH_DP("CAxHost_IOleInPlaceFrame Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle to one of the windows participating in in-place activation
' (frame, document, parent, or in-place object window).
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame GetWindow")
   IF phwnd = NULL THEN RETURN E_INVALIDARG
   *phwnd = m_pAxHost->m_hwnd
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether context-sensitive help mode should be entered during an in-place
' activation session.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame ContextSensitiveHelp")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies a RECT structure for toolbars and controls.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.GetBorder (BYVAL lprectBorder AS LPRECT) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame GetBorder")
   IF lprectBorder = NULL THEN RETURN E_INVALIDARG ELSE RETURN INPLACE_E_NOTOOLSPACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if tools can be installed around object's window frame.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.RequestBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame RequestBorderSpace")
   IF pborderwidths = NULL THEN RETURN E_INVALIDARG ELSE RETURN INPLACE_E_NOTOOLSPACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allocates space for the border.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.SetBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame SetBorderSpace")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides for direct communication between the object and each document and frame window.
' SetActiveObject will be called with both NULL and non-NULL pointers.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.SetActiveObject (BYVAL pActiveObject AS LPOLEINPLACEACTIVEOBJECT, BYVAL pszObjName AS LPCOLESTR) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame SetActiveObject")
   IF m_pAxHost->m_pIOleInPlaceActiveObject <> NULL THEN
      AfxSafeRelease(m_pAxHost->m_pIOleInPlaceActiveObject)
      m_pAxHost->m_pIOleInPlaceActiveObject = NULL
   END IF
   IF pActiveObject THEN
      AfxSafeAddRef(pActiveObject)
      m_pAxHost->m_pIOleInPlaceActiveObject = CAST(ANY PTR, pActiveObject)
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allows container to insert menus.
' Not required. This method is necessary only for containers with menu UI.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.InsertMenus (BYVAL hmenuShared AS HMENU, BYVAL lpMenuWidths AS LPOLEMENUGROUPWIDTHS) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame InsertMenus")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a composite menu to window frame.
' Not required. This method is necessary only for containers with menu UI.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.SetMenu (BYVAL hmenuShared AS HMENU, BYVAL holemenu AS HOLEMENU, BYVAL hwndActiveObject AS HWND) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame SetMenu")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes a container's menu elements.
' Not required. This method is necessary only for containers with menu UI.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.RemoveMenus (BYVAL hmenuShared AS HMENU) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame RemoveMenus")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets and displays status text about.
' Not required. This method is necessary only for containers that have a status line.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.SetStatusText (BYVAL pszStatusText AS LPCOLESTR) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame SetStatusText")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Optional. Enables or disables modeless dialog boxes.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.EnableModeless (BYVAL fEnable AS BOOL) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame EnableModeless")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Optional. Translates keystrokes.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceFrame.TranslateAccelerator (BYVAL lpmsg AS LPMSG, BYVAL wID AS WORD) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceFrame TranslateAccelerator")
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ########################################################################################
'                         *** CAxHost_ISimpleFrameSite class ***
' ########################################################################################

' ========================================================================================
' ISimpleFramaSite interface implemented
' Provides simple frame controls that act as simple containers for other nested controls.
' Some controls merely contain other controls. In such cases, the simple control container,
' called a simple frame, need not implement all container requirements. It can delegate
' most of the interface calls from its contained controls to the outer container that
' manages the simple frame. To support what are called simple frame controls, a container
' implements this interface as well as other site interfaces such as IOleControlSite.
' An example of a simple frame control is a group box that only needs to capture a few
' keystrokes for its contained controls to implement the correct tab or arrow key behavior,
' but does not want to handle every other message. Through the two methods of this interface,
' the simple frame control passes messages to its control site both before and after its
' own processing. If that site is itself a simple frame, it can continue to pass messages
' up the chain.
' Implement this interface on a control object to support nested controls without requiring
' the control to itself be a full container.
' ========================================================================================
TYPE CAxHost_ISimpleFrameSite EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION PreMessageFilter (BYVAL hWnd AS HWND, BYVAL msg AS UINT, BYVAL wp AS WPARAM, BYVAL lp AS LPARAM, BYVAL plResult AS LRESULT PTR, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION PostMessageFilter (BYVAL hWnd AS HWND, BYVAL msg AS UINT, BYVAL wp AS WPARAM, BYVAL lp AS LPARAM, BYVAL plResult AS LRESULT PTR, BYVAL dwCookie AS DWORD) AS HRESULT

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pAxHost AS CAxHost PTR)
   DECLARE DESTRUCTOR

Private:
   cRef AS ULONG = 1
   m_pAxHost AS CAxHost PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CAxHost_ISimpleFrameSite
   CAXH_DP("CAxHost_ISimpleFrameSite Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_ISimpleFrameSite (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_ISimpleFrameSite constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_ISimpleFrameSite
   CAXH_DP("### CAxHost_ISimpleFrameSite destructor")
   m_pAxHost->m_pISimpleFrameSiteImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_ISimpleFrameSite.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_ISimpleFrameSite.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_ISimpleFrameSite AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAxHost_ISimpleFrameSite.Release () AS DWORD
   this.cRef -= 1
   CAXH_DP("CAxHost_ISimpleFrameSite Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides a site with the opportunity to process a message that is received by a
' control's own window before the control itself does any processing.
' ========================================================================================
FUNCTION CAxHost_ISimpleFrameSite.PreMessageFilter (BYVAL hWnd AS HWND, BYVAL msg AS UINT, BYVAL wp AS WPARAM, BYVAL lp AS LPARAM, BYVAL plResult AS LRESULT PTR, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   CAXH_DP("CAxHost_ISimpleFrameSite PreMessageFilter")
   IF plResult = NULL OR pdwCookie = NULL THEN RETURN E_POINTER
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sends the simple frame site a message that is received by a control's own window after
' both ISimpleFrameSite.PreMessageFilter and the control have had a chance to process the
' message.
' ========================================================================================
FUNCTION CAxHost_ISimpleFrameSite.PostMessageFilter (ByVal hWnd as HWND, byval msg as UINT, byval wp as WPARAM, byval lp as LPARAM, byval plResult as LRESULT ptr, byval dwCookie as DWORD) as HRESULT
   CAXH_DP("CAxHost_ISimpleFrameSite PostMessageFilter")
   IF plResult = NULL THEN RETURN E_POINTER
   RETURN S_FALSE
END FUNCTION
' ========================================================================================
