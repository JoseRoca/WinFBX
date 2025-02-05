' ########################################################################################
' Microsoft Windows
' File: CRichEditOleCAllback.inc
' Contents: Rich Edit control
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' ========================================================================================
' The IRichEditOleCallback interface is used by a rich text edit control to retrieve OLE-related
' information from its client. A rich edit control client is responsible for implementing
' this interface and assigning it to the control by using the EM_SETOLECALLBACK message.
' ========================================================================================

#pragma once
#include once "win/richole.bi"

' ========================================================================================
' CRichEditOleCallback class
' Implements the IRichEditOleCallback interface
' ========================================================================================
TYPE CRichEditOleCallback EXTENDS Object

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION GetNewStorage (BYVAL lplpstg AS LPSTORAGE PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetInPlaceContext (BYVAL lplpFrame AS LPOLEINPLACEFRAME PTR, BYVAL lplpDoc AS LPOLEINPLACEUIWINDOW PTR, BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT
   DECLARE VIRTUAL FUNCTION ShowContainerUI (BYVAL fShow AS WINBOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION QueryInsertObject (BYVAL lpclsid AS LPCLSID, BYVAL lpstg AS LPSTORAGE, BYVAL cp AS LONG) AS HRESULT
   DECLARE VIRTUAL FUNCTION DeleteObject (BYVAL lpoleobj AS LPOLEOBJECT) AS HRESULT
   DECLARE VIRTUAL FUNCTION QueryAcceptData (BYVAL lpdataobj AS LPDATAOBJECT, BYVAL lpcfFormat AS CLIPFORMAT PTR, BYVAL reco AS DWORD, BYVAL fReally AS WINBOOL, BYVAL hMetaPict AS HGLOBAL) AS HRESULT
   DECLARE VIRTUAL FUNCTION ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetClipboardData (BYVAL lpchrg AS CHARRANGE PTR, BYVAL reco AS DWORD, BYVAL lplpdataobj AS LPDATAOBJECT PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetDragDropEffect (BYVAL fDrag AS WINBOOL, BYVAL grfKeyState AS DWORD, BYVAL pdwEffect AS LPDWORD) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetContextMenu (BYVAL seltype AS WORD, BYVAL lpoleobj AS LPOLEOBJECT, BYVAL lpchrg AS CHARRANGE PTR, BYVAL lphmenu AS HMENU PTR) AS HRESULT

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Private:
   cRef AS ULONG = 0

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CRichEditOleCallback
   CREC_DP("CRichEditOleCallback Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CRichEditOleCallback
   CREC_DP("### CRichEditOleCallback destructor")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CRichEditOleCallback.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CREC_DP("CRichEditOleCallback.QueryInterface")
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   DIM IID_IRichEditOleCallback_ AS GUID = (&h00020D03, &h0000, &h0000, {&hC0, &h00, &h00, &h00, &h00, &h00, &h00, &h46})
   IF IsEqualIID(riid, @IID_IRichEditOleCallback_) OR IsEqualIID(riid, @IID_IUnknown) OR IsEqualIID(riid, @IID_IDispatch) THEN
      *ppvObj = @this
      RETURN S_OK
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' ========================================================================================

' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CRichEditOleCallback.AddRef () AS ULONG
'   this.cRef += 1
'   CREC_DP("CRichEditOleCallback AddRef = " & WSTR(this.cRef))
'   RETURN this.cRef
   RETURN 2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CRichEditOleCallback.Release () AS DWORD
'   this.cRef -= 1
'   CREC_DP("CRichEditOleCallback Release = " & WSTR(this.cRef))
'   FUNCTION = this.cRef
'   IF this.cRef = 0 THEN DELETE @this
   RETURN 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides storage for a new object pasted from the clipboard or read in from an
' Rich Text Format (RTF) stream.
' Parameter:
'    lplpstg - The address of the IStorage interface created for the new object.
' Return value:
'    Returns S_OK on success. If the method fails, it can return one of the following values.
'    E_INVALIDARG - There was an invalid argument.
'    E_OUTOFMEMORY - There was not enough memory to do the operation.
' Remarks:
'    This method must be implemented to allow cut, copy, paste, drag, and drop operations
'    of Component Object Model (COM) objects.
' ========================================================================================
FUNCTION CRichEditOleCallback.GetNewStorage (BYVAL lplpstg AS LPSTORAGE PTR) AS HRESULT
   CREC_DP("CRichEditOleCallback.GetNewStorage")
   DIM hr AS HRESULT
   DIM pILockBytes AS ILockBytes PTR
   hr = CreateILockBytesOnHGlobal(NULL, TRUE, @pILockBytes)
   IF FAILED(hr) THEN RETURN hr
   hr = StgCreateDocfileOnILockBytes(pILockBytes, _
        STGM_SHARE_EXCLUSIVE OR STGM_READWRITE OR STGM_CREATE, 0, lplpstg)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides the application and document-level interfaces and information required to support in-place activation.
' ========================================================================================
FUNCTION CRichEditOleCallback.GetInPlaceContext (BYVAL lplpFrame AS LPOLEINPLACEFRAME PTR, _
BYVAL lplpDoc AS LPOLEINPLACEUIWINDOW PTR, BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT
   CREC_DP("CRichEditOleCallback.GetInPlaceContext")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether or not the application is to display its container UI.
' ========================================================================================
FUNCTION CRichEditOleCallback.ShowContainerUI (BYVAL fShow AS WINBOOL) AS HRESULT
   CREC_DP("CRichEditOleCallback.ShowContainerUI")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Queries the application AS to whether an object should be inserted.
' The member is called when pasting and when reading Rich Text Format (RTF).
' ========================================================================================
FUNCTION CRichEditOleCallback.QueryInsertObject (BYVAL lpclsid AS LPCLSID, BYVAL lpstg AS LPSTORAGE, BYVAL cp AS LONG) AS HRESULT
   CREC_DP("CRichEditOleCallback.QueryInsertObject" & " " & WSTR(cp))
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sends notification that an object is about to be deleted from a rich edit control.
' The object is not necessarily being released when this member is called.
' ========================================================================================
FUNCTION CRichEditOleCallback.DeleteObject (BYVAL lpoleobj AS LPOLEOBJECT) AS HRESULT
   CREC_DP("CRichEditOleCallback.DeleteObject")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' During a paste operation or a drag event, determines if the data that is pasted or dragged should be accepted.
' ========================================================================================
FUNCTION CRichEditOleCallback.QueryAcceptData (BYVAL lpdataobj AS LPDATAOBJECT, BYVAL lpcfFormat AS CLIPFORMAT PTR, _
BYVAL reco AS DWORD, BYVAL fReally AS WINBOOL, BYVAL hMetaPict AS HGLOBAL) AS HRESULT
   CREC_DP("CRichEditOleCallback.QueryAcceptData")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates if the application should transition into or out of context-sensitive help mode.
' This method should implement the functionality described for IOleWindow::ContextSensitiveHelp.
' ========================================================================================
FUNCTION CRichEditOleCallback.ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
   CREC_DP("CRichEditOleCallback.ContextSensitiveHelp")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allows the client to supply its own clipboard object.
' ========================================================================================
FUNCTION CRichEditOleCallback.GetClipboardData (BYVAL lpchrg AS CHARRANGE PTR, BYVAL reco AS DWORD, _
BYVAL lplpdataobj AS LPDATAOBJECT PTR) AS HRESULT
   CREC_DP("CRichEditOleCallback.GetClipboardData")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allows the client to specify the effects of a drop operation.
' ========================================================================================
FUNCTION CRichEditOleCallback.GetDragDropEffect (BYVAL fDrag AS WINBOOL, BYVAL grfKeyState AS DWORD, BYVAL pdwEffect AS LPDWORD) AS HRESULT
   CREC_DP("CRichEditOleCallback.GetDragDropEffect")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Queries the application for a context menu to use on a right-click event.
' ========================================================================================
FUNCTION CRichEditOleCallback.GetContextMenu (BYVAL seltype AS WORD, BYVAL lpoleobj AS LPOLEOBJECT, _
BYVAL lpchrg AS CHARRANGE PTR, BYVAL lphmenu AS HMENU PTR) AS HRESULT
   CREC_DP("CRichEditOleCallback.GetContextMenu")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================
