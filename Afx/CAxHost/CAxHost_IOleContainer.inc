' ########################################################################################
' Microsoft Windows
' File: CAxHost_IOleContainer.inc
' Contents: OLE Container class
' Implementation of the IOleCOntainer interface.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' this CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
'                         *** CAxHost_IParseDisplayName class ***
' ########################################################################################

' ========================================================================================
' IParseDisplayName interface implemented
' Parses a displayable name string to convert it into a moniker for custom moniker implementations.
' Display name parsing is necessary when the end user inputs a string to identify a component,
' as in the following situations:
' - A compound document application that supports linked components typically supports the
'   Edit.Links... dialog box. Through this dialog box, the end user can enter a display name
'   to specify a new link source for a specified linked component. The compound document needs
'   to have this input string converted into a moniker.
' - A script language such as the macro language of a spreadsheet can allow textual references
'   to a component. The language's interpreter needs to have such a reference converted into
'    a moniker in order to execute the macro.
' This interface is not supported for use across machine boundaries.
' ========================================================================================
TYPE CAxHost_IParseDisplayName EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION ParseDisplayName (BYVAL pbc AS IBindCtx PTR, BYVAL pszDisplayName AS LPOLESTR, BYVAL pchEaten AS ULONG PTR, BYVAL ppmkOut AS IMoniker PTR PTR) AS HRESULT

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
CONSTRUCTOR CAxHost_IParseDisplayName
   CAXH_DP("CAxHost_IParseDisplayName Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_IParseDisplayName (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_IParseDisplayName constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_IParseDisplayName
   CAXH_DP("### CAxHost_IParseDisplayName destructor")
   m_pAxHost->m_pIParseDisplayNameImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_IParseDisplayName.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IParseDisplayName.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_IParseDisplayName AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAxHost_IParseDisplayName.Release () AS DWORD
   this.cRef -= 1
   CAXH_DP("CAxHost_IParseDisplayName Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses the display name returning a moniker corresponding to it.
' Not required. This method is only necessary for containers that link to controls or
' other embedded objects.
' ========================================================================================
FUNCTION CAxHost_IParseDisplayName.ParseDisplayName (BYVAL pbc AS IBindCtx PTR, BYVAL pszDisplayName AS LPOLESTR, BYVAL pchEaten AS ULONG PTR, BYVAL ppmkOut AS IMoniker PTR PTR) AS HRESULT
   CAXH_DP("CAxHost_IParseDisplayName ParseDisplayName")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================


' ########################################################################################
'                         *** CAxHost_IOleContainer class ***
' ########################################################################################

' ========================================================================================
' Ole Container implementation of the IParseDisplayName interface.
' This interface, as well as support for nested simple frames is optional.
' The IParseDisplayName interface parses a displayable name string to convert it into a
' moniker for custom moniker implementations. Display name parsing is necessary when the
' end user inputs a string to identify a component, as in the following situations:
' * A compound document application that supports linked components typically supports
'   the Edit:Links... dialog box. Through this dialog box, the end user can enter a
'   display name to specify a new link source for a specified linked component. The
'   compound document needs to have this input string converted into a moniker.
' * A script language such as the macro language of a spreadsheet can allow textual
'   references to a component. The language's interpreter needs to have such a reference
'   converted into a moniker in order to execute the macro.
' ========================================================================================
TYPE CAxHost_IOleContainer EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION ParseDisplayName (BYVAL pbc AS IBindCtx PTR, BYVAL pszDisplayName AS LPOLESTR, BYVAL pchEaten AS ULONG PTR, BYVAL ppmkOut AS IMoniker PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION EnumObjects (BYVAL grfFlags AS DWORD, BYVAL ppenum AS IEnumUnknown PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION LockContainer (BYVAL fLock AS WINBOOL) AS HRESULT

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
CONSTRUCTOR CAxHost_IOleContainer
   CAXH_DP("CAxHost_IOleContainer Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_IOleContainer (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_IOleContainer constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_IOleContainer
   CAXH_DP("### CAxHost_IOleContainer destructor")
   m_pAxHost->m_pIOleContainerImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_IOleContainer.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IOleContainer.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_IOleContainer AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAxHost_IOleContainer.Release () AS DWORD
   this.cRef -= 1
   CAXH_DP("CAxHost_IOleContainer Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses the display name returning a moniker corresponding to it.
' Not required. This method is only necessary for containers that link to controls or
' other embedded objects.
' ========================================================================================
FUNCTION CAxHost_IOleContainer.ParseDisplayName (BYVAL pbc AS IBindCtx PTR, BYVAL pszDisplayName AS LPOLESTR, BYVAL pchEaten AS ULONG PTR, BYVAL ppmkOut AS IMoniker PTR PTR) AS HRESULT
   CAXH_DP("CAxHost_IOleContainer ParseDisplayName")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enumerates objects in a container.
' Not required. This method is only necessary to enumerate ActiveX controls.
' ========================================================================================
FUNCTION CAxHost_IOleContainer.EnumObjects (BYVAL grfFlags AS DWORD, BYVAL ppenum AS IEnumUnknown PTR PTR) AS HRESULT
   CAXH_DP("CAxHost_IOleContainer EnumObjects")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Keeps container running until explicitly released.
' Not required. This method is only necessary for containers that link to controls or
' other embedded objects.
' ========================================================================================
FUNCTION CAxHost_IOleContainer.LockContainer (BYVAL fLock AS WINBOOL) AS HRESULT
   CAXH_DP("CAxHost_IOleContainer LockContainer")
   m_pAxHost->m_bLocked = fLock
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================
