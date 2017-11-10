' ########################################################################################
' Microsoft Windows
' File: CAxHost_IServiceProvider.inc
' Contents: OLE Container class
' Implementation of the IServiceProvider interface.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' this CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
'                          *** CAxHost_IServiceProvider class ***
' ########################################################################################

' ========================================================================================
' This function must be called by ordinal because it isn't exported by name in XP and below.
' ========================================================================================
PRIVATE FUNCTION AfxAxHostQueryService ( _
   BYVAL punk AS IUnknown PTR _                           ' __in IUnknown* punk
 , BYVAL guidService AS REFGUID _                         ' __in REFGUID guidService
 , BYVAL riid AS REFGUID _                                ' __in REFIID riid
 , BYVAL ppvOut AS LPVOID PTR _                           ' __out void ** ppvOut
 ) AS HRESULT                                             ' HRESULT

   ' // Load the shell library
   DIM AS ANY PTR pLib = DyLibLoad("Shlwapi.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   ' // Get the address of the function (ordinal 176)
   DIM pAfxAxHostQueryService AS FUNCTION (BYVAL punk AS IUnknown PTR, BYVAL guidService AS REFGUID, _
       BYVAL riid AS REFGUID, BYVAL ppvOut AS LPVOID PTR) AS HRESULT
   pAfxAxHostQueryService = DyLibSymbol(pLib, 176)
   IF pAfxAxHostQueryService = NULL THEN
      FUNCTION = E_FAIL
   ELSE
      FUNCTION = pAfxAxHostQueryService(pUnk, guidService, riid, ppvOut)
   END IF
   FreeLibrary pLib

END FUNCTION
' ========================================================================================

' ========================================================================================
' IServiceProvider interface implemented
' Provides a generic access mechanism to locate a GUID-identified service.
' ========================================================================================
TYPE CAxHost_IServiceProvider EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION QueryService (BYVAL guidService AS const GUID const PTR, BYVAL riid AS const IID const PTR, BYVAL ppvObject AS ANY PTR PTR) AS HRESULT

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pAxHost AS CAxHost PTR)
   DECLARE DESTRUCTOR

Private:
   cRef AS ULONG = 0
   m_pAxHost AS CAxHost PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CAxHost_IServiceProvider
   CAXH_DP("CAxHost_IServiceProvider Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_IServiceProvider (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_IServiceProvider constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_IServiceProvider
   CAXH_DP("### CAxHost_IServiceProvider destructor")
   m_pAxHost->m_pIServiceProviderImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_IServiceProvider.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IServiceProvider.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_IServiceProvider AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAxHost_IServiceProvider.Release () AS DWORD
   this.cRef -= 1
   CAXH_DP("CAxHost_IServiceProvider Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Acts as the factory method for any services exposed through an implementation of
' IServiceProvider.
' ========================================================================================
FUNCTION CAxHost_IServiceProvider.QueryService (BYVAL guidService AS const GUID const PTR, BYVAL riid AS const IID const PTR, BYVAL ppvObject AS ANY PTR PTR) AS HRESULT
   DIM pwsz AS WSTRING PTR
   StringFromIID(guidService, CAST(LPOLESTR PTR, @pwsz))
   DIM s AS STRING : IF pwsz THEN s = *pwsz : CoTaskMemFree(pwsz)
   CAXH_DP("CAxHost_IServiceProvider QueryService : " & s)
   ' // Handle SID_SInPlaceBrowser (needed to enable the filesystem object to navigate
   ' // in-place within the WebBrowser control when running on Windows 7+)
   IF IsEqualIID(@guidService, @SID_SInPlaceBrowser) THEN
      RETURN AfxAxHostQueryService(CAST(IUnknown PTR, m_pAxHost->m_pOcx), @SID_SShellBrowser, riid, @ppvObject)
   END If
   RETURN E_NOINTERFACE
END FUNCTION
' ========================================================================================
