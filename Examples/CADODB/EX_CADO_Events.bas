' ########################################################################################
' Microsoft Windows
' Contents: ADO Events
' Compiler: FreeBasic 32 & 64 bit
' Note: Error checking ommited for brevity.
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

'#CONSOLE ON
#define UNICODE
#define _CADODB_DEBUG_ 1
#include "Afx/CADODB/CADODB.inc"
USING Afx

' ########################################################################################
' Interface name: ConnectionEventsVt
' IID: {00001402-0000-0010-8000-00AA006D2EA4}
' ########################################################################################
TYPE CConnectionEventsVtImpl EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION InfoMessage (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION BeginTransComplete (BYVAL TransactionLevel AS LONG, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION CommitTransComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION RollbackTransComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION WillExecute (BYVAL Source AS AFX_BSTR PTR, BYVAL CursorType AS CursorTypeEnum PTR, BYVAL LockType AS LockTypeEnum PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ExecuteComplete (BYVAL RecordsAffected AS LONG, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION WillConnect (BYVAL ConnectionString AS AFX_BSTR PTR, BYVAL UserID AS AFX_BSTR PTR, BYVAL Password AS AFX_BSTR PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ConnectComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION Disconnect (BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
CONSTRUCTOR CConnectionEventsVtImpl
   CADODB_DP("CConnectionEventsVtImpl CONSTRUCTOR")
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Destructor
' =====================================================================================
DESTRUCTOR CConnectionEventsVtImpl
   CADODB_DP("CConnectionEventsVtImpl DESTRUCTOR")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
FUNCTION CConnectionEventsVtImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CADODB_DP("CConnectionEventsVtImpl.QueryInterface - " & AfxGuidText(riid))
   DIM IID_ConnectionEvents_ AS GUID = (&h00001400, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
   DIM IID_ConnectionEventsVt_ AS GUID = (&h00001402, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @IID_ConnectionEventsVt_) OR _
      IsEqualIID(riid, @IID_IDispatch) THEN
'      IsEqualIID(riid, @IID_IUnknown) THEN
      *ppvObj = @this
      ' // Not really needed, since this is not a COM object
      cast(Afx_IUnknown PTR, *ppvObj)->AddRef
      RETURN S_OK
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' ========================================================================================

' ========================================================================================
FUNCTION CConnectionEventsVtImpl.AddRef () AS ULONG
   CADODB_DP("CConnectionEventsVtImpl.AddRef")
   RETURN 2
END FUNCTION
' ========================================================================================

' ========================================================================================
FUNCTION CConnectionEventsVtImpl.Release () AS ULONG
   CADODB_DP("CConnectionEventsVtImpl.Release")
   RETURN 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' The number of type information interfaces provided by the object. If the object provides
' type information, this number is 1; otherwise the number is 0.
' ========================================================================================
'FUNCTION CConnectionEventsVtImpl.GetTypeInfoCount (BYVAL pctInfo AS UINT PTR) AS HRESULT
'   CADODB_DP("CConnectionEventsVtImpl.GetTypeInfoCount")
'   *pctInfo = 0
'   RETURN E_NOTIMPL
'END FUNCTION
' ========================================================================================

' ========================================================================================
' The type information for an object, which can then be used to get the type information
' for an interface.
' ========================================================================================
'FUNCTION CConnectionEventsVtImpl.GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
'   CADODB_DP("CConnectionEventsVtImpl.GetTypeInfo")
'   RETURN E_NOTIMPL
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Maps a single member and an optional set of argument names to a corresponding set of
' integer DISPIDs, which can be used on subsequent calls to Invoke.
' ========================================================================================
'FUNCTION CConnectionEventsVtImpl.GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
'   CADODB_DP("CConnectionEventsVtImpl.GetIDsOfNames")
'   RETURN E_NOTIMPL
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides access to properties and methods exposed by an object.
' Parameters in the rgvarg array of variants of the DISPPARAMS structures are zero based
' and in reverse order.
' ========================================================================================
'FUNCTION CConnectionEventsVtImpl.Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
'   OutputDebugStringW("******* CConnectionEventsVtImpl.Invoke dispid = " & WSTR(dispIdMember))

''   SELECT CASE dispIdMember

''   END SELECT

'   RETURN S_OK

'END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
FUNCTION CConnectionEventsVtImpl.InfoMessage (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("CConnectionEventsVtImpl.InfoMessage")
   RETURN S_OK
END FUNCTION
' ========================================================================================


' ========================================================================================
' ========================================================================================
FUNCTION CConnectionEventsVtImpl.BeginTransComplete (BYVAL TransactionLevel AS LONG, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("CConnectionEventsVtImpl.BeginTransComplete")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
FUNCTION CConnectionEventsVtImpl.CommitTransComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("CConnectionEventsVtImpl.CommitTransComplete")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
FUNCTION CConnectionEventsVtImpl.RollbackTransComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("CConnectionEventsVtImpl.RollbackTransComplete")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
FUNCTION CConnectionEventsVtImpl.WillExecute (BYVAL Source AS AFX_BSTR PTR, BYVAL CursorType AS CursorTypeEnum PTR, BYVAL LockType AS LockTypeEnum PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("CConnectionEventsVtImpl.WillExecute")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
FUNCTION CConnectionEventsVtImpl.ExecuteComplete (BYVAL RecordsAffected AS LONG, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("CConnectionEventsVtImpl.ExecuteComplete")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
FUNCTION CConnectionEventsVtImpl.WillConnect (BYVAL ConnectionString AS AFX_BSTR PTR, BYVAL UserID AS AFX_BSTR PTR, BYVAL Password AS AFX_BSTR PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("CConnectionEventsVtImpl.WillConnect")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
FUNCTION CConnectionEventsVtImpl.ConnectComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("CConnectionEventsVtImpl.ConnectComplete")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
FUNCTION CConnectionEventsVtImpl.Disconnect (BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("CConnectionEventsVtImpl.Disconnect")
   RETURN S_OK
END FUNCTION
' ========================================================================================

'++++++++++
' ########################################################################################
' Main application
' ########################################################################################

DIM hr AS HRESULT
' // Open the connection
DIM pConnection AS CAdoConnection
' // Connect events
DIM dwConnCookie AS DWORD
DIM IID_ConnectionEventsVt_ AS GUID = (&h00001400, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
DIM pConnEvents AS CConnectionEventsVtImpl PTR = NEW CConnectionEventsVtImpl
print *pConnection, pConnection.m_pConnection
hr = AfxAdvise(*pConnection, pConnEvents, IID_ConnectionEventsVt_, @dwConnCookie)
print hex(hr)
' && it fails with error 80040200 - class not registered
' // Open the connection
'pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
'DIM pRecordset AS CAdoRecordset
'DIM cbsSource AS CBSTR = "SELECT * FROM Authors"
'hr = pRecordset.Open(cbsSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Parse the recordset
'DO
'   ' // While not at the end of the recordset...
'   IF pRecordset.EOF THEN EXIT DO
'   ' // Get the content of the "Author" column
'   DIM cvRes AS CVARIANT = pRecordset.Collect("Author")
'   PRINT cvRes
'   ' // Fetch the next row
'   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
'LOOP

IF pConnEvents THEN
'   IF dwConnCookie THEN AfxUnadvise(*pConnection, @IID_ConnectionEventsVt_, dwConnCookie)
'   Delete pConnEvents
END IF

PRINT
PRINT "Press any key..."
SLEEP
