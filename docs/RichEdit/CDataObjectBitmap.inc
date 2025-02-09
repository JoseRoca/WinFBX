' ########################################################################################
' Microsoft Windows
' File: CDataObjectBitmap.inc
' Contents: Enables data transfer and notification of changes in data.
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "win/richole.bi"

NAMESPACE Afx

' ========================================================================================
' CDataObjectBitmap class
' Implements the IDataObject interface
' ========================================================================================
TYPE CDataObjectBitmap EXTENDS Object

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
	DECLARE VIRTUAL FUNCTION GetData (BYREF fmtEtc AS FORMATETC, BYREF stgm AS STGMEDIUM) AS HRESULT
	DECLARE VIRTUAL FUNCTION GetDataHere (BYREF fmtEtc AS FORMATETC, BYREF stgm AS STGMEDIUM) AS HRESULT
	DECLARE VIRTUAL FUNCTION QueryGetData (BYREF fmtEtc AS FORMATETC) AS HRESULT
	DECLARE VIRTUAL FUNCTION GetCanonicalFormatEtc (BYREF fmtEtcIn AS FORMATETC, BYREF fmtEtcOut AS FORMATETC) AS HRESULT
	DECLARE VIRTUAL FUNCTION SetData (BYREF fmtEtc AS FORMATETC, BYREF stgm AS STGMEDIUM, BYVAL fRelease AS WINBOOL) AS HRESULT
	DECLARE VIRTUAL FUNCTION EnumFormatEtc (BYVAL dwDirection AS DWORD, BYREF enumFmtEtc AS IEnumFORMATETC PTR) AS HRESULT
	DECLARE VIRTUAL FUNCTION DAdvise (BYREF fmtEtc AS FORMATETC, BYVAL advf AS DWORD, BYVAL pAdvSink AS IAdviseSink PTR, BYREF dwConnection AS DWORD) AS HRESULT
	DECLARE VIRTUAL FUNCTION DUnadvise (BYVAL dwConnection AS DWORD) AS HRESULT
	DECLARE VIRTUAL FUNCTION EnumDAdvise (BYREF enumAdvise AS IEnumSTATDATA PTR) AS HRESULT

Public:
   DECLARE CONSTRUCTOR (BYREF fmtEtc AS FORMATETC, BYREF stgm AS STGMEDIUM)
   DECLARE DESTRUCTOR

Private:
   
   DIM m_fmtEtc AS FORMATETC
   DIM m_stgm AS STGMEDIUM
   ' // The reference count of this object must start at 1
   cRef AS ULONG = 1

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CDataObjectBitmap (BYREF fmtEtc AS FORMATETC, BYREF stgm AS STGMEDIUM)
   CREC_DP("CDataObjectBitmap constructor")
   m_fmtEtc = fmtEtc
   m_stgm = stgm
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CDataObjectBitmap
   CREC_DP("### CDataObjectBitmap destructor")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CDataObjectBitmap.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CREC_DP("CDataObjectBitmap.QueryInterface: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   DIM IID_IDataObject_ AS GUID = (&h0000010E, &h0000, &h0000, {&hC0, &h00, &h00, &h00, &h00, &h00, &h00, &h46})
   IF IsEqualIID(riid, @IID_IDataObject_) OR IsEqualIID(riid, @IID_IUnknown) THEN
      *ppvObj = @this
      RETURN S_OK
   END IF
   CREC_DP("CDataObjectBitmap.QueryInterface - END")
   RETURN E_NOINTERFACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CDataObjectBitmap.AddRef () AS ULONG
   InterlockedIncrement(@cRef)
   CREC_DP("CDataObjectBitmap AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CDataObjectBitmap.Release () AS DWORD
   InterlockedDecrement(@cRef)
   CREC_DP("CDataObjectBitmap Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Called by a data consumer to obtain data from a source data object. The GetData method
' renders the data described in the specified FORMATETC structure and transfers it through
' the specified STGMEDIUM structure. The caller then assumes responsibility for releasing'
' the STGMEDIUM structure.
' ========================================================================================
FUNCTION CDataObjectBitmap.GetData (BYREF fmtEtc AS FORMATETC, BYREF stgm AS STGMEDIUM) AS HRESULT
   CREC_DP("CDataObjectBitmap GetData")
   IF @fmtEtc = 0 THEN RETURN DV_E_FORMATETC
   IF @stgm = 0 THEN RETURN STG_E_MEDIUMFULL
   fmtEtc = m_fmtEtc
   stgm = m_stgm
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Called by a data consumer to obtain data from a source data object. This method differs
' from the GetData method in that the caller must allocate and free the specified storage medium.
' ========================================================================================
FUNCTION CDataObjectBitmap.GetDataHere (BYREF fmtEtc AS FORMATETC, BYREF stgm AS STGMEDIUM) AS HRESULT
   CREC_DP("CDataObjectBitmap GetDataHere")
'   IF @fmtEtc = 0 THEN RETURN DV_E_FORMATETC
'   IF @stgm = 0 THEN RETURN STG_E_MEDIUMFULL
'   fmtEtc = m_fmtEtc
'   stgm = m_stgm
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the data object is capable of rendering the data as specified. Objects
' attempting a paste or drop operation can call this method before calling IDataObject.GetData
' to get an indication of whether the operation may be successful.
' ========================================================================================
FUNCTION CDataObjectBitmap.QueryGetData (BYREF fmtEtc AS FORMATETC) AS HRESULT
   CREC_DP("CDataObjectBitmap QueryGetData")
   IF @fmtEtc = 0 THEN RETURN DV_E_FORMATETC
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides a potentially different but logically equivalent FORMATETC structure. You use this
' method to determine whether two different FORMATETC structures would return the same data,
' removing the need for duplicate rendering.
' ========================================================================================
FUNCTION CDataObjectBitmap.GetCanonicalFormatEtc (BYREF fmtEtcIn AS FORMATETC, BYREF fmtEtcOut AS FORMATETC) AS HRESULT
   CREC_DP("CDataObjectBitmap GetCanonicalFormatEtc")
   IF @fmtEtcIn = 0 THEN RETURN DV_E_FORMATETC
   IF @fmtEtcOut = 0 THEN RETURN DV_E_FORMATETC
   fmtEtcOut = fmtEtcIn
   fmtEtcOut.ptd = NULL
   RETURN DATA_S_SAMEFORMATETC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Called by an object containing a data source to transfer data to the object that implements this method.
' ========================================================================================
FUNCTION CDataObjectBitmap.SetData (BYREF fmtEtc AS FORMATETC, BYREF stgm AS STGMEDIUM, BYVAL fRelease AS WINBOOL) AS HRESULT
   CREC_DP("CDataObjectBitmap SetData")
   IF @fmtEtc = 0 THEN RETURN DV_E_FORMATETC
   IF @stgm = 0 THEN RETURN STG_E_MEDIUMFULL
   m_fmtEtc = fmtEtc
   m_stgm = stgm
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an object to enumerate the formats supported by a data object.
' ========================================================================================
FUNCTION CDataObjectBitmap.EnumFormatEtc (BYVAL dwDirection AS DWORD, BYREF enumFmtEtc AS IEnumFORMATETC PTR) AS HRESULT
   CREC_DP("CDataObjectBitmap EnumFormatEtc - Direction: " & WSTR(dwDirection))
   RETURN OLE_S_USEREG
END FUNCTION
' ========================================================================================

' ========================================================================================
' Called by an object supporting an advise sink to create a connection between a data object
' and the advise sink. This enables the advise sink to be notified of changes in the data of the object.
' ========================================================================================
FUNCTION CDataObjectBitmap.DAdvise (BYREF fmtEtc AS FORMATETC, BYVAL advf AS DWORD, BYVAL pAdvSink AS IAdviseSink PTR, BYREF dwConnection AS DWORD) AS HRESULT
   CREC_DP("CDataObjectBitmap DAdvise")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Destroys a notification connection that had been previously set up.
' ========================================================================================
FUNCTION CDataObjectBitmap.DUnadvise (BYVAL dwConnection AS DWORD) AS HRESULT
   CREC_DP("CDataObjectBitmap DUnadvise")
   RETURN OLE_E_ADVISENOTSUPPORTED
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an object that can be used to enumerate the current advisory connections.
' ========================================================================================
FUNCTION CDataObjectBitmap.EnumDAdvise (BYREF enumAdvise AS IEnumSTATDATA PTR) AS HRESULT
   CREC_DP("CDataObjectBitmap EnumDAdvise")
   RETURN OLE_E_ADVISENOTSUPPORTED
END FUNCTION
' ========================================================================================

END NAMESPACE



' ========================================================================================
' Inserts an image in the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.InsertObject2 ( _
   BYREF wszFileName AS WSTRING _                    ' // Name of the file to load
 ) AS BOOLEAN                                        ' // TRUE or FALSE

   CREC_DP("#### BEGIN - CRichEditCtrl.InsertObject2")

   CREC_DP("#### CRichEditCtrl.InsertObject2 - Get bitmap handle")
   ' // Initialize GDI+
   DIM token AS ULONG_PTR = AfxGdipInit
   ' // Load the image as a bitmap
   DIM pBitmap AS GpBitmap PTR
   GdipCreateBitmapFromFile(wszFileName, @pBitmap)
   IF pBitmap = NULL THEN RETURN FALSE
   ' // Get the bitmap handle
   DIM hBitmap AS HBITMAP
   GdipCreateHBITMAPFromBitmap(pBitmap, @hBitmap, 0)
   IF hBitmap = NULL THEN RETURN FALSE

   ' // Set up the STGMEDIUM structure
   DIM stgm AS STGMEDIUM
   stgm.tymed = TYMED_GDI
   stgm.hBitmap = hBitmap
   stgm.pUnkForRelease = NULL

   ' // Set up the data format
   DIM formatEtc AS FORMATETC
   formatEtc.cfFormat = CF_BITMAP
   formatEtc.ptd = NULL
   formatEtc.dwAspect = DVASPECT_CONTENT
   formatEtc.lindex = -1
   formatEtc.tymed = TYMED_GDI

   ' // Create an instance of the CDataObjectBitmap class
   CREC_DP("#### CRichEditCtrl.InsertObject2 - CDataObjectBitmap")
   DIM pDataObject AS CDataObjectBitmap = CDataObjectBitmap(formatEtc, stgm)
   ' // Check if the data object has a valid format
   CREC_DP("#### CRichEditCtrl.InsertObject2 - OleQueryCreateFromData")
   IF OleQueryCreateFromData(cast(ANY PTR, @pDataObject)) <> OLE_S_STATIC THEN RETURN FALSE

   ' // Get the IRichEditOle interface
   CREC_DP("#### CRichEditCtrl.InsertObject2 - GetOleInterface")
   DIM pRichEditOle AS IRichEditOle PTR
   this.SetResult(SendMessageW(m_hRichEdit, EM_GETOLEINTERFACE, 0, cast(LPARAM, @pRichEditOle)))
   IF pRichEditOle = NULL THEN RETURN FALSE

   ' // Get a pointer to the display site
   CREC_DP("#### CRichEditCtrl.InsertObject2 - GetClientSite")
   DIM pClientSite AS IOleClientSite PTR
   this.SetResult(pRichEditOle->lpvtbl->GetClientSite(pRichEditOle, @pClientSite))
   IF FAILED(this.GetLastResult) THEN
      pRichEditOle->lpvtbl->Release(pRichEditOle)
      RETURN FALSE
   END IF

   ' // Create structured storage
   CREC_DP("#### CRichEditCtrl.InsertObject2 - CreateILockBytesOnHGlobal")
   DIM pLockBytes AS ILockBytes PTR
   this.SetResult(CreateILockBytesOnHGlobal(NULL, TRUE, @pLockBytes))
   IF FAILED(this.GetLastResult)THEN
      pClientSite->lpvtbl->Release(pClientSite)
      pRichEditOle->lpvtbl->Release(pRichEditOle)
      RETURN FALSE
   END IF

   CREC_DP("#### CRichEditCtrl.InsertObject2 - StgCreateDocfileOnILockBytes")
   DIM pStorage AS IStorage PTR
   this.SetResult(StgCreateDocfileOnILockBytes(pLockBytes, _
      STGM_SHARE_EXCLUSIVE OR STGM_CREATE OR STGM_READWRITE, 0, @pStorage))
   IF FAILED(this.GetLastResult) THEN
      pLockBytes->lpvtbl->Release(pLockBytes)
      pClientSite->lpvtbl->Release(pClientSite)
      pRichEditOle->lpvtbl->Release(pRichEditOle)
      RETURN FALSE
   END IF

   ' // Creates a static object
   CREC_DP("#### CRichEditCtrl.InsertObject2 - OleCreateStaticFromData")
   DIM pObject AS IOleObject PTR
   OleCreateStaticFromData(cast(ANY PTR, @pDataObject), @IID_IOleObject, OLERENDER_FORMAT, @formatEtc, pClientSite, pStorage, @pObject)
   IF pObject = NULL THEN
      pStorage->lpvtbl->Release(pStorage)
      pLockBytes->lpvtbl->Release(pLockBytes)
      pClientSite->lpvtbl->Release(pClientSite)
      pRichEditOle->lpvtbl->Release(pRichEditOle)
      RETURN FALSE
   END IF

   ' // Get the object's class ifentifier
   CREC_DP("#### CRichEditCtrl.InsertObject2 - GetUserClassID")
   DIM clsid AS CLSID
   this.SetResult(pObject->lpvtbl->GetUserClassID(pObject, @clsid))
   IF FAILED(this.GetLastResult) THEN
      pObject->lpvtbl->Release(pObject)
      pStorage->lpvtbl->Release(pStorage)
      pLockBytes->lpvtbl->Release(pLockBytes)
      pClientSite->lpvtbl->Release(pClientSite)
      pRichEditOle->lpvtbl->Release(pRichEditOle)
      RETURN FALSE
   END IF

   ' // Set up object info
   DIM reobject AS REOBJECT
   reobject.cbStruct = SIZEOF(REOBJECT)
   reobject.clsid = clsid
   reobject.cp = REO_CP_SELECTION
   reobject.dvaspect = DVASPECT_CONTENT
   reobject.dwFlags = 0 'REO_RESIZABLE OR REO_BELOWBASELINE
   reobject.dwUser = 0
   reobject.poleobj = pObject
   reobject.polesite = pClientSite
   reobject.pstg = pStorage
   DIM sizel AS SIZEL
   DIM dwWidth AS DWORD, dwHeight AS DWORD
   DIM nStatus AS LONG = AfxGdipGetImageSizeFromFile(wszFileName, @dwWidth, @dwHeight)
   IF nStatus = 0 THEN
      sizel.cx = AfxPixelsToHiMetricX(dwWidth * m_ScalingRatio)
      sizel.cy = AfxPixelsToHiMetricY(dwHeight * m_ScalingRatio)
   END IF
   reobject.sizel = sizel

   ' // Insert the object
   CREC_DP("#### CRichEditCtrl.InsertObject2 - InsertObject")
   this.SetResult(pRichEditOle->lpvtbl->InsertObject(pRichEditOle, @reobject))
   pObject->lpvtbl->Release(pObject)
   pStorage->lpvtbl->Release(pStorage)
   pLockBytes->lpvtbl->Release(pLockBytes)
   pClientSite->lpvtbl->Release(pClientSite)
   pRichEditOle->lpvtbl->Release(pRichEditOle)
   IF FAILED(this.GetLastResult) THEN RETURN FALSE
   
   CREC_DP("#### END - CRichEditCtrl.InsertObject2 hr = " & WSTR(this.GetLastResult))
   ' // Shutdown GDI+
   GdipDisposeImage(cast(GpImage PTR, pBitmap))
   AfxGdipShutdown(token)

   ' // Return success
   RETURN TRUE

END FUNCTION
' ========================================================================================
