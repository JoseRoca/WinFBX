' ########################################################################################
' TypeLib Browser
' File: TLB_TpeLib.bi
' Contents: TypeLib Browser interface definitions
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2016 José Roca. All Rights Reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "win/unknwnbase.bi"

' // The definition for BSTR in the FreeBASIC headers was incorrectly changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

#ifndef __Afx_IUnknown_INTERFACE_DEFINED__
#define __Afx_IUnknown_INTERFACE_DEFINED__
TYPE Afx_IUnknown AS Afx_IUnknown_
TYPE Afx_IUnknown_ EXTENDS OBJECT
    DECLARE ABSTRACT FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
    DECLARE ABSTRACT FUNCTION AddRef () AS ULONG
    DECLARE ABSTRACT FUNCTION Release () AS ULONG
END TYPE
TYPE AFX_LPUNKNOWN AS Afx_IUnknown PTR
#endif

#ifndef __Afx_IDispatch_INTERFACE_DEFINED__
#define __Afx_IDispatch_INTERFACE_DEFINED__
TYPE Afx_IDispatch AS Afx_IDispatch_
TYPE Afx_IDispatch_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL _
puArgErr AS UINT PTR) AS HRESULT
END TYPE
TYPE AFX_LPDISPATCH AS Afx_IDispatch PTR
#endif

TYPE Afx_ITypeLib AS Afx_ITypeLib_

' ########################################################################################
' Interface name = ITypeInfo
' Extracts information about the characteristics and capabilities of objects from type libraries.
' Inherited interface = IUnknown
' ########################################################################################

TYPE Afx_ITypeInfo AS Afx_ITypeInfo_

#ifndef __Afx_ITypeInfo_INTERFACE_DEFINED__
#define __Afx_ITypeInfo_INTERFACE_DEFINED__

TYPE Afx_ITypeInfo_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetTypeAttr (BYVAL ppTypeAttr AS TYPEATTR PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeComp (BYVAL ppTComp AS ITypeComp PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFuncDesc (BYVAL index AS UINT, BYVAL ppFuncDesc AS FUNCDESC PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetVarDesc (BYVAL index AS UINT, BYVAL ppVarDesc AS VARDESC PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetNames (BYVAL memid AS MEMBERID, BYVAL rgBstrNames AS AFX_BSTR PTR, BYVAL cMaxNames AS UINT, BYVAL pcNames AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetRefTypeOfImplType (BYVAL index AS UINT, BYVAL pRefType AS HREFTYPE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetImplTypeFlags (BYVAL index AS UINT, BYVAL pImplTypeFlags AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL pMemId AS MEMBERID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL pvInstance AS PVOID, BYVAL memid AS MEMBERID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDocumentation (BYVAL memid AS MEMBERID, BYVAL pBstrName AS AFX_BSTR PTR, BYVAL pBstrDocString AS AFX_BSTR PTR, BYVAL pdwHelpContext AS DWORD PTR, BYVAL pBstrHelpFile AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDllEntry (BYVAL memid AS MEMBERID, BYVAL invKind AS INVOKEKIND, BYVAL pBstrDllName AS AFX_BSTR PTR, BYVAL pBstrName AS AFX_BSTR PTR, BYVAL pwOrdinal AS WORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetRefTypeInfo (BYVAL hRefType AS HREFTYPE, BYVAL ppTInfo AS Afx_ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddressOfMember (BYVAL memid AS MEMBERID, BYVAL invKind AS INVOKEKIND, BYVAL ppv AS PVOID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateInstance (BYVAL pUnkOuter AS IUnknown PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS PVOID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetMops (BYVAL memid AS MEMBERID, BYVAL pBstrMops AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetContainingTypeLib (BYVAL ppTLib AS Afx_ITypeLib PTR PTR, BYVAL pIndex AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT SUB      ReleaseTypeAttr (BYVAL pTypeAttr AS TYPEATTR PTR)
   DECLARE ABSTRACT SUB      ReleaseFuncDesc (BYVAL pFuncDesc AS FUNCDESC PTR)
   DECLARE ABSTRACT SUB      ReleaseVarDesc (BYVAL pVarDesc AS VARDESC PTR)
END TYPE

#endif

' ########################################################################################
' Interface name = ITypeInfo2
' Extracts information about the characteristics and capabilities of objects from type libraries.
' Can be cast to an ITypeInfo instead of using the calls QueryInterface and Release.
' Inherited interface = ITypeInfo
' ########################################################################################

TYPE Afx_ITypeInfo2 AS Afx_ITypeInfo2_

#ifndef __Afx_ITypeInfo2_INTERFACE_DEFINED__
#define __Afx_ITypeInfo2_INTERFACE_DEFINED__

TYPE Afx_ITypeInfo2_ EXTENDS Afx_ITypeInfo_
   DECLARE ABSTRACT FUNCTION GetTypeKind (BYVAL pTypeKind AS TYPEKIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeFlags (BYVAL pTypeFlags AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFuncIndexOfMemId (BYVAL memid AS MEMBERID, BYVAL invKind AS INVOKEKIND, BYVAL pFuncIndex AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetVarIndexOfMemId (BYVAL memid AS MEMBERID, BYVAL pVarIndex AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetCustData (BYVAL guid AS CONST GUID CONST PTR, BYVAL pVarVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFuncCustData (BYVAL index AS UINT, BYVAL guid AS CONST GUID CONST PTR, BYVAL pVarVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetParamCustData (BYVAL indexFunc AS UINT, BYVAL indexParam AS UINT, BYVAL guid AS CONST GUID CONST PTR, BYVAL pVarVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetVarCustData (BYVAL index AS UINT, BYVAL guid AS CONST GUID CONST PTR, BYVAL pVarVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetImplTypeCustData (BYVAL index AS UINT, BYVAL guid AS CONST GUID CONST PTR, BYVAL pVarVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDocumentation2 (BYVAL memid AS MEMBERID, BYVAL lcid AS LCID, BYVAL pbstrHelpString AS AFX_BSTR PTR, BYVAL pdwHelpStringContext AS DWORD PTR, BYVAL pbstrHelpStringDll AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAllCustData (BYVAL pCustData AS CUSTDATA PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAllFuncCustData (BYVAL index AS UINT, BYVAL pCustData AS CUSTDATA PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAllParamCustData (BYVAL indexFunc AS UINT, BYVAL indexParam AS UINT, BYVAL pCustData AS CUSTDATA PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAllVarCustData (BYVAL index AS UINT, BYVAL pCustData AS CUSTDATA PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAllImplTypeCustData (BYVAL index AS UINT, BYVAL pCustData AS CUSTDATA PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################
' Interface name = ITypeLib
' Extracts information about a type library, the data that describes a set of objects.
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ITypeLib_INTERFACE_DEFINED__
#define __Afx_ITypeLib_INTERFACE_DEFINED__

TYPE Afx_ITypeLib_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount () AS UINT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL index AS UINT, BYVAL ppTInfo AS Afx_ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfoType (BYVAL index AS UINT, BYVAL pTKind AS TYPEKIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfoOfGuid (BYVAL guid AS CONST GUID CONST PTR, BYVAL ppTinfo AS Afx_ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetLibAttr (BYVAL ppTLibAttr AS TLIBATTR PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeComp (BYVAL ppTComp AS ITypeComp PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDocumentation (BYVAL index AS INT_, BYVAL pBstrName AS AFX_BSTR PTR, BYVAL pBstrDocString AS AFX_BSTR PTR, BYVAL pdwHelpContext AS DWORD PTR, BYVAL pBstrHelpFile AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsName (BYVAL szNameBuf AS LPOLESTR, BYVAL lHashVal AS ULONG, BYVAL pfName AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION FindName (BYVAL szNameBuf AS LPOLESTR, BYVAL lHashVal AS ULONG, BYVAL ppTInfo AS Afx_ITypeInfo PTR PTR, BYVAL rgMemId AS MEMBERID PTR, BYVAL pcFound AS USHORT PTR) AS HRESULT
   DECLARE ABSTRACT SUB      ReleaseTLibAttr (BYVAL pTLibAttr AS TLIBATTR PTR)
END TYPE

#endif

' ########################################################################################
' Interface name = ITypeLib2
' Extracts information about a type library, the data that describes a set of objects.
' The ITypeLib2 interface inherits from the ITypeLibinterface. This allows ITypeLib to
' cast to an ITypeLib2 in performance-sensitive cases, rather than perform extra
' QueryInterface and Release calls.
' Inherited interface = ITypeLib
' ########################################################################################

TYPE Afx_ITypeLib2 AS Afx_ITypeLib2_

#ifndef __Afx_ITypeLib2_INTERFACE_DEFINED__
#define __Afx_ITypeLib2_INTERFACE_DEFINED__

TYPE Afx_ITypeLib2_ EXTENDS Afx_ITypeLib
   DECLARE ABSTRACT FUNCTION GetCustData (BYVAL guid AS CONST GUID CONST PTR, BYVAL pVarVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetLibStatistics (BYVAL pcUniqueNames AS ULONG PTR, BYVAL pcchUniqueNames AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDocumentation2 (BYVAL index AS INT_, BYVAL lcid AS LCID, BYVAL pbstrHelpString AS AFX_BSTR PTR, BYVAL pdwHelpStringContext AS DWORD PTR, BYVAL pbstrHelpStringDll AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAllCustData (BYVAL pCustData AS CUSTDATA PTR) AS HRESULT
END TYPE

#endif
