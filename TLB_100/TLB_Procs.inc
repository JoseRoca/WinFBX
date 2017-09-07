' ########################################################################################
' TypeLib Browser
' File: TLB_Procs.inc
' Contents: TypeLib Browser procedures
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2016 José Roca. All Rights Reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' ========================================================================================
' Customized message box.
' ========================================================================================
FUNCTION TLB_MsgBox (BYVAL hwnd AS HWND, BYVAL pwszMessage AS WSTRING PTR, BYVAL dwStyle AS DWORD = 0, BYVAL pwszCaption AS WSTRING PTR = NULL) AS LONG

   DIM wszCaption AS WSTRING * 260 = "TypeLib Browser"
   IF pwszCaption THEN
      IF LEN(*pwszCaption) THEN wszCaption = *pwszCaption
   END IF
   IF dwStyle = 0 THEN dwStyle = MB_OK

   ' // Initializes MSGBOXPARAMSW
   DIM mbp AS MSGBOXPARAMSW
   mbp.cbSize = SIZEOF(mbp)                 ' // Size of the structure
   mbp.hwndOwner = hWnd                     ' // Handle of main window
   mbp.hInstance = GetModuleHandleW(NULL)   ' // Instance of application
   mbp.lpszText = pwszMessage               ' // Text of the message
   mbp.lpszCaption = pwszCaption            ' // Caption
   mbp.dwStyle = dwStyle OR MB_USERICON     ' // Style
   mbp.lpszIcon = MAKEINTRESOURCEW(101)  ' // Icon identifier in the resource file

   FUNCTION = MessageBoxIndirectW(@mbp)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts LibFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_LibFlagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN FUNCTION = " [None]" : EXIT FUNCTION
   IF (wFlags AND LIBFLAG_FRESTRICTED) = LIBFLAG_FRESTRICTED THEN strFlags += " [Restricted]"
   IF (wFlags AND LIBFLAG_FCONTROL) = LIBFLAG_FCONTROL THEN strFlags += " [Control]"
   IF (wFlags AND LIBFLAG_FHIDDEN) = LIBFLAG_FHIDDEN THEN strFlags += " [Hidden]"
   IF (wFlags AND LIBFLAG_FHASDISKIMAGE) = LIBFLAG_FHASDISKIMAGE THEN strFlags += " [HasDiskImage]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts InterfaceFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_InterfaceFlagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN FUNCTION = " [None]" : EXIT FUNCTION
   IF (wFlags AND TYPEFLAG_FAPPOBJECT) = TYPEFLAG_FAPPOBJECT THEN strFlags += " [Application]"
   IF (wFlags AND TYPEFLAG_FCANCREATE) = TYPEFLAG_FCANCREATE THEN strFlags += " [Cancreate]"
   IF (wFlags AND TYPEFLAG_FLICENSED) = TYPEFLAG_FLICENSED THEN strFlags += " [Licensed]"
   IF (wFlags AND TYPEFLAG_FPREDECLID) = TYPEFLAG_FPREDECLID THEN strFlags += " [Predefined]"
   IF (wFlags AND TYPEFLAG_FHIDDEN) = TYPEFLAG_FHIDDEN THEN strFlags += " [Hidden]"
   IF (wFlags AND TYPEFLAG_FCONTROL) = TYPEFLAG_FCONTROL THEN strFlags += " [Control]"
   IF (wFlags AND TYPEFLAG_FDUAL) = TYPEFLAG_FDUAL THEN strFlags += " [Dual]"
   IF (wFlags AND TYPEFLAG_FNONEXTENSIBLE) = TYPEFLAG_FNONEXTENSIBLE THEN strFlags += " [Nonextensible]"
   IF (wFlags AND TYPEFLAG_FOLEAUTOMATION) = TYPEFLAG_FOLEAUTOMATION THEN strFlags += " [Oleautomation]"
   IF (wFlags AND TYPEFLAG_FRESTRICTED) = TYPEFLAG_FRESTRICTED THEN strFlags += " [Restricted]"
   IF (wFlags AND TYPEFLAG_FAGGREGATABLE) = TYPEFLAG_FAGGREGATABLE THEN strFlags += " [Aggregatable]"
   IF (wFlags AND TYPEFLAG_FREPLACEABLE) = TYPEFLAG_FREPLACEABLE THEN strFlags += " [Replaceable]"
   IF (wFlags AND TYPEFLAG_FDISPATCHABLE) = TYPEFLAG_FDISPATCHABLE THEN strFlags += " [Dispatchable]"
   IF (wFlags AND TYPEFLAG_FREVERSEBIND) = TYPEFLAG_FREVERSEBIND THEN strFlags += " [Reversebind]"
   IF (wFlags AND TYPEFLAG_FPROXY) = TYPEFLAG_FPROXY THEN strFlags += " [Proxy]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts ImplTypeFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_ImplTypeFlagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN FUNCTION = " [None]" : EXIT FUNCTION
   IF (wFlags AND IMPLTYPEFLAG_FDEFAULT) = IMPLTYPEFLAG_FDEFAULT THEN strFlags += " [Default]"
   IF (wFlags AND IMPLTYPEFLAG_FSOURCE) = IMPLTYPEFLAG_FSOURCE THEN strFlags += " [Source]"
   IF (wFlags AND IMPLTYPEFLAG_FRESTRICTED) = IMPLTYPEFLAG_FRESTRICTED THEN strFlags += " [Restricted]"
   IF (wFlags AND IMPLTYPEFLAG_FDEFAULTVTABLE) = IMPLTYPEFLAG_FDEFAULTVTABLE THEN strFlags += " [Default VTable]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts FuncFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_FuncFlagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN FUNCTION = " [None]" : EXIT FUNCTION
   IF (wFlags AND FUNCFLAG_FRESTRICTED) = FUNCFLAG_FRESTRICTED THEN strFlags += " [Restricted]"
   IF (wFlags AND FUNCFLAG_FSOURCE) = FUNCFLAG_FSOURCE THEN strFlags += " [Source]"
   IF (wFlags AND FUNCFLAG_FBINDABLE) = FUNCFLAG_FBINDABLE THEN strFlags += " [Bindable]"
   IF (wFlags AND FUNCFLAG_FREQUESTEDIT) = FUNCFLAG_FREQUESTEDIT THEN strFlags += " [RequestEdit]"
   IF (wFlags AND FUNCFLAG_FDISPLAYBIND) = FUNCFLAG_FDISPLAYBIND THEN strFlags += " [DisplayBind]"
   IF (wFlags AND FUNCFLAG_FDEFAULTBIND) = FUNCFLAG_FDEFAULTBIND THEN strFlags += " [DefaultBind]"
   IF (wFlags AND FUNCFLAG_FHIDDEN) = FUNCFLAG_FHIDDEN THEN strFlags += " [Hidden]"
   IF (wFlags AND FUNCFLAG_FUSESGETLASTERROR) = FUNCFLAG_FUSESGETLASTERROR THEN strFlags += " [UsesGetLastError]"
   IF (wFlags AND FUNCFLAG_FDEFAULTCOLLELEM) = FUNCFLAG_FDEFAULTCOLLELEM THEN strFlags += " [DefaultCollELem]"
   IF (wFlags AND FUNCFLAG_FUIDEFAULT) = FUNCFLAG_FUIDEFAULT THEN strFlags += " [UserInterfaceDefault]"
   IF (wFlags AND FUNCFLAG_FNONBROWSABLE) = FUNCFLAG_FNONBROWSABLE THEN strFlags += " [NonBrowsable]"
   IF (wFlags AND FUNCFLAG_FREPLACEABLE) = FUNCFLAG_FREPLACEABLE THEN strFlags += " [Replaceable]"
   IF (wFlags AND FUNCFLAG_FIMMEDIATEBIND) = FUNCFLAG_FIMMEDIATEBIND THEN strFlags += " [InmediateBind]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts ParamFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_ParamflagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN FUNCTION = " [None]" : EXIT FUNCTION
   IF (wFlags AND PARAMFLAG_FOPT) = PARAMFLAG_FOPT THEN strFlags += " [opt]"
   IF (wFlags AND PARAMFLAG_FRETVAL) = PARAMFLAG_FRETVAL THEN strFlags += " [retval]"
   IF (wFlags AND PARAMFLAG_FIN) = PARAMFLAG_FIN THEN strFlags += " [in]"
   IF (wFlags AND PARAMFLAG_FOUT) = PARAMFLAG_FOUT THEN strFlags += " [out]"
   IF (wFlags AND PARAMFLAG_FLCID) = PARAMFLAG_FLCID THEN strFlags += " [lcid]"
   IF (wFlags AND PARAMFLAG_FHASDEFAULT) = PARAMFLAG_FHASDEFAULT THEN strFlags += " [hasdefault]"
   IF (wFlags AND PARAMFLAG_FHASCUSTDATA) = PARAMFLAG_FHASCUSTDATA THEN strFlags += " [hascustdata]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts VarFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_VarFlagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN strFlags = " [None]"
   IF (wFlags AND VARFLAG_FREADONLY) = VARFLAG_FREADONLY THEN strFlags += " [ReadOnly]"
   IF (wFlags AND VARFLAG_FSOURCE) = VARFLAG_FSOURCE THEN strFlags += " [Source]"
   IF (wFlags AND VARFLAG_FBINDABLE) = VARFLAG_FBINDABLE THEN strFlags += " [Bindable]"
   IF (wFlags AND VARFLAG_FREQUESTEDIT) = VARFLAG_FREQUESTEDIT THEN strFlags += " [RequestEdit]"
   IF (wFlags AND VARFLAG_FDISPLAYBIND) = VARFLAG_FDISPLAYBIND THEN strFlags += " [DisplayBind]"
   IF (wFlags AND VARFLAG_FDEFAULTBIND) = VARFLAG_FDEFAULTBIND THEN strFlags += " [DefaultBind]"
   IF (wFlags AND VARFLAG_FHIDDEN) = VARFLAG_FHIDDEN THEN strFlags += " [Hidden]"
   IF (wFlags AND VARFLAG_FRESTRICTED) = VARFLAG_FRESTRICTED THEN strFlags += " [Restricted]"
   IF (wFlags AND VARFLAG_FDEFAULTCOLLELEM) = VARFLAG_FDEFAULTCOLLELEM THEN strFlags += " [DefaultCollElem]"
   IF (wFlags AND VARFLAG_FUIDEFAULT) = VARFLAG_FUIDEFAULT THEN strFlags += " [User interface default]"
   IF (wFlags AND VARFLAG_FNONBROWSABLE) = VARFLAG_FNONBROWSABLE THEN strFlags += " [NoBrowsable]"
   IF (wFlags AND VARFLAG_FREPLACEABLE) = VARFLAG_FREPLACEABLE THEN strFlags += " [Replaceable]"
   IF (wFlags AND VARFLAG_FIMMEDIATEBIND) = VARFLAG_FIMMEDIATEBIND THEN strFlags += " [ImmediateBind]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a type kind to a descriptive string.
' ========================================================================================
FUNCTION TLB_TypeKindToStr (BYVAL dwTypeKind AS DWORD) AS STRING
   DIM strType AS STRING
   SELECT CASE dwTypeKind
      CASE TKIND_ENUM      : strType = "TKIND_ENUM"
      CASE TKIND_RECORD    : strType = "TKIND_RECORD"
      CASE TKIND_MODULE    : strType = "TKIND_MODULE"
      CASE TKIND_INTERFACE : strType = "TKIND_INTERFACE"
      CASE TKIND_DISPATCH  : strType = "TKIND_DISPATCH"
      CASE TKIND_COCLASS   : strType = "TKIND_COCLASS"
      CASE TKIND_ALIAS     : strType = "TKIND_ALIAS"
      CASE TKIND_UNION     : strType = "TKIND_UNION"
   END SELECT
   FUNCTION = strType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the VarType.
' ========================================================================================
FUNCTION TLB_VarTypeToStr (BYVAL VarType AS LONG, BYVAL fReturnType AS LONG = 0) AS STRING
   DIM s AS STRING
   SELECT CASE VarType
      CASE     0 : s = "VT_EMPTY"
      CASE     1 : s = "VT_NULL"
      CASE     2 : s = "VT_I2 <Short>"
      CASE     3 : s = "VT_I4 <Long>"
      CASE     4 : s = "VT_R4 <Single>"
      CASE     5 : s = "VT_R8 <Double>"
      CASE     6 : s = "VT_CY <CY>"
      CASE     7 : s = "VT_DATE <DATE_>"
      CASE     8 : s = "VT_BSTR <BSTR>"
      CASE     9 : s = "VT_DISPATCH <IDispatch>"
      CASE    10 : s = "VT_ERROR <SCode>"
      CASE    11 : s = "VT_BOOL <VARIANT_BOOL>"
      CASE    12 : s = "VT_VARIANT <Variant>"
      CASE    13 : s = "VT_UNKNOWN <IUnknown>"
      CASE    14 : s = "VT_DECIMAL <DECIMAL>"
      CASE    16 : s = "VT_I1 <Byte>"
      CASE    17 : s = "VT_UI1 <UByte>"
      CASE    18 : s = "VT_UI2 <Short>"
      CASE    19 : s = "VT_UI4 <Ulong>"
      CASE    20 : s = "VT_I8 <LongInt>"
      CASE    21 : s = "VT_UI8 <ULongInt>"
      CASE    22 : s = "VT_INT <Int_>"
      CASE    23 : s = "VT_UINT <Uint>"
      CASE    24 :
         IF fReturnType THEN
            s = "VT_VOID <void>"
         ELSE
            s = "VT_VOID <void>"
         END IF
      CASE    25 : s = "VT_HRESULT <HRESULT>"
      CASE    26 : s = "VT_PTR <PTR>"
      CASE    27 : s = "VT_SAFEARRAY <SAFEARRAY>"
      CASE    28 : s = "VT_CARRAY"
      CASE    29 : s = "VT_USERDEFINED"
      CASE    30 : s = "VT_LPSTR"
      CASE    31 : s = "VT_LPWSTR"
      CASE    36 : s = "VT_RECORD"
      CASE    64 : s = "VT_FILETIME <FILETIME>"
      CASE    65 : s = "VT_BLOB <BLOB>"
      CASE    66 : s = "VT_STREAM <IStream PTR>"
      CASE    67 : s = "VT_STORAGE <IStorage PTR>"
      CASE    68 : s = "VT_STREAMED_OBJECT"
      CASE    69 : s = "VT_STORED_OBJECT"
      CASE    70 : s = "VT_BLOB_OBJECT"
      CASE    71 : s = "VT_CF"
      CASE    72 : s = "VT_CLSID <Guid>"
      CASE  4096 : s = "VT_VECTOR"
      CASE  8192 : s = "VT_ARRAY"
      CASE 16384 : s = "VT_BYREF"
      CASE 32768 : s = "VT_RESERVED"
   END SELECT
   FUNCTION = s
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the VarType
' ========================================================================================
FUNCTION TLB_VarTypeToConstant (BYVAL VarType AS LONG) AS STRING
   DIM s AS STRING
   SELECT CASE VarType
      CASE     0 : s = "VT_EMPTY"
      CASE     1 : s = "VT_NULL"
      CASE     2 : s = "VT_I2"
      CASE     3 : s = "VT_I4"
      CASE     4 : s = "VT_R4"
      CASE     5 : s = "VT_R8"
      CASE     6 : s = "VT_CY"
      CASE     7 : s = "VT_DATE"
      CASE     8 : s = "VT_BSTR"
      CASE     9 : s = "VT_DISPATCH"
      CASE    10 : s = "VT_ERROR"
      CASE    11 : s = "VT_BOOL"
      CASE    12 : s = "VT_VARIANT"
      CASE    13 : s = "VT_UNKNOWN"
      CASE    14 : s = "VT_DECIMAL"
      CASE    16 : s = "VT_I1"
      CASE    17 : s = "VT_UI1"
      CASE    18 : s = "VT_UI2"
      CASE    19 : s = "VT_UI4"
      CASE    20 : s = "VT_I8"
      CASE    21 : s = "VT_UI8"
      CASE    22 : s = "VT_INT"
      CASE    23 : s = "VT_UINT"
      CASE    24 : s = "VT_VOID"
      CASE    25 : s = "VT_HRESULT"
      CASE    26 : s = "VT_PTR"
      CASE    27 : s = "VT_SAFEARRAY"
      CASE    28 : s = "VT_CARRAY"
      CASE    29 : s = "VT_USERDEFINED"
      CASE    30 : s = "VT_LPSTR"
      CASE    31 : s = "VT_LPWSTR"
      CASE    36 : s = "VT_RECORD"
      CASE    64 : s = "VT_FILETIME"
      CASE    65 : s = "VT_BLOB"
      CASE    66 : s = "VT_STREAM"
      CASE    67 : s = "VT_STORAGE"
      CASE    68 : s = "VT_STREAMED_OBJECT"
      CASE    69 : s = "VT_STORED_OBJECT"
      CASE    70 : s = "VT_BLOB_OBJECT"
      CASE    71 : s = "VT_CF"
      CASE    72 : s = "VT_CLSID"
      CASE  4096 : s = "VT_VECTOR"
      CASE  8192 : s = "VT_ARRAY"
      CASE 16384 : s = "VT_BYREF"
      CASE 32768 : s = "VT_RESERVED"
   END SELECT
   FUNCTION = s
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the VarType as a keyword
' ========================================================================================
FUNCTION TLB_VarTypeToKeyword OVERLOAD (BYVAL VarType AS LONG, BYVAL cElements AS WORD = 0) AS STRING
   ' Note: VT_I1 is an array of bytes; translate it to a fixed string
   DIM s AS STRING
   SELECT CASE VarType
      CASE     0 : s = "VOID"                              ' VT_EMPTY
      CASE     1 : s = "VOID"                              ' VT_NULL
      CASE     2 : s = "SHORT"                             ' VT_I2
      CASE     3 : s = "LONG"                              ' VT_I4
      CASE     4 : s = "SINGLE"                            ' VT_R4
      CASE     5 : s = "DOUBLE"                            ' VT_R8
      CASE     6 : s = "CY"                                ' VT_CY
      CASE     7 : s = "DATE_"                             ' VT_DATE
      CASE     8 : s = "BSTR"                              ' VT_BSTR
      CASE     9 : s = "IDispatch"                         ' VT_DISPATCH
      CASE    10 : s = "SCODE"                             ' VT_ERROR
      CASE    11 : s = "VARIANT_BOOL"                      ' VT_BOOL
      CASE    12 : s = "VARIANT"                           ' VT_VARIANT
      CASE    13 : s = "IUnknown"                          ' VT_UNKNOWN
      CASE    14 : s = "DECIMAL"                           ' VT_DECIMAL
      CASE 16, 17                                          ' VT_I1, VT_UI1
'         IF cElements THEN
'            s = "WSTRING * " & STR(cElements)              ' Byte array
'         ELSE
'            s = "BYTE"
'         END IF
         IF cElements THEN
            s = "(0 TO " & STR(cElements) & " AS " & IIF&(VarType = 16, "BYTE", "UBYTE") & ")"
         ELSE
            s = IIF&(VarType = 16, "BYTE", "UBYTE")
         END IF
      CASE    18 : s = "USHORT"                            ' VT_UI2
      CASE    19 : s = "ULONG"                             ' VT_UI4
      CASE    20 : s = "LONGINT"                           ' VT_I8
      CASE    21 : s = "ULONGINT"                          ' VT_UI8
      CASE    22 : s = "INT_"                              ' VT_INT
      CASE    23 : s = "UINT"                              ' VT_UINT
      CASE    24 : s = "VOID"                              ' VT_VOID
      CASE    25 : s = "HRESULT"                           ' VT_HRESULT
      CASE    26 : s = "PTR"                               ' VT_PTR
      CASE    27 : s = "SAFEARRAY"                         ' VT_SAFEARRAY
      CASE    28 : s = "VOID"                              ' VT_CARRAY
      CASE    29 : s = "VOID"                              ' VT_USERDEFINED
      CASE    30 : s = "ZTRING"                            ' VT_LPSTR
      CASE    31 : s = "WSTRING"                           ' VT_LPWSTR
      CASE    36 : s = "VOID"                              ' VT_RECORD
      CASE    64 : s = "FILETIME"                          ' VT_FILETIME
      CASE    65 : s = "BLOB"                              ' VT_BLOB
      CASE    66 : s = "IStream"                           ' VT_STREAM
      CASE    67 : s = "IStorage"                          ' VT_STORAGE
      CASE    68 : s = "VOID"                              ' VT_STREAMED_OBJECT
      CASE    69 : s = "VOID"                              ' VT_STORED_OBJECT
      CASE    70 : s = "VOID"                              ' VT_BLOB_OBJECT
      CASE    71 : s = "VOID"                              ' VT_CF
      CASE    72 : s = "CLSID"                             ' VT_CLSID
      CASE  4096 : s = "VOID"                              ' VT_VECTOR
      CASE  8192 : s = "VOID"                              ' VT_ARRAY
      CASE 16384 : s = "VT_BYREF"
      CASE 32768 : s = "VT_RESERVED"
      CASE ELSE
         s = "VOID"
   END SELECT
   FUNCTION = s
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the VarType
' ========================================================================================
FUNCTION TLB_VarTypeToKeyword OVERLOAD (BYVAL VarType AS STRING) AS STRING
   DIM s AS STRING
   SELECT CASE VarType
      CASE "VT_EMPTY"           : s = "VOID"
      CASE "VT_NULL"            : s = "VOID"
      CASE "VT_I1"              : s = "BYTE"
      CASE "VT_UI1"             : s = "UBYTE"
      CASE "VT_I2"              : s = "SHORT"
      CASE "VT_UI2"             : s = "USHORT"
      CASE "VT_I4"              : s = "LONG"
      CASE "VT_UI4"             : s = "ULONG"
      CASE "VT_I8"              : s = "LONGINT"
      CASE "VT_UI8"             : s = "ULONGINT"
      CASE "VT_INT"             : s = "INT_"
      CASE "VT_UINT"            : s = "UINT"
      CASE "VT_R4"              : s = "SINGLE"
      CASE "VT_R8"              : s = "DOUBLE"
      CASE "VT_CY"              : s = "CY"
      CASE "VT_DATE"            : s = "DATE_"
      CASE "VT_BSTR"            : s = "BSTR"
      CASE "VT_UNKNOWN"         : s = "IUnknown"
      CASE "VT_DISPATCH"        : s = "IDispatch"
      CASE "VT_ERROR"           : s = "VOID"
      CASE "VT_BOOL"            : s = "BOOL"
      CASE "VT_VARIANT"         : s = "VARIANT"
      CASE "VT_DECIMAL"         : s = "DECIMAL"
      CASE "VT_VOID"            : s = "VOID"
      CASE "VT_HRESULT"         : s = "HRESULT"
      CASE "VT_PTR"             : s = "PTR"
      CASE "VT_SAFEARRAY"       : s = "SAFEARRAY"
      CASE "VT_CARRAY"          : s = "VOID"
      CASE "VT_USERDEFINED"     : s = "VOID"
      CASE "VT_LPSTR"           : s = "ZSTRING PTR"
      CASE "VT_LPWSTR"          : s = "WSTRING PTR"
      CASE "VT_RECORD"          : s = "VOID"
      CASE "VT_FILETIME"        : s = "FILETIME"
      CASE "VT_BLOB"            : s = "VOID"
      CASE "VT_STREAM"          : s = "IStream"
      CASE "VT_STORAGE"         : s = "IStorage"
      CASE "VT_STREAMED_OBJECT" : s = "VOID"
      CASE "VT_STORED_OBJECT"   : s = "VOID"
      CASE "VT_BLOB_OBJECT"     : s = "VOID"
      CASE "VT_CF"              : s = "VOID"
      CASE "VT_CLSID"           : s = "CLSID"
      CASE "VT_VECTOR"          : s = "VOID"
      CASE "VT_ARRAY"           : s = "VOID"
      CASE "VT_BYREF"           : s = "VOID"
      CASE "VT_RESERVED"        : s = "VOID"
      CASE ELSE
         s = VarType
   END SELECT
   FUNCTION = s
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the appropiate member name of the variant union for byref parameters.
' Note: VT_HRESULT isn't an automation compatible type, but the CreatePartnershipComplete
' event of Windows Media Player has a parameter of this type.
' ========================================================================================
FUNCTION TLB_GetUnionMemberName (BYVAL vt AS LONG) AS STRING
   DIM strvt AS STRING
   SELECT CASE vt
      CASE VT_I1, VT_UI1 : strvt = "pbVal"
      CASE VT_I2 : strvt = "piVal"
      CASE VT_I4, VT_INT, VT_UI4, VT_UINT, VT_HRESULT : strvt = "plVal"
      CASE VT_R4 : strvt = "pfltVal"
      CASE VT_R8, VT_I8, VT_UI8 : strvt = "pdblVal"
      CASE VT_BOOL : strvt = "pboolVal"
      CASE VT_ERROR : strvt = "pscode"
      CASE VT_CY : strvt = "pcyVal"
      CASE VT_DATE : strvt = "pdate"
      CASE VT_BSTR : strvt = "pbstrVal"
      CASE VT_UNKNOWN : strvt = "ppunkVal"
      CASE VT_DISPATCH : strvt = "ppdispVal"
      CASE VT_ARRAY : strvt = "psArray"
      CASE VT_VARIANT : strvt = "pVariant"
      CASE ELSE : strvt = "plVal"
   END SELECT
   FUNCTION = strvt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the ProgID from the registry.
' ========================================================================================
FUNCTION TLB_GetProgID (BYVAL pwszGuid AS WSTRING PTR) AS CWSTR

   ' // Name of the subkey to open
   DIM wszKey AS WSTRING * MAX_PATH = "CLSID\" & *pwszGuid & "\ProgID"
   DIM hKey AS HKEY   ' // Handle of the opened key
   RegOpenKeyExW HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey
   IF hKey = NULL THEN RETURN ""
   DIM dwIdx AS DWORD = 0                   ' // Index of the value to be retrieved
   DIM cValueName AS DWORD = MAX_PATH       ' // Size of wszValueName
   DIM cbData AS DWORD = MAX_PATH           ' // Size of wszKeyValue
   DIM keyType AS DWORD                     ' // Type of data
   DIM wszKeyValue AS WSTRING * MAX_PATH    ' // Buffer that receives the data
   DIM wszValueName AS WSTRING * MAX_PATH   ' // Name of the value
   RegEnumValueW hKey, dwIdx, @wszValueName, @cValueName, NULL, @keyType, cast(BYTE PTR, @wszKeyValue), @cbData
   RegCloseKey hKey
   RETURN wszKeyValue

END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the Version Independent ProgID from the registry.
' ========================================================================================
FUNCTION TLB_GetVersionIndependentProgID (BYVAL pwszGuid AS WSTRING PTR) AS CWSTR

   ' // Name of the subkey to open
   DIM wszKey AS WSTRING * MAX_PATH = "CLSID\" & *pwszGuid & "\VersionIndependentProgID"
   DIM hKey AS HKEY   ' // Handle of the opened key
   RegOpenKeyExW HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey
   IF hKey = NULL THEN RETURN ""
   DIM dwIdx AS DWORD = 0                   ' // Index of the value to be retrieved
   DIM cValueName AS DWORD = MAX_PATH       ' // Size of wszValueName
   DIM cbData AS DWORD = MAX_PATH           ' // Size of wszKeyValue
   DIM keyType AS DWORD                     ' // Type of data
   DIM wszKeyValue AS WSTRING * MAX_PATH    ' // Buffer that receives the data
   DIM wszValueName AS WSTRING * MAX_PATH   ' // Name of the value
   RegEnumValueW hKey, dwIdx, @wszValueName, @cValueName, NULL, @keyType, cast(BYTE PTR, @wszKeyValue), @cbData
   RegCloseKey hKey
   RETURN wszKeyValue

END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the InprocServer32 from the registry.
' ========================================================================================
FUNCTION TLB_GetInprocServer32 (BYVAL pwszGuid AS WSTRING PTR) AS CWSTR

   ' // Name of the subkey to open
   DIM wszKey AS WSTRING * MAX_PATH = "CLSID\" & *pwszGuid & "\InprocServer32"
   DIM hKey AS HKEY   ' // Handle of the opened key
   RegOpenKeyExW HKEY_CLASSES_ROOT, @wszKey , 0, KEY_READ, @hKey
   IF hKey = NULL THEN RETURN ""
   DIM dwIdx AS DWORD = 0                   ' // Index of the value to be retrieved
   DIM cValueName AS DWORD = MAX_PATH       ' // Size of wszValueName
   DIM cbData AS DWORD = MAX_PATH           ' // Size of wszKeyValue
   DIM keyType AS DWORD                     ' // Type of data
   DIM wszKeyValue AS WSTRING * MAX_PATH    ' // Buffer that receives the data
   DIM wszValueName AS WSTRING * MAX_PATH   ' // Name of the value
   RegEnumValueW hKey, dwIdx, @wszValueName, @cValueName, NULL, @keyType, cast(BYTE PTR, @wszKeyValue), @cbData
   RegCloseKey hKey
   RETURN wszKeyValue

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the implemented interface.
' ========================================================================================
FUNCTION TLB_GetImplementedInterface (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL idx AS LONG = 0) AS CWSTR

   DIM pRefType AS HREFTYPE   ' // Address to a referenced type description
   DIM hr AS HRESULT = pTypeInfo->GetRefTypeOfImplType(idx, @pRefType)
   IF hr <> S_OK OR pRefType = NULL THEN RETURN ""
   DIM pImplTypeInfo AS Afx_ITypeInfo PTR   ' // Implemented interface type info
   hr = pTypeInfo->GetRefTypeInfo(pRefType, @pImplTypeInfo)
   IF hr <> S_OK OR pImplTypeInfo = NULL THEN RETURN ""
   DIM cwsName AS CWSTR, bstrName AS AFX_BSTR   ' // interface name
   pImplTypeInfo->GetDocumentation(-1, @bstrName, NULL, NULL, NULL)
   pImplTypeInfo->Release
   cwsName = *bstrName : SysFreeString bstrName
   RETURN cwsName

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the inherited interface
' ========================================================================================
FUNCTION TLB_GetInheritedInterface (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL idx AS LONG = 0) AS CWSTR

   DIM pRefType AS HREFTYPE   ' // Address to a referenced type description
   DIM hr AS HRESULT = pTypeInfo->GetRefTypeOfImplType(idx, @pRefType)
   IF hr <> S_OK OR pRefType = NULL THEN RETURN ""
   DIM pImplTypeInfo AS Afx_ITypeInfo PTR   ' // Implied interface type info
   hr = pTypeInfo->GetRefTypeInfo (pRefType, @pImplTypeInfo)
   IF hr <> S_OK OR pImplTypeInfo = NULL THEN RETURN ""
   DIM pTypeAttr AS TYPEATTR PTR   ' // Address of a pointer to the TYPEATTR structure
   hr = pImplTypeInfo->GetTypeAttr(@pTypeAttr)
   DIM cwsInterfaceName AS CWSTR
   IF hr = S_OK AND pTypeAttr <> NULL THEN
      IF @pTypeAttr->cImplTypes = 1 THEN
         cwsInterfaceName = TLB_GetImplementedInterface(pImplTypeInfo, 0)
         pImplTypeInfo->ReleaseTypeAttr(pTypeAttr)
      END IF
   END IF
   pImplTypeInfo->Release
   RETURN cwsInterfaceName

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the base class
' ========================================================================================
FUNCTION TLB_GetBaseClass (BYVAL pTypeLib AS Afx_ITypeLib PTR, BYREF cwsItemName AS CWSTR) AS CWSTR

   DIM pTypeInfo               AS Afx_ITypeInfo PTR   ' // TypeInfo interface
   DIM pTypeAttr               AS TYPEATTR PTR        ' // Address of a pointer to the TYPEATTR structure
   DIM pRefType                AS DWORD               ' // Address to a referenced type description
   DIM pRefTypeInfo            AS Afx_ITypeInfo PTR   ' // Referenced TypeInfo interface
   DIM pRefTypeAttr            AS TYPEATTR PTR        ' // Referenced TYPEATTR structure
   DIM cwsInheritedInterface   AS CWSTR               ' // Inherited interface

   ' // Number of type infos
   DIM TypeInfoCount AS LONG = pTypeLib->GetTypeInfoCount
   IF TypeInfoCount = 0 THEN RETURN ""

   FOR i AS LONG = 0 TO TypeInfoCount - 1
      ' // Get the info type
      DIM pTKind AS TYPEKIND
      DIM hr AS HRESULT = pTypeLib->GetTypeInfoType(i, @pTKind)
      IF hr <> S_OK THEN EXIT FOR
      ' // Get the type info
      hr = pTypeLib->GetTypeInfo(i, @pTypeInfo)
      IF hr <> S_OK THEN EXIT FOR
      ' // Get the type attribute
      hr = pTypeInfo->GetTypeAttr(@pTypeAttr)
      IF hr <> S_OK OR pTypeAttr = NULL THEN EXIT FOR
      ' // If it is an interface...
      IF pTKind = TKIND_INTERFACE OR pTKind = TKIND_DISPATCH THEN
         ' // Get the name of the interface
         DIM cwsName AS CWSTR, bstrName AS AFX_BSTR
         hr = pTypeLib->GetDocumentation(i, @bstrName, NULL, NULL, NULL)
         cwsName = *bstrName : SysFreeString bstrName
         ' // If it is the one we are looking for...
         IF cwsName = cwsItemName THEN
            ' // If it inherits from another interface, recursively search the methods
            IF (pTypeAttr->wTypeFlags AND TYPEFLAG_FDUAL) = TYPEFLAG_FDUAL THEN
               cwsInheritedInterface = TLB_GetInheritedInterface(pTypeInfo, -1)
            ELSE
               cwsInheritedInterface = TLB_GetImplementedInterface(pTypeInfo)
            END IF
            ' // Check also that the interface doesn't inherit from itself!
            IF UCASE(cwsInheritedInterface) <> "IUNKNOWN" AND UCASE(cwsInheritedInterface) <> "IDISPATCH" AND UCASE(cwsInheritedInterface) <> UCASE(*bstrName) THEN
               cwsInheritedInterface = TLB_GetBaseClass(pTypeLib, cwsInheritedInterface)
            END IF
         END IF
      END IF
      pTypeInfo->ReleaseTypeAttr(pTypeAttr) : pTypeAttr = NULL
      pTypeInfo->Release : pTypeInfo = NULL
   NEXT

   IF pTypeAttr THEN pTypeInfo->ReleaseTypeAttr(pTypeAttr)
   IF pTypeInfo THEN pTypeInfo->Release

   RETURN cwsInheritedInterface

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the license key for licensed controls.
' Note: As CoGetClassObject creates an unitialized instance of the server to access the
' IClassFactory2 interface, it can take some time with remote servers such as WSCRIPT.EXE.
' ========================================================================================
'FUNCTION TLB_GetRuntimeLicenseKey (BYVAL ProgramID AS WSTRING PTR) AS WSTRING

'   DIM hr                 AS HRESULT              ' // HRESULT
'   DIM pIClassFactory2    AS IClassFactory2 PTR   ' // IClassFactory2 interface
'   DIM IID_IClassFactory2 AS GUID                 ' // IClassFactory2 GUID
'   DIM ClassClsid         AS GUID                 ' // CLSID
'   DIM tLicInfo           AS LICINFO              ' // LICINFO structure
'   DIM cbstrLicKey        AS WSTRING              ' // License key

'   ' // Retrieve the CLSID associated with the PROGID of the component
'   ClassClsid = CLSID$(ProgramID)

'   ' // Get a pointer to the IClassFactory2 interface
'   IID_IClassFactory2 = $IID_ICLASSFACTORY2
'   hr = CoGetClassObject(ClassClsid, CLSCTX_ALL, BYVAL NULL, IID_IClassFactory2, pIClassFactory2)
'   IF hr <> S_OK THEN EXIT FUNCTION

'   ' // Fill the LICINFO structure
'   tLicInfo.cbLicInfo = SIZEOF(tLicInfo)
'   hr = pIClassFactory2.GetLicInfo(tLicInfo)

'   ' // If there is a runtime key available retrieve it
'   IF hr = S_OK THEN
'      IF tLicInfo.fRuntimeKeyAvail THEN
'         pIClassFactory2.RequestLicKey(0, bstrLicKey)
'         FUNCTION = bstrLicKey
'      END IF
'   END IF

'   ' // Release the interface
'   pIClassFactory2 = NOTHING

'END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves the window placement
' ========================================================================================
FUNCTION TLB_SaveWindowPlacement (BYVAL hwnd AS HWND) AS BOOLEAN
   ' // Get the current window placement
   DIM WinPla AS WINDOWPLACEMENT
   WinPla.Length = SIZEOF(WinPla)
   GetWindowPlacement(hwnd, @WinPla)
   ' // Save the values in the .ini file
   DIM bRes AS BOOLEAN, wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   bRes = WritePrivateProfileStringW ("Window placement", "Left", WSTR(WinPla.rcNormalPosition.Left), @wszIniFileName)
   bRes = WritePrivateProfileStringW ("Window placement", "Right", WSTR(WinPla.rcNormalPosition.Right), @wszIniFileName)
   bRes = WritePrivateProfileStringW ("Window placement", "Top", WSTR(WinPla.rcNormalPosition.Top), @wszIniFileName)
   bRes = WritePrivateProfileStringW ("Window placement", "Bottom", WSTR(WinPla.rcNormalPosition.Bottom), @wszIniFileName)
   RETURN bRes
END FUNCTION
' ========================================================================================

' =====================================================================================
' Window placement
' =====================================================================================
SUB TLB_SetWindowPlacement (BYVAL hwnd AS HWND)
   ' // Get the current window placement
   DIM WinPla AS WINDOWPLACEMENT
   WinPla.Length = SIZEOF(WinPla)
   GetWindowPlacement(hwnd, @WinPla)
   ' // Read the values saved in the .ini file
   DIM wszBuffer AS WSTRING * 260, dwChars AS DWORD, wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   dwChars = GetPrivateProfileStringW("Window placement", "Left", NULL, @wszBuffer, MAX_PATH, @wszIniFileName)
   IF dwChars THEN WinPla.rcNormalPosition.Left = VAL(wszBuffer)
   dwChars = GetPrivateProfileStringW("Window placement", "Right", NULL, @wszBuffer, MAX_PATH, @wszIniFileName)
   IF dwChars THEN WinPla.rcNormalPosition.Right = VAL(wszBuffer)
   IF WinPla.rcNormalPosition.Right = 0 THEN WinPla.rcNormalPosition.Right = 750
   dwChars = GetPrivateProfileStringW("Window placement", "Top", NULL, @wszBuffer, MAX_PATH, @wszIniFileName)
   IF dwChars THEN WinPla.rcNormalPosition.Top = VAL(wszBuffer)
   dwChars = GetPrivateProfileStringW("Window placement", "Bottom", NULL, @wszBuffer, MAX_PATH, @wszIniFileName)
   IF dwChars THEN WinPla.rcNormalPosition.Bottom = VAL(wszBuffer)
   IF WinPla.rcNormalPosition.Bottom = 0 THEN WinPla.rcNormalPosition.Bottom = 450
   SetWindowPlacement hwnd, @WinPla
   ' // Desktop's real size
   DIM rcDeskTop AS RECT
   SystemParametersInfo SPI_GETWORKAREA, 0, @rcDesktop, 0
   ' // Make sure dialog is not horizontally oversized
   IF WinPla.rcNormalPosition.Right - WinPla.rcNormalPosition.Left > rcDesktop.Right THEN
      WinPla.rcNormalPosition.Left = 0
      WinPla.rcNormalPosition.Right = rcDesktop.Right
   END IF
   ' // Make sure dialog is not vertically oversized
   IF WinPla.rcNormalPosition.Bottom - WinPla.rcNormalPosition.Top > rcDesktop.Bottom THEN
      WinPla.rcNormalPosition.Top = 0
      WinPla.rcNormalPosition.Bottom = rcDesktop.Bottom
   END IF
   ' // Make sure left side of dialog is visible
   IF WinPla.rcNormalPosition.Left < 0 THEN
      WinPla.rcNormalPosition.Right = WinPla.rcNormalPosition.Right - WinPla.rcNormalPosition.Left
      WinPla.rcNormalPosition.Left = 0
   END IF
   ' // Make sure right side of dialog is visible
   IF WinPla.rcNormalPosition.Right > rcDesktop.Right THEN
      WinPla.rcNormalPosition.Left = WinPla.rcNormalPosition.Left - (WinPla.rcNormalPosition.Right - rcDesktop.Right)
      WinPla.rcNormalPosition.Right = rcDesktop.Right
   END IF
   ' // Make sure top side of dialog is visible
   IF WinPla.rcNormalPosition.Top < 0 THEN
      WinPla.rcNormalPosition.Bottom = WinPla.rcNormalPosition.Bottom - WinPla.rcNormalPosition.Top
      WinPla.rcNormalPosition.Top = 0
   END IF
   ' // Make sure bottom side of dialog is visible
   IF WinPla.rcNormalPosition.Bottom > rcDesktop.Bottom THEN
      WinPla.rcNormalPosition.Top = WinPla.rcNormalPosition.Top - (WinPla.rcNormalPosition.Bottom - rcDesktop.Bottom)
      WinPla.rcNormalPosition.Bottom = rcDesktop.Bottom
   END IF
   ' // Adjust the working area
   DIM rc AS RECT
   SystemParametersInfo SPI_GETWORKAREA, 0, @rc, 0
   IF WinPla.rcNormalPosition.Left = rc.Left AND WinPla.rcNormalPosition.Right = rc.Right THEN WinPla.rcNormalPosition.Right = WinPla.rcNormalPosition.Right
   IF WinPla.rcNormalPosition.Top = rc.Top AND WinPla.rcNormalPosition.Bottom = rc.Bottom THEN WinPla.rcNormalPosition.Bottom = WinPla.rcNormalPosition.Bottom
   ' // Place the window
   SetWindowPlacement hwnd, @WinPla
END SUB
' =====================================================================================

' =====================================================================================
' ListView columns order.
' =====================================================================================
FUNCTION TLB_SaveColumnsOrderArray (BYVAL hListView AS HWND) AS BOOLEAN
   DIM nCols AS LONG = ListView_GetColumnCount(hListView)
   REDIM rgLvOrderArray(nCols - 1) AS LONG
   IF ListView_GetColumnOrderArray(hListView, nCols, @rgLvOrderArray(0)) = FALSE THEN RETURN FALSE
   DIM wszBuffer AS WSTRING * MAX_PATH
   FOR idx AS LONG = LBOUND(rgLvOrderArray) TO UBOUND(rgLvOrderArray)
      wszBuffer += WSTR(rgLvOrderArray(idx))
      IF idx < UBOUND(rgLvOrderArray) THEN wszBuffer += ","
   NEXT
   DIM wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   RETURN WritePrivateProfileStringW ("Columns order", "ColumnsOrder", @wszBuffer, @wszIniFileName)
END FUNCTION
' =====================================================================================
' =====================================================================================
SUB TLB_SetColumnsOrderArray (BYVAL hListView AS HWND)
   DIM nCols AS LONG = ListView_GetColumnCount(hListView)
   DIM wszBuffer AS WSTRING * MAX_PATH, wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   DIM dwChars AS DWORD = GetPrivateProfileStringW("Columns order", "ColumnsOrder", NULL, @wszBuffer, MAX_PATH, @wszIniFileName)
   IF LEN(wszBuffer) THEN
      REDIM rgLvOrderArray(nCols - 1) AS LONG
      FOR idx AS LONG = LBOUND(rgLvOrderArray) TO UBOUND(rgLvOrderArray)
         DIM s AS STRING = AfxStrParse(wszBuffer, idx + 1, ",")
         rgLvOrderArray(idx) = VAL(s)
      NEXT
      ListView_SetColumnOrderArray(hListView, nCols, @rgLvOrderArray(0))
   END IF
END SUB
' =====================================================================================

' =====================================================================================
' ListView columns width.
' =====================================================================================
FUNCTION TLB_SaveColumnsWidth (BYVAL hListView AS HWND) AS BOOLEAN
   DIM nCols AS LONG = ListView_GetColumnCount(hListView)
   DIM wszBuffer AS WSTRING * MAX_PATH
   FOR idx AS LONG = 0 TO nCols - 1
      wszBuffer += WSTR(ListView_GetColumnWidth(hListView, idx))
      IF idx < nCols - 1 THEN wszBuffer += ","
   NEXT
   DIM wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   RETURN WritePrivateProfileStringW ("Columns width", "ColumnsWidth", @wszBuffer, @wszIniFileName)
END FUNCTION
' =====================================================================================
' =====================================================================================
SUB TLB_SetColumnsWidth (BYVAL hListView AS HWND)
   DIM nCols AS LONG = ListView_GetColumnCount(hListView)
   DIM wszBuffer AS WSTRING * MAX_PATH, wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   DIM dwChars AS DWORD = GetPrivateProfileStringW("Columns width", "ColumnsWidth", NULL, @wszBuffer, MAX_PATH, @wszIniFileName)
   IF LEN(wszBuffer) THEN
      FOR idx AS LONG = 0 TO nCols - 1
         DIM s AS STRING = AfxStrParse(wszBuffer, idx + 1, ",")
         DIM nWidth AS LONG = VAL(s)
         ListView_SetColumnWidth(hListView, idx, nWidth)
      NEXT
   END IF
END SUB
' =====================================================================================

' =====================================================================================
' ListView selected item
' =====================================================================================
FUNCTION TLB_SaveSelectedItem (BYVAL hListView AS HWND) AS BOOLEAN
   DIM wszSelItem AS WSTRING * MAX_PATH
   DIM nItem AS LONG = ListView_GetSelectionMark(hListView)
   ListView_GetItemText(hListView, nItem, 0, @wszSelItem, MAX_PATH)
   DIM wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   RETURN WritePrivateProfileStringW ("Selected Item", "SelectedItem", @wszSelItem, @wszIniFileName)
END FUNCTION
' =====================================================================================
' =====================================================================================
SUB TLB_SetSelectedItem (BYVAL hListView AS HWND)
   DIM wszSelItem AS WSTRING * MAX_PATH, wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   DIM dwChars AS DWORD = GetPrivateProfileStringW("Selected Item", "SelectedItem", NULL, @wszSelItem, MAX_PATH, @wszIniFileName)
   DIM nItem AS LONG = ListView_FindString(hListView, wszSelItem)
   ListView_SetItemState(hListView, nItem, LVIS_SELECTED OR LVIS_FOCUSED, LVIS_SELECTED OR LVIS_FOCUSED)
   ListView_EnsureVisible(hListView, nItem, CTRUE)
END SUB
' =====================================================================================

' =====================================================================================
' ListView column to sort
' =====================================================================================
FUNCTION TLB_SaveColumnToSort (BYVAL hListView AS HWND, BYVAL ColumnToSort AS LONG) AS BOOLEAN
   DIM wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   RETURN WritePrivateProfileStringW ("Browser options", "ColumnToSort", WSTR(ColumnToSort), @wszIniFileName)
END FUNCTION
' =====================================================================================

' =====================================================================================
' ListView column sort order
' =====================================================================================
FUNCTION TLB_SaveSortOrder (BYVAL hListView AS HWND, BYVAL SortOrder AS LONG) AS BOOLEAN
   DIM wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   RETURN WritePrivateProfileStringW ("Browser options", "SortOrder", WSTR(SortOrder), @wszIniFileName)
END FUNCTION
' =====================================================================================

' =====================================================================================
' ListView column sort order
' =====================================================================================
FUNCTION TLB_GetSortOptions (BYVAL hListView AS HWND, BYVAL ColumnToSort AS LONG PTR, BYVAL SortOrder AS LONG PTR) AS BOOLEAN
   IF ColumnToSort = NULL OR SortOrder = NULL THEN RETURN FALSE
   DIM wszBuffer AS WSTRING * MAX_PATH, wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   DIM dwChars AS DWORD = GetPrivateProfileStringW("Browser options", "ColumnToSort", NULL, @wszBuffer, MAX_PATH, @wszIniFileName)
   IF dwChars = 0 THEN RETURN FALSE
   *ColumnToSort = VAL(wszBuffer)
   wszBuffer = ""
   dwChars = GetPrivateProfileStringW("Browser options", "SortOrder", NULL, @wszBuffer, MAX_PATH, @wszIniFileName)
   IF dwChars = 0 THEN RETURN FALSE
   *SortOrder = VAL(wszBuffer)
   RETURN TRUE
END FUNCTION
' =====================================================================================
