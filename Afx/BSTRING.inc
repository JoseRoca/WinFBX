' ########################################################################################
' Platform: Microsoft Windows
' Filename: DWSTRING.inc
' Purpose:  Implements a data type for dynamic null terminated unicode strings.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2025 José Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the “Software”), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

#pragma once
#include once "DWString.inc"

' ========================================================================================
' Macro for debug
' To allow debugging, define _BSTRING_DEBUG_ 1 in your application before including this file.
' To capture and display the messages sent by the Windows function OutputDebugStringW, you
' can use the DebugView tool. See: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
' ========================================================================================
#ifndef _BSTRING_DEBUG_
   #define _BSTRING_DEBUG_ 0
#ENDIF
#ifndef _BSTRING_DP_
   #define _BSTRING_DP_ 1
   #MACRO BSTRING_DP(st)
      #IF (_BSTRING_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ########################################################################################
'                                *** BSTRING CLASS ***
' ########################################################################################
TYPE BSTRING EXTENDS WSTRING

Public:
   m_bstr AS BSTR   ' // The BSTR handle

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYREF bs AS BSTRING)
   DECLARE CONSTRUCTOR (BYREF dws AS DWSTRING)
   DECLARE CONSTRUCTOR (BYVAL pwszStr AS WSTRING PTR, BYVAL fAttach AS LONG = TRUE)
   DECLARE CONSTRUCTOR (BYREF ansiStr AS STRING = "", BYVAL nCodePage AS UINT = 0)
   DECLARE CONSTRUCTOR (BYVAL n AS LONGINT)
   DECLARE CONSTRUCTOR (BYVAL n AS DOUBLE)
   DECLARE DESTRUCTOR
   DECLARE SUB Append (BYVAL pwszStr AS WSTRING PTR)
   DECLARE SUB Clear
   DECLARE SUB Attach (BYVAL pbstrSrc AS BSTR)
   DECLARE FUNCTION Detach () AS BSTRING
   DECLARE FUNCTION Copy () AS BSTRING
   DECLARE FUNCTION wchar () AS WSTRING PTR
   DECLARE OPERATOR [] (BYVAL nIndex AS UINT) BYREF AS USHORT
   DECLARE OPERATOR CAST () BYREF AS CONST WSTRING
   DECLARE OPERATOR CAST () AS ANY PTR
   DECLARE OPERATOR LET (BYREF bs AS BSTRING)
   DECLARE OPERATOR LET (BYREF dws AS DWSTRING)
   DECLARE OPERATOR LET (BYREF s AS STRING)
   DECLARE OPERATOR LET (BYVAL pwszStr AS WSTRING PTR)
   DECLARE OPERATOR += (BYVAL pwszStr AS WSTRING PTR)
   DECLARE OPERATOR += (BYREF bs AS BSTRING)
   DECLARE OPERATOR += (BYREF dws AS DWSTRING)
   DECLARE OPERATOR += (BYVAL n AS LONGINT)
   DECLARE OPERATOR += (BYVAL n AS DOUBLE)
   DECLARE OPERATOR &= (BYVAL pwszStr AS WSTRING PTR)
   DECLARE OPERATOR &= (BYREF bs AS BSTRING)
   DECLARE OPERATOR &= (BYREF dws AS DWSTRING)
   DECLARE OPERATOR &= (BYVAL n AS LONGINT)
   DECLARE OPERATOR &= (BYVAL n AS DOUBLE)
   DECLARE PROPERTY Utf8 () AS STRING
   DECLARE PROPERTY Utf8 (BYREF utf8String AS STRING)

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
PRIVATE CONSTRUCTOR BSTRING
   BSTRING_DP("BSTRING CONSTRUCTOR - Default")
   ' // Do nothing.
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR BSTRING (BYREF bs AS BSTRING)
   BSTRING_DP("BSTRING CONSTRUCTOR - BSTRING")
   m_bstr = SysAllocString(bs)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR BSTRING (BYREF dws AS DWSTRING)
   BSTRING_DP("BSTRING CONSTRUCTOR - DWSTRING")
   m_bstr = SysAllocString(dws)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR BSTRING (BYVAL pwszStr AS WSTRING PTR, BYVAL fAttach AS LONG = TRUE)
   IF pwszStr = NULL THEN
      m_bstr = SysAllocString("")
   ELSE
      ' Detect if the passed handle is an OLE string.
      ' If it is an OLE string it must have a descriptor; otherwise, don't.
      ' Get the length in bytes looking at the descriptor and divide by 2 to get the number of
      ' unicode characters, that is the value returned by the FreeBASIC LEN operator.
      DIM res AS DWORD = PEEK(DWORD, CAST(ANY PTR, pwszStr) - 4) \ 2
      ' If the retrieved length is the same that the returned by LEN, then it must be an OLE string
      IF res = .LEN(*pwszStr) AND fAttach <> FALSE THEN
         ' Attach the passed handle to the class
         m_bstr = pwszStr
      ELSE
         ' Allocate an OLE string with the contents of the string pointed by pwszStr
         m_bstr = SysAllocString(pwszStr)
      END IF
   END IF
   BSTRING_DP("BSTRING CONSTRUCTOR - SysAllocString - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR BSTRING (BYREF ansiStr AS STRING = "", BYVAL nCodePage AS UINT = 0)
   IF nCodePage = CP_UTF8 THEN
      DIM dwLen AS DWORD = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), .LEN(ansiStr), NULL, 0)
      IF dwLen THEN
         m_bstr = SysAllocString(.WSTR(SPACE(dwLen)))
         MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), .LEN(ansiStr), m_bstr, dwLen * 2)
      ELSE
         m_bstr = SysAllocString("")
      END IF
   ELSE
      IF .LEN(ansiStr) THEN
         m_bstr = SysAllocString(.WSTR(ansiStr))
         MultiByteToWideChar(nCodePage, MB_PRECOMPOSED, STRPTR(ansiStr), -1, m_bstr, .LEN(ansiStr) * 2)
      ELSE
         m_bstr = SysAllocString("")
      END IF
   END IF
   BSTRING_DP("BSTRING CONSTRUCTOR - ansi string - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes the BSTRING with a number.
' These two constructors are needed to allow to use a number with the & operator.
' ========================================================================================
PRIVATE CONSTRUCTOR BSTRING (BYVAL n AS LONGINT)
   m_bstr = SysAllocString(.WSTR(n))
   BSTRING_DP("BSTRING CONSTRUCTOR - LONGINT - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR BSTRING (BYVAL n AS DOUBLE)
   m_bstr = SysAllocString(.WSTR(n))
   BSTRING_DP("BSTRING CONSTRUCTOR - DOUBLE - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR BSTRING
   IF m_bstr THEN SysFreeString m_bstr
   BSTRING_DP("BSTRING DESTRUCTOR - " & .WSTR(m_bstr))
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' * Appends a string to the BSTRING.
' ========================================================================================
PRIVATE SUB BSTRING.Append (BYVAL pwszStr AS WSTRING PTR)
   DIM n1 AS UINT = SysStringLen(m_bstr)
   DIM nLen AS UINT = .LEN(*pwszStr)
   IF nLen = 0 THEN EXIT SUB
   DIM b AS BSTR = SysAllocStringLen(NULL, n1 + nLen)
   IF b = NULL THEN EXIT SUB
   memcpy(b, m_bstr, n1 * SIZEOF(WSTRING))
   memcpy(b + n1, pwszStr, nLen * SIZEOF(WSTRING))
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = b
   BSTRING_DP("BSTRING Append - " & .WSTR(m_bstr))
END SUB
' ========================================================================================

' ========================================================================================
' Frees the m_bstr member.
' ========================================================================================
PRIVATE SUB BSTRING.Clear
   BSTRING_DP("BSTRING Clear - " & .WSTR(m_bstr))
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = NULL
END SUB
' ========================================================================================

' ========================================================================================
' Attaches a BSTR to the BSTRING object by setting the m_bstr member to pbstrSrc.
' The purpose of this method is to attach a BSTR pointer returned by a function that
' returns a BSTR pointer directly. Attaching it ensures that the BSTR will be freed when
' the BSTRING is destroyed. Don't attach a BSTRING to another BSTRING because each BSTRING
' will try to free the same BSTR when they are destroyed.
' ========================================================================================
PRIVATE SUB BSTRING.Attach (BYVAL pbstrSrc AS BSTR)
   BSTRING_DP("BSTRING Attach - " & .WSTR(m_bstr))
   ' // Check if the passed pointer is really a BSTR
   DIM res AS DWORD = PEEK(DWORD, CAST(ANY PTR, pbstrSrc) - 4) \ 2
   ' // If the retrieved length is the same that the returned by LEN, then it must be an OLE string
   IF res = .LEN(*pbstrSrc) THEN
      IF m_bstr THEN SysFreeString(m_bstr)
      m_bstr = pbstrSrc
   END IF
END SUB
' ========================================================================================

' ========================================================================================
' Detaches the BSTR pointer from the BSTRING object.
' ========================================================================================
PRIVATE FUNCTION BSTRING.Detach () AS BSTRING
   BSTRING_DP("BSTRING Detach - " & .WSTR(m_bstr))
   DIM pbstr AS BSTR = m_bstr
   m_bstr = NULL
   RETURN pbstr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allocates and returns a copy of this BSTRING.
' ========================================================================================
PRIVATE FUNCTION BSTRING.Copy () AS BSTRING
   BSTRING_DP("BSTRING Copy - " & .WSTR(m_bstr))
   RETURN SysAllocStringLen(m_bstr, SysStringLen(m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the BSTR pointer
' ========================================================================================
PRIVATE OPERATOR BSTRING.CAST () AS ANY PTR
   BSTRING_DP("BSTRING CAST ANY PTR - " & .WSTR(m_bstr))
   OPERATOR =  CAST(ANY PTR, m_bstr)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the string data (same as **)
' ========================================================================================
PRIVATE OPERATOR BSTRING.CAST () BYREF AS CONST WSTRING
   BSTRING_DP("BSTRING CAST BYREF AS CONST WSTRING - " & .WSTR(m_bstr))
   OPERATOR = *CAST(WSTRING PTR, m_bstr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the corresponding ASCII or Unicode integer representation of the character at
' the zero-based position specified by the nIndex parameter (0 for the first character,
' 1 for the second, etc.), e.g. value = dws[1]. To change the value: dws[1] = value.
' Example:
'   DIM dwsText AS BSTRING = "This is my text."
'   dwsText[1] = ASC("X")
'   PRINT dwsText
' ========================================================================================
PRIVATE OPERATOR BSTRING.[] (BYVAL nIndex AS UINT) BYREF AS USHORT
   DWSTRING_DP("BSTRING Operator []")
   OPERATOR = *CAST(USHORT PTR, CAST(ANY PTR, m_bstr + nIndex))
END OPERATOR
' ========================================================================================

' =====================================================================================
' Returns the contents of the BSTRING as a WSTRING allocated with CoTaskMemAlloc.
' Free the returned string later with CoTaskMemFree.
' Note: This is useful when we need to pass a pointer to a null terminated wide string to a
' function or method that will release it. If we pass a WSTRING it will GPF.
' If the length of the input string is 0, CoTaskMemAlloc allocates a zero-length item and
' returns a valid pointer to that item. If there is insufficient memory available,
' CoTaskMemAlloc returns NULL.
' =====================================================================================
PRIVATE FUNCTION BSTRING.wchar () AS WSTRING PTR
   DIM pwchar AS WSTRING PTR
   DIM nLen AS LONG = SysStringLen(m_bstr) * 2
   pwchar = CoTaskMemAlloc(nLen)
   IF pwchar = NULL THEN RETURN NULL
   IF nLen THEN memcpy pwchar, m_bstr, nLen
   IF nLen = 0 THEN *pwchar = CHR(0)
   BSTRING_DP("BSTRING WChar - " & .WSTR(m_bstr))
   RETURN pwchar
END FUNCTION
' =====================================================================================

' ========================================================================================
' Assigns new text to the BSTRING.
' ========================================================================================
PRIVATE OPERATOR BSTRING.LET (BYREF bs AS BSTRING)
   BSTRING_DP("BSTRING LET - BSTRING " & .WSTR(m_bstr))
   IF m_bstr <> bs.m_bstr THEN   ' // If the user has done cbs = cbs, ignore it
      IF m_bstr THEN SysFreeString(m_bstr)
      m_bstr = SysAllocString(bs)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR BSTRING.LET (BYREF dws AS DWSTRING)
   BSTRING_DP("BSTRING LET - DWSTRING " & .WSTR(m_bstr))
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = SysAllocString(dws)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR BSTRING.LET (BYREF s AS STRING)
   BSTRING_DP("BSTRING LET - STRING " & .WSTR(m_bstr))
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = SysAllocString(s)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR BSTRING.LET (BYVAL pwszStr AS WSTRING PTR)
   BSTRING_DP("BSTRING LET - WSTRING PTR " & .WSTR(pwszStr))
   ' Free the current OLE string
   IF m_bstr THEN SysFreeString(m_bstr)
   ' Detect if the passed handle is an OLE string.
   ' If it is an OLE string it must have a descriptor; otherwise, don't.
   ' Get the length in bytes looking at the descriptor and divide by 2 to get the number of
   ' unicode characters, that is the value returned by the FreeBASIC LEN operator.
   DIM res AS DWORD = PEEK(DWORD, CAST(ANY PTR, pwszStr) - 4) \ 2
   ' If the retrieved length is the same that the returned by LEN, then it must be an OLE string
   IF res = .LEN(*pwszStr) THEN
      ' Attach the passed handle to the class
      m_bstr = pwszStr
   ELSE
      ' Allocate an OLE string with the contents of the string pointed by pwszStr
      m_bstr = SysAllocString(pwszStr)
   END IF
END OPERATOR
' ========================================================================================

' ========================================================================================
' Converts the BSTRING to UTF8.
' ========================================================================================
PRIVATE PROPERTY BSTRING.Utf8 () AS STRING
   BSTRING_DP("BSTRING - TO Utf8")
   DIM cbLen AS INTEGER, bstrLen AS LONG
   IF m_bstr = NULL THEN RETURN ""
   bstrLen = SysStringLen(m_bstr)
   DIM buffer AS STRING = STRING(bstrLen * 5 + 1, 0)
   PROPERTY = *cast(ZSTRING PTR, WCharToUTF(1, cast(WSTRING PTR, m_bstr), bstrLen, STRPTR(buffer), @cbLen))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Converts UTF8 to unicode and assigns it to the BSTRING.
' ========================================================================================
PRIVATE PROPERTY BSTRING.Utf8 (BYREF utf8String AS STRING)
   BSTRING_DP("BSTRING - FROM Utf8")
   IF m_bstr THEN SysFreeString(m_bstr)
   DIM dwLen AS DWORD = MultiByteToWideChar(CP_UTF8, 0, STRPTR(utf8String), .LEN(utf8String), NULL, 0)
   IF dwLen THEN
      m_bstr = SysAllocString(.WSTR(SPACE(dwLen)))
      MultiByteToWideChar(CP_UTF8, 0, STRPTR(utf8String), .LEN(utf8String), m_bstr, dwLen * 2)
   ELSE
      m_bstr = SysAllocString("")
   END IF
END PROPERTY
' ========================================================================================

' ########################################################################################
' The following overloaded operators are needed because the length of the string needs to
' be updated and the intrinsic FreeBasic operators won't do it because they will treat the
' variable as a WSTRING, not as a BSTR.
' ########################################################################################

' ========================================================================================
' Appends a string to the BSTRING.
' ========================================================================================
PRIVATE OPERATOR BSTRING.+= (BYVAL pwszStr AS WSTRING PTR)
   BSTRING_DP("BSTRING += - WSTRING PTR")
   this.Append(pwszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Appends a BSTRING to the BSTRING.
' ========================================================================================
PRIVATE OPERATOR BSTRING.+= (BYREF bs AS BSTRING)
   BSTRING_DP("BSTRING += - BSTRING")
   this.Append(bs)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Appends a DWSTRING to the BSTRING.
' ========================================================================================
PRIVATE OPERATOR BSTRING.+= (BYREF dws AS DWSTRING)
   BSTRING_DP("BSTRING += - DWSTRING")
   this.Append(dws)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a number to the BSTRING.
' ========================================================================================
PRIVATE OPERATOR BSTRING.+= (BYVAL n AS LONGINT)
   BSTRING_DP("BSTRING += - DWSTRING")
   DIM wsz AS WSTRING * 40 = .WSTR(n)
   this.Append(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR BSTRING.+= (BYVAL n AS DOUBLE)
   BSTRING_DP("BSTRING += - DOUBLE")
   DIM wsz AS WSTRING * 40 = .WSTR(n)
   this.Append(wsz)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a string to the BSTRING.
' ========================================================================================
PRIVATE OPERATOR BSTRING.&= (BYVAL pwszStr AS WSTRING PTR)
   BSTRING_DP("BSTRING &= - WSTRING PTR")
   this.Append(pwszStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a BSTRING to the BSTRING.
' ========================================================================================
PRIVATE OPERATOR BSTRING.&= (BYREF bs AS BSTRING)
   BSTRING_DP("BSTRING &= - BSTRING")
   this.Append(bs)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a DWSTRING to the BSTRING.
' ========================================================================================
PRIVATE OPERATOR BSTRING.&= (BYREF dws AS DWSTRING)
   BSTRING_DP("BSTRING &= - DWSTRING")
   this.Append(dws)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a number to the BSTRING.
' ========================================================================================
PRIVATE OPERATOR BSTRING.&= (BYVAL n AS LONGINT)
   BSTRING_DP("BSTRING &= - LONGINT")
   DIM wsz AS WSTRING * 40 = .WSTR(n)
   this.Append(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR BSTRING.&= (BYVAL n AS DOUBLE)
   BSTRING_DP("BSTRING &= - DOUBLE")
   DIM wsz AS WSTRING * 40 = .WSTR(n)
   this.Append(wsz)
END OPERATOR
' ========================================================================================

' ########################################################################################

' ########################################################################################
'                         *** GLOBAL OPERATORS AND FUNCTIONS ***
' ########################################################################################

' ========================================================================================
' Global operator.
' Returns the length, in characters, of the DWSTRING.
' ========================================================================================
PRIVATE OPERATOR LEN (BYREF bs AS BSTRING) AS UINT
   OPERATOR = SysStringLen(bs)
   BSTRING_DP("BSTRING - LEN BSTRING - SysStringLen = " & .WSTR(SysStringLen(bs) & " - " & WSTR(bs.m_bstr)))
END OPERATOR
' ========================================================================================

' ========================================================================================
' Global operator.
' One * returns the BSTR handle.
' Two ** deferences de string data.
' ========================================================================================
PRIVATE OPERATOR * (BYREF bs AS BSTRING) AS WSTRING PTR
   BSTRING_DP("BSTRING - * BSTRING")
   OPERATOR = bs.m_bstr
END OPERATOR
' ========================================================================================

' ========================================================================================
' Concatenates two strings, converting non-strings to strings as needed
' Overloaded because the intrinsic LEFT, RIGHT and MID operators can fail when the string
' parameter is the result of a combination of different data types and literals,
' The operator + doesn't seem to have problems with this.
' ========================================================================================
PRIVATE OPERATOR & (BYREF bs1 AS BSTRING, BYREF bs2 AS BSTRING) AS BSTRING
   BSTRING_DP("BSTRING - & BSTRING")
   OPERATOR = bs1 + bs2
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the leftmost substring of a string
' Overloaded because the intrinsic LEFT operator fails with an ambiguous call error.
' ========================================================================================
PRIVATE FUNCTION Left (BYREF bs AS BSTRING, BYVAL nChars AS INTEGER) AS BSTRING
   BSTRING_DP("LEFT - BSTRING")
   RETURN LEFT(**bs, nChars)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the rightmost substring of a string
' Overloaded because the intrinsic LEFT operator fails with an ambiguous call error.
' ========================================================================================
PRIVATE FUNCTION Right (BYREF bs AS BSTRING, BYVAL nChars AS INTEGER) AS BSTRING
   BSTRING_DP("RIGHT - BSTRING")
   RETURN RIGHT(**bs, nChars)
END FUNCTION
' ========================================================================================

