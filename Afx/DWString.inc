' ########################################################################################
' Platform: Microsoft Windows
' Filename: DWSTRING.bi
' Purpose: Implements a data type for dynamic null terminated unicode strings.
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

' // Include files
#pragma ONCE
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "/crt/string.bi"
#INCLUDE ONCE "/crt/wchar.bi"
#INCLUDE ONCE "utf_conv.bi"

' ========================================================================================
' Macro for debug
' To allow debugging, define _DWSTRING_DEBUG_ 1 in your application before including this file.
' To capture and display the messages sent by the Windows function OutputDebugStringW, you
' can use the DebugView tool. See: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
' ========================================================================================
#ifndef _DWSTRING_DEBUG_
   #define _DWSTRING_DEBUG_ 0
#ENDIF
#ifndef _DWSTRING_DP_
   #define _DWSTRING_DP_ 1
   #MACRO DWSTRING_DP(st)
      #IF (_DWSTRING_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ########################################################################################
'                                *** DWSTRING CLASS ***
' ########################################################################################

TYPE DWSTRING EXTENDS WSTRING

   m_pBuffer AS WSTRING PTR   ' // Pointer to the buffer
   m_BufferLen AS ssize_t     ' // Length in UTF16 characters of the buffer
   m_Capacity AS ssize_t      ' // The total size of the buffer in UTF16 characters

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL nChars AS LONG, BYVAL bClear AS BOOLEAN)
   DECLARE CONSTRUCTOR (BYVAL pwszStr AS WSTRING PTR)
   DECLARE CONSTRUCTOR (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
   DECLARE CONSTRUCTOR (BYREF dws AS DWSTRING)
   DECLARE CONSTRUCTOR (BYVAL n AS LONGINT)
   DECLARE CONSTRUCTOR (BYVAL n AS DOUBLE)
   DECLARE DESTRUCTOR
   DECLARE PROPERTY Capacity () AS UINT
   DECLARE PROPERTY Capacity (BYVAL nValue AS UINT)
   DECLARE FUNCTION ResizeBuffer (BYVAL nChars AS UINT, BYVAL bClear AS BOOLEAN = FALSE) AS WSTRING PTR
   DECLARE FUNCTION AppendBuffer (BYVAL memAddr AS ANY PTR, BYVAL nChars AS UINT) AS BOOLEAN
   DECLARE FUNCTION GetBuffer () AS WSTRING PTR
   DECLARE SUB Clear
   DECLARE FUNCTION Add (BYVAL pwszStr AS WSTRING PTR) AS BOOLEAN
   DECLARE FUNCTION Add (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
   DECLARE FUNCTION Add (BYREF dws AS DWSTRING) AS BOOLEAN
   DECLARE OPERATOR [] (BYVAL nIndex AS UINT) BYREF AS USHORT
   DECLARE OPERATOR CAST () BYREF AS CONST WSTRING
   DECLARE OPERATOR CAST () AS ANY PTR
   DECLARE OPERATOR LET (BYVAL pwszStr AS WSTRING PTR)
   DECLARE OPERATOR LET (BYREF ansiStr AS STRING)
   DECLARE OPERATOR LET (BYREF dws AS DWSTRING)
   DECLARE OPERATOR LET (BYVAL n AS LONGINT)
   DECLARE OPERATOR LET (BYVAL n AS DOUBLE)
   DECLARE PROPERTY utf8 () AS STRING
   DECLARE PROPERTY utf8 (BYREF ansiStr AS STRING)

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR DWSTRING
   DWSTRING_DP("DWSTRING CONSTRUCTOR - Default")
   ' // Create a minimal initial buffer to make room for the teminating null
   this.ResizeBuffer(1, TRUE)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Attempts to allocate, or reserve, nChars number of characters from the free store (heap).
' bClear = FALSE: The newly allocated memory is not initialized.
' bClear = TRUE: The newly allocated memory is initialized.
' ========================================================================================
PRIVATE CONSTRUCTOR DWSTRING (BYVAL nChars AS LONG, BYVAL bClear AS BOOLEAN)
   DWSTRING_DP("DWSTRING CONSTRUCTOR - nChars, bClear")
   IF nChars < 1 THEN nChars = 1
   this.ResizeBuffer(nChars, bClear)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes the DWSTRING with the contents of the passed WSTRING.
' ========================================================================================
PRIVATE CONSTRUCTOR DWSTRING (BYVAL pwszStr AS WSTRING PTR)
   DWSTRING_DP("DWSTRING CONSTRUCTOR - WSTRING PTR")
   IF pwszStr = NULL THEN
      this.ResizeBuffer(1)
   ELSEIF .LEN(*pwszStr) = 0 THEN
      this.ResizeBuffer(1)
   ELSE
      this.Add(pwszStr)
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes the DWSTRING with the contents of the passed STRING.
' The default codepage is 0 (CP_ACP). You can pass CP_UTF8 to assign utf-8 strings.
' ========================================================================================
PRIVATE CONSTRUCTOR DWSTRING (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
   DWSTRING_DP("DWSTRING CONSTRUCTOR - STRING - " & .WSTR(nCodePage))
   IF .LEN(ansiStr) = 0 THEN
      this.ResizeBuffer(1)
   ELSE
      this.Add(ansiStr, nCodePage)
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes the DWSTRING with the contents of the passed DWSTRING.
' ========================================================================================
PRIVATE CONSTRUCTOR DWSTRING (BYREF dws AS DWSTRING)
   DWSTRING_DP("DWSTRING CONSTRUCTOR - DWSTRING")
   IF dws.m_BufferLen = 0 THEN
      this.ResizeBuffer(1)
   ELSE
      this.Add(dws)
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes the DWSTRING with a number.
' These two constructors are needed to allow to use a number with the & operator.
' ========================================================================================
PRIVATE CONSTRUCTOR DWSTRING (BYVAL n AS LONGINT)
   DWSTRING_DP("DWSTRING CONSTRUCTOR LONGINT")
   DIM wsz AS WSTRING * 40 = .WSTR(n)
   this.Add(wsz)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR DWSTRING (BYVAL n AS DOUBLE)
   DWSTRING_DP("DWSTRING CONSTRUCTOR DOUBLE")
   DIM wsz AS WSTRING * 40 = .WSTR(n)
   this.Add(wsz)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR DWSTRING
   DWSTRING_DP("DWSTRING DESTRUCTOR - buffer: " & .WSTR(m_pBuffer))
   IF m_pBuffer THEN Deallocate(m_pBuffer)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the DWSTRING buffer.
' To avoid excesive use of casting when used with my other data types and procedures.
' ========================================================================================
PRIVATE OPERATOR DWSTRING.CAST () AS ANY PTR
   DWSTRING_DP("DWSTRING CAST ANY PTR - buffer: " & .WSTR(m_pBuffer))
   OPERATOR = m_pBuffer
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the string data (same as **).
' ========================================================================================
PRIVATE OPERATOR DWSTRING.CAST () BYREF AS CONST WSTRING
   DWSTRING_DP("DWSTRING CAST BYREF AS WSTRING - buffer: " & .WSTR(m_pBuffer))
   OPERATOR = *m_pBuffer
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the corresponding ASCII or Unicode integer representation of the character at
' the zero-based position specified by the nIndex parameter (0 for the first character,
' 1 for the second, etc.), e.g. value = dws[1]. To change the value: dws[1] = value.
' Example:
'   DIM dwsText AS DWSTRING = "This is my text."
'   dwsText[1] = ASC("X")
'   PRINT dwsText
' ========================================================================================
PRIVATE OPERATOR DWSTRING.[] (BYVAL nIndex AS UINT) BYREF AS USHORT
   DWSTRING_DP("DWSTRING Operator []")
   OPERATOR = *CAST(USHORT PTR, m_pBuffer + nIndex)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns new text to the DWSTRING
' ========================================================================================
PRIVATE OPERATOR DWSTRING.Let (BYVAL pwszStr AS WSTRING PTR)
   DWSTRING_DP("DWSTRING Add - WSTRING PTR")
   this.Clear : this.Add(*pwszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DWSTRING.Let (BYREF ansiStr AS STRING)
   DWSTRING_DP("DWSTRING LET STRING")
   this.Clear : this.Add(ansiStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DWSTRING.Let (BYREF dws AS DWSTRING)
   DWSTRING_DP("DWSTRING LET DWSTRING")
   IF m_pBuffer = dws.m_pBuffer THEN EXIT OPERATOR   ' // Ignore dws = dws
   this.Clear : this.Add(dws)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a number to the DWSTRING
' ========================================================================================
PRIVATE OPERATOR DWSTRING.Let (BYVAL n AS LONGINT)
   DWSTRING_DP("DWSTRING OPERATOR Let LONGINT")
   this.Clear : DIM wsz AS WSTRING * 40 = .WSTR(n) : this.Add(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DWSTRING.Let (BYVAL n AS DOUBLE)
   DWSTRING_DP("DWSTRING OPERATOR Let DOUBLE")
   this.Clear : DIM wsz AS WSTRING * 40 = .WSTR(n) : this.Add(wsz)
END OPERATOR
' ========================================================================================

' ========================================================================================
' The size of the internal string buffer is retrieved and returned to the caller. The size
' is the number of characters which can be stored without further expansion.
' ========================================================================================
PRIVATE PROPERTY DWSTRING.Capacity() AS UINT
   DWSTRING_DP("DWSTRING PROPERTY GET Capacity")
   PROPERTY = m_Capacity
END PROPERTY
' ========================================================================================

' ========================================================================================
' The internal string buffer is expanded to the specified number of characters. If the new
' capacity is smaller than the current capacity, the buffer is shortened and the contents
' that exceed the new capacity are lost.
' ========================================================================================
PRIVATE PROPERTY DWSTRING.Capacity (BYVAL nValue AS UINT)
   DWSTRING_DP("DWSTRING PROPERTY SET Capacity")
   IF nValue = m_Capacity THEN EXIT PROPERTY
   this.ResizeBuffer(nValue)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Converts the DWSTRING to UTF8.
' ========================================================================================
PRIVATE PROPERTY DWSTRING.Utf8 () AS STRING
   DWSTRING_DP("DWSTRING Utf8 GET PROPERTY")
   DIM cbLen AS INTEGER
   IF m_BufferLen = 0 THEN RETURN ""
   DIM buffer AS STRING = STRING(m_BufferLen * 5 + 1, 0)
   PROPERTY = *cast(ZSTRING PTR, WCharToUTF(1, m_pBuffer, m_BufferLen * 2, STRPTR(buffer), @cbLen))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Converts UTF8 to unicode and assigns it to the DWSTRING.
' ========================================================================================
PRIVATE PROPERTY DWSTRING.Utf8 (BYREF utf8String AS STRING)
   DWSTRING_DP("DWSTRING Utf8 SET PROPERTY")
   this.Clear
   this.Add(utf8String, CP_UTF8)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Resizes the internal buffer capacity
' ========================================================================================
PRIVATE FUNCTION DWSTRING.ResizeBuffer (BYVAL nChars AS UINT, BYVAL bClear AS BOOLEAN = FALSE) AS WSTRING PTR
   DWSTRING_DP("DWSTRING ResizeBuffer - nChars = " & .WSTR(nChars))
   ' // Allocate a buffer of nChars utf16 characters + 1 for the terminating null
   DIM pNewBuffer AS WSTRING PTR = IIF(bClear, CAllocate((nChars + 1) * 2), Allocate((nChars + 1) * 2))
   ' // The user has requested to reduce the capacity.
   ' // The buffer is shortened and the contents that exceed the new capacity are lost.
   IF nChars < m_BufferLen THEN m_BufferLen = nChars
   IF m_pBuffer THEN
      ' // Copy the old buffer in the new one
      wmemmove(pNewBuffer, m_pBuffer, m_BufferLen)
      ' // Deallocate the old buffer
      Deallocate m_pBuffer
   END IF
   ' // Update the capacity
   m_Capacity = nChars
   ' // Store the new pointer
   m_pBuffer = pNewBuffer
   ' // Mark the end of the string with a double null
   m_pBuffer[m_BufferLen] = 0
   RETURN m_pBuffer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends the specified number of characters from the specified memory address to the end of the buffer.
' ========================================================================================
PRIVATE FUNCTION DWSTRING.AppendBuffer (BYVAL memAddr AS ANY PTR, BYVAL nChars AS UINT) AS BOOLEAN
   DWSTRING_DP("DWSTRING AppendBuffer - nChars = " & .WSTR(nChars))
   IF memAddr = NULL OR nChars = 0 THEN RETURN FALSE
   ' // Number of characters to append
   DIM nSize AS UINT = m_BufferLen + nChars
   ' // If there is not enough capacity, resize the buffer
   IF nSize > m_Capacity THEN this.ResizeBuffer(nSize * 2)
   ' // Copy the passed buffer
   IF m_pBuffer = NULL THEN RETURN FALSE
   wmemmove(m_pBuffer + m_BufferLen, memAddr, nChars)
   ' // Update the length of the buffer
   m_BufferLen += nChars
   ' // Mark the end of the string with a double null
   m_pBuffer[m_BufferLen] = 0
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the DWSTRING buffer.
' ========================================================================================
PRIVATE FUNCTION DWSTRING.GetBuffer () AS WSTRING PTR
   DWSTRING_DP("DWSTRING GetBuffer - buffer: " & .WSTR(m_pBuffer))
   RETURN m_pBuffer
END FUNCTION
' ========================================================================================

' ========================================================================================
' All data in the class object is erased. Actually, we only set the buffer length to zero,
' indicating no string in the buffer. The allocated memory for the buffer is deallocated
' when the class is destroyed.
' ========================================================================================
PRIVATE SUB DWSTRING.Clear
   DWSTRING_DP("DWSTRING Clear")
   m_BufferLen = 0
   m_pBuffer[m_BufferLen] = 0
END SUB
' ========================================================================================

' ========================================================================================
' The string parameter is appended to the string held in the class. If the internal string
' buffer overflows, the class will automatically extend it to an appropriate size.
' ========================================================================================
PRIVATE FUNCTION DWSTRING.Add (BYVAL pwszStr AS WSTRING PTR) AS BOOLEAN
   DWSTRING_DP("DWSTRING Add - WSTRING PTR")
   RETURN this.AppendBuffer(pwszStr, .LEN(*pwszStr))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DWSTRING.Add (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
   DWSTRING_DP("DWSTRING Add - STRING")
   IF .LEN(ansiStr) = 0 THEN RETURN FALSE
   ' // Create the wide string from the incoming ansi string
   DIM dwLen AS UINT, pbuffer AS ANY PTR
   DIM bRes AS BOOLEAN = TRUE   ' // assume success for now
   IF nCodePage = CP_UTF8 THEN
      dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), .LEN(ansiStr), NULL, 0)
      IF dwLen = 0 THEN RETURN FALSE
      dwLen *= 2
      pbuffer = Allocate(dwLen)
      dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), .LEN(ansiStr), pbuffer, dwLen)
      IF dwLen = 0 THEN bRes = FALSE
   ELSE
      dwLen = MultiByteToWideChar(nCodePage, MB_PRECOMPOSED, STRPTR(ansiStr), .LEN(ansiStr), pbuffer, dwLen)
      IF dwLen = 0 THEN RETURN FALSE
      dwLen *= 2
      pbuffer = Allocate(dwLen)
      dwLen = MultiByteToWideChar(nCodePage, MB_PRECOMPOSED, STRPTR(ansiStr), .LEN(ansiStr), pbuffer, dwLen)
      IF dwLen = 0 THEN bRes = FALSE
   END IF
   IF bRes = FALSE THEN
      IF pBuffer THEN Deallocate(pbuffer)
      RETURN bRes
   END IF
   ' // Copy the string into the buffer
   IF pbuffer THEN
      ' Copy the string into the buffer and update the length
      bRes = this.AppendBuffer(pbuffer, dwLen)
      ' // Deallocate the buffer
      IF pBuffer THEN Deallocate(pbuffer)
   END IF
   RETURN bRes
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DWSTRING.Add (BYREF dws AS DWSTRING) AS BOOLEAN
   DWSTRING_DP("DWSTRING Add - DWSTRING")
   RETURN this.AppendBuffer(dws.m_pBuffer, dws.m_BufferLen)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                *** HELPER FUNCTIONS ***
' ########################################################################################

' ========================================================================================
' qsort DWSTRING comparison function
' ========================================================================================
PRIVATE FUNCTION DWStringArrayCompare CDECL (BYVAL a AS DWSTRING PTR, BYVAL b AS DWSTRING PTR) AS LONG
   FUNCTION = wcscmp(a->m_pBuffer, b->m_pBuffer)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reverse qsort DWSTRING comparison function
' ========================================================================================
PRIVATE FUNCTION DWStringArrayReverseCompare CDECL (BYVAL a AS DWSTRING PTR, BYVAL b AS DWSTRING PTR) AS LONG
   DIM r AS LONG = wcscmp(a->m_pBuffer, b->m_pBuffer)
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sorts a one-dimensional DWSTRING array calling the C qsort function.
' Parameters:
' - rgwstr : Start of target array.
' - numElm : Number of elements in the array.
' - bAscend: TRUE for sorting in ascending order; FALSE for sorting in descending order.
' Example:
' DIM rg(1 TO 10) AS DWSTRING
' FOR i AS LONG = 1 TO 10
'    rg(i) = "string " & i
' NEXT
' FOR i AS LONG = 1 TO 10
'   print rg(i)
' NEXT
' print "---- after sorting ----"
' DWStringArrayQSort rg()
' FOR i AS LONG = 1 TO 10
'    print rg(i)
' NEXT
' ========================================================================================
PRIVATE SUB DWStringArrayQSort OVERLOAD (BYREF rgwstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgwstr = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgwstr, numElm, SIZEOF(DWSTRING), CPTR(ANY PTR, @DWStringArrayCompare)
   ELSE
      qsort rgwstr, numElm, SIZEOF(DWSTRING) , CPTR(ANY PTR, @DWStringArrayReverseCompare)
   END IF
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB DWStringArrayQSort OVERLOAD (rgwstr() AS DWSTRING, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgwstr) - LBOUND(rgwstr) + 1
   DWStringArrayQSort @rgwstr(LBOUND(rgwstr)), numElm, bAscend
END SUB
' ========================================================================================
