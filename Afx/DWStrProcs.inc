' ########################################################################################
' Platform: Microsoft Windows
' Filename: DWSTRING.bi
' Purpose: string procedures for the DWString data type
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
#ifndef _WIN32_WINNT
   #define _WIN32_WINNT &h0602
#endif
#include once "Afx/DWString.inc"
#include once "win/winnls.bi"

' ========================================================================================
' Returns a string with nCount characters removed from the left side of the string.
' If nCount is less than one then the entire string is returned.
' Example: DIM dws AS DWSTRING = DWSTRING_ClipLeft("1234567890", 3)
' Result: "4567890"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ClipLeft (BYREF wszMainStr AS CONST WSTRING, BYVAL nCount AS LONG) AS DWSTRING
   IF nCount <= 0 THEN RETURN ""
   DIM nLen AS LONG = LEN(wszMainStr)
   nCount = IIF(nLen < nCount, nLen, nCount)
   RETURN MID(wszMainStr, nCount + 1)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string with nCount characters removed from the right side of the string.
' If nCount is less than one then the entire string is returned.
' Example: DIM dws AS DWSTRING = DWSTRING_ClipRight("1234567890", 3)
' Result: "1234567"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ClipRight (BYREF wszMainStr AS CONST WSTRING, BYVAL nCount AS LONG) AS DWSTRING
   IF nCount <= 0 THEN RETURN ""
   DIM nLen AS LONG = LEN(wszMainStr)
   nCount = nLen - nCount
   nCount = IIF(nLen < nCount, nLen, nCount)
   RETURN LEFT(wszMainStr, nCount)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string with nCount characters removed starting at position nStart. The first
' character is considered position 1, the second position 2, etc...
' If nCount or nStart is less than one then the entire string is returned.
' Exaple: DIM dws AS DWSTRING = DWSTRING_ClipMid("1234567890", 3, 4)
' Result: "127890"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ClipMid (BYREF wszMainStr AS CONST WSTRING, BYVAL nStart AS LONG, BYVAL nCount AS LONG) AS DWSTRING
   DIM dws AS DWSTRING = wszMainStr
   IF (nCount <= 0) OR (nStart <= 0) THEN RETURN ""
   DIM nLen AS LONG = LEN(wszMainStr)
   RETURN LEFT(wszMainStr, nStart - 1) + MID(wszMainStr, nStart + nCount)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing a centered (padded) string.
' If the optional parameter wszPadCharacter not specified, the function pads the string with
' space characters to the left. Otherwise, the function pads the string with the first
' character of wszPadCharacter.
' Example: DIM dws AS DWSTRING = DWSTRING_CSet("FreeBasic", 20, "*")
' Result: "*****FreeBasic******"
' Example: DIM dws AS DWSTRING = DWSTRING_CSet("FreeBasic", 20)
' Result: "     FreeBasic      "
' ========================================================================================
PRIVATE FUNCTION DWSTRING_CSet (BYREF wszMainStr AS CONST WSTRING, BYVAL nStringLength AS LONG, BYREF wszPadCharacter AS CONST WSTRING = " ") AS DWSTRING
   IF LEN(wszMainStr) > nStringLength THEN RETURN LEFT(wszMainStr, nStringLength)
   DIM dws AS DWSTRING = WSTRING(nStringLength, wszPadCharacter)
   MID(dws, (nStringLength - LEN(wszMainStr)) \ 2 + 1, LEN(wszMainStr)) = wszMainStr
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing a centered string within the space of another string.
' Example: DIM dwsstr AS DWSTRING = "COOL COOL COOL COOL COOL"
'          DIM dws AS DWSTRING = DWSTRING_CSetAbs("..FreeBasic is..", dwsstr)
' Result: "COOL..FreeBasic is..COOL"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_CSetAbs (BYREF wszMainStr AS CONST WSTRING, BYREF wszPadString AS CONST WSTRING) AS DWSTRING
   IF LEN(wszMainStr) > LEN(wszPadString) THEN RETURN LEFT(wszMainStr, LEN(wszPadString))
   DIM dws AS DWSTRING = wszPadString
   MID(dws, (LEN(wszPadString) - LEN(wszMainStr)) \ 2 + 1, LEN(wszMainStr)) = wszMainStr
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a specified number of characters from a string expression.
' Returns a string based on wszMainStr but with nCount characters deleted
' starting at position nStart. The first character in the string is position 1, etc.
' Example: DIM dws AS DWSTRING = DWSTRING_Delete("1234567890", 4, 3)
' Result: "1237890"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_Delete (BYREF wszMainStr AS CONST WSTRING, BYVAL nStart AS LONG, BYVAL nCount AS LONG) AS DWSTRING
   DIM nLen AS LONG = LEN(wszMainStr)
   IF nLen = 0 OR nStart < 0 OR nCount <= 0 OR nStart > nLen THEN RETURN ""
   RETURN LEFT(wszMainStr, nStart - 1) & MID(wszMainStr, nStart + nCount)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts characters from a string up to a character or group of characters.
' Complementary function to DSTRING_Remain.
' Returns a substring of wszMainStr starting with its first character (or the character
' specified by nStart) and up to (but not including) the first occurrence of wszMatchStr
' If wszMatchStr is not present in wszMainStr (or is null) then all of wszMainStr is
' returned from the nStart position.
' This function is case-sensitive.
' The following line returns "aba" (match on "cad")
' DIM dws AS DWSTRING = DWSTRING_Extract(1, "abacadabra","cad")
' ========================================================================================
PRIVATE FUNCTION DWSTRING_Extract OVERLOAD (BYVAL nStart AS LONG = 1, BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS DWSTRING
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN RETURN ""
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   DIM nPos AS LONG = INSTR(nStart, wszMainStr, wszMatchStr)
   IF nPos THEN RETURN MID(wszMainStr, nStart, nPos - nStart)
   RETURN MID(wszMainStr, nStart)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Case insensitive version of Extract.
' The following line returns "aba" (match on "CaD")
' DIM dws AS DWSTRING = DWSTRING_ExtractI(1, "abacadabra","CaD")
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ExtractI (BYVAL nStart AS LONG = 1, BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS DWSTRING
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN RETURN ""
   DIM dws AS DWSTRING = wszMainStr
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   DIM nPos AS LONG = INSTR(nStart, UCASE(wszMainStr), UCASE(wszMatchStr))
   IF nPos THEN RETURN MID(wszMainStr, nStart, nPos - nStart )
   RETURN MID(wszMainStr, nStart)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the portion of a string following the occurrence of a specified delimiter up to
' the second delimiter. If one of the delimiters isn't found, it returns an empty string.
' Parameters:
' wszMainStr: The main string
' wszDelim1: The first delimiter
' wszDelim2: The second delimiter
' This function is case-sensitive.
' Example:
' DIM dws AS DWSTRING = "blah blah (text beween parentheses) blah blah"
' dws = DWSTRING_Extract(dws, "(", ")")
' Result: "text beween parentheses"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_Extract OVERLOAD (BYREF wszMainStr AS CONST WSTRING, BYREF wszDelim1 AS CONST WSTRING, BYREF wszDelim2 AS CONST WSTRING) AS DWSTRING
   DIM nPos1 AS LONG = INSTR(wszMainStr, wszDelim1)
   IF nPos1 = 0 THEN RETURN ""
   nPos1 += LEN(wszDelim1)
   DIM nPos2 AS LONG = INSTR(nPos1, wszMainStr, wszDelim2)
   IF nPos2 = 0 THEN RETURN ""
   DIM nLen AS LONG = nPos2 - nPos1
   RETURN MID(wszMainStr, nPos1, nLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the portion of a string following the occurrence of a specified delimiter up to
' the second delimiter. If one of the delimiters isn't found, it returns an empty string.
' Parameters:
' nStart: [Optional]. The one-based starting position where to start the search
' wszMainStr: The main string
' wszDelim1: The first delimiter
' wszDelim2: The second delimiter
' This function is case-sensitive.
' Example:
' DIM dws AS DWSTRING = "blah blah (text beween parentheses) blah blah"
' dws = DWSTRING_Extract(1, dws, "(", ")")
' Result: "text beween parentheses"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_Extract OVERLOAD (BYVAL nStart AS LONG = 1, BYREF wszMainStr AS CONST WSTRING, BYREF wszDelim1 AS CONST WSTRING, BYREF wszDelim2 AS CONST WSTRING) AS DWSTRING
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN RETURN ""
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   DIM nPos1 AS LONG = INSTR(nStart, wszMainStr, wszDelim1)
   IF nPos1 = 0 THEN RETURN ""
   nPos1 += LEN(wszDelim1)
   DIM nPos2 AS LONG = INSTR(nPos1, wszMainStr, wszDelim2)
   IF nPos2 = 0 THEN RETURN ""
   nLen = nPos2 - nPos1
   RETURN MID(wszMainStr, nPos1, nLen)
END FUNCTION
' ========================================================================================


' ========================================================================================
' Extract characters from a string up to a specific character.
' Returns a substring of wszMainStr starting with its first character (or the character
' specified by nStart) and up to (but not including) the first occurrence of wszMatchStr.
' wszMatchStr specifies a list of single characters to be searched for individually, a
' match on any one of which will cause the extract operation to be performed up to that character.
' If wszMatchStr is not present in wszMainStr (or is null) then all of wszMainStr is returned.
' This function is case-sensitive.
' The following line returns "aba" (match on "c")
' Example: DWSTRING_ExtractAny(1, "abacadabra","cd")
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ExtractAny (BYVAL nStart AS LONG = 1, BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS DWSTRING
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN RETURN ""
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   FOR i AS LONG = nStart TO nLen
      FOR x AS LONG = 1 TO LEN(wszMatchStr)
         IF MID(wszMainStr, i, 1) = MID(wszMatchStr, x, 1) THEN
            RETURN MID(wszMainStr, nStart, i - nStart)
         END IF
      NEXT
   NEXT
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' Case insensitive version of DWStrExtractAny.
' Example: DWSTRING_ExtractAnyI(1, "abacadabra","CD")
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ExtractAnyI (BYVAL nStart AS LONG = 1, BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS DWSTRING
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN RETURN ""
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   FOR i AS LONG = nStart TO nLen
      FOR x AS LONG = 1 TO LEN(wszMatchStr)
         IF MID(UCASE(wszMainStr), i, 1) = MID(UCASE(wszMatchStr), x, 1) THEN
            RETURN MID(wszMainStr, nStart, i - nStart)
         END IF
      NEXT
   NEXT
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a string at a specified position within another string expression.
' Returns a string consisting of wszMainStr with the string wszInsertString inserted
' at nPosition. If nPosition is greater than the length of wszMainStr or <= zero then
' wszInsertString is appended to wszMainStr. The first character in the string is position 1, etc.
' Example: DIM dws AS DWSTRING = DWSTRING_Insert("1234567890", "--", 6)
' Result: "12345--67890"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_Insert (BYREF wszMainStr AS CONST WSTRING, BYREF wszInsertString AS WSTRING, BYVAL nPosition AS LONG) AS DWSTRING
   IF nPosition <= 0 THEN RETURN ""
   IF nPosition > LEN(wszMainStr) THEN
      RETURN wszMainStr & wszInsertString
   ELSEIF nPosition = 1 THEN
      RETURN wszInsertString & MID(wszMainStr, 1)
   ELSE
      RETURN MID(wszMainStr, 1, nPosition - 1) + wszInsertString + MID(wszMainStr, nPosition)
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a string containing a left-justified (padded) string.
' If the optional parameter wszPadCharacter not specified, the function pads the string with
' space characters to the left. Otherwise, the function pads the string with the first
' character of wszPadCharacter
' Example: DIM dws AS DWSTRING = DWSTRING_LSet("FreeBasic", 20)
' Result: "FreeBasic           "
' Example: DIM dws AS DWSTRING = DWSTRING_LSet("FreeBasic", 20, "*")
' Result: "FreeBasic***********"
' ========================================================================================
PRIVATE FUNCTION DWSTRIG_LSet (BYREF wszMainStr AS CONST WSTRING, BYVAL nStringLength AS LONG, BYREF wszPadCharacter AS CONST WSTRING = " ") AS DWSTRING
   DIM dws AS DWSTRING = WSTRING(nStringLength, wszPadCharacter)
   MID(dws, 1, LEN(wszMainStr)) = wszMainStr
   RETURN dws
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a string containing a left-justified string within the space of another string.
' Example: DIM dwsstr AS DWSTRING = "BestBasic=BestBasic"
'          DIM dws AS DWSTRING = DWSTRING_LSetAbs("FreeBasic", dwsstr)
' Result: "FreeBasic=BestBasic"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_LSetAbs (BYREF wszMainStr AS CONST WSTRING, BYREF wszPadString AS CONST WSTRING) AS DWSTRING
   IF LEN(wszMainStr) > LEN(wszPadString) THEN RETURN LEFT(wszMainStr, LEN(wszPadString))
   DIM dws AS DWSTRING = wszPadString
   MID(dws, 1, LEN(wszMainStr)) = wszMainStr
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing a right-justified (padded) string.
' If the optional parameter wszPadCharacter not specified, the function pads the string with
' space characters to the left. Otherwise, the function pads the string with the first
' character of wszPadCharacter.
' Example: DIM dws AS DWSTRING = DWSTRIG_RSet("FreeBasic", 20, "*")
' Result: "***********FreeBasic"
' ========================================================================================
PRIVATE FUNCTION DWSTRIG_RSet (BYREF wszMainStr AS CONST WSTRING, BYVAL nStringLength AS LONG, BYREF wszPadCharacter AS CONST WSTRING = " ") AS DWSTRING
   IF LEN(wszMainStr) > nStringLength THEN RETURN LEFT(wszMainStr, nStringLength)
   DIM dws AS DWSTRING = WSTRING(nStringLength, wszPadCharacter)
   MID(dws, nStringLength - LEN(wszMainStr) + 1, LEN(wszMainStr)) = wszMainStr
   RETURN dws
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a string containing a right-justified string within the space of another string.
' Example: DIM dwsstr AS DWSTRING = "BestBasic=BestBasic"
'          DIM dws AS DWSTRING = DWSTRING_RSetAbs("FreeBasic", dwsstr)
' Result: "BestBasic=FreeBasic"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_RSetAbs (BYREF wszMainStr AS CONST WSTRING, BYREF wszPadString AS CONST WSTRING) AS DWSTRING
   IF LEN(wszMainStr) > LEN(wszPadString) THEN RETURN LEFT(wszMainStr, LEN(wszPadString))
   DIM dws AS DWSTRING = wszPadString
   DIM nStart AS LONG = LEN(wszPadString) - LEN(wszMainStr) + 1
   IF nStart < 1 THEN nStart = 1
   MID(dws, nStart, LEN(wszMainStr)) = wszMainStr
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the count of delimited fields from a string expression.
' If wszMainStr is empty (a null string) or contains no delimiter character(s), the string
' is considered to contain exactly one sub-field. In this case, DWStrParseCount returns the value 1.
' Delimiter contains a string (one or more characters) that must be fully matched.
' Delimiters are case-sensitive.
' Example: DIM nCount AS LONG = DWSTRING_ParseCount("one,two,three", ",")
' Result = 3
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ParseCount (BYREF wszMainStr AS CONST WSTRING, BYREF wszDelimiter AS CONST WSTRING = ",") AS LONG
   DIM nCount AS LONG = 1
   DIM nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, wszMainStr, wszDelimiter)
      IF nPos = 0 THEN EXIT DO
      nCount += 1
      nPos += LEN(wszDelimiter)
   LOOP
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Return the count of delimited fields from a string expression.
' If wszMainStr is empty (a null string) or contains no delimiter character(s), the string
' is considered to contain exactly one sub-field. In this case, DWStrParseCountAny returns the value 1.
' Delimiter contains a set of characters (one or more), any of which may act as a delimiter character.
' Delimiters are case-sensitive.
' Example: DIM nCount AS LONG = DWSTRING_ParseCountAny("1;2,3", ",;")
' Result = 5
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ParseCountAny (BYREF wszMainStr AS CONST WSTRING, BYREF wszDelimiter AS CONST WSTRING = ",") AS LONG
   DIM nCount AS LONG = 1
   FOR i AS LONG = 1 TO LEN(wszDelimiter)
      nCount += DWSTRING_ParseCount(wszMainStr, MID(wszDelimiter, i, 1))
   NEXT
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the nPosition-th substring in a string wszMainStr with separations wszDelimiter
' (one or more characters), beginning with nPosition = 1.
' ========================================================================================
PRIVATE FUNCTION DWSTRING_Parse OVERLOAD (BYREF wszMainStr AS CONST WSTRING, BYREF wszDelimiter AS CONST WSTRING, BYVAL nPosition AS LONG, BYVAL bIsAny AS BOOLEAN, BYVAL bIsCaseInsensitive AS BOOLEAN, BYVAL nLenDelimiter AS LONG) AS DWSTRING
   DIM nCount AS LONG, nStart AS LONG
   DIM fReverse AS BOOLEAN = IIF(nPosition < 0, TRUE, FALSE)
   nPosition = ABS(nPosition)
   DIM nPos AS LONG = 1
   IF fReverse THEN
      ' Reverse search
      ' Get the start of the token (j) by searching in reverse
      IF bIsAny THEN
         IF bIsCaseInsensitive THEN
            nPos = InstrRev(UCASE(wszMainStr), ANY UCASE(wszDelimiter))
         ELSE
            nPos = InstrRev(wszMainStr, ANY wszDelimiter)
         END IF
      ELSE
         IF bIsCaseInsensitive THEN
            nPos = InstrRev(UCAsE(wszMainStr), UCAsE(wszDelimiter))
         ELSE
            nPos = InstrRev(wszMainStr, wszDelimiter)
         END IF
      END IF
      DO WHILE nPos > 0        ' if not found loop will be skipped
         nStart = nPos + nLenDelimiter
         nCount += 1
         nPos = nPos - nLenDelimiter
         IF nCount = nPosition THEN EXIT DO
         IF bIsAny THEN
            IF bIsCaseInsensitive THEN
               nPos = InStrRev(UCAsE(wszMainStr), ANY UCASE(wszDelimiter), nPos)
            ELSE
               nPos = InStrRev(wszMainStr, ANY wszDelimiter, nPos)
            END IF
         ELSE
            IF bIsCaseInsensitive THEN
               nPos = InStrRev(UCASE(wszMainStr), UCASE(wszDelimiter), nPos)
            ELSE
               nPos = InStrRev(wszMainStr, wszDelimiter, nPos)
            END IF
         END IF
      LOOP
      IF nPos = 0 THEN nStart = 1
      ' Now continue forward to get the end of the token
      IF bIsAny THEN
         IF bIsCaseInsensitive THEN
            nPos = INSTR(nStart, UCASE(wszMainStr), ANY UCASE(wszDelimiter))
         ELSE
            nPos = INSTR(nStart, wszMainStr, ANY wszDelimiter)
         END IF
      ELSE
         nPos = INSTR(nStart, wszMainStr, wszDelimiter)
      END IF
      IF nPos > 0 OR nCount = nPosition THEN
         IF nPos = 0 THEN
            RETURN MID(wszMainStr, nStart)
         ELSE
            RETURN MID(wszMainStr, nStart, nPos - nStart)
         END IF
      END IF
   ELSE
      ' Forward search
      DO
         nStart = nPos
         IF bIsAny THEN
            IF bIsCaseInsensitive THEN
               nPos = INSTR(nPos, UCASE(wszMainStr), ANY UCASE(wszDelimiter))
            ELSE
               nPos = INSTR(nPos, wszMainStr, ANY wszDelimiter)
            END IF
         ELSE
            IF bIsCaseInsensitive THEN
               nPos = INSTR(nPos, UCASE(wszMainStr), UCASE(wszDelimiter))
            ELSE
               nPos = INSTR(nPos, wszMainStr, wszDelimiter)
            END IF
         END IF
         IF nPos THEN
            nCount += 1
            nPos += nLenDelimiter
         END IF
      LOOP UNTIL nPos = 0 OR nCount = nPosition
      IF nPos > 0 OR nCount = nPosition - 1 THEN
         IF nPos = 0 THEN
            RETURN MID(wszMainStr, nStart)
         ELSE
print nStart, nPos, nLenDelimiter - nStart
            RETURN MID(wszMainStr, nStart, nPos - nLenDelimiter - nStart)
         END IF
      END IF
   END IF
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a delimited field from a string expression.
' wszDelimiter contains a string of one or more characters that must be fully matched to be successful.
' If nPosition evaluates to zero or is outside of the actual field count, an empty string is returned.
' If nPosition is negative then fields are searched from the right to left of the wszMainStr
' Delimiters are case-sensitive.
' Example: DWSTRING_Parse("one,two,three", 2)
' Result = "two"
' Example: DWSTRING_Parse("one;two,three", 2)
' Result = "three"
' Example: DWSTRING_Parse("one", 2)
' Result = ""
' Example: DWSTRING_Parse("xyz", 1)
' Result = "xyz
' Example: DWSTRING_Parse("xx1x", "x", 3)
' Result: "1"
' Example: DWSTRING_ParseAny("1;2,3", ",;", 2)
' Result: "2"
' Example: DWSTRING_Parse("one;two,three", 1, ";")
' Result: "one"
' Example: DWSTRING_Parse("one;two,three", -1, ";")
' Result: "two,three"
' Example: DWSTRING_Parse("one,two,three", 1, "w")
' Result = "one,t"
' Example: DWSTRING_Parse("one,two,three", 1, "W")
' Result = "one,two,three"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_Parse OVERLOAD (BYREF wszMainStr AS CONST WSTRING, BYVAL nPosition AS LONG = 1, BYREF wszDelimiter AS CONST WSTRING = ",") AS DWSTRING
   ' The parse must match the entire deliminter string
   RETURN DWSTRING_Parse(wszMainStr, wszDelimiter, nPosition, FALSE, FALSE, Len(wszDelimiter))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Case insensitive version of Parse.
' Example: DWSTRING_ParseI("one,two,three", 1, "W")
' Result = "one,t"
' Example: DWSTRING_ParseI("one,two,three", 1, "T")
' Result = "one,"
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ParseI OVERLOAD (BYREF wszMainStr AS CONST WSTRING, BYVAL nPosition AS LONG = 1, BYREF wszDelimiter AS CONST WSTRING = ",") AS DWSTRING
   ' The parse must match the entire deliminter string
   RETURN DWSTRING_Parse(wszMainStr, wszDelimiter, nPosition, FALSE, TRUE, Len(wszDelimiter))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Return a delimited field from a string expression.
' Delimiter contains a set of characters (one or more), any of which may act as a delimiter character.
' If nPosition evaluates to zero or is outside of the actual field count, an empty string is returned.
' If nPosition is negative then fields are searched from the right to left of the MainString.
' Delimiters are case-sensitive.
' Example: DWSTRING_ParseAny("1;2,3", 2, ",;")
' Result: 2
' Example: DWSTRING_ParseAny("1;2,3", -1, ",;")
' Result: 3
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ParseAny (BYREF wszMainStr AS CONST WSTRING, BYVAL nPosition AS LONG = 1, BYREF wszDelimiter AS CONST WSTRING = ",") AS DWSTRING
   ' The parse must match one character (len = 1) in the delimiter string
   RETURN DWSTRING_Parse(wszMainStr, wszDelimiter, nPosition, TRUE, FALSE, 1)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Return a delimited field from a string expression.
' Delimiter contains a set of characters (one or more), any of which may act as a delimiter character.
' If nPosition evaluates to zero or is outside of the actual field count, an empty string is returned.
' If nPosition is negative then fields are searched from the right to left of the MainString.
' Delimiters are case insensitive.
' Example: DIM dws AS DWSTRING = DWSTRING_ParseAnyI("one;two,three", 2, "t;")
' Result: 2
' Example: DIM dws AS DWSTRING = DWSTRING_ParseAny("one;two,three", -1, "t;")
' Result: 3
' ========================================================================================
PRIVATE FUNCTION DWSTRING_ParseAnyI (BYREF wszMainStr AS CONST WSTRING, BYVAL nPosition AS LONG = 1, BYREF wszDelimiter AS CONST WSTRING = ",") AS DWSTRING
   ' The parse must match one character (len = 1) in the delimiter string
   RETURN DWSTRING_Parse(wszMainStr, wszDelimiter, nPosition, TRUE, TRUE, 1)
END FUNCTION
' ========================================================================================
