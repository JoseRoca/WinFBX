' ########################################################################################
' Microsoft Windows
' File: CRegExp.inc
' Contents: Regular expressions
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' Copyright (c) 2025 José Roca. Revised version.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxRegExp.bi"
#include once "Afx/AfxCOM.inc"
#include once "Afx/CVar.inc"
using Afx

NAMESPACE Afx

' ========================================================================================
' CRegExp class
' ========================================================================================
TYPE CRegExp

   Public:
      m_Result AS HRESULT                            ' // HRESULT
      m_pRegExp AS Afx_IRegExp2 PTR                  ' // IRegExp2 interface
      m_pMatches AS Afx_IMatchCollection2 PTR        ' // IMatchCollection2 interface

   Public:
      DECLARE CONSTRUCTOR
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
      DECLARE FUNCTION GetErrorInfo (BYVAL dwError AS LONG = -1) AS CWSTR
      DECLARE PROPERTY RegExpPtr () AS Afx_IRegExp2 PTR
      DECLARE PROPERTY Pattern () AS CBSTR
      DECLARE PROPERTY Pattern (BYREF cbsPattern AS CBSTR)
      DECLARE PROPERTY Global () AS BOOLEAN
      DECLARE PROPERTY Global (BYVAL bGlobal AS BOOLEAN)
      DECLARE PROPERTY IgnoreCase () AS BOOLEAN
      DECLARE PROPERTY IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
      DECLARE PROPERTY Multiline () AS BOOLEAN
      DECLARE PROPERTY Multiline (BYVAL bMultiline AS BOOLEAN)
      DECLARE FUNCTION Execute (BYREF cbsSourceString AS CBSTR) AS BOOLEAN
      DECLARE FUNCTION MatchesCount () AS LONG
      DECLARE FUNCTION MatchValue (BYVAL index AS LONG = 0) AS CBSTR
      DECLARE FUNCTION MatchPos (BYVAL index AS LONG = 0) AS LONG
      DECLARE FUNCTION MatchLen (BYVAL index AS LONG = 0) AS LONG
      DECLARE FUNCTION SubMatchesCount (BYVAL index AS LONG = 0) AS LONG
      DECLARE FUNCTION SubMatchValue (BYVAL MatchIndex AS LONG = 0, BYVAL SubMatchIndex AS LONG = 0) AS CVAR
      DECLARE FUNCTION Test (BYREF cbsSourceString AS CBSTR) AS BOOLEAN
      DECLARE FUNCTION Find (BYREF cbsSourceString AS CBSTR) AS LONG
      DECLARE FUNCTION Find (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR) AS LONG
      DECLARE FUNCTION FindEx (BYREF cbsSourceString AS CBSTR) AS CBSTR
      DECLARE FUNCTION Replace (BYREF cbsSourceString AS CBSTR, BYREF cvReplaceString AS CVAR) AS CBSTR
      DECLARE FUNCTION Remove (BYREF cbsSourceString AS CBSTR) AS CBSTR
      DECLARE FUNCTION Extract (BYREF cbsSourceString AS CBSTR) AS CBSTR
      DECLARE FUNCTION Extract (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR) AS CBSTR

END TYPE
' ========================================================================================

' ========================================================================================
' CRegExp default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CRegExp
   ' // Initialize the COM library
   CoInitialize(NULL)
   ' // Create an instance of the IRegExp2 interface
   m_pRegExp = AfxNewCom("VBScript.RegExp")
   IF m_pRegExp = NULL THEN
      CoUninitialize
      EXIT CONSTRUCTOR
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' CRegExp destructor
' ========================================================================================
PRIVATE DESTRUCTOR CRegExp
   ' // Release the interfaces
   IF m_pMatches THEN m_pMatches->Release
   IF m_pRegExp THEN m_pRegExp->Release
   ' // Uninitialize the COM library
   CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CRegExp.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CRegExp.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CRegExp.GetErrorInfo (BYVAL nError AS LONG = -1) AS CWSTR
   IF nError = -1 THEN nError = m_Result
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, cwsMsg AS CWSTR
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, nError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      cwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   IF nError THEN cwsMsg = "Error &h" & HEX(nError) & CHR(13, 10) & cwsMsg
   RETURN cwsMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the Afx_IRegExp2 interface
' ========================================================================================
PRIVATE PROPERTY CRegExp.RegExpPtr () AS Afx_IRegExp2 PTR
   m_Result = 0
   PROPERTY = m_pRegExp
END PROPERTY
' ========================================================================================

' ========================================================================================
' * The regular expression pattern to use.
' ========================================================================================
PRIVATE PROPERTY CRegExp.Pattern () AS CBSTR
   DIM bstrPattern AS AFX_BSTR
   this.SetResult(m_pRegExp->get_Pattern(@bstrPattern))
   PROPERTY = bstrPattern
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.Pattern (BYREF cbsPattern AS CBSTR)
   this.SetResult(m_pRegExp->put_Pattern(cbsPattern))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Global seach for all occurrences of the pattern.
' ========================================================================================
PRIVATE PROPERTY CRegExp.Global () AS BOOLEAN
   DIM bGlobal AS VARIANT_BOOL
   this.SetResult(m_pRegExp->get_Global(@bGlobal))
   PROPERTY = bGlobal
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.Global (BYVAL bGlobal AS BOOLEAN)
   this.SetResult(m_pRegExp->put_Global(cast(VARIANT_BOOL, bGlobal)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Ignore case.
' ========================================================================================
PRIVATE PROPERTY CRegExp.IgnoreCase () AS BOOLEAN
   DIM bIgnoreCase AS VARIANT_BOOL
   this.SetResult(m_pRegExp->get_IgnoreCase(@bIgnoreCase))
   PROPERTY = bIgnoreCase
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
   this.SetResult(m_pRegExp->put_IgnoreCase(cast(VARIANT_BOOL, bIgnoreCase)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Multiline search.
' ========================================================================================
PRIVATE PROPERTY CRegExp.Multiline () AS BOOLEAN
   DIM bMultiline AS VARIANT_BOOL
   this.SetResult(m_pRegExp->get_Multiline(@bMultiline))
   PROPERTY = bMultiline
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.Multiline (BYVAL bMultiline AS BOOLEAN)
   this.SetResult(m_pRegExp->put_Multiline(cast(VARIANT_BOOL, bMultiline)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Executes a regular expression search against a specified string.
' Example:
' #include once "Afx/CRegExp.inc"
' USING Afx
' DIM pRegExp AS CRegExp
' pRegExp.Pattern = "is."
' pRegExp.IgnoreCase = TRUE
' pRegExp.Global = TRUE
' IF pRegExp.Execute("IS1 is2 IS3 is4") = FALSE THEN
'    PRINT "No match found"
' ELSE
'    DIM nCount AS LONG = pRegExp.MatchesCount
'    FOR i AS LONG = 0 TO nCount - 1
'       PRINT "Value: ", pRegExp.MatchValue(i)
'       PRINT "Position: ", pRegExp.MatchPos(i)
'       PRINT "Length: ", pRegExp.MatchLen(i)
'       PRINT
'    NEXT
' END IF
' ========================================================================================
PRIVATE FUNCTION CRegExp.Execute (BYREF cbsSourceString AS CBSTR) AS BOOLEAN
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   IF m_Result = S_OK THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the number of matches.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchesCount () AS LONG
   IF m_pMatches = NULL THEN this.SetResult(E_POINTER): RETURN 0
   DIM nCount AS LONG
   this.SetResult(m_pMatches->get_Count(@nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the value or text of a match found in a search string.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchValue (BYVAL index AS LONG) AS CBSTR
   IF m_pMatches = NULL THEN this.SetResult(E_POINTER): RETURN ""
   DIM bstrValue AS AFX_BSTR
   DIM pMatch AS Afx_IMatch2 PTR
   this.SetResult(m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch)))
   IF pMatch THEN
      pMatch->get_Value(@bstrValue)
      AfxSafeRelease(pMatch)
   END IF
   RETURN bstrValue
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the position in a search string where a match occurs.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchPos (BYVAL index AS LONG) AS LONG
   IF m_pMatches = NULL THEN this.SetResult(E_POINTER): RETURN 0
   DIM nFirstIndex AS LONG
   DIM pMatch AS Afx_IMatch2 PTR
   this.SetResult(m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch)))
   IF pMatch THEN
      ' // FirstIndex property indicates the number of characters in the string to
      ' // the left of the match. If the match was found at the very start of the
      ' // string, FirstIndex will be zero. If the match starts at the second character
      ' // in the string, FirstIndex will be one, etc.
      IF pMatch->get_FirstIndex(@nFirstIndex) = S_OK THEN nFirstIndex += 1
      AfxSafeRelease(pMatch)
   END IF
   RETURN nFirstIndex
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the position in a search string where a match occurs.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchLen (BYVAL index AS LONG) AS LONG
   IF m_pMatches = NULL THEN this.SetResult(E_POINTER): RETURN 0
   DIM nLen AS LONG
   DIM pMatch AS Afx_IMatch2 PTR
   this.SetResult(m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch)))
   IF pMatch THEN
      pMatch->get_Length(@nLen)
      AfxSafeRelease(pMatch)
   END IF
   FUNCTION = nLen
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of submatches.
' ========================================================================================
PRIVATE FUNCTION CRegExp.SubMatchesCount (BYVAL index AS LONG) AS LONG
   IF m_pMatches = NULL THEN this.SetResult(E_POINTER): RETURN 0
   DIM nCount AS LONG
   DIM pMatch AS Afx_IMatch2 PTR
   this.SetResult(m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch)))
   IF pMatch THEN
      DIM pSubMatches AS Afx_ISubMatches PTR
      this.SetResult(pMatch->get_SubMatches(cast(Afx_IDispatch PTR PTR, @pSubMatches)))
      AfxSafeRelease(pMatch)
      IF pSubMatches THEN
         this.SetResult(pSubMatches->get_Count(@nCount))
         AfxSafeRelease(pSubMatches)
      END IF
   END IF
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value or text of a submatch found in a search string.
' ========================================================================================
PRIVATE FUNCTION CRegExp.SubMatchValue (BYVAL MatchIndex AS LONG, BYVAL SubMatchIndex AS LONG) AS CVAR
   IF m_pMatches = NULL THEN this.SetResult(E_POINTER): RETURN 0
   DIM vValue AS VARIANT
   DIM pMatch AS Afx_IMatch2 PTR
   this.SetResult(m_pMatches->get_Item(MatchIndex, cast(Afx_IDispatch PTR PTR, @pMatch)))
   IF pMatch THEN
      DIM pSubMatches AS Afx_ISubMatches PTR
      this.SetResult(pMatch->get_SubMatches(cast(Afx_IDispatch PTR PTR, @pSubMatches)))
      AfxSafeRelease(pMatch)
      IF pSubMatches THEN
         this.SetResult(pSubMatches->get_Item(SubMatchIndex, @vValue))
         AfxSafeRelease(pSubMatches)
      END IF
   END IF
   RETURN vValue
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Executes a regular expression search against a specified string and returns a Boolean
' value that indicates if a pattern match was found.
' - cbsSourceString : The source string.
' Note: The Global property has no effect on the Test method.
' Example:
' DIM pRegExp AS CRegExp
' pRegExp.Pattern = "is."
' pRegExp.IgnoreCase = TRUE
' IF pRegExp.Test("IS1 is2 IS3 is4") = FALSE THEN
'   PRINT "No match found"
' ELSE
'    PRINT "Match found"
' END IF
' ========================================================================================
PRIVATE FUNCTION CRegExp.Test (BYREF cbsSourceString AS CBSTR) AS BOOLEAN
   DIM bMatch AS VARIANT_BOOL
   this.SetResult(m_pRegExp->Test(cbsSourceString, @bMatch))
   RETURN bMatch
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Find function with VBScript regular expressions search patterns.
' Parameter:
' - cbsSourceString = The text to be parsed.
' Return value:
' - Returns the position of the match or 0 if not found.
'   The length of the match can be retrieved calling pRegExp.MatchLen.
' Usage example:
'   DIM pRegExp AS CRegExp
'   DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
'   pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
'   DIM nPos AS LONG = pRegExp.Find(cbsText)
'   PRINT nPos
' ========================================================================================
PRIVATE FUNCTION CRegExp.Find (BYREF cbsSourceString AS CBSTR) AS LONG
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   RETURN this.MatchPos(0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Find function with VBScript regular expressions search patterns.
' Parameters:
' - nStart: The position in the string at which the search will begin. The first character
'   starts at position 1.
' - cbsSourceString = The text to be parsed.
' Return value:
' - Returns the position of the match or 0 if not found.
'   The length of the match can be retrieved calling pRegExp.MatchLen.
' Usage example:
'   DIM pRegExp AS CRegExp
'   DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
'   pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
'   DIM nPos AS LONG = pRegExp.Find(15, cbsText)
' ========================================================================================
PRIVATE FUNCTION CRegExp.Find (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR) AS LONG
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   DIM pbstr AS BSTR
   IF nStart < 0 OR nStart > SysStringLen(cbsSourceString.m_bstr) THEN RETURN 0
   DIM cbs AS CBSTR = cbsSourceString.MidChars(nStart)
   this.SetResult(m_pRegExp->Execute(cbs, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   RETURN this.MatchPos(0) + nStart - 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Global, multiline find function with VBScript regular expressions search patterns.
' Parameters:
' - cbsSourceString = The text to be parsed.
' - cbsPattern = The pattern to match.
' Return value:
' - Returns a list of comma separated "index, length" value pairs. The pairs are separated
'   by a semicolon.
' Usage Example:
'   DIM pRegExp AS CRegExp
'   DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
'   pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
'   DIM cbsOut AS CBSTR = pRegExp.FindEx(cbsText)
'   PRINT cbsOut
' ========================================================================================
PRIVATE FUNCTION CRegExp.FindEx (BYREF cbsSourceString AS CBSTR) AS CBSTR
   DIM cbsOut AS CBSTR, pMatches AS Afx_IMatchCollection2 PTR
   m_pRegExp->put_Global(TRUE)
   m_pRegExp->put_Multiline(TRUE)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @pMatches)))
   IF pMatches THEN
      DIM nCount AS LONG
      pMatches->get_Count(@nCount)
      FOR i AS LONG = 0 TO nCount - 1
         DIM pMatch AS Afx_IMatch2 PTR
         this.SetResult(pMatches->get_Item(i, cast(Afx_IDispatch PTR PTR, @pMatch)))
         IF pMatch THEN
            DIM nFirstIndex AS LONG
            pMatch->get_FirstIndex(@nFirstIndex)
            DIM nLen AS LONG
            pMatch->get_Length(@nLen)
            IF i < nCount - 1 THEN
               cbsOut += STR(nFirstIndex + 1) & "," & STR(nLen) & ";"
            ELSE
               cbsOut += STR(nFirstIndex + 1) & "," & STR(nLen)
            END IF
            AfxSafeRelease(pMatch)
         END IF
      NEXT
      AfxSafeRelease(pMatches)
   END IF
   RETURN cbsOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a copy of a string with text removed using a regular expression as the search string.
' Parameters:
' - cbsSourceString : The source string.
' Examples:
' DIM pRegExp AS CRegExp
' pRegExp.Pattern = "ab"
' pRegExp.Global = TRUE
' PRINT pRegExp.Remove("abacadabra")
' Output: "acadra"
' ----------------------------------------------------------------------------------------
' DIM pRegExp AS CRegExp
' pRegExp.Pattern = "[bAc]"
' pRegExp.IgnoreCase = TRUE
' pRegExp.Global = TRUE
' PRINT pRegExp.Remove("abacadabra")
' Output: "dr"
' ========================================================================================
PRIVATE FUNCTION CRegExp.Remove (BYREF cbsSourceString AS CBSTR) AS CBSTR
   DIM bstrDestString AS AFX_BSTR
   this.SetResult(m_pRegExp->Replace(cbsSourceString, CVAR(""), @bstrDestString))
   RETURN bstrDestString
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a copy of a string with text replaced using a regular expression as the search string.
' Parameters:
' - cbsSourceString : The source string.
' - cvReplaceString : The replacement string.
' Examples:
' DIM pRegExp AS CRegExp
' pRegExp.Pattern = "World"
' PRINT pRegExp.Replace("Hello World", "Earth") - prints "Hello Earth"
' ----------------------------------------------------------------------------------------
' pRegExp.Pattern = "[bac]"
' pRegExp.Global = TRUE
' PRINT pRegExp.Replace("abacadabra", "*") - prints "*****d**r*"
' ----------------------------------------------------------------------------------------
' pRegExp.Pattern = "(\d{3})-(\d{3})-(\d{4})"
' PRINT pRegExp.Replace("555-123-4567", "($1) $2-$3") - prints "(555) 123-4567"
' ----------------------------------------------------------------------------------------
' pRegExp.Pattern = "(\S+), (\S+)"
' PRINT pRegExp.Replace("Squires, Paul", "$2 $1") - prints "Paul Squires"
' ----------------------------------------------------------------------------------------
' pRegExp.Pattern = $"\b0{1,}\."
' PRINT pRegExp.Replace("0000.34500044", ".") - prints ".34500044"
' ========================================================================================
PRIVATE FUNCTION CRegExp.Replace (BYREF cbsSourceString AS CBSTR, BYREF cvReplaceString AS CVAR) AS CBSTR
   DIM bstrDestString AS AFX_BSTR
   this.SetResult(m_pRegExp->Replace(cbsSourceString, cvReplaceString, @bstrDestString))
   RETURN bstrDestString
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a substring using VBScript regular expressions search patterns.
' Parameters:
' - cbsSourceString = The text to be parsed.
' Return value:
' - Returns the retrieved string on exit or an empty string on failure.
' Usage example:
'   DIM pRegExp AS CRegExp
'   DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
'   pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
'   DIM cbs AS CBSTR = pRegExp.Extract(cbsText)
'   PRINT cbs   ' prints "a234"
' ========================================================================================
PRIVATE FUNCTION CRegExp.Extract (BYREF cbsSourceString AS CBSTR) AS CBSTR
   DIM cbs AS CBSTR
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   DIM nPos AS LONG = this.MatchPos(0)
   DIM nLen AS LONG = this.MatchLen(0)
   cbs = MID(cbsSourceString, nPos, nLen)
   RETURN cbs
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a substring using VBScript regular expressions search patterns.
' Parameters:
' - nStart: The position in the string at which the search will begin. The first character
'   starts at position 1.
' - cbsSourceString = The text to be parsed.
' Return value:
' - Returns the retrieved string on exit or an empty string on failure.
' Usage example:
'   DIM pRegExp AS CRegExp
'   DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
'   pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
'   DIM cbs AS CBSTR = pRegExp.Extract(15, cbsText)
'   PRINT cbs   ' prints "x345"
' ========================================================================================
PRIVATE FUNCTION CRegExp.Extract (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR) AS CBSTR
   DIM cbs AS CBSTR
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   DIM pbstr AS BSTR
   IF nStart < 0 OR nStart > SysStringLen(cbsSourceString.m_bstr) THEN RETURN cbs
   cbs = cbsSourceString.MidChars(nStart)
   this.SetResult(m_pRegExp->Execute(cbs, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   DIM nPos AS LONG = this.MatchPos(0)
   DIM nLen AS LONG = this.MatchLen(0)
   RETURN MID(cbsSourceString, nPos + nStart - 1, nLen)
END FUNCTION
' ========================================================================================

END NAMESPACE
