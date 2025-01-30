' ########################################################################################
' Microsoft Windows
' File: CTextObjectBase.inc
' Contents: Base class to work with the Rich Edit control.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CTOM_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CTOM_DEBUG_
   #define _CTOM_DEBUG_ 0
#ENDIF
#ifndef _CTOM_DP_
   #define _CTOM_DP_ 1
   #MACRO CTOM_DP(st)
      #IF (_CTOM_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ########################################################################################
' CTextObjectBase class
' ########################################################################################
TYPE CTextObjectBase

   DIM m_Result AS LONG

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS LONG
   DECLARE FUNCTION SetResult (BYVAL Result AS LONG) AS LONG
   DECLARE FUNCTION GetErrorInfo (BYVAL dwError AS LONG = -1) AS CWSTR

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
CONSTRUCTOR CTextObjectBase
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Destructor
' ===========================================================================================
DESTRUCTOR CTextObjectBase
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the last result code
' ========================================================================================
PRIVATE FUNCTION CTextObjectBase.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CTextObjectBase.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CTextObjectBase.GetErrorInfo (BYVAL nError AS LONG = -1) AS CWSTR
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

END NAMESPACE
