' ########################################################################################
' Microsoft Windows
' File: CTOM.inc
' Contents: Classes to work with the Richd Edit Text Object Model (TOM).
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "Afx/CWSTR.inc"
#include once "Afx/AfxTOM.bi"

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
' CTOMBase class
' ########################################################################################
TYPE CTOMBase

   DIM m_Result AS HRESULT
   DIM m_bUninitCOM AS BOOLEAN

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo () AS CWSTR

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
CONSTRUCTOR CTOMBase
   CTOM_DP("*** BEGIN CTOMBase CONSTRUCTOR")
   ' // Initialize the COM library
   IF m_bUninitCOM = FALSE THEN
      DIM hr AS HRESULT = CoInitialize(NULL)
      IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   END IF
   CTOM_DP("*** END CTOMBase CONSTRUCTOR - CoInitialize")
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Destructor
' ===========================================================================================
DESTRUCTOR CTOMBase
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
   CTOM_DP("*** CTOMBase DESTRUCTOR - CoUninitialize")
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the last result
' ========================================================================================
PRIVATE FUNCTION CTOMBase.GetLastResult () AS HRESULT
   CTOM_DP("CTOMBase.GetLastResult - Result = " & HEX(m_Result))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CTOMBase.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   CTOM_DP("CTOMBase.SetLastResult - Result = " & HEX(Result))
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CTOMBase.GetErrorInfo () AS CWSTR
   IF SUCCEEDED(m_Result) THEN RETURN "Success"
   DIM s AS CWSTR = "Error &h" & HEX(m_Result, 8)
   SELECT CASE m_Result
      CASE E_POINTER : s += ": E_POINTER - Null pointer"
      CASE S_OK : s += ": S_OK - Success"
      CASE S_FALSE : s += ": S_FALSE - Failure"
      CASE E_NOTIMPL : s += ": E_NOTIMPL - Not implemented."
      CASE E_INVALIDARG : s += ": E_INVALIDARG - Invalid argument"
      CASE E_OUTOFMEMORY : s += ": E_OUTOFMEMORY - Insufficient memory"
'      CASE E_FILENOTFOUND : s += "E_FILENOTFOUND - File not found"
      CASE &h80070002 : s += "E_FILENOTFOUND - File not found"
      CASE E_ACCESSDENIED : s += "E_ACCESSDENIED - Access denied"
      CASE E_FAIL : s += ": E_FAIL - Access denied"
      CASE NOERROR : s += ": NOERROR - Success" '' (same as S_OK)
      CASE CO_E_RELEASED:  : s += ": CO_E_RELEASED: - The object has been released"
      CASE ELSE
         s += "Unknown error"
   END SELECT
   RETURN s

END FUNCTION
' ========================================================================================

#include once "Afx/CTOMDocument.inc"
#include once "Afx/CTOMRange.inc"
#include once "Afx/CTOMPara.inc"
#include once "Afx/CTOMSelection.inc"
#include once "Afx/CTOMFont.inc"
#include once "Afx/CTOMRow.inc"

END NAMESPACE
