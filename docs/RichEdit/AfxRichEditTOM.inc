' ########################################################################################
' Microsoft Windows
' File: AfxRichEditTOM.inc
' Contents: Rich Edit control wrapper functions using the Text Object Model (TOM) library
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "Afx/CTOM.inc"

DECLARE FUNCTION AfxRichEditTOM_GetText (BYVAL hRichEdit AS HWND, BYVAL cpActive AS LONG = 0, BYVAL cpAnchor AS LONG = 0) AS CBSTR

' ========================================================================================
' Gets the plain text in this range.
' ========================================================================================
PRIVATE FUNCTION AfxRichEditTOM_GetText (BYVAL hRichEdit AS HWND, BYVAL cpActive AS LONG = 0, BYVAL cpAnchor AS LONG = 0) AS CBSTR
   ' // Create an instance of the CTextDocument2 class
   DIM cbsText AS CBSTR
   DIM pCTextDoc AS CTextDocument2 = hRichEdit
   IF pCTextDoc THEN
      ' // Create a range
      IF cpActive < 0 THEN cpActive = 0
      DIM numChars AS LONG = RichEdit_GetTextLength(hRichEdit)   ' // Length of the text
      IF cpAnchor < 0 THEN cpAnchor = numChars
      IF cpAnchor > numChars THEN cpAnchor = numChars
      ' // Get the text
      DIM pCRange2 AS CTextRange2 = pCTextDoc.Range2(cpActive, cpAnchor)
      IF pCRange2 THEN cbsText = pCRange2.GetText2(0)
   END IF
   RETURN cbsText
END FUNCTION
' ========================================================================================

