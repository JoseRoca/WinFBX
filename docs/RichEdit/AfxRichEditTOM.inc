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

DECLARE FUNCTION AfxRichEditTOM_ChangeCase (BYVAL hRichEdit AS HWND, BYVAL nType AS LONG) AS HRESULT
DECLARE FUNCTION AfxRichEditTOM_GetText (BYVAL hRichEdit AS HWND, BYVAL cpActive AS LONG = 0, BYVAL cpAnchor AS LONG = 0, BYVAL Flags AS LONG = 0) AS CBSTR
DECLARE FUNCTION AfxRichEditTOM_GetCch (BYVAL hRichEdit AS HWND, BYVAL cpActive AS LONG = 0, BYVAL cpAnchor AS LONG = 0) AS LONG

' ========================================================================================
' Gets the plain text in this range.
' ========================================================================================
PRIVATE FUNCTION AfxRichEditTOM_GetText (BYVAL hRichEdit AS HWND, BYVAL cpActive AS LONG = 0, BYVAL cpAnchor AS LONG = 0, BYVAL Flags AS LONG = 0) AS CBSTR
   CTOM_DP("AfxRichEditTOM_GetText")
   DIM cbsText AS CBSTR
   ' // Create an instance of the CTextDocument2 class
   DIM pCTextDoc AS CTextDocument2 = hRichEdit
   IF pCTextDoc.TextDocumentPtr = NULL THEN RETURN cbsText   ' // Return an empty string
   ' // Create a range
   IF cpActive < 0 THEN cpActive = 0
   DIM numChars AS LONG = RichEdit_GetTextLength(hRichEdit)   ' // Length of the text
   IF cpAnchor < 0 THEN cpAnchor = numChars
   IF cpAnchor > numChars THEN cpAnchor = numChars
   ' // Get the text
   DIM pCRange2 AS CTextRange2 = pCTextDoc.Range2(cpActive, cpAnchor)
   IF pCRange2.TextRangePtr = NULL THEN RETURN cbsText   ' // Return an empty string
   cbsText = pCRange2.GetText2(Flags)
   RETURN cbsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the case of letters in this range according to the nType parameter.
' ========================================================================================
PRIVATE FUNCTION AfxRichEditTOM_ChangeCase (BYVAL hRichEdit AS HWND, BYVAL nType AS LONG) AS HRESULT
   CTOM_DP("AfxRichEditTOM_ChangeCase")
   ' // Create an instance of the CTextDocument2 class
   DIM pCTextDoc AS CTextDocument2 = hRichEdit
   IF pCTextDoc.TextDocumentPtr = NULL THEN RETURN S_FALSE
   ' // Get the selection
   DIM chrRange AS CHARRANGE
   RichEdit_ExGetSel(hRichEdit, @chrRange)
   IF chrRange.cpMin = chrRange.cpMax THEN RETURN S_FALSE
   ' // Create a range
   DIM pCRange2 AS CTextRange2 = pCTextDoc.Range2(chrRange.cpMin, chrRange.cpMax)
   IF pCRange2.TextRangePtr = NULL THEN RETURN S_FALSE
   ' // Change the case of the selected text according to the nType parameter
   DIM hr AS HRESULT = pCRange2.ChangeCase(nType)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the count of characters in a range.
' Returns the signed count of characters.
' ========================================================================================
PRIVATE FUNCTION AfxRichEditTOM_GetCch (BYVAL hRichEdit AS HWND, BYVAL cpActive AS LONG = 0, BYVAL cpAnchor AS LONG = 0) AS LONG
   CTOM_DP("AfxRichEditTOM_GetCch")
   ' // Create an instance of the CTextDocument2 class
   DIM pCTextDoc AS CTextDocument2 = hRichEdit
   IF pCTextDoc.TextDocumentPtr = NULL THEN RETURN S_FALSE
   ' // Create a range
   DIM pCRange2 AS CTextRange2 = pCTextDoc.Range2(cpActive, cpAnchor)
   IF pCRange2.TextRangePtr = NULL THEN RETURN S_FALSE
   ' // Gets the count of characters in the range.
   DIM cch AS LONG = pCRange2.GetCch
   RETURN cch
END FUNCTION
' ========================================================================================
