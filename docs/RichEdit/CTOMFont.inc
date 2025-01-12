' ########################################################################################
' Microsoft Windows
' File: CTOMFont.inc
' Contents: Class to work with the Richd Edit Text Object Model (TOM).
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' // To be included by CTOM.inc

' ########################################################################################
' CTextFont2 class
' ########################################################################################
TYPE CTextFont2

Public:
   m_Result AS HRESULT
   m_pTextFont2 AS ITextFont2 PTR

Public:
   DECLARE CONSTRUCTOR (BYVAL pTextFont2 AS ITextFont2 PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo () AS CWSTR
   DECLARE OPERATOR LET (BYVAL pTextFont2 AS ITextFont2 PTR)
   DECLARE OPERATOR CAST () AS ITextFont2 PTR
   DECLARE FUNCTION TextFont2Ptr () AS ITextFont2 PTR
   DECLARE FUNCTION Attach (BYVAL pTextFont2 AS ITextFont2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Detach () AS ITextFont2 PTR

   ' // Inherited ITextFont interface methods
   DECLARE FUNCTION GetDuplicate () AS ITextFont2 PTR
   DECLARE FUNCTION SetDuplicate (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   DECLARE FUNCTION CanChange () AS LONG
   DECLARE FUNCTION IsEqual (BYVAL pFont AS ITextFont2 PTR) AS LONG
   DECLARE FUNCTION Reset (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetStyle () AS LONG
   DECLARE FUNCTION SetStyle (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetAllCaps () AS LONG
   DECLARE FUNCTION SetAllCaps (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetAnimation () AS LONG
   DECLARE FUNCTION SetAnimation (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetBackColor () AS LONG
   DECLARE FUNCTION SetBackColor (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetBold () AS LONG
   DECLARE FUNCTION SetBold (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetEmboss () AS LONG
   DECLARE FUNCTION SetEmboss (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetForeColor () AS LONG
   DECLARE FUNCTION SetForeColor (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetHidden () AS LONG
   DECLARE FUNCTION SetHidden (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetEngrave () AS LONG
   DECLARE FUNCTION SetEngrave (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetItalic () AS LONG
   DECLARE FUNCTION SetItalic (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetKerning () AS SINGLE
   DECLARE FUNCTION SetKerning (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE FUNCTION GetLanguageID () AS LONG
   DECLARE FUNCTION SetLanguageID (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetName () AS CBSTR
   DECLARE FUNCTION SetName (BYVAL bstr AS BSTR) AS HRESULT
   DECLARE FUNCTION GetOutline () AS LONG
   DECLARE FUNCTION SetOutline (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetPosition () AS SINGLE
   DECLARE FUNCTION SetPosition (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE FUNCTION GetProtected () AS LONG
   DECLARE FUNCTION SetProtected (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetShadow () AS LONG
   DECLARE FUNCTION SetShadow (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetSize () AS SINGLE
   DECLARE FUNCTION SetSize (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE FUNCTION GetSmallCaps () AS LONG
   DECLARE FUNCTION SetSmallCaps (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetSpacing () AS SINGLE
   DECLARE FUNCTION SetSpacing (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE FUNCTION GetStrikeThrough () AS LONG
   DECLARE FUNCTION SetStrikeThrough (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetSubscript () AS LONG
   DECLARE FUNCTION SetSubscript (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetSuperscript () AS LONG
   DECLARE FUNCTION SetSuperscript (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetUnderline () AS LONG
   DECLARE FUNCTION SetUnderline (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetWeight () AS LONG
   DECLARE FUNCTION SetWeight (BYVAL Value AS LONG) AS HRESULT

   ' // ITextFont2 interface methods
   DECLARE FUNCTION GetCount () AS LONG
   DECLARE FUNCTION GetAutoLigatures () AS LONG
   DECLARE FUNCTION SetAutoLigatures (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetAutospaceAlpha () AS LONG
   DECLARE FUNCTION SetAutospaceAlpha (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetAutospaceNumeric () AS LONG
   DECLARE FUNCTION SetAutospaceNumeric (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetAutospaceParens () AS LONG
   DECLARE FUNCTION SetAutospaceParens (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCharRep () AS LONG
   DECLARE FUNCTION SetCharRep (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCompressionMode () AS LONG
   DECLARE FUNCTION SetCompressionMode (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCookie () AS LONG
   DECLARE FUNCTION SetCookie (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetDoubleStrike () AS LONG
   DECLARE FUNCTION SetDoubleStrike (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetDuplicate2 () AS ITextFont2 PTR
   DECLARE FUNCTION SetDuplicate2 (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   DECLARE FUNCTION GetLinkType () AS LONG
   DECLARE FUNCTION GetMathZone () AS LONG
   DECLARE FUNCTION SetMathZone (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetModWidthPairs () AS LONG
   DECLARE FUNCTION SetModWidthPairs (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetModWidthSpace () AS LONG
   DECLARE FUNCTION SetModWidthSpace (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetOldNumbers () AS LONG
   DECLARE FUNCTION SetOldNumbers (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetOverlapping () AS LONG
   DECLARE FUNCTION SetOverlapping (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetPositionSubSuper () AS LONG
   DECLARE FUNCTION SetPositionSubSuper (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetScaling () AS LONG
   DECLARE FUNCTION SetScaling (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetSpaceExtension () AS LONG
   DECLARE FUNCTION SetSpaceExtension (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetUnderlinePositionMode () AS LONG
   DECLARE FUNCTION SetUnderlinePositionMode (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetEffects (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   DECLARE FUNCTION GetEffects2 (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   DECLARE FUNCTION GetProperty (BYVAL nType AS LONG) AS LONG
   DECLARE FUNCTION GetPropertyInfo (BYVAL Index AS LONG, BYVAL pType AS LONG PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   DECLARE FUNCTION IsEqual2 (BYVAL pFont AS ITextFont2 PTR) AS LONG
   DECLARE FUNCTION SetEffects (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   DECLARE FUNCTION SetEffects2 (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   DECLARE FUNCTION SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CTextRange2 constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CTextFont2 (BYVAL pTextFont2 AS ITextFont2 PTR)
   ' // Store the pointer of ITextFont2 interface
   IF pTextFont2 = NULL THEN m_Result = E_INVALIDARG : EXIT CONSTRUCTOR
   m_pTextFont2 = pTextFont2
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTextRange2 destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTextFont2
   ' // Release the ITextFont2 interface
   IF m_pTextFont2 THEN m_pTextFont2->lpvtbl->Release(m_pTextFont2)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assignment operator
' ========================================================================================
PRIVATE OPERATOR CTextFont2.LET (BYVAL pTextFont2 AS ITextFont2 PTR)
   m_Result = 0
   IF pTextFont2 = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextFont2 THEN m_pTextFont2->lpvtbl->Release(m_pTextFont2)
   ' // Attach the passed interface pointer to the class
   m_pTextFont2 = pTextFont2
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operator
' ========================================================================================
PRIVATE OPERATOR CTextFont2.CAST () AS ITextFont2 PTR
   m_Result = 0
   OPERATOR = m_pTextFont2
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the underlying ITextFont2 interface
' ========================================================================================
PRIVATE FUNCTION CTextFont2.TextFont2Ptr () AS ITextFont2 PTR
   m_Result = 0
   RETURN m_pTextFont2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches the passed text document object to the class
' ========================================================================================
PRIVATE FUNCTION CTextFont2.Attach (BYVAL pTextFont2 AS ITextFont2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextFont2 = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextFont2 THEN m_Result = m_pTextFont2->lpvtbl->Release(m_pTextFont2)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextFont2->lpvtbl->AddRef(pTextFont2)
   m_pTextFont2 = pTextFont2
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the interface pointer from the class
' ========================================================================================
PRIVATE FUNCTION CTextFont2.Detach () AS ITextFont2 PTR
   m_Result = 0
   DIM pTextFont2 AS ITextFont2 PTR = m_pTextFont2
   m_pTextFont2 = NULL
   RETURN pTextFont2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a duplicate of this text font object.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetDuplicate () AS ITextFont2 PTR
   DIM pFont AS ITextFont2 PTR
   this.SetResult(m_pTextFont2->lpvtbl->GetDuplicate(m_pTextFont2, @pFont))
   FUNCTION = pFont
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the character formatting by copying another text font object.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetDuplicate (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetDuplicate(m_pTextFont2, pFont))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the font can be changed.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.CanChange () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->CanChange(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether this text font object has the same properties as the specified text font object.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.IsEqual (BYVAL pFont AS ITextFont2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->IsEqual(m_pTextFont2, pFont, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets the character formatting to the specified values.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.Reset (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->Reset(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character style handle of the characters in a range.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetStyle () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetStyle(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the character style handle of the characters in a range.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetStyle (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetStyle(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether the characters are all uppercase.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAllCaps () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAllCaps(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether the characters are all uppercase.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAllCaps (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAllCaps(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the animation type.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAnimation () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAnimation(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the animation type.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAnimation (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAnimation(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text background (highlight) color.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetBackColor () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetBackColor(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetBackColor (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetBackColor(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether the characters are bold.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetBold () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetBold(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are bold.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetBold (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetBold(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are embossed.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetEmboss () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetEmboss(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are embossed.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetEmboss (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetEmboss(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the foreground, or text, color.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetForeColor () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetForeColor(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the foreground (text) color.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetForeColor (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetForeColor(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are hidden.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetHidden () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetHidden(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are hidden.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetHidden (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetHidden(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as imprinted characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetEngrave () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetEngrave(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are displayed as imprinted characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetEngrave (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetEngrave(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are in italics.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetItalic () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetItalic(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are in italics.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetItalic (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetItalic(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the minimum font size at which kerning occurs.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetKerning () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextFont2->lpvtbl->GetKerning(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the minimum font size at which kerning occurs.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetKerning (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetKerning(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the language ID or language code identifier (LCID).
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetLanguageID () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetLanguageID(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the language ID or language code identifier (LCID).
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetLanguageID (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetLanguageID(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the font name.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetName () AS CBSTR
   DIM pName AS AFX_BSTR
   this.SetResult(m_pTextFont2->lpvtbl->GetName(m_pTextFont2, @pName))
   RETURN pName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font name.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetName (BYVAL bstr AS BSTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetName(m_pTextFont2, bstr))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as outlined characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetOutline () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetOutline(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are displayed as outlined characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetOutline (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetOutline(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the amount that characters are offset vertically relative to the baseline.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetPosition () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextFont2->lpvtbl->GetPosition(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the amount that characters are offset vertically relative to the baseline.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetPosition (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetPosition(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are protected against attempts to modify them.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetProtected () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetProtected(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are protected against attempts to modify them.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetProtected (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetProtected(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as shadowed characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetShadow () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetShadow(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are displayed as shadowed characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetShadow (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetShadow(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the font size.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSize () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextFont2->lpvtbl->GetSize(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font size.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSize (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSize(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are in small capital letters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSmallCaps () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetSmallCaps(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are in small capital letters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSmallCaps (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSmallCaps(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the amount of horizontal spacing between characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSpacing () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextFont2->lpvtbl->GetSpacing(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the amount of horizontal spacing between characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSpacing (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSpacing(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed with a horizontal line through the center.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetStrikeThrough () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetStrikeThrough(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are displayed with a horizontal line through the center.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetStrikeThrough (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetStrikeThrough(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as subscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSubscript () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetSubscript(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are displayed as subscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSubscript (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSubscript(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSuperscript () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetSuperscript(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSuperscript (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSuperscript(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetUnderline () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetUnderline(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetUnderline (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetUnderline(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the font weight for the characters in a range.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetWeight () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetWeight(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font weight for the characters in a range.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetWeight (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetWeight(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the count of extra properties in this character formatting collection.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetCount () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetCount(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether support for automatic ligatures is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAutoLigatures () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAutoLigatures(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether support for automatic ligatures is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAutoLigatures (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAutoLigatures(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the East Asian "autospace alphabetics" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAutospaceAlpha () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAutospaceAlpha(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the East Asian "autospace alpha" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAutospaceAlpha (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAutospaceAlpha(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the East Asian "autospace numeric" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAutospaceNumeric () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAutospaceNumeric(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the East Asian "autospace numeric" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAutospaceNumeric (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAutospaceNumeric(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the East Asian "autospace parentheses" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAutospaceParens () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAutospaceParens(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the East Asian "autospace parentheses" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAutospaceParens (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAutospaceParens(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character repertoire (writing system).
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetCharRep () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetCharRep(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the character repertoire (writing system).
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetCharRep (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetCharRep(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the East Asian compression mode.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetCompressionMode () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetCompressionMode(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the East Asian compression mode.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetCompressionMode (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetCompressionMode(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the client cookie.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetCookie () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetCookie(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the client cookie.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetCookie (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetCookie(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed with double horizontal lines through the center.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetDoubleStrike () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetDoubleStrike(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are displayed with double horizontal lines through the center.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetDoubleStrike (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetDoubleStrike(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a duplicate of this character format object.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetDuplicate2 () AS ITextFont2 PTR
   DIM pFont AS ITextFont2 PTR
   this.SetResult(m_pTextFont2->lpvtbl->GetDuplicate2(m_pTextFont2, @pFont))
   FUNCTION = pFont
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether characters are displayed with double horizontal lines through the center.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetDuplicate2 (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetDuplicate2(m_pTextFont2, pFont))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the link type.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetLinkType () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetLinkType(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether a math zone is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetMathZone () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetMathZone(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether a math zone is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetMathZone (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetMathZone(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether "decrease widths on pairs" is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetModWidthPairs () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetModWidthPairs(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether "decrease widths on pairs" is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetModWidthPairs (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetModWidthPairs(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether "increase width of whitespace" is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetModWidthSpace () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetModWidthSpace(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether "increase width of whitespace" is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetModWidthSpace (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetModWidthSpace(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether old-style numbers are active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetOldNumbers () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetOldNumbers(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether old-style numbers are active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetOldNumbers (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetOldNumbers(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether overlapping text is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetOverlapping () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetOverlapping(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether overlapping text is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetOverlapping (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetOverlapping(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the subscript or superscript position relative to the baseline.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetPositionSubSuper () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetPositionSubSuper(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the position of a subscript or superscript relative to the baseline, as a percentage
' of the font height.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetPositionSubSuper (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetPositionSubSuper(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the font horizontal scaling percentage.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetScaling () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetScaling(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font horizontal scaling percentage.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetScaling (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetScaling(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the East Asian space extension value.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSpaceExtension () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetSpaceExtension(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the East Asian space extension value.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSpaceExtension (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSpaceExtension(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the underline position mode.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetUnderlinePositionMode () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetUnderlinePositionMode(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the underline position mode.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetUnderlinePositionMode (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetUnderlinePositionMode(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetEffects (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->GetEffects(m_pTextFont2, pValue, pMask))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the additional character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetEffects2 (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->GetEffects2(m_pTextFont2, pValue, pMask))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the value of the specified property.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetProperty(m_pTextFont2, nType, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the additional character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetPropertyInfo (BYVAL Index AS LONG, BYVAL pType AS LONG PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->GetPropertyInfo(m_pTextFont2, Index, pType, pValue))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether this text font object has the same properties as the specified text font object.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.IsEqual2 (BYVAL pFont AS ITextFont2 PTR) AS LONG
   DIM B AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->IsEqual2(m_pTextFont2, pFont, @B))
   FUNCTION = B
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetEffects (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetEffects(m_pTextFont2, Value, Mask))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the additional character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetEffects2 (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetEffects2(m_pTextFont2, Value, Mask))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the additional character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetProperty(m_pTextFont2, nType, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================
