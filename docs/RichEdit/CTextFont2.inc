' ########################################################################################
' Microsoft Windows
' File: CTextFont2.inc
' Contents: Class to work with the Richd Edit Text Object Model (TOM).
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' // To be included by CRichEditCtrl.inc

NAMESPACE Afx

' ########################################################################################
' CTextFont2 class
' ########################################################################################
TYPE CTextFont2 EXTENDS CTextObjectBase

Public:
   m_pTextFont2 AS ITextFont2 PTR

Public:
   DECLARE CONSTRUCTOR (BYVAL pTextFont2 AS ITextFont2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
   ' // Inherited ITextFont interface methods
   DECLARE FUNCTION GetDuplicate () AS ITextFont2 PTR
   DECLARE FUNCTION SetDuplicate (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   DECLARE PROPERTY Duplicate () AS ITextFont2 PTR
   DECLARE PROPERTY Duplicate (BYVAL pFont AS ITextFont2 PTR)
   DECLARE FUNCTION CanChange () AS LONG
   DECLARE FUNCTION IsEqual (BYVAL pFont AS ITextFont2 PTR) AS LONG
   DECLARE FUNCTION Reset (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetStyle () AS LONG
   DECLARE FUNCTION SetStyle (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Style () AS LONG
   DECLARE PROPERTY Style (BYVAL Value AS LONG)
   DECLARE FUNCTION GetAllCaps () AS LONG
   DECLARE FUNCTION SetAllCaps (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY AllCaps () AS LONG
   DECLARE PROPERTY AllCaps (BYVAL Value AS LONG)
   DECLARE FUNCTION GetAnimation () AS LONG
   DECLARE FUNCTION SetAnimation (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Animation () AS LONG
   DECLARE PROPERTY Animation (BYVAL Value AS LONG)
   DECLARE FUNCTION GetBackColor () AS LONG
   DECLARE FUNCTION SetBackColor (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY BackColor (BYVAL Value AS LONG)
   DECLARE PROPERTY BackColor () AS LONG
   DECLARE FUNCTION GetBold () AS LONG
   DECLARE FUNCTION SetBold (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Bold () AS LONG
   DECLARE PROPERTY Bold (BYVAL Value AS LONG)
   DECLARE FUNCTION GetEmboss () AS LONG
   DECLARE FUNCTION SetEmboss (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Emboss () AS LONG
   DECLARE PROPERTY Emboss (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetForeColor () AS LONG
   DECLARE FUNCTION SetForeColor (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY ForeColor () AS LONG
   DECLARE PROPERTY ForeColor (BYVAL Value AS LONG)
   DECLARE FUNCTION GetHidden () AS LONG
   DECLARE FUNCTION SetHidden (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Hidden () AS LONG
   DECLARE PROPERTY Hidden (BYVAL Value AS LONG)
   DECLARE FUNCTION GetEngrave () AS LONG
   DECLARE FUNCTION SetEngrave (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Engrave () AS LONG
   DECLARE PROPERTY Engrave (BYVAL Value AS LONG)
   DECLARE FUNCTION GetItalic () AS LONG
   DECLARE FUNCTION SetItalic (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Italic () AS LONG
   DECLARE PROPERTY Italic (BYVAL Value AS LONG)
   DECLARE FUNCTION GetKerning () AS SINGLE
   DECLARE FUNCTION SetKerning (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE PROPERTY Kerning () AS SINGLE
   DECLARE PROPERTY Kerning (BYVAL Value AS SINGLE)
   DECLARE FUNCTION GetLanguageID () AS LONG
   DECLARE FUNCTION SetLanguageID (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY LanguageID () AS LONG
   DECLARE PROPERTY LanguageID (BYVAL Value AS LONG)
   DECLARE FUNCTION GetName () AS CBSTR
   DECLARE FUNCTION SetName (BYVAL fontName AS AFX_BSTR) AS HRESULT
   DECLARE PROPERTY Name () AS CBSTR
   DECLARE PROPERTY Name (BYVAL fontName AS AFX_BSTR)
   DECLARE FUNCTION GetOutline () AS LONG
   DECLARE FUNCTION SetOutline (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Outline () AS LONG
   DECLARE PROPERTY Outline (BYVAL Value AS LONG)
   DECLARE FUNCTION GetPosition () AS SINGLE
   DECLARE FUNCTION SetPosition (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE PROPERTY Position () AS SINGLE
   DECLARE PROPERTY Position (BYVAL Value AS SINGLE)
   DECLARE FUNCTION GetProtected () AS LONG
   DECLARE FUNCTION SetProtected (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Protected_ () AS LONG
   DECLARE PROPERTY Protected_ (BYVAL Value AS LONG)
   DECLARE FUNCTION GetShadow () AS LONG
   DECLARE FUNCTION SetShadow (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Shadow () AS LONG
   DECLARE PROPERTY Shadow (BYVAL Value AS LONG)
   DECLARE FUNCTION GetSize () AS SINGLE
   DECLARE FUNCTION SetSize (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE PROPERTY Size () AS SINGLE
   DECLARE PROPERTY Size (BYVAL Value AS SINGLE)
   DECLARE FUNCTION GetSmallCaps () AS LONG
   DECLARE FUNCTION SetSmallCaps (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY SmallCaps () AS LONG
   DECLARE PROPERTY SmallCaps (BYVAL Value AS LONG)
   DECLARE FUNCTION GetSpacing () AS SINGLE
   DECLARE FUNCTION SetSpacing (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE PROPERTY Spacing () AS SINGLE
   DECLARE PROPERTY Spacing (BYVAL Value AS SINGLE)
   DECLARE FUNCTION GetStrikeThrough () AS LONG
   DECLARE FUNCTION SetStrikeThrough (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY StrikeThrough () AS LONG
   DECLARE PROPERTY StrikeThrough (BYVAL Value AS LONG)
   DECLARE FUNCTION GetSubscript () AS LONG
   DECLARE FUNCTION SetSubscript (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Subscript () AS LONG
   DECLARE PROPERTY Subscript (BYVAL Value AS LONG)
   DECLARE FUNCTION GetSuperscript () AS LONG
   DECLARE FUNCTION SetSuperscript (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Superscript () AS LONG
   DECLARE PROPERTY Superscript (BYVAL Value AS LONG)
   DECLARE FUNCTION GetUnderline () AS LONG
   DECLARE FUNCTION SetUnderline (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Underline () AS LONG
   DECLARE PROPERTY Underline (BYVAL Value AS LONG)
   DECLARE FUNCTION GetWeight () AS LONG
   DECLARE FUNCTION SetWeight (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Weight () AS LONG
   DECLARE PROPERTY Weight (BYVAL Value AS LONG)
   ' // ITextFont2 interface methods
   DECLARE FUNCTION GetCount () AS LONG
   DECLARE FUNCTION Count () AS LONG
   DECLARE FUNCTION GetAutoLigatures () AS LONG
   DECLARE FUNCTION SetAutoLigatures (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY AutoLigatures () AS LONG
   DECLARE PROPERTY AutoLigatures (BYVAL Value AS LONG)
   DECLARE FUNCTION GetAutospaceAlpha () AS LONG
   DECLARE FUNCTION SetAutospaceAlpha (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY AutospaceAlpha () AS LONG
   DECLARE PROPERTY AutospaceAlpha (BYVAL Value AS LONG)
   DECLARE FUNCTION GetAutospaceNumeric () AS LONG
   DECLARE FUNCTION SetAutospaceNumeric (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY AutospaceNumeric () AS LONG
   DECLARE PROPERTY AutospaceNumeric (BYVAL Value AS LONG)
   DECLARE FUNCTION GetAutospaceParens () AS LONG
   DECLARE FUNCTION SetAutospaceParens (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY AutospaceParens () AS LONG
   DECLARE PROPERTY AutospaceParens (BYVAL Value AS LONG)
   DECLARE FUNCTION GetCharRep () AS LONG
   DECLARE FUNCTION SetCharRep (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY CharRep () AS LONG
   DECLARE PROPERTY CharRep (BYVAL Value AS LONG)
   DECLARE FUNCTION GetCompressionMode () AS LONG
   DECLARE FUNCTION SetCompressionMode (BYVAL Value AS LONG = tomCompressNone) AS HRESULT
   DECLARE PROPERTY CompressionMode () AS LONG
   DECLARE PROPERTY CompressionMode (BYVAL Value AS LONG)
   DECLARE FUNCTION GetCookie () AS LONG
   DECLARE FUNCTION SetCookie (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Cookie () AS LONG
   DECLARE PROPERTY Cookie (BYVAL Value AS LONG)
   DECLARE FUNCTION GetDoubleStrike () AS LONG
   DECLARE FUNCTION SetDoubleStrike (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY DoubleStrike () AS LONG
   DECLARE PROPERTY DoubleStrike (BYVAL Value AS LONG)
   DECLARE FUNCTION GetDuplicate2 () AS ITextFont2 PTR
   DECLARE FUNCTION SetDuplicate2 (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   DECLARE PROPERTY Duplicate2 () AS ITextFont2 PTR
   DECLARE PROPERTY Duplicate2 (BYVAL pFont AS ITextFont2 PTR)
   DECLARE FUNCTION GetLinkType () AS LONG
   DECLARE FUNCTION LinkType () AS LONG
   DECLARE FUNCTION GetMathZone () AS LONG
   DECLARE FUNCTION SetMathZone (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY MathZone () AS LONG
   DECLARE PROPERTY MathZone (BYVAL Value AS LONG)
   DECLARE FUNCTION GetModWidthPairs () AS LONG
   DECLARE FUNCTION SetModWidthPairs (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY ModWidthPairs () AS LONG
   DECLARE PROPERTY ModWidthPairs (BYVAL Value AS LONG)
   DECLARE FUNCTION GetModWidthSpace () AS LONG
   DECLARE FUNCTION SetModWidthSpace (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY ModWidthSpace () AS LONG
   DECLARE PROPERTY ModWidthSpace (BYVAL Value AS LONG)
   DECLARE FUNCTION GetOldNumbers () AS LONG
   DECLARE FUNCTION SetOldNumbers (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY OldNumbers () AS LONG
   DECLARE PROPERTY OldNumbers (BYVAL Value AS LONG)
   DECLARE FUNCTION GetOverlapping () AS LONG
   DECLARE FUNCTION SetOverlapping (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Overlapping () AS LONG
   DECLARE PROPERTY Overlapping (BYVAL Value AS LONG)
   DECLARE FUNCTION GetPositionSubSuper () AS LONG
   DECLARE FUNCTION SetPositionSubSuper (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY PositionSubSuper () AS LONG
   DECLARE PROPERTY PositionSubSuper (BYVAL Value AS LONG)
   DECLARE FUNCTION GetScaling () AS LONG
   DECLARE FUNCTION SetScaling (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Scaling () AS LONG
   DECLARE PROPERTY Scaling (BYVAL Value AS LONG)
   DECLARE FUNCTION GetSpaceExtension () AS LONG
   DECLARE FUNCTION SetSpaceExtension (BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY SpaceExtension () AS LONG
   DECLARE PROPERTY SpaceExtension (BYVAL Value AS LONG)
   DECLARE FUNCTION GetUnderlinePositionMode () AS LONG
   DECLARE FUNCTION SetUnderlinePositionMode (BYVAL Value AS LONG = tomUnderlinePositionAuto) AS HRESULT
   DECLARE PROPERTY UnderlinePositionMode () AS LONG
   DECLARE PROPERTY UnderlinePositionMode (BYVAL Value AS LONG)
   DECLARE FUNCTION GetEffects (BYREF Value AS LONG, BYREF Mask AS LONG) AS HRESULT
   DECLARE FUNCTION GetEffects2 (BYREF Value AS LONG, BYREF Mask AS LONG) AS HRESULT
   DECLARE FUNCTION GetProperty (BYVAL nType AS LONG) AS LONG
   DECLARE FUNCTION GetPropertyInfo (BYVAL Index AS LONG, BYVAL pType AS LONG PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   DECLARE FUNCTION PropertyInfo (BYVAL Index AS LONG, BYVAL pType AS LONG PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   DECLARE FUNCTION IsEqual2 (BYVAL pFont AS ITextFont2 PTR) AS LONG
   DECLARE FUNCTION SetEffects (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   DECLARE FUNCTION SetEffects2 (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   DECLARE FUNCTION SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   DECLARE PROPERTY Property_ (BYVAL nType AS LONG) AS LONG
   DECLARE PROPERTY Property_ (BYVAL nType AS LONG, BYVAL Value AS LONG)

END TYPE
' ========================================================================================

' ========================================================================================
' CTextFont2 constructor
' Called when a class variable is created.
' ========================================================================================
PRIVATE CONSTRUCTOR CTextFont2 (BYVAL pTextFont2 AS ITextFont2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CREC_DP("BEGIN CTextFont2 CONSTRUCTOR - pTextFont2 = " & WSTR(m_pTextFont2)) 
   ' // Store the pointer of ITextFont2 interface
   IF pTextFont2 =  NULL THEN EXIT CONSTRUCTOR
   IF fAddRef THEN pTextFont2->lpvtbl->AddRef(pTextFont2)
   m_pTextFont2 = pTextFont2
   CREC_DP("END CTextFont2 CONSTRUCTOR - m_pTextFont2 = " & WSTR(m_pTextFont2))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTextFont2 destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTextFont2
   CREC_DP("CTextFont2 DESTRUCTOR - m_pTextFont2 = " & WSTR(m_pTextFont2))
   ' // Release the ITextFont2 interface
   IF m_pTextFont2 THEN m_pTextFont2->lpvtbl->Release(m_pTextFont2)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Gets a duplicate of this text font object.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetDuplicate () AS ITextFont2 PTR
   DIM pFont AS ITextFont2 PTR
   m_Result = m_pTextFont2->lpvtbl->GetDuplicate2(m_pTextFont2, @pFont)
   RETURN pFont
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the character formatting by copying another text font object.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetDuplicate (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetDuplicate2(m_pTextFont2, pFont)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets a duplicate of this text font object.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Duplicate () AS ITextFont2 PTR
   DIM pFont AS ITextFont2 PTR
   m_Result = m_pTextFont2->lpvtbl->GetDuplicate2(m_pTextFont2, @pFont)
   RETURN pFont
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the character formatting by copying another text font object.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Duplicate (BYVAL pFont AS ITextFont2 PTR)
   m_Result = m_pTextFont2->lpvtbl->SetDuplicate2(m_pTextFont2, pFont)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets a duplicate of this character format object.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetDuplicate2 () AS ITextFont2 PTR
   DIM pFont AS ITextFont2 PTR
   m_Result = m_pTextFont2->lpvtbl->GetDuplicate2(m_pTextFont2, @pFont)
   RETURN pFont
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed with double horizontal lines through the center.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetDuplicate2 (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetDuplicate2(m_pTextFont2, pFont)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets a duplicate of this character format object.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Duplicate2 () AS ITextFont2 PTR
   DIM pFont AS ITextFont2 PTR
   m_Result = m_pTextFont2->lpvtbl->GetDuplicate2(m_pTextFont2, @pFont)
   RETURN pFont
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed with double horizontal lines through the center.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Duplicate2 (BYVAL pFont AS ITextFont2 PTR)
   m_Result = m_pTextFont2->lpvtbl->SetDuplicate2(m_pTextFont2, pFont)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Determines whether the font can be changed.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.CanChange () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->CanChange(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether this text font object has the same properties as the specified text font object.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.IsEqual (BYVAL pFont AS ITextFont2 PTR) AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->IsEqual(m_pTextFont2, pFont, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Determines whether this text font object has the same properties as the specified text font object.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.IsEqual2 (BYVAL pFont AS ITextFont2 PTR) AS LONG
   DIM B AS LONG
   m_Result = m_pTextFont2->lpvtbl->IsEqual2(m_pTextFont2, pFont, @B)
   RETURN B
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets the character formatting to the specified values.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.Reset (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->Reset(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character style handle of the characters in a range.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetStyle () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetStyle(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the character style handle of the characters in a range.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetStyle (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetStyle(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the character style handle of the characters in a range.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Style () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetStyle(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the character style handle of the characters in a range.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Style (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetStyle(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether the characters are all uppercase.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAllCaps () AS LONG
   CREC_DP("CTextFont2 GetAllCaps")
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAllCaps(m_pTextFont2, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether the characters are all uppercase.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAllCaps (BYVAL Value AS LONG) AS HRESULT
   CREC_DP("CTextFont2 SetAllCaps")
   this.SetResult(m_pTextFont2->lpvtbl->SetAllCaps(m_pTextFont2, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether the characters are all uppercase.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.AllCaps () AS LONG
   CREC_DP("CTextFont2 GET AllCaps")
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAllCaps(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether the characters are all uppercase.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.AllCaps (BYVAL Value AS LONG)
   CREC_DP("CTextFont2 SET AllCaps")
   m_Result = m_pTextFont2->lpvtbl->SetAllCaps(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the animation type.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAnimation () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAnimation(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the animation type.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAnimation (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetAnimation(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the animation type.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Animation () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAnimation(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the animation type.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Animation (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetAnimation(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the text background (highlight) color.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetBackColor () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetBackColor(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the background color.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetBackColor (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetBackColor(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the text background (highlight) color.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.BackColor () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetBackColor(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the background color.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.BackColor (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetBackColor(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether the characters are bold.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetBold () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetBold(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are bold.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetBold (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetBold(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether the characters are bold.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Bold () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetBold(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are bold.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Bold (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetBold(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are embossed.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetEmboss () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetEmboss(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are embossed.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetEmboss (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetEmboss(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are embossed.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Emboss () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetEmboss(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are embossed.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Emboss (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetEmboss(m_pTextFont2, Value)
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the foreground, or text, color.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetForeColor () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetForeColor(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the foreground (text) color.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetForeColor (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetForeColor(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the foreground, or text, color.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.ForeColor () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetForeColor(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the foreground (text) color.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.ForeColor (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetForeColor(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are hidden.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetHidden () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetHidden(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are hidden.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetHidden (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetHidden(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are hidden.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Hidden () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetHidden(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are hidden.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Hidden (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetHidden(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as imprinted characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetEngrave () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetEngrave(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as imprinted characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetEngrave (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetEngrave(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are displayed as imprinted characters.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Engrave () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetEngrave(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as imprinted characters.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Engrave (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetEngrave(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are in italics.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetItalic () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetItalic(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are in italics.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetItalic (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetItalic(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are in italics.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Italic () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetItalic(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are in italics.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Italic (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetItalic(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the minimum font size at which kerning occurs.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetKerning () AS SINGLE
   DIM Value AS SINGLE
   m_Result = m_pTextFont2->lpvtbl->GetKerning(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the minimum font size at which kerning occurs.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetKerning (BYVAL Value AS SINGLE) As HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetKerning(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' Gets the minimum font size at which kerning occurs.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Kerning () AS SINGLE
   DIM Value AS SINGLE
   m_Result = m_pTextFont2->lpvtbl->GetKerning(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the minimum font size at which kerning occurs.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Kerning (BYVAL Value AS SINGLE)
   m_Result = m_pTextFont2->lpvtbl->SetKerning(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the language ID or language code identifier (LCID).
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetLanguageID () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetLanguageID(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the language ID or language code identifier (LCID).
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetLanguageID (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetLanguageID(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the language ID or language code identifier (LCID).
' ========================================================================================
PRIVATE PROPERTY CTextFont2.LanguageID () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetLanguageID(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the language ID or language code identifier (LCID).
' ========================================================================================
PRIVATE PROPERTY CTextFont2.LanguageID (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetLanguageID(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the font name.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetName () AS CBSTR
   DIM pName AS AFX_BSTR
   m_Result = m_pTextFont2->lpvtbl->GetName(m_pTextFont2, @pName)
   RETURN pName
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the font name.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetName (BYVAL fontName AS AFX_BSTR) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetName(m_pTextFont2, fontName)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the font name.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Name () AS CBSTR
   DIM pName AS AFX_BSTR
   m_Result = m_pTextFont2->lpvtbl->GetName(m_pTextFont2, @pName)
   RETURN pName
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the font name.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Name (BYVAL fontName AS AFX_BSTR)
   m_Result = m_pTextFont2->lpvtbl->SetName(m_pTextFont2, fontName)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as outlined characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetOutline () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetOutline(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as outlined characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetOutline (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetOutline(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are displayed as outlined characters.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Outline () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetOutline(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as outlined characters.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Outline (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetOutline(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the amount that characters are offset vertically relative to the baseline.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetPosition () AS SINGLE
   DIM Value AS SINGLE
   m_Result = m_pTextFont2->lpvtbl->GetPosition(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the amount that characters are offset vertically relative to the baseline.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetPosition (BYVAL Value AS SINGLE) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetPosition(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the amount that characters are offset vertically relative to the baseline.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Position () AS SINGLE
   DIM Value AS SINGLE
   m_Result = m_pTextFont2->lpvtbl->GetPosition(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the amount that characters are offset vertically relative to the baseline.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Position (BYVAL Value AS SINGLE)
   m_Result = m_pTextFont2->lpvtbl->SetPosition(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are protected against attempts to modify them.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetProtected () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetProtected(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are protected against attempts to modify them.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetProtected (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetProtected(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are protected against attempts to modify them.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Protected_ () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetProtected(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are protected against attempts to modify them.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Protected_ (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetProtected(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as shadowed characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetShadow () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetShadow(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as shadowed characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetShadow (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetShadow(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are displayed as shadowed characters.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Shadow () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetShadow(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as shadowed characters.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Shadow (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetShadow(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the font size.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSize () AS SINGLE
   DIM Value AS SINGLE
   m_Result = m_pTextFont2->lpvtbl->GetSize(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the font size.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSize (BYVAL Value AS SINGLE) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetSize(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the font size.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Size () AS SINGLE
   DIM Value AS SINGLE
   m_Result = m_pTextFont2->lpvtbl->GetSize(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the font size.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Size (BYVAL Value AS SINGLE)
   m_Result = m_pTextFont2->lpvtbl->SetSize(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are in small capital letters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSmallCaps () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetSmallCaps(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are in small capital letters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSmallCaps (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetSmallCaps(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are in small capital letters.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.SmallCaps () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetSmallCaps(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are in small capital letters.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.SmallCaps (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetSmallCaps(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the amount of horizontal spacing between characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSpacing () AS SINGLE
   DIM Value AS SINGLE
   m_Result = m_pTextFont2->lpvtbl->GetSpacing(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the amount of horizontal spacing between characters.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSpacing (BYVAL Value AS SINGLE) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetSpacing(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the amount of horizontal spacing between characters.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Spacing () AS SINGLE
   DIM Value AS SINGLE
   m_Result = m_pTextFont2->lpvtbl->GetSpacing(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the amount of horizontal spacing between characters.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Spacing (BYVAL Value AS SINGLE)
   m_Result = m_pTextFont2->lpvtbl->SetSpacing(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed with a horizontal line through the center.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetStrikeThrough () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetStrikeThrough(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed with a horizontal line through the center.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetStrikeThrough (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetStrikeThrough(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are displayed with a horizontal line through the center.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.StrikeThrough () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetStrikeThrough(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed with a horizontal line through the center.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.StrikeThrough (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetStrikeThrough(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as subscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSubscript () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetSubscript(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as subscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSubscript (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetSubscript(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are displayed as subscript.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Subscript () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetSubscript(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as subscript.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Subscript (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetSubscript(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSuperscript () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetSuperscript(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSuperscript (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetSuperscript(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Superscript () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetSuperscript(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Superscript (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetSuperscript(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetUnderline () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetUnderline(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetUnderline (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetUnderline(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Underline () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetUnderline(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed as superscript.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Underline (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetUnderline(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the font weight for the characters in a range.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetWeight () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetWeight(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the font weight for the characters in a range.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetWeight (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetWeight(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the font weight for the characters in a range.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Weight () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetWeight(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the font weight for the characters in a range.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Weight (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetWeight(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the count of extra properties in this character formatting collection.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetCount () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetCount(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTextFont2.Count () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetCount(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether support for automatic ligatures is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAutoLigatures () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAutoLigatures(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether support for automatic ligatures is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAutoLigatures (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetAutoLigatures(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether support for automatic ligatures is active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.AutoLigatures () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAutoLigatures(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether support for automatic ligatures is active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.AutoLigatures (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetAutoLigatures(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the East Asian "autospace alphabetics" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAutospaceAlpha () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAutospaceAlpha(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the East Asian "autospace alpha" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAutospaceAlpha (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetAutospaceAlpha(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the East Asian "autospace alphabetics" state.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.AutospaceAlpha () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAutospaceAlpha(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the East Asian "autospace alpha" state.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.AutospaceAlpha (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetAutospaceAlpha(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the East Asian "autospace numeric" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAutospaceNumeric () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAutospaceNumeric(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the East Asian "autospace numeric" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAutospaceNumeric (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetAutospaceNumeric(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the East Asian "autospace numeric" state.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.AutospaceNumeric () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAutospaceNumeric(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the East Asian "autospace numeric" state.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.AutospaceNumeric (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetAutospaceNumeric(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the East Asian "autospace parentheses" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetAutospaceParens () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAutospaceParens(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the East Asian "autospace parentheses" state.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetAutospaceParens (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetAutospaceParens(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the East Asian "autospace parentheses" state.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.AutospaceParens () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetAutospaceParens(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the East Asian "autospace parentheses" state.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.AutospaceParens (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetAutospaceParens(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the character repertoire (writing system).
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetCharRep () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetCharRep(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the character repertoire (writing system).
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetCharRep (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetCharRep(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the character repertoire (writing system).
' ========================================================================================
PRIVATE PROPERTY CTextFont2.CharRep () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetCharRep(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the character repertoire (writing system).
' ========================================================================================
PRIVATE PROPERTY CTextFont2.CharRep (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetCharRep(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the East Asian compression mode.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetCompressionMode () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetCompressionMode(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the East Asian compression mode.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetCompressionMode (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetCompressionMode(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the East Asian compression mode.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.CompressionMode () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetCompressionMode(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the East Asian compression mode.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.CompressionMode (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetCompressionMode(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the client cookie.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetCookie () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetCookie(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the client cookie.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetCookie (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetCookie(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the client cookie.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Cookie () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetCookie(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the client cookie.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Cookie (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetCookie(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether characters are displayed with double horizontal lines through the center.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetDoubleStrike () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetDoubleStrike(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed with double horizontal lines through the center.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetDoubleStrike (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetDoubleStrike(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether characters are displayed with double horizontal lines through the center.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.DoubleStrike () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetDoubleStrike(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether characters are displayed with double horizontal lines through the center.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.DoubleStrike (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetDoubleStrike(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the link type.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetLinkType () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetLinkType(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTextFont2.LinkType () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetLinkType(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether a math zone is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetMathZone () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetMathZone(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether a math zone is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetMathZone (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetMathZone(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether a math zone is active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.MathZone () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetMathZone(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether a math zone is active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.MathZone (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetMathZone(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether "decrease widths on pairs" is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetModWidthPairs () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetModWidthPairs(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether "decrease widths on pairs" is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetModWidthPairs (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetModWidthPairs(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether "decrease widths on pairs" is active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.ModWidthPairs () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetModWidthPairs(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether "decrease widths on pairs" is active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.ModWidthPairs (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetModWidthPairs(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether "increase width of whitespace" is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetModWidthSpace () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetModWidthSpace(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether "increase width of whitespace" is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetModWidthSpace (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetModWidthSpace(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether "increase width of whitespace" is active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.ModWidthSpace () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetModWidthSpace(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether "increase width of whitespace" is active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.ModWidthSpace (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetModWidthSpace(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether old-style numbers are active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetOldNumbers () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetOldNumbers(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether old-style numbers are active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetOldNumbers (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetOldNumbers(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether old-style numbers are active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.OldNumbers () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetOldNumbers(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether old-style numbers are active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.OldNumbers (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetOldNumbers(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets whether overlapping text is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetOverlapping () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetOverlapping(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets whether overlapping text is active.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetOverlapping (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetOverlapping(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets whether overlapping text is active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Overlapping () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetOverlapping(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets whether overlapping text is active.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Overlapping (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetOverlapping(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the subscript or superscript position relative to the baseline.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetPositionSubSuper () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetPositionSubSuper(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the position of a subscript or superscript relative to the baseline, as a percentage
' of the font height.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetPositionSubSuper (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetPositionSubSuper(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the subscript or superscript position relative to the baseline.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.PositionSubSuper () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetPositionSubSuper(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the position of a subscript or superscript relative to the baseline, as a percentage
' of the font height.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.PositionSubSuper (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetPositionSubSuper(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the font horizontal scaling percentage.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetScaling () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetScaling(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the font horizontal scaling percentage.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetScaling (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetScaling(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the font horizontal scaling percentage.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Scaling () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetScaling(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the font horizontal scaling percentage.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Scaling (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetScaling(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the East Asian space extension value.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetSpaceExtension () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetSpaceExtension(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the East Asian space extension value.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetSpaceExtension (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetSpaceExtension(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the East Asian space extension value.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.SpaceExtension () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetSpaceExtension(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the East Asian space extension value.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.SpaceExtension (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetSpaceExtension(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the underline position mode.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetUnderlinePositionMode () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetUnderlinePositionMode(m_pTextFont2, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the underline position mode.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetUnderlinePositionMode (BYVAL Value AS LONG = tomUnderlinePositionAuto) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetUnderlinePositionMode(m_pTextFont2, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the underline position mode.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.UnderlinePositionMode () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetUnderlinePositionMode(m_pTextFont2, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the underline position mode.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.UnderlinePositionMode (BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetUnderlinePositionMode(m_pTextFont2, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetEffects (BYREF Value AS LONG, BYREF Mask AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->GetEffects(m_pTextFont2, @Value, @Mask)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetEffects (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetEffects(m_pTextFont2, Value, Mask)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the additional character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetEffects2 (BYREF Value AS LONG, BYREF Mask AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->GetEffects2(m_pTextFont2, @Value, @Mask)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the additional character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetEffects2 (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetEffects2(m_pTextFont2, Value, Mask)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the value of the specified property.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetProperty(m_pTextFont2, nType, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the additional character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextFont2->lpvtbl->SetProperty(m_pTextFont2, nType, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the value of the specified property.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Property_ (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   m_Result = m_pTextFont2->lpvtbl->GetProperty(m_pTextFont2, nType, @Value)
   RETURN Value
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the additional character format effects.
' ========================================================================================
PRIVATE PROPERTY CTextFont2.Property_ (BYVAL nType AS LONG, BYVAL Value AS LONG)
   m_Result = m_pTextFont2->lpvtbl->SetProperty(m_pTextFont2, nType, Value)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the additional character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.GetPropertyInfo (BYVAL Index AS LONG, BYVAL pType AS LONG PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IF m_pTextFont2 = NULL THEN m_Result = E_POINTER: RETURN m_Result
   this.SetResult(m_pTextFont2->lpvtbl->GetPropertyInfo(m_pTextFont2, Index, pType, pValue))
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the additional character format effects.
' ========================================================================================
PRIVATE FUNCTION CTextFont2.PropertyInfo (BYVAL Index AS LONG, BYVAL pType AS LONG PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IF m_pTextFont2 = NULL THEN m_Result = E_POINTER: RETURN m_Result
   m_Result = m_pTextFont2->lpvtbl->GetPropertyInfo(m_pTextFont2, Index, pType, pValue)
   RETURN m_Result
END FUNCTION
' ========================================================================================

END NAMESPACE
