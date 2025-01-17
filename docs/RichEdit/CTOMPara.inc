' ########################################################################################
' Microsoft Windows
' File: CTOMPara.inc
' Contents: Class to work with the Richd Edit Text Object Model (TOM).
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' // To be included by CTOM.inc

' ########################################################################################
' CTextPara2 class
' ########################################################################################
TYPE CTextPara2 EXTENDS CTOMBase

Public:
   m_pTextPara2 AS ITextPara2 PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR LET (BYVAL pTextPara2 AS ITextPara2 PTR)
   DECLARE OPERATOR CAST () AS ITextPara2 PTR
   DECLARE FUNCTION TextParaPtr () AS ITextPara2 PTR
   DECLARE FUNCTION Attach (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Detach () AS ITextPara2 PTR

   ' // Inherited ITextPara interface methods
   DECLARE FUNCTION GetDuplicate () AS ITextPara2 PTR
   DECLARE FUNCTION SetDuplicate (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   DECLARE FUNCTION CanChange (BYVAL pPara AS ITextPara2 PTR) AS LONG
   DECLARE FUNCTION IsEqual (BYVAL pPara AS ITextPara2 PTR) AS LONG
   DECLARE FUNCTION Reset (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetStyle () AS LONG
   DECLARE FUNCTION SetStyle (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetAlignment () AS LONG
   DECLARE FUNCTION SetAlignment (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetHyphenation () AS LONG
   DECLARE FUNCTION SetHyphenation (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetFirstLineIndent () AS SINGLE
   DECLARE FUNCTION GetKeepTogether () AS LONG
   DECLARE FUNCTION SetKeepTogether (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetKeepWithNext () AS LONG
   DECLARE FUNCTION SetKeepWithNext (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetLeftIndent () AS SINGLE
   DECLARE FUNCTION GetLineSpacing () AS SINGLE
   DECLARE FUNCTION GetLineSpacingRule () AS LONG
   DECLARE FUNCTION GetListAlignment () AS LONG
   DECLARE FUNCTION SetListAlignment (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetListLevelIndex () AS LONG
   DECLARE FUNCTION SetListLevelIndex (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetListStart () AS LONG
   DECLARE FUNCTION SetListStart (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetListTab () AS SINGLE
   DECLARE FUNCTION SetListTab (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE FUNCTION GetListType () AS LONG
   DECLARE FUNCTION SetListType (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetNoLineNumber () AS LONG
   DECLARE FUNCTION SetNoLineNumber (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetPageBreakBefore () AS LONG
   DECLARE FUNCTION SetPageBreakBefore (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetRightIndent () AS SINGLE
   DECLARE FUNCTION SetRightIndent (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE FUNCTION SetIndents (BYVAL First AS SINGLE, BYVAL Left AS SINGLE, BYVAL Right AS SINGLE) AS HRESULT
   DECLARE FUNCTION SetLineSpacing (BYVAL Rule AS LONG, BYVAL Spacing AS SINGLE) AS HRESULT
   DECLARE FUNCTION GetSpaceAfter () AS SINGLE
   DECLARE FUNCTION SetSpaceAfter (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE FUNCTION GetSpaceBefore () AS SINGLE
   DECLARE FUNCTION SetSpaceBefore (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE FUNCTION GetWidowControl () AS LONG
   DECLARE FUNCTION SetWidowControl (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetTabCount () AS LONG
   DECLARE FUNCTION AddTab (BYVAL tbPos AS SINGLE, BYVAL tbAlign AS LONG, BYVAL tbLeader AS LONG) AS HRESULT
   DECLARE FUNCTION ClearAllTabs () AS HRESULT
   DECLARE FUNCTION DeleteTab (BYVAL tbPos AS SINGLE) AS HRESULT
   DECLARE FUNCTION GetTab (BYVAL iTab AS LONG, BYVAL ptbPos AS SINGLE PTR, BYVAL ptbAlign AS LONG PTR, BYVAL ptbLeader AS LONG PTR) AS HRESULT

   ' // ITextPara2 interface methods
   DECLARE FUNCTION GetBorders () AS IUnknown PTR
   DECLARE FUNCTION GetDuplicate2 () AS ITextPara2 PTR
   DECLARE FUNCTION SetDuplicate2 (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   DECLARE FUNCTION GetFontAlignment () AS LONG
   DECLARE FUNCTION SetFontAlignment (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetHangingPunctuation () AS LONG
   DECLARE FUNCTION SetHangingPunctuation (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetSnapToGrid () AS LONG
   DECLARE FUNCTION SetSnapToGrid (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetTrimPunctuationAtStart () AS LONG
   DECLARE FUNCTION SetTrimPunctuationAtStart (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetEffects (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   DECLARE FUNCTION GetProperty (BYVAL nType AS LONG) AS LONG
   DECLARE FUNCTION IsEqual2 (BYVAL pPara AS ITextPara2 PTR) AS LONG
   DECLARE FUNCTION SetEffects (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   DECLARE FUNCTION SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CTextRange2 constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CTextPara2
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTextRange2 constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CTextPara2 (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Store the pointer of ITextPara2 interface
   IF pTextPara2 = NULL THEN m_Result = E_INVALIDARG : EXIT CONSTRUCTOR
   IF fAddRef THEN pTextPara2->lpvtbl->AddRef(pTextPara2)
   m_pTextPara2 = pTextPara2
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTextRange2 destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTextPara2
   ' // Release the ITextPara2 interface
   IF m_pTextPara2 THEN m_pTextPara2->lpvtbl->Release(m_pTextPara2)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assignment operator
' ========================================================================================
PRIVATE OPERATOR CTextPara2.LET (BYVAL pTextPara2 AS ITextPara2 PTR)
   m_Result = 0
   IF pTextPara2 = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextPara2 THEN m_pTextPara2->lpvtbl->Release(m_pTextPara2)
   ' // Attach the passed interface pointer to the class
   m_pTextPara2 = pTextPara2
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operator
' ========================================================================================
PRIVATE OPERATOR CTextPara2.CAST () AS ITextPara2 PTR
   m_Result = 0
   OPERATOR = m_pTextPara2
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the underlying ITextPara2 interface
' ========================================================================================
PRIVATE FUNCTION CTextPara2.TextParaPtr () AS ITextPara2 PTR
   m_Result = 0
   RETURN m_pTextPara2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches the passed text document object to the class
' ========================================================================================
PRIVATE FUNCTION CTextPara2.Attach (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextPara2 = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextPara2 THEN m_Result = m_pTextPara2->lpvtbl->Release(m_pTextPara2)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextPara2->lpvtbl->AddRef(pTextPara2)
   m_pTextPara2 = pTextPara2
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the interface pointer from the class
' ========================================================================================
PRIVATE FUNCTION CTextPara2.Detach () AS ITextPara2 PTR
   m_Result = 0
   DIM pTextPara2 AS ITextPara2 PTR = m_pTextPara2
   m_pTextPara2 = NULL
   RETURN pTextPara2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a duplicate of the specified paragraph format object. The duplicate property is
' the default property of an ITextPara object.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetDuplicate () AS ITextPara2 PTR
   DIM pPara AS ITextPara2 PTR
   this.SetResult(m_pTextPara2->lpvtbl->GetDuplicate(m_pTextPara2, @pPara))
   FUNCTION = pPara
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a duplicate of the specified paragraph format object. The duplicate property is
' the default property of an ITextPara object.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetDuplicate (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetDuplicate(m_pTextPara2, pPara))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the paragraph formatting can be changed.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.CanChange (BYVAL pPara AS ITextPara2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->CanChange(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if the current range has the same properties as a specified range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.IsEqual (BYVAL pPara AS ITextPara2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->CanChange(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets the paragraph formatting to a choice of default values.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.Reset (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->Reset(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the style handle to the paragraphs in the specified range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetStyle () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetStyle(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the paragraph style for the paragraphs in a range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetStyle (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetStyle(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current paragraph alignment value.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetAlignment () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetAlignment(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the paragraph alignment.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetAlignment (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetStyle(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether automatic hyphenation is enabled for the range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetHyphenation () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetHyphenation(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Controls hyphenation for the paragraphs in the range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetHyphenation (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetHyphenation(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the amount used to indent the first line of a paragraph relative to the left indent.
' The left indent is the indent for all lines of the paragraph except the first line.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetFirstLineIndent () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetFirstLineIndent(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether page breaks are allowed within paragraphs.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetKeepTogether () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetKeepTogether(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Controls whether page breaks are allowed within a paragraph in a range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetKeepTogether (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetKeepTogether(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether page breaks are allowed within paragraphs.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetKeepWithNext () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetKeepWithNext(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Controls whether page breaks are allowed within a paragraph in a range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetKeepWithNext (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetKeepWithNext(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the distance used to indent all lines except the first line of a paragraph.
' The distance is relative to the left margin.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetLeftIndent () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetLeftIndent(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the line-spacing value for the text range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetLineSpacing () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetLineSpacing(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the line-spacing rule for the text range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetLineSpacingRule () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetLineSpacingRule(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the list level index used with paragraphs.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetListLevelIndex () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetListLevelIndex(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the list level index used for paragraphs.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetListLevelIndex (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetListLevelIndex(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the starting value or code of a list numbering sequence.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetListStart () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetListStart(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the starting number or Unicode value for a numbered list.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetListStart (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetListStart(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the starting value or code of a list numbering sequence.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetListTab () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetListTab(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the list tab setting, which is the distance between the first indent and the start
' of the text on the first line.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetListTab (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetListTab(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the kind of numbering to use with paragraphs.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetListType () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetListType(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the type of list to be used for paragraphs.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetListType (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetListTab(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether paragraph numbering is enabled.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetNoLineNumber () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetNoLineNumber(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether to suppress line numbering of paragraphs in a range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetNoLineNumber (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetNoLineNumber(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether each paragraph in the range must begin on a new page.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetPageBreakBefore () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetPageBreakBefore(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Controls whether there is a page break before each paragraph in a range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetPageBreakBefore (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetPageBreakBefore(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the size of the right margin indent of a paragraph.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetRightIndent () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetRightIndent(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Controls whether there is a page break before each paragraph in a range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetRightIndent (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetRightIndent(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the first-line indent, the left indent, and the right indent for a paragraph.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetIndents (BYVAL First AS SINGLE, BYVAL Left AS SINGLE, BYVAL Right AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetIndents(m_pTextPara2, First, Left, Right))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the paragraph line-spacing rule and the line spacing for a paragraph.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetLineSpacing (BYVAL Rule AS LONG, BYVAL Spacing AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetLineSpacing(m_pTextPara2, Rule, Spacing))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the amount of vertical space below a paragraph.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetSpaceAfter () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetSpaceAfter(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the amount of space that follows a paragraph.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetSpaceAfter (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetSpaceAfter(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the amount of vertical space above a paragraph.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetSpaceBefore () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetSpaceBefore(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the amount of space preceding a paragraph.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetSpaceBefore (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetSpaceBefore(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the widow and orphan control state for the paragraphs in a range.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetWidowControl () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetWidowControl(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Controls the suppression of widows and orphans.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetWidowControl (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetWidowControl(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the tab count.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetTabCount () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetTabCount(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a tab at the displacement tbPos, with type tbAlign, and leader style, tbLeader.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.AddTab (BYVAL tbPos AS SINGLE, BYVAL tbAlign AS LONG, BYVAL tbLeader AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->AddTab(m_pTextPara2, tbPos, tbAlign, tbLeader))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Clears all tabs, reverting to equally spaced tabs with the default tab spacing.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.ClearAllTabs () AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->ClearAllTabs(m_pTextPara2))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a tab at a specified displacement.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.DeleteTab (BYVAL tbPos AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->DeleteTab(m_pTextPara2, tbPos))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a tab at a specified displacement.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetTab (BYVAL iTab AS LONG, BYVAL ptbPos AS SINGLE PTR, BYVAL ptbAlign AS LONG PTR, BYVAL ptbLeader AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->GetTab(m_pTextPara2, iTab, ptbPos, ptbAlign, ptbLeader))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the borders collection.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetBorders () AS IUnknown PTR
   DIM pBorders AS IUnknown PTR
   this.SetResult(m_pTextPara2->lpvtbl->GetBorders(m_pTextPara2, @pBorders))
   FUNCTION = pBorders
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a duplicate of this text paragraph format object.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetDuplicate2 () AS ITextPara2 PTR
   DIM pBorders AS ITextPara2 PTR
   this.SetResult(m_pTextPara2->lpvtbl->GetDuplicate2(m_pTextPara2, @pBorders))
   FUNCTION = pBorders
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a duplicate of this text paragraph format object.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetDuplicate2 (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetDuplicate2(m_pTextPara2, pPara))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the paragraph font alignment state.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetFontAlignment () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetFontAlignment(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the paragraph font alignment for Chinese, Japanese, Korean text.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetFontAlignment (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetFontAlignment(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether to hang punctuation symbols on the right margin when the paragraph is justified.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetHangingPunctuation () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetHangingPunctuation(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether to hang punctuation symbols on the right margin when the paragraph is justified.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetHangingPunctuation (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetHangingPunctuation(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether paragraph lines snap to a vertical grid that could be defined for the whole document.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetSnapToGrid () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetSnapToGrid(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether paragraph lines snap to a vertical grid that could be defined for the whole document.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetSnapToGrid (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetSnapToGrid(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether to trim the leading space of a punctuation symbol at the start of a line.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetTrimPunctuationAtStart () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetTrimPunctuationAtStart(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether to trim the leading space of a punctuation symbol at the start of a line.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetTrimPunctuationAtStart (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetTrimPunctuationAtStart(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether to trim the leading space of a punctuation symbol at the start of a line.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetEffects (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->GetEffects(m_pTextPara2, pValue, pMask))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the value of the specified property.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetProperty(m_pTextPara2, nType, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether this text paragraph object has the same properties as the specified
' text paragraph object.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.IsEqual2 (BYVAL pPara AS ITextPara2 PTR) AS LONG
   DIM B AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->IsEqual2(m_pTextPara2, pPara, @B))
   FUNCTION = B
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the paragraph format effects.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetEffects (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetEffects(m_pTextPara2, Value, Mask))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the property value.
' ========================================================================================
PRIVATE FUNCTION CTextPara2.SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetProperty(m_pTextPara2, nType, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================
