' ########################################################################################
' Microsoft Windows
' File: CTOMRow.inc
' Contents: Class to work with the Richd Edit Text Object Model (TOM).
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' // To be included by CTOM.inc

' ########################################################################################
' CTextRow class
' ########################################################################################
TYPE CTextRow EXTENDS CTOMBase

Public:
   m_pTextRow AS ITextRow PTR

Public:
   DECLARE CONSTRUCTOR (BYVAL pTextRow AS ITextRow PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR LET (BYVAL pTextRow AS ITextRow PTR)
   DECLARE OPERATOR CAST () AS ITextRow PTR
   DECLARE FUNCTION TextRowPtr () AS ITextRow PTR
   DECLARE FUNCTION Attach (BYVAL pTextRow AS ITextRow PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Detach () AS ITextRow PTR

   ' // ITextRow interface methods
   DECLARE FUNCTION GetAlignment () AS LONG
   DECLARE FUNCTION SetAlignment (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellCount () AS LONG
   DECLARE FUNCTION SetCellCount (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellCountCache () AS LONG
   DECLARE FUNCTION SetCellCountCache (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellIndex () AS LONG
   DECLARE FUNCTION SetCellIndex (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellMargin () AS LONG
   DECLARE FUNCTION SetCellMargin (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetHeight () AS LONG
   DECLARE FUNCTION SetHeight (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetIndent () AS LONG
   DECLARE FUNCTION SetIndent (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetKeepTogether () AS LONG
   DECLARE FUNCTION SetKeepTogether (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetKeepWithNext () AS LONG
   DECLARE FUNCTION SetKeepWithNext (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetNestLevel () AS LONG
   DECLARE FUNCTION GetRTL () AS LONG
   DECLARE FUNCTION SetRTL (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellAlignment () AS LONG
   DECLARE FUNCTION SetCellAlignment (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellColorBack () AS LONG
   DECLARE FUNCTION SetCellColorBack (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellColorFore () AS LONG
   DECLARE FUNCTION SetCellColorFore (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellMergeFlags () AS LONG
   DECLARE FUNCTION SetCellMergeFlags (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellShading () AS LONG
   DECLARE FUNCTION SetCellShading (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellVerticalText () AS LONG
   DECLARE FUNCTION SetCellVerticalText (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellWidth () AS LONG
   DECLARE FUNCTION SetCellWidth (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetCellBorderColors (BYVAL pcrLeft AS LONG PTR, BYVAL pcrTop AS LONG PTR, _
      BYVAL pcrRight AS LONG PTR, BYVAL pcrBottom AS LONG PTR) AS HRESULT
   DECLARE FUNCTION GetCellBorderWidths (BYVAL pduLeft AS LONG PTR, BYVAL pduTop AS LONG PTR, _
      BYVAL pduRight AS LONG PTR, BYVAL pduBottom AS LONG PTR) AS HRESULT
   DECLARE FUNCTION SetCellBorderColors (BYVAL crLeft AS LONG, BYVAL crTop AS LONG, _
      BYVAL crRight AS LONG, BYVAL crBottom AS LONG) AS HRESULT
   DECLARE FUNCTION SetCellBorderWidths (BYVAL duLeft AS LONG, BYVAL duTop AS LONG, _
      BYVAL duRight AS LONG, BYVAL duBottom AS LONG) AS HRESULT
   DECLARE FUNCTION Apply (BYVAL cRow AS LONG, BYVAL Flags AS LONG) AS HRESULT
   DECLARE FUNCTION CanChange () AS LONG
   DECLARE FUNCTION GetProperty (BYVAL nType AS LONG) AS LONG
   DECLARE FUNCTION Insert () AS LONG
   DECLARE FUNCTION IsEqual (BYVAL pRow AS ITextRow PTR) AS LONG
   DECLARE FUNCTION Reset (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CTextRange2 constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CTextRow (BYVAL pTextRow AS ITextRow PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Store the pointer of ITextRow interface
   IF pTextRow = NULL THEN m_Result = E_INVALIDARG : EXIT CONSTRUCTOR
   IF fAddRef THEN pTextRow->lpvtbl->AddRef(pTextRow)
   m_pTextRow = pTextRow
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTextRange2 destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTextRow
   ' // Release the ITextRow interface
   IF m_pTextRow THEN m_pTextRow->lpvtbl->Release(m_pTextRow)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assignment operator
' ========================================================================================
PRIVATE OPERATOR CTextRow.LET (BYVAL pTextRow AS ITextRow PTR)
   m_Result = 0
   IF pTextRow = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextRow THEN m_pTextRow->lpvtbl->Release(m_pTextRow)
   ' // Attach the passed interface pointer to the class
   m_pTextRow = pTextRow
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operator
' ========================================================================================
PRIVATE OPERATOR CTextRow.CAST () AS ITextRow PTR
   m_Result = 0
   OPERATOR = m_pTextRow
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the underlying ITextRow interface
' ========================================================================================
PRIVATE FUNCTION CTextRow.TextRowPtr () AS ITextRow PTR
   m_Result = 0
   RETURN m_pTextRow
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches the passed text document object to the class
' ========================================================================================
PRIVATE FUNCTION CTextRow.Attach (BYVAL pTextRow AS ITextRow PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextRow = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextRow THEN m_Result = m_pTextRow->lpvtbl->Release(m_pTextRow)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextRow->lpvtbl->AddRef(pTextRow)
   m_pTextRow = pTextRow
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the interface pointer from the class
' ========================================================================================
PRIVATE FUNCTION CTextRow.Detach () AS ITextRow PTR
   m_Result = 0
   DIM pTextRow AS ITextRow PTR = m_pTextRow
   m_pTextRow = NULL
   RETURN pTextRow
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the horizontal alignment of a row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetAlignment () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetAlignment(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the horizontal alignment of a row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetAlignment (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetAlignment(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the count of cells in this row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellCount () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellCount(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the count of cells in a row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellCount (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellCount(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the count of cells cached for this row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellCountCache () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellCountCache(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the count of cells cached for a row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellCountCache (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellCountCache(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the index of the active cell to get or set parameters for.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellIndex () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellIndex(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the index of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellIndex (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellIndex(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the cell margin of this row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellMargin () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellMargin(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the cell margin of a row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellMargin (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellMargin(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the height of the row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetHeight () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetHeight(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the height of a row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetHeight (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetHeight(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the indent of this row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetIndent () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetIndent(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the indent of a row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetIndent (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetIndent(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether this row is allowed to be broken across pages.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetKeepTogether () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetKeepTogether(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether this row is allowed to be broken across pages.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetKeepTogether (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetKeepTogether(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether this row should appear on the same page as the row that follows it.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetKeepWithNext () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetKeepWithNext(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether a row should appear on the same page as the row that follows it.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetKeepWithNext (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetKeepWithNext(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the nest level of a table.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetNestLevel () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetNestLevel(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether this row has right-to-left orientation.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetRTL () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetRTL(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether this row has right-to-left orientation.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetRTL (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetRTL(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets whether this row has right-to-left orientation.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellAlignment () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellAlignment(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets whether this row has right-to-left orientation.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellAlignment (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellAlignment(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the background color of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellColorBack () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellColorBack(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellColorBack (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellColorBack(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the foreground color of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellColorFore () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellColorFore(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the foreground color of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellColorFore (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellColorFore(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the merge flags of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellMergeFlags () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellMergeFlags(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the merge flags of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellMergeFlags (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellMergeFlags(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the shading of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellShading () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellShading(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the shading of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellShading (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellShading(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the vertical-text setting of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellVerticalText () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellVerticalText(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the vertical-text setting of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellVerticalText (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellVerticalText(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellWidth () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellWidth(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the active cell width in twips.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellWidth (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellWidth(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the border colors of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellBorderColors (BYVAL pcrLeft AS LONG PTR, BYVAL pcrTop AS LONG PTR, _
BYVAL pcrRight AS LONG PTR, BYVAL pcrBottom AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->GetCellBorderColors(m_pTextRow, pcrLeft, pcrTop, pcrRight, pcrBottom))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the border widths of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetCellBorderWidths (BYVAL pduLeft AS LONG PTR, BYVAL pduTop AS LONG PTR, _
BYVAL pduRight AS LONG PTR, BYVAL pduBottom AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->GetCellBorderWidths(m_pTextRow, pduLeft, pduTop, pduRight, pduBottom))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the border colors of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellBorderColors (BYVAL crLeft AS LONG, BYVAL crTop AS LONG, _
BYVAL crRight AS LONG, BYVAL crBottom AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellBorderColors(m_pTextRow, crLeft, crTop, crRight, crBottom))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the border widths of the active cell.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetCellBorderWidths (BYVAL duLeft AS LONG, BYVAL duTop AS LONG, _
BYVAL duRight AS LONG, BYVAL duBottom AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellBorderWidths(m_pTextRow, duLeft, duTop, duRight, duBottom))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Applies the formatting attributes of this text row object to the specified rows in the associated ITextRange2.
' ========================================================================================
PRIVATE FUNCTION CTextRow.Apply (BYVAL cRow AS LONG, BYVAL Flags AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->Apply(m_pTextRow, cRow, Flags))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether changes can be made to this row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.CanChange () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->CanChange(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the value of the specified property.
' ========================================================================================
PRIVATE FUNCTION CTextRow.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetProperty(m_pTextRow, nType, @Value))
   FUNCTION = Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a row, or rows, at the location identified by the associated ITextRange2 object.
' ========================================================================================
PRIVATE FUNCTION CTextRow.Insert () AS LONG
   DIM cRow AS LONG
   this.SetResult(m_pTextRow->lpvtbl->Insert(m_pTextRow, @cRow))
   FUNCTION = cRow
END FUNCTION
' ========================================================================================

' ========================================================================================
' Compares two table rows to determine if they have the same properties.
' ========================================================================================
PRIVATE FUNCTION CTextRow.IsEqual (BYVAL pRow AS ITextRow PTR) AS LONG
   DIM B AS LONG
   this.SetResult(m_pTextRow->lpvtbl->IsEqual(m_pTextRow, pRow, @B))
   FUNCTION = B
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets a row.
' ========================================================================================
PRIVATE FUNCTION CTextRow.Reset (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->Reset(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the value of the specified property.
' ========================================================================================
PRIVATE FUNCTION CTextRow.SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetProperty(m_pTextRow, nType, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================
