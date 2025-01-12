' ########################################################################################
' Microsoft Windows
' File: CTOMSelection.inc
' Contents: Class to work with the Richd Edit Text Object Model (TOM).
' A text selection is a text range with selection highlighting.
' The ITextSelection2 interface inherits from the ITextRange2 interface.
' Currently, this interface contains no methods other than those inherited from ITextRange2.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' // To be included by CTOM.inc

' ########################################################################################
' CTextSelection2 class
' ########################################################################################
TYPE CTextSelection2

Public:
   m_Result AS HRESULT
   m_pTextSelection2 AS ITextSelection2 PTR

Public:
   DECLARE CONSTRUCTOR (BYVAL pTextSelection2 AS ITextSelection2 PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo () AS CWSTR
   DECLARE OPERATOR LET (BYVAL pTextSelection2 AS ITextSelection2 PTR)
   DECLARE OPERATOR CAST () AS ITextSelection2 PTR
   DECLARE FUNCTION TextSelectionPtr () AS ITextSelection2 PTR
   DECLARE FUNCTION Attach (BYVAL pTextSelection2 AS ITextSelection2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Detach () AS ITextSelection2 PTR

   ' // ITextSelection interface methods (inherited from ITextRange)
   DECLARE FUNCTION GetText () AS CBSTR
   DECLARE FUNCTION SetText (BYVAL bstr AS AFX_BSTR) AS HRESULT
   DECLARE FUNCTION GetChar () AS LONG
   DECLARE FUNCTION SetChar (BYVAL Char AS LONG) AS HRESULT
   DECLARE FUNCTION GetDuplicate () AS ITextRange2 PTR
   DECLARE FUNCTION GetFormattedText () AS ITextRange2 PTR
   DECLARE FUNCTION SetFormattedText (BYVAl pRange AS ITextRange2 PTR) AS HRESULT
   DECLARE FUNCTION GetStart () AS LONG
   DECLARE FUNCTION SetStart (BYVAL cpFirst AS LONG) AS HRESULT
   DECLARE FUNCTION GetEnd () AS LONG
   DECLARE FUNCTION SetEnd (BYVAL cpLim AS LONG) AS HRESULT
   DECLARE FUNCTION GetFont () AS ITextFont PTR
   DECLARE FUNCTION SetFont (BYVAL pFont AS ITextFont PTR) AS HRESULT
   DECLARE FUNCTION GetPara () AS ITextPara PTR
   DECLARE FUNCTION SetPara (BYVAL pPara AS ITextPara PTR) AS HRESULT
   DECLARE FUNCTION GetStoryLength () AS LONG
   DECLARE FUNCTION GetStoryType () AS LONG
   DECLARE FUNCTION Collapse (BYVAL bStart AS LONG) AS HRESULT
   DECLARE FUNCTION Expand (BYVAL Unit AS LONG) AS LONG
   DECLARE FUNCTION GetIndex (BYVAL Unit AS LONG) AS LONG
   DECLARE FUNCTION SetIndex (BYVAL Unit AS LONG, BYVAL Index AS LONG, BYVAl Extend AS LONG) AS HRESULT
   DECLARE FUNCTION SetRange (BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   DECLARE FUNCTION InRange (BYVAL pRange AS ITextRange2 PTR) AS LONG
   DECLARE FUNCTION InStory (BYVAL pRange AS ITextRange2 PTR) AS LONG
   DECLARE FUNCTION IsEqual (BYVAL pRange AS ITextRange2 PTR) AS LONG
   DECLARE FUNCTION Select_ () AS HRESULT
   DECLARE FUNCTION StartOf (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DECLARE FUNCTION EndOf (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DECLARE FUNCTION Move (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION MoveStart (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION MoveEnd (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION MoveWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION MoveStartWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION MoveEndWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION MoveUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION MoveStartUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION MoveEndUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION FindText (BYVAL bstr AS AFX_BSTR, BYVAL cch AS LONG, BYVAL Flags AS LONG) AS LONG
   DECLARE FUNCTION FindTextStart (BYVAL bstr AS AFX_BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG) AS LONG
   DECLARE FUNCTION FindTextEnd (BYVAL bstr AS AFX_BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG) AS LONG
   DECLARE FUNCTION Delete_ (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION Cut (BYVAL pVar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION Copy (BYVAL pVar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION Paste (BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG) AS HRESULT
   DECLARE FUNCTION CanPaste (BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG) AS LONG
   DECLARE FUNCTION CanEdit () AS LONG
   DECLARE FUNCTION ChangeCase (BYVAL nType AS LONG) AS HRESULT
   DECLARE FUNCTION GetPoint (BYVAL nType AS LONG, BYVAL px AS LONG PTR, BYVAL py AS LONG PTR) AS HRESULT
   DECLARE FUNCTION SetPoint (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG, BYVAL Extend AS LONG) AS HRESULT
   DECLARE FUNCTION ScrollIntoView (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetEmbeddedObject () AS IUnknown PTR

   ' // ITextSelection2 methods (inherited from ITextSelection2)
   DECLARE FUNCTION GetFlags () AS LONG
   DECLARE FUNCTION SetFlags (BYVAL Flags AS LONG) AS HRESULT
   DECLARE FUNCTION GetType () AS LONG
   DECLARE FUNCTION MoveLeft (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DECLARE FUNCTION MoveRight (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DECLARE FUNCTION MoveUp (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DECLARE FUNCTION MoveDown (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DECLARE FUNCTION HomeKey (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DECLARE FUNCTION EndKey (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DECLARE FUNCTION TypeText (BYVAL bstr AS AFX_BSTR) AS HRESULT
   DECLARE FUNCTION GetCch () AS LONG
   DECLARE FUNCTION GetCells () AS IUnknown PTR
   DECLARE FUNCTION GetColumn () AS IUnknown PTR
   DECLARE FUNCTION GetCount () AS LONG
   DECLARE FUNCTION GetDuplicate2 () AS ITextRange2 PTR
   DECLARE FUNCTION GetFont2 () AS ITextFont2 PTR
   DECLARE FUNCTION SetFont2 (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   DECLARE FUNCTION GetFormattedText2 () AS ITextRange2 PTR
   DECLARE FUNCTION SetFormattedText2 (BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   DECLARE FUNCTION GetGravity () AS LONG
   DECLARE FUNCTION SetGravity (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetPara2 () AS ITextPara2 PTR
   DECLARE FUNCTION SetPara2 (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   DECLARE FUNCTION GetRow () AS ITextRow PTR
   DECLARE FUNCTION GetStartPara () AS LONG
   DECLARE FUNCTION GetTable () AS IUnknown PTR
   DECLARE FUNCTION GetURL () AS CBSTR
   DECLARE FUNCTION SetURL (BYVAL bstr AS AFX_BSTR) AS HRESULT
   DECLARE FUNCTION AddSubrange (BYVAL cp1 AS LONG, BYVAL cp2 AS LONG, BYVAL Activate AS LONG) AS HRESULT
   DECLARE FUNCTION BuildUpMath (BYVAL Flags AS LONG) AS HRESULT
   DECLARE FUNCTION DeleteSubrange (BYVAL cpFirst AS LONG, BYVAL cpLim AS LONG) AS HRESULT
   DECLARE FUNCTION Find (BYVAL pRange AS ITextRange2 PTR, BYVAL Count AS LONG, BYVAL Flags AS LONG) AS HRESULT
   DECLARE FUNCTION GetChar2 (BYVAL Offset AS LONG) AS HRESULT
   DECLARE FUNCTION GetDropCap (BYVAL pcLine AS LONG PTR, BYVAL pPosition AS LONG PTR) AS HRESULT
   DECLARE FUNCTION GetInlineObject (BYVAL pcLine AS LONG PTR, BYVAL pType AS LONG PTR, BYVAL pAlign AS LONG PTR, BYVAL pChar AS LONG PTR, BYVAL pChar1 AS LONG PTR, _
      BYVAL pChar2 AS LONG PTR, BYVAL pCount AS LONG PTR,  BYVAL pTeXStyle AS LONG PTR, BYVAL pcCol AS LONG PTR, BYVAL pLevel AS LONG PTR) AS HRESULT
   DECLARE FUNCTION GetProperty (BYVAL nType AS LONG) AS LONG
   DECLARE FUNCTION GetRect (BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, BYVAL pTop AS LONG PTR, BYVAL pRight AS LONG PTR, _
      BYVAL pBottom AS LONG PTR, BYVAL pHit AS LONG PTR) AS HRESULT
   DECLARE FUNCTION GetSubrange (BYVAL iSubrange AS LONG, BYVAL pcpFirst AS LONG PTR, BYVAL pcpLim AS LONG PTR) AS HRESULT
   DECLARE FUNCTION GetText2 (BYVAL Flags AS LONG) AS CBSTR
   DECLARE FUNCTION HexToUnicode () AS HRESULT
   DECLARE FUNCTION InsertTable (BYVAL cCol AS LONG, BYVAL cRow AS LONG, BYVAL AutoFit AS LONG) AS HRESULT
   DECLARE FUNCTION Linearize (BYVAL Flags AS LONG) AS HRESULT
   DECLARE FUNCTION SetActiveSubrange (BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   DECLARE FUNCTION SetDropCap (BYVAL cLine AS LONG, BYVAL Position AS LONG) AS HRESULT
   DECLARE FUNCTION SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION SetText2 (BYVAL Flags AS LONG, BYVAL bstr AS AFX_BSTR) AS HRESULT
   DECLARE FUNCTION UnicodeToHex () AS HRESULT
   DECLARE FUNCTION SetInlineObject (BYVAL nType AS Long, BYVAL Align AS LONG, BYVAL Char AS LONG, _
      BYVAL Char1 AS LONG, BYVAL Char2 AS LONG, BYVAL Count AS LONG, BYVAL TeXStyle AS LONG, BYVAL cCol AS LONG) AS HRESULT
   DECLARE FUNCTION GetMathFunctionType (BYVAL bstr AS AFX_BSTR) AS HRESULT
   DECLARE FUNCTION InsertImage (BYVAL width_ AS LONG, BYVAL height AS LONG, BYVAL ascent AS LONG, _
      BYVAL nType AS LONG, BYVAL bstrAltText AS AFX_BSTR, BYVAL pStream AS IStream PTR) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CTextSelection2 constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CTextSelection2 (BYVAL pTextSelection2 AS ITextSelection2 PTR)
   ' // Store the pointer of ITextSelection2 interface
   IF pTextSelection2 = NULL THEN m_Result = E_INVALIDARG : EXIT CONSTRUCTOR
   m_pTextSelection2 = pTextSelection2
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTextSelection2 destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTextSelection2
   ' // Release the ITextSelection2 interface
   IF m_pTextSelection2 THEN m_pTextSelection2->lpvtbl->Release(m_pTextSelection2)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assignment operator
' ========================================================================================
PRIVATE OPERATOR CTextSelection2.LET (BYVAL pTextSelection2 AS ITextSelection2 PTR)
   m_Result = 0
   IF pTextSelection2 = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextSelection2 THEN m_pTextSelection2->lpvtbl->Release(m_pTextSelection2)
   ' // Attach the passed interface pointer to the class
   m_pTextSelection2 = pTextSelection2
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operator
' ========================================================================================
PRIVATE OPERATOR CTextSelection2.CAST () AS ITextSelection2 PTR
   m_Result = 0
   OPERATOR = m_pTextSelection2
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the underlying ITextSelection2 interface
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.TextSelectionPtr () AS ITextSelection2 PTR
   m_Result = 0
   RETURN m_pTextSelection2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches the passed text document object to the class
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Attach (BYVAL pTextSelection2 AS ITextSelection2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextSelection2 = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextSelection2 THEN m_Result = m_pTextSelection2->lpvtbl->Release(m_pTextSelection2)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextSelection2->lpvtbl->AddRef(pTextSelection2)
   m_pTextSelection2 = pTextSelection2
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the interface pointer from the class
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Detach () AS ITextSelection2 PTR
   m_Result = 0
   DIM pTextSelection2 AS ITextSelection2 PTR = m_pTextSelection2
   m_pTextSelection2 = NULL
   RETURN pTextSelection2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the plain text in this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetText () AS CBSTR
   DIM pText AS AFX_BSTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetText(m_pTextSelection2, @pText))
   RETURN pText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text in this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetText (BYVAL bstr AS AFX_BSTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetText(m_pTextSelection2, bstr))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character at the start position of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetChar () AS LONG
   DIM Char AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetChar(m_pTextSelection2, @Char))
   RETURN Char
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the character at the starting position of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetChar (BYVAL Char AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetChar(m_pTextSelection2, Char))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character at the start position of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetDuplicate () AS ITextRange2 PTR
   DIM pRange AS ITextRange2 PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetDuplicate(m_pTextSelection2, @pRange))
   RETURN pRange
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an ITextSelection2 object with the specified range's formatted text.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetFormattedText () AS ITextRange2 PTR
   DIM pRange AS ITextRange2 PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetFormattedText(m_pTextSelection2, @pRange))
   RETURN pRange
END FUNCTION
' ========================================================================================
 
 ' ========================================================================================
' Sets the character at the starting position of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetFormattedText (BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetFormattedText(m_pTextSelection2, pRange))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the start character position of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetStart () AS LONG
   DIM cpFirst AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetStart(m_pTextSelection2, @cpFirst))
   RETURN cpFirst
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the character at the starting position of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetStart (BYVAL cpFirst AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetStart(m_pTextSelection2, cpFirst))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the end character position of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetEnd () AS LONG
   DIM cpLim AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetEnd(m_pTextSelection2, @cpLim))
   RETURN cpLim
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the new end character position of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetEnd (BYVAL cpLim AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetEnd(m_pTextSelection2, cpLim))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an ITextFont2 object with the character attributes of the specified range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetFont () AS ITextFont PTR
   DIM pFont AS ITextFont PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetFont(m_pTextSelection2, @pFont))
   RETURN pFont
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the new end character position of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetFont (BYVAL pFont AS ITextFont PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetFont(m_pTextSelection2, pFont))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an ITextPara object with the paragraph attributes of the specified range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetPara () AS ITextPara PTR
   DIM pPara AS ITextPara PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetPara(m_pTextSelection2, @pPara))
   RETURN pPara
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the paragraph attributes of this range to those of the specified ITextPara object.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetPara (BYVAL pPara AS ITextPara PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetPara(m_pTextSelection2, pPara))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the count of characters in the range's story.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetStoryLength () AS LONG
   DIM Count AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetStoryLength(m_pTextSelection2, @Count))
   RETURN Count
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the type of the range's story.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetStoryType () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetStoryType(m_pTextSelection2, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Collapses the specified text range into a degenerate point at either the beginning or end of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Collapse (BYVAL bStart AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->Collapse(m_pTextSelection2, bStart))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Expands this range so that any partial units it contains are completely contained.
' Returns the count of characters added to the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Expand (BYVAL Unit AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->Expand(m_pTextSelection2, Unit, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the story index of the Unit parameter at the specified range Start character
' position. The first Unit in a story has an index value of 1. The index of a Unit is the
' same for all character positions from that immediately preceding the Unit up to the last
' character in the Unit.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetIndex (BYVAL Unit AS LONG) AS LONG
   DIM Index AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetIndex(m_pTextSelection2, Unit, @Index))
   RETURN Index
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes this range to the specified unit of the story.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetIndex (BYVAL Unit AS LONG, BYVAL Index AS LONG, BYVAl Extend AS LONG) AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->SetIndex(m_pTextSelection2, Unit, Index, Extend))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adjusts the range endpoints to the specified values.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetRange (BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetRange(m_pTextSelection2, cpAnchor, cpActive))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether this range is within or at the same text as a specified range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.InRange (BYVAL pRange AS ITextRange2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->InRange(m_pTextSelection2, pRange, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether this range's story is the same as a specified range's story.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.InStory (BYVAL pRange AS ITextRange2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->InStory(m_pTextSelection2, pRange, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether this range has the same character positions and story as those of a specified range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.IsEqual (BYVAL pRange AS ITextRange2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->IsEqual(m_pTextSelection2, pRange, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the start and end positions, and story values of the active selection, to those of this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Select_ () AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->Select(m_pTextSelection2))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the range ends to the start of the first overlapping Unit in the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.StartOf (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->StartOf(m_pTextSelection2, Unit, Extend, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves this range's ends to the end of the last overlapping Unit in the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.EndOf (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->EndOf(m_pTextSelection2, Unit, Extend, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the insertion point forward or backward a specified number of units. If the range
' is nondegenerate, the range is collapsed to an insertion point at either end, depending
' on Count, and then is moved.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Move (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->Move(m_pTextSelection2, Unit, Count, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the start position of the range the specified number of units in the specified direction.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveStart (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveStart(m_pTextSelection2, Unit, Count, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the end position of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveEnd (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveEnd(m_pTextSelection2, Unit, Count, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Starts at a specified end of a range and searches while the characters belong to the set
' specified by Cset and while the number of characters is less than or equal to Count.
' The range is collapsed to an insertion point when a non-matching character is found.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveWhile(m_pTextSelection2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the start position of the range either Count characters, or just past all contiguous
' characters that are found in the set of characters specified by Cset, whichever is less.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveStartWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveStartWhile(m_pTextSelection2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the end of the range either Count characters or just past all contiguous characters
' that are found in the set of characters specified by Cset, whichever is less.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveEndWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveEndWhile(m_pTextSelection2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches up to Count characters for the first character in the set of characters specified
' by Cset. If a character is found, the range is collapsed to that point. The start of the
' search and the direction are also specified by Count.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveUntil(m_pTextSelection2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the start position of the range the position of the first character found that is
' in the set of characters specified by Cset, provided that the character is found within
' Count characters of the start position.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveStartUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveStartUntil(m_pTextSelection2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the start position of the range the position of the first character found that is
' in the set of characters specified by Cset, provided that the character is found within
' Count characters of the start position.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveEndUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveEndUntil(m_pTextSelection2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches up to Count characters for the text given by bstr. The starting position and direction
' are also specified by Count, and the matching criteria are given by Flags.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.FindText (BYVAL bstr AS AFX_BSTR, BYVAL cch AS LONG, BYVAL Flags AS LONG) AS LONG
   DIM Length AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->FindText(m_pTextSelection2, bstr, cch, Flags, @Length))
   RETURN Length
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches up to Count characters for the string, bstr, starting at the range's Start cp (cpFirst).
' The search is subject to the comparison parameter, Flags. If the string is found, the
' Start cp is changed to the matched string, and pLength is set equal to the length of the
' string. If the string is not found, the range is unchanged, and pLength is set equal to zero.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.FindTextStart (BYVAL bstr AS AFX_BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG) AS LONG
   DIM Length AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->FindTextStart(m_pTextSelection2, bstr, Count, Flags, @Length))
   RETURN Length
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches up to Count characters for the string, bstr, starting from the range's End cp.
' The search is subject to the comparison parameter, Flags. If the string is found, the
' End cp is changed to be the end of the matched string, and pLength is set equal to the
' length of the string. If the string is not found, the range is unchanged and pLength is
' set equal to zero.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.FindTextEnd (BYVAL bstr AS AFX_BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG) AS LONG
   DIM Length AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->FindTextEnd(m_pTextSelection2, bstr, Count, Flags, @Length))
   RETURN Length
END FUNCTION
' ========================================================================================

' ========================================================================================
' Mimics the DELETE and BACKSPACE keys, with and without the CTRL key depressed.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Delete_ (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->Delete_(m_pTextSelection2, Unit, Count, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Cuts the plain or rich text to a data object or to the Clipboard, depending on the pVar parameter.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Cut (BYVAL pVar AS VARIANT PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->Cut(m_pTextSelection2, pVar))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the text to a data object.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Copy (BYVAL pVar AS VARIANT PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->Copy(m_pTextSelection2, pVar))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Pastes text from a specified data object.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Paste (BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->Paste(m_pTextSelection2, pVar, Format))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if a data object can be pasted, using a specified format, into the current range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.CanPaste (BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->CanPaste(m_pTextSelection2, pVar, Format, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the specified range can be edited.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.CanEdit () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->CanEdit(m_pTextSelection2, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the case of letters in this range according to the Type parameter.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.ChangeCase (BYVAL nType AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->ChangeCase(m_pTextSelection2, nType))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves screen coordinates for the start or end character position in the text range,
' along with the intra-line position.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetPoint (BYVAL nType AS LONG, BYVAL px AS LONG PTR, BYVAL py AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->GetPoint(m_pTextSelection2, nType, px, py))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the range based on a specified point at or up through (depending on Extend) the
' point (x, y) aligned according to Type.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetPoint (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG, BYVAL Extend AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetPoint(m_pTextSelection2, x, y, nType, Extend))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the specified range into view.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.ScrollIntoView (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->ScrollIntoView(m_pTextSelection2, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a pointer to the embedded object at the start of the specified range, that is,
' at cpFirst. The range must either be an insertion point or it must select only the embedded object.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetEmbeddedObject () AS IUnknown PTR
   DIM pObject AS IUnknown PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetEmbeddedObject(m_pTextSelection2, @pObject))
   RETURN pObject
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text selection flags.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetFlags () AS LONG
   DIM pFlags AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetFlags(m_pTextSelection2, @pFlags))
   RETURN pFlags
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text selection flags.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetFlags (BYVAL Flags AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetFlags(m_pTextSelection2, Flags))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The selection type. The method returns pType with one of the values in the following table.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetType () AS LONG
   DIM nType AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetType(m_pTextSelection2, @nType))
   RETURN nType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Generalizes the functionality of the Left Arrow key.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveLeft (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveLeft(m_pTextSelection2, Unit, Count, Extend, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Generalizes the functionality of the Right Arrow key.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveRight (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveRight(m_pTextSelection2, Unit, Count, Extend, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Mimics the functionality of the Up Arrow and Page Up keys.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveUp (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveUp(m_pTextSelection2, Unit, Count, Extend, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Mimics the functionality of the Down Arrow and Page Down keys.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.MoveDown (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->MoveDown(m_pTextSelection2, Unit, Count, Extend, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Generalizes the functionality of the Home key.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.HomeKey (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->HomeKey(m_pTextSelection2, Unit, Extend, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Mimics the functionality of the End key.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.EndKey (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->EndKey(m_pTextSelection2, Unit, Extend, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Types the string given by bstr at this selection as if someone typed it. This is similar
' to the underlying SetText method, but is sensitive to the Insert/Overtype key state and
' UI settings like AutoCorrect and smart quotes.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.TypeText (BYVAL bstr AS AFX_BSTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->TypeText(m_pTextSelection2, bstr))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the count of characters in a range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetCch () AS LONG
   DIM cch AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetCch(m_pTextSelection2, @cch))
   RETURN cch
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a cells object with the parameters of cells in the currently selected table row or column.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetCells () AS IUnknown PTR
   DIM pCells AS IUnknown PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetCells(m_pTextSelection2, @pCells))
   RETURN pCells
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a cells object with the parameters of cells in the currently selected table row or column.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetColumn () AS IUnknown PTR
   DIM pColumn AS IUnknown PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetColumn(m_pTextSelection2, @pColumn))
   RETURN pColumn
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the count of subranges, including the active subrange in the current range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetCount () AS LONG
   DIM Count AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetCount(m_pTextSelection2, @Count))
   RETURN Count
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a duplicate of a range object.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetDuplicate2 () AS ITextRange2 PTR
   DIM pRange AS ITextRange2 PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetDuplicate2(m_pTextSelection2, @pRange))
   RETURN pRange
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an ITextFont2 object with the character attributes of the current range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetFont2 () AS ITextFont2 PTR
   DIM pFont AS ITextFont2 PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetFont2(m_pTextSelection2, @pFont))
   RETURN pFont
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the character formatting attributes of the range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetFont2 (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetFont2(m_pTextSelection2, pFont))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an ITextSelection2 object with the current range's formatted text.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetFormattedText2 () AS ITextRange2 PTR
   DIM pRange AS ITextRange2 PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetFormattedText2(m_pTextSelection2, @pRange))
   RETURN pRange
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets an ITextSelection2 object with the current range's formatted text.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetFormattedText2 (BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetFormattedText2(m_pTextSelection2, pRange))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the gravity of this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetGravity () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetGravity(m_pTextSelection2, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the gravity of this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetGravity (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetGravity(m_pTextSelection2, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an ITextPara2 object with the paragraph attributes of a range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetPara2 () AS ITextPara2 PTR
   DIM pPara AS ITextPara2 PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetPara2(m_pTextSelection2, @pPara))
   RETURN pPara
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets an ITextPara2 object with the paragraph attributes of a range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetPara2 (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetPara2(m_pTextSelection2, pPara))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the row properties in the currently selected row.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetRow () AS ITextRow PTR
   DIM pRow AS ITextRow PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetRow(m_pTextSelection2, @pRow))
   RETURN pRow
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character position of the start of the paragraph that contains the range's
' start character position.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetStartPara () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetStartPara(m_pTextSelection2, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the table properties in the currently selected table.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetTable () AS IUnknown PTR
   DIM pTable AS IUnknown PTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetTable(m_pTextSelection2, @pTable))
   RETURN pTable
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the URL text associated with a range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetURL () AS CBSTR
   DIM pbstr AS AFX_BSTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetURL(m_pTextSelection2, @pbstr))
   RETURN pbstr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text in this range to that of the specified URL.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetURL (BYVAL bstr AS AFX_BSTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetURL(m_pTextSelection2, bstr))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a subrange to this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.AddSubrange (BYVAL cp1 AS LONG, BYVAL cp2 AS LONG, BYVAL Activate AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->AddSubrange(m_pTextSelection2, cp1, cp2, Activate))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts the linear-format math in a range to a built-up form, or modifies the current built-up form.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.BuildUpMath (BYVAL Flags AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->BuildUpMath(m_pTextSelection2, Flags))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a subrange from a range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.DeleteSubrange (BYVAL cpFirst AS LONG, BYVAL cpLim AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->DeleteSubrange(m_pTextSelection2, cpFirst, cpLim))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches for math inline functions in text as specified by a source range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Find (BYVAL pRange AS ITextRange2 PTR, BYVAL Count AS LONG, BYVAL Flags AS LONG) AS HRESULT
   DIM Delta AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->Find(m_pTextSelection2, pRange, Count, Flags, @Delta))
   RETURN Delta
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character at the specified offset from the end of this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetChar2 (BYVAL Offset AS LONG) AS LONG
   DIM Char AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetChar2(m_pTextSelection2, @Char, Offset))
   RETURN Char
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the drop-cap parameters of the paragraph that contains this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetDropCap (BYVAL pcLine AS LONG PTR, BYVAL pPosition AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->GetDropCap(m_pTextSelection2, pcLine, pPosition))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the properties of the inline object at the range active end.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetInlineObject (BYVAL pcLine AS LONG PTR, BYVAL pType AS LONG PTR, BYVAL pAlign AS LONG PTR, BYVAL pChar AS LONG PTR, BYVAL pChar1 AS LONG PTR, _
BYVAL pChar2 AS LONG PTR, BYVAL pCount AS LONG PTR,  BYVAL pTeXStyle AS LONG PTR, BYVAL pcCol AS LONG PTR, BYVAL pLevel AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->GetInlineObject(m_pTextSelection2, pcLine, pType, pAlign, pChar, pChar1, pChar2, pCount, pTeXStyle, pcCol))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the drop-cap parameters of the paragraph that contains this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetProperty(m_pTextSelection2, nType, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a rectangle of the specified type for the current range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetRect (BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, BYVAL pTop AS LONG PTR, _
BYVAL pRight AS LONG PTR, BYVAL pBottom AS LONG PTR, BYVAL pHit AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->GetRect(m_pTextSelection2, nType, pLeft, pTop, pRight, pBottom, pHit))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a subrange in a range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetSubrange (BYVAL iSubrange AS LONG, BYVAL pcpFirst AS LONG PTR, BYVAL pcpLim AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->GetSubrange(m_pTextSelection2, iSubrange, pcpFirst, pcpLim))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text in this range according to the specified conversion flags.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetText2 (BYVAL Flags AS LONG) AS CBSTR
   DIM pText2 AS AFX_BSTR
   this.SetResult(m_pTextSelection2->lpvtbl->GetText2(m_pTextSelection2, Flags, @pText2))
   RETURN pText2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts and replaces the hexadecimal number at the end of this range to a Unicode character.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.HexToUnicode () AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->HexToUnicode(m_pTextSelection2))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a table in a range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.InsertTable (BYVAL cCol AS LONG, BYVAL cRow AS LONG, BYVAL AutoFit AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->InsertTable(m_pTextSelection2, cCol, cRow, AutoFit))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates the built-up math, ruby, and other inline objects in this range to linearized form.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.Linearize (BYVAL Flags AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->Linearize(m_pTextSelection2, Flags))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Makes the specified subrange the active subrange of this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetActiveSubrange (BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetActiveSubrange(m_pTextSelection2, cpAnchor, cpActive))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the drop-cap parameters for the paragraph that contains the current range.
' Not implemented.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetDropCap (BYVAL cLine AS LONG, BYVAL Position AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetDropCap(m_pTextSelection2, cLine, Position))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the drop-cap parameters for the paragraph that contains the current range.
' Not implemented.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetProperty(m_pTextSelection2, nType, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text in this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetText2 (BYVAL Flags AS LONG, BYVAL bstr AS AFX_BSTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetText2(m_pTextSelection2, Flags, bstr))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts the Unicode character(s) preceding the start position of this text range to a
' hexadecimal number, and selects it.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.UnicodeToHex () AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->UnicodeToHex(m_pTextSelection2))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets or inserts the properties of an inline object for a degenerate range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.SetInlineObject (BYVAL nType AS Long, BYVAL Align AS LONG, BYVAL Char AS LONG, _
BYVAL Char1 AS LONG, BYVAL Char2 AS LONG, BYVAL Count AS LONG, BYVAL TeXStyle AS LONG, BYVAL cCol AS LONG) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->SetInlineObject(m_pTextSelection2, nType, Align, Char, Char1, Char2, Count, TeXStyle, cCol))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the math function type associated with the specified math function name.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.GetMathFunctionType (BYVAL bstr AS AFX_BSTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextSelection2->lpvtbl->GetMathFunctionType(m_pTextSelection2, bstr, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts an image into this range.
' ========================================================================================
PRIVATE FUNCTION CTextSelection2.InsertImage (BYVAL width_ AS LONG, BYVAL height AS LONG, BYVAL ascent AS LONG, _
BYVAL nType AS LONG, BYVAL bstrAltText AS AFX_BSTR, BYVAL pStream AS IStream PTR) AS HRESULT
   this.SetResult(m_pTextSelection2->lpvtbl->InsertImage(m_pTextSelection2, width_, height, ascent, nType, bstrAltText, pStream))
   RETURN m_Result
END FUNCTION
' ========================================================================================
