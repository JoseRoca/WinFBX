' ########################################################################################
' Microsoft Windows
' File: CRichEditCtrl.inc
' Contents: Rich Edit control
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "Afx/CWindow.inc"
#include once "Afx/CTextObjectBase.inc"
#include once "Afx/AfxPrinter.inc"

NAMESPACE Afx

' ========================================================================================
' Custom structure used by the CRichEditCtrl.LoadRtfFromResource method.
' ========================================================================================
TYPE AFX_CRICHEDITCTRL_CUSTOMDATA
   pData  AS BYTE PTR
   nLen   AS LONG
   curPos AS LONG
END TYPE
' ========================================================================================

' ========================================================================================
' CRichEditCtrl class
' ========================================================================================
TYPE CRichEditCtrl EXTENDS CTextObjectBase

   Private:
      DIM m_Result AS HRESULT           ' // Last result code
      DIM m_hRichEdit AS HWND           ' // Rich Edit cotroll handle
      DIM m_hInstance AS HINSTANCE      ' // Instance handle
      DIM m_ScalingRatio AS SINGLE      ' // Scaling ratio

   Public:
      DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszTitle AS WSTRING = "", _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
         BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION hRichEdit () AS HWND
      DECLARE PROPERTY ScalingRatio (BYVAL ratio AS SINGLE)
      DECLARE PROPERTY ScalingRatio () AS SINGLE
      DECLARE FUNCTION SetWysiwygPrint (BYREF wszPrinterName AS WSTRING = "") AS BOOLEAN

      ' // Static callback procedures
      DECLARE STATIC FUNCTION CRichEditCtrlProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM, BYVAL uIdSubclass AS UINT_PTR, BYVAL dwRefData AS DWORD_PTR) AS LRESULT
      DECLARE STATIC FUNCTION CRichEditCtrlGetTextCallback (BYVAL dwCookie AS DWORD_PTR, BYVAL pbBuff AS BYTE PTR, BYVAL cb AS LONG, BYVAL pcb AS LONG PTR) AS DWORD
      DECLARE STATIC FUNCTION CRichEditCtrlEnumFontFamProcW (BYVAL lpelf AS ENUMLOGFONTW PTR, BYVAL lpntm AS NEWTEXTMETRICW PTR, _
         BYVAL FontType AS LONG, BYVAL lplf AS LOGFONTW PTR) AS LONG
      DECLARE STATIC FUNCTION CRichEditCtrlLoadRtfFromFileCallback (BYVAL hFile AS HANDLE, BYVAL lpBuff AS BYTE PTR, _
         BYVAL cb AS LONG, BYVAL pcb AS LONG PTR) AS UINT
      DECLARE STATIC FUNCTION CRichEditCtrlLoadRtfFromResourceCallback (BYVAL pCustData AS AFX_CRICHEDITCTRL_CUSTOMDATA PTR, _
         BYVAL lpBuff AS BYTE PTR, BYVAL cb AS LONG, BYVAL pcb AS LONG PTR) AS DWORD

      ' // Propereties
      DECLARE PROPERTY AutoCorrectProc () AS LONG_PTR
      DECLARE PROPERTY AutoCorrectProc (BYVAL pfn AS LONG_PTR)
      DECLARE PROPERTY AutoUrlDetect () AS BOOLEAN
      DECLARE PROPERTY AutoUrlDetect (BYVAL fUrlDetect AS LONG)
      DECLARE PROPERTY BidiOptions () AS .BIDIOPTIONS
      DECLARE PROPERTY BidiOptions (BYREF opt AS .BIDIOPTIONS)
      DECLARE PROPERTY CaretPos () AS DWORD
      DECLARE PROPERTY CaretPos (BYVAL dwPos AS DWORD)
      DECLARE PROPERTY CharFormat (BYVAL fOption AS DWORD) AS CHARFORMATW
      DECLARE PROPERTY CharFormat (BYVAL chfmt AS DWORD, BYREF cf AS CHARFORMATW)
      DECLARE PROPERTY CharFromPos (BYREF ptl AS .POINTL) AS LONG
      DECLARE PROPERTY DefaultCharFormat () AS CHARFORMATW
      DECLARE PROPERTY DefaultCharFormat (BYREF cf AS CHARFORMATW)
      DECLARE PROPERTY SelectionCharFormat () AS CHARFORMATW
      DECLARE PROPERTY SelectionCharFormat (BYREF cf AS CHARFORMATW)
      DECLARE PROPERTY WordCharFormat () AS CHARFORMATW
      DECLARE PROPERTY WordCharFormat (BYREF cf AS CHARFORMATW)
      DECLARE PROPERTY CTFModeBias () AS LONG
      DECLARE PROPERTY CTFModeBias (BYVAL nModeBias AS LONG)
      DECLARE PROPERTY CTFOpenStatus () AS BOOLEAN
      DECLARE PROPERTY CTFOpenStatus (BYVAL fTSFkbd AS LONG)
      DECLARE PROPERTY EditStyle () AS DWORD
      DECLARE PROPERTY EditStyle (BYVAL fStyle AS LONG, BYVAL fMask AS LONG)
      DECLARE PROPERTY EditStyleEx () AS DWORD
      DECLARE PROPERTY EditStyleEx (BYVAL fStyle AS LONG, BYVAL fMask AS LONG)
      DECLARE PROPERTY EllipsisMode () AS DWORD
      DECLARE PROPERTY EllipsisMode (BYVAL fMode AS DWORD)
      DECLARE PROPERTY EllipsisState () AS BOOLEAN
      DECLARE PROPERTY EventMask () AS DWORD
      DECLARE PROPERTY EventMask (BYVAL fMask AS DWORD)
      DECLARE PROPERTY FirstVisibleLine () AS LONG
      DECLARE PROPERTY HyphenateInfo () AS .HYPHENATEINFO
      DECLARE PROPERTY HyphenateInfo (BYREF hinfo AS .HYPHENATEINFO)
      DECLARE PROPERTY IMEModeBias () AS DWORD
      DECLARE PROPERTY IMEModeBias (BYVAL nModeBias AS LONG)
      DECLARE PROPERTY IMEOptions () AS DWORD
      DECLARE PROPERTY IMECompMode () AS DWORD
      DECLARE PROPERTY IMEOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
      DECLARE PROPERTY LangOptions () AS DWORD
      DECLARE PROPERTY LangOptions (BYVAL lgoptions AS LONG)
      DECLARE PROPERTY IMEProperty (BYVAL figp AS DWORD) AS DWORD
      DECLARE PROPERTY LeftMargin (BYVAL nWidth AS LONG)
      DECLARE PROPERTY LimitText () AS LONG
      DECLARE PROPERTY LimitText (BYVAL chMax AS DWORD)
      DECLARE PROPERTY Modify () AS LONG
      DECLARE PROPERTY Modify (BYVAL fModify AS LONG)
      DECLARE PROPERTY Options () AS DWORD
      DECLARE PROPERTY Options (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
      DECLARE PROPERTY PageRotate () AS DWORD
      DECLARE PROPERTY PageRotate (BYVAL txtlayout AS LONG)
      DECLARE PROPERTY ParaFormat () AS .PARAFORMAT
      DECLARE PROPERTY ParaFormat (BYREF pfmt AS .PARAFORMAT)
      DECLARE PROPERTY PasswordChar () AS LONG
      DECLARE PROPERTY PasswordChar (BYVAL dwchar AS DWORD)
      DECLARE PROPERTY Punctuation (BYVAL punctp AS DWORD) AS .PUNCTUATION
      DECLARE PROPERTY Punctuation (BYVAL ptype AS LONG, BYREF punct AS .PUNCTUATION)
      DECLARE PROPERTY Rect () AS .RECT
      DECLARE PROPERTY Rect (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
      DECLARE PROPERTY RectNP (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
      DECLARE PROPERTY RightMargin (BYVAL nWidth AS LONG)
      DECLARE PROPERTY ScrollPos () AS .POINT
      DECLARE PROPERTY ScrollPos (BYREF pt AS .POINT)
      DECLARE PROPERTY StoryType (BYVAL Index AS DWORD) AS DWORD
      DECLARE PROPERTY StoryType (BYVAL Index AS LONG, BYVAL dwType AS DWORD)
      DECLARE PROPERTY Text () AS CWSTR
      DECLARE PROPERTY Text (BYREF wszText AS WSTRING)
      DECLARE PROPERTY TextColor () AS COLORREF
      DECLARE PROPERTY TextColor (BYVAL crTxtColor AS COLORREF)
      DECLARE PROPERTY TextFontName () AS CWSTR
      DECLARE PROPERTY TextFontName (BYREF wszFaceName AS WSTRING)
      DECLARE PROPERTY TextHeight () AS LONG
      DECLARE PROPERTY TextHeight (BYVAL ptSize AS LONG)
      DECLARE PROPERTY TextMode () AS DWORD
      DECLARE PROPERTY TextMode (BYVAL pvalues AS LONG)
      DECLARE PROPERTY TextOffset () AS LONG
      DECLARE PROPERTY TextOffset (BYVAL offset AS LONG)
      DECLARE PROPERTY TouchOptions (BYVAL opt AS LONG PTR) AS DWORD
      DECLARE PROPERTY TouchOptions (BYVAL opt AS LONG, BYVAL fEnable AS LONG)
      DECLARE PROPERTY TypographyOptions () AS DWORD
      DECLARE PROPERTY TypographyOptions (BYVAL pto AS LONG, BYVAL fMask AS LONG)

      ' // Procedures
      DECLARE SUB AddLF (BYVAL atEnd AS BOOLEAN = FALSE)
      DECLARE SUB AddCR (BYVAL atEnd AS BOOLEAN = FALSE)
      DECLARE SUB AddCRLF (BYVAL atEnd AS BOOLEAN = FALSE)
      DECLARE FUNCTION CallAutocorrectProc (BYVAL char AS WCHAR) AS LONG
      DECLARE FUNCTION CanPaste (BYVAL clipformat AS LONG) AS BOOLEAN
      DECLARE FUNCTION CanRedo () AS BOOLEAN
      DECLARE FUNCTION CanUndo () AS BOOLEAN
      DECLARE FUNCTION DisableAutoUrlDetect () AS BOOLEAN
      DECLARE FUNCTION EnableAutoUrlDetect (BYVAL fUrlDetect AS LONG) AS HRESULT
      DECLARE FUNCTION DisableWordWrap (BYVAL LineWidth AS LONG = 32767) AS BOOLEAN
      DECLARE FUNCTION EnableWordWrap () AS BOOLEAN
      DECLARE FUNCTION DisplayBand (BYREF rc AS .RECT) AS BOOLEAN
      DECLARE SUB EmptyUndoBuffer ()
      DECLARE FUNCTION ExGetSel OVERLOAD () AS CHARRANGE
      DECLARE SUB ExGetSel OVERLOAD (BYREF cpMin AS LONG, BYVAL cpMax AS LONG)
      DECLARE SUB ExLimitText (BYVAL dwLimit AS DWORD)
      DECLARE FUNCTION ExLineFromChar (BYVAL iIndex AS LONG) AS LONG
      DECLARE FUNCTION ExSetSel OVERLOAD (BYREF cr AS CHARRANGE) AS DWORD
      DECLARE SUB ExSetSel OVERLOAD (BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1)
      DECLARE FUNCTION FindText OVERLOAD (BYVAL fOptions AS DWORD, BYREF ftw AS FINDTEXTW) AS LONG
      DECLARE FUNCTION FindText OVERLOAD (BYVAL fOptions AS DWORD = FR_DOWN, _
         BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1, BYREF wszText AS WSTRING) AS LONG
      DECLARE FUNCTION FindTextEx OVERLOAD (BYVAL fOptions AS DWORD, BYREF ftexw AS FINDTEXTEXW) AS LONG
      DECLARE FUNCTION FindTextEx OVERLOAD (BYVAL fOptions AS DWORD = FR_DOWN, BYVAL cpMin AS LONG = 0, _
         BYVAL cpMax AS LONG = -1, BYREF wszText AS WSTRING) AS CHARRANGE
      DECLARE FUNCTION FindWordBreak (BYVAL fOperation AS DWORD, BYVAL dwStartPos AS DWORD) AS DWORD
      DECLARE FUNCTION FormatRange (BYVAL fRender AS LONG, BYREF fr AS FORMATRANGE) AS DWORD
      DECLARE FUNCTION GetAutoCorrectProc () AS LONG_PTR
      DECLARE FUNCTION SetAutoCorrectProc (BYVAL pfn AS LONG_PTR) AS HRESULT
      DECLARE FUNCTION GetAutoUrlDetect () AS BOOLEAN
      DECLARE FUNCTION SetAutoUrlDetect (BYVAL fUrlDetect AS LONG) AS HRESULT
      DECLARE FUNCTION GetBidiOptions () AS .BIDIOPTIONS
      DECLARE SUB SetBidiOptions (BYREF opt AS .BIDIOPTIONS)
      DECLARE FUNCTION GetCaretPos () AS DWORD
      DECLARE SUB SetCaretPos (BYVAL dwPos AS DWORD)
      DECLARE FUNCTION GetCharFormat (BYVAL fOption AS DWORD) AS CHARFORMATW
      DECLARE FUNCTION SetCharFormat (BYVAL chfmt AS DWORD, BYREF cf AS CHARFORMATW) AS BOOLEAN
      DECLARE FUNCTION GetDefaultCharFormat () AS CHARFORMATW
      DECLARE FUNCTION SetDefaultCharFormat (BYREF cf AS CHARFORMATW) AS BOOLEAN
      DECLARE FUNCTION GetSelectionCharFormat () AS CHARFORMATW
      DECLARE FUNCTION SetSelectionCharFormat (BYREF cf AS CHARFORMATW) AS BOOLEAN
      DECLARE FUNCTION GetCharFromPos (BYREF ptl AS .POINTL) AS LONG
      DECLARE FUNCTION GetCTFModeBias () AS LONG
      DECLARE FUNCTION SetCTFModeBias (BYVAL nModeBias AS LONG) AS HRESULT
      DECLARE FUNCTION GetCTFOpenStatus () AS BOOLEAN
      DECLARE FUNCTION SetCTFOpenStatus (BYVAL fTSFkbd AS LONG) AS BOOLEAN
      DECLARE FUNCTION GetEditStyle () AS DWORD
      DECLARE FUNCTION SetEditStyle (BYVAL fStyle AS LONG, BYVAL fMask AS LONG) AS HRESULT
      DECLARE FUNCTION GetEditStyleEx () AS DWORD
      DECLARE FUNCTION SetEditStyleEx (BYVAL fStyle AS LONG, BYVAL fMask AS LONG) AS HRESULT
      DECLARE FUNCTION GetEllipsisMode () AS DWORD
      DECLARE FUNCTION SetEllipsisMode (BYVAL fMode AS DWORD) AS BOOLEAN
      DECLARE FUNCTION GetEllipsisState () AS BOOLEAN
      DECLARE FUNCTION GetEventMask () AS DWORD
      DECLARE FUNCTION SetEventMask (BYVAL fMask AS DWORD) AS HRESULT
      DECLARE FUNCTION GetFirstVisibleLine () AS LONG
      DECLARE FUNCTION GetHyphenateInfo () AS .HYPHENATEINFO
      DECLARE FUNCTION SetHyphenateInfo (BYREF info AS .HYPHENATEINFO) AS HRESULT
      DECLARE FUNCTION GetIMEColor (BYVAL rgCmpclr AS .COMPCOLOR PTR) AS LONG
      DECLARE FUNCTION SetIMEColor (BYVAL pcompcolor AS .COMPCOLOR PTR) AS LONG
      DECLARE FUNCTION GetIMECompMode () AS DWORD
      DECLARE FUNCTION GetIMECompText (BYREF ict AS .IMECOMPTEXT, BYVAL buffer AS ANY PTR) AS DWORD
      DECLARE FUNCTION GetIMEModeBias () AS DWORD
      DECLARE FUNCTION SetIMEModeBias (BYVAL nModeBias AS LONG) AS HRESULT
      DECLARE FUNCTION GetIMEOptions () AS DWORD
      DECLARE FUNCTION SetIMEOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG) AS HRESULT
      DECLARE FUNCTION GetIMEProperty (BYVAL figp AS DWORD) AS DWORD
      DECLARE FUNCTION GetLangOptions () AS DWORD
      DECLARE FUNCTION SetLangOptions (BYVAL lgoptions AS LONG) AS HRESULT
      DECLARE FUNCTION GetLimitText () AS LONG
      DECLARE SUB SetLimitText (BYVAL chMax AS DWORD)
      DECLARE FUNCTION GetModify () AS LONG
      DECLARE SUB SetModify (BYVAL fModify AS LONG)
      DECLARE FUNCTION GetOptions () AS DWORD
      DECLARE FUNCTION SetOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG) AS DWORD
      DECLARE FUNCTION GetPageRotate () AS DWORD
      DECLARE FUNCTION SetPageRotate (BYVAL txtlayout AS LONG) AS DWORD
      DECLARE FUNCTION GetParaFormat () AS .PARAFORMAT
      DECLARE FUNCTION SetParaFormat (BYREF pfmt AS .PARAFORMAT) AS BOOLEAN
      DECLARE FUNCTION GetPasswordChar () AS LONG
      DECLARE SUB SetPasswordChar (BYVAL dwchar AS DWORD)
      DECLARE FUNCTION GetPunctuation (BYVAL punctype AS DWORD) AS .PUNCTUATION
      DECLARE FUNCTION SetPunctuation (BYVAL punctype AS LONG, BYREF punct AS .PUNCTUATION) AS BOOLEAN
      DECLARE FUNCTION GetRect () AS .RECT
      DECLARE SUB SetRect (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
      DECLARE SUB SetRectNP (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
      DECLARE FUNCTION GetScrollPos () AS .POINT
      DECLARE FUNCTION SetScrollPos (BYREF pt AS .POINT) AS HRESULT
      DECLARE FUNCTION GetScalingRatio () AS SINGLE
      DECLARE SUB SetScalingRatio (BYVAL ratio AS SINGLE)
      DECLARE FUNCTION GetStoryType (BYVAL Index AS DWORD) AS DWORD
      DECLARE FUNCTION SetStoryType (BYVAL Index AS LONG, BYVAL dwType AS DWORD) AS DWORD
      DECLARE FUNCTION GetText () AS CWSTR
      DECLARE FUNCTION SetText (BYREF wszText AS WSTRING) AS BOOLEAN
      DECLARE FUNCTION GetTextColor () AS COLORREF
      DECLARE FUNCTION SetTextColor (BYVAL crTxtColor AS COLORREF) AS BOOLEAN
      DECLARE FUNCTION GetTextEx (BYREF gtex AS GETTEXTEX, BYVAL buffer AS ANY PTR) AS DWORD
      DECLARE FUNCTION SetTextEx (BYREF stex AS .SETTEXTEX, BYVAL buffer AS ANY PTR) AS DWORD
      DECLARE FUNCTION GetTextFontName () AS CWSTR
      DECLARE FUNCTION SetTextFontName (BYREF wszFaceName AS WSTRING) AS BOOLEAN
      DECLARE FUNCTION GetTextHeight () AS LONG
      DECLARE FUNCTION SetTextHeight (BYVAL ptSize AS LONG) AS BOOLEAN
      DECLARE FUNCTION GetTextMode () AS DWORD
      DECLARE FUNCTION SetTextMode (BYVAL values AS LONG) AS BOOLEAN
      DECLARE FUNCTION GetTextOffset () AS LONG
      DECLARE FUNCTION SetTextOffset (BYVAL offset AS LONG) AS BOOLEAN
      DECLARE FUNCTION GetTouchOptions (BYVAL opt AS LONG PTR) AS DWORD
      DECLARE FUNCTION SetTouchOptions (BYVAL opt AS LONG, BYVAL fEnable AS LONG) AS HRESULT
      DECLARE FUNCTION GetTypographyOptions () AS DWORD
      DECLARE FUNCTION SetTypographyOptions (BYVAL pto AS LONG, BYVAL fMask AS LONG) AS BOOLEAN

      DECLARE FUNCTION GetLine (BYVAL which AS DWORD) AS CWSTR
      DECLARE FUNCTION GetLineCount () AS LONG
      DECLARE FUNCTION GetOleInterface () AS IRichEditOle PTR
      DECLARE FUNCTION GetRedoName () AS LONG
      DECLARE FUNCTION GetSel OVERLOAD (BYREF dwStartPos AS DWORD, BYREF dwEndPos AS DWORD) AS LONG
      DECLARE FUNCTION GetSel OVERLOAD () AS .POINTL
      DECLARE FUNCTION GetSelText () AS CWSTR
      DECLARE FUNCTION GetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
      DECLARE FUNCTION GetTextLength () AS LONG
      DECLARE FUNCTION GetTextLengthEx OVERLOAD (BYREF gtex AS .GETTEXTLENGTHEX) AS LONG
      DECLARE FUNCTION GetTextLengthEx OVERLOAD (BYVAL dwFlags AS DWORD = GTL_DEFAULT) AS LONG
      DECLARE FUNCTION GetTextRange OVERLOAD (BYREF trg AS .TEXTRANGEW) AS DWORD
      DECLARE FUNCTION GetTextRange OVERLOAD (BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1) AS CWSTR
      DECLARE FUNCTION GetThumb () AS LONG
      DECLARE FUNCTION GetUndoName () AS DWORD
      DECLARE FUNCTION GetWordBreakProc () AS LONG_PTR
      DECLARE SUB SetWordBreakProc (BYVAL pfn AS LONG_PTR)
      DECLARE FUNCTION GetWordBreakProcEx () AS LONG_PTR
      DECLARE FUNCTION SetWordBreakProcEx (BYVAL pfn AS LONG_PTR) AS LONG_PTR
      DECLARE FUNCTION GetWordCharFormat () AS CHARFORMATW
      DECLARE FUNCTION SetWordCharFormat (BYREF cf AS CHARFORMATW) AS BOOLEAN
      DECLARE FUNCTION GetWordWrapMode () AS DWORD
      DECLARE FUNCTION SetWordWrapMode (BYVAL pvalues AS LONG) AS LONG
      DECLARE FUNCTION GetZoom (BYREF pzNum AS DWORD, BYREF pzDen AS DWORD) AS LONG
      DECLARE SUB HideSelection (BYVAL fHide AS DWORD)
      DECLARE FUNCTION InsertImage OVERLOAD (BYREF ip AS RICHEDIT_IMAGE_PARAMETERS) AS DWORD
      DECLARE FUNCTION InsertImage OVERLOAD (BYREF wszFileName AS WSTRING, BYVAL xWidth AS LONG = 0, BYVAL yHeight AS LONG = 0, _
         BYVAL Ascent AS LONG = 0, BYVAL nType AS LONG = TA_BASELINE, BYREF wszAlternateText AS WSTRING = "") AS BOOLEAN
      DECLARE FUNCTION InsertTable (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
      DECLARE FUNCTION IsIME () AS LONG
      DECLARE FUNCTION IsTextBold () AS BOOLEAN
      DECLARE SUB SetTextBold ()
      DECLARE FUNCTION IsTextItalic () AS BOOLEAN
      DECLARE SUB SetTextItalic ()
      DECLARE FUNCTION IsTextStrikeout () AS BOOLEAN
      DECLARE SUB SetTextStrikeout ()
      DECLARE FUNCTION IsTextUnderline () AS BOOLEAN
      DECLARE SUB SetTextUnderline ()
      DECLARE FUNCTION LineFromChar (BYVAL index AS DWORD) AS LONG
      DECLARE FUNCTION LineIndex (BYVAL nLine AS LONG) AS LONG
      DECLARE FUNCTION LineLength (BYVAL index AS DWORD) AS LONG
      DECLARE FUNCTION LineScroll (BYVAL y AS LONG) AS LONG
      DECLARE SUB PasteSpecial OVERLOAD (BYVAL clpfmt AS DWORD, BYREF rps AS REPASTESPECIAL)
      DECLARE SUB PasteSpecial OVERLOAD (BYVAL clpfmt AS DWORD, BYVAL dwAspect AS DWORD, BYVAL hMF AS HMETAFILE)
      DECLARE FUNCTION PosFromChar (BYVAL index as DWORD) AS .POINTL
      DECLARE SUB Reconversion ()
      DECLARE FUNCTION Redo () AS LONG
      DECLARE SUB ReplaceSel (BYVAL bCanBeUndone AS LONG = TRUE, BYREF wszText AS WSTRING)
      DECLARE SUB RequestResize ()
      DECLARE FUNCTION Scroll (BYVAL nAction AS LONG) AS LONG
      DECLARE SUB ScrollCaret ()
      DECLARE FUNCTION SelectionType () AS LONG
      DECLARE FUNCTION SetBkgndColor (BYVAL pSysColor AS DWORD, BYVAL pBkColor AS DWORD) AS DWORD
      DECLARE FUNCTION SetFontSize (BYVAL ptsize AS LONG) AS LONG
      DECLARE SUB SetMargins (BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
      DECLARE SUB SetLeftMargin (BYVAL nWidth AS LONG)
      DECLARE SUB SetRightMargin (BYVAL nWidth AS LONG)
      DECLARE FUNCTION SetOleCallback (BYVAL pCallback AS ANY PTR) AS LONG
      DECLARE SUB SetPalette (BYVAL newPalette AS HPALETTE)
      DECLARE FUNCTION SetReadOnly (BYVAL fReadOnly AS LONG) AS BOOLEAN
      DECLARE SUB SetSel OVERLOAD (BYVAL nStart AS LONG, BYVAL nEnd AS LONG)
      DECLARE SUB SetSel OVERLOAD (BYREF pt AS POINTL)
      DECLARE FUNCTION SetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
      DECLARE FUNCTION SetTabStops (BYVAL nTabs AS LONG, BYVAL rgTabStops AS LONG_PTR) AS LONG
      DECLARE FUNCTION SetTargetDevice (BYVAL hDC AS HDC, BYVAL lnwidth AS LONG) AS BOOLEAN
      DECLARE FUNCTION SetUIAName (BYREF wszName AS WSTRING) AS DWORD
      DECLARE FUNCTION SetUndoLimit (BYVAL maxactions AS DWORD) AS DWORD
      DECLARE FUNCTION SetZoom (BYVAL zNum AS DWORD, BYVAL zDen AS DWORD) AS LONG
      DECLARE SUB ShowScrollBar (BYVAL nScrollBar AS DWORD, BYVAL fShow AS LONG)
      DECLARE FUNCTION StopGroupTyping () AS DWORD
      DECLARE FUNCTION StreamIn (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
      DECLARE FUNCTION StreamOut (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
      DECLARE FUNCTION Undo () AS LONG

      ' // Helper methods and properties
      DECLARE FUNCTION SetFont (BYREF wszFaceName AS WSTRING, BYVAL ptsize AS LONG) AS HRESULT
      DECLARE FUNCTION GetRtf (BYVAL dwOptionsAndFlags AS DWORD = SF_RTF) AS STRING
      DECLARE FUNCTION LoadRtfFromFile (BYREF wszFileName AS WSTRING) AS BOOLEAN
      DECLARE FUNCTION LoadRtfFromResource (BYREF wszResourceName AS WSTRING) AS BOOLEAN
      DECLARE FUNCTION GetRtfFromFile (BYREF wszFileName AS WSTRING) AS STRING
      DECLARE FUNCTION AppendRtfFile (BYREF wszFileName AS WSTRING) AS DWORD
      DECLARE FUNCTION InsertRtfFile (BYREF wszFileName AS WSTRING, BYVAL nPos AS DWORD) AS DWORD
      DECLARE FUNCTION InsertObject (BYREF wszFileName AS WSTRING) AS BOOLEAN
      DECLARE FUNCTION SaveRtf (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
         BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
      DECLARE FUNCTION SaveSelRtf (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
         BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
      DECLARE FUNCTION SaveRtfNoObjs (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
         BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
      DECLARE FUNCTION SaveSelRtfNoObjs (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
         BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
      DECLARE FUNCTION SaveText (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
         BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
      DECLARE FUNCTION SaveSelText (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
         BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN

      ' // Wrapper methods and properties for the ITextDocument2 interface
      DECLARE FUNCTION GetDocInterface () AS ITextDocument2 PTR
      DECLARE FUNCTION NewDoc () AS HRESULT
      DECLARE FUNCTION OpenDoc (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG = 0, BYVAL CodePage AS LONG = CP_ACP) AS HRESULT
      DECLARE FUNCTION SaveDoc (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG = 0, BYVAL CodePage AS LONG = CP_ACP) AS HRESULT
      DECLARE FUNCTION DocName () AS CWSTR
      DECLARE PROPERTY Saved () AS BOOLEAN
      DECLARE PROPERTY Saved (BYVAL value AS BOOLEAN)
      DECLARE FUNCTION Freeze () AS LONG
      DECLARE FUNCTION Unfreeze () AS LONG
      DECLARE FUNCTION Update (BYVAL Value AS LONG) AS HRESULT
      DECLARE FUNCTION ChangeCase (BYVAL cpActive AS LONG = 0, BYVAL cpAnchor AS LONG = 0, BYVAL nType AS LONG = tomLowerCase) AS HRESULT
      DECLARE FUNCTION LowerCase (BYVAL nStart AS LONG, BYVAL nLen AS LONG) AS HRESULT
      DECLARE FUNCTION UpperCase (BYVAL nStart AS LONG, BYVAL nLen AS LONG) AS HRESULT
      DECLARE FUNCTION TitleCase (BYVAL nStart AS LONG, BYVAL nLen AS LONG) AS HRESULT
      DECLARE FUNCTION SentenceCase (BYVAL nStart AS LONG, BYVAL nLen AS LONG) AS HRESULT
      DECLARE FUNCTION ToggleCase (BYVAL nStart AS LONG, BYVAL nLen AS LONG) AS HRESULT
   
END TYPE
' ========================================================================================

' ========================================================================================
' Subclassed window procedure
' ========================================================================================
FUNCTION CRichEditCtrl.CRichEditCtrlProc ( _
   BYVAL hwnd   AS HWND, _                 ' // Control window handle
   BYVAL uMsg   AS UINT, _                 ' // Type of message
   BYVAL wParam AS WPARAM, _               ' // First message parameter
   BYVAL lParam AS LPARAM, _               ' // Second message parameter
   BYVAL uIdSubclass AS UINT_PTR, _        ' // The subclass ID
   BYVAL dwRefData AS DWORD_PTR _          ' // Pointer to reference data
   ) AS LRESULT

   SELECT CASE uMsg

      CASE WM_GETDLGCODE
         CREC_DP("CRichEditCtrlProc WM_GETDLGCODE " & WSTR(wParam))
         ' // Needed for VK_TAB to work
         RETURN DLGC_WANTALLKEYS

      CASE WM_KEYDOWN
         ' // Keyboard navigation
         SELECT CASE LOWORD(wParam)
            ' // Rich Edit multiline controls send a WM_CLOSE message to its parent window,
            ' // when the ESC key is pressed, causing the application to shut down.
            ' // Avoid it by cancelling the message.
            CASE VK_ESCAPE
               CREC_DP("CRichEditCtrlProc VK_ESCAPE")
               RETURN 0
            CASE VK_TAB
               CREC_DP("CRichEditCtrlProc VK_TAB")
'               IF HIWORD(.GetKeyState(VK_SHIFT)) = 0 THEN
'                  '// Set the focus in the previous control that has the WS_TABSTOP style
'                  .SetFocus .GetNextDlgTabItem(GetParent(hwnd), hwnd, 0)
'               ELSE
'                  '// Set the focus in the next control that has the WS_TABSTOP style
'                  .SetFocus .GetNextDlgTabItem(GetParent(hwnd), hwnd, -1)
'               END IF
         END SELECT

      CASE WM_DESTROY
         CREC_DP("CRichEditCtrlProc WM_DESTROY")
         ' // REQUIRED: Remove control subclassing
         RemoveWindowSubclass hwnd, @CRichEditCtrlProc, uIdSubclass
         EXIT FUNCTION

      ' // Returns the class pointer
      CASE WM_USER + 100
         RETURN dwRefData

   END SELECT

   ' // Default processing for other messages
   RETURN DefSubclassProc(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' CRichEditCtrl class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CRichEditCtrl (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszTitle AS WSTRING = "", _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)

   ' // Create the control
   IF pWindow THEN
      m_hRichEdit = pWindow->AddControl("RICHEDIT", pWindow->hWindow, cID, wszTitle, _
      x, y, nWidth, nHeight, WS_VISIBLE OR WS_CHILD OR WS_TABSTOP OR ES_LEFT OR WS_HSCROLL OR WS_VSCROLL OR _
         ES_AUTOHSCROLL OR ES_AUTOVSCROLL OR ES_MULTILINE OR ES_WANTRETURN OR ES_NOHIDESEL OR ES_SAVESEL, WS_EX_CLIENTEDGE, _
         NULL, CAST(SUBCLASSPROC, @CRichEditCtrlProc), cID, CAST(DWORD_PTR, @this))
      m_hInstance = pWindow->InstanceHandle
      m_ScalingRatio = pWindow->rxRatio
      ' // Set the ES_EX_ZOOMABLE (&H0010) extended style
      ' // EM_SETEXTENDEDSTYLE = ECM_FIRST + 10
      ' // EM_GETEXTENDEDSTYLE = ECM_FIRST + 11
      DIM dwExStyle AS DWORD = SendMessageW(m_hRichEdit, ECM_FIRST + 11, 0, 0)
      SendMessageW(m_hRichEdit, ECM_FIRST + 10, dwExStyle, &H0010)
   END IF

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CRichEditCtrl class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CRichEditCtrl
   CREC_DP("CRichEditCtrl DESTRUCTOR")
   ' // Destroy the Rich Edit control
   DestroyWindow m_hRichEdit
   m_hRichEdit = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the handle of the control
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.hRichEdit () AS HWND
   FUNCTION = m_hRichEdit
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets/sets the width scaling ratio
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.ScalingRatio () AS SINGLE
   PROPERTY = m_ScalingRatio
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.ScalingRatio (BYVAL ratio AS SINGLE)
   m_ScalingRatio = ratio
   IF m_ScalingRatio = 0 THEN m_ScalingRatio = 1
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetScalingRatio () AS SINGLE
   RETURN m_ScalingRatio
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetScalingRatio (BYVAL ratio AS SINGLE)
   m_ScalingRatio = ratio
   IF m_ScalingRatio = 0 THEN m_ScalingRatio = 1
END SUB
' ========================================================================================

' ========================================================================================
' Sets the target printer device and line width used for "what you see is what you get" (WYSIWYG)
' formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetWysiwygPrint (BYREF wszPrinterName AS WSTRING = "") AS BOOLEAN
   DIM _wszPrinterName AS WSTRING * MAX_PATH = wszPrinterName
   IF LEN(_wszPrinterName) = 0 THEN
      ' // Get the default printer
      DIM buffer AS WSTRING * MAX_PATH
      GetProfileStringW "WINDOWS", "DEVICE", "", buffer, SIZEOF(buffer)
      _wszPrinterName = AfxStrParse(buffer, 1, ",")
   END IF
   ' // Create a device context for the printer
   DIM hDCPrinter AS HDC = CreateDCW(NULL, _wszPrinterName, NULL, NULL)
   ' // Calculate the width in twips
   DIM cx AS LONG, cy AS LONG
   cx = (GetDeviceCaps(hDCPrinter, HORZRES) * 1440) \ GetDeviceCaps(hDCPrinter, LOGPIXELSX)
   ' // Set the target device and line width
   m_Result = SendMessageW(m_hRichEdit, EM_SETTARGETDEVICE, cast(WPARAM, hDCPrinter), cx)
   IF m_Result THEN m_Result = 0 : RETURN TRUE
   IF m_Result = 0 THEN m_Result = E_FAIL: RETURN FALSE
END FUNCTION
' ========================================================================================


' ########################################################################################
' Implemented procedures and properties
' ########################################################################################

' // Properties

' ========================================================================================
' Gets a pointer to the application-defined AutoCorrectProc callback function.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.AutoCorrectProc () AS LONG_PTR
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETAUTOCORRECTPROC, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets a pointer to the application-defined AutoCorrectProc callback procedure.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.AutoCorrectProc (BYVAL pfn AS LONG_PTR)
   m_Result = SendMessageW(m_hRichEdit, EM_SETAUTOCORRECTPROC, 0, cast(LPARAM, pfn))
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetAutoCorrectProc () AS LONG_PTR
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETAUTOCORRECTPROC, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets a pointer to the application-defined AutoCorrectProc callback procedure.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetAutoCorrectProc (BYVAL pfn AS LONG_PTR) AS HRESULT
   m_Result = SendMessageW(m_hRichEdit, EM_SETAUTOCORRECTPROC, 0, cast(LPARAM, pfn))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether the auto URL detection is turned on in the rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.AutoUrlDetect () AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_GETAUTOURLDETECT, 0, 0))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Enables or disables automatic detection of URLs by a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.AutoUrlDetect (BYVAL fUrlDetect AS LONG)
   m_Result = SendMessageW(m_hRichEdit, EM_AUTOURLDETECT, fUrlDetect, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Indicates whether the auto URL detection is turned on in the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetAutoUrlDetect () AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_GETAUTOURLDETECT, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Enables or disables automatic detection of URLs by a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetAutoUrlDetect (BYVAL fUrlDetect AS LONG) AS HRESULT
   m_Result = SendMessageW(m_hRichEdit, EM_AUTOURLDETECT, fUrlDetect, 0)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Disables automatic detection of URLs by a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.DisableAutoUrlDetect () AS BOOLEAN
   m_Result = SendMessageW(m_hRichEdit, EM_AUTOURLDETECT, 0, 0)
   IF m_Result THEN RETURN FALSE ELSE RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
' Enables automatic detection of URLs by a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.EnableAutoUrlDetect (BYVAL fUrlDetect AS LONG) AS HRESULT
   m_Result = SendMessageW(m_hRichEdit, EM_AUTOURLDETECT, fUrlDetect, 0)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates the current state of the bidirectional options in the rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.BidiOptions () AS .BIDIOPTIONS
   DIM bo AS .BIDIOPTIONS
   bo.cbSize = SIZEOF(.BIDIOPTIONS)
   m_Result = SendMessageW(m_hRichEdit, EM_GETBIDIOPTIONS, 0, cast(LPARAM, @bo))
   RETURN bo
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the current state of the bidirectional options in the rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.BidiOptions (BYREF opt AS .BIDIOPTIONS)
   IF opt.cbSize = 0 THEN opt.cbSize = SIZEOF(.BIDIOPTIONS)
   m_Result = SendMessageW(m_hRichEdit, EM_SETBIDIOPTIONS, 0, cast(LPARAM, @opt))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Indicates the current state of the bidirectional options in the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetBidiOptions () AS .BIDIOPTIONS
   DIM bo AS .BIDIOPTIONS
   bo.cbSize = SIZEOF(.BIDIOPTIONS)
   m_Result = SendMessageW(m_hRichEdit, EM_GETBIDIOPTIONS, 0, cast(LPARAM, @bo))
   RETURN bo
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the current state of the bidirectional options in the rich edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetBidiOptions (BYREF opt AS .BIDIOPTIONS)
   IF opt.cbSize = 0 THEN opt.cbSize = SIZEOF(.BIDIOPTIONS)
   m_Result = SendMessageW(m_hRichEdit, EM_SETBIDIOPTIONS, 0, cast(LPARAM, @opt))
END SUB
' ========================================================================================

' ========================================================================================
' Gets the current character formatting in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CharFormat (BYVAL fOption AS DWORD) AS CHARFORMATW
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_GETCHARFORMAT, fOption, cast(LPARAM, @cf))
   RETURN cf
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets character formatting in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CharFormat (BYVAL chfmt AS DWORD, BYREF cf AS CHARFORMATW)
   IF cf.cbSize = 0 THEN cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_SETCHARFORMAT, chfmt, cast (LPARAM, @cf))
   IF m_Result = 0 THEN m_Result = E_FAIL
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the current character formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetCharFormat (BYVAL fOption AS DWORD) AS CHARFORMATW
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_GETCHARFORMAT, fOption, cast(LPARAM, @cf))
   RETURN cf
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets character formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetCharFormat (BYVAL chfmt AS DWORD, BYREF cf AS CHARFORMATW) AS BOOLEAN
   m_Result = 0
   IF cf.cbSize = 0 THEN cf.cbSize = SIZEOF (cf)
   DIM hr AS HREsuLT = SendMessageW(m_hRichEdit, EM_SETCHARFORMAT, chfmt, cast (LPARAM, @cf))
   IF hr = 0 THEN RETURN FALSE ELSE RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the default character formatting in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.DefaultCharFormat () AS CHARFORMATW
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_GETCHARFORMAT, SCF_DEFAULT, cast(LPARAM, @cf))
   RETURN cf
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the default character formatting in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.DefaultCharFormat (BYREF cf AS CHARFORMATW)
   IF cf.cbSize = 0 THEN cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_SETCHARFORMAT, SCF_DEFAULT, cast (LPARAM, @cf))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the default character formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetDefaultCharFormat () AS CHARFORMATW
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_GETCHARFORMAT, SCF_DEFAULT, cast(LPARAM, @cf))
   RETURN cf
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the default character formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetDefaultCharFormat (BYREF cf AS CHARFORMATW) AS BOOLEAN
   IF cf.cbSize = 0 THEN cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_SETCHARFORMAT, SCF_DEFAULT, cast (LPARAM, @cf))
   IF m_Result = 0 THEN RETURN FALSE ELSE RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the selection character formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetSelectionCharFormat () AS CHARFORMATW
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_GETCHARFORMAT, SCF_SELECTION, cast(LPARAM, @cf))
   RETURN cf
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the selection character formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetSelectionCharFormat (BYREF cf AS CHARFORMATW) AS BOOLEAN
   IF cf.cbSize = 0 THEN cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_SETCHARFORMAT, SCF_SELECTION, cast (LPARAM, @cf))
   IF m_Result = 0 THEN RETURN FALSE ELSE RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the word character formatting in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.WordCharFormat () AS CHARFORMATW
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_GETCHARFORMAT, SCF_SELECTION OR SCF_WORD, cast(LPARAM, @cf))
   RETURN cf
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the word character formatting in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.WordCharFormat (BYREF cf AS CHARFORMATW)
   IF cf.cbSize = 0 THEN cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_SETCHARFORMAT, SCF_SELECTION OR SCF_WORD, cast (LPARAM, @cf))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the word character formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetWordCharFormat () AS CHARFORMATW
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_GETCHARFORMAT, SCF_SELECTION OR SCF_WORD, cast(LPARAM, @cf))
   RETURN cf
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the word character formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetWordCharFormat (BYREF cf AS CHARFORMATW) AS BOOLEAN
   IF cf.cbSize = 0 THEN cf.cbSize = SIZEOF (cf)
   m_Result = SendMessageW(m_hRichEdit, EM_SETCHARFORMAT, SCF_SELECTION OR SCF_WORD, cast (LPARAM, @cf))
   IF m_Result = 0 THEN RETURN FALSE ELSE RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the Text Services Framework mode bias values for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CTFModeBias () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETCTFMODEBIAS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the Text Services Framework (TSF) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CTFModeBias (BYVAL nModeBias AS LONG)
   m_Result = SendMessageW(m_hRichEdit, EM_SETCTFMODEBIAS, nModeBias, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the Text Services Framework mode bias values for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetCTFModeBias () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETCTFMODEBIAS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the Text Services Framework (TSF) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetCTFModeBias (BYVAL nModeBias AS LONG) AS HRESULT
   m_Result = SendMessageW(m_hRichEdit, EM_SETCTFMODEBIAS, nModeBias, 0)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if the Text Services Framework (TSF) keyboard is open or closed.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CTFOpenStatus () AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_GETCTFOPENSTATUS, 0, 0))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Opens or closes the Text Services Framework (TSF) keyboard.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CTFOpenStatus (BYVAL fTSFkbd AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETCTFOPENSTATUS, fTSFkbd, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Determines if the Text Services Framework (TSF) keyboard is open or closed.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetCTFOpenStatus () AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_GETCTFOPENSTATUS, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Opens or closes the Text Services Framework (TSF) keyboard.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetCTFOpenStatus (BYVAL fTSFkbd AS LONG) AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_SETCTFOPENSTATUS, fTSFkbd, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current edit style flags.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EditStyle () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETEDITSTYLE, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the current edit style flags.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EditStyle (BYVAL fStyle AS LONG, BYVAL fMask AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETEDITSTYLE, fStyle, fMask)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the current edit style flags.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetEditStyle () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETEDITSTYLE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the current edit style flags.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetEditStyle (BYVAL fStyle AS LONG, BYVAL fMask AS LONG) AS HRESULT
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETEDITSTYLE, fStyle, fMask)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the extended edit style flags.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EditStyleEx () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETEDITSTYLEEX, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the current extended edit style flags.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EditStyleEx (BYVAL fStyle AS LONG, BYVAL fMask AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETEDITSTYLEEX, fStyle, fMask)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the extended edit style flags.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetEditStyleEx () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETEDITSTYLEEX, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the current extended edit style flags.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetEditStyleEx (BYVAL fStyle AS LONG, BYVAL fMask AS LONG) AS HRESULT
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETEDITSTYLEEX, fStyle, fMask)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the current ellipsis mode.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EllipsisMode () AS DWORD
   DIM pmode AS DWORD
   m_Result = SendMessageW(m_hRichEdit, EM_GETELLIPSISMODE, 0, cast(LPARAM, @pMode))
   RETURN pmode
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the current ellipsis mode. 
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EllipsisMode (BYVAL fMode AS DWORD)
   m_Result = SendMessageW(m_hRichEdit, EM_SETELLIPSISMODE, 0, fMode)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Retrieves the current ellipsis mode.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetEllipsisMode () AS DWORD
   DIM pmode AS DWORD
   m_Result = SendMessageW(m_hRichEdit, EM_GETELLIPSISMODE, 0, cast(LPARAM, @pMode))
   RETURN pmode
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the current ellipsis mode. 
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetEllipsisMode (BYVAL fMode AS DWORD) AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_SETELLIPSISMODE, 0, fMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current ellipsis state.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EllipsisState () AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_GETELLIPSISSTATE, 0, 0))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the current ellipsis state.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetEllipsisState () AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_GETELLIPSISSTATE, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the event mask for a rich edit control. The event mask specifies which
' notification messages the control sends to its parent window.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EventMask () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETEVENTMASK, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the event mask for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EventMask (BYVAL fMask AS DWORD)
   m_Result = SendMessageW(m_hRichEdit, EM_SETEVENTMASK, 0, fMask)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the event mask for a rich edit control. The event mask specifies which
' notification messages the control sends to its parent window.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetEventMask () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETEVENTMASK, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the event mask for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetEventMask (BYVAL fMask AS DWORD) AS HRESULT
   m_Result = SendMessageW(m_hRichEdit, EM_SETEVENTMASK, 0, fMask)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about hyphenation for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.HyphenateInfo () AS .HYPHENATEINFO
   DIM info AS .HYPHENATEINFO
   info.cbSize = SIZEOF(info)
   m_Result = SendMessageW(m_hRichEdit, EM_GETHYPHENATEINFO, cast(WPARAM, @info), 0)
   RETURN info
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the way a Microsoft Rich Edit control does hyphenation.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.HyphenateInfo (BYREF info AS .HYPHENATEINFO)
   IF info.cbSize = 0 THEN info.cbSize = SIZEOF(info)
   m_Result = SendMessageW(m_hRichEdit, EM_SETHYPHENATEINFO, cast(WPARAM, @info), 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets information about hyphenation for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetHyphenateInfo () AS .HYPHENATEINFO
   DIM info AS .HYPHENATEINFO
   info.cbSize = SIZEOF(info)
   m_Result = SendMessageW(m_hRichEdit, EM_GETHYPHENATEINFO, cast(WPARAM, @info), 0)
   RETURN info
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the way a Microsoft Rich Edit control does hyphenation.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetHyphenateInfo (BYREF info AS .HYPHENATEINFO) AS HRESULT
   IF info.cbSize = 0 THEN info.cbSize = SIZEOF(info)
   m_Result = SendMessageW(m_hRichEdit, EM_SETHYPHENATEINFO, cast(WPARAM, @info), 0)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.IMEModeBias () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETIMEMODEBIAS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.IMEModeBias (BYVAL nModeBias AS LONG)
   m_Result = SendMessageW(m_hRichEdit,EM_SETIMEMODEBIAS, nModeBias, nModeBias)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetIMEModeBias () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETIMEMODEBIAS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetIMEModeBias (BYVAL nModeBias AS LONG) AS HRESULT
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit,EM_SETIMEMODEBIAS, nModeBias, nModeBias)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current Input Method Editor (IME) options. This message is available only
' in Asian-language versions of the operating system.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.IMEOptions () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETIMEOPTIONS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the Input Method Editor (IME) options.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.IMEOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETIMEOPTIONS, fCoop, fOptions)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the current Input Method Editor (IME) options. This message is available only
' in Asian-language versions of the operating system.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetIMEOptions () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETIMEOPTIONS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the Input Method Editor (IME) options.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetIMEOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG) AS HRESULT
   m_Result = 0
   DIM hr AS HRESULT = SendMessageW(m_hRichEdit, EM_SETIMEOPTIONS, fCoop, fOptions)
   IF hr = 0 THEN RETURN FALSE ELSE RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a rich edit control's option settings for Input Method Editor (IME) and Asian
' language support.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.LangOptions () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETLANGOPTIONS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets options for Input Method Editor (IME) and Asian language support in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.LangOptions (BYVAL lgoptions AS LONG) 
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETLANGOPTIONS, 0, lgoptions)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets a rich edit control's option settings for Input Method Editor (IME) and Asian
' language support.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetLangOptions () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETLANGOPTIONS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets options for Input Method Editor (IME) and Asian language support in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetLangOptions (BYVAL lgoptions AS LONG) AS HRESULT
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETLANGOPTIONS, 0, lgoptions)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current text limit for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.LimitText () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETLIMITTEXT, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the text limit of a rich edit control. The text limit is the maximum amount of
' text, in TCHARs, that the user can type into the edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.LimitText (BYVAL chMax AS DWORD)
   m_Result = SendMessageW(m_hRichEdit, EM_LIMITTEXT, chMax, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the current text limit for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetLimitText () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETLIMITTEXT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the text limit of a rich edit control. The text limit is the maximum amount of
' text, in TCHARs, that the user can type into the edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetLimitText (BYVAL chMax AS DWORD)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_LIMITTEXT, chMax, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the state of a rich edit control's modification flag. The flag indicates whether
' the contents of the rich edit control have been modified.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Modify () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETMODIFY, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets or clears the modification flag for a rich edit control. The modification flag
' indicates whether the text within the rich edit control has been modified.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Modify (BYVAL fModify AS LONG)
   m_Result = SendMessageW(m_hRichEdit, EM_SETMODIFY, fModify, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the state of a rich edit control's modification flag. The flag indicates whether
' the contents of the rich edit control have been modified.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetModify () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETMODIFY, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets or clears the modification flag for a rich edit control. The modification flag
' indicates whether the text within the rich edit control has been modified.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetModify (BYVAL fModify AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETMODIFY, fModify, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the options for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Options () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETOPTIONS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the options for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Options (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETOPTIONS, fCoop, fOptions)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Retrieves the options for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetOptions () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETOPTIONS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the options for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETOPTIONS, fCoop, fOptions)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deprecated. Gets the text layout for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.PageRotate () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETPAGEROTATE, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Deprecated. Sets the text layout for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.PageRotate (BYVAL txtlayout AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETPAGEROTATE, txtlayout, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Deprecated. Gets the text layout for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetPageRotate () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETPAGEROTATE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Deprecated. Sets the text layout for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetPageRotate (BYVAL txtlayout AS LONG) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETPAGEROTATE, txtlayout, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the paragraph formatting of the current selection in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.ParaFormat () AS .PARAFORMAT
   m_Result = 0
   DIM pfmt AS .PARAFORMAT
   pfmt.cbSize = SIZEOF(.PARAFORMAT)
   SendMessageW(m_hRichEdit, EM_GETPARAFORMAT, 0, cast(LPARAM, @pfmt))
   RETURN pfmt
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the paragraph formatting for the current selection in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.ParaFormat (BYREF pfmt AS .PARAFORMAT)
   IF pfmt.cbSize = 0 THEN pfmt.cbSize = SIZEOF(.PARAFORMAT)
   m_Result = SendMessageW(m_hRichEdit, EM_SETPARAFORMAT, 0, cast(LPARAM, @pfmt))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the paragraph formatting of the current selection in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetParaFormat () AS .PARAFORMAT
   m_Result = 0
   DIM pfmt AS .PARAFORMAT
   pfmt.cbSize = SIZEOF(.PARAFORMAT)
   SendMessageW(m_hRichEdit, EM_GETPARAFORMAT, 0, cast(LPARAM, @pfmt))
   RETURN pfmt
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the paragraph formatting for the current selection in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetParaFormat (BYREF pfmt AS .PARAFORMAT) AS BOOLEAN
   IF pfmt.cbSize = 0 THEN pfmt.cbSize = SIZEOF(.PARAFORMAT)
   m_Result = SendMessageW(m_hRichEdit, EM_SETPARAFORMAT, 0, cast(LPARAM, @pfmt))
   RETURN CBOOL(m_Result)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the password character that a rich edit control displays when the user enters text.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.PasswordChar () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETPASSWORDCHAR, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets or removes the password character for a rich edit control. When a password
' character is set, that character is displayed in place of the characters typed by the user.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.PasswordChar (BYVAL dwchar AS DWORD)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETPASSWORDCHAR, dwchar, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the password character that a rich edit control displays when the user enters text.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetPasswordChar () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETPASSWORDCHAR, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets or removes the password character for a rich edit control. When a password
' character is set, that character is displayed in place of the characters typed by the user.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetPasswordChar (BYVAL dwchar AS DWORD)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETPASSWORDCHAR, dwchar, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the current punctuation characters for the rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Punctuation (BYVAL punctype AS DWORD) AS .PUNCTUATION
   m_Result = 0
   DIM punct AS .PUNCTUATION
   SendMessageW(m_hRichEdit, EM_GETPUNCTUATION, punctype, cast(LPARAM, @punct))
   RETURN punct
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the punctuation characters for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Punctuation (BYVAL punctype AS LONG, BYREF punct AS .PUNCTUATION)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETPUNCTUATION, punctype, cast(LPARAM, @punct))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the current punctuation characters for the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetPunctuation (BYVAL punctype AS DWORD) AS .PUNCTUATION
   m_Result = 0
   DIM punct AS .PUNCTUATION
   SendMessageW(m_hRichEdit, EM_GETPUNCTUATION, punctype, cast(LPARAM, @punct))
   RETURN punct
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the punctuation characters for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetPunctuation (BYVAL punctype AS LONG, BYREF punct AS .PUNCTUATION) AS BOOLEAN
   m_Result = 0
   DIM hr As HRESULT = SendMessageW(m_hRichEdit, EM_SETPUNCTUATION, punctype, cast(LPARAM, @punct))
   IF hr = 0 THEN RETURN FALSE ELSE RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the formatting rectangle of a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Rect () AS .RECT
   DIM rc AS .RECT
   m_Result = SendMessageW(m_hRichEdit, EM_GETRECT, 0, cast(LPARAM, @rc))
   RETURN rc
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the formatting rectangle of a multiline rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Rect (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
   m_Result = SendMessageW(m_hRichEdit, EM_SETRECT, fCoord, cast(LPARAM, @rc))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the formatting rectangle of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetRect () AS .RECT
   DIM rc AS .RECT
   m_Result = SendMessageW(m_hRichEdit, EM_GETRECT, 0, cast(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the formatting rectangle of a multiline rich edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetRect (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETRECT, fCoord, cast(LPARAM, @rc))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the formatting rectangle of a multiline rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.RectNP (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
   m_Result = 0
   SendMessageW( m_hRichEdit, EM_SETRECTNP, fCoord, cast(LPARAM, @rc))
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetRectNP (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
   m_Result = 0
   SendMessageW( m_hRichEdit, EM_SETRECTNP, fCoord, cast(LPARAM, @rc))
END SUB
' ========================================================================================

' ========================================================================================
' Obtains the current scroll position of the edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.ScrollPos () AS .POINT
   DIM pt AS .POINT
   m_Result = SendMessageW(m_hRichEdit, EM_GETSCROLLPOS, 0, cast(LPARAM, @pt))
   RETURN pt
END PROPERTY
' ========================================================================================
' ========================================================================================
' Tells the rich edit control to scroll to a particular point.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.ScrollPos (BYREF pt AS .POINT)
   m_Result = SendMessageW(m_hRichEdit, EM_SETSCROLLPOS, 0, cast(LPARAM, @pt))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Obtains the current scroll position of the edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetScrollPos () AS .POINT
   DIM pt AS .POINT
   m_Result = SendMessageW(m_hRichEdit, EM_GETSCROLLPOS, 0, cast(LPARAM, @pt))
   RETURN pt
END FUNCTION
' ========================================================================================
' ========================================================================================
' Tells the rich edit control to scroll to a particular point.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetScrollPos (BYREF pt AS .POINT) AS HRESULT
   m_Result = SendMessageW(m_hRichEdit, EM_SETSCROLLPOS, 0, cast(LPARAM, @pt))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the story type.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.StoryType (BYVAL Index AS DWORD) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETSTORYTYPE, Index, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the story type.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.StoryType (BYVAL Index AS LONG, BYVAL dwType AS DWORD)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETSTORYTYPE, Index, dwType)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the story type.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetStoryType (BYVAL Index AS DWORD) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETSTORYTYPE, Index, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the story type.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetStoryType (BYVAL Index AS LONG, BYVAL dwType AS DWORD) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETSTORYTYPE, Index, dwType)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the text from a rich edit control.
' Note: GetWindowText cannot retrieve the text of a control in another application.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Text () AS CWSTR
   m_Result = 0
   DIM cbLen AS DWORD = SendMessageW(m_hRichEdit, WM_GETTEXTLENGTH, 0, 0)
   IF cbLen < 1 THEN RETURN ""
   DIM cwsText AS CWSTR = WSPACE(cbLen + 1)
   cbLen = SendMessageW(m_hRichEdit, WM_GETTEXT, cbLen + 1, cast(LPARAM, *cwsText))
   RETURN LEFT(**cwsText, cbLen)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the text of an edit control.
' Note: SetWindowText cannot change the text of a control in another application.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Text (BYREF wszText AS WSTRING)
   m_Result = SendMessageW(m_hRichEdit, WM_SETTEXT, 0, cast(LPARAM, @wszText))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the text from a rich edit control.
' Note: GetWindowText cannot retrieve the text of a control in another application.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetText () AS CWSTR
   m_Result = 0
   DIM cbLen AS DWORD = SendMessageW(m_hRichEdit, WM_GETTEXTLENGTH, 0, 0)
   IF cbLen < 1 THEN RETURN ""
   DIM cwsText AS CWSTR = WSPACE(cbLen + 1)
   cbLen = SendMessageW(m_hRichEdit, WM_GETTEXT, cbLen + 1, cast(LPARAM, *cwsText))
   RETURN LEFT(**cwsText, cbLen)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the text of an edit control.
' Note: SetWindowText cannot change the text of a control in another application.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetText (BYREF wszText AS WSTRING) AS BOOLEAN
   m_Result = SendMessageW(m_hRichEdit, WM_SETTEXT, 0, cast(LPARAM, @wszText))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current text mode and undo level of a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextMode () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTEXTMODE, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the text mode or undo level of a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextMode (BYVAL values AS LONG)
   m_Result = SendMessageW(m_hRichEdit, EM_SETTEXTMODE, 0, values)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the current text mode and undo level of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextMode () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTEXTMODE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the text mode or undo level of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTextMode (BYVAL values AS LONG) AS BOOLEAN
   m_Result = 0
   DIM hr AS HRESULT = SendMessageW(m_hRichEdit, EM_SETTEXTMODE, 0, values)
   IF hr = 0 THEN RETURN TRUE ELSE RETURN FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the touch options that are associated with a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TouchOptions (BYVAL opt AS LONG PTR) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTOUCHOPTIONS, cast(WPARAM, opt), 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the touch options associated with a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TouchOptions (BYVAL opt AS LONG, BYVAL fEnable AS LONG)
   m_Result = SendMessageW(m_hRichEdit, EM_SETTOUCHOPTIONS, opt, fEnable)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the touch options that are associated with a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTouchOptions (BYVAL opt AS LONG PTR) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTOUCHOPTIONS, cast(WPARAM, opt), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the touch options associated with a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTouchOptions (BYVAL opt AS LONG, BYVAL fEnable AS LONG) AS HRESULT
   m_Result = SendMessageW(m_hRichEdit, EM_SETTOUCHOPTIONS, opt, fEnable)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current state of the typography options of a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TypographyOptions () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTYPOGRAPHYOPTIONS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the current state of the typography options of a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TypographyOptions (BYVAL pto AS LONG, BYVAL fMask AS LONG)
   m_Result = 0
   m_Result = SendMessageW(m_hRichEdit, EM_SETTYPOGRAPHYOPTIONS, pto, fMask)
   IF m_Result = FALSE THEN m_Result = 1
END PROPERTY
' ========================================================================================
' ========================================================================================
' Returns the current state of the typography options of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTypographyOptions () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTYPOGRAPHYOPTIONS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the current state of the typography options of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTypographyOptions (BYVAL pto AS LONG, BYVAL fMask AS LONG) AS BOOLEAN
   m_Result = SendMessageW(m_hRichEdit, EM_SETTYPOGRAPHYOPTIONS, pto, fMask)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' // Procedures

' ========================================================================================
' Determines whether a rich edit control can paste a specified clipboard format.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.CanPaste (BYVAL clipformat AS LONG) AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_CANPASTE, clipformat, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether there are any actions in the control redo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.CanRedo () AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_CANREDO, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether there are any actions in an edit control's undo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.CanUndo () AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_CANUNDO, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calls the autocorrect callback function that is stored by the **CRichEditCtrl.SetAutocorrectProc**
' message, provided that the text preceding the insertion point is a candidate for autocorrection.
' ========================================================================================
FUNCTION CRichEditCtrl.CallAutocorrectProc (BYVAL char AS WCHAR) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_CALLAUTOCORRECTPROC, char, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Displays a portion of the contents of a rich edit control, as previously formatted for a
' device using the EM_FORMATRANGE message.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.DisplayBand (BYREF rc AS .RECT) AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_DISPLAYBAND, 0, cast(LPARAM, @rc)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets the undo flag of a rich edit control. The undo flag is set whenever an operation
' within the rich edit control can be undone.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.EmptyUndoBuffer ()
   m_Result = SendMessageW(m_hRichEdit, EM_EMPTYUNDOBUFFER, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the starting and ending character positions of the selection in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.ExGetSel () AS CHARRANGE
   DIM cr AS CHARRANGE
   m_Result = SendMessageW(m_hRichEdit, EM_EXGETSEL, 0, cast(LPARAM, @cr))
   RETURN cr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ExGetSel (BYREF cpMin AS LONG, BYVAL cpMax AS LONG)
   DIM cr AS CHARRANGE = TYPE<CHARRANGE>(cpMin, cpMax)
   m_Result = SendMessageW(m_hRichEdit, EM_EXGETSEL, 0, cast(LPARAM, @cr))
   IF VARPTR(cpMin) THEN cpMin = cr.cpMin
   IF VARPTR(cpMax) THEN cpMax = cr.cpMax
END SUB
' ========================================================================================

' ========================================================================================
' Sets an upper limit to the amount of text the user can type or paste into a rich edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ExLimitText (BYVAL dwLimit AS DWORD)
   m_Result = SendMessageW(m_hRichEdit, EM_EXLIMITTEXT, 0, dwLimit)
END SUB
' ========================================================================================

' ========================================================================================
' Determines which line contains the specified character in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.ExLineFromChar (BYVAL iIndex AS LONG) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_EXLINEFROMCHAR, 0, iIndex)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects a range of characters and/or Component Object Model (COM) objects in a Microsoft
' Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.ExSetSel (BYREF cr AS CHARRANGE) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_EXSETSEL, 0, cast(LPARAM, @cr))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ExSetSel (BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1)
   DIM cr AS CHARRANGE = TYPE<CHARRANGE>(cpMin, cpMax)
   m_Result = SendMessageW(m_hRichEdit, EM_EXSETSEL, 0, cast(LPARAM, @cr))
END SUB
' ========================================================================================

' ========================================================================================
' Finds text within a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FindText (BYVAL fOptions AS DWORD, BYREF ftw AS FINDTEXTW) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_FINDTEXTW, fOptions, cast(LPARAM, @ftw))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FindText (BYVAL fOptions AS DWORD = FR_DOWN, _
   BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1, BYREF wszText AS WSTRING) AS LONG
   m_Result = 0
   DIM ftw AS FINDTEXTW
   ftw.chrg.cpMin = cpMin
   ftw.chrg.cpMax = cpMax
   ftw.lpstrText = @wszText
   RETURN SendMessageW(m_hRichEdit, EM_FINDTEXTW, fOptions, cast(LPARAM, @ftw))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds text within a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FindTextEx (BYVAL fOptions AS DWORD, BYREF ftexw AS FINDTEXTEXW) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_FINDTEXTEXW, fOptions, cast(LPARAM, @ftexw))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FindTextEx (BYVAL fOptions AS DWORD = FR_DOWN, _
   BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1, BYREF wszText AS WSTRING) AS CHARRANGE
   DIM ftexw AS FINDTEXTEXW
   ftexw.chrg.cpMin = cpMin
   ftexw.chrg.cpMax = cpMax
   ftexw.lpstrText = @wszText
   m_Result = SendMessageW(m_hRichEdit, EM_FINDTEXTEXW, fOptions, cast(LPARAM, @ftexw))
   RETURN ftexw.chrgText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds the next word break before or after the specified character position or retrieves
' information about the character at that position.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FindWordBreak (BYVAL fOperation AS DWORD, BYVAL dwStartPos AS DWORD) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_FINDWORDBREAK, fOperation, dwStartPos)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the zero-based index of the uppermost visible line in a multiline rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.FirstVisibleLine () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETFIRSTVISIBLELINE, 0, 0)
END PROPERTY
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetFirstVisibleLine () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETFIRSTVISIBLELINE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Formats a range of text in a rich edit control for a specific device.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FormatRange (BYVAL fRender AS LONG, BYREF fr AS .FORMATRANGE) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_FORMATRANGE, fRender, cast(LPARAM, @fr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about the character closest to a specified point in the client area of
' an edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CharFromPos (BYREF ptl AS .POINTL) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_CHARFROMPOS, 0, cast(LPARAM, @ptl))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets information about the character closest to a specified point in the client area of
' an edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetCharFromPos (BYREF ptl AS .POINTL) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_CHARFROMPOS, 0, cast(LPARAM, @ptl))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current IME mode for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.IMECompMode () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETIMECOMPMODE, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetIMECompMode () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETIMECOMPMODE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the Input Method Editor (IME) composition color. This message is available
' only in Asian-language versions of the operating system.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetIMEColor (BYVAL rgCmpclr AS .COMPCOLOR PTR) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETIMECOLOR, 0, cast(LPARAM, rgCmpclr))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the Input Method Editor (IME) composition color.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetIMEColor (BYVAL pcompcolor AS .COMPCOLOR PTR) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETIMECOLOR, 0, cast(LPARAM, pcompcolor))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the property and capabilities of the Input Method Editor (IME) associated with the
' current input locale.
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.IMEProperty (BYVAL figp AS DWORD) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETIMEPROPERTY, figp, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetIMEProperty (BYVAL figp AS DWORD) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETIMEPROPERTY, figp, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies a line of text from a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetLine (BYVAL which AS DWORD) AS CWSTR
   m_Result = 0
   DIM buffer AS CWSTR = MKI(32765) + STRING(32765, 0)
   DIM n AS LONG = SendMessageW(m_hRichEdit, EM_GETLINE, which, cast(LPARAM, *buffer))
   RETURN LEFT(**buffer, n)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves an IRichEditOle object that a client can use to access a rich edit control's
' Component Object Model (COM) functionality.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetOleInterface () AS IRichEditOle PTR
   DIM ppObject AS IRichEditOle PTR
   m_Result = SendMessageW(m_hRichEdit, EM_GETOLEINTERFACE, 0, cast(LPARAM, @ppObject))
   RETURN ppObject
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the type of the next action, if any, in the control's redo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetRedoName () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETREDONAME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the starting and ending character positions of the current selection in a rich
' edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetSel (BYREF dwStartPos AS DWORD, BYREF dwEndPos AS DWORD) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETSEL, cast(WPARAM, @dwStartPos), cast(LPARAM, @dwEndPos))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetSel () AS .POINTL
   m_Result = 0
   DIM dwStartPos AS DWORD, dwEndPos AS DWORD, pt AS .POINTL
   SendMessageW(m_hRichEdit, EM_GETSEL, cast(WPARAM, @dwStartPos), cast(LPARAM, @dwEndPos))
   pt.x = dwStartPos
   pt.y = dwEndPos
   RETURN pt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects a range of characters in a rich edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetSel (BYVAL nStart AS LONG, BYVAL nEnd AS LONG)
   m_Result = SendMessageW(m_hRichEdit, EM_SETSEL, nStart, nEnd)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetSel (BYREF pt AS POINTL)
   m_Result = SendMessageW(m_hRichEdit, EM_SETSEL, pt.x, pt.y)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the caret position
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CaretPos () AS DWORD
   m_Result = 0
   DIM dwStartPos AS DWORD, dwEndPos AS DWORD
   SendMessageW(m_hRichEdit, EM_GETSEL, cast(WPARAM, @dwStartPos), cast(LPARAM, @dwEndPos))
   RETURN dwStartPos
END PROPERTY   
' ========================================================================================
' ========================================================================================
' Sets the caret position
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CaretPos (BYVAL dwPos AS DWORD)
   m_Result = 0
   IF CLNG(dwPos) = -1 THEN
      this.SetSel(0, -1)
      DIM dwStartPos AS DWORD, dwEndPos AS DWORD
      this.GetSel(dwStartPos, dwEndPos)
      this.SetSel(dwEndPos + 1, dwEndPos + 1)
      this.ReplaceSel(TRUE, BYVAL NULL)
   ELSE
      SendMessageW(m_hRichEdit, EM_SETSEL, cast(WPARAM, dwPos), cast(LPARAM, dwPos))
   END IF
END PROPERTY   
' ========================================================================================
' ========================================================================================
' Gets the caret position
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetCaretPos () AS DWORD
   m_Result = 0
   DIM dwStartPos AS DWORD, dwEndPos AS DWORD
   SendMessageW(m_hRichEdit, EM_GETSEL, cast(WPARAM, @dwStartPos), cast(LPARAM, @dwEndPos))
   RETURN dwStartPos
END FUNCTION   
' ========================================================================================
' ========================================================================================
' Sets the caret position
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetCaretPos (BYVAL dwPos AS DWORD)
   m_Result = 0
   IF CLNG(dwPos) = -1 THEN
      this.SetSel(0, -1)
      DIM dwStartPos AS DWORD, dwEndPos AS DWORD
      this.GetSel(dwStartPos, dwEndPos)
      this.SetSel(dwEndPos + 1, dwEndPos + 1)
      this.ReplaceSel(TRUE, BYVAL NULL)
   ELSE
      SendMessageW(m_hRichEdit, EM_SETSEL, cast(WPARAM, dwPos), cast(LPARAM, dwPos))
   END IF
END SUB   
' ========================================================================================

' ========================================================================================
' Retrieves the currently selected text in a rich edit control.
' Note: From version 4.1 (RichEdit50W class in MSFTEDIT.DLL), Windows XP, EM_GETSELTEXT
' returns the text in unicode.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetSelText () AS CWSTR
   m_Result = 0
   DIM dwStartPos AS DWORD, dwEndPos AS DWORD, cr AS CHARRANGE
   SendMessageW(m_hRichEdit, EM_EXGETSEL, 0, cast(LPARAM, @cr))
   DIM cbLen AS DWORD = cr.cpMax - cr.cpMin
   IF cbLen < 1 THEN RETURN ""
   DIM cwsText AS CWSTR = WSPACE(cbLen + 1)
   cbLen = SendMessageW(m_hRichEdit, EM_GETSELTEXT, 0, cast(LPARAM, *cwsText))
   RETURN LEFT(**cwsText, cbLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the type of the next undo action, if any.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetUndoName () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETUNDONAME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current word wrap and word-break options for the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetWordWrapMode () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETWORDWRAPMODE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the word-wrapping and word-breaking options for the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetWordWrapMode (BYVAL pvalues AS LONG) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETWORDWRAPMODE, pvalues, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current zoom ratio, which is always between 1/64 and 64.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetZoom (BYREF pzNum AS DWORD, BYREF pzDen AS DWORD) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETZOOM, cast(WPARAM, @pzNum), cast(LPARAM, @pzDen))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the zoom ratio anywhere between 1/64 and 64.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetZoom (BYVAL zNum AS DWORD, BYVAL zDen AS DWORD) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETZOOM, zNum, zDen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the table parameters for a table row and the cell parameters for the specified number of cells.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTABLEPARMS, cast(WPARAM, @tp), cast(LPARAM, @tcp))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Changes the parameters of rows in a table.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETTABLEPARMS, cast(WPARAM, @tp), cast(LPARAM, @tcp))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets all of the text from the rich edit control in any particular code base you want.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextEx (BYREF gtex AS .GETTEXTEX, BYVAL buffer AS ANY PTR) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTEXTEX, cast(WPARAM, @gtex), cast(LPARAM, buffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Combines the functionality of WM_SETTEXT and EM_REPLACESEL and adds the ability to set
' text using a code page and to use either Rich Text Format (RTF) rich text or plain text.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTextEx (BYREF stex AS .SETTEXTEX, BYVAL buffer AS ANY PTR) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETTEXTEX, cast(WPARAM, @stex), cast(LPARAM, buffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline
' rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetThumb () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTHUMB, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the address of the current Wordwrap function.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetWordBreakProc () AS LONG_PTR
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETWORDBREAKPROC, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Replaces a rich edit control's default Wordwrap function with an application-defined Wordwrap function.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetWordBreakProc (BYVAL pfn AS LONG_PTR)
   m_Result = SendMessageW(m_hRichEdit, EM_SETWORDBREAKPROC, 0, cast(LPARAM, pfn))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the address of the currently registered extended word-break procedure.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetWordBreakProcEx () AS LONG_PTR
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETWORDBREAKPROCEX, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the extended word-break procedure.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetWordBreakProcEx (BYVAL pfn AS LONG_PTR) AS LONG_PTR
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETWORDBREAKPROCEX, 0, cast(LPARAM, pfn))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Hides or shows the selection in a rich edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.HideSelection (BYVAL fHide AS DWORD)
   m_Result = SendMessageW(m_hRichEdit, EM_HIDESELECTION, fHide, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the Input Method Editor (IME) composition text.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetIMECompText (BYREF ict AS .IMECOMPTEXT, BYVAL buffer AS ANY PTR) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETIMECOMPTEXT, cast(WPARAM, @ict), cast(LPARAM, buffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces the selection with a blob that displays an image.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.InsertImage (BYREF ip AS RICHEDIT_IMAGE_PARAMETERS) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_INSERTIMAGE, 0, cast(LPARAM, @ip))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts an image from file
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.InsertImage (BYREF wszFileName AS WSTRING, BYVAL xWidth AS LONG = 0, BYVAL yHeight AS LONG = 0, _
BYVAL Ascent AS LONG = 0, BYVAL nType AS LONG = TA_BASELINE, BYREF wszAlternateText AS WSTRING = "") AS BOOLEAN
   m_Result = 0
   ' // Retrieve the size of the image
   IF xWidth = 0 AND yHeight = 0 THEN
      DIM dwWidth AS DWORD, dwHeight AS DWORD
      DIM nStatus AS LONG = AfxGdipGetImageSizeFromFile(wszFileName, @dwWidth, @dwHeight)
      IF nStatus = 0 THEN xWidth = dwWidth : yHeight = dwHeight
   END IF
   ' // Opens the file and creates an stream
   ' // See if the library is already loaded in the address space
   DIM AS ANY PTR pLib = GetModuleHandleW("shlwapi.dll")
   ' // If it is not loaded, load it
   IF pLib = NULL THEN pLib = DyLibLoad("shlwapi.dll")
   IF pLib = NULL THEN this.m_Result = E_FAIL : RETURN FALSE
   DIM pShCreateStreamOnFileEx AS FUNCTION (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD, BYVAL dwAttributes AS DWORD, _
       BYVAL fCreate AS WINBOOL, BYVAL pstmTemplate AS IStream PTR, BYVAL ppStream AS IStream PTR PTR) AS HRESULT
   pShCreateStreamOnFileEx = DyLibSymbol(pLib, "SHCreateStreamOnFileEx")
   IF pShCreateStreamOnFileEx = NULL THEN m_Result = E_FAIL : RETURN FALSE
   DIM pStream AS IStream PTR
   DIM hr AS HRESULT = pShCreateStreamOnFileEx(@wszFileName, STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, @pStream)
   ' For some reason, if we free the library, calling the methods that use the returned
   ' IStream pointer will GPF if compiled with the -O switch.
'   DyLibFree(pLib)
   m_Result = hr
   IF hr <> S_OK OR pStream = NULL THEN RETURN FALSE
   ' // Fill the RICHEDIT_IMAGE_PARAMETERS structure
   DIM ip AS RICHEDIT_IMAGE_PARAMETERS
   ip.xWidth = AfxPixelsToHiMetricX(xWidth * m_ScalingRatio)
   ip.yHeight = AfxPixelsToHiMetricY(yHeight * m_ScalingRatio)
   ip.Ascent = IIF(Ascent = 0, 0, AfxPixelsToHiMetricY(Ascent * m_ScalingRatio))
   ip.Type = nType
   ip.pwszAlternateText = @wszAlternateText
   ip.pIStream = pStream
   ' // Insert the image
   m_Result = SendMessageW(m_hRichEdit, EM_INSERTIMAGE, 0, cast(LPARAM, @ip))
   ' // Release the stream
   IF pStream THEN pStream->lpvtbl->Release(pStream)
   ' // Return TRUE for success or FALSE for failure
   IF m_Result <> S_OK THEN RETURN FALSE
   RETURN TRUE

END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts one or more identical table rows with empty cells.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.InsertTable (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_INSERTTABLE, cast(WPARAM, @tp), cast(LPARAM, @tcp))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if current input locale is an East Asian locale.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.IsIME () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_ISIME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of lines in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetLineCount () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETLINECOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the index of the line that contains the specified character index in a multiline
' rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LineFromChar (BYVAL index AS DWORD) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_LINEFROMCHAR, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character index of the first character of a specified line in a multiline
' rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LineIndex (BYVAL nLine AS LONG) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_LINEINDEX, nLine, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length, in characters, of a line in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LineLength (BYVAL index AS DWORD) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_LINELENGTH, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the text in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LineScroll (BYVAL y AS LONG) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_LINESCROLL, 0, y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Pastes a specific clipboard format in a rich edit control.
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.PasteSpecial (BYVAL clpfmt AS DWORD, BYREF rps AS REPASTESPECIAL)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_PASTESPECIAL, clpfmt, cast(LPARAM, @rps))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.PasteSpecial (BYVAL clpfmt AS DWORD, BYVAL dwAspect AS DWORD, BYVAL hMF AS HMETAFILE)
   m_Result = 0
   DIM rps AS REPASTESPECIAL= TYPE<REPASTESPECIAL>(dwAspect, cast(DWORD_PTR, hMF))
   SendMessageW(m_hRichEdit, EM_PASTESPECIAL, clpfmt, cast(LPARAM, @rps))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the client area coordinates of a specified character in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.PosFromChar (BYVAL index as DWORD) AS .POINTL
   DIM pt AS .POINTL
   m_Result = SendMessageW(m_hRichEdit, EM_POSFROMCHAR, cast(WPARAM, @pt), index)
   RETURN pt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Invokes the Input Method Editor (IME) reconversion dialog box.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.Reconversion
   m_Result = SendMessageW(m_hRichEdit, EM_RECONVERSION, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Redoes the next action in the control's redo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.Redo () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_REDO, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces the current selection in a rich edit control with the specified text.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ReplaceSel (BYVAL bCanBeUndone AS LONG = TRUE, BYREF wszText AS WSTRING)
   m_Result = SendMessageW(m_hRichEdit, EM_REPLACESEL, bCanBeUndone, cast(LPARAM, @wszText))
END SUB
' ========================================================================================

' ========================================================================================
' Forces a rich edit control to send an EN_REQUESTRESIZE notification message to its
' parent window.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.RequestResize ()
   m_Result = SendMessageW(m_hRichEdit, EM_REQUESTRESIZE, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Scrolls the text vertically in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.Scroll (BYVAL nAction AS LONG) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SCROLL, nAction, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the caret into view in a rich edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ScrollCaret ()
   m_Result = SendMessageW(m_hRichEdit, EM_SCROLLCARET, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Determines the selection type for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SelectionType () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SELECTIONTYPE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetBkgndColor (BYVAL SysColor AS DWORD, BYVAL BkColor AS DWORD) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETBKGNDCOLOR, SysColor, BkColor)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font size for the selected text.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetFontSize (BYVAL ptsize AS LONG) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETFONTSIZE, ptsize, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the widths of the left and right margins for a rich edit control. The message
' redraws the control to reflect the new margins.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetMargins (BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETMARGINS, nMargins, nWidth)
END SUB
' ========================================================================================
' ========================================================================================
' Sets the width of the left margin.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.LeftMargin (BYVAL nWidth AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETMARGINS, EC_LEFTMARGIN, MAKELONG(nWidth, 0))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the width of the left margin.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetLeftMargin (BYVAL nWidth AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETMARGINS, EC_LEFTMARGIN, MAKELONG(nWidth, 0))
END SUB
' ========================================================================================
' ========================================================================================
' Sets the width of the right margin.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.RightMargin (BYVAL nWidth AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETMARGINS, EC_RIGHTMARGIN, MAKELONG(0, nWidth))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the width of the right margin.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetRightMargin (BYVAL nWidth AS LONG)
   m_Result = 0
   SendMessageW(m_hRichEdit, EM_SETMARGINS, EC_RIGHTMARGIN, MAKELONG(0, nWidth))
END SUB
' ========================================================================================

' ========================================================================================
' Gives a rich edit control an IRichEditOleCallback object that the control uses to get
' OLE-related resources and information from the client.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetOleCallback (BYVAL pCallback AS ANY PTR) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETOLECALLBACK, 0, cast(LPARAM, pCallback))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the palette that a rich edit control uses for its display window.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetPalette (BYVAL newPalette AS HPALETTE)
   m_Result = SendMessageW(m_hRichEdit, EM_SETPALETTE, cast(WPARAM, newPalette), 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets or removes the read-only style (ES_READONLY) of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetReadOnly (BYVAL fReadOnly AS LONG) AS BOOLEAN
   m_Result = 0
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_SETREADONLY, fReadOnly, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the tab stops in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTabStops (BYVAL nTabs AS LONG, BYVAL rgTabStops AS LONG_PTR) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETTABSTOPS, nTabs, cast(LPARAM, rgTabStops))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the target device and line width used for WYSIWYG formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTargetDevice (BYVAL hDC AS HDC, BYVAL lnwidth AS LONG) AS BOOLEAN
   DIM hr AS HRESULT = SendMessageW(m_hRichEdit, EM_SETTARGETDEVICE, cast(WPARAM, hDC), lnwidth)
   ' // This method returns 0 for failure and <> 0 for success
   IF hr <> 0 THEN
      m_Result = S_OK
      RETURN TRUE
   ELSE
      m_Result = S_FALSE
      RETURN FALSE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the name of a rich edit control for UI Automation (UIA).
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetUIAName (BYREF wszName AS WSTRING) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETUIANAME, 0, cast(LPARAM, @wszName))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the maximum number of actions that can stored in the undo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetUndoLimit (BYVAL maxactions AS DWORD) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_SETUNDOLIMIT, maxactions, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Shows or hides one of the scroll bars in the Text Host window.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ShowScrollBar (BYVAL nScrollBar AS DWORD, BYVAL fShow AS LONG)
   m_Result = SendMessageW(m_hRichEdit, EM_SHOWSCROLLBAR, nScrollBar, fShow)
END SUB
' ========================================================================================

' ========================================================================================
' Stops the control from collecting additional typing actions into the current undo action.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.StopGroupTyping () AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_STOPGROUPTYPING, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces the contents of a rich edit control with a stream of data provided by an
' application defined EditStreamCallback callback function.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.StreamIn (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_STREAMIN, psf, cast(LPARAM, @edst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Causes a rich edit control to pass its contents to an application defined
' EditStreamCallback callback function.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.StreamOut (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_STREAMOUT, psf, cast(LPARAM, @edst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length of all text in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextLength () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, WM_GETTEXTLENGTH, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates text length in various ways. It is usually called before creating a buffer to
' receive the text from the control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextLengthEx (BYREF gtex AS .GETTEXTLENGTHEX) AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTEXTLENGTHEX, cast(WPARAM, @gtex), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Uses the unicode code page (1200).
' Returns the number of characters in the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextLengthEx (BYVAL dwFlags AS DWORD = GTL_DEFAULT) AS LONG
   m_Result = 0
   DIM gtex AS .GETTEXTLENGTHEX = TYPE<.GETTEXTLENGTHEX>(dwFlags, 1200)
   RETURN SendMessageW(m_hRichEdit, EM_GETTEXTLENGTHEX, cast(WPARAM, @gtex), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a specified range of characters from a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextRange (BYREF trg AS .TEXTRANGEW) AS DWORD
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_GETTEXTRANGE, 0, cast(LPARAM, @trg))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves a specified range of characters from a rich edit control.
' If the cpMin and cpMax members are equal, the range is empty.
' The range includes everything if cpMin is 0 and cpMax is –1.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextRange (BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1) AS CWSTR
   m_Result = 0
   ' // Calculate the length of the buffer
   DIM cbLen As LONG
   IF cpMin = cpMax THEN RETURN ""
   IF cpMin = 0 AND cpMax = -1 THEN
      cbLen = SendMessageW(m_hRichEdit, WM_GETTEXTLENGTH, 0, 0)
      IF cbLen = 0 THEN RETURN ""
   ELSE
      cbLen = cpMax - cpMin
      IF cbLen < 1 THEN
         m_Result = E_INVALIDARG
         RETURN ""
      END IF
   END IF
   DIM wstrText AS CWSTR
   wstrText = WSPACE(cbLen + 1)
   DIM trg AS TEXTRANGEW
   trg.chrg.cpMin = cpMin
   trg.chrg.cpMax = cpMax
   trg.lpstrText = wstrtext
   cbLen = SendMessageW(m_hRichEdit, EM_GETTEXTRANGE, 0, cast(LPARAM, @trg))
   RETURN LEFT(wstrtext, cbLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' This message undoes the last edit control operation in the control's undo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.Undo () AS LONG
   m_Result = 0
   RETURN SendMessageW(m_hRichEdit, EM_UNDO, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
'                                    HELPER PROCEDURES
' ========================================================================================

' ========================================================================================
' Enables word wrap.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.EnableWordWrap () AS BOOLEAN
   DIM hDC AS HDC = GetDC(m_hRichEdit)
   IF hDC = NULL THEN m_Result = E_FAIL : RETURN FALSE
   DIM hr AS HRESULT = SendMessageW(m_hRichEdit, EM_SETTARGETDEVICE, cast(WPARAM, hDC), 0)
   ReleaseDC(m_hRichEdit, hdc)
   IF hr <> 0 THEN
      m_Result = S_OK
      RETURN TRUE
   ELSE
      m_Result = S_FALSE
      RETURN FALSE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Disables word wrap.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.DisableWordWrap (BYVAL LineWidth AS LONG = 32767) AS BOOLEAN
   m_Result = 0
   DIM hDC AS HDC = GetDC(m_hRichEdit)
   IF hDC = NULL THEN m_Result = E_FAIL : RETURN FALSE
   DIM hr AS HRESULT = SendMessageW(m_hRichEdit, EM_SETTARGETDEVICE, cast(WPARAM, hDC), LineWidth)
   ReleaseDC(m_hRichEdit, hdc)
   IF hr <> 0 THEN
      m_Result = S_OK
      RETURN TRUE
   ELSE
      m_Result = S_FALSE
      RETURN FALSE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enumerates font families. Used by the SetFont method.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.CRichEditCtrlEnumFontFamProcW ( _
   BYVAL lpelf    AS ENUMLOGFONTW PTR, _     ' // Address of ENUMLOGFONT structure
   BYVAL lpntm    AS NEWTEXTMETRICW PTR, _   ' // Address of NEWTEXTMETRIC structure
   BYVAL FontType AS LONG, _                 ' // Font type
   BYVAL lplf     AS LOGFONTW PTR _          ' // Address of LOGFONT struct
   ) AS LONG

   lplf->lfCharSet        = lpelf->elfLogFont.lfCharSet
   lplf->lfPitchAndFamily = lpelf->elfLogFont.lfPitchAndFamily
   lplf->lfFaceName       = lpelf->elfLogFont.lfFaceName

   RETURN FALSE

END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font used by a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditctrl.SetFont ( _
   BYREF wszFaceName AS WSTRING, _        ' // Font name
   BYVAL ptsize AS LONG _                 ' // Font size in points
   ) AS HRESULT

   DIM hResult AS HRESULT                 ' // Result code
   DIM hDC AS HDC                         ' // Handle of the device context
   DIM tlf AS LOGFONTW                    ' // LOGFONT structure
   DIM tcf AS CHARFORMATW                 ' // CHARFORMATW structure

   hDC = GetDC(NULL)
   EnumFontFamiliesW(hDC, wszFaceName, cast(FONTENUMPROCW, @CRichEditCtrlEnumFontFamProcW), cast(LPARAM, @tlf))
   ReleaseDC NULL, hDC
   tcf.cbSize = SIZEOF(tcf)
   tcf.dwMask = CFM_BOLD OR CFM_ITALIC OR CFM_UNDERLINE OR CFM_STRIKEOUT OR _
                CFM_FACE OR CFM_CHARSET OR CFM_SIZE
   tcf.yHeight = ptsize * 20   ' // Expects it in 20ths of a point
   tcf.bCharSet = tlf.lfCharSet
   tcf.bPitchAndFamily = tlf.lfPitchAndFamily
   tcf.szFaceName = tlf.lfFaceName
   m_Result = SendMessageW(hRichEdit, EM_SETCHARFORMAT, SCF_ALL, cast(LPARAM, @tcf))
   RETURN m_Result

END FUNCTION
' ========================================================================================

' ========================================================================================
' Callback used by the GetRtf method.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.CRichEditCtrlGetTextCallback ( _
   BYVAL dwCookie AS DWORD_PTR _                      ' // Value of the dwCookie member of the EDITSTREAM structure.
 , BYVAL pbBuff AS BYTE PTR _                         ' // Pointer to the buffer to read from.
 , BYVAL cb AS LONG _                                 ' // Number of bytes to read.
 , BYVAL pcb AS LONG PTR _                            ' // Number of bytes actually read.
 ) AS DWORD                                           ' // 0 for success, or an error code

   DIM pcws AS CWSTR PTR = cast(CWSTR PTR, dwCookie)
   pcws->AppendBuffer(pbBuff, cb)
   *pcb = cb
   RETURN 0

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves RTF formatted text from a Rich Edit control
' Returns the retrieved text or a null string.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetRtf (BYVAL dwOptionsAndFlags AS DWORD = SF_RTF) AS STRING
   DIM eds AS EDITSTREAM, cws AS CWSTR
   eds.dwCookie = cast(DWORD_PTR, @cws)
   eds.pfnCallBack = cast(EDITSTREAMCALLBACK, @CRichEditCtrl.CRichEditCtrlGetTextCallback)
   SendMessageW(m_hRichEdit, EM_STREAMOUT, dwOptionsAndFlags, cast(LPARAM, @eds))
   ' // Copy the ansi contents of the CWSTR to a STRING
   DIM s AS STRING = SPACE(cws.m_BufferLen - 2)   ' // -2 to remove the ending nulls
   IF LEN(s) THEN CopyMemory(STRPTR(s), cws.m_pBuffer, cws.m_BufferLen - 2)
   m_Result = 0
   RETURN s
END FUNCTION
' ========================================================================================

' ========================================================================================
' Callback function used by the CRichEditCtrl.LoadRtfFromFile method.
' Transfers a stream of data into a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.CRichEditCtrlLoadRtfFromFileCallback ( _
   BYVAL hFile AS HANDLE _                  ' // Value of the dwCookie member of the EDITSTREAM structure.
 , BYVAL lpBuff AS BYTE PTR _               ' // Pointer to a buffer to write to
 , BYVAL cb AS LONG _                       ' // Maximum number of bytes to read
 , BYVAL pcb AS LONG PTR _                  ' // Number of bytes actually read
 ) AS UINT                                  ' // 0 for success, or an error code

   IF ReadFile(hFile, lpBuff, cb, pcb, NULL) = 0 THEN FUNCTION = GetLastError

END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads an RTF file.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LoadRtfFromFile ( _
   BYREF wszFileName AS WSTRING _           ' // Name of the file to load
 ) AS BOOLEAN                               ' // TRUE or FALSE

   DIM hFile AS HANDLE                      ' // File handle
   DIM eds AS EDITSTREAM                    ' // EDITSTREAM structure

   ' // Checks the validity of the parameters
   IF m_hRichEdit = 0 THEN m_Result = E_HANDLE : RETURN FALSE
   IF LEN(wszFileName) = 0 THEN m_Result = E_INVALIDARG : RETURN FALSE
   m_Result = 0

   ' // Opens the file and sends the message
   hFile = CreateFileW(wszFileName, GENERIC_READ, FILE_SHARE_READ, _
                       NULL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN m_Result = E_HANDLE: RETURN FALSE
   eds.dwCookie = cast(DWORD_PTR, hFile)
   eds.pfnCallback = cast(EDITSTREAMCALLBACK, @CRichEditCtrlLoadRtfFromFileCallback)
   IF SendMessageW(hRichEdit, EM_STREAMIN, SF_RTF, cast(LPARAM, @eds)) > 0 AND eds.dwError = 0 THEN FUNCTION = TRUE
   m_Result = eds.dwError
   CloseHandle hFile

END FUNCTION
' ========================================================================================

' ========================================================================================
' Callback function used by the CRichEditCtrl.LoadRtfFromResource method.
' Transfers a stream of data into a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.CRichEditCtrlLoadRtfFromResourceCallback ( _
   BYVAL pCustData AS AFX_CRICHEDITCTRL_CUSTOMDATA PTR _   ' // Value of the dwCookie member of the EDITSTREAM structure.
 , BYVAL lpBuff AS BYTE PTR _                              ' // Pointer to a buffer to write to.
 , BYVAL cb AS LONG _                                      ' // Number of bytes to write.
 , BYVAL pcb AS LONG PTR _                                 ' // Number of bytes actually written.
 ) AS DWORD                                                ' // 0 for success, or an error code

   DIM nBytes AS LONG
   IF pCustData->nLen - pCustData->curPos > cb THEN nBytes = cb ELSE nBytes = pCustData->nLen - pCustData->curPos
   IF nBytes THEN
      CopyMemory(lpBuff, pCustData->pData + pCustData->curPos, nBytes)
      pCustData->curPos = pCustData->curPos + nBytes
      FUNCTION = 0
   ELSE
      FUNCTION = 1
   END IF
   *pcb = nBytes

END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads a RTF resource file into a Rich Edit control.
' The EM_STREAMIN message replaces the contents of a rich edit control with a stream of
' data provided by an application defined EditStreamCallback callback function.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LoadRtfFromResource ( _
   BYREF wszResourceName AS WSTRING _                ' // Name of the resource to load
 ) AS BOOLEAN                                        ' // TRUE or FALSE

   DIM hResInfo AS HRSRC                             ' // Resource handle
   DIM pResData AS LPVOID                            ' // Pointer to the resource data
   DIM eds AS EDITSTREAM                             ' // EDITSTREAM structure
   DIM rtfCustData AS AFX_CRICHEDITCTRL_CUSTOMDATA   ' // AFX_CRICHEDITCTRL_CUSTOMDATA structure

   ' // Checks the validity of the parameters
   IF m_hRichEdit = 0 THEN m_Result = E_HANDLE : RETURN FALSE
   IF LEN(wszResourceName) = 0 THEN m_Result = E_INVALIDARG : RETURN FALSE
   m_Result = 0

   ' // Loads the resource
   hResInfo = FindResourceW(m_hInstance, wszResourceName, RT_RCDATA)
   IF hResInfo = NULL THEN m_Result = GetLastError : RETURN FALSE

   ' // Loads and locks the resource
   ' // Note  LockResource does not actually lock memory; it is just used to obtain
   ' // a pointer to the memory containing the resource data.
   pResData = LockResource(LoadResource(m_hInstance, hResInfo))
   IF pResData = NULL THEN m_Result = GetLastError : RETURN FALSE
   DIM cbSize AS LONG = SizeofResource(m_hInstance, hResInfo)
   DIM buffer AS STRING = SPACE(cbSize)
   CopyMemory(STRPTR(buffer), pResData, cbSize)

   ' // Sends the message
   rtfCustData.pData = STRPTR(buffer)
   rtfCustData.nLen = cbSize
   rtfCustData.curPos = 0
   eds.dwCookie = cast(DWORD_PTR, @rtfCustData)
   eds.pfnCallback = cast(EDITSTREAMCALLBACK, @CRichEditCtrlLoadRtfFromResourceCallback)
   IF SendMessageW(hRichEdit, EM_STREAMIN, SF_RTF, cast(LPARAM, @eds)) > 0 AND eds.dwError = 0 THEN
      RETURN TRUE
   ELSE
      m_Result = eds.dwError
      RETURN FALSE
   END IF

END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets RTF text from the specified file
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetRtfFromFile (BYREF wszFileName AS WSTRING) AS STRING
   DIM cb AS LONG                           ' // Length of the file
   DIM pcb AS LONG PTR                      ' // Bytes read
   DIM hFile AS HANDLE                      ' // File handle
   DIM buffer AS STRING                     ' // Buffer

   ' // Checks the validity of the parameters
   IF m_hRichEdit = 0 THEN m_Result = E_HANDLE : RETURN buffer
   IF LEN(wszFileName) = 0 THEN m_Result = E_INVALIDARG : RETURN buffer
   m_Result = 0
   ' // Opens the file and sends the message
   hFile = CreateFileW(wszFileName, GENERIC_READ, FILE_SHARE_READ, _
                       NULL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN m_Result = E_HANDLE: RETURN buffer
   ' // Get the
   cb = GetFileSize(hFile, NULL)
   buffer = SPACE(cb)
   IF ReadFile(hFile, STRPTR(buffer), cb, pcb, NULL) = 0 THEN m_Result = GetLastError
   CloseHandle hFile
   RETURN buffer

END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends the contents of the specified RTF file
' Usage example:
' DIM wszFileName AS WSTRING * MAX_PATH = AfxGetExePath & $"\Test.rtf"
' pRichEdit->AppendRtfFile(wszFileName)
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.AppendRtfFile (BYREF wszFileName AS WSTRING) AS DWORD
   ' // Checks the validity of the parameters
   IF m_hRichEdit = 0 THEN m_Result = E_HANDLE : RETURN 0
   IF LEN(wszFileName) = 0 THEN m_Result = E_INVALIDARG : RETURN 0
   m_Result = 0
   ' // Gets all the text from the file
   DIM strRtf AS STRING = this.GetRtfFromFile(wszFileName)
   IF LEN(strRtf) = 0 THEN RETURN 0
   ' // Fills the SETTEXTEX structure
   DIM stex AS .SETTEXTEX
   stex.flags = ST_SELECTION OR ST_KEEPUNDO
   stex.codepage = CP_ACP
   ' // Moves the caret at the end
   this.SetCaretPos(-1)
   ' // Gets the caret position
   DIM dwPos AS DWORD = this.GetCaretPos
   ' // Set the selection at the end
   this.SetSel(dwPos, dwPos)
   ' // Sets the Rtf text
   RETURN SendMessageW(m_hRichEdit, EM_SETTEXTEX, cast(WPARAM, @stex), cast(LPARAM, STRPTR(strRtf)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts the contents of the specified RTF file
' wszFilename : The name of the rtf file to append.
' dwPos : The position where to insert the contents of the rtf file.
'         If dwPos = -1, the rtf contents are inserted at the caret position.
' Usage example:
' DIM wszFileName AS WSTRING * MAX_PATH = AfxGetExePath & $"\Test.rtf"
' pRichEdit->InsertRtfFile(wszFileName, 500)
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.InsertRtfFile (BYREF wszFileName AS WSTRING, BYVAL dwPos AS DWORD) AS DWORD
   ' // Checks the validity of the parameters
   IF m_hRichEdit = 0 THEN m_Result = E_HANDLE : RETURN 0
   IF LEN(wszFileName) = 0 THEN m_Result = E_INVALIDARG : RETURN 0
   m_Result = 0
   ' // Gets all the text from the file
   DIM strRtf AS STRING = this.GetRtfFromFile(wszFileName)
   IF LEN(strRtf) = 0 THEN RETURN 0
   ' // Fills the SETTEXTEX structure
   DIM stex AS .SETTEXTEX
   stex.flags = ST_SELECTION OR ST_KEEPUNDO
   stex.codepage = CP_ACP
   ' // Moves the caret at the specified location
   IF CLNG(dwPos) = -1 THEN
      dwPos = this.GetCaretPos
   ELSE
      this.SetCaretPos(dwPos)
   END IF
   ' // Set the selection at the specified location
   this.SetSel(dwPos, dwPos)
   ' // Sets the Rtf text
   RETURN SendMessageW(m_hRichEdit, EM_SETTEXTEX, cast(WPARAM, @stex), cast(LPARAM, STRPTR(strRtf)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts an image or an Ole object in the rich edit control.
' See MSDN: https://learn.microsoft.com/en-us/windows/win32/controls/using-rich-edit-com
' Remarks: To insert images, use the InsertImage method, because the InsertObject method
' won't display the image, but an icon, when the image type is not a bitmap.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.InsertObject ( _
   BYREF wszFileName AS WSTRING _                    ' // Name of the file to load
 ) AS BOOLEAN                                        ' // TRUE or FALSE

   CREC_DP("#### BEGIN - CRichEditCtrl.InsertObject")

   ' // Get the IRichEditOle interface
   CREC_DP("#### CRichEditCtrl.InsertObject - GetOleInterface")
   DIM pRichEditOle AS IRichEditOle PTR
   m_Result = SendMessageW(m_hRichEdit, EM_GETOLEINTERFACE, 0, cast(LPARAM, @pRichEditOle))
   IF pRichEditOle = NULL THEN RETURN FALSE

   ' // Create structured storage
   CREC_DP("#### CRichEditCtrl.InsertObject - CreateILockBytesOnHGlobal")
   DIM pLockBytes AS ILockBytes PTR
   m_Result = CreateILockBytesOnHGlobal(NULL, TRUE, @pLockBytes)
   IF FAILED(m_Result)THEN RETURN FALSE

   CREC_DP("#### CRichEditCtrl.InsertObject - StgCreateDocfileOnILockBytes")
   DIM pStorage AS IStorage PTR
   m_Result = StgCreateDocfileOnILockBytes(pLockBytes, _
      STGM_SHARE_EXCLUSIVE OR STGM_CREATE OR STGM_READWRITE, 0, @pStorage)
   IF FAILED(m_Result) THEN RETURN FALSE

   ' // Set up the data format
   DIM formatEtc AS FORMATETC
   formatEtc.cfFormat = 0
   formatEtc.ptd = NULL
   formatEtc.dwAspect = DVASPECT_CONTENT
   formatEtc.lindex = -1
   formatEtc.tymed = TYMED_NULL

   ' // Get a pointer to the display site
   CREC_DP("#### CRichEditCtrl.InsertObject - GetClientSite")
   DIM pClientSite AS IOleClientSite PTR
   m_Result = pRichEditOle->lpvtbl->GetClientSite(pRichEditOle, @pClientSite)
   IF FAILED(m_Result) THEN RETURN FALSE

   ' // Create the object and retrieve its IUnknown interface
   CREC_DP("#### CRichEditCtrl.InsertObject - OleCreateFromFile")
   DIM pUnk AS IUnknown PTR
   m_Result = OleCreateFromFile(@CLSID_NULL, @wszFileName, @IID_IUnknown, OLERENDER_DRAW, _
           @formatEtc, pClientSite, pStorage, @pUnk)
   pClientSite->lpvtbl->Release(pClientSite)
   IF FAILED(m_Result) THEN RETURN FALSE

   ' // Get the IOleObject interface to the object
   CREC_DP("#### CRichEditCtrl.InsertObject - IOleObject")
   DIM pObject AS IOleObject PTR
   m_Result = pUnk->lpvtbl->QueryInterface(pUnk, @IID_IOleObject, @pObject)
   pUnk->lpvtbl->Release(pUnk)
   IF FAILED(m_Result) THEN RETURN FALSE

   ' // To ensure that references are counted correctly, notify the object that it is contained
   CREC_DP("#### CRichEditCtrl.InsertObject - GetUserClassID")
   DIM clsid AS CLSID
   OleSetContainedObject(cast(IUnknown PTR, pObject), TRUE)
   ' // Get the object class identifier
   m_Result = pObject->lpvtbl->GetUserClassID(pObject, @clsid)
   IF FAILED(m_Result) THEN
      pObject->lpvtbl->Release(pObject)
      RETURN FALSE
   END IF

   ' // Set up object info
   DIM reobject AS REOBJECT
   reobject.cbStruct = SIZEOF(REOBJECT)
   reobject.clsid = clsid
   reobject.cp = REO_CP_SELECTION
   reobject.dvaspect = DVASPECT_CONTENT
   reobject.dwFlags = REO_RESIZABLE 'OR REO_BELOWBASELINE
   reobject.dwUser = 0
   reobject.poleobj = pObject
   reobject.polesite = pClientSite
   reobject.pstg = pStorage
   DIM sizel AS SIZEL
   DIM dwWidth AS DWORD, dwHeight AS DWORD
   DIM nStatus AS LONG = AfxGdipGetImageSizeFromFile(wszFileName, @dwWidth, @dwHeight)
   IF nStatus = 0 THEN
      sizel.cx = AfxPixelsToHiMetricX(dwWidth * m_ScalingRatio)
      sizel.cy = AfxPixelsToHiMetricY(dwHeight * m_ScalingRatio)
   END IF
   reobject.sizel = sizel

   ' // Insert the object
   CREC_DP("#### CRichEditCtrl.InsertObject - InsertObject")
   m_Result = pRichEditOle->lpvtbl->InsertObject(pRichEditOle, @reobject)
   pObject->lpvtbl->Release(pObject)
   pRichEditOle->lpvtbl->Release(pRichEditOle)
   IF FAILED(m_Result) THEN RETURN FALSE
   
   CREC_DP("#### END - CRichEditCtrl.InsertObject hr = " & WSTR(m_Result))
   ' // Return success
   RETURN TRUE

END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a line feed at the cursor position.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.AddLF (BYVAL atEnd AS BOOLEAN = FALSE)
   IF atEnd THEN
      ' // Move the caret to the end of the text
      SendMessageW(m_hRichEdit, EM_SETSEL, 0, -1)
      DIM dwStart AS DWORD, dwEnd AS DWORD
      SendMessageW(m_hRichEdit, EM_GETSEL, cast(WPARAM, @dwStart), cast(LPARAM, @dwEnd))
      SendMessageW(m_hRichEdit, EM_SETSEL, cast(WPARAM, @dwEnd + 1), cast(LPARAM, @dwEnd + 1))
   END IF
   ' // Insert a line feed
   DIM wsz AS WSTRING * 2 = CHR(10)
   SendMessageW(m_hRichEdit, EM_REPLACESEL, TRUE, cast(LPARAM, @wsz))
END SUB
' ========================================================================================
' ========================================================================================
' Inserts a carriage return at the cursor position.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.AddCR (BYVAL atEnd AS BOOLEAN = FALSE)
   IF atEnd THEN
      ' // Move the caret to the end of the text
      SendMessageW(m_hRichEdit, EM_SETSEL, 0, -1)
      DIM dwStart AS DWORD, dwEnd AS DWORD
      SendMessageW(m_hRichEdit, EM_GETSEL, cast(WPARAM, @dwStart), cast(LPARAM, @dwEnd))
      SendMessageW(m_hRichEdit, EM_SETSEL, cast(WPARAM, @dwEnd + 1), cast(LPARAM, @dwEnd + 1))
   END IF
   ' // Insert a carriage return
   DIM wsz AS WSTRING * 2 = CHR(13)
   SendMessageW(m_hRichEdit, EM_REPLACESEL, TRUE, cast(LPARAM, @wsz))
END SUB
' ========================================================================================
' ========================================================================================
' Inserts a carriage return and a line feed at the cursor position.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.AddCRLF (BYVAL atEnd AS BOOLEAN = FALSE)
   IF atEnd THEN
      ' // Move the caret to the end of the text
      SendMessageW(m_hRichEdit, EM_SETSEL, 0, -1)
      DIM dwStart AS DWORD, dwEnd AS DWORD
      SendMessageW(m_hRichEdit, EM_GETSEL, cast(WPARAM, @dwStart), cast(LPARAM, @dwEnd))
      SendMessageW(m_hRichEdit, EM_SETSEL, cast(WPARAM, @dwEnd + 1), cast(LPARAM, @dwEnd + 2))
   END IF
   ' // Insert a carriage return and a line feed
   DIM wsz AS WSTRING * 3 = CHR(13, 10)
   SendMessageW(m_hRichEdit, EM_REPLACESEL, TRUE, cast(LPARAM, @wsz))
END SUB
' ========================================================================================

' ========================================================================================
' Saves the contents of the rich edit control to a file in rtf format
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SaveRtf (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
   ' // Get the Rtf text
   DIM strRTFData AS STRING = this.GetRTF(SF_RTF)
   ' // Create the file
   m_Result = 0
   DIM dwCreationFlag AS DWORD
   IF Overwrite THEN dwCreationFlag = CREATE_ALWAYS ELSE dwCreationFlag = CREATE_NEW
   DIM hFile AS HANDLE = CreateFileW(wszFileName, GENERIC_WRITE, 0, NULL, dwCreationFlag, dwFlagsAndAttributes, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN m_Result = GetLastError : RETURN FALSE
   ' // Save the rtf text
   DIM dwBytesWritten AS DWORD
   DIM bSuccess AS BOOLEAN = WriteFile(hFile, STRPTR(strRTFData), LEN(strRTFData), @dwBytesWritten, NULL)
   IF bSuccess = FALSE THEN m_Result = GetLastError
   CloseHandle(hFile)
   RETURN bSuccess
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves the contents of the rich edit control to file in rtf format
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SaveRtfNoObjs (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
   ' // Get the Rtf text
   DIM strRTFData AS STRING = this.GetRTF(SF_RTFNOOBJS)
   ' // Create the file
   m_Result = 0
   DIM dwCreationFlag AS DWORD
   IF Overwrite THEN dwCreationFlag = CREATE_ALWAYS ELSE dwCreationFlag = CREATE_NEW
   DIM hFile AS HANDLE = CreateFileW(wszFileName, GENERIC_WRITE, 0, NULL, dwCreationFlag, dwFlagsAndAttributes, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN m_Result = GetLastError : RETURN FALSE
   ' // Save the rtf text
   DIM dwBytesWritten AS DWORD
   DIM bSuccess AS BOOLEAN = WriteFile(hFile, STRPTR(strRTFData), LEN(strRTFData), @dwBytesWritten, NULL)
   IF bSuccess = FALSE THEN m_Result = GetLastError
   CloseHandle(hFile)
   RETURN bSuccess
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves selection of the rich edit control to a file in rtf format
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SaveSelRtf (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
   ' // Get the Rtf text
   DIM strRTFData AS STRING = this.GetRTF(SF_RTF OR SFF_SELECTION)
   ' // Create the file
   m_Result = 0
   DIM dwCreationFlag AS DWORD
   IF Overwrite THEN dwCreationFlag = CREATE_ALWAYS ELSE dwCreationFlag = CREATE_NEW
   DIM hFile AS HANDLE = CreateFileW(wszFileName, GENERIC_WRITE, 0, NULL, dwCreationFlag, dwFlagsAndAttributes, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN m_Result = GetLastError : RETURN FALSE
   ' // Save the rtf text
   DIM dwBytesWritten AS DWORD
   DIM bSuccess AS BOOLEAN = WriteFile(hFile, STRPTR(strRTFData), LEN(strRTFData), @dwBytesWritten, NULL)
   IF bSuccess = FALSE THEN m_Result = GetLastError
   CloseHandle(hFile)
   RETURN bSuccess
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves selection of the rich edit control to file in rtf format
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SaveSelRtfNoObjs (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
   ' // Get the Rtf text
   DIM strRTFData AS STRING = this.GetRTF(SF_RTFNOOBJS OR SFF_SELECTION)
   ' // Create the file
   m_Result = 0
   DIM dwCreationFlag AS DWORD
   IF Overwrite THEN dwCreationFlag = CREATE_ALWAYS ELSE dwCreationFlag = CREATE_NEW
   DIM hFile AS HANDLE = CreateFileW(wszFileName, GENERIC_WRITE, 0, NULL, dwCreationFlag, dwFlagsAndAttributes, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN m_Result = GetLastError : RETURN FALSE
   ' // Save the rtf text
   DIM dwBytesWritten AS DWORD
   DIM bSuccess AS BOOLEAN = WriteFile(hFile, STRPTR(strRTFData), LEN(strRTFData), @dwBytesWritten, NULL)
   IF bSuccess = FALSE THEN m_Result = GetLastError
   CloseHandle(hFile)
   RETURN bSuccess
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves the contents of the rich edit control in text format.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SaveText (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
   ' // Get the Rtf text
   DIM strText AS STRING = this.Text
   ' // Create the file
   m_Result = 0
   DIM dwCreationFlag AS DWORD
   IF Overwrite THEN dwCreationFlag = CREATE_ALWAYS ELSE dwCreationFlag = CREATE_NEW
   DIM hFile AS HANDLE = CreateFileW(wszFileName, GENERIC_WRITE, 0, NULL, dwCreationFlag, dwFlagsAndAttributes, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN m_Result = GetLastError : RETURN FALSE
   ' // Save the text
   DIM dwBytesWritten AS DWORD
   DIM bSuccess AS BOOLEAN = WriteFile(hFile, STRPTR(strText), LEN(strText), @dwBytesWritten, NULL)
   IF bSuccess = FALSE THEN m_Result = GetLastError
   CloseHandle(hFile)
   RETURN bSuccess
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves the selected text of the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SaveSelText (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
   m_Result = 0
   ' // Get the Rtf text
   DIM strText AS STRING = this.GetSelText
   ' // Create the file
   DIM dwCreationFlag AS DWORD
   IF Overwrite THEN dwCreationFlag = CREATE_ALWAYS ELSE dwCreationFlag = CREATE_NEW
   DIM hFile AS HANDLE = CreateFileW(wszFileName, GENERIC_WRITE, 0, NULL, dwCreationFlag, dwFlagsAndAttributes, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN m_Result = GetLastError : RETURN FALSE
   ' // Save the text
   DIM dwBytesWritten AS DWORD
   DIM bSuccess AS BOOLEAN = WriteFile(hFile, STRPTR(strText), LEN(strText), @dwBytesWritten, NULL)
   IF bSuccess = FALSE THEN m_Result = GetLastError
   CloseHandle(hFile)
   RETURN bSuccess
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the color of the selected word or text
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextColor () AS COLORREF
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN cf.crTextColor
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the color of the selected word or text
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextColor (BYVAL crTxtColor AS COLORREF)
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_COLOR
   cf.crTextColor = crTxtColor
   this.SetWordCharFormat(cf)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the color of the selected word or text
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextColor () AS COLORREF
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN cf.crTextColor
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the color of the selected word or text
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTextColor (BYVAL crTxtColor AS COLORREF) AS BOOLEAN
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_COLOR
   cf.crTextColor = crTxtColor
   RETURN this.SetWordCharFormat(cf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Checks if the text is bold
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.IsTextBold () AS BOOLEAN
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN IIF((cf.dwEffects AND CFE_BOLD) = CFE_BOLD, TRUE, FALSE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the selected text to bold.
' If it is already set, removes it.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetTextBold
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_BOLD
   cf.dwEffects = CFE_BOLD
   ' // if it is already set, remove it
   IF this.IsTextBold = TRUE THEN cf.dwEffects = 0
   this.SetWordCharFormat(cf)
END SUB
' ========================================================================================

' ========================================================================================
' Checks if the text is italic
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.IsTextItalic () AS BOOLEAN
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN IIF((cf.dwEffects AND CFE_ITAlIC) = CFE_ITAlIc, TRUE, FALSE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the selected text to italic
' If it is already set, removes it.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetTextItalic
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_ITALIC
   cf.dwEffects = CFE_ITALIC
   ' // if it is already set, remove it
   IF this.IsTextItalic = TRUE THEN cf.dwEffects = 0
   this.SetWordCharFormat(cf)
END SUB
' ========================================================================================

' ========================================================================================
' Checks if the text is striked out
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.IsTextStrikeOut () AS BOOLEAN
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN IIF((cf.dwEffects AND CFE_STRIKEOUT) = CFE_STRIKEOUT, TRUE, FALSE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the selected text to striked out
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetTextStrikeOut
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_STRIKEOUT
   cf.dwEffects = CFE_STRIKEOUT
   ' // if it is already set, remove it
   IF this.IsTextStrikeOut = TRUE THEN cf.dwEffects = 0
   this.SetWordCharFormat(cf)
END SUB
' ========================================================================================

' ========================================================================================
' Checks if the text is underlined
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.IsTextUnderline () AS BOOLEAN
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN IIF((cf.dwEffects AND CFE_UNDERLINE) = CFE_UNDERLINE, TRUE, FALSE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the selected text to underlined
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetTextUnderline
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_UNDERLINE
   cf.dwEffects = CFE_UNDERLINE
   ' // if it is already set, remove it
   IF this.IsTextUnderline = TRUE THEN cf.dwEffects = 0
   this.SetWordCharFormat(cf)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the font height of selected text in points
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextHeight () AS LONG
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN cf.yHeight / 20
END PROPERTY
' ========================================================================================
' ========================================================================================
' Changes the font height of selected text
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextHeight (BYVAL ptSize AS LONG)
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_SIZE
   ' Character height, in twips (1/1440 of an inch or 1/20 of a printer's point)
   cf.yHeight = ptSize * 20
   this.SetWordCharFormat(cf)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the font height of selected text in points
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextHeight () AS LONG
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN cf.yHeight / 20
END FUNCTION
' ========================================================================================
' ========================================================================================
' Changes the font height of selected text
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTextHeight (BYVAL ptSize AS LONG) AS BOOLEAN
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_SIZE
   ' Character height, in twips (1/1440 of an inch or 1/20 of a printer's point)
   cf.yHeight = ptSize * 20
   RETURN this.SetWordCharFormat(cf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the font offset of selected text
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextOffset () AS LONG
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN cf.yOffset / 20
END PROPERTY
' ========================================================================================
' ========================================================================================
' Changes the offset of selected text
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextOffset (BYVAL offset AS LONG)
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_OFFSET
   ' Character offset, in twips, from the baseline
   cf.yOffset = offset * 20
   this.SetWordCharFormat(cf)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the font offset of selected text
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextOffset () AS LONG
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN cf.yOffset / 20
END FUNCTION
' ========================================================================================
' ========================================================================================
' Changes the offset of selected text
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTextOffset (BYVAL offset AS LONG) AS BOOLEAN
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_OFFSET
   ' Character offset, in twips, from the baseline
   cf.yOffset = offset * 20
   RETURN this.SetWordCharFormat(cf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the font face name of selected text
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextFontName () AS CWSTR
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN cf.szFaceName
END PROPERTY
' ========================================================================================
' ========================================================================================
' Changes the font face name of selected text
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextFontName (BYREF wszFaceName AS WSTRING)
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_FACE
   cf.szFaceName = wszFaceName
   this.SetWordCharFormat(cf)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the font face name of selected text
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextFontName () AS CWSTR
   m_Result = 0
   DIM cf AS CHARFORMATW = this.GetWordCharFormat
   RETURN cf.szFaceName
END FUNCTION
' ========================================================================================
' ========================================================================================
' Changes the font face name of selected text
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTextFontName (BYREF wszFaceName AS WSTRING) AS BOOLEAN
   m_Result = 0
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   cf.dwMask = CFM_FACE
   cf.szFaceName = wszFaceName
   RETURN this.SetWordCharFormat(cf)
END FUNCTION
' ========================================================================================

' ########################################################################################
' Wrapper methods and propesties for the ITexDocument2 interface
' ########################################################################################

' ========================================================================================
' Retrieves a pointer to the ITextDocument2 interface
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetDocInterface () AS ITextDocument2 PTR
   ' // Retrieve a pointer to the IRichEditOle interface of the rich edit control
   DIM pUnk AS IUnknown PTR, pTextDocument2 AS ITextDocument2 PTR
   m_Result = SendMessageW(hRichEdit, EM_GETOLEINTERFACE, 0, cast(LPARAM, @pUnk))
   ' // Retrieve a pointer to its ITextDocument2 interface
   IF pUnk THEN
      DIM IID_ITextDocument2_ AS IID = AfxGuid(AFX_IID_ITextDocument2)
      m_Result = IUnknown_QueryInterface(pUnk, @IID_ITextDocument2_, @pTextDocument2)
      IUnknown_Release(pUnk)
   END IF
   RETURN pTextDocument2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens a new document. If another document is open, this method saves any current changes
' and closes the current document before opening a new one.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.NewDoc () AS HRESULT
   DIM pTextDocument2 AS ITextDocument2 PTR = this.GetDocInterface
   IF pTextDocument2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   m_Result = pTextDocument2->lpvtbl->New_(pTextDocument2)
   IUnknown_Release(pTextDocument2)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens a new document. If another document is open, this method saves any current changes
' and closes the current document before opening a new one.
' Example:
' DIM cv AS CVAR = AfxGetExePath & $"\Test.rtf"
' pRichEditCtrl.OpenDoc(cv)
' Default code page is CP_UNICODE (1200).
' If it is an ansi text file, pass CP_ANSI.
' If it UTF8 pas CP_UTF8.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.OpenDoc (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG = 0, BYVAL CodePage AS LONG = CP_ACP) AS HRESULT
   DIM pTextDocument2 AS ITextDocument2 PTR = this.GetDocInterface
   IF pTextDocument2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   m_Result = pTextDocument2->lpvtbl->Open(pTextDocument2, pVar, Flags, CodePage)
   IUnknown_Release(pTextDocument2)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves the document to a file in the specified format.
' pVar = Path and file name, e.g. 
'    DIM cv AS CVAR = AfxGetExePath & $"\Test3.rtf"
'    pRichEdit->SaveDoc(cv, tomCreateAlways OR tomRTF)
' Pass NULL to overwrite the file if it already exists.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SaveDoc (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG = 0, BYVAL CodePage AS LONG = CP_ACP) AS HRESULT
   DIM pTextDocument2 AS ITextDocument2 PTR = this.GetDocInterface
   IF pTextDocument2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   m_Result = pTextDocument2->lpvtbl->Save(pTextDocument2, pVar, Flags, CodePage)
   IUnknown_Release(pTextDocument2)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the file name of this document. This is the ITextDocument2 default property.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.DocName () AS CWSTR
   DIM cws AS CWSTR
   DIM pTextDocument2 AS ITextDocument2 PTR = this.GetDocInterface
   IF pTextDocument2 = NULL THEN m_Result = E_POINTER : RETURN cws
   DIM bstr AS AFX_BSTR
   m_Result = pTextDocument2->lpvtbl->GetName(pTextDocument2, @bstr)
   IUnknown_Release(pTextDocument2)
   cws = *bstr
   SysFreeString(bstr)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets/sets the Saved property
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Saved () AS BOOLEAN
   m_Result = 0
   DIM pTextDocument2 AS ITextDocument2 PTR = this.GetDocInterface
   IF pTextDocument2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM value AS LONG
   pTextDocument2->lpvtbl->GetSaved(pTextDocument2, @value)
   IUnknown_Release(pTextDocument2)
   IF value <> tomTrue THEN RETURN FALSE
   RETURN TRUE
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Saved (BYVAL value AS BOOLEAN)
   m_Result = 0
   DIM pTextDocument2 AS ITextDocument2 PTR = this.GetDocInterface
   IF pTextDocument2 = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   pTextDocument2->lpvtbl->SetSaved(pTextDocument2, IIF(value = TRUE, tomTrue, tomFalse))
   IUnknown_Release(pTextDocument2)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Increments the freeze count.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.Freeze () AS LONG
   DIM pTextDocument2 AS ITextDocument2 PTR = this.GetDocInterface
   IF pTextDocument2 = NULL THEN m_Result = E_POINTER : EXIT FUNCTION
   DIM nCount AS LONG
   m_Result = pTextDocument2->lpvtbl->Freeze(pTextDocument2, @nCount)
   IUnknown_Release(pTextDocument2)
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the freeze count.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.Unfreeze () AS LONG
   DIM pTextDocument2 AS ITextDocument2 PTR = this.GetDocInterface
   IF pTextDocument2 = NULL THEN m_Result = E_POINTER : EXIT FUNCTION
   DIM nCount AS LONG
   m_Result = pTextDocument2->lpvtbl->Unfreeze(pTextDocument2, @nCount)
   IUnknown_Release(pTextDocument2)
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the selection and caret.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.Update (BYVAL Value AS LONG) AS HRESULT
   DIM pTextDocument2 AS ITextDocument2 PTR = this.GetDocInterface
   IF pTextDocument2 = NULL THEN m_Result = E_POINTER : EXIT FUNCTION
   m_Result = pTextDocument2->lpvtbl->Update(pTextDocument2, Value)
   IUnknown_Release(pTextDocument2)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the case of letters in a range according to the nType parameter.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.ChangeCase (BYVAL cpActive AS LONG = 0, BYVAL cpAnchor AS LONG = 0, BYVAL nType AS LONG = tomLowerCase) AS HRESULT
   DIM pTextDocument2 AS ITextDocument2 PTR = this.GetDocInterface
   IF pTextDocument2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pRange2 AS ITextRange2 PTR
   m_Result = pTextDocument2->lpvtbl->Range2(pTextDocument2, cpActive, cpAnchor, @pRange2)
   IF pRange2 = NULL THEN IUnknown_Release(pTextDocument2) : RETURN m_Result
   m_Result = pRange2->lpvtbl->ChangeCase(pRange2, nType)
   IUnknown_Release(pRange2)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the case letters in a range to lower case
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LowerCase (BYVAL nStart AS LONG, BYVAL nLen AS LONG) AS HRESULT
   m_Result = this.ChangeCase(nStart, nStart + nLen, tomLowerCase)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Changes the case letters in a range to upper case
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.UpperCase (BYVAL nStart AS LONG, BYVAL nLen AS LONG) AS HRESULT
   m_Result = this.ChangeCase(nStart, nStart + nLen, tomUpperCase)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Changes the case letters in a range to title case
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.TitleCase (BYVAL nStart AS LONG, BYVAL nLen AS LONG) AS HRESULT
   m_Result = this.ChangeCase(nStart, nStart + nLen, tomTitleCase)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Changes the case letters in a range to sentence case
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SentenceCase (BYVAL nStart AS LONG, BYVAL nLen AS LONG) AS HRESULT
   m_Result = this.ChangeCase(nStart, nStart + nLen, tomSentenceCase)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Changes the case letters in a range to sentence case
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.ToggleCase (BYVAL nStart AS LONG, BYVAL nLen AS LONG) AS HRESULT
   m_Result = this.ChangeCase(nStart, nStart + nLen, tomToggleCase)
   RETURN m_Result
END FUNCTION
' ========================================================================================


' ########################################################################################
'                          H E L P E R    F U N C T I O N S
' ########################################################################################

' ========================================================================================
' Retrieves a pointer to the CRichEditCtrl class from the handle of the RichEdit control.
' hRichEdit = Handle of the Rich Edit control.
' Usage example:
' DIM pRichEdit AS CRichEditCtrl PTR = AfxCRichEditCtrlPtr(CAST(HWND, lParam))
' ========================================================================================
PRIVATE FUNCTION AfxCRichEditCtrlPtr OVERLOAD (BYVAL hRichEdit AS HWND) AS CRichEditCtrl PTR
   DIM pRichEditCtrl AS CRichEditCtrl PTR = CAST(CRichEditCtrl PTR, SendMessageW(hRichEdit, WM_USER + 100, 0, 0))
   RETURN pRichEditCtrl
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a pointer to the class given the handle of the parent window of the control and
' the identifier of the control.
' hParent = Handle of the parent window of the rich edit control.
' cID = Identifier of the rich edit control.
' ========================================================================================
PRIVATE FUNCTION AfxCRichEditCtrlPtr OVERLOAD (BYVAL hParent AS HWND, BYVAL cID AS LONG) AS CRichEditCtrl PTR
   DIM pRichEditCtrl AS CRichEditCtrl PTR = CAST(CRichEditCtrl PTR, SendMessageW(GetDlgItem(hParent, cID), WM_USER + 100, 0, 0))
   RETURN pRichEditCtrl
END FUNCTION
' ========================================================================================

END NAMESPACE
