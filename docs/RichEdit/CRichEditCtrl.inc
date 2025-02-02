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
#include once "win/richedit.bi"
#include once "win/richole.bi"
#include once "Afx/CTextObjectBase.inc"

NAMESPACE Afx
' ========================================================================================
' Macro for debug
' To allow debugging, define _CREC_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CREC_DEBUG_
   #define _CREC_DEBUG_ 0
#ENDIF
#ifndef _CREC_DP_
   #define _CREC_DP_ 1
   #MACRO CREC_DP(st)
      #IF (_CREC_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

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

   Public:
      DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszTitle AS WSTRING = "", _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
         BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION hRichEdit () AS HWND

      ' // Static callback procedure
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
      DECLARE PROPERTY AutoUrlDetect () AS LONG
      DECLARE PROPERTY AutoUrlDetect (BYVAL fUrlDetect AS LONG)
      DECLARE PROPERTY BidiOptions () AS .BIDIOPTIONS
      DECLARE PROPERTY BidiOptions (BYREF _Options AS .BIDIOPTIONS)
      DECLARE PROPERTY CharFormat (BYVAL fOption AS DWORD) AS CHARFORMATW
      DECLARE PROPERTY CharFormat (BYVAL chfmt AS DWORD, BYREF cf AS CHARFORMATW)
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
      DECLARE PROPERTY EventMask () AS DWORD
      DECLARE PROPERTY EventMask (BYVAL fMask AS LONG)
      DECLARE PROPERTY HyphenateInfo () AS .HYPHENATEINFO
      DECLARE PROPERTY HyphenateInfo (BYREF hinfo AS .HYPHENATEINFO)
      DECLARE PROPERTY IMEModeBias () AS DWORD
      DECLARE PROPERTY IMEModeBias (BYVAL nModeBias AS LONG)
      DECLARE PROPERTY IMEOptions () AS DWORD
      DECLARE PROPERTY IMEOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
      DECLARE PROPERTY LangOptions () AS DWORD
      DECLARE PROPERTY LangOptions (BYVAL lgoptions AS LONG)
      DECLARE PROPERTY LimitText () AS LONG
      DECLARE PROPERTY LimitText (BYVAL chMax AS DWORD)
      DECLARE PROPERTY Modify () AS LONG
      DECLARE PROPERTY Modify (BYVAL fModify AS LONG)
      DECLARE PROPERTY Options () AS DWORD
      DECLARE PROPERTY Options (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
      DECLARE PROPERTY PageRotate () AS DWORD
      DECLARE PROPERTY PageRotate (BYVAL txtlayout AS LONG)
      DECLARE PROPERTY ParaFormat () AS DWORD
      DECLARE PROPERTY ParaFormat (BYREF pfmt AS .PARAFORMAT)
      DECLARE PROPERTY PasswordChar () AS LONG
      DECLARE PROPERTY PasswordChar (BYVAL dwchar AS DWORD)
      DECLARE PROPERTY Punctuation (BYVAL punctp AS DWORD) AS .PUNCTUATION
      DECLARE PROPERTY Punctuation (BYVAL ptype AS LONG, BYREF punct AS .PUNCTUATION)
      DECLARE PROPERTY Rect () AS .RECT
      DECLARE PROPERTY Rect (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
      DECLARE PROPERTY RectNP (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
      DECLARE PROPERTY ScrollPos () AS .POINT
      DECLARE PROPERTY ScrollPos (BYREF pt AS .POINT)
      DECLARE PROPERTY StoryType (BYVAL Index AS DWORD) AS DWORD
      DECLARE PROPERTY StoryType (BYVAL Index AS LONG, BYVAL dwType AS DWORD)
      DECLARE PROPERTY Text () AS CWSTR
      DECLARE PROPERTY Text (BYREF wszText AS WSTRING)
      DECLARE PROPERTY TextMode () AS DWORD
      DECLARE PROPERTY TextMode (BYVAL pvalues AS LONG)
      DECLARE PROPERTY TouchOptions (BYVAL _Options AS LONG PTR) AS DWORD
      DECLARE PROPERTY TouchOptions (BYVAL _Options AS LONG, BYVAL fEnable AS LONG)
      DECLARE PROPERTY TypographyOptions () AS DWORD
      DECLARE PROPERTY TypographyOptions (BYVAL pto AS LONG, BYVAL fMask AS LONG)

      ' // Procedures
      DECLARE FUNCTION CallAutocorrectProc (BYVAL char AS WCHAR) AS LONG
      DECLARE FUNCTION CanPaste (BYVAL clipformat AS LONG) AS BOOLEAN
      DECLARE FUNCTION CanRedo () AS BOOLEAN
      DECLARE FUNCTION CanUndo () AS BOOLEAN
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
      DECLARE FUNCTION GetCharFromPos (BYREF ptl AS .POINTL) AS LONG
      DECLARE FUNCTION GetEllipsisState () AS BOOLEAN
      DECLARE FUNCTION GetFirstVisibleLine () AS LONG
      DECLARE FUNCTION GetIMEColor (BYVAL rgCmpclr AS .COMPCOLOR PTR) AS LONG
      DECLARE FUNCTION GetIMECompMode () AS DWORD
      DECLARE FUNCTION GetIMECompText (BYREF ict AS .IMECOMPTEXT, BYVAL buffer AS ANY PTR) AS DWORD
      DECLARE FUNCTION GetIMEProperty (BYVAL figp AS DWORD) AS DWORD
      DECLARE FUNCTION GetLine (BYVAL which AS DWORD) AS CWSTR
      DECLARE FUNCTION GetLineCount () AS LONG
      DECLARE FUNCTION GetOleInterface () AS IRichEditOle PTR
      DECLARE FUNCTION GetRedoName () AS LONG
      DECLARE FUNCTION GetSel (BYREF dwStartPos AS DWORD, BYREF dwEndPos AS DWORD) AS LONG
      DECLARE FUNCTION GetSelText () AS CWSTR
      DECLARE FUNCTION GetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
      DECLARE FUNCTION GetTextEx (BYREF gtex AS GETTEXTEX, BYVAL buffer AS ANY PTR) AS DWORD
      DECLARE FUNCTION GetTextLength () AS LONG
      DECLARE FUNCTION GetTextLengthEx (BYREF gtex AS .GETTEXTLENGTHEX) AS LONG
      DECLARE FUNCTION GetTextRange OVERLOAD (BYREF trg AS .TEXTRANGEW) AS DWORD
      DECLARE FUNCTION GetTextRange OVERLOAD (BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1) AS CWSTR
      DECLARE FUNCTION GetThumb () AS LONG
      DECLARE FUNCTION GetUndoName () AS DWORD
      DECLARE FUNCTION GetWordBreakProc () AS LONG_PTR
      DECLARE FUNCTION GetWordBreakProcEx () AS LONG_PTR
      DECLARE FUNCTION GetWordWrapMode () AS DWORD
      DECLARE FUNCTION GetZoom (BYREF pzNum AS DWORD, BYREF pzDen AS DWORD) AS LONG
      DECLARE SUB HideSelection (BYVAL fHide AS DWORD)
      DECLARE FUNCTION InsertImage (BYREF ip AS RICHEDIT_IMAGE_PARAMETERS) AS DWORD
      DECLARE FUNCTION InsertTable (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
      DECLARE FUNCTION IsIME () AS LONG
      DECLARE FUNCTION LineFromChar (BYVAL index AS DWORD) AS LONG
      DECLARE FUNCTION LineIndex (BYVAL nLine AS LONG) AS LONG
      DECLARE FUNCTION LineLength (BYVAL index AS DWORD) AS LONG
      DECLARE FUNCTION LineScroll (BYVAL y AS LONG) AS LONG
      DECLARE SUB PasteSpecial OVERLOAD (BYVAL clpfmt AS DWORD, BYREF rps AS REPASTESPECIAL)
      DECLARE SUB PasteSpecial OVERLOAD (BYVAL clpfmt AS DWORD, BYVAL dwAspect AS DWORD, BYVAL dwParam AS DWORD)
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
      DECLARE FUNCTION SetIMEColor (BYVAL pcompcolor AS .COMPCOLOR PTR) AS LONG
      DECLARE SUB SetMargins (BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
      DECLARE FUNCTION SetOleCallback (BYVAL pCallback AS ANY PTR) AS LONG
      DECLARE SUB SetPalette (BYVAL newPalette AS HPALETTE)
      DECLARE FUNCTION SetReadOnly (BYVAL fReadOnly AS LONG) AS LONG
      DECLARE SUB SetSel (BYVAL nStart AS LONG, BYVAL nEnd AS LONG)
      DECLARE FUNCTION SetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
      DECLARE FUNCTION SetTabStops (BYVAL nTabs AS LONG, BYVAL rgTabStops AS LONG_PTR) AS LONG
      DECLARE FUNCTION SetTargetDevice (BYVAL hDC AS HDC, BYVAL lnwidth AS LONG) AS LONG
      DECLARE FUNCTION SetTextEx (BYREF stex AS .SETTEXTEX, BYREF wszText AS WSTRING) AS DWORD
      DECLARE FUNCTION SetUIAName (BYREF wszName AS WSTRING) AS DWORD
      DECLARE FUNCTION SetUndoLimit (BYVAL maxactions AS DWORD) AS DWORD
      DECLARE SUB SetWordBreakProc (BYVAL pfn AS LONG_PTR)
      DECLARE FUNCTION SetWordBreakProcEx (BYVAL pfn AS LONG_PTR) AS LONG_PTR
      DECLARE FUNCTION SetWordWrapMode (BYVAL pvalues AS LONG) AS LONG
      DECLARE FUNCTION SetZoom (BYVAL zNum AS DWORD, BYVAL zDen AS DWORD) AS LONG
      DECLARE SUB ShowScrollBar (BYVAL nScrollBar AS DWORD, BYVAL fShow AS LONG)
      DECLARE FUNCTION StopGroupTyping () AS DWORD
      DECLARE FUNCTION StreamIn (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
      DECLARE FUNCTION StreamOut (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
      DECLARE FUNCTION Undo () AS LONG

      ' // Helper functions
      DECLARE FUNCTION SetFont (BYREF wszFaceName AS WSTRING, BYVAL ptsize AS LONG) AS HRESULT
      DECLARE FUNCTION GetRtfText () AS STRING
      DECLARE FUNCTION LoadRtfFromFile (BYREF wszFileName AS WSTRING) AS BOOLEAN
      DECLARE FUNCTION LoadRtfFromResource (BYREF wszResourceName AS WSTRING) AS BOOLEAN

END TYPE

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
               IF HIWORD(.GetKeyState(VK_SHIFT)) = 0 THEN
                  '// Set the focus in the previous control that has the WS_TABSTOP style
                  .SetFocus .GetNextDlgTabItem(GetParent(hwnd), hwnd, 0)
               ELSE
                  '// Set the focus in the next control that has the WS_TABSTOP style
                  .SetFocus .GetNextDlgTabItem(GetParent(hwnd), hwnd, -1)
               END IF
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


' ########################################################################################
' Implemented procedures and properties
' ########################################################################################

' // Properties

' ========================================================================================
' Gets a pointer to the application-defined AutoCorrectProc callback function.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.AutoCorrectProc () AS LONG_PTR
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETAUTOCORRECTPROC, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets a pointer to the application-defined AutoCorrectProc callback procedure.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.AutoCorrectProc (BYVAL pfn AS LONG_PTR)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETAUTOCORRECTPROC, 0, cast(LPARAM, pfn)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Indicates whether the auto URL detection is turned on in the rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.AutoUrlDetect () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETAUTOURLDETECT, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Enables or disables automatic detection of URLs by a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.AutoUrlDetect (BYVAL fUrlDetect AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_AUTOURLDETECT, fUrlDetect, 0))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Indicates the current state of the bidirectional options in the rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.BidiOptions () AS .BIDIOPTIONS
   DIM bo AS .BIDIOPTIONS
   bo.cbSize = SIZEOF(.BIDIOPTIONS)
   this.SetResult(SendMessageW(m_hRichEdit, EM_GETBIDIOPTIONS, 0, cast(LPARAM, @bo)))
   RETURN bo
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the current state of the bidirectional options in the rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.BidiOptions (BYREF _Options AS .BIDIOPTIONS)
   IF _Options.cbSize = 0 THEN _Options.cbSize = SIZEOF(.BIDIOPTIONS)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETBIDIOPTIONS, 0, cast(LPARAM, @_Options)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Determines the current character formatting in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CharFormat (BYVAL fOption AS DWORD) AS CHARFORMATW
   DIM cf AS CHARFORMATW
   cf.cbSize = SIZEOF (cf)
   this.SetResult(SendMessageW(m_hRichEdit, EM_GETCHARFORMAT, fOption, cast(LPARAM, @cf)))
   RETURN cf
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets character formatting in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CharFormat (BYVAL chfmt AS DWORD, BYREF cf AS CHARFORMATW)
   IF cf.cbSize = 0 THEN cf.cbSize = SIZEOF (cf)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETCHARFORMAT, chfmt, cast (LPARAM, @cf)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the Text Services Framework mode bias values for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CTFModeBias () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETCTFMODEBIAS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the Text Services Framework (TSF) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CTFModeBias (BYVAL nModeBias AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETCTFMODEBIAS, nModeBias, 0))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Determines if the Text Services Framework (TSF) keyboard is open or closed.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CTFOpenStatus () AS BOOLEAN
   this.SetResult(0)
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_GETCTFOPENSTATUS, 0, 0))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Opens or closes the Text Services Framework (TSF) keyboard.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.CTFOpenStatus (BYVAL fTSFkbd AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETCTFOPENSTATUS, fTSFkbd, 0))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Retrieves the current edit style flags.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EditStyle () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETEDITSTYLE, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the current edit style flags.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EditStyle (BYVAL fStyle AS LONG, BYVAL fMask AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETEDITSTYLE, fStyle, fMask))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the extended edit style flags.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EditStyleEx () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETEDITSTYLEEX, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the current extended edit style flags.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EditStyleEx (BYVAL fStyle AS LONG, BYVAL fMask AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETEDITSTYLEEX, fStyle, fMask))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Retrieves the current ellipsis mode.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EllipsisMode () AS DWORD
   DIM pmode AS DWORD
   this.SetResult(SendMessageW(m_hRichEdit, EM_GETELLIPSISMODE, 0, cast(LPARAM, @pMode)))
   RETURN pmode
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the current ellipsis mode. 
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EllipsisMode (BYVAL fMode AS DWORD)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETELLIPSISMODE, 0, fMode))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Retrieves the event mask for a rich edit control. The event mask specifies which
' notification messages the control sends to its parent window.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EventMask () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETEVENTMASK, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the event mask for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.EventMask (BYVAL fMask AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETEVENTMASK, 0, fMask))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets information about hyphenation for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.HyphenateInfo () AS .HYPHENATEINFO
   DIM info AS .HYPHENATEINFO
   info.cbSize = SIZEOF(info)
   this.SetResult(SendMessageW(m_hRichEdit, EM_GETHYPHENATEINFO, cast(WPARAM, @info), 0))
   RETURN info
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the way a Microsoft Rich Edit control does hyphenation.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.HyphenateInfo (BYREF info AS .HYPHENATEINFO)
   IF info.cbSize = 0 THEN info.cbSize = SIZEOF(info)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETHYPHENATEINFO, cast(WPARAM, @info), 0))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.IMEModeBias () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETIMEMODEBIAS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.IMEModeBias (BYVAL nModeBias AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit,EM_SETIMEMODEBIAS, nModeBias, nModeBias))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Retrieves the current Input Method Editor (IME) options. This message is available only
' in Asian-language versions of the operating system.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.IMEOptions () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETIMEOPTIONS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the Input Method Editor (IME) options.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.IMEOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETIMEOPTIONS, fCoop, fOptions))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets a rich edit control's option settings for Input Method Editor (IME) and Asian
' language support.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.LangOptions () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETLANGOPTIONS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets options for Input Method Editor (IME) and Asian language support in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.LangOptions (BYVAL lgoptions AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETLANGOPTIONS, 0, lgoptions))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the current text limit for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.LimitText () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETLIMITTEXT, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the text limit of a rich edit control. The text limit is the maximum amount of
' text, in TCHARs, that the user can type into the edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.LimitText (BYVAL chMax AS DWORD)
   this.SetResult(SendMessageW(m_hRichEdit, EM_LIMITTEXT, chMax, 0))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the state of a rich edit control's modification flag. The flag indicates whether
' the contents of the rich edit control have been modified.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Modify () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETMODIFY, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets or clears the modification flag for a rich edit control. The modification flag
' indicates whether the text within the rich edit control has been modified.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Modify (BYVAL fModify AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETMODIFY, fModify, 0))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Retrieves the options for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Options () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETOPTIONS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the options for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Options (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETOPTIONS, fCoop, fOptions))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Deprecated. Gets the text layout for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.PageRotate () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETPAGEROTATE, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Deprecated. Sets the text layout for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.PageRotate (BYVAL txtlayout AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETPAGEROTATE, txtlayout, 0))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Retrieves the paragraph formatting of the current selection in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.ParaFormat () AS DWORD
   DIM pfmt AS .PARAFORMAT
   pfmt.cbSize = SIZEOF(.PARAFORMAT)
   this.SetResult(SendMessageW(m_hRichEdit, EM_GETPARAFORMAT, 0, cast(LPARAM, @pfmt)))
   RETURN pfmt.dwMask
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the paragraph formatting for the current selection in a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.ParaFormat (BYREF pfmt AS .PARAFORMAT)
   IF pfmt.cbSize = 0 THEN pfmt.cbSize = SIZEOF(.PARAFORMAT)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETPARAFORMAT, 0, cast(LPARAM, @pfmt)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the password character that a rich edit control displays when the user enters text.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.PasswordChar () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETPASSWORDCHAR, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets or removes the password character for a rich edit control. When a password
' character is set, that character is displayed in place of the characters typed by the user.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.PasswordChar (BYVAL dwchar AS DWORD)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETPASSWORDCHAR, dwchar, 0))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the current punctuation characters for the rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Punctuation (BYVAL punctype AS DWORD) AS .PUNCTUATION
   DIM punct AS .PUNCTUATION
   this.SetResult(SendMessageW(m_hRichEdit, EM_GETPUNCTUATION, punctype, cast(LPARAM, @punct)))
   RETURN punct
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the punctuation characters for a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Punctuation (BYVAL punctype AS LONG, BYREF punct AS .PUNCTUATION)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETPUNCTUATION, punctype, cast(LPARAM, @punct)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the formatting rectangle of a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Rect () AS .RECT
   DIM rc AS .RECT
   this.SetResult(SendMessageW(m_hRichEdit, EM_GETRECT, 0, cast(LPARAM, @rc)))
   RETURN rc
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the formatting rectangle of a multiline rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Rect (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETRECT, fCoord, cast(LPARAM, @rc)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the formatting rectangle of a multiline rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.RectNP (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
   this.SetResult(SendMessageW( m_hRichEdit, EM_SETRECTNP, fCoord, cast(LPARAM, @rc)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Obtains the current scroll position of the edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.ScrollPos () AS .POINT
   DIM pt AS .POINT
   this.SetResult(SendMessageW(m_hRichEdit, EM_GETSCROLLPOS, 0, cast(LPARAM, @pt)))
   RETURN pt
END PROPERTY
' ========================================================================================
' ========================================================================================
' Tells the rich edit control to scroll to a particular point.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.ScrollPos (BYREF pt AS .POINT)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETSCROLLPOS, 0, cast(LPARAM, @pt)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the story type.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.StoryType (BYVAL Index AS DWORD) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETSTORYTYPE, Index, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the story type.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.StoryType (BYVAL Index AS LONG, BYVAL dwType AS DWORD)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETSTORYTYPE, Index, dwType))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Retrieves the text from a rich edit control.
' Note: GetWindowText cannot retrieve the text of a control in another application.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.Text () AS CWSTR
   this.SetResult(0)
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
   this.SetResult(SendMessageW(m_hRichEdit, WM_SETTEXT, 0, cast(LPARAM, @wszText)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the current text mode and undo level of a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextMode () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETTEXTMODE, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the text mode or undo level of a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TextMode (BYVAL values AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETTEXTMODE, 0, values))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Retrieves the touch options that are associated with a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TouchOptions (BYVAL _Options AS LONG PTR) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETTOUCHOPTIONS, cast(WPARAM, _Options), 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the touch options associated with a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TouchOptions (BYVAL _Options AS LONG, BYVAL fEnable AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETTOUCHOPTIONS, _Options, fEnable))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the current state of the typography options of a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TypographyOptions () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETTYPOGRAPHYOPTIONS, 0, 0)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the current state of the typography options of a rich edit control.
' ========================================================================================
PRIVATE PROPERTY CRichEditCtrl.TypographyOptions (BYVAL pto AS LONG, BYVAL fMask AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETTYPOGRAPHYOPTIONS, pto, fMask))
END PROPERTY
' ========================================================================================

' // Procedures

' ========================================================================================
' Determines whether a rich edit control can paste a specified clipboard format.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.CanPaste (BYVAL clipformat AS LONG) AS BOOLEAN
   this.SetResult(0)
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_CANPASTE, clipformat, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether there are any actions in the control redo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.CanRedo () AS BOOLEAN
   this.SetResult(0)
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_CANREDO, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether there are any actions in an edit control's undo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.CanUndo () AS BOOLEAN
   this.SetResult(0)
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_CANUNDO, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calls the autocorrect callback function that is stored by the **CRichEditCtrl.SetAutocorrectProc**
' message, provided that the text preceding the insertion point is a candidate for autocorrection.
' ========================================================================================
FUNCTION CRichEditCtrl.CallAutocorrectProc (BYVAL char AS WCHAR) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_CALLAUTOCORRECTPROC, char, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Displays a portion of the contents of a rich edit control, as previously formatted for a
' device using the EM_FORMATRANGE message.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.DisplayBand (BYREF rc AS .RECT) AS BOOLEAN
   this.SetResult(0)
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_DISPLAYBAND, 0, cast(LPARAM, @rc)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets the undo flag of a rich edit control. The undo flag is set whenever an operation
' within the rich edit control can be undone.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.EmptyUndoBuffer ()
   this.SetResult(SendMessageW(m_hRichEdit, EM_EMPTYUNDOBUFFER, 0, 0))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the starting and ending character positions of the selection in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.ExGetSel () AS CHARRANGE
   DIM cr AS CHARRANGE
   this.SetResult(SendMessageW(m_hRichEdit, EM_EXGETSEL, 0, cast(LPARAM, @cr)))
   RETURN cr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ExGetSel (BYREF cpMin AS LONG, BYVAL cpMax AS LONG)
   DIM cr AS CHARRANGE = TYPE<CHARRANGE>(cpMin, cpMax)
   this.SetResult(SendMessageW(m_hRichEdit, EM_EXGETSEL, 0, cast(LPARAM, @cr)))
END SUB
' ========================================================================================

' ========================================================================================
' Sets an upper limit to the amount of text the user can type or paste into a rich edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ExLimitText (BYVAL dwLimit AS DWORD)
   this.SetResult(SendMessageW(m_hRichEdit, EM_EXLIMITTEXT, 0, dwLimit))
END SUB
' ========================================================================================

' ========================================================================================
' Determines which line contains the specified character in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.ExLineFromChar (BYVAL iIndex AS LONG) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_EXLINEFROMCHAR, 0, iIndex)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects a range of characters and/or Component Object Model (COM) objects in a Microsoft
' Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.ExSetSel (BYREF cr AS CHARRANGE) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_EXSETSEL, 0, cast(LPARAM, @cr))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ExSetSel (BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1)
   this.SetResult(0)
   DIM cr AS CHARRANGE = TYPE<CHARRANGE>(cpMin, cpMax)
   this.SetResult(SendMessageW(m_hRichEdit, EM_EXSETSEL, 0, cast(LPARAM, @cr)))
END SUB
' ========================================================================================

' ========================================================================================
' Finds text within a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FindText (BYVAL fOptions AS DWORD, BYREF ftw AS FINDTEXTW) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_FINDTEXTW, fOptions, cast(LPARAM, @ftw))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FindText (BYVAL fOptions AS DWORD = FR_DOWN, _
   BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1, BYREF wszText AS WSTRING) AS LONG
   this.SetResult(0)
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
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_FINDTEXTEXW, fOptions, cast(LPARAM, @ftexw))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FindTextEx (BYVAL fOptions AS DWORD = FR_DOWN, _
   BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1, BYREF wszText AS WSTRING) AS CHARRANGE
   this.SetResult(0)
   DIM ftexw AS FINDTEXTEXW
   ftexw.chrg.cpMin = cpMin
   ftexw.chrg.cpMax = cpMax
   ftexw.lpstrText = @wszText
   this.SetResult(SendMessageW(m_hRichEdit, EM_FINDTEXTEXW, fOptions, cast(LPARAM, @ftexw)))
   RETURN ftexw.chrgText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds the next word break before or after the specified character position or retrieves
' information about the character at that position.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FindWordBreak (BYVAL fOperation AS DWORD, BYVAL dwStartPos AS DWORD) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_FINDWORDBREAK, fOperation, dwStartPos)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Formats a range of text in a rich edit control for a specific device.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.FormatRange (BYVAL fRender AS LONG, BYREF fr AS .FORMATRANGE) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_FORMATRANGE, fRender, cast(LPARAM, @fr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about the character closest to a specified point in the client area of
' an edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetCharFromPos (BYREF ptl AS .POINTL) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_CHARFROMPOS, 0, cast(LPARAM, @ptl))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current ellipsis state.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetEllipsisState () AS BOOLEAN
   this.SetResult(0)
   RETURN CBOOL(SendMessageW(m_hRichEdit, EM_GETELLIPSISSTATE, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the zero-based index of the uppermost visible line in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetFirstVisibleLine () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETFIRSTVISIBLELINE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current IME mode for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetIMECompMode () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETIMECOMPMODE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the Input Method Editor (IME) composition color. This message is available
' only in Asian-language versions of the operating system.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetIMEColor (BYVAL rgCmpclr AS .COMPCOLOR PTR) AS LONG
   RETURN this.SetResult(SendMessageW(m_hRichEdit, EM_GETIMECOLOR, 0, cast(LPARAM, rgCmpclr)))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the Input Method Editor (IME) composition color.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetIMEColor (BYVAL pcompcolor AS .COMPCOLOR PTR) AS LONG
   RETURN this.SetResult(SendMessageW(m_hRichEdit, EM_SETIMECOLOR, 0, cast(LPARAM, pcompcolor)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the property and capabilities of the Input Method Editor (IME) associated with the
' current input locale.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetIMEProperty (BYVAL figp AS DWORD) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETIMEPROPERTY, figp, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies a line of text from a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetLine (BYVAL which AS DWORD) AS CWSTR
   this.SetResult(0)
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
   this.SetResult(SendMessageW(m_hRichEdit, EM_GETOLEINTERFACE, 0, cast(LPARAM, @ppObject)))
   RETURN ppObject
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the type of the next action, if any, in the control's redo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetRedoName () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETREDONAME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the starting and ending character positions of the current selection in a rich
' edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetSel (BYREF dwStartPos AS DWORD, BYREF dwEndPos AS DWORD) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETSEL, cast(WPARAM, @dwStartPos), cast(LPARAM, @dwEndPos))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Selects a range of characters in a rich edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetSel (BYVAL nStart AS LONG, BYVAL nEnd AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETSEL, nStart, nEnd))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the currently selected text in a rich edit control.
' Note: From version 4.1 (RichEdit50W class in MSFTEDIT.DLL), Windows XP, EM_GETSELTEXT
' returns the text in unicode.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetSelText () AS CWSTR
   this.SetResult(0)
   DIM dwStartPos AS DWORD, dwEndPos AS DWORD, cr AS CHARRANGE
   SendMessageW(m_hRichEdit, EM_EXGETSEL, 0, cast(LPARAM, @cr))
   DIM cbLen AS DWORD = ABS(cr.cpMax - cr.cpMin)
   IF cbLen < 1 THEN RETURN ""
   DIM cwsText AS CWSTR = cbLen + 1
   cbLen = SendMessageW(m_hRichEdit, EM_GETSELTEXT, 0, cast(LPARAM, *cwsText))
   RETURN LEFT(**cwsText, cbLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the type of the next undo action, if any.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetUndoName () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETUNDONAME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current word wrap and word-break options for the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetWordWrapMode () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETWORDWRAPMODE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the word-wrapping and word-breaking options for the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetWordWrapMode (BYVAL pvalues AS LONG) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETWORDWRAPMODE, pvalues, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current zoom ratio, which is always between 1/64 and 64.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetZoom (BYREF pzNum AS DWORD, BYREF pzDen AS DWORD) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETZOOM, cast(WPARAM, @pzNum), cast(LPARAM, @pzDen))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the zoom ratio anywhere between 1/64 and 64.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetZoom (BYVAL zNum AS DWORD, BYVAL zDen AS DWORD) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETZOOM, zNum, zDen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the table parameters for a table row and the cell parameters for the specified number of cells.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETTABLEPARMS, cast(WPARAM, @tp), cast(LPARAM, @tcp))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Changes the parameters of rows in a table.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETTABLEPARMS, cast(WPARAM, @tp), cast(LPARAM, @tcp))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets all of the text from the rich edit control in any particular code base you want.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextEx (BYREF gtex AS .GETTEXTEX, BYVAL buffer AS ANY PTR) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETTEXTEX, cast(WPARAM, @gtex), cast(LPARAM, buffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Combines the functionality of WM_SETTEXT and EM_REPLACESEL and adds the ability to set
' text using a code page and to use either Rich Text Format (RTF) rich text or plain text.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTextEx (BYREF stex AS .SETTEXTEX, BYREF wszText AS WSTRING) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETTEXTEX, cast(WPARAM, @stex), cast(LPARAM, @wszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline
' rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetThumb () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETTHUMB, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the address of the current Wordwrap function.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetWordBreakProc () AS LONG_PTR
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETWORDBREAKPROC, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Replaces a rich edit control's default Wordwrap function with an application-defined Wordwrap function.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetWordBreakProc (BYVAL pfn AS LONG_PTR)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETWORDBREAKPROC, 0, cast(LPARAM, pfn)))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the address of the currently registered extended word-break procedure.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetWordBreakProcEx () AS LONG_PTR
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETWORDBREAKPROCEX, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the extended word-break procedure.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetWordBreakProcEx (BYVAL pfn AS LONG_PTR) AS LONG_PTR
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETWORDBREAKPROCEX, 0, cast(LPARAM, pfn))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Hides or shows the selection in a rich edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.HideSelection (BYVAL fHide AS DWORD)
   this.SetResult(SendMessageW(m_hRichEdit, EM_HIDESELECTION, fHide, 0))
END SUB
' ========================================================================================

' ========================================================================================
' Gets the Input Method Editor (IME) composition text.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetIMECompText (BYREF ict AS .IMECOMPTEXT, BYVAL buffer AS ANY PTR) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETIMECOMPTEXT, cast(WPARAM, @ict), cast(LPARAM, buffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces the selection with a blob that displays an image.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.InsertImage (BYREF ip AS RICHEDIT_IMAGE_PARAMETERS) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_INSERTIMAGE, 0, cast(LPARAM, @ip))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts one or more identical table rows with empty cells.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.InsertTable (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_INSERTTABLE, cast(WPARAM, @tp), cast(LPARAM, @tcp))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if current input locale is an East Asian locale.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.IsIME () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_ISIME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of lines in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetLineCount () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETLINECOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the index of the line that contains the specified character index in a multiline
' rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LineFromChar (BYVAL index AS DWORD) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_LINEFROMCHAR, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character index of the first character of a specified line in a multiline
' rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LineIndex (BYVAL nLine AS LONG) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_LINEINDEX, nLine, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length, in characters, of a line in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LineLength (BYVAL index AS DWORD) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_LINELENGTH, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the text in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.LineScroll (BYVAL y AS LONG) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_LINESCROLL, 0, y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Pastes a specific clipboard format in a rich edit control.
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.PasteSpecial (BYVAL clpfmt AS DWORD, BYREF rps AS REPASTESPECIAL)
   this.SetResult(0)
   SendMessageW(m_hRichEdit, EM_PASTESPECIAL, clpfmt, cast(LPARAM, @rps))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CRichEditCtrl.PasteSpecial (BYVAL clpfmt AS DWORD, BYVAL dwAspect AS DWORD, BYVAL dwParam AS DWORD)
   this.SetResult(0)
   DIM rps AS REPASTESPECIAL= TYPE<REPASTESPECIAL>(dwAspect, dwParam)
   SendMessageW(m_hRichEdit, EM_PASTESPECIAL, clpfmt, cast(LPARAM, @rps))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the client area coordinates of a specified character in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.PosFromChar (BYVAL index as DWORD) AS .POINTL
   DIM pt AS .POINTL
   this.SetResult(SendMessageW(m_hRichEdit, EM_POSFROMCHAR, cast(WPARAM, @pt), index))
   RETURN pt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Invokes the Input Method Editor (IME) reconversion dialog box.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.Reconversion
   this.SetResult(SendMessageW(m_hRichEdit, EM_RECONVERSION, 0, 0))
END SUB
' ========================================================================================

' ========================================================================================
' Redoes the next action in the control's redo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.Redo () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_REDO, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces the current selection in a rich edit control with the specified text.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ReplaceSel (BYVAL bCanBeUndone AS LONG = TRUE, BYREF wszText AS WSTRING)
   this.SetResult(SendMessageW(m_hRichEdit, EM_REPLACESEL, bCanBeUndone, cast(LPARAM, @wszText)))
END SUB
' ========================================================================================

' ========================================================================================
' Forces a rich edit control to send an EN_REQUESTRESIZE notification message to its
' parent window.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.RequestResize ()
   this.SetResult(SendMessageW(m_hRichEdit, EM_REQUESTRESIZE, 0, 0))
END SUB
' ========================================================================================

' ========================================================================================
' Scrolls the text vertically in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.Scroll (BYVAL nAction AS LONG) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SCROLL, nAction, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the caret into view in a rich edit control.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ScrollCaret ()
   this.SetResult(SendMessageW(m_hRichEdit, EM_SCROLLCARET, 0, 0))
END SUB
' ========================================================================================

' ========================================================================================
' Determines the selection type for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SelectionType () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SELECTIONTYPE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetBkgndColor (BYVAL SysColor AS DWORD, BYVAL BkColor AS DWORD) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETBKGNDCOLOR, SysColor, BkColor)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font size for the selected text.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetFontSize (BYVAL ptsize AS LONG) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETFONTSIZE, ptsize, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the widths of the left and right margins for a rich edit control. The message
' redraws the control to reflect the new margins.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetMargins (BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETMARGINS, nMargins, nWidth))
END SUB
' ========================================================================================

' ========================================================================================
' Gives a rich edit control an IRichEditOleCallback object that the control uses to get
' OLE-related resources and information from the client.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetOleCallback (BYVAL pCallback AS ANY PTR) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETOLECALLBACK, 0, cast(LPARAM, pCallback))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the palette that a rich edit control uses for its display window.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.SetPalette (BYVAL newPalette AS HPALETTE)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SETPALETTE, cast(WPARAM, newPalette), 0))
END SUB
' ========================================================================================

' ========================================================================================
' Sets or removes the read-only style (ES_READONLY) of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetReadOnly (BYVAL fReadOnly AS LONG) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETREADONLY, fReadOnly, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the tab stops in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTabStops (BYVAL nTabs AS LONG, BYVAL rgTabStops AS LONG_PTR) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETTABSTOPS, nTabs, cast(LPARAM, rgTabStops))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the target device and line width used for WYSIWYG formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetTargetDevice (BYVAL hDC AS HDC, BYVAL lnwidth AS LONG) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETTARGETDEVICE, cast(WPARAM, hDC), lnwidth)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the name of a rich edit control for UI Automation (UIA).
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetUIAName (BYREF wszName AS WSTRING) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETUIANAME, 0, cast(LPARAM, @wszName))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the maximum number of actions that can stored in the undo queue.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.SetUndoLimit (BYVAL maxactions AS DWORD) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_SETUNDOLIMIT, maxactions, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Shows or hides one of the scroll bars in the Text Host window.
' ========================================================================================
PRIVATE SUB CRichEditCtrl.ShowScrollBar (BYVAL nScrollBar AS DWORD, BYVAL fShow AS LONG)
   this.SetResult(SendMessageW(m_hRichEdit, EM_SHOWSCROLLBAR, nScrollBar, fShow))
END SUB
' ========================================================================================

' ========================================================================================
' Stops the control from collecting additional typing actions into the current undo action.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.StopGroupTyping () AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_STOPGROUPTYPING, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces the contents of a rich edit control with a stream of data provided by an
' application defined EditStreamCallback callback function.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.StreamIn (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_STREAMIN, psf, cast(LPARAM, @edst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Causes a rich edit control to pass its contents to an application defined
' EditStreamCallback callback function.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.StreamOut (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_STREAMOUT, psf, cast(LPARAM, @edst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length of all text in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextLength () AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, WM_GETTEXTLENGTH, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates text length in various ways. It is usually called before creating a buffer to
' receive the text from the control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextLengthEx (BYREF gtex AS .GETTEXTLENGTHEX) AS LONG
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETTEXTLENGTHEX, cast(WPARAM, @gtex), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a specified range of characters from a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextRange (BYREF trg AS .TEXTRANGEW) AS DWORD
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_GETTEXTRANGE, 0, cast(LPARAM, @trg))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves a specified range of characters from a rich edit control.
' If the cpMin and cpMax members are equal, the range is empty.
' The range includes everything if cpMin is 0 and cpMax is –1.
' ========================================================================================
PRIVATE FUNCTION CRichEditCtrl.GetTextRange (BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1) AS CWSTR
   this.SetResult(0)
   ' // Calculate the length of the buffer
   DIM cbLen As LONG
   DIM wstrText AS CWSTR
   IF cpMin = cpMax THEN RETURN ""
   IF cpMin = 0 AND cpMax = -1 THEN
      cbLen = SendMessageW(m_hRichEdit, WM_GETTEXTLENGTH, 0, 0)
      IF cbLen = 0 THEN RETURN ""
   ELSE
      cbLen = cpMax - cpMin
      wstrText = WSPACE(cbLen + 1)
   END IF
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
   this.SetResult(0)
   RETURN SendMessageW(m_hRichEdit, EM_UNDO, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Callback used by the GetRtfText method.
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
PRIVATE FUNCTION CRichEditCtrl.GetRtfText () AS STRING

   DIM eds AS EDITSTREAM, cws AS CWSTR
   eds.dwCookie = cast(DWORD_PTR, @cws)
   eds.pfnCallBack = cast(EDITSTREAMCALLBACK, @CRichEditCtrl.CRichEditCtrlGetTextCallback)
   this.SetResult(SendMessageW(m_hRichEdit, EM_STREAMOUT, SF_RTF, cast(LPARAM, @eds)))
   ' // Copy the ansi contents of the CWSTR to a STRING
   DIM s AS STRING = SPACE(cws.m_BufferLen - 2)   ' // -2 to remove the ending nulls
   IF LEN(s) THEN CopyMemory(STRPTR(s), cws.m_pBuffer, cws.m_BufferLen - 2)
   RETURN s

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
   hResult = SendMessageW(hRichEdit, EM_SETCHARFORMAT, SCF_ALL, cast(LPARAM, @tcf))
   ' // Specify which notifications the control sends to its parent window
   IF hResult <> 0 THEN hResult = SendMessageW(m_hRichEdit, EM_SETEVENTMASK, 0, ENM_CHANGE)
   this.SetResult(hResult)
   RETURN hResult

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
   IF m_hRichEdit = 0 THEN this.SetResult(E_HANDLE) : RETURN FALSE
   IF LEN(wszFileName) = 0 THEN this.SetResult(E_INVALIDARG) : RETURN FALSE
   this.SetResult(0)

   ' // Opens the file and sends the message
   hFile = CreateFileW(wszFileName, GENERIC_READ, FILE_SHARE_READ, _
                       NULL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN SetResult(E_HANDLE): RETURN FALSE
   eds.dwCookie = cast(DWORD_PTR, hFile)
   eds.pfnCallback = cast(EDITSTREAMCALLBACK, @CRichEditCtrlLoadRtfFromFileCallback)
   IF SendMessageW(hRichEdit, EM_STREAMIN, SF_RTF, cast(LPARAM, @eds)) > 0 AND eds.dwError = 0 THEN FUNCTION = TRUE
   this.SetResult(eds.dwError)
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
   IF m_hRichEdit = 0 THEN this.SetResult(E_HANDLE) : RETURN FALSE
   IF LEN(wszResourceName) = 0 THEN this.SetResult(E_INVALIDARG) : RETURN FALSE
   this.SetResult(0)

   ' // Loads the resource
   hResInfo = FindResourceW(m_hInstance, wszResourceName, RT_RCDATA)
   IF hResInfo = NULL THEN this.SetResult(GetLastError) : RETURN FALSE

   ' // Loads and locks the resource
   ' // Note  LockResource does not actually lock memory; it is just used to obtain
   ' // a pointer to the memory containing the resource data.
   pResData = LockResource(LoadResource(m_hInstance, hResInfo))
   IF pResData = NULL THEN this.SetResult(GetLastError) : RETURN FALSE
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
      this.SetResult(eds.dwError)
      RETURN FALSE
   END IF

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
