' ########################################################################################
' Microsoft Windows
' File: CTOMDocument.inc
' Contents: Class to work with the Richd Edit Text Object Model (TOM).
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' // To be included by CTOM.inc

' ########################################################################################
' CTextDocument2 class
' ########################################################################################
TYPE CTextDocument2 EXTENDS CTOMBase

Public:
   m_Result AS HRESULT
   m_pTextDocument2 AS ITextDocument2 PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL hRichEdit AS HWND)
   DECLARE CONSTRUCTOR (BYVAL pTextDocument2 AS ITextDocument2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR LET (BYVAL pTextDocument2 AS ITextDocument2 PTR)
   DECLARE OPERATOR CAST () AS ITextDocument2 PTR
   DECLARE FUNCTION TextDocumentPtr () AS ITextDocument2 PTR
   DECLARE FUNCTION Attach (BYVAL pTextDocument2 AS ITextDocument2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Detach () AS ITextDocument2 PTR

   ' // ITextDocument2 methods
   DECLARE FUNCTION GetName () AS CBSTR
   DECLARE FUNCTION GetSelection () AS ITextSelection PTR
   DECLARE FUNCTION GetStoryCount () AS LONG
   DECLARE FUNCTION GetStoryRanges () AS ITextStoryRanges PTR
   DECLARE FUNCTION GetSaved () AS LONG
   DECLARE FUNCTION SetSaved (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetDefaultTabStop () AS SINGLE
   DECLARE FUNCTION SetDefaultTabStop (BYVAL Value AS SINGLE) AS HRESULT
   DECLARE FUNCTION New_ () AS HRESULT
   DECLARE FUNCTION Open (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   DECLARE FUNCTION Save (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   DECLARE FUNCTION Freeze () AS LONG
   DECLARE FUNCTION Unfreeze () AS LONG
   DECLARE FUNCTION BeginEditCollection () AS HRESULT
   DECLARE FUNCTION EndEditCollection () AS HRESULT
   DECLARE FUNCTION Undo (BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION Redo (BYVAL Count AS LONG) AS LONG
   DECLARE FUNCTION Range (BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG) AS ITextRange PTR
   DECLARE FUNCTION RangeFromPoint (BYVAL x AS LONG, BYVAL y AS LONG) AS ITextRange PTR
   DECLARE FUNCTION GetCaretType () AS LONG
   DECLARE FUNCTION SetCaretType (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetDisplays () AS ITextDisplays PTR
   DECLARE FUNCTION GetDocumentFont () AS ITextFont2 PTR
   DECLARE FUNCTION SetDocumentFont (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   DECLARE FUNCTION GetDocumentPara () AS ITextPara2 PTR
   DECLARE FUNCTION SetDocumentPara (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   DECLARE FUNCTION GetEastAsianFlags () AS LONG
   DECLARE FUNCTION GetGenerator () AS CBSTR
   DECLARE FUNCTION SetIMEInProgress (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetNotificationMode () AS LONG
   DECLARE FUNCTION SetNotificationMode (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetSelection2 () AS ITextSelection2 PTR
   DECLARE FUNCTION GetStoryRanges2 () AS ITextStoryRanges2 PTR
   DECLARE FUNCTION GetTypographyOptions () AS LONG
   DECLARE FUNCTION GetVersion () AS LONG
   DECLARE FUNCTION GetWindow () AS __int64
   DECLARE FUNCTION AttachMsgFilter (BYVAL pFilter AS IUnknown PTR) AS HRESULT
   DECLARE FUNCTION CheckTextLimit (BYVAL cch AS LONG) AS LONG
   DECLARE FUNCTION GetCallManager () AS IUnknown PTR
   DECLARE FUNCTION GetClientRect (BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, _
      BYVAL pTop AS LONG PTR, BYVAL pRight AS LONG PTR, BYVAL pBottom AS LONG PTR) AS HRESULT
   DECLARE FUNCTION GetEffectColor (BYVAL Index AS LONG) AS ULONG
   DECLARE FUNCTION GetImmContext () AS __int64
   DECLARE FUNCTION GetPreferredFont (BYVAL cp AS LONG, BYVAL CodePage AS LONG, _
      BYVAL Options AS LONG, BYVAL curCodepage AS LONG, BYVAL curFontSize AS LONG, BYVAL pbstr AS BSTR PTR, _
      BYVAL pPitchAndFamily AS LONG PTR, BYVAL pNewFontSize AS LONG PTR) AS HRESULT
   DECLARE FUNCTION GetProperty (BYVAL nType AS LONG) AS LONG
   DECLARE FUNCTION GetStrings () AS ITextStrings PTR
   DECLARE FUNCTION Notify (BYVAL nNotify AS LONG) AS HRESULT
   DECLARE FUNCTION Range2 (BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG) AS ITextRange2 PTR
   DECLARE FUNCTION RangeFromPoint2 (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG) AS ITextRange2 PTR
   DECLARE FUNCTION ReleaseCallManager (BYVAL pVoid AS IUnknown PTR) AS HRESULT
   DECLARE FUNCTION ReleaseImmContext (BYVAL Context AS __int64) AS HRESULT
   DECLARE FUNCTION SetEffectColor (BYVAL Index AS LONG, BYVAL Value AS ULONG) AS HRESULT
   DECLARE FUNCTION SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION SetTypographyOptions (BYVAL Options AS LONG, BYVAL Mask AS LONG) AS HRESULT
   DECLARE FUNCTION SysBeep () AS HRESULT
   DECLARE FUNCTION Update (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION UpdateWindow () AS HRESULT
   DECLARE FUNCTION GetMathProperties () AS LONG
   DECLARE FUNCTION SetMathProperties (BYVAL Options AS LONG, BYVAL Mask AS LONG) AS HRESULT
   DECLARE FUNCTION GetActiveStory () AS ITextStory PTR
   DECLARE FUNCTION SetActiveStory (BYVAL pStory AS ITextStory PTR) AS HRESULT
   DECLARE FUNCTION GetMainStory () AS ITextStory PTR
   DECLARE FUNCTION GetNewStory () AS ITextStory PTR
   DECLARE FUNCTION GetStory (BYVAl Index AS LONG) AS ITextStory PTR

END TYPE
' ========================================================================================

' ========================================================================================
' CTextDocument constructor
' Called when a class variable is created.
' ========================================================================================
PRIVATE CONSTRUCTOR CTextDocument2
   CTOM_DP("CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTextDocument constructor
' Called when a class variable is created.
' ========================================================================================
PRIVATE CONSTRUCTOR CTextDocument2 (BYVAL hRichEdit AS HWND)
   CTOM_DP("BEGIN CTextDocument2 CONSTRUCTOR - hRichEdit " & WSTR(hRichEdit))
   IF hRichEdit = 0 THEN EXIT CONSTRUCTOR
   ' // Retrieve a pointer to a IRichEditOle object of the Rich Edit control
   DIM pUnk AS IUnknown PTR
   m_Result = SendMessageW(hRichEdit, EM_GETOLEINTERFACE, 0, cast(LPARAM, @pUnk))
   ' // Retrieve a pointer to its ITextDocument2 interface
   IF pUnk THEN
      DIM IID_ITextDocument2_ AS IID = AfxGuid(AFX_IID_ITextDocument2)
      m_Result = IUnknown_QueryInterface(pUnk, @IID_ITextDocument2_, @m_pTextDocument2)
      IUnknown_Release(pUnk)
   END IF
   CTOM_DP("END CTextDocument2 CONSTRUCTOR - m_pTextDocument2 = " & WSTR(m_pTextDocument2))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTextDocument constructor
' Called when a class variable is created.
' ========================================================================================
PRIVATE CONSTRUCTOR CTextDocument2 (BYVAL pTextDocument2 AS ITextDocument2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CTOM_DP("BEGIN CTextDocument2 CONSTRUCTOR - ITextDocument2 PTR = " & WSTR(pTextDocument2))
   IF fAddRef THEN pTextDocument2->lpvtbl->AddRef(pTextDocument2)
   m_pTextDocument2 = pTextDocument2
   CTOM_DP("END CTextDocument2 CONSTRUCTOR - m_pTextDocument2 = " & WSTR(m_pTextDocument2))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTextDocument destructor
' Called automatically when a class variable goes out of scope or is destroyed.
' ========================================================================================
PRIVATE DESTRUCTOR CTextDocument2
   CTOM_DP("CTextDocument2 DESTRUCTOR - m_pTextDocument2 = " & WSTR(m_pTextDocument2))
   ' // Release the interface
   DIM dwRefCount AS ULONG
   IF m_pTextDocument2 THEN dwRefCount = m_pTextDocument2->lpvtbl->Release(m_pTextDocument2)
   CTOM_DP("CTextDocument2 DESTRUCTOR - m_pTextDocument2 ref count = " & WSTR(dwRefCount))
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assignment operator
' ========================================================================================
PRIVATE OPERATOR CTextDocument2.LET (BYVAL pTextDocument AS ITextDocument2 PTR)
   m_Result = 0
   IF pTextDocument = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextDocument2 THEN m_pTextDocument2->lpvtbl->Release(m_pTextDocument2)
   ' // Attach the passed interface pointer to the class
   m_pTextDocument2 = pTextDocument
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operator
' ========================================================================================
PRIVATE OPERATOR CTextDocument2.CAST () AS ITextDocument2 PTR
   m_Result = 0
   OPERATOR = m_pTextDocument2
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the underlying ITextDocument2 interface
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.TextDocumentPtr () AS ITextDocument2 PTR
   m_Result = 0
   RETURN m_pTextDocument2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches an ITextDocument2 interface pointer to the class.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Attach (BYVAL pTextDocument AS ITextDocument2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextDocument = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextDocument2 THEN m_Result = m_pTextDocument2->lpvtbl->Release(m_pTextDocument2)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextDocument->lpvtbl->AddRef(pTextDocument)
   m_pTextDocument2 = pTextDocument
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the underlying ITextDocument2 interface pointer from the class
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Detach () AS ITextDocument2 PTR
   m_Result = 0
   DIM pTextDocument AS ITextDocument2 PTR = m_pTextDocument2
   m_pTextDocument2 = NULL
   RETURN pTextDOcument
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the file name of this document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetName () AS CBSTR
   DIM pName AS AFX_BSTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetName(m_pTextDocument2, @pName))
   RETURN pName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the active selection.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetSelection () AS ITextSelection PTR
   DIM pSelection AS ITextSelection PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetSelection(m_pTextDocument2, @pSelection))
   RETURN pSelection
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the count of stories in this document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetStoryCount () AS LONG
   DIM nCount AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetStoryCount(m_pTextDocument2, @nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the story collection object used to enumerate the stories in a document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetStoryRanges () AS ITextStoryRanges PTR
   DIM pStoryRanges AS ITextStoryRanges PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetStoryRanges(m_pTextDocument2, @pStoryRanges))
   RETURN pStoryRanges
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a value that indicates whether changes have been made since the file was last saved.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetSaved () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetSaved(m_pTextDocument2, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the document Saved property.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetSaved (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetSaved(m_pTextDocument2, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the default tab width.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetDefaultTabStop () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextDocument2->lpvtbl->GetDefaultTabStop(m_pTextDocument2, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the default tab stop, which is used when no tab exists beyond the current display position.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetDefaultTabStop (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetDefaultTabStop(m_pTextDocument2, Value))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens a new document. If another document is open, this method saves any current changes
' and closes the current document before opening a new one.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.New_ () AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->New_(m_pTextDocument2))
   FUNCTION = m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens a new document. If another document is open, this method saves any current changes
' and closes the current document before opening a new one.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Open (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->Open(m_pTextDocument2, pVar, Flags, CodePage))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves the document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Save (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->Save(m_pTextDocument2, pVar, Flags, CodePage))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the freeze count. Returns the updated freeze count.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Freeze () AS LONG
   DIM nCount AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->Freeze(m_pTextDocument2, @nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the freeze count. Returns the updated freeze count.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Unfreeze () AS LONG
   DIM nCount AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->Unfreeze(m_pTextDocument2, @nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Turns on edit collection (also called undo grouping).
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.BeginEditCollection () AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->BeginEditCollection(m_pTextDocument2))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Turns off edit collection (also called undo grouping).
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.EndEditCollection () AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->EndEditCollection(m_pTextDocument2))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Performs a specified number of undo operations.
' Returns the actual count of undo operations performed.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Undo (BYVAL Count AS LONG) AS LONG
   DIM nCount AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->Undo(m_pTextDocument2, Count, @nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Performs a specified number of redo operations.
' Returns the actual count of redo operations performed.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Redo (BYVAL Count AS LONG) AS LONG
   DIM nCount AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->Redo(m_pTextDocument2, Count, @nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a text range object for a specified range of content in the active story of the document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Range (BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG) AS ITextRange PTR
   DIM pTextRange AS ITextRange PTR
   this.SetResult(m_pTextDocument2->lpvtbl->Range(m_pTextDocument2, cpActive, cpAnchor, @pTextRange))
   RETURN pTextRange
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a text range object for a specified range of content in the active story of the document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.RangeFromPoint (BYVAL x AS LONG, BYVAL y AS LONG) AS ITextRange PTR
   DIM pTextRange AS ITextRange PTR
   this.SetResult(m_pTextDocument2->lpvtbl->RangeFromPoint(m_pTextDocument2, x, y, @pTextRange))
   RETURN pTextRange
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the caret type.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetCaretType () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetCaretType(m_pTextDocument2, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the caret type.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetCaretType (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetCaretType(m_pTextDocument2, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the displays collection for this Text Object Model (TOM) engine instance.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetDisplays () AS ITextDisplays PTR
   DIM pTextDisplays AS ITextDisplays PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetDisplays(m_pTextDocument2, @pTextDisplays))
   RETURN pTextDisplays
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an object that provides the default character format information for this instance
' of the Text Object Model (TOM) engine.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetDocumentFont () AS ITextFont2 PTR
   DIM pITextFont2 AS ITextFont2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetDocumentFont(m_pTextDocument2, @pITextFont2))
   RETURN pITextFont2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the default character formatting for this instance of the Text Object Model (TOM) engine.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetDocumentFont (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetDocumentFont(m_pTextDocument2, pFont))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an object that provides the default paragraph format information for this instance
' of the Text Object Model (TOM) engine.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetDocumentPara () AS ITextPara2 PTR
   DIM pITextPara2 AS ITextPara2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetDocumentPara(m_pTextDocument2, @pITextPara2))
   RETURN pITextPara2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the default paragraph formatting for this instance of the Text Object Model (TOM) engine.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetDocumentPara (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetDocumentPara(m_pTextDocument2, pPara))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the East Asian flags.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetEastAsianFlags () AS LONG
   DIM pFlags AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetEastAsianFlags(m_pTextDocument2, @pFlags))
   RETURN pFlags
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the East Asian flags.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetGenerator () AS CBSTR
   DIM bstr AS AFX_BSTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetGenerator(m_pTextDocument2, @bstr))
   RETURN bstr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the state of the Input Method Editor (IME) in-progress flag.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetIMEInProgress (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetIMEInProgress(m_pTextDocument2, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the notification mode.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetNotificationMode () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetNotificationMode(m_pTextDocument2, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the notification mode.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetNotificationMode (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetNotificationMode(m_pTextDocument2, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the active selection.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetSelection2 () AS ITextSelection2 PTR
   DIM ppSel AS ITextSelection2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetSelection2(m_pTextDocument2, @ppSel))
   RETURN ppSel
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an object for enumerating the stories in a document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetStoryRanges2 () AS ITextStoryRanges2 PTR
   DIM pStories AS ITextStoryRanges2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetStoryRanges2(m_pTextDocument2, @pStories))
   RETURN pStories
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the typography options.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetTypographyOptions () AS LONG
   DIM Options AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetTypographyOptions(m_pTextDocument2, @Options))
   RETURN Options
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the version number of the Text Object Model (TOM) engine.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetVersion () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetVersion(m_pTextDocument2, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the handle of the window that the Text Object Model (TOM) engine is using to display output.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetWindow () AS __int64
   DIM pHwnd AS __int64
   this.SetResult(m_pTextDocument2->lpvtbl->GetWindow(m_pTextDocument2, @pHwnd))
   RETURN pHwnd
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches a new message filter to the edit instance. All window messages that the edit
' instance receives are forwarded to the message filter.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.AttachMsgFilter (BYVAL pFilter AS IUnknown PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->AttachMsgFilter(m_pTextDocument2, pFilter))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Checks whether the number of characters to be added would exceed the maximum text limit.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.CheckTextLimit (BYVAL cch AS LONG) AS LONG
   DIM pcch AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->CheckTextLimit(m_pTextDocument2, cch, @pcch))
   RETURN pcch
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the call manager.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetCallManager () AS IUnknown PTR
   DIM pVoid AS IUnknown PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetCallManager(m_pTextDocument2, @pVoid))
   RETURN pVoid
END FUNCTION
' ========================================================================================
'   GetCallManager AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppVoid AS IUnknown PTR) AS HRESULT

' ========================================================================================
' Retrieves the client rectangle of the rich edit control.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetClientRect (BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, _
BYVAL pTop AS LONG PTR, BYVAL pRight AS LONG PTR, BYVAL pBottom AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->GetClientRect(m_pTextDocument2, nType, pLeft, pTop, pRight, pBottom))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the color used for special text attributes.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetEffectColor (BYVAL Index AS LONG) AS ULONG
   DIM pcr AS ULONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetEffectColor(m_pTextDocument2, Index, @pcr))
   RETURN pcr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the preferred font for a particular character repertoire and character position.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetPreferredFont (BYVAL cp AS LONG, BYVAL CharRep AS LONG, _
BYVAL Options AS LONG, BYVAL curCharRep AS LONG, BYVAL curFontSize AS LONG, BYVAL pbstr AS BSTR PTR, _
BYVAL pPitchAndFamily AS LONG PTR, BYVAL pNewFontSize AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->GetPreferredFont(m_pTextDocument2, cp, CharRep, _
   Options, curCharRep, curFontSize, pbstr, pPitchAndFamily, pNewFontSize))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the Input Method Manager (IMM) input context from the Text Object Model (TOM) host.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetImmContext () AS __int64
   DIM Context AS __int64
   this.SetResult(m_pTextDocument2->lpvtbl->GetImmContext(m_pTextDocument2, @Context))
   RETURN Context
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the value of a property.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetProperty(m_pTextDocument2, nType, @Value))
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a collection of rich-text strings.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetStrings () AS ITextStrings PTR
   DIM pStrs AS ITextStrings PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetStrings(m_pTextDocument2, @pStrs))
   RETURN pStrs
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the Text Object Model (TOM) engine client of particular Input Method Editor (IME) events.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Notify (BYVAL nNotify AS LONG) AS HRESULT
   DIM pStrs AS ITextStrings PTR
   this.SetResult(m_pTextDocument2->lpvtbl->Notify(m_pTextDocument2, nNotify))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a new text range for the active story of the document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Range2 (BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG) AS ITextRange2 PTR
   CTOM_DP("BEGIN CTextRange2.Range2 - m_pTextDocument2 = " & WSTR(m_pTextDocument2))
   DIM pRange2 AS ITextRange2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->Range2(m_pTextDocument2, cpActive, cpAnchor, @pRange2))
   CTOM_DP("END CTextRange2.Range2 - pRange2 = " & WSTR(pRange2))
   RETURN pRange2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the degenerate range at (or nearest to) a particular point on the screen.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.RangeFromPoint2 (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG) AS ITextRange2 PTR
   DIM pRange2 AS ITextRange2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->RangeFromPoint2(m_pTextDocument2, x, y, nType, @pRange2))
   RETURN pRange2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the call manager.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.ReleaseCallManager (BYVAL pVoid AS IUnknown PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->ReleaseCallManager(m_pTextDocument2, pVoid))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases an Input Method Manager (IMM) input context.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.ReleaseImmContext (BYVAL Context AS __int64) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->ReleaseImmContext(m_pTextDocument2, Context))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the color to use for special text attributes.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetEffectColor (BYVAL Index AS LONG, BYVAL Value AS ULONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetEffectColor(m_pTextDocument2, Index, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies a new value for a property.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetProperty(m_pTextDocument2, nType, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the typography options for the document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetTypographyOptions (BYVAL Options AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetTypographyOptions(m_pTextDocument2, Options, Mask))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Generates a system beep.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SysBeep () AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SysBeep(m_pTextDocument2))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the selection and caret.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.Update (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->Update(m_pTextDocument2, Value))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the client that the view has changed and the client should update the view if
' the Text Object Model (TOM) engine is in-place active.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.UpdateWindow () AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->UpdateWindow(m_pTextDocument2))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the math properties for the document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetMathProperties () AS LONG
   DIM Options AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetMathProperties(m_pTextDocument2, @Options))
   RETURN Options
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the math properties to use for the document.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetMathProperties (BYVAL Options AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetMathProperties(m_pTextDocument2, Options, Mask))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the active story; that is, the story that receives keyboard and mouse input.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetActiveStory () AS ITextStory PTR
   DIM pStory AS ITextStory PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetActiveStory(m_pTextDocument2, @pStory))
   RETURN pStory
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the active story; that is, the story that receives keyboard and mouse input.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.SetActiveStory (BYVAL pStory AS ITextStory PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetActiveStory(m_pTextDocument2, pStory))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the main story.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetMainStory () AS ITextStory PTR
   DIM pStory AS ITextStory PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetMainStory(m_pTextDocument2, @pStory))
   RETURN pStory
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a new story.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetNewStory () AS ITextStory PTR
   DIM pStory AS ITextStory PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetNewStory(m_pTextDocument2, @pStory))
   RETURN pStory
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the story that corresponds to a particular index.
' ========================================================================================
PRIVATE FUNCTION CTextDocument2.GetStory (BYVAl Index AS LONG) AS ITextStory PTR
   DIM pStory AS ITextStory PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetStory(m_pTextDocument2, Index, @pStory))
   RETURN pStory
END FUNCTION
' ========================================================================================

