# CRichEditCtrl Class

Provides the functionality of the rich edit control.

A "rich edit control" is a window in which the user can enter and edit text. The text can be assigned character and paragraph formatting, and can include embedded OLE objects. Rich edit controls provide a programming interface for formatting text. However, an application must implement any user interface components necessary to make formatting operations available to the user.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTOR](#CONSTRUCTOR) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [hRichEdit](#hRichEdit) | Returns the handle of the rich edit control. |

# <a name="CONSTRUCTOR"></a>CONSTRUCTOR

Creates an instance of the rich edit control.

```
CONSTRUCTOR CRichEditCtrl (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszTitle AS WSTRING = "", _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
```
| Parameter | Description |
| --------- | ----------- |
| *pWindow* | Pointer to the parent `CWindow`class. |
| *cID* | The control identifier, an integer value used to notify its parent about events. The application determines the control identifier; it must be unique for all controls with the same parent window. |
| *wszTitle* | Optional. The text of the control. |
| *x* | Optional. The x-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *y* | Optional. The initial y-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *nWidth* | Optional. The width of the control. |
| *nHeight* | Optional. The height of the control. |
| *dwStyle* | Optional. The window styles of the control being created.<br>Default styles: WS_VISIBLE OR WS_CHILD OR WS_TABSTOP OR ES_LEFT OR WS_HSCROLL OR WS_VSCROLL OR ES_AUTOHSCROLL OR ES_AUTOVSCROLL OR ES_MULTILINE OR ES_WANTRETURN OR ES_NOHIDESEL OR ES_SAVESEL |
| *dwExStyle* | Optional. The extended window styles of the control being created.<br>Default extended style: WS_EX_CLIENTEDGE |
| *lParam* | Optional. Pointer to custom data. |

# <a name="DESTRUCTOR"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

# <a name="hRichEdit"></a>hRichEdit

Returns the handle of the rich edit control.

### CRichEditCtrl Properties

| Name       | Description |
| ---------- | ----------- |
| [AutoCorrectProc (Get)](#AutoCorrectProc1) | Gets a pointer to the application-defined AutoCorrectProc callback function. |
| [AutoCorrectProc (Set)](#AutoCorrectProc2) | Sets a pointer to the application-defined AutoCorrectProc callback function. |
| [AutoUrlDetect (Get)](#AutoUrlDetect1) | Indicates whether the auto URL detection is turned on in the rich edit control. |
| [AutoUrlDetect (Set)](#AutoUrlDetect2) | Enables or disables automatic detection of URLs by a rich edit control. |
| [BidiOptions (Get)](#BidiOptions1) | Indicates the current state of the bidirectional options in the rich edit control. |
| [BidiOptions (Set)](#BidiOptions2) | Sets the current state of the bidirectional options in the rich edit control. |
| [CharFormat (Get)](#CharFormat1) | Determines the current character formatting in a rich edit control. |
| [CharFormat (Set)](#CharFormat2) | Sets character formatting in a rich edit control. |
| [CTFModeBias (Get)](#CTFModeBias1) | Gets the Text Services Framework mode bias values for a Microsoft Rich Edit control. |
| [CTFModeBias (Set)](#CTFModeBias2) | Sets the Text Services Framework mode bias values for a Microsoft Rich Edit control. |
| [CTFOpenStatus (Get)](#CTFOpenStatus1) | Determines if the Text Services Framework (TSF) keyboard is open or closed. |
| [CTFOpenStatus (Set)](#CTFOpenStatus2) | Opens or closes the Text Services Framework (TSF) keyboard. |
| [EditStyle (Get)](#EditStyle1) | Retrieves the current edit style flags. |
| [EditStyle (Set)](#EditStyle2) | Sets the current edit style flags. |
| [EditStyleEx (Get)](#EditStyleEx1) | Returns the extended edit style flags. |
| [EditStyleEx (Set)](#EditStyleEx2) | Sets the current extended edit style flags. |
| [EllipsisMode (Get)](#EllipsisMode1) | Retrieves the current ellipsis mode. |
| [EllipsisMode (Set)](#EllipsisMode2) | Sets the current ellipsis mode. |
| [EventMask (Get)](#EventMask1) | Retrieves the event mask for a rich edit control. The event mask specifies which notification messages the control sends to its parent window. |
| [EventMask (Set)](#EventMask2) | Sets the event mask for a rich edit control. |
| [HyphenateInfo (Get)](#HyphenateInfo1) | Gets information about hyphenation for a Microsoft Rich Edit control. |
| [HyphenateInfo (Set)](#HyphenateInfo2) | Sets the way a Microsoft Rich Edit control does hyphenation.|
| [IMEColor (Get)](#IMEColor1) | Retrieves the Input Method Editor (IME) composition color. This message is available only in Asian-language versions of the operating system. |
| [IMEColor (Set)](#IMEColor2) | Sets the Input Method Editor (IME) composition color. |
| [IMEModeBias (Get)](#IMEModeBias1) | Gets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control. |
| [IMEModeBias (Set)](#IMEModeBias2) | Sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control. |
| [IMEOptions (Get)](#IMEOptions1) | Retrieves the current Input Method Editor (IME) options. This message is available only in Asian-language versions of the operating system. |
| [IMEOptions (Set)](#IMEOptions2) | Sets the Input Method Editor (IME) options. |
| [LangOptions (Get)](#LangOptions1) | Gets a rich edit control's option settings for Input Method Editor (IME) and Asian language support. |
| [LangOptions (Set)](#LangOptions2) | Sets options for Input Method Editor (IME) and Asian language support in a rich edit control. |
| [LimitText (Get)](#LimitText1) | Gets the current text limit for a rich edit control. |
| [LimitText (Set)](#LimitText2) | Sets the text limit of a rich edit control. The text limit is the maximum amount of text, in TCHARs, that the user can type into the edit control. |
| [Modify (Get)](#Modify1) | Gets the state of a rich edit control's modification flag. The flag indicates whether the contents of the rich edit control have been modified. |
| [Modify (Set)](#Modify2) | Sets or clears the modification flag for a rich edit control. The modification flag indicates whether the text within the rich edit control has been modified. |
| [Options (Get)](#Options1) | Retrieves the options for a rich edit control. |
| [Options (Set)](#Options2) | Sets the options for a rich edit control. |
| [PageRotate (Get)](#PageRotate1) | Deprecated. Gets the text layout for a Microsoft Rich Edit control. |
| [PageRotate (Set)](#PageRotate2) | Deprecated. Sets the text layout for a Microsoft Rich Edit control. |
| [ParaFormat (Get)](#ParaFormat1) | Retrieves the paragraph formatting of the current selection in a rich edit control. |
| [ParaFormat (Set)](#ParaFormat2) | Sets the paragraph formatting for the current selection in a rich edit control. |
| [PasswordChar (Get)](#PasswordChar1) | Gets the password character that a rich edit control displays when the user enters text. |
| [PasswordChar (Set)](#PasswordChar2) | Sets or removes the password character for a rich edit control. When a password character is set, that character is displayed in place of the characters typed by the user. |
| [Punctuation (Get)](#Punctuation1) | Gets the current punctuation characters for the rich edit control. |
| [Punctuation (Set)](#Punctuation2) | Sets the punctuation characters for a rich edit control. |
| [Rect (Get)](#Rect1) | Gets the formatting rectangle of a rich edit control. |
| [Rect (Set)](#Rect2) | Sets the formatting rectangle of a multiline rich edit control. |
| [RectNP (Set)](#RectNP) | Sets the formatting rectangle of a multiline rich edit control. |
| [ScrollPos (Get)](#ScrollPos1) | Obtains the current scroll position of the edit control. |
| [ScrollPos (Set)](#ScrollPos2) | Tells the rich edit control to scroll to a particular point. |
| [StoryType (Get)](#StoryType1) | Gets the story type. |
| [StoryType (Set)](#StoryType2) | Sets the story type. |
| [Text (Get)](#Text1) | Retrieves the text from a rich edit control. |
| [Text (Set)](#Text2) | Sets the text of an edit control. |
| [TextMode (Get)](#TextMode1) | Gets the current text mode and undo level of a rich edit control. |
| [TextMode (Set)](#TextMode2) | Sets the text mode or undo level of a rich edit control. |
| [TouchOptions (Get)](#TouchOptions1) | Retrieves the touch options that are associated with a rich edit control. |
| [TouchOptions (Set)](#TouchOptions2) | Sets the touch options associated with a rich edit control. |
| [TypographyOptions (Get)](#TypographyOptions1) | Returns the current state of the typography options of a rich edit control. |
| [TypographyOptions (Set)](#TypographyOptions2) | Sets the current state of the typography options of a rich edit control. |

### CRichEditCtrl Methods

| Name       | Description |
| ---------- | ----------- |
| [CanPaste](#CanPaste) | Determines whether a rich edit control can paste a specified clipboard format. |
| [CanRedo](#CanRedo) | Determines whether there are any actions in the control redo queue. |
| [CanUndo](#CanUndo) | Determines whether there are any actions in an edit control's undo queue. |
| [CallAutocorrectProc](#CallAutocorrectProc) | Calls the autocorrect callback function that is stored by the **CRichEditCtrl.SetAutocorrectProc** message, provided that the text preceding the insertion point is a candidate for autocorrection. |
| [DisplayBand](#DisplayBand) | Displays a portion of the contents of a rich edit control, as previously formatted for a device using the EM_FORMATRANGE message. |
| [EmptyUndoBuffer](#EmptyUndoBuffer) | Resets the undo flag of a rich edit control. The undo flag is set whenever an operation within the rich edit control can be undone. |
| [ExGetSel](#ExGetSel) | Retrieves the starting and ending character positions of the selection in a rich edit control. |
| [ExLimitText](#ExLimitText) | Sets an upper limit to the amount of text the user can type or paste into a rich edit control. |
| [ExLineFromChar](#ExLineFromChar) | Determines which line contains the specified character in a rich edit control. |
| [ExSetSel](#ExSetSel) | Selects a range of characters and/or Component Object Model (COM) objects in a Microsoft Rich Edit control. |
| [FindText](#FindText) | Finds text within a rich edit control. |
| [FindTextEx](#FindTextEx) | Finds text within a rich edit control. |
| [FindWordBreak](#FindWordBreak) | Finds the next word break before or after the specified character position or retrieves information about the character at that position. |
| [GetFirstVisibleLine](#GetFirstVisibleLine) | Gets the zero-based index of the uppermost visible line in a multiline rich edit control. |
| [FormatRange](#FormatRange) | Formats a range of text in a rich edit control for a specific device. |
| [GetCharFromPos](#GetCharFromPos) | Gets information about the character closest to a specified point in the client area of an edit control. |
| [GetEllipsisState](#GetEllipsisState) | Retrieves the current ellipsis state. |
| [GetIMECompMode](#GetIMECompMode) | Gets the current IME mode for a rich edit control. |
| [GetIMEProperty](#GetIMEProperty) | Gets the property and capabilities of the Input Method Editor (IME) associated with the current input locale. |
| [GetLine](#GetLine) | Copies a line of text from a rich edit control. |
| [GetOleInterface](#GetOleInterface) | Retrieves an IRichEditOle object that a client can use to access a rich edit control's Component Object Model (COM) functionality. |
| [GetRedoName](#GetRedoName) | Retrieves the type of the next action, if any, in the control's redo queue. |
| [GetSel](#GetSel) | Gets the starting and ending character positions of the current selection in a rich edit control. |
| [SetSel](#SetSel) | Selects a range of characters in a rich edit control. |
| [GetSelText](#GetSelText) | Retrieves the currently selected text in a rich edit control. |
| [GetUndoName](#GetUndoName) | Retrieves the type of the next undo action, if any. |
| [GetWordWrapMode](#GetWordWrapMode) | Gets the current word wrap and word-break options for the rich edit control. |
| [SetWordWrapMode](#SetWordWrapMode) | Sets the word-wrapping and word-breaking options for the rich edit control. |
| [GetZoom](#GetZoom) | Gets the current zoom ratio, which is always between 1/64 and 64. |
| [SetZoom](#SetZoom) | Sets the zoom ratio anywhere between 1/64 and 64. |
| [GetTableParams](#GetTableParams) | Retrieves the table parameters for a table row and the cell parameters for the specified number of cells. |
| [SetTableParams](#SetTableParams) | Changes the parameters of rows in a table. |
| [GetTextEx](#GetTextEx) | Gets all of the text from the rich edit control in any particular code base you want. |
| [SetTextExW](#SetTextExW) | Combines the functionality of WM_SETTEXT and EM_REPLACESEL and adds the ability to set text using a code page and to use either Rich Text Format (RTF) rich text or plain text. |
| [GetThumb](#GetThumb) | Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline rich edit control. |
| [GetWordBreakProc](#GetWordBreakProc) | Gets the address of the current Wordwrap function. |
| [SetWordBreakProc](#SetWordBreakProc) | Replaces a rich edit control's default Wordwrap function with an application-defined Wordwrap function. |
| [GetWordBreakProcEx](#GetWordBreakProcEx) | Retrieves the address of the currently registered extended word-break procedure. |
| [SetWordBreakProcEx](#SetWordBreakProcEx) | Sets the extended word-break procedure. |
| [HideSelection](#HideSelection) | Hides or shows the selection in a rich edit control. |
| [GetIMECompText](#GetIMECompText) | Gets the Input Method Editor (IME) composition text. |
| [InsertImage](#InsertImage) | Replaces the selection with a blob that displays an image. |
| [InsertTable](#InsertTable) | Inserts one or more identical table rows with empty cells. |
| [IsIME](#IsIME) | Determines if current input locale is an East Asian locale. |
| [GetLineCount](#GetLineCount) | Gets the number of lines in a multiline rich edit control. |
| [LineFromChar](#LineFromChar) | Gets the index of the line that contains the specified character index in a multiline rich edit control. |
| [LineIndex](#LineIndex) | Gets the character index of the first character of a specified line in a multiline rich edit control. |
| [LineLength](#LineLength) | Retrieves the length, in characters, of a line in a rich edit control. |
| [LineScroll](#LineScroll) | Scrolls the text in a multiline rich edit control. |
| [PasteSpecial](#PasteSpecial) | Pastes a specific clipboard format in a rich edit control. |
| [PosFromChar](#PosFromChar) | Retrieves the client area coordinates of a specified character in a rich edit control. |
| [Redo](#Redo) | Redoes the next action in the control's redo queue. |
| [ReplaceSel](#ReplaceSel) | Replaces the current selection in a rich edit control with the specified text. |
| [RequestResize](#RequestResize) | Forces a rich edit control to send an EN_REQUESTRESIZE notification message to its parent window. |
| [Reconversion](#Reconversion) | Invokes the Input Method Editor (IME) reconversion dialog box. |
| [Scroll](#Scroll) | Scrolls the text vertically in a multiline rich edit control. |



| Name       | Description |
| ---------- | ----------- |
| [](#) |  |


