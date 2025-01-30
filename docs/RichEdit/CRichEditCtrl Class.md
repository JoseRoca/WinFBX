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

#### Usage example
```
' // Add a rich edit control without coordinates (it will be resized in WM_SIZE)
DIM hRichEdit AS HWND = pWindow.AddControl("RichEdit", , IDC_RICHEDIT, "RichEdit box test")
```

# <a name="DESTRUCTOR"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.
```
DESTRUCTOR CRichEditCtrl
```

# <a name="hRichEdit"></a>hRichEdit

Returns the handle of the rich edit control.

```
FUNCTION CRichEditCtrl.hRichEdit () AS HWND
```
#### Usage example
```
' // Add a rich edit control without coordinates (it will be resized in WM_SIZE)
DIM hRichEdit AS HWND = pWindow.AddControl("RichEdit", , IDC_RICHEDIT, "RichEdit box test")
'...
DIM hRichEdit AS HWND = pRichEdit.hRichEdit
```

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
| [GetFirstVisibleLine](#GetFirstVisibleLine) | Gets the zero-based index of the uppermost visible line in a  rich edit control. |
| [FormatRange](#FormatRange) | Formats a range of text in a rich edit control for a specific device. |
| [GetCharFromPos](#GetCharFromPos) | Gets information about the character closest to a specified point in the client area of an edit control. |
| [GetEllipsisState](#GetEllipsisState) | Retrieves the current ellipsis state. |
| [GetIMECompMode](#GetIMECompMode) | Gets the current IME mode for a rich edit control. |
| [GetIMECompText](#GetIMECompText) | Gets the Input Method Editor (IME) composition text. |
| [GetIMEProperty](#GetIMEProperty) | Gets the property and capabilities of the Input Method Editor (IME) associated with the current input locale. |
| [GetLine](#GetLine) | Copies a line of text from a rich edit control. |
| [GetLineCount](#GetLineCount) | Gets the number of lines in a multiline rich edit control. |
| [GetOleInterface](#GetOleInterface) | Retrieves an IRichEditOle object that a client can use to access a rich edit control's Component Object Model (COM) functionality. |
| [GetRedoName](#GetRedoName) | Retrieves the type of the next action, if any, in the control's redo queue. |
| [GetSel](#GetSel) | Gets the starting and ending character positions of the current selection in a rich edit control. |
| [GetSelText](#GetSelText) | Retrieves the currently selected text in a rich edit control. |
| [GetUndoName](#GetUndoName) | Retrieves the type of the next undo action, if any. |
| [GetWordWrapMode](#GetWordWrapMode) | Gets the current word wrap and word-break options for the rich edit control. |
| [GetZoom](#GetZoom) | Gets the current zoom ratio, which is always between 1/64 and 64. |
| [GetTableParams](#GetTableParams) | Retrieves the table parameters for a table row and the cell parameters for the specified number of cells. |
| [GetTextEx](#GetTextEx) | Gets all of the text from the rich edit control in any particular code base you want. |
| [GetTextLength](#GetTextLength) | Retrieves the length of all text in a rich edit control. |
| [GetTextLengthEx](#GetTextLengthEx) | Calculates text length in various ways. It is usually called before creating a buffer to receive the text from the control. |
| [GetTextRange](#GetTextRange) | Retrieves a specified range of characters from a rich edit control. |
| [GetThumb](#GetThumb) | Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline rich edit control. |
| [GetWordBreakProc](#GetWordBreakProc) | Gets the address of the current Wordwrap function. |
| [GetWordBreakProcEx](#GetWordBreakProcEx) | Retrieves the address of the currently registered extended word-break procedure. |
| [HideSelection](#HideSelection) | Hides or shows the selection in a rich edit control. |
| [InsertImage](#InsertImage) | Replaces the selection with a blob that displays an image. |
| [InsertTable](#InsertTable) | Inserts one or more identical table rows with empty cells. |
| [IsIME](#IsIME) | Determines if current input locale is an East Asian locale. |
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
| [ScrollCaret](#ScrollCaret) | Scrolls the caret into view in a rich edit control. |
| [SelectionType](#SelectionType) | Determines the selection type for a rich edit control. |
| [SetBkgndColor](#SetBkgndColor) | Sets the background color for a rich edit control. |
| [SetFontSize](#SetFontSize) | Sets the font size for the selected text. |
| [SetMargins](#SetMargins) | Sets the widths of the left and right margins for a rich edit control. The message redraws the control to reflect the new margins. |
| [SetOleCallback](#SetOleCallback) | Gives a rich edit control an IRichEditOleCallback object that the control uses to get OLE-related resources and information from the client. |
| [SetPalette](#SetPalette) | Changes the palette that a rich edit control uses for its display window. |
| [SetReadOnly](#SetReadOnly) | Changes the palette that a rich edit control uses for its display window. |
| [SetSel](#SetSel) | Selects a range of characters in a rich edit control. |
| [SetTableParams](#SetTableParams) | Changes the parameters of rows in a table. |
| [SetTabStops](#SetTabStops) | Sets the tab stops in a multiline rich edit control. |
| [SetTargetDevice](#SetTargetDevice) | Sets the target device and line width used for WYSIWYG formatting in a rich edit control. |
| [SetTextExW](#SetTextExW) | Combines the functionality of WM_SETTEXT and EM_REPLACESEL and adds the ability to set text using a code page and to use either Rich Text Format (RTF) rich text or plain text. |
| [SetUIAName](#SetUIAName) | Sets the maximum number of actions that can stored in the undo queue. |
| [SetUndoLimit](#SetUndoLimit) | Sets the maximum number of actions that can stored in the undo queue. |
| [SetWordWrapMode](#SetWordWrapMode) | Sets the word-wrapping and word-breaking options for the rich edit control. |
| [SetWordBreakProc](#SetWordBreakProc) | Replaces a rich edit control's default Wordwrap function with an application-defined Wordwrap function. |
| [SetWordBreakProcEx](#SetWordBreakProcEx) | Sets the extended word-break procedure. |
| [SetZoom](#SetZoom) | Sets the zoom ratio anywhere between 1/64 and 64. |
| [ShowScrollBar](#ShowScrollBar) | Shows or hides one of the scroll bars in the Text Host window. |
| [StopGroupTyping](#StopGroupTyping) | Stops the control from collecting additional typing actions into the current undo action. |
| [StreamIn](#StreamIn) | Replaces the contents of a rich edit control with a stream of data provided by an application defined EditStreamCallback callback function. |
| [StreamOut](#StreamOut) | Causes a rich edit control to pass its contents to an application defined EditStreamCallback callback function. |
| [Undo](#Undo) | This message undoes the last edit control operation in the control's undo queue. |

### Methods inherited from CTextObjectBase

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#GetLastResult) | Returns the last result code. |
| [SetResult](#SetResult) | Sets the last result code. |
| [GetErrorInfo](#GetErrorInfo) | Returns a localized description of the last result code. |

# <a name="AutoCorrectProc (Get)"></a>AutoCorrectProc1

Gets a pointer to the application-defined AutoCorrectProc callback function.

# <a name="AutoCorrectProc (Set)"></a>AutoCorrectProc2

Sets a pointer to the application-defined AutoCorrectProc callback function.

# <a name="AutoUrlDetect (Get)"></a>AutoUrlDetect1

Indicates whether the auto URL detection is turned on in the rich edit control.

# <a name="AutoUrlDetect (Set)"></a>AutoUrlDetect2

Enables or disables automatic detection of URLs by a rich edit control.

# <a name="BidiOptions (Get)"></a>BidiOptions1

Indicates the current state of the bidirectional options in the rich edit control.

# <a name="BidiOptions (Set)"></a>BidiOptions2

Sets the current state of the bidirectional options in the rich edit control.

# <a name="CharFormat (Get)"></a>CharFormat1

Determines the current character formatting in a rich edit control.

# <a name="CharFormat (Set)"></a>CharFormat2

Sets character formatting in a rich edit control.

# <a name="CTFModeBias (Get)"></a>CTFModeBias1

Gets the Text Services Framework mode bias values for a Microsoft Rich Edit control.

# <a name="CTFModeBias (Set)"></a>CTFModeBias2

Sets the Text Services Framework mode bias values for a Microsoft Rich Edit control.

# <a name="CTFOpenStatus (Get)"></a>CTFOpenStatus1

Determines if the Text Services Framework (TSF) keyboard is open or closed.

# <a name="CTFOpenStatus (Set)"></a>CTFOpenStatus2

Opens or closes the Text Services Framework (TSF) keyboard.

# <a name="EditStyle (Get)"></a>EditStyle1

Retrieves the current edit style flags.

# <a name="EditStyle (Set)"></a>EditStyle2

Sets the current edit style flags.

# <a name="EditStyleEx (Get)"></a>EditStyleEx1

Returns the extended edit style flags.

# <a name="EditStyleEx (Set)"></a>EditStyleEx2

Sets the current extended edit style flags.

# <a name="EllipsisMode (Get)"></a>EllipsisMode1

Retrieves the current ellipsis mode.

# <a name="EllipsisMode (Set)"></a>EllipsisMode2

Sets the current ellipsis mode.

# <a name="EventMask (Get)"></a>EventMask1

Retrieves the event mask for a rich edit control. The event mask specifies which notification messages the control sends to its parent window.

# <a name="EventMask (Set)"></a>EventMask2

Sets the event mask for a rich edit control.

# <a name="HyphenateInfo (Get)"></a>HyphenateInfo1

Gets information about hyphenation for a Microsoft Rich Edit control.

# <a name="HyphenateInfo (Set)"></a>HyphenateInfo2

Sets the way a Microsoft Rich Edit control does hyphenation.|

# <a name="IMEColor (Get)"></a>IMEColor1

Retrieves the Input Method Editor (IME) composition color. This message is available only in Asian-language versions of the operating system.

# <a name="IMEColor (Set)"></a>IMEColor2

Sets the Input Method Editor (IME) composition color.

# <a name="IMEModeBias (Get)"></a>IMEModeBias1

Gets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.

# <a name="IMEModeBias (Set)"></a>IMEModeBias2

Sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.

# <a name="IMEOptions (Get)"></a>IMEOptions1

Retrieves the current Input Method Editor (IME) options. This message is available only in Asian-language versions of the operating system.

# <a name="IMEOptions (Set)"></a>IMEOptions2

Sets the Input Method Editor (IME) options.

# <a name="LangOptions (Get)"></a>LangOptions1

Gets a rich edit control's option settings for Input Method Editor (IME) and Asian language support.

# <a name="LangOptions (Set)"></a>LangOptions2

Sets options for Input Method Editor (IME) and Asian language support in a rich edit control.

# <a name="LimitText (Get)"></a>LimitText1

Gets the current text limit for a rich edit control.

# <a name="LimitText (Set)"></a>LimitText2

Sets the text limit of a rich edit control. The text limit is the maximum amount of text, in TCHARs, that the user can type into the edit control.

# <a name="Modify (Get)"></a>Modify1

Gets the state of a rich edit control's modification flag. The flag indicates whether the contents of the rich edit control have been modified.

# <a name="Modify (Set)"></a>Modify2

Sets or clears the modification flag for a rich edit control. The modification flag indicates whether the text within the rich edit control has been modified.

# <a name="Options (Get)"></a>Options1

Retrieves the options for a rich edit control.

# <a name="Options (Set)"></a>Options2

Sets the options for a rich edit control.

# <a name="PageRotate (Get)"></a>PageRotate1

Deprecated. Gets the text layout for a Microsoft Rich Edit control.

# <a name="PageRotate (Set)"></a>PageRotate2

Deprecated. Sets the text layout for a Microsoft Rich Edit control.

# <a name="ParaFormat (Get)"></a>ParaFormat1

Retrieves the paragraph formatting of the current selection in a rich edit control.

# <a name="ParaFormat (Set)"></a>ParaFormat2

Sets the paragraph formatting for the current selection in a rich edit control.

# <a name="PasswordChar (Get)"></a>PasswordChar1

Gets the password character that a rich edit control displays when the user enters text.

# <a name="PasswordChar (Set)"></a>PasswordChar2

Sets or removes the password character for a rich edit control. When a password character is set, that character is displayed in place of the characters typed by the user.

# <a name="Punctuation (Get)"></a>Punctuation1

Gets the current punctuation characters for the rich edit control.

# <a name="Punctuation (Set)"></a>Punctuation2

Sets the punctuation characters for a rich edit control.

# <a name="Rect (Get)"></a>Rect1

Gets the formatting rectangle of a rich edit control.

# <a name="Rect (Set)"></a>Rect2

Sets the formatting rectangle of a multiline rich edit control.

# <a name="RectNP (Set)"></a>RectNP

Sets the formatting rectangle of a multiline rich edit control.

# <a name="ScrollPos (Get)"></a>ScrollPos1

Obtains the current scroll position of the edit control.

# <a name="ScrollPos (Set)"></a>ScrollPos2

Tells the rich edit control to scroll to a particular point.

# <a name="StoryType (Get)"></a>StoryType1

Gets the story type.

# <a name="StoryType (Set)"></a>StoryType2

Sets the story type.

# <a name="Text (Get)"></a>Text1

Retrieves the text from a rich edit control.

# <a name="Text (Set)"></a>Text2

Sets the text of an edit control.

# <a name="TextMode (Get)"></a>TextMode1

Gets the current text mode and undo level of a rich edit control.

# <a name="TextMode (Set)"></a>TextMode2

Sets the text mode or undo level of a rich edit control.

# <a name="TouchOptions (Get)"></a>TouchOptions1

Retrieves the touch options that are associated with a rich edit control.

# <a name="TouchOptions (Set)"></a>TouchOptions2

Sets the touch options associated with a rich edit control.

# <a name="TypographyOptions (Get)"></a>TypographyOptions1

Returns the current state of the typography options of a rich edit control.

# <a name="TypographyOptions (Set)"></a>TypographyOptions2

Sets the current state of the typography options of a rich edit control.


### CRichEditCtrl Methods

# <a name="CanPaste"></a>CanPaste

Determines whether a rich edit control can paste a specified clipboard format.

# <a name="CanRedo"></a>CanRedo

Determines whether there are any actions in the control redo queue.

# <a name="CanUndo"></a>CanUndo

Determines whether there are any actions in an edit control's undo queue.

# <a name="CallAutocorrectProc"></a>CallAutocorrectProc

Calls the autocorrect callback function that is stored by the **CRichEditCtrl.SetAutocorrectProc** message, provided that the text preceding the insertion point is a candidate for autocorrection.

# <a name="DisplayBand"></a>DisplayBand

Displays a portion of the contents of a rich edit control, as previously formatted for a device using the EM_FORMATRANGE message.

# <a name="EmptyUndoBuffer"></a>EmptyUndoBuffer

Resets the undo flag of a rich edit control. The undo flag is set whenever an operation within the rich edit control can be undone.

# <a name="ExGetSel"></a>ExGetSel

Retrieves the starting and ending character positions of the selection in a rich edit control.

# <a name="ExLimitText"></a>ExLimitText

Sets an upper limit to the amount of text the user can type or paste into a rich edit control.

# <a name="ExLineFromChar"></a>ExLineFromChar

Determines which line contains the specified character in a rich edit control.

# <a name="ExSetSel"></a>ExSetSel

Selects a range of characters and/or Component Object Model (COM) objects in a Microsoft Rich Edit control.

# <a name="FindText"></a>FindText

Finds text within a rich edit control.

# <a name="FindTextEx"></a>FindTextEx

Finds text within a rich edit control.

# <a name="FindWordBreak"></a>FindWordBreak

Finds the next word break before or after the specified character position or retrieves information about the character at that position.

# <a name="GetFirstVisibleLine"></a>GetFirstVisibleLine

Gets the zero-based index of the uppermost visible line in a multiline rich edit control.

# <a name="FormatRange"></a>FormatRange

Formats a range of text in a rich edit control for a specific device.

# <a name="GetCharFromPos"></a>GetCharFromPos

Gets information about the character closest to a specified point in the client area of an edit control.

# <a name="GetEllipsisState"></a>GetEllipsisState

Retrieves the current ellipsis state.

# <a name="GetIMECompMode"></a>GetIMECompMode

Gets the current IME mode for a rich edit control.

# <a name="GetIMECompText"></a>GetIMECompText

Gets the Input Method Editor (IME) composition text.

# <a name="GetIMEProperty"></a>GetIMEProperty

Gets the property and capabilities of the Input Method Editor (IME) associated with the current input locale.

# <a name="GetLine"></a>GetLine

Copies a line of text from a rich edit control.

# <a name="GetLineCount"></a>GetLineCount

Gets the number of lines in a multiline rich edit control.

# <a name="GetOleInterface"></a>GetOleInterface

Retrieves an IRichEditOle object that a client can use to access a rich edit control's Component Object Model (COM) functionality.

# <a name="GetRedoName"></a>GetRedoName

Retrieves the type of the next action, if any, in the control's redo queue.

# <a name="GetSel"></a>GetSel

Gets the starting and ending character positions of the current selection in a rich edit control.

# <a name="GetSelText"></a>GetSelText

Retrieves the currently selected text in a rich edit control.

# <a name="GetUndoName"></a>GetUndoName

Retrieves the type of the next undo action, if any.

# <a name="GetWordWrapMode"></a>GetWordWrapMode

Gets the current word wrap and word-break options for the rich edit control.

# <a name="GetZoom"></a>GetZoom

Gets the current zoom ratio, which is always between 1/64 and 64.

# <a name="GetTableParams"></a>GetTableParams

Retrieves the table parameters for a table row and the cell parameters for the specified number of cells.

# <a name="GetTextEx"></a>GetTextEx

Gets all of the text from the rich edit control in any particular code base you want.

# <a name="GetTextLength"></a>GetTextLength

Retrieves the length of all text in a rich edit control.

# <a name="GetTextLengthEx"></a>GetTextLengthEx

Calculates text length in various ways. It is usually called before creating a buffer to receive the text from the control.

# <a name="GetTextRange"></a>GetTextRange

Retrieves a specified range of characters from a rich edit control.

# <a name="GetThumb"></a>GetThumb

Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline rich edit control.

# <a name="GetWordBreakProc"></a>GetWordBreakProc

Gets the address of the current Wordwrap function.

# <a name="GetWordBreakProcEx"></a>GetWordBreakProcEx

Retrieves the address of the currently registered extended word-break procedure.

# <a name="HideSelection"></a>HideSelection

Hides or shows the selection in a rich edit control.

# <a name="InsertImage"></a>InsertImage

Replaces the selection with a blob that displays an image.

# <a name="InsertTable"></a>InsertTable

Inserts one or more identical table rows with empty cells.

# <a name="IsIME"></a>IsIME

Determines if current input locale is an East Asian locale.

# <a name="LineFromChar"></a>LineFromChar

Gets the index of the line that contains the specified character index in a multiline rich edit control.

# <a name="LineIndex"></a>LineIndex

Gets the character index of the first character of a specified line in a multiline rich edit control.

# <a name="LineLength"></a>LineLength

Retrieves the length, in characters, of a line in a rich edit control.

# <a name="LineScroll"></a>LineScroll

Scrolls the text in a multiline rich edit control.

# <a name="PasteSpecial"></a>PasteSpecial

Pastes a specific clipboard format in a rich edit control.

# <a name="PosFromChar"></a>PosFromChar

Retrieves the client area coordinates of a specified character in a rich edit control.

# <a name="Redo"></a>Redo

Redoes the next action in the control's redo queue.

# <a name="ReplaceSel"></a>ReplaceSel

Replaces the current selection in a rich edit control with the specified text.

# <a name="RequestResize"></a>RequestResize

Forces a rich edit control to send an EN_REQUESTRESIZE notification message to its parent window.

# <a name="Reconversion"></a>Reconversion

Invokes the Input Method Editor (IME) reconversion dialog box.

# <a name="Scroll"></a>Scroll

Scrolls the text vertically in a multiline rich edit control.

# <a name="ScrollCaret"></a>ScrollCaret

Scrolls the caret into view in a rich edit control.

# <a name="SelectionType"></a>SelectionType

Determines the selection type for a rich edit control.

# <a name="SetBkgndColor"></a>SetBkgndColor

Sets the background color for a rich edit control.

# <a name="SetFontSize"></a>SetFontSize

Sets the font size for the selected text.

# <a name="SetMargins"></a>SetMargins

Sets the widths of the left and right margins for a rich edit control. The message redraws the control to reflect the new margins.

# <a name="SetOleCallback"></a>SetOleCallback

Gives a rich edit control an IRichEditOleCallback object that the control uses to get OLE-related resources and information from the client.

# <a name="SetPalette"></a>SetPalette

Changes the palette that a rich edit control uses for its display window.

# <a name="SetReadOnly"></a>SetReadOnly

Changes the palette that a rich edit control uses for its display window.

# <a name="SetSel"></a>SetSel

Selects a range of characters in a rich edit control.

# <a name="SetTableParams"></a>SetTableParams

Changes the parameters of rows in a table.

# <a name="SetTabStops"></a>SetTabStops

Sets the tab stops in a multiline rich edit control.

# <a name="SetTargetDevice"></a>SetTargetDevice

Sets the target device and line width used for WYSIWYG formatting in a rich edit control.

# <a name="SetTextExW"></a>SetTextExW

Combines the functionality of WM_SETTEXT and EM_REPLACESEL and adds the ability to set text using a code page and to use either Rich Text Format (RTF) rich text or plain text.

# <a name="SetUIAName"></a>SetUIAName

Sets the maximum number of actions that can stored in the undo queue.

# <a name="SetUndoLimit"></a>SetUndoLimit

Sets the maximum number of actions that can stored in the undo queue.

# <a name="SetWordWrapMode"></a>SetWordWrapMode

Sets the word-wrapping and word-breaking options for the rich edit control.

# <a name="SetWordBreakProc"></a>SetWordBreakProc

Replaces a rich edit control's default Wordwrap function with an application-defined Wordwrap function.

# <a name="SetWordBreakProcEx"></a>SetWordBreakProcEx

Sets the extended word-break procedure.

# <a name="SetZoom"></a>SetZoom

Sets the zoom ratio anywhere between 1/64 and 64.

# <a name="ShowScrollBar"></a>ShowScrollBar

Shows or hides one of the scroll bars in the Text Host window.

# <a name="StopGroupTyping"></a>StopGroupTyping

Stops the control from collecting additional typing actions into the current undo action.

# <a name="StreamIn"></a>StreamIn

Replaces the contents of a rich edit control with a stream of data provided by an application defined EditStreamCallback callback function.

# <a name="StreamOut"></a>StreamOut

Causes a rich edit control to pass its contents to an application defined EditStreamCallback callback function.

# <a name="Undo"></a>Undo

This message undoes the last edit control operation in the control's undo queue.
