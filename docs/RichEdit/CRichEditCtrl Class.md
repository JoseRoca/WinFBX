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

#### Usage example (dotted syntax)
```
DIM pRichEdit AS CRichEditCtrl = CRichEditCtrl(@pWindow, IDC_RICHEDIT, "RichEditbox", 100, 50, 300, 200)
```
#### Usage example (pointer syntax)
```
DIM pRichEdit AS CRichEditCtrl PTR = NEW CRichEditCtrl(@pWindow, IDC_RICHEDIT, "RichEditbox", 100, 50, 300, 200)

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
' // Create an instance of the CRichEditCtrl class
DIM pRichEdit AS CRichEditCtrl = CRichEditCtrl(@pWindow, IDC_RICHEDIT, "RichEditbox", 100, 50, 300, 200)
' // Set the focus inthe control
DIM hRichEdit AS HWND = pRichEdit.hRichEdit
```

# CRichEditCtrl Properties

| Name       | Description |
| ---------- | ----------- |
| [AutoCorrectProc](#AutoCorrectProc) | Gets/sets a pointer to the application-defined AutoCorrectProc callback function. |
| [AutoUrlDetect](#AutoUrlDetect) | Gets/sets whether the auto URL detection is turned on in the rich edit control. |
| [BidiOptions](#BidiOptions) | Gets/sets the current state of the bidirectional options in the rich edit control. |
| [CharFormat](#CharFormat) | Gets/sets the current character formatting in a rich edit control. |
| [CTFModeBias](#CTFModeBias) | Gets/sets the Text Services Framework mode bias values for a Microsoft Rich Edit control. |
| [CTFOpenStatus](#CTFOpenStatus) | Gets/sets if the Text Services Framework (TSF) keyboard is open or closed. |
| [EditStyle](#EditStyle) | Gets/sets the current edit style flags. |
| [EditStyleEx](#EditStyleEx) | Gets/sets the extended edit style flags. |
| [EllipsisMode](#EllipsisMode) | Gets/sets the current ellipsis mode. |
| [EventMask](#EventMask) | Gets/sets the event mask for a rich edit control. The event mask specifies which notification messages the control sends to its parent window. |
| [HyphenateInfo](#HyphenateInfo) | Gets/sets information about hyphenation for a Microsoft Rich Edit control. |
| [IMEColor](#IMEColor) | Gets/sets the Input Method Editor (IME) composition color. This message is available only in Asian-language versions of the operating system. |
| [IMEModeBias](#IMEModeBias) | Gets/sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control. |
| [IMEOptions](#IMEOptions) | Gets/sets the current Input Method Editor (IME) options. This message is available only in Asian-language versions of the operating system. |
| [LangOptions](#LangOptions) | Gets/sets a rich edit control's option settings for Input Method Editor (IME) and Asian language support. |
| [LimitText](#LimitText) | Gets/sets the current text limit for a rich edit control. |
| [Modify](#Modify) | Gets/sets the state of a rich edit control's modification flag. The flag indicates whether the contents of the rich edit control have been modified. |
| [Options](#Options) | Gets/sets the options for a rich edit control. |
| [PageRotate](#PageRotate) | Deprecated. Gets/sets the text layout for a Microsoft Rich Edit control. |
| [ParaFormat](#ParaFormat) | Gets/sets the paragraph formatting of the current selection in a rich edit control. |
| [PasswordChar](#PasswordChar) | Gets/sets the password character that a rich edit control displays when the user enters text. |
| [Punctuation](#Punctuation) | Gets/sets the current punctuation characters for the rich edit control. |
| [Rect](#Rect) | Gets/sets the formatting rectangle of a rich edit control. |
| [RectNP](#RectNP) | Sets the formatting rectangle of a multiline rich edit control. |
| [ScrollPos](#ScrollPos) | Gets/sets the current scroll position of the edit control. |
| [StoryType](#StoryType) | Gets/sets the story type. |
| [Text](#Text) | Gets/sets the text from a rich edit control. |
| [Text](#Text) | Gets/sets the text of an edit control. |
| [TextMode](#TextMode) | Gets/sets the current text mode and undo level of a rich edit control. |
| [TouchOptions](#TouchOptions) | Gets/sets the touch options that are associated with a rich edit control. |
| [TypographyOptions](#TypographyOptions) | Gets/sets the current state of the typography options of a rich edit control. |

# CRichEditCtrl Methods

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

# <a name="AutoCorrectProc"></a>AutoCorrectProc

Gets/sets a pointer to the application-defined [AutoCorrectProc](https://learn.microsoft.com/en-us/windows/win32/api/richedit/nc-richedit-autocorrectproc) callback function.
```
(GET) PROPERTY AutoCorrectProc () AS LONG_PTR
(SET) PROPERTY AutoCorrectProc (BYVAL pfn AS LONG_PTR)
```
| Parameter  | Description |
| ---------- | ----------- |
| *pfn* | Pointer to an [AutoCorrectProc](https://learn.microsoft.com/en-us/windows/win32/api/richedit/nc-richedit-autocorrectproc) function. |

#### Return value

(GET) A pointer to the application-defined [AutoCorrectProc](https://learn.microsoft.com/en-us/windows/win32/api/richedit/nc-richedit-autocorrectproc) callback function.

(SET) If the operation succeeds, the return value is zero. If the operation fails, the return value is a nonzero value. Call **GetLastResult** and/or **GetErrorInfo** to get information about the result.

# <a name="AutoUrlDetect"></a>AutoUrlDetect

Gets/sets whether the auto URL detection is turned on in the rich edit control.
```
(GET) PROPERTY AutoUrlDetect () AS LONG
(SET) PROPERTY AutoUrlDetect (BYVAL fUrlDetect AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *fUrlDetect* | Specify 0 to disable automatic link detection, or one of the following values to enable various kinds of detection. |

| fUrlDetect value  | Description |
| --------------- | ----------- |
| **AURL_DISABLEMIXEDLGC** | **Windows 8**: Disable recognition of domain names that contain labels with characters belonging to more than one of the following scripts: Latin, Greek, and Cyrillic. |
| **AURL_ENABLEDRIVELETTERS** | **Windows 8**: Recognize file names that have a leading drive specification, such as c:\temp. |
| **AURL_ENABLEEA** | This value is deprecated; use **AURL_ENABLEEAURLS** instead. |
| **AURL_ENABLEEAURLS** | Recognize URLs that contain East Asian characters. |
| **AURL_ENABLEEMAILADDR** | **Windows 8**: Recognize email addresses. |
| **AURL_ENABLETELNO** | **Windows 8**: Recognize telephone numbers. |
| **AURL_ENABLEUR** | **Windows 8**: Recognize URLs that include the path. |

#### Return value

(GET) If auto-URL detection is active, the return value is 1. If auto-URL detection is inactive, the return value is 0.

(SET) If the message succeeds, the return value is zero. If the message fails, the return value is a nonzero value. For example, the message might fail due to insufficient memory or an invalid detection option. Call **GetLastResult** and/or **GetErrorInfo** to get information about the result.

#### Remarks

If automatic URL detection is enabled (that is, *fUrlDetect* includes **AURL_ENABLEURL**), the rich edit control scans any modified text to determine whether the text matches the format of a URL (or more generally in Windows 8 or later an IRI International Resource Identifier). The control detects URLs that begin with the following scheme names:

- callto
- file
- ftp
- gopher
- http
- https
- mailto
- news
- notes
- nntp
- onenote
- outlook
- prospero
- tel
- telnet
- wais
- webcal

When automatic link detection is enabled, the rich edit control removes the **CFE_LINK** effect from modified text that does not have a format recognized by the control. If your application uses the **CFE_LINK** effect to mark other types of text, do not enable automatic link detection. The rich edit control does not check whether a detected link exists; that responsibility belongs to the client.

A rich edit control sends the [EN_LINK](https://learn.microsoft.com/en-us/windows/win32/controls/en-link) notification when it receives various messages while the mouse pointer is over text that has the **CFE_LINK** effect. 

# <a name="BidiOptions"></a>BidiOptions

Gets/sets the current state of the bidirectional options in the rich edit control.
```
(GET) PROPERTY BidiOptions () AS .BIDIOPTIONS
(SET) PROPERTY BidiOptions (BYREF _Options AS .BIDIOPTIONS)
```
| Parameter  | Description |
| ---------- | ----------- |
| *_Options* | A [BIDIOPTIONS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-bidioptions) structure that indicates how to set the state of the bidirectional options in the rich edit control. |

#### Return value

(GET) A [BIDIOPTIONS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-bidioptions) structure with the current state of the bidirectional options in the rich edit control. The values of the **wMask** and **wEffects** contain the current state of the bidirectional options in the rich edit control.

(SET) This property does not return a result.

#### Remarks (SET property)

The rich edit control must be in plain text mode or **BidiOptions** will not do anything.

In plain text controls, **BidiOptions** automatically determines the paragraph direction and/or alignment based on the context rules. These rules state that the direction and/or alignment is derived from the first strong character in the control. A strong character is one from which text direction can be determined (see Unicode Standard version 2.0). The paragraph direction and/or alignment is applied to the default format.

**BidiOptions** only switches the default paragraph format to RTL (right to left) if it finds an RTL character.


# <a name="CharFormat"></a>CharFormat

Gets/sets the current character formatting in a rich edit control.
```
(GET) PROPERTY CharFormat (BYVAL fOption AS DWORD) AS CHARFORMATW
(SET) PROPERTY CharFormat (BYVAL chfmt AS DWORD, BYREF cf AS CHARFORMATW)
```
| Parameter  | Description |
| ---------- | ----------- |
| *fOption* | (GET) Specifies the range of text from which to retrieve formatting. It can be one of the following values.<br>**SCF_DEFAULT** The default character formatting.<br>**SCF_SELECTION** The current selection's character formatting. |
| *chfmt* | (SET) Character formatting that applies to the control. If this parameter is zero, the default character format is set. Otherwise, it can be one of the following values (see Formatting values below). |
| *cf* | (SET) A [CHARFORMATW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformatw) structure specifying the character formatting to use. Only the formatting attributes specified by the **dwMask** member are changed. The **szFaceName** and **bCharSet** members may be overruled when invalid for characters, for example: Arial on kanji characters. |

| Formatting value  | Meaning |
| ----------------- | ------- |
| **SCF_ALL** | Applies the formatting to all text in the control. Not valid with **SCF_SELECTION** or **SCF_WORD**. |
| **SCF_ASSOCIATEFONT** | **RichEdit 4.1**: Associates a font to a given script, thus changing the default font for that script. To specify the font, use the following members of **CHARFORMAT2**: yHeight, bCharSet, bPitchAndFamily, szFaceName, and lcid. |
| **SCF_ASSOCIATEFONT2** | **RichEdit 4.1**: Associates a surrogate (plane-2) font to a given script, thus changing the default font for that script. To specify the font, use the following members of **CHARFORMAT2**: yHeight, bCharSet, bPitchAndFamily, szFaceName, and lcid. |
| **SCF_CHARREPFROMLCID** | Gets the character repertoire from the LCID. |
| **SCF_DEFAULT** | **RichEdit 4.1**: Sets the default font for the control. |
| **SPF_DONTSETDEFAULT** | Prevents setting the default paragraph format when the rich edit control is empty. |
| **SCF_NOKBUPDATE** | **RichEdit 4.1**: Prevents keyboard switching to match the font. For example, if an Arabic font is set, normally the automatic keyboard feature for Bidi languages changes the keyboard to an Arabic keyboard. |
| **SCF_SELECTION** | Applies the formatting to the current selection. If the selection is empty, the character formatting is applied to the insertion point, and the new character format is in effect only until the insertion point changes. |
| **SPF_SETDEFAULT** | Sets the default paragraph formatting attributes. |
| **SCF_SMARTFONT** | Apply the font only if it can handle script. |
| **SCF_USEUIRULES** | **RichEdit 4.1**: Used with **SCF_SELECTION**. Indicates that format came from a toolbar or other UI tool, so UI formatting rules should be used instead of literal formatting. |
| **SCF_WORD** | Applies the formatting to the selected word or words. If the selection is empty but the insertion point is inside a word, the formatting is applied to the word. The **SCF_WORD** value must be used in conjunction with the **SCF_SELECTION** value. |

#### Return value

(GET) Returns the value of the **dwMask** member of the [CHARFORMATW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformatw) structure.

(SET) If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero. Call **GetLastResult** and/or **GetErrorInfo** to get information about the result.

# <a name="CTFModeBias"></a>CTFModeBias

Gets/sets the Text Services Framework mode bias values for a Microsoft Rich Edit control.
```
(GET) PROPERTY CTFModeBias () AS LONG
(SET) PROPERTY CTFModeBias (BYVAL nModeBias AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *nModeBias* | (SET) Mode bias value. This can be one of the following values below. |

| Mode bias value  | Meaning |
| ---------------- | ------- |
| **CTFMODEBIAS_DEFAULT** | There is no mode bias. |
| **CTFMODEBIAS_FILENAME** | The bias is to a filename. |
| **CTFMODEBIAS_NAME** | The bias is to a name. |
| **CTFMODEBIAS_READING** | The bias is to the reading. |
| **CTFMODEBIAS_DATETIME** | The bias is to a date or time. |
| **CTFMODEBIAS_CONVERSATION** | The bias is to a conversation. |
| **CTFMODEBIAS_NUMERIC** | The bias is to a number. |
| **CTFMODEBIAS_HIRAGANA** | The bias is to hiragana strings. |
| **CTFMODEBIAS_KATAKANA** | The bias is to katakana strings. |
| **CTFMODEBIAS_HANGUL** | The bias is to Hangul characters. |
| **CTFMODEBIAS_HALFWIDTHKATAKANA** | The bias is to half-width katakana strings. |
| **CTFMODEBIAS_FULLWIDTHALPHANUMERIC** | The bias is to full-width alphanumeric characters. |
| **CTFMODEBIAS_HALFWIDTHALPHANUMERIC** | The bias is to half-width alphanumeric characters. |

#### Return value

(GET) The current Text Services Framework mode bias value.

(SET) If successful, the return value is the new TSF mode bias value. If unsuccessful, the return value is the old TSF mode bias value. Call the (GET) **CTFModeBias** property to check if the value has changed.

#### Remarks

When a Microsoft Rich Edit application uses TSF, it can select the TSF mode bias. This message sets the criteria by which an alternative choice appears at the top of the list for selection.

# <a name="CTFOpenStatus"></a>CTFOpenStatus

Gets/sets if the Text Services Framework (TSF) keyboard is open or closed.
```
(GET) PROPERTY CTFOpenStatus () AS BOOLEAN
(SET) PROPERTY CTFOpenStatus (BYVAL fTSFkbd AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *fTSFkbd* | To turn on the TSF keyboard, use **TRUE**. To turn off the TSF keyboard, use **FALSE**. |

#### Return value

(GET) If the TSF keyboard is open, the return value is **TRUE**. Otherwise, it is **FALSE**.

(SET) If successful, it returns **TRUE**. If unsuccessful, it returns **FALSE**. Call the (GET) **CTFOpenStatus** property to check if the value has changed.

# <a name="EditStyle"></a>EditStyle

Gets/sets the current edit style flags.
```
PROPERTY EditStyle () AS DWORD
PROPERTY EditStyle (BYVAL fStyle AS LONG, BYVAL fMask AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *fStyle* | Specifies one or more edit style flags. For a list of possible values, see the table below. |
| *fMask* | A mask consisting of one or more of the *fStyle* values. Only the values specified in this mask will be set or cleared. This allows a single flag to be set or cleared without reading the current flag states. |

| Edit Style Flag  | Description |
| ---------------- | ----------- |
| **SES_BEEPONMAXTEXT** | Rich Edit will call the system beeper if the user attempts to enter more than the maximum characters. |
| **SES_BIDI** | Turns on bidirectional processing. This is automatically turned on by Rich Edit if any of the following window styles are active: WS_EX_RIGHT, WS_EX_RTLREADING, WS_EX_LEFTSCROLLBAR. However, this setting is useful for handling these window styles when using a custom implementation of ITextHost (default: 0). |
| **SES_CTFALLOWEMBED** | **Windows XP with SP1**: Allow embedded objects to be inserted using TSF (default: 0). |
| **SES_CTFALLOWPROOFING** | **Windows XP with SP1**: Allows TSF proofing tips (default: 0). |
| **SES_CTFALLOWSMARTTAG** | **Windows XP with SP1**: Allows TSF SmartTag tips (default: 0). |
| **SES_CTFNOLOCK** | **Windows 8**: Do not allow the TSF lock read/write access. This pauses TSF input. |
| **SES_DEFAULTLATINLIGA** | **Windows 8**: Fonts with an fi ligature are displayed with default OpenType features resulting in improved typography (default: 0). |
| **SES_DRAFTMODE** | **Windows XP with SP1**: Use draft mode fonts to display text. Draft mode is an accessibility option where the control displays the text with a single font; the font is determined by the system setting for the font used in message boxes. For example, accessible users may read text easier if it is uniform, rather than a mix of fonts and styles (default: 0). |
| **SES_EMULATE10** | **Windows 8**: Emulate RichEdit 1.0 behavior.<br>**Note**: If you really want this behavior, use the Windows riched32.dll instead of riched20.dll or msftedit.dll. Riched32.dll had more functionality. |
| **SES_EMULATESYSEDIT** | When this bit is on, rich edit attempts to emulate the system edit control (default: 0). |
| **SES_EXTENDBACKCOLOR** | Extends the background color all the way to the edges of the client rectangle (default: 0). |
| **SES_HIDEGRIDLINES** | **Windows XP with SP1**: If the width of table gridlines is zero, gridlines are not displayed. This is equivalent to the hide gridlines feature in Word's table menu (default: 0). |
| **SES_HYPERLINKTOOLTIPS** | **Windows 8**: When the cursor is over a link, display a tooltip with the target link address (default: 0). |
| **SES_LOGICALCARET** | **Windows 8**: Provide logical caret information instead of a caret bitmap as described in [ITextHost::TxSetCaretPos]([url](https://learn.microsoft.com/en-us/windows/win32/api/textserv/nf-textserv-itexthost-txsetcaretpos))(default: 0). |
| **SES_LOWERCASE** | Converts all input characters to lowercase (default: 0). |
| **SES_MAPCPS** | Obsolete. Do not use. |
| **SES_MULTISELECT** | **Windows 8**: Enable multiselection with individual mouse selections made while the Ctrl key is pressed (default: 0). |
| **SES_NOEALINEHEIGHTADJUST** | **Windows 8**: Do not adjust line height for East Asian text (default: 0 which adjusts the line height by 15%). |
| **SES_NOFOCUSLINKNOTIFY** | Sends EN_LINK notification from links that do not have focus. |
| **SES_NOIME** | Disallows IMEs for this instance of the rich edit control (default: 0). |
| **SES_NOINPUTSEQUENCECHK** | When this bit is on, rich edit does not verify the sequence of typed text. Some languages (such as Thai and Vietnamese) require verifying the input sequence order before submitting it to the backing store (default: 0). |
| **SES_SCROLLONKILLFOCUS** | When KillFocus occurs, scroll to the beginning of the text (character position equal to 0) (default: 0). |
| **SES_SMARTDRAGDROP** | **Windows 8**: Add or delete a space according to the context when dropping text (default: 0). |
| **SES_USECRLF** | Obsolete. Do not use. |
| **SES_WORDDRAGDROP** | **Windows 8**: If word select is active, ensure that the drop location is at a word boundary (default: 0). |
| **SES_UPPERCASE** | Converts all input characters to uppercase (default: 0). |
| **SES_USEAIMM** | Uses the Active IMM input method component that ships with Internet Explorer 4.0 or later (default: 0). |
| **SES_USEATFONT** | **Windows XP with SP1**: Uses an @ font, which is designed for vertical text; this is used with the ES_VERTICAL window style. The name of an @ font begins with the @ symbol, for example, "@Batang" (default: 0, but is automatically turned on for vertical text layout). |
| **SES_USECTF** | **Windows XP with SP1**: Turns on TSF support. (default: 0). |
| **SES_XLTCRCRLFTOCR** | Turns on translation of CRCRLFs to CRs. When this bit is on and a file is read in, all instances of CRCRLF will be converted to hard CRs internally. This will affect the text wrapping. Note that if such a file is saved as plain text, the CRs will be replaced by CRLFs. This is the .txt standard for plain text (default: 0, which deletes CRCRLFs on input). |

#### Return value

(GET) Returns the current edit style flags, which can include one or more of the following values (see table above).

(SET) The return value is the state of the edit style flags after the rich edit control has attempted to implement your edit style changes. The edit style flags are a set of flags that indicate the current edit style. Call the (GET) **EditStyle** property to check if the value has changed.

# <a name="EditStyleEx"></a>EditStyleEx

Gets/sets the extended edit style flags.
```
(GET) PROPERTY EditStyleEx () AS DWORD
(SET) PROPERTY EditStyleEx (BYVAL fStyle AS LONG, BYVAL fMask AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *fStyle* | Specifies one or more edit style flags. For a list of possible values, see table below. |
| *fMask* | A mask consisting of one or more of the *fStyle* values. Only the values specified in this mask will be set or cleared. This allows a single flag to be set or cleared without reading the current flag states. |

| Edit style flag | Description |
| --------------- | ----------- |
| **SES_EX_HANDLEFRIENDLYURL** | Display friendly name links with the same text color and underlining as automatic links, provided that temporary formatting isn't used or uses text autocolor (default: 0). |
| **SES_EX_MULTITOUCH** | Enable touch support in Rich Edit. This includes selection, caret placement, and context-menu invocation. When this flag is not set, touch is emulated by mouse commands, which do not take touch-mode specifics into account (default: 0). |
| **SES_EX_NOACETATESELECTION** | Display selected text using classic Windows selection text and background colors instead of background acetate color (default: 0). |
| **SES_EX_NOMATH** | Disable insertion of math zones (default: 1). To enable math editing and display, send the **RichEdit_SetEditStyleEx** message with *fStyle* set to 0, and *fMask* set to SES_EX_NOMATH. |
| **SES_EX_NOTABLE** | Disable insertion of tables. The **RichEdit_InsertTable** message returns **E_FAIL** and RTF tables are skipped (default: 0). |
| **SES_EX_USESINGLELINE** | Enable a multiline control to act like a single-line control with the ability to scroll vertically when the single-line height is greater than the window height (default: 0). |
| **SES_HIDETEMPFORMAT** | Hide temporary formatting that is created when **ITextFont.Reset** is called with **tomApplyTmp**. For example, such formatting is used by spell checkers to display a squiggly underline under possibly misspelled words. |
| **SES_EX_USEMOUSEWPARAM** | Use *wParam* when handling the **WM_MOUSEMOVE** message and do not call **GetAsyncKeyState**. |

#### Return value

(GET) Returns the state of the edit style flags.

(SET) The return value is the state of the edit style flags after the rich edit control has attempted to implement your edit style changes. The edit style flags are a set of flags that indicate the current edit style. Call the (GET) **EditStyleEx** property to check if the value has changed.


# <a name="EllipsisMode"></a>EllipsisMode

Gets/sets the current ellipsis mode. When enabled, an ellipsis ( ) is displayed for text that doesn't fit in the display window. The ellipsis is only used when the control isn't active. When active, scroll bars are used to reveal text that doesn't fit into the display window.
```
(GET) PROPERTY EllipsisMode () AS DWORD
(SET) PROPERTY EllipsisMode (BYVAL fMode AS DWORD)
```
| Parameter  | Description |
| ---------- | ----------- |
| *fMode* | (SET) One of the values listed below. |

| Value  | Meaning |
| ------ | ----------- |
| **ELLIPSIS_NONE** | No ellipsis is used. |
| **ELLIPSIS_END** | Ellipsis at the end (forced break). |
| **ELLIPSIS_WORD** | Ellipsis at the end (word break). |

The bits for these values all fit in the **ELLIPSIS_MASK**.

# <a name="EventMask"></a>EventMask

Gets/sets the event mask for a rich edit control. The event mask specifies which notification messages the control sends to its parent window.
```
(GET) PROPERTY EventMask () AS DWORD
(SET) PROPERTY EventMask (BYVAL fMask AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *fMask* | New event mask for the rich edit control. For a list of event masks, see the following table. |

| Flag  | Description |
| ----- | ----------- |
| **ENM_CHANGE** | Sends **EN_CHANGE** notifications. |
| **ENM_CLIPFORMAT** | Sends **EN_CLIPFORMAT** notifications. |
| **ENM_CORRECTTEXT** | Sends **EN_CORRECTTEXT** notifications. |
| **ENM_DRAGDROPDONE** | Sends **EN_DRAGDROPDONE** notifications. |
| **ENM_DROPFILES** | Sends **EN_DROPFILES** notifications. |
| **ENM_IMECHANGE** | **Microsoft Rich Edit 1.0 only**: Sends **EN_IMECHANGE** notifications when the IME conversion status has changed. Only for Asian-language versions of the operating system. |
| **ENM_KEYEVENTS** | Sends **EN_MSGFILTER** notifications for keyboard events. |
| **ENM_LINK** | **Rich Edit 2.0 and later**: Sends **EN_LINK** notifications when the mouse pointer is over text that has the CFE_LINK and one of several mouse actions is performed. |
| **ENM_LOWFIRTF** | Sends **EN_LOWFIRTF** notifications. |
| **ENM_MOUSEEVENTS** | Sends **EN_MSGFILTER** notifications for mouse events. |
| **ENM_OBJECTPOSITIONS** | Sends **EN_OBJECTPOSITIONS** notifications. |
| **ENM_PARAGRAPHEXPANDED** | Sends **EN_PARAGRAPHEXPANDED** notifications. |
| **ENM_PROTECTED** | Sends **EN_PROTECTED** notifications. |
| **ENM_REQUESTRESIZE** | Sends **EN_REQUESTRESIZE** notifications. |
| **ENM_SCROLL** | Sends **EN_HSCROLL** and **EN_VSCROLL** notifications. |
| **ENM_SCROLLEVENTS** | Sends **EN_MSGFILTER** notifications for mouse wheel events. |
| **ENM_SELCHANGE** | Sends **EN_SELCHANGE** notifications. |
| **ENM_UPDATE** | Sends **EN_UPDATE** notifications.<br>**Rich Edit 2.0 and later**: this flag is ignored and the **EN_UPDATE** notifications are always sent. However, if Rich Edit 3.0 emulates Microsoft Rich Edit 1.0, you must use this flag to send **EN_UPDATE** notifications. |

#### Return value

(GET)The event mask for this rich edit control.

(SET) Returns the previous event mask. Call the (GET) **EventMask** property to check if the value has changed.

#### Remarks

The default event mask is **ENM_NONE** in which case no notifications are sent to the parent window.

# <a name="HyphenateInfo"></a>HyphenateInfo

Gets/sets information about hyphenation for a Microsoft Rich Edit control.
```
(GET) PROPERTY HyphenateInfo () AS .HYPHENATEINFO
(SET) PROPERTY HyphenateInfo (BYREF hinfo AS .HYPHENATEINFO)
```
| Parameter  | Description |
| ---------- | ----------- |
| *hinfo* | (SET) A [HYPHENATEINFO](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-hyphenateinfo) structure. |

#### Return value

(GET) A [HYPHENATEINFO](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-hyphenateinfo) structure.

# <a name="IMEColor"></a>IMEColor

Gets/sets the Input Method Editor (IME) composition color. This message is available only in Asian-language versions of the operating system.
```
PROPERTY IMEColor () AS .COMPCOLOR
PROPERTY IMEColor (BYREF cmpcolor AS .COMPCOLOR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cmpcolor* | (SET) A [COMPCOLOR](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-compcolor) structure that contains the composition color to be set. |

#### Return value

(GET) A four-element array of [COMPCOLOR](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-compcolor) structures.

(SET) If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero. Call **GetLastResult** and/or **GetErrorInfo** to get information about the result.

#### Note

This message is supported only in Asian-language versions of Microsoft Rich Edit 1.0. It is not supported in any later versions.

# <a name="IMEModeBias"></a>IMEModeBias

Gets/sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.
```
PROPERTY IMEModeBias () AS DWORD
PROPERTY IMEModeBias (BYVAL nModeBias AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *nModeBias* | Mode bias value. This can be one of the following values below. |

| Mode bias value  | Meaning |
| ---------------- | ------- |
| **CTFMODEBIAS_DEFAULT** | There is no mode bias. |
| **CTFMODEBIAS_FILENAME** | The bias is to a filename. |
| **CTFMODEBIAS_NAME** | The bias is to a name. |
| **CTFMODEBIAS_READING** | The bias is to the reading. |
| **CTFMODEBIAS_DATETIME** | The bias is to a date or time. |
| **CTFMODEBIAS_CONVERSATION** | The bias is to a conversation. |
| **CTFMODEBIAS_NUMERIC** | The bias is to a number. |
| **CTFMODEBIAS_HIRAGANA** | The bias is to hiragana strings. |
| **CTFMODEBIAS_KATAKANA** | The bias is to katakana strings. |
| **CTFMODEBIAS_HANGUL** | The bias is to Hangul characters. |
| **CTFMODEBIAS_HALFWIDTHKATAKANA** | The bias is to half-width katakana strings. |
| **CTFMODEBIAS_FULLWIDTHALPHANUMERIC** | The bias is to full-width alphanumeric characters. |
| **CTFMODEBIAS_HALFWIDTHALPHANUMERIC** | The bias is to half-width alphanumeric characters. |

#### Return value

(GET) This message returns the current IME mode bias setting.

(SET) If successful, the return value is the new TSF mode bias value. If unsuccessful, the return value is the old TSF mode bias value. Call the (GET) **IMEModeBias** property to check if the value has changed.

#### Remarks

(GET) To get the Text Services Framework mode bias, usethe **CTFModeBias** property.

(SET) When a Microsoft Rich Edit application uses TSF, it can select the TSF mode bias. This message sets the criteria by which an alternative choice appears at the top of the list for selection.

The application should call he **IsIME** function before calling these properties.


# <a name="IMEOptions"></a>IMEOptions

Gets/sets the current Input Method Editor (IME) options. This message is available only in Asian-language versions of the operating system.
```
(GET) PROPERTY IMEOptions () AS DWORD
(SET) PROPERTY IMEOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *fCoop* | (SET) Specifies one of the following values.<br>**ECOOP_SET**. Sets the options to those specified by *fOptions*.<br>**COOP_OR**. Combines the specified options with the current options.<br>**ECOOP_AND**. Retains only those current options that are also specified by *fOptions*.<br>**ECOOP_XOR**. Logically exclusive OR the current options with those specified by *fOptions*. |
| *fOptions* | (SET) Specifies one of more of the following values.<br>**IMF_CLOSESTATUSWINDOW**. Closes the IME status window when the control receives the input focus.<br>**IMF_FORCEACTIVE**. Activates the IME when the control receives the input focus.<br>**IMF_FORCEDISABLE**. Disables the IME when the control receives the input focus.<br>**IMF_FORCEENABLE**. Enables the IME when the control receives the input focus.<br>**IMF_FORCEINACTIVE**. Inactivates the IME when the control receives the input focus.<br>**IMF_FORCENONE**. Disables IME handling.<br>**IMF_FORCEREMEMBER**. Restores the previous IME status when the control receives the input focus.<br>**IMF_MULTIPLEEDIT**. Specifies that the composition string will not be canceled or determined by focus changes. This allows an application to have separate composition strings on each rich edit control.<br>**IMF_VERTICAL**. Note: used in Rich Edit 2.0 and later. |

#### Return value

(GET) Returns one or more of the IME option flag values described in the *fOptions* parameter of the SET property.

(SET) If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero. Call the (GET) **ImeOptions** property to check if the value has changed.

# <a name="LangOptions"></a>LangOptions

Gets/sets a rich edit control's option settings for Input Method Editor (IME) and Asian language support.
```
(GET) PROPERTY LangOptions () AS DWORD
(SET) PROPERTY LangOptions (BYVAL lgoptions AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *lgoptions* | (SET) Specifies the language options. For a list of possible values, see table below. |

| Flag  | Description |
| ----- | ----------- |
| **IMF_AUTOFONT** | If this flag is set, the control automatically changes fonts when the user explicitly changes to a different keyboard layout. It is useful to turn off **IMF_AUTOFONT** for universal Unicode fonts. This option is turned on by default (1). |
| **IMF_AUTOFONTSIZEADJUST** | If this flag is set, the control scales font-bound font sizes from insertion point size according to script. For example, Asian fonts are slightly larger than Western ones. This option is turned on by default (1). |
| **IMF_AUTOKEYBOARD** | If this flag is set, the control automatically changes the keyboard layout when the user explicitly changes to a different font, or when the user explicitly changes the insertion point to a new location in the text. Will be turned on automatically for bidirectional controls. For all other controls, it is turned off by default. This option is turned off by default (0). |
| **IMF_DISABLEAUTOBIDIAUTOKEYBOARD** | **Windows 8**: If this flag is set, the control uses language neutral logic for automatic keyboard switching. This option is turned off by default (0). |
| **IMF_DUALFONT** | If this flag is set, the control uses dual-font mode. Used for Asian language support. The control uses an English font for ASCII text and a Asian font for Asian text. This option is turned on by default (1). |
| **IMF_IMEALWAYSSENDNOTIFY** | This flag controls how the rich edit control notifies the client during IME composition:<br>0: No **EN_CHANGE** or **EN_SELCHANGE** notifications during undetermined state. Send notification when the final string comes in. This is the default.<br>1: Send **EN_CHANGE** and **EN_SELCHANGE** events during undetermined state. |
| **IMF_IMECANCELCOMPLETE** | This flag determines how the control uses the composition string of an IME if the user cancels it. If this flag is set, the control discards the composition string. If this flag is not set, the control uses the composition string as the result string. This option is turned off by default (0). |
| **IMF_NOIMPLICITLANG** | **Windows 8**: If this flag is set, disable stamping keyboard input with the keyboard language and ensuring that non-East Asian language IDss are compatible with the character repertoire. This option is turned off by default (0). |
| **IMF_NOKBDLIDFIXUP** | **Windows 8**: If this flag is set, the rich edit control disables stamping keyboard language on an empty control. This option is turned off by default (0). |
| **IMF_SPELLCHECKING** | **Windows 8**: If this flag is set, the rich edit control turns on spell checking. This option is turned off by default (0). |
| **IMF_TKBAUTOCORRECTION** | **Windows 8**: If this flag is set, enable touch keyboard autocorrect. This option is turned off by default (0). |
| **IMF_TKBPREDICTION** | **Windows 10**: Ignored.<br>**Windows 8**: If this flag is set, the rich edit control enables touch keyboard prediction. This option is turned off by default (0). |
| **IMF_UIFONTS** | Use user-interface default fonts. This option is turned off by default (0). |

#### Return value

(GET) Returns the IME and Asian language settings, which can be zero or more of the above flags.

(SET) Returns a value of 1. Call the (GET) **LangOptions" property to check the value.

#### Remarks

The **IMF_AUTOFONT** flag is set by default. The **IMF_AUTOKEYBOARD** and **IMF_IMECANCELCOMPLETE** flags are cleared by default.

# <a name="LimitText"></a>LimitText

Gets/sets the current text limit for a rich edit control. The text limit is the maximum amount of text that the user can type into the edit control.
```
(GET) PROPERTY LimitText () AS LONG
(SET) PROPERTY LimitText (BYVAL chMax AS DWORD)
```
| Parameter  | Description |
| ---------- | ----------- |
| *chMax* | (SET) The maximum number of characters the user can enter. If this parameter is zero, the text length is set to 64,000 characters. |

#### Return value

(GET) The return value is the text limit.

(SET) The set property does not return a value.

#### Remarks

(SET) **LimitText** limits only the text the user can enter. It does not affect any text already in the edit control when the message is sent, nor does it affect the length of the text copied to the edit control by the **Text** property. If an application uses the **Text** property to place more text into an edit control than is specified by the **LimitText** property, the user can edit the entire contents of the edit control.

# <a name="Modify"></a>Modify

Gets/sets the state of a rich edit control's modification flag. The flag indicates whether the contents of the rich edit control have been modified.
```
(GET) PROPERTY Modify () AS LONG
(SET) PROPERTY Modify (BYVAL fModify AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *fModify* | (SET) The new value for the modification flag. A value of **TRUE** indicates the text has been modified, and a value of **FALSE** indicates it has not been modified. |

#### Return value

(GET) If the contents of edit control have been modified, the return value is nonzero; otherwise, it is zero.

(SET) The set property does not return a value.

### Remarks

The system automatically clears the modification flag to zero when the control is created. If the user changes the control's text, the system sets the flag to nonzero. You can use the (SET) **Modify** property to set or clear the flag.


# <a name="Options"></a>Options

Gets/sets the options for a rich edit control.
```
PROPERTY Options () AS DWORD
PROPERTY Options (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
```

# <a name="PageRotate"></a>PageRotate

Deprecated. Gets/sets the text layout for a Microsoft Rich Edit control.
```
PROPERTY PageRotate () AS DWORD
PROPERTY PageRotate (BYVAL txtlayout AS LONG)
```

# <a name="ParaFormat"></a>ParaFormat

Gets/sets the paragraph formatting of the current selection in a rich edit control.
```
PROPERTY ParaFormat () AS .PARAFORMAT
PROPERTY ParaFormat (BYREF pfmt AS .PARAFORMAT)
```

# <a name="PasswordChar"></a>PasswordChar

Gets/sets the password character that a rich edit control displays when the user enters text.
```
PROPERTY PasswordChar () AS LONG
PROPERTY PasswordChar (BYVAL dwchar AS DWORD)
```

# <a name="Punctuation"></a>Punctuation

Gets/sets the current punctuation characters for the rich edit control.
```
PROPERTY Punctuation (BYVAL punctp AS DWORD) AS .PUNCTUATION
PROPERTY Punctuation (BYVAL ptype AS LONG, BYREF punct AS .PUNCTUATION)
```

# <a name="Rect"></a>Rect

Gets/sets the formatting rectangle of a rich edit control.
```
PROPERTY Rect () AS .RECT
PROPERTY Rect (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
```

# <a name="RectNP"></a>RectNP

Sets the formatting rectangle of a multiline rich edit control.
```
PROPERTY RectNP (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
```

# <a name="ScrollPos"></a>ScrollPos

Gets/sets the current scroll position of the edit control.
```
PROPERTY ScrollPos () AS .POINT
PROPERTY ScrollPos (BYREF pt AS .POINT)
```

# <a name="StoryType"></a>StoryType

Gets/sets the story type.
```
PROPERTY StoryType (BYVAL Index AS DWORD) AS DWORD
PROPERTY StoryType (BYVAL Index AS LONG, BYVAL dwType AS DWORD)
```

# <a name="Text"></a>Text

Gets/sets the text from a rich edit control.
```
PROPERTY Text () AS CWSTR
DECLARE PROPERTY Text (BYVAL pwszText AS WSTRING PTR)
```

# <a name="TextMode"></a>TextMode

Gets/sets the current text mode and undo level of a rich edit control.
```
PROPERTY TextMode () AS DWORD
PROPERTY TextMode (BYVAL pvalues AS LONG)
```

# <a name="TouchOptions"></a>TouchOptions

Gets/sets the touch options that are associated with a rich edit control.
```
PROPERTY TouchOptions (BYVAL _Options AS LONG PTR) AS DWORD
PROPERTY TouchOptions (BYVAL _Options AS LONG, BYVAL fEnable AS LONG)
```

# <a name="TypographyOptions"></a>TypographyOptions

Gets/sets the current state of the typography options of a rich edit control.
```
PROPERTY TypographyOptions () AS DWORD
PROPERTY TypographyOptions (BYVAL pto AS LONG, BYVAL fMask AS LONG)
```

# <a name="CallAutocorrectProc"></a>CallAutocorrectProc

Calls the autocorrect callback function that is stored by the **CRichEditCtrl.SetAutocorrectProc** message, provided that the text preceding the insertion point is a candidate for autocorrection.
```
FUNCTION CallAutocorrectProc (BYVAL char AS WCHAR) AS LONG
```

# <a name="CanPaste"></a>CanPaste

Determines whether a rich edit control can paste a specified clipboard format.
```
FUNCTION CanPaste (BYVAL clipformat AS LONG) AS BOOLEAN
```

# <a name="CanRedo"></a>CanRedo

Determines whether there are any actions in the control redo queue.
```
FUNCTION CanRedo () AS BOOLEAN
```

# <a name="CanUndo"></a>CanUndo

Determines whether there are any actions in an edit control's undo queue.
```
FUNCTION CanUndo () AS BOOLEAN
```

# <a name="DisplayBand"></a>DisplayBand

Displays a portion of the contents of a rich edit control, as previously formatted for a device using the EM_FORMATRANGE message.
```
FUNCTION DisplayBand (BYREF rc AS .RECT) AS LONG
```

# <a name="EmptyUndoBuffer"></a>EmptyUndoBuffer

Resets the undo flag of a rich edit control. The undo flag is set whenever an operation within the rich edit control can be undone.
```
SUB EmptyUndoBuffer ()
```

# <a name="ExGetSel"></a>ExGetSel

Retrieves the starting and ending character positions of the selection in a rich edit control.
```
FUNCTION ExGetSel () AS CHARRANGE
```

# <a name="ExLimitText"></a>ExLimitText

Sets an upper limit to the amount of text the user can type or paste into a rich edit control.
```
SUB ExLimitText (BYVAL dwLimit AS DWORD)
```

# <a name="ExLineFromChar"></a>ExLineFromChar

Determines which line contains the specified character in a rich edit control.
```
FUNCTION ExLineFromChar (BYVAL iIndex AS LONG) AS LONG
```

# <a name="ExSetSel"></a>ExSetSel

Selects a range of characters and/or Component Object Model (COM) objects in a Microsoft Rich Edit control.
```
FUNCTION ExSetSel (BYREF cr AS CHARRANGE) AS DWORD
```

# <a name="FindText"></a>FindText

Finds text within a rich edit control.
```
FUNCTION FindText (BYVAL fOptions AS DWORD, BYREF ft AS FINDTEXTW) AS LONG
```

# <a name="FindTextEx"></a>FindTextEx

Finds text within a rich edit control.
```
FUNCTION FindTextEx (BYVAL fOptions AS DWORD, BYREF ftexw AS FINDTEXTEXW) AS LONG
```

# <a name="FindWordBreak"></a>FindWordBreak

Finds the next word break before or after the specified character position or retrieves information about the character at that position.
```
FUNCTION FindWordBreak (BYVAL fOperation AS DWORD, BYVAL dwStartPos AS DWORD) AS DWORD
```

# <a name="FormatRange"></a>FormatRange

Formats a range of text in a rich edit control for a specific device.
```
FUNCTION FormatRange (BYVAL fRender AS LONG, BYREF fr AS FORMATRANGE) AS DWORD
```

# <a name="GetCharFromPos"></a>GetCharFromPos

Gets information about the character closest to a specified point in the client area of an edit control.
```
FUNCTION GetCharFromPos (BYREF ptl AS .POINTL) AS LONG
```

# <a name="GetEllipsisState"></a>GetEllipsisState

Retrieves the current ellipsis state.
```
FUNCTION GetEllipsisState () AS BOOLEAN
```

# <a name="GetFirstVisibleLine"></a>GetFirstVisibleLine

Gets the zero-based index of the uppermost visible line in a multiline rich edit control.
```
FUNCTION GetFirstVisibleLine () AS LONG
```

# <a name="GetIMECompMode"></a>GetIMECompMode

Gets the current IME mode for a rich edit control.
```
FUNCTION GetIMECompMode () AS DWORD
```

# <a name="GetIMECompText"></a>GetIMECompText

Gets the Input Method Editor (IME) composition text.
```
FUNCTION GetIMECompText (BYREF ict AS .IMECOMPTEXT, BYVAL buffer AS ANY PTR) AS DWORD
```

# <a name="GetIMEProperty"></a>GetIMEProperty

Gets the property and capabilities of the Input Method Editor (IME) associated with the current input locale.
```
FUNCTION GetIMEProperty (BYVAL figp AS DWORD) AS DWORD
```

# <a name="GetLine"></a>GetLine

Copies a line of text from a rich edit control.
```
FUNCTION GetLine (BYVAL which AS DWORD) AS CWSTR
```

# <a name="GetLineCount"></a>GetLineCount

Gets the number of lines in a multiline rich edit control.
```
FUNCTION GetLineCount () AS LONG
```

# <a name="GetOleInterface"></a>GetOleInterface

Retrieves an IRichEditOle object that a client can use to access a rich edit control's Component Object Model (COM) functionality.
```
FUNCTION GetOleInterface () AS IRichEditOle PTR
```

# <a name="GetRedoName"></a>GetRedoName

Retrieves the type of the next action, if any, in the control's redo queue.
```
FUNCTION GetRedoName () AS LONG
```

# <a name="GetSel"></a>GetSel

Gets the starting and ending character positions of the current selection in a rich edit control.
```
FUNCTION GetSel (BYREF dwStartPos AS DWORD, BYREF dwEndPos AS DWORD) AS LONG
```

# <a name="GetSelText"></a>GetSelText

Retrieves the currently selected text in a rich edit control.
```
FUNCTION GetSelText () AS CWSTR
```

# <a name="GetTableParams"></a>GetTableParams

Retrieves the table parameters for a table row and the cell parameters for the specified number of cells.
```
FUNCTION GetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
```

# <a name="GetTextEx"></a>GetTextEx

Gets all of the text from the rich edit control in any particular code base you want.
```
FUNCTION GetTextEx (BYREF gtex AS GETTEXTEX, BYVAL buffer AS ANY PTR) AS DWORD
```

# <a name="GetTextLength"></a>GetTextLength

Retrieves the length of all text in a rich edit control.
```
FUNCTION GetTextLength () AS LONG
```

# <a name="GetTextLengthEx"></a>GetTextLengthEx

Calculates text length in various ways. It is usually called before creating a buffer to receive the text from the control.
```
FUNCTION GetTextLengthEx (BYREF gtex AS .GETTEXTLENGTHEX) AS LONG
```

# <a name="GetTextRange"></a>GetTextRange

Retrieves a specified range of characters from a rich edit control.
```
FUNCTION GetTextRange (BYREF trg AS .TEXTRANGEW) AS DWORD
```

# <a name="GetThumb"></a>GetThumb

Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline rich edit control.
```
FUNCTION GetThumb () AS LONG
```

# <a name="GetUndoName"></a>GetUndoName

Retrieves the type of the next undo action, if any.
```
FUNCTION GetUndoName () AS DWORD
```

# <a name="GetWordBreakProc"></a>GetWordBreakProc

Gets the address of the current Wordwrap function.
```
FUNCTION GetWordBreakProc () AS LONG_PTR
```

# <a name="GetWordBreakProcEx"></a>GetWordBreakProcEx

Retrieves the address of the currently registered extended word-break procedure.
```
FUNCTION GetWordBreakProcEx () AS LONG_PTR
```

# <a name="GetWordWrapMode"></a>GetWordWrapMode

Gets the current word wrap and word-break options for the rich edit control.
```
FUNCTION GetWordWrapMode () AS DWORD
```

# <a name="GetZoom"></a>GetZoom

Gets the current zoom ratio, which is always between 1/64 and 64.
```
FUNCTION GetZoom (BYVAL pzNum AS DWORD PTR, BYVAL pzDen AS DWORD PTR) AS LONG
```

# <a name="HideSelection"></a>HideSelection

Hides or shows the selection in a rich edit control.
```
SUB HideSelection (BYVAL fHide AS DWORD)
```

# <a name="InsertImage"></a>InsertImage

Replaces the selection with a blob that displays an image.
```
FUNCTION InsertImage (BYREF ip AS RICHEDIT_IMAGE_PARAMETERS) AS DWORD
```

# <a name="InsertTable"></a>InsertTable

Inserts one or more identical table rows with empty cells.
```
FUNCTION InsertTable (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
```

# <a name="IsIME"></a>IsIME

Determines if current input locale is an East Asian locale.
```
FUNCTION IsIME () AS LONG
```

# <a name="LineFromChar"></a>LineFromChar

Gets the index of the line that contains the specified character index in a multiline rich edit control.
```
FUNCTION LineFromChar (BYVAL index AS DWORD) AS LONG
```

# <a name="LineIndex"></a>LineIndex

Gets the character index of the first character of a specified line in a multiline rich edit control.
```
FUNCTION LineIndex (BYVAL nLine AS LONG) AS LONG
```

# <a name="LineLength"></a>LineLength

Retrieves the length, in characters, of a line in a rich edit control.
```
FUNCTION LineLength (BYVAL index AS DWORD) AS LONG
```

# <a name="LineScroll"></a>LineScroll

Scrolls the text in a multiline rich edit control.
```
FUNCTION LineScroll (BYVAL y AS LONG) AS LONG
```

# <a name="PasteSpecial"></a>PasteSpecial

Pastes a specific clipboard format in a rich edit control.
```
SUB PasteSpecial (BYVAL clpfmt AS DWORD, BYREF rps AS REPASTESPECIAL)
```

# <a name="PosFromChar"></a>PosFromChar

Retrieves the client area coordinates of a specified character in a rich edit control.
```
FUNCTION PosFromChar (BYVAL index as DWORD) AS .POINTL
```

# <a name="Reconversion"></a>Reconversion

Invokes the Input Method Editor (IME) reconversion dialog box.
```
SUB Reconversion ()
```

# <a name="Redo"></a>Redo

Redoes the next action in the control's redo queue.
```
FUNCTION Redo () AS LONG
```

# <a name="ReplaceSel"></a>ReplaceSel

Replaces the current selection in a rich edit control with the specified text.
```
SUB ReplaceSel (BYVAL bCanBeUndone AS LONG, BYVAL pwszText AS WSTRING PTR)
```

# <a name="RequestResize"></a>RequestResize

Forces a rich edit control to send an EN_REQUESTRESIZE notification message to its parent window.
```
SUB RequestResize ()
```

# <a name="Scroll"></a>Scroll

Scrolls the text vertically in a multiline rich edit control.
```
FUNCTION Scroll (BYVAL nAction AS LONG) AS LONG
```

# <a name="ScrollCaret"></a>ScrollCaret

Scrolls the caret into view in a rich edit control.
```
SUB ScrollCaret ()
```

# <a name="SelectionType"></a>SelectionType

Determines the selection type for a rich edit control.
```
FUNCTION SelectionType () AS LONG
```

# <a name="SetBkgndColor"></a>SetBkgndColor

Sets the background color for a rich edit control.
```
FUNCTION SetBkgndColor (BYVAL pSysColor AS DWORD, BYVAL pBkColor AS DWORD) AS DWORD
```

# <a name="SetFontSize"></a>SetFontSize

Sets the font size for the selected text.
```
FUNCTION SetFontSize (BYVAL ptsize AS LONG) AS LONG
```

# <a name="SetMargins"></a>SetMargins

Sets the widths of the left and right margins for a rich edit control. The message redraws the control to reflect the new margins.
```
SUB SetMargins (BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
```

# <a name="SetOleCallback"></a>SetOleCallback

Gives a rich edit control an IRichEditOleCallback object that the control uses to get OLE-related resources and information from the client.
```
FUNCTION SetOleCallback (BYVAL pCallback AS ANY PTR) AS LONG
```

# <a name="SetPalette"></a>SetPalette

Changes the palette that a rich edit control uses for its display window.
```
SUB SetPalette (BYVAL newPalette AS HPALETTE)
```

# <a name="SetReadOnly"></a>SetReadOnly

Changes the palette that a rich edit control uses for its display window.
```
FUNCTION SetReadOnly (BYVAL fReadOnly AS LONG) AS LONG
```

# <a name="SetSel"></a>SetSel

Selects a range of characters in a rich edit control.
```
SUB SetSel (BYVAL nStart AS LONG, BYVAL nEnd AS LONG)
```

# <a name="SetTableParams"></a>SetTableParams

Changes the parameters of rows in a table.
```
FUNCTION SetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
```

# <a name="SetTabStops"></a>SetTabStops

Sets the tab stops in a multiline rich edit control.
```
FUNCTION SetTabStops (BYVAL nTabs AS LONG, BYVAL rgTabStops AS LONG_PTR) AS LONG
```

# <a name="SetTargetDevice"></a>SetTargetDevice

Sets the target device and line width used for WYSIWYG formatting in a rich edit control.
```
FUNCTION SetTargetDevice (BYVAL hDC AS HDC, BYVAL lnwidth AS LONG) AS LONG
```

# <a name="SetTextExW"></a>SetTextExW

Combines the functionality of WM_SETTEXT and EM_REPLACESEL and adds the ability to set text using a code page and to use either Rich Text Format (RTF) rich text or plain text.
```
FUNCTION SetTextExW (BYREF stex AS SETTEXTEX, BYVAL pwszText AS WSTRING PTR) AS DWORD
```

# <a name="SetUIAName"></a>SetUIAName

Sets the maximum number of actions that can stored in the undo queue.
```
FUNCTION SetUIAName (BYVAL bstrName AS AFX_BSTR) AS DWORD
```

# <a name="SetUndoLimit"></a>SetUndoLimit

Sets the maximum number of actions that can stored in the undo queue.
```
FUNCTION SetUndoLimit (BYVAL maxactions AS DWORD) AS DWORD
```

# <a name="SetWordBreakProc"></a>SetWordBreakProc

Replaces a rich edit control's default Wordwrap function with an application-defined Wordwrap function.
```
SUB SetWordBreakProc (BYVAL pfn AS LONG_PTR)
```

# <a name="SetWordBreakProcEx"></a>SetWordBreakProcEx

Sets the extended word-break procedure.
```
FUNCTION SetWordBreakProcEx (BYVAL pfn AS LONG_PTR) AS LONG_PTR
```

# <a name="SetWordWrapMode"></a>SetWordWrapMode

Sets the word-wrapping and word-breaking options for the rich edit control.
```
FUNCTION SetWordWrapMode (BYVAL pvalues AS LONG) AS LONG
```

# <a name="SetZoom"></a>SetZoom

Sets the zoom ratio anywhere between 1/64 and 64.
```
FUNCTION SetZoom (BYVAL zNum AS DWORD, BYVAL zDen AS DWORD) AS LONG
```

# <a name="ShowScrollBar"></a>ShowScrollBar

Shows or hides one of the scroll bars in the Text Host window.
```
SUB ShowScrollBar (BYVAL nScrollBar AS DWORD, BYVAL fShow AS LONG)
```

# <a name="StopGroupTyping"></a>StopGroupTyping

Stops the control from collecting additional typing actions into the current undo action.
```
FUNCTION StopGroupTyping () AS DWORD
```

# <a name="StreamIn"></a>StreamIn

Replaces the contents of a rich edit control with a stream of data provided by an application defined EditStreamCallback callback function.
```
FUNCTION StreamIn (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
```

# <a name="StreamOut"></a>StreamOut

Causes a rich edit control to pass its contents to an application defined EditStreamCallback callback function.
```
FUNCTION StreamOut (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
```

# <a name="Undo"></a>Undo

This message undoes the last edit control operation in the control's undo queue.
```
FUNCTION Undo () AS LONG
```
