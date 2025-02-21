# CRichEditCtrl Class

Provides the functionality of the rich edit control.

A "rich edit control" is a window in which the user can enter and edit text. The text can be assigned character and paragraph formatting, and can include embedded OLE objects. Rich edit controls provide a programming interface for formatting text. However, an application must implement any user interface components necessary to make formatting operations available to the user.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTOR](#constructor) | Called when a class variable is created. |
| [DESTRUCTOR](#destructor) | Called automatically when a class variable goes out of scope or is destroyed. |
| [hRichEdit](#hrichedit) | Returns the handle of the rich edit control. |
| [ScalingRatio](#scalingratio) | Gets/sets the scaling ratio. |
| [SetWysiwygPrint](#setwysiwygprint) | Sets the target printer device and line width used for "what you see is what you get" (WYSIWYG) formatting in a rich edit control. |

# Methods inherited from CTextObjectBase

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#getlastresult) | Returns the last result code. |
| [SetResult](#setresult) | Sets the last result code. |
| [GetErrorInfo](#geterrorinfo) | Returns a localized description of the last result code. |

#### Helper procedure

| Name       | Description |
| ---------- | ----------- |
| [AfxCRichEditCtrlPtr](#afxcricheditctrlptr) | Overloaded function that retrieves a pointer to the `CRichEditCtrl`class from the handle of the rich edit control or from the handle of its parent window and the control's identifier. |

---

# CRichEditCtrl Properties

| Name       | Description |
| ---------- | ----------- |
| [AutoCorrectProc](#autocorrectproc) | Gets/sets a pointer to the application-defined **AutoCorrectProc** callback function. |
| [AutoUrlDetect](#autourldetect) | Gets/sets whether the auto URL detection is turned on in the rich edit control. |
| [BidiOptions](#bidioptions) | Gets/sets the current state of the bidirectional options in the rich edit control. |
| [CaretPos](#caretpos) | Gets/sets the caret position. |
| [CharFormat](#charformat) | Gets/sets the current character formatting in a rich edit control. |
| [CTFModeBias](#ctfmodebias) | Gets/sets the Text Services Framework mode bias values for a Microsoft Rich Edit control. |
| [CTFOpenStatus](#ctfopenstatus) | Gets/sets if the Text Services Framework (TSF) keyboard is open or closed. |
| [EditStyle](#editstyle) | Gets/sets the current edit style flags. |
| [EditStyleEx](#editstyleex) | Gets/sets the extended edit style flags. |
| [EllipsisMode](#ellipsismode) | Gets/sets the current ellipsis mode. |
| [EllipsisState](#ellipsisstate) | Gets the current ellipsis state. |
| [EventMask](#eventmask) | Gets/sets the event mask for a rich edit control. |
| [FirstVisibleLine](#firstvisibleline) | Gets the zero-based index of the uppermost visible line in a  rich edit control. |
| [HyphenateInfo](#hyphenateinfo) | Gets/sets information about hyphenation for a Microsoft Rich Edit control. |
| [IMEModeBias](#imemodebias) | Gets/sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control. |
| [IMEOptions](#imeoptions) | Gets/sets the current Input Method Editor (IME) options. This message is available only in Asian-language versions of the operating system. |
| [LangOptions](#langoptions) | Gets/sets a rich edit control's option settings for Input Method Editor (IME) and Asian language support. |
| [LimitText](#limittext) | Gets/sets the current text limit for a rich edit control. |
| [Margins](#margins) | Sets the width of the specified margin. |
| [Modify](#modify) | Gets/sets the state of a rich edit control's modification flag. The flag indicates whether the contents of the rich edit control have been modified. |
| [Options](#options) | Gets/sets the options for a rich edit control. |
| [PageRotate](#pagerotate) | Deprecated. Gets/sets the text layout for a Microsoft Rich Edit control. |
| [ParaFormat](#paraformat) | Gets/sets the paragraph formatting of the current selection in a rich edit control. |
| [PasswordChar](#passwordchar) | Gets/sets the password character that a rich edit control displays when the user enters text. |
| [Punctuation](#punctuation) | Gets/sets the current punctuation characters for the rich edit control. |
| [Rect](#rect) | Gets/sets the formatting rectangle of a rich edit control. |
| [RectNP](#rectnp) | Sets the formatting rectangle of a multiline rich edit control. |
| [RedoName](#redoname) | Retrieves the type of the next action, if any, in the control's redo queue. |
| [ScrollPos](#scrollpos) | Gets/sets the current scroll position of the edit control. |
| [StoryType](#storytype) | Gets/sets the story type. |
| [Text](#text) | Gets/sets the text from a rich edit control. |
| [TextColor](#textcolor) | Gets/sets the text color of the selected text or the word under the cursor. |
| [TextFontName](#textfontname) | Gets/sets the font name of the selected text or the word under the cursor. |
| [TextHeight](#textheight) | Gets/sets the text height of the selected text or the word under the cursor. |
| [TextLength](#textlength) | Retrieves the length of all text in a rich edit control. |
| [TextLengthEx](#textlengthex) | Calculates text length in various ways. It is usually called before creating a buffer to receive the text from the control. |
| [TextMode](#textmode) | Gets/sets the current text mode and undo level of a rich edit control. |
| [TextOffset](#textoffset) | Gets/sets the text offset of the selected text or the word under the cursor. |
| [TouchOptions](#touchoptions) | Gets/sets the touch options that are associated with a rich edit control. |
| [TypographyOptions](#typographyoptions) | Gets/sets the current state of the typography options of a rich edit control. |
| [UndoName](#undoname) | Retrieves the type of the next undo action, if any. |
| [WordBreakProc](#wordbreakproc) | Gets/sets the address of the currently registered word-break procedure. |
| [WordWrap](#wordwrap) | Enables/disables word wrap. |
| [WordWrapMode](#wordwrapmode) | Sets the word-wrapping and word-breaking options for the rich edit control. |

# CRichEditCtrl Methods

| Name       | Description |
| ---------- | ----------- |
| [AddLF/AddCR/AddCRLF](#addlfcr) | Inserts a line feed, a carriage return or a carriage return and line feed at the cursor position or at the end of the text. |
| [CallAutocorrectProc](#callautocorrectproc) | Calls the autocorrect callback function that is stored by the (SET) **AutocorrectProc** property, provided that the text preceding the insertion point is a candidate for autocorrection. |
| [CanPaste](#canpaste) | Determines whether a rich edit control can paste a specified clipboard format. |
| [CanRedo](#canredo) | Determines whether there are any actions in the control redo queue. |
| [CanUndo](#canundo) | Determines whether there are any actions in an edit control's undo queue. |
| [DisplayBand](#displayband) | Displays a portion of the contents of a rich edit control, as previously formatted for a device using the EM_FORMATRANGE message. |
| [EmptyUndoBuffer](#emptyundobuffer) | Resets the undo flag of a rich edit control. The undo flag is set whenever an operation within the rich edit control can be undone. |
| [ExGetSel](#exgetsel) | Retrieves the starting and ending character positions of the selection in a rich edit control. |
| [ExLimitText](#exlimittext) | Sets an upper limit to the amount of text the user can type or paste into a rich edit control. |
| [ExLineFromChar](#exlinefromchar) | Determines which line contains the specified character in a rich edit control. |
| [ExSetSel](#exsetsel) | Selects a range of characters and/or Component Object Model (COM) objects in a Microsoft Rich Edit control. |
| [FindText](#findtext) | Finds text within a rich edit control. |
| [FindTextEx](#findtextex) | Finds text within a rich edit control. |
| [FindWordBreak](#findwordbreak) | Finds the next word break before or after the specified character position or retrieves information about the character at that position. |
| [FormatRange](#formatrange) | Formats a range of text in a rich edit control for a specific device. |
| [GetCharFromPos](#getcharfrompos) | Gets information about the character closest to a specified point in the client area of an edit control. |
| [GetIMEColor](#getimecolor) | Retrieves the Input Method Editor (IME) composition color. |
| [GetIMECompMode](#getimecompmode) | Gets the current IME mode for a rich edit control. |
| [GetIMECompText](#getimecompText) | Gets the Input Method Editor (IME) composition text. |
| [GetIMEProperty](#getimeproperty) | Gets the property and capabilities of the Input Method Editor (IME) associated with the current input locale. |
| [GetLimitText](#getlimittext) | Gets the current text limit for a rich edit control. |
| [GetLine](#getline) | Copies a line of text from a rich edit control. |
| [GetLineCount](#getlinecount) | Gets the number of lines in a multiline rich edit control. |
| [GetOleInterface](#getoleinterface) | Retrieves an IRichEditOle object that a client can use to access a rich edit control's Component Object Model (COM) functionality. |
| [GetRedoName](#getredoname) | Retrieves the type of the next action, if any, in the control's redo queue. |
| [GetSel](#getsel) | Gets the starting and ending character positions of the current selection in a rich edit control. |
| [GetSelText](#getseltext) | Retrieves the currently selected text in a rich edit control. |
| [GetTableParams](#gettableparams) | Retrieves the table parameters for a table row and the cell parameters for the specified number of cells. |
| [GetTextEx](#gettextex) | Gets all of the text from the rich edit control in any particular code base you want. |
| [GetTextRange](#gettextrange) | Retrieves a specified range of characters from a rich edit control. |
| [GetThumb](#getthumb) | Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline rich edit control. |
| [GetZoom](#getzoom) | Gets the current zoom ratio, which is always between 1/64 and 64. |
| [HideSelection](#hideselection) | Hides or shows the selection in a rich edit control. |
| [InsertImage](#insertimage) | Replaces the selection with a blob that displays an image. |
| [InsertObject](#insertobject) | Inserts an image or a Ole object in the rich edit control. |
| [InsertTable](#inserttable) | Inserts one or more identical table rows with empty cells. |
| [IsIME](#isime) | Determines if current input locale is an East Asian locale. |
| [IsTextBold](#istextbold) | Checks if the selected text or the word under the cursor is bolded. |
| [IsTextItalic](#istextitalic) | Checks if the selected text or word under the cursor is italicised. |
| [IsTextStrikeOut](#istextstrikeout) | Checks if the selected text or word under the cursor is striked out. |
| [IsTextUnderline](#istextunderline) | Checks if the selected text or word under the cursor is underlined. |
| [LineFromChar](#linefromchar) | Gets the index of the line that contains the specified character index in a multiline rich edit control. |
| [LineIndex](#lineindex) | Gets the character index of the first character of a specified line in a multiline rich edit control. |
| [LineLength](#linelength) | Retrieves the length, in characters, of a line in a rich edit control. |
| [LineScroll](#linescroll) | Scrolls the text in a multiline rich edit control. |
| [PasteSpecial](#pastespecial) | Pastes a specific clipboard format in a rich edit control. |
| [PosFromChar](#posfromchar) | Retrieves the client area coordinates of a specified character in a rich edit control. |
| [Redo](#redo) | Redoes the next action in the control's redo queue. |
| [ReplaceSel](#replacesel) | Replaces the current selection in a rich edit control with the specified text. |
| [RequestResize](#requestresize) | Forces a rich edit control to send an **EN_REQUESTRESIZE** notification message to its parent window. |
| [Reconversion](#reconversion) | Invokes the Input Method Editor (IME) reconversion dialog box. |
| [Scroll](#scroll) | Scrolls the text vertically in a multiline rich edit control. |
| [ScrollCaret](#scrollcaret) | Scrolls the caret into view in a rich edit control. |
| [SelectionType](#selectiontype) | Determines the selection type for a rich edit control. |
| [SetBkgndColor](#setbkgndcolor) | Sets the background color for a rich edit control. |
| [SetFont](#setfont) | Sets the font used by a rich edit control. |
| [SetFontSize](#setfontsize) | Sets the font size for the selected text. |
| [SetIMEColor](#setimecolor) | Sets the Input Method Editor (IME) composition color. |
| [SetLimitText](#setlimittext) | Sets the current text limit for a rich edit control. |
| [SetOleCallback](#setolecallback) | Gives a rich edit control an **IRichEditOleCallback** object that the control uses to get OLE-related resources and information from the client. |
| [SetOptions](#setoptions) | Sets the options for a rich edit control. |
| [SetPalette](#setpalette) | Changes the palette that a rich edit control uses for its display window. |
| [SetReadOnly](#setreadonly) | Sets or removes the read-only style (ES_READONLY) of a rich edit control. |
| [SetRectNP](#setrectNP) | Sets the formatting rectangle of a rich edit control. |
| [SetSel](#setsel) | Selects a range of characters in a rich edit control. |
| [SetTableParams](#settableparams) | Changes the parameters of rows in a table. |
| [SetTabStops](#settabstops) | Sets the tab stops in a multiline rich edit control. |
| [SetTargetDevice](#settargetdevice) | Sets the target device and line width used for WYSIWYG formatting in a rich edit control. |
| [SetTextEx](#settextex) | Combines the functionality of WM_SETTEXT and EM_REPLACESEL and adds the ability to set text using a code page and to use either Rich Text Format (RTF) rich text or plain text. |
| [SetTextBold](#settextbold) | Sets the attribute of selected text or word to bold. |
| [SetTextItalic](#settextitalic) | Sets the attribute of selected text or word to italic. |
| [SetTextStrikeOut](#settextstrikeout) | Sets the attribute of selected text or word to strike out. |
| [SetTextUnderline](#settextunderline) | Sets the attribute of selected text or word to underline. |
| [SetUIAName](#setuianame) | Sets the maximum number of actions that can stored in the undo queue. |
| [SetUndoLimit](#setundolimit) | Sets the maximum number of actions that can stored in the undo queue. |
| [SetZoom](#setzoom) | Sets the zoom ratio anywhere between 1/64 and 64. |
| [ShowScrollBar](#showscrollbar) | Shows or hides one of the scroll bars in the Text Host window. |
| [StopGroupTyping](#stopgrouptyping) | Stops the control from collecting additional typing actions into the current undo action. |
| [StreamIn](#streamin) | Replaces the contents of a rich edit control with a stream of data provided by an application defined EditStreamCallback callback function. |
| [StreamOut](#streamout) | Causes a rich edit control to pass its contents to an application defined EditStreamCallback callback function. |
| [Undo](#undo) | This message undoes the last edit control operation in the control's undo queue. |


# CRichEdit File Operations Methods

| Name       | Description |
| ---------- | ----------- |
| [NewDoc](#newdoc) | Opens a new document. |
| [OpenDoc](#opendoc) | Opens a new document. |
| [SaveDoc](#savedoc) | Saves the document. |
| [AppendDocFile](#appenddocfile) | Appends the contents of the specified RTF file. |
| [GetRtf](#getrtf) | Retrieves formatted text from a rich edit control. |
| [InsertdocFile](#insertdocfile) | Inserts the contents of the specified RTF file. |
| [LoadRtfFromResource](#loadrtffromresource) | Loads a rich text resource file into a rich edit control. |
| [SaveRtf](#savertf) | Saves the contents of the rich edit control to a file in rtf format. |
| [SaveRtfNoObjs](#savertfnoobjs) | Saves the contents of the rich edit control to a file in rtf format with spaces in place of COM objects. |
| [SaveSelRtf](#saveselrtf) | Saves selection of the rich edit control to a file in rtf format. |
| [SaveText](#savetext) | Saves the contents of the rich edit control in text format. |
| [SaveSelText](#saveseltext) | Saves selection of the rich edit control in text format. |

# ITextDocument2 Methods

| Name       | Description |
| ---------- | ----------- |

---

# <a name="afxcricheditctrlptr"></a>AfxCRichEditCtrlPtr

Overloaded function that retrieves a pointer to the `CRichEditCtrl`class from the handle of the rich edit control or from the handle of its parent window and the control's identifier.
```
FUNCTION AfxCRichEditCtrlPtr OVERLOAD (BYVAL hRichEdit AS HWND) AS CRichEditCtrl PTR
FUNCTION AfxCRichEditCtrlPtr OVERLOAD (BYVAL hParent AS HWND, BYVAL cID AS LONG) AS CRichEditCtrl PTR
```
| Parameter | Description |
| --------- | ----------- |
| *hRichEdit* | Handle of the rich edit control. |

| Parameter | Description |
| --------- | ----------- |
| *hParent* | Handle of the parent window of the rich edit control. |
| *cID* | Identifier of the rich edit control. |

#### Return value

A pointer to the `CRichEditCtrl` class.

---

# <a name="constructor"></a>CONSTRUCTOR

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
DIM pRichEdit AS CRichEditCtrl = CRichEditCtrl(@pWindow, IDC_RICHEDIT, _
    "Rich Edit box", 100, 50, 300, 200)
```
#### Usage example (pointer syntax)
```
DIM pRichEdit AS CRichEditCtrl PTR = NEW CRichEditCtrl(@pWindow, IDC_RICHEDIT, _
    "Rich Edit box", 100, 50, 300, 200)

```

---

# <a name="destructor"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.
```
DESTRUCTOR CRichEditCtrl
```

---

# <a name="hrichedit"></a>hRichEdit

Returns the handle of the rich edit control.

```
FUNCTION CRichEditCtrl.hRichEdit () AS HWND
```
#### Usage example
```
' // Create an instance of the CRichEditCtrl class
DIM pRichEdit AS CRichEditCtrl = CRichEditCtrl(@pWindow, IDC_RICHEDIT, "RichEditbox", 100, 50, 300, 200)
' // Set the focus in the control
DIM hRichEdit AS HWND = pRichEdit.hRichEdit
```
---

# <a name="scalingratio"></a>ScalingRatio

Gets/sets the scaling ratio.

```
(GET) PROPERTY ScalingRatio () AS SINGLE
(SET) PROPERTY ScalingRatio (BYVAL Ratio AS SINGLE)
```
FUNCTION GetScalingRatio () AS SINGLE
SUB SetScalingRatio (BYVAL ratio AS SINGLE)
```
```
#### Remarks

The scaling ratio is used by the **InsertImage** and **InsertObject** methods to scale images according to the DPI settings of your computer. Its initial value is the scaling ratio used by its parent window, but you can change it with this property. Setting a value or 1 disables scaling.

#### Usage examples

```
DIM ratio AS LONG = pRichEdit.ScalingRatio
pRichEdit.ScalingRatio = ratio
```
```
DIM ratio AS LONG = pRichEdit.GetScalingRatio
pRichEdit.SetScalingRatio(ratio)
```

---

# <a name="setwysiwygprint"></a>SetWysiwygPrint

Sets the target printer device and line width used for "what you see is what you get" (WYSIWYG) formatting in a rich edit control.

```
FUNCTION SetWysiwygPrint (BYREF wszPrinterName AS WSTRING = "") AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name, e.g. "OKI B410d" or the "Microsoft Print to PDF" virtual printer. If wszPrinterName is empty, the control will use the default printer. |

#### Return value

A boolean true (-1) or false (0).

# <a name="autocorrectproc"></a>AutoCorrectProc

Gets/sets a pointer to the application-defined [AutoCorrectProc](https://learn.microsoft.com/en-us/windows/win32/api/richedit/nc-richedit-autocorrectproc) callback function.

```
(GET) PROPERTY AutoCorrectProc () AS LONG_PTR
(SET) PROPERTY AutoCorrectProc (BYVAL pfn AS LONG_PTR)
```
```
FUNCTION GetAutoCorrectProc () AS LONG_PTR
FUNCTION SetAutoCorrectProc (BYVAL pfn AS LONG_PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfn* | Pointer to the [AutoCorrectProc](https://learn.microsoft.com/en-us/windows/win32/api/richedit/nc-richedit-autocorrectproc) application-defined callback function. |

#### Return value

(GET) A pointer to the application-defined [AutoCorrectProc](https://learn.microsoft.com/en-us/windows/win32/api/richedit/nc-richedit-autocorrectproc) callback function.

(SET) If the operation succeeds, the return value is zero. If the operation fails, the return value is a nonzero value.

---

# <a name="autourldetect"></a>AutoUrlDetect

Gets/sets whether the auto URL detection is turned on in the rich edit control.

```
(GET) PROPERTY AutoUrlDetect () AS BOOLEAN
(SET) PROPERTY AutoUrlDetect (BYVAL fUrlDetect AS LONG)
```
```
FUNCTION GetAutoUrlDetect () AS BOOLEAN
FUNCTION SetAutoUrlDetect (BYVAL fUrlDetect AS LONG) AS HRESULT
```
```
FUNCTION DisableAutoUrlDetect () AS BOOLEAN
FUNCTION EnableAutoUrlDetect (BYVAL fUrlDetect AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *fUrlDetect* | (SET) Specify 0 to disable automatic link detection, or one of the following values to enable various kinds of detection. |

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

(GET) If auto-URL detection is active, the return value is true (-1). If auto-URL detection is inactive, the return value is false (0).

(SET) If the message succeeds, the return value is zero. If the message fails, the return value is a nonzero value. For example, the message might fail due to insufficient memory or an invalid detection option.

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

---

# <a name="bidioptions"></a>BidiOptions

Gets/sets the current state of the bidirectional options in the rich edit control.

```
(GET) PROPERTY BidiOptions () AS .BIDIOPTIONS
(SET) PROPERTY BidiOptions (BYREF opt AS .BIDIOPTIONS)
```
```
FUNCTION GetBidiOptions () AS .BIDIOPTIONS
SUB SetBidiOptions (BYREF opt AS .BIDIOPTIONS)
```

| Parameter  | Description |
| ---------- | ----------- |
| *opt* |(SET) A [BIDIOPTIONS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-bidioptions) structure that indicates how to set the state of the bidirectional options in the rich edit control. |

#### Return value

(GET) A [BIDIOPTIONS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-bidioptions) structure with the current state of the bidirectional options in the rich edit control. The values of the **wMask** and **wEffects** contain the current state of the bidirectional options in the rich edit control.

(SET) This property does not return a result.

#### Remarks (SET property)

The rich edit control must be in plain text mode or **BidiOptions** will not do anything.

In plain text controls, **BidiOptions** automatically determines the paragraph direction and/or alignment based on the context rules. These rules state that the direction and/or alignment is derived from the first strong character in the control. A strong character is one from which text direction can be determined (see Unicode Standard version 2.0). The paragraph direction and/or alignment is applied to the default format.

**BidiOptions** only switches the default paragraph format to RTL (right to left) if it finds an RTL character.

---

# <a name="caretpos"></a>CaretPos

Gets/sets the caret position.
```
(GET) PROPERTY CaretPos () AS DWORD
(SET) PROPERTY CaretPos (BYVAL dwPos AS DWORD)
```
```
FUNCTION GetCaretPos () AS DWORD
SUB SetCaretPos (BYVAL dwPos AS DWORD)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwPos* | (SET) The caret position. If you pass a value of -1, the caret will be positioned at the end of the document. |

#### Return value

(GET) The caret position.

(SET) This property does not return a result.

---

# <a name="charformat"></a>CharFormat

Gets/sets the current character formatting in a rich edit control.

```
(GET) PROPERTY CharFormat (BYVAL fOption AS DWORD) AS CHARFORMATW
(SET) PROPERTY CharFormat (BYVAL chfmt AS DWORD, BYREF cf AS CHARFORMATW)
```
```
FUNCTION GetCharFormat (BYVAL fOption AS DWORD) AS CHARFORMATW
FUNCTION SetCharFormat (BYVAL chfmt AS DWORD, BYREF cf AS CHARFORMATW) AS BOOLEAN
```

Gets/sets the default character formatting in a rich edit control.

```
(GET) PROPERTY DefaultCharFormat () AS CHARFORMATW
(SET) PROPERTY DefaultCharFormat (BYREF cf AS CHARFORMATW)
```
```
FUNCTION GetDefaultCharFormat () AS CHARFORMATW
FUNCTION SetDefaultCharFormat (BYREF cf AS CHARFORMATW) AS BOOLEAN
```

Gets/sets the character formatting attributes of the current selection.

```
(GET) PROPERTY SelectionCharFormat () AS CHARFORMATW
(SET) PROPERTY SelectionCharFormat (BYREF cf AS CHARFORMATW)
```
```
FUNCTION GetSelectionCharFormat () AS CHARFORMATW
FUNCTION SetSelectionCharFormat (BYREF cf AS CHARFORMATW) AS BOOLEAN
```

Gets/sets the character formatting attributes for the currently selected word.
```
(GET) PROPERTY WordCharFormat () AS CHARFORMATW
(SET) PROPERTY WordCharFormat (BYREF cf AS CHARFORMATW)
```
```
FUNCTION GetWordCharFormat () AS CHARFORMATW
FUNCTION SetWordCharFormat (BYREF cf AS CHARFORMATW) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *fOption* | (GET) Specifies the range of text from which to retrieve formatting. It can be one of the following values.<br>**SCF_DEFAULT** The default character formatting.<br>**SCF_SELECTION** The current selection's character formatting. |

| Parameter  | Description |
| ---------- | ----------- |
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

(GET) Returns a [CHARFORMATW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformatw) structure with the attributes of the first character. The **dwMask** member specifies which attributes are consistent throughout the entire selection. For example, if the entire selection is either in italics or not in italics, CFM_ITALIC is set; if the selection is partly in italics and partly not, CFM_ITALIC is not set.

(SET) If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero. Call **GetLastResult** and/or **GetErrorInfo** to get information about the result.


#### Usage examples

Select text and change color
```
pRichEdit.ExSetSel(98, 113)         ' // Select word at position 98, 113
DIM cf AS CHARFORMAT
cf.dwMask = CFM_COLOR               ' // Let's set the color
cf.crTextColor = BGR(255, 0, 0)     ' // Red color
pRichEdit.SelectionCharFormat = cf  ' // Set the color
pRichEdit.HideSelection(TRUE)       ' // Hide selection
```
Select text and make it bold
```
pRichEdit.ExSetSel(98, 113)         ' // Select word at position 98, 113
DIM cf AS CHARFORMAT
cf.dwMask = CFM_BOLD                ' // The CFE_BOLD value of the dwEffects member is valid.
cf.dwEffects = CFE_BOLD             ' // Characters are bold
pRichEdit.SelectionCharFormat = cf  ' // Set the format
pRichEdit.HideSelection(TRUE)       ' // Hide selection
```
Select text and change the font height
```
pRichEdit.ExSetSel(98, 113)          ' // Select word at position 98, 113
DIM cf AS CHARFORMAT
cf.dwMask = CFM_SIZE                 ' // The yHeight member is valid.
cf.yHeight = 12 * 20                 ' // Character height, in twips (1/1440 of an inch or 1/20 of a printer's point)
pRichEdit.SelectionCharFormat = cf   ' // Set the format
pRichEdit.HideSelection(TRUE)        ' // Hide selection
```

---

# <a name="ctfmodebias"></a>CTFModeBias

Gets/sets the Text Services Framework mode bias values for a Microsoft Rich Edit control.

```
(GET) PROPERTY CTFModeBias () AS LONG
(SET) PROPERTY CTFModeBias (BYVAL nModeBias AS LONG)
```
```
FUNCTION GetCTFModeBias () AS LONG
FUNCTION SetCTFModeBias (BYVAL nModeBias AS LONG) AS HRESULT
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

(SET) If successful, the return value is the new TSF mode bias value. If unsuccessful, the return value is the old TSF mode bias value.

#### Remarks

When a Microsoft Rich Edit application uses TSF, it can select the TSF mode bias. This message sets the criteria by which an alternative choice appears at the top of the list for selection.

---

# <a name="ctfopenstatus"></a>CTFOpenStatus

Gets/sets if the Text Services Framework (TSF) keyboard is open or closed.
```
(GET) PROPERTY CTFOpenStatus () AS BOOLEAN
(SET) PROPERTY CTFOpenStatus (BYVAL fTSFkbd AS LONG)
```
```
FUNCTION GetCTFOpenStatus () AS BOOLEAN
FUNCTION SetCTFOpenStatus (BYVAL fTSFkbd AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *fTSFkbd* | To turn on the TSF keyboard, use **TRUE**. To turn off the TSF keyboard, use **FALSE**. |

#### Return value

(GET) If the TSF keyboard is open, the return value is **TRUE**. Otherwise, it is **FALSE**.

(SET) If successful, it returns **TRUE**. If unsuccessful, it returns **FALSE**. Call the (GET) **CTFOpenStatus** property to check if the value has changed.

---

# <a name="editstyle"></a>EditStyle

Gets/sets the current edit style flags.

```
(GET) PROPERTY EditStyle () AS DWORD
(SET) PROPERTY EditStyle (BYVAL fStyle AS LONG, BYVAL fMask AS LONG)
```
```
FUNCTION GetEditStyle () AS DWORD
FUNCTION SetEditStyle (BYVAL fStyle AS LONG, BYVAL fMask AS LONG) AS HRESULT
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
| **SES_LOGICALCARET** | **Windows 8**: Provide logical caret information instead of a caret bitmap as described in [ITextHost.TxSetCaretPos]([url](https://learn.microsoft.com/en-us/windows/win32/api/textserv/nf-textserv-itexthost-txsetcaretpos)) (default: 0). |
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

---

# <a name="editstyleex"></a>EditStyleEx

Gets/sets the extended edit style flags.

```
(GET) PROPERTY EditStyleEx () AS DWORD
(SET) PROPERTY EditStyleEx (BYVAL fStyle AS LONG, BYVAL fMask AS LONG)
```
```
FUNCTION GetEditStyleEx () AS DWORD
FUNCTION SetEditStyleEx (BYVAL fStyle AS LONG, BYVAL fMask AS LONG) AS HRESULT
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
| **SES_EX_NOMATH** | Disable insertion of math zones (default: 1). To enable math editing and display, call the **EditStyleEx** property with *fStyle* set to 0, and *fMask* set to SES_EX_NOMATH. |
| **SES_EX_NOTABLE** | Disable insertion of tables. The **InsertTable** method returns **E_FAIL** and RTF tables are skipped (default: 0). |
| **SES_EX_USESINGLELINE** | Enable a multiline control to act like a single-line control with the ability to scroll vertically when the single-line height is greater than the window height (default: 0). |
| **SES_HIDETEMPFORMAT** | Hide temporary formatting that is created when **Reset** method of the **ITextFont** is called with **tomApplyTmp**. For example, such formatting is used by spell checkers to display a squiggly underline under possibly misspelled words. |
| **SES_EX_USEMOUSEWPARAM** | Use *wParam* when handling the **WM_MOUSEMOVE** message and do not call **GetAsyncKeyState**. |

#### Return value

(GET) Returns the state of the edit style flags.

(SET) The return value is the state of the edit style flags after the rich edit control has attempted to implement your edit style changes. The edit style flags are a set of flags that indicate the current edit style. Call the (GET) **EditStyleEx** property to check if the value has changed.

---

# <a name="ellipsismode"></a>EllipsisMode

Gets/sets the current ellipsis mode. When enabled, an ellipsis ( ) is displayed for text that doesn't fit in the display window. The ellipsis is only used when the control isn't active. When active, scroll bars are used to reveal text that doesn't fit into the display window.

```
(GET) PROPERTY EllipsisMode () AS DWORD
(SET) PROPERTY EllipsisMode (BYVAL fMode AS DWORD)
```
```
FUNCTION GetEllipsisMode () AS DWORD
FUNCTION SetEllipsisMode (BYVAL fMode AS DWORD) AS BOOLEAN
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

#### Return value

A boolean true(-1) or false (0).

---

# <a name="ellipsisstate"></a>EllipsisState

Gets the current ellipsis state.
```
PROPERTY EllipsisState () AS BOOLEAN
FUNCTION GetEllipsisState () AS BOOLEAN
```
#### Return value

Returns a boolean true (-1) if an ellipsis is being displayed of false (0) otherwise.

---

# <a name="eventmask"></a>EventMask

Gets/sets the event mask for a rich edit control. The event mask specifies which notification messages the control sends to its parent window.

```
(GET) PROPERTY EventMask () AS DWORD
(SET) PROPERTY EventMask (BYVAL fMask AS DWORD)
```
```
FUNCTION GetEventMask () AS DWORD
FUNCTION SetEventMask (BYVAL fMask AS DWORD) AS HRESULT
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

(GET) The event mask for this rich edit control.

(SET) Returns the previous event mask. Call the (GET) **EventMask** property to check if the value has changed.

#### Remarks

The default event mask is **ENM_NONE** in which case no notifications are sent to the parent window.

---

# <a name="firstvisibleline"></a>FirstVisibleLine

Gets the zero-based index of the uppermost visible line in a multiline rich edit control.

```
PROPERTY FirstVisibleLine () AS LONG
FUNCTION GetFirstVisibleLine () AS LONG
```
#### Return value

The return value is the zero-based index of the uppermost visible line in a multiline edit control.

For single-line rich edit controls, the return value is zero.

---

# <a name="hyphenateinfo"></a>HyphenateInfo

Gets/sets information about hyphenation for a Microsoft Rich Edit control.

```
(GET) PROPERTY HyphenateInfo () AS .HYPHENATEINFO
(SET) PROPERTY HyphenateInfo (BYREF hinfo AS .HYPHENATEINFO)
```
```
FUNCTION GetHyphenateInfo () AS .HYPHENATEINFO
FUNCTION SetHyphenateInfo (BYREF info AS .HYPHENATEINFO) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hinfo* | (SET) A [HYPHENATEINFO](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-hyphenateinfo) structure. |

#### Return value

(GET) A [HYPHENATEINFO](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-hyphenateinfo) structure.

---

# <a name="imemodebias"></a>IMEModeBias

Gets/sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.

```
PROPERTY IMEModeBias () AS DWORD
PROPERTY IMEModeBias (BYVAL nModeBias AS LONG)
```
```
FUNCTION GetIMEModeBias () AS DWORD
FUNCTION SetIMEModeBias (BYVAL nModeBias AS LONG) AS HRESULT
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

(GET) To get the Text Services Framework mode bias, use the **CTFModeBias** property.

(SET) When a Microsoft Rich Edit application uses TSF, it can select the TSF mode bias. This message sets the criteria by which an alternative choice appears at the top of the list for selection.

The application should call the **IsIME** method before calling these properties.

---

# <a name="imeoptions"></a>IMEOptions

Gets/sets the current Input Method Editor (IME) options. This message is available only in Asian-language versions of the operating system.

```
(GET) PROPERTY IMEOptions () AS DWORD
(SET) PROPERTY IMEOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
```
```
FUNCTION GetIMEOptions () AS DWORD
FUNCTION SetIMEOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *fCoop* | (SET) Specifies one of the following values.<br>**ECOOP_SET**. Sets the options to those specified by *fOptions*.<br>**COOP_OR**. Combines the specified options with the current options.<br>**ECOOP_AND**. Retains only those current options that are also specified by *fOptions*.<br>**ECOOP_XOR**. Logically exclusive OR the current options with those specified by *fOptions*. |
| *fOptions* | (SET) Specifies one of more of the following values.<br>**IMF_CLOSESTATUSWINDOW**. Closes the IME status window when the control receives the input focus.<br>**IMF_FORCEACTIVE**. Activates the IME when the control receives the input focus.<br>**IMF_FORCEDISABLE**. Disables the IME when the control receives the input focus.<br>**IMF_FORCEENABLE**. Enables the IME when the control receives the input focus.<br>**IMF_FORCEINACTIVE**. Inactivates the IME when the control receives the input focus.<br>**IMF_FORCENONE**. Disables IME handling.<br>**IMF_FORCEREMEMBER**. Restores the previous IME status when the control receives the input focus.<br>**IMF_MULTIPLEEDIT**. Specifies that the composition string will not be canceled or determined by focus changes. This allows an application to have separate composition strings on each rich edit control.<br>**IMF_VERTICAL**. Note: used in Rich Edit 2.0 and later. |

#### Return value

(GET) Returns one or more of the IME option flag values described in the *fOptions* parameter of the SET property.

(SET) If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero. Call the (GET) **ImeOptions** property to check if the value has changed.

---

# <a name="langoptions"></a>LangOptions

Gets/sets a rich edit control's option settings for Input Method Editor (IME) and Asian language support.

```
(GET) PROPERTY LangOptions () AS DWORD
(SET) PROPERTY LangOptions (BYVAL lgoptions AS LONG)
```
```
FUNCTION GetLangOptions () AS DWORD
FUNCTION SetLangOptions (BYVAL lgoptions AS LONG) AS HRESULT
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

(SET) Returns a value of 1. Call the (GET) **LangOptions** property to check the value.

#### Remarks

The **IMF_AUTOFONT** flag is set by default. The **IMF_AUTOKEYBOARD** and **IMF_IMECANCELCOMPLETE** flags are cleared by default.

---

# <a name="limittext"></a>LimitText

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

---

# <a name="getlimittext"></a>GetLimitText

Gets the current text limit for a rich edit control. The text limit is the maximum amount of text that the user can type into the edit control.
```
FUNCTION GetLimitText () AS LONG
```

#### Return value

The return value is the text limit.

---

# <a name="setlimittext"></a>SetLimitText

Sets the current text limit for a rich edit control. The text limit is the maximum amount of text that the user can type into the edit control.
```
SUB SetLimitText (BYVAL chMax AS DWORD)
```
| Parameter  | Description |
| ---------- | ----------- |
| *chMax* | The maximum number of characters the user can enter. If this parameter is zero, the text length is set to 64,000 characters. |

#### Return value

The set property does not return a value.

#### Remarks

**SetLimitText** limits only the text the user can enter. It does not affect any text already in the edit control when the message is sent, nor does it affect the length of the text copied to the edit control by the **SetText** method. If an application uses the **SetText** method to place more text into an edit control than is specified by the **SetLimitText** method, the user can edit the entire contents of the edit control.

---

# <a name="modify"></a>Modify

Gets/sets the state of a rich edit control's modification flag. The flag indicates whether the contents of the rich edit control have been modified.

```
(GET) PROPERTY Modify () AS LONG
(SET) PROPERTY Modify (BYVAL fModify AS LONG)
```
```
FUNCTION GetModify () AS LONG
SUB SetModify (BYVAL fModify AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *fModify* | (SET) The new value for the modification flag. A value of **TRUE** indicates the text has been modified, and a value of **FALSE** indicates it has not been modified. |

#### Return value

(GET) If the contents of edit control have been modified, the return value is nonzero; otherwise, it is zero.

(SET) The set property does not return a value.

### Remarks

The system automatically clears the modification flag to zero when the control is created. If the user changes the control's text, the system sets the flag to nonzero. You can use the (SET) **Modify** property to set or clear the flag.

---

# <a name="options"></a>Options

Gets/sets the options for a rich edit control.

```
(GET) PROPERTY Options () AS DWORD
(SET) PROPERTY Options (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG)
```
```
FUNCTION GetOptions () AS DWORD
FUNCTION SetOptions (BYVAL fCoop AS LONG, BYVAL fOptions AS LONG) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *fCoop* | Specifies the operation, which can be one of these values.<br>**ECOOP_SET**. Sets the options to those specified by *fOptions*.<br>**ECOOP_OR**. Combines the specified options with the current options.<br>**ECOOP_AND**. Retains only those current options that are also specified by *fOptions*.<br>**ECOOP_XOR**. Logically exclusive OR the current options with those specified by *fOptions*. |
| *fOptions* | Specifies one or more of the following values.<br>**ECO_AUTOWORDSELECTION**- Automatic selection of word on double-click.<br>**ECO_AUTOVSCROLL**. Same as **ES_AUTOVSCROLL** style.<br>**ECO_AUTOHSCROLL**. Same as **ES_AUTOHSCROLL** style.<br>**ECO_NOHIDESEL**. Same as **ES_NOHIDESEL** style.<br>**ECO_READONLY**. Same as **ES_READONLY** style.<br>**ECO_WANTRETURN**. Same as **ES_WANTRETURN** style.<br>**ECO_SELECTIONBAR**. Same as **ES_SELECTIONBAR** style.<br>**ECO_VERTICAL**. Same as **ES_VERTICAL** style. Available in Asian-language versions only. |

#### Return value

(GET) This message returns a combination of the current option flag values described in the *fOptions* parameter of the (SET) **Options** property.

(SET) This message returns the current options of the edit control. Call the (GET) **Options** property to get the values.

---

# <a name="pagerotate"></a>PageRotate

Deprecated. Gets/sets the text layout for a Microsoft Rich Edit control.

```
(GET) PROPERTY PageRotate () AS DWORD
(SET) PROPERTY PageRotate (BYVAL txtlayout AS LONG)
```
```
FUNCTION GetPageRotate () AS DWORD
FUNCTION SetPageRotate (BYVAL txtlayout AS LONG) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *txtlayout* | Text layout value. This can be one of the following values.<br>**EPR_0**. Text flows from left to right and from top to bottom.<br>**EPR_90**. Text flows from bottom to top and from left to right.<br>**EPR_180**. Text flows from right to left and from bottom to top.<br>**EPR_270**. Text flows from top to bottom and from right to left.<br>**EPR_SE**. Windows 8: Text flows top to bottom and left to right (Mongolian text layout). |

#### Return value

(GET) The current text layout. For a list of possible text layout values, see the *txtlayout* parameter above.

(SET) Return value is the new text layout value. Call the (GET) **PageRotate** property to get the value.

#### Remarks

(SET) This message sets the text layout for the entire document. However, embedded contents are not rotated and must be rotated separately by the application.

---

# <a name="paraformat"></a>ParaFormat

Gets/sets the paragraph formatting of the current selection in a rich edit control.

```
(GET) PROPERTY ParaFormat () AS .PARAFORMAT
(SET) PROPERTY ParaFormat (BYREF pfmt AS .PARAFORMAT)
```
```
FUNCTION GetParaFormat () AS .PARAFORMAT
FUNCTION SetParaFormat (BYREF pfmt AS .PARAFORMAT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfmt* | (SET) A [PARAFORMAT](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-paraformat) structure specifying the new paragraph formatting attributes. Only the attributes specified by the **dwMask** member are changed. |

#### Return value

(GET) Returns a [PARAFORMAT](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-paraformat) structure.

(SET) If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero.

---

# <a name="passwordchar"></a>PasswordChar

Gets/sets the password character that a rich edit control displays when the user enters text.

```
(GET) PROPERTY PasswordChar () AS LONG
(SET) PROPERTY PasswordChar (BYVAL dwchar AS DWORD)
```
```
FUNCTION GetPasswordChar () AS LONG
FUNCTION SetPasswordChar (BYVAL dwchar AS DWORD)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwchar* | (SET) The character to be displayed in place of the characters typed by the user. If this parameter is zero, the control removes the current password character and displays the characters typed by the user. |

#### Return value

(GET) The return value specifies the character to be displayed in place of any characters typed by the user. If the return value is **NULL**, there is no password character, and the control displays the characters typed by the user.

(SET) The set property does not return a value.

#### Remarks

When an edit control receives the **EM_SETPASSWORDCHAR** message, the control redraws all visible characters using the character specified by the dwchar parameter. If *dwchar* is zero, the control redraws all visible characters using the characters typed by the user.

If an edit control is created with the **ES_PASSWORD** style, the default password character is set to an asterisk (*). If an edit control is created without the **ES_PASSWORD** style, there is no password character. The **ES_PASSWORD** style is removed if an **EM_SETPASSWORDCHAR** is sent with the *dwchar* parameter set to zero.

**Edit controls**: Multiline edit controls do not support the password style or messages.

---

# <a name="punctuation"></a>Punctuation

Gets/sets the current punctuation characters for the rich edit control.

```
(GET) PROPERTY Punctuation (BYVAL punctype AS DWORD) AS .PUNCTUATION
(SET) PROPERTY Punctuation (BYVAL punctype AS LONG, BYREF punct AS .PUNCTUATION)
```
```
FUNCTION GetPunctuation (BYVAL punctype AS DWORD) AS .PUNCTUATION
FUNCTION SetPunctuation (BYVAL punctype AS LONG, BYREF punct AS .PUNCTUATION) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *punctype* | Specifies the punctuation type, which can be one of the following values.<br>**PC_LEADING**. Leading punctuation characters.<br>**PC_FOLLOWING**. Following punctuation characters.<br>**PC_DELIMITER**. Delimiter.<br>**PC_OVERFLOW**. Not supported. |
| *punct* | A [PUNCTUATION](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-punctuation) structure that contains the punctuation characters. |

#### Return value

If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero.

### Note

This message is supported only in Asian-language versions of Microsoft Rich Edit 1.0. It is not supported in any later versions.

---

# <a name="rect"></a>Rect

Gets/sets the formatting rectangle of a rich edit control.

```
(GET) PROPERTY Rect () AS .RECT
(SET) PROPERTY Rect (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
```
```
FUNCTION GetRect () AS .RECT
SUB SetRect (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *fCoord* | (SET) **Rich Edit 2.0 and later**: Indicates whether *rc* specifies absolute or relative coordinates. A value of zero indicates absolute coordinates. A value of 1 indicates offsets relative to the current formatting rectangle. (The offsets can be positive or negative.)<br>**Edit controls and Rich Edit 1.0**: This parameter is not used and must be zero. |
| *rc* | (SET) A [RECT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-rect) structure that specifies the new dimensions of the rectangle. If this parameter is **NULL**, the formatting rectangle is set to its default values. |

#### Return value

(GET) A **RECT** structure with the formatting rectangle.

(SET) The set property does not return a value.

#### Remarks

Setting *rc* to **NULL** has no effect if a touch device is installed, or if the message is sent from a thread that has a hook installed (see [SetWindowsHookEx](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowshookexw)). In these cases, *rc* should contain a valid pointer to a [RECT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-rect) structure.

The (SET) **Rect** property causes the text of the rich edit control to be redrawn. To change the size of the formatting rectangle without redrawing the text, use the **RectNP** property.

When a rich edit control is first created, the formatting rectangle is set to a default size. You can use the (SET) **Rect** message to make the formatting rectangle larger or smaller than the rich edit control window.

If the rich edit control does not have a horizontal scroll bar, and the formatting rectangle is set to be larger than the rich edit control window, lines of text exceeding the width of the rich edit control window (but smaller than the width of the formatting rectangle) are clipped instead of wrapped.

If the rich edit control contains a border, the formatting rectangle is reduced by the size of the border. If you are adjusting the rectangle returned by the (GET) **Rect** property, you must remove the size of the border before using the rectangle with the (SET) **Rect** property.

The formatting rectangle does not include the selection bar, which is an unmarked area to the left of each paragraph. When the user clicks in the selection bar, the corresponding line is selected.

---

# <a name="rectnp"></a>RectNP

Sets the formatting rectangle of a multiline rich edit control. It is identical to the **Rect** property, except that **RectNP** does not redraw the edit control window.

The formatting rectangle is the limiting rectangle into which the control draws the text. The limiting rectangle is independent of the size of the edit control window.

This message is processed only by multiline edit controls. You can send this message to either an edit control or a rich edit control.

```
(SET) PROPERTY RectNP (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
```
```
SUB SetRectNP (BYVAL fCoord AS LONG, BYREF rc AS .RECT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *fCoord* | (SET) **Rich Edit 3.0 and later**: Indicates whether *prect* specifies absolute or relative coordinates. A value of zero indicates absolute coordinates. A value of 1 indicates offsets relative to the current formatting rectangle. (The offsets can be positive or negative.) |
| *rc* | (SET) A [RECT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-rect) structure that specifies the new dimensions of the rectangle. If this parameter is **NULL**, the formatting rectangle is set to its default values. |

#### Return value

This property does not return a value.

---

# <a name="redoname"></a>RedoName

Retrieves the type of the next action, if any, in the control's redo queue.

```
FUNCTION RedoName () AS LONG
FUNCTION GetRedoName () AS LONG
```

#### Return value

If the redo queue for the control is not empty, the value returned is an [UNDONAMEID](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ne-richedit-undonameid) enumeration value that indicates the type of the next action in the control's redo queue.

If there are no redoable actions or the type of the next redoable action is unknown, the return value is zero.

#### UNDONAMEID enumeration

| Name              | Value | Description |
| ----------------- | ----- | ----------- |
| **UID_UNKNOWN**   |   0   | The type of undo action is unknown. |
| **UID_TYPING**    |   1   | Typing operation. |
| **UID_DELETE**    |   2   | Delete operation. |
| **UID_DRAGDROP**  |   3   | Drag-and-drop operation. |
| **UID_CUT**       |   4   | Cut operation. |
| **UID_PASTE**     |   5   | Paste operation. |
| **UID_AUTOTABLE** |   6   | Automatic table insertion; for example, typing +---+---+<Enter> to insert a table row. |

#### Remarks

The types of actions that can be undone or redone include typing, delete, drag-drop, cut, and paste operations. This information can be useful for applications that provide an extended user interface for undo and redo operations, such as a drop-down list box of redoable actions.

---

# <a name="scrollpos"></a>ScrollPos

Gets/sets the current scroll position of the edit control.

```
(GET) PROPERTY ScrollPos () AS .POINT
(SET) PROPERTY ScrollPos (BYREF pt AS .POINT)
```
```
FUNCTION GetScrollPos () AS .POINT
FUNCTION SetScrollPos (BYREF pt AS .POINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pt* | A [POINT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-point) structure which specifies a point in the virtual text space of the document, expressed in pixels. The document will be scrolled until this point is located in the upper-left corner of the edit control window. If you want to change the view such that the upper left corner of the view is two lines down and one character in from the left edge. You would pass a point of (7, 22).<br>The rich edit control checks the x and y coordinates and adjusts them if necessary, so that a complete line is displayed at the top. It also ensures that the text is never completely scrolled off the view rectangle. |

#### Return value

(GET) A [POINT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-point) structure containing a point in the virtual text space of the document, expressed in pixels. This point will be the point that is currently located in the upper-left corner of the edit control window.

(SET) This message always returns 1.

#### Remarks

The values returned in the [POINT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-point) structure are 16-bit values (even in the 32-bit wide fields).

---

# <a name="storytype"></a>StoryType

Gets/sets the story type.

```
PROPERTY StoryType (BYVAL Index AS DWORD) AS DWORD
PROPERTY StoryType (BYVAL Index AS LONG, BYVAL dwType AS DWORD)
```
```
FUNCTION GetStoryType (BYVAL Index AS DWORD) AS DWORD
FUNCTION SetStoryType (BYVAL Index AS LONG, BYVAL dwType AS DWORD) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *Index* | (GET/SET) The story index. |
| *dwType* | (SET) The new story type. See the list below. |

| Constant  | Value | Description |
| --------- | ----- | ----------- |
| **tomCommentsStory** | 4 | The story used for comments. |
| **tomEndnotesStory** | 3 | The story used for endnotes. |
| **tomEvenPagesFooterStory** | 8 | The story containing footers for even pages. |
| **tomEvenPagesHeaderStory** | 6 | The story containing headers for even pages. |
| **tomFindStory** | 128 | The story used for a Find dialog. |
| **tomFirstPageFooterStory** | 11 | The story containing the footer for the first page. |
| **tomFirstPageHeaderStory** | 10 | The story containing the header for the first page. |
| **tomFootnotesStory** | 2 | The story used for footnotes. |
| **tomMainTextStory** | 1 | The main story always exists for a rich edit control. |
| **tomPrimaryFooterStory** | 9 | The story containing footers for odd pages. |
| **tomPrimaryFooterStory** | 7 | The story containing headers for odd pages. |
| **tomReplaceStory** | 129 | The story used for a Replace dialog. |
| **tomScratchStory** | 127 | The scratch story. |
| **tomTextFrameStory** | 5 | The story used for a text box. |
| **tomUnknownStory** | 0 | No special type. |

#### Return value

(GET) Returns the story type, which can be a client-defined custom value, or one of the values listed in the table above.

(SET) The story type that was set. Call the (GET) **StoryType** property to get the value.

---

# <a name="text"></a>Text

Gets/sets the text from a rich edit control.

```
(GET) PROPERTY Text () AS CWSTR
(SET) PROPERTY Text (BYREF wszText AS WSTRING)
```
```
FUNCTION GetText () AS CWSTR
FUNCTION SetText (BYREF wszText AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | A **WSTRING** with the new text. |

#### Return value

(GET) The retrieved text.

(SET) The return value is **TRUE** if the text is set. Call **GetLastResult** and/or **GetErrorInfo** to get information about the result.

#### Remarks

The Windows API function **GetWindowTextW** can also be used to retrieve the text of a rich edit control, but it cannot retrieve the text of a control in another application.

#### Usage example (GET)
```
DIM cws AS CWSTR = pRichEdit.Text
```

#### Usage example (SET)
```
DIM cws AS CWSTR = "New text"
pRichEdit.Text = cws
```

# <a name="textlength"></a>TextLength

Retrieves the length of all text in a rich edit control.

```
FUNCTION TextLength () AS LONG
FUNCTION GetTextLength () AS LONG
```

#### Return value

The return value is the length of the text in characters, not including the terminating null character.

#### Remarks

When the **GetTextLength** method is called, the **DefWindowProc** function returns the length, in characters, of the text. Under certain conditions, the **DefWindowProc** function returns a value that is larger than the actual length of the text. This occurs with certain mixtures of ANSI and Unicode, and is due to the system allowing for the possible existence of double-byte character set (DBCS) characters within the text. The return value, however, will always be at least as large as the actual length of the text; you can thus always use it to guide buffer allocation. This behavior can occur when an application uses both ANSI functions and common dialogs, which use Unicode.

To retrieve the text, you can also use the **AfxGetWindowText** function, the **WM_GETTEXT** message, or the Windows API **GetWindowTextW** function.

---

# <a name="textlengthex"></a>TextLengthEx

Calculates text length in various ways. It is usually called before creating a buffer to receive the text from the control. Since this `CRichEditCtrl` class uses unicode, it is easier to use the simplified second overloaded function.

```
FUNCTION TextLengthEx (BYREF gtex AS .GETTEXTLENGTHEX) AS LONG
FUNCTION TextLengthEx (BYVAL dwFlags AS DWORD = GTL_DEFAULT) AS LONG
FUNCTION GetTextLengthEx OVERLOAD (BYREF gtex AS .GETTEXTLENGTHEX) AS LONG
FUNCTION GetTextLengthEx OVERLOAD (BYVAL dwFlags AS DWORD = GTL_DEFAULT) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *gtex* | A [GETTEXTLENGTHEX](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-gettextlengthex) structure that receives the text length information. |
| *dwFlags* | Value specifying the method to be used in determining the text length. This member can be one or more of the following values (some values are mutually exclusive). |

| Flag  | Meaning |
| ----- | ------- |
| **GTL_DEFAULT** | Returns the number of characters. This is the default. |
| **GTL_USECRLF** | Computes the answer by using CR/LFs at the end of paragraphs. |
| **GTL_PRECISE** | Computes a precise answer. This approach could necessitate a conversion and thereby take longer. This flag cannot be used with the **GTL_CLOSE** flag. **E_INVALIDARG** will be returned if both are used. |
| **GTL_CLOSE** | Computes an approximate (close) answer. It is obtained quickly and can be used to set the buffer size. This flag cannot be used with the **GTL_PRECISE** flag. **E_INVALIDARG** will be returned if both are used. |
| **GTL_NUMCHARS** | Returns the number of characters. This flag cannot be used with the **GTL_NUMBYTES** flag. **E_INVALIDARG** will be returned if both are used. |
| **GTL_NUMBYTES** | Returns the number of bytes. This flag cannot be used with the **GTL_NUMCHARS** flag. **E_INVALIDARG** will be returned if both are used |

#### Return value

**First overloaded method**: The method returns the number of characters in the edit control, depending on the setting of the flags in the [GETTEXTLENGTHEX](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-gettextlengthex) structure. If incompatible flags were set in the **flags** member, the message returns **E_INVALIDARG**.

**Second overloaded method**: The method returns the number of characters in the edit control, depending on the setting of the flags in the *dwFlags* parameter. If incompatible flags were set in the **flags** member, the message returns **E_INVALIDARG**.

#### Remarks

This message is a fast and easy way to determine the number of characters in the Unicode version of the rich edit control. However, for a non-Unicode target code page you will potentially be converting to a combination of single-byte and double-byte characters.

#### Usage examples
```
--- First overloaded method:
DIM gtex AS GETTEXTLENGTHEX = TYPE<GETTEXTLENGTHEX>(GTL_NUMCHARS, 1200)
DIM Result AS LONG = pRichEdit.GetTextLEngthEx(gtex)
DIM cbLen AS LONG
IF Result <> E_INVALIDARG THEN cbLen = Result
```
```
--- Second overloaded method:
DIM Result AS LONG = pRichEdit.GetTextLEngthEx                 ' // Uses the GTL_DEFAULT flag
DIM Result AS LONG = pRichEdit.GetTextLEngthEx(GTL_NUMCHARS)   ' // Uses the GTL_NUMCHARS flag
DIM cbLen AS LONG
IF Result <> E_INVALIDARG THEN cbLen = Result
```
---

# <a name="textmode"></a>TextMode

Gets/sets the current text mode and undo level of a rich edit control.

```
(GET) PROPERTY TextMode () AS DWORD
(SET) PROPERTY TextMode (BYVAL values AS LONG)
```
```
FUNCTION GetTextMode () AS DWORD
FUNCTION SetTextMode (BYVAL values AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *values* | One or more values from the [TEXTMODE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ne-richedit-textmode) enumeration type. The values specify the new settings for the control's text mode and undo level parameters. |

#### TextMode enumeration type

Specify one of the following values to set the text mode parameter. If you do not specify a text mode value, the text mode remains at its current setting.

| Value  | Meaning |
| ------ | ----------- |
| **TM_PLAINTEXT** | Indicates plain text mode, in which the control is similar to a standard edit control. For more information about plain text mode, see the **Remarks** section. |
| **TM_RICHTEXT** | Indicates rich text mode, in which the control has standard rich edit functionality. Rich text mode is the default setting. |

Specify one of the following values to set the undo level parameter. If you do not specify an undo level value, the undo level remains at its current setting.

| Value  | Meaning |
| ------ | ----------- |
| **TM_SINGLELEVELUNDO** | The control allows the user to undo only the last action that can be undone. |
| **TM_MULTILEVELUNDO** | The control supports multiple undo operations. This is the default setting. Use the **RichEdit_SetUndoLimit** message to set the maximum number of undo actions. |

Specify one of the following values to set the code page parameter. If you do not specify an code page value, the code page remains at its current setting.

| Value  | Meaning |
| ------ | ----------- |
| **TM_SINGLECODEPAGE** | The control only allows the English keyboard and a keyboard corresponding to the default character set. For example, you could have Greek and English. Note that this prevents Unicode text from entering the control. For example, use this value if a Rich Edit control must be restricted to ANSI text. |
| **TM_MULTICODEPAGE** | The control allows multiple code pages and Unicode text into the control. This is the default setting. |

#### Return value

(GET) The return value is one or more values from the [TEXTMODE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ne-richedit-textmode) enumeration type. The values indicate the current text mode and undo level of the control.

(SET) If the message succeeds, the result code is zero. If the message fails, the result code is a nonzero value.

#### Remarks

In rich text mode, a rich edit control has standard rich edit functionality. However, in plain text mode, the control is similar to a standard edit control:

- The text in a plain text control can have only one format (such as Bold, 10pt Arial).
- The user cannot paste rich text formats, such as Rich Text Format (RTF) or embedded objects into a plain text control.
- Rich text mode controls always have a default end-of-document marker or carriage return, to format paragraphs. Plain text controls, on the other hand, do not need the default, end-of-document marker, so it is omitted.
- The control must contain no text when it receives the **EM_SETTEXTMODE** message. To ensure there is no text, call the (SET) **Text** property with an empty string ("").

---

# <a name="touchoptions"></a>TouchOptions

Gets/sets the touch options that are associated with a rich edit control.

```
(GET) PROPERTY TouchOptions (BYVAL opt AS LONG PTR) AS DWORD
(SET) PROPERTY TouchOptions (BYVAL opt AS LONG, BYVAL fEnable AS LONG)
```
```
FUNCTION GetTouchOptions (BYVAL opt AS LONG PTR) AS DWORD
FUNCTION SetTouchOptions (BYVAL opt AS LONG, BYVAL fEnable AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *opt* | (GET/SET) The touch options to set. It can be one of the following values:<br>**RTO_SHOWHANDLES**. Show or hide the touch gripper handles, depending on the value of *opt*.<br>**RTO_DISABLEHANDLES**. Enable or disable the touch gripper handles, depending on the value of *opt*. When handles are disabled, they are hidden if they are visible and remain hidden until **TouchOptions** changes their status. |
| *fEnable* | (SET) Set to **TRUE** to show/enable the touch selection handles, or **FALSE** to hide/disable the touch selection handles. |

#### Return value

(GET) Returns the value of the option specified by the *opt* parameter. It is nonzero if *opt* is **RTO_SHOWHANDLES** and the touch grippers are visible; zero, otherwise.

(SET) Returns **TRUE** if *opt* is valid, otherwise **FALSE**.

#### Remarks

Advanced line breaking is turned on automatically by the rich edit control when needed, such as for handling complex scripts like Arabic and Hebrew, and for mathematics. It s also needed for justified paragraphs, hyphenation, and other typographic features.

---

# <a name="typographyoptions"></a>TypographyOptions

Gets/sets the current state of the typography options of a rich edit control.

```
(GET) PROPERTY TypographyOptions () AS DWORD
(SET) PROPERTY TypographyOptions (BYVAL pto AS LONG, BYVAL fMask AS LONG)
```
```
FUNCTION GetTypographyOptions () AS DWORD
FUNCTION SetTypographyOptions (BYVAL pto AS LONG, BYVAL fMask AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pto* | Specifies one or both of the following values.<br>**TO_ADVANCEDTYPOGRAPHY**. Advanced line breaking and line formatting is turned on.<br>**TO_SIMPLELINEBREAK**. Faster line breaking for simple text (requires **TO_ADVANCEDTYPOGRAPHY**). |
| *fMask* | A mask consisting of one or more of the flags in *pto*. Only the flags that are set in this mask will be set or cleared. This allows a single flag to be set or cleared without reading the current flag states. |

#### Return value

(GET) Returns the current typography options.

(SET) A boolean true (-1) or false (0).

#### Remarks

You can turn on advanced line breaking by sending calling the (SET) **TypographyOPtions** property. Advanced line breaking is turned on automatically by the rich edit control when needed, such as for handling complex scripts like Arabic and Hebrew, and for mathematics. It s also needed for justified paragraphs, hyphenation, and other typographic features.

---

# <a name="wordbreakproc"></a>WordBreakProc

Gets/sets the address of the currently registered word-break procedure.

```
PROPERTY WordBreakProc () AS LONG_PTR
PROPERTY WordBreakProc (BYVAL pfn AS LONG_PTR)
```
```
FUNCTION GetWordBreakProc () AS LONG_PTR
SUB SetWordBreakProc (BYVAL pfn AS LONG_PTR)
```

Gets/sets the address of the currently registered extended word-break procedure.

```
PROPERTY WordBreakProcEx () AS LONG_PTR
PROPERTY WordBreakProcEx (BYVAL pfn AS LONG_PTR)
```
```
FUNCTION GetWordBreakProcEx () AS LONG_PTR
SUB SetWordBreakProcEx (BYVAL pfn AS LONG_PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfn* | Pointer to an [EditWordBreakProcEx](https://learn.microsoft.com/en-us/windows/win32/api/richedit/nc-richedit-editwordbreakprocex) function, or **NULL** to use the default procedure. |

#### Return value

The return value specifies the address of the application-defined word-break function. The return value is **NULL** if no word-break function exists.

#### Remarks

A Wordwrap function scans a text buffer that contains text to be sent to the display, looking for the first word that does not fit on the current display line. The wordwrap function places this word at the beginning of the next line on the display. A Wordwrap function defines the point at which the system should break a line of text for multiline edit controls, usually at a space character that separates two words.

---

# <a name="wordwrap"></a>WordWrap

Enables/disables word wrap.

```
FUNCTION DisableWordWrap (BYVAL LineWidth AS LONG = 32767) AS BOOLEAN
FUNCTION EnableWordWrap () AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *LineWidth* | The line width. Default value: 32767 characters. |

#### Return value

If the method succeeds it returns the boolean value true (-1); if it fails, it returns false (0).

---

# <a name="wordwrapmode"></a>WordWrapMode

Gets/sets the current word wrap and word-break options for the rich edit control.

```
(GET) PROPERTY WordWrapMode () AS DWORD
(SET) PROPERTY WordWrapMode (BYVAL values AS DWORD) AS DWORD
```
```
FUNCTION GetWordWrapMode () AS DWORD
FUNCTION SetWordWrapMode (BYVAL values AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *values* | Specifies one or more of the following values. |

| Value  | Meaning |
| ------ | ------- |
| **WBF_WORDWRAP** | Enables Asian-specific word wrap operations, such as kinsoku in Japanese. |
| **WBF_WORDBREAK** | Enables English word-breaking operations in Japanese and Chinese. Enables Hangeul word-breaking operation. |
| **WBF_OVERFLOW** | Recognizes overflow punctuation. (Not currently supported.) |
| **WBF_LEVEL1** | Sets the Level 1 punctuation table as the default. |
| **WBF_LEVEL2** | Sets the Level 2 punctuation table as the default. |
| **WBF_CUSTOM** | Sets the application-defined punctuation table. |

#### Return value

(GET) The the current word wrap and word-break options.

(SET) This method returns the current word-wrapping and word-breaking options.

#### Remarks

This property is supported only in Asian-language versions of Microsoft Rich Edit 1.0. It is not supported in any later versions of Rich Edit. This message must not be sent by the application-defined, word-break procedure.

---

# <a name="callautocorrectproc"></a>CallAutocorrectProc

Calls the autocorrect callback function that is stored by the (SET) **AutocorrectProc** property, provided that the text preceding the insertion point is a candidate for autocorrection.
```
FUNCTION CallAutocorrectProc (BYVAL char AS WCHAR) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *char* | A character of type **WCHAR**. If this character is a tab (U+0009), and the character preceding the insertion point isn't a tab, then the character preceding the insertion point is treated as part of the autocorrect candidate string instead of as a string delimiter; otherwise, it has no effect. |

#### Return value

The return value is zero if the message succeeds, or nonzero if an error occurs.

---

# <a name="canpaste"></a>CanPaste

Determines whether a rich edit control can paste a specified clipboard format.
```
FUNCTION CanPaste (BYVAL clipformat AS LONG) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *clipformat* | Specifies the [Clipboard Formats](https://learn.microsoft.com/en-us/windows/win32/dataxchg/clipboard-formats) to try. To try any format currently on the clipboard, set this parameter to zero. |

#### Return value

If the clipboard format can be pasted, the return value is a nonzero value. If the clipboard format cannot be pasted, the return value is zero.

---

# <a name="canredo"></a>CanRedo

Determines whether there are any actions in the control redo queue.
```
FUNCTION CanRedo () AS BOOLEAN
```
#### Return value

If there are actions in the control redo queue, the return value is a nonzero value. If the redo queue is empty, the return value is zero.

# <a name="canundo"></a>CanUndo

Determines whether there are any actions in an edit control's undo queue.
```
FUNCTION CanUndo () AS BOOLEAN
```
#### Return value

If there are actions in the control's undo queue, the return value is nonzero. If the undo queue is empty, the return value is zero.

---

# <a name="displayband"></a>DisplayBand

Displays a portion of the contents of a rich edit control, as previously formatted for a device using the **FormatRange** method.
```
FUNCTION DisplayBand (BYREF rc AS .RECT) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *rc* | A [RECT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-rect) structure specifying the display area of the device. |

#### Return value

If the operation succeeds, the return value is **TRUE**. If the operation fails, the return value is **FALSE**.

#### Remarks

Text and Component Object Model (COM) objects are clipped by the rectangle. The application does not need to set the clipping region.

Banding is the process by which a single page of output is generated using one or more separate rectangles, or bands. When all bands are placed on the page, a complete image results. This approach is often used by raster printers that do not have sufficient memory or ability to image a full page at one time. Banding devices include most dot matrix printers as well as some laser printers.

---

# <a name="emptyundobuffer"></a>EmptyUndoBuffer

Resets the undo flag of a rich edit control. The undo flag is set whenever an operation within the rich edit control can be undone.
```
SUB EmptyUndoBuffer ()
```
#### Return value

This method does not return a value.

#### Remarks

The undo flag is automatically reset whenever the edit control receives a **WM_SETTEXT** or **EM_SETHANDLE** message.

---

# <a name="exgetsel"></a>ExGetSel

Oveloaded method to retrieve the starting and ending character positions of the selection in a rich edit control.
```
FUNCTION ExGetSel OVERLOAD () AS CHARRANGE
FUNCTION ExGetSel OVERLOAD (BYREF cpMin AS LONG, BYVAL cpMax AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *cpMin* | Character position index immediately preceding the first character in the range. |
| *cpMaz* | Character position immediately following the last character in the range. |

#### Return value

The first overloaded method returns a [CHARRANGE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charrange) structure with the selection range.

The second overloades method does not return a value. The values are filled in the passed parameters.

#### CHARRANGE structure
```
type _charrange field = 4
   cpMin as LONG
   cpMax as LONG
end type

type CHARRANGE as _charrange
```
| Member  | Description |
| ------- | ----------- |
| *cpMin* | Character position index immediately preceding the first character in the range. |
| *cpMax* | Character position immediately following the last character in the range. |

---

# <a name="exlimittext"></a>ExLimitText

Sets an upper limit to the amount of text the user can type or paste into a rich edit control.
```
SUB ExLimitText (BYVAL dwLimit AS DWORD)
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwLimit* | Specifies the maximum amount of text that can be entered. If this parameter is zero, the default maximum is used, which is 64K characters. If this parameter is -1 (&hFFFFFFFF), raises the the 32,767 default characters limit to no limit. A COM object counts as a single character. |

#### Remarks

The text limit set by the **ExLimitText** method does not limit the amount of text that you can stream into a rich edit control using the **StreamIn** method with the *pedst* parameter set to **SF_TEXT**. However, it does limit the amount of text that you can stream into a rich edit control using the **StreamIn** method with the *pedst* parameter set set to **SF_RTF**.

Before **ExLimitText** is called, the default limit to the amount of text a user can enter is 32,767 characters.

#### Return value

This method does not return a value.

---

# <a name="exlinefromchar"></a>ExLineFromChar

Determines which line contains the specified character in a rich edit control.
```
FUNCTION ExLineFromChar (BYVAL iIndex AS LONG) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *iIndex* | Zero-based index of the character. |

#### Return value

Returns the zero-based index of the line.

---

# <a name="exsetsel"></a>ExSetSel

Overloaded method that selects a range of characters and/or Component Object Model (COM) objects in a Microsoft Rich Edit control.
```
FUNCTION ExSetSel OVERLOAD (BYREF cr AS CHARRANGE) AS DWORD
SUB ExSetSel OVERLOAD (BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1)
```
| Parameter  | Description |
| ---------- | ----------- |
| *cr* | A [CHARRANGE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charrange) structure that specifies the selection range. |

| Parameter  | Description |
| ---------- | ----------- |
| *cpMin* | Character position index immediately preceding the first character in the range. |
| *cpMax* | Character position immediately following the last character in the range. |

#### CHARRANGE structure
```
type _charrange field = 4
   pMin as LONG
   cpMax as LONG
end type

type CHARRANGE as _charrange
```
| Member  | Description |
| ------- | ----------- |
| *cpMin* | Character position index immediately preceding the first character in the range. |
| *cpMax* | Character position immediately following the last character in the range. |

#### Return value

The return value is the selection that is actually set.

#### Usage example
```
DIM chrRange AS CHARRANGE = TYPE<CHARRANGE>(3, 12)
pRichEditCtro.ExSetSel(@chrRange)
```

---

# <a name="findtext"></a>FindText

Overloaded method to find text within a rich edit control.
```
FUNCTION FindText OVERLOAD (BYVAL fOptions AS DWORD, BYREF ft AS FINDTEXTW) AS LONG
FUNCTION FindText OVERLOAD (BYVAL fOptions AS DWORD = FR_DOWN, BYVAL cpMin AS LONG = 0, _
   BYVAL cpMax AS LONG = -1, BYREF wszText AS WSTRING) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *fOptions* | Specifies the parameters of the search operation. This parameter can be one or more of the following values.<br>**FR_DOWN**. If set, the operation searches from the end of the current selection to the end of the document. If not set, the operation searches from the end of the current selection to the beginning of the document.<br>**FR_MATCHALEFHAMZA**. By default, Arabic and Hebrew alefs with different accents are all matched by the alef character. Set this flag if you want the search to differentiate between alefs with different accents.<br>**FR_MATCHCASE**. If set, the search operation is case-sensitive. If not set, the search operation is case-insensitive.<br>**FR_MATCHDIAC**. By default, Arabic and Hebrew diacritical marks are ignored. Set this flag if you want the search operation to consider diacritical marks.<br>**FR_MATCHKASHIDA**. By default, Arabic and Hebrew kashidas are ignored. Set this flag if you want the search operation to consider kashidas.<br>**FR_WHOLEWORD**. If set, the operation searches only for whole words that match the search string. If not set, the operation also searches for word fragments that match the search string.|
| *ft* | A [FINDTEXTW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-findtextw) structure containing information about the find operation. |
| *cpMin* | Character position index immediately preceding the first character in the range. |
| *cpMax* | Character position immediately following the last character in the range. If the *cpMin* and *cpMax* members are equal, the range is empty. The range includes everything if *cpMin* is 0 and *cpMax* is –1. |
| *pwszText* | A **WSTRING** containing the text to find. |

#### Return value

If the target string is found, the return value is the zero-based position of the first character of the match. If the target is not found, the return value is -1.

Search in a range:
```
DIM cws AS CWSTR = "box"
DIM cbStart AS LONG = pRichEdit.FindText(FR_DOWN, 5, 20, cws)
pRichEdit.ExSetSel(cbStart, cbStart + LEN(cws))
```
Search in the entire document:
```
DIM cws AS CWSTR = "box"
DIM cbStart AS LONG = pRichEdit.FindText(FR_DOWN, 0, -1, cws)
--or--
DIM cbStart AS LONG = pRichEdit.FindText(FR_DOWN, , , cws)
pRichEdit.ExSetSel(cbStart, cbStart + LEN(cws))
```
Case sensitive search:
```
DIM cws AS CWSTR = "box"
DIM cbStart AS LONG = pRichEdit.FindText(FR_DOWN OR FR_MATCHCASE, 5, 20, cws)
pRichEdit.ExSetSel(cbStart, cbStart + LEN(cws))
```
Search backwards from the end of the document to the beginning:
```
DIM cbStart AS LONG = pRichEdit.GetTextLength
DIM cbStart AS LONG = pRichEdit.FindText(0, cbStart, 0, cws)
pRichEdit.ExSetSel(cbStart, cbStart + LEN(cws))
```

---

# <a name="findtextex"></a>FindTextEx

Overloaded method to find text within a rich edit control.
```
FUNCTION FindTextEx OVERLOAD (BYVAL fOptions AS DWORD, BYREF ftexw AS FINDTEXTEXW) AS LONG
FUNCTION FindTextEx OVERLOAD (BYVAL fOptions AS DWORD = FR_DOWN, BYVAL cpMin AS LONG = 0, _
   BYVAL cpMax AS LONG = -1, BYREF wszText AS WSTRING) AS CHARRANGE
```
| Parameter  | Description |
| ---------- | ----------- |
| *fOptions* | Specifies the parameters of the search operation. This parameter can be one or more of the following values.<br>**FR_DOWN**. If set, the operation searches from the end of the current selection to the end of the document. If not set, the operation searches from the end of the current selection to the beginning of the document.<br>**FR_MATCHALEFHAMZA**. By default, Arabic and Hebrew alefs with different accents are all matched by the alef character. Set this flag if you want the search to differentiate between alefs with different accents.<br>**FR_MATCHCASE**. If set, the search operation is case-sensitive. If not set, the search operation is case-insensitive.<br>**FR_MATCHDIAC**. By default, Arabic and Hebrew diacritical marks are ignored. Set this flag if you want the search operation to consider diacritical marks.<br>**FR_MATCHKASHIDA**. By default, Arabic and Hebrew kashidas are ignored. Set this flag if you want the search operation to consider kashidas.<br>**FR_WHOLEWORD**. If set, the operation searches only for whole words that match the search string. If not set, the operation also searches for word fragments that match the search string.|
| *ftexw* | A [FINDTEXTEXW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-findtextexw) structure containing information about the find operation. |
| *cpMin* | Character position index immediately preceding the first character in the range. |
| *cpMax* | Character position immediately following the last character in the range. If the *cpMin* and *cpMax* members are equal, the range is empty. The range includes everything if *cpMin* is 0 and *cpMax* is –1. |
| *pwszText* | A **WSTRING** containing the text to find. |

#### Return value

First overloaded function: If the target string is found, the return value is the zero-based position of the first character of the match. If the target is not found, the return value is -1.

Second overloaded function: A **CHARRANGE** structure. The range of characters in which the text was found. If the text was not found, *cpMin* and *cpMax* are -1.

#### Remarks

**FindTextEx** uses the [FINDTEXTEXW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-findtextexw) structure, while **FindTex** uses the [FINDTEXTW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-findtextw) structure. The difference is that **FindTextEx** reports the range of text that was found.

#### Usage examples

Search in a range:
```
DIM cws AS CWSTR =  "box"
DIM chrg AS CHARRANGE = pRichEdit.FindTextEx(FR_DOWN, 5, 20, cws)
pRichEdit.ExSetSel(chrg.cpMin, chrg.cpMax)
```
Search in the entire document:
```
DIM cws AS CWSTR =  "box"
DIM chrg AS CHARRANGE = pRichEdit.FindTextEx(FR_DOWN, 0, -1, cws)
--or--
DIM chrg AS CHARRANGE = pRichEdit.FindTextEx(FR_DOWN, , , cws)
pRichEdit.ExSetSel(chrg.cpMin, chrg.cpMax)
```
Search backwards from the end of the document to the beginning:
```
DIM cbStart AS LONG = pRichEdit.GetTextLength
DIM chrg AS CHARRANGE = pRichEdit.FindTextEx(0, cbStart, 0, cws)
pRichEdit.ExSetSel(chrg.cpMin, chrg.cpMax)
```

---

# <a name="findwordbreak"></a>FindWordBreak

Finds the next word break before or after the specified character position or retrieves information about the character at that position.
```
FUNCTION FindWordBreak (BYVAL fOperation AS DWORD, BYVAL dwStartPos AS DWORD) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *fOperation* | Specifies the find operation. This parameter can be one of the following values.<br>**WB_CLASSIFY**. Returns the character class and word-break flags of the character at the specified position.<br>**WB_ISDELIMITER**. Returns **TRUE** if the character at the specified position is a delimiter, or **FALSE** otherwise.<br>**WB_LEFT**. Finds the nearest character before the specified position that begins a word.<br>**WB_LEFTBREAK**. Finds the next word end before the specified position. This value is the same as WB_PREVBREAK.<br>**WB_MOVEWORDLEFT**. Finds the next character that begins a word before the specified position. This value is used during CTRL+LEFT ARROW key processing. This value is the similar to WB_MOVEWORDPREV. See Remarks for more information.<br>**WB_MOVEWORDRIGHT**. Finds the next character that begins a word after the specified position. This value is used during CTRL+right key processing. This value is similar to WB_MOVEWORDNEXT. See Remarks for more information.<br>**WB_RIGHT**. Finds the next character that begins a word after the specified position.<br>**WB_RIGHTBREAK**. Finds the next end-of-word delimiter after the specified position. This value is the same as WB_NEXTBREAK. |
| *dwStartPos* | Zero-based character starting position. |

#### Return value

The method returns a value based on the *fOperation* parameter.

| Return code  | Description |
| ------------ | ----------- |
| **WB_CLASSIFY** | Returns the character class and word-break flags of the character at the specified position. |
| **WB_ISDELIMITER** | Returns **TRUE** if the character at the specified position is a delimiter; otherwise it returns **FALSE**. |
| **Others** | Returns the character index of the word break. |

#### Remarks

If *fOperation* is **WB_LEFT** and **WB_RIGHT**, the word-break procedure finds word breaks only after delimiters. This matches the functionality of an edit control. If *fOperation* is **WB_MOVEWORDLEFT** or **WB_MOVEWORDRIGHT**, the word-break procedure also compares character classes and word-break flags.

For information about character classes and word-break flags, see [Word and Line Breaks](https://learn.microsoft.com/en-us/windows/win32/controls/use-word-and-line-break-information).

---

# <a name="formatrange"></a>FormatRange

Formats a range of text in a rich edit control for a specific device.
```
FUNCTION FormatRange (BYVAL fRender AS LONG, BYREF fr AS FORMATRANGE) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *fRender* | Specifies whether to render the text. If this parameter is not zero, the text is rendered. Otherwise, the text is just measured. |
| *fr* | A [FORMATRANGE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-formatrange) structure containing information about the output device, or **NULL** to free information cached by the control. |

#### Return value

This method returns the index of the last character that fits in the region, plus 1.

#### Remarks

This method is typically used to format the content of rich edit control for an output device such as a printer.

After using this method to format a range of text, it is important that you free cached information by sending **FormatRange** again, but with *fr* set to **NULL**; otherwise, a memory leak will occur. Also, after using this message for one device, you must free cached information before using it again for a different device.

---

# <a name="getcharfrompos"></a>GetCharFromPos

Gets information about the character closest to a specified point in the client area of an edit control.
```
FUNCTION GetCharFromPos (BYREF ptl AS .POINTL) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *ptl* | A [POINTL](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-pointl) structure that contains the horizontal and vertical coordinates of a point in the control's client area. The coordinates are in screen units and are relative to the upper-left corner of the control's client area. |

#### Return value

The return value specifies the zero-based character index of the character nearest the specified point. The return value indicates the last character in the edit control if the specified point is beyond the last character in the control.

---

# <a name="getimecolor"></a>GetIMEColor

Gets the Input Method Editor (IME) composition color. This message is available only in Asian-language versions of the operating system.

```
FUNCTION CRichEditCtrl.GetIMEColor (BYVAL rgCmpclr AS .COMPCOLOR PTR) AS LONG
   RETURN this.SetResult(SendMessageW(m_hRichEdit, EM_GETIMECOLOR, 0, cast(LPARAM, rgCmpclr)))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *rgCmpclr* | A pointer to a four-element array of [COMPCOLOR](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-compcolor) structures. |

#### Return value

If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero. Call **GetLastResult** and/or **GetErrorInfo** to get information about the result.

#### Note

This message is supported only in Asian-language versions of Microsoft Rich Edit 1.0. It is not supported in any later versions.

---

# <a name="getimecompMode"></a>GetIMECompMode

Gets the current IME mode for a rich edit control.
```
FUNCTION GetIMECompMode () AS DWORD
```
#### Return value

The return value is one of the following values.

| Return code | Description |
| ----------- | ----------- |
| **ICM_NOTOPEN** | IME is not open. |
| **ICM_LEVEL3** | True inline mode. |
| **ICM_LEVEL2** | Level 2. |
| **ICM_LEVEL2_5** | Level 2.5. |
| **ICM_LEVEL2_SUI** | Special UI. |

# <a name="getimecompText"></a>GetIMECompText

Gets the Input Method Editor (IME) composition text.
```
FUNCTION GetIMECompText (BYREF ict AS .IMECOMPTEXT, BYVAL buffer AS ANY PTR) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *ict* | A [IMECOMPTEXT](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-imecomptext) structure. |
| *buffer* | The buffer that receives the composition text. The size of this buffer is contained in the *cb* member of the **IMECOMPTEXT** structure. |

#### Return value

If successful, the return value is the number of Unicode characters copied to the buffer. Otherwise, it is zero.

#### Remarks

This message only takes Unicode strings.

**Security Warning**: Be sure to have a buffer sufficient for the size of the input. Failure to do so could cause problems for your application.

---

# <a name="getimeproperty"></a>GetIMEProperty

Gets the property and capabilities of the Input Method Editor (IME) associated with the current input locale.
```
FUNCTION GetIMEProperty (BYVAL figp AS DWORD) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *figp* | Specifies the type of property information to retrieve. This parameter can be one of the following values. |

| Value  | Meaning |
| ------ | ----------- |
| **IGP_PROPERTY** | Property information. |
| **IGP_CONVERSION** | Conversion capabilities. |
| **IGP_SENTENCE** | Sentence mode capabilities. |
| **IGP_UI** | User interface capabilities. |
| **IGP_SETCOMPSTR** | Composition string capabilities. |
| **IGP_SELECT** | Selection inheritance capabilities. |
| **IGP_GETIMEVERSION** | Retrieves the system version number for which the specified IME was created. |

#### Return value

Returns the property or capability value, depending on the value of the *figp* parameter.

#### Remarks

If *figp* is **IGP_PROPERTY**, it returns one or more of the following values.

| Requirement  | Meaning |
| ------------ | ----------- |
| **IME_PROP_AT_CARET** | If set, conversion window is at the caret position. If clear, the window is near caret position. |
| **IME_PROP_SPECIAL_UI** | If set, IME has a nonstandard user interface. The application should not draw in the IME window. |
| **IME_PROP_CANDLIST_START_FROM_1** | If set, strings in the candidate list are numbered starting at 1. If clear, strings start at zero. |
| **IME_PROP_UNICODE** | If set, the IME is viewed as a UnicodeIME. The system and the IME will communicate through the UnicodeIME interface. If clear, IME will use the ANSI interface to communicate with the system. |
| **IME_PROP_COMPLETE_ON_UNSELECT** | If set, conversion window is at the caret position. If clear, the window is near caret position. |
| **IME_PROP_ACCEPT_WIDE_VKEY** | If set, the IME processes the injected Unicode that came from the **SendInput** function by using VK_PACKET. If clear, the IME might not process the injected Unicode, and the injected Unicode might be sent to the application directly. |

If *figp* is **IGP_UI**, it returns one or more of the following values.

| Requirement  | Meaning |
| ------------ | ----------- |
| **UI_CAP_2700** | Supports text escapement values of 0 or 2700. For more information, see **lfEscapement**. |
| **UI_CAP_ROT90** | Supports text escapement values of 0, 900, 1800, or 2700. For more information, see **lfEscapement**. |
| **UI_CAP_ROTANY** | Supports any text escapement value. For more information, see **lfEscapement**. |
	
If *figp* is **IGP_SETCOMPSTR**, it returns one or more of the following values.

| Requirement  | Meaning |
| ------------ | ----------- |
| **SCS_CAP_COMPSTR** | Can create the composition string by calling the **ImmSetCompositionString** function with the SCS_SETSTR value. |
| **SCS_CAP_MAKEREAD** | Can create the reading string from corresponding composition string when using the **ImmSetCompositionString** function with SCS_SETSTR and without setting *lpRead*. |
| **SCS_CAP_SETRECONVERTSTRING** | This IME can support reconversion. Use ImmSetCompositionString to do the reconversion. |

If *figp* is **IGP_SELECT**, it returns one or more of the following values.

| Requirement  | Meaning |
| ------------ | ----------- |
| **SELECT_CAP_CONVMODE** | Inherits conversion mode when a new IME is selected. |
| **SELECT_CAP_SENTENCE** | Inherits sentence mode when a new IME is selected. |

If *figp* is **IGP_GETIMEVERSION**, it returns one or more of the following values.

| Requirement  | Meaning |
| ------------ | ----------- |
| **IMEVER_0310** | The IME was created for Windows 3.1. |
| **IMEVER_0400** | The IME was created for Windows 95 or later. |

This message is similar to **ImmGetProperty**, except that it uses the current input locale. The application should call **IsIME** method before calling this function.

---

# <a name="getline"></a>GetLine

Copies a line of text from a rich edit control.
```
FUNCTION GetLine (BYVAL which AS DWORD) AS CWSTR
```
| Parameter  | Description |
| ---------- | ----------- |
| *which* | The zero-based index of the line to retrieve from a multiline edit control. A value of zero specifies the topmost line. |

#### Return value

A copy of the line.

---

# <a name="getlinecount"></a>GetLineCount

Gets the number of lines in a multiline rich edit control.
```
FUNCTION GetLineCount () AS LONG
```
#### Return value

The return value is an integer specifying the total number of text lines in the multiline edit control or rich edit control. If the control has no text, the return value is 1. The return value will never be less than 1.

### Remarks

The **GetLineCount** method retrieves the total number of text lines, not just the number of lines that are currently visible.

If the Wordwrap feature is enabled, the number of lines can change when the dimensions of the editing window change.

---

# <a name="getoleinterface"></a>GetOleInterface

Retrieves an IRichEditOle object that a client can use to access a rich edit control's Component Object Model (COM) functionality.
```
FUNCTION GetOleInterface () AS IRichEditOle PTR
```
#### Return value

A pointer to the **IRichEditOle** interface. The control calls the **AddRef** method for the object before returning, so the calling application must call the **Release** method when it is done with the object. |

---

# <a name="getsel"></a>GetSel

Gets the starting and ending character positions of the current selection in a rich edit control.
```
FUNCTION GetSel OVERLOAD (BYREF dwStartPos AS DWORD, BYREF dwEndPos AS DWORD) AS LONG
FUNCTION GetSel OVERLOAD () AS .POINTL
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwStartPos* | A **DWORD** variable that receives the starting position of the selection. This parameter can be **NULL**. |
| *dwEndPos* | A **DWORD** variable that receives the position of the first unselected character after the end of the selection. This parameter can be **NULL**. |

#### Return value

The return value is a zero-based value with the starting position of the selection in the **LOWORD** and the position of the first character after the last selected character in the **HIWORD**. If either of these values exceeds 65,535, the return value is -1. It is better to use the values returned in *dwStartPos* and *dwEndPos* because they are full 32-bit values.

The second overloaded method returns a [POINTL](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-pointl) structure. The **x** and **y** members of this structure receive the starting and ending positions of the selection.

#### Remarks

If there is no selection, the starting and ending values are both the position of the caret.

You can also use the **ExtGetSel** method to retrieve the same information. **ExtGetSel** also returns starting and ending character positions as 32-bit values. A combination of the use of **ExtGetSel** and **GetSelText** are used in the **GetSelText** method to retrieve the selected text as a **CWSTR**.

---

# <a name="getseltext"></a>GetSelText

Retrieves the currently selected text in a rich edit control.
```
FUNCTION GetSelText () AS CWSTR
```
#### Return value

The selected text as a **CWSTR** (dynamic unicode string).

# <a name="gettableparams"></a>GetTableParams

Retrieves the table parameters for a table row and the cell parameters for the specified number of cells.
```
FUNCTION GetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *tp* | A [TABLEROWPARMS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-tablerowparms) structure. |
| *tcp* | A [TABLECELLPARMS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-tablecellparms) structure. |

#### Return value

Returns S_OK if successful, or one of the following error codes.

| Return code  | Description |
| ------------ | ----------- |
| **E_FAIL** | Changes cannot be made. This can occur if the control is a plain-text or single-line control, or if the insertion point is inside a math object. It also occurs if tables are disabled if the **EditStyleEx** property sets the **SES_EX_NOTABLE** value. |
| **E_INVALIDARG** | The *lptp* or *lptcp* parameters are NULL or point to an invalid structure. The **cbRow** member of the **TABLEROWPARMS** structure must equal sizeof(TABLEROWPARMS) or sizeof(TABLEROWPARMS) 2*sizeof(long). The latter value is the size of the RichEdit 4.1 **TABLEROWPARMS** structure. The **cbCell** member of the **TABLEROWPARMS** structure must equal sizeof(TABLECELLPARMS). The query character position must be at a table row delimiter. |
| **E_OUTOFMEMORY** | Insufficient memory is available. |

#### Remarks

This method gets the table parameters for the row at the character position specified by the **cpStartRow** member of the **TABLEROWPARMS** structure, and the number of cells specified by the **cCells** member of the **TABLECELLPARMS** structure.

The character position specified by the **cpStartRow** member of the **TABLEROWPARMS** structure should be at the start of the table row, or at the end delimiter of the table row. If **cpStartRow** is set to 1, the character position is given by the current selection. In this case, position the selection at the end of the row (between the cell mark and the end delimiter of the table row), or select the row.

---

# <a name="gettextex"></a>GetTextEx

Gets all of the text from the rich edit control in any particular code base you want.
```
FUNCTION GetTextEx (BYREF gtex AS GETTEXTEX, BYVAL buffer AS ANY PTR) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *gtex* | A [GETTEXTEX](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-gettextex) structure which indicates how to translate the text before putting it into the output buffer. |
| *buffer* | Pointer to the buffer to receive the text. The size of this buffer, in bytes, is specified by the *cb* member of the [GETTEXTEX](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-gettextex) structure. Use the **GetTextLength** method to get the required size of the buffer. |

#### Return value

The return value is the number of characters copied into the output buffer, not including the null terminator.

#### Remarks

If the size of the output buffer is less than the size of the text in the control, the edit control will copy text from its beginning and place it in the buffer until the buffer is full. A terminating null character will still be placed at the end of the buffer.

---

# <a name="gettextrange"></a>GetTextRange

Overloaded method that retrieves a specified range of characters from a rich edit control.
```
FUNCTION GetTextRange (BYREF trg AS .TEXTRANGEW) AS DWORD
FUNCTION GetTextRange (BYVAL cpMin AS LONG = 0, BYVAL cpMax AS LONG = -1) AS CWSTR
```
| Parameter  | Description |
| ---------- | ----------- |
| *trg* | A [TEXTRANGEW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-textrangew) structure that specifies the range of characters to retrieve and a buffer to copy the characters to. |

| Parameter  | Description |
| ---------- | ----------- |
| *cpMin* | Character position index immediately preceding the first character in the range. |
| *cpMax* | Character position immediately following the last character in the range. If the *cpMin* and *cpMax* members are equal, the range is empty. The range includes everything if *cpMin* is 0 and *cpMax* is –1. |

#### Return value

The first overloaded method returns the number of characters copied, not including the terminating null character.

The second overloaded method returns the retrieved text.

#### Usage examples

Retrieve all the text using the first overloaded method:

```
DIM cbLen AS LONG = pRichEdit.GetTextLEngth
DIM wstrText AS CWSTR = WSPACE(cbLen + 1)
DIM trg AS TEXTRANGEW
trg.chrg.cpMin = 0
trg.chrg.cpMax = -1
trg.lpstrText = wstrtext
pRichEdit.GetTextRange(trg)
AfxMsg wstrtext
```
Retrieve all the text using the second overloaded method:
```
DIM cws AS CWSTR = pRichEdit.GetTextRange
AfxMsg cws
```
Retrieve a range of text.
```
DIM cws AS CWSTR = pRichEdit.GetTextRange(5, 15)
AfxMsg cws
```

---

# <a name="getthumb"></a>GetThumb

Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline rich edit control.
```
FUNCTION GetThumb () AS LONG
```
#### Return value

The return value is the position of the scroll box.

# <a name="undoname"></a>UndoName

Retrieves the type of the next undo action, if any.

```
FUNCTION UndoName () AS DWORD
FUNCTION GetUndoName () AS DWORD
```

#### Return value

If there is an undo action, the value returned is an [UNDONAMEID](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ne-richedit-undonameid) enumeration value that indicates the type of the next action in the control's undo queue.

If there are no actions that can be undone or the type of the next undo action is unknown, the return value is zero.

#### UNDONAMEID enumeration
4
| Name              | Value | Description |
| ----------------- | ----- | ----------- |
| **UID_UNKNOWN**   |   0   | The type of undo action is unknown. |
| **UID_TYPING**    |   1   | Typing operation. |
| **UID_DELETE**    |   2   | Delete operation. |
| **UID_DRAGDROP**  |   3   | Drag-and-drop operation. |
| **UID_CUT**       |   4   | Cut operation. |
| **UID_PASTE**     |   5   | Paste operation. |
| **UID_AUTOTABLE** |   6   | Automatic table insertion; for example, typing +---+---+<Enter> to insert a table row. |

#### Remarks

The types of actions that can be undone or redone include typing, delete, drag, drop, cut, and paste operations. This information can be useful for applications that provide an extended user interface for undo and redo operations, such as a drop-down list box of actions that can be undone.

---

# <a name="getzoom"></a>GetZoom

Gets the current zoom ratio, which is always between 1/64 and 64.
```
FUNCTION GetZoom (BYREF pzNum AS DWORD, BYREF pzDen AS DWORD) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| zNum* | A **DWORD** variable that receives the numerator of the zoom ratio. |
| zDen* | A **DWORD** variable that receives the denominator of the zoom ratio.. |

#### Return value
The message returns **TRUE** if message is processed, which it will be if both *pzNum* and *pzDen* are not **NULL**.

---

# <a name="hideselection"></a>HideSelection

Hides or shows the selection in a rich edit control.
```
SUB HideSelection (BYVAL fHide AS DWORD)
```
| Parameter  | Description |
| ---------- | ----------- |
| *fHide* | Value specifying whether to hide or show the selection. If this parameter is zero, the selection is shown. Otherwise, the selection is hidden. |

#### Return value

This message does not return a value.

---

# <a name="insertimage"></a>InsertImage

Overloaded method that replaces the selection with a blob that displays an image. It uses pixels instead of HIMETRIC units and it is DPI aware. Can insert and display .bmp, .jpg, .png and .gif files.
```
FUNCTION InsertImage OVERLOAD (BYREF wszFileName AS WSTRING, BYVAL xWidth AS LONG = 0, BYVAL yHeight AS LONG = 0, _
BYVAL Ascent AS LONG = 0, BYVAL nType AS LONG = TA_BASELINE, BYREF wszAlternateText AS WSTRING = "") AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Path and name of the image file to load. |
| *xWidth* | Optional. Image width in pixels. If both *xWidth* and *yWidth* parameters are 0, the method retrieves the size of the image using GDI+. |
| *yHeight* | Optional. Image heigh in pixels. If both *xWidth* and *yWidth* parameters are 0, the method retrieves the size of the image using GDI+. |
| *Ascent* | Optional. If *nType* is TA_BASELINE (the default value), this parameter is the distance, in pixels, that the top of the image extends above the text baseline. If *nType* is TA_BASELINE and ascent is zero, the bottom of the image is placed at the text baseline. |
| *nType* | Optional. The vertical alignment of the image. It can be one of the following values.<br>**TA_BASELINE**. Align the image relative to the text baseline.<br>**TA_BOTTOM**. Align the bottom of the image at the bottom of the text line.<br>**TA_TOP**. Align the top of the image at the top of the text line. |
| *wszAlternateText* | Optional. The alternate text for the image. |

#### Return value

Returns a boolean value: 0 for success and -1 for failure. To get extended error information call **GetLastResult** and/or **GetErrorInfo**, which can return **S_OK** if successful, or one of the following error codes.

| Return code  | Description |
| ------------ | ----------- |
| **E_FAIL** | Cannot insert the image. |
| **E_INVALIDARG** | Invaid argument |
| **E_OUTOFMEMORY** | Insufficient memory is available. |

---

Overloaded method that replaces the selection with a blob that displays an image. Uses HIMETRIC units.
```
FUNCTION InsertImage OVERLOAD (BYREF ip AS RICHEDIT_IMAGE_PARAMETERS) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *ip* | A [RICHEDIT_IMAGE_PARAMETERS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-richedit_image_parameters) structure that contains the image blob. |

#### Return value

Returns S_OK if successful, or one of the following error codes.

| Return code  | Description |
| ------------ | ----------- |
| **E_FAIL** | Cannot insert the image. |
| **E_INVALIDARG** | The *ip* parameter is NULL or points to an invalid image. |
| **E_OUTOFMEMORY** | Insufficient memory is available. |

# <a name="insertobject"></a>InsertObject

Inserts an image or a Ole object in the rich edit control.
See MSDN: https://learn.microsoft.com/en-us/windows/win32/controls/using-rich-edit-com
Remarks: To insert images, use the **InsertImage** method, because the **InsertObject** method won't display the image, but an icon, when the image type is not a bitmap.
```
FUNCTION InsertObject (BYREF wszFileName AS WSTRING) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | The path and file name of the image file. |

#### Return value

Returns a boolean true (-1) for success of false (0) for failure. To get extended information call the **GetLastResult** and/or the **GetErrorInfo** methods.

---

# <a name="inserttable"></a>InsertTable

Inserts one or more identical table rows with empty cells.
```
FUNCTION InsertTable (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *tp* | A [TABLEROWPARMS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-tablerowparms) structure. |
| *tcp* | A [TABLECELLPARMS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-tablecellparms) structure. |

#### Return value

Returns **S_OK** if the table is inserted, or an error code if not.

#### Remarks

If the **cpStartRow** member of the **TABLEROWPARMS** is –1, this message deletes the selected text (if any), and then inserts empty table rows with the row and cell parameters given by *lptp^* and *lptcp*. It leaves the selection pointing to the start of the first cell in the first row. The client can then populate the table cells by pointing the selection (or an **ITextRange**) to the various cell end marks and inserting and formatting the desired text. Such text can include nested table rows. Alternatively, if the **cpStartRow** member of the **TABLEROWPARMS** is 0 or greater, table rows are inserted at the character position given by **cpStartRow**. This only changes the current selection if the table is inserted inside the selected text.

A Microsoft Rich Edit table consists of a sequence of table rows which, in turn, consist of sequences of paragraphs. A table row starts with the special two-character delimiter paragraph U+FFF9 U+000D and ends with the two-character delimiter paragraph U+FFFB U+000D. Each cell is terminated by the cell mark U+0007, which is treated as a hard end-of-paragraph mark just as U+000D (CR) is. The table row and cell parameters are treated as special paragraph formatting of the table-row delimiters. The formatting contains the information in the **TABLEROWPARMS** structure. The cell parameters given by the **TABLECELLPARMS** structure are stored in an expanded version of the tabs array. This format allows tables to be nested within other tables, up to fifteen levels deep.

---

# <a name="isime"></a>IsIME

Determines if current input locale is an East Asian locale.
```
FUNCTION IsIME () AS LONG
```
#### Return value

Returns **TRUE** if it is an East Asian locale. Otherwise, it returns **FALSE**.

---

# <a name="linefromchar"></a>LineFromChar

Gets the index of the line that contains the specified character index in a multiline rich edit control.
```
FUNCTION LineFromChar (BYVAL index AS DWORD) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *index* | The character index of the character contained in the line whose number is to be retrieved. If this parameter is -1, **LineFromChar** retrieves either the line number of the current line (the line containing the caret) or, if there is a selection, the line number of the line containing the beginning of the selection. |

#### Return value

The return value is the zero-based line number of the line containing the character index specified by *index*.

---

# <a name="lineindex"></a>LineIndex

Gets the character index of the first character of a specified line in a multiline rich edit control.
```
FUNCTION LineIndex (BYVAL nLine AS LONG) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *nLine* | The zero-based line number. A value of -1 specifies the current line number (the line that contains the caret). |

#### Return value

The return value is the character index of the line specified in the *nLine* parameter, or it is -1 if the specified line number is greater than the number of lines in the edit control.

---

# <a name="linelength"></a>LineLength

Retrieves the length, in characters, of a line in a rich edit control.
```
FUNCTION LineLength (BYVAL index AS DWORD) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *index* | The character index of a character in the line whose length is to be retrieved. If this parameter is greater than the number of characters in the control, the return value is zero.<br>This parameter can be -1. In this case, the message returns the number of unselected characters on lines containing selected characters. For example, if the selection extended from the fourth character of one line through the eighth character from the end of the next line, the return value would be 10 (three characters on the first line and seven on the next). |

#### Return value

If *index* is greater than the number of characters in the control, the return value is zero.

---

# <a name="linescroll"></a>LineScroll

Scrolls the text in a multiline rich edit control.
```
FUNCTION LineScroll (BYVAL y AS LONG) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *y* | The number of lines to scroll vertically. |

#### Return value

**TRUE** or **FALSE**.

#### Remarks

The control does not scroll vertically past the last line of text in the edit control. If the current line plus the number of lines specified by the *y* parameter exceeds the total number of lines in the edit control, the value is adjusted so that the last line of the edit control is scrolled to the top of the edit-control window.

---

# <a name="margins"></a>Margins

Sets the widths of the left and right margins for a rich edit control. The message redraws the control to reflect the new margins.

```
FUNCTION Margins (BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
SUB SetMargins (BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
PROPERTY LeftMargin (BYVAL nWidth AS LONG)
PROPERTY RightMargin (BYVAL nWidth AS LONG)
SUB SetLeftMargin (BYVAL nWidth AS LONG)
SUB SetRightMargin (BYVAL nWidth AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMargins* | The margins to set. This parameter can be one or more of the following values.<br>**EC_LEFTMARGIN**. Sets the left margin.<br>**EC_RIGHTMARGIN**. Sets the right margin.<br>**EC_USEFONTINFO**. Sets the left and right margins to a narrow width calculated using the text metrics of the control's current font. If no font has been set for the control, the margins are set to zero. The *nWidth* parameter is ignored. |
| *nWidth* | The **LOWORD** specifies the new width of the left margin, in pixels. This value is ignored if *nMargins* does not include **EC_LEFTMARGIN**.<br>**Rich Edit 3.0 and later**: The **LOWORD** can specify the **EC_USEFONTINFO** value to set the left margin to a narrow width calculated using the text metrics of the control's current font. If no font has been set for the control, the margin is set to zero.<br>The **HIWORD** specifies the new width of the right margin, in pixels. This value is ignored if *nMargins* does not include **EC_RIGHTMARGIN**.<br>**Rich Edit 3.0 and later**: The **HIWORD** can specify the **EC_USEFONTINFO** value to set the right margin to a narrow width calculated using the text metrics of the control's current font. If no font has been set for the control, the margin is set to zero |

#### Usage examples

Set the left margin
```
pRichEdit.SetMargins(EC_LEFTMARGIN, MAKELONG(50, 0))
--or--
pRichEdit.SetMargins(EC_LEFTMARGIN, MAKELONG(50))
--or--
pRichEdit.Margins(EC_LEFTMARGIN) = MAKELONG(50, 0)
--or--
pRichEdit.LeftMargin = 50
```
Set the right margin
```
pRichEdit.SetMargins(EC_RIGHTMARGIN, MAKELONG(0, 50))
--or--
pRichEdit.Margins(EC_RIGHTMARGIN) = MAKELONG(0, 50)
--or--
pRichEdit.RightMargin = 50
```
Set both margins
```
pRichEdit.SetMargins(EC_LEFTMARGIN OR EC_RIGHTMARGIN, MAKELONG(50, 50))
--or--
pRichEdit.Margins(EC_LEFTMARGIN OR EC_RIGHTMARGIN) = MAKELONG(50, 50)
```

---

# <a name="pastespecial"></a>PasteSpecial

Pastes a specific clipboard format in a rich edit control.
```
SUB PasteSpecial OVERLOAD (BYVAL clpfmt AS DWORD, BYREF rps AS REPASTESPECIAL)
SUB PasteSpecial OVERLOAD (BYVAL clpfmt AS DWORD, BYVAL dwAspect AS DWORD, BYVAL hMF AS HMETAFILE)
```
| Parameter  | Description |
| ---------- | ----------- |
| *clpfmt* | Specifies the [ Clipboard Formats](https://learn.microsoft.com/en-us/windows/win32/dataxchg/clipboard-formats). |
| *rps* | A [REPASTESPECIAL](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-repastespecial) structure or **NULL**. If an object is being pasted, the [REPASTESPECIAL](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-repastespecial) structure is filled in with the desired display aspect. If *clpfmt* is **NULL** or the *dwAspect* member is zero, the display aspect used will be the contents of the object descriptor. |
| *dwAspect* | Display aspect. It can be one of the following values.<br>**DVASPECT_CONTENT**. Aspect is based on the content of the object.<br>**DVASPECT_ICON**. Aspect is based on the icon view of the object. |
| *hMF* | Aspect data. If *dwAspect* is **DVASPECT_ICON**, this member contains the handle to the metafile with the icon view of the object. |

---

# <a name="posfromchar"></a>PosFromChar

Retrieves the client area coordinates of a specified character in a rich edit control.
```
FUNCTION PosFromChar (BYVAL index as DWORD) AS .POINTL
```
| Parameter  | Description |
| ---------- | ----------- |
| *index* | The zero-based index of the character. |

#### Return value

A [POINTL](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-pointl) structure with receives the client area coordinates of the character. The coordinates are in screen units and are relative to the upper-left corner of the control's client area.

#### Remarks

A returned coordinate can be a negative value if the specified character is not displayed in the edit control's client area. The coordinates are truncated to integer values.

If the character is a line delimiter, the returned coordinates indicate a point just beyond the last visible character in the line. If the specified index is greater than the index of the last character in the control, the control returns -1.

---

# <a name="reconversion"></a>Reconversion

Invokes the Input Method Editor (IME) reconversion dialog box.
```
SUB Reconversion ()
```
#### Return value

Thismethod does not return a value.

---

# <a name="redo"></a>Redo

Redoes the next action in the control's redo queue.
```
FUNCTION Redo () AS LONG
```
#### Return value

If the **Redo** operation succeeds, the return value is a nonzero value. If the **Redo** operation fails, the return value is zero.

#### Remarks

To determine whether there are any actions in the control's redo queue, call the **CanRedo** method.

---

# <a name="replacesel"></a>ReplaceSel

Replaces the current selection in a rich edit control with the specified text.
```
SUB ReplaceSel (BYVAL bCanBeUndone AS LONG = TRUE, BYREF wszText AS WSTRING)
```
| Parameter  | Description |
| ---------- | ----------- |
| *bCanBeUndone* | Specifies whether the replacement operation can be undone. If this is **TRUE**, the operation can be undone. If this is **FALSE**, the operation cannot be undone. |
| *wszText* | A **WSTRING** containing the replacement text. |

#### Remarks

Use the **ReplaceSel** method to replace only a portion of the text in an edit control. To replace all of the text, use the **Text** property.

If there is no selection, the replacement text is inserted at the caret.

In a rich edit control, the replacement text takes the formatting of the character at the caret or, if there is a selection, of the first character in the selection.

---

# <a name="requestresize"></a>RequestResize

Forces a rich edit control to send an **EN_REQUESTRESIZE** notification message to its parent window.
```
SUB RequestResize ()
```
#### Remarks

This message is useful during **WM_SIZE** processing for the parent of a bottomless rich edit control.

---

# <a name="scroll"></a>Scroll

Scrolls the text vertically in a multiline rich edit control.
```
FUNCTION Scroll (BYVAL nAction AS LONG) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *nAction* | The action the scroll bar is to take. This parameter can be one of the following values:<br>**SB_LINEDOWN**. Scrolls down one line.<br>**SB_LINEUP**. Scrolls up one line.<br>**SB_PAGEDOWN**. Scrolls down one page.<br>**SB_PAGEUP**. Scrolls up one page. |

#### Return value

If the method is successful, the **HIWORD** of the return value is **TRUE**, and the **LOWORD** is the number of lines that the command scrolls. The number returned may not be the same as the actual number of lines scrolled if the scrolling moves to the beginning or the end of the text. If the *nAction* parameter specifies an invalid value, the return value is **FALSE**.

#### Remarks

To scroll to a specific line or character position, use the **LineScroll** method. To scroll the caret into view, use the **ScrollCaret** method.

---

# <a name="scrollcaret"></a>ScrollCaret

Scrolls the caret into view in a rich edit control.
```
SUB ScrollCaret ()
```

#### Return value

The return value is not meaningful.

---

# <a name="selectiontype"></a>SelectionType

Determines the selection type for a rich edit control.
```
FUNCTION SelectionType () AS LONG
```
#### Return value

If the selection is not empty, the return value is a set of flags containing one or more of the following values.

| Return code  | Description |
| ------------ | ----------- |
| **SEL_TEXT** | Text.|
| **SEL_OBJECT** | At least one COM object. |
| **SEL_MULTICHAR** | More than one character of text. |
| **SEL_MULTIOBJECT** | More than one COM object. |

#### Remarks

This message is useful during **WM_SIZE** processing for the parent of a bottomless rich edit control.

---

# <a name="setbkgndcolor"></a>SetBkgndColor

Sets the background color for a rich edit control.
```
FUNCTION SetBkgndColor (BYVAL SysColor AS DWORD, BYVAL BkColor AS DWORD) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *SysColor* | Specifies whether to use the system color. If this parameter is a nonzero value, the background is set to the window background system color. Otherwise, the background is set to the specified color. |
| *BkColor* | A [COLORREF](https://learn.microsoft.com/en-us/windows/win32/gdi/colorref) structure specifying the color if *SysColor* is zero. To generate a COLORREF, use the RGB macro (renamed as BGR in the FreeBasic Windows headers). |

#### Return value

This message returns the original background color.

---

# <a name="setfontsize"></a>SetFontSize

Sets the font size for the selected text.
```
FUNCTION SetFontSize (BYVAL ptsize AS LONG) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *ptsize* | Change in point size of the selected text. The result will be rounded according to values shown in the following table. This parameter should be in the range of -1637 to 1638. The resulting font size will be within the range of 1 to 1638. |

#### Return value

If no error occurred, the return value is TRUE. If an error occurred, the return value is FALSE.

#### Remarks

You can easily get the font size by calling the **CharFormat** property.

Rich Edit first adds *ptsize* to the current font size and then uses the resulting size and the following table to determine the rounding value.

| Band  | Rounding value |
| ----- | ----------- |
| <=12 | 1 |
| 28 | 2 |
| 36 | 0 |
| 48 | 0 |
| 72 | 0 |
| 80 | 0 |
| 80 | 10 |

If the resulting font size is not evenly divisible by the rounding value, the font size is then rounded to a number evenly divisible by the rounding value. So if the font size is less than or equal to 12, the rounding value will be 1. Similarly, if the font size is less than or equal to 28, the rounding value is 2. For values greater than 28, font sizes are rounded to the next band. So, the font size jumps to 36, 48, 72, 80. After 80, all rounding is done in increments of ten points.

The font size is rounded up or down depending on the sign of *ptsize*. If *ptsize* is positive, the rounding is always up. Otherwise, rounding is always down. So, if the current font size is 10 and *ptsize* is 3, the resulting font size would be 14 (10 + 3 = 13, which is not divisible by 2, so the size rounds up to 14). Conversely, if the current font size is 14 and *ptsize* is -3, the resulting font size would be 10 (14 - 3 = 11, which is not divisible by 2, so the size rounds down to 10).

The change is applied to each part of the selection. So, if some of the text is 10pt and some 20pt, after a call with *ptsize* set to 1, the font sizes become 11pt and 22pt, respectively.

Additional examples are shown in the following table.

| Original font size  | ptsize | Resulting font size |
| ------------------- | ------ | ------------------- |
| 7 | 1 | 8 |
| 7 | 3 | 10 |
| 10 | 3 | 14 |
| 14 | -3 | 10 |
| 28 | 1 | 36 |
| 28 | 3 | 36 |
| 80 | 1 | 90 |
| 80 | -1 | 72 |

# <a name="setimecolor"></a>SetIMEColor

Sets the Input Method Editor (IME) composition color. This message is available only in Asian-language versions of the operating system.
```
FUNCTION CRichEditCtrl.SetIMEColor (BYVAL pcompcolor AS .COMPCOLOR PTR) AS LONG
   RETURN this.SetResult(SendMessageW(m_hRichEdit, EM_SETIMECOLOR, 0, cast(LPARAM, pcompcolor)))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *pcompcolor* | Pointer to a [COMPCOLOR](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-compcolor) structure that contains the composition color to be set.. |

#### Return value

If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero. Call **GetLastResult** and/or **GetErrorInfo** to get information about the result.

#### Note

This message is supported only in Asian-language versions of Microsoft Rich Edit 1.0. It is not supported in any later versions.

---

# <a name="setolecallback"></a>SetOleCallback

Gives a rich edit control an IRichEditOleCallback object that the control uses to get OLE-related resources and information from the client.
```
FUNCTION SetOleCallback (BYVAL pCallback AS ANY PTR) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *pCallback* | Pointer to an **IRichEditOleCallback** object. The control calls the **AddRef** method for the object before returning. |

#### Return value

If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero.

---

# <a name="setpalette"></a>SetPalette

Changes the palette that a rich edit control uses for its display window.
```
SUB SetPalette (BYVAL newPalette AS HPALETTE)
```
| Parameter  | Description |
| ---------- | ----------- |
| *newPalette* | Handle to the new palette used by the rich edit control. |

#### Return value

This method does notreturn a value.

#### Remarks

The rich edit control does not check whether the new palette is valid.

---

# <a name="setreadonly"></a>SetReadOnly

Sets or removes the read-only style (ES_READONLY) of a rich edit control.
```
FUNCTION SetReadOnly (BYVAL fReadOnly AS LONG) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *fReadOnly* | Specifies whether to set or remove the **ES_READONLY** style. A value of **TRUE** sets the **ES_READONLY+* style; a value of **FALSE** removes the **ES_READONLY** style. |

#### Return value

Returns the boolean true (-1) or false (0).

#### Remarks

When an edit control has the **ES_READONLY** style, the user cannot change the text within the edit control.

To determine whether an edit control has the **ES_READONLY** style, use the Windows API **GetWindowLong** function with the **GWL_STYLE** flag.

---

# <a name="setsel"></a>SetSel

Selects a range of characters in a rich edit control.
```
SUB SetSel OVERLOAD (BYVAL nStart AS LONG, BYVAL nEnd AS LONG)
SUB SetSel OVERLOAD (BYREF pt AS POINTL)
```
| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The starting character position of the selection. |
| *nEnd* | The ending character position of the selection. |

| Parameter  | Description |
| ---------- | ----------- |
| *pt* | A [POINTL](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-pointl) structure with its **x** and **y** members set to the starting and ending characters positions of the selection. |

#### Remarks

The start value can be greater than the end value. The lower of the two values specifies the character position of the first character in the selection. The higher value specifies the position of the first character beyond the selection.

The start value is the anchor point of the selection, and the end value is the active end. If the user uses the SHIFT key to adjust the size of the selection, the active end can move but the anchor point remains the same.

If the start is 0 and the end is -1, all the text in the edit control is selected. If the start is -1, any current selection is deselected.

If the edit control has the **ES_NOHIDESEL** style, the selected text is highlighted regardless of whether the control has focus. Without the **ES_NOHIDESEL** style, the selected text is highlighted only when the edit control has the focus.

---

# <a name="settableparams"></a>SetTableParams

Changes the parameters of rows in a table.
```
FUNCTION SetTableParams (BYREF tp AS TABLEROWPARMS, BYREF tcp AS TABLECELLPARMS) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lptp* | A pointer to a [TABLEROWPARMS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-tablerowparms) structure. |
| *lptcp* | A pointer to a [TABLECELLPARMS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-tablecellparms) structure. |

#### Return value

Returns **S_OK** if successful, or one of the following error codes.

| Return code  | Description |
| ------------ | ----------- |
| **E_FAIL** | Changes cannot be made. This can occur if the control is a plain-text or single-line control, or if the insertion point is inside a math object. It also occurs if tables are disabled if the **EditStyleEx** property sets the **SES_EX_NOTABLE** value. |
| **E_INVALIDARG** | The *lptp* or *lptcp* parameters are NULL or point to an invalid structure. The **cbRow** member of the **TABLEROWPARMS** structure must equal sizeof(TABLEROWPARMS) or sizeof(TABLEROWPARMS) 2*sizeof(long). The latter value is the size of the RichEdit 4.1 **TABLEROWPARMS** structure. The **cbCell** member of the **TABLEROWPARMS** structure must equal sizeof(TABLECELLPARMS). The query character position must be at a table row delimiter. |
| **E_OUTOFMEMORY** | Insufficient memory is available. |

#### Remarks

This message changes the parameters of the number of rows specified by the **cRow** member of the **TABLEROWPARMS** structure, if the table has that many consecutive rows. If **cRow** is less than 0, the message iterates until the end of the table. If the new cell count differs from the current cell count by +1 or 1, it inserts or deletes the cell at the index specified by the **iCell** member of **TABLEROWPARMS**. The starting table row is identified by a character position. This position is specified by **cpStartRow** members with values that are greater than or equal to zero. The position should be inside the table row, but not inside a nested table, unless you want to change that table s parameters. If the **cpStartRow** member is 1, the character position is given by the current selection. For this, position the selection anywhere inside the table row, or select the row with the active end of the selection at the end of the table row.

---

# <a name="settabstops"></a>SetTabStops

Sets the tab stops in a multiline rich edit control.
```
FUNCTION SetTabStops (BYVAL nTabs AS LONG, BYVAL rgTabStops AS LONG_PTR) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *nTabs* | The number of tab stops contained in the array. If this parameter is zero, the *rgTabStops* parameter is ignored and default tab stops are set at every 32 dialog template units. If this parameter is 1, tab stops are set at every n dialog template units, where n is the distance pointed to by the *rgTabStops* parameter. If this parameter is greater than 1, *rgTabStops* is a pointer to an array of tab stops. |
| *rgTabStops* | A pointer to an array of unsigned integers specifying the tab stops, in dialog template units. If the *nTabs* parameter is 1, this parameter is a pointer to an unsigned integer containing the distance between all tab stops, in dialog template units. |

#### Return value

If all the tabs are set, the return value is *TRUE*.

If all the tabs are not set, the return value is *FALSE*.

#### Remarks

The **EM_SETTABSTOPS** message does not automatically redraw the edit control window. If the application is changing the tab stops for text already in the edit control, it should call the Windows API **InvalidateRect** function to redraw the edit control window.

The values specified in the array are in dialog template units, which are the device-independent units used in dialog box templates. To convert measurements from dialog template units to screen units (pixels), use the Windows API **MapDialogRect** function.

**Rich Edit**: Supported in Microsoft Rich Edit 3.0 and later. A rich edit control can have the maximum number of tab stops specified by **MAX_TAB_STOPS**.

---

# <a name="settargetdevice"></a>SetTargetDevice

Sets the target device and line width used for WYSIWYG formatting in a rich edit control.
```
FUNCTION SetTargetDevice (BYVAL hDC AS HDC, BYVAL lnwidth AS LONG) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *hDC* | HDC [Handle to a Device Context] for the target device. |
| *lnwidth* | Line width to use for formatting. |

#### Return value

If the method succeeds it returns the boolean value true (-1); if it fails, it returns false (0).

#### Remarks

The HDC for the default printer can be obtained as follows.

```
DIM hdc AS HDC
DIM pd AS PRINTDLGW
pd.lStructSize = SIZEOF(pd)
pd.flags = PD_RETURNDC OR PD_RETURNDEFAULT
IF PrintDlgW(@pd) THEN
   hdc =  = pd.hDC
END IF
```

If *lnwidth* is zero, no line breaks are created.

---

# <a name="settextex"></a>SetTextEx

Combines the functionality of WM_SETTEXT and EM_REPLACESEL and adds the ability to set text using a code page and to use either Rich Text Format (RTF) rich text or plain text.
```
FUNCTION SetTextExW (BYREF stex AS SETTEXTEX, BYVAL buffer AS ANY PTR) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *stex* | A [SETTEXTEX](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-settextex) structure that specifies flags and an optional code page to use in translating to Unicode. |
| *buffer* | Pointer to the null-terminated text to insert. This text is an ANSI string, unless the code page is 1200 (Unicode). If *pwszText* starts with a valid RTF ASCII sequence for example, "{\rtf" or "{urtf" the text is read in using the RTF reader. |

**SETTEXTEX flags**

Option flags. It can be any reasonable combination of the following flags.

| Flag  | Value | Description |
| ----- | ----- | ----------- |
| **ST_DEFAULT** | &h00 | Deletes the undo stack, discards rich-text formatting, replaces all text. |
| **ST_KEEPUNDO** | &h01 | Keeps the undo stack. |
| **ST_SELECTION** | &h02 | Replaces selection and keeps rich-text formatting. |
| **ST_NEWCHARS** | &h04 | Act as if new characters are being entered. |
| **ST_UNICODE** | &h08 | The text is UTF-16 (the WCHAR data type). |
| **ST_PLACEHOLDERTEXT** | &h10 | Placeholder text that is visible only when focus is not on the RichEdit control and the control does not contain any user-specified text. |
| **ST_PLAINTEXTONLY** | &h20 | RichEdit control supports plain text only. |

**SETTEXTEX codepage**

The code page used to translate the text to Unicode. If codepage is 1200 (Unicode code page), no translation is done. If codepage is CP_ACP, the system code page is used.

#### Return value

If the operation is setting all of the text and succeeds, the return value is 1.

If the operation is setting the selection and succeeds, the return value is the number of bytes or characters copied.

If the operation fails, the return value is zero.

---

# <a name="setuianame"></a>SetUIAName

Sets the name of a rich edit control for UI Automation (UIA).
```
FUNCTION SetUIAName (BYREF wszName AS WSTRING) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *wszName* | A **WSTRING** with the name. |

#### Return value

**TRUE** if the name for UIA is successfully set, otherwise **FALSE**.

---

# <a name="setundolimit"></a>SetUndoLimit

Sets the maximum number of actions that can stored in the undo queue.
```
FUNCTION SetUndoLimit (BYVAL maxactions AS DWORD) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *maxactions* | Specifies the maximum number of actions that can be stored in the undo queue. |

#### Return value

The return value is the new maximum number of undo actions for the rich edit control. This value may be less than *maxactions* if memory is limited.

#### Remarks

By default, the maximum number of actions in the undo queue is 100. If you increase this number, there must be enough available memory to accommodate the new number. For better performance, set the limit to the smallest possible value.

Setting the limit to zero disables the **Undo** feature.

---

# <a name="setzoom"></a>SetZoom

Sets the zoom ratio for a multiline edit control or a rich edit control. The ratio must be a value between 1/64 and 64. The edit control needs to have the **ES_EX_ZOOMABLE** extended style set, for this message to have an effect, see [Edit Control Extended Styles](https://learn.microsoft.com/en-us/windows/win32/controls/edit-control-window-extended-styles).

```
FUNCTION SetZoom (BYVAL zNum AS DWORD, BYVAL zDen AS DWORD) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *zNum* | Numerator of the zoom ratio. |
| *zDen* | Denominator of the zoom ratio. These parameters can have the following values. |

| Value  | Meaning |
| ------ | ------- |
| **Both 0** | Turns off zooming by using the **EM_SETZOOM** message (zooming may still occur using [TxGetExtent](https://learn.microsoft.com/en-us/windows/win32/api/textserv/nf-textserv-itexthost-txgetextent). |
| **1/64 < (wParam / lParam) < 64** | Zooms display by the zoom ratio numerator/denominator |

#### Return value

If the new zoom setting is accepted, the return value is **TRUE**. If the new zoom setting is not accepted, the return value is **FALSE**.

---

# <a name="showscrollbar"></a>ShowScrollBar

Shows or hides one of the scroll bars in the Text Host window.
```
SUB ShowScrollBar (BYVAL nScrollBar AS DWORD, BYVAL fShow AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *nScrollBar* | Identifies which scroll bar to display: horizontal or vertical. This parameter must be **SB_VERT** or **SB_HORZ**. |
| *fShow* | Specifies whether to show the scroll bar or hide it. Specify **TRUE** to show the scroll bar and **FALSE** to hide it. |

#### Remarks

This method is only valid when the control is in-place active. Calls made while the control is inactive may fail.

---

# <a name="stopgrouptyping"></a>StopGroupTyping

Stops the control from collecting additional typing actions into the current undo action.
```
FUNCTION StopGroupTyping () AS DWORD
```
#### Return value

The return value is zero. This message cannot fail.

#### Remarks
A rich edit control groups consecutive typing actions, including characters deleted by using the BackSpace key, into a single undo action until one of the following events occurs:

- The control receives an **EM_STOPGROUPTYPING** message.
- The control loses focus.
- The user moves the current selection, either by using the arrow keys or by clicking the mouse.
- The user presses the **Delete** key.
- The user performs any other action, such as a paste operation that does not involve typing.

You can send the **RichEdit_StopGroupTyping** message to break consecutive typing actions into smaller undo groups. For example, you could send **RichEdit_StopGroupTyping** after each character or at each word break.

---

# <a name="streamin"></a>StreamIn

Replaces the contents of a rich edit control with a stream of data provided by an application defined [EditStreamCallback](https://learn.microsoft.com/en-us/windows/win32/api/richedit/nc-richedit-editstreamcallback) callback function.
```
FUNCTION StreamIn (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *psf* | Specifies the data format and replacement options. This value must be one of the following values (see table below). |
| *pedst* | A [EDITSTREAM](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-editstream) structure. On input, the **pfnCallback** member of this structure must point to an application defined **EditStreamCallback** function. On output, the **dwError** member can contain a nonzero error code if an error occurred. |

#### Data format and replacement options

| Value  | Meaning |
| ------ | ------- |
| **SF_RTF** | RTF |
| **SF_TEXT** | Text |

In addition, you can specify the following flags.

| Value  | Meaning |
| ------ | ------- |
| **SFF_PLAINRTF** | If specified, only keywords common to all languages are streamed in. Language-specific RTF keywords in the stream are ignored. If not specified, all keywords are streamed in. You can combine this flag with the **SF_RTF** flag. |
| **SFF_SELECTION** | If specified, the data stream replaces the contents of the current selection. If not specified, the data stream replaces the entire contents of the control. You can combine this flag with the **SF_TEXT** or **SF_RTF** flags. |
| **SF_UNICODE** | **Rich Edit 2.0 and later**: Indicates Unicode text. You can combine this flag with the **SF_TEXT** flag. |
| **SF_USECODEPAGE** | **Rich Edit 3.0 and later**: Reads UTF-8 RTF and text using other code pages. The code page is set in the high word of *psf*. For example, for UTF-8 RTF, set *psf* to (CP_UTF8 << 16) | SF_USECODEPAGE | SF_RTF. |

#### Return value

This message returns the number of characters read.

#### Remarks

When you call the **StreamIn** methos, the rich edit control makes repeated calls to the **EditStreamCallback** function specified by the **pfnCallback** member of the [EDITSTREAM](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-editstream) structure. Each time the callback function is called, it fills a buffer with data to read into the control. This continues until the callback function indicates that the stream-in operation has been completed or an error occurs.

---

# <a name="streamout"></a>StreamOut

Causes a rich edit control to pass its contents to an application defined [EditStreamCallback](https://learn.microsoft.com/en-us/windows/win32/api/richedit/nc-richedit-editstreamcallback) callback function. The callback function can then write the stream of data to a file or any other location that it chooses.

```
FUNCTION StreamOut (BYVAL psf AS LONG, BYREF edst AS EDITSTREAM) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *psf* | Specifies the data format and replacement options. This value must be one of the following values (see table below). |
| *pedst* | A [EDITSTREAM](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-editstream) structure. On input, the **pfnCallback** member of this structure must point to an application defined **EditStreamCallback** function. On output, the **dwError** member can contain a nonzero error code if an error occurred. |

#### Data format and replacement options

| Value  | Meaning |
| ------ | ------- |
| **SF_RTF** | RTF |
| **SF_RTFNOOBJS** | RTF with spaces in place of COM objects. |
| **SF_TEXT** | Text |
| **SF_TEXTIZED** | Text with a text representation of COM objects. |

The **SF_RTFNOOBJS** option is useful if an application stores COM objects itself, as RTF representation of COM objects is not very compact. The control word, \objattph, followed by a space denotes the object position.

In addition, you can specify the following flags.

| Value  | Meaning |
| ------ | ------- |
| **SFF_PLAINRTF** | If specified, the rich edit control streams out only the keywords common to all languages, ignoring language-specific keywords. If not specified, the rich edit control streams out all keywords. You can combine this flag with the SF_RTF or **SF_RTFNOOBJS** flag. |
| **SFF_SELECTION** | If specified, the rich edit control streams out only the contents of the current selection. If not specified, the control streams out the entire contents. You can combine this flag with any of data format values. |
| **SF_UNICODE** | **Rich Edit 2.0 and later**: Indicates Unicode text. You can combine this flag with the **SF_TEXT** flag. |
| **SF_USECODEPAGE** | **Rich Edit 3.0 and later**: Reads UTF-8 RTF and text using other code pages. The code page is set in the high word of *psf*. For example, for UTF-8 RTF, set *psf* to (CP_UTF8 << 16) | SF_USECODEPAGE | SF_RTF. |

#### Return value

This message returns the number of characters written to the data stream.

#### Remarks

When you call the **StreamOut** method, the rich edit control makes repeated calls to the **EditStreamCallback** function specified by the **pfnCallback** member of the [EDITSTREAM](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-editstream) structure.  Each time it calls the callback function, the control passes a buffer containing a portion of the contents of the control. This process continues until the control has passed all its contents to the callback function, or until an error occurs.

---

# <a name="undo"></a>Undo

This message undoes the last edit control operation in the control's undo queue.
```
FUNCTION Undo () AS LONG
```
#### Return value

For a single-line edit control, the return value is always **TRUE**.

For a multiline edit control, the return value is **TRUE** if the undo operation is successful, or **FALSE** if the undo operation fails.

#### Remarks

**Edit controls and Rich Edit 1.0**: An undo operation can also be undone. For example, you can restore deleted text with the first **EM_UNDO** message, and remove the text again with a second **EM_UNDO** message as long as there is no intervening edit operation.

**Rich Edit 2.0 and later**: The undo feature is multilevel so sending two **EM_UNDO** messages will undo the last two operations in the undo queue. To redo an operation, send the **EM_REDO** message.

---

# <a name="getlastresult"></a>GetLastResult

Returns the last result code
```
FUNCTION GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
```

# <a name="setresult"></a>SetResult

Sets the last result code.
```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *Result* | The **HRESULT** error code returned by the methods. |

---

# <a name="geterrorinfo"></a>GetErrorInfo

Returns a description of the last result code.
```
PRIVATE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS CWSTR
```

---

# <a name="newdoc"></a>NewDoc

Opens a new document. If another document is open, this method saves any current changes and closes the current document before opening a new one.
```
FUNCTION NewDoc () AS HRESULT
```
#### Return value

If this method succeeds, it returns **S_OK**.

#### Usage example
```
pRichEdit.NewDoc
```

---

# <a name="opendoc"></a>OpenDoc

Opens a new document. If another document is open, this method saves any current changes and closes the current document before opening a new one.
```
FUNCTION OpenDoc (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG = 0, _
   BYVAL CodePage AS LONG = CP_ACP) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pVar* | A VARIANT that specifies the name of the file to open. |
| *Flags* | **tomRTF**: Open as RTF. **tomText**: Open as text ANSI or Unicode. |
| *CodePage* | The code page to use for the file. Zero (the default value) means **CP_ACP** (ANSI code page) unless the file begins with a Unicode BOM &hfeff, in which case the file is considered to be Unicode. Note that code page 1200 is Unicode, **CP_UTF8** is UTF-8. |

| Flag | Description |
| ---- | ----------- |
| *tomOpenExisting* | Open an existing file. Fail if the file does not exist. |
| *tomOpenAlways* | Open an existing file. Create a new file if the file does not exist. |
| *tomTruncateExisting* | Open an existing file, but truncate it to zero length. |

#### Return value

The return value can be an **HRESULT** value that corresponds to a system error or COM error code, including one of the following values.

| Result code | Description |
| ----------- | ----------- |
| **S_OK** | Method succeeds. |
| **E_INVALIDARG** | Invalid argument. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **E_NOTIMPL** | Feature not implemented. |

#### Usage example

```
DIM cv AS CVAR = AfxGetExePath & $"\Test.rtf"
pRichEdit.OpenDoc(cv)
```
# <a name="savedoc"></a>SaveDoc

Saves the document.

```
FUNCTION SaveDoc (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG = tomRTF, _
   BYVAL CodePage AS LONG = CP_ACP) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *pVar* | The save target. This parameter is a **VARIANT**, which can be a file name, or **NULL**. See **Remarks**. |
| *Flags* | File creation, share, and conversion flags. For a list of possible values, see the tables below. |
| *CodePage* | The specified code page. Common values are **CP_ACP** (zero: system ANSI code page), 1200 (Unicode), and 1208 (UTF-8). |

Any combination of these values may be used.

| Flag | Description |
| ---- | ----------- |
| *tomReadOnly* | Read only. |
| *tomShareDenyRead* | Other programs cannot read. |
| *tomShareDenyWrite* | Other programs cannot write. |

These values are mutually exclusive.

| Flag | Description |
| ---- | ----------- |
| *tomCreateNew* | Create a new file. Fail if the file already exists. |
| *tomCreateAlways* | Create a new file. Destroy the existing file if it exists. |
| *tomRTF* | Save as RTF. |
| *tomText* | Save as text. |

#### Return value

The return value can be an **HRESULT** value that corresponds to a system error or COM error code, including one of the following values.

| Result code | Description |
| ----------- | ----------- |
| **S_OK** | Method succeeds. |
| **E_INVALIDARG** | Invalid argument. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **E_NOTIMPL** | Feature not implemented. |

#### Remarks

To use the parameters that were specified for opening the file, use zero values for the parameters.

If *pVar* is null or missing, the file name given by this document's name is used. If both of these are missing or null, the method fails.

If *pVar* specifies a file name, that name should replace the current **Name** property. Similarly, the *Flags* and *CodePage* arguments can overrule those supplied in the **OpenDoc** method and define the values to use for files created with the **NewDoc** method.

Unicode plain-text files should be saved with the Unicode byte-order mark (&hFEFF) as the first character. This character should be removed when the file is read in; that is, it is only used for import/export to identify the plain text as Unicode and to identify the byte order of that text. Microsoft Notepad adopted this convention, which is now recommended by the Unicode standard.

#### Usage examples

Overwrites the current document in RTF format:
```
pRichEdit->SaveDoc(BYVAL NULL)
-- or --
pRichEdit->SaveDoc(BYVAL NULL, tomRTF)
```
Overwrites the current document in text format:
```
pRichEdit->SaveDoc(BYVAL NULL, tomText)
```
Overwrites the current document in utf-8 format:
```
pRichEdit->SaveDoc(BYVAL NULL, tomText, CP_UTF8)
```
Overwrites the current document in unicode format:
```
pRichEdit->SaveDoc(BYVAL NULL, tomText, 1200)
```
Saves the current document in RTF format:
```
DIM cv AS CVAR = AfxGetExePath & $"\Test02.rtf"
pRichEdit->SaveDoc(cv, tomCreateAlways OR tomRTF)
```
Saves the current document in text format:
```
DIM cv AS CVAR = AfxGetExePath & $"\Test02.txt"
pRichEdit->SaveDoc(cv, tomCreateAlways OR tomText)
```
Saves the current document in utf-8 (with BOM):
```
DIM cv AS CVAR = AfxGetExePath & $"\Test02.txt"
pRichEdit->SaveDoc(cv, tomCreateAlways OR tomText, CP_UTF8)
```
Saves the current document in unicode:
```
DIM cv AS CVAR = AfxGetExePath & $"\Test02.txt"
pRichEdit->SaveDoc(cv, tomCreateAlways OR tomText, 1200)
```

---

# <a name="appenddocfile"></a>AppendDocFile

Appends the contents of the specified file. Besides .rtf files, this method also insert the content of ansi, unicode or utf8 files if you specify the appropriate code page.
```
FUNCTION AppendDocFile (BYREF wszFileName AS WSTRING, BYVAL CodePage AS DWORD = CP_ACP) AS DWORD
```
| Parameter | Description |
| --------- | ----------- |
| *wszFileName* | Path and name of the RTF file to append. |
| *CodePage* | The code page to use for the conversion. Besides the default (CP_ACP), you can use CP_ANSI (ansi), CP_UTF8 (utf8) or 1200 (unicode). |

#### Usage example

Append a RTF file.
```
IM wszFileName AS WSTRING * MAX_PATH = AfxGetExePath & $"\Test.rtf"
RichEdit->AppendDocFile(wszFileName)
```
Append a file with ansi contents.
```
IM wszFileName AS WSTRING * MAX_PATH = AfxGetExePath & $"\Test.txt"
RichEdit->AppendDocFile(wszFileName, 500, CP_ANSI)
-- or --
RichEdit->AppendDocFile(wszFileName, -1, CP_ANSI)
```
Append a file with utf8 contents.
```
IM wszFileName AS WSTRING * MAX_PATH = AfxGetExePath & $"\Test.txt"
RichEdit->AppendDocFile(wszFileName, 500, CP_UTF8)
-- or --
RichEdit->AppendDocFile(wszFileName, -1, CP_UTF8)
```
Append a file with unicode contents.
```
IM wszFileName AS WSTRING * MAX_PATH = AfxGetExePath & $"\Test.txt"
RichEdit->AppendDocFile(wszFileName, 500, 1200)
-- or --
RichEdit->AppendDocFile(wszFileName, -1, 1200)
```

---

# <a name="insertdocfile"></a>InsertDocFile

Inserts the contents of the specified file at the specified location. Besides .rtf files, this method also insert the content of ansi, unicode or utf8 files if you specify the appropriate code page.
```
FUNCTION InsertDocFile (BYREF wszFileName AS WSTRING, BYVAL dwPos AS DWORD, BYVAL CodePage AS DWORD = CP_ACP) AS DWORD
```
| Parameter | Description |
| --------- | ----------- |
| *wszFileName* | Path and name of the RTF file to insert. |
| *dwPos* | Location where to insert the RTF file. If dwPos = -1, the contents of the RTF file are inserted at the caret position. |
| *CodePage* | The code page to use for the conversion. Besides the default (CP_ACP), you can use CP_ANSI (ansi), CP_UTF8 (utf8) or 1200 (unicode). |

#### Usage examples

Insert a RTF file.
```
IM wszFileName AS WSTRING * MAX_PATH = AfxGetExePath & $"\Test.rtf"
RichEdit->InsertDocFile(wszFileName, 500)
-- or --
RichEdit->InsertDocFile(wszFileName, -1)
```
Insert a file with ansi contents.
```
IM wszFileName AS WSTRING * MAX_PATH = AfxGetExePath & $"\Test.txt"
RichEdit->InsertDocFile(wszFileName, 500, CP_ANSI)
-- or --
RichEdit->InsertDocFile(wszFileName, -1, CP_ANSI)
```
Insert a file with utf8 contents.
```
IM wszFileName AS WSTRING * MAX_PATH = AfxGetExePath & $"\Test.txt"
RichEdit->InsertDocFile(wszFileName, 500, CP_UTF8)
-- or --
RichEdit->InsertDocFile(wszFileName, -1, CP_UTF8)
```
Insert a file with unicode contents.
```
IM wszFileName AS WSTRING * MAX_PATH = AfxGetExePath & $"\Test.txt"
RichEdit->InsertDocFile(wszFileName, 500, 1200)
-- or --
RichEdit->InsertDocFile(wszFileName, -1, 1200)
```

---

# <a name="getrtf"></a>GetRtf

Retrieves RTF formatted text from a Rich Edit control.
```
FUNCTION GetRtf (BYVAL dwOptionsAndFlags AS DWORD = SF_RTF) AS STRING
```
| Parameter | Description |
| --------- | ----------- |
| *dwOptionsAndFlags* | Specifies the data format and replacement options. |

| Value | Meaning |
| ----- | ------- |
| **SF_RTF** | RTF. |
| **SF_RTFNOOBJS** | RTF with spaces in place of COM objects. |
| **SF_TEXT** | Text with spaces in place of COM objects. |
| **SF_TEXTIZED** | Text with a text representation of COM objects. |

The **SF_RTFNOOBJS** option is useful if an application stores COM objects itself, as RTF representation of COM objects is not very compact. The control word, \objattph, followed by a space denotes the object position.

In addition, you can specify the following flags.

| Value | Meaning |
| ----- | ------- |
| **SFF_PLAINRTF** | If specified, the rich edit control streams out only the keywords common to all languages, ignoring language-specific keywords. If not specified, the rich edit control streams out all keywords. You can combine this flag with the **SF_RTF** or **SF_RTFNOOBJS** flag. |
| **SFF_SELECTION** | If specified, the rich edit control streams out only the contents of the current selection. If not specified, the control streams out the entire contents. You can combine this flag with any of data format values. |
| **SF_UNICODE** | Microsoft Rich Edit 2.0 and later: Indicates Unicode text. You can combine this flag with the **SF_TEXT** flag. |
| **SF_USECODEPAGE** | Rich Edit 3.0 and later: Generates UTF-8 RTF and text using other code pages. The code page is set in the high word of *dwOptionsAndFlags*. For example, for UTF-8 RTF, set *dwOptionsAndFlags* to (CP_UTF8 << 16) OR SF_USECODEPAGE OR SF_RTF. |

#### Return value

Returns the retrieved text or a null string.

---

# <a name="loadrtffromfile"></a>LoadRtfFromFile

Loads the contents of a RTF file into a Rich Edit control.
Deprecated and removed. Use **OpenDoc** instead.
```
FUNCTION LoadRtfFromFile (BYREF wszFileName AS WSTRING) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | The name of the RTF file to load. |

#### Return value

If the operation succeeds, the return value is **TRUE**. If the operation fails, the return value is **FALSE**.

#### Remarks

First you have to create an instance of the `CRichEditOleCallback` class and pass a pointer to it to the rich edit control using the **SetOleCallback** method.
```
DIM pRichEditOleCallback AS CRichEditOleCallback = CRichEditOleCallback
pRichEdit.SetOleCallback(@pRichEditOleCallback)
```
Then load the .rtf file using the **LoadRtfFromFile** method.
```
pRichEdit.LoadRtfFromFile(<path and name of the RTF file>)
e.g. pRichEdit->LoadRtfFromFile(AfxGetExePath & "\Test.rtf")
```

---

# <a name="loadrtffromresource"></a>LoadRtfFromResource

Loads a RTF resource file into a Rich Edit control.
```
FUNCTION LoadRtfFromResource (BYREF wszResourceName AS WSTRING) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *wszResourceName* | The name of the resource to load. |

#### Return value

If the operation succeeds, the return value is **TRUE**. If the operation fails, the return value is **FALSE**.

#### Remarks

First you have to create an instance of the `CRichEditOleCallback` class and pass a pointer to it to the rich edit control using the **SetOleCallback** method.
```
DIM pRichEditOleCallback AS CRichEditOleCallback = CRichEditOleCallback
pRichEdit.SetOleCallback(@pRichEditOleCallback)
```
Then load the .rtf file using the **LoadRtfFromResource** method.
```
pRichEdit.LoadRtfFromResource(<resource name>, e.g. RTFILE)
```
**Resource file**
1 24 ".\Resources\Manifest.xml"
RTFFILE       RCDATA ".\Resources\<title of the rtf file>.rtf"

---

# <a name="setfont"></a>SetFont

Sets the font used by a rich edit control.
```
FUNCTION SetFont (BYREF wszFaceName AS WSTRING, BYVAL ptsize AS LONG) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *wszFaceName* | The font name. |
| *ptsize* | The font size in points. |

#### Return value

If the operation succeeds, the return value is a nonzero value. If the operation fails, the return value is zero.

---

# <a name="addlfcr"></a>AddLF/AddCR/AddCRLF

Inserts a line feed, a carriage return or a carriage return and line feed at the cursor position or at the end of the text.
```
SUB AddLF (BYVAL atEnd AS BOOLEAN = FALSE)
SUB AddCR (BYVAL atEnd AS BOOLEAN = FALSE)
SUB AddCRLF (BYVAL atEnd AS BOOLEAN = FALSE)
```
| Parameter  | Description |
| ---------- | ----------- |
| *atEnd* | Optional. Boolean true (-1) or false (0). If true, the character is added at the end of the text; if false, it is added at the cursor position. |

---

# <a name="savertf"></a>SaveRtf

Saves the contents of the rich edit control to a file in rtf format.
```
FUNCTION SaveRtf (BYREF wszFilename AS WSTRING, BYVAL Overwrite AS BOOLEAN = FALSE, _
BYVAL dwFlagsAndAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFilename* | Path and file name of the file where the RTF content will be saved. |
| *Overwrite* | If true, the file will be overwritten if it already exists and the function will return true; if false, the file will not be overwritten, the function will return false and **GetLastResult" will return the **ERROR_ALREADY_EXISTS** error. |
| *dwFlagsAndAttributes* | The file or device attributes and flags. Default is **FILE_ATTRIBUTE_NORMAL**, which is the most common default value for files. To see more options see [CreateFileW](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew) |

#### Return value

Boolean true on success or false on failure.

---

# <a name="savertfnoobjs"></a>SaveRtfNoObjs

Saves the contents of the rich edit control to a file in rtf format with spaces in place of COM objects.

For parameters and return values see the **SaveRtf** method.

---

# <a name="saveselrtf"></a>SaveSelRtf

Saves selection of the rich edit control to a file in rtf format.

For parameters and return values see the **SaveRtf** method.

---

# <a name="savetext"></a>SaveText

Saves the contents of the rich edit control in text format.

For parameters and return values see the **SaveRtf** method.

---

# <a name="saveseltext"></a>SaveSelText

Saves selection of the rich edit control in text format.

For parameters and return values see the **SaveRtf** method.

---

# <a name="istextbold"></a>IsTextBold

Checks if the selected text or the word under the cursor is bolded.
```
FUNCTION IsTextBold () AS BOOLEAN
```
#### Return value

Boolean true (-1) or false (0).

---

# <a name="istextitalic"></a>IsTextItalic

Checks if the selected text or the word under the cursor is italicised.
```
FUNCTION IsTextItalic () AS BOOLEAN
```
#### Return value

Boolean true (-1) or false (0).

---

# <a name="istextstrikeout"></a>IsTextStrikeOut

Checks if the selected text or the word under the cursor is striked out.
```
FUNCTION IsTextStrikeOut () AS BOOLEAN
```
#### Return value

Boolean true (-1) or false (0).

---

# <a name="istextuderline"></a>IsTextUnderline

Checks if the selected text or the word under the cursor is underlined.
```
FUNCTION IsTextUnderline () AS BOOLEAN
```
#### Return value

Boolean true (-1) or false (0).

---

# <a name="settextbold"></a>SetTextBold

Changes the selected text or word under the cursor to bold. If it is already set, it removes it.
```
SUB SetTextBold ()
```

---

# <a name="settextitalic"></a>SetTextItalic

Changes the selected text or word under the cursor to italic. If it is already set, it removes it.
```
SUB SetTextItalic ()
```

---

# <a name="settextstrikeout"></a>SetTextStrikeOut

Changes the selected text or word under the cursor to strike out. If it is already set, it removes it.
```
SUB SetTextStrikeOut ()
```

---

# <a name="settextunderline"></a>SetTextUnderline

Changes the selected text or word under the cursor to underline. If it is already set, it removes it.
```
SUB SetTextUderline ()
```

---

# <a name="textcolor"></a>TextColor

Gets/sets the color of the selected text or the word under the cursor.

```
(GET) PROPERTY TextColor () AS COLORREF
(SET) PROPERTY TextColor (BYVAL crTxtColor AS COLORREF)
```
```
FUNCTION GetTextColor () AS COLORREF
FUNCTION SetTextColor (BYVAL crTxtColor AS COLORREF) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *crTxtColor* | The new text color. Use the macro RGB (BGR in FreeBasic), e.g. BGR(255,0,0). |

#### Return value

(GET) The color of the selected text or the word under the cursor if there is not selection.

(SET) A boolean true (-1) or false (0).

#### Remarks

To select text programatically, use the **SetSel** method.

---

# <a name="textfontname"></a>TextFontName

Gets/sets the font face name of the selected text or the word under the cursor if there is not selection.

```
(GET) PROPERTY TextFontName () AS CWSTR
(SET) PROPERTY TextFontName (BYREF wszFaceName AS WSTRING)
```
```
FUNCTION GetTextFontName () AS CWSTR
FUNCTION SetTextFontName (BYREF wszFaceName AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFaceName* | (SET) The new font face name, e.g. "Times New Roman". |

#### Return value

(GET) The font face name.

(SET) A boolean true (-1) or false (0).

#### Remarks

To select text programatically, use the **SetSel** method.

---

# <a name="textheight"></a>TextHeight

Gets/sets the height of the selected text or the word under the cursor.

```
(GET) PROPERTY TextHeight () AS LONG
(SET) PROPERTY TextHeight (BYVAL ptSize AS LONG)
```
```
FUNCTION GetTextHeight () AS LONG
FUNCTION SetTextHeight (BYVAL ptSize AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ptSize* | The new text height in points. |

#### Return value

(GET) The height of the selected text or the word under the cursor if there is not selection.

#### Remarks

To select text programatically, use the **SetSel** method.

---

# <a name="textoffset"></a>TextOffset

Gets/sets the offset of the selected text or the word under the cursor.

```
(GET) PROPERTY TextOffset () AS LONG
(SET) PROPERTY TextOffset (BYVAL offset AS LONG)
```
```
FUNCTION GetTextOffset () AS LONG
FUNCTION SetTextOffset (BYVAL offset AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *offset* | Character offset, in pixels, from the baseline. If the value of this member is positive, the character is a superscript; if it is negative, the character is a subscript. |

#### Return value

(GET) The offset of the selected text or the word under the cursor if there is not selection.

(SET) A boollean true (-1) or false (0).

#### Remarks

This property converts the pixels to twips internally.

To select text programatically, use the **SetSel** method.

---
