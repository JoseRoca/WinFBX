# RichEdit Procedures

| Name       | Description |
| ---------- | ----------- |
| [RichEdit_AutoUrlDetect](#RichEdit_AutoUrlDetect) | Enables or disables automatic detection of URLs by a rich edit control. |
| [RichEdit_CanPaste](#RichEdit_CanPaste) | Determines whether a rich edit control can paste a specified clipboard format. |
| [RichEdit_CanRedo](#RichEdit_CanRedo) | Determines whether there are any actions in the rich edit control redo queue. |
| [RichEdit_CanUndo](#RichEdit_CanUndo) | Determines whether there are any actions in the rich edit control undo queue. |
| [RichEdit_CharFromPos](#RichEdit_CharFromPos) | Gets information about the character closest to a specified point in the client area of a rich edit control. |
| [RichEdit_DisplayBand](#RichEdit_DisplayBand) | Displays a portion of the contents of a rich edit control, as previously formatted for a device using the EM_FORMATRANGE message. |
| [RichEdit_EmptyUndoBuffer](#RichEdit_EmptyUndoBuffer) | Resets the undo flag of a rich edit control. The undo flag is set whenever an operation within the rich edit control can be undone. |
| [RichEdit_ExGetSel](#RichEdit_ExGetSel) | Retrieves the starting and ending character positions of the selection in a rich edit control. |
| [RichEdit_ExLimitText](#RichEdit_ExLimitText) | Sets an upper limit to the amount of text the user can type or paste into a rich edit control. |
| [RichEdit_ExLineFromChar](#RichEdit_ExLineFromChar) | Determines which line contains the specified character in a rich edit control. |
| [RichEdit_ExSetSel](#RichEdit_ExSetSel) | Selects a range of characters and/or Component Object Model (COM) objects in a Microsoft Rich Edit control. |
| [RichEdit_FindText](#RichEdit_FindText) | Finds text within a rich edit control. |
| [RichEdit_FindTextEx](#RichEdit_FindTextEx) | Finds text within a rich edit control. |
| [RichEdit_FindWordBreak](#RichEdit_FindWordBreak) | Finds the next word break before or after the specified character position or retrieves information about the character at that position. |
| [RichEdit_FormatRange](#RichEdit_FormatRange) | Formats a range of text in a rich edit control for a specific device. |
| [RichEdit_GetAutoUrlDetect](#RichEdit_GetAutoUrlDetect) | Indicates whether the auto URL detection is turned on in the rich edit control. |
| [RichEdit_GetBidiOptions](#RichEdit_GetBidiOptions) | Indicates the current state of the bidirectional options in the rich edit control. |
| [RichEdit_GetCharFormat](#RichEdit_GetCharFormat) | Determines the current character formatting in a rich edit control. |
| [RichEdit_GetCTFModeBias](#RichEdit_GetCTFModeBias) | Gets the Text Services Framework mode bias values for a Microsoft Rich Edit control. |
| [RichEdit_GetCTFOpenStatus](#RichEdit_GetCTFOpenStatus) | Determines if the Text Services Framework (TSF) keyboard is open or closed. |
| [RichEdit_GetEditStyle](#RichEdit_GetEditStyle) | Retrieves the current edit style flags. |
| [RichEdit_GetEventMask](#RichEdit_GetEventMask) | Retrieves the event mask for a rich edit control. The event mask specifies which notification messages the control sends to its parent window. |
| [RichEdit_GetFirstVisibleLine](#RichEdit_GetFirstVisibleLine) | Gets the zero-based index of the uppermost visible line in a multiline rich edit control. |
| [RichEdit_GetHyphenateInfo](#RichEdit_GetHyphenateInfo) | Gets information about hyphenation for a Microsoft Rich Edit control. |
| [RichEdit_GetIMEColor](#RichEdit_GetIMEColor) | Retrieves the Input Method Editor (IME) composition color. This message is available only in Asian-language versions of the operating system. |
| [RichEdit_GetIMECompMode](#RichEdit_GetIMECompMode) | Gets the current IME mode for a rich edit control. |
| [RichEdit_GetIMECompText](#RichEdit_GetIMECompText) | Gets the Input Method Editor (IME) composition text. |
| [RichEdit_GetIMEModeBias](#RichEdit_GetIMEModeBias) | Gets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control. |
| [RichEdit_GetIMEOptions](#RichEdit_GetIMEOptions) | Retrieves the current Input Method Editor (IME) options. This message is available only in Asian-language versions of the operating system. |
| [RichEdit_GetIMEProperty](#RichEdit_GetIMEProperty) | Gets the property and capabilities of the Input Method Editor (IME) associated with the current input locale. |
| [RichEdit_GetLangOptions](#RichEdit_GetLangOptions) | Gets a rich edit control's option settings for Input Method Editor (IME) and Asian language support. |
| [RichEdit_GetLimitText](#RichEdit_GetLimitText) | Gets the current text limit for a rich edit control. |
| [RichEdit_GetLine](#RichEdit_GetLine) | Copies a line of text from a rich edit control. |
| [RichEdit_GetLineCount](#RichEdit_GetLineCount) | Gets the number of lines in a multiline rich edit control. |
| [RichEdit_GetModify](#RichEdit_GetModify) | Gets the state of a rich edit control's modification flag. The flag indicates whether the contents of the rich edit control have been modified. |
| [RichEdit_GetOleInterface](#RichEdit_GetOleInterface) | Retrieves an IRichEditOle object that a client can use to access a rich edit control's Component Object Model (COM) functionality. |
| [RichEdit_GetOptions](#RichEdit_GetOptions) | Retrieves rich edit control options. |
| [RichEdit_GetPageRotate](#RichEdit_GetPageRotate) | Deprecated. Gets the text layout for a Microsoft Rich Edit control. |
| [RichEdit_GetParaFormat](#RichEdit_GetParaFormat) | Retrieves the paragraph formatting of the current selection in a rich edit control. |
| [RichEdit_GetPasswordChar](#RichEdit_GetPasswordChar) | Gets the password character that a rich edit control displays when the user enters text. |
| [RichEdit_GetPunctuation](#RichEdit_GetPunctuation) | Gets the current punctuation characters for the rich edit control. |
| [RichEdit_GetRect](#RichEdit_GetRect) | Gets the formatting rectangle of a rich edit control. |
| [RichEdit_GetRedoName](#RichEdit_GetRedoName) | Retrieves the type of the next action, if any, in the control's redo queue. |
| [RichEdit_GetScrollPos](#RichEdit_GetScrollPos) | Obtains the current scroll position of the edit control. |
| [RichEdit_GetSel](#RichEdit_GetSel) | Gets the starting and ending character positions of the current selection in a rich edit control. |
| [RichEdit_GetSelText](#RichEdit_GetSelText) | Retrieves the currently selected text in a rich edit control. |
| [RichEdit_GetText](#RichEdit_GetText) | Retrieves the text from a rich edit control. |
| [RichEdit_GetTextEx](#RichEdit_GetTextEx) | Gets all of the text from the rich edit control in any particular code base you want. |
| [RichEdit_GetTextLength](#RichEdit_GetTextLength) | Retrieves the length of all text in a rich edit control. |
| [RichEdit_GetTextLengthEx](#RichEdit_GetTextLengthEx) | Calculates text length in various ways. It is usually called before creating a buffer to receive the text from the control. |
| [RichEdit_GetTextMode](#RichEdit_GetTextMode) | Gets the current text mode and undo level of a rich edit control. |
| [RichEdit_GetTextRange](#RichEdit_GetTextRange) | Retrieves a specified range of characters from a rich edit control. |
| [RichEdit_GetThumb](#RichEdit_GetThumb) | Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline rich edit control. |
| [RichEdit_GetTypographyOptions](#RichEdit_GetTypographyOptions) | Returns the current state of the typography options of a rich edit control. |
| [RichEdit_GetUndoName](#RichEdit_GetUndoName) | Retrieves the type of the next undo action, if any. |
| [RichEdit_GetWordBreakProc](#RichEdit_GetWordBreakProc) | Gets the address of the current Wordwrap function. |
| [RichEdit_GetWordBreakProcEx](#RichEdit_GetWordBreakProcEx) | Retrieves the address of the currently registered extended word-break procedure. |
| [RichEdit_GetWordWrapMode](#RichEdit_GetWordWrapMode) | Gets the current word wrap and word-break options for the rich edit control. |
| [RichEdit_GetZoom](#RichEdit_GetZoom) | Gets the current zoom ratio, which is always between 1/64 and 64. |
| [RichEdit_HideSelection](#RichEdit_HideSelection) | Hides or shows the selection in a rich edit control. |
| [RichEdit_IsIME](#RichEdit_IsIME) | Determines if current input locale is an East Asian locale. |
| [RichEdit_LimitText](#RichEdit_LimitText) | Sets the text limit of a rich edit control. The text limit is the maximum amount of text, in TCHARs, that the user can type into the edit control. |
| [RichEdit_LineFromChar](#RichEdit_LineFromChar) | Gets the index of the line that contains the specified character index in a multiline rich edit control. |
| [RichEdit_LineIndex](#RichEdit_LineIndex) | Gets the character index of the first character of a specified line in a multiline rich edit control. |
| [RichEdit_LineLength](#RichEdit_LineLength) | Retrieves the length, in characters, of a line in a rich edit control. |
| [RichEdit_LineScroll](#RichEdit_LineScroll) | Scrolls the text in a multiline rich edit control. |
| [RichEdit_PasteSpecial](#RichEdit_PasteSpecial) | Pastes a specific clipboard format in a rich edit control. |
| [RichEdit_PosFromChar](#RichEdit_PosFromChar) | Retrieves the client area coordinates of a specified character in a rich edit control. |
| [RichEdit_Reconversion](#RichEdit_Reconversion) | Invokes the Input Method Editor (IME) reconversion dialog box. |
| [RichEdit_Redo](#RichEdit_Redo) | Redoes the next action in the control's redo queue. |
| [RichEdit_ReplaceSel](#RichEdit_ReplaceSel) | Replaces the current selection in a rich edit control with the specified text. |
| [RichEdit_RequestResize](#RichEdit_RequestResize) | Forces a rich edit control to send an EN_REQUESTRESIZE notification message to its parent window. |
| [RichEdit_Scroll](#RichEdit_Scroll) | Scrolls the text vertically in a multiline rich edit control. |
| [RichEdit_ScrollCaret](#RichEdit_ScrollCaret) | Scrolls the caret into view in a rich edit control. |
| [RichEdit_SelectionType](#RichEdit_SelectionType) | Determines the selection type for a rich edit control. |
| [RichEdit_SetBidiOptions](#RichEdit_SetBidiOptions) | Sets the current state of the bidirectional options in the rich edit control. |
| [RichEdit_SetBkgndColor](#RichEdit_SetBkgndColor) | Sets the background color for a rich edit control. |
| [RichEdit_SetCharFormat](#RichEdit_SetCharFormat) | Sets character formatting in a rich edit control. |
| [RichEdit_SetCTFModeBias](#RichEdit_SetCTFModeBias) | Sets the Text Services Framework (TSF) mode bias for a Microsoft Rich Edit control. |
| [RichEdit_SetCTFModeBias](#RichEdit_SetCTFModeBias) | Sets the Text Services Framework (TSF) mode bias for a Microsoft Rich Edit control. |
| [RichEdit_SetCTFOpenStatus](#RichEdit_SetCTFOpenStatus) | Opens or closes the Text Services Framework (TSF) keyboard. |
| [RichEdit_SetEditStyle](#RichEdit_SetEditStyle) | Sets the current edit style flags. |
| [RichEdit_SetEventMask](#RichEdit_SetEventMask) | Sets the event mask for a rich edit control. |
| [RichEdit_SetFontSize](#RichEdit_SetFontSize) | Sets the font size for the selected text. |
| [RichEdit_SetHyphenateInfo](#RichEdit_SetHyphenateInfo) | Sets the way a Microsoft Rich Edit control does hyphenation. |
| [RichEdit_SetIMEColor](#RichEdit_SetIMEColor) | Sets the Input Method Editor (IME) composition color. |
| [RichEdit_SetIMEModeBias](#RichEdit_SetIMEModeBias) | Sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control. |
| [RichEdit_SetIMEOptions](#RichEdit_SetIMEOptions) | Sets the Input Method Editor (IME) options. |
| [RichEdit_SetLangOptions](#RichEdit_SetLangOptions) | Sets options for Input Method Editor (IME) and Asian language support in a rich edit control. |
| [RichEdit_SetLimitText](#RichEdit_SetLimitText) | Sets the text limit of a rich edit control. The text limit is the maximum amount of text, in TCHARs, that the user can type into the edit control. |
| [RichEdit_SetMargins](#RichEdit_SetMargins) | Sets the widths of the left and right margins for a rich edit control. The message redraws the control to reflect the new margins. |
| [RichEdit_SetModify](#RichEdit_SetModify) | Sets or clears the modification flag for a rich edit control. The modification flag indicates whether the text within the rich edit control has been modified. |
| [RichEdit_SetOleCallback](#RichEdit_SetOleCallback) | Gives a rich edit control an IRichEditOleCallback object that the control uses to get OLE-related resources and information from the client. |
| [RichEdit_SetOptions](#RichEdit_SetOptions) | Sets the options for a rich edit control. |
| [RichEdit_SetPageRotate](#RichEdit_SetPageRotate) | Deprecated. Sets the text layout for a Microsoft Rich Edit control. |
| [RichEdit_SetPalette](#RichEdit_SetPalette) | Deprecated. Sets the text layout for a Microsoft Rich Edit control. |
| [RichEdit_SetParaFormat](#RichEdit_SetParaFormat) | Sets the paragraph formatting for the current selection in a rich edit control. |
| [RichEdit_SetPasswordChar](#RichEdit_SetPasswordChar) | Sets or removes the password character for a rich edit control. When a password character is set, that character is displayed in place of the characters typed by the user. |
| [RichEdit_SetPunctuation](#RichEdit_SetPunctuation) | Sets the punctuation characters for a rich edit control. |
| [RichEdit_SetReadOnly](#RichEdit_SetReadOnly) | Sets or removes the read-only style (ES_READONLY) of a rich edit control. |
| [RichEdit_SetRect](#RichEdit_SetRect) | Sets the formatting rectangle of a multiline rich edit control. |
| [RichEdit_SetRectNP](#RichEdit_SetRectNP) | Sets the formatting rectangle of a multiline rich edit control. |
| [RichEdit_SetScrollPos](#RichEdit_SetScrollPos) | Tells the rich edit control to scroll to a particular point. |
| [RichEdit_SetSel](#RichEdit_SetSel) | Selects a range of characters in a rich edit control. |
| [RichEdit_SetTabStops](#RichEdit_SetTabStops) | Sets the tab stops in a multiline rich edit control. |
| [RichEdit_SetTargetDevice](#RichEdit_SetTargetDevice) | Sets the target device and line width used for WYSIWYG formatting in a rich edit control. |
| [RichEdit_SetText](#RichEdit_SetText) | Sets the text of an edit control. |
| [RichEdit_SetTextExW](#RichEdit_SetTextExW) | Combines the functionality of WM_SETTEXT and EM_REPLACESEL and adds the ability to set text using a code page and to use either Rich Text Format (RTF) rich text or plain text. |
| [RichEdit_SetTextMode](#RichEdit_SetTextMode) | Sets the text mode or undo level of a rich edit control. |
| [RichEdit_SetTypographyOptions](#RichEdit_SetTypographyOptions) | Sets the text mode or undo level of a rich edit control. |
| [RichEdit_SetTypographyOptions](#RichEdit_SetTypographyOptions) | Sets the text mode or undo level of a rich edit control. |
| [RichEdit_SetUndoLimit](#RichEdit_SetUndoLimit) | Sets the maximum number of actions that can stored in the undo queue. |
| [RichEdit_SetWordBreakProc](#RichEdit_SetWordBreakProc) | Replaces a rich edit control's default Wordwrap function with an application-defined wordwrap function. |
| [RichEdit_SetWordBreakProcEx](#RichEdit_SetWordBreakProcEx) | Sets the extended word-break procedure. |
| [RichEdit_SetWordWrapMode](#RichEdit_SetWordWrapMode) | Sets the word-wrapping and word-breaking options for the rich edit control. |
| [RichEdit_SetZoom](#RichEdit_SetZoom) | Sets the zoom ratio anywhere between 1/64 and 64. |
| [RichEdit_ShowScrollBar](#RichEdit_ShowScrollBar) | Shows or hides one of the scroll bars in the Text Host window. |
| [RichEdit_StopGroupTyping](#RichEdit_StopGroupTyping) | Stops the control from collecting additional typing actions into the current undo action. |
| [RichEdit_StreamIn](#RichEdit_StreamIn) | Replaces the contents of a rich edit control with a stream of data provided by an application defined–EditStreamCallback callback function. |
| [RichEdit_StreamOut](#RichEdit_StreamOut) | Causes a rich edit control to pass its contents to an application–defined EditStreamCallback callback function. |
| [RichEdit_Undo](#RichEdit_Undo) | This message undoes the last edit control operation in the control's undo queue. |

# RichEdit Helper Procedures

| Name       | Description |
| ---------- | ----------- |
| [RichEdit_GetRtfText](#RichEdit_GetRtfText) | Retrieves formatted text from a Rich Edit control |
| [RichEdit_LoadRtfFromFileW](#RichEdit_LoadRtfFromFileW) | Loads a Rich Text File into a Rich Edit control. |
| [RichEdit_LoadRtfFromResourceW](#RichEdit_LoadRtfFromResourceW) | Loads a Rich Text Resource File into a Rich Edit control. |
| [RichEdit_SetFontW](#RichEdit_SetFontW) | Sets the font used by a rich edit control. |

# <a name="RichEdit_AutoUrlDetect"></a>RichEdit_AutoUrlDetect

Enables or disables automatic detection of URLs by a rich edit control.

```
FUNCTION RichEdit_AutoUrlDetect (BYVAL hRichEdit AS HWND, BYVAL fUrlDetect AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_AUTOURLDETECT, fUrlDetect, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *fUrlDetect* | Specify 0 to disable automatic link detection, or one of the following values to enable various kinds of detection. |

| fUrlDetect value  | Description |
| --------------- | ----------- |
| AURL_DISABLEMIXEDLGC | **Windows 8 or later**: Disable recognition of domain names that contain labels with characters belonging to more than one of the following scripts: Latin, Greek, and Cyrillic. |
| AURL_ENABLEDRIVELETTERS | **Windows 8 or later**: Recognize file names that have a leading drive specification, such as c:\temp. |
| AURL_ENABLEEA | This value is deprecated; use **AURL_ENABLEEAURLS** instead. |
| AURL_ENABLEEAURLS | Recognize URLs that contain East Asian characters. |
| AURL_ENABLEEMAILADDR | **Windows 8 or later**: Recognize email addresses. |
| AURL_ENABLETELNO | **Windows 8 or later**: Recognize telephone numbers. |
| AURL_ENABLEURL | **Windows 8 or later**: Recognize URLs that include the path. |

#### Return value

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

#### Return value

If the message succeeds, the return value is zero.

If the message fails, the return value is a nonzero value. For example, the message might fail due to insufficient memory or an invalid detection option.

# <a name="RichEdit_CanPaste"></a>RichEdit_CanPaste

Determines whether a rich edit control can paste a specified clipboard format.

```
FUNCTION RichEdit_CanPaste (BYVAL hRichEdit AS HWND, BYVAL clipformat AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_CANPASTE, clipformat, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *clipformat* | Specifies the [Clipboard Formats](https://learn.microsoft.com/en-us/windows/win32/dataxchg/clipboard-formats) to try. To try any format currently on the clipboard, set this parameter to zero. |

#### Return value

If the clipboard format can be pasted, the return value is a nonzero value.

If the clipboard format cannot be pasted, the return value is zero.

# <a name="RichEdit_CanRedo"></a>RichEdit_CanRedo

Determines whether there are any actions in the rich control redo queue.

```
FUNCTION RichEdit_CanRedo (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_CANREDO, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

If there are actions in the control redo queue, the return value is a nonzero value.

If the redo queue is empty, the return value is zero.

# <a name="RichEdit_CanUndo"></a>RichEdit_CanUndo

Determines whether there are any actions in the rich edit control undo queue.

```
FUNCTION RichEdit_CanUndo (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_CANUNDO, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

If there are actions in the control's undo queue, the return value is nonzero.

If the undo queue is empty, the return value is zero.

# <a name="RichEdit_CharFromPos"></a>RichEdit_CharFromPos

Gets information about the character closest to a specified point in the client area of a rich edit control.

```
FUNCTION RichEdit_CharFromPos (BYVAL hRichEdit AS HWND, BYVAL lppl AS POINTL PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_CHARFROMPOS, 0, cast(LPARAM, lppl))
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lppl* | A pointer to a [POINTL](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-pointl) structure that contains the horizontal and vertical coordinates of a point in the control's client area. The coordinates are in screen units and are relative to the upper-left corner of the control's client area. |

#### Return value

The return value specifies the zero-based character index of the character nearest the specified point. The return value indicates the last character in the edit control if the specified point is beyond the last character in the control.

# <a name="RichEdit_DisplayBand"></a>RichEdit_DisplayBand

Displays a portion of the contents of a rich edit control, as previously formatted for a device using the [EM_FORMATRANGE]([url](https://learn.microsoft.com/en-us/windows/win32/controls/em-formatrange)) message.

```
FUNCTION RichEdit_DisplayBand (BYVAL hRichEdit AS HWND, BYVAL lprc AS RECT PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_DISPLAYBAND, 0, cast(LPARAM, lprc))
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lprc* | A pointer to a [RECT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-rect) structure specifying the display area of the device. |

#### Return value

If the operation succeeds, the return value is TRUE.

If the operation fails, the return value is FALSE.

#### Remarks

Text and Component Object Model (COM) objects are clipped by the rectangle. The application does not need to set the clipping region.

Banding is the process by which a single page of output is generated using one or more separate rectangles, or bands. When all bands are placed on the page, a complete image results. This approach is often used by raster printers that do not have sufficient memory or ability to image a full page at one time. Banding devices include most dot matrix printers as well as some laser printers.

# <a name="RichEdit_EmptyUndoBuffer"></a>RichEdit_EmptyUndoBuffer

Resets the undo flag of an edit control. The undo flag is set whenever an operation within the edit control can be undone.

```
SUB RichEdit_EmptyUndoBuffer (BYVAL hRichEdit AS HWND)
   SendMessageW hRichEdit, EM_EMPTYUNDOBUFFER, 0, 0
END SUB
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

# <a name="RichEdit_ExGetSel"></a>RichEdit_ExGetSel

Retrieves the starting and ending character positions of the selection in a rich edit control.

```
SUB RichEdit_ExGetSel (BYVAL hRichEdit AS HWND, BYVAL lpchr AS CHARRANGE PTR)
   SendMessageW hRichEdit, EM_EXGETSEL, 0, cast(LPARAM, lpchr)
END SUB
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lpchr* | A pointer to a [CHARRANGE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charrange) structure that receives the selection range. |

# <a name="RichEdit_ExLimitText"></a>RichEdit_ExLimitText

Retrieves the starting and ending character positions of the selection in a rich edit control.

```
SUB RichEdit_ExLimitText (BYVAL hRichEdit AS HWND, BYVAL dwLimit AS DWORD)
   SendMessageW hRichEdit, EM_EXLIMITTEXT, 0, dwLimit
END SUB
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *dwLimit* | Specifies the maximum amount of text that can be entered. If this parameter is zero, the default maximum is used, which is 64K characters. A COM object counts as a single character. |

#### Remarks

The text limit set by the EM_EXLIMITTEXT message does not limit the amount of text that you can stream into a rich edit control using the EM_STREAMIN message with lParam set to SF_TEXT. However, it does limit the amount of text that you can stream into a rich edit control using the EM_STREAMIN message with lParam set to SF_RTF.

Before EM_EXLIMITTEXT is called, the default limit to the amount of text a user can enter is 32,767 characters.

# <a name="RichEdit_ExLineFromChar"></a>RichEdit_ExLineFromChar

Determines which line contains the specified character in a rich edit control.

```
FUNCTION RichEdit_ExLineFromChar (BYVAL hRichEdit AS HWND, BYVAL iIndex AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_EXLINEFROMCHAR, 0, iIndex)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *iIndex* | Zero-based index of the character. |

#### Return value

Returns the zero-based index of the line.

# <a name="RichEdit_ExSetSel"></a>RichEdit_ExSetSel

Selects a range of characters or Component Object Model (COM) objects in a Microsoft Rich Edit control.

```
FUNCTION RichEdit_ExSetSel (BYVAL hRichEdit AS HWND, BYVAL lpcr AS CHARRANGE PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_EXSETSEL, 0, cast(LPARAM, lpcr))
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lpcr* | A pointer to a [CHARRANGE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charrange) structure that specifies the selection range. |

#### Return value

The return value is the selection that is actually set.

# <a name="RichEdit_FindText"></a>RichEdit_FindText

Finds Unicode text within a rich edit control.

```
FUNCTION RichEdit_FindText (BYVAL hRichEdit AS HWND, BYVAL fOptions AS DWORD, BYVAL lpft AS FINDTEXTW PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_FINDTEXTW, fOptions, cast(LPARAM, lpft))
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *fOptions* | Specifies the parameters of the search operation. This parameter can be one or more of the following values.<br>**FR_DOWN**. If set, the operation searches from the end of the current selection to the end of the document. If not set, the operation searches from the end of the current selection to the beginning of the document.<br>**FR_MATCHALEFHAMZA**. By default, Arabic and Hebrew alefs with different accents are all matched by the alef character. Set this flag if you want the search to differentiate between alefs with different accents.<br>**FR_MATCHCASE**. If set, the search operation is case-sensitive. If not set, the search operation is case-insensitive.<br>**FR_MATCHDIAC**. By default, Arabic and Hebrew diacritical marks are ignored. Set this flag if you want the search operation to consider diacritical marks.<br>**FR_MATCHKASHIDA**. By default, Arabic and Hebrew kashidas are ignored. Set this flag if you want the search operation to consider kashidas.<br>**FR_WHOLEWORD**. If set, the operation searches only for whole words that match the search string. If not set, the operation also searches for word fragments that match the search string.|
| *lpft* | A pointer to a [FINDTEXTW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-findtextw) structure containing information about the find operation. |

#### Return value

If the target string is found, the return value is the zero-based position of the first character of the match. If the target is not found, the return value is -1.

# <a name="RichEdit_FindTextEx"></a>RichEdit_FindTextEx

Finds Unicode text within a rich edit control.

```
FUNCTION RichEdit_FindTextEx (BYVAL hRichEdit AS HWND, BYVAL fOptions AS DWORD, BYVAL lpftexw AS FINDTEXTEXW PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_FINDTEXTEXW, fOptions, cast(LPARAM, lpftexw))
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *fOptions* | Specifies the parameters of the search operation. This parameter can be one or more of the following values.<br>**FR_DOWN**. If set, the operation searches from the end of the current selection to the end of the document. If not set, the operation searches from the end of the current selection to the beginning of the document.<br>**FR_MATCHALEFHAMZA**. By default, Arabic and Hebrew alefs with different accents are all matched by the alef character. Set this flag if you want the search to differentiate between alefs with different accents.<br>**FR_MATCHCASE**. If set, the search operation is case-sensitive. If not set, the search operation is case-insensitive.<br>**FR_MATCHDIAC**. By default, Arabic and Hebrew diacritical marks are ignored. Set this flag if you want the search operation to consider diacritical marks.<br>**FR_MATCHKASHIDA**. By default, Arabic and Hebrew kashidas are ignored. Set this flag if you want the search operation to consider kashidas.<br>**FR_WHOLEWORD**. If set, the operation searches only for whole words that match the search string. If not set, the operation also searches for word fragments that match the search string.|
| *lpftexw* | A pointer to a [FINDTEXTEXW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-findtextexw) structure containing information about the find operation. |

#### Return value

If the target string is found, the return value is the zero-based position of the first character of the match. If the target is not found, the return value is -1.

#### Remarks

**RichEdit_FindTextEx** uses the [FINDTEXTEXW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-findtextexw) structure, while **RichEdit_FindText** uses the [FINDTEXTW](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-findtextw) structure. The difference is that EM_FINDTEXTEXW reports the range of text that was found.

# <a name="RichEdit_FindWordBreak"></a>RichEdit_FindWordBreak

Finds the next word break before or after the specified character position or retrieves information about the character at that position.

```
FUNCTION RichEdit_FindWordBreak (BYVAL hRichEdit AS HWND, BYVAL fOperation AS DWORD, BYVAL dwStartPos AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_FINDWORDBREAK, fOperation, dwStartPos)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *fOperation* | pecifies the find operation. This parameter can be one of the following values.<br>**WB_CLASSIFY**. Returns the character class and word-break flags of the character at the specified position.<br>**WB_ISDELIMITER**. Returns TRUE if the character at the specified position is a delimiter, or FALSE otherwise.<br>**WB_LEFT**. Finds the nearest character before the specified position that begins a word.<br>**WB_LEFTBREAK**. Finds the next word end before the specified position. This value is the same as WB_PREVBREAK.<br>**WB_MOVEWORDLEFT**. Finds the next character that begins a word before the specified position. This value is used during CTRL+LEFT ARROW key processing. This value is the similar to WB_MOVEWORDPREV. See Remarks for more information.<br>**WB_MOVEWORDRIGHT**. Finds the next character that begins a word after the specified position. This value is used during CTRL+right key processing. This value is similar to WB_MOVEWORDNEXT. See Remarks for more information.<br>**WB_RIGHT**. Finds the next character that begins a word after the specified position.<br>**WB_RIGHTBREAK**. Finds the next end-of-word delimiter after the specified position. This value is the same as WB_NEXTBREAK. |
| *dwStartPos* | Zero-based character starting position. |

#### Return value

The message returns a value based on the *fOperation* parameter.

| Return code  | Description |
| ------------ | ----------- |
| **WB_CLASSIFY** | Returns the character class and word-break flags of the character at the specified position. |
| **WB_ISDELIMITER** | Returns TRUE if the character at the specified position is a delimiter; otherwise it returns FALSE. |
| **Others** | Returns the character index of the word break. |

#### Remarks
If fOperation is WB_LEFT and WB_RIGHT, the word-break procedure finds word breaks only after delimiters. This matches the functionality of an edit control. If fOperation is WB_MOVEWORDLEFT or WB_MOVEWORDRIGHT, the word-break procedure also compares character classes and word-break flags.

For information about character classes and word-break flags, see [Word and Line Breaks](https://learn.microsoft.com/en-us/windows/win32/controls/use-word-and-line-break-information).

# <a name="RichEdit_FormatRange"></a>RichEdit_FormatRange

Formats a range of text in a rich edit control for a specific device.

```
FUNCTION RichEdit_FormatRange (BYVAL hRichEdit AS HWND, BYVAL fRender AS LONG, BYVAL lpfr AS FORMATRANGE PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_FORMATRANGE, fRender, cast(LPARAM, lpfr))
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *fRender* | Specifies whether to render the text. If this parameter is not zero, the text is rendered. Otherwise, the text is just measured. |
| *lpfr* | A pointer to a [FORMATRANGE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-formatrange) structure containing information about the output device, or **NULL** to free information cached by the control. |

#### Return value

This message returns the index of the last character that fits in the region, plus 1.

#### Remarks

This message is typically used to format the content of rich edit control for an output device such as a printer.

After using this message to format a range of text, it is important that you free cached information by sending **EM_FORMATRANGE** again, but with lParam set to **NULL**; otherwise, a memory leak will occur. Also, after using this message for one device, you must free cached information before using it again for a different device.
