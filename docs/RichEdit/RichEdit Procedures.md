# RichEdit Procedures

| Name       | Description |
| ---------- | ----------- |
| [RichEdit_AutoUrlDetect](#RichEdit_AutoUrlDetect) | Enables or disables automatic detection of URLs by a rich edit control. |
| [RichEdit_CanPaste](#RichEdit_CanPaste) | Determines whether a rich edit control can paste a specified clipboard format. |
| [RichEdit_CanRedo](#RichEdit_CanRedo) | Determines whether there are any actions in the rich edit control redo queue. |
| [RichEdit_CanUndo](#RichEdit_CanUndo) | Determines whether there are any actions in the rich edit control undo queue. |
| [RichEdit_CharFromPos](#RichEdit_CharFromPos) | Retrieves information about the character closest to a specified point in the client area of a rich edit control. |
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
| [RichEdit_GetCTFModeBias](#RichEdit_GetCTFModeBias) | Retrieves the Text Services Framework mode bias values for a Microsoft Rich Edit control. |
| [RichEdit_GetCTFOpenStatus](#RichEdit_GetCTFOpenStatus) | Determines if the Text Services Framework (TSF) keyboard is open or closed. |
| [RichEdit_GetEditStyle](#RichEdit_GetEditStyle) | Retrieves the current edit style flags. |
| [RichEdit_GetEventMask](#RichEdit_GetEventMask) | Retrieves the event mask for a rich edit control. The event mask specifies which notification messages the control sends to its parent window. |
| [RichEdit_GetFirstVisibleLine](#RichEdit_GetFirstVisibleLine) | Retrieves the zero-based index of the uppermost visible line in a multiline rich edit control. |
| [RichEdit_GetHyphenateInfo](#RichEdit_GetHyphenateInfo) | Retrieves information about hyphenation for a Microsoft Rich Edit control. |
| [RichEdit_GetIMEColor](#RichEdit_GetIMEColor) | Retrieves the Input Method Editor (IME) composition color. This message is available only in Asian-language versions of the operating system. |
| [RichEdit_GetIMECompMode](#RichEdit_GetIMECompMode) | Retrieves the current IME mode for a rich edit control. |
| [RichEdit_GetIMECompText](#RichEdit_GetIMECompText) | Retrieves the Input Method Editor (IME) composition text. |
| [RichEdit_GetIMEModeBias](#RichEdit_GetIMEModeBias) | Retrieves the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control. |
| [RichEdit_GetIMEOptions](#RichEdit_GetIMEOptions) | Retrieves the current Input Method Editor (IME) options. This message is available only in Asian-language versions of the operating system. |
| [RichEdit_GetIMEProperty](#RichEdit_GetIMEProperty) | Retrieves the property and capabilities of the Input Method Editor (IME) associated with the current input locale. |
| [RichEdit_GetLangOptions](#RichEdit_GetLangOptions) | Retrieves a rich edit control's option settings for Input Method Editor (IME) and Asian language support. |
| [RichEdit_GetLimitText](#RichEdit_GetLimitText) | Retrieves the current text limit for a rich edit control. |
| [RichEdit_GetLine](#RichEdit_GetLine) | Copies a line of text from a rich edit control. |
| [RichEdit_GetLineCount](#RichEdit_GetLineCount) | Retrieves the number of lines in a multiline rich edit control. |
| [RichEdit_GetModify](#RichEdit_GetModify) | Retrieves the state of a rich edit control's modification flag. The flag indicates whether the contents of the rich edit control have been modified. |
| [RichEdit_GetOleInterface](#RichEdit_GetOleInterface) | Retrieves an IRichEditOle object that a client can use to access a rich edit control's Component Object Model (COM) functionality. |
| [RichEdit_GetOptions](#RichEdit_GetOptions) | Retrieves rich edit control options. |
| [RichEdit_GetPageRotate](#RichEdit_GetPageRotate) | Deprecated. Retrieves the text layout for a Microsoft Rich Edit control. |
| [RichEdit_GetParaFormat](#RichEdit_GetParaFormat) | Retrieves the paragraph formatting of the current selection in a rich edit control. |
| [RichEdit_GetPasswordChar](#RichEdit_GetPasswordChar) | Retrieves the password character that a rich edit control displays when the user enters text. |
| [RichEdit_GetPunctuation](#RichEdit_GetPunctuation) | Retrieves the current punctuation characters for the rich edit control. |
| [RichEdit_GetRect](#RichEdit_GetRect) | Retrieves the formatting rectangle of a rich edit control. |
| [RichEdit_GetRedoName](#RichEdit_GetRedoName) | Retrieves the type of the next action, if any, in the control's redo queue. |
| [RichEdit_GetScrollPos](#RichEdit_GetScrollPos) | Obtains the current scroll position of the edit control. |
| [RichEdit_GetSel](#RichEdit_GetSel) | Retrieves the starting and ending character positions of the current selection in a rich edit control. |
| [RichEdit_GetSelText](#RichEdit_GetSelText) | Retrieves the currently selected text in a rich edit control. |
| [RichEdit_GetText](#RichEdit_GetText) | Retrieves the text from a rich edit control. |
| [RichEdit_GetTextEx](#RichEdit_GetTextEx) | Retrieves all of the text from the rich edit control in any particular code base you want. |
| [RichEdit_GetTextLength](#RichEdit_GetTextLength) | Retrieves the length of all text in a rich edit control. |
| [RichEdit_GetTextLengthEx](#RichEdit_GetTextLengthEx) | Calculates text length in various ways. It is usually called before creating a buffer to receive the text from the control. |
| [RichEdit_GetTextMode](#RichEdit_GetTextMode) | Retrieves the current text mode and undo level of a rich edit control. |
| [RichEdit_GetTextRange](#RichEdit_GetTextRange) | Retrieves a specified range of characters from a rich edit control. |
| [RichEdit_GetThumb](#RichEdit_GetThumb) | Retrieves the position of the scroll box (thumb) in the vertical scroll bar of a multiline rich edit control. |
| [RichEdit_GetTypographyOptions](#RichEdit_GetTypographyOptions) | Returns the current state of the typography options of a rich edit control. |
| [RichEdit_GetUndoName](#RichEdit_GetUndoName) | Retrieves the type of the next undo action, if any. |
| [RichEdit_GetWordBreakProc](#RichEdit_GetWordBreakProc) | Retrieves the address of the current Wordwrap function. |
| [RichEdit_GetWordBreakProcEx](#RichEdit_GetWordBreakProcEx) | Retrieves the address of the currently registered extended word-break procedure. |
| [RichEdit_GetWordWrapMode](#RichEdit_GetWordWrapMode) | Retrieves the current word wrap and word-break options for the rich edit control. |
| [RichEdit_GetZoom](#RichEdit_GetZoom) | Retrieves the current zoom ratio, which is always between 1/64 and 64. |
| [RichEdit_HideSelection](#RichEdit_HideSelection) | Hides or shows the selection in a rich edit control. |
| [RichEdit_IsIME](#RichEdit_IsIME) | Determines if current input locale is an East Asian locale. |
| [RichEdit_LimitText](#RichEdit_LimitText) | Sets the text limit of a rich edit control. The text limit is the maximum amount of text, in characters, that the user can type into the edit control. |
| [RichEdit_LineFromChar](#RichEdit_LineFromChar) | Retrieves the index of the line that contains the specified character index in a multiline rich edit control. |
| [RichEdit_LineIndex](#RichEdit_LineIndex) | Retrieves the character index of the first character of a specified line in a multiline rich edit control. |
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
| [RichEdit_SetLimitText](#RichEdit_SetLimitText) | Sets the text limit of a rich edit control. The text limit is the maximum amount of text, in characters, that the user can type into the edit control. |
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
| AURL_DISABLEMIXEDLGC | **Windows 8**: Disable recognition of domain names that contain labels with characters belonging to more than one of the following scripts: Latin, Greek, and Cyrillic. |
| AURL_ENABLEDRIVELETTERS | **Windows 8**: Recognize file names that have a leading drive specification, such as c:\temp. |
| AURL_ENABLEEA | This value is deprecated; use **AURL_ENABLEEAURLS** instead. |
| AURL_ENABLEEAURLS | Recognize URLs that contain East Asian characters. |
| AURL_ENABLEEMAILADDR | **Windows 8**: Recognize email addresses. |
| AURL_ENABLETELNO | **Windows 8**: Recognize telephone numbers. |
| AURL_ENABLEURL | **Windows 8**: Recognize URLs that include the path. |

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

If the operation succeeds, the return value is **TRUE**.

If the operation fails, the return value is **FALSE**.

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
| *fOperation* | pecifies the find operation. This parameter can be one of the following values.<br>**WB_CLASSIFY**. Returns the character class and word-break flags of the character at the specified position.<br>**WB_ISDELIMITER**. Returns **TRUE** if the character at the specified position is a delimiter, or **FALSE** otherwise.<br>**WB_LEFT**. Finds the nearest character before the specified position that begins a word.<br>**WB_LEFTBREAK**. Finds the next word end before the specified position. This value is the same as WB_PREVBREAK.<br>**WB_MOVEWORDLEFT**. Finds the next character that begins a word before the specified position. This value is used during CTRL+LEFT ARROW key processing. This value is the similar to WB_MOVEWORDPREV. See Remarks for more information.<br>**WB_MOVEWORDRIGHT**. Finds the next character that begins a word after the specified position. This value is used during CTRL+right key processing. This value is similar to WB_MOVEWORDNEXT. See Remarks for more information.<br>**WB_RIGHT**. Finds the next character that begins a word after the specified position.<br>**WB_RIGHTBREAK**. Finds the next end-of-word delimiter after the specified position. This value is the same as WB_NEXTBREAK. |
| *dwStartPos* | Zero-based character starting position. |

#### Return value

The message returns a value based on the *fOperation* parameter.

| Return code  | Description |
| ------------ | ----------- |
| **WB_CLASSIFY** | Returns the character class and word-break flags of the character at the specified position. |
| **WB_ISDELIMITER** | Returns **TRUE** if the character at the specified position is a delimiter; otherwise it returns **FALSE**. |
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

# <a name="RichEdit_GetAutoUrlDetect"></a>RichEdit_GetAutoUrlDetect

Indicates whether the auto URL detection is turned on in the rich edit control.

```
FUNCTION RichEdit_GetAutoUrlDetect (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETAUTOURLDETECT, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value
If auto-URL detection is active, the return value is 1.

If auto-URL detection is inactive, the return value is 0.

#### Remarks

When auto URL detection is on, Microsoft Rich Edit is constantly checking typed text for a valid URL. Rich Edit recognizes URLs that start with these prefixes:

- http:
- file:
- mailto:
- ftp:
- https:
- gopher:
- nntp:
- prospero:
- telnet:
- news:
- wais:
- outlook:

Rich Edit also recognizes standard path names that start with \\\. When Rich Edit locates a URL, it changes the URL text color, underlines the text, and notifies the client using [EN_LINK](https://learn.microsoft.com/en-us/windows/win32/controls/en-link).

# <a name="RichEdit_GetBidiOptions"></a>RichEdit_GetBidiOptions

Indicates the current state of the bidirectional options in the rich edit control.

```
SUB RichEdit_GetBidiOptions (BYVAL hRichEdit AS HWND, BYVAL lpbo AS BIDIOPTIONS PTR)
   SendMessageW hRichEdit, EM_GETBIDIOPTIONS, 0, cast(LPARAM, lpbo)
END SUB
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lpbo* | A pointer to a [BIDIOPTIONS](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-bidioptions) structure that receives the current state of the bidirectional options in the rich edit control. |

#### Remarks

This message sets the values of the **wMask** and **wEffects** members to the value of the current state of the bidirectional options in the rich edit control.

# <a name="RichEdit_GetCharFormat"></a>RichEdit_GetCharFormat

Determines the character formatting in a rich edit control.

```
FUNCTION RichEdit_GetCharFormat (BYVAL hRichEdit AS HWND, BYVAL fOption AS DWORD, BYVAL lpcf AS CHARFORMATW PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETCHARFORMAT, fOption, cast(LPARAM, lpcf))
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *fOption* | Specifies the range of text from which to retrieve formatting. It can be one of the following values.<br>**SCF_DEFAULT**The default character formatting.<br>**SCF_SELECTION** The current selection's character formatting. |
| *lpcf* | A pointer tp a [CHARFORMAT](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformata) structure that receives the attributes of the first character. The **dwMask** member specifies which attributes are consistent throughout the entire selection. For example, if the entire selection is either in italics or not in italics, CFM_ITALIC is set; if the selection is partly in italics and partly not, CFM_ITALIC is not set. |

#### Return value

This message returns the value of the **dwMask** member of the [CHARFORMAT](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformata) structure.

# <a name="RichEdit_GetCTFModeBias"></a>RichEdit_GetCTFModeBias

Retrieves the Text Services Framework mode bias values for a Microsoft Rich Edit control.

```
FUNCTION RichEdit_GetCTFModeBias (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETCTFMODEBIAS, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value
The current Text Services Framework mode bias value.

#### Remarks

To get the IME mode bias, call **RichEdit_GetIMEModeBias**.

# <a name="RichEdit_GetCTFOpenStatus"></a>RichEdit_GetCTFOpenStatus

Determines if the Text Services Framework (TSF) keyboard is open or closed.

```
FUNCTION RichEdit_GetCTFOpenStatus (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETCTFOPENSTATUS, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

If the TSF keyboard is open, the return value is **TRUE**. Otherwise, it is **FALSE**.

# <a name="RichEdit_GetEditStyle"></a>RichEdit_GetEditStyle

Retrieves the current edit style flags.

```
FUNCTION RichEdit_GetEditStyle (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETEDITSTYLE, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

Returns the current edit style flags, which can include one or more of the following values:

| Return code  | Description |
| ------------ | ----------- |
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

# <a name="RichEdit_GetEventMask"></a>RichEdit_GetEventMask

Retrieves the event mask for a rich edit control. The event mask specifies which notification messages the control sends to its parent window.

```
FUNCTION RichEdit_GetEventMask (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETEVENTMASK, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

This message returns the event mask for the rich edit control.

# <a name="RichEdit_GetFirstVisibleLine"></a>RichEdit_GetFirstVisibleLine

Retrieves the zero-based index of the uppermost visible line in a multiline edit control. You can send this message to either an edit control or a rich edit control.

```
FUNCTION RichEdit_GetFirstVisibleLine (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETFIRSTVISIBLELINE, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The return value is the zero-based index of the uppermost visible line in a multiline edit control.

For single-line rich edit controls, the return value is zero.

# <a name="RichEdit_GetHyphenateInfo"></a>RichEdit_GetHyphenateInfo

Retrieves information about hyphenation for a Microsoft Rich Edit control.

```
SUB RichEdit_GetHyphenateInfo (BYVAL hRichEdit AS HWND, BYVAL lphi AS HYPHENATEINFO PTR)
   SendMessageW hRichEdit, EM_GETHYPHENATEINFO, cast(WPARAM, lphi), 0
END SUB
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lphi* | A pointer to a [HYPHENATEINFO](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-hyphenateinfo) structure. |

# <a name="RichEdit_GetIMEColor"></a>RichEdit_GetIMEColor

Retrieves the Input Method Editor (IME) composition color. This message is available only in Asian-language versions of the operating system.

```
FUNCTION RichEdit_GetIMEColor (BYVAL hRichEdit AS HWND, BYVAL rgCmpclr AS COMPCOLOR PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMECOLOR, 0, cast(LPARAM, rgCmpclr))
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *rgCmpclr* | A four-element array of [COMPCOLOR](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-compcolor) structures that receives the composition color. |

#### Return value

If the operation succeeds, the return value is a nonzero value.

If the operation fails, the return value is zero.

# <a name="RichEdit_GetIMECompMode"></a>RichEdit_GetIMECompMode

Retrieves the current Input Method Editor (IME) mode for a rich edit control.

```
FUNCTION RichEdit_GetIMECompMode (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMECOMPMODE, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The return value is one of the following values.

| Return code | Description |
| ----------- | ----------- |
| **ICM_NOTOPEN** | IME is not open. |
| **ICM_LEVEL3** | True inline mode. |
| **ICM_LEVEL2** | Level 2. |
| **ICM_LEVEL2_5** | Level 2.5. |
| **ICM_LEVEL2_SUI** | Special UI. |

# <a name="RichEdit_GetIMECompText"></a>RichEdit_GetIMECompText

Retrieves the Input Method Editor (IME) composition text.

```
FUNCTION RichEdit_GetIMECompText (BYVAL hRichEdit AS HWND, BYVAL lpict AS IMECOMPTEXT PTR, BYVAL buffer AS ANY PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMECOMPTEXT, cast(WPARAM, lpict), cast(LPARAM, buffer))
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lpict* | A pointer to the [IMECOMPTEXT](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-imecomptext) structure. |
| *buffer* | The buffer that receives the composition text. The size of this buffer is contained in the *cb* member of the IMECOMPTEXT structure. |

#### Return value
If successful, the return value is the number of Unicode characters copied to the buffer. Otherwise, it is zero.

#### Remarks

This message only takes Unicode strings.

**Security Warning**: Be sure to have a buffer sufficient for the size of the input. Failure to do so could cause problems for your application.

# <a name="RichEdit_GetIMEModeBias"></a>RichEdit_GetIMEModeBias

Retrieves the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.

```
FUNCTION RichEdit_GetIMEModeBias (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMEMODEBIAS, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

This message returns the current IME mode bias setting.

#### Remarks

To get the Text Services Framework mode bias, use **RichEdit_GetCTFModeBias**.

The application should call **RichEdit_IsIME** before calling this function.

# <a name="RichEdit_GetIMEOptions"></a>RichEdit_GetIMEOptions

Retrieves the current Input Method Editor (IME) options. This message is available only in Asian-language versions of the operating system.

```
FUNCTION RichEdit_GetIMEOptions (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMEOPTIONS, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

This message returns one or more of the IME option flag values described in the **RichEdit_SetIMEOptions** message.

# <a name="RichEdit_GetIMEProperty"></a>RichEdit_GetIMEProperty

Retrieves the property and capabilities of the Input Method Editor (IME) associated with the current input locale.

```
FUNCTION RichEdit_GetIMEOptions (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMEOPTIONS, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

This message returns one or more of the IME option flag values described in the **RichEdit_SetIMEOptions** message.

# <a name="RichEdit_GetLangOptions"></a>RichEdit_GetLangOptions

Retrieves a rich edit control's option settings for Input Method Editor (IME) and Asian language support.

```
FUNCTION RichEdit_GetLangOptions (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETLANGOPTIONS, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

Returns the IME and Asian language settings, which can be zero or more of the following values.

| Return code  | Description |
| ------------ | ----------- |
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

**Remarks**

The **IMF_AUTOFONT** flag is set by default. The **IMF_AUTOKEYBOARD** and **IMF_IMECANCELCOMPLETE** flags are cleared by default.

# <a name="RichEdit_GetLimitText"></a>RichEdit_GetLimitText

Retrieves the current text limit for an edit control. You can send this message to either an edit control or a rich edit control.

```
FUNCTION RichEdit_GetLimitText (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETLIMITTEXT, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The return value is the text limit.

#### Remarks

The text limit is the maximum amount of text, in characters, that the control can contain. For ANSI text, this is the number of bytes; for Unicode text, this is the number of characters. Two documents with the same character limit will yield the same text limit, even if one is ANSI and the other is Unicode.

# <a name="RichEdit_GetLine"></a>RichEdit_GetLine

Copies a line of text from a rich edit control.

```
FUNCTION RichEdit_GetLine (BYVAL hRichEdit AS HWND, BYVAL which AS DWORD) AS CWSTR
   DIM buffer AS CWSTR = MKI(32765) + STRING(32765, 0)
   DIM n AS LONG = SendMessageW(hRichEdit, EM_GETLINE, which, cast(LPARAM, *buffer))
   RETURN LEFT(**buffer, n)
END FUNCTION
```
**Note**: Before sending the EM_GETLINE message, the first word of the buffer has to be set to the size, in characters, of the buffer. The size in the first word is overwritten by the copied line.

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *which* | The zero-based index of the line to retrieve from a multiline edit control. A value of zero specifies the topmost line. This parameter is ignored by a single-line edit control. |

#### Retuen value

A copy of the line.

# <a name="RichEdit_GetLineCount"></a>RichEdit_GetLineCount

Retrieves the number of lines in a multiline rich edit control.

```
FUNCTION RichEdit_GetLineCount (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETLINECOUNT, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The return value is an integer specifying the total number of text lines in the multiline edit control or rich edit control. If the control has no text, the return value is 1. The return value will never be less than 1.

### Remarks

The **EM_GETLINECOUNT** message retrieves the total number of text lines, not just the number of lines that are currently visible.

If the Wordwrap feature is enabled, the number of lines can change when the dimensions of the editing window change.

# <a name="RichEdit_GetModify"></a>RichEdit_GetModify

Retrieves the state of a rich edit control's modification flag. The flag indicates whether the contents of the rich edit control have been modified.

```
FUNCTION RichEdit_GetModify (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETMODIFY, 0, 0)
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

If the contents of edit control have been modified, the return value is nonzero; otherwise, it is zero.

### Remarks

The system automatically clears the modification flag to zero when the control is created. If the user changes the control's text, the system sets the flag to nonzero. You can send the **RichEdit_SetModify** message to the edit control to set or clear the flag.

# <a name="RichEdit_GetOleInterface"></a>RichEdit_GetOleInterface

Retrieves an **IRichEditOle** object that a client can use to access a rich edit control's Component Object Model (COM) functionality.

```
FUNCTION RichEdit_GetOleInterface (BYVAL hRichEdit AS HWND, BYVAL ppObject AS IUnknown PTR PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETOLEINTERFACE, 0, cast(LPARAM, @ppObject))
END FUNCTION
```
| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *ppObject* | Pointer to a pointer that receives the **IRichEditOle** object. The control calls the **AddRef** method for the object before returning, so the calling application must call the **Release** method when it is done with the object. |

#### Return value

If the operation succeeds, the return value is a nonzero value.

If the operation fails, the return value is zero.

# <a name="RichEdit_GetOptions"></a>RichEdit_GetOptions

Retrieves rich edit control options.

```
FUNCTION RichEdit_GetOptions (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETOPTIONS, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

This message returns a combination of the current option flag values described in the **RichEdit_SetOptions** message.

# <a name="RichEdit_GetPageRotate"></a>RichEdit_GetPageRotate

Deprecated. Retrieves the text layout for a Microsoft Rich Edit control.

```
FUNCTION RichEdit_GetPageRotate (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETPAGEROTATE, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The current text layout. For a list of possible text layout values, see **RichEdit_SetPageRotate**.

# <a name="RichEdit_GetParaFormat"></a>RichEdit_GetParaFormat

Retrieves the paragraph formatting of the current selection in a rich edit control.

```
FUNCTION RichEdit_GetParaFormat (BYVAL hRichEdit AS HWND, BYVAL pParaFmt AS PARAFORMAT PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETPARAFORMAT, 0, cast(LPARAM, pParaFmt))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *pParaFmt* | Pointer to a [PARAFORMAT](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-paraformat) structure that receives the paragraph formatting attributes of the current selection. If more than one paragraph is selected, the structure receives the attributes of the first paragraph, and the **dwMask** member specifies which attributes are consistent throughout the entire selection. |

#### Return value

This message returns the value of the dwMask member of the [PARAFORMAT](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-paraformat) structure.

# <a name="RichEdit_GetPasswordChar"></a>RichEdit_GetPasswordChar

Retrieves the password character that a rich edit control displays when the user enters text.

```
FUNCTION RichEdit_GetPasswordChar (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETPASSWORDCHAR, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The return value specifies the character to be displayed in place of any characters typed by the user. If the return value is **NULL**, there is no password character, and the control displays the characters typed by the user.

#### Remarks

If an edit control is created with the **ES_PASSWORD** style, the default password character is set to an asterisk (*). If an edit control is created without the **ES_PASSWORD** style, there is no password character. To change the password character, send the **RichEdit_SetPasswordChar** message.

# <a name="RichEdit_GetPunctuation"></a>RichEdit_GetPunctuation

Retrieves the current punctuation characters for the rich edit control.

```
FUNCTION RichEdit_GetPunctuation (BYVAL hRichEdit AS HWND, BYVAL punctp AS DWORD, BYVAL lppunct AS PUNCTUATION PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETPUNCTUATION, punctp, cast(LPARAM, lppunct))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *punctp* | The punctuation type can be one of the following values.<br>**PC_LEADING**. Leading punctuation characters.<br>**PC_FOLLOWING**. Following punctuation characters.<br>**PC_DELIMITER**. Delimiter.<br>**PC_OVERFLOW**. Not supported |
| *lppunct* | Pointer to a [PUNCTUATION](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-punctuation) structure that receives the punctuation characters. |

#### Return value

If the operation succeeds, the return value is a nonzero value.

If the operation fails, the return value is zero.

# <a name="RichEdit_GetRect"></a>RichEdit_GetRect

Retrieves the formatting rectangle of an edit control. The formatting rectangle is the limiting rectangle into which the control draws the text. The limiting rectangle is independent of the size of the edit-control window.

```
SUB RichEdit_GetRect (BYVAL hRichEdit AS HWND, BYVAL prc AS RECT PTR)
   SendMessageW hRichEdit, EM_GETRECT, 0, cast(LPARAM, prc)
END SUB
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *prc* | A pointer to a [RECT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-rect) structure that receives the formatting rectangle |

#### Remarks

You can modify the formatting rectangle of a multiline edit control by using the **RichEdit_SeRect** and **RichEdit_SeRectNP** messages.

Under certain conditions, **RichEdit_GetRect** might not return the exact values that **RichEdit_SeRect** and **RichEdit_SeRectNP** set it will be approximately correct, but it can be off by a few pixels.

The formatting rectangle does not include the selection bar, which is an unmarked area to the left of each paragraph. When clicked, the selection bar selects the line.

# <a name="RichEdit_GetRedoName"></a>RichEdit_GetRedoName

Retrieves the type of the next action, if any, in the rich edit control's redo queue.

```
FUNCTION RichEdit_GetRedoName (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETREDONAME, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

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

# <a name="RichEdit_GetScrollPos"></a>RichEdit_GetScrollPos

Retrieves the current scroll position of the edit control.

```
FUNCTION RichEdit_GetScrollPos (BYVAL hRichEdit AS HWND, BYVAL lppt AS POINT PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETSCROLLPOS, 0, cast(LPARAM, lppt))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lppt* | Pointer to a [POINT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-point) structure. After calling **RichEdit_GetScrollPos**, this parameters contains a point in the virtual text space of the document, expressed in pixels. This point will be the point that is currently located in the upper-left corner of the edit control window. |

#### Remarks

The values returned in the [POINT](https://learn.microsoft.com/en-us/windows/win32/api/windef/ns-windef-point) structure are 16-bit values (even in the 32-bit wide fields).

# <a name="RichEdit_GetSel"></a>RichEdit_GetSel

Retrieves the starting and ending character positions of the current selection in a rich edit control.

```
FUNCTION RichEdit_GetSel (BYVAL hRichEdit AS HWND, BYVAL pdwStartPos AS DWORD PTR, BYVAL pdwEndPos AS DWORD PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETSEL, cast(WPARAM, pdwStartPos), cast(LPARAM, pdwEndPos))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *pdwStartPos* | A pointer to a **DWORD** value that receives the starting position of the selection. This parameter can be **NULL**. |
| *pdwEndPos* | A pointer to a **DWORD** value that receives the position of the first unselected character after the end of the selection. This parameter can be **NULL**. |

#### Return value

The return value is a zero-based value with the starting position of the selection in the **LOWORD** and the position of the first character after the last selected character in the **HIWORD**. If either of these values exceeds 65,535, the return value is -1.

It is better to use the values returned in *pdwStartPos* and *pdwEndPos* because they are full 32-bit values.

#### Remarks

If there is no selection, the starting and ending values are both the position of the caret.

You can also use the **EM_EXGETSEL** message to retrieve the same information. **EM_EXGETSEL** also returns starting and ending character positions as 32-bit values. A combination of the use of **EM_EXGETSEL** and **EM_GETSELTEXT** are used ine the **RichEdit_GetSelText** function to retrieve the selected text as a **CWSTR**.

# <a name="RichEdit_GetSelText"></a>RichEdit_GetSelText

Retrieves the currently selected text in a rich edit control.

```
FUNCTION RichEdit_GetSelText (BYVAL hRichEdit AS HWND) AS CWSTR
   DIM dwStartPos AS DWORD, dwEndPos AS DWORD, cr AS CHARRANGE
   SendMessageW(hRichEdit, EM_EXGETSEL, 0, cast(LPARAM, @cr))
   DIM cbLen AS DWORD = ABS(cr.cpMax - cr.cpMin)
   IF cbLen < 1 THEN RETURN ""
   DIM cwsText AS CWSTR = cbLen + 1
   cbLen = SendMessageW(hRichEdit, EM_GETSELTEXT, 0, cast(LPARAM, *cwsText))
   RETURN LEFT(**cwsText, cbLen)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The selected text as a **CWSTR** (dynamic unicode string).

# <a name="RichEdit_GetText"></a>RichEdit_GetText

Retrieves the text from a rich edit control.

```
FUNCTION RichEdit_GetText (BYVAL hRichEdit AS HWND) AS CWSTR
   DIM cbLen AS DWORD = SendMessageW(hRichEdit, WM_GETTEXTLENGTH, 0, 0)
   IF cbLen < 1 THEN RETURN ""
   DIM cwsText AS CWSTR = cbLen + 1
   cbLen = SendMessageW(hRichEdit, WM_GETTEXT, cbLen + 1, cast(LPARAM, *cwsText))
   RETURN LEFT(**cwsText, cbLen)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Remarks

The Windows API function **GetWindowTextW** can also be used to retrive the text of a rich edit control, but it cannot retrieve the text of a control in another application.

# <a name="RichEdit_GetTextEx"></a>RichEdit_GetTextEx

Retrieves all of the text from the rich edit control in any particular code base you want.

```
FUNCTION RichEdit_GetTextEx (BYVAL hRichEdit AS HWND, BYVAL lpgtex AS GETTEXTEX PTR, BYVAL buffer AS ANY PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETTEXTEX, cast(WPARAM, lpgtex), cast(LPARAM, buffer))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lpgtex* | Pointer to a [GETTEXTEX](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-gettextex) structure, which indicates how to translate the text before putting it into the output buffer. |
| *buffer* | Pointer to the buffer to receive the text. The size of this buffer, in bytes, is specified by the *cb* member of the [GETTEXTEX](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-gettextex) structure. Use the **RichEdit_GetTextLength** message to get the required size of the buffer. |

#### Return value

The return value is the number of characters copied into the output buffer, not including the null terminator.

#### Remarks

If the size of the output buffer is less than the size of the text in the control, the edit control will copy text from its beginning and place it in the buffer until the buffer is full. A terminating null character will still be placed at the end of the buffer.

# <a name="RichEdit_GetTextLength"></a>RichEdit_GetTextLength

Retrieves the length of all text in a rich edit control.

```
FUNCTION RichEdit_GetTextLength (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, WM_GETTEXTLENGTH, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

####Return value

The return value is the length of the text in characters, not including the terminating null character.

#### Remarks

When the **WM_GETTEXTLENGTH** message is sent, the **DefWindowProc** function returns the length, in characters, of the text. Under certain conditions, the **DefWindowProc** function returns a value that is larger than the actual length of the text. This occurs with certain mixtures of ANSI and Unicode, and is due to the system allowing for the possible existence of double-byte character set (DBCS) characters within the text. The return value, however, will always be at least as large as the actual length of the text; you can thus always use it to guide buffer allocation. This behavior can occur when an application uses both ANSI functions and common dialogs, which use Unicode.

To retrieve the text, you can also use the **AfxGetWindowText** function, the **WM_GETTEXT** message, or the Windows API **GetWindowTextW** function.

# <a name="RichEdit_GetTextLengthEx"></a>RichEdit_GetTextLengthEx

Calculates text length in various ways. It is usually called before creating a buffer to receive the text from the control.

```
FUNCTION RichEdit_GetTextLengthEx (BYVAL hRichEdit AS HWND, BYVAL lpgtex AS GETTEXTLENGTHEX PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETTEXTLENGTHEX, cast(WPARAM, lpgtex), 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lpgtex* | Pointer to a [GETTEXTLENGTHEX](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-gettextlengthex) structure that receives the text length information. |

#### Return value

The message returns the number of characters in the edit control, depending on the setting of the flags in the [GETTEXTLENGTHEX](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-gettextlengthex) structure. If incompatible flags were set in the *flags* member, the message returns E_INVALIDARG .

#### Remarks

This message is a fast and easy way to determine the number of characters in the Unicode version of the rich edit control. However, for a non-Unicode target code page you will potentially be converting to a combination of single-byte and double-byte characters.

# <a name="RichEdit_GetTextMode"></a>RichEdit_GetTextMode

Retrieves the current text mode and undo level of a rich edit control.

```
FUNCTION RichEdit_GetTextMode (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETTEXTMODE, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The return value is one or more values from the [TEXTMODE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ne-richedit-textmode) enumeration type. The values indicate the current text mode and undo level of the control.

#### TEXTMODE enumeration

| Name              | Value | Description |
| ----------------- | ----- | ----------- |
| **TM_PLAINTEXT**  |   1   | Indicates plain-text mode, in which the control is similar to a standard edit control. |
| **TM_RICHTEXT**   |   2   | Indicates rich-text mode, in which the control has the standard rich edit functionality. Rich-text mode is the default setting. |
| **TM_SINGLELEVELUNDO**  |   4   | The control allows the user to undo only the last action in the undo queue. |
| **TM_MULTILEVELUNDO**       |   8   | The control supports multiple undo actions. This is the default setting. Use the *RichEdit_SetUndoLimit* message to set the maximum number of undo actions. |
| **TM_SINGLECODEPAGE** |  16   | The control only allows the English keyboard and a keyboard corresponding to the default character set. For example, you could have Greek and English. Note that this prevents Unicode text from entering the control. For example, use this value if a Rich Edit control must be restricted to ANSI text. |
| **TM_MULTICODEPAGE** |  32   | The control allows multiple code pages and Unicode text into the control. This is the default setting. |

# <a name="RichEdit_GetTextRange"></a>RichEdit_GetTextRange

Retrieves a specified range of characters from a rich edit control.

```
FUNCTION RichEdit_GetTextRange (BYVAL hRichEdit AS HWND, BYVAL lptrg AS TEXTRANGE PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETTEXTRANGE, 0, cast(LPARAM, lptrg))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *lptrg* | Pointer to a [TEXTRANGE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-textrangea) structure that specifies the range of characters to retrieve and a buffer to copy the characters to. |

#### Return value

The message returns the number of characters copied, not including the terminating null character.

# <a name="RichEdit_GetThumb"></a>RichEdit_GetThumb

Retrieves the position of the scroll box (thumb) in the vertical scroll bar of a multiline edit control.

```
FUNCTION RichEdit_GetThumb (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETTHUMB, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The return value is the position of the scroll box.

# <a name="RichEdit_GetTypographyOptions"></a>RichEdit_GetTypographyOptions

Returns the current state of the typography options of a rich edit control.

```
FUNCTION RichEdit_GetTypographyOptions (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETTYPOGRAPHYOPTIONS, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

Returns the current typography options. For a list of options, see [EM_SETTYPOGRAPHYOPTIONS](https://learn.microsoft.com/en-us/windows/win32/controls/em-settypographyoptions).

#### Remarks

You can turn on advanced line breaking by sending the **RichEdit_SetTypographyOPtions** message. Advanced and normal line breaking may also be turned on automatically by the rich edit control if it is needed for certain languages.

# <a name="RichEdit_GetUndoName"></a>RichEdit_GetUndoName

Retrieves the type of the next undo action, if any.

```
FUNCTION RichEdit_GetUndoName (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETUNDONAME, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

If there is an undo action, the value returned is an [UNDONAMEID](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ne-richedit-undonameid) enumeration value that indicates the type of the next action in the control's undo queue.

If there are no actions that can be undone or the type of the next undo action is unknown, the return value is zero.

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

The types of actions that can be undone or redone include typing, delete, drag, drop, cut, and paste operations. This information can be useful for applications that provide an extended user interface for undo and redo operations, such as a drop-down list box of actions that can be undone.

# <a name="RichEdit_GetWordBreakProc"></a>RichEdit_GetWordBreakProc

Retrieves the address of the current Wordwrap function.

```
FUNCTION RichEdit_GetWordBreakProc (BYVAL hRichEdit AS HWND) AS LONG_PTR
   FUNCTION = SendMessageW(hRichEdit, EM_GETWORDBREAKPROC, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The return value specifies the address of the application-defined Wordwrap function. The return value is **NULL** if no Wordwrap function exists.

#### Remarks

A Wordwrap function scans a text buffer that contains text to be sent to the display, looking for the first word that does not fit on the current display line. The wordwrap function places this word at the beginning of the next line on the display. A Wordwrap function defines the point at which the system should break a line of text for multiline edit controls, usually at a space character that separates two words.

# <a name="RichEdit_GetWordBreakProcEx"></a>RichEdit_GetWordBreakProcEx

Retrieves the address of the currently registered extended word-break procedure.

```
FUNCTION RichEdit_GetWordBreakProcEx (BYVAL hRichEdit AS HWND) AS LONG_PTR
   FUNCTION = SendMessageW(hRichEdit, EM_GETWORDBREAKPROCEX, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The message returns the address of the current procedure.

# <a name="RichEdit_GetWordWrapMode"></a>RichEdit_GetWordWrapMode

Retrieves the current word wrap and word-break options for the rich edit control.

```
FUNCTION RichEdit_GetWordWrapMode (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETWORDWRAPMODE, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

The message returns the current word wrap and word-break options.

#### Remarks

This message is supported only in Asian-language versions of Microsoft Rich Edit 1.0. It is not supported in any later versions of Rich Edit. This message must not be sent by the application-defined, word-break procedure.

# <a name="RichEdit_GetZoom"></a>RichEdit_GetZoom

Retrieves the current zoom ratio, which is always between 1/64 and 64.

```
FUNCTION RichEdit_GetZoom (BYVAL hRichEdit AS HWND, BYVAL pzNum AS DWORD PTR, BYVAL pzDen AS DWORD PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETZOOM, cast(WPARAM, pzNum), cast(LPARAM, pzDen))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *pzNum* | A pointer to a **DWORD** value that receives the numerator of the zoom ratio. |
| *pzDen* | A pointer to a **DWORD** value that receives the denominator of the zoom ratio.. |

#### Return value
The message returns **TRUE** if message is processed, which it will be if both *pzNum* and *pzDen* are not **NULL**.

# <a name="RichEdit_HideSelection"></a>RichEdit_HideSelection

Hides or shows the selection in a rich edit control.

```
SUB RichEdit_HideSelection (BYVAL hRichEdit AS HWND, BYVAL fHide AS DWORD)
   SendMessageW hRichEdit, EM_HIDESELECTION, fHide, 0
END SUB
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *fHide* | Value specifying whether to hide or show the selection. If this parameter is zero, the selection is shown. Otherwise, the selection is hidden. |

# <a name="RichEdit_IsIME"></a>RichEdit_IsIME

Determines if current input locale is an East Asian locale.

```
FUNCTION RichEdit_IsIME (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_ISIME, 0, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |

#### Return value

Returns **TRUE** if it is an East Asian locale. Otherwise, it returns **FALSE**.

# <a name="RichEdit_LimitText"></a>RichEdit_LimitText

Sets the text limit of a rich edit control. The text limit is the maximum amount of text, in characters, that the user can type into the edit control.

```
SUB RichEdit_LimitText (BYVAL hRichEdit AS HWND, BYVAL chMax AS DWORD)
   SendMessageW hRichEdit, EM_LIMITTEXT, chMax, 0
END SUB
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *chMax* | The maximum number of characters the user can enter. If this parameter is zero, the text length is set to 64,000 characters. |

#### Remarks

The **RichEdit_LimitText** message limits only the text the user can enter. It does not affect any text already in the edit control when the message is sent, nor does it affect the length of the text copied to the edit control by the **RichEdit_SetText** message. If an application uses the **RichEdit_SetTExt** message to place more text into an edit control than is specified in the **RichEdit_LimitText** message, the user can edit the entire contents of the edit control.

# <a name="RichEdit_LineFromChar"></a>RichEdit_LineFromChar

Retrieves the index of the line that contains the specified character index in a multiline rich edit control.

```
FUNCTION RichEdit_LineFromChar (BYVAL hRichEdit AS HWND, BYVAL index AS DWORD) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_LINEFROMCHAR, index, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *index* | The character index of the character contained in the line whose number is to be retrieved. If this parameter is -1, **RichEdit_LineFromChar** retrieves either the line number of the current line (the line containing the caret) or, if there is a selection, the line number of the line containing the beginning of the selection. |

#### Return value

The return value is the zero-based line number of the line containing the character index specified by wParam.

# <a name="RichEdit_LineIndex"></a>RichEdit_LineIndex

Retrieves the character index of the first character of a specified line in a multiline rich edit control.

```
FUNCTION RichEdit_LineIndex (BYVAL hRichEdit AS HWND, BYVAL nLine AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_LINEINDEX, nLine, 0)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRichEdit* | The handle of the rich edit control. |
| *nLine* | The zero-based line number. A value of -1 specifies the current line number (the line that contains the caret). |

#### Return value

The return value is the character index of the line specified in the wParam parameter, or it is -1 if the specified line number is greater than the number of lines in the edit control.

