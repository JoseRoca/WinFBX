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
| [RichEdit_LimitText](#RichEdit_LimitText) | Sets the text limit of a rich edit control. The text limit is the maximum amount of text, in TCHARs, that the user can type into the edit control. |
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

If the TSF keyboard is open, the return value is TRUE. Otherwise, it is FALSE.

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

The text limit is the maximum amount of text, in **TCHAR**s, that the control can contain. For ANSI text, this is the number of bytes; for Unicode text, this is the number of characters. Two documents with the same character limit will yield the same text limit, even if one is ANSI and the other is Unicode.

# <a name="RichEdit_GetLine"></a>RichEdit_GetLine

Copies a line of text from a rich edit control.

```
FUNCTION RichEdit_GetLine (BYVAL hRichEdit AS HWND, BYVAL which AS DWORD) AS CWSTR
   DIM buffer AS CWSTR = MKI(32765) + STRING(32765, 0)
   DIM n AS LONG = SendMessageW(hRichEdit, EM_GETLINE, which, cast(LPARAM, *buffer))
   RETURN LEFT(**buffer, n)
END FUNCTION
```
**Note**: Before sending the EM_GETLINE message, the first word of the buffer has to be set to the size, in **TCHAR**s, of the buffer. For ANSI text, this is the number of bytes; for Unicode text, this is the number of characters. The size in the first word is overwritten by the copied line.

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

Retrieves an IRichEditOle object that a client can use to access a rich edit control's Component Object Model (COM) functionality.

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
