# CTextDocument2 Class

### ITextDocument Interface

The **ITextDocument** interface is the Text Object Model (TOM) top-level interface, which retrieves the active selection and range objects for any story in the document—whether active or not. It enables the application to:

- Open and save documents.
- Control undo behavior and screen updating.
- Find a range from a screen position.
- Get an **ITextStoryRanges** story enumerator.

#### When to implement

Applications typically do not implement the **ITextDocument** interface. Microsoft text solutions, such as rich edit controls, implement **ITextDocument** as part of their TOM implementation.

#### When to Use

Applications can retrieve an **ITextDocument** pointer from a rich edit control. To do this, send an **EM_GETOLEINTERFACE** message to retrieve an **IRichEditOle** object from a rich edit control. Then, call the object's **IUnknown::QueryInterface** method to retrieve an **ITextDocument** pointer.

#### Inheritance

The **ITextDocument** interface inherits from the IUnknown interface. **ITextDocument** also has these types of members:

| Name       | Description |
| ---------- | ----------- |
| [GetName](#GetName) | Gets the file name of this document. |
| [GetSelection](#GetSelection) | Gets the active selection. |
| [GetStoryCount](#GetStoryCount) | Gets the count of stories in this document. |
| [GetStoryRanges](#GetStoryRanges) | Gets the story collection object used to enumerate the stories in a document. |
| [GetSaved](#GetSaved) | Gets a value that indicates whether changes have been made since the file was last saved. |
| [SetSaved](#SetSaved) | Sets the document **Saved** property. |
| [GetDefaultTabStop](#GetDefaultTabStop) | Gets the default tab width. |
| [SetDefaultTabStop](#SetDefaultTabStop) | Sets the default tab stop, which is used when no tab exists beyond the current display position. |
| [New_](#New_) | Opens a new document. |
| [Open](#Open) | Opens a specified document. There are parameters to specify access and sharing privileges, creation and conversion of the file, as well as the code page for the file. |
| [Save](#Save) | Saves the document. |
| [Freeze](#Freeze) | Increments the freeze count. |
| [Unfreeze](#Unfreeze) | Decrements the freeze count. |
| [BeginEditCollection](#BeginEditCollection) | Turns on edit collection (also called *undo grouping*). |
| [EndEditCollection](#EndEditCollection) | Turns off edit collection (also called *undo grouping*). |
| [Undo](#Undo) | Performs a specified number of undo operations. |
| [Redo](#Redo) | Performs a specified number of redo operations. |
| [Range](#Range) | Retrieves a text range object for a specified range of content in the active story of the document. |
| [RangeFromPoint](#RangeFromPoint) | Retrieves a range for the content at or nearest to the specified point on the screen. |

### ITextDocument2 Interface

Extends the **ITextDocument** interface, adding methods that enable the Input Method Editor (IME) to drive the rich edit control, and methods to retrieve other interfaces such as **ITextDisplays**, **ITextRange2**, **ITextFont2**, **ITextPara2**, and so on.

Some **ITextDocument2** methods used with the IME need access to the current window handle (HWND). Use the **GetWindow** method of the **ITextDocument2** interface to retrieve the handle.

| Name       | Description |
| ---------- | ----------- |
| [GetCaretType](#GetCaretType) | Gets the caret type. |
| [SetCaretType](#SetCaretType) | Sets the caret type. |
| [GetDisplays](#GetDisplays) | Gets the displays collection for this Text Object Model (TOM) engine instance. |
| [GetDocumentFont](#GetDocumentFont) | Gets an object that provides the default character format information for this instance of the Text Object Model (TOM) engine. |
| [SetDocumentFont](#SetDocumentFont) | Sets the default paragraph formatting for this instance of the Text Object Model (TOM) engine. |
| [GetDocumentPara](#GetDocumentPara) | Gets an object that provides the default paragraph format information for this instance of the Text Object Model (TOM) engine. |
| [SetDocumentPara](#SetDocumentPara) | Sets the default paragraph formatting for this instance of the Text Object Model (TOM) engine. |
| [GetEastAsianFlags](#GetEastAsianFlags) | Gets the East Asian flags. |
| [GetGenerator](#GetGenerator) | Gets the name of the Text Object Model (TOM) engine. |
| [SetIMEInProgress](#SetIMEInProgress) | Sets the state of the Input Method Editor (IME) in-progress flag. |
| [GetNotificationMode](#GetNotificationMode) | Gets the notification mode. |
| [SetNotificationMode](#SetNotificationMode) | Sets the notification mode. Use **tomTrue** to turn on notifications, or **tomFalse** to turn them off. |
| [GetSelection2](#GetSelection2) | Gets the active selection. |
| [GetStoryRanges2](#GetStoryRanges2) | Gets an object for enumerating the stories in a document. |
| [GetTypographyOptions](#GetTypographyOptions) | Gets the typography options. |
| [GetVersion](#GetVersion) | Gets the version number of the Text Object Model (TOM) engine. |
| [GetWindow](#GetWindow) | Gets the handle of the window that the Text Object Model (TOM) engine is using to display output. |
| [AttachMsgFilter](#AttachMsgFilter) | Attaches a new message filter to the edit instance. All window messages that the edit instance receives are forwarded to the message filter. |
| [CheckTextLimit](#CheckTextLimit) | Checks whether the number of characters to be added would exceed the maximum text limit. |
| [GetCallManager](#GetCallManager) | Gets the call manager. |
| [GetClientRect](#GetClientRect) | Retrieves the client rectangle of the rich edit control. |
| [GetEffectColor](#GetEffectColor) | Retrieves the color used for special text attributes. |
| [GetImmContext](#GetImmContext) | Gets the Input Method Manager (IMM) input context from the Text Object Model (TOM) host. |
| [GetPreferredFont](#GetPreferredFont) | Retrieves the preferred font for a particular character repertoire and character position. |
| [GetProperty](#GetProperty) | Retrieves the value of a property. |
| [GetStrings](#GetStrings) | Gets a collection of rich-text strings. |
| [Notify](#Notify) | Notifies the Text Object Model (TOM) engine client of particular Input Method Editor (IME) events. |
| [Range2](#Range2) | Retrieves a new text range for the active story of the document. |
| [RangeFromPoint2](#RangeFromPoint2) | Retrieves the degenerate range at (or nearest to) a particular point on the screen. |
| [ReleaseCallManager](#ReleaseCallManager) | Releases the call manager. |
| [ReleaseImmContext](#ReleaseImmContext) | Releases an Input Method Manager (IMM) input context. |
| [SetEffectColor](#SetEffectColor) | Specifies the color to use for special text attributes. |
| [SetProperty](#SetProperty) | Specifies a new value for a property. |
| [SetTypographyOptions](#SetTypographyOptions) | Specifies the typography options for the document. |
| [SysBeep](#SysBeep) | Generates a system beep. |
| [Update](#Update) | Updates the selection and caret. |
| [UpdateWindow](#UpdateWindow) | Notifies the client that the view has changed and the client should update the view if the Text Object Model (TOM) engine is in-place active. |
| [GetMathProperties](#GetMathProperties) | Gets the math properties for the document. |
| [SetMathProperties](#SetMathProperties) | Specifies the math properties to use for the document. |
| [GetActiveStory](#GetActiveStory) | Gets the active story; that is, the story that receives keyboard and mouse input. |
| [SetActiveStory](#SetActiveStory) | Sets the active story; that is, the story that receives keyboard and mouse input. |
| [GetMainStory](#GetMainStory) | Gets the main story. |
| [GetNewStory](#GetNewStory) | Gets a new story. Not implemented. |
| [GetStory](#GetStory) | Retrieves the story that corresponds to a particular index. |

# <a name="GetName"></a>GetName

Gets the file name of this document. This is the **ITextDocument** default property.

```
FUNCTION CTextDocument2.GetName () AS CBSTR
   DIM pName AS AFX_BSTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetName(m_pTextDocument2, @pName))
   RETURN pName
END FUNCTION
```
#### Return value

The filename of this document, or an empty setring if there is not a filename associated with this object.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **S_FALSE** | No file name associated with this object. |
| **E_INVALIDARG** | Invalid argument. |
| **E_OUTOFMEMORY** | Insufficient memory for output string. |

# <a name="GetSelection"></a>GetSelection

Gets the active selection.

```
FUNCTION CTextDocument2.GetSelection () AS ITextSelection PTR
   DIM pSelection AS ITextSelection PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetSelection(m_pTextDocument2, @pSelection))
   RETURN pSelection
END FUNCTION
```
#### Return value

The **ITextSelection** pointer of the active selection.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **S_FALSE** | Indicates no active selection. |
| **E_INVALIDARG** | Invalid argument. |

# <a name="GetStoryCount"></a>GetStoryCount

Gets the count of stories in this document.

```
FUNCTION CTextDocument2.GetStoryCount () AS LONG
   DIM nCount AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetStoryCount(m_pTextDocument2, @nCount))
   RETURN nCount
END FUNCTION
```
#### Return value

The number of stories in the document.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |

# <a name="GetStoryRanges"></a>GetStoryRanges

Gets the story collection object used to enumerate the stories in a document.

```
FUNCTION CTextDocument2.GetStoryRanges () AS ITextStoryRanges PTR
   DIM pStoryRanges AS ITextStoryRanges PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetStoryRanges(m_pTextDocument2, @pStoryRanges))
   RETURN pStoryRanges
END FUNCTION
```
#### Return value

A pointer to the **ITextStoryRanges** interface.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_NOTIMPL** | Not implemented; only one story in this document. |

#### Remarks

Invoke this method only if **GetStoryCount** returns a value greater than 1.

# <a name="GetSaved"></a>GetSaved

Gets a value that indicates whether changes have been made since the file was last saved.

```
FUNCTION CTextDocument2.GetSaved () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetSaved(m_pTextDocument2, @Value))
   RETURN Value
END FUNCTION
```
#### Return value

The value **tomTrue** if no changes have been made since the file was last saved, or the value tomFalse if there are unsaved changes.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |

# <a name="SetSaved"></a>SetSaved

Sets the document **Saved** property.

```
FUNCTION CTextDocument2.SetSaved (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetSaved(m_pTextDocument2, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | New value of the **Saved** property. It can be one of the following values:<br>**tomTrue**. No changes to the file since the last time it was saved.<br>**tomFalse**. There are changes to the file. |

#### Return value

The return value is S_OK.

# <a name="GetDefaultTabStop"></a>GetDefaultTabStop

Gets the default tab width.

```
FUNCTION CTextDocument2.GetDefaultTabStop () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextDocument2->lpvtbl->GetDefaultTabStop(m_pTextDocument2, @Value))
   RETURN Value
END FUNCTION
```

#### Return value

The default tab width.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |

# <a name="SetDefaultTabStop"></a>SetDefaultTabStop

Sets the default tab stop, which is used when no tab exists beyond the current display position.

```
FUNCTION CTextDocument2.SetDefaultTabStop (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetDefaultTabStop(m_pTextDocument2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | New default tab setting, in floating-point points. Default value is 36.0 points, that is, 0.5 inches. |

If the method succeeds it returns S_OK. If the method fails, it returns one of the following COM error codes. For more information on COM error codes, see Error Handling in COM.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="New_"></a>New_

Opens a new document.

```
FUNCTION CTextDocument2.New_ () AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->New_(m_pTextDocument2))
   FUNCTION = m_Result
END FUNCTION
```

#### Return value

If the method succeeds, it returns S_OK.

#### Remarks

If another document is open, this method saves any current changes and closes the current document before opening a new one.

# <a name="Open"></a>Open

Opens a specified document. There are parameters to specify access and sharing privileges, creation and conversion of the file, as well as the code page for the file.

```
FUNCTION CTextDocument2.Open (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->Open(m_pTextDocument2, pVar, Flags, CodePage))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pVar* | A **VARIANT** that specifies the name of the file to open. |
| *Flags* | The file creation, open, share, and conversion flags. Default value is zero, which gives read/write access and read/write sharing, open always, and automatic recognition of the file format (unrecognized file formats are treated as text). Other values are defined in the following groups (see table below). |
| *CodePage* | The code page to use for the file. Zero (the default value) means **CP_ACP** (ANSI code page) unless the file begins with a Unicode BOM 0xfeff, in which case the file is considered to be Unicode. Note that code page 1200 is Unicode, **CP_UTF8** is UTF-8. |

Any combination of these values may be used:

| Constant | Value | Description |
| -------- | ----- | ----------- |
| **tomReadOnly** | &h100 | Read only. |
| **tomShareDenyRead** | &h200 | Other programs cannot read. |
| **tomShareDenyWrite** | &h400 | Other programs cannot write. |
| **tomPasteFile** | &h1000 | Replace the selection with a file. |

These values are mutually exclusive:

| Constant | Value | Description |
| ---- | ----- | ----------- |
| **tomCreateNew** | &h10 | Create a new file. Fail if the file already exists. |
| **tomCreateAlways** | &h20 | Create a new file. Destroy the existing file if it exists. |
| **tomOpenExisting** | &h30 | Open an existing file. Fail if the file does not exist. |
| **tomOpenAlways** | &h40 | Open an existing file. Create a new file if the file does not exist. |
| **tomTruncateExisting** | &h50 | Open an existing file, but truncate it to zero length. |
| **tomRTF** | &h1 | Open as RTF. |
| **tomText** | &h2 | Open as text ANSI or Unicode. |
| **tomHTML** | &h3 | Open as HTML. |
| **tomWordDocument* | &h4 | Open as Word document. |

#### Return value

The return value can be an **HRESULT** value that corresponds to a system error or COM error code, including one of the following values.

| Result code | Description |
| ----------- | ----------- |
| **S_OK** | Method succeeds. |
| **E_INVALIDARG** | Invalid argument. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **E_NOTIMPL** | Feature not implemented. |

# <a name="Save"></a>Save

Saves the document.

```
FUNCTION CTextDocument2.Save (BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->Save(m_pTextDocument2, pVar, Flags, CodePage))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pVar* | The save target. This parameter is a **VARIANT**, which can be a file name, or NULL. |
| *Flags* | File creation, open, share, and conversion flags. For a list of possible values, see table below. |
| *CodePage* | The specified code page. Common values are **CP_ACP** (zero: system ANSI code page), 1200 (Unicode), and 1208 (UTF-8). |

Any combination of these values may be used:

| Constant | Value | Description |
| -------- | ----- | ----------- |
| **tomReadOnly** | &h100 | Read only. |
| **tomShareDenyRead** | &h200 | Other programs cannot read. |
| **tomShareDenyWrite** | &h400 | Other programs cannot write. |
| **tomPasteFile** | &h1000 | Replace the selection with a file. |

These values are mutually exclusive:

| Constant | Value | Description |
| -------- | ----- | ----------- |
| **tomCreateNew** | &h10 | Create a new file. Fail if the file already exists. |
| **tomCreateAlways** | &h20 | Create a new file. Destroy the existing file if it exists. |
| **tomOpenExisting** | &h30 | Open an existing file. Fail if the file does not exist. |
| **tomOpenAlways** | &h40 | Open an existing file. Create a new file if the file does not exist. |
| **tomTruncateExisting** | &h50 | Open an existing file, but truncate it to zero length. |
| **tomRTF** | &h1 | Open as RTF. |
| **tomText** | &h2 | Open as text ANSI or Unicode. |
| **tomHTML** | &h3 | Open as HTML. |
| **tomWordDocument* | &h4 | Open as Word document. |

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

If *pVar* specifies a file name, that name should replace the current Name property. Similarly, the *Flags* and *CodePage* arguments can overrule those supplied in the **Open** method and define the values to use for files created with the **New_** method.

Unicode plain-text files should be saved with the Unicode byte-order mark (0xFEFF) as the first character. This character should be removed when the file is read in; that is, it is only used for import/export to identify the plain text as Unicode and to identify the byte order of that text. Microsoft Notepad adopted this convention, which is now recommended by the Unicode standard.

# <a name="Freeze"></a>Freeze

Increments the freeze count.

```
FUNCTION CTextDocument2.Freeze () AS LONG
   DIM nCount AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->Freeze(m_pTextDocument2, @nCount))
   RETURN nCount
END FUNCTION
```

#### Return value

The updated freeze count.

#### Result code

If the count is nonzero, **GetLastResult** returns **S_OK**. If the count is zero, it returns **FALSE**.

#### Remarks

If the freeze count is nonzero, screen updating is disabled. This allows a sequence of editing operations to be performed without the performance loss and flicker of screen updating. To decrement the freeze count, call the **Unfreeze** method.

# <a name="Unfreeze"></a>Unfreeze

Decrements the freeze count.

```
FUNCTION CTextDocument2.Unfreeze () AS LONG
   DIM nCount AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->Unfreeze(m_pTextDocument2, @nCount))
   RETURN nCount
END FUNCTION
```

#### Return value

The updated freeze count.

#### Result code

If the freeze count is zero, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**, indicating that the freeze count is nonzero.

#### Remarks

If the freeze count goes to zero, screen updating is enabled. This method cannot decrement the count below zero, and no error occurs if it is executed with a zero freeze count.

Note, if edit collection is active, screen updating is suppressed, even if the freeze count is zero.

# <a name="BeginEditCollection"></a>BeginEditCollection

Turns on edit collection (also called *undo grouping*).

```
FUNCTION CTextDocument2.BeginEditCollection () AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->BeginEditCollection(m_pTextDocument2))
   RETURN m_Result
END FUNCTION
```

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| **S_OK** | Method succeeds. |
| **S_FALSE** | Undo is not enabled. |
| **E_NOTIMPL** | Feature not implemented. |

#### Remarks

A single **Undo** command undoes all changes made while edit collection is turned on.

# <a name="EndEditCollection"></a>EndEditCollection

Turns off edit collection (also called *undo grouping*).

```
FUNCTION CTextDocument2.EndEditCollection () AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->EndEditCollection(m_pTextDocument2))
   RETURN m_Result
END FUNCTION
```

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns a COM error code.

| Return value | Description |
| ------------ | ----------- |
| **S_OK** | Method succeeds. |
| **E_NOTIMPL** | Feature not implemented. |

#### Remarks

The screen is unfrozen unless the freeze count is nonzero.

# <a name="Undo"></a>Undo

Performs a specified number of undo operations.

```
FUNCTION CTextDocument2.Undo (BYVAL Count AS LONG) AS LONG
   DIM nCount AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->Undo(m_pTextDocument2, Count, @nCount))
   RETURN nCount
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Count* | The specified number of undo operations. If the value of this parameter is **tomFalse**, undo processing is suspended. If this parameter is **tomTrue**, undo processing is restored. |

#### Return value

The actual count of undo operations performed.

#### Result code

If all of the *Count* undo operations were performed, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**, indicating that less than *Count* undo operations were performed.

# <a name="Redo"></a>Redo

Performs a specified number of redo operations.

```
FUNCTION CTextDocument2.Redo (BYVAL Count AS LONG) AS LONG
   DIM nCount AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->Redo(m_pTextDocument2, Count, @nCount))
   RETURN nCount
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Count* | The number of redo operations specified. |

#### Return value

The actual count of redo operations performed.

#### Result code

If the method succeeds **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **S_OK** | Method succeeds. |
| **S_FALSE** | Less than Count redo operations were performed. |

# <a name="Range"></a>Range

Retrieves a text range object for a specified range of content in the active story of the document.

```
FUNCTION CTextDocument2.Range (BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG) AS ITextRange PTR
   DIM pTextRange AS ITextRange PTR
   this.SetResult(m_pTextDocument2->lpvtbl->Range(m_pTextDocument2, cpActive, cpAnchor, @pTextRange))
   RETURN pTextRange
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *cpActive* | The start position of new range. The default value is zero, which represents the start of the document. |
| *cpAnchor* | The end position of new range. The default value is zero. |

### Return value

Pointer to a **ITextRange** interface to the specified text range.

# <a name="RangeFromPoint"></a>RangeFromPoint

Retrieves a range for the content at or nearest to the specified point on the screen.

```
UNCTION CTextDocument2.RangeFromPoint (BYVAL x AS LONG, BYVAL y AS LONG) AS ITextRange PTR
   DIM pTextRange AS ITextRange PTR
   this.SetResult(m_pTextDocument2->lpvtbl->RangeFromPoint(m_pTextDocument2, x, y, @pTextRange))
   RETURN pTextRange
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *x* | The horizontal coordinate of the specified point, in screen coordinates. |
| *y* | The vertical coordinate of the specified point, in screen coordinates. |

### Return value

Pointer to a **ITextRange** interface that corresponds to the specified point.

#### Result code

If the method succeeds **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **S_OK** | Method succeeds. |
| **E_INVALIDARG** | Invalid argument. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetCaretType"></a>GetCaretType

Gets the caret type.

```
FUNCTION CTextDocument2.GetCaretType () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetCaretType(m_pTextDocument2, @Value))
   RETURN Value
END FUNCTION
```

### Return value

The caret type. It can be one of the following values:

| Constant | Value | Description |
| -------- | ----- | ----------- |
| **tomKoreanBlockCaret** | &h1 | The Korean block caret. |
| **tomNormalCaret** | 0 | Normal caret. |
| **tomNullCaret** | &h2 | NULL caret (caret suppressed) |

#### Result code

If the method succeeds **GetLastResult** returns **S_OK**. If the method fails, it returns an **HRESULT** error code.

# <a name="SetCaretType"></a>SetCaretType

Sets the caret type.

```
FUNCTION CTextDocument2.SetCaretType (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetCaretType(m_pTextDocument2, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new caret type. It can be one of the following values: |

| Constant | Value | Description |
| -------- | ----- | ----------- |
| **tomKoreanBlockCaret** | &h1 | The Korean block caret. |
| **tomNormalCaret** | 0 | Normal caret. |
| **tomNullCaret** | &h2 | NULL caret (caret suppressed) |

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetDisplays"></a>GetDisplays

Gets the displays collection for this Text Object Model (TOM) engine instance.

```
FUNCTION CTextDocument2.GetDisplays () AS ITextDisplays PTR
   DIM pTextDisplays AS ITextDisplays PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetDisplays(m_pTextDocument2, @pTextDisplays))
   RETURN pTextDisplays
END FUNCTION
```

### Return value

A pointer to the **ITextDisplays** interface.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The rich edit control doesn't implement this method.

# <a name="GetDocumentFont"></a>GetDocumentFont

Gets an object that provides the default character format information for this instance of the Text Object Model (TOM) engine.

```
FUNCTION CTextDocument2.GetDocumentFont () AS ITextFont2 PTR
   DIM pITextFont2 AS ITextFont2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetDocumentFont(m_pTextDocument2, @pITextFont2))
   RETURN pITextFont2
END FUNCTION
```

### Return value

A pointer to the ITextFont2 interface.

### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.
