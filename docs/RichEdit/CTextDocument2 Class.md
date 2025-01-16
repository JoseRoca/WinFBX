# CTextDocument2 Class

Class that wraps all the methods of the **ITextDocument** and **ITextDocument2** interfaces.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTORS](#CONSTRUCTORS) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [LET](#LET) | Assignment operator. |
| [CAST](#CAST) | Cast operator. |
| [TextDocumentPtr](#TextDocumentPtr) | Returns a pointer to the underlying **ITextDocument2** interface. |
| [Attach](#Attach) | Attaches an **ITextDocument2** interface pointer to the class. |
| [Detach](#Detach) | Detaches the underlying **ITextDocument2** interface pointer from the class. |

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

The **ITextDocument** interface inherits from the IDispatch interface. **ITextDocument** also has these types of members:

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

### Methods inherited from CTOMBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#GetLastResult) | Returns the last result code |
| [SetResult](#SetResult) | Sets the last result code. |
| [GetErrorInfo](#GetErrorInfo) | Returns a description of the last result code. |

# <a name="CONSTRUCTORS"></a>CONSTRUCTORS

Called when a **CTextDocument2** class variable is created.

```
DECLARE CONSTRUCTOR
DECLARE CONSTRUCTOR (BYVAL hRichEdit AS HWND)
DECLARE CONSTRUCTOR (BYVAL pTextDocument2 AS ITextDocument2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

## CONSTRUCTOR (Empty)

Can be used, for example, when we have an **ITextDocument2** interface pointer returned by a function and we want to attach it to a new instance of the **CTextDocument2** class.

```
DIM DIM pCTextDocument2 AS CTextDocument2
pCTextDocument2.Attach(pTextDocument2)
```

## CONSTRUCTOR (hRichEdit)

| Parameter | Description |
| --------- | ----------- |
| *hRichEdit* | Handle of the Rich Edit control. |

```
CONSTRUCTOR CTextDocument2 (BYVAL hRichEdit AS HWND)
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
END CONSTRUCTOR
```

## CONSTRUCTOR (ITextDocument2 PTR)

```
CONSTRUCTOR CTextDocument2 (BYVAL pTextDocument2 AS ITextDocument2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   IF fAddRef THEN pTextDocument2->lpvtbl->AddRef(pTextDocument2)
   m_pTextDocument2 = pTextDocument2
END CONSTRUCTOR
```

| Parameter | Description |
| --------- | ----------- |
| *pTextDocument2* | An **ITextDocument2** interface pointer. |
| *fAddRef* | Optional. TRUE to increment the reference count of the passed **ITextDocument2** interface pointer; otherise, FALSE. Default is FALSE. |

#### Return value

A pointer to the new instance of the class.

#### Usage examples

To use with the dotted syntax.

```
SCOPE
   ' // Create a new instance of the CTextDocument2 class
   DIM pTextDocument2 AS CTextDocument2 = hRichEdit
   ' // Get the number of characters of the text in the Rich Edit control
   DIM numChars AS LONG = RichEdit_GetTextLength(hRichEdit)
   ' // Get the 0-based range of all the text
   DIM pCRange2 AS CTextRange2 = pCTextDoc.Range2(0, numChars)
   ' // Get the text
   DIM cbsText AS CBSTR = pCRange2.GetText2(0)
   ' // The CTextDocument2 class and the CTextRange2 class will be destroyed when the scope ends
END SCOPE
```

To use with the pointer syntax.

```
' // Create a new instance of the CTextDocument2 class
DIM pCTextDocument2 AS CTextDocument2 PTR = NEW CTextDocument2(hRichEdit)
' // Get the number of characters of the text in the Rich Edit control
DIM numChars AS LONG = RichEdit_GetTextLength(hRichEdit)
' // Get the 0-based range of all the text
DIM pCRange2 AS CTextRange2 = pCTextDoc->Range2(0, numChars)
' // Get the text
DIM cbsText AS CBSTR = pCRange2->GetText2(0)
' // Delete the range
Delete pCRange2
' // Delete the class
Delete pCTextDocument2
```

# <a name="DESTRUCTOR"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

```
DESTRUCTOR CTextDocument2
   ' // Release the ITextDocument2 interface
   IF m_pTextDocument2 THEN m_pTextDocument2->lpvtbl->Release(m_pTextDocument2)
END DESTRUCTOR
```

# <a name="LET"></a>LET

Assignment operator.

```
OPERATOR CTextDocument2.LET (BYVAL pTextDocument2 AS ITextDocument2 PTR)
   m_Result = 0
   IF pTextDocument2 = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextDocument2 THEN m_pTextDocument2->lpvtbl->Release(m_pTextDocument2)
   ' // Attach the passed interface pointer to the class
   m_pTextDocument2 = pTextDocument2
END OPERATOR
```

# <a name="CAST"></a>CAST

Cast operator.

```
OPERATOR CTextDocument2.CAST () AS ITextDocument2 PTR
   m_Result = 0
   OPERATOR = m_pTextDocument2
END OPERATOR
```

# <a name="TextDocumentPtr"></a>TextDocumentPtr

Returns a pointer to the underlying **ITextDocument2** interface

```
FUNCTION CTextDocument2.TextDocumentPtr () AS ITextDocument2 PTR
   m_Result = 0
   RETURN m_pTextDocument2
END FUNCTION
```

# <a name="Attach"></a>Attach

Attaches an **ITextDocument2** interface pointer to the class.

```
FUNCTION CTextDocument2.Attach (BYVAL pTextDocument2 AS ITextDocument2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextDocument = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextDocument2 THEN m_Result = m_pTextDocument2->lpvtbl->Release(m_pTextDocument2)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextDocument2->lpvtbl->AddRef(pTextDocument2)
   m_pTextDocument2 = pTextDocument2
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pTextDocument2* | The **ITextDocument2** interface pointer to attach. |
| *fAddRef* | **TRUE** to increment the reference count of te object. Default is FALSE. |

# <a name="Detach"></a>Detach

Detaches the underlying **ITextDocument2** interface pointer from the class

```
FUNCTION CTextDocument2.Detach () AS ITextDocument2 PTR
   m_Result = 0
   DIM pTextDocument2 AS ITextDocument2 PTR = m_pTextDocument2
   m_pTextDocument2 = NULL
   RETURN pTextDOcument2
END FUNCTION
```

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

A pointer to the **ITextFont2** interface.

### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetDocumentFont"></a>SetDocumentFont

Sets the default character formatting for this instance of the Text Object Model (TOM) engine.

```
FUNCTION CTextDocument2.SetDocumentFont (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetDocumentFont(m_pTextDocument2, pFont))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pFont* | The font object that provides the default character formatting. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

You can also set the default character formatting by calling the **Reset** method of the **ITextFont** interface with a value of **tomDefault**.

# <a name="GetDocumentPara"></a>GetDocumentPara

Gets an object that provides the default paragraph format information for this instance of the Text Object Model (TOM) engine.

```
FUNCTION CTextDocument2.GetDocumentPara () AS ITextPara2 PTR
   DIM pITextPara2 AS ITextPara2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetDocumentPara(m_pTextDocument2, @pITextPara2))
   RETURN pITextPara2
END FUNCTION
```

#### Return value

A pointer to the **ITextPara2** interface.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetDocumentPara"></a>SetDocumentPara

Sets the default paragraph formatting for this instance of the Text Object Model (TOM) engine.

```
FUNCTION CTextDocument2.SetDocumentPara (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetDocumentPara(m_pTextDocument2, pPara))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pPara* | The paragraph object that provides the default paragraph formatting. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

You can also set the default character formatting by calling the **Reset** method of the **ITextFont** interface with a value of **tomDefault**.

# <a name="GetEastAsianFlags"></a>GetEastAsianFlags

Gets an object that provides the default paragraph format information for this instance of the Text Object Model (TOM) engine.

```
FUNCTION CTextDocument2.GetEastAsianFlags () AS LONG
   DIM pFlags AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetEastAsianFlags(m_pTextDocument2, @pFlags))
   RETURN pFlags
END FUNCTION
```

#### Return value

The East Asian flags. This parameter can be a combination of the following values.

| Value | Meaning |
| ----- | ------- |
| *tomRE10Mode* | TOM version 1.0 emulation mode. |
| *tomUseAtFont* | Use @ fonts for CJK vertical text. |
| *tomTextFlowMask* | A mask for the following four text orientations. |
| *tomTextFlowES* | Ordinary left-to-right horizontal text. |
| *tomTextFlowSW* | Ordinary East Asian vertical text. |
| *tomTextFlowWN* | An alternative orientation. |
| *tomTextFlowNE* | An alternative orientation. |
| *tomUsePassword* | Use password control. |
| *tomNoIME* | Turn off IME operation (see **ES_NOIME**). |
| *tomSelfIME* | The rich edit host handles IME operation (see **ES_SELFIME**) . |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetGenerator"></a>GetGenerator

Gets the name of the Text Object Model (TOM) engine.

```
FUNCTION CTextDocument2.GetGenerator () AS CBSTR
   DIM bstr AS AFX_BSTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetGenerator(m_pTextDocument2, @bstr))
   RETURN bstr
END FUNCTION
```

#### Return value

The name of the TOM engine.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetIMEInProgress"></a>SetIMEInProgress

Sets the state of the Input Method Editor (IME) in-progress flag.

```
FUNCTION CTextDocument2.SetIMEInProgress (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetIMEInProgress(m_pTextDocument2, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | Use **tomTrue** to turn on the IME in-progress flag, or **tomFalse** to turn it off. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetNotificationMode"></a>GetNotificationMode

Gets the notification mode.

```
FUNCTION CTextDocument2.GetNotificationMode () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetNotificationMode(m_pTextDocument2, @Value))
   RETURN Value
END FUNCTION
```

#### Return value

The notification mode. This parameter is set to **tomTrue** if notifications are active, or **tomFalse** if not.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetNotificationMode"></a>SetNotificationMode

Sets the notification mode.

```
FUNCTION CTextDocument2.SetNotificationMode (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetNotificationMode(m_pTextDocument2, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The notification mode. Use **tomTrue** to turn on notifications, or **tomFalse** to turn them off. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetSelection2"></a>GetSelection2

Gets the active selection.

```
FUNCTION CTextDocument2.GetSelection2 () AS ITextSelection2 PTR
   DIM ppSel AS ITextSelection2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetSelection2(m_pTextDocument2, @ppSel))
   RETURN ppSel
END FUNCTION
```

#### Return value

A pointer to the **ITextSelection2** interface of the selection. This pointer is **NULL** if the rich edit control is not in-place active.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetStoryRanges2"></a>GetStoryRanges2

Gets an object for enumerating the stories in a document.

```
FUNCTION CTextDocument2.GetStoryRanges2 () AS ITextStoryRanges2 PTR
   DIM pStories AS ITextStoryRanges2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetStoryRanges2(m_pTextDocument2, @pStories))
   RETURN pStories
END FUNCTION
```

#### Return value

A pointer to the **ITextStoryRanges2** interface used for enumerating stories.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


#### Remarks

Call this method only if the **GetStoryCount** method returns a value that is greater than one.

# <a name="GetTypographyOptions"></a>GetTypographyOptions

Gets the typography options.

```
FUNCTION CTextDocument2.GetTypographyOptions () AS LONG
   DIM Options AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetTypographyOptions(m_pTextDocument2, @Options))
   RETURN Options
END FUNCTION
```

#### Return value

A combination of the following typography options.

| Value | Meaning |
| ----- | ------- |
| **TO_ADVANCEDTYPOGRAPHY** | Advanced typography (special line breaking and line formatting) is turned on. |
| **TO_SIMPLELINEBREAK** | Normal line breaking and formatting is used. |

#### Return value

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetVersion"></a>GetVersion

Gets the version number of the Text Object Model (TOM) engine.

```
FUNCTION CTextDocument2.GetVersion () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetVersion(m_pTextDocument2, @Value))
   RETURN Value
END FUNCTION
```

#### Return value

The version number. Byte 3 gives the major version number, byte 2 the minor version number, and the low-order 16 bits give the build number.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetWindow"></a>GetWindow

Gets the handle of the window that the Text Object Model (TOM) engine is using to display output.

```
FUNCTION CTextDocument2.GetWindow () AS __int64
   DIM pHwnd AS __int64
   this.SetResult(m_pTextDocument2->lpvtbl->GetWindow(m_pTextDocument2, @pHwnd))
   RETURN pHwnd
END FUNCTION
```

#### Return value

The handle of the window that the TOM engine is using.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

A rich edit control doesn't need to own the window that the TOM engine is using. For example, the rich edit control might be windowless.

The Input Method Editor (IME) needs the handle of the window that is receiving keyboard messages. This method retrieves that handle.

# <a name="AttachMsgFilter"></a>AttachMsgFilter

Attaches a new message filter to the edit instance. All window messages that the edit instance receives are forwarded to the message filter.

```
FUNCTION CTextDocument2.AttachMsgFilter (BYVAL pFilter AS IUnknown PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->AttachMsgFilter(m_pTextDocument2, pFilter))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pFilter* | The message filter. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The message filter must be bound to the document before it can be used.

# <a name="CheckTextLimit"></a>CheckTextLimit

Checks whether the number of characters to be added would exceed the maximum text limit.

```
FUNCTION CTextDocument2.CheckTextLimit (BYVAL cch AS LONG) AS LONG
   DIM pcch AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->CheckTextLimit(m_pTextDocument2, cch, @pcch))
   RETURN pcch
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *cch* | The number of characters to be added. |

#### Return value

The number of characters that exceed the maximum text limit. This parameter is 0 if the number of characters does not exceed the limit.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetCallManager"></a>GetCallManager

Gets the call manager.

```
FUNCTION CTextDocument2.GetCallManager () AS IUnknown PTR
   DIM pVoid AS IUnknown PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetCallManager(m_pTextDocument2, @pVoid))
   RETURN pVoid
END FUNCTION
```

#### Return value

The call manager object.

#### Result code

If the method succeeds,**GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The call manager object is opaque to the caller. The Text Object Model (TOM) engine uses the object to handle internal notifications for particular scenarios.

# <a name="GetClientRect"></a>GetClientRect

Retrieves the client rectangle of the rich edit control.

```
FUNCTION CTextDocument2.GetClientRect (BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, _
BYVAL pTop AS LONG PTR, BYVAL pRight AS LONG PTR, BYVAL pBottom AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->GetClientRect(m_pTextDocument2, nType, pLeft, pTop, pRight, pBottom))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *nType* | The client rectangle retrieval options. It can be a combination of the following values.<br>**tomClientCoord**. Retrieve the rectangle in client coordinates. If this value isn't specified, the function retrieves screen coordinates.<br>**tomIncludeInset**. Add left and top insets to the left and top coordinates of the client rectangle, and subtract right and bottom insets from the right and bottom coordinates.<br>**tomTransform**. Use a world transform (XFORM) provided by the host application to transform the retrieved rectangle coordinates. |
| *pLeft* | The x-coordinate of the upper-left corner of the rectangle. |
| *pTop* | The y-coordinate of the upper-left corner of the rectangle. |
| *pRight* | The x-coordinate of the lower-right corner of the rectangle. |
| *pBottom* | The y-coordinate of the lower-right corner of the rectangle. |

#### Return value

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetEffectColor"></a>GetEffectColor

Retrieves the color used for special text attributes.

```
FUNCTION CTextDocument2.GetEffectColor (BYVAL Index AS LONG) AS ULONG
   DIM pcr AS ULONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetEffectColor(m_pTextDocument2, Index, @pcr))
   RETURN pcr
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Index* | The index of the color to retrieve. It can be one of the following values. |

| Index | Meaning |
| --------- | ----------- |
| 0 | Text color. |
| 1 | RGB(0, 0, 0) |
| 2 | RGB(0, 0, 255) |
| 3 | RGB(0, 255, 255) |
| 4 | RGB(0, 255, 0) |
| 5 | RGB(255, 0, 255) |
| 6 | RGB(255, 0, 0) |
| 7 | RGB(255, 255, 0) |
| 8 | RGB(255, 255, 255) |
| 9 | RGB(0, 0, 128) |
| 10 | RGB(0, 128, 128) |
| 11 | RGB(0, 128, 0) |
| 12 | RGB(128, 0, 128) |
| 13 | RGB(128, 0, 0) |
| 14 | RGB(128, 128, 0) |
| 15 | RGB(128, 128, 128) |
| 16 | RGB(192, 192, 192) |

#### Return value

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

### Remarks

The first 16 index values are for special underline colors. If an index between 1 and 16 hasn't been defined by a call to the **etEffectColor** method, **GetEffectColor** returns the corresponding Microsoft Word default color.

# <a name="GetPreferredFont"></a>GetPreferredFont

Retrieves the preferred font for a particular character repertoire and character position.

```
FUNCTION CTextDocument2.GetPreferredFont (BYVAL cp AS LONG, BYVAL CodePage AS LONG, _
BYVAL Options AS LONG, BYVAL curCodepage AS LONG, BYVAL curFontSize AS LONG, BYVAL pbstr AS BSTR PTR, _
BYVAL pPitchAndFamily AS LONG PTR, BYVAL pNewFontSize AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->GetPreferredFont(m_pTextDocument2, cp, CodePage, _
   Options, curCodePage, curFontSize, pbstr, pPitchAndFamily, pNewFontSize))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *cp* | The character position for the preferred font. |
| *CharRep* | The character repertoire index for the preferred font. It can be one of the following values: **tomAboriginal**, **tomAnsi**, **tomArabic**, **tomArmenian**, **tomBaltic**, **tomBengali**, **tomBIG5**, **tomBraille**, **tomCherokee**, **tomCyrillic**, **tomDefaultCharRep**, **tomDevanagari**, **tomEastEurope**, **tomEmoji**, **tomEthiopic**, **tomGB2312**, **tomGeorgian**, **tomGreek**, **tomGujarati**, **tomGurmukhi**, **tomHangul**, **tomHebrew**, **tomJamo**, **tomKannada**, **tomKayahli**, **tomKharoshthi**, **tomKhmer**, **tomLao**, **tomLimbu**, **tomMac**, **tomMalayalam**, **tomMongolian**, **tomMyanmar**, **tomNewTaiLu**, **tomOEM**, **tomOgham**, **tomOriya**, **tomPC437**, **tomRunic**, **tomShiftJIS**, **tomSinhala**, **tomSylotinagr**, **tomSymbol**, **tomSyriac**, **tomTaiLe**, **tomTamil**, **tomTelugu**, **tomThaana**, **tomThai**, **tomTibetan**, **tomTurkish**, **tomUsymbol**, **tomVietnamese**, **tomYi**. |
| *Options* | The preferred font options. The low-order word can be a combination of the following values.<br>**tomIgnoreCurrentFont**, **tomMatchCharRep**, **tomMatchFontSignature**, **tomMatchAscii**, **tomGetHeightOnly**, **tomMatchMathFont**. If the high-order word of Options is tomUseTwips, the font heights are given in twips. |
| *curCharRep* | The index of the current character repertoire. |
| *curFontSize* | The current font size. |
| *pPitchAndFamily* | The font pitch and family. |
| *pNewFontSize* | The new font size. |

#### Return value

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetImmContext"></a>GetImmContext

Gets the Input Method Manager (IMM) input context from the Text Object Model (TOM) host.

```
FUNCTION CTextDocument2.GetImmContext () AS __int64
   DIM Context AS __int64
   this.SetResult(m_pTextDocument2->lpvtbl->GetImmContext(m_pTextDocument2, @Context))
   RETURN Context
END FUNCTION
```

#### Return value

The IMM input context.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetProperty"></a>GetProperty

Retrieves the value of a property.

```
FUNCTION CTextDocument2.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetProperty(m_pTextDocument2, nType, @Value))
   RETURN Value
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *nType* | The identifier of the property to retrieve. It can be one of the following property IDs: **tomCanCopy**, **tomCanRedo**, **tomCanUndo**, **tomDocMathBuild**, **tomMathInterSpace**, **tomMathIntraSpace**, **tomMathLMargin**, **tomMathPostSpace**, **tomMathPreSpace**, **tomMathRMargin**, **tomMathWrapIndent**, **tomMathWrapRight**, **tomUndoLimit**, **tomEllipsisMode**, **tomEllipsisState** |

#### Return value

The value of the property.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetStrings"></a>GetStrings

Gets a collection of rich-text strings.

```
FUNCTION CTextDocument2.GetStrings () AS ITextStrings PTR
   DIM pStrs AS ITextStrings PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetStrings(m_pTextDocument2, @pStrs))
   RETURN pStrs
END FUNCTION
```

#### Return value

A pointer to the **ITextStrings** interface of the collection of rich-text strings.

# <a name="Notify"></a>Notify

Notifies the Text Object Model (TOM) engine client of particular Input Method Editor (IME) events.

```
FUNCTION CTextDocument2.Notify (BYVAL nNotify AS LONG) AS HRESULT
   DIM pStrs AS ITextStrings PTR
   this.SetResult(m_pTextDocument2->lpvtbl->Notify(m_pTextDocument2, nNotify))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *nNotify* | An IME notification code. |

#### Rerturn value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="Range2"></a>Range2

Notifies the Text Object Model (TOM) engine client of particular Input Method Editor (IME) events.

```
FUNCTION CTextDocument2.Range2 (BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG) AS ITextRange2 PTR
   DIM pRange2 AS ITextRange2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->Range2(m_pTextDocument2, cpActive, cpAnchor, @pRange2))
   RETURN pRange2
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *cpActive* | The active end of the new text range. The default value is 0; that is, the beginning of the story. |
| *cpAnchor* | The anchor end of the new text range. The default value is 0. |

#### Return value

The new text range.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="RangeFromPoint2"></a>RangeFromPoint2

Retrieves the degenerate range at (or nearest to) a particular point on the screen.

```
FUNCTION CTextDocument2.RangeFromPoint2 (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG) AS ITextRange2 PTR
   DIM pRange2 AS ITextRange2 PTR
   this.SetResult(m_pTextDocument2->lpvtbl->RangeFromPoint2(m_pTextDocument2, x, y, nType, @pRange2))
   RETURN pRange2
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *x* | The x-coordinate of a point, in screen coordinates. |
| *y* | The y-coordinate of a point, in screen coordinates. |
| *nType* | The alignment type of the specified point. For a list of valid values, see the [GetPoint](https://learn.microsoft.com/en-us/windows/win32/api/tom/nf-tom-itextrange-getpoint) method of the **ITextRange** interface. |

#### Result code

If the method succeeds, **GetLastResult^^ returns **NOERROR**. Otherwise, it returns an **HRESUL** error code.

# <a name="ReleaseCallManager"></a>ReleaseCallManager

Releases the call manager.

```
FUNCTION CTextDocument2.ReleaseCallManager (BYVAL pVoid AS IUnknown PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->ReleaseCallManager(m_pTextDocument2, pVoid))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pvoid* | The call manager object to release. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="ReleaseImmContext"></a>ReleaseImmContext

Releases an Input Method Manager (IMM) input context.

```
FUNCTION CTextDocument2.ReleaseImmContext (BYVAL Context AS __int64) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->ReleaseImmContext(m_pTextDocument2, Context))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Context* | The IMM input context to release. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetEffectColor"></a>SetEffectColor

Specifies the color to use for special text attributes.

```
FUNCTION CTextDocument2.SetEffectColor (BYVAL Index AS LONG, BYVAL Value AS ULONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetEffectColor(m_pTextDocument2, Index, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Index* | The index of the color to retrieve. For a list of values, see the table below. |
| *Value* | The new color for the specified index. |

| Index | Meaning |
| --------- | ----------- |
| 0 | Text color. |
| 1 | RGB(0, 0, 0) |
| 2 | RGB(0, 0, 255) |
| 3 | RGB(0, 255, 255) |
| 4 | RGB(0, 255, 0) |
| 5 | RGB(255, 0, 255) |
| 6 | RGB(255, 0, 0) |
| 7 | RGB(255, 255, 0) |
| 8 | RGB(255, 255, 255) |
| 9 | RGB(0, 0, 128) |
| 10 | RGB(0, 128, 128) |
| 11 | RGB(0, 128, 0) |
| 12 | RGB(128, 0, 128) |
| 13 | RGB(128, 0, 0) |
| 14 | RGB(128, 128, 0) |
| 15 | RGB(128, 128, 128) |
| 16 | RGB(192, 192, 192) |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The first 16 index values are for special underline colors. If an index between 1 and 16 hasn't been defined by a call to the **SetEffectColor** method of the **ITextDocument2** interface, the corresponding Microsoft Word default color is used.

# <a name="SetProperty"></a>SetProperty

Specifies a new value for a property.

```
FUNCTION CTextDocument2.SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetProperty(m_pTextDocument2, nType, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *nType* | The identifier of the property. It can be one of the following property IDs: **tomCanCopy**, **tomCanRedo**, **tomCanUndo**, **tomDocMathBuild**, **tomMathInterSpace**, **tomMathIntraSpace**, **tomMathLMargin**, **tomMathPostSpace**, **tomMathPreSpace**, **tomMathRMargin**, **tomMathWrapIndent**, **tomMathWrapRight**, **tomUndoLimit**, **tomEllipsisMode**, **tomEllipsisState**. |
| *Value* | The new property value. |

#### Return value

If the method succeeds, it returns *NOERRO*. Otherwise, it returns an *HRESULT* error code.

# <a name="SetTypographyOptions"></a>SetTypographyOptions

Specifies the typography options for the document.

```
FUNCTION CTextDocument2.SetTypographyOptions (BYVAL Options AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetTypographyOptions(m_pTextDocument2, Options, Mask))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Options* | The typography options to set. For a list of possible options, see the table below. |
| *Mask* | A mask identifying the options to set. For example, to turn on **TO_ADVANCEDTYPOGRAPHY**, call *SetTypographyOptions (TO_ADVANCEDTYPOGRAPHY, TO_ADVANCEDTYPOGRAPHY)*. |

Typography options.

| Value | Meaning |
| ----- | ------- |
| **TO_ADVANCEDTYPOGRAPHY** | Advanced typography (special line breaking and line formatting) is turned on. |
| **TO_SIMPLELINEBREAK** | Normal line breaking and formatting is used. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SysBeep"></a>SysBeep

Generates a system beep.

```
FUNCTION CTextDocument2.SysBeep () AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SysBeep(m_pTextDocument2))
   RETURN m_Result
END FUNCTION
```

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="Update"></a>Update

Updates the selection and caret.

```
FUNCTION CTextDocument2.Update (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->Update(m_pTextDocument2, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | Scroll flag. Use tomTrue to scroll the caret into view. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="UpdateWindow"></a>UpdateWindow

Notifies the client that the view has changed and the client should update the view if the Text Object Model (TOM) engine is in-place active.

```
FUNCTION CTextDocument2.UpdateWindow () AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->UpdateWindow(m_pTextDocument2))
   RETURN m_Result
END FUNCTION
```

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetMathProperties"></a>GetMathProperties

Gets the math properties for the document.

```
FUNCTION CTextDocument2.GetMathProperties () AS LONG
   DIM Options AS LONG
   this.SetResult(m_pTextDocument2->lpvtbl->GetMathProperties(m_pTextDocument2, @Options))
   RETURN Options
END FUNCTION
```

#### Return value

A combination of the following math properties:

| Property | Meaning |
| -------- | ----------- |
| *tomMathDispAlignMask* | Display-mode alignment mask. |
| *tomMathDispAlignCenter* | Center (default) alignment. |
| *tomMathDispAlignLeft* | Left alignment. |
| *tomMathDispAlignRight* | Right alignment. |
| *tomMathDispIntUnderOver* | Display-mode integral limits location. |
| *tomMathDispFracTeX* | Display-mode nested fraction script size. |
| *tomMathDispNaryGrow* | Math-paragraph n-ary grow. |
| *tomMathDocEmptyArgMask* | Empty arguments display mask. |
| *tomMathDocEmptyArgAuto* | Automatically use a dotted square to denote empty arguments, if necessary. |
| *tomMathDocEmptyArgAlways* | Always use a dotted square to denote empty arguments. |
| *tomMathDocEmptyArgNever* | Don't denote empty arguments. |
| *tomMathDocSbSpOpUnchanged* | Display the underscore (_) and caret (^) as themselves. |
| *tomMathDocDiffMask* | Style mask for the **tomMathDocDiffUpright**, **tomMathDocDiffItalic**, **tomMathDocDiffOpenItalic** options. |
| *tomMathDocDiffItalic* | Use italic (default) for math differentials. |
| *tomMathDocDiffUpright* | Use an upright font for math differentials. |
| *tomMathDocDiffOpenItalic* | Use open italic (default) for math differentials. |
| *tomMathDispNarySubSup* | Math-paragraph non-integral n-ary limits location. |
| *tomMathDispDef* | Math-paragraph spacing defaults. |
| *tomMathEnableRtl* | Enable right-to-left (RTL) math zones in RTL paragraphs. |
| *tomMathBrkBinMask* | Equation line break mask. |
| *tomMathBrkBinBefore* | Break before binary/relational operator. |
| *tomMathBrkBinAfter* | Break after binary/relational operator. |
| *tomMathBrkBinDup* | Duplicate binary/relational before/after. |
| *tomMathBrkBinSubMask* | Duplicate mask for minus operator. |
| *tomMathBrkBinSubMM* | - - (minus on both lines). |
| *tomMathBrkBinSubPM* | + - |
| *tomMathBrkBinSubMP* | - + |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetMathProperties"></a>SetMathProperties

Sets the math properties for the document.

```
FUNCTION CTextDocument2.SetMathProperties (BYVAL Options AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetMathProperties(m_pTextDocument2, Options, Mask))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Options* | The math properties to set. For a list of possible properties, see the table below. |
| *Mask* | The math mask. For a list of possible masks, see the table below. |

| Property | Meaning |
| -------- | ----------- |
| *tomMathDispAlignMask* | Display-mode alignment mask. |
| *tomMathDispAlignCenter* | Center (default) alignment. |
| *tomMathDispAlignLeft* | Left alignment. |
| *tomMathDispAlignRight* | Right alignment. |
| *tomMathDispIntUnderOver* | Display-mode integral limits location. |
| *tomMathDispFracTeX* | Display-mode nested fraction script size. |
| *tomMathDispNaryGrow* | Math-paragraph n-ary grow. |
| *tomMathDocEmptyArgMask* | Empty arguments display mask. |
| *tomMathDocEmptyArgAuto* | Automatically use a dotted square to denote empty arguments, if necessary. |
| *tomMathDocEmptyArgAlways* | Always use a dotted square to denote empty arguments. |
| *tomMathDocEmptyArgNever* | Don't denote empty arguments. |
| *tomMathDocSbSpOpUnchanged* | Display the underscore (_) and caret (^) as themselves. |
| *tomMathDocDiffMask* | Style mask for the **tomMathDocDiffUpright**, **tomMathDocDiffItalic**, **tomMathDocDiffOpenItalic** options. |
| *tomMathDocDiffItalic* | Use italic (default) for math differentials. |
| *tomMathDocDiffUpright* | Use an upright font for math differentials. |
| *tomMathDocDiffOpenItalic* | Use open italic (default) for math differentials. |
| *tomMathDispNarySubSup* | Math-paragraph non-integral n-ary limits location. |
| *tomMathDispDef* | Math-paragraph spacing defaults. |
| *tomMathEnableRtl* | Enable right-to-left (RTL) math zones in RTL paragraphs. |
| *tomMathBrkBinMask* | Equation line break mask. |
| *tomMathBrkBinBefore* | Break before binary/relational operator. |
| *tomMathBrkBinAfter* | Break after binary/relational operator. |
| *tomMathBrkBinDup* | Duplicate binary/relational before/after. |
| *tomMathBrkBinSubMask* | Duplicate mask for minus operator. |
| *tomMathBrkBinSubMM* | - - (minus on both lines). |
| *tomMathBrkBinSubPM* | + - |
| *tomMathBrkBinSubMP* | - + |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetActiveStory"></a>GetActiveStory

Gets the active story; that is, the story that receives keyboard and mouse input.

```
FUNCTION CTextDocument2.GetActiveStory () AS ITextStory PTR
   DIM pStory AS ITextStory PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetActiveStory(m_pTextDocument2, @pStory))
   RETURN pStory
END FUNCTION
```

#### Return value

A pointer to the **ITextStory** interface of the active story.

#### Result code

If the method succeeds, **GetLAstREsult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetActiveStory"></a>SetActiveStory

Sets the active story; that is, the story that receives keyboard and mouse input.

```
FUNCTION CTextDocument2.SetActiveStory (BYVAL pStory AS ITextStory PTR) AS HRESULT
   this.SetResult(m_pTextDocument2->lpvtbl->SetActiveStory(m_pTextDocument2, pStory))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pStory* | The story to set as active. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetMainStory"></a>GetMainStory

Gets the main story.

```
FUNCTION CTextDocument2.GetMainStory () AS ITextStory PTR
   DIM pStory AS ITextStory PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetMainStory(m_pTextDocument2, @pStory))
   RETURN pStory
END FUNCTION
```

#### Return value

A pointer to the **ITextStory** interface of the main story.

#### Result code

If this method succeeds, **GetLastResult** returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

A rich edit control automatically includes the main story; a call to the **GetNewStory** method is not required.

# <a name="GetNewStory"></a>GetNewStory

Gets a new story. Not implemented.

```
FUNCTION CTextDocument2.GetNewStory () AS ITextStory PTR
   DIM pStory AS ITextStory PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetNewStory(m_pTextDocument2, @pStory))
   RETURN pStory
END FUNCTION
```

#### Return value

A pointer to the **ITextStory** interface of the new story.

#### Result code

If this method succeeds, **GetLastResult** returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetStory"></a>GetStory

Retrieves the story that corresponds to a particular index.

```
FUNCTION CTextDocument2.GetStory (BYVAl Index AS LONG) AS ITextStory PTR
   DIM pStory AS ITextStory PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetStory(m_pTextDocument2, Index, @pStory))
   RETURN pStory
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Index* | The index of the story to retrieve. |

#### Return value

A pointer to the **ITextStory** interface of the requested story.

#### Result code

If this method succeeds, **GetLastResult** returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetLastResult"></a>GetLastResult

Returns the last result code

```
FUNCTION CTOMBase.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
```

# <a name="SetResult"></a>SetResult

Sets the last result code.

```
FUNCTION CTOMBase.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Result* | The **HRESULT** error code returned by the methods. |

# <a name="GetErrorInfo"></a>GetErrorInfo

Returns a description of the last result code.

```
FUNCTION CTOMBase.GetErrorInfo () AS CWSTR
   IF SUCCEEDED(m_Result) THEN RETURN "Success"
   DIM s AS CWSTR = "Error &h" & HEX(m_Result, 8)
   SELECT CASE m_Result
      CASE E_POINTER : s += ": E_POINTER - Null pointer"
      CASE S_OK : s += ": S_OK - Success"
      CASE S_FALSE : s += ": S_FALSE - Failure"
      CASE E_NOTIMPL : s += ": E_NOTIMPL - Not implemented."
      CASE E_INVALIDARG : s += ": E_INVALIDARG - Invalid argument"
      CASE E_OUTOFMEMORY : s += ": E_OUTOFMEMORY - Insufficient memory"
'      CASE E_FILENOTFOUND : s += "E_FILENOTFOUND - File not found"
      CASE &h80070002 : s += "E_FILENOTFOUND - File not found"
      CASE E_ACCESSDENIED : s += "E_ACCESSDENIED - Access denied"
      CASE E_FAIL : s += ": E_FAIL - Access denied"
      CASE NOERROR : s += ": NOERROR - Success" '' (same as S_OK)
      CASE CO_E_RELEASED:  : s += ": CO_E_RELEASED: - The object has been released"
      CASE ELSE
         s += "Unknown error"
   END SELECT
   RETURN s
END FUNCTION
```
