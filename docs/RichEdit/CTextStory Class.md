# CTextStory Class

Class that wraps all the methods of the **ITextStory** interface.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTORS](#CONSTRUCTORS) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [LET](#LET) | Assignment operator. |
| [CAST](#CAST) | Cast operator. |
| [TextStoryPtr](#TextStoryPtr) | Returns a pointer to the underlying **ITextStory** interface. |
| [Attach](#Attach) | Attaches an **ITextStory** interface pointer to the class. |
| [Detach](#Detach) | Detaches the underlying **ITextStory** interface pointer from the class. |

### ITextStory Interface

The **ITextStory** interface methods are used to access shared data from multiple stories, which is stored in the parent **ITextServices** instance.

The stories can be "edited" simultaneously by using individual **ITextRange2** methods, and displayed independently of one another. In addition, one story at a time can be UI active; that is, it receives keyboard and mouse input.

The **ITextStory** is a lightweight interface that does not require an **ITextRange2** object. This allows the client to manipulate a story, which is a faster, smaller object than a complete editing instance.

| Name       | Description |
| ---------- | ----------- |
| [GetActive](#GetActive) | Gets the active state of a story. |
| [SetActive](#SetActive) | Sets the active state of a story. |
| [GetDisplay](#GetDisplay) | Gets a new display for a story. |
| [GetIndex](#GetIndex) | Gets the index of a story. |
| [GetType](#GetType) | Gets this story's type. |
| [SetType](#SetType) | Sets this story's type. |
| [GetProperty](#GetProperty) | Gets the value of the specified property. |
| [GetRange](#GetRange) | Gets a text range object for the story. |
| [GetText](#GetText) | Gets the text in a story according to the specified conversion flags. |
| [SetFormattedText](#SetFormattedText) | Replaces a story’s text with specified formatted text. |
| [SetProperty](#SetProperty) | Sets the value of the specified property. |
| [SetText](#SetText) | Sets the text in this range. |

### Methods inherited from CTOMBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#GetLastResult) | Returns the last result code |
| [SetResult](#SetResult) | Sets the last result code. |
| [GetErrorInfo](#GetErrorInfo) | Returns a description of the last result code. |

# <a name="CONSTRUCTORS"></a>CONSTRUCTORS

Called when a **CTextStory** class variable is created.

```
DECLARE CONSTRUCTOR
DECLARE CONSTRUCTOR (BYVAL pTextStory AS ITextStory PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

## CONSTRUCTOR (Empty)

Can be used, for example, when we have an **ITextStory** interface pointer returned by a function and we want to attach it to a new instance of the **CTextStory** class.

```
DIM pCTextStory AS CTextStory
pCTextStory.Attach(pTextStory)
```
## CONSTRUCTOR (ITextStory PTR)

```
CONSTRUCTOR CTextStory (BYVAL pTextStory AS ITextStory PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Store the pointer of ITextStory interface
   IF pTextStory THEN
      IF fAddRef THEN pTextStory->lpvtbl->AddRef(pTextStory)
   END IF
   m_pTextStory = pTextStory
END CONSTRUCTOR
```

| Parameter | Description |
| --------- | ----------- |
| *pTextStory* | An **ITextStory** interface pointer. |
| *fAddRef* | Optional. **TRUE** to increment the reference count of the passed **ITextStory** interface pointer; otherwise, **FALSE**. Default is **FALSE**. |

#### Return value

A pointer to the new instance of the class.

# <a name="DESTRUCTOR"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

```
DESTRUCTOR CTextStory
   ' // Release the ITextStory interface
   IF m_pTextStory THEN m_pTextStory->lpvtbl->Release(m_pTextStory)
END DESTRUCTOR
```

# <a name="LET"></a>LET

Assignment operator. The assigned pointer must be an "addrefed" one.

```
OPERATOR CTextStory.LET (BYVAL pTextStory AS ITextStory PTR)
   m_Result = 0
   IF pTextStory = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF pTextStory THEN pTextStory->lpvtbl->Release(m_pTextStory)
   ' // Attach the passed interface pointer to the class
   m_pTextStory = pTextStory
END OPERATOR
```

# <a name="CAST"></a>CAST

Cast operator.

```
OPERATOR CTextStory.CAST () AS ITextStory PTR
   m_Result = 0
   OPERATOR = m_pTextStory
END OPERATOR
```

# <a name="TextStoryPtr"></a>TextStoryPtr

Returns a pointer to the underlying **ITextStory** interface

```
FUNCTION CTextStory.TextRangePtr () AS ITextStory PTR
   m_Result = 0
   RETURN m_pTextStory
END FUNCTION
```

# <a name="Attach"></a>Attach

Attaches an **ITextStory** interface pointer to the class.

```
FUNCTION CTextStory.Attach (BYVAL pTextStory AS ITextStory PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextStory = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextStory THEN m_Result = m_pTextStory->lpvtbl->Release(m_pTextStory)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextStory->lpvtbl->AddRef(pTextStory)
   m_pTextStory = pTextStory
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pTextStory* | The **ITextStory** interface pointer to attach. |
| *fAddRef* | **TRUE** to increment the reference count of the object; otherwise, **FALSE**. Default is **FALSE**. |

# <a name="Detach"></a>Detach

Detaches the interface pointer from the class

```
FUNCTION CTextStory.Detach () AS ITextStory PTR
   m_Result = 0
   DIM pTextStory AS ITextStory PTR = m_pTextStory
   m_pTextStory = NULL
   RETURN pTextStory
END FUNCTION
```

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

# <a name="GetActive"></a>GetActive

Gets the active state of a story.

```
FUNCTION CTextStory.GetActive () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextStory->lpvtbl->GetActive(m_pTextStory, @Value))
   RETURN Value
END FUNCTION
```

#### Return value

The active state. It can be one of the following values.

| Value | Meaning |
| ----- | ------- |
| **tomDisplayActive** | The story has no UI or display (fast and lightweight). |
| **tomDisplayUIActive** | The story is UI active; that is, gets keyboard and mouse interactions. |
| **tomInactive** | The story has display, but no UI. |
| **tomUIActive** | The story has display and UI activity. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetActive"></a>SetActive

Sets the active state of a story.

```
FUNCTION CTextStory.SetActive (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextStory->lpvtbl->SetActive(m_pTextStory, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The active state. It can be one of the following values. |

| Value | Meaning |
| ----- | ------- |
| **tomDisplayActive** | The story has no UI or display (fast and lightweight). |
| **tomDisplayUIActive** | The story is UI active; that is, gets keyboard and mouse interactions. |
| **tomInactive** | The story has display, but no UI. |
| **tomUIActive** | The story has display and UI activity. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetDisplay"></a>GetDisplay

Gets a new display for a story.

```
FUNCTION CTextStory.GetDisplay () AS IUnknown PTR
   DIM pUnk AS IUnknown PTR
   this.SetResult(m_pTextStory->lpvtbl->GetDisplay(m_pTextStory, @pUnk))
   RETURN pUnk
END FUNCTION
```

#### Return value

The IUnknown interface for a display.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

A story can be displayed by calling **SetActive(tomDisplayActive)**. The **GetDisplay** method is included, in case it might be advantageous to have more than one display for a set of **ITextStory** interfaces.


# <a name="GetIndex"></a>GetIndex

Gets the index of this story.

```
FUNCTION CTextStory.GetIndex () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextStory->lpvtbl->GetIndex(m_pTextStory, @Value))
   RETURN Value
END FUNCTION
```

#### Return value

The index of the story.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The index is used with the **GetStory** method of the **ITextDocument2** interface.


# <a name="GetType"></a>GetType

Gets this story's type.

```
FUNCTION CTextStory.GetType () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextStory->lpvtbl->GetType(m_pTextStory, @Value))
   RETURN Value
END FUNCTION
```

#### Return value

This story's type. It can be any of the following values, or a custom client value from 100 to 65535.

| Type | Description |
| ---- | ----------- |
| **tomCommentsStory** | The story used for comments. |
| **tomEndnotesStory** | The story used for endnotes. |
| **tomEvenPagesFooterStory** | The story containing footers for even pages. |
| **tomEvenPagesHeaderStory** |The story containing headers for even pages.|
| **tomFindStory** | The story used for a Find dialog. |
| **tomFirstPageFooterStory** | The story containing the footer for the first page. |
| **tomFirstPageHeaderStory** | The story containing the header for the first page. |
| **tomFootnotesStory** | The story used for footnotes. |
| **tomMainTextStory** | The main story always exists for a rich edit control. |
| **tomPrimaryFooterStory** | The story containing footers for odd pages. |
| **tomPrimaryHeaderStory** | The story containing headers for odd pages. |
| **tomReplaceStory** | The story used for a Replace dialog. |
| **tomScratchStory** | The scratch story. |
| **tomTextFrameStory** | The story used for a text box. |
| **tomUnknownStory** | No special type. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetType"></a>SetType

Sets this story's type.

```
FUNCTION CTextStory.SetType (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextStory->lpvtbl->SetType(m_pTextStory, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | This story's type. It can be any of the following values, or a custom client value from 100 to 65535. |

| Type | Description |
| ---- | ----------- |
| **tomCommentsStory** | The story used for comments. |
| **tomEndnotesStory** | The story used for endnotes. |
| **tomEvenPagesFooterStory** | The story containing footers for even pages. |
| **tomEvenPagesHeaderStory** |The story containing headers for even pages.|
| **tomFindStory** | The story used for a Find dialog. |
| **tomFirstPageFooterStory** | The story containing the footer for the first page. |
| **tomFirstPageHeaderStory** | The story containing the header for the first page. |
| **tomFootnotesStory** | The story used for footnotes. |
| **tomMainTextStory** | The main story always exists for a rich edit control. |
| **tomPrimaryFooterStory** | The story containing footers for odd pages. |
| **tomPrimaryHeaderStory** | The story containing headers for odd pages. |
| **tomReplaceStory** | The story used for a Replace dialog. |
| **tomScratchStory** | The scratch story. |
| **tomTextFrameStory** | The story used for a text box. |
| **tomUnknownStory** | No special type. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetProperty"></a>GetProperty

Gets the value of the specified property.

```
FUNCTION CTextStory.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextStory->lpvtbl->GetProperty(m_pTextStory, nType, @Value))
   RETURN Value
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *nType* | The ID of the property. Currently, no extra properties are defined. |

#### Return value

The property value.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetRange"></a>GetRange

Gets a text range object for the story.

```
FUNCTION CTextStory.GetRange (BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG) AS ITextRange2 PTR
   DIM pRange AS ITextRange2 PTR
   IF m_pTextStory = NULL THEN m_Result = E_POINTER: RETURN pRange
   this.SetResult(m_pTextStory->lpvtbl->GetRange(m_pTextStory, cpActive, cpAnchor, @pRange))
   RETURN pRange
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *cpActive* | The active end of the range. |
| *cpAnchor* | The anchor end of the range. |

#### Return value

The text range object.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetText"></a>GetText

Gets the text in a story according to the specified conversion flags.

```
FUNCTION CTextStory.GetText (BYVAL Flags AS LONG) AS CBSTR
   DIM pText AS AFX_BSTR
   this.SetResult(m_pTextStory->lpvtbl->GetText(m_pTextStory, Flags, @pText))
   RETURN pText
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Flags* | The conversion flags. A *Flags* value of 0 retrieves text the same as **ITextRange.GetText**. Other values include the following: **tomAdjustCRLF, tomAllowFinalEOP, tomFoldMathAlpha, tomIncludeNumbering, tomNoHidden, tomNoMathZoneBrackets, tomLanguageTag, tomTextize, tomTranslateTableCell, tomUseCRLF**. |

#### Return value

The text in the story.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

This method is similar to using **ITextRange2.GetText2** for a whole story, but it doesn’t require a range.


# <a name="SetFormattedText"></a>SetFormattedText

Replaces a story’s text with specified formatted text.

```
FUNCTION CTextStory.SetFormattedText (BYVAL pUnk AS IUnknown PTR) AS HRESULT
   this.SetResult(m_pTextStory->lpvtbl->SetFormattedText(m_pTextStory, pUnk))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pUnk* | The formatted text to replace the story’s text. |

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

This method calls **IUnknown.QueryInterface** for an **ITextRange2** interface.


# <a name="SetProperty"></a>SetProperty

Sets the value of the specified property.

```
FUNCTION CTextStory.SetProperty (BYVAL nType AS LONG, BYVAL VaLue AS LONG) AS HRESULT
   this.SetResult(m_pTextStory->lpvtbl->SetProperty(m_pTextStory, nType, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *nType* | The Microsoft accountID that identifies the property. Currently, no extra properties are defined. |
| *Value* | The new property value. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetText"></a>SetText

Replaces the text in a story with the specified text.

```
FUNCTION CTextStory.SetText (BYVAL Flags AS LONG, BYREF cbs AS CBSTR) AS HRESULT
   this.SetResult(m_pTextStory->lpvtbl->SetText(m_pTextStory, Flags, cbs))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Flags* | Flags controlling how the text is inserted as defined in the following table: **tomCheckTextLimit, tomMathCFCheck, tomUnhide, tomUnicodeBiDi, tomUnlink**. |
| *cbs* | The new text for this story. If this parameter is **""**, the text in the story is deleted. |

#### Return value

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
