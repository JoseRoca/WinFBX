# CTextRange2 Class

Class that wraps all the methods of the **ITextRange** and **ITextRange2** interfaces.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTORS](#CONSTRUCTORS) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [LET](#LET) | Assignment operator. |
| [CAST](#CAST) | Cast operator. |
| [TextDocumentPtr](#TextDocumentPtr) | Returns a pointer to the underlying **ITextRange2** interface. |
| [Attach](#Attach) | Attaches an **ITextRange2** interface pointer to the class. |
| [Detach](#Detach) | Detaches the underlying **ITextRange2** interface pointer from the class. |

### ITextRange Interface

The **ITextRange** objects are powerful editing and data-binding tools that allow a program to select text in a story and then examine or change that text.

The **ITextRange** interface inherits from the **IDispatch** interface. **ITextRange** also has these types of members:

| Name       | Description |
| ---------- | ----------- |
| [GetText](#GetText) | Gets the plain text in this range. |
| [SetText](#SetText) | Sets the plain text in this range. |
| [GetChar](#GetChar) | Gets the character at the start position of the range. |
| [SetChar](#SetChar) | Sets the character at the starting position of the range. |
| [GetDuplicate](#GetDuplicate) | Gets a duplicate of this range object. |
| [GetFormattedText](#GetFormattedText) | Gets an **ITextRange** object with the specified range's formatted text. |
| [SetFormattedText](#SetFormattedText) | Sets the formatted text of this range text to the formatted text of the specified range. |
| [GetStart](#GetStart) | Gets the start character position of the range. |
| [SetStart](#SetStart) | Sets the character position for the start of this range. |
| [GetEnd](#GetEnd) | Gets the end character position of the range. |
| [SetEnd](#SetEnd) | Sets the end position of the range. |
| [GetFont](#GetFont) | Gets an **ITextFont** object with the character attributes of the specified range. |
| [SetFont](#SetFont) | Sets this range's character attributes to those of the specified **ITextFont** object. |
| [GetPara](#GetPara) | Gets an **ITextPara** object with the paragraph attributes of the specified range. |
| [SetPara](#SetPara) | Sets the paragraph attributes of this range to those of the specified **ITextPara** object. |
| [GetStoryLength](#GetStoryLength) | Gets the count of characters in the range's story. |
| [GetStoryType](#GetStoryType) | Get the type of the range's story. |
| [Collapse](#Collapse) | Collapses the specified text range into a degenerate point at either the beginning or end of the range. |
| [Expand](#Expand) | Expands this range so that any partial units it contains are completely contained. |
| [GetIndex](#GetIndex) | Retrieves the story index of the *Unit* parameter at the specified range Start character position. |
| [SetIndex](#SetIndex) | Changes this range to the specified unit of the story. |
| [SetRange](#SetRange) | Adjusts the range endpoints to the specified values. |
| [InRange](#InRange) | Determines whether this range is within or at the same text as a specified range. |
| [InStory](#InStory) | Determines whether this range's story is the same as a specified range's story. |
| [IsEqual](#IsEqual) | Determines whether this range has the same character positions and story as those of a specified range. |
| [Select](#Select) | Sets the start and end positions, and story values of the active selection, to those of this range. |
| [StartOf](#StartOf) | Moves the range ends to the start of the first overlapping *Unit* in the range. |
| [EndOf](#EndOf) | Moves this range's ends to the end of the last overlapping *Unit* in the range. |
| [Move](#Move) | Moves the insertion point forward or backward a specified number of units. |
| [MoveStart](#MoveStart) | Moves the start position of the range the specified number of units in the specified direction. |
| [MoveEnd](#MoveEnd) | Moves the end position of the range. |
| [MoveWhile](#MoveWhile) | Starts at a specified end of a range and searches while the characters belong to the set specified by *Cset* and while the number of characters is less than or equal to *Coun*t. |
| [MoveStartWhile](#MoveStartWhile) | Moves the start position of the range either *Count* characters, or just past all contiguous characters that are found in the set of characters specified by *Cset*, whichever is less. |
| [MoveEndWhile](#MoveEndWhile) | Moves the end of the range either *Count* characters or just past all contiguous characters that are found in the set of characters specified by *Cset*, whichever is less. |
| [MoveUntil](#MoveUntil) | Searches up to *Count* characters for the first character in the set of characters specified by *Cset*. |
| [MoveStartUntil](#MoveStartUntil) | Moves the start position of the range either *Count* characters, or just past all contiguous characters that are found in the set of characters specified by *Cset*, whichever is less. |
| [MoveEndUntil](#MoveEndUntil) | Moves the end of the range either *Count* characters or just past all contiguous characters that are found in the set of characters specified by *Cset*, whichever is less. |
| [FindText](#FindText) | Searches up to *Count* characters for the text given by *bstr*. The starting position and direction are also specified by *Count*, and the matching criteria are given by *Flags*. |
| [FindTextStart](#FindTextStart) | Searches up to *Count* characters for the string, *bstr*, starting at the range's Start *cp (cpFirst)*. The search is subject to the comparison parameter, *Flags*. |
| [FindTextEnd](#FindTextEnd) | Searches up to *Count* characters for the string, *bstr*, starting from the range's End *cp*. The search is subject to the comparison parameter, *Flags*. |
| [Delete](#Delete) | Mimics the DELETE and BACKSPACE keys, with and without the CTRL key depressed. |
| [Cut](#Cut) | Cuts the plain or rich text to a data object or to the Clipboard, depending on the *pVar* parameter. |
| [Copy](#Copy) | Copies the text to a data object. |
| [Paste](#Paste) | Pastes text from a specified data object. |
| [CanPaste](#CanPaste) | Determines if a data object can be pasted, using a specified format, into the current range. |
| [CanEdit](#CanEdit) | Determines whether the specified range can be edited. |
| [ChangeCase](#ChangeCase) | Changes the case of letters in this range according to the *nType* parameter. |
| [GetPoint](#GetPoint) | Retrieves screen coordinates for the start or end character position in the text range, along with the intra-line position. |
| [SetPoint](#SetPoint) | Changes the range based on a specified point at or up through (depending on *Extend*) the point (x, y) aligned according to *nType*. |
| [ScrollIntoView](#ScrollIntoView) | Scrolls the specified range into view. |
| [GetEmbeddedObject](#GetEmbeddedObject) | Retrieves a pointer to the embedded object at the start of the specified range, that is, at *cpFirst*. |

### ITextSelection Interface

A text selection is a text range with selection highlighting.

The **ITextSelection** interface inherits from **ITextRange**. **ITextSelection** also has these types of members:

| Name       | Description |
| ---------- | ----------- |
| [GetFlags](#GetFlags) | Gets the text selection flags. |
| [SetFlags](#SetFlags) | Sets the text selection flags. |
| [GetType](#GetType) | Gets the type of text selection. |
| [MoveLeft](#MoveLeft) | Generalizes the functionality of the Left Arrow key. |
| [MoveRight](#MoveRight) | Generalizes the functionality of the Right Arrow key. |
| [MoveUp](#MoveUp) | Mimics the functionality of the Up Arrow and Page Up keys. |
| [MoveDown](#MoveDown) | Mimics the functionality of the Down Arrow and Page Down keys. |
| [HomeKey](#HomeKey) | Generalizes the functionality of the Home key. |
| [EndKey](#EndKey) | Mimics the functionality of the End key. |
| [TypeText](#TypeText) | Types the string given by *bstr* at this selection as if someone typed it.  |

### ITextRange2 Interface

The **ITextRange2** objects are powerful editing and data-binding tools that enable a program to select text in a story and then examine or change that text.

The **ITextRange2** interface inherits from **ITextSelection**, that in turn inherits from **ITextRange**. **ITextRange2** also has these types of members:

| Name       | Description |
| ---------- | ----------- |
| [GetCch](#GetCch) | Gets the count of characters in a range. |
| [GetCells](#GetCells) | Gets a cells object with the parameters of cells in the currently selected table row or column. |
| [GetColumn](#GetColumn) | Gets the column properties for the currently selected column. |
| [GetCount](#GetCount) | Gets the count of subranges, including the active subrange in the current range. |
| [GetDuplicate2](#GetDuplicate2) | Gets a duplicate of a range object. |
| [GetFont2](#GetFont2) | Gets an **ITextFont2** object with the character attributes of the current range. |
| [SetFont2](#SetFont2) | Sets the character formatting attributes of the range. |
| [GetFormattedText2](#GetFormattedText2) | Gets an **ITextRange2** object with the current range's formatted text. |
| [SetFormattedText2](#SetFormattedText2) | Sets the text of this range to the formatted text of the specified range. |
| [GetGravity](#GetGravity) | Gets the gravity of this range. |
| [SetGravity](#SetGravity) | Sets the gravity of this range. |
| [GetPara2](#GetPara2) | Gets an **ITextPara2** object with the paragraph attributes of a range. |
| [SetPara2](#SetPara2) | Sets the paragraph format attributes of a range. |
| [GetRow](#GetRow) | Gets the row properties in the currently selected row. |
| [GetStartPara](#GetStartPara) | Gets the character position of the start of the paragraph that contains the range's start character position. |
| [GetTable](#GetTable) |  Not implemented. Gets the table properties in the currently selected table. |
| [GetURL](#GetURL) | Returns the URL text associated with a range. |
| [SetURL](#SetURL) | Sets the text in this range to that of the specified URL. |
| [AddSubrange](#AddSubrange) | Adds a subrange to this range. |
| [BuildUpMath](#BuildUpMath) | Converts the linear-format math in a range to a built-up form, or modifies the current built-up form. |
| [DeleteSubrange](#DeleteSubrange) | Deletes a subrange from a range. |
| [Find](#Find) | Searches for math inline functions in text as specified by a source range. |
| [GetChar2](#GetChar2) | Gets the character at the specified offset from the end of this range. |
| [GetDropCap](#GetDropCap) | Gets the drop-cap parameters of the paragraph that contains this range. |
| [GetInlineObject](#GetInlineObject) | Gets the properties of the inline object at the range active end. |
| [GetProperty](#GetProperty) | Gets the value of a property. |
| [GetRect](#GetRect) | Retrieves a rectangle of the specified type for the current range. |
| [GetSubrange](#GetSubrange) | Retrieves a subrange in a range. |
| [GetText2](#GetText2) | Gets the text in this range according to the specified conversion flags. |
| [HexToUnicode](#HexToUnicode) | Converts and replaces the hexadecimal number at the end of this range to a Unicode character. |
| [InsertTable](#InsertTable) | Inserts a table in a range. |
| [Linearize](#Linearize) | Translates the built-up math, ruby, and other inline objects in this range to linearized form. |
| [SetActiveSubrange](#SetActiveSubrange) | Makes the specified subrange the active subrange of this range. |
| [SetDropCap](#SetDropCap) | Not implemented. Sets the drop-cap parameters for the paragraph that contains the current range. |
| [SetProperty](#SetProperty) | Sets the value of the specified property. |
| [SetText2](#SetText2) | Sets the text of this range. |
| [UnicodeToHex](#UnicodeToHex) | Converts the Unicode character(s) preceding the start position of this text range to a hexadecimal number, and selects it. |
| [SetInlineObject](#SetInlineObject) | Sets the text of this range. |
| [GetMathFunctionType](#GetMathFunctionType) | Retrieves the math function type associated with the specified math function name. |
| [InsertImage](#InsertImage) | Inserts an image into this range. |

### Methods inherited from CTOMBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#GetLastResult) | Returns the last result code |
| [SetResult](#SetResult) | Sets the last result code. |
| [GetErrorInfo](#GetErrorInfo) | Returns a description of the last result code. |

# <a name="CONSTRUCTORS"></a>CONSTRUCTORS

Called when a **CTextRange2** class variable is created.

```
DECLARE CONSTRUCTOR
DECLARE CONSTRUCTOR (BYVAL pTextRange2 AS ITextRange2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

## CONSTRUCTOR (Empty)

Can be used, for example, when we have an **ITextRange2** interface pointer returned by a function and we want to attach it to a new instance of the **CTextRange2** class.

```
DIM DIM pCTextRange2 AS CTextRange2
pCCTextRange2.Attach(pTextRange2)
```
## CONSTRUCTOR (ITextRange2 PTR)

```
CONSTRUCTOR CTextRange2 (BYVAL pTextRange2 AS ITextRange2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   IF pTextRange2 = NULL THEN m_Result = E_INVALIDARG : EXIT CONSTRUCTOR
   IF fAddRef THEN pTextRange2->lpvtbl->AddRef(pTextRange2)
   ' // Store the pointer of ITextRange2 interface
   m_pTextRange2 = pTextRange2
END CONSTRUCTOR
```

| Parameter | Description |
| --------- | ----------- |
| *pTextTextRange2* | An **ITextRange2** interface pointer. |
| *fAddRef* | Optional. TRUE to increment the reference count of the passed **ITextDocument2** interface pointer; otherise, FALSE. Default is FALSE. |

#### Return value

A pointer to the new instance of the class.

#### Usage examples

To use with the dotted syntax.

```
SCOPE
   ' // Create a new instance of the CTextDocument2 class
   DIM pCTextDocument2 AS CTextDocument2 = hRichEdit
   ' // Get the number of characters of the text in the Rich Edit control
   DIM numChars AS LONG = RichEdit_GetTextLength(hRichEdit)
   ' // Get the 0-based range of all the text
   DIM pCRange2 AS CTextRange2 = pCTextDocument2.Range2(0, numChars)
   ' // Get the text
   DIM cbsText AS CBSTR = pCRange2.GetText2(0)
   ' // The CTextDocument2 class and the CTextRange2 class will be destroyed when the scope ends
END SCOPE
```
**Note**: The following lines of code:
```
DIM pCRange2 AS CTextRange2 = pCTextDocument2.Range2(0, numChars)
DIM cbsText AS CBSTR = pCRange2.GetText2(0)
```
can be replaced with a compound syntax:
```
DIM cbsText AS CBSTR = CTextRange2(pCTextDocument2.Range2(0, numChars)).GetText2(0)
```
To use with the pointer syntax.

```
' // Create a new instance of the CTextDocument2 class
DIM pCTextDocument2 AS CTextDocument2 PTR = NEW CTextDocument2(hRichEdit)
' // Get the number of characters of the text in the Rich Edit control
DIM numChars AS LONG = RichEdit_GetTextLength(hRichEdit)
' // Get the 0-based range of all the text
DIM pCRange2 AS CTextRange2 PTR = NEW CTextRange2(pCTextDocument2->Range2(0, numChars))
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
DESTRUCTOR CTextRange2
   ' // Release the ITextRange2 interface
   IF m_pTextRange2 THEN m_pTextRange2->lpvtbl->Release(m_pTextRange2)
END DESTRUCTOR
```

# <a name="LET"></a>LET

Assignment operator.

```
OPERATOR CTextRange2.LET (BYVAL pTextRange2 AS ITextRange2 PTR)
   m_Result = 0
   IF pTextRange2 = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextRange2 THEN m_pTextRange2->lpvtbl->Release(m_pTextRange2)
   ' // Attach the passed interface pointer to the class
   m_pTextRange2 = pTextRange2
END OPERATOR
```

# <a name="CAST"></a>CAST

Cast operator.

```
OPERATOR CTextRange2.CAST () AS ITextRange2 PTR
   m_Result = 0
   OPERATOR = m_pTextRange2
END OPERATOR
```

# <a name="TextRangePtr"></a>TextRangePtr

Returns a pointer to the underlying **ITextRange2** interface

```
FUNCTION CTextRange2.TextRangePtr () AS ITextRange2 PTR
   m_Result = 0
   RETURN m_pTextRange2
END FUNCTION
```

# <a name="Attach"></a>Attach

Attaches an **ITextRange2** interface pointer to the class.

```
PRIVATE FUNCTION CTextRange2.Attach (BYVAL pTextRange2 AS ITextRange2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextRange2 = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextRange2 THEN m_Result = m_pTextRange2->lpvtbl->Release(m_pTextRange2)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextRange2->lpvtbl->AddRef(pTextRange2)
   m_pTextRange2 = pTextRange2
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pTextRange2* | The **ITextDocument2** interface pointer to attach. |
| *fAddRef* | **TRUE** to increment the reference count of te object. Default is FALSE. |

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

# <a name="GetText"></a>GetText

Gets the plain text in this range.

```
FUNCTION CTextRange2.GetText () AS CBSTR
   DIM pText AS AFX_BSTR
   this.SetResult(m_pTextRange2->lpvtbl->GetText(m_pTextRange2, @pText))
   RETURN pText
END FUNCTION
```

#### Return value

The retrieved text.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following error code.

| Result code | Description |
| ----------- | ----------- |
| **E_OUTOFMEMORY** | Insufficient memory to hold the text. |

#### Usage example

```
' // Create a new instance of the CTextDocument2 class
DIM pTextDocument2 AS CTextDocument2 = hRichEdit
' // Get the number of characters of the text in the Rich Edit control
DIM numChars AS LONG = RichEdit_GetTextLength(hRichEdit)
' // Get the 0-based range of all the text
DIM pCRange2 AS CTextRange2 = pCTextDocument.Range2(0, numChars)
' // Get the text
DIM cbsText AS CBSTR = pCRange2.GetText
```

**Note**: The following lines of code:
```
DIM pCRange2 AS CTextRange2 = pCTextDocument2.Range2(0, numChars)
DIM cbsText AS CBSTR = pCRange2.GetText2(0)
```
can be replaced with a compound syntax:
```
DIM cbsText AS CBSTR = CTextRange2(pCTextDocument2.Range2(0, numChars)).GetText2(0)
```

# <a name="SetText"></a>SetText

Sets the text in this range.

```
FUNCTION CTextRange2.SetText (BYREF cbs AS CBSTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetText(m_pTextRange2, cbs))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *cbs* | Text that replaces the current text in this range. If "", the current text is deleted. |

The method returns an **HRESULT** value. If the method succeeds, it returns S_OK. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_ACCESSDENIED** | Text is write-protected. |
| **E_OUTOFMEMORY** | Insufficient memory to hold the text. |

#### Usage example
```
DIM pCTextDocument AS CTextDocument2 = hRichEdit
DIM pCRange2 AS CTextRange2 = pCTextDocument.Range2(3, 8)
DIM cbsText AS CBSTR = "new text"
pCRange2.SetText(cbsText)
' You can also use a string literal
pCRange2.SetText("new text")
' or pass an empty string to delete the range
pCRange2.SetText("")
```

# <a name="GetChar"></a>GetChar

Gets the character at the start position of the range.

```
FUNCTION CTextRange2.GetChar () AS CBSTR
   DIM Char AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetChar(m_pTextRange2, @Char))
   RETURN WCHR(Char)
END FUNCTION
```
#### Return value

The character at the start position of the range, or an empty string on failure.

#### Usage example

```
DIM pCTextDocument AS CTextDocument2 = hRichEdit
IF pCTextDocument THEN
   DIM pCRange2 AS CTextRange2 = pCTextDocument.Range2(3, 8)
   IF pCRange2 THEN
      DIM cbs AS CBSTR = pcRange2.GetChar
   END IF
END IF
```

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**; if it fails, it returns **S_FALSE**.

# <a name="SetChar"></a>SetChar

Sets the character at the starting position of the range.

```
FUNCTION CTextRange2.SetChar (BYREF cbsChar AS CBSTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetChar(m_pTextRange2, ASC(cbsChar)))
   FUNCTION = m_Result
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *cbsChar* | New value for character at the starting position. |

#### Return code

The method returns an **HRESULT** value. If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following error codes.

| Return code | Description |
| ----------- | ----------- |
| **E_ACCESSDENIED** | Text is write-protected. |
| **E_OUTOFMEMORY** | Out of memory. |

#### Usage example

```
DIM pCTextDocument AS CTextDocument2 = hRichEdit
IF pCTextDocument THEN
   DIM pCRange2 AS CTextRange2 = pCTextDocument.Range2(3, 8)
   IF pCRange2 THEN pcRange2.SetChar("X")
END IF
```

# <a name="GetDuplicate"></a>GetDuplicate

```
FUNCTION CTextRange2.GetDuplicate () AS ITextRange2 PTR
   DIM pRange AS ITextRange2 PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetDuplicate(m_pTextRange2, @pRange))
   RETURN pRange
END FUNCTION
```

# <a name="GetFormattedText"></a>GetFormattedText

```
FUNCTION CTextRange2.GetFormattedText () AS ITextRange2 PTR
   DIM pRange AS ITextRange2 PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetFormattedText(m_pTextRange2, @pRange))
   RETURN pRange
END FUNCTION
```

# <a name="SetFormattedText"></a>SetFormattedText

```
FUNCTION CTextRange2.SetFormattedText (BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetFormattedText(m_pTextRange2, pRange))
   FUNCTION = m_Result
END FUNCTION
```

# <a name="GetStart"></a>GetStart

```
FUNCTION CTextRange2.GetStart () AS LONG
   DIM cpFirst AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetStart(m_pTextRange2, @cpFirst))
   RETURN cpFirst
END FUNCTION
```

# <a name="SetStart"></a>SetStart

```
FUNCTION CTextRange2.SetStart (BYVAL cpFirst AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetStart(m_pTextRange2, cpFirst))
   FUNCTION = m_Result
END FUNCTION
```

# <a name="GetEnd"></a>GetEnd

```
FUNCTION CTextRange2.GetEnd () AS LONG
   DIM cpLim AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetEnd(m_pTextRange2, @cpLim))
   RETURN cpLim
END FUNCTION
```

# <a name="SetEnd"></a>SetEnd

```
FUNCTION CTextRange2.SetEnd (BYVAL cpLim AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetEnd(m_pTextRange2, cpLim))
   FUNCTION = m_Result
END FUNCTION
```

# <a name="GetFont"></a>GetFont

```
FUNCTION CTextRange2.GetFont () AS ITextFont PTR
   DIM pFont AS ITextFont PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetFont(m_pTextRange2, @pFont))
   RETURN pFont
END FUNCTION
```

# <a name="SetFont"></a>SetFont

```
FUNCTION CTextRange2.SetFont (BYVAL pFont AS ITextFont PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetFont(m_pTextRange2, pFont))
   FUNCTION = m_Result
END FUNCTION
```

# <a name="GetPara"></a>GetPara

```
FUNCTION CTextRange2.GetPara () AS ITextPara PTR
   DIM pPara AS ITextPara PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetPara(m_pTextRange2, @pPara))
   RETURN pPara
END FUNCTION
```

# <a name="SetPara"></a>SetPara

```
FUNCTION CTextRange2.SetPara (BYVAL pPara AS ITextPara PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetPara(m_pTextRange2, pPara))
   FUNCTION = m_Result
END FUNCTION
```

# <a name="GetStoryLength"></a>GetStoryLength

```
FUNCTION CTextRange2.GetStoryLength () AS LONG
   DIM Count AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetStoryLength(m_pTextRange2, @Count))
   RETURN Count
END FUNCTION
```

# <a name="GetStoryType"></a>GetStoryType

```
FUNCTION CTextRange2.GetStoryType () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetStoryType(m_pTextRange2, @Value))
   RETURN Value
END FUNCTION
```

# <a name="Collapse"></a>Collapse

```
FUNCTION CTextRange2.Collapse (BYVAL bStart AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->Collapse(m_pTextRange2, bStart))
   RETURN m_Result
END FUNCTION
```

# <a name="Expand"></a>Expand

```
FUNCTION CTextRange2.Expand (BYVAL Unit AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->Expand(m_pTextRange2, Unit, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="GetIndex"></a>GetIndex

```
FUNCTION CTextRange2.GetIndex (BYVAL Unit AS LONG) AS LONG
   DIM Index AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetIndex(m_pTextRange2, Unit, @Index))
   RETURN Index
END FUNCTION
```

# <a name="SetIndex"></a>SetIndex

```
FUNCTION CTextRange2.SetIndex (BYVAL Unit AS LONG, BYVAL Index AS LONG, BYVAl Extend AS LONG) AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->SetIndex(m_pTextRange2, Unit, Index, Extend))
   RETURN m_Result
END FUNCTION
```

# <a name="SetRange"></a>SetRange

```
FUNCTION CTextRange2.SetRange (BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetRange(m_pTextRange2, cpAnchor, cpActive))
   RETURN m_Result
END FUNCTION
```

# <a name="InRange"></a>InRange

```
FUNCTION CTextRange2.InRange (BYVAL pRange AS ITextRange2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->InRange(m_pTextRange2, pRange, @Value))
   RETURN Value
END FUNCTION
```

# <a name="InStory"></a>InStory

```
FUNCTION CTextRange2.InStory (BYVAL pRange AS ITextRange2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->InStory(m_pTextRange2, pRange, @Value))
   RETURN Value
END FUNCTION
```

# <a name="IsEqual"></a>IsEqual

```
FUNCTION CTextRange2.IsEqual (BYVAL pRange AS ITextRange2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->IsEqual(m_pTextRange2, pRange, @Value))
   RETURN Value
END FUNCTION
```

# <a name="Select_"></a>Select_

```
FUNCTION CTextRange2.Select_ () AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->Select(m_pTextRange2))
   RETURN m_Result
END FUNCTION
```

# <a name="StartOf"></a>StartOf

```
FUNCTION CTextRange2.StartOf (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->StartOf(m_pTextRange2, Unit, Extend, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="EndOf"></a>EndOf

```
FUNCTION CTextRange2.EndOf (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->EndOf(m_pTextRange2, Unit, Extend, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="Move"></a>Move

```
FUNCTION CTextRange2.Move (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->Move(m_pTextRange2, Unit, Count, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveStart"></a>MoveStart

```
FUNCTION CTextRange2.MoveStart (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveStart(m_pTextRange2, Unit, Count, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveEnd"></a>MoveEnd

```
FUNCTION CTextRange2.MoveEnd (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveEnd(m_pTextRange2, Unit, Count, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveWhile"></a>MoveWhile

```
FUNCTION CTextRange2.MoveWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveWhile(m_pTextRange2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveStartWhile"></a>MoveStartWhile

```
FUNCTION CTextRange2.MoveStartWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveStartWhile(m_pTextRange2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveEndWhile"></a>MoveEndWhile

```
FUNCTION CTextRange2.MoveEndWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveEndWhile(m_pTextRange2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveUntil"></a>MoveUntil

```
FUNCTION CTextRange2.MoveUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveUntil(m_pTextRange2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveStartUntil"></a>MoveStartUntil

```
FUNCTION CTextRange2.MoveStartUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveStartUntil(m_pTextRange2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveEndUntil"></a>MoveEndUntil

```
FUNCTION CTextRange2.MoveEndUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveEndUntil(m_pTextRange2, Cset, Count, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="FindText"></a>FindText

```
FUNCTION CTextRange2.FindText (BYREF cbs AS CBSTR, BYVAL cch AS LONG, BYVAL Flags AS LONG) AS LONG
   DIM Length AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->FindText(m_pTextRange2, cbs, cch, Flags, @Length))
   RETURN Length
END FUNCTION
```

# <a name="FindTextStart"></a>FindTextStart

```
FUNCTION CTextRange2.FindTextStart (BYREF cbs AS CBSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG) AS LONG
   DIM Length AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->FindTextStart(m_pTextRange2, cbs, Count, Flags, @Length))
   RETURN Length
END FUNCTION
```

# <a name="FindTextEnd"></a>FindTextEnd

```
FUNCTION CTextRange2.FindTextEnd (BYREF cbs AS CBSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG) AS LONG
   DIM Length AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->FindTextEnd(m_pTextRange2, cbs, Count, Flags, @Length))
   RETURN Length
END FUNCTION
```

# <a name="Delete_"></a>Delete_

```
FUNCTION CTextRange2.Delete_ (BYVAL Unit AS LONG, BYVAL Count AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->Delete_(m_pTextRange2, Unit, Count, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="Cut"></a>Cut

```
FUNCTION CTextRange2.Cut (BYVAL pVar AS VARIANT PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->Cut(m_pTextRange2, pVar))
   RETURN m_Result
END FUNCTION
```

# <a name="Copy"></a>Copy

```
FUNCTION CTextRange2.Copy (BYVAL pVar AS VARIANT PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->Copy(m_pTextRange2, pVar))
   RETURN m_Result
END FUNCTION
```

# <a name="Paste"></a>Paste

```
FUNCTION CTextRange2.Paste (BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->Paste(m_pTextRange2, pVar, Format))
   RETURN m_Result
END FUNCTION
```

# <a name="CanPaste"></a>CanPaste

```
FUNCTION CTextRange2.CanPaste (BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->CanPaste(m_pTextRange2, pVar, Format, @Value))
   RETURN Value
END FUNCTION
```

# <a name="CanEdit"></a>CanEdit

```
FUNCTION CTextRange2.CanEdit () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->CanEdit(m_pTextRange2, @Value))
   RETURN Value
END FUNCTION
```

# <a name="ChangeCase"></a>ChangeCase

```
FUNCTION CTextRange2.ChangeCase (BYVAL nType AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->ChangeCase(m_pTextRange2, nType))
   RETURN m_Result
END FUNCTION
```

# <a name="GetPoint"></a>GetPoint

```
FUNCTION CTextRange2.GetPoint (BYVAL nType AS LONG, BYVAL px AS LONG PTR, BYVAL py AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->GetPoint(m_pTextRange2, nType, px, py))
   RETURN m_Result
END FUNCTION
```

# <a name="SetPoint"></a>SetPoint

```
FUNCTION CTextRange2.SetPoint (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG, BYVAL Extend AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetPoint(m_pTextRange2, x, y, nType, Extend))
   RETURN m_Result
END FUNCTION
```

# <a name="ScrollIntoView"></a>ScrollIntoView

```
FUNCTION CTextRange2.ScrollIntoView (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->ScrollIntoView(m_pTextRange2, Value))
   RETURN m_Result
END FUNCTION
```
# <a name="GetEmbeddedObject"></a>GetEmbeddedObject

```
FUNCTION CTextRange2.GetEmbeddedObject () AS IUnknown PTR
   DIM pObject AS IUnknown PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetEmbeddedObject(m_pTextRange2, @pObject))
   RETURN pObject
END FUNCTION
```
# <a name="GetFlags"></a>GetFlags

```
FUNCTION CTextRange2.GetFlags () AS LONG
   DIM pFlags AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetFlags(m_pTextRange2, @pFlags))
   RETURN pFlags
END FUNCTION
```

# <a name="SetFlags"></a>SetFlags

```
FUNCTION CTextRange2.SetFlags (BYVAL Flags AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetFlags(m_pTextRange2, Flags))
   RETURN m_Result
END FUNCTION
```

# <a name="GetType"></a>GetType

```
FUNCTION CTextRange2.GetType () AS LONG
   DIM nType AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetType(m_pTextRange2, @nType))
   RETURN nType
END FUNCTION
```

# <a name="MoveLeft"></a>MoveLeft

```
FUNCTION CTextRange2.MoveLeft (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveLeft(m_pTextRange2, Unit, Count, Extend, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveRight"></a>MoveRight

```
FUNCTION CTextRange2.MoveRight (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveRight(m_pTextRange2, Unit, Count, Extend, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveUp"></a>MoveUp

```
FUNCTION CTextRange2.MoveUp (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveUp(m_pTextRange2, Unit, Count, Extend, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="MoveDown"></a>MoveDown

```
FUNCTION CTextRange2.MoveDown (BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->MoveDown(m_pTextRange2, Unit, Count, Extend, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="HomeKey"></a>HomeKey

```
FUNCTION CTextRange2.HomeKey (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->HomeKey(m_pTextRange2, Unit, Extend, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="EndKey"></a>EndKey

```
FUNCTION CTextRange2.EndKey (BYVAL Unit AS LONG, BYVAL Extend AS LONG) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->EndKey(m_pTextRange2, Unit, Extend, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="TypeText"></a>TypeText

```
FUNCTION CTextRange2.TypeText (BYREF cbs AS CBSTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->TypeText(m_pTextRange2, cbs))
   RETURN m_Result
END FUNCTION
```

# <a name="GetCch"></a>GetCch

```
FUNCTION CTextRange2.GetCch () AS LONG
   DIM cch AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetCch(m_pTextRange2, @cch))
   RETURN cch
END FUNCTION
```

# <a name="GetCells"></a>GetCells

```
FUNCTION CTextRange2.GetCells () AS IUnknown PTR
   DIM pCells AS IUnknown PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetCells(m_pTextRange2, @pCells))
   RETURN pCells
END FUNCTION
```

# <a name="GetColumn"></a>GetColumn

```
FUNCTION CTextRange2.GetColumn () AS IUnknown PTR
   DIM pColumn AS IUnknown PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetColumn(m_pTextRange2, @pColumn))
   RETURN pColumn
END FUNCTION
```

# <a name="GetCount"></a>GetCount

```
FUNCTION CTextRange2.GetCount () AS LONG
   DIM Count AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetCount(m_pTextRange2, @Count))
   RETURN Count
END FUNCTION
```

# <a name="GetDuplicate2"></a>GetDuplicate2

```
FUNCTION CTextRange2.GetDuplicate2 () AS ITextRange2 PTR
   DIM pRange AS ITextRange2 PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetDuplicate2(m_pTextRange2, @pRange))
   RETURN pRange
END FUNCTION
```

# <a name="GetFont2"></a>GetFont2

```
FUNCTION CTextRange2.GetFont2 () AS ITextFont2 PTR
   DIM pFont AS ITextFont2 PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetFont2(m_pTextRange2, @pFont))
   RETURN pFont
END FUNCTION
```

# <a name="SetFont2"></a>SetFont2

```
FUNCTION CTextRange2.SetFont2 (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetFont2(m_pTextRange2, pFont))
   RETURN m_Result
END FUNCTION
```

# <a name="GetFormattedText2"></a>GetFormattedText2

```
FUNCTION CTextRange2.GetFormattedText2 () AS ITextRange2 PTR
   DIM pRange AS ITextRange2 PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetFormattedText2(m_pTextRange2, @pRange))
   RETURN pRange
END FUNCTION
```

# <a name="SetFormattedText2"></a>SetFormattedText2

```
FUNCTION CTextRange2.SetFormattedText2 (BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetFormattedText2(m_pTextRange2, pRange))
   RETURN m_Result
END FUNCTION
```

# <a name="GetGravity"></a>GetGravity

```
FUNCTION CTextRange2.GetGravity () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetGravity(m_pTextRange2, @Value))
   RETURN Value
END FUNCTION
```

# <a name="SetGravity"></a>SetGravity

```
FUNCTION CTextRange2.SetGravity (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetGravity(m_pTextRange2, Value))
   RETURN m_Result
END FUNCTION
```

# <a name="GetPara2"></a>GetPara2

```
FUNCTION CTextRange2.GetPara2 () AS ITextPara2 PTR
   DIM pPara AS ITextPara2 PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetPara2(m_pTextRange2, @pPara))
   RETURN pPara
END FUNCTION
```

# <a name="SetPara2"></a>SetPara2

```
FUNCTION CTextRange2.SetPara2 (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetPara2(m_pTextRange2, pPara))
   RETURN m_Result
END FUNCTION
```

# <a name="GetRow"></a>GetRow

```
FUNCTION CTextRange2.GetRow () AS ITextRow PTR
   DIM pRow AS ITextRow PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetRow(m_pTextRange2, @pRow))
   RETURN pRow
END FUNCTION
```

# <a name="GetStartPara"></a>GetStartPara

```
FUNCTION CTextRange2.GetStartPara () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetStartPara(m_pTextRange2, @Value))
   RETURN Value
END FUNCTION
```

# <a name="GetTable"></a>GetTable

```
FUNCTION CTextRange2.GetTable () AS IUnknown PTR
   DIM pTable AS IUnknown PTR
   this.SetResult(m_pTextRange2->lpvtbl->GetTable(m_pTextRange2, @pTable))
   RETURN pTable
END FUNCTION
```

# <a name="GetURL"></a>GetURL

```
FUNCTION CTextRange2.GetURL () AS CBSTR
   DIM pbstr AS AFX_BSTR
   this.SetResult(m_pTextRange2->lpvtbl->GetURL(m_pTextRange2, @pbstr))
   RETURN pbstr
END FUNCTION
```

# <a name="SetURL"></a>SetURL

```
FUNCTION CTextRange2.SetURL (BYREF cbs AS CBSTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetURL(m_pTextRange2, cbs))
   RETURN m_Result
END FUNCTION
```

# <a name="AddSubrange"></a>AddSubrange

```
FUNCTION CTextRange2.AddSubrange (BYVAL cp1 AS LONG, BYVAL cp2 AS LONG, BYVAL Activate AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->AddSubrange(m_pTextRange2, cp1, cp2, Activate))
   RETURN m_Result
END FUNCTION
```

# <a name="BuildUpMath"></a>BuildUpMath

```
FUNCTION CTextRange2.BuildUpMath (BYVAL Flags AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->BuildUpMath(m_pTextRange2, Flags))
   RETURN m_Result
END FUNCTION
```

# <a name="DeleteSubrange"></a>DeleteSubrange

```
FUNCTION CTextRange2.DeleteSubrange (BYVAL cpFirst AS LONG, BYVAL cpLim AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->DeleteSubrange(m_pTextRange2, cpFirst, cpLim))
   RETURN m_Result
END FUNCTION
```

# <a name="Find"></a>Find

```
FUNCTION CTextRange2.Find (BYVAL pRange AS ITextRange2 PTR, BYVAL Count AS LONG, BYVAL Flags AS LONG) AS HRESULT
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->Find(m_pTextRange2, pRange, Count, Flags, @Delta))
   RETURN Delta
END FUNCTION
```

# <a name="GetChar2"></a>GetChar2

```
FUNCTION CTextRange2.GetChar2 (BYVAL Offset AS LONG) AS LONG
   DIM Char AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetChar2(m_pTextRange2, @Char, Offset))
   RETURN Char
END FUNCTION
```

# <a name="GetDropCap"></a>GetDropCap

```
FUNCTION CTextRange2.GetDropCap (BYVAL pcLine AS LONG PTR, BYVAL pPosition AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->GetDropCap(m_pTextRange2, pcLine, pPosition))
   RETURN m_Result
END FUNCTION
```

# <a name="GetInlineObject"></a>GetInlineObject

```
FUNCTION CTextRange2.GetInlineObject (BYVAL pcLine AS LONG PTR, BYVAL pType AS LONG PTR, BYVAL pAlign AS LONG PTR, BYVAL pChar AS LONG PTR, BYVAL pChar1 AS LONG PTR, _
BYVAL pChar2 AS LONG PTR, BYVAL pCount AS LONG PTR,  BYVAL pTeXStyle AS LONG PTR, BYVAL pcCol AS LONG PTR, BYVAL pLevel AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->GetInlineObject(m_pTextRange2, pcLine, pType, pAlign, pChar, pChar1, pChar2, pCount, pTeXStyle, pcCol))
   RETURN m_Result
END FUNCTION
```

# <a name="GetProperty"></a>GetProperty

```
FUNCTION CTextRange2.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetProperty(m_pTextRange2, nType, @Value))
   RETURN Value
END FUNCTION
```

# <a name="GetRect"></a>GetRect

```
FUNCTION CTextRange2.GetRect (BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, BYVAL pTop AS LONG PTR, _
BYVAL pRight AS LONG PTR, BYVAL pBottom AS LONG PTR, BYVAL pHit AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->GetRect(m_pTextRange2, nType, pLeft, pTop, pRight, pBottom, pHit))
   RETURN m_Result
END FUNCTION
```

# <a name="GetSubrange"></a>GetSubrange

```
FUNCTION CTextRange2.GetSubrange (BYVAL iSubrange AS LONG, BYVAL pcpFirst AS LONG PTR, BYVAL pcpLim AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->GetSubrange(m_pTextRange2, iSubrange, pcpFirst, pcpLim))
   RETURN m_Result
END FUNCTION
```

# <a name="GetText2"></a>GetText2

Gets the text in this range according to the specified conversion flags.

```
FUNCTION CTextRange2.GetText2 (BYVAL Flags AS LONG) AS CBSTR
   DIM pText2 AS AFX_BSTR
   this.SetResult(m_pTextRange2->lpvtbl->GetText2(m_pTextRange2, Flags, @pText2))
   RETURN pText2
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Flags* | The flags controlling how the text is retrieved. The flags can include a combination of the values in the table below. Specifying a *Flags* value of 0 is the same as calling the **GetText** method. |

| Flag | Value | Description |
| ---- | ----- | ----------- |
| **tomAdjustCRLF** | 1 | Adjust CR/LFs at the start. |
| **tomUseCRLF** | 2 | Use CR/LF in place of a carriage return or a line feed. |
| **tomIncludeNumbering** | 64 | Include list numbers. |
| **tomNoHidden** | 32 | Don't include hidden text. |
| **tomNoMathZoneBrackets** | &h100 | Don't include math zone brackets. |
| **tomTextize** | 4 | Copy up to &hFFFC (OLE object). |
| **tomAllowFinalEOP** | 8 | Allow a final end-of-paragraph (EOP) marker. |
| **tomTranslateTableCell** | 128 | Replace table row delimiter characters with spaces. |
| **tomFoldMathAlpha** | 128 | Replace table row delimiter characters with spaces. |
| **tomLanguageTag** | 16 | Fold math alphanumerics to ASCII/Greek. |

#### Return value

The text in the range

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied.. |
| **E_OUTOFMEMORY** | Insufficient memory to hold the text. |

#### Remarks

This method includes the special flag **tomLanguageTag** to get the BCP-47 language tag for the range. This is an industry standard language tag which may be preferable to the language code identifier (LCID) obtained by calling the **GetLanguageID** method of the **ITextFont** interface.

#### Usage example

```
' // Create a new instance of the CTextDocument2 class
DIM pTextDocument2 AS CTextDocument2 = hRichEdit
' // Get the number of characters of the text in the Rich Edit control
DIM numChars AS LONG = RichEdit_GetTextLength(hRichEdit)
' // Get the 0-based range of all the text
DIM pCRange2 AS CTextRange2 = pCTextDoc.Range2(0, numChars)
' // Get the text
DIM cbsText AS CBSTR = pCRange2.GetText2(0)
```

# <a name="HexToUnicode"></a>HexToUnicode

```
FUNCTION CTextRange2.HexToUnicode () AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->HexToUnicode(m_pTextRange2))
   RETURN m_Result
END FUNCTION
```

# <a name="InsertTable"></a>InsertTable

```
FUNCTION CTextRange2.InsertTable (BYVAL cCol AS LONG, BYVAL cRow AS LONG, BYVAL AutoFit AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->InsertTable(m_pTextRange2, cCol, cRow, AutoFit))
   RETURN m_Result
END FUNCTION
```

# <a name="Linearize"></a>Linearize

```
FUNCTION CTextRange2.Linearize (BYVAL Flags AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->Linearize(m_pTextRange2, Flags))
   RETURN m_Result
END FUNCTION
```

# <a name="SetActiveSubrange"></a>SetActiveSubrange

```
FUNCTION CTextRange2.SetActiveSubrange (BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetActiveSubrange(m_pTextRange2, cpAnchor, cpActive))
   RETURN m_Result
END FUNCTION
```

# <a name="SetDropCap"></a>SetDropCap

```
FUNCTION CTextRange2.SetDropCap (BYVAL cLine AS LONG, BYVAL Position AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetDropCap(m_pTextRange2, cLine, Position))
   RETURN m_Result
END FUNCTION
```

# <a name="SetProperty"></a>SetProperty

```
FUNCTION CTextRange2.SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetProperty(m_pTextRange2, nType, Value))
   RETURN m_Result
END FUNCTION
```

# <a name="SetText2"></a>SetText2

Sets the text of this range.

```
FUNCTION CTextRange2.SetText2 (BYVAL Flags AS LONG, BYREF cbs AS CBSTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetText2(m_pTextRange2, Flags, cbs))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Flags* | Flags controlling how the text is inserted in the range. The flag can be one of the following values (see table below) |
| *cbs* | Text that replaces the current text in this range. If "", the current text is deleted. |

| Flag | Value | Description |
| ---- | ----- | ----------- |
| **tomUnicodeBiDi** | 1 | Use the Unicode bidirectional (bidi) algorithm. |
| **tomMathCFCheck** | 4 | Check math-zone character formatting. |
| **tomUnlink** | 8 | Don't include text as part of a hyperlink. |
| **tomUnhide** | &h10 | Don't insert as hidden text. |
| **tomCheckTextLimit** | &h20 | Obey the current text limit instead of increasing the text to fit. |
| **tomLanguageTag** | &h1000 | Get the BCP-47 language tag for this range. |

#### Return value

The method returns an **HRESULT** value. If the method succeeds, it returns S_OK. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Text is write-protected. |
| **E_OUTOFMEMORY** | Insufficient memory to hold the text. |

#### Remarks

This method is similar to **SetText**, but lets the client specify flags that control various insertion options, including the special flag **tomLanguageTag** to get the BCP-47 language tag for the range. This is an industry standard language tag that may be preferable **SetLanguageID**, which uses a language code identifier (LCID).

#### Usage example
```
DIM pCTextDocument AS CTextDocument2 = hRichEdit
DIM pCRange2 AS CTextRange2 = pCTextDoc.Range2(3, 8)
DIM cbsText AS CBSTR = "new text"
pCRange2.SetText2(0, cbsText)
' You can also use a string literal
pCRange2.SetText2(0, "new text")
' or pass an empty string to delete the range
pCRange2.SetText2(0, "")
```

# <a name="UnicodeToHex"></a>UnicodeToHex

```
FUNCTION CTextRange2.UnicodeToHex () AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->UnicodeToHex(m_pTextRange2))
   RETURN m_Result
END FUNCTION
```

# <a name="SetInlineObject"></a>SetInlineObject

```
FUNCTION CTextRange2.SetInlineObject (BYVAL nType AS Long, BYVAL Align AS LONG, BYVAL Char AS LONG, _
BYVAL Char1 AS LONG, BYVAL Char2 AS LONG, BYVAL Count AS LONG, BYVAL TeXStyle AS LONG, BYVAL cCol AS LONG) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->SetInlineObject(m_pTextRange2, nType, Align, Char, Char1, Char2, Count, TeXStyle, cCol))
   RETURN m_Result
END FUNCTION
```

# <a name="GetMathFunctionType"></a>GetMathFunctionType

```
FUNCTION CTextRange2.GetMathFunctionType (BYREF cbs AS CBSTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->GetMathFunctionType(m_pTextRange2, cbs, @Value))
   RETURN Value
END FUNCTION
```

# <a name="InsertImage"></a>InsertImage

```
FUNCTION CTextRange2.InsertImage (BYVAL width_ AS LONG, BYVAL height AS LONG, BYVAL ascent AS LONG, _
BYVAL nType AS LONG, BYREF cbsAltText AS CBSTR, BYVAL pStream AS IStream PTR) AS HRESULT
   this.SetResult(m_pTextRange2->lpvtbl->InsertImage(m_pTextRange2, width_, height, ascent, nType, cbsAltText, pStream))
   RETURN m_Result
END FUNCTION
```
