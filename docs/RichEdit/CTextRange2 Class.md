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
| *pTextDocpTextRange2ument2* | An **ITextRange2** interface pointer. |
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
