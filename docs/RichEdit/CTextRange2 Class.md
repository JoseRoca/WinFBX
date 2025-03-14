# CTextRange2 Class

Class that wraps all the methods of the **ITextRange**, **ITextSelection** and **ITextRange2** interfaces.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTOR](#constructor) | Called when a class variable is created. |
| [DESTRUCTOR](#destructor) | Called automatically when a class variable goes out of scope or is destroyed. |

### ITextRange Interface

The **ITextRange** objects are powerful editing and data-binding tools that allow a program to select text in a story and then examine or change that text.

| Name       | Description |
| ---------- | ----------- |
| [GetText](#gettext) | Gets the text in this range according to the specified conversion flags. |
| [SetText](#settext) | Sets the plain text in this range. |
| [GetChar](#getchar) | Gets the character at the start position of the range. |
| [SetChar](#setchar) | Sets the character at the starting position of the range. |
| [GetChar2](#getchar2) | Gets the character at the specified offset from the end of this range. |
| [GetDuplicate](#getduplicate) | Gets a duplicate of this range object. |
| [GetFormattedText](#getformattedtext) | Gets an **ITextRange** object with the specified range's formatted text. |
| [SetFormattedText](#setformattedtext) | Sets the formatted text of this range text to the formatted text of the specified range. |
| [GetStart](#getstart) | Gets the start character position of the range. |
| [SetStart](#setstart) | Sets the character position for the start of this range. |
| [GetEnd](#getend) | Gets the end character position of the range. |
| [SetEnd](#setend) | Sets the end position of the range. |
| [GetLenght](#getlength) | Gets the length of the range. |
| [SetLenght](#setlength) | Sets the length of the range. |
| [GetFont](#getfont) | Gets an **ITextFont** object with the character attributes of the specified range. |
| [SetFont](#setfont) | Sets this range's character attributes to those of the specified **ITextFont** object. |
| [GetPara](#getpara) | Gets an **ITextPara** object with the paragraph attributes of the specified range. |
| [SetPara](#setpara) | Sets the paragraph attributes of this range to those of the specified **ITextPara** object. |
| [GetStoryLength](#getstorylength) | Gets the count of characters in the range's story. |
| [GetStoryType](#getstorytype) | Get the type of the range's story. |
| [Collapse](#collapse) | Collapses the specified text range into a degenerate point at either the beginning or end of the range. |
| [Expand](#expand) | Expands this range so that any partial units it contains are completely contained. |
| [GetIndex](#getIndex) | Retrieves the story index of the *Unit* parameter at the specified range Start character position. |
| [SetIndex](#setIndex) | Changes this range to the specified unit of the story. |
| [SetRange](#setrange) | Adjusts the range endpoints to the specified values. |
| [InRange](#inrange) | Determines whether this range is within or at the same text as a specified range. |
| [InStory](#instory) | Determines whether this range's story is the same as a specified range's story. |
| [IsEqual](#isequal) | Determines whether this range has the same character positions and story as those of a specified range. |
| [Select_](#select_) | Sets the start and end positions, and story values of the active selection, to those of this range. |
| [StartOf](#startof) | Moves the range ends to the start of the first overlapping *Unit* in the range. |
| [EndOf](#endof) | Moves this range's ends to the end of the last overlapping *Unit* in the range. |
| [Move](#move) | Moves the insertion point forward or backward a specified number of units. |
| [Movestart](#movestart) | Moves the start position of the range the specified number of units in the specified direction. |
| [MoveEnd](#moveend) | Moves the end position of the range. |
| [MoveWhile](#movewhile) | Starts at a specified end of a range and searches while the characters belong to the set specified by *Cset* and while the number of characters is less than or equal to *Coun*t. |
| [MoveStartWhile](#movestartwhile) | Moves the start position of the range either *Count* characters, or just past all contiguous characters that are found in the set of characters specified by *Cset*, whichever is less. |
| [MoveEndWhile](#moveendwhile) | Moves the end of the range either *Count* characters or just past all contiguous characters that are found in the set of characters specified by *Cset*, whichever is less. |
| [MoveUntil](#moveuntil) | Searches up to *Count* characters for the first character in the set of characters specified by *Cset*. |
| [MoveStartUntil](#movestartuntil) | Moves the start position of the range either *Count* characters, or just past all contiguous characters that are found in the set of characters specified by *Cset*, whichever is less. |
| [MoveEndUntil](#moveenduntil) | Moves the end of the range either *Count* characters or just past all contiguous characters that are found in the set of characters specified by *Cset*, whichever is less. |
| [FindText](#findtext) | Searches up to *Count* characters for the text given by *bstr*. The starting position and direction are also specified by *Count*, and the matching criteria are given by *Flags*. |
| [FindTextStart](#findtextstart) | Searches up to *Count* characters for the string, *bstr*, starting at the range's Start *cp (cpFirst)*. The search is subject to the comparison parameter, *Flags*. |
| [FindTextEnd](#findtextend) | Searches up to *Count* characters for the string, *bstr*, starting from the range's End *cp*. The search is subject to the comparison parameter, *Flags*. |
| [Delete_](#delete_) | Mimics the DELETE and BACKSPACE keys, with and without the CTRL key depressed. |
| [Cut](#cut) | Cuts the plain or rich text to a data object or to the Clipboard, depending on the *pVar* parameter. |
| [Copy](#copy) | Copies the text to a data object. |
| [Paste](#paste) | Pastes text from a specified data object. |
| [CanPaste](#canpaste) | Determines if a data object can be pasted, using a specified format, into the current range. |
| [CanEdit](#canedit) | Determines whether the specified range can be edited. |
| [ChangeCase](#changecase) | Changes the case of letters in this range according to the *nType* parameter. |
| [GetPoint](#getpoint) | Retrieves screen coordinates for the start or end character position in the text range, along with the intra-line position. |
| [SetPoint](#setpoint) | Changes the range based on a specified point at or up through (depending on *Extend*) the point (x, y) aligned according to *nType*. |
| [ScrollIntoView](#scrollintoview) | Scrolls the specified range into view. |
| [GetEmbeddedObject](#getembeddedobject) | Retrieves a pointer to the embedded object at the start of the specified range, that is, at *cpFirst*. |

## Helper Method

| Name       | Description |
| ---------- | ----------- |
| [GetRtf](#getrtf) | Gets the content of the range as a Rich Text Format (RTF) string. |

### ITextSelection Interface

A text selection is a text range with selection highlighting.

The **ITextSelection** interface inherits from **ITextRange**. **ITextSelection** also has these types of members:

| Name       | Description |
| ---------- | ----------- |
| [GetFlags](#getflags) | Gets the text selection flags. |
| [SetFlags](#setflags) | Sets the text selection flags. |
| [GetType](#gettype) | Gets the type of text selection. |
| [MoveLeft](#moveleft) | Generalizes the functionality of the Left Arrow key. |
| [MoveRight](#moveright) | Generalizes the functionality of the Right Arrow key. |
| [MoveUp](#moveup) | Mimics the functionality of the Up Arrow and Page Up keys. |
| [MoveDown](#movedown) | Mimics the functionality of the Down Arrow and Page Down keys. |
| [HomeKey](#homekey) | Generalizes the functionality of the Home key. |
| [EndKey](#endkey) | Mimics the functionality of the End key. |
| [TypeText](#typetext) | Types the string given by *cbs* at this selection as if someone typed it.  |

### ITextRange2 Interface

The **ITextRange2** objects are powerful editing and data-binding tools that enable a program to select text in a story and then examine or change that text.

| Name       | Description |
| ---------- | ----------- |
| [GetCch](#getcch) | Gets the count of characters in a range. |
| [Getcells](#getcells) | Not implemented. Gets a cells object with the parameters of cells in the currently selected table row or column. |
| [GetColumn](#getcolumn) | Not implemented. Gets the column properties for the currently selected column. |
| [GetCount](#getcount) | Gets the count of subranges, including the active subrange in the current range. |
| [GetGravity](#getgravity) | Gets the gravity of this range. |
| [SetGravity](#setgravity) | Sets the gravity of this range. |
| [GetRow](#getrow) | Gets the row properties in the currently selected row. |
| [GetStartPara](#getstartpara) | Gets the character position of the start of the paragraph that contains the range's start character position. |
| [GetTable](#gettable) |  Not implemented. Gets the table properties in the currently selected table. |
| [GetURL](#geturl) | Returns the URL text associated with a range. |
| [SetURL](#seturl) | Sets the text in this range to that of the specified URL. |
| [AddSubrange](#addsubrange) | Adds a subrange to this range. |
| [BuildUpMath](#buildupmath) | Converts the linear-format math in a range to a built-up form, or modifies the current built-up form. |
| [DeleteSubrange](#deletesubrange) | Deletes a subrange from a range. |
| [Find](#find) | Searches for math inline functions in text as specified by a source range. |
| [GetDropCap](#getdropcap) | Not implemented. Gets the drop-cap parameters of the paragraph that contains this range. |
| [GetInlineObject](#getinlineobject) | Gets the properties of the inline object at the range active end. |
| [GetProperty](#getproperty) | Gets the value of a property. |
| [GetRect](#getrect) | Retrieves a rectangle of the specified type for the current range. |
| [GetSubrange](#getsubrange) | Retrieves a subrange in a range. |
| [HexToUnicode](#hextounicode) | Converts and replaces the hexadecimal number at the end of this range to a Unicode character. |
| [InsertTable](#inserttable) | Inserts a table in a range. |
| [Linearize](#linearize) | Translates the built-up math, ruby, and other inline objects in this range to linearized form. |
| [SetActiveSubrange](#setactivesubrange) | Makes the specified subrange the active subrange of this range. |
| [SetDropCap](#setdropcap) | Not implemented. Sets the drop-cap parameters for the paragraph that contains the current range. |
| [SetProperty](#SetProperty) | Sets the value of the specified property. |
| [UnicodeToHex](#unicodetohex) | Converts the Unicode character(s) preceding the start position of this text range to a hexadecimal number, and selects it. |
| [SetInlineObject](#setinlineobject) | Sets the text of this range. |
| [GetMathFunctionType](#getmathfunctiontype) | Retrieves the math function type associated with the specified math function name. |
| [InsertImage](#insertimage) | Inserts an image into this range. |

### Methods inherited from CTextObjectBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#getlastresult) | Returns the last result code |
| [SetResult](#setresult) | Sets the last result code. |
| [GetErrorInfo](#geterrorinfo) | Returns a description of the last result code. |

## <a name="constructor"></a>CONSTRUCTOR

Called when a **CTextRange2** class variable is created.

```
DECLARE CONSTRUCTOR (BYVAL pTextRange2 AS ITextRange2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

| Parameter | Description |
| --------- | ----------- |
| *pTextRange2* | An **ITextRange2** interface pointer. |
| *fAddRef* | Optional. **TRUE** to increment the reference count of the passed **ITextRange2** interface pointer; otherwise, **FALSE**. Default is **FALSE**. |

#### Return value

A pointer to the new instance of the class.

#### Usage examples

<ins>To use with the dotted syntax.</ins>

```
SCOPE
   ' // Get the number of characters of the text in the Rich Edit control
   DIM numChars AS LONG = pRichEdit->TextLength
   ' // Get the 0-based range of all the text
   DIM pRange2 AS CTextRange2 = pRichEdit->Range2(0, numChars)
   ' // Get the text
   DIM cbsText AS CBSTR = pRange2.GetText(tomUseCRLF)
   ' // The CTextRange2 class will be destroyed when the scope ends
END SCOPE
```
<ins>To use with the pointer syntax.</ins>
```
' // Get the number of characters of the text in the Rich Edit control
DIM numChars AS LONG = pRichEdit->TextLength
' // Get the 0-based range of all the text
DIM pRange2 AS CTextRange2 PTR = NEW CTextRange2(pRichEdit->Range2(0, numChars))
' // Get the text
DIM cbsText AS CBSTR = pRange2->GetText(tomUseCRLF)
' // Delete the range
Delete pRange2
```
---
## <a name="destructor"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

```
DESTRUCTOR CTextRange2
```
---

## <a name="getlastresult"></a>GetLastResult

Returns the last result code

```
FUNCTION GetLastResult () AS HRESULT
```
---

## <a name="setresult"></a>SetResult

Sets the last result code.

```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Result* | The **HRESULT** error code returned by the methods. |

---

## <a name="GetErrorInfo"></a>GetErrorInfo

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS CWSTR
```
---

## <a name="gettext"></a>GetText

Gets the text in this range according to the specified conversion flags.

```
FUNCTION GetText (BYVAL Flags AS LONG = tomUseCRLF) AS CBSTR
```

| Parameter | Description |
| --------- | ----------- |
| *Flags* | The flags controlling how the text is retrieved. The flags can include a combination of the values in the table below. |

| Flag | Description |
| ---- | ----------- |
| **tomAdjustCRLF** | Adjust CR/LFs at the start. |
| **tomUseCRLF** | Use CR/LF in place of a carriage return or a line feed. |
| **tomIncludeNumbering** | Include list numbers. |
| **tomNoHidden** | Don't include hidden text. |
| **tomNoMathZoneBrackets** | Don't include math zone brackets. |
| **tomTextize** | Copy up to &hFFFC (OLE object). |
| **tomAllowFinalEOP** | Allow a final end-of-paragraph (EOP) marker. |
| **tomTranslateTableCell** | Replace table row delimiter characters with spaces. |
| **tomFoldMathAlpha** | Replace table row delimiter characters with spaces. |
| **tomLanguageTag** | Fold math alphanumerics to ASCII/Greek. |

#### Return value

CBSTR. The text in the range in unicode.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied.. |
| **E_OUTOFMEMORY** | Insufficient memory to hold the text. |

#### Remarks

This method uses the **GetText2** method of the **ITextRange2** interface to retrieve the text because it allows to use the conversion flags. Passing 0 in the *Flags* parameter is the same that using the **GetText** method of the **ITextRange** interface, which does not return correct results when dealing with files that use CR (MacIntosh) or LF (Linux), instead of CRLF (Windows).

This method includes the special flag **tomLanguageTag** to get the BCP-47 language tag for the range. This is an industry standard language tag which may be preferable to the language code identifier (LCID) obtained by calling the **GetLanguageID** method of the **ITextFont** interface.

<ins>To use with the dotted syntax.</ins>
```
SCOPE
   ' // Get the number of characters of the text in the Rich Edit control
   DIM numChars AS LONG = pRichEdit.TextLength
   ' // Get the 0-based range of all the text
   DIM pRange2 AS CTextRange2 = pRichEdit.Range2(0, numChars)
   ' // Get the text
   DIM cbsText AS CBSTR = pRange2.GetText(tomUseCRLF)
   ' // The CTextRange2 class will be destroyed when the scope ends
END SCOPE
```
<ins>To use with the pointer syntax.</ins>
```
' // Get the number of characters of the text in the Rich Edit control
DIM numChars AS LONG = pRichEdit->TextLength
' // Get the 0-based range of all the text
DIM pRange2 AS CTextRange2 PTR = NEW CTextRange2(pRichEdit->Range2(0, numChars))
' // Get the text
DIM cbsText AS CBSTR = pRange2->GetText(tomUseCRLF)
' // Delete the range
Delete pRange2
```
---

## <a name="settext"></a>SetText

Sets the text of this range.

```
FUNCTION SetText (BYREF cbs AS CBSTR, BYVAL Flags AS LONG = 0) AS HRESULT
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
DIM pRange2 AS CTextRange2 = pRichEdit.Range2(3, 8)
DIM cbsText AS CBSTR = "new text"
pRange2.SetText(cbsText)
' You can also use a string literal
pRange2.SetText("new text")
' or pass an empty string to delete the range
pRange2.SetText("")
```
---

## <a name="getchar"></a>GetChar

Gets the character at the start position of the range.

```
FUNCTION GetChar () AS LONG
```
#### Return value

The character at the start position of the range.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**; if it fails, it returns **S_FALSE**.

#### Usage example

```
DIM pRange2 AS CTextRange2 = pRichEdit.Range2(3, 8)
DIM char AS LONG = pcRange2.GetChar
```

#### Remarks

The character retrieved by this method is a **LONG**, which hide the way that they are stored in the backing store (as bytes, words, variable-length, and so forth), and they do not require using a **BSTR**.

The **Char** property, which can do most things that a characters collection can, has two big advantages:

- It can reference any character in the parent story instead of being limited to the parent range.
- It is significantly faster, since **LONG**s are involved instead of range objects.

Accordingly, the Text Object Model (TOM) does not support a characters collection.

---

## <a name="getchar2"></a>GetChar2

Gets the character at the specified offset from the end of this range.

```
FUNCTION GetChar2 (BYVAL Offset AS LONG) AS LONG
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Offset* | The offset from the end of the range. An offset of 0 gets the character at the end of the range. |

#### Return value

The character value.

#### Remarks

This method differs from **GetChar** in the following ways:

- It returns the UTF-32 character for the surrogate pair instead of the pair's lead code.
- It gets the character code, or codes, at the specified offset from the end of the range instead of the character at the start of the range.

If the character is the lead code for a surrogate pair, the corresponding UTF-32 character is returned.

If *Offset* specifies a character before the start of the story or at the end of the story, this method returns the character code 0.

| If the Offset value is | This character is returned |
| ---------------------- | -------------------------- |
| 0 | The character at the end of the range. |
| Negative and accesses the middle of a surrogate pair | The corresponding UTF-32 character. |
| Positive and accesses the middle of a surrogate pair | The UTF-32 character following that pair. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="setchar"></a>SetChar

Sets the character at the starting position of the range.

```
FUNCTION SetChar (BYVAL char AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *char* | New value for character at the starting position. |

#### Return code

The method returns an **HRESULT** value. If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_ACCESSDENIED** | Text is write-protected. |
| **E_OUTOFMEMORY** | Out of memory. |

#### Usage example

```
DIM pRange2 AS CTextRange2 = pRichEdit.Range2(3, 8)
pRange2.SetChar(ASC("X"))
```
---

## <a name="getduplicate"></a>GetDuplicate

Gets a duplicate of this range object.

```
FUNCTION GetDuplicate () AS ITextRange2 PTR
```

#### Return value

The duplicate of the range.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **E_FAIL**.

#### Usage example

```
DIM pRange2 AS CTextRange2 = pRichEdit.Range2(3, 8)
DIM pRangeDup AS CTextRange2 = pRange2.GetDuplicate
DIM cbsText AS CBSTR = pRangeDup.GetText
```
---

## <a name="getformattedtext"></a>GetFormattedText

Gets an **ITextRange2** object with the specified range's formatted text.

```
FUNCTION GetFormattedText () AS ITextRange2 PTR
```

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **E_OUTOFMEMORY**. 

---

## <a name="setformattedtext"></a>SetFormattedText

Sets the formatted text of this range text to the formatted text of the specified range.

```
FUNCTION SetFormattedText (BYVAL pRange AS ITextRange2 PTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *pRange* | The formatted text to replace this range's text. |

### Return code

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_ACCESSDENIED** | Text is write-protected. |
| **E_OUTOFMEMORY** | Out of memory. |

---

## <a name="getstart"></a>GetStart

Gets the start character position of the range.

```
FUNCTION GetStart () AS LONG
```
#### Return value

The start character position.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**.

---

## <a name="setstart"></a>SetStart

Sets the character position for the start of this range.

```
FUNCTION SetStart (BYVAL cpFirst AS LONG) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *cpFirst* | The new character position for the start of the range. |

#### Return value

The method returns an HRESULT value. If the method succeeds, it returns **S_OK**. If the method fails, it returns **S_FALSE**.

---

## <a name="getend"></a>GetEnd

Gets the end character position of the range.

```
FUNCTION GetEnd () AS LONG
```
#### Return value

The end character position.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If method fails, it returns **S_FALSE**.

#### Remarks

Although a pointer to a range remains valid when the text is edited, this is not the case for the character position. A character position is volatile; that is, it becomes invalid as soon as text is inserted or deleted before the character position. Be careful about using methods that return character position values, especially if the values are to be stored for any duration.

This method is similar to the **GetStart** method which gets the start character position of the range.

---

## <a name="setend"></a>SetEnd

Sets the end position of the range.

```
FUNCTION SetEnd (BYVAL cpLim AS LONG) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *cpLim* | The new end position. |

#### Return value

The method returns an **HRESULT** value. If the method succeeds, it returns **S_OK**. If the method fails, it returns **S_FALSE**.

#### Remarks

If the new end position is less than the start position, this method also sets the start position to *cp*; that is, the range becomes an insertion point.

If this range is actually the selection, the end position becomes the active end and, if the display is not frozen, it is scrolled into view.

**SetStart** sets the range's start position and **SetRange** sets both range ends simultaneously.

---

## <a name="getlength"></a>GetLength

Gets the length of the range.

```
FUNCTION GetLength () AS LONG
```
#### Return value

The lenght of the range

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If method fails, it returns **S_FALSE**.

---


## <a name="setlength"></a>SetLength

Sets the length of the range.

```
FUNCTION SetLength (BYVAL Value AS LONG) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | The new length. |

#### Return value

The method returns an **HRESULT** value. If the method succeeds, it returns **S_OK**. If the method fails, it returns **S_FALSE**.

---

## <a name="getfont"></a>GetFont

Gets an **ITextFont2** object with the character attributes of the current range. In this implementation of the class, **GetFont** and **GetFont2** are the same method.

```
FUNCTION GetFont () AS ITextFont PTR
```
#### Return value

The **ITextFont2** object.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an HRESULT error code.

---

## <a name="setfont"></a>SetFont

Sets the character formatting attributes of the range.

```
FUNCTION SetFont (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *pFont* | The font object with the desired character formatting attributes. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes:

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="getpara"></a>GetPara

Gets an **ITextPara2** object with the paragraph attributes of a range.

```
FUNCTION GetPara () AS ITextPara2 PTR
```
#### Return value

The **ITextPara2** object.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="setpara"></a>SetPara

Sets the paragraph format attributes of a range.

```
FUNCTION SetPara (BYVAL pPara AS ITextPara PTR) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *pPara* | The desired paragraph format. |

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="getstorylength"></a>GetStoryLength

Gets the count of characters in the range's story.

```
FUNCTION GetStoryLength () AS LONG
```
#### Return value

The count of characters in the range's story.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**.

---

## <a name="getstorytype"></a>GetStoryType

Get the type of the range's story.

```
FUNCTION GetStoryType () AS LONG
```
#### Return value

The type of the range's story. It can be one of the following values:

| Story type | Value |
| ---------- | ----- |
| **tomUnknownStory** | 0 |
| **tomMainTextStory** | 1 |
| **tomFootnotesStory** | 2 |
| **tomEndnotesStory** | 3 |
| **tomCommentsStory** | 4 |
| **tomTextFrameStory** | 5 |
| **tomEvenPagesHeaderStory** | 6 |
| **tomPrimaryHeaderStory** | 7 |
| **tomEvenPagesFooterStory** | 8 |
| **tomPrimaryFooterStory** | 9 |
| **tomFirstPageHeaderStory** | 10 |
| **tomFirstPageFooterStory** | 11 |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**.

---

## <a name="collapse"></a>Collapse

Collapses the specified text range into a degenerate point at either the beginning or end of the range.

```
FUNCTION Collapse (BYVAL bStart AS LONG = tomStart) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *bStart* | Flag specifying the end to collapse at. It can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomStart or tomTrue** | Range is collapsed to the start of the range. This is the default. |
| **tomEnd or tomFalse** | Range is collapsed to the end of the range. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**.

---

## <a name="expand"></a>Expand

Expands this range so that any partial units it contains are completely contained.

```
FUNCTION Expand (BYVAL Unit AS LONG = tomWord) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit to include, if it is partially within the range. The default value is **tomWord**.<br> For a list of *Unit* values, see the table below. |

| Unit | Value | Meaning |
| ---- | ----- | ------- |
| **tomCharacter** | 1 | Character. |
| **tomWord** | 2 | Word. |
| **tomSentence** | 3 | Sentence. |
| **tomParagraph** | 4 | Paragraph. |
| **tomLine** | 5 | Line (on display). |
| **tomStory** | 6 | Story. |
| **tomScreen** | 7 | Screen (as for PAGE UP/PAGE DOWN). |
| **tomSection** | 8 | Section. |
| **tomColumn** | 9 | Table column. |
| **tomRow** | 10 | Table row. |
| **tomWindow** | 11 | Upper-left or lower-right of the window. |
| **tomCell** | 12 | Table cell. |
| **tomCharFormat** | 13 | Run of constant character formatting. |
| **tomParaFormat** | 14 | Run of constant paragraph formatting. |
| **tomTable** | 15 | Table. |
| **tomObject** | 16 | Embedded object. |

#### Return value

The count of characters added to the range. The value can be null.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes:

| Result code | Description |
| ----------- | ----------- |
| **E_NOTIMPL** | Unit is not supported. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

For example, if an insertion point is at the beginning, the end, or within a word, **Expand** expands the range to include that word. If the range already includes one word and part of another, **Expand** expands the range to include both words. **Expand** expands the range to include the visible portion of the range's story.

---

## <a name="getindex"></a>GetIndex

Retrieves the story index of the *Unit* parameter at the specified range Start character position. The first *Unit* in a story has an index value of 1. The index of a *Unit* is the same for all character positions from that immediately preceding the *Unit* up to the last character in the *Unit*.

```
FUNCTION GetIndex (BYVAL Unit AS LONG = tomWord) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit to include, if it is partially within the range. The default value is **tomWord**.<br> For a list of *Unit* values, see the table below. |

| Unit | Value | Meaning |
| ---- | ----- | ------- |
| **tomCharacter** | 1 | Character. |
| **tomWord** | 2 | Word. |
| **tomSentence** | 3 | Sentence. |
| **tomParagraph** | 4 | Paragraph. |
| **tomLine** | 5 | Line (on display). |
| **tomStory** | 6 | Story. |
| **tomScreen** | 7 | Screen (as for PAGE UP/PAGE DOWN). |
| **tomSection** | 8 | Section. |
| **tomColumn** | 9 | Table column. |
| **tomRow** | 10 | Table row. |
| **tomWindow** | 11 | Upper-left or lower-right of the window. |
| **tomCell** | 12 | Table cell. |
| **tomCharFormat** | 13 | Run of constant character formatting. |
| **tomParaFormat** | 14 | Run of constant paragraph formatting. |
| **tomTable** | 15 | Table. |
| **tomObject** | 16 | Embedded object. |

#### Return value

The index value. The value is zero if Unit does not exist.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_NOTIMPL** | Unit is not supported. |

---

## <a name="setindex"></a>SetIndex

Changes this range to the specified unit of the story.

```
FUNCTION SetIndex (BYVAL Unit AS LONG, BYVAL Index AS LONG, BYVAl Extend AS LONG = 0) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit used to index the range. For a list of *Unit* values, see the table below. |
| *Index* | Index for the *Unit*. This range is relocated to the *Unit* that has this index number. If positive, the numbering of *Units* begins at the start of the story and proceeds forward. If negative, the numbering begins at the end of the story and proceeds backward. The start of the story corresponds to an *Index* of 1 for all units that exist, and the last unit in the story corresponds to an *Index* of -1. |
| *Extend* | Flag that indicates the extent of the range. If zero (the default), the range is collapsed to an insertion point at the start position of the specified *Unit*. If nonzero, the range is set to the entire *Unit*. |

| Unit | Value | Meaning |
| ---- | ----- | ------- |
| **tomCharacter** | 1 | Character. |
| **tomWord** | 2 | Word. |
| **tomSentence** | 3 | Sentence. |
| **tomParagraph** | 4 | Paragraph. |
| **tomLine** | 5 | Line (on display). |
| **tomStory** | 6 | Story. |
| **tomScreen** | 7 | Screen (as for PAGE UP/PAGE DOWN). |
| **tomSection** | 8 | Section. |
| **tomColumn** | 9 | Table column. |
| **tomRow** | 10 | Table row. |
| **tomWindow** | 11 | Upper-left or lower-right of the window. |
| **tomCell** | 12 | Table cell. |
| **tomCharFormat** | 13 | Run of constant character formatting. |
| **tomParaFormat** | 14 | Run of constant paragraph formatting. |
| **tomTable** | 15 | Table. |
| **tomObject** | 16 | Embedded object. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Index is not valid. |
| **E_NOTIMPL** | Unit is not supported. |
| **S_FALSE** | Failure for some other reason. |

---

## <a name="setrange"></a>SetRange

Adjusts the range endpoints to the specified values.

```
FUNCTION SetRange (BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *cpAnchor* | The character position for the anchor end of the range. |
| *cpActive* | The character position for the active end of the range. |

#### Return value

If this method succeeds, it returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

This method sets the range's start position to *min(cpActive, cpAnchor)*, and the end position to *max(cpActive, cpAnchor)*. If the range is a nondegenerate selection, *cpAnchor* is the active end, and *cpAnchor* is the anchor end. If the range is a degenerate selection, the selection is displayed at the start of the line, rather than at the end of the previous line.

This method removes any other subranges this range may have. To preserve the current subranges, use **SetActiveSubrange**.

If the text range is a selection, you can set the attributes of the selection by using the **SetFlags** method.

---

## <a name="inrange"></a>InRange

Determines whether this range is within or at the same text as a specified range.

```
FUNCTION InRange (BYVAL pRange AS ITextRange2 PTR) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *pRange* | Text that is compared to the current range. |

#### Return value

The comparison result. The pointer can be null. The method returns **tomTrue** only if the range is in or at the same text as *pRange*.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**.

#### Remarks
For range2 to be contained in range1, both ranges must be in the same story, and the limits of range2 must satisfy either of the following statements.

- The start and end character positions of range1 are the same as range2. That is, both ranges are degenerate and have identical insertion points.
- Range2 is a nondegenerate range with start and end character positions at or within those of range1.

The following example shows how to walk one range with another.

```
range2 = range1.GetDuplicate
range2.SetEnd = range2.SetStart       ' Collapse range2 to its start position 
While range2.InRange(range1)          ' Iterate so long as range2 remains within range1
   ...   ' This code would change the range2 character positions 
Wend
```
When the **FindText**, **MoveWhile**, and **MoveUntil** method families are used, you can use one range to walk another by specifying the appropriate limit count of characters (for an example, see the Remarks in **Find**).

**IsEqual** is a special case of **InRange** that returns **tomTrue** if the *pRange* has the same start and end character positions and belongs to the same story.

---

## <a name="instory"></a>InStory

Determines whether this range's story is the same as a specified range's story.

```
FUNCTION InStory (BYVAL pRange AS ITextRange2 PTR) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *pRange* | The **ITextRange2** object whose story is compared to this range's story. |

#### Return value

The comparison result. Returns **tomTrue** if this range's story is the same as that of the *pRange*; otherwise it receives **tomFalse**.

#### Result code

If the two stories are the same, **GetLastResult** returns **S_OK**. Otherwise, it returns **S_FALSE**.

---

## <a name="isequal"></a>IsEqual

Determines whether this range has the same character positions and story as those of a specified range.

```
FUNCTION IsEqual (BYVAL pRange AS ITextRange2 PTR) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *pRange* | The **ITextRange2** object that is compared to this range. |

#### Return value

The comparison result. Returns **tomTrue** if this range points at the same text (has the same start and end character positions and story) as *pRange*; otherwise it returns **tomFalse**.

#### Result code

If the ranges have the same character positions and story, **GetLastResult** returns **S_OK**. Otherwise, it returns **S_FALSE**.

The **IsEqual** method returns **tomTrue** only if the range points at the same text as *pRange*.

---

## <a name="Select_"></a>Select_

Sets the start and end positions, and story values of the active selection, to those of this range.

```
FUNCTION Select_ () AS HRESULT
```

#### Return value

The method returns an **HRESULT** value. If the method succeeds, it returns **S_OK**. If the method fails, it returns **S_FALSE**.

#### Remarks

The active end of the new selection is at the end position.

The caret for an ambiguous character position is displayed at the beginning of the line.

---

## <a name="startof"></a>StartOf

Moves the range ends to the start of the first overlapping *Unit* in the range.

```
FUNCTION StartOf (BYVAL Unit AS LONG, BYVAL Extend AS LONG = 0) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *Unit* | *Unit* to use in the move operation. For a list of *Unit* values, see the table below. |
| *Extend* | How to move the ends of the range. It can be one of the following values. |

|     |     |
| --- | --- |
| 0 (or tomMove) | Collapses a nondegenerate range to the start position by moving the insertion point. This is the default. |
| 1 (or tomExtend) | Moves the start position to the beginning of the overlapping Unit. Does not move the end position. |

Table of Unit values:

| Unit | Value | Meaning |
| ---- | ----- | ------- |
| **tomCharacter** | 1 | Character. |
| **tomWord** | 2 | Word. |
| **tomSentence** | 3 | Sentence. |
| **tomParagraph** | 4 | Paragraph. |
| **tomLine** | 5 | Line (on display). |
| **tomStory** | 6 | Story. |
| **tomScreen** | 7 | Screen (as for PAGE UP/PAGE DOWN). |
| **tomSection** | 8 | Section. |
| **tomColumn** | 9 | Table column. |
| **tomRow** | 10 | Table row. |
| **tomWindow** | 11 | Upper-left or lower-right of the window. |
| **tomCell** | 12 | Table cell. |
| **tomCharFormat** | 13 | Run of constant character formatting. |
| **tomParaFormat** | 14 | Run of constant paragraph formatting. |
| **tomTable** | 15 | Table. |
| **tomObject** | 16 | Embedded object. |

#### Return value

The signed number of characters that the start position is moved. It can be null. This value is always less than or equal to zero, because the motion is always toward the beginning of the story.

The method returns an HRESULT value. If the method succeeds, it returns S_OK. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_NOTIMPL** | Unit is not supported. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

If the range is an insertion point on a boundary between Units, **StartOf** does not change the start position.

The **StartOf** and **EndOf** methods differ from the **HomeKey** and **EndKey** methods in that the latter extend from the active end, whereas **StartOf** extends from the start position and **EndOf** extends from the end position.

---

## <a name="endof"></a>EndOf

Moves this range's ends to the end of the last overlapping Unit in the range.

```
FUNCTION EndOf (BYVAL Unit AS LONG, BYVAL Extend AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Unit* | *Unit* to use. efault value: **tomWord**. For a list of *Unit* values, see the table below. |
| *Extend* | Indicator of how the shifting of the range ends is to proceed. It can be one of the following. |

|     |     |
| --- | --- |
| 0 (or tomMove) | Collapses a nondegenerate range to the End of the original range by moving the insertion point. This is the default. |
| 1 (or tomExtend) | Moves End to the end of the overlapping *Unit*. Does not move Start. |

Table of Unit values:

| Unit | Value | Meaning |
| ---- | ----- | ------- |
| **tomCharacter** | 1 | Character. |
| **tomWord** | 2 | Word. |
| **tomSentence** | 3 | Sentence. |
| **tomParagraph** | 4 | Paragraph. |
| **tomLine** | 5 | Line (on display). |
| **tomStory** | 6 | Story. |
| **tomScreen** | 7 | Screen (as for PAGE UP/PAGE DOWN). |
| **tomSection** | 8 | Section. |
| **tomColumn** | 9 | Table column. |
| **tomRow** | 10 | Table row. |
| **tomWindow** | 11 | Upper-left or lower-right of the window. |
| **tomCell** | 12 | Table cell. |
| **tomCharFormat** | 13 | Run of constant character formatting. |
| **tomParaFormat** | 14 | Run of constant paragraph formatting. |
| **tomTable** | 15 | Table. |
| **tomObject** | 16 | Embedded object. |

#### Return value

The number of characters the insertion point or End is moved plus 1 if a collapse occurs to the entry End. If the range includes the final CR (carriage return) (at the end of the story) and *Extend* = **tomMove**, then the method returns –1 to indicate that the collapse occurred before the end of the range (because an insertion point cannot exist beyond the final CR).

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_NOTIMPL** | Unit is not supported. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

For comparison, the **StartOf** method moves the range ends to the beginning of the first overlapping *Unit* in the range. Note, the **StartOf** and **EndOf** methods differ from the **HomeKey** and **EndKey** methods in that the latter extend from the active end, whereas **StartOf** extends from Start and **EndOf** extends from End. If the range is an insertion point on a boundary between *Units*, **EndOf** does not change End. In particular, calling *EndOf (tomCharacter, 0)* does not change End except for an insertion point at the beginning of a story.

---

## <a name="Move"></a>Move

Moves the insertion point forward or backward a specified number of units. If the range is nondegenerate, the range is collapsed to an insertion point at either end, depending on *Count*, and then is moved.
```
FUNCTION Move (BYVAL Unit AS LONG = tomCharacter, BYVAL Count AS LONG = 1) AS LONG
   DIM Delta AS LONG
   this.SetResult(m_pTextRange2->lpvtbl->Move(m_pTextRange2, Unit, Count, @Delta))
   RETURN Delta
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit to use. The default value is **tomCharacter**. For information on other values, For a list of *Unit* values, see the table below. |
| *Count* | Number of *Units* to move past. The default value is 1. If *Count* is greater than zero, motion is forward—toward the end of the story—and if *Count* is less than zero, motion is backward—toward the beginning. If *Count* is zero, the range is unchanged. |

| Unit | Value | Meaning |
| ---- | ----- | ------- |
| **tomCharacter** | 1 | Character. |
| **tomWord** | 2 | Word. |
| **tomSentence** | 3 | Sentence. |
| **tomParagraph** | 4 | Paragraph. |
| **tomLine** | 5 | Line (on display). |
| **tomStory** | 6 | Story. |
| **tomScreen** | 7 | Screen (as for PAGE UP/PAGE DOWN). |
| **tomSection** | 8 | Section. |
| **tomColumn** | 9 | Table column. |
| **tomRow** | 10 | Table row. |
| **tomWindow** | 11 | Upper-left or lower-right of the window. |
| **tomCell** | 12 | Table cell. |
| **tomCharFormat** | 13 | Run of constant character formatting. |
| **tomParaFormat** | 14 | Run of constant paragraph formatting. |
| **tomTable** | 15 | Table. |
| **tomObject** | 16 | Embedded object. |

#### Return value

The actual number of *Units* the insertion point moves past. For more information, see the **Remarks** section.

#### Result code

If the method succeeds in moving the insertion point, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes:

| Result code | Description |
| ----------- | ----------- |
| **E_NOTIMPL** | Unit is not supported. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

If the range is degenerate (an insertion point), this method tries to move the insertion point *Count Units*.

If the range is nondegenerate and *Count* is greater than zero, this method collapses the range at the end character position, moves the resulting insertion point forward to a *Unit* boundary (if it is not already at one), and then tries to move *Count* - 1 *Units* forward. If the range is nondegenerate and *Count* is less than zero, this method collapses the range at the start character position, moves the resulting insertion point backward to a *Unit* boundary (if it isn't already at one), and then tries to move *Count* - 1 *Units* backward. Thus, in both cases, collapsing a nondegenerate range to an insertion point, whether moving to the start or end of the Unit following the collapse, counts as a *Unit*.

The **Move** method returns the number of *Units* actually moved. This method never moves the insertion point beyond the story of this range. If *Count Units* would move the insertion point before the beginning of the story, it is moved to the story beginning. Similarly, if *Count  Units* would move it beyond the end of the story, it is moved to the story end.

The **Move** method works similarly to the UI-oriented **MoveLeft** and **MoveRight** methods, except that the direction of motion is logical rather than geometrical. That is, with **Move** the direction is either toward the end or toward the start of the story. Depending on the language, moving toward the end of the story could be moving to the left or to the right. To get a feel for *Count*, press Ctrl+Right Arrow in a Microsoft Word document for a variety of selections. In left-to-right text, this keystroke behaves the same as *Move(tomWord, 1)*, and *MoveRight(tomWord, 1)*. *Count* corresponds to the number of times you press Ctrl+Right Arrow.

The return value is equal to the number of *Units* that the insertion point is moved including one *Unit* for collapsing a nondegenerate range and moving it to a *Unit* boundary. So, if no motion and no collapse occur, as when the range is an insertion point at the end of the story, the return value is zero. This approach is useful for controlling program loops that process a whole story.

In both of the cases mentioned above, calling *Move(tomWord, 1)* returns a value of 1 because the ranges were collapsed. Similarly, calling *Move(tomWord, -1)* returns a value of -1 for both cases. Collapsing, with or without moving part of a Unit to a *Unit* boundary, counts as a *Unit* moved.

The direction of motion refers to the logical character ordering in the plain-text backing store. This approach avoids the problems of geometrical ordering, such as left versus right and up versus down, in international software. Such geometrical methods are still needed in the edit engine, of course, since keyboards have arrow keys to invoke them. If the range is really an **ITextSelection** object, then methods like **MoveLeft** and **MoveRight** can be used.

If *Unit* specifies characters (**tomCharacter**), the Text Object Model (TOM) uses the Unicode character set. To convert between Unicode and multibyte character sets the **MultiByteToWideChar** and **WideCharToMultiByte** functions provide easy ways to convert between Unicode and multibyte character sets on import and export, respectively. For more information, see **Open**. In this connection, the use of a carriage return/line feed (CR/LF) to separate paragraphs is as problematic as double-byte character set (DBCS). The **ITextSelection** UI methods back up over a CR/LF as if it were a single character, but the **Move** methods count CR/LFs as two characters. It's clearly better to use a single character as a paragraph separator, which in TOM is represented by a character return, although the Unicode paragraph separator character, &h2029, is accepted. In general, TOM engines should support CR/LF, carriage return (CR), line feed (LF), vertical tab, form feed, and &h2029. Microsoft Rich Edit 2.0 also supports CR/CR/LF for backward compatibility.

See also the **MoveStart** and **MoveEnd** methods, which move the range Start or End position *Count  Units*, respectively.

## <a name="movestart"></a>MoveStart

Moves the start position of the range the specified number of units in the specified direction.

```
FUNCTION MoveStart (BYVAL Unit AS LONG = tomCharacter, BYVAL Count AS LONG = 1) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit used in the move. The default value is **tomCharacter**. For a list of *Unit* values, see the table below. |
| *Count* |Number of units to move. The default value is 1. If *Count* is greater than zero, motion is forward—toward the end of the story—and if *Count* is less than zero, motion is backward—toward the beginning. If *Count* is zero, the start position is unchanged. |

| Unit | Value | Meaning |
| ---- | ----- | ------- |
| **tomCharacter** | 1 | Character. |
| **tomWord** | 2 | Word. |
| **tomSentence** | 3 | Sentence. |
| **tomParagraph** | 4 | Paragraph. |
| **tomLine** | 5 | Line (on display). |
| **tomStory** | 6 | Story. |
| **tomScreen** | 7 | Screen (as for PAGE UP/PAGE DOWN). |
| **tomSection** | 8 | Section. |
| **tomColumn** | 9 | Table column. |
| **tomRow** | 10 | Table row. |
| **tomWindow** | 11 | Upper-left or lower-right of the window. |
| **tomCell** | 12 | Table cell. |
| **tomCharFormat** | 13 | Run of constant character formatting. |
| **tomParaFormat** | 14 | Run of constant paragraph formatting. |
| **tomTable** | 15 | Table. |
| **tomObject** | 16 | Embedded object. |

#### Return value

The actual number of units that the end is moved.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_NOTIMPL** | Unit is not supported. |
| **S_FALSE** | Failure for some other reason. |

---

## <a name="moveend"></a>MoveEnd

Moves the end position of the range.

```
FUNCTION MoveEnd (BYVAL Unit AS LONG = tomCharacter, BYVAL Count AS LONG = 1) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Unit* | The units by which to move the end of the range. The default value is **tomCharacter**. For a list of *Unit* values, see the table below. |
| *Count* | The number of units to move past. The default value is 1. If *Count* is greater than zero, motion is forward—toward the end of the story—and if *Count* is less than zero, motion is backward—toward the beginning. If *Count* is zero, the end position is unchanged. |

#### Return value

The actual number of units that the end position of the range is moved past.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_NOTIMPL** | Unit is not supported. |
| **S_FALSE** | Failure for some other reason. |

---

## <a name="movewhile"></a>MoveWhile

Starts at a specified end of a range and searches while the characters belong to the set specified by *Cset* and while the number of characters is less than or equal to *Count*. The range is collapsed to an insertion point when a non-matching character is found.

```
FUNCTION MoveWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG = tomForward) AS LONG
```

Starts at a specified end of a range and searches while the characters belong to the set specified by *Cset* and while the number of characters is less than or equal to *Count*. The range is collapsed to an insertion point when a non-matching character is found.

| Parameter | Description |
| --------- | ----------- |
| *Cset* | The character set to use in the match. This could be an explicit string of characters or a character-set index. For more information, see [Character Match Sets](https://learn.microsoft.com/en-us/windows/win32/controls/about-text-object-model#character-match-sets). |
| *Count* | Maximum number of characters to move past. The default value is **tomForward**, which searches to the end of the story. If *Count* is less than zero, the search starts at the start position and goes backward — toward the beginning of the story. If *Count* is greater than zero, the search starts at the end position and goes forward — toward the end of the story. |

#### Return value

The actual count of characters end is moved. 

#### Return code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | *Cset* is not valid. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

The motion described by **MoveUntil** is logical rather than geometric. That is, motion is toward the end or toward the start of a story. Depending on the language, moving to the end of the story could be moving left or moving right.

For more information, see the discussion in **ITextRange** and the **Remarks** section of **Move**.

The **MoveWhile** method is similar to **MoveUntil**, but **MoveWhile** searches as long as it finds members of the set specified by *Cset*, and there is no additional increment to the returned value.

The **MoveStartWhile** and **MoveEndWhile** methods move the start and end, respectively, just past all contiguous characters that are found in set of characters specified by the *Cset* parameter.

The following code illustrates how to initialize and use the **VARIANT** argument for matching a span of digits in the range *pRange*.

```
DIM varg AS VARIANT
varg.vt = VT_I4
varg.lVal = C1_DIGIT
DIM Delta AS LONG = pRange.MoveWhile(@varg, tomForward)   ' // Move IP past span of digits
```

Alternatively, an explicit string could be used, as in the following sample.

```
DIM varg AS VARIANT
varg.vt = VT_BSTR
varg.bstr = SysAllocString("0123456789")
DIM Delta AS LONG = pRange.MoveWhile(@varg, tomForward)   ' // Move IP past span of digits
```

## <a name="movestartwhile"></a>MoveStartWhile

Moves the start position of the range either *Count* characters, or just past all contiguous characters that are found in the set of characters specified by *Cset*, whichever is less.

```
FUNCTION MoveStartWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG = tomForward) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Cset* | The character set to use in the match. This could be an explicit string of characters or a character-set index. For more information, see [Character Match Sets](https://learn.microsoft.com/en-us/windows/win32/controls/about-text-object-model#character-match-sets). |
| *Count* | Maximum number of characters to move past. The default value is **tomForward**, which searches to the end of the story. If *Count* is greater than zero, the search is forward—toward the end of the story—and if *Count* is less than zero, search is backward—toward the beginning. If *Count* is zero, the start position is unchanged. |

#### Return value

The actual count of characters that the start position is moved.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | *Cset* is not valid. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

If the new start follows the old end, the new end is set equal to the new start.

The motion described by **MoveStartWhile** is logical rather than geometric. That is, motion is toward the end or toward the start of a story. Depending on the language, moving to the end of the story could be moving left or moving right.

For more information, see **Move**.

---

## <a name="moveendwhile"></a>MoveEndWhile

Moves the end of the range either *Count* characters or just past all contiguous characters that are found in the set of characters specified by *Cset*, whichever is less.

```
FUNCTION MoveEndWhile (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG = tomForward) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Cset* | The character set to use in the match. This could be an explicit string of characters or a character-set index. For more information, see [Character Match Sets](https://learn.microsoft.com/en-us/windows/win32/controls/about-text-object-model#character-match-sets). |
| *Count* | Maximum number of characters to move past. The default value is **tomForward**, which searches to the end of the story. If *Count* is greater than zero, the search moves forward (toward the end of the story). If *Count* is less than zero, the search moves backward (toward the beginning of the story). If *Count* is zero, the end position is unchanged. |

#### Return value

The actual number of characters that the end is moved.

#### Return code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | *Cset* is not valid. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

If the new end precedes the old start, the new start is set equal to the new end.

The motion described by **MoveEndWhile** is logical rather than geometric. That is, motion is toward the end or toward the start of a story. Depending on the language, moving to the end of the story could be moving left or moving right.

For more information, see **Move**.

---

## <a name="moveuntil"></a>MoveUntil

Searches up to *Count* characters for the first character in the set of characters specified by *Cset*. If a character is found, the range is collapsed to that point. The start of the search and the direction are also specified by *Count*.

```
FUNCTION MoveUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG = tomForward) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Cset* | The character set used in the match. This could be an explicit string of characters or a character-set index. For more information, see [Character Match Sets](https://learn.microsoft.com/en-us/windows/win32/controls/about-text-object-model#character-match-sets). |
| *Count* | Maximum number of characters to move past. The default value is **tomForward**, which searches to the end of the story. If *Count* is less than zero, the search is backward starting at the start position. If *Count* is greater than zero, the search is forward starting at the end. |

#### Return value

The number of characters the insertion point is moved, plus 1 for a match if *Count* is greater than zero, and –1 for a match if *Count* less than zero.

#### Return code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | *Cset* is not valid. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

If no character is matched, the range is unchanged.

The motion described by **MoveUntil** is logical rather than geometric. That is, motion is toward the end or toward the start of a story. Depending on the language, moving to the end of the story could be moving left or moving right.

For more information, see the Remarks section of **Move**.

The **MoveStartUntil** and **MoveEndUntil** methods move the start and end, respectively, until it finds the first character that is also in the set specified by the *Cset* parameter.

The **MoveUntil** method is similar to **MoveWhile**, but there are two differences. First, **MoveUntil** moves an insertion point until it finds the first character that belongs to the character set specified by *Cset*. Second, in **MoveUntil** the character matched counts as an additional character in the value returned. This lets you know that the character at one end of the range or the other belongs to the *Cset* even though the insertion point stays at one of the range ends.

For example, suppose the range, *pRange*, is an insertion point. To see if the character at *pRange* (that is, given by *pRange.GetChar()*) is in *Cset*, call

```
pRange.MoveUntil(Cset, 1)
```
If the character is in *Cset*, the return value is 1 and the insertion point does not move. Similarly, to see if the character preceding *pRange* is in Cset, call

```
pRange.MoveUntil(Cset, -1)
```

If the character is in *Cset*, the return value is –1.

The following subroutine prints all numbers in the story identified by the range, *pRange*.

```
SUB PrintNumbers (BYVAL pRange AS ITextRange2 PTR)
   pRange.SetRange(0, 0)   ' // pRange = insertion point at start of story
   WHILE pRange.MoveUntil(C1_DIGIT, tomForward)   ' // Move pRange to 1st digit in next number
      DIM Delta AS LONG = pRange.MoveEndWhile(C1_DIGIT, tomForward)   ' // Select number (span of digits)
      ' PRINT ---   ' // Print it
   WEND
END SUB
```

## <a name="movestartuntil"></a>MoveStartUntil

Moves the start position of the range the position of the first character found that is in the set of characters specified by *Cset*, provided that the character is found within *Count* characters of the start position.

```
FUNCTION MoveStartUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG = tomForward) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Cset* | The character set to use in the match. This could be an explicit string of characters or a character-set index. For more information, see [Character Match Sets](https://learn.microsoft.com/en-us/windows/win32/controls/about-text-object-model#character-match-sets). |
| *Count* | Maximum number of characters to move past. The default value is **tomForward**, which searches to the end of the story. If *Count* is greater than zero, the search is forward—toward the end of the story—and if *Count* is less than zero, search is backward—toward the beginning. If *Count* is zero, the start position is unchanged. |

#### Return value

The actual number of characters the start of the range is moved, plus 1 for a match if *Count* is greater than zero, and –1 for a match if *Count* is less than zero. 

#### Return code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | *Cset* is not valid. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

If no character from *Cset* is found within *Count* positions of the start position, the range is left unchanged.

If the new start follows the old end, the new end is set equal to the new start.

The motion described by **MoveStartUntil** is logical rather than geometric. That is, motion is toward the end or toward the start of a story. Depending on the language, moving to the end of the story could be moving left or moving right.

For more information, see **Move**.

---

## <a name="moveenduntil"></a>MoveEndUntil

Moves the range's end to the character position of the first character found that is in the set of characters specified by *Cset*, provided that the character is found within *Count* characters of the range's end.

```
FUNCTION MoveEndUntil (BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG = tomForward) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Cset* | The character set to use in the match. This could be an explicit string of characters or a character-set index. For more information, see [Character Match Sets](https://learn.microsoft.com/en-us/windows/win32/controls/about-text-object-model#character-match-sets). |
| *Count* | Maximum number of characters to move past. The default value is **tomForward**, which searches to the end of the story. If *Count* is greater than zero, the search moves forward (toward the end of the story). If *Count* is less than zero, the search moves backward (toward the beginning of the story). If *Count* is zero, the end position is unchanged. |

#### Return value

The actual number of characters that the range end is moved, plus 1 for a match if *Count* is greater than zero, and –1 for a match if *Count* is less than zero.

#### Return code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | *Cset* is not valid. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

If no character from the set specified by *Cset* is found within *Count* positions of the range's end, the range is left unchanged. If the new end precedes the old start, the new start is set equal to the new end.

The motion described by **MoveEndUntil** is logical rather than geometric. That is, motion is toward the end or toward the start of a story. Depending on the language, moving to the end of the story could be moving left or moving right.

For more information, see **Move**.

---

## <a name="findtext"></a>FindText

Searches up to *Count* characters for the text given by *cbs*. The starting position and direction are also specified by *Count*, and the matching criteria are given by *Flags*.

```
FUNCTION FindText (BYREF cbs AS CBSTR, BYVAL Count AS LONG = tomForward, BYVAL Flags AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *cbs* | String to find. |
| *Count* | Maximum number of characters to search. It can be one of the following.<br>- *tomForward*. Searches to the end of the story. This is the default value.<br>- *n* (greater than 0). Searches forward for *n* chars, starting from *cpFirst*. If the range itself matches *cbs*, another search is attempted from *cpFirst* + 1.<br>- *n* (less than 0). Searches backward for *n* chars, starting from *cpLim*. If the range itself matches *cbs*, another search is attempted from *cpLim*– 1.<br>- *0* (degenerate range). Search begins after the range.<br>- *0* (nondegenerate range)	Search is limited to the range.<br>In all cases, if a string is found, the range limits are changed to be those of the matched string and the method returns the length of the string. If the string is not found, the range remains unchanged and the method returns zero. |
| *Flags* | Flags governing comparisons. It can be 0 (the default) or any combination of the following values. |

| Unit | Value | Meaning |
| ---- | ----- | ------- |
| **tomMatchWord** | 2 | Matches whole words. |
| **tomMatchCase** | 4 | Matches case. |
| **tomMatchPattern** | 8 | Matches regular expressions. |

#### Return value

The length of string matched.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**.

The **FindText** method can also match special characters by using a caret (^) followed by a special letter. For a list of special characters, see the Special list available in the Microsoft Word **Find and Replace** dialog box. For example, *^p* matches the next paragraph mark. Note, *^c* can be used to represent the Clipboard contents in the string to be replaced. Thus, using *^c* in the find string enables you to search for rich text. For more details, see the Word Help files.

As a comparison with the **FindText** method, the **FindTextStart** method searches forward or backward from the range's Start *cp*, and the **FindTextEnd** method searches forward or backward from the range's End *cp*. For more details, see the descriptions of these methods.

The following are several code snippets that show the **FindText** methods.

**Example #1**. The following code prints all the /* ... */ comments in a story identified by the range *pRange*.

```
SUB PrintComments (pRange AS ITextRange PTR)
    pRange.SetRange(0, 0)    ' pRange = insertion point at start of story
    DO WHILE pRange.FindText("/*") AND pRange.FindTextEnd("*/"   'Select comment
        pRange.MoveStart(tomCharacter, 2)   ' But do not include the opening
        pRange.MoveEnd(tomCharacter, -2)   '  or closing comment brackets
        ' PRINT --- Print ' Show the folks
    LOOP
END SUB
```
Instead of these comments being printed, they could be inserted into another edit instance and saved to a file, or they could be inserted into separate cells in a table or spreadsheet.

To print all lines containing one or more occurrences of the word "laser", replace the loop by the following code:

```
WHILE pRange.FindText("laser")   // Select next occurrence of "laser"
   pRange.Expand(tomLine)    ' // Select enclosing line
   ' PRINT ---   ' // Print the line
WEND
```

**Example #2**. The following code prints a telephone list, given a story that contains an address list. The address list entries are separated by two or more paragraph marks, and each entry has the following form.

```
Person/Business Name
Address (one or more lines)
(area code) telephone number 
Note the use of the character ^p in the FindText string argument to locate a pair of consecutive paragraph marks.
```

```
SUB PrintTelephoneList (pRange AS ITextRange PTR)
    pRange.SetRange(0, 0)   ' // pRange = insertion point at start of story
    pRange.MoveWhile(C1_WHITE)   ' // Bypass any initial white space
    DO
        pRange.EndOf tomParagraph, 1     // Select next para (line): has name
        ' PRINT  ---   ' // Print it
        DO
            pRange.MoveWhile(C1_SPACE)   ' // Bypass possible space chars
            IF pRange.Char = ASC("(") THEN EXIT DO   ' // Look for start of telephone #
        LOOP WHILE pRange.Move(tomParagraph)   ' // Go to next paragraph
        pRange.EndOf(tomParagraph, 1)   ' // Select line with telephone number
        ' PRINT ---   ' // Print it
    LOOP WHILE pRange.FindText("^p^p")   ' // Find two consecutive para marks
END FUB
```

**Example #3**. The following subroutine replaces all occurrences of the string, *str1*, in a range by *str2*:

```
SUB Replace (BYVAL tr AS ITextRange PTR, BYREF str1 AS STRING, BYREF str2 AS STRING)
    DIM r AS ITextRange PTR
    r = tr.Duplicate   ' // Copy tr parameters to r
    r.End = r.Start   ' // Convert to insertion point at Start
    WHILE r.FindText(str1, tr.End - r.End)   ' // Match next occurrence of str1
        pRange.SetText(str2)   ' // Replace it with str2
    WEND   ' // Iterate till no more matches
END SUB
```

**Example #4**. The following line of code inserts a blank before the first occurrence of a right parenthesis, "(", that follows an occurrence of HRESULT.
```
    If pRange.FindText("HRESULT") AND pRange.FindText("(") Then pRange.SetText(" (")
```
To do this for all such occurrences, change the IF into a WHILE/WEND loop in the above line of code.

---

# <a name="findtextstart"></a>FindTextStart

Searches up to *Count* characters for the string, *cbs*, starting at the range's Start *cp (cpFirst)*. The search is subject to the comparison parameter, *Flags*. If the string is found, the Start *cp* is changed to the matched string, and the method returns the length of the string. If the string is not found, the range is unchanged, and the method returns zero.

```
FUNCTION FindTextStart (BYREF cbs AS CBSTR, BYVAL Count AS LONG = tomForward, BYVAL Flags AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *cbs* | The string to search for. |
| *Count* | Maximum number of characters to search. It can be one of the following.<br>- *tomForward*. Searches to the end of the story. This is the default value.<br>- *n* (greater than 0). Searches forward for *n* chars, starting from *cpFirst*. If the range itself matches *cbs*, another search is attempted from *cpFirst* + 1.<br>- *n* (less than 0). Searches backward for *n* chars, starting from *cpLim*. |
| *Flags* | Flags governing the comparisons. It can be zero (the default) or any combination of the following values. |

| Unit | Value | Meaning |
| ---- | ----- | ------- |
| **tomMatchWord** | 2 | Matches whole words. |
| **tomMatchCase** | 4 | Matches case. |
| **tomMatchPattern** | 8 | Matches regular expressions. |

#### Return value

The length of the matched string.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**.

## <a name="findtextend"></a>FindTextEnd

Searches up to *Count* characters for the string, *cbs*, starting from the range's End *cp*. The search is subject to the comparison parameter, *Flags*. If the string is found, the End *cp* is changed to be the end of the matched string, and the method returns the length of the string. If the string is not found, the range is unchanged and the method returns zero.

```
FUNCTION FindTextEnd (BYREF cbs AS CBSTR, BYVAL Count AS LONG = tomForward, BYVAL Flags AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *cbs* | The string to search for. |
| *Count* | Maximum number of characters to search. It can be one of the following.<br>- *tomForward*. Searches to the end of the story. This is the default value.<br>- *n* (greater than 0). Searches forward for *n* chars, starting from *cpFirst*. If the range itself matches *cbs*, another search is attempted from *cpFirst* + 1.<br>- *n* (less than 0). Searches backward for *n* chars, starting from *cpLim*. |
| *Flags* | Flags governing the comparisons. It can be zero (the default) or any combination of the following values. |

| Unit | Value | Meaning |
| ---- | ----- | ------- |
| **tomMatchWord** | 2 | Matches whole words. |
| **tomMatchCase** | 4 | Matches case. |
| **tomMatchPattern** | 8 | Matches regular expressions. |

#### Return value

The length of the matched string.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**.

---

## <a name="delete_"></a>Delete_

Mimics the DELETE and BACKSPACE keys, with and without the CTRL key depressed.

```
FUNCTION Delete_ (BYVAL Unit AS LONG = tomCharacter, BYVAL Count AS LONG = 1) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit to use. *Unit* can be **tomCharacter** (the default value) or **tomWord**. |
| *Count* | Number of *Units* to delete. If *Count* = zero, it deletes the text in the range only. If *Count* is greater than zero, **Delete_** acts as if the DELETE key was pressed *Count* times. If *Count* is less than zero, it acts as if the BACKSPACE key was pressed *Count* times. The default value is 1. For more information, see the **Remarks**. |

#### Return value

The count of units deleted. Deleting the text in a nondegenerate range counts as one Unit.

#### Result code

If successful, **GetLastResult** returns **S_OK**. Otherwise it returns one of the following values.

| Result code | Description |
| ----------- | ----------- |
| **E_ACCESSDENIED** | Text is write-protected. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

If *Count* = zero, this method deletes the text in the range, that is, it deletes nothing if the range is only an insertion point.

If *Count* is not zero, and the range is an insertion point (that is, degenerate), | *Count*| (absolute value of *Count*) *Units* are deleted in the logical direction given by the sign of *Count*, where a positive value is the direction toward the end of the story, and a negative value is toward the start of the story.

If *Count* is not zero, and the range is nondegenerate (contains text), the text in the range is deleted (regardless of the values of *Unit* and *Count*), thereby creating an insertion point. Then, | *Count*| - 1   Units are deleted in the logical direction given by the sign of Count.

The text in the range can also be deleted by assigning a null string to the range. However, Delete__ does not require allocating a **BSTR**.

Deleting the end-of-paragraph mark (CR) results in the special behavior of the Microsoft Word UI. Four cases are of particular interest:

- If you delete just the CR but the paragraph includes text, then the CR is deleted, and the following paragraph gets the same paragraph formatting as current one.
- If you delete the CR as well as some, but not all, of the characters in the following paragraph, the characters left over from the current paragraph get the paragraph formatting of the following paragraph.
- If you select to the end of a paragraph, but not the whole paragraph, the CR is not deleted.
- If you delete the whole paragraph (from the beginning through the CR), you delete the CR as well (unless it is the final CR in the file).

---

## <a name="cut"></a>Cut

Cuts the plain or rich text to a data object or to the Clipboard, depending on the *pVar* parameter.

```
FUNCTION Cut (BYVAL pVar AS VARIANT PTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *pVar* | The cut text. pVar->ppunkVal is the out parameter for an **IDataObject** object, provided that the following conditions exist:<br>- pVar->vt = (VT_UNKNOWN OR VT_BYREF)<br>- pVar is not null<br>- pVar->ppunkVal is not null<br>Otherwise, the clipboard is used. |

---

## <a name="copy"></a>Copy

Copies the text to a data object.

```
FUNCTION Copy (BYVAL pVar AS VARIANT PTR) AS HRESULT
```

#### Return value

If successful, it returns **S_OK**. Otherwise, it returns **E_OUTOFMEMORY**.

#### Remarks

The **Cut**, **Copy**, and **Paste** methods let you perform the usual **Cut**, **Copy**, and **Paste** operations on a range object using an **IDataObject**, thereby not changing the contents of the clipboard. Among clipboard formats typically supported are **CF_TEXT** and **CF_RTF**. In addition, private clipboard formats can be used to reference a text solution's own internal rich text formats.

To copy and replace plain text, you can use the **GetText**  and **SetText**  methods. To copy formatted text from range r1 to range r2 without using the clipboard, you can use **Copy** and **Paste** and also the **GetFormattedText** and **SetFormattedText** methods.

---

## <a name="paste"></a>Paste

Pastes text from a specified data object.

```
FUNCTION Paste (BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG = 0) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *pVar* | The **IDataObject** to paste. However, the contents of the clipboard are used if any of the following are true.<br>- *pVar* is null<br>- *pVar* punkVal is null<br>- *pVar* is not VT_UNKNOWN<br>- *pVar* punkVal does not return an **IDataObject** when queried for one |
| *Format* | The clipboard format to use in the paste operation. Zero is best format, which usually is RTF, but **CF_UNICODETEXT** and other formats are also possible. The default value is zero. For more information, see [Clipboard Formats](https://learn.microsoft.com/en-us/windows/win32/dataxchg/clipboard-formats). |

#### Return value

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_ACCESSDENIED** | Destination is write-protected. |
| **E_OUTOFMEMORY** | Destination cannot contain the text to be pasted. |

#### Remarks
For more information, see **Copy**.

---

## <a name="canpaste"></a>CanPaste

Determines if a data object can be pasted, using a specified format, into the current range.

```
FUNCTION CanPaste (BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *pVar* | The **IDataObject** to be pasted. However, the Clipboard contents are checked for pasting if any of the following are true:<br>- *pVar* is null<br>- *pVar->punkVal* is null<br>- *pVar->vt* is not VT_UNKNOWN<br>- *pVar->punkVal* does not return an **IDataObject** object when queried for one |
| *Format* | Clipboard format that is used. Zero represents the best format, which usually is RTF, but **CF_UNICODETEXT** and other formats are also possible. The default value is zero. |

#### Return value

A **tomBool** value that is tomTrue only if the data object identified by pVar can be pasted, using the specified format, into the range.

#### Result code

The method returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **S_OK** | The clipboard contents or IDataObject can be pasted. |
| **S_FALSE** | The clipboard contents or **IDataObject** cannot be pasted. |

---

## <a name="canedit"></a>CanEdit

Determines whether the specified range can be edited.

```
FUNCTION CanEdit () AS LONG
```

#### Return value

A tomBool value indicating whether the range can be edited. It is **tomTrue** only if the specified range can be edited.

#### Result code

If the range can be edited, the method succeeds and **GetLastREsult** returns S_OK. If the range cannot be edited, the method fails and  **GetLastREsult** returns **S_FALSE**.

#### Remarks

The range cannot be edited if any part of it is protected or if the document is read-only.

---

## <a name="changecase"></a>ChangeCase

Changes the case of letters in this range according to the *nType* parameter.

```
FUNCTION ChangeCase (BYVAL nType AS LONG  = tomLowerCase) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *nType* | Type of case change. The default value is **tomLower**. |

| Type | Value | Meaning |
| ---- | ----- | ------- |
| **tomLowerCase** | 0 | Sets all text to lowercase. |
| **tomUpperCase** | 1 | Sets all text to uppercase. |
| **tomTitleCase** | 2 | Capitalizes the first letter of each word. |
| **tomSentenceCase** | 4 | Capitalizes the first letter of each sentence. |
| **tomToggleCase** | 5 | Toggles the case of each letter. |

#### Return value
This method returns an **HRESULT** value. If successful, it returns **S_OK**. Otherwise, it returns **S_FALSE**.

---

## <a name="getpoint"></a>GetPoint

Retrieves screen coordinates for the start or end character position in the text range, along with the intra-line position.

```
FUNCTION GetPoint (BYVAL nType AS LONG = tomStart + TA_BASELINE + TA_LEFT, _
   BYVAL px AS LONG PTR, BYVAL py AS LONG PTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *nType* | Flag that indicates the position to retrieve. This parameter can include one value from each of the tables below. The default value is tomStart + TA_BASELINE + TA_LEFT. |
| *px* | The x-coordinate. |
| *py* | The y-coordinate. |

| Constant | Value | Meaning |
| -------- | ----- | ----------- |
| **tomAllowOffClient** | 512 | Allow points outside of the client area. |
| **tomClientCoord** | 256 | Use client coordinates instead of screen coordinates. |
| **tomObjectArg** | 2048 | Get a point inside an inline object argument; for example, inside the numerator of a fraction. |
| **tomTransform** | 1024 | Transform coordinates using a world transform (XFORM) supplied by the host application. |

Use one of the following values to indicate the start or end of the range.

| Constant | Value | Meaning |
| -------- | ----- | ----------- |
| **tomStart** | 32 | The start of text range. |
| **tomEnd** | 0 | The end of a text range. |

Use one of the following values to indicate the vertical position.

| Constant | Value | Meaning |
| -------- | ----- | ----------- |
| **TA_TOP** | 0 | Top edge of the bounding rectangle. |
| **TA_BASELINE** | 24 | Base line of the text. |
| **TA_BOTTOM** | 8 | Bottom edge of the bounding rectangle. |

Use one of the following values to indicate the horizontal position.

| Constant | Value | Meaning |
| -------- | ----- | ----------- |
| **TA_LEFT** | 0 | Left edge of the bounding rectangle. |
| **TA_CENTER** | 6 | Center of the bounding rectangle. |
| **TA_RIGHT** | 2 | ight edge of the bounding rectangle. |

The method returns an **HRESULT** value. If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following error codes.

| Return code | Description |
| --------- | ----------- |
| **E_INVALIDARG** | Either px or py is null. |
| **S_FALSE** | Failure for some other reason. |

### Remarks

The **GetPoint** method gives **ITextRange** the ability to emulate UI-pointer commands; it is also handy for accessibility purposes.

---

## <a name="setpoint"></a>SetPoint

Changes the range based on a specified point at or up through (depending on Extend) the point (x, y) aligned according to *nType*.

```
FUNCTION SetPoint (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG, BYVAL Extend AS LONG = 0) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *x* | Horizontal coordinate of the specified point, in absolute screen coordinates. |
| *y* | Vertical coordinate of the specified point, in absolute screen coordinates. |
| *nType* | The end to move to the specified point. It can be one of the following.<br>**tomStart**.	Move the start of range.<br>**tomEnd**. Move the end of range. |
| *Extend* | How to set the endpoints of the range. If **Extend** is zero (the default), the range is an insertion point at the specified point (or at the nearest point with selectable text). If **Extend** is 1, the end specified by *nType* is moved to the point and the other end is left alone. |

#### Return value

The method returns an **HRESULT** value. If the method succeeds, it returns **S_OK**. If the method fails, it returns **S_FALSE**.

#### Remarks
An application can use the specified point in the **WindowFromPoint** function to get the handle of the window, which usually can be used to find the client-rectangle coordinates (although a notable exception is with [Windowless Controls](https://learn.microsoft.com/en-us/windows/win32/controls/windowless-rich-edit-controls)).

---

## <a name="scrollintoview"></a>ScrollIntoView

Scrolls the specified range into view.

```
FUNCTION ScrollIntoView (BYVAL Value AS LONG = tomStart) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | Flag specifying the end to scroll into view. It can be one of the following. |


| Flag     | Value | Meaning |
| -------- | ----- | ----------- |
| **tomEnd** | 0 | Scrolls the end character position to appear on the bottom line. |
| **tomStart** | 32 | Scrolls the start character position to appear on the top line. (Default value). |
| **tomNoUpScroll** | &h10000 | Horizontal scrolling is disabled. |
| **tomNoVpScroll** | &h40000 | Vertical scrolling is disabled. |

#### Return value

The method returns an **HRESULT** value. If the method succeeds, it returns **S_OK**. If the method fails, it returns **S_FALSE**.

---

## <a name="getembeddedobject"></a>GetEmbeddedObject

Retrieves a pointer to the embedded object at the start of the specified range, that is, at *cpFirst*. The range must either be an insertion point or it must select only the embedded object.

```
FUNCTION GetEmbeddedObject () AS IUnknown PTR
```

#### Return value

The pointer to the object.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns **S_FALSE**.

---

## <a name="getflags"></a>GetFlags

Gets the text selection flags.

```
FUNCTION GetFlags () AS LONG
```

#### Return value

Any combination of the following selection flags.

| Selection flag | Value | Meaning |
| -------------- | ----- | ----------- |
| **tomSelStartActive** | 1 | Start end is active. |
| **tomSelAtEOL** | 2 | For degenerate selections, the ambiguous character position corresponding to both the beginning of a line and the end of the preceding line should have the caret displayed at the end of the preceding line. |
| **tomSelOvertype** | 4 | Insert/Overtype mode is set to overtype. |
| **tomSelActive** | 8 | Selection is active. |
| **tomSelReplace** | 16 | Typing and pasting replaces selection. |

Each of the table values is binary. Thus, if any value is not set, the text selection has the opposite property.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**.

---

## <a name="setflags"></a>SetFlags

Sets the text selection flags.

```
FUNCTION SetFlags (BYVAL Flags AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Flags* | Flag specifying the end to scroll into view. It can be one of the following. |

| Selection flag | Value | Meaning |
| -------------- | ----- | ----------- |
| **tomSelStartActive** | 1 | Start end is active. |
| **tomSelAtEOL** | 2 | For degenerate selections, the ambiguous character position corresponding to both the beginning of a line and the end of the preceding line should have the caret displayed at the end of the preceding line. |
| **tomSelOvertype** | 4 | Insert/Overtype mode is set to overtype. |
| **tomSelActive** | 8 | Selection is active. |
| **tomSelReplace** | 16 | Typing and pasting replaces selection. |

Each of the table values is binary. Thus, if any value is not set, the text selection has the opposite property.

#### Return value

The method returns **S_OK**.

To make sure that the start end is active and that the ambiguous character position is displayed at the end of the line, execute the following code:

```
selection.Flags = tomSelStartActive + tomSelAtEOL
```

The *Flags* property is useful because an **ITextRange** object can select itself. With **SetFlags**, you can change the active end from the default value of End, select the caret position for an ambiguous character position, or change the Insert/Overtype mode.

---

## <a name="gettype"></a>GetType

Gets the type of text selection.

```
FUNCTION GetType () AS LONG
```

#### Return value

The selection type. Can be one of the values in the following table.

| Selection type | Value | Meaning |
| -------------- | ----- | ----------- |
| **tomNoSelection** | 0 | No selection and no insertion point. |
| **tomSelectionIP** | 1 | Insertion point. |
| **tomSelectionNormal** | 2 | Single nondegenerate range. |
| **tomSelectionFrame** | 3 | Frame. |
| **tomSelectionColumn** | 4 | Table column. |
| **tomSelectionRow** | 5 | Table rows. |
| **tomSelectionBlock** | 6 | Block selection. |
| **tomSelectionInlineShape** | 7 | Picture. |
| **tomSelectionShape** | 8 | Shape. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**.

---

## <a name="moveleft"></a>MoveLeft

Generalizes the functionality of the Left Arrow key.

```
FUNCTION MoveLeft (BYVAL Unit AS LONG = tomCharacter, BYVAL Count AS LONG = 1, _
   BYVAL Extend AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit to use. It can be one of the following (see table below). |
| *Count* | Number of *Units* to move past. The default value is 1. If *Count* is less than zero, movement is to the right. |
| *Extend* | Flag that indicates how to change the selection. If *Extend* is zero (or **tomMove**), the method collapses the selection to an insertion point at the active end and then moves it. If *Extend* is 1 (or **tomExtend**), the method moves the active end and leaves the other end alone. The default value is zero. A nonzero *Extend* value corresponds to the Shift key being pressed in addition to the key combination described in *Unit*. |

| Value | Corresponding key combination | Meaning |
| ----- | ----------------------------- | ------- |
| **tomCharacter** | Left Arrow | Move one character position to the left. This is the default. |
| **tomWord** | Ctrl+Left Arrow | Move one word to the left. |
 
Note: If *Count* is less than zero, movement is to the right.

#### Return value

The actual count of units the insertion point or active end is moved left. Collapsing the selection, when *Extend* is 0, counts as one unit.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Return code | Description |
| -------------- | ----------- |
| **E_INVALIDARG** | Unit is not valid. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

The WordBasic move methods like **CharRight**, **CharLeft**, **WordRight**, and **WordLeft** are hybrids that can do four things that are closely related to the standard arrow-key edit behavior:

- Move the current insertion point if there's no selection.
- Move the active end of the selection if there is a selection.
- Turn an insertion point into a selection and vice versa.
- Return a Boolean stating if movement occurred.

The *Extend* argument of **MoveLeft** and **MoveRight** enables you to be consistent with the three items above. For example, given a selection consisting of a single range, you have the following correspondences (for left-to-right characters).

| ITextSelection | WordBasic | Function |
| -------------- | --------- | -------- |
| selection.MoveRight tomWord, 1, 1 | WordRight 1,1 | Moves active end one word right. |
| selection.MoveLeft tomCharacter, 1, 1 | CharLeft 1,1 | Moves active end one character left. |
 
As in WordBasic, if *Count* is less than zero, the meanings of left and right are interchanged, that is MoveLeft (Unit, Count, Extend) is equivalent to MoveRight (Unit, -Count, Extend).

Similar to WordBasic and the Left Arrow key UI behavior, calling MoveLeft(Unit, Count) on a degenerate selection moves the insertion point the specified number of *Units*. On a degenerate range, calling MoveLeft(Unit, Count, 1) where *Count* is greater than zero causes the range to become nondegenerate with the left end being the active end.

When *Extend* is **tomExtend** (or is nonzero), **MoveLeft** moves only the active end of the selection, leaving the other end where it is. However, if **Extend** equals zero and the selection starts as a nondegenerate range, MoveLeft(Unit, Count) where *Count* is greater than zero moves the active end Count - 1 units left, and then moves the other end to the active end. In other words, it makes an insertion point at the active end. Collapsing the range counts as one unit. Thus, MoveLeft(tomCharacter) converts a nondegenerate selection into a degenerate one at the selection's left end. Here, *Count* has the default value of 1 and *Extend* has the default value of zero. This example corresponds to pressing the Left Arrow key. **MoveLeft** and **MoveRight** are related to the **ITextRange** move methods, but differ in that they explicitly use the active end (the end moved by pressing the Shift key).

---

## <a name="moveright"></a>MoveRight

Generalizes the functionality of the Right Arrow key.

```
FUNCTION MoveRight (BYVAL Unit AS LONG = tomCharacter, BYVAL Count AS LONG = 1, _
   BYVAL Extend AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit to use. It can be one of the following (sse table below). |
| *Count* | Number of *Units* to move past. The default value is 1. If Count is less than zero, movement is to the left. |
| *Extend* | Flag that indicates how to change the selection. If *Extend* is zero (or **tomMove**), the method collapses the selection to an insertion point at the active end and then moves it. If *Extend* is 1 (or **tomExtend**), the method moves the active end and leaves the other end alone. The default value is zero. A nonzero *Extend* value corresponds to the Shift key being pressed in addition to the key combination described in *Unit*. |

| Value | Corresponding key combination | Meaning |
| ----- | ----------------------------- | ------- |
| **tomCharacter** | Right Arrow | Move one character position to the right. This is the default. |
| **tomWord** | Ctrl+Right Arrow | Move one word to the right. |
 
Note: If *Count* is less than zero, movement is to the left.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Return code | Description |
| -------------- | ----------- |
| **E_INVALIDARG** | Unit is not valid. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

Microsoft WordBasic move methods like **CharRigh**t, **CharLeft**, **WordRight**, and **WordLeft** are hybrids that can do four things that are closely related to the standard arrow-key edit behavior:

- Move the current insertion point if there's no selection.
- Move the active end of the selection if there is a selection.
- Turn an insertion point into a selection and vice versa.
- Return a Boolean stating if movement occurred.

The *Extend* argument of **MoveLeft** and **MoveRight** enables you to be consistent with the three items above. For example, given a selection, s, consisting of a single range, you have the following correspondences (for left-to-right characters).

| ITextSelection | WordBasic | Function |
| -------------- | --------- | -------- |
| s.MoveRight tomWord, 1, 1 | WordRight 1,1 | Moves active end one word right. |
| s.MoveLeft tomCharacter, 1, 1 | CharLeft 1,1 | Moves active end one character left. |

As in WordBasic, if *Count* is less than zero, the meanings of left and right are interchanged, that is MoveLeft (Unit, Count, Extend) is equivalent to MoveRight(Unit, -Count, Extend).

Similar to WordBasic and the Right Arrow key UI behavior, calling MoveRight(Unit, Count) on a degenerate selection moves the insertion point the specified number of units. On a degenerate range, calling MoveRight(Unit, Count, 1) where *Count* is greater than zero causes the range to become nondegenerate with the right end being the active end.

When *Extend* is **tomExtend** (or is nonzero), **MoveRight** moves only the active end of the selection, leaving the other end where it is. However, if Extend equals zero and the selection starts as a nondegenerate range, MoveRight(Unit, Count) where Count is greater than zero moves the active end Count - 1 units right, and then moves the other end to the active end. In other words, it makes an insertion point at the active end. Collapsing the range counts as one unit. Thus, MoveRight(tomCharacter) converts a nondegenerate selection into a degenerate one at the selection's right end. Here, *Count* has the default value of 1 and *Extend* has the default value of zero. This example corresponds to pressing the Right Arrow key. **MoveLeft** and **MoveRight** are related to the **ITextRange** move methods, but differ in that they explicitly use the active end (the end moved by pressing the Shift key).

#### Return value

The actual count of units the insertion point or active end is moved left. Collapsing the selection, when *Extend* is 0, counts as one unit.

---

## <a name="moveup"></a>MoveUp

Mimics the functionality of the Up Arrow and Page Up keys.

```
FUNCTION MoveUp (BYVAL Unit AS LONG = tomLine, BYVAL Count AS LONG = 1, BYVAL Extend AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit to use. It can be one of the following (sse table below). |
| *Count* | Number of *Units* to move past. The default value is 1. If *Count* is less than zero, movement is to the left. |
| *Extend* | Flag that indicates how to change the selection. If *Extend* is zero (or **tomMove**), the method collapses the selection to an insertion point at the active end and then moves it. If *Extend* is 1 (or **tomExtend**), the method moves the active end and leaves the other end alone. The default value is zero. A nonzero *Extend* value corresponds to the Shift key being pressed in addition to the key combination described in *Unit*. |

| Value | Corresponding key combination | Meaning |
| ----- | ----------------------------- | ------- |
| **tomLine** | Up Arrow | Moves up one line. This is the default. |
| **tomParagraph** | Ctrl+Up Arrow | Moves up one paragraph. |
| **tomScreen** | Page Up | Moves up one screen. |
| **tomWindow** | Ctrl+Page Up | Moves to first character in window. |

Note: If *Count* is less than zero, movement is to the left.

#### Return value

The actual count of units the insertion point or active end is moved down. Collapsing the selection counts as one unit.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Return code | Description |
| -------------- | ----------- |
| **E_INVALIDARG** | Unit is not valid. |
| **S_FALSE** | Failure for some other reason. |

---

## <a name="movedown"></a>MoveDown

Mimics the functionality of the Down Arrow and Page Down keys.

```
FUNCTION MoveDown (BYVAL Unit AS LONG = tomLine, BYVAL Count AS LONG = 1, BYVAL Extend AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit to use. It can be one of the following (sse table below). |
| *Count* | Number of *Units* to move past. The default value is 1. |
| *Extend* | Flag that indicates how to change the selection. If *Extend* is zero (or **tomMove**), the method collapses the selection to an insertion point and then moves. If *Extend* is 1 (or **tomExtend**), the method moves the active end and leaves the other end alone. The default value is zero. A nonzero **Extend** value corresponds to the Shift key being pressed in addition to the key combination described in *Unit*. |

| Value | Corresponding key combination | Meaning |
| ----- | ----------------------------- | ------- |
| **tomLine** | Down Arrow | Moves down one line. This is the default. |
| **tomParagraph** | Ctrl+Down Arrow | Moves down one paragraph. |
| **tomScreen** | Page Donw | Moves down one screen. |
| **tomWindow** | Ctrl+Page Down | Moves to last character in window. |

#### Return value

The actual count of units the insertion point or active end is moved down. Collapsing the selection counts as one unit. 

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Return code | Description |
| -------------- | ----------- |
| **E_INVALIDARG** | Unit is not valid. |
| **S_FALSE** | Failure for some other reason. |

---

## <a name="homekey"></a>HomeKey

Generalizes the functionality of the Home key.

```
FUNCTION HomeKey (BYVAL Unit AS LONG = tomLine, BYVAL Extend AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit to use in the Home key operation. It can take on one of the following values (see table below). |
| *Extend* | Flag that indicates how to change the selection. If *Extend* is zero (or **tomMove**), the method collapses the selection to an insertion point. If **Extend** is 1 (or **tomExtend**), the method moves the active end and leaves the other end alone. The default value is zero. |

| Value | Meaning |
| ----- | ------- |
| **tomLine** | Depending on *Extend*, it moves either the insertion point or the active end to the beginning of the first line in the selection. This is the default. |
| **tomStory** | Depending on *Extend*, it moves either the insertion point or the active end to the beginning of the first line in the story. |
| **tomColumn** | Depending on *Extend*, it moves either the insertion point or the active end to the beginning of the first column in the selection. This is available only if the TOM engine supports tables. |
| **tomRow** | Depending on *Extend*, it moves either the insertion point or the active end to the beginning of the first row in the selection. This is available only if the TOM engine supports tables. |

#### Return value

The count of characters that the insertion point or the active end is moved.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Return code | Description |
| -------------- | ----------- |
| **E_INVALIDARG** | Unit is not valid. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

The **HomeKey** and **EndKey** methods are used to mimic the standard Home/End key behavior.

**tomLine** mimics the Home or End key behavior without the Ctrl key pressed, while **tomStory** mimics the behavior with the Ctrl key pressed. Similarly, **tomMove** mimics the Home or End key behavior without the Shift key pressed, while **tomExtend** mimics the behavior with the Shift key pressed. So HomeKey(tomStory) converts the selection into an insertion point at the beginning of the associated story, while HomeKey(tomStory, tomExtend) moves the active end of the selection to the beginning of the story and leaves the other end where it was.

The **HomeKey** and **EndKey** methods are logical methods like the **Move** methods, rather than directional methods. Thus, they depend on the language that is involved. For example, in Arabic text, HomeKey moves to the right end of a line, whereas in English text, it moves to the left. Thus, **HomeKey** and **EndKey** methods are different than the **MoveLeft** and **MoveRight** methods. Also, note that the **HomeKey** method is quite different from the **Start** property, which is the cp at the beginning of the selection. **HomeKey** and **EndKey** also differ from the **StartOf** and **EndOf** methods in that they extend from the active end, whereas **StartOf** extends from Start and **EndOf** extends from End.

---

## <a name="endkey"></a>EndKey

Mimics the functionality of the End key.

```
FUNCTION EndKey (BYVAL Unit AS LONG = tomLine, BYVAL Extend AS LONG = 0) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *Unit* | Unit to use in the End key operation. It can take on one of the following values (see table below). |
| *Extend* | Flag that indicates how to change the selection. If *Extend* is zero (or **tomMove**), the method collapses the selection to an insertion point. If *Extend* is 1 (or **tomExtend**), the method moves the active end and leaves the other end alone. The default value is zero. |

| Value | Meaning |
| ----- | ------- |
| **tomLine** | Depending on *Extend*, it moves either the insertion point or the active end to the end of the last line in the selection. This is the default. |
| **tomStory** | Depending on *Extend*, it moves either the insertion point or the active end to the end of the last line in the story. |
| **tomColumn** | Depending on *Extend*, it moves either the insertion point or the active end to the end of the last column in the selection. This is available only if the TOM engine supports tables. |
| **tomRow** | Depending on *Extend*, it moves either the insertion point or the active end to the end of the last row in the selection. This is available only if the TOM engine supports tables. |

#### Return value

The count of characters that the insertion point or the active end is moved. 

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Return code | Description |
| -------------- | ----------- |
| **E_INVALIDARG** | Unit is neither **tomLine** nor **tomStory**. |
| **S_FALSE** | Failure for some other reason. |

#### Remarks

Setting *Extend* to **tomExtend** (or nonzero) corresponds to the Shift key being pressed. Setting *Unit* to **tomLine** corresponds to the Ctrl key not being pressed. Setting *Unit* to **tomStory** to Ctrl being pressed.

The **HomeKey** and **EndKey** methods are used to mimic the standard Home/End key behavior.

The **tomLine** value mimics the Home or End key behavior without the Ctrl key pressed, while **tomStory** mimics the behavior with the Ctrl key pressed. Similarly, **tomMove** mimics the Home or End key behavior without the Shift key pressed, while **tomExtend** mimics the behavior with the Shift key pressed. So EndKey(tomStory) converts the selection into an insertion point at the end of the associated story, while EndKey(tomStory, tomExtend) moves the active end of the selection to the end of the story and leaves the other end where it was.

The **HomeKey** and **EndKey** methods are logical methods like the **Move** methods, rather than directional methods. Thus, they depend on the language that is involved. For example, in Arabic text, **HomeKey** moves to the right end of a line, whereas in English text, it moves to the left. Thus, **HomeKey** and **EndKey** are different than the **MoveLeft** and **MoveRight** methods. Also, note that the **EndKey** method is quite different from the **End** property, which is the cp at the end of the selection. **HomeKey** and **EndKey** also differ from the **StartOf** and **EndOf** methods in that they extend from the active end, whereas **StartOf** extends from Start and **EndOf** extends from End.

---

## <a name="typetext"></a>TypeText

Types the string given by *cbs* at this selection as if someone typed it. This is similar to the underlying **SetText** method, but is sensitive to the Insert/Overtype key state and UI settings like AutoCorrect and smart quotes.

```
FUNCTION TypeText (BYREF cbs AS CBSTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *cbs* | String to type into this selection. |

#### Return value

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following error codes.

| Return code | Description |
| -------------- | ----------- |
| **E_ACCESSDENIED** | Text is write-protected. |
| **E_OUTOFMEMORY** | Out of memory. |

#### Remarks

This method types the string given by *cbs* at this selection as if someone typed it. Using **TypeText** is faster than sending characters through the **SendMessage** function, but it is slower than using **SetText**.

**TypeText** is similar to the underlying **SetText** method, however, it is sensitive to the Insert/Overtype key state and UI settings like AutoCorrect and smart quotes. For example, it deletes any nondegenerate selection and then inserts or overtypes (depending on the Insert/Overtype key state—see the **SetFlags** method) the string *cbs* at the insertion point, leaving this selection as an insertion point following the inserted text.

---

## <a name="getCch"></a>GetCch

Gets the count of characters in a range.

```
FUNCTION GetCch () AS LONG
```

#### Return value

The signed count of characters.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The count of characters is the difference between the character position of the active end of the range, and the character position of the anchor end. Some Text Object Model (TOM) implementations might include active ends only for a selection (represented by the **ITextSelection** interface). The rich edit control's TOM implementation of a text range (represented by the **ITextRange** interface) also has active ends.

---

## <a name="getcells"></a>GetCells

Not implemented.

Gets a cells object with the parameters of cells in the currently selected table row or column.

```
FUNCTION GetCells () AS IUnknown PTR
```

#### Return value

The cells object.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="getcolumn"></a>GetColumn

Not implemented.

Gets the column properties for the currently selected column.

```
FUNCTION GetColumn () AS IUnknown PTR
```

#### Return value

The column properties for the currently selected column.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="getcount"></a>GetCount

Gets the count of subranges, including the active subrange in the current range.

```
FUNCTION GetCount () AS LONG
```

#### Return value

The count of subranges not including the active one.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

If you select a range with no or one character, the count will be 1. But if you select a word and then move to a different location, and select a second word not touching the first, then the count is 2.

See **AddSubrange** to add subranges programmatically.

---

## <a name="getgravity"></a>GetGravity

Gets the gravity of this range.

```
FUNCTION GetGravity () AS LONG
```
#### Return value

The gravity value, which can be one of the following:

| Constant | Value | Meaning |
| -------- | ----- | ------- |
| **tomGravityUI** | 0 | Use selection user interface rules. |
| **tomGravityBack** | 1 | Both ends have backward gravity. |
| **tomGravityFore** | 2 | Both ends have forward gravity. |
| **tomGravityIn** | 3 | Inward gravity; that is, the start is forward, and the end is backward. |
| **tomGravityOut** | 4 | Outward gravity; that is, the start is backward, and the end is forward. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="setgravity"></a>SetGravity

Sets the gravity of this range.

```
FUNCTION SetGravity (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new gravity value, which can be one of the following. |

| Constant | Value | Meaning |
| -------- | ----- | ------- |
| **tomGravityUI** | 0 | Use selection user interface rules. |
| **tomGravityBack** | 1 | Both ends have backward gravity. |
| **tomGravityFore** | 2 | Both ends have forward gravity. |
| **tomGravityIn** | 3 | Inward gravity; that is, the start is forward, and the end is backward. |
| **tomGravityOut** | 4 | Outward gravity; that is, the start is backward, and the end is forward. |

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="getrow"></a>GetRow

Gets the row properties in the currently selected row.

```
FUNCTION GetRow () AS ITextRow PTR
```

#### Return value

The row properties.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="getstartpara"></a>GetStartPara

Gets the character position of the start of the paragraph that contains the range's start character position.

```
FUNCTION GetStartPara () AS LONG
```

#### Return value

The start of the paragraph that contains the range's start character position.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="gettable"></a>GetTable

Gets the table properties in the currently selected table.

```
FUNCTION GetTable () AS IUnknown PTR
```

#### Return value

The table properties.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="geturl"></a>GetURL

Returns the URL text associated with a range.

```
FUNCTION GetURL () AS CBSTR
```

#### Return value

The URL text associated with the range.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="seturl"></a>SetURL

Sets the text in this range to that of the specified URL.

```
FUNCTION SetURL (BYREF cbs AS CBSTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *cbs* | The text to use as a URL for the selected friendly name. |

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Return code | Description |
| -------------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Out of memory. |

#### Remarks

The URL string is not validated. The text it contains must be enclosed in quotes, optionally preceded by the sentinel character 0xFDDF. For example: "http://www.msn.com" or 0xFDDF"http://www.msn.com". The range must be nondegenerate.

The following actions are possible:

- If part of a link's friendly name is selected, the URL part is replaced with bstr.
- If part of a regular URL is selected, it becomes the link's friendly name, with bstr as the URL.
- If nonlink text is selected:
  - If the text immediately follows a link's friendly name and bstr matches the URL, the text is appended to the friendly name.
  - Otherwise, the text becomes the friendly name of a link, with bstr as the URL.

The text range be adjusted to different character positions after calling SetURL.

---

## <a name="addsubrange"></a>AddSubrange

Adds a subrange to this range.

```
FUNCTION AddSubrange (BYVAL cp1 AS LONG, BYVAL cp2 AS LONG, BYVAL Activate AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *cp1* | The active-end character position of the subrange. |
| *cp2* | The anchor-end character position of the subrange. |
| *Activate* | The activate parameter. If this parameter is **tomTrue**, the new subrange is the active subrange, with *cp1* as the active end, and *cp2* the anchor end. |

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="buildupmath"></a>BuildUpMath

Converts the linear-format math in a range to a built-up form, or modifies the current built-up form.

```
FUNCTION BuildUpMath (BYVAL Flags AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Flags* | A combination of the following flags. |

| Constant | Value |
| -------- | ----- |
| **tomChemicalFormula** | &h1000 |
| **tomHaveDelimiter** | &h20 |
| **tomMathAlphabetics** | &h4 |
| **tomMathApplyTemplate** | &h4000 |
| **tomMathArabicAlphabetics** | &h800000 |
| **tomMathAutoCorrect** | &h100 |
| **tomMathAutoCorrectExt** | &h2000000 |
| **tomMathAutoCorrectOpPairs** | &h1000000 |
| **tomMathBackspace** | &h100000 |
| **tomMathBuildDown** | &h2000 |
| **tomMathBuildDownOutermost** | &h800 |
| **tomMathBuildUpArgOrZone** | &h200 |
| **tomMathBuildUpRecurse** | &h400 |
| **tomMathChangeMask** | &h1F0000 |
| **tomMathCollapseSel** | &h80 |
| **tomMathDeleteArg** | &h70000 |
| **tomMathDeleteArg1** | &h80000 |
| **tomMathDeleteArg2** | &h90000 |
| **tomMathDeleteCol** | &h60000 |
| **tomMathDeleteRow** | &h50000 |
| **tomMathEnter** | &h110000 |
| **tomMathInsColAfter** | &h40000 |
| **tomMathInsColBefore** | &h30000 |
| **tomMathInsRowAfter** | &h20000 |
| **tomMathInsRowBefore** | &h10000 |
| **tomMathMakeFracLinear** | &hA0000 |
| **tomMathMakeFracSlashed** | &hC0000 |
| **tomMathMakeFracStacked** | &hB0000 |
| **tomMathMakeLeftSubSup** | &hD0000 |
| **tomMathMakeSubSup** | &hE0000 |
| **tomMathRemoveOutermost** | &h8000 |
| **tomMathRichEdit** | &h40000000 |
| **tomMathShiftTab** | &h120000 |
| **tomMathSingleChar** | &h8 |
| **tomMathSubscript** | &h170000 |
| **tomMathSuperscript** | &h180000 |
| **tomMathTab** | &h130000 |
| **tomNeedTermOp** | &h2 |
| **tomPlain** | &h10 |
| **tomShowEmptyArgPlaceholders** | &h4000000 |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

If the **BuildUpMath** method is called on a nondegenerate range, the method checks the text for math italic conversions (if **tomMathAlphabetics** is specified) and math autocorrect conversions (if **tomMathAutoCorrect** or **tomMathAutoCorrectExt** is specified). Then, the method attempts to build up the selected text. If successful, the method replaces the previous text in the range with the built-up text. If the method makes any changes to the range, the function returns **NOERROR** and the range selects the result. If the method does change the range, it returns **S_FALSE** or a Component Object Model (COM) error code.

If the **BuildUpMath** method is called on a degenerate range, the **BuildUpMath** method treats the range as an insertion point (IP) immediately following the last character input. The method converts that character, possibly along with some preceding characters, to math italic (if **tomMathAlphabetics** is specified), internal math autocorrect (if **tomMathAutoCorrect** is specified), negated operators, and some operator pairs (if **tomMathAutoCorrectOpPairs** is specified). If the IP is inside an argument, the method scans a range of text from the IP back to the start of a math object argument; otherwise, the method scans to the start of the current math zone. The scan is terminated by a hard carriage return or a soft end-of-paragraph mark, because math zones are terminated by these marks. A scan forward from start of the math object argument or math zone bypasses text that has no chance of being built up. If the scan reaches the original entry IP, one of the following outcomes can occur:

- If the method made any changes, the function returns **NOERROR** and the range updated with the changed text.
- If the method made no changes, the function returns **S_FALSE** and leaves the range unchanged.

If the scan finds text that might get built up, the **BuildUpMath** method attempts to build up the text up to the insertion point. If successful, the method returns **NOERROR**, and the range is updated with the corresponding built-up text.

If this full build-up attempt fails, the **BuildUpMath** method does a partial build-up check for the expression immediately preceding the IP. If this succeeds, the method returns **NOERROR** and the range contains the linear text to be replaced by the built-up text.

If full and partial build-up attempts fail, the function returns as described previously for the cases where no build-up text was found. Other possible return values include **E_INVALIDARG** (if either interface pointer is **NULL**) and **E_OUTOFMEMORY**.

You should set the **tomNeedTermOp** flag should for formula autobuildup unless autocorrection has occurred that deletes the terminating blank. Autocorrection can occur when correcting text like \alpha when the user types a blank to force autocorrection.

---

## <a name="deletesubrange"></a>DeleteSubrange

Deletes a subrange from a range.

```
FUNCTION DeleteSubrange (BYVAL cpFirst AS LONG, BYVAL cpLim AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *cpFirst* | The start character position of the subrange. |
| *cpLim* | The end character position of the subrange. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

## <a name="find"></a>Find

Searches for math inline functions in text as specified by a source range.

```
FUNCTION Find (BYVAL pRange AS ITextRange2 PTR, BYVAL Count AS LONG, BYVAL Flags AS LONG) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *pRange* | The formatted text to find in the range's text. |
| *Count* | The number of characters to search through. |
| *Flags* | Flags that control the search as defined for **FindText**. |

#### Return value

A count of the number of characters bypassed.

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="getdropcap"></a>GetDropCap

Not implemented.

Gets the drop-cap parameters of the paragraph that contains this range.

```
FUNCTION GetDropCap (BYVAL pcLine AS LONG PTR, BYVAL pPosition AS LONG PTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *pcLine* | The count of lines for the drop cap. A value of 0 means no drop cap. |
| *pPosition* | The position of the drop cap. The position can be one of the following:<br>**tomDropMargin**<br>**tomDropNone**<br>**tomDropNormal** |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="getinlineobject"></a>GetInlineObject

Gets the properties of the inline object at the range active end.

See full MSDN documentation: [ITextRange2::GetInlineObject method](https://learn.microsoft.com/en-us/windows/win32/api/tom/nf-tom-itextrange2-getinlineobject)

```
FUNCTION GetInlineObject (BYVAL pType AS LONG PTR, BYVAL pAlign AS LONG PTR, _
   BYVAL pChar AS LONG PTR, BYVAL pChar1 AS LONG PTR, BYVAL pChar2 AS LONG PTR, _
   BYVAL pCount AS LONG PTR,  BYVAL pTeXStyle AS LONG PTR, BYVAL pcCol AS LONG PTR, _
   BYVAL pLevel AS LONG PTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *pType* | The inline object type. |
| *pAlign* | The inline object alignment. |
| *pChar* | The inline object character. |
| *pChar1* | The closing **tomBrackets** character. |
| *pChar2* | The separator character for **tomBracketsWithSep**. |
| *pCount* | The inline object count of arguments. |
| *pTeXStyle* | The inline object TeX style. |
| *pcCol* | The inline object count of columns (**tomMatrix** only). |
| *pLevel* | The inline object 0-based nesting level. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

[Unicode Technical Note 28](https://www.unicode.org/notes/tn28/) describes the alignment and character values in detail when the active end character is an inline object start delimiter.

When that character is not a start delimiter, the character and column parameters are set to 0, the count is set to the 0-based argument index, and the other parameters are set according to the active-end character properties of the innermost inline object argument.

## <a name="getproperty"></a>GetProperty

Gets the value of a property.

```
FUNCTION GetProperty (BYVAL nType AS LONG) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *pType* | The property ID. |

#### Return value

The property value.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="getrect"></a>GetRect

Retrieves a rectangle of the specified type for the current range.

```
FUNCTION GetRect (BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, BYVAL pTop AS LONG PTR, _
   BYVAL pRight AS LONG PTR, BYVAL pBottom AS LONG PTR, BYVAL pHit AS LONG PTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *pType* | The type of rectangle to return. This parameter can include one value from each of the tables below. |
| *pLeft* | The left rectangle coordinate. |
| *pTop* | The top rectangle coordinate. |
| *pRight* | The right rectangle coordinate. |
| *pBottom* | The bottom rectangle coordinate. |
| *pHit* | The hit-test value for the range. |

| Constant | Value | Meaning |
| -------- | ----- | ----------- |
| **tomAllowOffClient** | 512 | Allow points outside of the client area. |
| **tomClientCoord** | 256 | Use client coordinates instead of screen coordinates. |
| **tomObjectArg** | 2048 | Get a point inside an inline object argument; for example, inside the numerator of a fraction. |
| **tomTransform** | 1024 | Transform coordinates using a world transform (XFORM) supplied by the host application. |

Use one of the following values to indicate the vertical position.

| Constant | Value | Meaning |
| -------- | ----- | ----------- |
| **TA_TOP** | 0 | Top edge of the bounding rectangle. |
| **TA_BASELINE** | 24 | Base line of the text. |
| **TA_BOTTOM** | 8 | Bottom edge of the bounding rectangle. |

Use one of the following values to indicate the horizontal position.

| Constant | Value | Meaning |
| -------- | ----- | ----------- |
| **TA_LEFT** | 0 | Left edge of the bounding rectangle. |
| **TA_CENTER** | 6 | Center of the bounding rectangle. |
| **TA_RIGHT** | 2 | ight edge of the bounding rectangle. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="getsubrange"></a>GetSubrange

Retrieves a subrange in a range.

```
FUNCTION GetSubrange (BYVAL iSubrange AS LONG, BYVAL pcpFirst AS LONG PTR, _
   BYVAL pcpLim AS LONG PTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *iSubrange* | The subrange index. |
| *pcpFirst* | The character position for the start of the subrange. |
| *pcpLim* | The character position for the end of the subrange. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

Subranges are selected as follows.

| iSubrange value | Subrange |
| --------------- | -------- |
| Equals zero | Gets the current active subrange. |
| Greater than zero | Gets the subrange at the index specified by iSubrange, in the order in which the subranges were added. This requires extra calculation. |
| Less than zero | Gets the subrange at the index specified by *iSubrange*, in increasing character position order. |
 
See **GetCount** for the count of subranges not including the active subrange.

---

## <a name="hextounicode"></a>HexToUnicode

Converts and replaces the hexadecimal number at the end of this range to a Unicode character.

```
FUNCTION HexToUnicode () AS HRESULT
```

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

Some Unicode surrogates for hex values from &h10000 up to &h10FFFF are for internal use:

| Hex values | Available for use |
| ---------- | ----------------- |
| 7, &hFDD0 — &hFDEF, &hFFF9 — &hFFFF | Internal use only |
| &hA — &hD in the C0 range (0-&h1F) | Available for use |
| C1 range (&h80 — &h9F) | Internal use only |

---

## <a name="inserttable"></a>InsertTable

Inserts a table in a range.

```
FUNCTION InsertTable (BYVAL cCol AS LONG, BYVAL cRow AS LONG, BYVAL AutoFit AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *cCol* | The number of columns in the table. |
| *cRow* | The number of rows in the table. |
| *AutoFit* | Specifies how the cells fit the target space. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns an **HRESULT** error code.

---

## <a name="linearize"></a>Linearize

Translates the built-up math, ruby, and other inline objects in this range to linearized form.

```
FUNCTION Linearize (BYVAL Flags AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Flags* | A combination of the following flags: **tomMathAlphabetics**, **tomMathBuildDownOutermost**, **tomMathBuildUpArgOrZone**, **tomMathRemoveOutermost**, **tomPlain**, **tomTeX**. |

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="setactivesubrange"></a>SetActiveSubrange

Makes the specified subrange the active subrange of this range.

```
FUNCTION SetActiveSubrange (BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *cpAnchor* | The anchor end character position of the subrange to make active. |
| *cpActive* | The active end character position of the subrange to make active. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The active subrange is the one affected by operations such as Shift+Arrow keys if this range is the selection.

---

## <a name="setdropcap"></a>SetDropCap

Not implemented.

Sets the drop-cap parameters for the paragraph that contains the current range.

```
FUNCTION SetDropCap (BYVAL cLine AS LONG, BYVAL Position AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *cLine* | The count of lines for drop cap. Zero means no drop cap. |
| *Position* | The position of drop cap. It can be one of the following.<br>**tomDropMargin**,<br> **tomDropNone**, <br> **tomDropNormal** |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks
The current range can be degenerate, or you can select up to the complete drop-cap paragraph. If the range contains more than one paragraph, this method returns **E_FAIL**.

---

## <a name="setproperty"></a>SetProperty

Sets the value of the specified property.

```
FUNCTION SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *nType* | The ID of the property to set. |
| *Value* | The new property value. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="unicodetohex"></a>UnicodeToHex

Converts the Unicode character(s) preceding the start position of this text range to a hexadecimal number, and selects it.

```
FUNCTION UnicodeToHex () AS HRESULT
```

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

Some Unicode surrogates for hex values from &h10000 up to &h10FFFF are for internal use:

| Hex values | Available for use |
| ---------- | ----------------- |
| 7, &hFDD0 — &hFDEF, &hFFF9 — &hFFFF | Internal use only |
| &hA — &hD in the C0 range (0-&h1F) | Available for use |
| C1 range (&h80 — &h9F) | Internal use only |

---

## <a name="setinlineobject"></a>SetInlineObject

Sets or inserts the properties of an inline object for a degenerate range.

```
FUNCTION SetInlineObject (BYVAL nType AS Long, BYVAL Align AS LONG, BYVAL Char AS LONG, _
   BYVAL Char1 AS LONG, BYVAL Char2 AS LONG, BYVAL Count AS LONG, BYVAL TeXStyle AS LONG, _
   BYVAL cCol AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *nType* | The object type as defined in **GetInlineObject**. |
| *Align* | The object alignment as defined in **GetInlineObject**. |
| *Char* | The object character as defined in **GetInlineObject**. |
| *Char1* | The closing bracket (**tomBrackets**) character. See [Unicode Technical Note 28](https://www.unicode.org/notes/tn28/) for a list of characters. |
| *Char2* | The separator character for **tomBracketsWithSeps**, which can be one of the following values. |
| *Count* | The number of arguments in the inline object. |
| *TeXStyle* | The TeX style, as defined in **GetInlineObject**. |
| *cCol* | The number of columns in the inline object. For **tomMatrix** only. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="getmathfunctiontype"></a>GetMathFunctionType

Retrieves the math function type associated with the specified math function name.

```
FUNCTION GetMathFunctionType (BYREF cbs AS CBSTR) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *cbs* | The math function name that is checked to determine the math function type. |

#### Return value

The math function type of the function name specified by bstr. It can be one of the following values.

| Constant | Value | Description |
| -------- | ----- | ----------- |
| **tomFunctionTypeNone** | 0 | Not in the function list. |
| **tomFunctionTypeTakesArg** | 1 | An ordinary math function that takes arguments. |
| **tomFunctionTypeTakesLim** | 2 | Use the lower limit for _, and so on. |
| **tomFunctionTypeTakesLim2** | 3 | Turn the preceding FA into an NBSP. |
| **tomFunctionTypeIsLim** | 4 | A "lim" function. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="insertimage"></a>InsertImage

Inserts an image into this range.

```
FUNCTION InsertImage (BYVAL width_ AS LONG, BYVAL height AS LONG, BYVAL ascent AS LONG, _
   BYVAL nType AS LONG, BYVAL bstrAltText AS AFX_BSTR, BYVAL pStream AS IStream PTR) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *width_* | The width, in HIMETRIC units (0.01 mm), of the image. |
| *height* | The height, in HIMETRIC units, of the image. |
| *ascent* | If *nType* is TA_BASELINE, this parameter is the distance, in HIMETRIC units, that the top of the image extends above the text baseline. If *nType* is TA_BASELINE and ascent is zero, the bottom of the image is placed at the text baseline. |
| *nType* | The vertical alignment of the image. It can be one of the following values:<br>**TA_BASELINE**. Align the image relative to the text baseline.<br>**TA_BOTTOM**. Align the bottom of the image at the bottom of the text line.<br>**TA_TOP**. Align the top of the image at the top of the text line. |
| *bstrAltText* | The alternate text for the image. |
| *pStream* | The stream that contains the image data. |

#### Return value

If this method succeeds, it returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

If the range is nondegenerate, the image replaces the text in the range.

---

## <a name="getrtf"></a>GetRtf

Gets the content of the range as a Rich Text Format (RTF) string.

```
FUNCTION CTextRange2.GetRtf () AS STRING
```

#### Return value

If the method succeeds, it returns the content of the range as a Rich Text Format (RTF) string. If the method fails, it returns an empty string. Call **GetErrorInfo** to get information about the error.

---
