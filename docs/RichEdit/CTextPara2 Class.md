# CTextPara2 Class

Class that wraps all the methods of the **ITextPara** and **ITextPara2** interfaces.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTORS](#CONSTRUCTORS) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [LET](#LET) | Assignment operator. |
| [CAST](#CAST) | Cast operator. |
| [TextDocumentPtr](#TextDocumentPtr) | Returns a pointer to the underlying **ITextPara2** interface. |
| [Attach](#Attach) | Attaches an **ITextPara2** interface pointer to the class. |
| [Detach](#Detach) | Detaches the underlying **ITextPara2** interface pointer from the class. |

### ITextPara Interface

Text Object Model (TOM) rich text-range attributes are accessed through a pair of dual interfaces, **ITextFont** and **ITextPara**.

The **ITextPara** interface inherits from the **IDispatch** interface. **ITextPara** also has these types of members:

| Name       | Description |
| ---------- | ----------- |
| [GetDuplicate](#GetDuplicate) | Gets a duplicate of this text paragraph format object. |
| [SetDuplicate](#SetDuplicate) | Sets the formatting for an existing paragraph by copying a given format. |
| [CanChange](#CanChange) | Determines whether the paragraph formatting can be changed. |
| [IsEqual](#IsEqual) | Determines if the current range has the same properties as a specified range. |
| [Reset](#Reset) | Resets the paragraph formatting to a choice of default values. |
| [GetStyle](#GetStyle) | Retrieves the style handle to the paragraphs in the specified range. |
| [SetStyle](#SetStyle) | Sets the paragraph style for the paragraphs in a range. |
| [GetAlignment](#GetAlignment) | Retrieves the current paragraph alignment value. |
| [SetAlignment](#SetAlignment) | Sets the paragraph alignment. |
| [GetHyphenation](#GetHyphenation) | Determines whether automatic hyphenation is enabled for the range. |
| [SetHyphenation](#SetHyphenation) | Controls hyphenation for the paragraphs in the range. |
| [GetFirstLineIndent](#GetFirstLineIndent) | Retrieves the amount used to indent the first line of a paragraph relative to the left indent. The left indent is the indent for all lines of the paragraph except the first line. |
| [GetKeepTogether](#GetKeepTogether) | Determines whether page breaks are allowed within paragraphs. |
| [SetKeepTogether](#SetKeepTogether) | Controls whether page breaks are allowed within a paragraph in a range. |
| [GetKeepWithNext](#GetKeepWithNext) | Determines whether page breaks are allowed between paragraphs in the range. |
| [SetKeepWithNext](#SetKeepWithNext) | Controls whether page breaks are allowed between the paragraphs in a range. |
| [GetLeftIndent](#GetLeftIndent) | Retrieves the distance used to indent all lines except the first line of a paragraph. The distance is relative to the left margin. |
| [GetLineSpacing](#GetLineSpacing) | Retrieves the line-spacing value for the text range. |
| [GetLineSpacingRule](#GetLineSpacingRule) | Retrieves the line-spacing rule for the text range. |
| [GetListAlignment](#GetListAlignment) | Retrieves the kind of alignment to use for bulleted and numbered lists. |
| [SetListAlignment](#SetListAlignment) | Sets the alignment of bulleted or numbered text used for paragraphs. |
| [GetListLevelIndex](#GetListLevelIndex) | Retrieves the list level index used with paragraphs. |
| [SetListLevelIndex](#SetListLevelIndex) | Sets the list level index used for paragraphs. |
| [GetListStart](#GetListStart) | Retrieves the starting value or code of a list numbering sequence. |
| [SetListStart](#SetListStart) | Sets the starting number or Unicode value for a numbered list. |
| [GetListTab](#GetListTab) | Retrieves the list tab setting, which is the distance between the first-line indent and the text on the first line. The numbered or bulleted text is left-justified, centered, or right-justified at the first-line indent value. |
| [SetListTab](#SetListTab) | Sets the list tab setting, which is the distance between the first indent and the start of the text on the first line. |
| [GetListType](#GetListType) | Retrieves the kind of numbering to use with paragraphs. |
| [SetListType](#SetListType) | Sets the type of list to be used for paragraphs. |
| [GetNoLineNumber](#GetNoLineNumber) | Determines whether paragraph numbering is enabled. |
| [SetNoLineNumber](#SetNoLineNumber) | Determines whether to suppress line numbering of paragraphs in a range. |
| [GetPageBreakBefore](#GetPageBreakBefore) | Determines whether each paragraph in the range must begin on a new page. |
| [SetPageBreakBefore](#SetPageBreakBefore) | Controls whether there is a page break before each paragraph in a range. |
| [GetRightIndent](#GetRightIndent) | Retrieves the size of the right margin indent of a paragraph. |
| [SetRightIndent](#SetRightIndent) | Sets the right margin of paragraph. |
| [SetIndents](#SetIndents) | Sets the first-line indent, the left indent, and the right indent for a paragraph. |
| [SetLineSpacing](#SetLineSpacing) | Sets the paragraph line-spacing rule and the line spacing for a paragraph. |
| [GetSpaceAfter](#GetSpaceAfter) | The space-after value, in floating-point points. |
| [SetSpaceAfter](#SetSpaceAfter) | Sets the amount of space that follows a paragraph. |
| [GetSpaceBefore](#GetSpaceBefore) |  |
| [SetSpaceBefore](#SetSpaceBefore) |  |
| [GetWidowControl](#GetWidowControl) |  |
| [SetWidowControl](#SetWidowControl) |  |
| [GetTabCount](#GetTabCount) |  |
| [AddTab](#AddTab) |  |
| [ClearAllTabs](#ClearAllTabs) |  |
| [DeleteTab](#DeleteTab) |  |
| [GetTab](#GetTab) |  |

### ITextPara2 Interface

The **ITextPara2** interface extends ITextPara, providing the equivalent of the Microsoft Word format-paragraph dialog.

The **ITextPara2** interface has these methods.

| Name       | Description |
| ---------- | ----------- |
| [GetBorders](#GetBorders) |  |
| [GetDuplicate2](#GetDuplicate) | Gets a duplicate of this text paragraph format object. |
| [SetDuplicate2](#SetDuplicate) | Sets the formatting for an existing paragraph by copying a given format. |
| [GetFontAlignment](#GetFontAlignment) |  |
| [SetFontAlignment](#SetFontAlignment) |  |
| [GetHangingPunctuation](#GetHangingPunctuation) |  |
| [SetHangingPunctuation](#SetHangingPunctuation) |  |
| [GetSnapToGrid](#GetSnapToGrid) |  |
| [SetSnapToGrid](#SetSnapToGrid) |  |
| [GetTrimPunctuationAtStart](#GetTrimPunctuationAtStart) |  |
| [SetTrimPunctuationAtStart](#SetTrimPunctuationAtStart) |  |
| [GetEffects](#GetEffects) |  |
| [GetProperty](#GetProperty) |  |
| [IsEqual2](#IsEqual) | Determines if the current range has the same properties as a specified range. |
| [SetEffects](#SetEffects) |  |
| [SetProperty](#SetProperty) |  |

### Methods inherited from CTOMBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#GetLastResult) | Returns the last result code |
| [SetResult](#SetResult) | Sets the last result code. |
| [GetErrorInfo](#GetErrorInfo) | Returns a description of the last result code. |

# <a name="CONSTRUCTORS"></a>CONSTRUCTORS

Called when a **CTextPara2** class variable is created.

```
DECLARE CONSTRUCTOR
DECLARE CONSTRUCTOR (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

## CONSTRUCTOR (Empty)

Can be used, for example, when we have an **ITextPara2** interface pointer returned by a function and we want to attach it to a new instance of the **CTextPara2** class.

```
DIM DIM pCTextPara2 AS CTextPara2
pCTextPara2.Attach(pTextPara2)
```
## CONSTRUCTOR (ITextPara2 PTR)

```
CONSTRUCTOR CTextPara2 (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   IF pTextPara2 THEN
      IF fAddRef THEN pTextPara2->lpvtbl->AddRef(pTextPara2)
   End IF
   m_pTextPara2 = pTextPara2
END CONSTRUCTOR
```

| Parameter | Description |
| --------- | ----------- |
| *pTextPara2* | An **ITextPara2** interface pointer. |
| *fAddRef* | Optional. **TRUE** to increment the reference count of the passed **ITextPara2** interface pointer; otherwise, **FALSE**. Default is **FALSE**. |

#### Return value

A pointer to the new instance of the class.

# <a name="DESTRUCTOR"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

```
DESTRUCTOR CTextPara2
   ' // Release the ITextPara2 interface
   IF m_pTextPara2 THEN m_pTextPara2->lpvtbl->Release(m_pTextPara2)
END DESTRUCTOR
```

# <a name="LET"></a>LET

Assignment operator. The assigned pointer must be an "addrefed" one.

```
OPERATOR CTextPara2.LET (BYVAL pTextPara2 AS ITextPara2 PTR)
   m_Result = 0
   IF pTextPara2 = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextPara2 THEN m_pTextPara2->lpvtbl->Release(m_pTextPara2)
   ' // Attach the passed interface pointer to the class
   m_pTextPara2 = pTextPara2
END OPERATOR
```

# <a name="CAST"></a>CAST

Cast operator.

```
OPERATOR CTextPara2.CAST () AS ITextPara2 PTR
   m_Result = 0
   OPERATOR = m_pTextPara2
END OPERATOR
```

# <a name="TextParaPtr"></a>TextParaPtr

Returns a pointer to the underlying **ITextPara2** interface

```
FUNCTION CTextPara2.TextParaPtr () AS ITextPara2 PTR
   m_Result = 0
   RETURN m_pTextPara2
END FUNCTION
```

# <a name="Attach"></a>Attach

Attaches an **ITextPara2** interface pointer to the class.

```
FUNCTION CTextPara2.Attach (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextPara2 = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextPara2 THEN m_Result = m_pTextPara2->lpvtbl->Release(m_pTextPara2)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextPara2->lpvtbl->AddRef(pTextPara2)
   m_pTextPara2 = pTextPara2
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pTextPara2* | The **ITextPara2** interface pointer to attach. |
| *fAddRef* | **TRUE** to increment the reference count of te object; otherwise, **FALSE**. Default is **FALSE**. |

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

# <a name="GetDuplicate"></a>GetDuplicate

Gets a duplicate of this range object. In this implementation of the class, **GetDuplicate** and **GetDuplicate2** are the same method.

```
FUNCTION CTextPara2.GetDuplicate () AS ITextPara2 PTR
   DIM pPara AS ITextPara2 PTR
   this.SetResult(m_pTextPara2->lpvtbl->GetDuplicate2(m_pTextPara2, @pPara))
   FUNCTION = pPara
END FUNCTION
```
```
FUNCTION CTextPara2.GetDuplicate2 () AS ITextPara2 PTR
   DIM pPara AS ITextPara2 PTR
   this.SetResult(m_pTextPara2->lpvtbl->GetDuplicate2(m_pTextPara2, @pPara))
   FUNCTION = pPara
END FUNCTION
```

#### Return value

The duplicate of the range.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_OUTOFMEMORY** | Memory could not be allocated for the new object. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="SetDuplicate"></a>SetDuplicate

Sets the properties of this object by copying the properties of another text paragraph object. In this implementation of the class, **SetDuplicate** and **SetDuplicate2** are the same method.

```
FUNCTION CTextPara2.SetDuplicate (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetDuplicate(m_pTextPara2, pPara))
   FUNCTION = m_Result
END FUNCTION
```
```
FUNCTION CTextPara2.SetDuplicate2 (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetDuplicate2(m_pTextPara2, pPara))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pPara* | The text paragraph object to copy from. |

#### Return value

If **SetDuplicate2** succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="CanChange"></a>CanChange

Determines whether the paragraph formatting can be changed.

```
FUNCTION CTextPara2.CanChange (BYVAL pPara AS ITextPara2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->CanChange(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

**tomTrue** if the paragraph formatting can be changed or **tomFalse** if it cannot be changed.

#### Result code

If paragraph formatting can change, **CanChange** succeeds and **GetLastResult** returns **S_OK**. If paragraph formatting cannot change, the method fails and returns **S_FALSE**.

# <a name="IsEqual"></a>IsEqual

Determines if the current range has the same properties as a specified range.  In this implementation of the class, **IsEqual** and **IsEqual2** are the same method.

```
FUNCTION CTextPara2.IsEqual (BYVAL pPara AS ITextPara2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->CanChange(m_pTextPara2, pPara, @Value))
   FUNCTION = Value
END FUNCTION
```
```
FUNCTION CTextPara2.IsEqual2 (BYVAL pPara AS ITextPara2 PTR) AS LONG
   DIM B AS LONG
   IF m_pTextPara2 = NULL THEN m_Result = E_POINTER: RETURN B
   this.SetResult(m_pTextPara2->lpvtbl->IsEqual2(m_pTextPara2, pPara, @B))
   FUNCTION = B
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pPara* | The **ITextPara2** range that is compared to the current range. |

#### Return value

A **tomBool** value that is **tomTrue** if the text paragraph objects have the same properties, or **tomFalse** if they don't. 

#### Result code

If paragraph formatting can change, **IsEqual** succeeds and **GetLastResult** returns **S_OK**. If paragraph formatting cannot change, the method fails and returns **S_FALSE**.

# <a name="Reset"></a>Reset

Determines whether the paragraph formatting can be changed.

```
FUNCTION CTextPara2.Reset (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->Reset(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | Type of reset. It can be one of the following possible values. |

| Value | Meaning |
| ----- | ------- |
| **tomDefault** | Used for paragraph formatting that is defined by the RTF \pard, that is, the paragraph default control word. |
| **tomUndefined** | Used for all undefined values. The tomUndefined value is only valid for duplicate (clone) ITextPara objects. |

#### Return value

If **Reset** succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| E_ACCESSDENIED | Write access is denied. |
| E_OUTOFMEMORY | Insufficient memory. |
| CO_E_RELEASED | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetStyle"></a>GetStyle

Retrieves the style handle to the paragraphs in the specified range.

```
FUNCTION CTextPara2.GetStyle () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetStyle(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The paragraph style handle.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

Remarks
The Text Object Model (TOM) version 1.0 has no way to specify the meanings of user-defined style handles. They depend on other facilities of the text system implementing TOM. Negative style handles are reserved for built-in character and paragraph styles. Currently defined values are listed in the following table. For a description of the following styles, see the Microsoft Word documentation.

| Style | Value | Style | Value |
| ----- | ----- | ----- | ----- |
| StyleNormal | -1 | StyleTableofAuthorities | -45 |
| StyleHeading1 | -2 | StyleMacroText | -46 |
| StyleHeading2 | -3 | StyleTOAHeading | -47 |
| StyleHeading3 | -4 | StyleList | -48 |
| StyleHeading4 | -5 | StyleListBullet | -49 |
| StyleHeading5 | -6 | StyleListNumber | -50 |
| StyleHeading6 | -7 | StyleList2 | -51 |
| StyleHeading7 | -8 | StyleList3 | -52 |
| StyleHeading8 | -9 | StyleList4 | -53 |
| StyleHeading9 | -10 | StyleList5 | -54 |
| StyleIndex1 | -11 | StyleListBullet2 | -55 |
| StyleIndex2 | -12 | StyleListBullet3 | -56 |
| StyleIndex3 | -13 | StyleListBullet4 | -57 |
| StyleIndex4 | -14 | StyleListBullet5 | -58 |
| StyleIndex5 | -15 | StyleListNumber2 | -59 |
| StyleIndex6 | -16 | StyleListNumber3 | -60 |
| StyleIndex7 | -17 | StyleListNumber4 | -61 |
| StyleIndex8 | -18 | StyleListNumber5 | -62 |
| StyleIndex9 | -19 | StyleTitle | -63 |
| StyleTOC1 | -20 | StyleClosing | -64 |
| StyleTOC2 | -21 | StyleClosing | -65 |
| StyleTOC3 | -22 | StyleDefaultParagraphFont | -66 |
| StyleTOC4 | -23 | StyleBodyText | -67 |
| StyleTOC5 | -24 | StyleBodyTextIndent | -68 |
| StyleTOC6 | -25 | StyleListContinue | -69 |
| StyleTOC7 | -26 | StyleListContinue2 | -70 |
| StyleTOC8 | -27 | StyleListContinue3 | -71 |
| StyleTOC9 | -28 | StyleListContinue4 | -72 |
| StyleNormalIndent | -29 | StyleListContinue5 | -73 |
| StyleFootnoteText | -30 | StyleMessageHeader | -74 |
| StyleAnnotationText | -31 | StyleSubtitle | -75 |
| StyleHeader | -32 | StyleSalutation | -76 |
| StyleFooter | -33 | StyleDate | -77 |
| StyleIndexHeading | -34 | StyleBodyTextFirstIndent | -78 |
| StyleCaption | -35 | StyleBodyTextFirstIndent2 | -79 |
| StyleTableofFigures | -36 | StyleNoteHeading | -80 |
| StyleEnvelopeAddress | -37 | StyleBodyText2 | -81 |
| StyleEnvelopeReturn | -38 | StyleBodyText3 | -82 |
| StyleFootnoteReference | -39 | StyleBodyTextIndent2 | -83 |
| StyleAnnotationReference | -40 | StyleBodyTextIndent3 | -84 |
| StyleLineNumber | -41 | StyleBlockQuotation | -85 |
| StylePageNumber | -42 | StyleHyperlink | -86 |
| StyleEndnoteReference | -43 | StyleHyperlinkFollowed | -87 |
| StyleEndnoteText | -44 |  |  |

# <a name="SetStyle"></a>SetStyle

Sets the paragraph style for the paragraphs in a range.

```
FUNCTION CTextPara2.SetStyle (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetStyle(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | New paragraph style handle. |

| Style | Value | Style | Value |
| ----- | ----- | ----- | ----- |
| StyleNormal | -1 | StyleTableofAuthorities | -45 |
| StyleHeading1 | -2 | StyleMacroText | -46 |
| StyleHeading2 | -3 | StyleTOAHeading | -47 |
| StyleHeading3 | -4 | StyleList | -48 |
| StyleHeading4 | -5 | StyleListBullet | -49 |
| StyleHeading5 | -6 | StyleListNumber | -50 |
| StyleHeading6 | -7 | StyleList2 | -51 |
| StyleHeading7 | -8 | StyleList3 | -52 |
| StyleHeading8 | -9 | StyleList4 | -53 |
| StyleHeading9 | -10 | StyleList5 | -54 |
| StyleIndex1 | -11 | StyleListBullet2 | -55 |
| StyleIndex2 | -12 | StyleListBullet3 | -56 |
| StyleIndex3 | -13 | StyleListBullet4 | -57 |
| StyleIndex4 | -14 | StyleListBullet5 | -58 |
| StyleIndex5 | -15 | StyleListNumber2 | -59 |
| StyleIndex6 | -16 | StyleListNumber3 | -60 |
| StyleIndex7 | -17 | StyleListNumber4 | -61 |
| StyleIndex8 | -18 | StyleListNumber5 | -62 |
| StyleIndex9 | -19 | StyleTitle | -63 |
| StyleTOC1 | -20 | StyleClosing | -64 |
| StyleTOC2 | -21 | StyleClosing | -65 |
| StyleTOC3 | -22 | StyleDefaultParagraphFont | -66 |
| StyleTOC4 | -23 | StyleBodyText | -67 |
| StyleTOC5 | -24 | StyleBodyTextIndent | -68 |
| StyleTOC6 | -25 | StyleListContinue | -69 |
| StyleTOC7 | -26 | StyleListContinue2 | -70 |
| StyleTOC8 | -27 | StyleListContinue3 | -71 |
| StyleTOC9 | -28 | StyleListContinue4 | -72 |
| StyleNormalIndent | -29 | StyleListContinue5 | -73 |
| StyleFootnoteText | -30 | StyleMessageHeader | -74 |
| StyleAnnotationText | -31 | StyleSubtitle | -75 |
| StyleHeader | -32 | StyleSalutation | -76 |
| StyleFooter | -33 | StyleDate | -77 |
| StyleIndexHeading | -34 | StyleBodyTextFirstIndent | -78 |
| StyleCaption | -35 | StyleBodyTextFirstIndent2 | -79 |
| StyleTableofFigures | -36 | StyleNoteHeading | -80 |
| StyleEnvelopeAddress | -37 | StyleBodyText2 | -81 |
| StyleEnvelopeReturn | -38 | StyleBodyText3 | -82 |
| StyleFootnoteReference | -39 | StyleBodyTextIndent2 | -83 |
| StyleAnnotationReference | -40 | StyleBodyTextIndent3 | -84 |
| StyleLineNumber | -41 | StyleBlockQuotation | -85 |
| StylePageNumber | -42 | StyleHyperlink | -86 |
| StyleEndnoteReference | -43 | StyleHyperlinkFollowed | -87 |
| StyleEndnoteText | -44 |  |  |

#### Return value

If **SetStyle** succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetAlignment"></a>GetAlignment

Retrieves the current paragraph alignment value.

```
FUNCTION CTextPara2.GetAlignment () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetAlignment(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The paragraph alignment, which can be one of the following values.

| Value | Description |
| --------- | ----------- |
| **tomAlignLeft** | Text aligns with the left margin. |
| **tomAlignCenter** | Text is centered between the margins. |
| **tomAlignRight** | Text aligns with the right margin. |
| **tomAlignJustify** | Text starts at the left margin and, if the line extends beyond the right margin, all the spaces in the line are adjusted to be even. |
| **tomAlignInterWord** | Same as **tomAlignJustify**. |
| **tomAlignNewspaper** | Same as **tomAlignInterLetter**, but uses East Asian metrics. |
| **tomAlignInterLetter** | The first and last characters of each line (except the last line) are aligned to the left and right margins, and lines are filled by adding or subtracting the same amount from each character. |
| **tomAlignScaled** | Same as **tomAlignInterLetter**, but uses East Asian metrics, and scales the spacing by the width of characters. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="SetAlignment"></a>SetAlignment

Sets the paragraph alignment.

```
FUNCTION CTextPara2.SetAlignment (BYVAL Value AS LONG) AS HRESULT
   IF m_pTextPara2 = NULL THEN m_Result = E_POINTER: RETURN m_Result
   this.SetResult(m_pTextPara2->lpvtbl->SetStyle(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | New paragraph alignment. |

| Value | Description |
| --------- | ----------- |
| **tomAlignLeft** | Text aligns with the left margin. |
| **tomAlignCenter** | Text is centered between the margins. |
| **tomAlignRight** | Text aligns with the right margin. |
| **tomAlignJustify** | Text starts at the left margin and, if the line extends beyond the right margin, all the spaces in the line are adjusted to be even. |
| **tomAlignInterWord** | Same as **tomAlignJustify**. |
| **tomAlignNewspaper** | Same as **tomAlignInterLetter**, but uses East Asian metrics. |
| **tomAlignInterLetter** | The first and last characters of each line (except the last line) are aligned to the left and right margins, and lines are filled by adding or subtracting the same amount from each character. |
| **tomAlignScaled** | Same as **tomAlignInterLetter**, but uses East Asian metrics, and scales the spacing by the width of characters. |

#### Return value

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetHyphenation"></a>GetHyphenation

Determines whether automatic hyphenation is enabled for the range.

```
FUNCTION CTextPara2.GetHyphenation () AS LONG
   DIM Value AS LONG
   IF m_pTextPara2 = NULL THEN m_Result = E_POINTER: RETURN Value
   this.SetResult(m_pTextPara2->lpvtbl->GetHyphenation(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

One of the following values:

| Value | Description |
| --------- | ----------- |
| **tomTrue** | Automatic hyphenation is enabled. |
| **tomFalse** | Automatic hyphenation is disabled. |
| **tomUndefined** | The hyphenation property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the **PFE_DONOTHYPHEN** effect described in the [PARAFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-paraformat2) structure.

# <a name="SetHyphenation"></a>SetHyphenation

Controls hyphenation for the paragraphs in the range.

```
FUNCTION CTextPara2.SetHyphenation (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetHyphenation(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | Indicates how hyphenation is controlled. It can be one of the following possible values. |

| Value | Description |
| --------- | ----------- |
| **tomTrue** | Automatic hyphenation is enabled. |
| **tomFalse** | Automatic hyphenation is disabled. |
| **tomUndefined** | The hyphenation property is undefined. |

#### Return value

If **SetHyphenation** succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes. 

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetFirstLineIndent"></a>GetFirstLineIndent

Retrieves the amount used to indent the first line of a paragraph relative to the left indent. The left indent is the indent for all lines of the paragraph except the first line.

```
FUNCTION CTextPara2.GetFirstLineIndent () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetFirstLineIndent(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The first-line indentation amount in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

To set the first line indentation amount, call the **SetIndents** method.

To get and set the indent for all other lines of the paragraph (that is, the left indent), use **GetLeftIndent** and **SetIndents**.

# <a name="GetKeepTogether"></a>GetKeepTogether

Determines whether page breaks are allowed within paragraphs.

```
FUNCTION CTextPara2.GetKeepTogether () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetKeepTogether(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

One of the following values.

| Value | Description |
| --------- | ----------- |
| **tomTrue** | Page breaks are not allowed within a paragraph. |
| **tomFalse** | Page breaks are allowed within a paragraph. |
| **tomUndefined** | The property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the **PFE_KEEP** effect described in the [PARAFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-paraformat2) structure.

# <a name="SetKeepTogether"></a>SetKeepTogether

Controls whether page breaks are allowed within a paragraph in a range.

```
FUNCTION CTextPara2.SetKeepTogether (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetKeepTogether(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | Indicates whether page breaks are allowed within a paragraph in a range. It can be one of the following possible values. |

| Value | Description |
| --------- | ----------- |
| **tomTrue** | Page breaks are not allowed within a paragraph. |
| **tomFalse** | Page breaks are allowed within a paragraph. |
| **tomUndefined** | The property is undefined. |

#### Return value

If **SetKeepTogether** succeeds, it returns S_OK. If the method fails, it returns one of the following COM error codes. 

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetKeepWithNext"></a>GetKeepWithNext

Determines whether page breaks are allowed between paragraphs in the range.

```
FUNCTION CTextPara2.GetKeepWithNext () AS LONG
   DIM Value AS LONG
   IF m_pTextPara2 = NULL THEN m_Result = E_POINTER: RETURN Value
   this.SetResult(m_pTextPara2->lpvtbl->GetKeepWithNext(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

One of the following values.

| Value | Description |
| --------- | ----------- |
| **tomTrue** | Page breaks are not allowed between paragraphs. |
| **tomFalse** | Page breaks are allowed between paragraphs. |
| **tomUndefined** | The property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the PFE_KEEPNEXT effect described in the [PARAFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-paraformat2) structure.

# <a name="SetKeepWithNext"></a>SetKeepWithNext

Controls whether page breaks are allowed between the paragraphs in a range.

```
FUNCTION CTextPara2.SetKeepWithNext (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetKeepWithNext(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | Indicates if page breaks can be used between the paragraphs of a range. It can be one of the following possible values. |

| **tomTrue** | Page breaks are not allowed between paragraphs. |
| **tomFalse** | Page breaks are allowed between paragraphs. |
| **tomUndefined** | The property is undefined. |

#### Return value

If **SetKeepWithNext** succeeds, it returns S_OK. If the method fails, it returns one of the following COM error codes. 

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks
This property corresponds to the PFE_KEEPNEXT effect described in the [PARAFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-paraformat2) structure.

# <a name="GetLeftIndent"></a>GetLeftIndent

Retrieves the distance used to indent all lines except the first line of a paragraph. The distance is relative to the left margin.

```
FUNCTION CTextPara2.GetLeftIndent () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetLeftIndent(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The left indentation, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetLineSpacing"></a>GetLineSpacing

Retrieves the line-spacing value for the text range.

```
FUNCTION CTextPara2.GetLineSpacing () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetLineSpacing(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The line-spacing value. The following table shows how this value is interpreted for the different line-spacing rules.

| Line spacing rule | Meaning |
| ----------------- | ------- |
| **tomLineSpaceSingle** | Single spacing. The line-spacing value is ignored. |
| **tomLineSpace1pt5** | One-and-a-half line spacing. The line-spacing value is ignored. |
| **tomLineSpaceDouble** | Double line spacing. The line-spacing value is ignored. |
| **tomLineSpaceAtLeast** | The line-spacing value specifies the spacing, in floating-point points, from one line to the next. However, if the value is less than single spacing, the control displays single-spaced text. |
| **tomLineSpaceExactly** | The line-spacing value specifies the exact spacing, in floating-point points, from one line to the next (even if the value is less than single spacing). |
| **tomLineSpaceMultiple** | The line-spacing value specifies the line spacing, in lines. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetLineSpacingRule"></a>GetLineSpacingRule

Retrieves the line-spacing rule for the text range.

```
FUNCTION CTextPara2.GetLineSpacingRule () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetLineSpacingRule(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

One of the following values that indicates the line-spacing rule.

| Line spacing rule | Meaning |
| ----------------- | ------- |
| **tomLineSpaceSingle** | Single space. The line-spacing value is ignored. |
| **tomLineSpace1pt5** | One-and-a-half line spacing. The line-spacing value is ignored. |
| **tomLineSpaceDouble** | Double line spacing. The line-spacing value is ignored. |
| **tomLineSpaceAtLeast** | The line-spacing value specifies the spacing, in floating-point points, from one line to the next. However, if the value is less than single spacing, the control displays single-spaced text. |
| **tomLineSpaceExactly** | The line-spacing value specifies the exact spacing, in floating-point points, from one line to the next (even if the value is less than single spacing). |
| **tomLineSpaceMultiple** | The line-spacing value specifies the line spacing, in lines.|
| **tomLineSpacePercent** | The line-spacing value specifies the line spacing by percent of line height. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetListAlignment"></a>GetListAlignment

Retrieves the kind of alignment to use for bulleted and numbered lists.

```
FUNCTION CTextPara2.GetListAlignment () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetListAlignment(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

One of the following values that indicates the kind of bullet and numbering alignment.

| Line spacing rule | Meaning |
| ----------------- | ------- |
| **tomAlignLeft** | Text is left aligned. |
| **tomAlignCenter** | Text is centered in the line. |
| **tomAlignRight** | Text is right aligned. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="SetListAlignment"></a>SetListAlignment

Sets the alignment of bulleted or numbered text used for paragraphs.

```
FUNCTION CTextPara2.SetListAlignment (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetListAlignment(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```
| Parameter | Description |
| ----------------- | ------- |
| **Value** | New list alignment value. |

| Line spacing rule | Meaning |
| ----------------- | ------- |
| **tomAlignLeft** | Text is left aligned. |
| **tomAlignCenter** | Text is centered in the line. |
| **tomAlignRight** | Text is right aligned. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetListLevelIndex"></a>GetListLevelIndex

Retrieves the list level index used with paragraphs.

```
FUNCTION CTextPara2.GetListLevelIndex () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetListLevelIndex(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The list level index.

| Value | Meaning |
| ----- | ------- |
| 0 | No list. |
| 1 | First-level (outermost) list. |
| 2 | Second-level (nested) list. This is nested under a level 1 list item. |
| 3 | Third-level (nested) list. This is nested under a level 2 list item. |
| and so forth | Nesting continues similarly. |

Up to three levels are common in HTML documents.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="SetListLevelIndex"></a>SetListLevelIndex

Sets the list level index used for paragraphs.

```
FUNCTION CTextPara2.SetListLevelIndex (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetListLevelIndex(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | The new list level index value. |

| Value | Meaning |
| ----- | ------- |
| 0 | No list. |
| 1 | First-level (outermost) list. |
| 2 | Second-level (nested) list. This is nested under a level 1 list item. |
| 3 | Third-level (nested) list. This is nested under a level 2 list item. |
| and so forth | Nesting continues similarly. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetListStart"></a>GetListStart

Retrieves the starting value or code of a list numbering sequence.

```
FUNCTION CTextPara2.GetListStart () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetListStart(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The starting value or code of a list numbering sequence.

| Value | Meaning |
| ----- | ------- |
| **tomListNone** | Not a list paragraph. |
| **tomListBullet** | List uses bullets (&h2022); other bullets are given by > 32. |
| **tomListNumberAsArabic** | List is numbered with Arabic numerals (0, 1, 2, ...). |
| **tomListNumberAsLCLetter** | List is ordered with lowercase letters (a, b, c, ...). |
| **tomListNumberAsUCLetter** | List is ordered with uppercase Arabic letters (A, B, C, ...). |
| **tomListNumberAsLCRoman** | List is ordered with lowercase Roman letters (i, ii, iii, ...). |
| **tomListNumberAsUCRoman** | List is ordered with uppercase Roman letters (I, II, III, ...). |
| **tomListNumberAsSequence** | The value returned by **GetListStart** is treated as the first code in a Unicode sequence. |
| **tomListNumberedCircle** | List is ordered with Unicode circled numbers. |
| **tomListNumberedBlackCircleWingding** | List is ordered with Wingdings black circled digits. |
| **tomListNumberedArabicWide** | Full-width ASCII (０, １, ２, ３, …). |
| **tomListNumberedChS** | Chinese with 十 only in items 10 through 99. (一, 二, 三, 四…). |
| **tomListNumberedChT** | Chinese with 十 only in items 10 through 19. |
| **tomListNumberedJpnChs** | Chinese with a full-width period, no 十. |
| **tomListNumberedJpnKor** | Chinese with no 十. |
| **tomListNumberedArabic1** | Arabic alphabetic ( أ ,ب ,ت ,ث ,…). |
| **tomListNumberedArabic2** | Arabic abjadi ( أ ,ب ,ج ,د ,…). |
| **tomListNumberedHebrew** | Hebrew alphabet (א, ב, ג, ד, …). |
| **tomListNumberedThaiAlpha** | Thai alphabetic (ก, ข,ค, ง, …). |
| **tomListNumberedThaiNum** | Thai numbers (๑, ๒,๓, ๔…). |
| **tomListNumberedHindiAlpha** | Hindi vowels (अ, आ, इ, ई, …). |
| **tomListNumberedHindiAlpha1** | Hindi consonants (क, ख, ग, घ, …). |
| **tomListNumberedHindiNum** | Hindi numbers (१, २, ३, ४, …). |

By default, numbers are followed by a right parenthesis, for example: 1). However, the returnes value can include one of the following flags to indicate a different formatting.

| Value | Meaning |
| ----- | ------- |
| **tomListMinus** | Follows the number with a hyphen (-). |
| **tomListParentheses** | Encloses the number in parentheses, as in: (1). |
| **tomListPeriod** | Follows the number with a period. |
| **tomListPlain** | Uses the number alone. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

For a discussion on which sequence to use, see the **GetListType** method.

# <a name="SetListStart"></a>SetListStart

Sets the starting number or Unicode value for a numbered list.

```
FUNCTION CTextPara2.SetListStart (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetListStart(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | New starting number or Unicode value for a numbered list. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

Other characteristics of a list are specified by **SetListType**.

# <a name="GetListTab"></a>GetListTab

Retrieves the list tab setting, which is the distance between the first-line indent and the text on the first line. The numbered or bulleted text is left-justified, centered, or right-justified at the first-line indent value.

```
FUNCTION CTextPara2.GetListTab () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetListTab(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The list tab setting. The list tab value is in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="SetListTab"></a>SetListTab

Sets the list tab setting, which is the distance between the first indent and the start of the text on the first line.

```
FUNCTION CTextPara2.SetListTab (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetListTab(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | New list tab value, in floating-point points. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetListType"></a>GetListType

Retrieves the kind of numbering to use with paragraphs.

```
FUNCTION CTextPara2.GetListType () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetListType(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

One of the following values to indicate the kind of list numbering.

| Value | Meaning |
| ----- | ------- |
| **tomListNone** | Not a list paragraph. |
| **tomListBullet** | List uses bullets (&h2022); other bullets are given by > 32. |
| **tomListNumberAsArabic** | List is numbered with Arabic numerals (0, 1, 2, ...). |
| **tomListNumberAsLCLetter** | List is ordered with lowercase letters (a, b, c, ...). |
| **tomListNumberAsUCLetter** | List is ordered with uppercase Arabic letters (A, B, C, ...). |
| **tomListNumberAsLCRoman** | List is ordered with lowercase Roman letters (i, ii, iii, ...). |
| **tomListNumberAsUCRoman** | List is ordered with uppercase Roman letters (I, II, III, ...). |
| **tomListNumberAsSequence** | The value returned by **GetListStart** is treated as the first code in a Unicode sequence. |
| **tomListNumberedCircle** | List is ordered with Unicode circled numbers. |
| **tomListNumberedBlackCircleWingding** | List is ordered with Wingdings black circled digits. |
| **tomListNumberedArabicWide** | Full-width ASCII (０, １, ２, ３, …). |
| **tomListNumberedChS** | Chinese with 十 only in items 10 through 99. (一, 二, 三, 四…). |
| **tomListNumberedChT** | Chinese with 十 only in items 10 through 19. |
| **tomListNumberedJpnChs** | Chinese with a full-width period, no 十. |
| **tomListNumberedJpnKor** | Chinese with no 十. |
| **tomListNumberedArabic1** | Arabic alphabetic ( أ ,ب ,ت ,ث ,…). |
| **tomListNumberedArabic2** | Arabic abjadi ( أ ,ب ,ج ,د ,…). |
| **tomListNumberedHebrew** | Hebrew alphabet (א, ב, ג, ד, …). |
| **tomListNumberedThaiAlpha** | Thai alphabetic (ก, ข,ค, ง, …). |
| **tomListNumberedThaiNum** | Thai numbers (๑, ๒,๓, ๔…). |
| **tomListNumberedHindiAlpha** | Hindi vowels (अ, आ, इ, ई, …). |
| **tomListNumberedHindiAlpha1** | Hindi consonants (क, ख, ग, घ, …). |
| **tomListNumberedHindiNum** | Hindi numbers (१, २, ३, ४, …). |

By default, numbers are followed by a right parenthesis, for example: 1). However, the returnes value can include one of the following flags to indicate a different formatting.

| Value | Meaning |
| ----- | ------- |
| **tomListMinus** | Follows the number with a hyphen (-). |
| **tomListParentheses** | Encloses the number in parentheses, as in: (1). |
| **tomListPeriod** | Follows the number with a period. |
| **tomListPlain** | Uses the number alone. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

Values above 32 correspond to Unicode values for bullets.

# <a name="SetListType"></a>SetListType

Sets the type of list to be used for paragraphs.

```
FUNCTION CTextPara2.SetListType (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetListTab(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | New list type. |

| Value | Meaning |
| ----- | ------- |
| **tomListNone** | Not a list paragraph. |
| **tomListBullet** | List uses bullets (&h2022); other bullets are given by > 32. |
| **tomListNumberAsArabic** | List is numbered with Arabic numerals (0, 1, 2, ...). |
| **tomListNumberAsLCLetter** | List is ordered with lowercase letters (a, b, c, ...). |
| **tomListNumberAsUCLetter** | List is ordered with uppercase Arabic letters (A, B, C, ...). |
| **tomListNumberAsLCRoman** | List is ordered with lowercase Roman letters (i, ii, iii, ...). |
| **tomListNumberAsUCRoman** | List is ordered with uppercase Roman letters (I, II, III, ...). |
| **tomListNumberAsSequence** | The value returned by **GetListStart** is treated as the first code in a Unicode sequence. |
| **tomListNumberedCircle** | List is ordered with Unicode circled numbers. |
| **tomListNumberedBlackCircleWingding** | List is ordered with Wingdings black circled digits. |
| **tomListNumberedArabicWide** | Full-width ASCII (０, １, ２, ３, …). |
| **tomListNumberedChS** | Chinese with 十 only in items 10 through 99. (一, 二, 三, 四…). |
| **tomListNumberedChT** | Chinese with 十 only in items 10 through 19. |
| **tomListNumberedJpnChs** | Chinese with a full-width period, no 十. |
| **tomListNumberedJpnKor** | Chinese with no 十. |
| **tomListNumberedArabic1** | Arabic alphabetic ( أ ,ب ,ت ,ث ,…). |
| **tomListNumberedArabic2** | Arabic abjadi ( أ ,ب ,ج ,د ,…). |
| **tomListNumberedHebrew** | Hebrew alphabet (א, ב, ג, ד, …). |
| **tomListNumberedThaiAlpha** | Thai alphabetic (ก, ข,ค, ง, …). |
| **tomListNumberedThaiNum** | Thai numbers (๑, ๒,๓, ๔…). |
| **tomListNumberedHindiAlpha** | Hindi vowels (अ, आ, इ, ई, …). |
| **tomListNumberedHindiAlpha1** | Hindi consonants (क, ख, ग, घ, …). |
| **tomListNumberedHindiNum** | Hindi numbers (१, २, ३, ४, …). |

By default, numbers are followed by a right parenthesis, for example: 1). However, the returnes value can include one of the following flags to indicate a different formatting.

| Value | Meaning |
| ----- | ------- |
| **tomListMinus** | Follows the number with a hyphen (-). |
| **tomListParentheses** | Encloses the number in parentheses, as in: (1). |
| **tomListPeriod** | Follows the number with a period. |
| **tomListPlain** | Uses the number alone. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetNoLineNumber"></a>GetNoLineNumber

Determines whether paragraph numbering is enabled.

```
FUNCTION CTextPara2.GetNoLineNumber () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetNoLineNumber(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

One of the following values to indicate the kind of list numbering.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Line numbering is disabled. |
| **tomFalse** | Line numbering is enabled. |
| **tomUndefined** | The property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

Paragraph numbering is when the paragraphs of a range are numbered. The number appears on the first line of a paragraph.

# <a name="SetNoLineNumber"></a>SetNoLineNumber

Determines whether to suppress line numbering of paragraphs in a range.

```
FUNCTION CTextPara2.SetNoLineNumber (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetNoLineNumber(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | Indicates if line numbering is suppressed. It can be one of the following possible values. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Line numbering is disabled. |
| **tomFalse** | Line numbering is enabled. |
| **tomUndefined** | The property is undefined. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="GetPageBreakBefore"></a>GetPageBreakBefore

Determines whether each paragraph in the range must begin on a new page.

```
FUNCTION CTextPara2.GetPageBreakBefore () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextPara2->lpvtbl->GetPageBreakBefore(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

One of the following values:

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Each paragraph in this range must begin on a new page. |
| **tomFalse** | The paragraphs in this range do not need to begin on a new page. |
| **tomUndefined** | The property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="SetPageBreakBefore"></a>SetPageBreakBefore

Controls whether there is a page break before each paragraph in a range.

```
FUNCTION CTextPara2.SetPageBreakBefore (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetPageBreakBefore(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that controls page breaks before paragraphs. It can be one of the following values. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Paragraphs in this range must begin on a new page. |
| **tomFalse** | Paragraphs in this range do not need to begin on a new page. |
| **tomToggle** | Toggle the property value. |
| **tomUndefined** | The property is undefined. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

This method is included for compatibility with Microsoft Word; it does not affect how the rich edit control displays text.

# <a name="GetRightIndent"></a>GetRightIndent

Retrieves the size of the right margin indent of a paragraph.

```
FUNCTION CTextPara2.GetRightIndent () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetRightIndent(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The right indentation, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="SetRightIndent"></a>SetRightIndent

Sets the right margin of paragraph.

```
FUNCTION CTextPara2.SetRightIndent (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetRightIndent(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | Right indent, in floating-point points. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="SetIndents"></a>SetIndents

Sets the first-line indent, the left indent, and the right indent for a paragraph.

```
FUNCTION CTextPara2.SetIndents (BYVAL First AS SINGLE, BYVAL Left AS SINGLE, BYVAL Right AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetIndents(m_pTextPara2, First, Left, Right))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *First* | Indent of the first line in a paragraph, relative to the left indent. The value is in floating-point points and can be positive or negative. |
| *Left* | Left indent of all lines except the first line in a paragraph, relative to left margin. The value is in floating-point points and can be positive or negative. |
| *Right* | Right indent of all lines in paragraph, relative to the right margin. The value is in floating-point points and can be positive or negative. This value is optional. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

Line indents are not allowed to position text in the margins. If the first-line indent is set to a negative value (for an outdented paragraph) while the left indent is zero, the first-line indent is reset to zero. To avoid this problem while retaining property sets, set the first-line indent value equal to zero either explicitly or by calling the **Reset** method. Then, call **SetIndents** to set a nonnegative, left-indent value and set the desired first-line indent.

# <a name="SetLineSpacing"></a>SetLineSpacing

Sets the paragraph line-spacing rule and the line spacing for a paragraph.

```
FUNCTION CTextPara2.SetLineSpacing (BYVAL Rule AS LONG, BYVAL Spacing AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetLineSpacing(m_pTextPara2, Rule, Spacing))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Rule* | Value of new line-spacing rule.  |
| *Spacing* | Value of new line spacing. If the line-spacing rule treats the *Spacing* value as a linear dimension, then *Spacing* is given in floating-point points. |

| Line spacing rule | Meaning |
| ----------------- | ------- |
| **tomLineSpaceSingle** | Single spacing. The line-spacing value is ignored. |
| **tomLineSpace1pt5** | One-and-a-half line spacing. The line-spacing value is ignored. |
| **tomLineSpaceDouble** | Double line spacing. The line-spacing value is ignored. |
| **tomLineSpaceAtLeast** | The line-spacing value specifies the spacing, in floating-point points, from one line to the next. However, if the value is less than single spacing, the control displays single-spaced text. |
| **tomLineSpaceExactly** | The line-spacing value specifies the exact spacing, in floating-point points, from one line to the next (even if the value is less than single spacing). |
| **tomLineSpaceMultiple** | The line-spacing value specifies the line spacing, in lines. |
| **tomLineSpacePercent** | The line-spacing value specifies the line spacing by percent of line height. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

The line-spacing rule and line spacing work together, and as a result, they must be set together, much as the first and left indents need to be set together.

# <a name="GetSpaceAfter"></a>GetSpaceAfter

Retrieves the amount of vertical space below a paragraph.

```
FUNCTION CTextPara2.GetSpaceAfter () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextPara2->lpvtbl->GetSpaceAfter(m_pTextPara2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The space-after value, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

# <a name="SetSpaceAfter"></a>SetSpaceAfter

Sets the amount of space that follows a paragraph.

```
FUNCTION CTextPara2.SetSpaceAfter (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetSpaceAfter(m_pTextPara2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | New space-after value, in floating-point points. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |
