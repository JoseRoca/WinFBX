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
| [GetHyphenation](#GetHyphenation) |  |
| [SetHyphenation](#SetHyphenation) |  |
| [GetFirstLineIndent](#GetFirstLineIndent) |  |
| [GetKeepTogether](#GetKeepTogether) |  |
| [SetKeepTogether](#SetKeepTogether) |  |
| [SetKeepWithNext](#SetKeepWithNext) |  |
| [GetLeftIndent](#GetLeftIndent) |  |
| [GetLineSpacing](#GetLineSpacing) |  |
| [GetLineSpacingRule](#GetLineSpacingRule) |  |
| [GetListAlignment](#GetListAlignment) |  |
| [SetListAlignment](#SetListAlignment) |  |
| [GetListLevelIndex](#GetListLevelIndex) |  |
| [SetListLevelIndex](#SetListLevelIndex) |  |
| [GetListStart](#GetListStart) |  |
| [SetListStart](#SetListStart) |  |
| [GetListTab](#GetListTab) |  |
| [SetListTab](#SetListTab) |  |
| [GetListType](#GetListType) |  |
| [SetListType](#SetListType) |  |
| [GetNoLineNumber](#GetNoLineNumber) |  |
| [SetNoLineNumber](#SetNoLineNumber) |  |
| [GetPageBreakBefore](#GetPageBreakBefore) |  |
| [SetPageBreakBefore](#SetPageBreakBefore) |  |
| [GetRightIndent](#GetRightIndent) |  |
| [SetRightIndent](#SetRightIndent) |  |
| [SetIndents](#SetIndents) |  |
| [SetLineSpacing](#SetLineSpacing) |  |
| [GetSpaceAfter](#GetSpaceAfter) |  |
| [SetSpaceAfter](#SetSpaceAfter) |  |
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
   this.SetResult(m_pTextPara2->lpvtbl->CanChange(m_pTextPara2, @Value))
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
| tomAlignLeft | Text aligns with the left margin. |
| tomAlignCenter | Text is centered between the margins. |
| tomAlignRight | Text aligns with the right margin. |
| tomAlignJustify | Text starts at the left margin and, if the line extends beyond the right margin, all the spaces in the line are adjusted to be even. |
| tomAlignInterWord | Same as **tomAlignJustify**. |
| tomAlignNewspaper | Same as **tomAlignInterLetter**, but uses East Asian metrics. |
| tomAlignInterLetter | The first and last characters of each line (except the last line) are aligned to the left and right margins, and lines are filled by adding or subtracting the same amount from each character. |
| tomAlignScaled | Same as **tomAlignInterLetter**, but uses East Asian metrics, and scales the spacing by the width of characters. |

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
| tomAlignLeft | Text aligns with the left margin. |
| tomAlignCenter | Text is centered between the margins. |
| tomAlignRight | Text aligns with the right margin. |
| tomAlignJustify | Text starts at the left margin and, if the line extends beyond the right margin, all the spaces in the line are adjusted to be even. |
| tomAlignInterWord | Same as **tomAlignJustify**. |
| tomAlignNewspaper | Same as **tomAlignInterLetter**, but uses East Asian metrics. |
| tomAlignInterLetter | The first and last characters of each line (except the last line) are aligned to the left and right margins, and lines are filled by adding or subtracting the same amount from each character. |
| tomAlignScaled | Same as **tomAlignInterLetter**, but uses East Asian metrics, and scales the spacing by the width of characters. |

#### Return value

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |
