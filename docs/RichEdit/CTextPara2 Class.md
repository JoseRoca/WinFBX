# CTextPara2 Class

Class that wraps all the methods of the **ITextPara** and **ITextPara2** interfaces.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTOR](#constructor) | Called when a class variable is created. |
| [DESTRUCTOR](#destructor) | Called automatically when a class variable goes out of scope or is destroyed. |

### ITextPara Interface

Text Object Model (TOM) rich text-range attributes are accessed through a pair of dual interfaces, **ITextFont** and **ITextPara**.

The **ITextPara** interface inherits from the **IDispatch** interface. **ITextPara** also has these types of members:

| Name       | Description |
| ---------- | ----------- |
| [GetDuplicate](#getduplicate) | Gets a duplicate of this text paragraph format object. |
| [SetDuplicate](#setduplicate) | Sets the formatting for an existing paragraph by copying a given format. |
| [CanChange](#canchange) | Determines whether the paragraph formatting can be changed. |
| [IsEqual](#isequal) | Determines if the current range has the same properties as a specified range. |
| [Reset](#reset) | Resets the paragraph formatting to a choice of default values. |
| [GetStyle](#getstyle) | Retrieves the style handle to the paragraphs in the specified range. |
| [SetStyle](#setstyle) | Sets the paragraph style for the paragraphs in a range. |
| [GetAlignment](#getalignment) | Retrieves the current paragraph alignment value. |
| [SetAlignment](#setalignment) | Sets the paragraph alignment. |
| [GetHyphenation](#gethyphenation) | Determines whether automatic hyphenation is enabled for the range. |
| [SetHyphenation](#sethyphenation) | Controls hyphenation for the paragraphs in the range. |
| [GetFirstLineIndent](#getfirstlineindent) | Retrieves the amount used to indent the first line of a paragraph relative to the left indent. The left indent is the indent for all lines of the paragraph except the first line. |
| [GetKeepTogether](#getkeeptogether) | Determines whether page breaks are allowed within paragraphs. |
| [SetKeepTogether](#setkeeptogether) | Controls whether page breaks are allowed within a paragraph in a range. |
| [GetKeepWithNext](#getkeepwithnext) | Determines whether page breaks are allowed between paragraphs in the range. |
| [SetKeepWithNext](#setkeepwithnext) | Controls whether page breaks are allowed between the paragraphs in a range. |
| [GetLeftIndent](#getleftindent) | Retrieves the distance used to indent all lines except the first line of a paragraph. The distance is relative to the left margin. |
| [GetLineSpacing](#getlinespacing) | Retrieves the line-spacing value for the text range. |
| [GetLineSpacingRule](#getlinespacingrule) | Retrieves the line-spacing rule for the text range. |
| [GetListAlignment](#getlistalignment) | Retrieves the kind of alignment to use for bulleted and numbered lists. |
| [SetListAlignment](#setlistalignment) | Sets the alignment of bulleted or numbered text used for paragraphs. |
| [GetListLevelIndex](#getlistlevelindex) | Retrieves the list level index used with paragraphs. |
| [SetListLevelIndex](#setlistlevelindex) | Sets the list level index used for paragraphs. |
| [GetListStart](#getliststart) | Retrieves the starting value or code of a list numbering sequence. |
| [SetListStart](#setliststart) | Sets the starting number or Unicode value for a numbered list. |
| [GetListTab](#getlisttab) | Retrieves the list tab setting, which is the distance between the first-line indent and the text on the first line. The numbered or bulleted text is left-justified, centered, or right-justified at the first-line indent value. |
| [SetListTab](#setlisttab) | Sets the list tab setting, which is the distance between the first indent and the start of the text on the first line. |
| [GetListType](#getlisttype) | Retrieves the kind of numbering to use with paragraphs. |
| [SetListType](#setlisttype) | Sets the type of list to be used for paragraphs. |
| [GetNoLineNumber](#getnolinenumber) | Determines whether paragraph numbering is enabled. |
| [SetNoLineNumber](#setnolinenumber) | Determines whether to suppress line numbering of paragraphs in a range. |
| [GetPageBreakBefore](#getpagebreakbefore) | Determines whether each paragraph in the range must begin on a new page. |
| [SetPageBreakBefore](#setpagebreakbefore) | Controls whether there is a page break before each paragraph in a range. |
| [GetRightIndent](#getrightindent) | Retrieves the size of the right margin indent of a paragraph. |
| [SetRightIndent](#setrightindent) | Sets the right margin of paragraph. |
| [SetIndents](#setindents) | Sets the first-line indent, the left indent, and the right indent for a paragraph. |
| [SetLineSpacing](#setlinespacing) | Sets the paragraph line-spacing rule and the line spacing for a paragraph. |
| [GetSpaceAfter](#getspaceafter) | The space-after value, in floating-point points. |
| [SetSpaceAfter](#setspaceafter) | Sets the amount of space that follows a paragraph. |
| [GetSpaceBefore](#getspacebefore) | Retrieves the amount of vertical space above a paragraph. |
| [SetSpaceBefore](#setspacebefore) | Sets the amount of space preceding a paragraph. |
| [GetWidowControl](#getwidowcontrol) | Retrieves the widow and orphan control state for the paragraphs in a range. |
| [SetWidowControl](#setwidowcontrol) | Controls the suppression of widows and orphans. |
| [GetTabCount](#gettabcount) | Retrieves the tab count. |
| [AddTab](#addtab) | Adds a tab at the displacement *tbPos*, with type *tbAlign*, and leader style, *tbLeader*. |
| [ClearAllTabs](#clearalltabs) | Clears all tabs, reverting to equally spaced tabs with the default tab spacing. |
| [DeleteTab](#deletetab) | Deletes a tab at a specified displacement. |
| [GetTab](#gettab) | Retrieves tab parameters (displacement, alignment, and leader style) for a specified tab. |

### ITextPara2 Interface

The **ITextPara2** interface extends ITextPara, providing the equivalent of the Microsoft Word format-paragraph dialog.

The **ITextPara2** interface has these methods.

| Name       | Description |
| ---------- | ----------- |
| [GetBorders](#getborders) | Not implemented. Gets the borders collection. |
| [GetDuplicate2](#getduplicate) | Gets a duplicate of this text paragraph format object. |
| [SetDuplicate2](#setduplicate) | Sets the formatting for an existing paragraph by copying a given format. |
| [GetFontAlignment](#getfontalignment) | Gets the paragraph font alignment state. |
| [SetFontAlignment](#setfontalignment) | Sets the paragraph font alignment for Chinese, Japanese, Korean text. |
| [GetHangingPunctuation](#gethangingpunctuation) | Gets whether to hang punctuation symbols on the right margin when the paragraph is justified. |
| [SetHangingPunctuation](#sethangingpunctuation) | Sets whether to hang punctuation symbols on the right margin when the paragraph is justified. |
| [GetSnapToGrid](#getsnaptogrid) | Gets whether paragraph lines snap to a vertical grid that could be defined for the whole document. |
| [SetSnapToGrid](#setsnaptogrid) | Sets whether paragraph lines snap to a vertical grid that could be defined for the whole document. |
| [GetTrimPunctuationAtStart](#gettrimpunctuationatstart) | Gets whether to trim the leading space of a punctuation symbol at the start of a line. |
| [SetTrimPunctuationAtStart](#settrimpunctuationatstart) | Sets whether to trim the leading space of a punctuation symbol at the start of a line. |
| [GetEffects](#geteffects) | Gets the paragraph format effects. |
| [GetProperty](#getproperty) | Gets the value of the specified property. |
| [IsEqual2](#isequal) | Determines if the current range has the same properties as a specified range. |
| [SetEffects](#seteffects) | Sets the paragraph format effects. |
| [SetProperty](#setproperty) | Sets the property value. |

### Methods inherited from CTOMBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#GetLastResult) | Returns the last result code |
| [SetResult](#SetResult) | Sets the last result code. |
| [GetErrorInfo](#GetErrorInfo) | Returns a description of the last result code. |

# <a name="constructor"></a>CONSTRUCTOR

Called when a **CTextPara2** class variable is created.

```
CONSTRUCTOR (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

| Parameter | Description |
| --------- | ----------- |
| *pTextPara2* | An **ITextPara2** interface pointer. |
| *fAddRef* | Optional. **TRUE** to increment the reference count of the passed **ITextPara2** interface pointer; otherwise, **FALSE**. Default is **FALSE**. |

#### Return value

A pointer to the new instance of the class.

---

# <a name="destructor"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

```
DESTRUCTOR
```

---

# <a name="getlastresult"></a>GetLastResult

Returns the last result code

```
FUNCTION GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
```
---

# <a name="setresult"></a>SetResult

Sets the last result code.

```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Result* | The **HRESULT** error code returned by the methods. |

---

# <a name="geterrorinfo"></a>GetErrorInfo

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS CWSTR
```
---

# <a name="getduplicate"></a>GetDuplicate

Gets a duplicate of this range object. In this implementation of the class, **GetDuplicate** and **GetDuplicate2** are the same method.

```
FUNCTION GetDuplicate () AS ITextPara2 PTR
```

#### Return value

The duplicate of the range.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_OUTOFMEMORY** | Memory could not be allocated for the new object. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="setduplicate"></a>SetDuplicate

Sets the properties of this object by copying the properties of another text paragraph object. In this implementation of the class, **SetDuplicate** and **SetDuplicate2** are the same method.

```
FUNCTION SetDuplicate (BYVAL pPara AS ITextPara2 PTR) AS HRESULT
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

---

# <a name="canchange"></a>CanChange

Determines whether the paragraph formatting can be changed.

```
FUNCTION CanChange (BYVAL pPara AS ITextPara2 PTR) AS LONG
```

#### Return value

**tomTrue** if the paragraph formatting can be changed or **tomFalse** if it cannot be changed.

#### Result code

If paragraph formatting can change, **CanChange** succeeds and **GetLastResult** returns **S_OK**. If paragraph formatting cannot change, the method fails and returns **S_FALSE**.

---

# <a name="isequal"></a>IsEqual

Determines if the current range has the same properties as a specified range.  In this implementation of the class, **IsEqual** and **IsEqual2** are the same method.

```
FUNCTION IsEqual (BYVAL pPara AS ITextPara2 PTR) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *pPara* | The **ITextPara2** range that is compared to the current range. |

#### Return value

A **tomBool** value that is **tomTrue** if the text paragraph objects have the same properties, or **tomFalse** if they don't. 

#### Result code

If paragraph formatting can change, **IsEqual** succeeds and **GetLastResult** returns **S_OK**. If paragraph formatting cannot change, the method fails and returns **S_FALSE**.

---

# <a name="reset"></a>Reset

Determines whether the paragraph formatting can be changed.

```
FUNCTION Reset (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | Type of reset. It can be one of the following possible values. |

| Value | Meaning |
| ----- | ------- |
| **tomDefault** | Used for paragraph formatting that is defined by the RTF \pard, that is, the paragraph default control word. |
| **tomUndefined** | Used for all undefined values. The **tomUndefined** value is only valid for duplicate (clone) **ITextPara2** objects. |

#### Return value

If **Reset** succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| E_ACCESSDENIED | Write access is denied. |
| E_OUTOFMEMORY | Insufficient memory. |
| CO_E_RELEASED | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="getstyle"></a>GetStyle

Retrieves the style handle to the paragraphs in the specified range.

```
FUNCTION GetStyle () AS LONG
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

---

# <a name="setstyle"></a>SetStyle

Sets the paragraph style for the paragraphs in a range.

```
FUNCTION SetStyle (BYVAL Value AS LONG) AS HRESULT
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

---

# <a name="getalignment"></a>GetAlignment

Retrieves the current paragraph alignment value.

```
FUNCTION GetAlignment () AS LONG
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

---

# <a name="setalignment"></a>SetAlignment

Sets the paragraph alignment.

```
FUNCTION SetAlignment (BYVAL Value AS LONG) AS HRESULT
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

---

# <a name="gethyphenation"></a>GetHyphenation

Determines whether automatic hyphenation is enabled for the range.

```
FUNCTION GetHyphenation () AS LONG
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

---

# <a name="sethyphenation"></a>SetHyphenation

Controls hyphenation for the paragraphs in the range.

```
FUNCTION SetHyphenation (BYVAL Value AS LONG) AS HRESULT
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

---

# <a name="getfirstlineindent"></a>GetFirstLineIndent

Retrieves the amount used to indent the first line of a paragraph relative to the left indent. The left indent is the indent for all lines of the paragraph except the first line.

```
FUNCTION GetFirstLineIndent () AS SINGLE
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

---

# <a name="getkeeptogether"></a>GetKeepTogether

Determines whether page breaks are allowed within paragraphs.

```
FUNCTION GetKeepTogether () AS LONG
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

---

# <a name="setkeeptogether"></a>SetKeepTogether

Controls whether page breaks are allowed within a paragraph in a range.

```
FUNCTION SetKeepTogether (BYVAL Value AS LONG) AS HRESULT
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

---

# <a name="getkeepwithnext"></a>GetKeepWithNext

Determines whether page breaks are allowed between paragraphs in the range.

```
FUNCTION GetKeepWithNext () AS LONG
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

---

# <a name="setkeepwithnext"></a>SetKeepWithNext

Controls whether page breaks are allowed between the paragraphs in a range.

```
FUNCTION SetKeepWithNext (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextPara2->lpvtbl->SetKeepWithNext(m_pTextPara2, Value))
   RETURN m_Result
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

---

# <a name="getleftindent"></a>GetLeftIndent

Retrieves the distance used to indent all lines except the first line of a paragraph. The distance is relative to the left margin.

```
FUNCTION GetLeftIndent () AS SINGLE
```
#### Return value

The left indentation, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="getlinespacing"></a>GetLineSpacing

Retrieves the line-spacing value for the text range.

```
FUNCTION GetLineSpacing () AS SINGLE
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

---

# <a name="getlinespacingrule"></a>GetLineSpacingRule

Retrieves the line-spacing rule for the text range.

```
FUNCTION GetLineSpacingRule () AS LONG
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

---

# <a name="getlistalignment"></a>GetListAlignment

Retrieves the kind of alignment to use for bulleted and numbered lists.

```
FUNCTION GetListAlignment () AS LONG
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

---

# <a name="setlistalignment"></a>SetListAlignment

Sets the alignment of bulleted or numbered text used for paragraphs.

```
FUNCTION SetListAlignment (BYVAL Value AS LONG) AS HRESULT
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

---

# <a name="getlistlevelindex"></a>GetListLevelIndex

Retrieves the list level index used with paragraphs.

```
FUNCTION GetListLevelIndex () AS LONG
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

---

# <a name="setlistlevelindex"></a>SetListLevelIndex

Sets the list level index used for paragraphs.

```
FUNCTION SetListLevelIndex (BYVAL Value AS LONG) AS HRESULT
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

---

# <a name="getliststart"></a>GetListStart

Retrieves the starting value or code of a list numbering sequence.

```
FUNCTION GetListStart () AS LONG
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

---

# <a name="setliststart"></a>SetListStart

Sets the starting number or Unicode value for a numbered list.

```
FUNCTION SetListStart (BYVAL Value AS LONG) AS HRESULT
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

---

# <a name="getlisttab"></a>GetListTab

Retrieves the list tab setting, which is the distance between the first-line indent and the text on the first line. The numbered or bulleted text is left-justified, centered, or right-justified at the first-line indent value.

```
FUNCTION GetListTab () AS SINGLE
```
#### Return value

The list tab setting. The list tab value is in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="setlisttab"></a>SetListTab

Sets the list tab setting, which is the distance between the first indent and the start of the text on the first line.

```
FUNCTION SetListTab (BYVAL Value AS SINGLE) AS HRESULT
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

---

# <a name="getlisttype"></a>GetListType

Retrieves the kind of numbering to use with paragraphs.

```
FUNCTION GetListType () AS LONG
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

---

# <a name="setlisttype"></a>SetListType

Sets the type of list to be used for paragraphs.

```
FUNCTION SetListType (BYVAL Value AS LONG) AS HRESULT
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

---

# <a name="getnolinenumber"></a>GetNoLineNumber

Determines whether paragraph numbering is enabled.

```
FUNCTION GetNoLineNumber () AS LONG
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

---

# <a name="setnolinenumber"></a>SetNoLineNumber

Determines whether to suppress line numbering of paragraphs in a range.

```
FUNCTION SetNoLineNumber (BYVAL Value AS LONG) AS HRESULT
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

---

# <a name="getpagebreakbefore"></a>GetPageBreakBefore

Determines whether each paragraph in the range must begin on a new page.

```
FUNCTION GetPageBreakBefore () AS LONG
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

---

# <a name="setpagebbreakbefore"></a>SetPageBreakBefore

Controls whether there is a page break before each paragraph in a range.

```
FUNCTION SetPageBreakBefore (BYVAL Value AS LONG) AS HRESULT
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

---

# <a name="getrightindent"></a>GetRightIndent

Retrieves the size of the right margin indent of a paragraph.

```
FUNCTION GetRightIndent () AS SINGLE
```
#### Return value

The right indentation, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="setrightindent"></a>SetRightIndent

Sets the right margin of paragraph.

```
FUNCTION SetRightIndent (BYVAL Value AS SINGLE) AS HRESULT
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

---

# <a name="setindents"></a>SetIndents

Sets the first-line indent, the left indent, and the right indent for a paragraph.

```
FUNCTION SetIndents (BYVAL First AS SINGLE, BYVAL Left AS SINGLE, BYVAL Right AS SINGLE) AS HRESULT
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

---

# <a name="setlinespacing"></a>SetLineSpacing

Sets the paragraph line-spacing rule and the line spacing for a paragraph.

```
FUNCTION SetLineSpacing (BYVAL Rule AS LONG, BYVAL Spacing AS SINGLE) AS HRESULT
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

---

# <a name="getspaceafter"></a>GetSpaceAfter

Retrieves the amount of vertical space below a paragraph.

```
FUNCTION GetSpaceAfter () AS SINGLE
```
#### Return value

The space-after value, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="setspaceafter"></a>SetSpaceAfter

Sets the amount of space that follows a paragraph.

```
FUNCTION SetSpaceAfter (BYVAL Value AS SINGLE) AS HRESULT
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

---

# <a name="getspacebefore"></a>GetSpaceBefore

Retrieves the amount of vertical space above a paragraph.

```
FUNCTION GetSpaceBefore () AS SINGLE
```
#### Return value

The space-before value, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="setspacebefore"></a>SetSpaceBefore

Sets the amount of space preceding a paragraph.

```
FUNCTION SetSpaceBefore (BYVAL Value AS SINGLE) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | New space-before value, in floating-point points. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="getwidowcontrol"></a>GetWidowControl

Retrieves the widow and orphan control state for the paragraphs in a range.

```
FUNCTION GetWidowControl () AS LONG
```
#### Return value

A **tomBool** value that indicates the state of widow and orphan control. It can be one of the following values.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Prevents the printing of a widow or orphan |
| **tomFalse** | Allows the printing of a widow or orphan. |
| **tomUndefined** | The widow-control property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="setwidowcontrol"></a>SetWidowControl

Controls the suppression of widows and orphans.

```
FUNCTION SetWidowControl (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A tomBool value that controls the suppression of widows and orphans. It can be one of the following possible values. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Prevents printing of widows and orphans. |
| **tomFalse** | Allows printing of widows and orphans. |
| **tomToggle** | The value is toggled. |
| **tomUndefined** | No change. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="gettabcount"></a>GetTabCount

Retrieves the tab count.

```
FUNCTION GetTabCount () AS LONG
```
#### Return value

The tab count.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Prevents the printing of a widow or orphan. |
| **tomFalse** | Allows the printing of a widow or orphan. |
| **tomUndefined** | The widow-control property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

The tab count of a new instance can be nonzero, depending on the underlying text engine. For example, Microsoft Word stories begin with no explicit tabs defined, while rich edit instances start with a single explicit tab. To be sure there are no explicit tabs (that is, to set the tab count to zero), call **ClearAllTabs**.

---

# <a name="addtab"></a>AddTab

Adds a tab at the displacement *tbPos*, with type *tbAlign*, and leader style, *tbLeader*.

```
FUNCTION AddTab (BYVAL tbPos AS SINGLE, BYVAL tbAlign AS LONG, BYVAL tbLeader AS LONG) AS HRESULT
```

| Parameter | Description |
| ----- | ------- |
| **tbPos** | New tab displacement, in floating-point points. |
| **tbAlign** | Alignment options for the tab position. It can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomAlignLeft** | Text is left justified from the tab position. This is the default. |
| **tomAlignCenter** | Text is centered on the tab position. |
| **tomAlignDecimal** | The decimal point is set at the tab position. This is useful for aligning a column of decimal numbers. |
| **tomAlignBar** | A vertical bar is positioned at the tab position. Text is not affected. Alignment bars on nearby lines at the same position form a continuous vertical line. |

| Parameter | Description |
| ----- | ------- |
| **tbLeader** | Leader character style. A leader character is the character that is used to fill the space taken by a tab character. It can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomSpaces** | Spaces are used. This is the default. |
| **tomDots** | Dots are used. |
| **tomDashes** | A dashed line is used. |
| **tomLines** | A solid line is used. |

#### Return value

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

#### Remarks

It is assumed that there is never a tab at position zero. If multiple paragraphs are described, the common subset of tabs will be returned with &h8000 in the upper word of the tab type.

---

# <a name="clearalltabs"></a>ClearAllTabs

Clears all tabs, reverting to equally spaced tabs with the default tab spacing.

```
FUNCTION ClearAllTabs () AS HRESULT
```
#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="deletetab"></a>DeleteTab

Deletes a tab at a specified displacement.

```
FUNCTION DeleteTab (BYVAL tbPos AS SINGLE) AS HRESULT
```

| Parameter | Description |
| ----- | ------- |
| **tbPos** | Displacement, in floating-point points, at which a tab should be deleted. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |

---

# <a name="gettab"></a>GetTab

Retrieves tab parameters (displacement, alignment, and leader style) for a specified tab.

```
FUNCTION GetTab (BYVAL iTab AS LONG, BYVAL ptbPos AS SINGLE PTR, BYVAL ptbAlign AS LONG PTR, _
   BYVAL ptbLeader AS LONG PTR) AS HRESULT
```

| Parameter | Description |
| ----- | ------- |
| **iTab** | Index of tab for which to retrieve info. It can be either a numerical index or a special value (see the following table). Since tab indexes are zero-based, *iTab* = zero gets the first tab defined, *iTab = 1 gets the second tab defined, and so forth. The following table summarizes all of the possible values of ^*iTab*. |

| iTab | Value | Meaning |
| ---- | ----- | ------- |
| **tomTabBack** | -3 | Get tab previous to * *ptbPos* |
| **tomTabNext** | -2 | Get tab following * *ptbPos* |
| **tomTabHere** | -1 | Get tab at * *ptbPos* |
|  | >= 0 | Get tab with index of iTab (and ignore *ptbPos*). |

| Parameter | Description |
| ----- | ------- |
| **ptbPos** | The tab displacement, in floating-point points. The value of * *ptbPos* is zero if the tab does not exist and the value of * *ptbPos* is tomUndefined if there are multiple values in the associated range. |
| **ptbAlign** | The tab alignment. |

| Value | Meaning |
| ----- | ------- |
| **tomAlignLeft** | Text is left justified from the tab position. This is the default. |
| **tomAlignCenter** | Text is centered on the tab position. |
| **tomAlignDecimal** | The decimal point is set at the tab position. This is useful for aligning a column of decimal numbers. |
| **tomAlignBar** | A vertical bar is positioned at the tab position. Text is not affected. Alignment bars on nearby lines at the same position form a continuous vertical line. |

| Parameter | Description |
| ----- | ------- |
| **ptbLeader** | The tab leader-character style. |

| Value | Meaning |
| ----- | ------- |
| **tomSpaces** | Spaces are used. This is the default. |
| **tomDots** | Dots are used. |
| **tomDashes** | A dashed line is used. |
| **tomLines** | A solid line is used. |

#### Return value

If **GetTab** succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The paragraph formatting object is attached to a range that has been deleted. |
| **S_FALSE** | There is no tab corresponding to iTab. |

---

# <a name="getborders"></a>GetBorders

Not implemented. Gets the borders collection.

```
FUNCTION GetBorders () AS IUnknown PTR
```

#### Return value

The borders collection.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

# <a name="getfontalignment"></a>GetFontAlignment

Gets the paragraph font alignment state.

```
FUNCTION GetFontAlignment () AS LONG
```

#### Return value

The paragraph font alignment state. It can be one of the following values.

| Font Alignment | States |
| -------------- | ------ |
| **tomFontAlignmentAuto** (default) | For horizontal layout, align CJK characters on the baseline. For vertical layout, center align CJK characters. |
| **tomFontAlignmentTop** | For horizontal layout, top align CJK characters. For vertical layout, right align CJK characters. |
| **tomFontAlignmentBaseline** | For horizontal or vertical layout, align CJK characters on the baseline. |
| **tomFontAlignmentBottom** | For horizontal layout, bottom align CJK characters. For vertical layout, left align CJK characters. |
| **tomFontAlignmentCenter** | For horizontal layout, center CJK characters vertically. For vertical layout, center align CJK characters horizontally. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

# <a name="setfontalignment"></a>SetFontAlignment

Sets the paragraph font alignment for Chinese, Japanese, Korean text.

```
FUNCTION SetFontAlignment (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| ----- | ------- |
| **Value** | The paragraph font alignment. It can be one of the following values. |

#### Return value

The paragraph font alignment state. It can be one of the following values.

| Font Alignment | States |
| -------------- | ------ |
| **tomFontAlignmentAuto** (default) | For horizontal layout, align CJK characters on the baseline. For vertical layout, center align CJK characters. |
| **tomFontAlignmentTop** | For horizontal layout, top align CJK characters. For vertical layout, right align CJK characters. |
| **tomFontAlignmentBaseline** | For horizontal or vertical layout, align CJK characters on the baseline. |
| **tomFontAlignmentBottom** | For horizontal layout, bottom align CJK characters. For vertical layout, left align CJK characters. |
| **tomFontAlignmentCenter** | For horizontal layout, center CJK characters vertically. For vertical layout, center align CJK characters horizontally. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

# <a name="gethangingpunctuation"></a>GetHangingPunctuation

Gets whether to hang punctuation symbols on the right margin when the paragraph is justified.

```
FUNCTION GetHangingPunctuation () AS LONG
```

#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Hang punctuation symbols on the right margin. |
| **tomFalse** | Do not hang punctuation symbols on the right margin. |
| **tomUndefined** | The HangingPunctuation property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

# <a name="sethangingpunctuation"></a>SetHangingPunctuation

Sets whether to hang punctuation symbols on the right margin when the paragraph is justified.

```
FUNCTION SetHangingPunctuation (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| ----- | ------- |
| **Value** | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Hang punctuation symbols on the right margin. |
| **tomFalse** | Do not hang punctuation symbols on the right margin. |
| **tomToggle** | Toggle the HangingPunctuation property. |
| **tomUndefined** | The HangingPunctuation property is undefined. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

# <a name="getsnapogrid"></a>GetSnapToGrid

Gets whether paragraph lines snap to a vertical grid that could be defined for the whole document.

```
FUNCTION GetSnapToGrid () AS LONG
```

#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Paragraph lines snap to a vertical grid. |
| **tomFalse** | Paragraph lines do not snap to a grid. |
| **tomUndefined** | The SnapToGrid property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

# <a name="setsnaptogrid"></a>SetSnapToGrid

Sets whether paragraph lines snap to a vertical grid that could be defined for the whole document.

```
FUNCTION SetSnapToGrid (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| ----- | ------- |
| **Value** | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Paragraph lines snap to a vertical grid. |
| **tomFalse** | Paragraph lines do not snap to a grid. |
| **tomToggle** | Toggle the SnapToGrid property. |
| **tomUndefined** | The SnapToGrid property is undefined. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

# <a name="gettrimpunctuationatstart"></a>GetTrimPunctuationAtStart

Gets whether to trim the leading space of a punctuation symbol at the start of a line.

```
FUNCTION GetTrimPunctuationAtStart () AS LONG
```

#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Trim the leading space of a punctuation symbol at the start of a line. |
| **tomFalse** | Do not trim the leading space of a punctuation symbol at the start of a line. |
| **tomUndefined** | The TrimPunctuationAtStart property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

# <a name="settrimpunctuationatstart"></a>SetTrimPunctuationAtStart

Sets whether to trim the leading space of a punctuation symbol at the start of a line.

```
FUNCTION SetTrimPunctuationAtStart (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| ----- | ------- |
| **Value** | A **tomBool** that indicates whether to trim the leading space of a punctuation symbol. It can be one of the following values. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Trim the leading space of a punctuation symbol at the start of a line. |
| **tomFalse** | Do not trim the leading space of a punctuation symbol at the start of a line. |
| **tomToggle** | Toggle the TrimPunctuationAtStart property. |
| **tomUndefined** | The TrimPunctuationAtStart property is undefined. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

# <a name="geteffects"></a>GetEffects

Gets the paragraph format effects.

```
FUNCTION GetEffects (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
```

| Parameter | Description |
| ----- | ------- |
| **pValue** | The paragraph effects. This value can be a combination of the following flags. |

| Flag | Meaning |
| ---- | ------- |
| **tomParaEffectRTL** | Right-to-left paragraph. |
| **tomParaEffectKeep** | Keep the paragraph together. |
| **tomParaEffectPageBreakBefore** | Put a page break before this paragraph. |
| **tomParaEffectNoLineNumber** | No line number for this paragraph. |
| **tomParaEffectNoWidowControl** | No widow control. |
| **tomParaEffectDoNotHyphen** | Don't hyphenate this paragraph. |
| **tomParaEffectSideBySide** | Side by side. |
| **tomParaEffectCollapsed** | Heading contents are collapsed (in outline view). |
| **tomParaEffectOutlineLevel** | Outline view nested level. |
| **tomParaEffectBox** | Paragraph has boxed effect (is not displayed). |
| **tomParaEffectTableRowDelimiter** | At or inside table delimiter. |
| **tomParaEffectTable** | Inside or at the start of a table. |

| Parameter | Description |
| ----- | ------- |
| **pMask** | The differences in the flags over the range. A value of 1 indicates that the corresponding effect is the same over the range. For an insertion point, the values equal 1 for all defined effects. |

#### Return value

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

If the **tomTable** flag is set, you can use the **GetTable** method to get more table properties.

---

# <a name="seteffects"></a>SetEffects

Sets the paragraph format effects.

```
FUNCTION SetEffects (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
```

| Parameter | Description |
| ----- | ------- |
| **Value** | The paragraph effects value. This value can be a combination of the flags defined in the table below. |
| **Mask** | The desired mask. This value can be a combination of the flags defined in the table below. Only effects with the corresponding mask flag set are modified. |

| Flag | Meaning |
| ---- | ------- |
| **tomParaEffectRTL** | Right-to-left paragraph. |
| **tomParaEffectKeep** | Keep the paragraph together. |
| **tomParaEffectPageBreakBefore** | Put a page break before this paragraph. |
| **tomParaEffectNoLineNumber** | No line number for this paragraph. |
| **tomParaEffectNoWidowControl** | No widow control. |
| **tomParaEffectDoNotHyphen** | Don't hyphenate this paragraph. |
| **tomParaEffectSideBySide** | Side by side. |
| **tomParaEffectCollapsed** | Heading contents are collapsed (in outline view). |
| **tomParaEffectOutlineLevel** | Outline view nested level. |
| **tomParaEffectBox** | Paragraph has boxed effect (is not displayed). |
| **tomParaEffectTableRowDelimiter** | At or inside table delimiter. |
| **tomParaEffectTable** | Inside or at the start of a table. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

# <a name="getproperty"></a>GetProperty

Gets the value of the specified property.

```
FUNCTION GetProperty (BYVAL nType AS LONG) AS LONG
```

| Parameter | Description |
| ----- | ------- |
| **nType** | The ID of the property value to retrieve. |

#### Return value

The property value.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The **tomParaPropMathAlign** property sets the math alignment for math paragraphs in a text paragraph. It can have one of the following values.

| Value | Meaning |
| ----- | ------- |
| **tomMathParaAlignDefault** | The default alignment for math paragraphs. |
| **tomMathParaAlignCenterGroup** | Center math paragraphs as a group. |
| **tomMathParaAlignCenter** | Center math paragraphs. |
| **tomMathParaAlignLeft** | Left-align math paragraphs. |
| **tomMathParaAlignRight** | Right-align math paragraphs. |

---

# <a name="setproperty"></a>SetProperty

Sets the value of the specified property.

```
FUNCTION SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| ----- | ------- |
| **nType** | The ID of the property value to set. |
| **nType** | The property value to set. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The **tomParaPropMathAlign** property sets the math alignment for math paragraphs in a text paragraph. It can have one of the following values.

| Value | Meaning |
| ----- | ------- |
| **tomMathParaAlignDefault** | The default alignment for math paragraphs. |
| **tomMathParaAlignCenterGroup** | Center math paragraphs as a group. |
| **tomMathParaAlignCenter** | Center math paragraphs. |
| **tomMathParaAlignLeft** | Left-align math paragraphs. |
| **tomMathParaAlignRight** | Right-align math paragraphs. |

---
