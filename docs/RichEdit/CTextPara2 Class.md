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
| [GetDuplicate](#GetDuplicate) |  |
| [SetDuplicate](#SetDuplicate) |  |
| [CanChange](#CanChange) |  |
| [IsEqual](#IsEqual) |  |
| [Reset](#Reset) |  |
| [GetStyle](#GetStyle) |  |
| [SetStyle](#SetStyle) |  |
| [GetAlignment](#GetAlignment) |  |
| [SetAlignment](#SetAlignment) |  |
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

| [GetBorders](#GetBorders) |  |
| [GetDuplicate2](#GetDuplicate2) |  |
| [SetDuplicate2](#SetDuplicate2) |  |
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
| [IsEqual2](#IsEqual2) |  |
| [SetEffects](#SetEffects) |  |
| [SetProperty](#SetProperty) |  |
