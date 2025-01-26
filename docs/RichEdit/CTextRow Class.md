# CTextRow Class

Class that wraps all the methods of the **ITextRow** interface.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTORS](#CONSTRUCTORS) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [LET](#LET) | Assignment operator. |
| [CAST](#CAST) | Cast operator. |
| [TextDocumentPtr](#TextDocumentPtr) | Returns a pointer to the underlying **ITextRow** interface. |
| [Attach](#Attach) | Attaches an **ITextRow** interface pointer to the class. |
| [Detach](#Detach) | Detaches the underlying **ITextRow** interface pointer from the class. |

### ITextRow Interface

The **ITextRow** interface provides methods to insert one or more identical table rows, and to retrieve and change table row properties. To insert nonidentical rows, call **Insert** for each different row configuration.

The **ITextRow** interface inherits from the **IDispatch** interface. **ITextRow** also has these types of members:

| Name       | Description |
| ---------- | ----------- |
| [GetAlignment](#GetAlignment) |  |
| [SetAlignment](#SetAlignment) |  |
| [GetCellCount](#GetCellCount) |  |
| [SetCellCount](#SetCellCount) |  |
| [GetCellCountCache](#GetCellCountCache) |  |
| [SetCellCountCache](#SetCellCountCache) |  |
| [GetCellIndex](#GetCellIndex) |  |
| [SetCellIndex](#SetCellIndex) |  |
| [GetCellMargin](#GetCellMargin) |  |
| [SetCellMargin](#SetCellMargin) |  |
| [GetHeight](#GetHeight) |  |
| [SetHeight](#SetHeight) |  |
| [GetIndent](#GetIndent) |  |
| [SetIndent](#SetIndent) |  |
| [GetKeepTogether](#GetKeepTogether) |  |
| [SetKeepTogether](#SetKeepTogether) |  |
| [GetKeepWithNext](#GetKeepWithNext) |  |
| [SetKeepWithNext](#SetKeepWithNext) |  |
| [GetNestLevel](#GetNestLevel) |  |
| [GetRTL](#GetRTL) |  |
| [SetRTL](#SetRTL) |  |
| [GetCellAlignment](#GetCellAlignment) |  |
| [SetCellAlignment](#SetCellAlignment) |  |
| [GetCellColorBack](#GetCellColorBack) |  |
| [SetCellColorBack](#SetCellColorBack) |  |
| [GetCellColorFore](#GetCellColorFore) |  |
| [SetCellColorFore](#SetCellColorFore) |  |
| [GetCellMergeFlags](#GetCellMergeFlags) |  |
| [SetCellMergeFlags](#SetCellMergeFlags) |  |
| [GetCellShading](#GetCellShading) |  |
| [SetCellShading](#SetCellShading) |  |
| [GetCellVerticalText](#GetCellVerticalText) |  |
| [SetCellVerticalText](#SetCellVerticalText) |  |
| [GetCellWidth](#GetCellWidth) |  |
| [SetCellWidth](#SetCellWidth) |  |
| [GetCellBorderColors](#GetCellBorderColors) |  |
| [GetCellBorderWidths](#GetCellBorderWidths) |  |
| [SetCellBorderColors](#SetCellBorderColors) |  |
| [SetCellBorderWidths](#SetCellBorderWidths) |  |
| [Apply](#Apply) |  |
| [GetProperty](#GetProperty) |  |
| [Insert](#Insert) |  |
| [IsEqual](#IsEqual) |  |
| [Reset](#Reset) |  |
| [SetProperty](#SetProperty) |  |

### Methods inherited from CTOMBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#GetLastResult) | Returns the last result code |
| [SetResult](#SetResult) | Sets the last result code. |
| [GetErrorInfo](#GetErrorInfo) | Returns a description of the last result code. |

