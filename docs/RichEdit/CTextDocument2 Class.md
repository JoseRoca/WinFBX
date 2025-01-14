# CTextDocument2 Class

### ITextDocument Interface

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

