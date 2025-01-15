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
| [SetFont](#SetFont) | Sets this range's character attributes to those of the specified ITextFont object. |
