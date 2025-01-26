# CTextStory Class

Class that wraps all the methods of the **ITextStory** interface.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTORS](#CONSTRUCTORS) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [LET](#LET) | Assignment operator. |
| [CAST](#CAST) | Cast operator. |
| [TextDocumentPtr](#TextDocumentPtr) | Returns a pointer to the underlying **ITextStory** interface. |
| [Attach](#Attach) | Attaches an **ITextStory** interface pointer to the class. |
| [Detach](#Detach) | Detaches the underlying **ITextStory** interface pointer from the class. |

### ITextStory Interface

The **ITextStory** interface methods are used to access shared data from multiple stories, which is stored in the parent **ITextServices** instance.

The stories can be "edited" simultaneously by using individual **ITextRange2** methods, and displayed independently of one another. In addition, one story at a time can be UI active; that is, it receives keyboard and mouse input.

The **ITextStory** is a lightweight interface that does not require an **ITextRange2** object. This allows the client to manipulate a story, which is a faster, smaller object than a complete editing instance.

| Name       | Description |
| ---------- | ----------- |
| [GetActive](#GetActive) | Gets the active state of a story. |
| [SetActive](#SetActive) | Sets the active state of a story. |
| [GetDisplay](#GetDisplay) | Gets a new display for a story. |
| [GetIndex](#GetIndex) | Gets the index of a story. |
| [GetType](#GetType) | Gets this story's type. |
| [SetType](#SetType) | Sets this story's type. |
| [GetProperty](#GetProperty) | Gets the value of the specified property. |
| [GetRange](#GetRange) | Gets a text range object for the story. |
| [GetText](#GetText) | Gets the text in a story according to the specified conversion flags. |
| [SetFormattedText](#SetFormattedText) | Replaces a story’s text with specified formatted text. |
| [SetProperty](#SetProperty) | Sets the value of the specified property. |
| [SetText](#SetText) | Sets the text in this range. |
