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
| [SetIMEInProgress](#SetIMEInProgress) | Sets the state of the Input Method Editor (IME) in-progress flag. |
| [GetNotificationMode](#GetNotificationMode) | Gets the notification mode. |
| [SetNotificationMode](#SetNotificationMode) | Sets the notification mode. Use **tomTrue** to turn on notifications, or **tomFalse** to turn them off. |
| [GetSelection2](#GetSelection2) | Gets the active selection. |
| [GetStoryRanges2](#GetStoryRanges2) | Gets an object for enumerating the stories in a document. |
| [GetTypographyOptions](#GetTypographyOptions) | Gets the typography options. |
| [GetVersion](#GetVersion) | Gets the version number of the Text Object Model (TOM) engine. |
| [GetWindow](#GetWindow) | Gets the handle of the window that the Text Object Model (TOM) engine is using to display output. |
| [AttachMsgFilter](#AttachMsgFilter) | Attaches a new message filter to the edit instance. All window messages that the edit instance receives are forwarded to the message filter. |
| [CheckTextLimit](#CheckTextLimit) | Checks whether the number of characters to be added would exceed the maximum text limit. |
| [GetCallManager](#GetCallManager) | Gets the call manager. |
| [GetClientRect](#GetClientRect) | Retrieves the client rectangle of the rich edit control. |
| [GetEffectColor](#GetEffectColor) | Retrieves the color used for special text attributes. |
| [GetImmContext](#GetImmContext) | Gets the Input Method Manager (IMM) input context from the Text Object Model (TOM) host. |
| [GetPreferredFont](#GetPreferredFont) | Retrieves the preferred font for a particular character repertoire and character position. |
| [GetProperty](#GetProperty) | Retrieves the value of a property. |
| [GetStrings](#GetStrings) | Gets a collection of rich-text strings. |
| [Notify](#Notify) | Notifies the Text Object Model (TOM) engine client of particular Input Method Editor (IME) events. |
| [Range2](#Range2) | Retrieves a new text range for the active story of the document. |
| [RangeFromPoint2](#RangeFromPoint2) | Retrieves the degenerate range at (or nearest to) a particular point on the screen. |
| [ReleaseCallManager](#ReleaseCallManager) | Releases the call manager. |
| [ReleaseImmContext](#ReleaseImmContext) | Releases an Input Method Manager (IMM) input context. |
| [SetEffectColor](#SetEffectColor) | Specifies the color to use for special text attributes. |
| [SetProperty](#SetProperty) | Specifies a new value for a property. |
| [SetTypographyOptions](#SetTypographyOptions) | Specifies the typography options for the document. |
| [SysBeep](#SysBeep) | Generates a system beep. |
| [Update](#Update) | Updates the selection and caret. |
| [UpdateWindow](#UpdateWindow) | Notifies the client that the view has changed and the client should update the view if the Text Object Model (TOM) engine is in-place active. |
| [GetMathProperties](#GetMathProperties) | Gets the math properties for the document. |
| [SetMathProperties](#SetMathProperties) | Specifies the math properties to use for the document. |
| [GetActiveStory](#GetActiveStory) | Gets the active story; that is, the story that receives keyboard and mouse input. |
| [SetActiveStory](#SetActiveStory) | Sets the active story; that is, the story that receives keyboard and mouse input. |
| [GetMainStory](#GetMainStory) | Gets the main story. |
| [GetNewStory](#GetNewStory) | Gets a new story. Not implemented. |
| [GetStory](#GetStory) | Retrieves the story that corresponds to a particular index. |

# <a name="GetName"></a>GetName

Gets the file name of this document. This is the ITextDocument default property.

```
PRIVATE FUNCTION CTextDocument2.GetName () AS CBSTR
   DIM pName AS AFX_BSTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetName(m_pTextDocument2, @pName))
   RETURN pName
END FUNCTION
```
#### Return value

The filename of this document, or an emmpty setring if there is not a filename associated with this object.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **S_FALSE** | No file name associated with this object. |
| **E_INVALIDARG** | Invalid argument. |
| **E_OUTOFMEMORY** | Insufficient memory for output string. |

# <a name="GetSelection"></a>GetSelection

Gets the active selection.

```
PRIVATE FUNCTION CTextDocument2.GetSelection () AS ITextSelection PTR
   DIM pSelection AS ITextSelection PTR
   this.SetResult(m_pTextDocument2->lpvtbl->GetSelection(m_pTextDocument2, @pSelection))
   RETURN pSelection
END FUNCTION
```
#### Return value

The **ITextSelection** pointer of the active selection.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **S_FALSE** | Indicates no active selection. |
| **E_INVALIDARG** | Invalid argument. |
