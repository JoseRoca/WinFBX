# CRichEditOleCallback Class

Implements the IRichEditOleCallback interface.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTOR](#CONSTRUCTOR) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |

The `IRichEditOleCallback`interface is used by a rich text edit control to retrieve OLE-related information from its client. A rich edit control client is responsible for implementing this interface and assigning it to the control by using the **EM_SETOLECALLBACK** message.

The `IRichEditOleCallback`interface inherits from the **IUnknown** interface. `IRichEditOleCallback`also has these types of methods:

# CRichEditOleCAllback Methods

| Name       | Description |
| ---------- | ----------- |
| [GetNewStorage](#GetNewStorage) | Provides storage for a new object pasted from the clipboard or read in from an Rich Text Format (RTF) stream. |
| [GetInPlaceContext](#GetInPlaceContext) | Provides the application and document-level interfaces and information required to support in-place activation. |
| [ShowContainerUI](#ShowContainerUI) | Indicates whether or not the application is to display its container UI. The rich edit control looks ahead for double-clicks and defers the call if appropriate. Applications may defer hiding adornments until an [IOleInPlaceUIWindow.SetBorderSpace](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-ioleinplaceuiwindow-setborderspace) call is received. |
| [QueryInsertObject](#QueryInsertObject) | Queries the application as to whether an object should be inserted. The member is called when pasting and when reading Rich Text Format (RTF). |
| [DeleteObject](#DeleteObject) | Sends notification that an object is about to be deleted from a rich edit control. The object is not necessarily being released when this member is called. |
| [QueryAcceptData](#QueryAcceptData) | During a paste operation or a drag event, determines if the data that is pasted or dragged should be accepted. |
| [ContextSensitiveHelp](#ContextSensitiveHelp) | Indicates if the application should transition into or out of context-sensitive help mode. This method should implement the functionality described for [IOleWindow.ContextSensitiveHelp](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-iolewindow-contextsensitivehelp). |
| [GetClipboardData](#GetClipboardData) | Allows the client to supply its own clipboard object. |
| [GetDragDropEffect](#GetDragDropEffect) | Allows the client to specify the effects of a drop operation. |
| [GetContextMenu](#GetContextMenu) | Queries the application for a context menu to use on a right-click event. |
