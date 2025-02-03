# CRichEditOleCallback Class

Implements the `IRichEditOleCallback`interface.

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

# <a name="CONSTRUCTOR"></a>CONSTRUCTOR

Creates an instance of the `CRichEditOleCallback`.

```
CONSTRUCTOR CRichEditOLeCallback
```

# <a name="DESTRUCTOR"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.
```
DESTRUCTOR CRichEditOLeCallback
```

# <a name="GetNewStorage"></a>GetNewStorage

Provides storage for a new object pasted from the clipboard or read in from an Rich Text Format (RTF) stream.
```
FUNCTION GetNewStorage (BYVAL lplpstg AS LPSTORAGE PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *lplpstg* | The address of the **IStorage** interface created for the new object. |

#### Return value

Returns **S_OK** on success. If the method fails, it can return one of the following values.

| Return code  | Description |
| ------------ | ----------- |
| **E_INVALIDARG** | There was an invalid argument. |
| **E_OUTOFMEMORY** | There was not enough memory to do the operation. |

#### Remarks

This method must be implemented to allow cut, copy, paste, drag, and drop operations of Component Object Model (COM) objects.

#### Implementation
```
FUNCTION CRichEditOleCallback.GetNewStorage (BYVAL lplpstg AS LPSTORAGE PTR) AS HRESULT
   DIM hr AS HRESULT
   DIM pILockBytes AS ILockBytes PTR
   hr = CreateILockBytesOnHGlobal(NULL, TRUE, @pILockBytes)
   IF FAILED(hr) THEN RETURN hr
   hr = StgCreateDocfileOnILockBytes(pILockBytes, _
        STGM_SHARE_EXCLUSIVE OR STGM_READWRITE OR STGM_CREATE, 0, lplpstg)
   RETURN hr
END FUNCTION
```

# <a name="GetInPlaceContext"></a>GetInPlaceContext

Provides the application and document-level interfaces and information required to support in-place activation.
```
GetInPlaceContext (BYVAL lplpFrame AS LPOLEINPLACEFRAME PTR, BYVAL lplpDoc AS LPOLEINPLACEUIWINDOW PTR, _
   BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *lplpFrame* | The address of the [IOleInPlaceFrame](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nn-oleidl-ioleinplaceframe) interface that represents the frame window of a rich edit control client. Use the **AddRef** method to increment the reference count. The rich edit control releases the interface when it is no longer needed. |
| *lplpDoc* | The address of the [IOleInPlaceUIWindow(https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nn-oleidl-ioleinplaceuiwindow) interface that represents the document window of the rich edit control client. An interface need not be returned if the frame and document windows are the same. Use the AddRef method to increment the reference count. The rich edit control releases the interface when it is no longer needed. |
| *lpFrameInfo* | The accelerator information. [OLEINPLACEFRAMEINFO](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/ns-oleidl-oleinplaceframeinfo) |

#### Return value

Returns **S_OK** on success. If the method fails, it can return the following value.
| Return code  | Description |
| ------------ | ----------- |
| **E_INVALIDARG** | There was an invalid argument. |
