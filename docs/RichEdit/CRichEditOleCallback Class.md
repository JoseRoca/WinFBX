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
| *lplpDoc* | The address of the [IOleInPlaceUIWindow(https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nn-oleidl-ioleinplaceuiwindow) interface that represents the document window of the rich edit control client. An interface need not be returned if the frame and document windows are the same. Use the **AddRef** method to increment the reference count. The rich edit control releases the interface when it is no longer needed. |
| *lpFrameInfo* | The accelerator information. [OLEINPLACEFRAMEINFO](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/ns-oleidl-oleinplaceframeinfo) |

#### Return value

Returns **S_OK** on success. If the method fails, it can return the following value.

| Return code  | Description |
| ------------ | ----------- |
| **E_INVALIDARG** | There was an invalid argument. |

# <a name="ShowContainerUI"></a>ShowContainerUI

Indicates whether or not the application is to display its container UI. The rich edit control looks ahead for double-clicks and defers the call if appropriate. Applications may defer hiding adornments until an [IOleInPlaceUIWindow.SetBorderSpace](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-ioleinplaceuiwindow-setborderspace) call is received.

```
FUNCTION ShowContainerUI (BYVAL fShow AS WINBOOL) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *fShow* | Show container UI flag. The value is **TRUE** if the container UI is displayed, and **FALSE** if it is not. |

#### Return value

Returns **S_OK** on success. If the method fails, it can return the following value.

| Return code  | Description |
| ------------ | ----------- |
| **E_INVALIDARG** | There was an invalid argument. |

#### Remarks
The **ShowContainerUI** method is called by the [IOleInPlaceSite.OnUIActivate](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-ioleinplacesite-onuiactivate) and [IOleInPlaceSite.OnUIDeactivate](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-ioleinplacesite-onuideactivate) methods of the [IOleInPlaceSite](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nn-oleidl-ioleinplacesite) interface.

# <a name="QueryInsertObject"></a>QueryInsertObject

Queries the application as to whether an object should be inserted. The member is called when pasting and when reading Rich Text Format (RTF).
```
FUNCTION QueryInsertObject (BYVAL lpclsid AS LPCLSID, BYVAL lpstg AS LPSTORAGE, BYVAL cp AS LONG) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *lpclsid* | Class identifier of the object to be inserted. |
| *lpstg* | Storage containing the object. |
| *cp* | Character position, at which the object will be inserted. |

#### Return value

Returns **S_OK** on success. If the return value is not S_OK, the object was not inserted. If the method fails, it can return the following value.

| Return code  | Description |
| ------------ | ----------- |
| **E_INVALIDARG** | There was an invalid argument. |

# <a name="DeleteObject"></a>DeleteObject

Sends notification that an object is about to be deleted from a rich edit control. The object is not necessarily being released when this member is called.

```
FUNCTION DeleteObject (BYVAL lpoleobj AS LPOLEOBJECT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *lpoleobj* | The object that is being deleted. |

#### Return value

Returns **S_OK**.

# <a name="QueryAcceptData"></a>QueryAcceptData

During a paste operation or a drag event, determines if the data that is pasted or dragged should be accepted.

**QueryAcceptData** is called when the user attempts to drag or paste an object into the control. If your application can accept the data, your implementation of the method returns **S_OK**; otherwise, it returns an error code. The method might also take some other action, such as warning the user that files of that type cannot be placed in the control.

```
FUNCTION QueryAcceptData (BYVAL lpdataobj AS LPDATAOBJECT, BYVAL lpcfFormat AS CLIPFORMAT PTR, _
   BYVAL reco AS DWORD, BYVAL fReally AS WINBOOL, BYVAL hMetaPict AS HGLOBAL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *lpdataobj* | The data object being pasted or dragged. |
| *lpcfFormat* | The clipboard format that will be used for the paste or drop operation. If the value pointed to by *lpcfFormat* is zero, the best available format will be used. If the callback changes the value pointed to by *lpcfFormat*, the rich edit control only uses that format and the operation will fail if the format is not available. |
| *reco* | A clipboard operation flag, which can be one of these values:<br>**RECO_DROP**. Drop operation (drag-and-drop).<br>**RECO_PASTE**. Paste from the clipboard. |
| *fReally* | Indicates whether the drag-drop is actually happening or if it is just a query. A nonzero value indicates the paste or drop is actually happening. A zero value indicates the operation is just a query, such as for **EM_CANPASTE**. |
| *hMetaPict* | Handle to a metafile containing the icon view of an object if **DVASPECT_ICON** is being imposed on an object by a paste special operation. |

#### Return value

Returns **S_OK** on success. See **Remarks**.

#### Remarks

On failure, the rich edit control refuses the data and terminates the operation. Otherwise, the control checks the data itself for acceptable formats. A success code other than **S_OK** means that the callback either checked the data itself (if *fReally* is **FALSE**) or imported the data itself (if *fReally* is **TRUE**). If the application returns a success code other than **S_OK**, the control does not check the read-only state of the edit control.

# <a name="ContextSensitiveHelp"></a>ContextSensitiveHelp

Indicates if the application should transition into or out of context-sensitive help mode. This method should implement the functionality described for [IOleWindow.ContextSensitiveHelp](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-iolewindow-contextsensitivehelp).

```
FUNCTION ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *fEnterMode* | If **TRUE**, the application should enter context-sensitive help mode. If **FALSE**, the application should leave context-sensitive help mode. |

#### Return value

Returns **S_OK** on success. If the method fails, it can return the following value.

| Return code  | Description |
| ------------ | ----------- |
| **E_INVALIDARG** | There was an invalid argument. |

# <a name="GetClipboardData"></a>GetClipboardData

Allows the client to supply its own clipboard object.

```
FUNCTION GetClipboardData (BYVAL lpchrg AS CHARRANGE PTR, BYVAL reco AS DWORD, _
   BYVAL lplpdataobj AS LPDATAOBJECT PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *lpchrg* | The clipboard object range. |
| *reco* | The clipboard operation flag can be one of these values.<br>**ECO_COPY**. Copy to the clipboard.<br>**RECO_CUT**. Cut to the clipboard. |
| *lplpdataobj* | Pointer to the pointer variable that receives the address of the [IDataObject](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nn-objidl-idataobject) implementation representing the range specified in the *lpchrg* parameter. The value of **lplpdataobj** is ignored if an error is returned. |

#### Return value

Returns **S_OK** on success. If the return value is **E_NOTIMPL**, the rich edit control created its own clipboard object. If the return value is a failure other than **E_NOTIMPL**, the operation failed.

# <a name="GetDragDropEffect"></a>GetDragDropEffect

Allows the client to specify the effects of a drop operation.

```
FUNCTION GetDragDropEffect (BYVAL fDrag AS WINBOOL, BYVAL grfKeyState AS DWORD, _
   BYVAL pdwEffect AS LPDWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *fDrag* | **TRUE** if the query is for a [IDropTarget.DragEnter](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-idroptarget-dragenter) or [IDropTarget.DragOver](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-idroptarget-dragover). **FALSE** if the query is for [IDropTarget.Drop](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-idroptarget-drop). |
| *grfKeyState* | Key state as defined by OLE. |
| *pdwEffect* | The effect used by a rich edit control. When *fDrag* is **TRUE**, on return, its content is set to the effect allowable by the rich edit control. When *fDrag* is **FALSE**, on return, the variable is set to the effect to use. |

#### Return value

This method returns **S_OK**.

# <a name="GetContextMenu"></a>GetContextMenu

Queries the application for a context menu to use on a right-click event.

```
FUNCTION GetContextMenu (BYVAL seltype AS WORD, BYVAL lpoleobj AS LPOLEOBJECT, _
   BYVAL lpchrg AS CHARRANGE PTR, BYVAL lphmenu AS HMENU PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *seltype* | Selection type. The value, which specifies the contents of the new selection, can be one or more of the following values.<br>**SEL_EMPTY**. The selection is empty.<br>**SEL_TEXT**. Text.<br>**SEL_OBJECT**. At least one COM object.<br>**SEL_MULTICHAR**. More than one character of text.<br>**SEL_MULTIOBJECT**. More than one COM object.<br>**GCM_RIGHTMOUSEDROP**. Indicates that a context menu for a right-mouse drag drop should be generated. The *lpoleobj* parameter is a pointer to the [IDataObject](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nn-objidl-idataobject) interface for the object being dropped. |
| *lpoleobj* | Pointer to an interface. If the *seltype* parameter includes the **SEL_OBJECT** flag, *lpoleobj* is a pointer to the [IOleObject](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nn-oleidl-ioleobject) interface for the first selected COM object. If seltype includes the **GCM_RIGHTMOUSEDROP** flag, *lpoleobj* is a pointer to an [IDataObject](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nn-objidl-idataobject) interface. Otherwise, *lpoleobj* is NULL. If you hold on to the interface pointer, you must call the **AddRef** method to increment the object's reference count. |
| *lpchrg* | Pointer to a [CHARRANGE](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charrange) structure containing the current selection. |
| *lphmenu* | The handle of a context menu to use. This parameter is ignored if an error is returned. A rich edit control destroys the menu when it is finished with it so the client should not. |

Returns **S_OK** on success. If the method fails, it can return the following value.

| Return code  | Description |
| ------------ | ----------- |
| **E_INVALIDARG** | There was an invalid argument. |
