# CGraphCtx Class

**CGraphCtx** is a graphic control for pictures, text and graphics. You can use both GDI and GDI+ to draw graphics and text and to load and manipulate images. Optionally, you can add support for OpenGL by passing "OPENGL" in the caption.

This control features persistence and uses a virtual buffer (set initially equal to the size of the control when it is created) to allow to display images bigger than the size of the control. Scrollbars are automatically added when the size of the virtual buffer is bigger than the size of the control and removed when unneeded. It also features keyboard navigation and sends command messages to the parent window or dialog when the return or Escape keys are pressed, and notification messages for mouse clicks.

To use the control, include the CGraphCtx.inc file in your program and use the namespace **Afx**.

**Include file**: CGraphCtx.inc

### Constructor

Registers the class name and creates an instance of the control.

```
CONSTRUCTOR CGraphCtx (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, _
   BYREF wszTitle AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
   BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, _
   BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWindow* | Pointer to the **CWindow** class of the parent window. |
| *cID* | Control identifier. |
| *wszTitle* | The window caption. If "OPENGL" is used as the caption, support for OpenGL is added to the control. |
| *x* | The x-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *y* | The initial y-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *nWidth* | The width of the control. |
| *nHeight* | The height of the control. |
| *dwStyle* | The style of the window being created. Pass -1 to use the default styles. |
| *dwExStyle* | The extended window style of the control being created. Pass -1 to use the default styles. |
| *lpParam* | Pointer to custom data. |

### Helper Procedure: AfxCGraphPtr

Returns a pointer to the CGraphCtx class given the handle of its associated window.

```
FUNCTION AfxCGraphPtr (BYVAL hwnd AS HWND) AS CGraphCtx PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle of the window associated with the graphic control. Call the **hWindow** method of the **CGraphCtx** class to retrieve it. |

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Clear](#Clear) | Clears the graphic control with the specified RGB color. |
| [CreateBitmapFromFile](#CreateBitmapFromFile) | Loads and displays the specified image in the Graphic Control. It also allows to convert the image to gray scale and/or dim the image. |
| [DrawBitmap](#DrawBitmap) | Draws a bitmap in the Graphic Control. |
| [GetBits](#GetBits) | Returns the location of the DIB bit values. |
| [GethBmp](#GethBmp) | Returns the handle of the compatible bitmap. |
| [GethRC](#GethRC) | If OpenGL is enabled, it returns the handle of the rendering context of the control. |
| [GetMemDC](#GetMemDC) | Handle to  the memory device context of the control. |
| [GetVirtualBufferHeight](#GetVirtualBufferHeight) | Returns the height of the virtual buffer. |
| [GetVirtualBufferWidth](#GetVirtualBufferWidth) | Returns the width of the virtual buffer. |
| [hWindow](#hWindow) | Returns the handle of the control. |
| [LoadImageFromFile](#LoadImageFromFile) | Loads and displays the specified image in the Graphic Control. |
| [LoadImageFromRes](#LoadImageFromRes) | Loads the specified image from a resource file in the Graphic Control. It also allows to convert the image to gray scale and/or dim the image. |
| [MakeCurrent](#MakeCurrent) | Redirects OpenGL calls are directed to the correct rendering context. |
| [PrintImage](#PrintImage) | Prints the image in the default printer. |
| [Resizable](#Resizable) | Gets/sets the value of the Resizable property. |
| [SaveImage](#SaveImage) | Saves the image to a file. |
| [SetVirtualBufferSize](#SetVirtualBufferSize) | Sets the size of the virtual buffer. |
| [Stretchable](#Stretchable) | Gets/sets the value of the Stretchable property. |
| [StretchMode](#StretchMode) | Gets/sets the value of the StretchMode property. |

### Notification Messages

| Name       | Description |
| ---------- | ----------- |
| [NM_CLICK](#NM_CLICK) | Sent by the control when the user clicks it with the left mouse button. |
| [NM_DBLCLK](#NM_DBLCLK) | Sent by the control when the user double clicks it with the left mouse button. |
| [NM_KILLFOCUS](#NM_KILLFOCUS) | Notifies a control's parent window that the control has lost the input focus. |
| [NM_RCLICK](#NM_RCLICK) | Sent by the control when the user clicks it with the right mouse button. |
| [NM_RDBLCLK](#NM_RDBLCLK) | Sent by the control when the user double clicks it with the right mouse button. |
| [NM_SETFOCUS](#NM_SETFOCUS) | Notifies a control's parent window that the control has received the input focus. |

# <a name="NM_CLICK"></a>NM_CLICK Notification Message

Sent by the control when the user clicks it with the left mouse button. This notification code is sent in the form of a WM_NOTIFY message.

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_CLICK
            ' Left button clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="NM_DBLCLK"></a>NM_DBLCLK Notification Message

Sent by the control when the user double clicks it with the left mouse button. This notification code is sent in the form of a WM_NOTIFY message.

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_DBLCLK
            ' Left button double clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="NM_KILLFOCUS"></a>NM_KILLFOCUS Notification Message

Notifies a control's parent window that the control has lost the input focus. This notification code is sent in the form of a WM_NOTIFY message. 

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_KILLFOCUS
            ' The control has lost focus
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="NM_RCLICK"></a>NM_RCLICK Notification Message

Notifies a control's parent window that the control has lost the input focus. This notification code is sent in the form of a WM_NOTIFY message. 

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_RCLICK
            ' Right button clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="NM_RDBLCLK"></a>NM_RDBLCLK Notification Message

Sent by the control when the user double clicks it with the right mouse button. This notification code is sent in the form of a WM_NOTIFY message.

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_RDBLCLK
            ' Right button double clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="NM_SETFOCUS"></a>NM_SETFOCUS Notification Message

Notifies a control's parent window that the control has received the input focus. This notification code is sent in the form of a WM_NOTIFY message. 

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_SETFOCUS
            ' The control has gained focus
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="Clear"></a>Clear

Clears the graphic control with the specified RGB color.

```
FUNCTION Clear (BYVAL RGBColor AS COLORREF) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RGBColor* | RGB color used to fill the control. |

#### Return value

TRUE or FALSE.

# <a name="CreateBitmapFromFile"></a>CreateBitmapFromFile

Loads and displays the specified image in the Graphic Control. It also allows to convert the image to gray scale and/or dim the image.

```
SUB CreateBitmapFromFile (BYREF wszFileName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Absolute path to the file. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | TRUE or FALSE. Convert to gray scale. |

#### Return value

TRUE or FALSE.

# <a name="DrawBitmap"></a>DrawBitmap

Draws a bitmap in the Graphic Control.

```
SUB DrawBitmap (BYVAL hbmp AS HBITMAP, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
   BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus
```
```
SUB DrawBitmap (BYVAL pBitmap AS GpBitmap PTR, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
   BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus
```
```
SUB DrawBitmap (BYREF pMemBmp AS CMemBmp, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
   BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hbmp* | A handle to the bitmap. |
| *pBitmap* | Pointer to a GDI+ bitmap. |
| *pMemBmp* | Reference to a memory bitmap (CMemBmp class). |
| *x* | Optional. Left position. |
| *y* | Optional. Top position. |
| *nRight* | Optional. Right position.. |
| *nBottom* | Optional. Bottom position.. |

#### Return value

Returns OK (0) on success or a GdiPlus status code on failure.

#### Remarks

If both *x* and *y* are ommited, the image is draw starting at position 0, 0.

If *nRight* and *nBottom* are specified, the image is draw stretched in the bounding rectangle formed by *x*, *y*, *nRight* and *nBottom*.

#### Sample code

```
DIM pMemBmp AS CMemBmp = CMemBmp($"C:\Users\Pepe\Pictures\Cole_Kyla_01.jpg")
Rectangle pMemBmp.GetMemDC, 10, 10, 150, 150
LineTo pMemBmp.GetMemDC, 30, 180
pGraphCtx.DrawBitmap pMemBmp
```

# <a name="GetBits"></a>GetBits

Returns the location of the DIB bit values.

```
FUNCTION GetBits () AS ANY PTR
```

#### Return value

Pointer to the location of the DIB bit values.

# <a name="GethBmp"></a>GethBmp

Returns the handle of the compatible bitmap.

```
FUNCTION GethBmp () AS HBITMAP
```

# <a name="GethRC"></a>GethRC

If OpenGL is enabled, it returns the handle of the rendering context of the control.

```
FUNCTION GethRC () AS HGLRC
```

# <a name="GetMemDC"></a>GetMemDC

Returns the handle of the memory device context of the control.

```
FUNCTION GetMemDC () AS HDC
```

# <a name="GetVirtualBufferHeight"></a>GetVirtualBufferHeight

Returns the height of the virtual buffer.

```
FUNCTION GetVirtualBufferHeight () AS LONG
```

# <a name="GetVirtualBufferWidth"></a>GetVirtualBufferWidth

Returns the width of the virtual buffer.

```
FUNCTION GetVirtualBufferWidth () AS LONG
```

# <a name="hWindow"></a>hWindow

Returns the handle of the control.

```
FUNCTION hWindow () AS HWND
```

# <a name="LoadImageFromFile"></a>LoadImageFromFile

Loads and displays the specified image in the Graphic Control.

```
SUB LoadImageFromFile (BYREF wszFileName AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Absolute path to the file. |

#### Remarks

A quirk in the GDI+ **GdipCreateBitmapFromFile** function causes that black and white images are loaded with increased contrast. Therefore, it's better to use the **CreateBitmapFromFile** for black and white images.

# <a name="LoadImageFromRes"></a>LoadImageFromRes

Loads the specified image from a resource file in the Graphic Control. It also allows to convert the image to gray scale and/or dim the image.

```
SUB LoadImageFromRes (BYVAL hInst AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL clrBackground AS ARGB = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInst* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszImageName* | Name of the image in the resource file (.RES). If the image resource uses an integral identifier, wszImage should begin with a number symbol (#) followed by the identifier in an ASCII format, e.g., "#998". Otherwise, use the text identifier name for the image. Only images embedded as raw data (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff. |
| *dimPercent* | Percent of dimming (1-99). |
| *bGrayScale* | TRUE or FALSE. Convert to gray scale. |
| *clrBackground* | The background color. This parameter is ignored if the bitmap is totally opaque. |

# <a name="MakeCurrent"></a>MakeCurrent

As more than one instance of this control can be used on a form, we need to make sure that OpenGL calls are directed to the correct rendering context. This is achieved by calling the **MakeCurrent** method.

```
FUNCTION MakeCurrent () AS BOOLEAN
```
#### Return value

TRUE or FALSE.

# <a name="PrintImage"></a>PrintImage

Prints the image in the default printer.

```
FUNCTION PrintImage (BYVAL bStretch AS BOOLEAN = FALSE, _
   BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *bStretch* | Optional. Stretch the image. |
| *nStretchMode* | Optional. Stretching mode.<br>Default value = InterpolationModeHighQualityBicubic. |

### InterpolationMode Enumeration

| Constant   | Description |
| ---------- | ----------- |
| **InterpolationModeDefault** | Specifies the default interpolation mode. |
| **InterpolationModeLowQuality** | Specifies a low-quality mode. |
| **InterpolationModeHighQuality** | Specifies a high-quality mode. |
| **InterpolationModeBilinear** | Specifies bilinear interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 50 percent of its original size. |
| **InterpolationModeBicubic** | Specifies bicubic interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 25 percent of its original size. |
| **InterpolationModeNearestNeighbor** | Specifies nearest-neighbor interpolation. |
| **InterpolationModeHighQualityBilinear** | Specifies high-quality, bilinear interpolation. Prefiltering is performed to ensure high-quality shrinking. |
| **InterpolationModeHighQualityBicubic** | Specifies high-quality, bicubic interpolation. Prefiltering is performed to ensure high-quality shrinking. This mode produces the highest quality transformed images. |

#### Return value

Returns TRUE if the bitmap has been printed successfully, or FALSE otherwise.

# <a name="Resizable"></a>Resizable

Gets/sets the value of the **Resizable** property.

```
PROPERTY Resizable () AS BOOLEAN
PROPERTY Resizable (BYVAL bResizable AS BOOLEAN)
```

#### Return value

TRUE or FALSE.

##### Remarks

Resizable and stretchable are mutually exclusive.

# <a name="SaveImage"></a>SaveImage

Saves the image to a file.

```
FUNCTION SaveImage (BYREF wszFileName AS WSTRING, BYREF wszMimeType AS WSTRING) AS LONG
```
```
FUNCTION SaveImageAsBmp (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveImageAsJpeg (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveImageAsPng (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveImageAsGif (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveImageAsTiff (BYREF wszFileName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Absolute path name of the file. |
| *wszMimeType* | The mime type.<br>"image/bmp" = Bitmap (.bmp)<br>"image/gif" = GIF (.gif)<br>"image/jpeg" = JPEG (.jpg)<br>"image/png" = PNG (.png)<br>"image/tiff" = TIFF (.tiff) |

#### Return value

If the method succeeds, it returns Ok, which is an element of the GDI+ Status enumeration.

If the method fails, it returns one of the other elements of the GDI+ Status enumeration.

# <a name="SetVirtualBufferSize"></a>SetVirtualBufferSize

Sets the size of the virtual buffer.

```
FUNCTION SetVirtualBufferSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | Width, in pixels, of the virtual buffer. |
| *nHeight* | Height, in pixels, of the virtual buffer. |

#### Return value

If the function succeeds, the return value is S_OK (0).

# <a name="Stretchable"></a>Stretchable

Gets/sets the value of the **Stretchable** property.

```
PROPERTY Stretchable () AS BOOLEAN
PROPERTY Stretchable (BYVAL bStretchable AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bStretchable* | TRUE or FALSE. |

#### Return value

TRUE or FALSE.

# <a name="StretchMode"></a>StretchMode

Gets/sets the value of the **StretchMode** property.

```
PROPERTY StretchMode () AS LONG
PROPERTY StretchMode (BYVAL nStretchMode AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStretchMode* | The stretching mode. |

| Mode       | Description |
| ---------- | ----------- |
| **BLACKONWHITE** | Performs a Boolean AND operation using the color values for the eliminated and existing pixels. If the bitmap is a monochrome bitmap, this mode preserves black pixels at the expense of white pixels. |
| **COLORONCOLOR** | Deletes the pixels. This mode deletes all eliminated lines of pixels without trying to preserve their information. |
| **HALFTONE** | Maps pixels from the source rectangle into blocks of pixels in the destination rectangle. The average color over the destination block of pixels approximates the color of the source pixels. After setting the HALFTONE stretching mode, an application must call the **SetBrushOrgEx** function to set the brush origin. If it fails to do so, brush misalignment occurs. |
| **STRETCH_ANDSCANS** | Same as BLACKONWHITE. |
| **STRETCH_DELETESCANS** | Same as COLORONCOLOR. |
| **STRETCH_HALFTONE** | Same as HALFTONE. |
| **STRETCH_ORSCANS** | Same as WHITEONBLACK. |
| **WHITEONBLACK** | Performs a Boolean OR operation using the color values for the eliminated and existing pixels. If the bitmap is a monochrome bitmap, this mode preserves white pixels at the expense of black pixels. |

#### Return value

The previous value of the property.
