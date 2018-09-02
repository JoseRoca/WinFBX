# CGpGraphics Class

The **CGpGraphics** class provides methods for drawing lines, curves, figures, images, and text. A **Graphics** object stores attributes of the display device and attributes of the items to be drawn.

**Inherits from**: CGpBase.<br>
**Imclude file**: CGpGraphics.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsGraphics) | Creates a **Graphics** object. |
| [AddMetafileComment](#AddMetafileComment) | Adds a text comment to an existing metafile. |
| [BeginContainer](#BeginContainer) | Begins a new graphics container. |
| [Clear](#Clear) | Clears a **Graphics** object to a specified color. |
| [DrawArc](#DrawArc) | Draws an arc. The arc is part of an ellipse. |
| [DrawBezier](#DrawBezier) | Draws a Bézier spline. |
| [DrawBeziers](#DrawBeziers) | Draws a sequence of connected Bézier splines. |
| [DrawCachedBitmap](#DrawCachedBitmap) | Draws the image stored in a **CachedBitmap** object. |
| [DrawClosedCurve](#DrawClosedCurve) | Draws a closed cardinal spline. |
| [DrawCurve](#DrawCurve) | Draws a cardinal spline. |
| [DrawDriverString](#DrawDriverString) | Draws characters at the specified positions. |
| [DrawEllipse](#DrawEllipse) | Draws an ellipse. |
| [DrawImage](#DrawImage) | Draws an image. |
| [DrawImageFX](#DrawImageFX) | Draws a portion of an image after applying a specified effect. |
| [DrawLine](#DrawLine) | Draws a line that connects two points. |
| [DrawLines](#DrawLines) | Draws a sequence of connected lines. |
| [DrawPath](#DrawPath) | Draws a sequence of lines and curves defined by a **GraphicsPath** object. |
| [DrawPie](#DrawPie) | Draws a pie. |
| [DrawPolygon](#DrawPolygon) | Draws a polygon. |
| [DrawRectangle](#DrawRectangle) | Draws a rectangle. |
| [DrawRectangles](#DrawRectangles) | Draws a sequence of rectangles. |
| [DrawString](#DrawString) | Draws a string based on a font and an origin for the string. |
| [EndContainer](#EndContainer) | Closes a graphics container that was previously opened by the **BeginContainer** method. |
| [EnumerateMetafile](#EnumerateMetafile) | Calls an application-defined callback function for each record in a specified metafile. |
| [ExcludeClip](#ExcludeClip) | Updates the clipping region to the portion of itself that does not intersect the specified rectangle. |
| [FillClosedCurve](#FillClosedCurve) | Creates a closed cardinal spline from an array of points and uses a brush to fill the interior of the spline. |
| [FillEllipse](#FillEllipse) | Uses a brush to fill the interior of an ellipse that is specified by coordinates and dimensions. |
| [FillPath](#FillPath) | Uses a brush to fill the interior of a path. |
| [FillPie](#FillPie) | Uses a brush to fill the interior of a pie. |
| [FillPolygon](#FillPolygon) | Uses a brush to fill the interior of a polygon. |
| [FillRectangle](#FillRectangle) | Uses a brush to fill the interior of a rectangle. |
| [FillRectangles](#FillRectangles) | Uses a brush to fill the interior of a sequence of rectangles. |
| [FillRegion](#FillRegion) | Uses a brush to fill a specified region. |
| [Flush](#Flush) | Flushes all pending graphics operations. |
| [FromHDC](#FromHDC) | Creates a **Graphics** object that is associated with a specified device context. |
| [FromHWND](#FromHWND) | Creates a **Graphics** object that is associated with a specified window. |
| [FromImage](#FromImage) | Creates a **Graphics** object that is associated with a specified device context. |
| [GetClip](#GetClip) | Gets the clipping region of this **Graphics** object. |
| [GetClipBounds](#GetClipBounds) | Gets a rectangle that encloses the clipping region of this **Graphics** object. |
| [GetCompositingMode](#GetCompositingMode) | Gets the compositing mode currently set for this **Graphics** object. |
| [GetCompositingQuality](#GetCompositingQuality) | Gets the compositing quality currently set for this **Graphics** object. |
| [GetDpiX](#GetDpiX) | Gets the horizontal resolution, in dots per inch, of the display device associated with this **Graphics** object. |
| [GetDpiY](#GetDpiY) | Gets the vertical resolution, in dots per inch, of the display device associated with this **Graphics** object. |
| [GetHalftonePalette](#GetHalftonePalette) | Gets a Windows halftone palette. |
| [GetHDC](#GetHDC) | Gets a handle to the device context associated with this **Graphics** object. |
| [GetInterpolationMode](#GetInterpolationMode) | Gets the interpolation mode currently set for this **Graphics** object. |
| [GetNearestColor](#GetNearestColor) | Gets the nearest color to the color that is passed in. |
| [GetPageScale](#GetPageScale) | Gets the scaling factor currently set for the page transformation of this **Graphics** object. |
| [GetPageUnit](#GetPageUnit) | Gets the unit of measure currently set for this **Graphics** object. |
| [GetPixelOffsetMode](#GetPixelOffsetMode) | Gets the pixel offset mode currently set for this **Graphics** object. |
| [GetRenderingOrigin](#GetRenderingOrigin) | Gets the rendering origin currently set for this **Graphics** object. |
| [GetSmoothingMode](#GetSmoothingMode) | Determines whether smoothing (antialiasing) is applied to the **Graphics** object. |
| [GetTextContrast](#GetTextContrast) | Gets the contrast value currently set for this **Graphics** object. |
| [GetTextRenderingHint](#GetTextRenderingHint) | Returns the text rendering mode currently set for this **Graphics** object. |
| [GetTransform](#GetTransform) | Gets the world transformation matrix of this **Graphics** object. |
| [GetVisibleClipBounds](#GetVisibleClipBounds) | Gets a rectangle that encloses the visible clipping region of this **Graphics** object. |
| [IntersectClip](#IntersectClip) | Updates the clipping region of this **Graphics** object to the portion of the specified rectangle that intersects with the current clipping region of this **Graphics** object. |
| [IsClipEmpty](#IsClipEmpty) | Determines whether the clipping region of this **Graphics** object is empty. |
| [IsVisible](#IsVisible) | Determines whether the specified point is inside the visible clipping region of this **Graphics** object. |
| [IsVisibleClipEmpty](#IsVisibleClipEmpty) | Determines whether the visible clipping region of this **Graphics** object is empty. |
| [MeasureCharacterRanges](#MeasureCharacterRanges) | Gets a set of regions each of which bounds a range of character positions within a string. |
| [MeasureDriverString](#MeasureDriverString) | Measures the bounding box for the specified characters and their corresponding positions. |
| [MeasureString](#MeasureString) | Measures the extent of the string in the specified font, format, and layout rectangle. |
| [MultiplyTransform](#MultiplyTransform) | Updates this **Graphics** object's world transformation matrix with the product of itself and another matrix. |
| [ReleaseHDC](#ReleaseHDC) | Releases a device context handle obtained by a previous call to the **GetHDC** method of this **Graphics** object. |
| [ResetClip](#ResetClip) | Sets the clipping region of this **Graphics** object to an infinite region. |
| [ResetTransform](#ResetTransform) | Sets the world transformation matrix of this Graphics object to the identity matrix. |
| [Restore](#Restore) | Sets the state of this **Graphics** object to the state stored by a previous call to the **Save** method of this Graphics object. |
| [RotateTransform](#RotateTransform) | Updates the world transformation matrix of this **Graphics** object with the product of itself and a rotation matrix. |
| [Save](#Save) | Saves the current state (transformations, clipping region, and quality settings) of this **Graphics** object. |
| [ScaleTransform](#ScaleTransform) | Updates this **Graphics** object's world transformation matrix with the product of itself and a scaling matrix. |
| [SetClip](#SetClip) | Updates the clipping region of this **Graphics** object to a region that is the combination of itself and the clipping region of another **Graphics** object. |
| [SetCompositingMode](#SetCompositingMode) | Sets the compositing mode of this **Graphics** object. |
| [SetCompositingQuality](#SetCompositingQuality) | Sets the compositing quality of this **Graphics** object. |
| [SetInterpolationMode](#SetInterpolationMode) | Sets the interpolation mode of this **Graphics** object. |
| [SetPageScale](#SetPageScale) | Sets the scaling factor for the page transformation of this **Graphics** object. |
| [SetPageUnit](#SetPageUnit) | Sets the unit of measure for this **Graphics** object. |
| [SetPixelOffsetMode](#SetPixelOffsetMode) | Sets the pixel offset mode of this **Graphics** object. |
| [SetRenderingOrigin](#SetRenderingOrigin) | Sets the rendering origin of this **Graphics** object. |
| [SetSmoothingMode](#SetSmoothingMode) | Sets the rendering quality of the **Graphics** object. |
| [SetTextContrast](#SetTextContrast) | Sets the contrast value of this **Graphics** object. |
| [SetTextRenderingHint](#SetTextRenderingHint) | Sets the text rendering mode of this **Graphics** object. |
| [SetTransform](#SetTransform) | Sets the world transformation of this **Graphics** object. |
| [TransformPoints](#TransformPoints) | Converts an array of points from one coordinate space to another. |
| [TranslateClip](#TranslateClip) | Translates the clipping region of this **Graphics** object. |
| [TranslateTransform](#TranslateTransform) | Updates this **Graphics** object's world transformation matrix with the product of itself and a translation matrix. |

# CGpGraphicsPath Class

The **CGpGraphicsPath** allows the creation of **GraphicPath** objects. A **GraphicsPath** object stores a sequence of lines, curves, and shapes. You can draw the entire sequence by calling the **DrawPath** method of a **Graphics** object. You can partition the sequence of lines, curves, and shapes into figures, and with the help of a **GraphicsPathIterator** object, you can draw selected figures. You can also place markers in the sequence, so that you can draw selected portions of the path.

**Inherits from**: CGpBase.<br>
**Imclude file**: CGpPath.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsGraphicsPath) | Creates a **Graphics** object. |
| [AddArc](#AddArc) | Adds an elliptical arc to the current figure of this path. |
| [AddBezier](#AddBezier) | Adds a Bézier spline to the current figure of this path. |
| [AddBeziers](#AddBeziers) | Adds a sequence of connected Bézier splines to the current figure of this path. |
| [AddClosedCurve](#AddClosedCurve) | Adds a closed cardinal spline to this path. |
| [AddCurve](#AddCurve) | Adds a cardinal spline to this path. |
| [AddEllipse](#AddEllipse) | Adds an ellipse to this path. |
| [AddLine](#AddLine) | Adds a line to the current figure of this path. |
| [AddLines](#AddLines) | Adds a sequence of connected lines to the current figure of this path. |
| [AddPath](#AddPath) | Adds a path to this path. |
| [AddPie](#AddPie) | Adds a pie to this path. An arc is a portion of an ellipse, and a pie is a portion of the area enclosed by an ellipse. |
| [AddPolygon](#AddPolygon) | Adds a polygon to this path. |
| [AddRectangle](#AddRectangle) | Adds a rectangle to this path. |
| [AddRectangles](#AddRectangles) | Adds a sequence of rectangles to this path. |
| [AddString](#AddString) | Adds the outline of a string to this path. |
| [ClearMarkers](#ClearMarkers) | Clears the markers from this path. |
| [Clone](#Clone) | Copies the contents of the existing **GraphicsPath** object into a new **GraphicsPath** object. |
| [CloseAllFigures](#CloseAllFigures) | Clears the markers from this path. |
| [CloseFigure](#CloseFigure) | Clears the current figure of this path. |
| [Flatten](#Flatten) | Applies a transformation to this path and converts each curve in the path to a sequence of connected lines. |
| [GetBounds](#GetBounds) | Gets a bounding rectangle for this path. |
| [GetFillMode](#GetFillMode) | Gets the fill mode of this path. |
| [GetLastPoint](#GetLastPoint) | Gets the ending point of the last figure in this path. |
| [GetPathData](#GetPathData) | Gets an array of points and an array of point types from this path. |
| [GetPathPoints](#GetPathPoints) | Gets this path's array of points. |
| [GetPathTypes](#GetPathTypes) | Gets this path's array of point types. |
| [GetPointCount](#GetPointCount) | Gets the number of points in this path's array of data points. |
| [IsOutlineVisible](#IsOutlineVisible) | Determines whether a specified point touches the outline of this path when the path is drawn by a specified Graphics object and a specified pen. |
| [IsVisible](#IsVisible) | Determines whether a specified point lies in the area that is filled when this path is filled by a specified Graphics object. |
| [Outline](#Outline) | Transforms and flattens this path, and then converts this path's data points so that they represent only the outline of the path. |
| [Reset](#Reset) | Empties the path and sets the fill mode to **FillModeAlternate**. |
| [Reverse](#Reverse) | Reverses the order of the points that define this path's lines and curves. |
| [SetFillMode](#SetFillMode) | Sets the fill mode of this path. |
| [SetMarker](#SetMarker) | Designates the last point in this path as a marker point. |
| [StartFigure](#StartFigure) | Starts a new figure without closing the current figure. |
| [Transform](#Transform) | Multiplies each of this path's data points by a specified matrix. |
| [Warp](#Warp) | Applies a warp transformation to this path. |
| [Widen](#Widen) | Replaces this path with curves that enclose the area that is filled when this path is drawn by a specified pen. |

# CGpGraphicsPathIterator Class

The **CGpGraphicsPathIterator** class provides methods for isolating selected subsets of the path stored in a **GraphicsPath** object. A path consists of one or more figures. You can use a **GraphicsPathIterator** to isolate one or more of those figures. A path can also have markers that divide the path into sections. You can use a **GraphicsPathIterator** object to isolate one or more of those sections.

**Inherits from**: CGpBase.<br>
**Imclude file**: CGpPath.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsGraphicsPathIterator) | Creates a new **GraphicsPathIterator** object and associates it with a GraphicsPath object. |
| [CopyData](#CopyData) | Copies a subset of the path's data points to a PointF array and copies a subset of the path's point types to a BYTE array. |
| [Enumerate](#Enumerate) | Copies the path's data points to a PointF array and copies the path's point types to a BYTE array. |
| [GetCount](#GetCount) | Returns the number of data points in the path. |
| [GetSubpathCount](#GetSubpathCount) | Returns the number of subpaths (also called figures) in the path. |
| [HasCurve](#HasCurve) | Determines whether the path has any curves. |
| [NextMarker](#NextMarker) | Gets the starting index and the ending index of the next marker-delimited section in this iterator's associated path. |
| [NextPathType](#NextPathType) | Gets the starting index and the ending index of the next group of data points that all have the same type. |
| [NextSubpath](#NextSubpath) | Gets the starting index and the ending index of the next subpath (figure) in this iterator's associated path. |
| [Rewind](#Rewind) | Rewinds this iterator to the beginning of its associated path. |

# <a name="ConstructorsGraphics"></a>Constructors (CGpGraphics)

Creates a **Graphics** object that is associated with a specified device context. When you use this method to create a **Graphics** object, make sure that the **Graphics** object is deleted before the device context is released.

```
CONSTRUCTOR CGpGraphics (BYVAL hdc AS HDC)
```
Creates a **Graphics** object that is associated with a specified device context and a specified device. When you use this constructor to create a **Graphics** object, make sure that the Graphics object is deleted or goes out of scope before the device context is released.

```
CONSTRUCTOR CGpGraphics (BYVAL hdc AS HDC, BYVAL hDevice AS HANDLE)
```

Creates a **Graphics** object that is associated with a specified window.

```
CONSTRUCTOR CGpGraphics (BYVAL pImage AS CGpImage PTR)
```

Creates a **Graphics** object that is associated with an **Image** object. This constructor fails if the **Image** object is based on a metafile that was opened for reading. The Image(file) and Metafile(file) constructors open a metafile for reading. To open a metafile for recording, use a **Metafile** constructor that receives a device context handle.

This constructor also fails if the image uses one of the following pixel formats:

PixelFormatUndefined<br>
PixelFormatDontCare<br>
PixelFormat1bppIndexed<br>
PixelFormat4bppIndexed<br>
PixelFormat8bppIndexed<br>
PixelFormat16bppGrayScale<br>
PixelFormat16bppARGB1555

```
CONSTRUCTOR CGpGraphics (BYVAL pImage AS CGpImage PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | Handle to the device context that will be associated with the new **Graphics** object. |
| *hDevice* | Handle to a device that will be associated with the new **Graphics** object. |
| *hwnd* | Handle to a window that will be associated with the new **Graphics** object. |
| *icm* | Optional. Boolean value that specifies whether the new **Graphics** object applies color adjustment according to the ICC profile associated with the display device. TRUE specifies that color adjustment is applied, and FALSE specifies that color adjustment is not applied. The default value is FALSE. |
| *pImage* | Pointer to an **Image** object that will be associated with the new **Graphics** object. |


# <a name="AddMetafileComment"></a>AddMetafileComment (CGpGraphics)

Creates a **Graphics** object that is associated with a specified device context. When you use this method to create a **Graphics** object, make sure that the **Graphics** object is deleted before the device context is released.

```
FUNCTION AddMetafileComment (BYVAL pdata AS BYTE PTR, BYVAL sizeData AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rgData* | Pointer to a buffer that contains the comment. |
| *sizeData* | Integer that specifies the number of bytes in the value of the *data* parameter.  |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="BeginContainer"></a>BeginContainer (CGpGraphics)

Begins a new graphics container.

```
FUNCTION BeginContainer () AS GraphicsContainer
FUNCTION BeginContainer (BYVAL destrect AS GpRectF PTR, BYVAL srcrect AS GpRectF PTR, _
   BYVAL nUnit AS GpUnit) AS GraphicsContainer
FUNCTION BeginContainer (BYVAL destrect AS GpRect PTR, BYVAL srcrect AS GpRect PTR, _
   BYVAL nUnit AS GpUnit) AS GraphicsContainer
```

| Parameter  | Description |
| ---------- | ----------- |
| *destrect* | Reference to a rectangle that, together with *srcrect*, specifies a transformation for the container. |
| *srcrect* | Reference to a rectangle that, together with *dstrect*, specifies a transformation for the container. |

#### Return value

This method returns a value that identifies the container.

#### Remarks

Use this method to create nested graphics containers. Graphics containers are used to retain graphics state, such as transformations, clipping regions, and various rendering properties.

The **BeginContainer** method returns a value of type **GraphicsContainer**. When you have finished using a container, pass that value to the **EndContainer** method. The **GraphicsContainer** data type is defined in Gdiplusenums.inc.

When you call the **BeginContainer** method of a **Graphics** object, an information block that holds the state of the **Graphics** object is put on a stack. The **BeginContainer** method returns a value that identifies that information block. When you pass the identifying value to the EndContainer method, the information block is removed from the stack and is used to restore the **Graphics** object to the state it was in at the time of the **BeginContainer** call.

Containers can be nested; that is, you can call the **BeginContainer** method several times before you call the **EndContainer** method. Each time you call the **BeginContainer** method, an information block is put on the stack, and you receive an identifier for the information block. When you pass one of those identifiers to the **EndContainer** method, the **Graphics** object is returned to the state it was in at the time of the **BeginContainer** call that returned that particular identifier. The information block placed on the stack by that **BeginContainer** call is removed from the stack, and all information blocks placed on that stack after that **BeginContainer** call are also removed.

Calls to the **Save** method place information blocks on the same stack as calls to the **BeginContainer** method. Just as an **EndContainer** call is paired with a **BeginContainer** call, a **Restore** call is paired with a **Save** call.

**Caution**: When you call **EndContainer**, all information blocks placed on the stack (by **Save** or by **BeginContainer**) after the corresponding call to **BeginContainer** are removed from the stack. Likewise, when you call **Restore**, all information blocks placed on the stack (by Save or by **BeginContainer**) after the corresponding call to Save are removed from the stack.

For more information about graphics containers, see [Nested Graphics Containers](https://docs.microsoft.com/en-us/windows/desktop/gdiplus/-gdiplus-nested-graphics-containers-use)


# <a name="Clear"></a>Clear (CGpGraphics)

Clears a **Graphics** object to a specified color.

```
FUNCTION Clear (BYVAL colour AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *Colour* | The color to paint the background. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="DrawArc"></a>DrawArc (CGpGraphics)

Draws an arc. The arc is part of an ellipse.

```
FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, _
   BYVAL sweepAngle AS SINGLE) AS LONG
FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, _
   BYVAL sweepAngle AS SINGLE) AS LONG
FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS LONG
FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the arc. |
| *x* | The x-coordinate of the upper-left corner of the bounding rectangle for the ellipse that contains the arc. |
| *y* | The y-coordinate of the upper-left corner of the bounding rectangle for the ellipse that contains the arc. |
| *nWidth* | The width of the ellipse that contains the arc. |
| *nHeight* | The height of the ellipse that contains the arc. |
| *startAngle* | The angle between the x-axis and the starting point of the arc.  |
| *sweepAngle* | The angle between the starting and ending points of the arc. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a 90-degree arc.
' ========================================================================================
SUB Example_DrawArc (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the arc
   DIM redPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 3)
   graphics.DrawArc(@redPen, 0, 0, 200, 100, 0.0, 90.0)

END SUB
' ========================================================================================
```

# <a name="DrawBezier"></a>DrawBezier (CGpGraphics)

Draws a Bézier spline.

```
FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, _
   BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE, BYVAL x3 AS SINGLE, BYVAL y3 AS SINGLE, _
   BYVAL x4 AS SINGLE, BYVAL y4 AS SINGLE) AS GpStatus
FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, _
   BYVAL x2 AS INT_, BYVAL y2 AS INT_, BYVAL x3 AS INT_, BYVAL y3 AS INT_, _
   BYVAL x4 AS INT_, BYVAL y4 AS INT_) AS GpStatus
FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPointF) AS GpStatus
FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPoint) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the Bézier spline. |
| *x1* | The x-coordinate of the starting point of the Bézier spline. |
| *y1* | The y-coordinate of the starting point of the Bézier spline. |
| *x2* | The x-coordinate of the first control point of the Bézier spline. |
| *y2* | The y-coordinate of the first control point of the Bézier spline. |
| *x3* | The x-coordinate of the second control point of the Bézier spline. |
| *y3* | The y-coordinate of the second control point of the Bézier spline. |
| *x4* | The x-coordinate of the ending point of the Bézier spline. |
| *y4* | The y-coordinate of the ending point of the Bézier spline. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a Bézier curve.
' ========================================================================================
SUB Example_DrawBezier (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the curve.
   DIM greenPen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   graphics.DrawBezier(@greenPen, 100.0, 100.0, 200.0, 10.0, 350.0, 50.0, 500.0, 100.0)

   ' // Draw the end points and control points.
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM blueBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   graphics.FillEllipse(@redBrush, 100 - 5, 100 - 5, 10, 10)
   graphics.FillEllipse(@redBrush, 500 - 5, 100 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 200 - 5, 10 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 350 - 5, 50 - 5, 10, 10)

END SUB
' ========================================================================================
```


# <a name="DrawBeziers"></a>DrawBeziers (CGpGraphics)

Draws a sequence of connected Bézier splines.

```
FUNCTION DrawBeziers (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the Bézier spline. |
| *pts* | Pointer to an array of GpPointF objects that specify the starting, ending, and control points of the Bézier splines. |
| *count* | Integer that specifies the number of elements in the points array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A Bézier spline does not pass through its control points. The control points act as magnets, pulling the curve in certain directions to influence the way a Bézier spline bends. Each Bézier spline requires a starting point and an ending point. Each ending point is the starting point for the next Bézier spline.

#### Example

```
' ========================================================================================
' The following example draws a pair of Bézier curves.
' ========================================================================================
SUB Example_DrawBeziers (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Define a Pen object and an array of PointF objects.
   DIM greenPen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   DIM startPoint AS GpPointF : startPoint.x = 100.0 : startPoint.y = 100.0
   DIM ctrlPoint1 AS GpPointF : ctrlPoint1.x = 200.0 : ctrlPoint1.y = 50.0
   DIM ctrlPoint2 AS GpPointF : ctrlPoint2.x = 400.0 : ctrlPoint2.y = 10.0
   DIM endPoint1  AS GpPointF : endPoint1.x  = 500.0 : endPoint1.y  = 100.0
   DIM ctrlPoint3 AS GpPointF : ctrlPoint3.x = 600.0 : ctrlPoint3.y = 200.0
   DIM ctrlPoint4 AS GpPointF : ctrlPoint4.x = 700.0 : ctrlPoint4.y = 400.0
   DIM endPoint2  AS GpPointF : endPoint2.x  = 500.0 : endPoint2.y  = 500.0

   DIM curvePoints(6) AS GpPointF
   curvePoints(0) = startPoint
   curvePoints(1) = ctrlPoint1
   curvePoints(2) = ctrlPoint2
   curvePoints(3) = endPoint1
   curvePoints(4) = ctrlPoint3
   curvePoints(5) = ctrlPoint4
   curvePoints(6) = endPoint2

   ' // Draw the Bezier curves.
   graphics.DrawBeziers(@greenPen, @curvePoints(0), 7)

   ' // Draw the control and end points.
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   graphics.FillEllipse(@redBrush, 100 - 5, 100 - 5, 10, 10)
   graphics.FillEllipse(@redBrush, 500 - 5, 100 - 5, 10, 10)
   graphics.FillEllipse(@redBrush, 500 - 5, 500 - 5, 10, 10)
   DIM blueBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   graphics.FillEllipse(@blueBrush, 200 - 5, 50 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 400 - 5, 10 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 600 - 5, 200 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 700 - 5, 400 - 5, 10, 10)

END SUB
' ========================================================================================
```

# <a name="DrawCachedBitmap"></a>DrawCachedBitmap (CGpGraphics)

Draws the image stored in a **CachedBitmap** object.

```
FUNCTION DrawCachedBitmap (BYVAL pCachedBitmap AS CGpCachedBitmap PTR, _
   BYVAL x AS LONG, BYVAL y AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCachedBitmap* | Pointer to a **CachedBitmap** object that contains the image to be drawn. |
| *x* | The x-coordinate of the upper-left corner of the image. |
| *y* | The y-coordinate of the upper-left corner of the image. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **CachedBitmap** object stores an image in a format that is optimized for a particular display screen. You cannot draw a cached bitmap to a printer or to a metafile.

Cached bitmaps will not work with any transformations other than translation.

When you construct a **CachedBitmap** object, you must pass the address of a **Graphics** object to the constructor. If the screen associated with that Graphics object has its bit depth changed after the cached bitmap is constructed, then the **DrawCachedBitmap** method will fail, and you should reconstruct the cached bitmap. Alternatively, you can hook the display change notification message and reconstruct the cached bitmap at that time.


# <a name="DrawClosedCurve"></a>DrawClosedCurve (CGpGraphics)

Draws a closed cardinal spline.

```
FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_, _
   BYVAL tension AS SINGLE) AS GpStatus
FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_, _
   BYVAL tension AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the closed cardinal spline. |
| *pts* | Pointer to an array of **GpPointF** objects that specify the coordinates of the closed cardinal spline. The array of **GpPointF** objects must contain a minimum of three elements. |
| *count* | Integer that specifies the number of elements in the points array. |
| *tension* | Simple precision number that specifies how tightly the curve bends through the coordinates of the closed cardinal spline. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Each ending point is the starting point for the next cardinal spline. In a closed cardinal spline, the curve continues through the last point in the points array and connects with the first point in the array.

#### Example

```
' ========================================================================================
' The following example draws a closed cardinal spline.
' ========================================================================================
SUB Example_DrawClosedCurve (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Define a Pen object and an array of PointF objects.
   DIM greenPen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   DIM point1 AS GpPointF : point1.x = 100.0 : point1.y = 100.0
   DIM point2 AS GpPointF : point2.x = 200.0 : point2.y = 50.0
   DIM point3 AS GpPointF : point3.x = 400.0 : point3.y = 10.0
   DIM point4 AS GpPointF : point4.x = 500.0 : point4.y = 100.0
   DIM point5 AS GpPointF : point5.x = 600.0 : point5.y = 200.0
   DIM point6 AS GpPointF : point6.x = 700.0 : point6.y = 400.0
   DIM point7 AS GpPointF : point7.x = 500.0 : point7.y = 500.0

   DIM curvePoints(6) AS GpPointF
   curvePoints(0) = point1
   curvePoints(1) = point2
   curvePoints(2) = point3
   curvePoints(3) = point4
   curvePoints(4) = point5
   curvePoints(5) = point6
   curvePoints(6) = point7

   ' // Draw the closed curve.
   graphics.DrawClosedCurve(@greenPen, @curvePoints(0), 7, 1.0)

   ' // Draw the points in the curve.
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   graphics.FillEllipse(@redBrush, 95, 95, 10, 10)
   graphics.FillEllipse(@redBrush, 495, 95, 10, 10)
   graphics.FillEllipse(@redBrush, 495, 495, 10, 10)
   graphics.FillEllipse(@redBrush, 195, 45, 10, 10)
   graphics.FillEllipse(@redBrush, 395, 5, 10, 10)
   graphics.FillEllipse(@redBrush, 595, 195, 10, 10)
   graphics.FillEllipse(@redBrush, 695, 395, 10, 10)

END SUB
' ========================================================================================
```


# <a name="DrawCurve"></a>DrawCurve (CGpGraphics)

Draws a cardinal spline.

```
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, _
   BYVAL count AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, _
   BYVAL count AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, _
   BYVAL count AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_, _
   BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_, _
   BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the cardinal spline. |
| *pts* | Pointer to an array of GpPointF objects that specify the coordinates that the cardinal spline passes through. |
| *count* | Integer that specifies the number of elements in the points array. |
| *offset* | Integer that specifies the element in the points array that specifies the point at which the cardinal spline begins. |
| *numberOfSegments* | Integer that specifies the number of segments in the cardinal spline. |
| *tension* | Simple precision number that specifies how tightly the curve bends through the coordinates of the cardinal spline. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A segment is defined as a curve that connects two consecutive points in the cardinal spline. The ending point of each segment is the starting point for the next.

#### Example

```
' ========================================================================================
' The following example draws a cardinal spline.
' ========================================================================================
SUB Example_DrawCurve (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM greenPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 3)

   DIM point1 AS GpPointF : point1.x = 100.0 : point1.y = 100.0
   DIM point2 AS GpPointF : point2.x = 200.0 : point2.y = 50.0
   DIM point3 AS GpPointF : point3.x = 400.0 : point3.y = 10.0
   DIM point4 AS GpPointF : point4.x = 500.0 : point4.y = 100.0

   DIM curvePoints(3) AS GpPointF
   curvePoints(0) = point1
   curvePoints(1) = point2
   curvePoints(2) = point3
   curvePoints(3) = point4

   ' // Draw the curve.
   graphics.DrawCurve(@greenPen, @curvePoints(0), 4)

   ' // Draw the points in the curve.
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   graphics.FillEllipse(@redBrush, 95, 95, 10, 10)
   graphics.FillEllipse(@redBrush, 195, 45, 10, 10)
   graphics.FillEllipse(@redBrush, 395, 5, 10, 10)
   graphics.FillEllipse(@redBrush, 495, 95, 10, 10)

END SUB
' ========================================================================================
```


# <a name="DrawDriverString"></a>DrawDriverString (CGpGraphics)

Draws characters at the specified positions. The method gives the client complete control over the appearance of text. The method assumes that the client has already set up the format and layout to be applied.

```
FUNCTION DrawDriverString (BYVAL pText AS UINT16 PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL pBrush AS CGpBrush PTR, BYVAL positions AS ANY PTR, BYVAL flags AS INT_, _
   BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pText* | Pointer to an array of 16-bit values. If the **DriverStringOptionsCmapLookup** flag is set, each value specifies a Unicode character to be displayed. Otherwise, each value specifies an index to a font glyph that defines a character to be displayed. |
| *length* | Integer that specifies the number of values in the text array. The length parameter can be set to –1 if the string is null terminated. |
| *pFont* | Pointer to a **Font** object that specifies the family name, size, and style of the font that is to be applied to the string. |
| *pBrush* | Pointer to a **Brush** object that is used to fill the string. |
| *positions* | If the **DriverStringOptionsRealizedAdvance** flag is set, positions is a pointer to a **GpPointF** structure that specifies the position of the first glyph. Otherwise, positions is an array of **GpPointF** structures, each of which specifies the origin of an individual glyph. |
| *flags* | Integer that specifies the options for the appearance of the string. This value must be an element of the **DriverStringOptions** enumeration or the result of a bitwise OR applied to two or more of these elements. |
| *pMatrix* | Pointer to a **Matrix** object that specifies the transformation matrix to apply to each value in the text array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A segment is defined as a curve that connects two consecutive points in the cardinal spline. The ending point of each segment is the starting point for the next. The numberOfSegments parameter must not be greater than the count parameter minus the offset parameter plus one.


# <a name="DrawEllipse"></a>DrawEllipse (CGpGraphics)

Draws an ellipse.

```
FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF) AS GpStatus
FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the ellipse. |
| *x* | The x-coordinate of the upper-left corner of the rectangle that bounds the ellipse. |
| *y* | The y-coordinate of the upper-left corner of the rectangle that bounds the ellipse. |
| *nWidth* | The width of the rectangle that bounds the ellipse. |
| *nHeight* | The height of the rectangle that bounds the ellipse. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws an ellipse.
' ========================================================================================
SUB Example_DrawEllipse (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the ellipse
   DIM bluePen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 3)
   graphics.DrawEllipse(@bluePen, 100.0, 70.0, 200.0, 100.0)

END SUB
' ========================================================================================
```


# <a name="DrawImage"></a>DrawImage (CGpGraphics)

Draws an image.

```
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL rc AS GpRect PTR) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_ _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL pt AS GpPointF PTR) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL pt AS GpPoint PTR) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPointF, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPoint, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, _
   BYVAL srcUnit AS GpUnit) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL srcx AS INT_, _
   BYVAL srcy AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL srcUnit AS GpUnit) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRectF PTR, _
   BYVAL sourceRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRect PTR, _
   BYVAL sourceRect AS GpRect PTR, BYVAL srcUnit AS GpUnit, _
    BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS SINGLE, BYVAL desty AS SINGLE, _
   BYVAL destWidth AS SINGLE, BYVAL destHeight AS SINGLE, BYVAL sourcex AS SINGLE, _
   BYVAL sourcey AS SINGLE, BYVAL sourceWidth AS SINGLE, BYVAL sourceHeight AS SINGLE, _
   BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS INT_, BYVAL desty AS INT_, _
   BYVAL destWidth AS INT_, BYVAL destHeight AS INT_, BYVAL sourcex AS INT_, BYVAL sourcey AS INT_, _
   BYVAL sourceWidth AS INT_, BYVAL sourceHeight AS INT_, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRectF PTR, _
   BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, _
   BYVAL srcheight AS SINGLE, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRect PTR, _
   BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, _
   BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS SINGLE, BYVAL desty AS SINGLE, _
   BYVAL destWidth AS SINGLE, BYVAL destHeigh AS SINGLE, BYVAL srcx AS SINGLE, _
   BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, BYVAL srcheight AS SINGLE, _
   BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS INT_, BYVAL desty AS INT_, _
   BYVAL destWidth AS INT_, BYVAL destHeigh AS INT_, BYVAL srcx AS INT_, BYVAL srcy AS INT_, _
   BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPointF PTR, _
   BYVAL nCount AS INT_, BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, _
   BYVAL srcheight AS SINGLE, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPoint PTR, _
   BYVAL nCount AS INT_, BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, _
   BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pImage* | Pointer to an **Image** object that specifies the source image. |
| *x* | The x-coordinate of the upper-left corner of the destination rectangle at which to draw the image. |
| *y* | The y-coordinate of the upper-left corner of the destination rectangle at which to draw the image. |
| *nWidth* | Optional. The width of the destination rectangle at which to draw the image. |
| *nHeight* | Optional. The height of the destination rectangle at which to draw the image. |
| *pt* | Reference to a **GpPointF** object that specifies the coordinates of the upper-left corner of the destination position at which to draw the image. |
| *srcx* | Simple precision number that specifies the x-coordinate of the upper-left corner of the portion of the source image to be drawn. |
| *srcy* | Simple precision number that specifies the y-coordinate of the upper-left corner of the portion of the source image to be drawn. |
| *srcwidth* | Simple precision number that specifies the width of the portion of the source image to be drawn. |
| *srcheight* | Simple precision number that specifies the height of the portion of the source image to be drawn. |
| *srcUnit* | Element of the Unit enumeration that specifies the unit of measure for the image. The default value is **UnitPixel**. |
| *destPoints* | Pointer to an array of **GpPointF** objects that specify the area, in a parallelogram, in which to draw the image. |
| *nCount* | Integer that specifies the number of elements in the *destPoints* array. |
| *pImageAttributes* | Pointer to an **ImageAttributes** object that specifies the color and size attributes of the image to be drawn. The default value is NULL. |
| *pCallback* | Callback method used to cancel the drawing in progress. The default value is NULL. |
| *pCallbackData* | Pointer to additional data used by the method specified by the *pCallback* parameter. The default value is NULL. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws an image.
' ========================================================================================
SUB Example_DrawImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set scaling
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(rxRatio)

   ' // Create an Image object.
   DIM pImage AS CGpImage = "climber.jpg"

   ' // Draw the original source image.
   graphics.DrawImage(@pImage, 10, 10)

   ' // Draw the scaled image.
   graphics.DrawImage(@pImage, 200, 50, 150, 75)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example draws a portion of an image. The portion of the source image to be
' drawn is scaled to fit a specified parallelogram.
' ========================================================================================
SUB Example_DrawImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set scaling
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(rxRatio)

   ' // Create an Image object.
   DIM pImage AS CGpImage = "climber.jpg"

   ' // Draw the original source image.
   graphics.DrawImage(@pImage, 10, 10)

   ' // Draw the scaled image.
   graphics.DrawImage(@pImage, 200.0, 30.0, 70.0, 20.0, 100.0, 200.0, UnitPixel)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example draws an image.
' ========================================================================================
SUB Example_DrawImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set scaling
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(rxRatio)

   ' // Create an Image object.
   DIM pImage AS CGpImage = "climber.jpg"

   ' // Create an array of PointF objects that specify the destination of the image.
   DIM destPoints(0 TO 2) AS GpPointF
   destPoints(0).x =  30 : destPoints(0).y =  30
   destPoints(1).x = 250 : destPoints(1).y =  50
   destPoints(2).x = 175 : destPoints(2).y = 120

   ' // Draw the image.
   graphics.DrawImage(@pImage, @destPoints(0), 3)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example draws the original source image and then draws a portion of the
' image in a specified parallelogram.
' ========================================================================================
SUB Example_DrawImageRectRect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set scaling
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(rxRatio)

   ' // Create an Image object.
   DIM pImage AS CGpImage = "pattern.png"

   ' // Draw the original source image.
   graphics.DrawImage(@pImage, 10, 10)

   ' // Define the portion of the image to draw.
   DIM srcX AS SINGLE = 70.0
   DIM srcY AS SINGLE = 20.0
   DIM srcWidth AS SINGLE = 100.0
   DIM srcHeight AS SINGLE = 100.0

   ' // Create an array of Point objects that specify the destination of the cropped image.
   DIM destPoints(0 TO 2) AS GpPointF
   destPoints(0).x = 230 : destPoints(0).y = 30
   destPoints(1).x = 350 : destPoints(1).y = 50
   destPoints(2).x = 275 : destPoints(2).y = 120

   ' Yet another mess of the FB GdiPlus declares.
'#ifdef __FB_64BIT__
'   DIM redToBlue AS ColorMap_
'   redToBlue.oldColor.value = GDIP_ARGB(255, 255, 0, 0)
'   redToBlue.newColor.value = GDIP_ARGB(255, 0, 0, 255)
'#else
'   DIM redToBlue AS ColorMap
'   redToBlue.from = GDIP_ARGB(255, 255, 0, 0)
'   redToBlue.to = GDIP_ARGB(255, 0, 0, 255)
'#endif

   ' // GDIP_COLORMAP is an union that solves the 32/64-bit incompatibility
   DIM redToBlue AS GDIP_COLORMAP = (GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   ' // Create an ImageAttributes object that specifies a recoloring from red to blue.
   DIM remapAttributes AS CGpImageAttributes
   RemapAttributes.SetRemapTable(1, @redToBlue)

   ' // Draw the cropped image
   graphics.DrawImage(@pImage, @destPoints(0), 3, srcX, srcY, srcWidth, srcHeight, _
                     UnitPixel, @remapAttributes, NULL, NULL)


END SUB
' ========================================================================================
```


# <a name="DrawImageFX"></a>DrawImageFX (CGpGraphics)

Draws a portion of an image after applying a specified effect.

```
FUNCTION DrawImageFX (BYVAL pImage AS CGpImage PTR, BYREF sourceRect AS GpRectF, _
   BYVAL pMatrix AS CGpMatrix PTR, BYREF gpEffect AS CGpEffect, _
   BYREF gpImageAttributes AS CGpImageAttributes, BYVAL srcUnit AS GpUnit) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pImage* | Pointer to an **Image** object that specifies the image to be drawn. |
| *sourceRect* | Pointer to a **GpRectF** structure that specifies the portion of the image to be drawn. |
| *pMatrix* | Pointer to a **Matrix** object that specifies the parallelogram in which the image portion is rendered. The destination parallelogram is calculated by applying the affine transformation stored in the matrix to the source rectangle. |
| *pEffect* | Pointer to a instance of a descendant of the **CGpEffect** class. The descendant specifies an effect or adjustment (for example, a change in contrast) that is applied to the image before rendering. The image is not permanently altered by the effect. |
| *imageAttributes* | Pointer to an **ImageAttributes** object that specifies color adjustments to be applied when the image is rendered. Can be NULL. |
| *srcUnit* | Element of the **GpUnit** enumeration that specifies the unit of measure for the source rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="DrawLine"></a>DrawLine (CGpGraphics)

Draws a line that connects two points.

```
FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, _
   BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE) AS GpStatus
FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, _
   BYVAL x2 AS INT_, BYVAL y2 AS INT_) AS GpStatus
FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR) AS GpStatus
FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the line. |
| *x1* | The x-coordinate of the starting point of the line. |
| *y1* | The y-coordinate of the starting point of the line. |
| *x2* | The x-coordinate of the ending point of the line. |
| *y2* | The y-coordinate of the ending point of the line. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a line.
' ========================================================================================
SUB Example_DrawLine (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the line
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)
   graphics.DrawLine(@blackPen, 100.0, 100.0, 500.0, 100.0)

END SUB
' ========================================================================================
```


# <a name="DrawLines"></a>DrawLines (CGpGraphics)

Draws a sequence of connected lines.

```
FUNCTION DrawLines (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS LONG) AS GpStatus
FUNCTION DrawLines (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the lines. |
| *pts* | Pointer to an array of **GpPointF** structures that specify the starting and ending points of the lines. |
| *nCount* | Integer that specifies the number of elements in the points array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a sequence of connected lines.
' ========================================================================================
SUB Example_DrawLines (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)

   ' // Create an array of PointF objects that define the lines to draw
   DIM point1 AS GpPointF : point1.x =  10 : point1.y = 10
   DIM point2 AS GpPointF : point2.x =  10 : point2.y = 100
   DIM point3 AS GpPointF : point3.x = 200 : point3.y = 50
   DIM point4 AS GpPointF : point4.x = 250 : point4.y = 300

   DIM pts(0 TO 3) AS GpPointF
   pts(0) = point1
   pts(1) = point2
   pts(2) = point3
   pts(3) = point4

   ' // Draw the lines
   graphics.DrawLines(@blackPen, @pts(0), 4)

END SUB
' ========================================================================================
```


# <a name="DrawPath"></a>DrawPath (CGpGraphics)

Draws a sequence of lines and curves defined by a **GraphicsPath** object.

```
FUNCTION DrawPath (BYVAL pPen AS CGpPen PTR, BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the path. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a GraphicsPath object.
' ========================================================================================
SUB Example_DrawPath (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GraphicsPath, and add an ellipse
   DIM ellipsePath AS CGpGraphicsPath
   ellipsePath.AddEllipse(100, 70, 200, 100)

   ' // Create a Pen object
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)

   ' // Draw ellipsePath.
   graphics.DrawPath(@blackPen, @ellipsePath)

END SUB
' ========================================================================================
```


# <a name="DrawPie"></a>DrawPie (CGpGraphics)

Draws a pie.

```
FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, _
   BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS INT_, _
   BYVAL sweepAngle AS INT_) AS GpStatus
FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect, _
   BYVAL startAngle AS INT_, BYVAL sweepAngle AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the pie. |
| *x* | The x-coordinate of the upper-left corner of the rectangle that bounds the ellipse in which to draw the pie. |
| *y* | The y-coordinate of the upper-left corner of the rectangle that bounds the ellipse in which to draw the pie. |
| *nWidth* | The width of the rectangle that bounds the ellipse in which to draw the pie. |
| *nHeight* | The height of the rectangle that bounds the ellipse in which to draw the pie. |
| *startAngle* | The angle, in degrees, between the x-axis and the starting point of the arc that defines the pie. A positive value specifies clockwise rotation. |
| *sweepAngle* | The angle, in degrees, between the starting and ending points of the arc that defines the pie. A positive value specifies clockwise rotation. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a pie.
' ========================================================================================
SUB Example_DrawPie (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the pie
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)
   graphics.DrawPie(@blackPen, 0.0, 0.0, 200.0, 100.0, 0.0, 45.0)

END SUB
' ========================================================================================
```
