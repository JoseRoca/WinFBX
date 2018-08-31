# CGpGraphics Class

The **CGpGraphics** class provides methods for drawing lines, curves, figures, images, and text. A **Graphics** object stores attributes of the display device and attributes of the items to be drawn.

**Inherits from**: CGpBase.
**Imclude file**: CGpGraphics.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsGraphics) | Creates a Graphics object. |
| [AddMetafileComment](#AddMetafileComment) | Adds a text comment to an existing metafile. |
| [BeginContainer](#BeginContainer) | Begins a new graphics container. |
| [Clear](#Clear) | Clears a Graphics object to a specified color. |
| [DrawArc](#DrawArc) | Draws an arc. The arc is part of an ellipse. |
| [DrawBezier](#DrawBezier) | Draws a Bézier spline. |
| [DrawBeziers](#DrawBeziers) | Draws a sequence of connected Bézier splines. |
| [DrawCachedBitmap](#DrawCachedBitmap) | Draws the image stored in a CachedBitmap object. |
| [DrawClosedCurve](#DrawClosedCurve) | Draws a closed cardinal spline. |
| [DrawCurve](#DrawCurve) | Draws a cardinal spline. |
| [DrawDriverString](#DrawDriverString) | Draws characters at the specified positions. |
| [DrawEllipse](#DrawEllipse) | Draws an ellipse. |
| [DrawImage](#DrawImage) | Draws an image. |
| [DrawImageFX](#DrawImageFX) | Draws a portion of an image after applying a specified effect. |
| [DrawLine](#DrawLine) | Draws a line that connects two points. |
| [DrawLines](#DrawLines) | Draws a sequence of connected lines. |
| [DrawPath](#DrawPath) | Draws a sequence of lines and curves defined by a GraphicsPath object. |
| [DrawPie](#DrawPie) | Draws a pie. |
| [DrawPolygon](#DrawPolygon) | Draws a polygon. |
| [DrawRectangle](#DrawRectangle) | Draws a rectangle. |
| [DrawRectangles](#DrawRectangles) | Draws a sequence of rectangles. |
| [DrawString](#DrawString) | Draws a string based on a font and an origin for the string. |
| [EndContainer](#EndContainer) | Closes a graphics container that was previously opened by the BeginContainer method. |
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
| [FromHDC](#FromHDC) | Creates a Graphics object that is associated with a specified device context. |
| [FromHWND](#FromHWND) | Creates a Graphics object that is associated with a specified window. |
| [FromImage](#FromImage) | Creates a Graphics object that is associated with a specified device context. |
| [GetClip](#GetClip) | Gets the clipping region of this Graphics object. |
| [GetClipBounds](#GetClipBounds) | Gets a rectangle that encloses the clipping region of this Graphics object. |
| [GetCompositingMode](#GetCompositingMode) | Gets the compositing mode currently set for this Graphics object. |
| [GetCompositingQuality](#GetCompositingQuality) | Gets the compositing quality currently set for this Graphics object. |
| [GetDpiX](#GetDpiX) | Gets the horizontal resolution, in dots per inch, of the display device associated with this Graphics object. |
| [GetDpiY](#GetDpiY) | Gets the vertical resolution, in dots per inch, of the display device associated with this Graphics object. |
| [GetHalftonePalette](#GetHalftonePalette) | Gets a Windows halftone palette. |
| [GetHDC](#GetHDC) | Gets a handle to the device context associated with this Graphics object. |
| [GetInterpolationMode](#GetInterpolationMode) | Gets the interpolation mode currently set for this Graphics object. |
| [GetNearestColor](#GetNearestColor) | Gets the nearest color to the color that is passed in. |
| [GetPageScale](#GetPageScale) | Gets the scaling factor currently set for the page transformation of this Graphics object. |
| [GetPageUnit](#GetPageUnit) | Gets the unit of measure currently set for this Graphics object. |
| [GetPixelOffsetMode](#GetPixelOffsetMode) | Gets the pixel offset mode currently set for this Graphics object. |
| [GetRenderingOrigin](#GetRenderingOrigin) | Gets the rendering origin currently set for this Graphics object. |
| [GetSmoothingMode](#GetSmoothingMode) | Determines whether smoothing (antialiasing) is applied to the Graphics object. |
| [GetTextContrast](#GetTextContrast) | Gets the contrast value currently set for this Graphics object. |
| [GetTextRenderingHint](#GetTextRenderingHint) | Returns the text rendering mode currently set for this Graphics object. |
| [GetTransform](#GetTransform) | Gets the world transformation matrix of this Graphics object. |
| [GetVisibleClipBounds](#GetVisibleClipBounds) | Gets a rectangle that encloses the visible clipping region of this Graphics object. |
| [IntersectClip](#IntersectClip) | Updates the clipping region of this Graphics object to the portion of the specified rectangle that intersects with the current clipping region of this Graphics object. |
| [IsClipEmpty](#IsClipEmpty) | Determines whether the clipping region of this Graphics object is empty. |
| [IsVisible](#IsVisible) | Determines whether the specified point is inside the visible clipping region of this Graphics object. |
| [IsVisibleClipEmpty](#IsVisibleClipEmpty) | Determines whether the visible clipping region of this Graphics object is empty. |
| [MeasureCharacterRanges](#MeasureCharacterRanges) | Gets a set of regions each of which bounds a range of character positions within a string. |
| [MeasureDriverString](#MeasureDriverString) | Measures the bounding box for the specified characters and their corresponding positions. |
| [MeasureString](#MeasureString) | Measures the extent of the string in the specified font, format, and layout rectangle. |
| [MultiplyTransform](#MultiplyTransform) | Updates this Graphics object's world transformation matrix with the product of itself and another matrix. |
| [ReleaseHDC](#ReleaseHDC) | Releases a device context handle obtained by a previous call to the GetHDC method of this Graphics object. |
| [ResetClip](#ResetClip) | Sets the clipping region of this Graphics object to an infinite region. |
| [ResetTransform](#ResetTransform) | Sets the world transformation matrix of this Graphics object to the identity matrix. |
| [Restore](#Restore) | Sets the state of this Graphics object to the state stored by a previous call to the Save method of this Graphics object. |
| [RotateTransform](#RotateTransform) | Updates the world transformation matrix of this Graphics object with the product of itself and a rotation matrix. |
| [Save](#Save) | Saves the current state (transformations, clipping region, and quality settings) of this Graphics object. |
| [ScaleTransform](#ScaleTransform) | Updates this Graphics object's world transformation matrix with the product of itself and a scaling matrix. |
| [SetClip](#SetClip) | Updates the clipping region of this Graphics object to a region that is the combination of itself and the clipping region of another Graphics object. |
| [SetCompositingMode](#SetCompositingMode) | Sets the compositing mode of this Graphics object. |
| [SetCompositingQuality](#SetCompositingQuality) | Sets the compositing quality of this Graphics object. |
| [SetInterpolationMode](#SetInterpolationMode) | Sets the interpolation mode of this Graphics object. |
| [SetPageScale](#SetPageScale) | Sets the scaling factor for the page transformation of this Graphics object. |
| [SetPageUnit](#SetPageUnit) | Sets the unit of measure for this Graphics object. |
| [SetPixelOffsetMode](#SetPixelOffsetMode) | Sets the pixel offset mode of this Graphics object. |
| [SetRenderingOrigin](#SetRenderingOrigin) | Sets the rendering origin of this Graphics object. |
| [SetSmoothingMode](#SetSmoothingMode) | Sets the rendering quality of the Graphics object. |
| [SetTextContrast](#SetTextContrast) | Sets the contrast value of this Graphics object. The contrast value is used for antialiasing text. |
| [SetTextRenderingHint](#SetTextRenderingHint) | Sets the text rendering mode of this Graphics object. |
| [SetTransform](#SetTransform) | Sets the world transformation of this Graphics object. |
| [TransformPoints](#TransformPoints) | Converts an array of points from one coordinate space to another. |
| [TranslateClip](#TranslateClip) | Translates the clipping region of this Graphics object. |
| [TranslateTransform](#TranslateTransform) | Updates this Graphics object's world transformation matrix with the product of itself and a translation matrix. |

# CGpGraphicsPath Class

The **CGpGraphicsPath** allows the creation of **GraphicPath** objects. A **GraphicsPath** object stores a sequence of lines, curves, and shapes. You can draw the entire sequence by calling the **DrawPath** method of a **Graphics** object. You can partition the sequence of lines, curves, and shapes into figures, and with the help of a **GraphicsPathIterator** object, you can draw selected figures. You can also place markers in the sequence, so that you can draw selected portions of the path.

**Inherits from**: CGpBase.
**Imclude file**: CGpPath.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsGraphicsPath) | Creates a Graphics object. |
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
| [Clone](#Clone) | Copies the contents of the existing GraphicsPath object into a new GraphicsPath object. |
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
| [Reset](#Reset) | Empties the path and sets the fill mode to FillModeAlternate. |
| [Reverse](#Reverse) | Reverses the order of the points that define this path's lines and curves. |
| [SetFillMode](#SetFillMode) | Sets the fill mode of this path. |
| [SetMarker](#SetMarker) | Designates the last point in this path as a marker point. |
| [StartFigure](#StartFigure) | Starts a new figure without closing the current figure. |
| [Transform](#Transform) | Multiplies each of this path's data points by a specified matrix. |
| [Warp](#Warp) | Applies a warp transformation to this path. |
| [Widen](#Widen) | Replaces this path with curves that enclose the area that is filled when this path is drawn by a specified pen. |

# CGpGraphicsPathIterator Class

The **CGpGraphicsPathIterator** class provides methods for isolating selected subsets of the path stored in a **GraphicsPath** object. A path consists of one or more figures. You can use a **GraphicsPathIterator** to isolate one or more of those figures. A path can also have markers that divide the path into sections. You can use a **GraphicsPathIterator** object to isolate one or more of those sections.

**Inherits from**: CGpBase.
**Imclude file**: CGpPath.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsGraphicsPathIterator) | Creates a new GraphicsPathIterator object and associates it with a GraphicsPath object. |
| [CopyData](#CopyData) | Copies a subset of the path's data points to a PointF array and copies a subset of the path's point types to a BYTE array. |
| [Enumerate](#Enumerate) | Copies the path's data points to a PointF array and copies the path's point types to a BYTE array. |
| [GetCount](#GetCount) | Returns the number of data points in the path. |
| [GetSubpathCount](#GetSubpathCount) | Returns the number of subpaths (also called figures) in the path. |
| [HasCurve](#HasCurve) | Determines whether the path has any curves. |
| [NextMarker](#NextMarker) | Gets the starting index and the ending index of the next marker-delimited section in this iterator's associated path. |
| [NextPathType](#NextPathType) | Gets the starting index and the ending index of the next group of data points that all have the same type. |
| [NextSubpath](#NextSubpath) | Gets the starting index and the ending index of the next subpath (figure) in this iterator's associated path. |
| [Rewind](#Rewind) | Rewinds this iterator to the beginning of its associated path. |
