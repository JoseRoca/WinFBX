# CGpRegion Class

Provides methods to manage **Regions**. **Regions** are used in clipping and hit-testing operations.

The **Region** object describes an area of the display surface. The area can be any shape. In other words, the boundary of the area can be a combination of curved and straight lines. Regions can also be created from the interiors of rectangles, paths, or a combination of these. Regions are used in clipping and hit-testing operations.

**Inherits from**: CGpBase.<br>
**Imclude file**: CGpRegion.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Create a **Region** object. |
| [Clone](#Clone) | Copies the contents of the existing **Region** object into a new **Region** object. |
| [Complement](#Complement) | Updates this region to the portion of the specified rectangle's interior that does not intersect this region. |
| [Equals](#Equals) | Determines whether this region is equal to a specified region. |
| [Exclude](#Exclude) | Updates this region to the portion of itself that does not intersect the specified rectangle's interior. |
| [GetBounds](#GetBounds) | Gets a rectangle that encloses this region. |
| [GetData](#GetData) | Gets data that describes this region. |
| [GetDataSize](#GetDataSize) | Gets the number of bytes of data that describes this region. |
| [GetHRGN](#GetHRGN) | Creates a Windows Graphics Device Interface (GDI) region from this region. |
| [GetRegionScans](#GetRegionScans) | Gets an array of rectangles that approximate this region. |
| [GetRegionScansCount](#GetRegionScansCount) | Gets the number of rectangles that approximate this region. |
| [Intersect](#Intersect) | Updates this region to the portion of itself that intersects the specified rectangle's interior. |
| [IsEmpty](#IsEmpty) | Determines whether this region is empty. |
| [IsInfinite](#IsInfinite) | Determines whether this region is infinite. |
| [IsVisible](#IsVisible) | Determines whether a point is inside this region. |
| [MakeEmpty](#MakeEmpty) | Updates this region to an empty region. In other words, the region occupies no space on the display device. |
| [MakeInfinite](#MakeInfinite) | Updates this region to an infinite region. |
| [Transform](#Transform) | Transforms this region by multiplying each of its data points by a specified matrix. |
| [Translate](#Translate) | Offsets this region by specified amounts in the horizontal and vertical directions. |
| [Union](#Union) | Updates this region to all portions (intersecting and nonintersecting) of itself and all portions of the specified rectangle's interior. |
| [Xor](#Xor) | Updates this region to the nonintersecting portions of itself and the specified rectangle's interior. |

# <a name="Constructors"></a>Constructors

Creates a region that is defined by a rectangle.

```
CONSTRUCTOR CGpRegion (BYVAL rc AS GpRectF PTR)
CONSTRUCTOR CGpRegion (BYVAL rc AS GpRect PTR)
CONSTRUCTOR CGpRegion (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE)
CONSTRUCTOR CGpRegion (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_)
```

Creates a region that is defined by a path (a **GraphicsPath** object) and has a fill mode that is contained in the **GraphicsPath** object.

```
CONSTRUCTOR CGpRegion (BYVAL pPath AS CGpGraphicsPath PTR)
```

Creates a region that is defined by data obtained from another region.

```
CONSTRUCTOR CGpRegion (BYVAL regionData AS BYTE PTR, BYVAL nSize AS LONG)
```

Creates a region that is identical to the region that is specified by a handle to a Windows Graphics Device Interface (GDI) region.

```
CONSTRUCTOR CGpRegion (BYVAL hRgn AS HRGN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Reference to a rectangle. |
| *x, y, nWidth, nHeight* | left coordinate, top coordinate, width and height of a rectangle. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the path |
| *regionData* | Pointer to an array of bytes that specifies a region. The data contained in the bytes is obtained from another region by using the **GetData** method. |
| *nSize* | Integer that specifies the number of bytes in the *regionData* array. |
| *hRgn* | Handle to an existing GDI region.  |


# <a name="Clone"></a>Clone

Copies the contents of the existing **Region** object into a new **Region** object.

```
FUNCTION Clone (BYVAL pCloneRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCloneRegion* | Pointer to the **Region** object where to copy the contents of the existing object. |

#### Example

```
' ========================================================================================
' The following example creates two regions, one from a rectangle and the other from a path.
' Next, the code clones the region that was created from a path and uses a solid brush to
' fill the cloned region. Then, it forms the union of the two regions and fills it.
' ========================================================================================
SUB Example_CloneRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM alphaBrush AS CGpSolidBrush = GDIP_ARGB(128, 0, 0, 255)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a rectangle.
   DIM rc AS GpRect = GDIP_RECT(65, 15, 70, 45)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (65, 15, 70, 45)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.x = 65 : rc.y = 15 : rc.Width = 70 : rc.Height = 45
'#endif
   DIM rectRegion AS CGpRegion = @rc

   ' // Create a region from a curved path.
   DIM pathRegion AS CGpRegion = @pPath

   ' // Make a copy (clone) of the curved region.
   DIM pClonedRegion AS CGpRegion
   pathRegion.Clone(@pClonedRegion)

   ' // Fill the cloned region with a red brush.
   graphics.FillRegion(@solidBrush, @pClonedRegion)

   ' // Form the union of the cloned region and the rectangular region.
   pClonedRegion.Union_(@rectRegion)

   ' // Fill the union of the two regions with a semitransparent blue brush.
   graphics.FillRegion(@alphaBrush, @pClonedRegion)

END SUB
' ========================================================================================
```

# <a name="Complement"></a>Complement

Updates this region to the portion of the specified rectangle's interior that does not intersect this region.

```
FUNCTION Complement (BYVAL rcf AS GpRectF PTR) AS GpStatus
FUNCTION Complement (BYVAL rcf AS GpRect PTR) AS GpStatus
```

Updates this region to the portion of the specified path's interior that does not intersect this region.

```
FUNCTION Complement (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
```

Updates this region to the portion of another region that does not intersect this region.

```
FUNCTION Complement (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rcf* | Reference to a rectangle to use to update this **Region** object. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the path to use to update this **Region** object. |
| *pRegion* | Pointer to a **Region** object to use to update this **Region** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update the region.
' ========================================================================================
SUB Example_ComplementRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   ' // Create a region from a path.
   DIM pRegion AS CGpRegion = @pPath

   ' // Update the region so that it consists of all points inside the
   ' // rectangle but outside the path region.
   pRegion.Complement(@rcf)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update the region.
' ========================================================================================
SUB Example_ComplementRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rc AS GpRect = GDIP_RECT(65, 15, 70, 45)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (65, 15, 70, 45)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.x = 65 : rc.y = 15 : rc.Width = 70 : rc.Height = 45
'#endif

   ' // Create a region from a rectangle
   DIM pRegion AS CGpRegion = @rc

   ' // Update the region so that it consists of all points inside the
   ' // rectangle but outside the path region
   pRegion.Complement(@pPath)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update the region.
' ========================================================================================
SUB Example_ComplementRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rc AS GpRect = GDIP_RECT(65, 15, 70, 45)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (65, 15, 70, 45)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.x = 65 : rc.y = 15 : rc.Width = 70 : rc.Height = 45
'#endif

   ' // Create a region from a rectangle
   DIM rectRegion AS CGpRegion = @rc
   ' // Create a region from a path
   DIM pathRegion AS CGpRegion = @pPath

   ' // Update the region so that it consists of all points inside the
   ' // rectangle but outside the path region
   rectRegion.Complement(@pathRegion)
   graphics.FillRegion(@solidBrush, @rectRegion)

END SUB
' ========================================================================================
```
