# CGpBrush Class

The **CGpBrush** class is a base interface that defines a **Brush** object. A **Brush** object is used to paint the interior of graphics shapes, such as rectangles, ellipses, pies, polygons, and paths. You must not instantiate the **CGpBrush** class, but to use one of its derived classes.

A closed shape, such as a rectangle or an ellipse, consists of an outline and an interior. The outline is drawn with a pen and the interior is filled with a brush. GDI+ provides several brush classes for filling the interiors of closed shapes: **CGpSolidBrush**, **CGpHatchBrush**, **CGpTextureBrush**, **CGpLinearGradientBrush**, and **CGpPathGradientBrush**. All of these classes inherit from the **CGpBrush** class.

**Inherits from**: CGpBase.
**Imclude file**: CGpBrush.inc.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Clone](#CloneBrush) | Copies the contents of the existing **Brush** object into a new **Brush** object. |
| [GetType](#GetTypeBrush) | Gets the type of this brush. |

# CGpSolidBrush Class

The **SolidBrush** object defines a solid color Brush object. A **Brush** object is used to fill in shapes similar to the way a paint brush can paint the inside of a shape.

**Inherits from**: CGpBrush.
**Imclude file**: CGpBrush.inc.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#ConstructorSolidBrush) | Creates a **SolidBrush** object based on a color. |
| [GetColor](#GetColorSolidBrush) | Gets the color of this brush. |
| [SetColor](#SetColorSolidBrush) | Sets the color of this brush. |

# CGpHatchBrush Class

Creates a **HatchBrush** object based on a hatch style, a foreground color, and a background color.

**Inherits from**: CGpBrush.
**Imclude file**: CGpBrush.inc.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#ConstructorHatchBrush) | Creates a **HatchBrush** object based on a hatch style, a foreground color, and a background color. |
| [GetBackgroundColor](#GetBackgroundColor) | Gets the background color of this hatch brush. |
| [GetForegroundColor](#GetForegroundColor) | Gets the foreground color of this hatch brush. |
| [GetHatchStyle](#GetHatchStyle) | Gets the hatch style of this hatch brush. |

# CGpLinearGradientBrush Class

Defines a brush that paints a color gradient in which the color changes evenly from the starting boundary line of the linear gradient brush to the ending boundary line of the linear gradient brush. The boundary lines of a linear gradient brush are two parallel straight lines. The color gradient is perpendicular to the boundary lines of the linear gradient brush, changing gradually across the stroke from the starting boundary line to the ending boundary line. The color gradient has one color at the starting boundary line and another color at the ending boundary line.

**Inherits from**: CGpBrush.
**Imclude file**: CGpBrush.inc.

### Constructors and Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorLGBrush) | Creates a **LinearGradientBrush** object. |
| [GetBlend](#GetBlendLGBrush) | Gets the blend factors and their corresponding blend positions. |
| [GetBlendCount](#GetBlendCountLGBrush) | Gets the number of blend factors currently set. |
| [GetGammaCorrection](#GetGammaCorrectionLGBrush) | Determines whether gamma correction is enabled for this brush. |
| [GetInterpolationColorCount](#GetInterpolationColorCountLGBrush) | Gets the number of colors currently set to be interpolated. |
| [GetInterpolationColors](#GetInterpolationColorsLGBrush) | Gets the blend factors and their corresponding blend positions. |
| [GetLinearColors](#GetLinearColors) | Gets the starting color and ending color. |
| [GetRectangle](#GetRectangle) | Gets the rectangle that defines the boundaries of the gradient. |
| [GetTransform](#GetTransform) | Gets the transformation matrix. |
| [GetWrapMode](#GetWrapMode) | Gets the wrap mode currently set for this brush. |
| [MultiplyTransform](#MultiplyTransform) | Updates this brush's transformation matrix with the product of itself and another matrix. |
| [ResetTransform](#ResetTransform) | Resets the transformation matrix to the identity matrix. |
| [RotateTransform](#RotateTransform) | Updates this brush's current transformation matrix with the product of itself and a rotation matrix. |
| [ScaleTransform](#ScaleTransform) | Updates this brush's current transformation matrix with the product of itself and a scaling matrix. |
| [SetBlend](#SetBlend) | Sets the blend factors and the blend positions to create a custom blend. |
| [SetBlendBellShape](#SetBlendBellShape) | Sets the blend bell shape. |
| [SetBlendTriangularShape](#SetBlendTriangularShape) | Sets the blend triangular shape. |
| [SetGammaCorrection](#SetBlendTriangularShape) | Specifies whether gamma correction is enabled. |
| [SetInterpolationColors](#SetInterpolationColors) | Sets the colors to be interpolated and their corresponding blend positions. |
| [SetLinearColors](#SetLinearColors) | Sets the starting color and ending color. |
| [SetTransform](#SetTransform) | Sets the transformation matrix. |
| [SetWrapMode](#SetWrapMode) | Sets the wrap mode. |
| [TranslateTransform](#TranslateTransform) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |

# CGpPathGradientBrush Class

A **PathGradientBrush** object stores the attributes of a color gradient that you can use to fill the interior of a path with a gradually changing color. A path gradient brush has a boundary path, a boundary color, a center point, and a center color. When you paint an area with a path gradient brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point.

**Inherits from**: CGpBrush.
**Imclude file**: CGpBrush.inc.

### Constructors and Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorPGBrush) | Creates a PathGradientBrush object based on an array of points. Initializes the wrap mode of the path gradient brush. |
| [GetBlend](#GetBlendPGBrush) | Gets the blend factors and their corresponding blend positions. |
| [GetBlendCount](#GetBlendCountPGBrush) | Gets the number of blend factors currently set. |
| [GetCenterColor](#GetCenterColor) | Gets center cpñpr of the brush. |
| [GetCenterPoint](#GetCenterPoint) | Gets the center point of the brush. |
| [GetFocusScales](#GetFocusScales) | Gets the focus scales of the brush. |
| [GetGammaCorrection](#GetGammaCorrectionPGBrush) | Determines whether gamma correction is enabled for this brush. |
| [GetInterpolationColorCount](#GetInterpolationColorCountPGBrush) | Gets the number of preset colors currently specified for this brush. |
| [GetInterpolationColors](#GetInterpolationColors^GBrush) | Gets preset colors and blend positions currently specified for this brush. |
| [GetPointCount](#GetPointCount) | Gets the number of points in the array of points that defines this brush's boundary path. |
| [GetRectangle](#GetRectangle) | Gets the smallest rectangle that encloses the boundary path of this brush. |
| [GetSurroundColorCount](#GetSurroundColorCount) | Gets the number of colors that have been specified for the boundary path of this brush. |
| [GetSurroundColors](#GetSurroundColors) | Gets the surround colors currently specified for this brush. |
| [GetTransform](#GetTransform) | Gets the transformation matrix. |
| [GetWrapMode](#GetWrapMode) | Gets the wrap mode currently set for this brush. |
| [MultiplyTransform](#MultiplyTransform) | Updates this brush's transformation matrix with the product of itself and another matrix. |
| [ResetTransform](#ResetTransform) | Resets the transformation matrix to the identity matrix. |
| [RotateTransform](#RotateTransform) | Updates this brush's current transformation matrix with the product of itself and a rotation matrix. |
| [ScaleTransform](#ScaleTransform) | Updates this brush's current transformation matrix with the product of itself and a scaling matrix. |
| [SetBlend](#SetBlend) | Sets the blend factors and the blend positions to create a custom blend. |
| [SetBlendBellShape](#SetBlendBellShape) | Sets the blend bell shape. |
| [SetBlendTriangularShape](#SetBlendTriangularShape) | Sets the blend triangular shape. |
| [SetCenterColor](#SetCenterColor) | Sets the center color of this brush. |
| [SetCenterPoint](#SetCenterPoint) | Sets the center point of this brush. |
| [SetFocusScales](#SetFocusScales) | Sets the focus scales of this brush. |
| [SetGammaCorrection](#SetBlendTriangularShape) | Specifies whether gamma correction is enabled. |
| [SetInterpolationColors](#SetInterpolationColors) | Sets the colors to be interpolated and their corresponding blend positions. |
| [SetSurroundColors](#SetSurroundColors) | Sets the surround colors of this brush. |
| [SetTransform](#SetTransform) | Sets the transformation matrix. |
| [SetWrapMode](#SetWrapMode) | Sets the wrap mode. |
| [TranslateTransform](#TranslateTransform) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |

# CGpTextureBrush Class

Defines a **Brush** object that contains an **Image** object that is used for the fill. The fill image can be transformed by using the local **Matrix** object contained in the **Brush** object.

**Inherits from**: CGpBrush.
**Imclude file**: CGpBrush.inc.

### Constructors and Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorTBrush) | Creates a texture brush. |
| [GetImage](#GetImage) | Gets a pointer to the **Image** object that is defined by this brush. |
| [GetTransform](#GetTransform) | Gets the transformation matrix. |
| [GetWrapMode](#GetWrapMode) | Gets the wrap mode currently set for this brush. |
| [MultiplyTransform](#MultiplyTransform) | Updates this brush's transformation matrix with the product of itself and another matrix. |
| [ResetTransform](#ResetTransform) | Resets the transformation matrix to the identity matrix. |
| [RotateTransform](#RotateTransform) | Updates this brush's current transformation matrix with the product of itself and a rotation matrix. |
| [ScaleTransform](#ScaleTransform) | Updates this brush's current transformation matrix with the product of itself and a scaling matrix. |
| [SetTransform](#SetTransform) | Sets the transformation matrix. |
| [SetWrapMode](#SetWrapMode) | Sets the wrap mode. |
| [TranslateTransform](#TranslateTransform) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |

# <a name="CloneBrush"></a>Clone (CGpBrush)

Copies the contents of the existing **Brush** object into a new **Brush** object.

```
FUNCTION Clone (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a variable that will receive a pointer to the cloned **Brush** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a SolidBrush object, clones it, and then uses the clone
' to fill a rectangle.
' ========================================================================================
SUB Example_CloneBrush (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a SolidBrush object
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   ' // Create a clone of solidBrush
   DIM cloneBrush AS CGpSolidBrush
   solidBrush.Clone(@cloneBrush)

   ' // Use cloneBrush to fill a rectangle
   graphics.FillRectangle(@cloneBrush, 0, 0, 100, 100)

END SUB
' ========================================================================================
```

# <a name="GetTypeBrush"></a>GetType (CGpBrush)

Gets the type of this brush.

```
FUNCTION GetType () AS GpBrushType
```

#### Return value

This method returns the type of this brush. The value returned is one of the elements of the **BrushType** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a SolidBrush object, checks the type of the object, and
' then, if the type is BrushTypeSolidColor, uses the brush to fill a rectangle.
' ========================================================================================
SUB Example_GetType (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a SolidBrush object
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)

   ' // Get the type of solidBrush
   DIM nType AS BrushType = solidBrush.GetType

   ' // If the type of solidBrush is BrushTypeSolidColor, use it to fill a rectangle
   IF nType = BrushTypeSolidColor THEN
      graphics.FillRectangle(@solidBrush, 0, 0, 100, 100)
   END IF

END SUB
' ========================================================================================
```

# <a name="ConstructorSolidBrush"></a>Constructor (CGpSolidBrush)

Create a **SolidBrush** object based on a color.

```
CONSTRUCTOR CGpSolidBrush (BYVAL colour AS ARGB = &hFF000000)
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | An ARGB color that specifies the initial color of this solid brush. |

# <a name="GetColorSolidBrush"></a>GetColor (CGpSolidBrush)

Gets the color of this solid brush.

```
FUNCTION GetColor (BYVAL colour AS ARGB PTR) AS GpStatus
FUNCTION GetColor () AS ARGB
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a variable that receives the color of this solid brush. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

The second overloaded function returns the ARGB color as the result of the function.

# <a name="SetColorSolidBrush"></a>SetColor (CGpSolidBrush)

Sets the color of this solid brush.

```
FUNCTION SetColor (BYVAL colour AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | The color to be set in this solid brush. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a solid brush and uses it to fill a rectangle. The code
' uses GdipSetSolidFillColor to change the color of the solid brush and then paints a
' second rectangle the new color.
' ========================================================================================
SUB Example_SetColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a solid brush, and use it to fill a rectangle
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   graphics.FillRectangle(@solidBrush, 10, 10, 200, 100)

   ' // Change the color of the brush to red, and fill another rectangle
   solidBrush.SetColor(GDIP_ARGB(255, 255, 0, 0))
   graphics.FillRectangle(@solidBrush, 220, 10, 200, 100)

END SUB
' ========================================================================================
```

# <a name="ConstructorHatchBrush"></a>Constructor (CGpHatchBrush)

Creates a **HatchBrush** object based on a hatch style, a foreground color, and a background color.

```
FUNCTION HatchBrush (BYVAL hatchStyle AS HatchStyle, BYVAL foreColor AS ARGB, _
   BYVAL backColor AS ARGB = &HFF000000)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hatchStyle* | Element of the **HatchStyle** enumeration that specifies the pattern of hatch lines that will be used. |
| *foreColor* | Reference to a color to use for the hatch lines. |
| *backColor* | Optional. Reference to a color to use for the background. |

# <a name="GetBackgroundColor"></a>GetBackgroundColor (CGpHatchBrush)

Gets the background color of this hatch brush.

```
FUNCTION GetBackgroundColor (BYVAL colour AS ARGB PTR) AS GpStatus
FUNCTION GetBackgroundColor () AS ARGB
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a variable that receives the background color. The background color defines the color over which the hatch lines are drawn. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

The second overloaded function returns the ARGB color as the result of the function.

#### Example

```
' ========================================================================================
' The following example sets up three Color objects: black, turquoise, and current
' (initialized to black). A rectangle is painted by using turquoise as the background
' color and black as the foreground color. Then the HatchBrush.GetBackgroundColor method
' is used to get the current color of the brush (which at the time is turquoise). The
' address of the current Color object (initialized to black) is passed as the return point
' for the call to HatchBrush,GetBackgroundColor. When the rectangle is painted again,
' note that the background color is again turquoise (not black). This shows that the call
' to HatchBrush.GetBackgroundColor was successful.
' ========================================================================================
SUB Example_HatchBrushGetBackgroundColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Set colors
   DIM black AS ARGB = GDIP_ARGB(255, 0, 0, 0)           ' // foreground
   DIM turquoise AS ARGB = GDIP_ARGB(255, 0, 255, 255)   ' // background
   DIM current AS ARGB = GDIP_ARGB(255, 0, 0, 0)         ' // new foreground

   ' // Set and then draw the first hatch style.
   DIM brush AS CGpHatchBrush = CGpHatchBrush(HatchStyleHorizontal, black, turquoise)
   graphics.FillRectangle(@brush, 20, 20, 100, 50)

   ' // Get the current background color of the brush.
   brush.GetBackgroundColor(@current)

   ' // Draw the rectangle again using the current color.
   DIM brush2 AS CGpHatchBrush = CGpHatchBrush(HatchStyleDiagonalCross, black, current)
   graphics.FillRectangle(@brush2, 130, 20, 100, 50)

END SUB
' ========================================================================================
```

# <a name="GetForegroundColor"></a>GetForegroundColor (CGpHatchBrush)

Gets the foreground color of this hatch brush.

```
FUNCTION GetForegroundColor (BYVAL colour AS ARGB PTR) AS GpStatus
FUNCTION GetForegroundColor () AS ARGB
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a variable that receives the foreground color. The foreground color defines the color of the hatch lines. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

The second overloaded function returns the ARGB color as the result of the function.

#### Example

```
' ========================================================================================
' The following example sets up three Color objects: black, turquoise, and current
' (initialized to black). A rectangle is painted by using turquoise as the background
' color and black as the foreground color. Then the HatchBrush.GetBackgroundColor method
' is used to get the current color of the brush (which at the time is turquoise). The
' address of the current Color object (initialized to black) is passed as the return point
' for the call to HatchBrush,GetBackgroundColor. When the rectangle is painted again,
' note that the background color is again turquoise (not black). This shows that the call
' to HatchBrush.GetBackgroundColor was successful.
' ========================================================================================
SUB Example_HatchBrushGetForegroundColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Set colors
   DIM blue AS ARGB = GDIP_ARGB(255, 0, 0, 255)          ' // foreground
   DIM turquoise AS ARGB = GDIP_ARGB(255, 0, 255, 255)   ' // background
   DIM current AS ARGB = GDIP_ARGB(255, 0, 0, 0)         ' // new foreground

   ' // Set and then draw the first hatch style.
   DIM brush AS CGpHatchBrush = CGpHatchBrush(HatchStyleHorizontal, blue, turquoise)
   graphics.FillRectangle(@brush, 20, 20, 100, 50)

   ' // Get the current foreground color of the brush.
   brush.GetForegroundColor(@current)

   ' // Draw the rectangle again using the current color.
   DIM brush2 AS CGpHatchBrush = CGpHatchBrush(HatchStyleDiagonalCross, current, turquoise)
   graphics.FillRectangle(@brush2, 130, 20, 100, 50)

END SUB
' ========================================================================================
```

# <a name="GetHatchStyle"></a>GetHatchStyle (CGpHatchBrush)

Gets the hatch style of this hatch brush.

```
FUNCTION GetHatchStyle () AS HatchStyle
```

#### Return value

This method returns the hatch style, which is one of the elements of the **HatchStyle** enumeration.

#### Example

```
' ========================================================================================
' The following example sets up two hatch styles: horiz and current (initialized to
' HatchStyleDiagonalCross). A rectangle that uses horiz as the hatch style is painted.
' Then the HatchBrush.GetHatchStyle method is used to get the current hatch style of the
' brush (which at the time is HatchStyleHorizontal). The address of the current HatchStyle
' object (initialized to HatchStyleDiagonalCross) is passed as the return point for the
' call to GetHatchStyle. When the rectangle is painted again, notice that the hatch style
' is again HatchStyleHorizontal (not HatchStyleDiagonalCross). This shows that the call to
' HatchBrush.GetHatchStyle was successful. 
' ========================================================================================
SUB Example_HatchBrushGetHatchStyle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Set colors
   DIM blue AS ARGB = GDIP_ARGB(255, 0, 0, 255)          ' // foreground
   DIM turquoise AS ARGB = GDIP_ARGB(255, 0, 255, 255)   ' // background

   ' // Set hatch styles
   DIM horiz AS HatchStyle = HatchStyleHorizontal
   DIM current AS HatchStyle = HatchStyleDiagonalCross

   ' // Set and then draw the first hatch style.
   DIM brush AS CGpHatchBrush = CGpHatchBrush(horiz, blue, turquoise)
   graphics.FillRectangle(@brush, 20, 20, 100, 50)

   ' // Get the current hatch style of the brush.
   current = brush.GetHatchStyle
   
   ' // Get the current background color of the brush.
   brush.GetBackgroundColor(@current)

   ' // Draw the rectangle again using the current hatch style.
   DIM brush2 AS CGpHatchBrush = CGpHatchBrush(current, blue, turquoise)
   graphics.FillRectangle(@brush2, 130, 20, 100, 50)

END SUB
' ========================================================================================
```

# <a name="ConstructorLGBrush"></a>Constructors (CGpLinearGradientBrush)

Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors.

```
CONSTRUCTOR CGpLinearGradientBrush (BYVAL point1 AS GpPointF PTR, BYVAL point2 AS GpPointF PTR, _
   BYVAL color1 AS ARGB, BYVAL color2 AS ARGB)
CONSTRUCTOR CGpLinearGradientBrush (BYVAL point1 AS GpPoint PTR, BYVAL point2 AS GpPoint PTR, _
   BYVAL color1 AS ARGB, BYVAL color2 AS ARGB)
```

Creates a **LinearGradientBrush** object based on a rectangle and mode of direction.

```
CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB, BYVAL mode AS LinearGradientMode)
CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB, BYVAL mode AS LinearGradientMode)
```

Creates a **LinearGradientBrush** object from a rectangle and angle of direction.

```
CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL)
CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *point1* | Reference to a **GpPoint** structure that specifies the starting point of the gradient. The starting boundary line passes through the starting point. |
| *point2* | Reference to a **GpPoint** structure that specifies the ending point of the gradient. The ending boundary line passes through the ending point. |
| *color1* | An ARGB value that specifies the color at the starting boundary line of this linear gradient brush. |
| *color2* | An ARGB value that specifies the color at the ending boundary line of this linear gradient brush. |
| *rc* | Reference to a rectangle that specifies the starting and ending points of the gradient. The upper-left corner of the rectangle is the starting point. The lower-right corner is the ending point. |
| *mode* | Element of the **LinearGradientMode** enumeration that specifies the direction of the gradient. |
| *angle* | Real number that, if *isAngleScalable* is TRUE, specifies the base angle from which the angle of the directional line is calculated, or that, if *isAngleScalable* is FALSE, specifies the angle of the directional line. The angle is measured from the top of the rectangle that is specified by rect and must be in degrees. The gradient follows the directional line. |
| *isAngleScalable* | BOOL value that specifies whether the angle is scalable. If isAngleScalable is TRUE, the angle of the directional line is scalable; otherwise, the angle is not scalable. |

# <a name="GetBlendLGBrush"></a>GetBlend (CGpLinearGradientBrush)

Gets the blend factors and their corresponding blend positions from a **LinearGradientBrush** object.

```
FUNCTION GetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, _
   BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *blendFactors* | Pointer to an array that receives the blend factors. Each number in the array indicates a percentage of the ending color and is in the range from 0.0 through 1.0. |
| *blendPositions* | Pointer to an array that receives the blend positions. Each number in the array indicates a percentage of the distance between the starting boundary and the ending boundary and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the gradient and 1.0 indicates the ending boundary. A blend position between 0.0 and 1.0 indicates a line, parallel to the boundary lines, that is a certain fraction of the distance from the starting boundary to the ending boundary. For example, a blend position of 0.7 indicates the line that is 70 percent of the distance from the starting boundary to the ending boundary. The color is constant on lines that are parallel to the boundary lines. |
| *count* | Integer that specifies the number of blend factors to retrieve. Before calling the **GetBlend** method of a **LinearGradientBrush** object, call the **GetBlendCount** method of that same **LinearGradientBrush** object to determine the current number of blend factors. The number of blend positions retrieved is the same as the number of blend factors retrieved. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **LinearGradientBrush** object has two parallel boundaries: a starting boundary and an ending boundary. A color is associated with each of these two boundaries. Each boundary is a straight line that passes through a specified point — the starting boundary passes through the starting point; the ending boundary passes through the ending point — and is perpendicular to the direction of the linear gradient brush. The direction of the linear gradient brush follows the line that is defined by the starting and ending points. This line, the "directional line," may be horizontal, vertical, or diagonal. All points that lie on a line that is parallel to the boundaries are the same color. When you fill an area with a linear gradient brush, the color changes gradually from one line to the next as you move along the directional line from the starting boundary to the ending boundary. By default, the change in color is proportional to the change in distance; that is, a line 30 percent of the distance between the starting boundary and the ending boundary has a color that is 30 percent of the distance between the starting boundary color and the ending boundary color. The color pattern is repeated outside of the starting and ending boundaries.

You can call the **SetBlend** method of a LinearGradientBrush object to customize the relationship between color and distance. For example, suppose you set the blend positions to {0, 0.5, 1} and you set the blend factors to {0, 0.3, 1}. Then a line 50 percent of the distance between the starting boundary and the ending boundary will have a color that is 30 percent of the distance between the starting boundary color and the ending boundary color.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets its blend, and uses the brush
' to fill a rectangle. The code then gets the blend. The blend factors and positions can
' then be inspected or used in some way.
' ========================================================================================
SUB Example_GetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.6, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.2, 0.8, 1.0}
   DIM rcf AS GpRectF = GDIP_RECTF(0, 0, 100, 50)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rcf, GDIP_ARGB(255, 255, 0, 0), _
       GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   linGrBrush.SetBlend(@factors(0), @positions(0), 4)
   graphics.FillRectangle(@linGrBrush, @rcf)

   DIM blendCount AS LONG = linGrBrush.GetBlendCount
   DIM rgFactors(blendCount - 1) AS SINGLE
   DIM rgPositions(blendCount - 1) AS SINGLE

   linGrBrush.GetBlend(@rgFactors(0), @rgPositions(0), blendCount)

   FOR j AS LONG = 0 TO blendCount - 1
'      // Inspect or use the value in rgFactors(j)
'      // Inspect or use the value in rgPositions(j)
      OutputDebugString STR(rgFactors(j)) & STR(rgPositions(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetBlendPGBrush"></a>GetBlend (CGpPathGradientBrush)

Gets the blend factors and the corresponding blend positions currently set for this path gradient brush.

```
FUNCTION GetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, _
   BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *blendFactors* | Pointer to an array that receives the blend factors. |
| *blendPositions* | Pointer to an array that receives the blend positions. |
| *count* | Integer that specifies the number of blend factors to retrieve. Before calling the **GetBlend** method of a **PathGradientBrush** object, call the **GetBlendCount** method of that same **PathGradientBrush** object to determine the current number of blend factors. The number of blend positions retrieved is the same as the number of blend factors retrieved. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **PathGradientBrush** object has a boundary path and a center point. When you fill an area with a path gradient brush, the color changes gradually as you move from the boundary path to the center point. By default, the color is linearly related to the distance, but you can customize the relationship between color and distance by calling the SetBlend method.

#### Example

```
' ========================================================================================
' The following example demonstrates several methods of the PathGradientBrush class including
' PathGradientBrush.SetBlend, PathGradientBrush.GetBlendCount, and PathGradientBrush.GetBlend.
' The code creates a PathGradientBrush object and calls the PathGradientBrush.SetBlend method
' to establish a set of blend factors and blend positions for the brush. Then the code calls
' the PathGradientBrush.GetBlendCount method to retrieve the number of blend factors. After
' the number of blend factors is retrieved, the code allocates two buffers: one to receive
' the array of blend factors and one to receive the array of blend positions. Then the code
' calls the PathGradientBrush.GetBlend method to retrieve the blend factors and the blend
' positions.
' ========================================================================================
SUB Example_GetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}

   pthGrBrush.SetBlend(@factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM blendCount AS LONG = pthGrBrush.GetBlendCount
   DIM rgFactors(blendCount - 1) AS SINGLE
   DIM rgPositions(blendCount - 1) AS SINGLE

   pthGrBrush.GetBlend(@rgFactors(0), @rgPositions(0), blendCount)

   FOR j AS LONG = 0 TO blendCount - 1
'      // Inspect or use the value in rgFactors(j)
'      // Inspect or use the value in rgPositions(j)
      OutputDebugString STR(rgFactors(j)) & STR(rgPositions(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetBlendCountLGBrush"></a>GetBlendCount (CGpLinearGradientBrush)

Gets the number of blend factors currently set for this **LinearGradientBrush** object.

```
FUNCTION GetBlendCount () AS INT_
```

#### Return value

This method returns the number of blend factors currently set for this **LinearGradientBrush** object. If no custom blend has been set by using **SetBlend**, or if invalid positions were passed to **SetBlend**, then **GetBlend** returns 1.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets its blend, and uses the brush
' to fill a rectangle. The code then gets the blend. The blend factors and positions can
' then be inspected or used in some way.
' ========================================================================================
SUB Example_GetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.6, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.2, 0.8, 1.0}

   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(100, 0)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, _
      GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))
   linGrBrush.SetBlend(@factors(0), @positions(0), 4)

   ' // Use the linear gradient brush to fill a rectangle.
   graphics.FillRectangle(@linGrBrush, 0, 0, 100, 50)

   ' // Obtain information about the linear gradient brush.
   DIM blendCount AS LONG = linGrBrush.GetBlendCount
   DIM rgFactors(blendCount - 1) AS SINGLE
   DIM rgPositions(blendCount - 1) AS SINGLE

   linGrBrush.GetBlend(@rgFactors(0), @rgPositions(0), blendCount)

   FOR j AS LONG = 0 TO blendCount - 1
'      // Inspect or use the value in rgFactors(j)
'      // Inspect or use the value in rgPositions(j)
      OutputDebugString STR(rgFactors(j)) & STR(rgPositions(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetBlendCountPGBrush"></a>GetBlendCount (CGpPathGradientBrush)

Gets the number of blend factors currently set for this path gradient brush.

```
FUNCTION GetBlendCount () AS INT_
```

#### Return value

Before you call the **GetBlend** method of a **PathGradientBrush** object, you must allocate two buffers: one to receive an array of blend factors and one to receive an array of blend positions. To determine the size of the required buffers, call the **GetBlendCount** method of the **PathGradientBrush** object. The size (in bytes) of each buffer should be the return value of **GetBlendCount** multiplied by 4 (the size of a simple precision number).

#### Example

```
' ========================================================================================
' The following example demonstrates several methods of the PathGradientBrush class including
' PathGradientBrush.SetBlend, PathGradientBrush.GetBlendCount, and PathGradientBrush.GetBlend.
' The code creates a PathGradientBrush object and calls the PathGradientBrush.SetBlend method
' to establish a set of blend factors and blend positions for the brush. Then the code calls
' the PathGradientBrush.GetBlendCount method to retrieve the number of blend factors. After
' the number of blend factors is retrieved, the code allocates two buffers: one to receive
' the array of blend factors and one to receive the array of blend positions. Then the code
' calls the PathGradientBrush.GetBlend method to retrieve the blend factors and the blend
' positions.
' ========================================================================================
SUB Example_GetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}

   pthGrBrush.SetBlend(@factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM blendCount AS LONG = pthGrBrush.GetBlendCount
   DIM rgFactors(blendCount - 1) AS SINGLE
   DIM rgPositions(blendCount - 1) AS SINGLE

   pthGrBrush.GetBlend(@rgFactors(0), @rgPositions(0), blendCount)

   FOR j AS LONG = 0 TO blendCount - 1
'      // Inspect or use the value in rgFactors(j)
'      // Inspect or use the value in rgPositions(j)
      OutputDebugString STR(rgFactors(j)) & STR(rgPositions(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetCenterColor"></a>GetCenterColor (CGpPathGradientBrush)

Gets the center color of this path gradient brush.

```
FUNCTION GetCenterColor (BYVAL colour AS ARGB PTR) AS GpStatus
FUNCTION GetCenterColor () AS ARGB
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a variable that receives the color of the center point. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

The second overloaded function returns the ARGB color as the result of the function.

#### Remarks

By default, the center point of a **PathGradientBrush** object is the centroid of the brush's boundary path, but you can set the center point to any location, inside or outside the path, by calling the **SetCenterPoint** method of the **PathGradientBrush** object.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object and uses it to fill an ellipse.
' Then the code calls the PathGradientBrush.GetCenterColor method of the PathGradientBrush
' object to obtain the center color.
' ========================================================================================
SUB Example_GetCenterColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM colour AS ARGB
   pthGrBrush.GetCenterColor(@colour)

   ' // Fill a rectangle with the retrieved color.
   DIM solidBrush AS CGpSolidBrush = colour
   graphics.FillRectangle(@solidBrush, 0, 120, 200, 30)

END SUB
' ========================================================================================
```

# <a name="GetCenterPoint"></a>GetCenterPoint (CGpPathGradientBrush)

Gets the center point of this path gradient brush.

```
FUNCTION GetCenterPoint (BYVAL pt AS PointF PTR) AS GpStatus
FUNCTION GetCenterPoint (BYVAL pt AS Point PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pt* | Pointer to a **PointF** or **Point** structure that receives the center point. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

By default, the center point of a **PathGradientBrush** object is at the centroid of the brush's boundary path, but you can set the center point to any location, inside or outside the path, by calling the **SetCenterPoint** method of the **PathGradientBrush** object.

#### Example

```
' ========================================================================================
' The following example demonstrates several methods of the PathGradientBrush class including
' PathGradientBrush.GetCenterPoint and PathGradientBrush.SetCenterColor. The code creates
' a PathGradientBrush object and then sets the brush's center color and boundary color.
' The code calls the PathGradientBrush.GetCenterPoint method to determine the center point
' of the path gradient and then draws a line from the origin to that center point.
' ========================================================================================
SUB Example_GetCenterPoint (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM centerPoint AS GpPointF
   pthGrBrush.GetCenterPoint(@centerPoint)

   ' // Draw a line from the origin to the center of the ellipse.
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   DIM pt AS GpPointF = GDIP_POINTF(0, 0)
   graphics.DrawLine(@pen, @pt, @centerPoint)
   
END SUB
' ========================================================================================
```

# <a name="GetFocusScales"></a>GetFocusScales (CGpPathGradientBrush)

Gets the focus scales of this path gradient brush.

```
FUNCTION GetFocusScales (BYVAL xScale AS SINGLE PTR, BYVAL yScale AS SINGLE PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *xScale* | Pointer to a simple precision that receives the x focus scale value. |
| *yScale* | Pointer to a simple precision that receives the y focus scale value. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

By default, the center color of a path gradient is at the center point. By calling **SetFocusScales**, you can specify that the center color should appear along a path that surrounds the center point. For example, suppose the boundary path is a triangle and the center point is at the centroid of that triangle. Also assume that the boundary color is red and the center color is blue. If you set the focus scales to (0.2, 0.2), the color is blue along the boundary of a small triangle that surrounds the center point. That small triangle is the main boundary path scaled by a factor of 0.2 in the x direction and 0.2 in the y direction. When you paint with the path gradient brush, the color will change gradually from red to blue as you move from the boundary of the large triangle to the boundary of the small triangle. The area inside the small triangle will be filled with blue.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path.
' The code sets the focus scales of the path gradient brush to (0.2, 0.2) and then uses
' the path gradient brush to fill an area that contains the triangular path. Finally, the
' code calls the PathGradientBrush.GetFocusScales method of the PathGradientBrush object
' to obtain the values of the x focus scale and the y focus scale.
' ========================================================================================
SUB Example_GetFocusScales (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}

   ' // No GraphicsPath object is created. The PathGradientBrush
   ' // object is constructed directly from the array of points.
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM colors(0 TO 1) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255)}

   ' // red at the boundary of the outer triangle
   ' // blue at the boundary of the inner triangle
   DIM relativePositions(0 TO 1) AS SINGLE = {0.0, 1.0}
   pthGrBrush.SetInterpolationColors(@colors(0), @relativePositions(0), 2)

   ' // The inner triangle is formed by scaling the outer triangle
   ' // about its centroid. The scaling factor is 0.2 in both the x and y directions.
   pthGrBrush.SetFocusScales(0.2, 0.2)

   ' // Fill a rectangle that is larger than the triangle
   ' // specified in the Point array. The portion of the
   ' // rectangle outside the triangle will not be painted.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 200, 200)

   ' // Obtain information about the path gradient brush.
   DIM xScale AS SINGLE = 0.0
   DIM yScale AS SINGLE = 0.0
   pthGrBrush.GetFocusScales(@xScale, @yScale)

   ' // The value of xScale is now 0.2.
   ' // The value of yScale is now 0.2. 

END SUB
' ========================================================================================
```

# <a name="GetGammaCorrectionLGBrush"></a>GetGammaCorrection (CGpLinearGradientBrush)

Gets the focus scales of this path gradient brush.

```
FUNCTION GetGammaCorrection () AS BOOL
```

#### Return value

If gamma correction is enabled, this method returns TRUE; otherwise, it returns FALSE.

# <a name="GetGammaCorrectionPGBrush"></a>GetGammaCorrection (CGpPathGradientBrush)

Determines whether gamma correction is enabled for this path gradient brush.

```
FUNCTION GetGammaCorrection () AS BOOL
```

#### Return value

If gamma correction is enabled, this method returns TRUE; otherwise, it returns FALSE.

# <a name="GetInterpolationColorCountLGBrush"></a>GetInterpolationColorCount (CGpLinearGradientBrush)

Gets the number of colors currently set to be interpolated for this linear gradient brush.

```
FUNCTION GetInterpolationColorCount () AS INT_
```

#### Return value

This method returns the number of colors to be interpolated for this linear gradient brush. If no colors have been set by using **SetInterpolationColors**, or if invalid positions were passed to **SetInterpolationColors**, then **GetInterpolationColorCount** returns 0.

#### Remarks

A simple linear gradient brush has two colors: a color at the starting boundary and a color at the ending boundary. When you paint with such a brush, the color changes gradually from the starting color to the ending color as you move from the starting boundary to the ending boundary. You can create a more complex gradient by using the **SetInterpolationColors** method to specify an array of colors and their corresponding blend positions to be interpolated for this linear gradient brush.

You can obtain the colors and blend positions currently set for a linear gradient brush by calling its **GetInterpolationColors** method. Before you call the **GetInterpolationColors** method, you must allocate two buffers: one to hold the array of colors and one to hold the array of blend positions. You can call the **GetInterpolationColorCount** method to determine the required size of those buffers. The size of the colors buffer is the return value of **GetInterpolationColorCount** multiplied by **sizeof(Color)**. The size of the blend positions buffer is the value of **GetInterpolationColorCount** multiplied by **sizeof(REAL)**.

#### Example

```
' ========================================================================================
' The following example sets the colors that are interpolated for this linear gradient
' brush to red, blue, and green and sets the blend positions to 0, 0.3, and 1. The code
' calls the LinearGradientBrush::GetInterpolationColorCount method of a LinearGradientBrush
' object to obtain the number of colors currently set to be interpolated for the brush.
' Next, the code gets the colors and their positions. Then, the code fills a small
' rectangle with each color.
' ========================================================================================
SUB Example_GetInterpolationColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush, and set the colors to be interpolated.
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 0)}
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.3, 1.0}

   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(100, 0)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 255, 255, 255))
   linGrBrush.SetInterpolationColors(@colors(0), @positions(0), 3)

   ' // Obtain information about the linear gradient brush.
   ' // How many colors have been specified to be interpolated for this brush?
   DIM colorCount AS LONG = linGrBrush.GetInterpolationColorCount

   ' // Allocate a buffer large enough to hold the set of colors.
   DIM rgcolors(0 TO colorCount - 1) AS ARGB

   ' // Allocate a buffer to hold the relative positions of the colors.
   DIM rgPositions(0 TO colorCount - 1) AS SINGLE

   ' // Get the colors and their relative positions.
   linGrBrush.GetInterpolationColors(@rgcolors(0), @rgPositions(0), colorCount)

   ' // Fill a small rectangle with each of the colors.
   DIM pSolidBrush AS CGpSolidBrush PTR
   FOR j AS LONG = 0 TO colorCount - 1
      pSolidBrush = NEW CGpSolidBrush(rgcolors(j))
      graphics.FillRectangle(pSolidBrush, 15 * j, 0, 10, 10)
      Delete pSolidBrush
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetInterpolationColorCountPGBrush"></a>GetInterpolationColorCount (CGpPathGradientBrush)

Gets the number of preset colors currently specified for this path gradient brush.

```
FUNCTION GetInterpolationColorCount () AS INT_
```

#### Return value

This method returns the number of preset colors currently specified for this path gradient brush.

#### Remarks

Remarks

A simple path gradient brush has two colors: a boundary color and a center color. When you paint with such a brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point. You can create a more complex gradient by specifying an array of preset colors and an array of blend positions.

You can obtain the interpolation colors and interpolation positions currently set for a **PathGradientBrush** object by calling the **GetInterpolationColors** method of that **PathGradientBrush** object. Before you call the **GetInterpolationColors** method, you must allocate two buffers: one to hold the array of interpolation colors and one to hold the array of interpolation positions. You can call the **GetInterpolationColorCount** method of the **PathGradientBrush** object to determine the required size of those buffers. The size of the color buffer is the return value of **GetInterpolationColorCount** multiplied by 4. The size of the position buffer is the value of **GetInterpolationColorCount** multiplied by 4 (the size of a simple precision number).

# <a name="GetInterpolationColorsLGBrush"></a>GetInterpolationColors (CGpLinearGradientBrush)

Gets the blend factors and their corresponding blend positions from a **LinearGradientBrush** object.

```
FUNCTION GetInterpolationColors (BYVAL presetColors AS ARGB PTR, _
   BYVAL blendPositions AS SINGLE PTR, BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *presetColors* | Pointer to an array that receives the colors. A color of a given index in the presetColors array corresponds to the blend position of that same index in the *blendPositions* array. |
| *blendPositions* | Pointer to an array that receives the blend positions. Each number in the array indicates a percentage of the distance between the starting boundary and the ending boundary and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the gradient and 1.0 indicates the ending boundary. A blend position between 0.0 and 1.0 indicates a line, parallel to the boundary lines, that is a certain fraction of the distance from the starting boundary to the ending boundary. For example, a blend position of 0.7 indicates the line that is 70 percent of the distance from the starting boundary to the ending boundary. The color is constant on lines that are parallel to the boundary lines. |
| *count* | Integer that specifies the number of elements in the presetColors array. This is the same as the number of elements in the blendPositions array. Before calling the **GetInterpolationColors** method of a **LinearGradientBrush** object, call the **GetInterpolationColorCount** method of that same **LinearGradientBrush** object to determine the current number of colors. The number of blend positions retrieved is the same as the number of colors retrieved. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example sets the colors that are interpolated for this linear gradient
' brush to red, blue, and green and sets the blend positions to 0, 0.3, and 1. The code
' calls the LinearGradientBrush::GetInterpolationColorCount method of a LinearGradientBrush
' object to obtain the number of colors currently set to be interpolated for the brush.
' Next, the code gets the colors and their positions. Then, the code fills a small
' rectangle with each color.
' ========================================================================================
SUB Example_GetInterpolationColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush, and set the colors to be interpolated.
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 0)}
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.3, 1.0}

   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(100, 0)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 255, 255, 255))
   linGrBrush.SetInterpolationColors(@colors(0), @positions(0), 3)

   ' // Obtain information about the linear gradient brush.
   ' // How many colors have been specified to be interpolated for this brush?
   DIM colorCount AS LONG = linGrBrush.GetInterpolationColorCount

   ' // Allocate a buffer large enough to hold the set of colors.
   DIM rgcolors(0 TO colorCount - 1) AS ARGB

   ' // Allocate a buffer to hold the relative positions of the colors.
   DIM rgPositions(0 TO colorCount - 1) AS SINGLE

   ' // Get the colors and their relative positions.
   linGrBrush.GetInterpolationColors(@rgcolors(0), @rgPositions(0), colorCount)

   ' // Fill a small rectangle with each of the colors.
   DIM pSolidBrush AS CGpSolidBrush PTR
   FOR j AS LONG = 0 TO colorCount - 1
      pSolidBrush = NEW CGpSolidBrush(rgcolors(j))
      graphics.FillRectangle(pSolidBrush, 15 * j, 0, 10, 10)
      Delete pSolidBrush
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetInterpolationColorsPGBrush"></a>GetInterpolationColors (CGpPathGradientBrush)

Gets preset colors and blend positions currently specified for this path gradient brush.

```
FUNCTION GetInterpolationColors (BYVAL presetColors AS ARGB PTR, _
   BYVAL blendPositions AS SINGLE PTR, BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *presetColors* | Pointer to an array that receives the colors. A color of a given index in the presetColors array corresponds to the blend position of that same index in the *blendPositions* array. |
| *blendPositions* | Pointer to an array that receives the blend positions. Each blend position is a number from 0 through 1, where 0 indicates the boundary of the gradient and 1 indicates the center point. A blend position between 0 and 1 indicates the set of all points that are a certain fraction of the distance from the boundary to the center point. For example, a blend position of 0.7 indicates the set of all points that are 70 percent of the way from the boundary to the center point. |
| *count* | Integer that specifies the number of elements in the *presetColorsarray*. This is the same as the number of elements in the *blendPositions* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A simple path gradient brush has two colors: a boundary color and a center color. When you paint with such a brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point. You can create a more complex gradient by specifying an array of preset colors and an array of blend positions.

Before you call the GetInterpolationColors method, you must allocate two buffers: one to hold the array of preset colors and one to hold the array of blend positions. You can call the **GetInterpolationColorCount** method of the **PathGradientBrush** object to determine the required size of those buffers. The size of the color buffer is the return value of **GetInterpolationColorCount** multiplied by 4. The size of the position buffer is the value of **GetInterpolationColorCount** multiplied by 4 (the size of a simple precision number).

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object from a triangular path. The code
' sets the preset colors to red, blue, and aqua and sets the blend positions to 0, 0.6, and 1.
' The code calls the PathGradientBrush.GetInterpolationColorCount method of the PathGradientBrush
' object to obtain the number of preset colors currently set for the brush. Next, the code
' allocates two buffers: one to hold the array of preset colors, and one to hold the array
' of blend positions. The call to the PathGradientBrush.GetInterpolationColors method of
' the PathGradientBrush object fills the buffers with the preset colors and the blend positions.
' Finally the code fills a small square with each of the preset colors.
' ========================================================================================
SUB Example_GetInterpolationColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 255)}
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.6, 1.0}

   pthGrBrush.SetInterpolationColors(@colors(0), @positions(0), 3)

   ' // Obtain information about the path gradient brush.
   DIM colorCount AS LONG = pthGrBrush.GetInterpolationColorCount
   DIM rgColors(colorCount - 1) AS ARGB
   DIM rgPositions(colorCount - 1) AS SINGLE
   pthGrBrush.GetInterpolationColors(@rgColors(0), @rgPositions(0), colorCount)

   ' // Fill a small square with each of the interpolation colors.
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 255, 255)

   FOR j AS LONG = 0 TO colorCount - 1
      solidBrush.SetColor(rgColors(j))
      graphics.FillRectangle(@solidBrush, 15 * j, 0, 10, 10)
   NEXT

END SUB
' ========================================================================================
```
