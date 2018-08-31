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
| [GetBlendCount](#GetBlendCount) | Gets the number of blend factors currently set. |
| [GetGammaCorrection](#GetGammaCorrection) | Determines whether gamma correction is enabled for this brush. |
| [GetInterpolationColorCount](#GetInterpolationColorCount) | Gets the number of colors currently set to be interpolated. |
| [GetInterpolationColors](#GetInterpolationColors) | Gets the blend factors and their corresponding blend positions. |
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
| [GetBlend](#GetBlend) | Gets the blend factors and their corresponding blend positions. |
| [GetBlendCount](#GetBlendCount) | Gets the number of blend factors currently set. |
| [GetCenterColor](#GetCenterColor) | Gets center cpñpr of the brush. |
| [GetCenterPoint](#GetCenterPoint) | Gets the center point of the brush. |
| [GetFocusScales](#GetFocusScales) | Gets the focus scales of the brush. |
| [GetGammaCorrection](#GetGammaCorrection) | Determines whether gamma correction is enabled for this brush. |
| [GetInterpolationColorCount](#GetInterpolationColorCount) | Gets the number of preset colors currently specified for this brush. |
| [GetInterpolationColors](#GetInterpolationColors) | Gets preset colors and blend positions currently specified for this brush. |
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

# <a name="GetBlendÑGBrush"></a>GetBlend (CGpLinearGradientBrush)

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
