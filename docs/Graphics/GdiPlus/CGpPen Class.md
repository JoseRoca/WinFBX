# CGpPen Class

Encapsulates a **Pen** object. A **Pen** object is a Windows GDI+ object used to draw lines and curves.

**Inherits from**: CGpBase.<br>
**Include file**: CGpPen.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Creates a **Pen** object. |
| [Clone](#Clone) | Copies the contents of the existing **Pen** object into a new **Pen** object. |
| [GetAlignment](#GetAlignment) | Gets the alignment currently set for this **Pen** object. |
| [GetBrush](#GetBrush) | Gets the the **Brush** object that is currently set for this **Pen** object. |
| [GetColor](#GetColor) | Gets the color currently set for this **Pen** object. |
| [GetCompoundArray](#GetCompoundArray) | Gets the the compound array currently set for this **Pen** object. |
| [GetCompoundArrayCount](#GetCompoundArrayCount) | Gets the number of elements in a compound array. |
| [GetCustomEndCap](#GetCustomEndCap) | Gets the custom end cap currently set for this **Pen** object. |
| [GetCustomStartCap](#GetCustomStartCap) | Gets the custom end cap currently set for this **Pen** object. |
| [GetDashCap](#GetDashCap) | Gets the dash cap style currently set for this **Pen** object. |
| [GetDashOffset](#GetDashOffset) | Gets the distance from the start of the line to the start of the first space in a dashed line. |
| [GetDashPattern](#GetDashPattern) | Gets an array of custom dashes and spaces currently set for this **Pen** object. |
| [GetDashPatternCount](#GetDashPatternCount) | Gets the number of elements in a dash pattern array. |
| [GetDashStyle](#GetDashStyle) | Gets the dash style currently set for this **Pen** object. |
| [GetEndCap](#GetEndCap) | Gets the end cap currently set for this **Pen** object. |
| [GetLineJoin](#GetLineJoin) | Gets the line join style currently set for this **Pen** object. |
| [GetMiterLimit](#GetMiterLimit) | Gets the miter length currently set for this **Pen** object. |
| [GetPenType](#GetPenType) | Gets the type currently set for this **Pen** object. |
| [GetStartCap](#GetStartCap) | Gets the start cap currently set for this **Pen** object. |
| [GetTransform](#GetTransform) | Gets the world transformation matrix currently set for this **Pen** object. |
| [GetWidth](#GetWidth) | Gets the width currently set for this **Pen** object. |
| [MultiplyTransform](#MultiplyTransform) | Updates the world transformation matrix of this **Pen** object with the product of itself and another matrix. |
| [ResetTransform](#ResetTransform) | Sets the world transformation matrix of this **Pen** object to the identity matrix. |
| [RotateTransform](#RotateTransform) | Updates the world transformation matrix of this **Pen** object with the product of itself and a rotation matrix. |
| [ScaleTransform](#ScaleTransform) | Sets the **Pen** object's world transformation matrix equal to the product of itself and a scaling matrix. |
| [SetAlignment](#SetAlignment) | Sets the alignment for this **Pen** object relative to the line. |
| [SetBrush](#SetBrush) | Sets the **Brush** object that a pen uses to fill a line. |
| [SetColor](#SetColor) | Sets the color for this **Pen** object. |
| [SetCompoundArray](#SetCompoundArray) | Sets the compound array for this **Pen** object. |
| [SetCustomEndCap](#SetCustomEndCap) | Sets the custom end cap for this **Pen** object. |
| [SetCustomStartCap](#SetCustomStartCap) | Sets the custom start cap for this **Pen** object. |
| [SetDashCap](#SetDashCap) | Sets the dash cap style for this **Pen** object. |
| [SetDashOffset](#SetDashOffset) | Sets the distance from the start of the line to the start of the first dash in a dashed line. |
| [SetDashPattern](#SetDashPattern) | Sets an array of custom dashes and spaces for this **Pen** object. |
| [SetDashStyle](#SetDashStyle) | Sets the dash style for this **Pen** object. |
| [SetEndCap](#SetEndCap) | Sets the end cap for this **Pen** object. |
| [SetLineCap](#SetLineCap) | Sets the cap styles for the start, end, and dashes in a line drawn with this pen. |
| [SetLineJoin](#SetLineJoin) | Sets the line join for this **Pen** object. |
| [SetMiterLimit](#SetMiterLimit) | Sets the miter limit of this **Pen** object. |
| [SetStartCap](#SetStartCap) | Sets the start cap for this **Pen** object. |
| [SetTransform](#SetTransform) | Sets the world transformation of this **Pen** object. |
| [SetTransform](#SetTransform) | Sets the world transformation of this **Pen** object. |
| [SetWidth](#SetWidth) | Sets the width for this **Pen** object. |
| [TranslateTransform](#TranslateTransform) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |

# <a name="Constructors"></a>Constructors

Creates a **Pen** object that uses a specified color and width.

```
CONSTRUCTOR CGpPen (BYVAL pBrush AS CGpBrush PTR, BYVAL nWidth AS SINGLE = 1.0)
CONSTRUCTOR CGpPen (BYVAL colour AS ARGB, BYVAL nWidth AS SINGLE = 1.0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a brush to base this pen on. |
| *colour* | An ARGB color that specifies the color for this **Pen** object. |
| *nWidth* | The width of this pen's stroke. The default value is 1.0. |

#### Remarks

If you pass the address of a pen to one of the draw methods of a **Graphics** object, the width of the pen's stroke is dependent on the unit of measure specified in the **Graphics** object. The default unit of measure is **UnitPixel**, which is an element of the **GpUnit** enumeration.

# <a name="Clone"></a>Clone

Copies the contents of the existing **Pen** object into a new **Pen** object.

```
FUNCTION Clone (BYVAL pClonePen AS CGpPen PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pClonePen* | Pointer to the **Pen** object where to copy the contents of the existing object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, creates a copy of the Pen object, and then
' draws an ellipse using the copied Pen object.
' ========================================================================================
SUB Example_ClonePen (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create and clone a Pen object.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 4)
   DIM clonedPen AS CGpPen
   pen.Clone(@clonedPen)

   ' // Draw a rectangle using the cloned Pen object.
   graphics.DrawRectangle(@clonedPen, 10, 10, 100, 50)

END SUB
' ========================================================================================
```

# <a name="GetAlignment"></a>GetAlignment

Gets the alignment currently set for this **Pen** object.

```
FUNCTION GetAlignment () AS PenAlignment
```

#### Return value

This method returns an element of the **PenAlignment** enumeration that indicates the current alignment setting for this **Pen** object. The default value of **PenAlignment** is **PenAlignmentCenter**.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the alignment, draws a line, and then
' gets the pen alignment settings.
' ========================================================================================
SUB Example_GetAlignment (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object and set its alignment.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 15)
   pen.SetAlignment(PenAlignmentCenter)
   
   ' // Draw a line.
   graphics.DrawLine(@pen, 0, 0, 100, 50)

   ' // Obtain information about the Pen object.
   DIM nPenAlignment AS PenAlignment = pen.GetAlignment

   IF nPenAlignment = PenAlignmentCenter THEN
      ' // The pixels will be centered on the theoretical line.
   ELSEIF nPenAlignment = PenAlignmentInset THEN
      '  // The pixels will lie inside the filled area  of the theoretical line.
   END IF

END SUB
' ========================================================================================
```

# <a name="GetBrush"></a>GetBrush

Gets the the **Brush** object that is currently set for this **Pen** object.

```
FUNCTION GetBrush (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
```

#### Return value

This method returns a pointer to a **Brush** object that is currently used to fill a line.

#### Example

```
' ========================================================================================
' The following example creates a Brush object, creates a Pen object based on the Brush
' object, draws a line with the pen, gets the Brush from the pen, and then uses the Brush
' to fill a rectangle.
' ========================================================================================
SUB Example_GetBrush (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a HatchBrush object
   DIM hatchBrush AS CGpHatchBrush = CGpHatchBrush(HatchStyleVertical, GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   ' // Create a pen, and set the brush for the pen
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 10)
   pen.SetBrush(@hatchBrush)

   ' // Draw a line with the pen
   graphics.DrawLine(@pen, 0, 0, 200, 100)

   ' // Get the pen's brush, and use that brush to fill a rectangle.
   DIM pBrush AS CGpBrush
   pen.GetBrush(@pBrush)
   graphics.FillRectangle(@pBrush, 0, 100, 200, 100)

END SUB
' ========================================================================================
```

# <a name="GetColor"></a>GetColor

Gets the color currently set for this **Pen** object.

```
FUNCTION GetColor (BYVAL colour AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a variable that receives the color of this **Pen** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object and draws a line. The code then gets the color
' of the pen and creates a Brush object based on that color. Finally, the code uses the
' Brush object to fill a rectangle.
' ========================================================================================
SUB Example_GetColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen, and use it to draw a line.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 200, 150, 100), 5)
   graphics.DrawLine(@pen, 0, 0, 200, 100)

   ' // Get the pen's color, and use that color to create a brush.
   DIM colour AS ARGB
   pen.GetColor(@colour)
   DIM solidBrush AS CGpSolidBrush = colour

   ' // Use the brush to fill a rectangle.
   graphics.FillRectangle(@solidBrush, 0, 100, 200, 100)

END SUB
' ========================================================================================
```

# <a name="GetCompoundArray"></a>GetCompoundArray

Gets the the compound array currently set for this **Pen** object.

```
FUNCTION GetCompoundArray (BYVAL compoundArray AS SINGLE PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *compoundArray* | Pointer to an array that receives the compound array. |
| *count* | Integer that specifies the number of elements in the *compoundArray* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example declares an array, sets the compound array, draws a line, and gets
' the number of elements in the compound array.
' ========================================================================================
SUB Example_GetCompoundArray (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create an array of real numbers and a Pen object.
   DIM compVals(0 TO 5) AS SINGLE = {0.0, 0.2, 0.5, 0.7, 0.9, 1.0}
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 30)

   ' // Set the compound array of the pen.
   pen.SetCompoundArray(@compVals(0), 6)

   ' // Draw a line with the pen.
   graphics.DrawLine(@pen, 5, 20, 405, 200)

   ' // Obtain information about the pen
   DIM compValues(ANY) AS SINGLE
   DIM nCount AS LONG = pen.GetCompoundArrayCount
   REDIM compValues(nCount -1)
   pen.GetCompoundArray(@compValues(0), nCount)

   FOR j AS LONG = 0 TO nCount - 1
      ' // Inspect or use the value in compValues(j)
      PRINT compValues(j)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetCompoundArrayCount"></a>GetCompoundArrayCount

Gets the number of elements in a compound array.

```
FUNCTION GetCompoundArrayCount () AS GpStatus
```

#### Example

```
' ========================================================================================
' The following example declares an array, sets the compound array, draws a line, and gets
' the number of elements in the compound array.
' ========================================================================================
SUB Example_GetCompoundArray (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create an array of real numbers and a Pen object.
   DIM compVals(0 TO 5) AS SINGLE = {0.0, 0.2, 0.5, 0.7, 0.9, 1.0}
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 30)

   ' // Set the compound array of the pen.
   pen.SetCompoundArray(@compVals(0), 6)

   ' // Draw a line with the pen.
   graphics.DrawLine(@pen, 5, 20, 405, 200)

   ' // Obtain information about the pen
   DIM compValues(ANY) AS SINGLE
   DIM nCount AS LONG = pen.GetCompoundArrayCount
   REDIM compValues(nCount -1)
   pen.GetCompoundArray(@compValues(0), nCount)

   FOR j AS LONG = 0 TO nCount - 1
      ' // Inspect or use the value in compValues(j)
      PRINT compValues(j)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetCustomEndCap"></a>GetCustomEndCap

Gets the custom end cap currently set for this **Pen** object.

```
FUNCTION GetCustomEndCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCustomLineCap* | Pointer to a **CustomLineCap** object that receives the custom end cap of this **Pen** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a rectangle to it. The code
' then creates a Pen object, sets the custom end cap using the GraphicsPath object, and
' draws a line. Finally, the code gets the custom end cap of the pen and creates another
' Pen object using the same custom end cap. It then draws a second line.
' ========================================================================================
SUB Example_GetCustomEndCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GraphicsPath object, and add a rectangle to it.
   DIM pStrokePath AS CGpGraphicsPath
   pStrokePath.AddRectangle(-10, -5, 20, 10)

   ' // Create a pen, and set the custom end cap based on the GraphicsPath object.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255))
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @pStrokePath)
   pen.SetCustomEndCap(@custCap)

   ' // Draw a line with the custom end cap.
   graphics.DrawLine(@pen, 0, 0, 200, 100)

   ' // Obtain the custom end cap for the pen.
   DIM customLineCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, NULL)
   pen.GetCustomEndCap(@customLineCap)

   ' // Create another pen, and use the same custom end cap.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 3)
   pen2.SetCustomEndCap(@customLineCap)

   ' // Draw a line using the second pen.
   graphics.DrawLine(@pen2, 0, 100, 200, 150)

END SUB
' ========================================================================================
```

# <a name="GetCustomStartCap"></a>GetCustomStartCap

Gets the custom end cap currently set for this **Pen** object.

```
FUNCTION GetCustomStartCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCustomLineCap* | Pointer to a **CustomLineCap** object that receives the custom start cap of this **Pen** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a rectangle to it. The code
' then creates a Pen object, sets the custom start cap using the GraphicsPath object, and
' draws a line. Finally, the code gets the custom start cap of the pen and creates another
' Pen object using the same custom end cap. It then draws a second line.
' ========================================================================================
SUB Example_GetCustomStartCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GraphicsPath object, and add a rectangle to it.
   DIM pStrokePath AS CGpGraphicsPath
   pStrokePath.AddRectangle(-10, -5, 20, 10)

   ' // Create a pen, and set the custom start cap based on the GraphicsPath object.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255))
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @pStrokePath)
   pen.SetCustomStartCap(@custCap)

   ' // Draw a line with the custom start cap.
   graphics.DrawLine(@pen, 20, 20, 200, 100)

   ' // Obtain the custom start cap for the pen.
   DIM customLineCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, NULL)
   pen.GetCustomStartCap(@customLineCap)

   ' // Create another pen, and use the same custom end cap.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 3)
   pen2.SetCustomStartCap(@customLineCap)

   ' // Draw a line using the second pen.
   graphics.DrawLine(@pen2, 50, 100, 200, 150)

END SUB
' ========================================================================================
```

# <a name="GetDashCap"></a>GetDashCap

Gets the dash cap style currently set for this **Pen** object.

```
FUNCTION GetDashCap () AS DashCap
```

#### Return value

This method returns an element of the **DashCap** enumeration that indicates the dash cap style currently set for this Pen object.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash cap style, and draws a line.
' The code then gets the dash cap style of the pen and creates a second Pen object with
' the same dash cap style. Finally, the code draws a second line using the second pen.
' ========================================================================================
SUB Example_GetDashCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 20)

   ' // Set the dash style for the pen
   pen.SetDashStyle(DashStyleDash)

   ' // Set a rounded dash cap for the pen
   pen.SetDashCap(DashCapRound)

   ' // Draw a line using the pen
   graphics.DrawLine(@pen, 50, 50, 400, 200)

   ' // Obtain the dash cap for the pen
   DIM nDashCap AS DashCap = pen.GetDashCap

   ' // Create another pen, and use the same dash cap.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 15)
   pen2.SetDashStyle(DashStyleDash)
   pen2.SetDashCap(nDashCap)

   ' // Draw a second line.
   graphics.DrawLine(@pen2, 50, 100, 400, 300)

END SUB
' ========================================================================================
```

# <a name="GetDashOffset"></a>GetDashOffset

Gets the distance from the start of the line to the start of the first space in a dashed line.

```
FUNCTION GetDashOffset () AS SINGLE
```

#### Return value

This method returns a real number that indicates the distance from the start of the line to the start of the dashes.

# <a name="GetDashPattern"></a>GetDashPattern

Gets an array of custom dashes and spaces currently set for this **Pen** object.

```
FUNCTION GetDashPattern (BYVAL dashArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dashArray* | Pointer to an array that receives the length of the dashes and spaces in a custom dashed line. |
| *nCount* | Integer that specifies the number of elements in the *dashArray* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Return value

This method returns a real number that indicates the distance from the start of the line to the start of the dashes.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash pattern array, and draws a
' custom dashed line. The code then gets the number of elements in the custom dashed array.
' ========================================================================================
SUB Example_GetDashPattern (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create and set an array of real numbers.
   DIM dashVals(0 TO 3) AS SINGLE = {5.0, 2.0, 15.0, 4.0}

   ' // Create a Pen
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 5)

   ' // Set the dash pattern for the custom dashed line.
   pen.SetDashPattern(@dashVals(0), 4)

   ' // Draw a line using the pen.
   graphics.DrawLine(@pen, 5, 20, 405, 200)

   ' // Obtain information about the pen.
   DIM nCount AS INT_
   nCount = pen.GetDashPatternCount
   DIM dashValues(nCount - 1) AS SINGLE
   pen.GetDashPattern(@dashValues(0), nCount)

   FOR j AS LONG = 0 TO nCount - 1
      OutputDebugStringW WSTR(dashValues(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetDashPatternCount"></a>GetDashPatternCount

Gets the number of elements in a dash pattern array.

```
FUNCTION GetDashPatternCount () AS INT_
```

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash pattern array, and draws a
' custom dashed line. The code then gets the number of elements in the custom dashed array.
' ========================================================================================
SUB Example_GetDashPattern (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create and set an array of real numbers.
   DIM dashVals(0 TO 3) AS SINGLE = {5.0, 2.0, 15.0, 4.0}

   ' // Create a Pen
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 5)

   ' // Set the dash pattern for the custom dashed line.
   pen.SetDashPattern(@dashVals(0), 4)

   ' // Draw a line using the pen.
   graphics.DrawLine(@pen, 5, 20, 405, 200)

   ' // Obtain information about the pen.
   DIM nCount AS INT_
   nCount = pen.GetDashPatternCount
   DIM dashValues(nCount - 1) AS SINGLE
   pen.GetDashPattern(@dashValues(0), nCount)

   FOR j AS LONG = 0 TO nCount - 1
      OutputDebugStringW WSTR(dashValues(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetDashStyle"></a>GetDashStyle

Gets the dash style currently set for this **Pen** object.

```
FUNCTION GetDashStyle () AS DashStyle
```

#### Return value

This method returns an element of the DashStyle enumeration that indicates the dash style for this **Pen** object.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash style, and draws a dashed line.
' The code then gets the dash style, creates a second pen with the dash style of the first
' pen, and draws a second dashed line.
' ========================================================================================
SUB Example_GetDashStyle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 9)

   ' // Set the dash style for the pen
   pen.SetDashStyle(DashStyleDashDot)
   graphics.DrawLine(@pen, 20, 20, 200, 100)

   ' // Obtain the dash style for the pen.
   DIM nDashStyle AS DashStyle
   nDashStyle = pen.GetDashStyle

   ' // Create another pen, and use the same dash style.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 9)
   pen2.SetDashStyle(nDashStyle)

   ' // Draw a second dashed line.
   graphics.DrawLine(@pen2, 20, 60, 200, 140)

END SUB
' ========================================================================================
```

# <a name="GetEndCap"></a>GetEndCap

Gets the end cap currently set for this **Pen** object.

```
FUNCTION GetEndCap () AS LineCap
```

#### Return value

This method returns an element of the **LineCap** enumeration that indicates the end cap for this **Pen** object.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the end cap, and draws a line. The code
' then resets the end cap, draws a second line, resets the end cap again, and draws a third line.
' ========================================================================================
SUB Example_GetEndCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 15)

   ' // Set the end cap of the pen, and draw a line.
   pen.SetEndCap(LineCapArrowAnchor)
   graphics.DrawLine(@pen, 20, 20, 200, 100)

   ' // Obtain the end cap for the pen.
   DIM nLineCap AS LineCap
   nLineCap = pen.GetEndCap

   ' // Create another pen, and use the same end cap.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 9)
   pen2.SetEndCap(nLineCap)

   ' // Draw a second line.
   graphics.DrawLine(@pen2, 20, 60, 200, 140)

END SUB
' ========================================================================================
```
