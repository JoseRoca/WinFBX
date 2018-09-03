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

# <a name="Constructors"></a>Constructors (CGpPen)

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

# <a name="Clone"></a>Clone (CGpPen)

Copies the contents of the existing **Pen** object into a new **Pen** object.

```
FUNCTION Clone (BYVAL pClonePen AS CGpPen PTR) AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

| Parameter  | Description |
| ---------- | ----------- |
| *pClonePen* | Pointer to the **Pen** object where to copy the contents of the existing object. |

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

# <a name="GetAlignment"></a>GetAlignment (CGpPen)

Gets the alignment currently set for this **Pen** object.

```
FUNCTION GetAlignment () AS PenAlignment
```

#### Return value

This method returns an element of the PenAlignment enumeration that indicates the current alignment setting for this **Pen** object. The default value of **PenAlignment** is **PenAlignmentCenter**.

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
