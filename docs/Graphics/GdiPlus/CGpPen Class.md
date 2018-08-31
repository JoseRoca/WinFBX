# CGpPen Class

Encapsulates a **Pen** object. A **Pen** object is a Windows GDI+ object used to draw lines and curves.

**Inherits from**: CGpBase.
**Imclude file**: CGpPen.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Create a Pen object. |
| [Clone](#Clone) | Copies the contents of the existing Pen object into a new Pen object. |
| [GetAlignment](#GetAlignment) | Gets the alignment currently set for this Pen object. |
| [GetBrush](#GetBrush) | Gets the the Brush object that is currently set for this Pen object. |
| [GetColor](#GetColor) | Gets the color currently set for this Pen object. |
| [GetCompoundArray](#GetCompoundArray) | Gets the the compound array currently set for this Pen object. |
| [GetCompoundArrayCount](#GetCompoundArrayCount) | Gets the number of elements in a compound array. |
| [GetCustomEndCap](#GetCustomEndCap) | Gets the custom end cap currently set for this Pen object. |
| [GetCustomStartCap](#GetCustomStartCap) | Gets the custom end cap currently set for this Pen object. |
| [GetDashCap](#GetDashCap) | Gets the dash cap style currently set for this Pen object. |
| [GetDashOffset](#GetDashOffset) | Gets the distance from the start of the line to the start of the first space in a dashed line. |
| [GetDashPattern](#GetDashPattern) | Gets an array of custom dashes and spaces currently set for this Pen object. |
| [GetDashPatternCount](#GetDashPatternCount) | Gets the number of elements in a dash pattern array. |
| [GetDashStyle](#GetDashStyle) | Gets the dash style currently set for this Pen object. |
| [GetEndCap](#GetEndCap) | Gets the end cap currently set for this Pen object. |
| [GetEndCap](#GetEndCap) | Gets the end cap currently set for this Pen object. |
| [GetLineJoin](#GetLineJoin) | Gets the line join style currently set for this Pen object. |
| [GetMiterLimit](#GetMiterLimit) | Gets the miter length currently set for this Pen object. |
| [GetPenType](#GetPenType) | Gets the type currently set for this Pen object. |
| [GetStartCap](#GetStartCap) | Gets the start cap currently set for this Pen object. |
| [GetTransform](#GetTransform) | Gets the world transformation matrix currently set for this Pen object. |
| [GetWidth](#GetWidth) | Gets the width currently set for this Pen object. |
| [MultiplyTransform](#MultiplyTransform) | Updates the world transformation matrix of this Pen object with the product of itself and another matrix. |
| [ResetTransform](#ResetTransform) | Sets the world transformation matrix of this Pen object to the identity matrix. |
| [RotateTransform](#RotateTransform) | Updates the world transformation matrix of this Pen object with the product of itself and a rotation matrix. |
| [ScaleTransform](#ScaleTransform) | Sets the Pen object's world transformation matrix equal to the product of itself and a scaling matrix. |
| [SetAlignment](#SetAlignment) | Sets the alignment for this Pen object relative to the line. |
| [SetBrush](#SetBrush) | Sets the Brush object that a pen uses to fill a line. |
| [SetColor](#SetColor) | Sets the color for this Pen object. |
| [SetCompoundArray](#SetCompoundArray) | Sets the compound array for this Pen object. |
| [SetCustomEndCap](#SetCustomEndCap) | Sets the custom end cap for this Pen object. |
| [SetCustomStartCap](#SetCustomStartCap) | Sets the custom start cap for this Pen object. |
| [SetDashCap](#SetDashCap) | Sets the dash cap style for this Pen object. |
| [SetDashOffset](#SetDashOffset) | Sets the distance from the start of the line to the start of the first dash in a dashed line. |
| [SetDashPattern](#SetDashPattern) | Sets an array of custom dashes and spaces for this Pen object. |
| [SetDashStyle](#SetDashStyle) | Sets the dash style for this Pen object. |
| [SetEndCap](#SetEndCap) | Sets the end cap for this Pen object. |
| [SetLineCap](#SetLineCap) | Sets the cap styles for the start, end, and dashes in a line drawn with this pen. |
| [SetLineJoin](#SetLineJoin) | Sets the line join for this Pen object. |
| [SetMiterLimit](#SetMiterLimit) | Sets the miter limit of this Pen object. |
| [SetStartCap](#SetStartCap) | Sets the start cap for this Pen object. |
| [SetTransform](#SetTransform) | Sets the world transformation of this Pen object. |
| [SetTransform](#SetTransform) | Sets the world transformation of this Pen object. |
| [SetWidth](#SetWidth) | Sets the width for this Pen object. |
| [TranslateTransform](#TranslateTransform) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |
