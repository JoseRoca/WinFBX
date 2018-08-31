# CGpCustomeLineCap Class

The **CGpCustomLineCap** class encapsulates a custom line cap. A line cap defines the style of graphic used to draw the ends of a line. It can be various shapes, such as a square, circle, or diamond. A custom line cap is defined by the path that draws it. The path is drawn by using a **Pen** object to draw the outline of a shape or by using a **Brush** object to fill the interior. The cap can be used on either or both ends of the line. Spacing can be adjusted between the end caps and the line.

**Inherits from**: CGpBase.
**Imclude file**: CGpLineCaps.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#ConstructorCustomLineCap) | Creates a CustomLineCap object. |
| [Clone](#Clone) | Copies the contents of the existing CustomLineCap object into a new CustomLineCap object. |
| [GetBaseCap](#GetBaseCap) | Gets the style of the base cap. |
| [GetBaseInset](#GetBaseInset) | Gets the distance between the base cap to the start of the line. |
| [GetStrokeCaps](#GetStrokeCaps) | Gets the end cap styles for both the start line cap and the end line cap. |
| [GetStrokeJoin](#GetStrokeJoin) | Returns the style of LineJoin used to join multiple lines in the same GraphicsPath object. |
| [GetWidthScale](#GetWidthScale) | Gets the value of the scale width. |
| [SetBaseCap](#SetBaseCap) | Sets the LineCap that appears as part of this CustomLineCap at the end of a line. |
| [SetBaseInset](#SetBaseInset) | Sets the base inset value of this custom line cap. |
| [SetStrokeCap](#SetStrokeCap) | Sets the LineCap object used to start and end lines within the GraphicsPath object that defines this CustomLineCap object. |
| [SetStrokeCaps](#SetStrokeCaps) | Sets the LineCap objects used to start and end lines within the GraphicsPath object that defines this CustomLineCap object. |
| [SetStrokeJoin](#SetStrokeJoin) | Sets the style of line join for the stroke. |
| [SetWidthScale](#SetWidthScale) | Sets the value of the scale width.  |

# CGpAdjustableArrowCap Class

The **CGpAdjustableArrowCap** object extends **CGpCustomLineCap**. This object builds a line cap that looks like an arrow.

**Inherits from**: CGpCustomLineCap.
**Imclude file**: CGpLineCaps.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#ConstructorArrowCap) | Creates an adjustable arrow line cap with the specified height and width. |
| [GetHeight](#GetHeight) | Gets the height of the arrow cap. |
| [GetMiddleInset](#GetMiddleInset) | Gets the value of the inset. |
| [GetWidth](#GetWidth) | Gets the width of the arrow cap. |
| [IsFilled](#IsFilled) | Determines whether the arrow cap is filled. |
| [SetFillState](#SetFillState) | Sets the fill state of the arrow cap. |
| [SetHeight](#SetHeight) | Sets the height of the arrow cap. |
| [SetMiddleInset](#SetMiddleInset) | Sets the number of units that the midpoint of the base shifts towards the vertex. |
| [SetWidth](#SetWidth) | Sets the width of the arrow cap. |
