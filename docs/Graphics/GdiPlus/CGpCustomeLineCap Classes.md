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

# <a name="ConstructorCustomLineCap"></a>Constructor (CGpCustomLineCap)

Creates a **CustomLineCap** object.

```
CONSTRUCTOR CGpCustomLineCap (BYVAL pFillPath AS CGpGraphicsPath PTR, _
   BYVAL pStrokePath AS CGpGraphicsPath PTR, BYVAL baseCap AS LineCap = LinecapFLat, _
   BYVAL baseInset AS SINGLE = 0.0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFillPath* | Pointer to a path. |
| *pStrokePath* | Pointer to a path. |
| *baseCap* | Optional. Element of the **LineCap** enumeration that specifies the line cap that will be used. The default value is **LineCapFlat**.  |
| *baseInset* | Optional. The default value is 0.0. |

#### Remarks

The *fillPath* and *strokePath* parameters cannot be used at the same time. You should pass NULL to one of those two parameters. If you pass nonnull values to both parameters, then *fillPath* is ignored.

The **CustomLineCap** class uses the winding fill mode regardless of the fill mode that is set for the **GraphicsPath** object passed to the **CustomLineCap** constructor.


# <a name="Clone"></a>Clone (CGpCustomLineCap)

Copies the contents of the existing **CustomLineCap** object into a new **CustomLineCap** object.

```
FUNCTION Clone (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCustomLineCap* | Pointer to the **CustomLineCap** object where to copy the contents of the existing object. |

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object with a stroke path, creates a second
' CustomLineCap object by cloning the first, and then assigns the cloned cap to a Pen object.
' It then draws a line by using the Pen object.
' ========================================================================================
SUB Example_Clone (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object
   DIM firstCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)

  ' // Create a copy of firstCap
   DIM secondCap AS CGpCustomLineCap
   firstCap.Clone(@secondCap)

  ' // Create a Pen object, assign second cap as the custom end cap, and draw a line
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 1)
   pen.SetCustomEndCap(@secondCap)
   graphics.DrawLine(@pen, 0, 0, 100, 100)

END SUB
' ========================================================================================
```

# <a name="GetBaseCap"></a>GetBaseCap (CGpCustomLineCap)

Gets the style of the base cap. The base cap is used as a cap at the end of a line along with this CustomLineCap object.

```
FUNCTION GetBaseCap () AS LineCap
```

#### Return value

This method returns the value of the line cap used at the base of the line.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object, sets its base cap, and then gets
' the base cap and assigns it to a Pen object. It then uses the Pen object to draw a line.
' ========================================================================================
SUB Example_GetBaseCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' //Create a Path object
   DIM capPath AS CGpGraphicsPath

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)
   custCap.SetBaseCap(LineCapRound)

   ' // Get the base cap of custCap
   DIM baseCap AS LineCap = custCap.GetBaseCap

   ' // Create a Pen object, assign baseCap as the end cap, and draw a line
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 10)
   pen.SetEndCap(baseCap)
   graphics.DrawLine(@pen, 0, 0, 100, 100)

END SUB
' ========================================================================================
```
