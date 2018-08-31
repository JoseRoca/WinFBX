# CGpBrush Class

The **CGpBrush** class is a base interface that defines a **Brush** object. A **Brush** object is used to paint the interior of graphics shapes, such as rectangles, ellipses, pies, polygons, and paths. You must not instantiate the **CGpBrush** class, but to use one of its derived classes.

A closed shape, such as a rectangle or an ellipse, consists of an outline and an interior. The outline is drawn with a pen and the interior is filled with a brush. GDI+ provides several brush classes for filling the interiors of closed shapes: **CGpSolidBrush**, **CGpHatchBrush**, **CGpTextureBrush**, **CGpLinearGradientBrush**, and **CGpPathGradientBrush**. All of these classes inherit from the **CGpBrush** class.

**Inherits from**: CGpBase.
**Imclude file**: CGpBrush.inc.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Clone](#Clone) | Copies the contents of the existing **Brush** object into a new **Brush** object. |
| [GetType](#GetType) | Gets the type of this brush. |

# CGpSolidBrush Class

The **SolidBrush** object defines a solid color Brush object. A **Brush** object is used to fill in shapes similar to the way a paint brush can paint the inside of a shape.

**Inherits from**: CGpBrush.
**Imclude file**: CGpBrush.inc.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#ConstructorSBrush) | Create a **SolidBrush** object based on a color. |
| [GetColor](#GetColor) | Gets the color of this brush. |
| [SetColor](#SetColor) | Sets the color of this brush. |

# CGpHatchBrush Class

Creates a **HatchBrush** object based on a hatch style, a foreground color, and a background color.

**Inherits from**: CGpBrush.
**Imclude file**: CGpBrush.inc.

### Methods

| Name       | Description |
| ---------- | ----------- |
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
| [Constructors](#ConstructorLGBrush) | Create a **LinearGradientBrush** object from a set of boundary points and boundary colors. |
| [GetBlend](#GetBlend) | Gets the blend factors and their corresponding blend positions. |
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
| [Constructors](#ConstructorPGBrush) | Create a PathGradientBrush object based on an array of points. Initializes the wrap mode of the path gradient brush. |
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
| [Constructors](#ConstructorTBrush) | Create a texture brush. |
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
