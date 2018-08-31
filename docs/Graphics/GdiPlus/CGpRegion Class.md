# CGpRegion Class

Provides methods to manage Regions. Regions are used in clipping and hit-testing operations.

The Region object describes an area of the display surface. The area can be any shape. In other words, the boundary of the area can be a combination of curved and straight lines. Regions can also be created from the interiors of rectangles, paths, or a combination of these. Regions are used in clipping and hit-testing operations.

**Inherits from**: CGpBase.
**Imclude file**: CGpRegion.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Create a Region object. |
| [Clone](#Clone) | Copies the contents of the existing Region object into a new Region object. |
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
