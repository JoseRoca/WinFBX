# OpenGL Procedures

Assorted helper OpenGL procedures.

**Include File**: AfxGlut.inc.

| Name       | Description |
| ---------- | ----------- |
| [AfxGlutSolidCone](#AfxGlutCone) | Draws a solid cone. |
| [AfxGlutWireCone](#AfxGlutCone) | Draws a wireframe cone. |
| [AfxGlutSolidCube](#AfxGlutCube) | Draws a solid cube. |
| [AfxGlutWireCube](#AfxGlutCube) | Draws a wireframe cube. |
| [AfxGlutSolidCylinder](#AfxGlutCylinder) | Draws a solid cylinder. |
| [AfxGlutWireCylinder](#AfxGlutCylinder) | Draws a wireframe cylinder. |
| [AfxGlutSolidDodecahedron](#AfxGlutDodecahedron) | Draws a solid dodecahedron. |
| [AfxGlutWireDodecahedron](#AfxGlutDodecahedron) | Draws a wireframe dodecahedron. |
| [AfxGlutSolidIcosahedron](#AfxGlutIcosahedron) | Draws a solid icosahedron. |
| [AfxGlutWireIcosahedron](#AfxGlutIcosahedron) | Draws a wireframe icosahedron. |
| [AfxGlutSolidOctahedron](#AfxGlutOctahedron) | Draws a solid octahedron. |
| [AfxGlutWireOctahedron](#AfxGlutOctahedron) | Draws a wireframe octahedron. |
| [AfxGlutSolidRhombicDodecahedron](#AfxGlutRhombicDodecahedron) | Draws a solid rhombic dodecahedron. |
| [AfxGlutWireRhombicDodecahedron](#AfxGlutRhombicDodecahedron) | Draws a wireframe rhombic dodecahedron. |
| [AfxGlutSolidSphere](#AfxGlutSphere) | Renders a solid sphere centered at the origin of the modeling coordinate system. |
| [AfxGlutWireSphere](#AfxGlutSphere) | Renders a wireframe sphere centered at the origin of the modeling coordinate system. |
| [AfxGlutSolidTeapot](#AfxGlutTeapot) | Renders a solid teapot of the desired size, centered at the origin. |
| [AfxGlutWireTeapot](#AfxGlutTeapot) | Renders a wireframe teapot of the desired size, centered at the origin. |
| [AfxGlutSolidTetrahedron](#AfxGlutTetrahedron) | Renders a solid tetrahedron. |
| [AfxGlutWireTetrahedron](#AfxGlutTetrahedron) | Renders a wireframe tetrahedron. |
| [AfxGlutSolidTorus](#AfxGlutTorus) | Renders a solid torus (doughnut shape). |
| [AfxGlutWireTorus](#AfxGlutTorus) | Renders a wireframe torus (doughnut shape). |

# <a name="AfxGlutCone"></a>AfxGlutSolidCone / AfxGlutWireCone

Renders a right circular cone with a base centered at the origin and in the X-Y plane and its tip on the positive Z-axis. The wire cone is rendered with triangular elements.

```
SUB AfxGlutSolidCone (BYVAL dbase AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```
```
SUB AfxGlutWireCone (BYVAL dbase AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dbase* | The desired radius of the base of the cone. |
| *height* | The desired height of the cone. |
| *slices* | The desired number of slices around the cone. |
| *stacks* | The desired number of segments between the base and the tip of the cone (the number of points, including the tip, is stacks + 1). |

# <a name="AfxGlutCube"></a>AfxGlutSolidCube / AfxGlutWireCube

Renders a cube of the desired size, centered at the origin. Its faces are normal to the coordinate directions.

```
SUB AfxGlutSolidCube (BYVAL dbSize AS DOUBLE)
SUB AfxGlutWireCube (BYVAL dbSize AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dbSize* | The desired length of an edge of the cube. |

# <a name="AfxGlutCylinder"></a>AfxGlutSolidCylinder / AfxGlutWireCylinder

Draws a cylinder.

```
SUB AfxGlutSolidCylinder (BYVAL radius AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```
```
SUB AfxGlutWireCylinder (BYVAL radius AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *radius* | The desired radius of the cylinder. |
| *height* | The desired height of the cylinder. |
| *slices* | The desired number of slices around the cylinder. |
| *stacks* | The desired number of segments between the base and the top of the cylinder (the number of points, including the tip, is stacks + 1). |

# <a name="AfxGlutDodecahedron"></a>AfxGlutSolidDodecahedron / AfxGlutWireDodecahedron

Renders a dodecahedron whose corners are each a distance of sqrt(3) from the origin. The length of each side is sqrt(5)-1. There are twenty corners; interestingly enough, eight of them coincide with the corners of a cube with sizes of length 2.

```
SUB AfxGlutSolidDodecahedron
SUB AfxGlutWireDodecahedron
```

# <a name="AfxGlutIcosahedron"></a>AfxGlutSolidIcosahedron / AfxGlutWireIcosahedron

Renders an icosahedron whose corners are each a unit distance from the origin. The length of each side is slightly greater than one. Two of the corners lie on the positive and negative X-axes.

```
SUB AfxGlutSolidIcosahedron
SUB AfxGlutWireIcosahedron
```

# <a name="AfxGlutOctahedron"></a>AfxGlutSolidOctahedron / AfxGlutWireOctahedron

Renders an octahedron whose corners are each a distance of one from the origin. The length of each side is sqrt(2). The corners are on the positive and negative coordinate axes.

```
SUB AfxGlutSolidOctahedron
SUB AfxGlutWireOctahedron
```

# <a name="AfxGlutRhombicDodecahedron"></a>AfxGlutSolidRhombicDodecahedron / AfxGlutWireRhombicDodecahedron

Renders a rhombic dodecahedron whose corners are at most a distance of one from the origin. The rhombic dodecahedron has faces which are identical rhombuses (rhombi?) but which have some vertices at which three faces meet and some vertices at which four faces meet. The length of each side is sqrt(3)/2. Vertices at which four faces meet are found at (0, 0, +/- 1) and (+/- sqrt(2)/2, +/- sqrt(2)/2, 0). 

```
SUB AfxGlutSolidRhombicDodecahedron
SUB AfxGlutWireRhombicDodecahedron
```

# <a name="AfxGlutSphere"></a>AfxGlutSolidSphere / AfxGlutWireSphere

Renders a sphere centered at the origin of the modeling coordinate system. The north and south poles of the sphere are on the positive and negative Z-axes respectively and the prime meridian crosses the positive X-axis.

```
SUB AfxGlutSolidSphere (BYVAL radius AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)
SUB AfxGlutWireSphere (BYVAL radius AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *radius* | The desired radius of the sphere. |
| *slices* | The desired number of slices (divisions in the longitudinal direction) in the sphere. |
| *stacks* | The desired number of stacks (divisions in the latitudinal direction) in the sphere. The number of points in this direction, including the north and south poles, is stacks+1. |

# <a name="AfxGlutTeapot"></a>AfxGlutSolidTeapot / AfxGlutWireTeapot

Renders a solid teapot of the desired size, centered at the origin.

| Parameter  | Description |
| ---------- | ----------- |
| *dbSize* | The desired size of the teapot. |

```
SUB AfxGlutSolidTeapot
SUB AfxGlutWireTeapot
```

# <a name="AfxGlutTetrahedron"></a>AfxGlutSolidTetrahedron / AfxGlutWireTetrahedron

Renders a tetrahedron whose corners are each a distance of one from the origin. The length of each side is 2/3 sqrt(6). One corner is on the positive X-axis and another is in the X-Y plane with a positive Y-coordinate.

| Parameter  | Description |
| ---------- | ----------- |
| *dbSize* | The desired size of the teapot. |

```
SUB AfxGlutSolidTetrahedron
SUB AfxGlutWireTetrahedron
```

# <a name="AfxGlutTorus"></a>AfxGlutSolidTorus / AfxGlutWireTorus

Renders a torus centered at the origin of the modeling coordinate system. The torus is circularly symmetric about the Z-axis and starts at the positive X-axis.

```
SUB AfxGlutSolidTorus (BYVAL innerRadius AS DOUBLE, BYVAL outerRadius AS DOUBLE, _
   BYVAL nsides AS LONG, BYVAL rings AS LONG)
```
```
SUB AfxGlutWireTorus (BYVAL innerRadius AS DOUBLE, BYVAL outerRadius AS DOUBLE, _
   BYVAL nsides AS LONG, BYVAL rings AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *innerRadius* | Inner radius of the torus. |
| *outerRadius* | Outer radius of the torus. |
| *nsides* | Number of sides for each radial section. |
| *rings* | Number of radial divisions for the torus. |
