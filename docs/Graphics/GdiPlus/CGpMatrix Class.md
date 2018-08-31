# CGpMatrix Class

Encapsulates a 3-by-3 affine matrix that represents a geometric transform. A **Matrix** object represents a 3×3 matrix that, in turn, represents an affine transformation. A **Matrix** object stores only six of the 9 numbers in a 3×3 matrix because all 3×3 matrices that represent affine transformations have the same third column (0, 0, 1).

**Inherits from**: CGpBase.
**Imclude file**: CGpMatrix.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Creates and initializes a Matrix object that represents the identity matrix. |
| [Clone](#Clone) | Copies the contents of the existing Matrix object into a new Matrix object. |
| [Equals](#Equals) | Determines whether the elements of this matrix are equal to the elements of another matrix. |
| [GetElements](#GetElements) | Gets the elements of this matrix. |
| [Invert](#Invert) | If this matrix is invertible, the Invert method replaces the elements of this matrix with the elements of its inverse. |
| [IsIdentity](#IsIdentity) | Determines whether this matrix is the identity matrix. |
| [IsInvertible](#IsInvertible) | Determines whether this matrix is invertible. |
| [Multiply](#Multiply) | Updates this matrix with the product of itself and another matrix. |
| [OffsetX](#OffsetX) | Gets the horizontal translation value of this matrix, which is the element in row 3, column 1. |
| [OffsetY](#OffsetY) | Gets the vertical translation value of this matrix, which is the element in row 3, column 1. |
| [Reset](#Reset) | Updates this matrix with the elements of the identity matrix. |
| [Rotate](#Rotate) | Updates this matrix with the product of itself and a rotation matrix. |
| [RotateAt](#RotateAt) | Updates this matrix with the product of itself and a matrix that represents rotation about a specified point. |
| [Scale](#Scale) | Scales this matrix with the product of itself and a scaling matrix. |
| [SetElements](#SetElements) | Sets the elements of this matrix. |
| [Shear](#Shear) | Updates this matrix with the product of itself and a shearing matrix. |
| [TransformPoints](#TransformPoints) | Multiplies each point in an array by this matrix. Each point is treated as a row matrix. The multiplication is performed with the row matrix on the left and this matrix on the right. |
| [TransformVectors](#TransformVectors) | Multiplies each vector in an array by this matrix. |
| [Translate](#Translate) | Updates this matrix with the product of itself and a translation matrix. |
