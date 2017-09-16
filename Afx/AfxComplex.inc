' ########################################################################################
' Microsoft Windows
' File: AfxComplex.inc
' Contents: Complex numbers.
' Compiler: FreeBasic 32 & 64-bit
' Written in 2017 by José Roca. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' ########################################################################################
' Note: Some functions are based in the .NET Complex.cs class
' https://github.com/Microsoft/referencesource/blob/master/System.Numerics/System/Numerics/Complex.cs
' and, therefore, subject to the MIT license.
' Copyright (c) Microsoft Corporation.  All rights reserved.
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "crt/math.bi"
#include once "crt/limits.bi"

NAMESPACE Afx

' ########################################################################################
' Additional math functions.
' ########################################################################################

#define DBL_EPSILON      2.2204460492503131e-016   ' /* smallest such that 1.0+DBL_EPSILON != 1.0 */
#define SQRT_DBL_EPSILON 1.4901161193847656e-08

' ========================================================================================
' A function returns this value when the result of a mathematical operation yields a value
' that is too large in magnitude to be representable with its return type. This is one of
' the possible range errors, and is signaled by setting errno to ERANGE.
' Actually, functions may either return a positive or a negative HUGE_VAL (HUGE_VAL or -HUGE_VAL)
' to indicate the sign of the result.
' Returns: 1.#INF
' Note: Use HUGE_VALF, that is defined as #define HUGE_VALF (1.0F/0.0F) and returns the same value.
' ========================================================================================
PRIVATE FUNCTION HUGE_VAL () AS DOUBLE
   DIM hLib AS HMODULE, phuge AS DOUBLE PTR
   hLib = LoadLibraryW("msvcrt.dll")
   IF hLib THEN
      phuge = cast(DOUBLE PTR, GetProcAddress(hLib, "_HUGE"))
      IF phuge THEN FUNCTION = *phuge
      FreeLibrary hLib
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Computes the value of exp(x)-1 in a way that is accurate for small x.
' The true value of exp(1e-10) - 1 is 1.00000000005e-010 to about 32 significant digits.
' This example shows the superiority of expm1 in this case.
' expm1(1e-10)
'   1.00000000005e-010
' exp(1e-10) - 1
'   1.000000082740371e-010
' ========================================================================================
PRIVATE FUNCTION expm1 (BYVAL x AS DOUBLE) AS DOUBLE
   DIM AS DOUBLE i, sum, term
   IF fabs(x) < M_LN2 THEN
      ' /* Compute the Taylor series S = x + (1/2!) x^2 + (1/3!) x^3 + ... */
      i = 1
      sum = x
      term = x / 1
      DO
         i += 1
         term *= x / i
         sum += term
         IF fabs(term) <= fabs(sum) * DBL_EPSILON THEN EXIT DO
      LOOP
      FUNCTION = sum
   ELSE
      FUNCTION = exp(x) - 1
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Calculates the inverse hyperbolic cosine.
' Example:
'   DIM AS double pi = 3.1415926535
'   DIM AS double x, y
'   x = cosh(pi / 4)
'   y = acosh(x)
'   print "cosh = ", pi/4, x
'   print "acosh = ", x, y
' Output:
'   cosh =         0.785398163375              1.324609089232506
'   acosh =        1.324609089232506           0.7853981633749999
' ========================================================================================
PRIVATE FUNCTION acosh (BYVAL x AS DOUBLE) AS DOUBLE
   DIM t AS DOUBLE
   IF x > 1.0 / SQRT_DBL_EPSILON THEN
      FUNCTION = log(x) + M_LN2
   ELSEIF x > 2 THEN
      FUNCTION = log(2 * x - 1 / (sqr(x * x - 1) + x))
   ELSEIF x > 1 THEN
      t = x - 1
      FUNCTION = log1p(t + sqr(2 * t + t * t))
   ELSEIF x = 1 THEN
      FUNCTION = 0
   ELSE
      FUNCTION = 0 / 0  ' NAN
   END IF
END FUNCTION
' ========================================================================================
#define ArcCosH acosh

' ========================================================================================
' * Calculates the inverse hyperbolic sine.
' Example:
'   DIM AS double pi = 3.1415926535
'   DIM AS double x, y
'   x = sinh(pi / 4)
'   y = asinh(x)
'   print "sinh = ", pi/4, x
'   print "asinh = ", x, y
' Output:
'   sinh =         0.785398163375              0.8686709614562744
'   asinh =        0.8686709614562744          0.7853981633750001
' ========================================================================================
PRIVATE FUNCTION asinh (BYVAL x AS DOUBLE) AS DOUBLE
   DIM AS DOUBLE a, s, a2
   a = fabs(x)
   s = IIF(x < 0, -1, 1)
   IF a > 1 / SQRT_DBL_EPSILON THEN
      FUNCTION = s * (log(a) + M_LN2)
   ELSEIF a > 2 THEN
      FUNCTION = s * log(2 * a + 1 / (a + sqr(a * a + 1)))
   ELSEIF a > SQRT_DBL_EPSILON THEN
      a2 = a * a
      FUNCTION = s * log1p(a + a2 / (1 + sqr(1 + a2)))
   ELSE
      FUNCTION = x
   END IF
END FUNCTION
' ========================================================================================
#define ArcSinH asinh

' ========================================================================================
' * Returns the inverse hyperbolic tangent of a number.
' Examples:
' print atanh(0.76159416)
'    1.00000000962972
' print atanh(-0.1)
'   -0.1003353477310756
' ========================================================================================
PRIVATE FUNCTION atanh (BYVAL x AS DOUBLE) AS DOUBLE
   DIM AS DOUBLE a, s
   a = fabs(x)
   s = IIF(x < 0, -1, 1)
   IF a > 1 THEN
      FUNCTION = 0 / 0   ' NAN
   ELSEIF a = 1 THEN
      FUNCTION = IIF(x < 0, HUGE_VALF, -HUGE_VALF)
   ELSEIF a >= 0.5 THEN
      FUNCTION = s * 0.5 * log1p(2 * a / (1 - a))
   ELSEIF a > DBL_EPSILON THEN
      FUNCTION = s * 0.5 * log1p(2 * a + 2 * a * a / (1 - a))
   ELSE
      FUNCTION = x
   END IF
END FUNCTION
' ========================================================================================
#define ArcTanH atanh

' ========================================================================================
' Computes the value of sqr(x^2 + y^2 + z^2).
' ========================================================================================
PRIVATE FUNCTION hypot3 (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE, BYVAL z AS DOUBLE) AS DOUBLE
   RETURN hypot(hypot(x, y), z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Computes the value of sqr(x^2 + y^2 + z^2) in a way that avoids overflow.
' ========================================================================================
'PRIVATE FUNCTION hypot3 (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE, BYVAL z AS DOUBLE) AS DOUBLE
'   DIM AS DOUBLE xabs, yabs, zabs, w, r
'   xabs = fabs(x)
'   yabs = fabs(y)
'   zabs = fabs(z)
'   w = MAX(xabs, MAX(yabs, zabs))
'   IF w = 0 THEN
'      FUNCTION = 0
'   ELSE
'      r = w * SQR((xabs / w) * (xabs / w) + _
'                  (yabs / w) * (yabs / w) + _
'                  (zabs / w) * (zabs / w))
'      FUNCTION = r
'   END IF
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether x and y are approximately equal to a relative accuracy EPSILON.
' ========================================================================================
PRIVATE FUNCTION fcmp (BYVAL x1 AS DOUBLE, BYVAL x2 AS DOUBLE, BYVAL epsilon AS DOUBLE) AS LONG
   DIM exponent AS LONG
   DIM AS DOUBLE delta, diff, dmax
   ' /* Find exponent of largest absolute value */
   dmax = IIF(fabs(x1) > fabs(x2), x1, x2)
   frexp(dmax, @exponent)
   ' /* Form a neighborhood of size  2 * delta */
   delta = ldexp(epsilon, exponent)
   diff = x1 - x2
   IF diff > delta THEN        ' /* x1 > x2 */
      FUNCTION = 1
   ELSEIF diff < -delta THEN   ' /* x1 < x2 */
      FUNCTION = -1
   ELSE                        ' /* -delta <= diff <= delta */
      FUNCTION = 0             ' /* x1 ~=~ x2 */
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the argument is an infinity.
' Returns +1 if x is positive infinity, -1 if x is negative infinity and 0 otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxIsInf (BYVAL x AS DOUBLE) AS LONG
   IF NOT _finite(x) AND NOT _isnan(x) THEN
      FUNCTION = IIF&(x > 0, 1, -1)
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsInfinity (BYVAL x AS DOUBLE) AS LONG
   IF _finite(x) = TRUE AND _isnan(x) = FALSE THEN
      FUNCTION = IIF&(x > 0, 1, -1)
   END IF
END FUNCTION
' ========================================================================================

' ########################################################################################
' Complex numbers wrapper functions
' ########################################################################################

' ========================================================================================
' * Sets the real and imaginary parts of the complex number.
' Example: DIM z AS _complex = AfxCSet(3, 4)
' But you can al so use: DIM z AS _complex = (3, 4)
' ========================================================================================
PRIVATE FUNCTION AfxCSet (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(x, y)
END FUNCTION
' ========================================================================================
' =====================================================================================
' * Sets the real and imaginary parts of the complex number.
' Example: DIM z AS _complex = AfxCRect(3, 4)
' =====================================================================================
PRIVATE FUNCTION AfxCRect (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(x, y)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the real and imaginary parts of the complex number in polar format.
' Example: DIM z AS _complex = AfxCPolar(3, 4)
' =====================================================================================
PRIVATE FUNCTION AfxCPolar (BYVAL r AS DOUBLE, BYVAL theta AS DOUBLE) AS _complex
   RETURN TYPE<_complex> (r * cos(theta), r * sin(theta))
END FUNCTION
' =====================================================================================

' ========================================================================================
' * Returns the real part of a complex number.
' ========================================================================================
PRIVATE FUNCTION AfxCGetReal (BYREF z AS _complex) AS DOUBLE
   RETURN z.x
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the imaginary part of a complex number.
' ========================================================================================
PRIVATE FUNCTION AfxCGetImag (BYREF z AS _complex) AS DOUBLE
   RETURN z.y
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the real part of a complex number.
' ========================================================================================
PRIVATE SUB AfxCSetReal (BYREF z AS _complex, BYVAL x AS DOUBLE)
   z.x = x
END SUB
' ========================================================================================

' ========================================================================================
' * Sets the imaginary part of a complex number.
' ========================================================================================
PRIVATE SUB AfxCSetImag (BYREF z AS _complex, BYVAL y AS DOUBLE)
   z.y = y
END SUB
' ========================================================================================

'/* Complex numbers */

' ========================================================================================
' * Returns true if the two complex numbers are equal, or false otherwise.
' ========================================================================================
PRIVATE OPERATOR = (BYREF z1 AS _complex, BYREF z2 AS _complex) AS BOOLEAN
   RETURN (z1.x = z2.x AND z1.y = z2.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns true if the two complex numbers are different, or false otherwise.
' ========================================================================================
PRIVATE OPERATOR <> (BYREF z1 AS _complex, BYREF z2 AS _complex) AS BOOLEAN
   RETURN (z1.x <> z2.x OR z1.y <> z2.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Exchanges the contents of two complex numbers.
' ========================================================================================
PRIVATE SUB AfxCSwap (BYREF z1 AS _complex, BYREF z2 AS _complex)
   DIM z AS _complex = z1 : z1 = z2 : z2 = z
END SUB
' ========================================================================================

'/* Complex arithmetic operators */

' ========================================================================================
' * Returns the negative of a complex number.
' ========================================================================================
PRIVATE FUNCTION AfxCNegative (BYREF z AS _complex) AS _complex
   RETURN TYPE<_complex>(-z.x, -z.y)
END FUNCTION
' ========================================================================================
#define AfxCNegate AfxCNegative

' ========================================================================================
' * Returns the negative of a complex number.
' ========================================================================================
PRIVATE OPERATOR - (BYREF z AS _complex) AS _complex
   RETURN TYPE<_complex>(-z.x, -z.y) 
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the sum of complex numbers.
' ========================================================================================
PRIVATE OPERATOR + (BYREF z1 AS _complex, BYREF z2 AS _complex) AS _complex
   RETURN TYPE<_complex>(z1.x + z2.x, z1.y + z2.y) 
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYVAL a AS DOUBLE, BYREF z AS _complex) AS _complex
   RETURN TYPE<_complex>(z.x + a, z.y) 
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYREF z AS _complex, BYVAL a AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(z.x + a, z.y) 
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the difference of complex numbers.
' ========================================================================================
PRIVATE OPERATOR - (BYREF z1 AS _complex, BYREF z2 AS _complex) AS _complex
   RETURN TYPE<_complex>(z1.x - z2.x, z1.y - z2.y) 
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYVAL a AS DOUBLE, BYREF z AS _complex) AS _complex
   RETURN TYPE<_complex>(a - z.x, -z.y) 
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF z AS _complex, BYVAL a AS DOUBLE) AS _complex
  RETURN TYPE<_complex> (z.x - a, z.y) 
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the product of complex numbers.
' ========================================================================================
PRIVATE OPERATOR * (BYREF z1 AS _complex, BYREF z2 AS _complex) AS _complex
   RETURN TYPE<_complex>(z1.x * z2.x - z1.y * z2.y, z1.x * z2.y + z1.y * z2.x) 
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYVAL a AS DOUBLE, BYREF z AS _complex) AS _complex
   RETURN TYPE<_complex> (a * z.x, a * z.y) 
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF z AS _complex, BYVAL a AS DOUBLE) AS _complex
   RETURN TYPE<_complex> (a * z.x, a * z.y) 
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the quotient of complex numbers.
' - .NET 4.7 code:
' public static Complex operator /(Complex left, Complex right) {
'    // Division : Smith's formula.
'    double a = left.m_real;
'    double b = left.m_imaginary;
'    double c = right.m_real;
'    double d = right.m_imaginary;
'    if (Math.Abs(d) < Math.Abs(c)) {
'       double doc = d / c;
'       return new Complex((a + b * doc) / (c + d * doc), (b - a * doc) / (c + d * doc));
'    } else {
'       double cod = c / d;
'       return new Complex((b + a * cod) / (d + c * cod), (-a + b * cod) / (d + c * cod));
'    }
' }
' ========================================================================================
PRIVATE OPERATOR / (BYREF leftside AS _complex, BYREF rightside AS _complex) AS _complex
   ' // Division : Smith's formula.
   DIM a AS double = leftside.x
   DIM b AS DOUBLE = leftside.y
   DIM c AS DOUBLE = rightside.x
   DIM d AS DOUBLE = rightside.y
   IF ABS(d) < ABS(c) THEN
      DIM doc AS DOUBLE = d / c
      RETURN TYPE<_complex>((a + b * doc) / (c + d * doc), (b - a * doc) / (c + d * doc))
   ELSE
      DIM cod AS DOUBLE = c / d
      RETURN TYPE<_complex>((b + a * cod) / (d + c * cod), (-a + b * cod) / (d + c * cod))
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
'PRIVATE OPERATOR / (BYREF z1 AS _complex, BYREF z2 AS _complex) AS _complex
'   DIM d AS DOUBLE = 1 / (z2.x * z2.x + z2.y * z2.y)
'   RETURN TYPE <_complex>((z1.x * z2.x + z1.y * z2.y) * d, _
'                          (z1.y * z2.x - z1.x * z2.y) * d)
'END OPERATOR
' ========================================================================================
' ========================================================================================
' * Returns the quotient of the complex number a and the real number.
' ========================================================================================
PRIVATE OPERATOR / (BYVAL a AS DOUBLE, BYREF z AS _complex) AS _complex
   DIM d AS DOUBLE = a / (z.x * z.x + z.y * z.y)
   RETURN TYPE<_complex>(z.x * d, -z.y * d)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYREF z AS _complex, BYVAL a AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(z.x / a, z.y / a)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the sum of a complex number and a real number.
' ========================================================================================
PRIVATE FUNCTION AfxCAddReal (BYREF z AS _complex, BYVAL x AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(z.x + x, z.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * This function returns the sum of a complex number and an imaginary number.
' ========================================================================================
PRIVATE FUNCTION AfxCAddImag (BYREF z AS _complex, BYVAL y AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(z.x, z.y + y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the difference of a complex number and a real number.
' ========================================================================================
PRIVATE FUNCTION AfxCSubReal (BYREF z AS _complex, BYVAL x AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(z.x - x, z.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Returns the difference of a complex number and an imaginary number.
' ========================================================================================
PRIVATE FUNCTION AfxCSubImag (BYREF z AS _complex, BYVAL y AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(z.x, z.y - y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the product of a complex number and a real number.
' ========================================================================================
PRIVATE FUNCTION AfxCMulReal (BYREF z AS _complex, BYVAL x AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(z.x * x, z.y * x)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Returns the product of a complex number and an imaginary number.
' ========================================================================================
PRIVATE FUNCTION AfxCMulImag (BYREF z AS _complex, BYVAL y AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(-y * z.y, y * z.x)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the quotient of a complex number and a real number.
' ========================================================================================
PRIVATE FUNCTION AfxCDivReal (BYREF z AS _complex, BYVAL x AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(z.x / x, z.y / x)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the quotient of a complex number and an imaginary number.
' ========================================================================================
PRIVATE FUNCTION AfxCDivImag (BYREF z AS _complex, BYVAL y AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(z.y / y, -z.x / y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex conjugate of a complex number.
' Example:
'   DIM z AS _complex = (2, 3)
'   z = AfxCConj(z)
'   PRINT AfxCStr(z)
' Output: 2 - 3 * i
' ========================================================================================
' ========================================================================================
' - .NET 4.7 code:
' public static Complex Conjugate(Complex value) {
'   // Conjugate of a Complex number: the conjugate of x+i*y is x-i*y 
'   return (new Complex(value.m_real, (-value.m_imaginary)));
' }
' ========================================================================================
PRIVATE FUNCTION AfxCConj (BYREF z AS _complex) AS _complex
   RETURN TYPE<_complex>(z.x, -z.y)
END FUNCTION
' ========================================================================================
#define AfxCConjugate AfxCConj

'/* Properties of complex numbers */

' =====================================================================================
' * Returns the argument of a complex number.
' Examples:
'   DIM z AS _complex = (1, 0)
'   DIM d AS DOUBLE = AfxCArg(z)
'   PRINT d
' Output: 0.0
'   DIM z AS _complex = (0, 1)
'   DIM d AS DOUBLE = AfxCArg(z)
'   PRINT d
' Output: 1.570796326794897
'   DIM z AS _complex = (0, -1)
'   DIM d AS DOUBLE = AfxCArg(z)
' PRINT d
' Output: -1.570796326794897
'   DIM z AS _complex = (-1, 0)
'   DIM d AS DOUBLE = AfxCArg(z)
' PRINT d
' Output: 3.141592653589793
' =====================================================================================
PRIVATE FUNCTION AfxCArg (BYREF z AS _complex) AS DOUBLE
   FUNCTION = atan2(z.y, z.x)
END FUNCTION
' =====================================================================================
#define AfxCArgument AfxCArg
' =====================================================================================
' * Gets the phase of a complex number (same as argument, above).
' - NET 4.7 code:
' public Double Phase {
'   get { return Math.Atan2(m_imaginary, m_real); }
' }
' =====================================================================================
#define AfxCPhase AfxCArg

' =====================================================================================
' * Returns the magnitude of a complex number.
' Example:
'   DIM z AS _complex = (2, 3)
'   PRINT AfxCAbs(z)
' Output: 3.60555127546399
' =====================================================================================
' =====================================================================================
' - NET 4.7 code:
' public static Double Abs(Complex value) {
'    if(Double.IsInfinity(value.m_real) || Double.IsInfinity(value.m_imaginary)) {
'       return double.PositiveInfinity;
'    }
'    // |value| == sqrt(a^2 + b^2)
'    // sqrt(a^2 + b^2) == a/a * sqrt(a^2 + b^2) = a * sqrt(a^2/a^2 + b^2/a^2)
'    // Using the above we can factor out the square of the larger component to dodge overflow.
'    double c = Math.Abs(value.m_real);
'    double d = Math.Abs(value.m_imaginary);
'    if (c > d) {
'       double r = d / c;
'       return c * Math.Sqrt(1.0 + r * r);
'    } else if (d == 0.0) {
'       return c;  // c is either 0.0 or NaN
'    } else {
'       double r = c / d;
'       return d * Math.Sqrt(1.0 + r * r);
'    }
' }
' =====================================================================================
PRIVATE FUNCTION AfxCAbs (BYREF z AS _complex) AS DOUBLE
   IF AfxIsInfinity(z.x) OR AfxIsInfinity(z.y) THEN RETURN HUGE_VALF
   ' // |value| == sqrt(a^2 + b^2)
   ' // sqrt(a^2 + b^2) == a/a * sqrt(a^2 + b^2) = a * sqrt(a^2/a^2 + b^2/a^2)
   ' // Using the above we can factor out the square of the larger component to dodge overflow.
   DIM c AS DOUBLE = ABS(z.x)
   DIM d AS DOUBLE = ABS(z.y)
   IF c > d THEN
      DIM r AS DOUBLE = d / c
      RETURN c * SQR(1.0 + r * r)
   ELSEIF d = 0 THEN
      RETURN c   ' // c is either 0.0 or NaN
   ELSE
      DIM r AS DOUBLE = c / d
      RETURN d * SQR(1.0 + r * r)
   END IF
END FUNCTION
' =====================================================================================
#define AfxCMagnitude AfxCAbs

' ========================================================================================
' * Returns the natural logarithm of the magnitude of the complex number z, log|z|.
' It allows an accurate evaluation of \log|z| when |z| is close to one. The direct
' evaluation of log(CAbs(z)) would lead to a loss of precision in this case.
' Example:
'   DIM z AS _complex = (1.1, 0.1)
'   DIM d AS DOUBLE = AfxCLogAbs(z)
'   print d
' Output: 0.09942542937258279
' ========================================================================================
PRIVATE FUNCTION AfxCLogAbs (BYREF value AS _complex) AS DOUBLE
   RETURN LOG(AfxCAbs(value))
END FUNCTION
' ========================================================================================

' =====================================================================================
' * Returns the squared magnitude of a complex number, otherwise known as the complex norm.
' Examples:
'   DIM z AS _complex = (2, 3)
'   DIM d AS DOUBLE = AfxCAbs2(z)
'   PRINT d
' Output: 13
' --or--
'   DIM z AS _complex = (2, 3)
'   DIM d AS DOUBLE = AfxCNorm(z)
'   PRINT d
' Output: 13
' =====================================================================================
PRIVATE FUNCTION AfxCAbs2 (BYREF z AS _complex) AS DOUBLE
   FUNCTION = (z.x * z.x) + (z.y * z.y)
END FUNCTION
' =====================================================================================
#define AfxCNorm AfxCAbs2

' ========================================================================================
' * Returns the inverse, or reciprocal, of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCReciprocal(z)
'   PRINT AfxCStr(z)
' Output: 0.5 -0.5 * i
' ========================================================================================
' ========================================================================================
' .NET 4.7: System.Numerics/System/Numerics/Complex.cs
' public static Complex Reciprocal(Complex value) {
'    // Reciprocal of a Complex number : the reciprocal of x+i*y is 1/(x+i*y)
'    if ((value.m_real == 0) && (value.m_imaginary == 0)) {
'       return Complex.Zero;
'    }
'       return Complex.One / value;
' }
' ========================================================================================
PRIVATE FUNCTION AfxCReciprocal (BYREF value AS _complex) AS _complex
   IF value.x = 0 AND value.y = 0 THEN RETURN TYPE<_complex>(0, 0)
   RETURN TYPE<_complex>(1, 0) / value
END FUNCTION
' ========================================================================================
#define AfxCInverse AfxCReciprocal

'/* Elementary Complex Functions */

' ========================================================================================
' * Returns the square root of the complex number z. The branch cut is the negative
' real axis. The result always lies in the right half of the complex plane.
' Example:
'   DIM z AS _complex = (2, 3)
'   z = AfxCSqr(z)
'   PRINT AfxCStr(z)
' Output: 1.67414922803554 +0.895977476129838 * i
' Compute the square root of -1:
'   DIM z AS _complex = (-1)
'   z = AfxCSqr(z)
'   PRINT AfxCStr(z)
' Output: 0 +1.0 * i
' ========================================================================================
' // Make a complex number from polar coordinates
' - NET 4.7 code:
'public static Complex FromPolarCoordinates(Double magnitude, Double phase) /* Factory method to take polar inputs and create a Complex object */
'   { return new Complex((magnitude * Math.Cos(phase)), (magnitude * Math.Sin(phase))); }
FUNCTION AfxCFromPolarCoordinates (BYVAL magnitude AS DOUBLE, BYVAL phase AS DOUBLE) AS _complex
   RETURN TYPE<_complex>(magnitude * COS(phase), magnitude * SIN(phase))
END FUNCTION
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
'public static Complex Sqrt(Complex value) /* Square root ot the complex number */
'{ return Complex.FromPolarCoordinates(Math.Sqrt(value.Magnitude), value.Phase / 2.0); }
' ========================================================================================
FUNCTION AfxCSqr (BYREF value AS _complex) AS _complex
   DIM z AS _complex = AfxCFromPolarCoordinates(SQR(AfxCMagnitude(value)), AfxCPhase(value) / 2.0)
   IF value.x < 0 AND value.y = 0 THEN z.x = 0   ' // For negative real numbers
   RETURN z
END FUNCTION
' ========================================================================================
#define AfxCSqrt AfxCSqr

' ========================================================================================
' * Returns the complex square root of the real number x, where x may be negative.
' ========================================================================================
PRIVATE FUNCTION AfxCSqrReal (BYVAL value AS DOUBLE) AS _complex
   IF value >= 0 THEN RETURN TYPE<_complex>(SQR(value), 0)
   RETURN TYPE<_complex>(0, SQR(-value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex exponential of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCExp(z)
'   PRINT AfxCStr(z)
' Output: 1.468693939915885 +2.287355287178842 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
'public static Complex Exp(Complex value) /* The complex number raised to e */
'{
'   Double temp_factor = Math.Exp(value.m_real);
'   Double result_re = temp_factor * Math.Cos(value.m_imaginary);
'   Double result_im = temp_factor * Math.Sin(value.m_imaginary);
'   return (new Complex(result_re, result_im));
'}
' ========================================================================================
PRIVATE FUNCTION AfxCExp (BYREF value AS _complex) AS _complex
   DIM factor AS DOUBLE = EXP(value.x)
   RETURN TYPE<_complex>(factor * COS(value.y), factor * SIN(value.y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex natural logarithm (base e) of a complex number.
' The branch cut is the negative real axis.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCLog(z)
'   PRINT AfxCStr(z)
' Output: 0.3465735902799727 +0.7853981633974483 * i
'   DIM z AS _complex = (0, 0)
'   z = AfxCLog(z)
'   PRINT AfxCStr(z)
' Output: -1.#INF
' ========================================================================================
' ========================================================================================
' .NET 4.7:
' public static Complex Log(Complex value) /* Log of the complex number value to the base of 'e' */
' { return (new Complex((Math.Log(Abs(value))), (Math.Atan2(value.m_imaginary, value.m_real)))); }
' ========================================================================================
PRIVATE FUNCTION AfxCLog OVERLOAD (BYREF value AS _complex) AS _complex
   RETURN TYPE<_complex>(LOG(AfxCAbs(value)), atan2(value.y, value.x))
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Returns the complex base-value logarithm of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCLog(z, 10)
'   PRINT AfxCStr(z)
' Output: 0.1505149978319906 +0.3410940884604603 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7:
' public static Complex Log(Complex value, Double baseValue) /* Log of the complex number to a the base of a double */
' { return (Log(value) / Log(baseValue)); }
' ========================================================================================
PRIVATE FUNCTION AfxCLog OVERLOAD (BYREF value AS _complex, BYVAL baseValue AS DOUBLE) AS _complex
   DIM z AS _complex = AfxCLog(value)
   z.x /= LOG(baseValue)
   z.y /= LOG(baseValue)
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex base-10 logarithm of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCLog10(z)
'   PRINT AfxCStr(z)
' Output: 0.1505149978319906 +0.3410940884604603 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' private static Complex Scale(Complex value, Double factor) {
'    Double result_re = factor * value.m_real;
'    Double result_im = factor * value.m_imaginary;
'    return (new Complex(result_re, result_im));
' }
' public static Complex Log10(Complex value) /* Log to the base of 10 of the complex number */
' {
'    Complex temp_log = Log(value);
'    return (Scale(temp_log, (Double)LOG_10_INV));
' }
' ========================================================================================
PRIVATE FUNCTION AfxCLog10 (BYREF value AS _complex) AS _complex
   DIM z AS _complex = AfxCLog(value)
   z.x = z.x * M_LOG10E
   z.y = z.y * M_LOG10E
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex number a raised to the complex power b.
' Example:
'   DIM z AS _complex = (1, 1)
'   DIM b AS _complex = (2, 2)
'   print AfxCStr(AfxCPow(z, b))
' Output: -0.2656539988492412 +0.3198181138561362 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' public static Complex Pow(Complex value, Complex power) /* A complex number raised to another complex number */
' {
'    if (power == Complex.Zero) { return Complex.One; }
'    if (value == Complex.Zero) { return Complex.Zero; }
'    double a = value.m_real;
'    double b = value.m_imaginary;
'    double c = power.m_real;
'    double d = power.m_imaginary;
'    double rho = Complex.Abs(value);
'    double theta = Math.Atan2(b, a);
'    double newRho = c * theta + d * Math.Log(rho);
'    double t = Math.Pow(rho, c) * Math.Pow(Math.E, -d * theta);
'    return new Complex(t * Math.Cos(newRho), t * Math.Sin(newRho));
' }
' ========================================================================================
PRIVATE FUNCTION AfxCPow OVERLOAD (BYREF value AS _complex, BYREF power AS _complex) AS _complex
   IF power = TYPE<_complex>(0, 0) THEN RETURN TYPE<_complex>(1, 0)
   IF value = TYPE<_complex>(0, 0) THEN RETURN TYPE<_complex>(0, 0)
   DIM rho AS DOUBLE = AfxCAbs(value)
   DIM theta AS DOUBLE = atan2(value.y, value.x)
   DIM newRho AS DOUBLE = power.x * theta + power.y * LOG(rho)
   DIM t AS DOUBLE = POW(rho, power.x) * POW(M_E, -power.y * theta)
   RETURN TYPE<_complex>(t * COS(newRho), t * SIN(newRho))
END FUNCTION
' ========================================================================================
' ========================================================================================
OPERATOR ^ (BYREF value AS _complex, BYREF power AS _complex) AS _complex
   OPERATOR = AfxCPow(value, power)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns a complex number raised to a real number.
' Example:
'   DIM z AS _complex = (1, 1)
'   PRINT AfxCStr(AfxCPow(z, 2))
' Output: 1.224606353822378e-016 +2 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
'public static Complex Pow(Complex value, Double power) // A complex number raised to a real number 
'{ return Pow(value, new Complex(power, 0)); }
' ========================================================================================
PRIVATE FUNCTION AfxCPow OVERLOAD (BYREF value AS _complex, BYVAL power AS DOUBLE) AS _complex
   RETURN AfxCPow(value, TYPE<_complex>(power, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
OPERATOR ^ (BYREF value AS _complex, BYVAL power AS DOUBLE) AS _complex
   OPERATOR = AfxCPow(value, power)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns a a real number raised to a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCStr(2 ^ z)
'   PRINT z
' Output: 1.538477802727944 +1.27792255262727 * i
' ========================================================================================
OPERATOR ^ (BYVAL value AS DOUBLE, BYREF power AS _complex) AS _complex
   OPERATOR = AfxCExp(power * LOG(value))
END OPERATOR
' ========================================================================================


'/* Complex Trigonometric Functions */

' ========================================================================================
' * Returns the complex sine of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCSin(z)
'   PRINT AfxCStr(z)
' Output: 1.298457581415977 +0.6349639147847361 * i
' ========================================================================================
PRIVATE FUNCTION AfxCSin (BYREF value AS _complex) AS _complex
   IF value.y = 0 THEN RETURN TYPE<_complex>(SIN(value.x), 0)
   RETURN TYPE<_complex>(SIN(value.x) * cosh(value.y), COS(value.x) * sinh(value.y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex cosine of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCCos(z)
'   PRINT AfxCStr(z)
' Output: 0.8337300251311491 -0.9888977057628651 * i
' ========================================================================================
PRIVATE FUNCTION AfxCCos (BYREF value AS _complex) AS _complex
   IF value.y = 0 THEN RETURN TYPE<_complex>(COS(value.x), 0)
   RETURN TYPE<_complex>(COS(value.x) * cosh(value.y), SIN(value.x) * sinh(-value.y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex secant of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCSec(z)
'   PRINT AfxCStr(z)
' Output: 0.4983370305551869 +0.591083841721045 * i
' ========================================================================================
PRIVATE FUNCTION AfxCSec (BYREF value AS _complex) AS _complex
   RETURN AfxCReciprocal(AfxCCos(value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex cosecant of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCCsc(z)
'   PRINT AfxCStr(z)
' Output: 0.6215180171704285 -0.3039310016284265 * i
' ========================================================================================
PRIVATE FUNCTION AfxCCsc (BYREF value AS _complex) AS _complex
   RETURN AfxCReciprocal(AfxCSin(value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex tangent of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCTan(z)
'   PRINT AfxCStr(z)
' Output: 0.2717525853195117 +1.083923327338695 * i
' ========================================================================================
' - NET 4.7 code:
' public static Complex Tan(Complex value) { return (Sin(value) / Cos(value)); }
' ========================================================================================
PRIVATE FUNCTION AfxCTan (BYREF value AS _complex) AS _complex
   RETURN TYPE<_complex>(AfxCSin(value) / AfxCCos(value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex cotangent of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCCot(z)
'   PRINT AfxCStr(z)
' Output: 0.2176215618544027 -0.8680141428959249 * i
' ========================================================================================
PRIVATE FUNCTION AfxCCot (BYREF value AS _complex) AS _complex
   RETURN TYPE<_complex>(AfxCReciprocal(AfxCTan(value)))
END FUNCTION
' ========================================================================================

'/* Inverse Complex Trigonometric Functions */

' ========================================================================================
' Returns the complex arcsine of a real number.
' For a between -1 and 1, the function returns a real value in the range [-pi/2, pi/2].
' For a less than -1 the result has a real part of -pi/2 and a positive imaginary part.
' For a greater than 1 the result has a real part of pi/2 and a negative imaginary part.
' Example:
'   DIM z AS _complex
'   z = AfxCArcSinReal(1)
'   PRINT AfxCStr(z)
' Output: 1.570796326794897 +0 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcSinReal (BYVAL value AS DOUBLE) AS _complex
   IF ABS(value) <= 1 THEN RETURN TYPE<_complex>(ASIN(value))
   IF value < 0 THEN RETURN TYPE<_complex>(-M_PI_2, acosh(-value))
   RETURN TYPE<_complex>(M_PI_2, acosh(value))
END FUNCTION
' ========================================================================================
#define AfxCASinReal AfxCArcSinReal

' ========================================================================================
' * Returns the complex arcsine of a complex number.
' The branch cuts are on the real axis, less than -1 and greater than 1.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcSin(z)
'   PRINT AfxCStr(z)
' Output: 0.6662394324925152 +1.061275061905036 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' public static Complex Asin(Complex value) /* Arcsin */
' { return (-ImaginaryOne) * Log(ImaginaryOne * value + Sqrt(One - value * value)); }
' ========================================================================================
PRIVATE FUNCTION AfxCArcSin (BYREF value AS _complex) AS _complex
   DIM ImaginaryOne AS _complex = TYPE<_complex>(0.0, 1.0)
   DIM One AS _complex = TYPE<_complex>(1.0, 0.0)
   RETURN (-ImaginaryOne) * AfxCLog(ImaginaryOne * value + AfxCSqr(One - value * value))
END FUNCTION
' ========================================================================================
#define AfxCAsin AfxCArcSin

' ========================================================================================
' Returns the complex arccosine of a real number.
' For a between -1 and 1, the function returns a real value in the range [0, pi].
' For a less than -1 the result has a real part of pi and a negative imaginary part.
' For a greater than 1 the result is purely imaginary and positive.
' Examples:
' print AfxCStr(AfxCArcCosReal(1)) ' = 0 0 * i
' print AfxCStr(AfxCArcCosReal(-1)) ' = 3.141592653589793 0 * i
' print AfxCStr(AfxCArcCosReal(2)) ' = 0 +1.316957896924817 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcCosReal (BYVAL value AS DOUBLE) AS _complex
   IF ABS(value) <= 1.0 THEN RETURN TYPE<_complex>(acos(value))
   IF value < 0 THEN RETURN TYPE<_complex>(M_PI, -acosh(-value))
   RETURN TYPE<_complex>(0, acosh(value))
END FUNCTION
' ========================================================================================
#define AfxCACosReal AfxCArcCosReal

' ========================================================================================
' * Returns the complex arccosine of a complex number.
' The branch cuts are on the real axis, less than -1 and greater than 1.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcCos(z)
'   print AfxCStr(z)
' Output: 0.9045568943023814 -1.061275061905036 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' public static Complex Acos(Complex value) /* Arccos */
' { return (-ImaginaryOne) * Log(value + ImaginaryOne*Sqrt(One - (value * value))); }
' ========================================================================================
PRIVATE FUNCTION AfxCArcCos (BYREF value AS _complex) AS _complex
   DIM ImaginaryOne AS _complex = TYPE<_complex>(0.0, 1.0)
   DIM One AS _complex = TYPE<_complex>(1.0, 0.0)
   RETURN (-ImaginaryOne) * AfxCLog(value + ImaginaryOne * AfxCSqr(One - (value * value)))
END FUNCTION
' ========================================================================================
#define AfxCAcos AfxCArcCos

' ========================================================================================
' Returns the complex arcsecant of the a number.
' Example
' print AfxCStr(AfxCArcSecReal(1.1))
' Output: 0.4296996661514246 0 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcSecReal (BYVAL value AS DOUBLE) AS _complex
   IF value <= -1 OR value >= 1 THEN RETURN TYPE<_complex>(acos(1 / value), 0)
   IF value >= 0 THEN RETURN TYPE<_complex>(0, acosh(1 / value))
   RETURN TYPE<_complex>(M_PI, -acosh(-1 / value))
END FUNCTION
' ========================================================================================
#define AfxCASecReal AfxCArcSecReal

' ========================================================================================
' * Returns the complex arcsecant of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcSec(z)
'   PRINT AfxCStr(z)
' Output: 1.118517879643706 +0.5306375309525176 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcSec (BYREF value AS _complex) AS _complex
   RETURN AfxCArcCos(AfxCReciprocal(value))
END FUNCTION
' ========================================================================================
#define AfxCASec AfxCArcSec

' ========================================================================================
' Returns the complex arccosecant of a real number.
' Example:
'   DIM z AS _complex = (1, 1)
'   print AfxCStr(AfcCArcCscReal(1))
' Output: 1.570796326794897 0 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcCscReal (BYVAL value AS DOUBLE) AS _complex
   IF value <= -1 OR value >= 1 THEN RETURN TYPE<_complex>(ASIN(1 / value))
   IF value >= 0 THEN RETURN TYPE<_complex>(M_PI_2, -acosh(1 / value))
   RETURN TYPE<_complex>(-M_PI_2, acosh(-1 / value))
END FUNCTION
' ========================================================================================
#define AfxCACscReal AfxCArcCscReal

' ========================================================================================
' * Returns the complex arccosecant of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcCsc(z)
'   PRINT AfxCStr(z)
' Output: 0.4522784471511907 -0.5306375309525178 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcCsc (BYREF value AS _complex) AS _complex
   RETURN AfxCArcSin(AfxCReciprocal(value))
END FUNCTION
' ========================================================================================
#define AfxCACsc AfxCArcCsc

' ========================================================================================
' * This function returns the complex arctangent of a complex number.
' The branch cuts are on the imaginary axis, below -i and above i.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcTan(z)
'   PRINT AfxCStr(z)
' Output: 1.017221967897851 +0.4023594781085251 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' public static Complex Atan(Complex value) /* Arctan */
' {
'    Complex Two = new Complex(2.0, 0.0);
'    return (ImaginaryOne / Two) * (Log(One - ImaginaryOne * value) - Log(One + ImaginaryOne * value));
' }
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxCArcTan (BYREF value AS _complex) AS _complex
   DIM ImaginaryOne AS _complex = TYPE<_complex>(0.0, 1.0)
   DIM One AS _complex = TYPE<_complex>(1.0, 0.0)
   DIM Two AS _complex = TYPE<_complex>(2.0, 0.0)
   RETURN (ImaginaryOne / Two) * (AfxCLog(One - ImaginaryOne * value) - AfxCLog(One + ImaginaryOne * value))
END FUNCTION
' ========================================================================================
#define AfxCATan AfxCArcTan

' ========================================================================================
' * Returns the complex arccotangent of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcCot(z)
'   PRINT AfxCStr(z)
' Output: 0.5535743588970452 -0.4023594781085251 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcCot (BYREF value AS _complex) AS _complex
   IF value.x = 0 AND value.y = 0 THEN RETURN TYPE<_complex>(M_PI_2)
   RETURN AfxCArcTan(AfxCReciprocal(value))
END FUNCTION
' ========================================================================================
#define AfxCACot AfxCArcCot

'/* Complex Hyperbolic Functions */

' ========================================================================================
' * Returns the complex hyperbolic sine of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCSinH(z)
'   PRINT AfxCStr(z)
' Output: 0.6349639147847361 +1.298457581415977 * i
' ========================================================================================
PRIVATE FUNCTION AfxCSinH (BYREF value AS _complex) AS _complex
   RETURN TYPE<_complex>(sinh(value.x) * cos(value.y), cosh(value.x) * sin(value.y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic cosine of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCCosH(z)
'   PRINT AfxCStr(z)
' Output: 0.8337300251311491 +0.9888977057628651 * i
' ========================================================================================
PRIVATE FUNCTION AfxCCosH (BYREF value AS _complex) AS _complex
   RETURN TYPE<_complex>(cosh(value.x) * cos(value.y), sinh(value.x) * sin(value.y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic secant of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCSecH(z)
'   PRINT AfxCStr(z)
' Output: 0.4983370305551869 -0.591083841721045 * i
' ========================================================================================
PRIVATE FUNCTION AfxCSecH (BYREF value AS _complex) AS _complex
   RETURN AfxCReciprocal(AfxCCosH(value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic cosecant of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCCscH(z)
'   PRINT AfxCStr(z)
' Output: 0.3039310016284265 -0.6215180171704285 * i
' ========================================================================================
PRIVATE FUNCTION AfxCCscH (BYREF value AS _complex) AS _complex
   RETURN AfxCReciprocal(AfxCSinH(value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic tangent of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCTanH(z)
'   PRINT AfxCStr(z)
' Output: 1.083923327338695 +0.2717525853195117 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' public static Complex Tanh(Complex value) /* Hyperbolic tan */
' { return (Sinh(value) / Cosh(value)); }
' ========================================================================================
PRIVATE FUNCTION AfxCTanH (BYREF value AS _complex) AS _complex
   RETURN AfxCSinH(value) / AfxCCosh(value)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic cotangent of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCCotH(cpx)
'   PRINT AfxCStr(z)
' Output: 0.8680141428959249 -0.2176215618544027 * i
' ========================================================================================
PRIVATE FUNCTION AfxCCotH (BYREF value AS _complex) AS _complex
   RETURN TYPE<_complex>(AfxCReciprocal(AfxCTanH(value)))
END FUNCTION
' ========================================================================================

'/* Inverse Complex Hyperbolic Functions */

' ========================================================================================
' * Returns the complex hyperbolic arcsine of a complex number.
' The branch cuts are on the imaginary axis, below -i and above i.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcSinH(z)
'   PRINT AfxCStr(z)
' Output: 1.061275061905036 +0.6662394324925152 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcSinH (BYREF value AS _complex) AS _complex
   RETURN AfxCMulImag(AfxCArcSin(AfxCMulImag(value, 1.0)), -1.0)
END FUNCTION
' ========================================================================================
#define AfxCASinH AfxCArcSinH

' ========================================================================================
' Returns the complex hyperbolic arccosine of a real number.
' ========================================================================================
PRIVATE FUNCTION AfxCArcCoshReal (BYVAL value AS DOUBLE) AS _complex
   IF value >= 1 THEN RETURN TYPE<_complex>(acosh(value), 0)
   IF value >= -1 THEN RETURN TYPE<_complex>(0, acos(value))
   RETURN TYPE<_complex>(acosh(-value), M_PI)
END FUNCTION
' ========================================================================================
#define AfxCACoshReal AfxCArcCoshReal

' ========================================================================================
' * Returns the complex hyperbolic arccosine of a complex number.
' The branch cut is on the real axis, less than 1.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcCosH(z)
'   PRINT AfxCStr(z)
' Output: 1.061275061905036 +0.9045568943023814 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcCosH (BYREF value AS _complex) AS _complex
   DIM z AS _complex = AfxCArcCos(value)
   RETURN AfxCMulImag(z, IIF(z.y > 0, -1.0, 1.0))
END FUNCTION
' ========================================================================================
#define AfxCACosH AfxCArcCosH

' ========================================================================================
' * Returns the complex hyperbolic arcsecant of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcSecH(z)
'   PRINT AfxCStr(z)
' Output: 0.5306375309525179 -1.118517879643706 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcSecH (BYREF value AS _complex) AS _complex
   RETURN AfxCArcCosH(AfxCReciprocal(value))
END FUNCTION
' ========================================================================================
#define AfxCAsecH AfxCArcSecH

' ========================================================================================
' * This function returns the complex hyperbolic arccosecant of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcCscH(z)
'   PRINT AfxCStr(z)
' Output: 0.5306375309525179 -0.4522784471511906 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcCscH (BYREF value AS _complex) AS _complex
   RETURN AfxCArcSinH(AfxCReciprocal(value))
END FUNCTION
' ========================================================================================
#define AfxCACscH AfxCArcCscH

' ========================================================================================
' Returns the complex hyperbolic arctangent of a real number.
PRIVATE FUNCTION AfxCArcTanhReal (BYVAL value AS DOUBLE) AS _complex
   IF value > -1 AND value < 1 THEN RETURN TYPE<_complex>(atanh(value), 0)
   RETURN TYPE<_complex>(atanh(1 / value), IIF(value < 0, M_PI_2, -M_PI_2))
END FUNCTION
' ========================================================================================
#define AfxCATanhReal AfxCArcTanhReal

' ========================================================================================
' * Returns the complex hyperbolic arctangent of a complex number.
' The branch cuts are on the real axis, less than -1 and greater than 1.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcTanH(z)
'   PRINT AfxCStr(z)
' Output: 0.4023594781085251 +1.017221967897851 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcTanH (BYREF value AS _complex) AS _complex
   IF value.y = 0 THEN RETURN AfxCArcTanhReal(value.x)
   RETURN AfxCMulImag(AfxCArcTan(AfxCMulImag(value, 1)), -1)
END FUNCTION
' ========================================================================================
#define AfxCATanH AfxCArcTanH

' ========================================================================================
' * Returns the complex hyperbolic arccotangent of a complex number.
' Example:
'   DIM z AS _complex = (1, 1)
'   z = AfxCArcCotH(z)
'   PRINT AfxCStr(z)
' Output: 0.4023594781085251 -0.5535743588970453 * i
' ========================================================================================
PRIVATE FUNCTION AfxCArcCotH (BYREF value AS _complex) AS _complex
   RETURN AfxCArcTanH(AfxCReciprocal(value))
END FUNCTION
' ========================================================================================
#define AfxCACotH AfxCArcCotH

' ========================================================================================
' Returns the sign of a complex number.
' If number is greater than zero, then CSgn returns 1.
' If number is equal to zero, then CSgn returns 0.
' If number is less than zero, then CSgn returns -1.
' ========================================================================================
PRIVATE FUNCTION AfxCSgn (BYREF z AS _complex) AS LONG
   IF z.x > 0 THEN RETURN 1
   IF z.x < 0 THEN RETURN -1
   IF z.y > 0 THEN RETURN 1
   IF z.y < 0 THEN RETURN -1
   RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a complex number as a formated string.
' ========================================================================================
PRIVATE FUNCTION AfxCStr (BYREF z AS _complex) AS STRING
   FUNCTION = STR(z.x) & IIF(SGN(z.y) = -1 OR SGN(z.y) = 0, " ", " +") & STR(z.y) & " * i"
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a complex number in polar form: Z = |Z| * exp(Arg(Z) * i)
' ========================================================================================
PRIVATE FUNCTION AfxCStrPolar (BYREF z AS _complex) AS STRING
   FUNCTION = STR(AfxCAbs(z)) & " * exp(" & STR(AfxCArg(z)) & " * i)"
END FUNCTION
' ========================================================================================

END NAMESPACE
