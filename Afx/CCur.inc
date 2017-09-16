' ########################################################################################
' Microsoft Windows
' File: CCur.inc
' Contents: Currency data type
' Compiler: FreeBasic 32 & 64-bit
' Written in 2017 by José Roca. Freeware. Use at your own risk.
' Portions of this code are based in the ATL's CComCurrency class.
' Copyright (C) Microsoft Corporation.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"

NAMESPACE Afx

'union tagCY
'   type
'      Lo as ulong
'      Hi as long
'   end type

'   int64 as LONGLONG
'end union

const CY_MIN_INTEGER   = -922337203685477LL
const CY_MAX_INTEGER   = 922337203685477LL
const CY_MIN_FRACTION  = -9999
const CY_MAX_FRACTION  = 9999
const CY_SCALE         = 10000

' ========================================================================================
' CCur class
' CCUR is a wrapper for the CURRENCY data type. CURRENCY is implemented as an 8-byte
' two's-complement integer value scaled by 10,000. This gives a fixed-point number with
' 15 digits to the left of the decimal point and 4 digits to the right. The CURRENCY data
' type is extremely useful for calculations involving money, or for any fixed-point
' calculations where accuracy is important.
' The CCUR wrapper implements arithmetic, assignment, and comparison operations for this
' fixed-point type, and provides access to the numbers on either side of the decimal point
' in the form of two components: an integer component which stores the value to the left
' of the decimal point, and a fractional component which stores the value to the right of
' the decimal point. The fractional component is stored internally as an integer value
' between -9999 (CY_MIN_FRACTION) and +9999 (CY_MAX_FRACTION). The function GetFraction
' returns a value scaled by a factor of 10000 (CY_SCALE).
' When specifying the integer and fractional components of a CCUR object, remember that
' the fractional component is a number in the range 0 to 9999. This is important when
' dealing with a currency such as the US dollar that expresses amounts using only two
' significant digits after the decimal point. Even though the last two digits are not
' displayed, they must be taken into account.
' ========================================================================================
TYPE CCur

Public:
   DIM m_cur AS CY                    ' // The underlying CURRENCY structure

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYREF cSrc AS CCUR)
   DECLARE CONSTRUCTOR (BYVAL cySrc AS CURRENCY)
   DECLARE CONSTRUCTOR (BYVAL nInteger AS LONGLONG)
   DECLARE CONSTRUCTOR (BYVAL nInteger AS LONGLONG, BYVAL nFraction AS SHORT)
   DECLARE CONSTRUCTOR (BYVAL bSrc AS BYTE)
   DECLARE CONSTRUCTOR (BYVAL ubSrc AS UBYTE)
   DECLARE CONSTRUCTOR (BYVAL sSrc AS SHORT)
   DECLARE CONSTRUCTOR (BYVAL usSrc AS USHORT)
   DECLARE CONSTRUCTOR (BYVAL lSrc AS LONG)
   DECLARE CONSTRUCTOR (BYVAL ulSrc AS ULONG)
   DECLARE CONSTRUCTOR (BYVAL fSrc AS SINGLE)
   DECLARE CONSTRUCTOR (BYVAL dSrc AS DOUBLE)
   DECLARE CONSTRUCTOR (BYVAL dSrc AS DECIMAL)
   DECLARE CONSTRUCTOR (BYREF pwszSrc AS WSTRING)
   DECLARE CONSTRUCTOR (BYVAL varSrc AS VARIANT)
   DECLARE CONSTRUCTOR (BYVAL pDispSrc AS IDispatch PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION SetInteger (BYVAL nInteger AS LONGLONG) AS HRESULT
   DECLARE FUNCTION SetFraction (BYVAL nFraction AS SHORT) AS HRESULT
   DECLARE OPERATOR LET (BYREF cSrc AS CCUR)
   DECLARE OPERATOR LET (BYVAL bSrc AS BYTE)
   DECLARE OPERATOR LET (BYVAL ubSrc AS UBYTE)
   DECLARE OPERATOR LET (BYVAL sSrc AS SHORT)
   DECLARE OPERATOR LET (BYVAL usSrc AS USHORT)
   DECLARE OPERATOR LET (BYVAL lSrc AS LONG)
   DECLARE OPERATOR LET (BYVAL ulSrc AS ULONG)
   DECLARE OPERATOR LET (BYVAL fSrc AS SINGLE)
   DECLARE OPERATOR LET (BYVAL dSrc AS DOUBLE)
   DECLARE OPERATOR LET (BYVAL dSrc AS DECIMAL)
   DECLARE OPERATOR LET (BYREF pwszSrc AS WSTRING)
   DECLARE OPERATOR LET (BYVAL varSrc AS VARIANT)
   DECLARE OPERATOR LET (BYVAL pDispSrc AS IDispatch PTR)
   DECLARE OPERATOR += (BYREF cur AS CCUR)
   DECLARE OPERATOR += (BYVAL nValue AS LONG)
   DECLARE OPERATOR += (BYVAL nValue AS DOUBLE)
   DECLARE OPERATOR -= (BYREF cur AS CCUR)
   DECLARE OPERATOR -= (BYREF nValue AS LONG)
   DECLARE OPERATOR -= (BYREF nValue AS DOUBLE)
   DECLARE OPERATOR *= (BYREF cur AS CCUR)
   DECLARE OPERATOR *= (BYVAL nOperand AS LONG)
   DECLARE OPERATOR *= (BYVAL nOperand AS DOUBLE)
   DECLARE OPERATOR /= (BYREF cOperand AS CCUR)
   DECLARE OPERATOR /= (BYVAL nOperand AS LONG)
   DECLARE OPERATOR /= (BYVAL nOperand AS DOUBLE)
   DECLARE OPERATOR CAST () AS CURRENCY
   DECLARE OPERATOR CAST () AS DOUBLE
   DECLARE OPERATOR CAST () AS STRING
   DECLARE OPERATOR CAST () AS VARIANT
   DECLARE FUNCTION Round (BYVAL nDecimals AS LONG) AS CCUR
   DECLARE FUNCTION GetInteger () AS LONGLONG
   DECLARE FUNCTION GetFraction () AS SHORT

END TYPE
' ========================================================================================

' ========================================================================================
' CCur constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CCur
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYREF cSrc AS CCUR)
   m_cur.int64 = cSrc.m_cur.int64
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL cySrc AS CURRENCY)
   m_cur.int64 = cySrc.int64
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL nInteger AS LONGLONG)
   this.SetInteger(nInteger)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL nInteger AS LONGLONG, BYVAL nFraction AS SHORT)
   this.SetInteger(nInteger)
   this.SetFraction(nFraction)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL bSrc AS BYTE)
   m_cur.int64 = bSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL ubSrc AS UBYTE)
   m_cur.int64 = ubSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL sSrc AS SHORT)
   m_cur.int64 = sSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL usSrc AS USHORT)
   m_cur.int64 = usSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL lSrc AS LONG)
   m_cur.int64 = lSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL ulSrc AS ULONG)
   m_cur.int64 = ulSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL fSrc AS SINGLE)
   m_cur.int64 = fSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL dSrc AS DOUBLE)
   VarCyFromR8(dSrc, @m_cur)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL dSrc AS DECIMAL)
   VarCyFromDec(@dSrc, @m_cur)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYREF pwszSrc AS WSTRING)
   VarCyFromStr(@pwszSrc, GetThreadLocale(), LOCALE_NOUSEROVERRIDE, @m_cur)
   m_Cur.int64 /= CY_SCALE
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL varSrc AS VARIANT)
   DIM v AS VARIANT
   DIM hr AS HRESULT = VariantChangeType(@v, @varSrc, 0, VT_CY)
   IF FAILED(hr) THEN EXIT CONSTRUCTOR
   m_cur = v.cyVal
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CCur (BYVAL pDispSrc AS IDispatch PTR)
   VarCyFromDisp(pDispSrc, GetThreadLocale(), @m_cur)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Sets the integer component.
' ========================================================================================
PRIVATE FUNCTION CCur.SetInteger (BYVAL nInteger AS LONGLONG) AS HRESULT
   ' // Check if within range
   IF nInteger < CY_MIN_INTEGER OR nInteger > CY_MAX_INTEGER THEN RETURN E_INVALIDARG
   ' // Add new integer part scaled by CY_SCALE
   IF m_cur.int64 = 0 THEN m_cur.int64 += nInteger * CY_SCALE : RETURN S_OK
   ' // Signs must match
   IF (m_cur.int64 < 0 AND nInteger > 0) OR _
      (m_cur.int64 > 0 AND nInteger < 0) THEN RETURN TYPE_E_TYPEMISMATCH
   DIM cyTemp AS CURRENCY
   ' // Get fractional part
   cyTemp.int64 = m_cur.int64 MOD CY_SCALE
   ' // Check if within range again
   IF (nInteger = CY_MAX_INTEGER AND cyTemp.int64 > 5807) OR _
      (nInteger = CY_MIN_INTEGER AND cyTemp.int64 < -5808) THEN RETURN TYPE_E_OUTOFBOUNDS
   ' // set to fractional part, wiping out integer part
   m_cur.int64 = cyTemp.int64
   ' // add new integer part scaled by CY_SCALE
   m_cur.int64 += nInteger * CY_SCALE
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the fractional component.
' Based on 4 digits. To set .2, pass 2000, to set .0002, pass a 2
' ========================================================================================
PRIVATE FUNCTION CCur.SetFraction (BYVAL nFraction AS SHORT) AS HRESULT
   ' // Check if within range
   IF (nFraction < CY_MIN_FRACTION OR nFraction > CY_MAX_FRACTION) THEN RETURN E_INVALIDARG
   IF m_cur.int64 = 0 THEN m_cur.int64 += nFraction : RETURN S_OK
   ' // Signs must match
   IF (m_cur.int64 < 0 AND nFraction > 0) OR _
      (m_cur.int64 > 0 AND nFraction < 0) THEN RETURN TYPE_E_TYPEMISMATCH
   DIM cyTemp AS CURRENCY
  '  // Get integer part, wiping out fractional part
   cyTemp.int64 = m_cur.int64 / CY_SCALE
   ' // Check if within range again
   IF (cyTemp.int64 = CY_MAX_INTEGER AND nFraction > 5807) OR _
      (cyTemp.int64 = CY_MIN_INTEGER AND nFraction < -5808) THEN RETURN TYPE_E_OUTOFBOUNDS
   ' // Scale to CY_SCALE
   m_cur.int64 = cyTemp.int64 * CY_SCALE
   m_cur.int64 += nFraction
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' CCur destructor
' ========================================================================================
PRIVATE DESTRUCTOR CCur
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assignment operators
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYREF cSrc AS CCUR)
   m_cur.int64 = cSrc.m_cur.int64
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL bSrc AS BYTE)
   VarCyFromI1(bSrc, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL ubSrc AS UBYTE)
   VarCyFromUI1(ubSrc, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL sSrc AS SHORT)
   VarCyFromI2(sSrc, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL usSrc AS USHORT)
   VarCyFromUI2(usSrc, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL lSrc AS LONG)
   VarCyFromI4(lSrc, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL ulSrc AS ULONG)
   VarCyFromUI4(ulSrc, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL fSrc AS SINGLE)
   VarCyFromR4(fSrc, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL dSrc AS DOUBLE)
   VarCyFromR8(dSrc, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL dSrc AS DECIMAL)
   VarCyFromDec(@dSrc, @m_cur)
END OPERATOR
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYREF pwszSrc AS WSTRING)
   VarCyFromStr(@pwszSrc, GetThreadLocale(), LOCALE_NOUSEROVERRIDE, @m_cur)
   m_Cur.int64 /= CY_SCALE
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL varSrc AS VARIANT)
   DIM v AS VARIANT
   DIM hr AS HRESULT = VariantChangeType(@v, @varSrc, 0, VT_CY)
   IF FAILED(hr) THEN EXIT OPERATOR
   m_cur = v.cyVal
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.LET (BYVAL pDispSrc AS IDispatch PTR)
   VarCyFromDisp(pDispSrc, GetThreadLocale(), @m_cur)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators.
' ========================================================================================
PRIVATE OPERATOR = (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
   RETURN (VarCyCmp(cur1.m_cur, cur2.m_cur) = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
   RETURN (VarCyCmp(cur1.m_cur, cur2.m_cur) <> VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
   RETURN (VarCyCmp(cur1.m_cur, cur2.m_cur) = VARCMP_LT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
   RETURN (VarCyCmp(cur1.m_cur, cur2.m_cur) = VARCMP_GT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
   DIM hr AS HRESULT = VarCyCmp(cur1.m_cur, cur2.m_cur)
   RETURN (hr = VARCMP_LT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
   DIM hr AS HRESULT = VarCyCmp(cur1.m_cur, cur2.m_cur)
   RETURN (hr = VARCMP_GT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Math operators.
' ========================================================================================
PRIVATE OPERATOR + (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   VarCyAdd(cur1.m_cur, cur2.m_cur, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYREF cur AS CCUR, BYVAL nValue AS LONG) AS CCUR
   DIM cy AS CURRENCY
   DIM c AS CURRENCY
   VarCyFromI4(nValue, @c)
   VarCyAdd(cur.m_cur, c, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYVAL nValue AS LONG, BYREF cur AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   DIM c AS CURRENCY
   VarCyFromI4(nValue, @c)
   VarCyAdd(c, cur.m_cur, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYREF cur AS CCUR, BYVAL nValue AS DOUBLE) AS CCUR
   DIM cy AS CURRENCY
   DIM c AS CURRENCY
   VarCyFromR8(nValue, @c)
   VarCyAdd(cur.m_cur, c, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYVAL nValue AS DOUBLE, BYREF cur AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   DIM c AS CURRENCY
   VarCyFromR8(nValue, @c)
   VarCyAdd(c, cur.m_cur, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.+= (BYREF cur AS CCUR)
   VarCyAdd(m_cur, cur.m_cur, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.+= (BYVAL nValue AS LONG)
   DIM c AS CURRENCY
   VarCyFromI4(nValue, @c)
   VarCyAdd(m_cur, c, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.+= (BYVAL nValue AS DOUBLE)
   DIM c AS CURRENCY
   VarCyFromR8(nValue, @c)
   VarCyAdd(m_cur, c, @m_cur)
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR - (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS CCUR
   DIM c AS CURRENCY
   VarCySub(cur1.m_cur, cur2.m_cur, @c)
   RETURN c
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF cur AS CCUR, BYVAL nValue AS LONG) AS CCUR
   DIM cy AS CURRENCY
   DIM c AS CURRENCY
   VarCyFromI4(nValue, @c)
   VarCySub(cur.m_cur, c, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYVAL nValue AS LONG, BYREF cur AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   DIM c AS CURRENCY
   VarCyFromI4(nValue, @c)
   VarCySub(c, cur.m_cur, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF cur AS CCUR, BYVAL nValue AS DOUBLE) AS CCUR
   DIM cy AS CURRENCY
   DIM c AS CURRENCY
   VarCyFromR8(nValue, @c)
   VarCySub(cur.m_cur, c, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYVAL nValue AS DOUBLE, BYREF cur AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   DIM c AS CURRENCY
   VarCyFromR8(nValue, @c)
   VarCySub(c, cur.m_cur, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.-= (BYREF cur AS CCUR)
   VarCySub(m_cur, cur.m_cur, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.-= (BYREF nValue AS LONG)
   DIM c AS CURRENCY
   VarCyFromI4(nValue, @c)
   VarCySub(m_cur, c, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.-= (BYREF nValue AS DOUBLE)
   DIM c AS CURRENCY
   VarCyFromR8(nValue, @c)
   VarCySub(m_cur, c, @m_cur)
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR - (BYREF cur AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   VarCyNeg(cur.m_cur, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR * (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   VarCyMul(cur1.m_cur, cur2.m_cur, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF cur AS CCUR, BYVAL nOperand AS LONG) AS CCUR
   DIM cy AS CURRENCY
   VarCyMulI4(cur.m_cur, nOperand, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYVAL nOperand AS LONG, BYREF cur AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   VarCyMulI4(cur.m_cur, nOperand, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF cur AS CCUR, BYVAL nOperand AS DOUBLE) AS CCUR
   DIM cy AS CURRENCY
   DIM c AS CCUR = nOperand
   VarCyMul(cur.m_cur, c, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYVAL nOperand AS DOUBLE, BYREF cur AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   DIM c AS CCUR = nOperand
   VarCyMul(cur.m_cur, c, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.*= (BYREF cur AS CCUR)
   VarCyMul(m_cur, cur.m_cur, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.*= (BYVAL nOperand AS LONG)
   VarCyMulI4(m_cur, nOperand, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.*= (BYVAL nOperand AS DOUBLE)
   DIM c AS CCUR = nOperand
   m_cur *= c
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR / (BYREF cur AS CCUR, BYVAL cOperand AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   IF cOperand.m_cur.int64 = 0 THEN RETURN cur.m_cur.int64
   cy.int64 = cur.m_cur.int64 / (cOperand.m_cur.int64 / CY_SCALE)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYREF cur AS CCUR, BYVAL nOperand AS LONG) AS CCUR
   DIM cy AS CURRENCY
   IF nOperand = 0 THEN RETURN cur.m_cur.int64
   cy.int64 = cur.m_cur.int64 / nOperand
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYREF cur AS CCUR, BYVAL nOperand AS DOUBLE) AS CCUR
   DIM cy AS CURRENCY
   IF nOperand = 0 THEN RETURN cur.m_cur.int64
   cy.int64 = cur.m_cur.int64 / nOperand
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYVAL nValue AS LONG, BYREF cur AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   IF nValue = 0 THEN RETURN CCUR(0, 0)
   cy = CCUR(nValue, 0) / (cur.m_cur.int64 / CY_SCALE)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYVAL nValue AS DOUBLE, BYREF cur AS CCUR) AS CCUR
   DIM cy AS CURRENCY
   IF nValue = 0 THEN RETURN CCUR(0, 0)
   cy = CCUR(nValue) / (cur.m_cur.int64 / CY_SCALE)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur./= (BYREF cOperand AS CCUR)
   IF cOperand.m_cur.int64 = 0 THEN EXIT OPERATOR
   m_cur.int64 /= (cOperand.m_cur.int64 / CY_SCALE)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur./= (BYVAL nOperand AS LONG)
   DIM cy AS CURRENCY
   IF nOperand = 0 THEN EXIT OPERATOR
   m_cur.int64 /= nOperand
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur./= (BYVAL nOperand AS DOUBLE)
   DIM cy AS CURRENCY
   IF nOperand = 0 THEN EXIT OPERATOR
   m_cur.int64 /= nOperand
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operators.
' ========================================================================================
PRIVATE OPERATOR CCur.CAST () AS CURRENCY
   OPERATOR = m_cur
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.CAST () AS DOUBLE
   OPERATOR = m_cur.int64 / CY_SCALE
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.CAST () AS STRING
   OPERATOR = STR(m_cur.int64 / CY_SCALE)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CCur.CAST () AS VARIANT
   DIM v AS VARIANT, cy AS CURRENCY
   VarCyFromR8(m_cur.int64 / CY_SCALE, @cy)
   v.vt = VT_CY
   v.cyVal = cy
   RETURN v
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF cur AS CCUR) AS CURRENCY PTR
   OPERATOR = VARPTR(cur.m_cur)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Rounds the currency to a specified number of decimal places.
' ========================================================================================
PRIVATE FUNCTION CCur.Round (BYVAL nDecimals AS LONG) AS CCUR
   DIM cy AS CURRENCY
   VarCyRound(m_cur, nDecimals, @cy)
   RETURN cy
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the integer component of the m_currency data member.
' ========================================================================================
PRIVATE FUNCTION CCur.GetInteger () AS LONGLONG
   IF m_cur.int64 THEN RETURN FIX(m_cur.int64 / CY_SCALE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the fractional component of the m_currency data member.
' ========================================================================================
PRIVATE FUNCTION CCur.GetFraction () AS SHORT
   IF m_cur.int64 THEN RETURN m_cur.int64 MOD CY_SCALE
END FUNCTION
' ========================================================================================

END NAMESPACE
