' ########################################################################################
' Microsoft Windows
' File: CInt96.inc
' Contents: 96-bit (12 bytes) integer data type.
' Compiler: FreeBasic 32 & 64-bit
' Written in 2018 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"

NAMESPACE Afx

' ========================================================================================
' CInt96 class
' ========================================================================================
TYPE CInt96

Public:
   m_int96 AS DECIMAL   ' // The underlying DECIMAL structure

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYREF cSrc AS CInt96)
   DECLARE CONSTRUCTOR (BYREF decSrc AS DECIMAL)
   DECLARE CONSTRUCTOR (BYVAL nInteger AS LONGINT)
   DECLARE CONSTRUCTOR (BYVAL nUInteger AS ULONGINT)
   DECLARE CONSTRUCTOR (BYREF wszSrc AS WSTRING)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR LET (BYREF cSrc AS CInt96)
   DECLARE OPERATOR LET (BYVAL nInteger AS LONGINT)
   DECLARE OPERATOR LET (BYVAL nUInteger AS ULONGINT)
   DECLARE OPERATOR LET (BYREF wszSrc AS WSTRING)
   DECLARE OPERATOR += (BYREF int96 AS CInt96)
   DECLARE OPERATOR += (BYVAL nValue AS LONGINT)
   DECLARE OPERATOR += (BYVAL nValue AS ULONGINT)
   DECLARE OPERATOR += (BYREF wszValue AS WSTRING)
   DECLARE OPERATOR -= (BYREF int96 AS CInt96)
   DECLARE OPERATOR -= (BYVAL nValue AS LONGINT)
   DECLARE OPERATOR -= (BYVAL nValue AS ULONGINT)
   DECLARE OPERATOR -= (BYREF wszValue AS WSTRING)
   DECLARE OPERATOR *= (BYREF int96 AS CInt96)
   DECLARE OPERATOR *= (BYVAL nValue AS LONGINT)
   DECLARE OPERATOR *= (BYVAL nValue AS ULONGINT)
   DECLARE OPERATOR *= (BYREF wszValue AS WSTRING)
   DECLARE OPERATOR /= (BYREF cOperand AS CInt96)
   DECLARE OPERATOR /= (BYVAL nValue AS LONGINT)
   DECLARE OPERATOR /= (BYVAL nValue AS ULONGINT)
   DECLARE OPERATOR /= (BYREF wszValue AS WSTRING)
   DECLARE OPERATOR CAST () AS BYTE
   DECLARE OPERATOR CAST () AS UBYTE
   DECLARE OPERATOR CAST () AS SHORT
   DECLARE OPERATOR CAST () AS USHORT
   DECLARE OPERATOR CAST () AS LONG
   DECLARE OPERATOR CAST () AS ULONG
   DECLARE OPERATOR CAST () AS LONGINT
   DECLARE OPERATOR CAST () AS ULONGINT
   DECLARE OPERATOR CAST () AS SINGLE
   DECLARE OPERATOR CAST () AS DOUBLE
   DECLARE OPERATOR CAST () AS CURRENCY
   DECLARE OPERATOR CAST () AS VARIANT
   DECLARE OPERATOR CAST () AS STRING
   DECLARE FUNCTION Abs_ () AS CInt96
   DECLARE FUNCTION ToVar () AS VARIANT
   DECLARE FUNCTION Sign () AS UBYTE
   DECLARE FUNCTION IsSigned () AS BOOLEAN
   DECLARE FUNCTION IsUnsigned () AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' CInt96 constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CInt96
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CInt96 (BYREF cSrc AS CInt96)
   m_int96 = cSrc.m_int96
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CInt96 (BYREF decSrc AS DECIMAL)
   IF decSrc.Scale THEN SetLastError(VarDecInt(@decSrc, @m_int96)) ELSE m_int96 = decSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CInt96 (BYVAL nInteger AS LONGINT)
   SetLastError(VarDecFromI8(nInteger, @m_int96))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CInt96 (BYVAL nUInteger AS ULONGINT)
   SetLastError(VarDecFromUI8(nuInteger, @m_int96))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CInt96 (BYREF wszSrc AS WSTRING)
   IF INSTR(wszSrc, ".") THEN SetLastError(E_INVALIDARG) : EXIT CONSTRUCTOR
   SetLastError(VarDecFromStr(@wszSrc, 0, 0, @m_int96))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CInt96 destructor
' ========================================================================================
PRIVATE DESTRUCTOR CInt96
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assignment operators
' ========================================================================================
PRIVATE OPERATOR CInt96.LET (BYREF cSrc AS CInt96)
   m_int96 = cSrc.m_int96
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.LET (BYVAL nInteger AS LONGINT)
   SetLastError(VarDecFromI8(nInteger, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.LET (BYVAL nUInteger AS ULONGINT)
   SetLastError(VarDecFromUI8(nuInteger, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.LET (BYREF wszSrc AS WSTRING)
   IF INSTR(wszSrc, ".") THEN SetLastError(E_INVALIDARG) : EXIT OPERATOR
   SetLastError(VarDecFromStr(@wszSrc, 0, 0, @m_int96))
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators.
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR = (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   RETURN (VarDecCmp(@int961.m_int96, @int962.m_int96) = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR <> (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   RETURN (VarDecCmp(@int961.m_int96, @int962.m_int96) <> VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR < (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   RETURN (VarDecCmp(@int961.m_int96, @int962.m_int96) = VARCMP_LT)
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR > (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   RETURN (VarDecCmp(@int961.m_int96, @int962.m_int96) = VARCMP_GT)
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR <= (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   DIM hr AS HRESULT = VarDecCmp(@int961.m_int96, @int962.m_int96)
   RETURN (hr = VARCMP_LT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR >= (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   DIM hr AS HRESULT = VarDecCmp(@int961.m_int96, @int962.m_int96)
   RETURN (hr = VARCMP_GT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Math operators.
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR + (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM hr AS LONG = VarDecAdd(@int961.m_int96, @int962.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int961
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR + (BYREF int96Src AS CInt96, BYREF wszValue AS WSTRING) AS CInt96
   DIM d AS DECIMAL
   DIM dec AS DECIMAL
   IF INSTR(wszValue, ".") THEN SetLastError(E_INVALIDARG) : RETURN int96Src
   DIM hr AS LONG = VarDecFromStr(@wszValue, 0, 0, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   hr = VarDecAdd(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
' Make sure that nValue is greater that -9223372036854775808 or it will be truncated.
' For values smaller than -9223372036854775808 use a string.
' ========================================================================================
PRIVATE OPERATOR + (BYREF int96Src AS CInt96, BYVAL nValue AS LONGINT) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nValue, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   hr = VarDecAdd(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYVAL nValue AS LONGINT, BYREF int96Src AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nValue, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nValue
   hr = VarDecAdd(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nValue
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
' Make sure that nValue does not exceed 18446744073709551615 or it will be truncated.
' For values greater than 18446744073709551615 use a string.
' ========================================================================================
PRIVATE OPERATOR + (BYREF int96Src AS CInt96, BYVAL nValue AS ULONGINT) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nValue, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   hr = VarDecAdd(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYVAL nValue AS ULONGINT, BYREF int96Src AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nValue, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nValue
   hr = VarDecAdd(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nValue
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96.+= (BYREF int96Src AS CInt96)
    SetLastError(VarDecAdd(@m_int96, @int96Src.m_int96, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
' Make sure that nValue is greater that -9223372036854775808 or it will be truncated.
' For values smaller than -9223372036854775808 use a string.
PRIVATE OPERATOR CInt96.+= (BYVAL nValue AS LONGINT)
   DIM dec AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nValue, @dec)
   IF hr <> S_OK THEN SetLastError(hr) ELSE SetLastError(VarDecAdd(@m_int96, @dec, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
' Make sure that nValue does not exceed 18446744073709551615 or it will be truncated.
' For values greater than 18446744073709551615 use a string.
PRIVATE OPERATOR CInt96.+= (BYVAL nValue AS ULONGINT)
   DIM dec AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nValue, @dec)
   IF hr <> S_OK THEN SetLastError(hr) ELSE SetLastError(VarDecAdd(@m_int96, @dec, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96.+= (BYREF wszValue AS WSTRING)
   DIM d AS DECIMAL
   DIM dec AS DECIMAL
   IF INSTR(wszValue, ".") THEN SetLastError(E_INVALIDARG) : EXIT OPERATOR
   DIM hr AS LONG = VarDecFromStr(@wszValue, 0, 0, @d)
   IF hr <> S_OK THEN SetLastError(hr) ELSE SetLastError(VarDecAdd(@m_int96, @d, @m_int96))
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR - (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM hr AS LONG = VarDecSub(@int961.m_int96, @int962.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int961
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
' Make sure that nValue is greater that -9223372036854775808 or it will be truncated.
' For values smaller than -9223372036854775808 use a string.
PRIVATE OPERATOR - (BYREF int96Src AS CInt96, BYVAL nValue AS LONGINT) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nValue, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   hr = VarDecSub(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
' Make sure that nValue does not exceed 18446744073709551615 or it will be truncated.
' For values greater than 18446744073709551615 use a string.
PRIVATE OPERATOR - (BYREF int96Src AS CInt96, BYVAL nValue AS ULONGINT) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nValue, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   hr = VarDecSub(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
' Make sure that nValue is greater that -9223372036854775808 or it will be truncated.
' For values smaller than -9223372036854775808 use a string.
PRIVATE OPERATOR - (BYVAL nValue AS LONGINT, BYREF int96Src AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nValue, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nValue
   hr = VarDecSub(@d, @int96Src.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nValue
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
' Make sure that nValue does not exceed 18446744073709551615 or it will be truncated.
' For values greater than 18446744073709551615 use a string.
PRIVATE OPERATOR - (BYVAL nValue AS ULONGINT, BYREF int96Src AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nValue, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nValue
   hr = VarDecSub(@d, @int96Src.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nValue
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR - (BYREF int96Src AS CInt96, BYREF wszValue AS WSTRING) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   IF INSTR(wszValue, ".") THEN SetLastError(E_INVALIDARG) : RETURN int96Src
   DIM hr AS LONG = VarDecFromStr(@wszValue, 0, 0, @d)
   IF VarDecSub(@int96Src.m_int96, @d, @dec) <> S_OK THEN RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96.-= (BYREF int96Src AS CInt96)
   SetLastError(VarDecSub(@m_int96, @int96Src.m_int96, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
' Make sure that nValue does not exceed 18446744073709551615 or it will be truncated.
' For values greater than 18446744073709551615 use a string.
PRIVATE OPERATOR CInt96.-= (BYVAL nValue AS LONGINT)
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nValue, @d)
   IF hr <> S_OK THEN SetLastError(hr) ELSE SetLastError(VarDecSub(@m_int96, @d, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
' Make sure that nValue does not exceed 18446744073709551615 or it will be truncated.
' For values greater than 18446744073709551615 use a string.
PRIVATE OPERATOR CInt96.-= (BYVAL nValue AS ULONGINT)
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nValue, @d)
   IF hr <> S_OK THEN SetLastError(hr) ELSE SetLastError(VarDecSub(@m_int96, @d, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96.-= (BYREF wszValue AS WSTRING)
   DIM d AS DECIMAL
   DIM dec AS DECIMAL
   IF INSTR(wszValue, ".") THEN EXIT OPERATOR
   DIM hr AS LONG = VarDecFromStr(@wszValue, 0, 0, @d)
   IF hr <> S_OK THEN SetLastError(hr) ELSE SetLastError(VarDecSub(@m_int96, @d, @m_int96))
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR - (BYREF int96Src AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   SetLastError(VarDecNeg(@int96Src.m_int96, @dec))
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR * (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM hr AS LONG = VarDecMul(@int961.m_int96, @int962.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int961
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR * (BYVAL nOperand AS LONGINT, BYREF int96Src AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nOperand, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nOperand
   hr = VarDecMul(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nOperand
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR * (BYREF int96Src AS CInt96, BYVAL nOperand AS LONGINT) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nOperand, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   hr = VarDecMul(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR * (BYREF int96Src AS CInt96, BYVAL nOperand AS ULONGINT) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr  AS LONG = VarDecFromUI8(nOperand, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   hr = VarDecMul(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR * (BYVAL nOperand AS ULONGINT, BYREF int96Src AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nOperand, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nOperand
   hr = VarDecMul(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nOperand
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR * (BYREF int96Src AS CInt96, BYREF wszValue AS WSTRING) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   IF INSTR(wszValue, ".") THEN SetLastError(E_INVALIDARG) : RETURN int96Src
   DIM hr AS LONG = VarDecFromStr(@wszValue, 0, 0, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   hr = VarDecMul(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96.*= (BYREF int96Src AS CInt96)
   SetLasterror(VarDecMul(@m_int96, @int96Src.m_int96, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96.*= (BYVAL nOperand AS LONGINT)
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nOperand, @d)
   IF hr THEN SetLastError(hr) ELSE SetLastError(VarDecMul(@m_int96, @d, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96.*= (BYVAL nOperand AS ULONGINT)
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nOperand, @d)
   IF hr <> S_OK THEN SetLasterror(hr) ELSE SetLastError(VarDecMul(@m_int96, @d, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96.*= (BYREF wszValue AS WSTRING)
   DIM d AS DECIMAL
   DIM dec AS DECIMAL
   IF INSTR(wszValue, ".") THEN SetLastError(E_INVALIDARG) : EXIT OPERATOR
   DIM hr AS LONG = VarDecFromStr(@wszValue, 0, 0, @d)
   IF hr <> S_OK THEN SetLastError(hr) ELSE SetLastError(VarDecMul(@m_int96, @d, @m_int96))
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR / (BYREF int96Src AS CInt96, BYREF cOperand AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM hr AS LONG = VarDecDiv(@int96Src.m_int96, @cOperand.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   IF dec.Scale THEN hr = VarDecInt(@dec, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR / (BYREF int96Src AS CInt96, BYVAL nOperand AS LONGINT) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nOperand, @d)
   hr = VarDecDiv(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   IF dec.Scale THEN hr = VarDecInt(@dec, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR / (BYVAL nOperand AS LONGINT, BYREF int96Src AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nOperand, @d)
   hr = VarDecDiv(@d, @int96Src.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nOperand
   IF dec.Scale THEN hr = VarDecInt(@dec, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR / (BYREF int96Src AS CInt96, BYVAL nOperand AS ULONGINT) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nOperand, @d)
   hr = VarDecDiv(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   IF dec.Scale THEN hr = VarDecInt(@dec, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR / (BYVAL nOperand AS ULONGINT, BYREF int96Src AS CInt96) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nOperand, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nOperand
   hr = VarDecDiv(@d, @int96Src.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nOperand
   IF dec.Scale THEN hr = VarDecInt(@dec, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN nOperand
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR / (BYREF int96Src AS CInt96, BYREF wszValue AS WSTRING) AS CInt96
   DIM dec AS DECIMAL
   DIM d AS DECIMAL
   IF INSTR(wszValue, ".") THEN SetLastError(E_INVALIDARG) : RETURN int96Src
   DIM hr AS LONG = VarDecFromStr(@wszValue, 0, 0, @d)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   hr = VarDecDiv(@int96Src.m_int96, @d, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   IF dec.Scale THEN hr = VarDecInt(@dec, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96./= (BYREF cOperand AS CInt96)
   DIM hr AS LONG = VarDecDiv(@m_int96, @cOperand.m_int96, @m_int96)
   IF hr <> S_OK THEN SetLastError(hr) : EXIT OPERATOR
   IF m_int96.Scale THEN SetLastError(VarDecInt(@m_int96, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96./= (BYVAL nOperand AS LONGINT)
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromI8(nOperand, @d)
   IF hr <> S_OK THEN SetLastError(hr) : EXIT OPERATOR
   hr = VarDecDiv(@m_int96, @d, @m_int96)
   IF hr <> S_OK THEN SetLastError(hr) : EXIT OPERATOR
   IF m_int96.Scale THEN SetLastError(VarDecInt(@m_int96, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96./= (BYVAL nOperand AS ULONGINT)
   DIM d AS DECIMAL
   DIM hr AS LONG = VarDecFromUI8(nOperand, @d)
   IF hr <> S_OK THEN SetLastError(hr) : EXIT OPERATOR
   hr = VarDecDiv(@m_int96, @d, @m_int96)
   IF hr <> S_OK THEN SetLastError(hr) : EXIT OPERATOR
   IF m_int96.Scale THEN SetLastError(VarDecInt(@m_int96, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' *** Ok ***
PRIVATE OPERATOR CInt96./= (BYREF wszValue AS WSTRING)
   DIM d AS DECIMAL
   DIM dec AS DECIMAL
   IF INSTR(wszValue, ".") THEN SetLastError(E_INVALIDARG) : EXIT OPERATOR
   DIM hr AS LONG = VarDecFromStr(@wszValue, 0, 0, @d)
   IF hr <> S_OK THEN SetLastError(hr) : EXIT OPERATOR
   hr = VarDecDiv(@m_int96, @d, @m_int96)
   IF hr <> S_OK THEN SetLastError(hr) : EXIT OPERATOR
   IF m_int96.Scale THEN SetLastError(VarDecInt(@m_int96, @m_int96))
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operators.
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS BYTE
   DIM nValue AS BYTE
   SetLastError(VarI1FromDec(@m_int96, @nValue))
   OPERATOR = nValue
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS UBYTE
   DIM nValue AS UBYTE
   SetLastError(VarUI1FromDec(@m_int96, @nValue))
   OPERATOR = nValue
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS SHORT
   DIM nValue AS SHORT
   SetLastError(VarI2FromDec(@m_int96, @nValue))
   OPERATOR = nValue
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS USHORT
   DIM nValue AS USHORT
   SetLastError(VarUI2FromDec(@m_int96, @nValue))
   OPERATOR = nValue
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS LONG
   DIM nValue AS LONG
   SetLastError(VarI4FromDec(@m_int96, @nValue))
   OPERATOR = nValue
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS ULONG
   DIM nValue AS ULONG
   SetLastError(VarUI4FromDec(@m_int96, @nValue))
   OPERATOR = nValue
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS LONGINT
   DIM nValue AS LONGINT
   SetLastError(VarI8FromDec(@m_int96, @nValue))
   OPERATOR = nValue
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS ULONGINT
   DIM nValue AS LONGINT
   SetLastError(VarUI8FromDec(@m_int96, @nValue))
   OPERATOR = nValue
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS SINGLE
   DIM nValue AS SINGLE
   SetLastError(VarR4FromDec(@m_int96, @nValue))
   OPERATOR = nValue
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS DOUBLE
   DIM nValue AS DOUBLE
   SetLastError(VarR8FromDec(@m_int96, @nValue))
   OPERATOR = nValue
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS CURRENCY
   DIM cy AS CURRENCY
   SetLastError(VarCyFromDec(@m_int96, @cy))
   OPERATOR = cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS VARIANT
   DIM v AS VARIANT
   v.vt = VT_DECIMAL
   v.decVal = m_int96
   RETURN v
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS STRING
   DIM bstrOut AS BSTR
   VarBstrFromDec(@m_int96, 0, 0, @bstrOut)
   DIM s AS STRING = *cast(WSTRING PTR, bstrOut)
   SysFreeString(bstrOut)
   OPERATOR = s
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the address of the underlying decimal value.
' ========================================================================================
PRIVATE OPERATOR * (BYREF int96 AS CInt96) AS DECIMAL PTR
   OPERATOR = VARPTR(int96.m_int96)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Retrieves the absolute value of a CInt96 data type.
' ========================================================================================
PRIVATE FUNCTION CInt96.Abs_ () AS CInt96
   DIM dec AS DECIMAL
   SetLastError(VarDecAbs(@m_int96, @dec))
   RETURN dec
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts the currency to a VT_CY variant.
' To allow to assign the decimal to a CVAR variable
' DIM dec AS CInt96 = 12345.1234
' DIM cv AS CVAR = dec.ToVar
' ========================================================================================
PRIVATE FUNCTION CInt96.ToVar () AS VARIANT
   DIM v AS VARIANT
   v.vt = VT_DECIMAL
   v.decVal = m_int96
   RETURN v
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the sign of this decimal number.
' Returns 0 if it is not signed of &h80 (128) if it is signed.
' ========================================================================================
PRIVATE FUNCTION CInt96.Sign () AS UBYTE
   RETURN m_int96.sign
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns true if this number is signed or false otherwise.
' ========================================================================================
PRIVATE FUNCTION CInt96.IsSigned () AS BOOLEAN
   IF m_int96.sign THEN RETURN TRUE ELSE RETURN FALSE
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns true if this number is unsigned or false otherwise.
' ========================================================================================
PRIVATE FUNCTION CInt96.IsUnsigned () AS BOOLEAN
   IF m_int96.sign = 0 THEN RETURN TRUE ELSE RETURN FALSE
END FUNCTION
' ========================================================================================

END NAMESPACE
