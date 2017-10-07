' ########################################################################################
' Microsoft Windows
' File: COleDateTime.inc
' Contents: Date and time classes
' Compiler: FreeBasic 32 & 64-bit
' Based on the ATL classes COleDateTimeSpan and COleDateTime.
' Copyright (C) Microsoft Corporation
' Written in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "crt/math.bi"
#include once "win/ole2.bi"
#include once "Afx/CWStr.inc"
USING Afx

NAMESPACE Afx

' ########################################################################################
'                            *** ColeDateTimeSpan Class ***
' ########################################################################################

' ========================================================================================
' Represents a relative time, a time span.
' ========================================================================================
TYPE COleDateTimeSpan

Public:
   m_span AS DOUBLE
   m_status AS BOOLEAN
   OLE_DATETIME_HALFSECOND AS DOUBLE = 1.0 / (2.0 * (60.0 * 60.0 * 24.0))
   maxDaysInSpan AS LONG = 3615897

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL dblSpanSrc AS DOUBLE)
   DECLARE CONSTRUCTOR (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS DOUBLE
   DECLARE OPERATOR LET (BYVAL dblSpanSrc AS DOUBLE)
   DECLARE OPERATOR += (BYREF dateSpan AS COleDateTimeSpan)
   DECLARE OPERATOR -= (BYREF dateSpan AS COleDateTimeSpan)
   DECLARE SUB CheckRange
   DECLARE FUNCTION GetStatus () AS BOOLEAN
   DECLARE SUB SetStatus (BYVAL nStatus AS BOOLEAN)
   DECLARE SUB SetDateTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   DECLARE FUNCTION GetTotalDays () AS DOUBLE
   DECLARE FUNCTION GetTotalHours () AS DOUBLE
   DECLARE FUNCTION GetTotalMinutes () AS DOUBLE
   DECLARE FUNCTION GetTotalSeconds () AS DOUBLE
   DECLARE FUNCTION GetDays () AS LONG
   DECLARE FUNCTION GetHours () AS LONG
   DECLARE FUNCTION GetMinutes () AS LONG
   DECLARE FUNCTION GetSeconds () AS LONG

END TYPE
' ========================================================================================

' ========================================================================================
' COleDateTimeSpan constructor
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTimeSpan
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTimeSpan (BYVAL dblSpanSrc AS DOUBLE)
   m_span = dblSpanSrc
   m_status = TRUE
   this.CheckRange
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR ColeDateTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   this.SetDateTimeSpan(lDays, nHours, nMins, nSecs)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' COleDateTimeSpan destructor
' ========================================================================================
PRIVATE DESTRUCTOR COleDateTimeSpan
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR COleDateTimeSpan.CAST () AS DOUBLE
   IF m_status = TRUE THEN RETURN m_span
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR COleDateTimeSpan.LET (BYVAL dblSpanSrc AS DOUBLE)
   m_span = dblSpanSrc
   m_status = TRUE
   this.CheckRange
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR COleDateTimeSpan.+= (BYREF dateSpan AS COleDateTimeSpan)
   IF m_status = FALSE THEN EXIT OPERATOR
   m_span = m_span + dateSpan.m_span
   this.CheckRange
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR COleDateTimeSpan.-= (BYREF dateSpan AS COleDateTimeSpan)
   IF m_status = FALSE THEN EXIT OPERATOR
   m_span = m_span + dateSpan.m_span
   this.CheckRange
END OPERATOR
' ========================================================================================

' ========================================================================================
' Checks if the span value is between ranges.
' ========================================================================================
PRIVATE SUB COleDateTimeSpan.CheckRange
   IF m_span < -maxDaysInSpan OR m_span > maxDaysInSpan THEN m_Status = FALSE
END SUB
' ========================================================================================
' ========================================================================================
' Gets/sets the status (validity) of a given COleDateTime object.
' ========================================================================================
PRIVATE FUNCTION COleDateTimeSpan.GetStatus () AS BOOLEAN
   RETURN m_status
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB COleDateTimeSpan.SetStatus (BYVAL nStatus AS BOOLEAN)
   m_status = nStatus
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB COleDateTimeSpan.SetDateTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   ' // Set date span by breaking into fractional days (all input ranges valid)
   m_span = lDays + (nHours)/24 + (nMins)/(24*60) + (nSecs)/(24*60*60)
   m_status = TRUE
   this.CheckRange
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION COleDateTimeSpan.GetTotalDays () AS DOUBLE
   IF m_status = FALSE THEN RETURN 0
   RETURN CLNGINT(m_span + IIF(m_span < 0, -OLE_DATETIME_HALFSECOND, OLE_DATETIME_HALFSECOND))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleDateTimeSpan.GetTotalHours () AS DOUBLE
   IF m_status = FALSE THEN RETURN 0
   DIM d AS DOUBLE = IIF(m_span < 0, -OLE_DATETIME_HALFSECOND, OLE_DATETIME_HALFSECOND)
   RETURN CLNGINT((m_span + d) * 24)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleDateTimeSpan.GetTotalMinutes () AS DOUBLE
   IF m_status = FALSE THEN RETURN 0
   DIM d AS DOUBLE = IIF(m_span < 0, -OLE_DATETIME_HALFSECOND, OLE_DATETIME_HALFSECOND)
   RETURN CLNGINT((m_span + d) * (24 * 60))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleDateTimeSpan.GetTotalSeconds () AS DOUBLE
   IF m_status = FALSE THEN RETURN 0
   DIM d AS DOUBLE = IIF(m_span < 0, -OLE_DATETIME_HALFSECOND, OLE_DATETIME_HALFSECOND)
   RETURN CLNGINT((m_span + d) * (24 * 60 * 60))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleDateTimeSpan.GetDays () AS LONG
   RETURN this.GetTotalDays
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleDateTimeSpan.GetHours () AS LONG
   RETURN this.GetTotalHours
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleDateTimeSpan.GetMinutes () AS LONG
   RETURN this.GetTotalMinutes
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleDateTimeSpan.GetSeconds () AS LONG
   RETURN this.GetTotalSeconds
END FUNCTION
' ========================================================================================


' ########################################################################################
'                               *** ColeDateTime Class ***
' ########################################################################################

'enum DateTimeStatus
'	{
'		error = -1,
'		valid = 0,
'		invalid = 1,    // Invalid date (out of range, etc.)
'		null = 2,       // Literally has no value
'};

#ifndef DBTIMESTAMP
TYPE DBTIMESTAMP
   year AS SHORT
   month AS USHORT
   day AS USHORT
   hour AS USHORT
   minute AS USHORT
   second AS USHORT
   fraction AS ULONG
END TYPE
#endif

' ========================================================================================
' Encapsulates the DATE data type that is used in OLE automation.
' The COleDateTime class handles dates from January 1, 100, through December 31, 9999.
' The COleDateTime class uses the Gregorian calendar; it does not support Julian dates.
' COleDateTime ignores Daylight Saving Time.
' ========================================================================================
TYPE COleDateTime

Public:
   m_dt AS DATE_
   m_status AS BOOLEAN
   OLE_DATETIME_HALFSECOND AS DOUBLE = 1.0 / (2.0 * (60.0 * 60.0 * 24.0))

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, BYVAL nHour AS WORD, BYVAL nMin AS WORD, BYVAL nSec AS WORD)
   DECLARE CONSTRUCTOR (BYVAL timeSrc AS LONGINT)
   DECLARE CONSTRUCTOR (BYVAL dtSrc AS DATE_)
   DECLARE CONSTRUCTOR (BYREF varSrc AS VARIANT)
   DECLARE CONSTRUCTOR (BYREF systimeSrc AS SYSTEMTIME)
   DECLARE CONSTRUCTOR (BYREF filetimeSrc AS FILETIME)
   DECLARE CONSTRUCTOR (BYVAL wDosDate AS DWORD, BYVAL wDosTime AS WORD)
   DECLARE CONSTRUCTOR (BYREF dbts AS DBTIMESTAMP)
   DECLARE CONSTRUCTOR (BYREF ud AS UDATE)
   DECLARE CONSTRUCTOR (BYREF dt AS COleDateTime)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS DATE_
   DECLARE OPERATOR LET (BYVAL timeSrc AS LONGINT)
   DECLARE OPERATOR LET (BYVAL dtSrc AS DATE_)
   DECLARE OPERATOR LET (BYREF varSrc AS VARIANT)
   DECLARE OPERATOR LET (BYREF systimeSrc AS SYSTEMTIME)
   DECLARE OPERATOR LET (BYREF filetimeSrc AS FILETIME)
   DECLARE OPERATOR LET (BYREF dbts AS DBTIMESTAMP)
   DECLARE OPERATOR LET (BYREF ud AS UDATE)
   DECLARE OPERATOR LET (BYREF dt AS COleDateTime)
   DECLARE OPERATOR += (BYREF dateSpan AS COleDateTimeSpan)
   DECLARE OPERATOR -= (BYREF dateSpan AS COleDateTimeSpan)
   DECLARE FUNCTION GetStatus () AS BOOLEAN
   DECLARE SUB SetStatus (BYVAL nStatus AS BOOLEAN)
   DECLARE FUNCTION SetDateTime (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, BYVAL nHour AS WORD, BYVAL nMin AS WORD, BYVAL nSec AS WORD) AS LONG
   DECLARE FUNCTION SetDate (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD) AS LONG
   DECLARE FUNCTION SetTime (BYVAL nHour AS WORD, BYVAL nMin AS WORD, BYVAL nSec AS WORD) AS LONG
   DECLARE FUNCTION GetCurrentTime () AS COleDateTime
   DECLARE FUNCTION GetAsSystemTime (BYREF sysTime AS SYSTEMTIME) AS BOOLEAN
   DECLARE FUNCTION GetYear () AS WORD
   DECLARE FUNCTION GetMonth () AS WORD
   DECLARE FUNCTION GetDay () AS WORD
   DECLARE FUNCTION GetHour () AS WORD
   DECLARE FUNCTION GetMinute () AS WORD
   DECLARE FUNCTION GetSecond () AS WORD
   DECLARE FUNCTION GetDayOfWeek () AS WORD
   DECLARE FUNCTION GetAsUDATE (BYREF ud AS UDATE) AS BOOLEAN
   DECLARE FUNCTION GetDayOfYear () AS WORD
   DECLARE FUNCTION GetAsDBTIMESTAMP (BYREF dbts AS DBTIMESTAMP) AS BOOLEAN
   DECLARE FUNCTION DoubleFromDate (BYVAL dt AS DATE_) AS DOUBLE
   DECLARE FUNCTION DateFromDouble (BYVAL f AS DOUBLE) AS DATE_

Private:
   DECLARE SUB CheckRange


'Format
'ParseDateTime
'operator =
'operator +
'operator *

END TYPE
' ========================================================================================

' ========================================================================================
' COleDateTime constructor
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, BYVAL nHour AS WORD, BYVAL nMin AS WORD, BYVAL nSec AS WORD)
   this.SetDateTime(nYear, nMonth, nDay, nHour, nMin, nSec)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime (BYVAL timeSrc AS LONGINT)
   this = timeSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime (BYVAL dtSrc AS DATE_)
   this = dtSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime (BYREF varSrc AS VARIANT)
   this = varSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime (BYREF systimeSrc AS SYSTEMTIME)
'   m_status = SystemTimeToVariantTime(@systimeSrc, @m_dt)
   this = systimeSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime (BYREF filetimeSrc AS FILETIME)
   this = filetimeSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime (BYVAL wDosDate AS DWORD, BYVAL wDosTime AS WORD)
   m_status = DosDateTimeToVariantTime(wDosDate, wDosTime, @m_dt)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime (BYREF dbts AS DBTIMESTAMP)
   this = dbts
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime (BYREF ud aS UDATE)
   this = ud
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleDateTime (BYREF dt AS COleDateTime)
   this = dt
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' COleDateTime destructor
' ========================================================================================
PRIVATE DESTRUCTOR COleDateTime
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR COleDateTime.CAST () AS DATE_
   IF m_status = TRUE THEN RETURN m_dt
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR COleDateTime.LET (BYVAL timeSrc AS LONGINT)
   m_dt = timeSrc
   m_status = TRUE
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR COleDateTime.LET (BYVAL dtSrc AS DATE_)
   m_dt = dtSrc
   m_status = TRUE
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR COleDateTime.LET (BYREF varSrc AS VARIANT)
   IF varSrc.vt <> VT_DATE THEN
      DIM varDest AS VARIANT
      m_status = FALSE
      IF SUCCEEDED(VariantChangeType(@varDest, @varSrc, 0, VT_DATE)) THEN
         m_dt = varDest.date
         m_status = TRUE
      END IF
   ELSE
      m_dt = varSrc.date
      m_status = TRUE
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR COleDateTime.LET (BYREF systimeSrc AS SYSTEMTIME)
   m_status = SystemTimeToVariantTime(@systimeSrc, @m_dt)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR COleDateTime.LET (BYREF filetimeSrc AS FILETIME)
   DIM ft AS FILETIME, st AS SYSTEMTIME
   IF FileTimeToLocalFileTime(@filetimesrc, @ft) THEN
      IF FileTimeToSystemTime(@ft, @st) THEN
         m_status = SystemTimeToVariantTime(@st, @m_dt)
      END IF
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR COleDateTime.LET (BYREF dbts AS DBTIMESTAMP)
	DIM st AS SYSTEMTIME
	st.wYear = dbts.year
	st.wMonth = dbts.month
	st.wDay = dbts.day
	st.wHour = dbts.hour
	st.wMinute = dbts.minute
	st.wSecond = dbts.second
   m_status = SystemTimeToVariantTime(@st, @m_dt)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR COleDateTime.LET (BYREF ud AS UDATE)
   m_Status = VarDateFromUdate(@ud, 0, @m_dt)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR COleDateTime.LET (BYREF dt AS COleDateTime)
   m_dt = dt.m_dt
   m_status = dt.m_status
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR COleDateTime.+= (BYREF dateSpan AS COleDateTimeSpan)
   IF m_status = FALSE THEN EXIT OPERATOR
   m_dt = this.DateFromDouble(this.DoubleFromDate(m_dt) + dateSpan.m_Span)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR COleDateTime.-= (BYREF dateSpan AS COleDateTimeSpan)
   IF m_status = FALSE THEN EXIT OPERATOR
   m_dt = this.DateFromDouble(this.DoubleFromDate(m_dt) - dateSpan.m_Span)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Gets/sets the status (validity) of a given COleDateTime object.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetStatus () AS BOOLEAN
   RETURN m_status
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB COleDateTime.SetStatus (BYVAL nStatus AS BOOLEAN)
   m_status = nStatus
END SUB
' ========================================================================================

' ========================================================================================
' Sets the date and time of this COleDateTime object.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.SetDateTime (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, BYVAL nHour AS WORD, BYVAL nMin AS WORD, BYVAL nSec AS WORD) AS LONG
   DIM st AS SYSTEMTIME = (nYear, nMonth, 0, nDay, nHour, nMin, nSec, 0)
   m_status = SystemTimeToVariantTime(@st, @m_dt)
   RETURN m_status
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION COleDateTime.SetDate (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD) AS LONG
   DIM st AS SYSTEMTIME = (nYear, nMonth, 0, nDay, 0, 0, 0, 0)
   m_status = SystemTimeToVariantTime(@st, @m_dt)
   RETURN m_status
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleDateTime.SetTime (BYVAL nHour AS WORD, BYVAL nMin AS WORD, BYVAL nSec AS WORD) AS LONG
   ' // Set date to zero date - 12/30/1899
   DIM st AS SYSTEMTIME = (1899, 12, 0, 30, nHour, nMin, nSec, 0)
   m_status = SystemTimeToVariantTime(@st, @m_dt)
   RETURN m_status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current date/time value.
' Usage example: DIM ct AS COleDateTime = COleDateTime().GetCurrentTime()
' --or--
' DIM ct AS COleDateTime
' ct = ct.GetCurrentTime()
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetCurrentTime () AS COleDateTime
   DIM st AS SYSTEMTIME, d AS DOUBLE
   GetLocalTime @st
   m_status = SystemTimeToVariantTime(@st, @d)
   RETURN COleDateTime(d)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Call this method to obtain the time in the COleDateTime object as a SYSTEMTIME data structure.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetAsSystemTime (BYREF sysTime AS SYSTEMTIME) AS BOOLEAN
   IF m_status = FALSE THEN RETURN m_status
   m_Status = VariantTimeToSystemTime(m_dt, @sysTime)
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the year represented by this date/time value.
' Valid return values range between 100 and 9999, which includes the century.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetYear() AS WORD
   DIM st AS SYSTEMTIME
   IF this.GetAsSystemTime(st) THEN RETURN st.wYear
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the month represented by this date/time value.
' Valid return values range between 1 and 12.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetMonth() AS WORD
   DIM st AS SYSTEMTIME
   IF this.GetAsSystemTime(st) THEN RETURN st.wMonth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the day of the month represented by this date/time value.
' Valid return values range between 1 and 31.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetDay() AS WORD
   DIM st AS SYSTEMTIME
   IF this.GetAsSystemTime(st) THEN RETURN st.wDay
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the hour represented by this date/time value.
' Valid return values range between 0 and 23.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetHour() AS WORD
   DIM st AS SYSTEMTIME
   IF this.GetAsSystemTime(st) THEN RETURN st.wHour
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the minute represented by this date/time value.
' Valid return values range between 0 and 59.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetMinute() AS WORD
   DIM st AS SYSTEMTIME
   IF this.GetAsSystemTime(st) THEN RETURN st.wMinute
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the second represented by this date/time value.
' Valid return values range between 0 and 59.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetSecond() AS WORD
   DIM st AS SYSTEMTIME
   IF this.GetAsSystemTime(st) THEN RETURN st.wSecond
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the day of the month represented by this date/time value.
' Valid return values range between 1 and 7, where 1=Sunday, 2=Monday, and so on.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetDayOfWeek () AS WORD
   DIM st AS SYSTEMTIME
   IF this.GetAsSystemTime(st) THEN RETURN st.wDayOfWeek + 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Call this method to obtain the time in the COleDateTime object as a SYSTEMTIME data structure.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetAsUDATE (BYREF ud AS UDATE) AS BOOLEAN
   RETURN SUCCEEDED(VarUdateFromDate(m_dt, 0, @ud))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the day of the month represented by this date/time value.
' Valid return values range between 1 and 366, where January 1 = 1.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetDayOfYear () AS WORD
   DIM ud AS UDATE
   IF this.GetAsUDATE(ud) THEN RETURN ud.wDayOfYear
END FUNCTION
' ========================================================================================

' ========================================================================================
' Call this method to obtain the time in the COleDateTime object as a DBTIMESTAMP data structure.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.GetAsDBTIMESTAMP(BYREF dbts AS DBTIMESTAMP) AS BOOLEAN
	DIM ud AS UDATE
   IF VarUdateFromDate(m_dt, 0, @ud) <> S_OK THEN RETURN FALSE
	dbts.year = ud.st.wYear
	dbts.month = ud.st.wMonth
	dbts.day = ud.st.wDay
	dbts.hour = ud.st.wHour
	dbts.minute = ud.st.wMinute
	dbts.second = ud.st.wSecond
	dbts.fraction = 0
   return TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Internal procedure.
' Converts a DATE_ to a double.
' ========================================================================================
PRIVATE FUNCTION ColeDateTime.DoubleFromDate (BYVAL dt AS DATE_) AS DOUBLE
   ' // We treat it as positive from -OLE_DATETIME_HALFSECOND because of numeric errors
   ' // If value is positive it doesn't need conversion
   IF dt > -OLE_DATETIME_HALFSECOND THEN RETURN dt
   ' // If negative, must convert since negative dates not continuous
   ' // (examples: -1.25 to -.75, -1.50 to -.50, -1.75 to -.25)
   DIM fTemp AS DOUBLE = ceil(dt)
   RETURN fTemp - (dt - fTemp)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Internal procedure.
' Converts a doible to a DATE_.
' ========================================================================================
PRIVATE FUNCTION COleDateTime.DateFromDouble (BYVAL f AS DOUBLE) AS DATE_
   ' // We treat it as positive from -OLE_DATETIME_HALFSECOND because of numeric errors
   ' // If value is positive it doesn't need conversion
   IF f > -OLE_DATETIME_HALFSECOND THEN RETURN f
   ' // If negative, must convert since negative dates not continuous
   ' // (examples: -1.25 to -.75, -1.50 to -.50, -1.75 to -.25)
   DIM fTemp AS DOUBLE = floor(f)   ' // fTemp is now whole part
   RETURN fTemp + (fTemp - f)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Internal procedure.
' Checks if the date is between the maximum and minimum value.
' ========================================================================================
PRIVATE SUB COleDateTime.CheckRange
   IF m_dt > VTDATEGRE_MAX OR m_dt < VTDATEGRE_MIN THEN m_Status = FALSE
END SUB
' ========================================================================================

END NAMESPACE

' ########################################################################################
'                                *** Global operators ***
' ########################################################################################

USING AFx

' ========================================================================================
PRIVATE OPERATOR - (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS COleDateTimeSpan
   IF dt1.m_status = FALSE OR dt2.m_status = FALSE THEN EXIT OPERATOR
   DIM d AS DOUBLE = dt1.DoubleFromDate(dt1.m_dt) - dt2.DoubleFromDate(dt2.m_dt)
   OPERATOR = COleDateTimeSpan(d)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYREF dt AS COleDateTime, BYREF dateSpan AS COleDateTimeSpan) AS COleDateTime
   IF dt.m_status = FALSE OR dateSpan.m_status = FALSE THEN EXIT OPERATOR
   DIM d AS DOUBLE = dt.DateFromDouble(dt.DoubleFromDate(dt.m_dt) + dateSpan.m_span)
   OPERATOR = COleDateTime(d)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF dt AS COleDateTime, BYREF dateSpan AS COleDateTimeSpan) AS COleDateTime
   IF dt.m_status = FALSE OR dateSpan.m_status = FALSE THEN EXIT OPERATOR
   DIM d AS DOUBLE = dt.DateFromDouble(dt.DoubleFromDate(dt.m_dt) - dateSpan.m_span)
   OPERATOR = COleDateTime(d)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR = (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
   IF dt1.m_status = FALSE OR dt2.m_status = FALSE THEN EXIT OPERATOR
   ' // it has to be in precision range to say that it is equal         
   IF (dt1.m_dt + dt1.OLE_DATETIME_HALFSECOND > dt2.m_dt AND _
       dt1.m_dt - dt1.OLE_DATETIME_HALFSECOND < dt2.m_dt) THEN OPERATOR = TRUE
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
   IF dt1.m_status = FALSE OR dt2.m_status = FALSE THEN EXIT OPERATOR
   ' // it has to be in precision range to say that it is equal         
   IF (dt1.m_dt + dt1.OLE_DATETIME_HALFSECOND > dt2.m_dt AND _
       dt1.m_dt - dt1.OLE_DATETIME_HALFSECOND < dt2.m_dt) THEN _
       OPERATOR = FALSE ELSE OPERATOR = TRUE
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
   IF dt1.m_status = FALSE OR dt2.m_status = FALSE THEN EXIT OPERATOR
   OPERATOR = dt1.DoubleFromDate(dt1.m_dt) < dt2.DoubleFromDate(dt2.m_dt)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
   IF dt1.m_status = FALSE OR dt2.m_status = FALSE THEN EXIT OPERATOR
   OPERATOR = dt1.DoubleFromDate(dt1.m_dt) > dt2.DoubleFromDate(dt2.m_dt)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
   IF dt1.m_status = FALSE OR dt2.m_status = FALSE THEN EXIT OPERATOR
   OPERATOR = dt1 < dt2 OR dt1 = dt2
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
   IF dt1.m_status = FALSE OR dt2.m_status = FALSE THEN EXIT OPERATOR
   OPERATOR = dt1 > dt2 OR dt1 = dt2
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR - (BYREF dateSpan AS COleDateTimeSpan) AS COleDateTimeSpan
   OPERATOR = -dateSpan.m_span
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS COleDateTimeSpan
   IF dateSpan1.m_status = FALSE OR dateSpan2.m_status = FALSE THEN EXIT OPERATOR
   DIM dateSpanTemp AS COleDateTimeSpan
   ' // Add spans and validate within legal range
   dateSpanTemp.m_span = dateSpan1.m_span + dateSpan2.m_span
   dateSpanTemp.CheckRange
   OPERATOR = dateSpanTemp
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS COleDateTimeSpan
   IF dateSpan1.m_status = FALSE OR dateSpan2.m_status = FALSE THEN EXIT OPERATOR
   DIM dateSpanTemp AS COleDateTimeSpan
   ' // Subtract spans and validate within legal range
   dateSpanTemp.m_span = dateSpan1.m_span - dateSpan2.m_span
   dateSpanTemp.CheckRange
   OPERATOR = dateSpanTemp
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR = (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSPan2 AS COleDateTimeSpan) AS BOOLEAN
   IF dateSpan1.m_status = FALSE OR dateSpan2.m_status = FALSE THEN EXIT OPERATOR
   ' // it has to be in precision range to say that it is equal         
   IF (dateSpan1.m_span + dateSpan1.OLE_DATETIME_HALFSECOND > dateSpan2.m_span AND _
       dateSpan1.m_span - dateSpan1.OLE_DATETIME_HALFSECOND < dateSpan2.m_span) THEN OPERATOR = TRUE
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSPan2 AS COleDateTimeSpan) AS BOOLEAN
   IF dateSpan1.m_status = FALSE OR dateSpan2.m_status = FALSE THEN EXIT OPERATOR
   ' // it has to be in precision range to say that it is equal         
   IF (dateSpan1.m_span + dateSpan1.OLE_DATETIME_HALFSECOND > dateSpan2.m_span AND _
       dateSpan1.m_span - dateSpan1.OLE_DATETIME_HALFSECOND < dateSpan2.m_span) THEN _
       OPERATOR = FALSE ELSE OPERATOR = TRUE
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS BOOLEAN
   IF dateSpan1.m_status = FALSE OR dateSpan2.m_status = FALSE THEN EXIT OPERATOR
   OPERATOR = dateSpan1.m_span < dateSpan2.m_span
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS BOOLEAN
   IF dateSpan1.m_status = FALSE OR dateSpan2.m_status = FALSE THEN EXIT OPERATOR
   OPERATOR = dateSpan1.m_span > dateSpan2.m_span
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS BOOLEAN
   IF dateSpan1.m_status = FALSE OR dateSpan2.m_status = FALSE THEN EXIT OPERATOR
   OPERATOR = dateSpan1 < dateSpan2 OR dateSpan1 = dateSpan2
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS BOOLEAN
   IF dateSpan1.m_status = FALSE OR dateSPan2.m_status = FALSE THEN EXIT OPERATOR
   OPERATOR = dateSpan1 > dateSpan2 OR dateSpan1 = dateSpan2
END OPERATOR
' ========================================================================================
