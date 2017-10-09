' ########################################################################################
' Microsoft Windows
' File: CTime64.inc
' Contents: Date and time classes
' Compiler: FreeBasic 32 & 64-bit
' Based on the MFC classes CTimeSpan and CTime.
' Copyright (C) Microsoft Corporation
' Written in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "Afx/AfxTime.inc"
USING Afx

NAMESPACE Afx

' ########################################################################################
'                                *** ColeTimeSpan Class ***
' ########################################################################################

' ========================================================================================
' An amount of time, which is internally stored as the number of seconds in the time span.
' ========================================================================================
TYPE CTimeSpan

Public:
   m_timeSpan AS LONGLONG

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL lltime AS LONGLONG)
   DECLARE CONSTRUCTOR (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS LONGLONG
   DECLARE OPERATOR LET (BYVAL nSpan AS LONGLONG)
   DECLARE OPERATOR LET (BYREF cSpan AS CTimeSpan)
   DECLARE OPERATOR += (BYREF cSpan AS CTimeSpan)
   DECLARE OPERATOR -= (BYREF cSpan AS CTimeSpan)
   DECLARE SUB SetTimeSpan (BYVAL lltime AS LONGLONG)
   DECLARE SUB SetTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   DECLARE FUNCTION GetTimeSpan () AS LONGLONG
   DECLARE FUNCTION GetTotalHours () AS LONGLONG
   DECLARE FUNCTION GetTotalMinutes () AS LONGLONG
   DECLARE FUNCTION GetTotalSeconds () AS LONGLONG
   DECLARE FUNCTION GetDays () AS LONG
   DECLARE FUNCTION GetHours () AS LONG
   DECLARE FUNCTION GetMinutes () AS LONG
   DECLARE FUNCTION GetSeconds () AS LONG

END TYPE
' ========================================================================================

' ========================================================================================
' CTimeSpan constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CTimeSpan
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' - lltime: A time value, which is the number of seconds in the time span.
' ========================================================================================
PRIVATE CONSTRUCTOR CTimeSpan (BYVAL lltime AS LONGLONG)
   m_timeSpan = lltime
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' - lDays, nHours, nMins, nSecs
' Indicate the day and time values to be copied into the new CTimeSpan object.
' ========================================================================================
PRIVATE CONSTRUCTOR CTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   m_timeSpan = nSecs + 60* (nMins + 60* (nHours + 24 * lDays))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' CTimeSpan destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTimeSpan
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the CTimeSpan value.
' ========================================================================================
PRIVATE OPERATOR CTimeSpan.CAST () AS LONGLONG
   RETURN m_timeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Sets the value of this date/time-span value.
' - lltime: A time value, which is the number of seconds in the time span.
' ========================================================================================
PRIVATE OPERATOR CTimeSpan.LET (BYVAL lltime AS LONGLONG)
   m_timeSpan = lltime
END OPERATOR
' ========================================================================================
' ========================================================================================
' Sets the value of this date/time-span value.
' ========================================================================================
PRIVATE OPERATOR CTimeSpan.LET (BYREF cSpan AS CTimeSpan)
   m_timeSpan = cSpan.m_timeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Sets the value of this date/time-span value.
' - lltime: A time value, which is the number of seconds in the time span.
' ========================================================================================
PRIVATE SUB CTimeSpan.SetTimeSpan (BYVAL lltime AS LONGLONG)
   m_timeSpan = lltime
END SUB
' ========================================================================================
' ========================================================================================
' Sets the value of this date/time-span value.
' ========================================================================================
PRIVATE SUB CTimeSpan.SetTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   m_timeSpan = nSecs + 60* (nMins + 60* (nHours + 24 * lDays))
END SUB
' ========================================================================================
' ========================================================================================
' Returns the value of the CTimeSpan object.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetTimeSpan () AS LONGLONG
   RETURN m_timeSpan
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a CTimeSpan object to this CTimeSpan.
' ========================================================================================
PRIVATE OPERATOR CTimeSpan.+= (BYREF cSpan AS CTimeSpan)
   m_timeSpan += cSpan.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
' Subtracts a CTimeSpan object to from this CTimeSpan.
' ========================================================================================
PRIVATE OPERATOR CTimeSpan.-= (BYREF cSpan AS CTimeSpan)
   m_timeSpan -= cSpan.m_timeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a value that represents the total number of complete hours in this CTimeSpan.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetTotalHours () AS LONGLONG
   RETURN m_timeSpan/3600
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a value that represents the total number of complete minutes in this CTimeSpan.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetTotalMinutes () AS LONGLONG
   RETURN m_timeSpan/60
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves this date/time-span value expressed in seconds.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetTotalSeconds () AS LONGLONG
   RETURN m_timeSpan
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a value that represents the number of complete days in this CTimeSpan.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetDays () AS LONG
   RETURN m_timeSpan/(24*3600) 
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a value that represents the number of hours in the current day (–23 through 23).
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetHours () AS LONG
   RETURN this.GetTotalHours()-(this.GetDays()*24)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a value that represents the number of minutes in the current hour (–59 through 59).
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetMinutes () AS LONG
   RETURN this.GetTotalMinutes()-(this.GetTotalHours()*60)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the seconds portion of this date/time-span value.
' The return values from this function range between – 59 and 59.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetSeconds () AS LONG
   RETURN this.GetTotalSeconds()-(this.GetTotalMinutes()*60)
END FUNCTION
' ========================================================================================


'' ########################################################################################
''                                 *** CTime64 Class ***
'' ########################################################################################

' ========================================================================================
' Represents an absolute time and date.
' Note: Named CTime64 because CTime conflicts with the C function ctime.
' ========================================================================================
TYPE CTime64

Public:
   m_time AS LONGLONG

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL timeSrc AS LONGINT)
   DECLARE CONSTRUCTOR (BYVAL nYear AS LONG, BYVAL nMonth AS LONG, BYVAL nDay AS LONG, BYVAL nHour AS LONG, BYVAL nMin AS LONG, BYVAL nSec AS LONG)
   DECLARE CONSTRUCTOR (BYREF systimeSrc AS SYSTEMTIME)
   DECLARE CONSTRUCTOR (BYREF filetimeSrc AS FILETIME)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS LONGLONG
   DECLARE OPERATOR LET (BYVAL timeSrc AS LONGINT)
   DECLARE OPERATOR += (BYREF dateSpan AS CTimeSpan)
   DECLARE OPERATOR -= (BYREF dateSpan AS CTimeSpan)
   DECLARE FUNCTION SetDateTime (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, BYVAL nHour AS WORD = 0, BYVAL nMin AS WORD = 0, BYVAL nSec AS WORD = 0) AS BOOLEAN
   DECLARE FUNCTION GetCurrentTime () AS CTime64
   DECLARE FUNCTION GetTime() AS LONGLONG
   DECLARE FUNCTION GetYear () AS LONG
   DECLARE FUNCTION GetMonth () AS LONG
   DECLARE FUNCTION GetDay () AS LONG
   DECLARE FUNCTION GetHour () AS LONG
   DECLARE FUNCTION GetMinute () AS LONG
   DECLARE FUNCTION GetSecond () AS LONG
   DECLARE FUNCTION GetDayOfWeek () AS LONG
   DECLARE FUNCTION GetGmtTime () AS tm
   DECLARE FUNCTION GetLocalTime () AS tm
   DECLARE FUNCTION GetAsSystemTime () AS SYSTEMTIME

END TYPE
' ========================================================================================

' ========================================================================================
' CTime64 constructors
' All these constructors create new CTime64 objects initialized to the specified value.
' - timeSrc : Indicates __time64_t (LONGLONG) value.
' - systimeSrc: A SYSTEMTIME structure to be converted to a __time64_t (LONGLONG) value
'   and copied into the new CTime64 object.
' - filetimeSrc: A FILETIME structure to be converted to a __time64_t (LONGLONG) value
'   and copied into  the new CTime64 object. Note that FILETIME uses Universal Coordinated
'   Time (UTC), so if you pass a local time in the structure, your results will be incorrect.
' - nYear, nMonth, nDay, nHour, nMin, nSec, bDTS
'   Indicates the date and time values to be copied into the new CTime object.
' ========================================================================================
PRIVATE CONSTRUCTOR CTime64
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTime64 (BYVAL timeSrc AS LONGINT)
   m_time = timeSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Parameters:
' - nYear  : The year
' - nMonth : The month
' - nDay   : The day
' - nHour  : The hour
' - nMin   : The minutes
' - nSec   : The seconds
' ========================================================================================
PRIVATE CONSTRUCTOR CTime64 (BYVAL nYear AS LONG, BYVAL nMonth AS LONG, BYVAL nDay AS LONG, BYVAL nHour AS LONG, BYVAL nMin AS LONG, BYVAL nSec AS LONG)
	IF nYear < 1900 THEN EXIT CONSTRUCTOR
	IF nMonth < 1 OR nMonth > 12 THEN EXIT CONSTRUCTOR
	IF nDay < 1 OR nDay >= 31 THEN EXIT CONSTRUCTOR
	IF nHour < 0 OR nHour > 23 THEN EXIT CONSTRUCTOR
	IF nMin < 0 OR nMin > 59 THEN EXIT CONSTRUCTOR
	IF nSec < 0 OR nSec > 59 THEN EXIT CONSTRUCTOR
   DIM atm AS tm
   atm.tm_year = nYear - 1900   ' // tm_year is 1900 based
   atm.tm_mon = nMonth - 1      ' // tm_mon is 0 based
   atm.tm_mday = nDay
   atm.tm_hour = nHour
   atm.tm_min = nMin
   atm.tm_sec = nSec
   m_time = AfxMakeTime64(atm)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTime64 (BYREF systimeSrc AS SYSTEMTIME)
   this = CTime64(systimeSrc.wYear, systimeSrc.wMonth, systimeSrc.wDay, _
          systimeSrc.wHour, systimeSrc.wMinute, systimeSrc.wSecond)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTime64 (BYREF filetimeSrc AS FILETIME)
   ' // first convert file time (UTC time) to local time
   DIM ftlocalTime AS FILETIME
   IF FileTimeToLocalFileTime(@filetimeSrc, @ftlocalTime) = 0 THEN EXIT CONSTRUCTOR
   ' // then convert that time to system time
   DIM sysTime AS SYSTEMTIME
   IF FileTimeToSystemTime(@ftlocalTime, @sysTime) = 0 THEN EXIT CONSTRUCTOR
   this = CTime64(systime.wYear, systime.wMonth, systime.wDay, systime.wHour, _
          systime.wMinute, systime.wSecond)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTime destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTime64
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the CTime value.
' ========================================================================================
PRIVATE OPERATOR CTime64.CAST () AS LONGLONG
   RETURN m_time
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a __time64_t (LONGLONG) value.
' ========================================================================================
PRIVATE OPERATOR CTime64.LET (BYVAL timeSrc AS LONGINT)
   m_time = timeSrc
END OPERATOR
' ========================================================================================

' ========================================================================================
' Adss a CTimeSpan value to this CTime64 object.
' ========================================================================================
PRIVATE OPERATOR CTime64.+= (BYREF dateSpan AS CTimeSpan)
   m_time += dateSpan.GetTimeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts a CTimeSpan value from this CTime64 object.
' ========================================================================================
PRIVATE OPERATOR CTime64.-= (BYREF dateSpan AS CTimeSpan)
   m_time -= dateSpan.GetTimeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Sets the date and time of this CTime object.
' Parameters:
' - nYear  : The year
' - nMonth : The month
' - nDay   : The day
' - nHour  : The hour
' - nMin   : The minutes
' - nSec   : The seconds
' ========================================================================================
PRIVATE FUNCTION CTime64.SetDateTime (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, BYVAL nHour AS WORD = 0, BYVAL nMin AS WORD = 0, BYVAL nSec AS WORD = 0) AS BOOLEAN
	IF nYear < 1900 THEN RETURN FALSE
	IF nMonth < 1 OR nMonth > 12 THEN RETURN FALSE
	IF nDay < 1 OR nDay >= 31 THEN RETURN FALSE
	IF nHour > 23 THEN RETURN FALSE
	IF nMin > 59 THEN RETURN FALSE
	IF nSec > 59 THEN RETURN FALSE
   DIM atm AS tm
   atm.tm_year = nYear - 1900   ' // tm_year is 1900 based
   atm.tm_mon = nMonth - 1      ' // tm_mon is 0 based
   atm.tm_mday = nDay
   atm.tm_hour = nHour
   atm.tm_min = nMin
   atm.tm_sec = nSec
   m_time = AfxMakeTime64(atm)
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a CTime object that represents the current time.
' Usage example: DIM ct AS CTime64 = CTime64().GetCurrentTime()
' --or--
' DIM ct AS CTime64
' ct = ct.GetCurrentTime()
' ========================================================================================
PRIVATE FUNCTION CTime64.GetCurrentTime () AS CTime64
   RETURN CTime64(AfxTime64)
END FUNCTION
' ========================================================================================

' ========================================================================================
' eturns a __time64_t value for the given CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetTime() AS LONGLONG
   RETURN m_time
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the year represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetYear() AS LONG
   DIM _tm AS tm = AfxGetLocalTime64(m_time)
   RETURN _tm.tm_year + 1900
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the month represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetMonth() AS LONG
   DIM _tm AS tm = AfxGetLocalTime64(m_time)
   RETURN _tm.tm_mon + 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the day represent by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetDay() AS LONG
   DIM _tm AS tm = AfxGetLocalTime64(m_time)
   RETURN _tm.tm_mday
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the hour represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetHour() AS LONG
   DIM _tm AS tm = AfxGetLocalTime64(m_time)
   RETURN _tm.tm_hour
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the minute represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetMinute() AS LONG
   DIM _tm AS tm = AfxGetLocalTime64(m_time)
   RETURN _tm.tm_min
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the second represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetSecond() AS LONG
   DIM _tm AS tm = AfxGetLocalTime64(m_time)
   RETURN _tm.tm_sec
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the day of the week represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetDayOfWeek () AS LONG
   DIM _tm AS tm = AfxGetLocalTime64(m_time)
   RETURN _tm.tm_wday + 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a tm structure that contains a decomposition of the time contained in this CTime object.
' The fields of the structure type tm store the following values, each of which is an int:
' tm_sec : Seconds after minute (0 – 59).
' tm_min : Minutes after hour (0 – 59).
' tm_hour : Hours after midnight (0 – 23).
' tm_mday : Day of month (1 – 31).
' tm_mon : Month (0 – 11; January = 0).
' tm_year : Year (current year minus 1900).
' tm_wday : Day of week (0 – 6; Sunday = 0).
' tm_yday : Day of year (0 – 365; January 1 = 0).
' tm_isdst : Positive value if daylight saving time is in effect; 0 if daylight saving
' time is not in effect; negative value if status of daylight saving time is unknown.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetGmtTime () AS tm
   RETURN AfxGmtTime64(m_time)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a tm  structure containing a decomposition of the time contained in this CTime object.
' The fields of the structure type tm store the following values, each of which is an int:
' tm_sec : Seconds after minute (0 – 59).
' tm_min : Minutes after hour (0 – 59).
' tm_hour : Hours after midnight (0 – 23).
' tm_mday : Day of month (1 – 31).
' tm_mon : Month (0 – 11; January = 0).
' tm_year : Year (current year minus 1900).
' tm_wday : Day of week (0 – 6; Sunday = 0).
' tm_yday : Day of year (0 – 365; January 1 = 0).
' tm_isdst : Positive value if daylight saving time is in effect; 0 if daylight saving
' time is not in effect; negative value if status of daylight saving time is unknown.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetLocalTime () AS tm
   RETURN AfxGetLocalTime64(m_time)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Convertsthe time information stored in the CTime object to a Win32–compatible
' SYSTEMTIME structure.
' ========================================================================================
PRIVATE FUNCTION Ctime64.GetAsSystemTime () AS SYSTEMTIME
   DIM ttm AS tm = this.GetLocalTime
   DIM ts AS SYSTEMTIME
   ts.wYear = 1900 + ttm.tm_year
	ts.wMonth = 1 + ttm.tm_mon
	ts.wDayOfWeek = ttm.tm_wday
	ts.wDay = ttm.tm_mday
	ts.wHour = ttm.tm_hour
	ts.wMinute = ttm.tm_min
	ts.wSecond = ttm.tm_sec
	ts.wMilliseconds = 0
   RETURN ts
END FUNCTION
' ========================================================================================

END NAMESPACE

' ########################################################################################
'                                *** Global operators ***
' ########################################################################################

USING Afx

' ========================================================================================
' Adds CTimeSpan values.
' ========================================================================================
PRIVATE OPERATOR + (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS CTimeSpan
   OPERATOR = CTimeSpan(cSpan1.m_timeSpan + cSpan2.m_timeSpan )
END OPERATOR
' ========================================================================================
' ========================================================================================
' Subtracts CTimeSpan values.
' ========================================================================================
PRIVATE OPERATOR - (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS CTimeSpan
   OPERATOR = CTimeSpan(cSpan1.m_timeSpan - cSpan2.m_timeSpan )
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators
' ========================================================================================
PRIVATE OPERATOR = (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan = cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan <> cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan < cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan > cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan <= cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan >= cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Adds a CTimeSpan value to a CTime64 value
' ========================================================================================
PRIVATE OPERATOR + (BYREF dt AS CTime64, BYREF dateSpan AS CTimeSpan) AS CTime64
   RETURN CTime64(dt.m_time + dateSpan.GetTimeSpan)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts a CTimeSpan value from a CTime64 value
' ========================================================================================
PRIVATE OPERATOR - (BYREF dt AS CTime64, BYREF dateSpan AS CTimeSpan) AS CTime64
   RETURN CTime64(dt.m_time - dateSpan.GetTimeSpan)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts a CTime64 value from another CTime64 value
' ========================================================================================
PRIVATE OPERATOR - (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS CTimeSpan
   RETURN CTimeSpan(dt1.m_time - dt2.m_time)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators
' ========================================================================================
PRIVATE OPERATOR = (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time = dt2.m_time
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time <> dt2.m_time
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time < dt2.m_time
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time > dt2.m_time
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time <= dt2.m_time
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time >= dt2.m_time
END OPERATOR
' ========================================================================================
