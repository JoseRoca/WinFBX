# CPowerTime Class

A **CPowerTime** object contains a date and time value, allowing easy calculations. The date and time value is stored as a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601. A nanosecond is one-billionth of a second.

The following static const member variables are provided to simplify calculations (number of 100-nanosecond intervals):

```
CPowerTime_Millisecond    10000ull
CPowerTime_Second         CPowerTime_Millisecond * 1000
CPowerTime_Minute         CPowerTime_Second * 60
CPowerTime_Hour           CPowerTime_Minute * 60
CPowerTime_Day            CPowerTime_Hour * 24
CPowerTime_Week           CPowerTime_Day * 7
```

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors1) | Create new **CPowerTime** objects initialized to the specified value. |
| [CAST Operator](#CastOp1) | Returns the **CPowerTime** value as a long integer. |
| [LET Operator](#LetOp1) | Assigns a value to a **CPowerTime** object. |
| [Operators](#Operators1) | Adds, subtracts or compares **CPowerTime** objects. |
| [AstroDay](#AstroDay) | Returns the Astronomical Day for any given date. |
| [DateString](#DateString) | Retuns the date as a string based on the specified mask, e.g. "dd-MM-yyyy". |
| [Day](#Day) | Returns the Day component of the **CPowerTime** object. It is a  value in the range of 1-31. |
| [Format](#Format) | Converts a **CPowerTime** object to a string. |
| [GetAsFileTime](#GetAsFileTime) | Returns the time as a **FILETIME** structure. |
| [GetAsSystemTime](#GetAsSystemTime) | Returns the time as a **SYSTEMTIME** structure. |
| [GetCurrentTime](#GetCurrentTime) | Returns a **CPowerTime** object that represents the current system date and time. |
| [GetFileTime](#GetFileTime) | Returns the value of the **CPowerTime** object. |
| [Hour](#Hour) | Returns the Hour component of the **CPowerTime** object. It is a numeric value in the range of 0-23. |
| [LocalToUTC](#LocalToUTC) | Converts a local file time to a file time based on the Coordinated Universal Time (UTC). |
| [Minute](#Minute) | Returns the Minute component of the **CPowerTime** object. This is a numeric value in the range of 0-59. |
| [Month](#Month) | Returns the Month component of the **CPowerTime** object. It is a  value in the range of 1-12. |
| [MSecond](#MSecond) | Returns the millisecond component of the **CPowerTime** object.This is a numeric value in the range of 0-999. |
| [Now](#Now) | Assigns the current local date and time on this computer to this **CPowerTime^^ object. |
| [NowUTC](#NowUTC) | Assigns the current Coordinated Universal date and time (UTC) to this **CPowerTime** object. |
| [Second](#Second) | Returns the Second component of the **CPowerTime** object. This is a numeric value in the range of 0-59. |
| [SetFileTime](#SetFileTime) | Sets the date and time of this **CPowerTime** object. |
| [TimeString](#TimeString) | Retuns the time as a string based on the specified mask, e.g. "dd-MM-yyyy". |
| [UTCToLocal](#UTCToLocal) | Converts time based on the Coordinated Universal Time (UTC) to local file time. |
| [Year](#Year) | Returns the Year component of the **CPowerTime** object. |

# <a name="Constructors1"></a>Constructors

Create new **CPowerTime** objects initialized to the specified value.

```
CONSTRUCTOR CPowerTime
CONSTRUCTOR CFileTime (BYVAL nTime AS ULONGLONG)
CONSTRUCTOR CFileTime (BYREF ft AS FILETIME)
CONSTRUCTOR CFileTime (BYREF st AS SYSTEMTIME)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTime* | A date and time expressed as a 64-bit value. |
| *ft* | A FILETIME structure. |
| *st* | A SYSTEMTIME structure. |

#### Examples

```
DIM cft AS CPowerTime = CPowerTime().GetCurrentTime()
```
```
DIM cft AS CPowerTime = AfxLocalFileTime
print cft.GetFileTime
```
```
DIM cft AS CPowerTime = AfxLocalSystemTime
print cft.GetFileTime
```

# <a name="CastOp1"></a>CAST Operator

Returns the **CPowerTime** value as a long integer.

```
OPERATOR CAST () AS LONGLONG
```

#### Examples

```
DIM cft AS CPowerTime = CPowerTime().GetCurrentTime()
DIM nTime AS LONGLONG = cft
print nTime
```
```
DIM cft AS CPowerTime = CPowerTime().GetCurrentTime()
DIM cft2 AS CPowerTime = cft
```

# <a name="LetOp1"></a>LET Operator (=)

Assigns a value to a **CPowerTime** object.

```
OPERATOR LET (BYVAL nTime AS ULONGLONG)
OPERATOR LET (BYREF ft AS FILETIME)
OPERATOR LET (BYREF st AS SYSTEMTIME)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTime* | A date and time expressed as a 64-bit value. |
| *ft* | A FILETIME structure. |
| *st* | A SYSTEMTIME structure. |

#### Examples

```
DIM cft AS CPowerTime
cft = AfxLocalFileTime
print cft.GetFileTime
```
```
DIM cft AS CPowerTime
cft = AfxLocalSystemTime
print cft.GetFileTime
```

# <a name="Operators1"></a>Operators

Compares **CPowerTime** objects.

```
OPERATOR = (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
OPERATOR <> (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
OPERATOR < (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
OPERATOR > (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
OPERATOR <= (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
OPERATOR >= (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
```
# <a name="AstroDay"></a>AstroDay

Returns the Astronomical Day for any given date.

```
FUNCTION AstroDay (BYVAL nYear AS LONG, BYVAL nMonth AS LONG, BYVAL nDay AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | A four digit year. |
| *nMonth* | A month number (1-12). |
| *nDay* | A day number (1-31). |

#### Return value

The astronomical day.

#### Remarks

Among other things, can be used to find the number of days between any two dates, e.g.:

```
DIM cpt AS CPowerTime
PRINT cpt.AstroDay(-12400, 3, 1) - cpt.AstroDay(-12400, 2, 28)  ' Prints 2
PRINT cpt.AstroDay(12000, 3, 1) - cpt.AstroDay(-12000, 2, 28) ' Prints 8765822
PRINT cpt.AstroDay(1902, 2, 28) - cpt.AstroDay(1898, 3, 1)  ' Prints 1459 days
```

# <a name="Day"></a>Day

Returns the Day component of the **CPowerTime** object. It is a  value in the range of 1-31.

```
FUNCTION Day () AS LONG
```

# <a name="Hour"></a>Hour

Returns the Hour component of the **CPowerTime** object. It is a numeric value in the range of 0-23.

```
FUNCTION Hour () AS LONG
```

# <a name="Minute"></a>Minute

Returns the Minute component of the **CPowerTime** object. This is a numeric value in the range of 0-59.

```
FUNCTION Minute () AS LONG
```

# <a name="Month"></a>Month

Returns the Month component of the **CPowerTime** object. It is a  value in the range of 1-12.

```
FUNCTION Month () AS LONG
```

# <a name="MSecond"></a>MSecond

Returns the millisecond component of the **CPowerTime** object. This is a numeric value in the range of 0-999.

```
FUNCTION MSecond () AS LONG
```

# <a name="Second"></a>Second

Returns the Second component of the **CPowerTime** object. This is a numeric value in the range of 0-59.

```
FUNCTION Second () AS LONG
```

# <a name="DsateString"></a>DateString

Retuns the date as a string based on the specified mask, e.g. "dd-MM-yyyy".

```
FUNCTION DateString (BYREF wszMask AS WSTRING, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMask* | A picture string that is used to form the date.<br>The format types "d", and "y" must be lowercase and the letter "M" must be uppercase.<br>For example, to get the date string "Wed, Aug 31 94", the application uses the picture string "ddd',' MMM dd yy". |
| *lcid* | Optional. The language identifier used for the conversion. Default is LOCALE_USER_DEFAULT. |

The following table defines the format types used to represent days.

| Format type | Meaning |
| ----------- | ----------- |
| d | Day of the month as digits without leading zeros for single-digit days. |
| dd | Day of the month as digits with leading zeros for single-digit days. |
| ddd | Abbreviated day of the week, for example, "Mon" in English (United States). |
| dddd | Day of the week. |

The following table defines the format types used to represent months.

| Format type | Meaning |
| ----------- | ----------- |
| M | Month as digits without leading zeros for single-digit months. |
| MM | Month as digits with leading zeros for single-digit months. |
| MMM | Abbreviated month, for example, "Nov" in English (United States). |
| MMMM | Month value, for example, "November" for English (United States), and "Noviembre" for Spanish (Spain). |

The following table defines the format types used to represent years.

| Format type | Meaning |
| ----------- | ----------- |
| y | Year represented only by the last digit. |
| yy | Year represented only by the last two digits. A leading zero is added for single-digit years. |
| yyyy | Year represented by a full four or five digits, depending on the calendar used. Thai Buddhist and Korean calendars have five-digit years. The "yyyy" pattern shows five digits for these two calendars, and four digits for all other supported calendars. Calendars that have single-digit or two-digit years, such as for the Japanese Emperor era, are represented differently. A single-digit year is represented with a leading zero, for example, "03". A two-digit year is represented with two digits, for example, "13". No additional leading zeros are displayed. |
| yyyyy | Behaves identically to "yyyy". |

#### Return value

The formatted date.


# <a name="TimeString"></a>TimeString

Retuns the time as a string based on the specified mask, e.g. "hh':'mm':'ss".

```
FUNCTION TimeString (BYREF wszMask AS WSTRING, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ft* | A FILETIME structure. |
| *wszMask* | A picture string that is used to form the time. |
| *lcid* | Optional. The language identifier used for the conversion. Default is LOCALE_USER_DEFAULT. |


The application can use the following elements to construct a format picture string. If spaces are used to separate the elements in the format string, these spaces appear in the same location in the output string. The letters must be in uppercase or lowercase as shown, for example, "ss", not "SS". Characters in the format string that are enclosed in single quotation marks appear in the same location and unchanged in the output string.

| Picture    | Meaning |
| ---------- | ----------- |
| h | Hours with no leading zero for single-digit hours; 12-hour clock |
| hh | Hours with leading zero for single-digit hours; 12-hour clock |
| H | Hours with no leading zero for single-digit hours; 24-hour clock |
| HH | Hours with leading zero for single-digit hours; 24-hour clock |
| m | Minutes with no leading zero for single-digit minutes |
| mm | Minutes with leading zero for single-digit minutes |
| s | Seconds with no leading zero for single-digit seconds |
| ss | Seconds with leading zero for single-digit seconds |
| t | One character time marker string, such as A or P |
| tt | Multi-character time marker string, such as AM or PM |

#### Return value

The formatted time.

# <a name="Format"></a>Format

Converts a **CPowerTime** object to a string.

```
FUNCTION Format (BYREF wszFmt AS WSTRING) AS CWSTR
```
Formatting codes:

| Code       | Meaning     |
| ---------- | ----------- |
| %a | Abbreviated weekday name |
| %A | Full weekday name |
| %b | Abbreviated month name |
| %B | Full month name |
| %c | Date and time representation appropriate for locale |
| %d | Day of month as decimal number (01 – 31) |
| %H | Hour in 24-hour format (00 – 23) |
| %I | Hour in 12-hour format (01 – 12) |
| %j | Day of year as decimal number (001 – 366) |
| %m | Month as decimal number (01 – 12) |
| %M | Minute as decimal number (00 – 59) |
| %p | Current locale's A.M./P.M. indicator for 12-hour clock |
| %S | Second as decimal number (00 – 59) |
| %U | Week of year as decimal number, with Sunday as first day of week (00 – 53) |
| %w | Weekday as decimal number (0 – 6; Sunday is 0) |
| %W | Week of year as decimal number, with Monday as first day of week (00 – 53) |
| %x | Date representation for current locale |
| %X | Time representation for current locale |
| %y | Year without century, as decimal number (00 – 99) |
| %Y | Year with century, as decimal number |
| %z, %Z | Either the time-zone name or time zone abbreviation, depending on registry settings; no characters if time zone is unknown |
| %% | Percent sign |

The # flag may prefix any formatting code. In that case, the meaning of the format code is changed as follows.

* %#a, %#A, %#b, %#B, %#p, %#X, %#z, %#Z, %#%: # flag is ignored.
* %#c: Long date and time representation, appropriate for current locale. For example: "Tuesday, March 14, 1995, 12:41:29".
* %#x* Long date representation, appropriate to current locale. For example: "Tuesday, March 14, 1995".
* %#d, %#H, %#I, %#j, %#m, %#M, %#S, %#U, %#w, %#W, %#y, %#Y: Remove leading zeros (if any).

#### Remarks

Formats the value by using the format string which contains special formatting codes that are preceded by a percent sign (%).

#### Examples

```
print cft.Format("%A, %B %d, %Y %H:%M:%S")
```
```
DIM cft AS CPowerTime
cft = AfxLocalFileTime
print cft.Format("%A, %B %d, %Y %H:%M:%S")
```

# <a name="GetAsFileTime"></a>GetAsFileTime

Returns the time as a FILETIME structure.

```
FUNCTION GetAsFileTime () AS FILETIME
```

# <a name="GetAsSystemTime"></a>GetAsSystemTime

Returns the time as a **SYSTEMTIME** structure.

```
FUNCTION GetAsSystemTime () AS SYSTEMTIME
```

# <a name="GetCurrentTime"></a>GetCurrentTime

Returns a **CPowerTime** object that represents the current system date and time.

```
FUNCTION GetCurrentTime () AS CPowerTime
```

# <a name="GetFileTime"></a>GetFileTime

Returns the value of the **CFileTime** object.

```
FUNCTION GetFileTime () AS LONGLONG
```

#### Example

```
DIM cft AS CPowerTime = CPowerTime().GetCurrentTime()
print cft.GetFileTime
```

# <a name="Now"></a>Now

Assigns the current local date and time on this computer to this **CPowerTime** object.

```
SUB Now
```

# <a name="NowUTC"></a>NowUTC

Assigns the current Coordinated Universal date and time (UTC) to this **CPowerTime** object.

```
SUB NowUTC
```

# <a name="SetFileTime"></a>SetFileTime

Sets the date and time of this **CPowerTime** object

```
FUNCTION SetFileTime (BYVAL nTime AS ULONGLONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTime* | The date and time expressed as a 64-bit value. |

# <a name="ToUTC"></a>ToUTC

The CPowerTime object is converted to Coordinated Universal Time (UTC). It is assumed that previous value was in local time.

```
SUB ToUTC
```

# <a name="ToLocalTime"></a>ToLocalTime

The CPowerTime object is converted to local time. It is assumed that the previous value was in Coordinated Universal Time (UTC).

```
SUB ToLocalTime
```
# <a name="Year"></a>Year

Returns the Year component of the **CPowerTime** object.

```
FUNCTION Year () AS LONG
```
