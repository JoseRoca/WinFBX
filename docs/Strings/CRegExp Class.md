# CRegExp Class

**CRegExp** is a wrapper class on top of the VB Script Regular Expressions.

In a typical search and replace operation, you must provide the exact text for which you are searching. That technique may be adequate for simple search and replace tasks in static text, but it lacks flexibility and makes searching dynamic text difficult, if not impossible.

With regular expressions, you can:

* Test for a pattern within a string. For example, you can test an input string to see if a telephone number pattern or a credit card number pattern occurs within the string. This is called data validation.

* Replace text. You can use a regular expression to identify specific text in a document and either remove it completely or replace it with other text.

* Extract a substring from a string based upon a pattern match. You can find specific text within a document or input field.

#### Include file

CRegExp.inc

### Constructors

```
CONSTRUCTOR CRegExp (BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE)
```
```
CONSTRUCTOR CRegExp (BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsPattern* | The regular expression pattern being searched for. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

```
CONSTRUCTOR CRegExp (BYREF pRegExp AS CRegExp)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pRegExp* | Reference to an instance of a **CRegExp** class to be cloned. |

### Methods

| Name  | Description |
| ---------- | ----------- |
| [Execute](#Execute) | Executes a regular expression search against a specified string. |
| [Extract](#Extract) | Extracts a substring using VBScript regular expressions search patterns. |
| [Find](#Find) | Find function with VBScript regular expressions search patterns. |
| [FindEx](#FindEx) | Global, multiline find function with VBScript regular expressions search patterns. |
| [GetLastResult](#GetLastResult) | Returns the last result code. |
| [MatchCount](#MatchCount) | Returns the number of matches found. |
| [RegExpPtr](#RegExpPtr) | Returns a direct pointer to the **Afx_IRegExp2** interface. |
| [Remove](#Remove) | Returns a copy of a string with text removed using a regular expression as the search string. |
| [Replace](#Replace) | Replaces text found in a regular expression search. |
| [SubMatchValue](#SubMatchValue) | Retrieves the content of the specified submatch. |
| [Test](#Test) | Executes a regular expression search against a specified string and returns a boolean value that indicates if a pattern match was found. |

### Properties

| Name  | Description |
| ---------- | ----------- |
| [Global](#Global) | Sets or returns a boolean value that indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| [IgnoreCase](#IgnoreCase) | Sets or returns a boolean value that indicates if a pattern search is case-sensitive or not. |
| [MatchLen](#MatchLen) | Returns the length of a match found in a search string. |
| [MatchPos](#MatchPos) | Returns the position in a search string where a match occurs. |
| [MatchValue](#MatchValue) | Returns the value or text of a match found in a search string. |
| [Multiline](#Multiline) | Sets or returns a boolean value that indicates whether or not to search in strings across multiple lines. |
| [Pattern](#Pattern) | Sets or returns a boolean value that indicates whether or not to search in strings across multiple lines. |
| [SubMatchesCount](#SubMatchesCount) | Returns the number of submatches. |

# <a name="Execute"></a>Execute

Executes a regular expression search against a specified string.

```
FUNCTION Execute (BYREF cbsSourceString AS CBSTR, BYREF cvReplaceString AS CVAR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
```
```
FUNCTION Execute (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYREF cvReplaceString AS CVAR,BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSourceString* | The main string. |
| *cbsPattern* | The regular string expression being searched for. |
| *cvReplaceString* | The replacement text string. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not.  |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Remarks

In the first overloaded method, the actual pattern for the regular expression search is set using the **Pattern** property.

#### Return value

BOOLEAN. True on success or False on failure.

#### Example

```
'#CONSOLE ON
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
pRegExp.Pattern = "(\w+)@(\w+)\.(\w+)"
pRegExp.IgnoreCase = TRUE
DIM cbsText AS CBSTR = "Please send mail to dragon@xyzzy.com. Thanks!"
IF pRegExp.Execute(cbsText) = FALSE THEN
   print "No match found"
ELSE
   ' Get the number of submatches
   DIM nCount AS LONG = pRegExp.SubMatchesCount(0)
   print "Sub matches: ", nCount
   FOR i AS LONG = 0 TO nCount - 1
      print pRegExp.SubMatchValue(0, i)
   NEXT
END IF

PRINT
PRINT "Press any key..."
SLEEP
```


# <a name="Extract"></a>Extract

Extracts a substring using VBScript regular expressions search patterns.

```
FUNCTION Extract (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```
```
FUNCTION Extract (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = FALSE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```
```
FUNCTION Extract (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS CBSTR
```
```
FUNCTION Extract (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, _
   BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The position in the string at which the search will begin. The first character starts at position 1. |
| *cbsSourceString* | The text to be parsed. |
| *cbsPattern* | The pattern to match. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Return value

CBSTR. The retrieved string.

#### Usage examples

```
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbs AS CBSTR = pRegExp.Extract(cbsText, cbsPattern)
' Output: a234
```
```
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbs AS CBSTR = pRegExp.Extract(15, cbsText, cbsPattern)
' Output: x345
```
```
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbs AS CBSTR = CRegExp(cbsPattern).Extract(cbsText)
' Output: a234
```
```
' // Ignore case
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsText AS CBSTR = "blah blah A234 blah blah x345 blah blah"
DIM cbs AS CBSTR = CRegExp(cbsPattern).Extract(cbsText, TRUE)
' Output: A234
```

# <a name="Find"></a>Find

Find function with VBScript regular expressions search patterns.

```
FUNCTION Find (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
```
```
FUNCTION Find (BYREF cbsSourceString AS CBSTR, _
   BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
```
```
FUNCTION Find (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
```
```
FUNCTION Find (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, _
   BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The position in the string at which the search will begin. The first character starts at position 1. |
| *cbsSourceString* | The text to be parsed. |
| *cbsPattern* | The pattern to match. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Return value

Returns the position of the match or 0 if not found. The length of the match can be retrieved calling the **MatchLen** property.

#### Usage examples

```
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM nPos AS LONG = pRegExp.Find(cbsText, cbsPattern)
' Output: 11
```
```
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM nPos AS LONG = pRegExp.Find(15, cbsText, cbsPattern)
' Output: 26
```
```
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM nPos AS LONG = CRegExp(cbsPattern).Find(cbsText)
' Output: 11
```
```
DIM nPos AS LONG = CRegExp("[a-z][0-9][0-9][0-9]").Find("blah blah a234 blah blah x345 blah blah")
' Output: 11
```

# <a name="FindEx"></a>FindEx

Global, multiline find function with VBScript regular expressions search patterns.

```
FUNCTION FindEx (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = TRUE) AS CBSTR
```
```
FUNCTION FindEx (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, _
   BYVAL bMultiline AS BOOLEAN = TRUE) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSourceString* | The text to be parsed. |
| *cbsPattern* | The pattern to match. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Return value

Returns a list of comma separated "index, length" value pairs. The pairs are separated by a semicolon.

#### Usage example

```
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsOut AS CBSTR = pRegExp.FindEx(cbsText, cbsPattern)
' Output: 11,4;26,4
```
```
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsOut AS CBSTR = CRegExp(cbsPattern).FindEx(cbsText)
' Output: 11,4;26,4
```
# <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```
#### Return value

S_OK (0) on success, or an error code on failure.

# <a name="MatchCount"></a>MatchCount

Returns the number of matches found.

```
FUNCTION MatchCount () AS LONG
```

# <a name="RegExpPtr"></a>RegExpPtr

Returns a direct pointer to the Afx_IRegExp2 interface.

```
FUNCTION RegExpPtr () AS Afx_IRegExp2 PTR
```
#### Remarks

Since it is a direct pointer, you don't have to release it calling the **Release** method.

# <a name="Remove"></a>Remove

Returns a copy of a string with text removed using a regular expression as the search string.

```
FUNCTION Remove (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```
```
FUNCTION Remove (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSourceString* | The main string. |
| *cbsPattern* | The pattern to be removed. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Usage examples

```
DIM pRegExp AS CRegExp
PRINT pRegExp.Remove("abacadabra", "ab") ' - prints "acadra"
PRINT pRegExp.Remove("abacadabra", "[bAc]", TRUE) ' - prints "dr"
PRINT pRegExp.Remove("World, worldx, world", $"\bworld\b", TRUE) ' prints ", worldx,"
```
```
PRINT CRegExp("ab").Remove("abacadabra") ' - prints "acadra"
PRINT CRegExp("[bAc]").Remove("abacadabra", TRUE) ' - prints "dr"
PRINT CRegExp($"\bworld\b").Remove("World, worldx, world", TRUE) ' prints ", worldx,"
```

# <a name="Replace"></a>Replace

Replaces text found in a regular expression search.

```
FUNCTION Replace (BYREF cbsSourceString AS CBSTR, BYREF cvReplaceString AS CVAR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```
```
FUNCTION Replace (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYREF cvReplaceString AS CVAR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSourceString* | The main string. |
| *cvReplaceString* | The replacement text string. |
| *cbsPattern* | The regular string expression being searched for. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Remarks

In the first overloaded function, the actual pattern for the text being replaced is set using the Pattern property.

The **Replace** method returns a copy of *cbsSourceString* with the text of *cbsPattern* replaced with *cvsReplaceString*. If no match is found, a copy of *cbsSourceString* is returned unchanged.

#### Examples

```
'#CONSOLE ON
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
pRegExp.Pattern = "fox"
pRegExp.IgnoreCase = TRUE
DIM cbsText AS CBSTR = "The quick brown fox jumped over the lazy dog."
' Make replacement
DIM cbsRes AS CBSTR = pRegExp.Replace(cbsText, "cat")
print cbsRes
' Output: The quick brown cat jumped over the lazy dog.
```
In addition, the **Replace** method can replace subexpressions in the pattern.

The following call to the function shown in the previous example swaps the first pair of words in the original string:

```
DIM pRegExp AS CRegExp
pRegExp.Pattern = "(\S+)(\s+)(\S+)"
pRegExp.IgnoreCase = TRUE
DIM cbsText AS CBSTR = "The quick brown fox jumped over the lazy dog."
' Make replacement
DIM cbsRes AS CBSTR = pRegExp.Replace(cbsText, "$3$2$1")
print cbsRes
```

Suppose that you have a telephone directory, and all the phone numbers are formatted like this:
555-123-4567. Now, you decide that all the phone numbers should be formatted to look like this: (555) 123-4567.

```
DIM pRegExp AS CRegExp
pRegExp.Global = TRUE
pRegExp.Pattern = "(\d{3})-(\d{3})-(\d{4})"
DIM cbsText AS CBSTR = "555-123-4567"
DIM cbsRes AS CBSTR = pRegExp.Replace(cbsText, "($1) $2-$3")
print cbsRes
```
--or--

```
DIM cbsText AS CBSTR = "555-123-4567"
DIM cbsPattern AS CBSTR = "(\d{3})-(\d{3})-(\d{4})"
DIM cbsRes AS CBSTR = CRegExp(cbsPattern).Replace(cbsText, "($1) $2-$3")
print cbsRes
```

What we have done is to search for 3 digits (\d{3}) followed by a dash, followed by 3 more digits and a dash, followed by 4 digits and add () to the first three digits and change the first dash with a space.  $1, $2, and $3 are examples of a regular expression "back reference." A back reference is simply a portion of the found text that can be saved and then reused.

# <a name="SubMatchValue"></a>SubMatchValue

Retrieves the content of the specified submatch.

```
FUNCTION SubMatchValue (BYVAL MatchIndex AS LONG = 0, BYVAL SubMatchIndex AS LONG = 0) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *MatchIndex* | 0-based index of the match to retrieve. |
| *SubMatchIndex* | 0-based index of the submatch to retrieve. |

# <a name="Test"></a>Test

Executes a regular expression search against a specified string and returns a boolean value that indicates if a pattern match was found.

```
FUNCTION Test (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
```
```
FUNCTION Test (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSourceString* | The main string. |
| *cbsPattern* | The regular string expression being searched for. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Return value

BOOLEAN. True if a pattern match is found; False if no match is found.

#### Remarks

In the first overloaded method, the actual pattern for the regular expression search is set using the **Pattern** property. The **Global** property has no effect on the **Test** method.

# <a name="Global"></a>Global

Sets or returns a boolean value that indicates if a pattern should match all occurrences in an entire search string or just the first one.

```
PROPERTY Global () AS BOOLEAN
PROPERTY Global (BYVAL bGlobal AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bGlobal* | True if the search applies to the entire string, False if it does not. Default is False. |

# <a name="IgnoreCase"></a>IgnoreCase

Sets or returns a boolean value that indicates if a pattern search is case-sensitive or not.

```
PROPERTY IgnoreCase () AS BOOLEAN
PROPERTY IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bIgnoreCase* | False if the search is case-sensitive, True if it is not. Default is False. |

# <a name="MatchLen"></a>MatchLen

Returns the length of a match found in a search string.

```
PROPERTY MatchLen (BYVAL index AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

# <a name="MatchPos"></a>MatchPos

Returns the position in a search string where a match occurs.

```
PROPERTY MatchPos (BYVAL index AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

#### Remarks

The **MatchPos** property uses a zero-based offset from the beginning of the search string. In other words, the first character in the string is identified as character zero (0).

# <a name="MatchValue"></a>MatchValue

Returns the value or text of a match found in a search string.

```
PROPERTY MatchValue (BYVAL index AS LONG = 0) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

# <a name="Multiline"></a>Multiline

Sets or returns a boolean value that indicates whether or not to search in strings across multiple lines.

```
PROPERTY Multiline () AS BOOLEAN
PROPERTY Multiline (BYVAL bMultiline AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bMultiline* | True if the search is performed across multiple lines, False if it is not. Default is False. |

# <a name="Pattern"></a>Pattern

Sets or returns the regular expression pattern being searched for.

```
PROPERTY Pattern () AS CWSTR
PROPERTY Pattern (BYREF cbsPattern AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsPattern* | Regular string expression being searched for. May include any of the regular expression characters defined in the table in the Settings section. |

#### Settings

Special characters and sequences are used in writing patterns for regular expressions. The following table describes and gives an example of the characters and sequences that can be used.

| Character  | Description |
| ---------- | ----------- |
| \\ | Marks the next character as either a special character or a literal. For example, "n" matches the character "n". "\\n" matches a newline character. The sequence "\\\\" matches "\\" and "\\(" matches "(". |
| ^ | Matches the beginning of input. |
| $ | Matches the end of input. |
| * | Matches the preceding character zero or more times. For example, "zo*" matches either "z" or "zoo". |
| + | Matches the preceding character one or more times. For example, "zo+" matches "zoo" but not "z". |
| ? | Matches the preceding character zero or one time. For example, "a?ve?" matches the "ve" in "never". |
| . | Matches any single character except a newline character. |
| (pattern) | Matches pattern and remembers the match. The matched substring can be retrieved from the resulting Matches collection, using Item \[0]...\[n]. To match parentheses characters ( ), use "\\(" or "\\)". |
| x\|y | Matches either x or y. For example, "z\|wood" matches "z" or "wood". "(z\|w)oo" matches "zoo" or "wood". |
| {n} | n is a nonnegative integer. Matches exactly n times. For example, "o{2}" does not match the "o" in "Bob," but matches the first two o's in "foooood". |
| {n,} | n is a nonnegative integer. Matches at least n times. For example, "o{2,}" does not match the "o" in "Bob" and matches all the o's in "foooood." "o{1,}" is equivalent to "o+". "o{0,}" is equivalent to "o*". |
| { n , m } | m and n are nonnegative integers. Matches at least n and at most m times. For example, "o{1,3}" matches the first three o's in "fooooood." "o{0,1}" is equivalent to "o?". |
| \[xyz] | A character set. Matches any one of the enclosed characters. For example, "\[abc]" matches the "a" in "plain".  |
| \[^xyz] | A negative character set. Matches any character not enclosed. For example, "\[^abc]" matches the "p" in "plain".  |
| \[a-z] | A range of characters. Matches any character in the specified range. For example, "\[a-z]" matches any lowercase alphabetic character in the range "a" through "z". |
| \[^m-z] | A negative range characters. Matches any character not in the specified range. For example, "\[m-z]" matches any character not in the range "m" through "z".  |
| \\b | Matches a word boundary, that is, the position between a word and a space. For example, "er\b" matches the "er" in "never" but not the "er" in "verb". |
| \\B | Matches a non-word boundary. "ea\*r\B" matches the "ear" in "never early". |
| \\d | Matches a digit character. Equivalent to \[0-9]. |
| \\D | Matches a non-digit character. Equivalent to \[^0-9]. |
| \\f | Matches a form-feed character. |
| \\n | Matches a newline character. |
| \\r | Matches a carriage return character. |
| \\s | Matches any white space including space, tab, form-feed, etc. Equivalent to "\[ \f\n\r\t\v]". |
| \\S | Matches any nonwhite space character. Equivalent to "\[^ \f\n\r\t\v]".  |
| \\t | Matches a tab character. |
| \\v | Matches a vertical tab character. |
| \\w | Matches any word character including underscore. Equivalent to "\[A-Za-z0-9_]". |
| \\W | Matches any non-word character. Equivalent to "\[^A-Za-z0-9_]". |
| \\num | Matches num, where num is a positive integer. A reference back to remembered matches. For example, "(.)\1" matches two consecutive identical characters. |
| \\n | Matches n, where n is an octal escape value. Octal escape values must be 1, 2, or 3 digits long. For example, "\11" and "\011" both match a tab character. "\0011" is the equivalent of "\001" & "1". Octal escape values must not exceed 256. If they do, only the first two digits comprise the expression. Allows ASCII codes to be used in regular expressions. |
| \\xn | Matches n, where n is a hexadecimal escape value. Hexadecimal escape values must be exactly two digits long. For example, "\x41" matches "A". "\x041" is equivalent to "\x04" & "1". Allows ASCII codes to be used in regular expressions. |

####Example

```
'#CONSOLE ON
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
pRegExp.Pattern = "is."
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute("IS1 is2 IS3 is4") = FALSE THEN
   print "No match found"
ELSE
   DIM nCount AS LONG = pRegExp.MatchesCount
   FOR i AS LONG = 0 TO nCount - 1
      print "Value: ", pRegExp.MatchValue(i)
      print "Position: ", pRegExp.MatchPos(i)
      print "Length: ", pRegExp.MatchLen(i)
      print
   NEXT
END IF

PRINT
PRINT "Press any key..."
SLEEP
```

# <a name="SubMatchesCount"></a>SubMatchesCount

Returns the number of submatches.

```
PROPERTY SubMatchesCount (BYVAL index AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

# Introduction to Regular Expressions

Unless you have worked with regular expressions before, the term and the concept may be unfamiliar to you. However, they may not be as unfamiliar as you think.

### The Concept of Regular Expressions

Think about how you search for files on your hard disk. You most likely use the ? and * characters to find files. The ? character matches a single character in a file name, while the * matches zero or more characters. A pattern such as 'data?.dat' would find the following files:

data1.dat
data2.dat
datax.dat
dataN.dat

Using the * character instead of the ? character expands the number of files found. 'data*.dat' matches all of the following:

data.dat
data1.dat
data2.dat
data12.dat
datax.dat
dataXYZ.dat

While this method of searching for files can certainly be useful, it is also very limited. The limited ability of the ? and * wildcard characters give you an idea of what regular expressions can do, but regular expressions are much more powerful and flexible.

### Uses for Regular Expressions

In a typical search and replace operation, you must provide the exact text for which you are searching. That technique may be adequate for simple search and replace tasks in static text, but it lacks flexibility and makes searching dynamic text difficult, if not impossible. 

Flexibility of Regular Expressions

With regular expressions, you can: 

* Test for a pattern within a string. For example, you can test an input string to see if a telephone number pattern or a credit card number pattern occurs within the string. This is called data validation.

* Replace text. You can use a regular expression to identify specific text in a document and either remove it completely or replace it with other text.

* Extract a substring from a string based upon a pattern match. You can find specific text within a document or input field

For example, if you need to search an entire web site to remove some outdated material and replace some HTML formatting tags, you can use a regular expression to test each file to see if the material or the HTML formatting tags you are looking for exists in that file. That way, you can narrow down the affected files to only those that contain the material that has to be removed or changed. You can then use a regular expression to remove the outdated material, and finally, you can use regular expressions to search for and replace the tags that need replacing.

Another example of where a regular expression is useful occurs in a language that is not known for its string-handling ability. VBScript, a subset of Visual Basic, has a rich set of string-handling functions. JScript, like C, does not. Regular expressions provide a significant improvement in string-handling for JScript. However, regular expressions may also be more efficient to use in VBScript as well, allowing you do perform multiple string manipulations in a single expression.

### Regular Expression Syntax

A regular expression is a pattern of text that consists of ordinary characters (for example, letters a through z) and special characters, known as metacharacters. The pattern describes one or more strings to match when searching a body of text. The regular expression serves as a template for matching a character pattern to the string being searched.

The following table contains the complete list of metacharacters and their behavior in the context of regular expressions:

| Character  | Description |
| ---------- | ----------- |
| \\ | Marks the next character as either a special character or a literal. For example, "n" matches the character "n". "\\n" matches a newline character. The sequence "\\\\" matches "\\" and "\\(" matches "(". |
| ^ | Matches the beginning of input. |
| $ | Matches the end of input. |
| * | Matches the preceding character zero or more times. For example, "zo*" matches either "z" or "zoo". |
| + | Matches the preceding character one or more times. For example, "zo+" matches "zoo" but not "z". |
| ? | Matches the preceding character zero or one time. For example, "a?ve?" matches the "ve" in "never". |
| . | Matches any single character except a newline character. |
| (pattern) | Matches pattern and remembers the match. The matched substring can be retrieved from the resulting Matches collection, using Item \[0]...\[n]. To match parentheses characters ( ), use "\\(" or "\\)". |
| x\|y | Matches either x or y. For example, "z\|wood" matches "z" or "wood". "(z\|w)oo" matches "zoo" or "wood". |
| {n} | n is a nonnegative integer. Matches exactly n times. For example, "o{2}" does not match the "o" in "Bob," but matches the first two o's in "foooood". |
| {n,} | n is a nonnegative integer. Matches at least n times. For example, "o{2,}" does not match the "o" in "Bob" and matches all the o's in "foooood." "o{1,}" is equivalent to "o+". "o{0,}" is equivalent to "o*". |
| { n , m } | m and n are nonnegative integers. Matches at least n and at most m times. For example, "o{1,3}" matches the first three o's in "fooooood." "o{0,1}" is equivalent to "o?". |
| \[xyz] | A character set. Matches any one of the enclosed characters. For example, "\[abc]" matches the "a" in "plain".  |
| \[^xyz] | A negative character set. Matches any character not enclosed. For example, "\[^abc]" matches the "p" in "plain".  |
| \[a-z] | A range of characters. Matches any character in the specified range. For example, "\[a-z]" matches any lowercase alphabetic character in the range "a" through "z". |
| \[^m-z] | A negative range characters. Matches any character not in the specified range. For example, "\[m-z]" matches any character not in the range "m" through "z".  |
| \\b | Matches a word boundary, that is, the position between a word and a space. For example, "er\b" matches the "er" in "never" but not the "er" in "verb". |
| \\B | Matches a non-word boundary. "ea\*r\B" matches the "ear" in "never early". |
| \\d | Matches a digit character. Equivalent to \[0-9]. |
| \\D | Matches a non-digit character. Equivalent to \[^0-9]. |
| \\f | Matches a form-feed character. |
| \\n | Matches a newline character. |
| \\r | Matches a carriage return character. |
| \\s | Matches any white space including space, tab, form-feed, etc. Equivalent to "\[ \f\n\r\t\v]". |
| \\S | Matches any nonwhite space character. Equivalent to "\[^ \f\n\r\t\v]".  |
| \\t | Matches a tab character. |
| \\v | Matches a vertical tab character. |
| \\w | Matches any word character including underscore. Equivalent to "\[A-Za-z0-9_]". |
| \\W | Matches any non-word character. Equivalent to "\[^A-Za-z0-9_]". |
| \\num | Matches num, where num is a positive integer. A reference back to remembered matches. For example, "(.)\1" matches two consecutive identical characters. |
| \\n | Matches n, where n is an octal escape value. Octal escape values must be 1, 2, or 3 digits long. For example, "\11" and "\011" both match a tab character. "\0011" is the equivalent of "\001" & "1". Octal escape values must not exceed 256. If they do, only the first two digits comprise the expression. Allows ASCII codes to be used in regular expressions. |
| \\xn | Matches n, where n is a hexadecimal escape value. Hexadecimal escape values must be exactly two digits long. For example, "\x41" matches "A". "\x041" is equivalent to "\x04" & "1". Allows ASCII codes to be used in regular expressions. |

Examples of Regular Expressions

"^\\s\*$" Match a blank line.<br>
"\\d{2}-\\d{5}" Validate an ID number consisting of 2 digits, a hyphen, and another 5 digits.

#### Constructing a Regular Expression

Regular expressions are constructed in the same way that arithmetic expressions are created. That is, small expressions are combined using a variety of metacharacters and operators to create larger expressions.

You construct a regular expression by putting the various components of the expression pattern between a pair of quotation marks ("") delimit regular expressions. For example: "expression". The regular expression pattern (expression) is stored in the **Pattern** property.

The components of a regular expression can be individual characters, sets of characters, ranges of characters, choices between characters, or any combination of all of these components.

Once you have constructed a regular expression, it is evaluated much like an arithmetic expression, that is, it is evaluated from left to right and follows an order of precedence. 

#### Order of Precedence

From highest to lowest, the order of precedence of the regular expression operators:

| Operator(s) | Description |
| ---------- | ----------- |
| \\ | Escape |
| (), (?:), (?=), \[] | Parentheses and brackets |
| \*, +, ?, {n}, {n,}, {n,m} | Quantifiers |
| ^, $, \\anymetacharacter | Anchors and Sequences |
| \| | Alternation |

Characters have higher precedence than the alternation operator, which allows 'm|food' to match "m" or "food". To match "mood" or "food", use parentheses to create a subexpression, which results in '(m|f)ood'.

#### Ordinary Characters

Ordinary characters consist of all printable and non-printable characters that are not explicitly designated as metacharacters. This includes all uppercase and lowercase alphabetic characters, all digits, all punctuation marks, and some symbols. 

The simplest form of a regular expression is a single, ordinary character that matches itself in a searched string. For example, the single-character pattern 'A' matches the letter 'A' wherever it appears in the searched string. Here are some examples of single-character regular expression patterns:

"a"<br>
"7"<br>
"M"

You can combine a number of single characters to form a larger expression. For example, the following JScript regular expression is nothing more than an expression created by combining the single-character expressions 'a', '7', and 'M'. 

"a7M"

Notice that there is no concatenation operator. All that is required is that you just put one character after another.

#### Special Characters

There are a number of metacharacters that require special treatment when trying to match them. To match these special characters, you must first escape those characters, that is, precede them with a backslash character (\\). 

| Special character | Meaning |
| ---------- | ----------- |
| $ | Matches the position at the end of an input string. If the RegExp object's Multiline property is set, $ also matches the position preceding '\\n' or '\\r'. To match the $ character itself, use \\$. |
| ( ) | Marks the beginning and end of a subexpression. Subexpressions can be captured for later use. To match these characters, use \\( and \\). |
| * | Matches the preceding character or subexpression zero or more times. To match the * character, use \\\*. |
| + | Matches the preceding character or subexpression one or more times. To match the + character, use \\+. |
| . | Matches any single character except the newline character \\n. To match ., use \\. |
| \[ ] | Marks the beginning of a bracket expression. To match these characters, use \\\[ and \\]. |
| ? | Matches the preceding character or subexpression zero or one time, or indicates a non-greedy quantifier. To match the ? character, use \\?. |
| \ | Marks the next character as either a special character, a literal, a backreference, or an octal escape. For example, 'n' matches the character 'n'. '\\n' matches a newline character. The sequence '\\\\' matches "\\" and '\\(' matches "(". |
| / | Denotes the start or end of a literal regular expression. To match the '/' character, use '\\/'. |
| ^ | Matches the position at the beginning of an input string except when used in a bracket expression where it negates the character set. To match the ^ character itself, use \\^. |
| { } | Marks the beginning of a quantifier expression. To match these characters, use \\{ and \}. |
| \| | Indicates a choice between two items. To match \|, use \\\|. |

#### Escape Sequences that Represent Non-printing Characters

There are a number of useful non-printing characters that must be used occasionally.

| Character  | Meaning     |
| ---------- | ----------- |
| \\cx | Matches the control character indicated by x. For example, \\cM matches a Control-M or carriage return character. The value of x must be in the range of A-Z or a-z. If not, c is assumed to be a literal 'c' character. |
| \\f | Matches a form-feed character. Equivalent to \\x0c and \\cL. |
| \\n | Matches a newline character. Equivalent to \\x0a and \\cJ. |
| \\r | Matches a carriage return character. Equivalent to \\x0d and \\cM. |
| \\s | Matches any white space character including space, tab, form-feed, and so on. Equivalent to \[\f\n\r\t\v]. |
| \\S | Matches any non-white space character. Equivalent to \[^\f\n\r\t\v]. |
| \\t | Matches a tab character. Equivalent to \\x09 and \\cI. |
| \\v | Matches a vertical tab character. Equivalent to \\x0b and \\cK. |

#### Character Matching

The period (.) matches any single printing or non-printing character in a string, except a newline character (\n). The following regular expression matches 'aac', 'abc', 'acc', 'adc', and so on, as well as 'a1c', 'a2c', a-c', and a#c': 

"a.c"

If you are trying to match a string containing a file name where a period (.) is part of the input string, you do so by preceding the period in the regular expression with a backslash (\\) character. To illustrate, the following regular expression matches 'filename.ext':

"filename\\.ext"

These expressions are still pretty limited. They only let you match any single character. Many times, it is useful to match specified characters from a list. For example, if you have an input text that contains chapter headings that are expressed numerically as Chapter 1, Chapter 2, and so on, you might want to find those chapter headings. 

#### Bracket Expressions

You can create a list of matching characters by placing one or more individual characters within square brackets (\[ and ]). When characters are enclosed in brackets, the list is called a bracket expression. Within brackets, as anywhere else, ordinary characters represent themselves, that is, they match an occurrence of themselves in the input text. Most special characters lose their meaning when they occur inside a bracket expression. Here are some exceptions: 

The ']' character ends a list if it is not the first item. To match the ']' character in a list, place it first, immediately following the opening '\['.

The '\\' character continues to be the escape character. To match the '\\' character, use '\\\\'.

Characters enclosed in a bracket expression match only a single character for the position in the regular expression where the bracket expression appears. The following regular expression matches 'Chapter 1', 'Chapter 2', 'Chapter 3', 'Chapter 4', and 'Chapter 5':

"Chapter \[12345]"

Notice that the word 'Chapter' and the space that follows are fixed in position relative to the characters within brackets. The bracket expression then, is used to specify only the set of characters that matches the single character position immediately following the word 'Chapter' and a space. That is the ninth character position.

If you want to express the matching characters using a range instead of the characters themselves, you can separate the beginning and ending characters in the range using the hyphen (-) character. The character value of the individual characters determines their relative order within a range. The following regular expression contains a range expression that is equivalent to the bracketed list shown above.

"Chapter \[1-5]"

When a range is specified in this manner, both the starting and ending values are included in the range. It is important to note that the starting value must precede the ending value in Unicode sort order.

If you want to include the hyphen character in your bracket expression, you must do one of the following: 

Escape it with a backslash: 

\[\\-]

Put the hyphen character at the beginning or the end of the bracketed list. The following expressions matches all lowercase letters and the hyphen: 

\[-a-z]
\[a-z-]

Create a range where the beginning character value is lower than the hyphen character and the ending character value is equal to or greater than the hyphen. Both of the following regular expressions satisfy this requirement: 

\[!--]
\[!-~]

You can also find all the characters not in the list or range by placing the caret (^) character at the beginning of the list. If the caret character appears in any other position within the list, it matches itself, that is, it has no special meaning. The following regular expression matches chapter headings with numbers greater than 5':

"Chapter \[^12345]"

In the examples shown above, the expression matches any digit character in the ninth position except 1, 2, 3, 4, or 5. So, for example, 'Chapter 7' is a match and so is 'Chapter 9'. 

The same expressions above can be represented using the hyphen character (-):

"Chapter \[^1-5]"

A typical use of a bracket expression is to specify matches of any upper- or lowercase alphabetic characters or any digits. The following JScript expression specifies such a match:

"\[A-Za-z0-9]"


