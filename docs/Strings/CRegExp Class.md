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

Returns a copy of a string with text removed using a regular expression as the search string

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

#### Example

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
