# CGpStringFormat Class

The **CGpStringFormat** class encapsulates text layout information (such as alignment, orientation, tab stops, and clipping) and display manipulations (such as trimming, font substitution for characters that are not supported by the requested font, and digit substitution for languages that do not use Western European digits). A **StringFormat** object can be passed to the **DrawString** methods to format a string.

**Inherits from**: CGpBase.<br>
**Include file**: CGpStringFormat.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#Constructor) | Creates a StringFormat object based on string format flags and a language. |
| [Clone](#Clone) | Copies the contents of the existing StringFormat object into a new StringFormat object. |
| [GenericDefault](#GenericDefault) | Creates a generic, default StringFormat object. |
| [GenericTypographic](#GenericTypographic) | Creates a generic, typographic StringFormat object. |
| [GetAlignment](#GetAlignment) | Gets an element of the StringAlignment enumeration that indicates the character alignment of this StringFormat object in relation to the origin of the layout rectangle. |
| [GetDigitSubstitutionLanguage](#GetDigitSubstitutionLanguage) | Gets the language that corresponds with the digits that are to be substituted for Western European digits. |
| [GetDigitSubstitutionMethod](#GetDigitSubstitutionMethod) | Gets an element of the StringDigitSubstitute enumeration that indicates the digit substitution method that is used by this StringFormat object. |
| [GetFormatFlags](#GetFormatFlags) | Gets the string format flags for this StringFormat object. |
| [GetHotkeyPrefix](#GetHotkeyPrefix) | Gets an element of the HotkeyPrefix enumeration that indicates the type of processing that is performed on a string when a hot key prefix, an ampersand (&), is encountered. |
| [GetLineAlignment](#GetLineAlignment) | Gets an element of the StringAlignment enumeration that indicates the line alignment of this StringFormat object in relation to the origin of the layout rectangle. |
| [GetMeasurableCharacterRangeCount](#GetMeasurableCharacterRangeCount) | Gets the number of measurable character ranges that are currently set. |
| [GetTabStopCount](#GetTabStopCount) | Gets the number of tab-stop offsets in this StringFormat object. |
| [GetTabStops](#GetTabStops) | Gets the offsets of the tab stops in this StringFormat object. |
| [GetTrimming](#GetTrimming) | Gets an element of the StringTrimming enumeration that indicates the trimming style of this StringFormat object. |
| [SetAlignment](#SetAlignment) | Sets the character alignment of this StringFormat object in relation to the origin of the layout rectangle. |
| [SetDigitSubstitution](#SetDigitSubstitution) | Sets the digit substitution method and the language that corresponds to the digit substitutes. |
| [SetFormatFlags](#SetFormatFlags) | Sets the format flags for this StringFormat object. |
| [SetHotkeyPrefix](#SetHotkeyPrefix) | Sets the type of processing that is performed on a string when the hot key prefix, an ampersand (&), is encountered. |
| [SetLineAlignment](#SetLineAlignment) | Sets the line alignment of this StringFormat object in relation to the origin of the layout rectangle. |
| [SetMeasurableCharacterRanges](#SetMeasurableCharacterRanges) | Sets a series of character ranges for this StringFormat object that, when in a string, can be measured by the MeasureCharacterRanges method. |
| [SetTabStops](#SetTabStops) | Sets the offsets for tab stops in this StringFormat object. |
| [SetTrimming](#SetTrimming) | Sets the trimming style for this StringFormat object. |

# <a name="Constructor"></a>Constructor

Creates a **StringFormat** object based on string format flags and a language.

```
CONSTRUCTOR CGpStringFormat (BYVAL formatFlags AS INT_ = 0, BYVAL language AS LANGID = LANG_NEUTRAL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *formatFlags* | Value that specifies the format flags that control most of the characteristics of the **StringFormat** object. The flags are set by applying a bitwise OR to elements of the **StringFormatFlags** enumeration. The default value is 0 (no flags set). |
| *language* | Sixteen-bit value that specifies the language to use. The default value is LANG_NEUTRAL, which is the user's default language. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="Clone"></a>Clone

Copies the contents of the existing **StringFormat** object into a new **StringFormat** object.

```
FUNCTION Clone (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pStringFormat* | Pointer to the **StringFormat** object where to copy the contents of the existing object. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, clones it, and then uses the clone
' to draw a formatted string.
' ========================================================================================
SUB Example_Clone (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the alignment
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetAlignment(StringAlignmentCenter)

   ' // Clone the StringFormat object
   DIM pStringFormat AS CGpStringFormat
   stringFormat.Clone(@pStringFormat)

   ' // Use the cloned StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "This text was formatted by a StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 200, 200, @pStringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 200, 200)

END SUB
' ========================================================================================
```

# <a name="GenericDefault"></a>GenericDefault

Creates a generic, default **StringFormat** object.

```
FUNCTION GenericDefault (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pStringFormat* | Pointer to the **StringFormat** object where to return the new **StringFormat** object. |

#### Remarks

A generic, default **StringFormat** object has the following characteristics:

* No string format flags are set.
* Character alignment and line alignment are set to **StringAlignmentNear**.
* Language ID is set to neutral language, which means that the current language associated with the calling thread is used.
* String digit substitution is set to **StringDigitSubstituteUser**.
* Hot key prefix is set to **HotkeyPrefixNone**.
* Number of tab stops is set to zero.
* String trimming is set to **StringTrimmingCharacter**.

#### Example

```
' ========================================================================================
' The following example creates a generic, default StringFormat object and then uses it to
' draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GenericDefault (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   ' // Points must not be scaled for a generic font, so we are unscaling them dividing by rxRatio
   DIM pFont AS CGpFont = CGpFont(@fontFamily, AfxPointsToPixelsX(12) / rxRatio, FontStyleRegular, UnitPixel)

   ' // Create a generic default string format
   DIM stringFormat AS CGpStringFormat
   stringFormat.GenericDefault(@stringFormat)

   ' // Use the generic StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 260 = "This text was formatted by a generic StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 100, 120, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text.
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 100, 120)

END SUB
' ========================================================================================
```

# <a name="GenericTypographic"></a>GenericTypographic

Creates a generic, typographic **StringFormat** object.

```
FUNCTION GenericTypographic (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pStringFormat* | Pointer to the **StringFormat** object where to return the new **StringFormat** object. |

#### Remarks

A generic, typographic **StringFormat** object has the following characteristics:

* String format flags **StringFormatFlagsLineLimit**, **StringFormatFlagsNoClip**, and **StringFormatFlagsNoFitBlackBox** are set.
* Character alignment and line alignment are set to **StringAlignmentNear**.
* Language ID is set to neutral language, which means that the current language associated with the calling thread is used.
* String digit substitution is set to **StringDigitSubstituteUser**.
* Hot key prefix is set to **HotkeyPrefixNone**.
* Number of tab stops is set to zero.
* String trimming is set to **StringTrimmingNone**.

#### Example

```
' ========================================================================================
' The following example creates a generic, default StringFormat object and then uses it to
' draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GenericTypographic (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   ' // Points must not be scaled for a typographic font, so we are unscaling them dividing by rxRatio
   DIM pFont AS CGpFont = CGpFont(@fontFamily, AfxPointsToPixelsX(12) / rxRatio, FontStyleRegular, UnitPixel)

   ' // Create a typographic string format object
   DIM stringFormat AS CGpStringFormat
   stringFormat.GenericTypographic(@stringFormat)

   ' // Use the generic StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 260 = "Formatted by a generic typographic StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 100, 120, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text.
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 100, 120)

END SUB
' ========================================================================================
```

# <a name="GetAlignment"></a>GetAlignment

Gets an element of the **StringAlignment** enumeration that indicates the character alignment of this **StringFormat** object in relation to the origin of the layout rectangle. A layout rectangle is used to position the displayed string.

```
FUNCTION GetAlignment () AS StringAlignment
```

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the string's character alignment,
' gets the alignment value, and stores it in a variable. The code then creates a second
' StringFormat object and uses the stored alignment value to set the character alignment
' of the second StringFormat object. Next, the code uses the second StringFormat object to
' draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetAlignment (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the alignment
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetAlignment(StringAlignmentFar)

   ' // Get the alignment setting from the StringFormat object.
   DIM nStringAlignment AS StringAlignment = stringFormat.GetAlignment

   ' // Create a second StringFormat object with the same alignment.
   DIM stringFormat2 AS CGpStringFormat
   stringFormat2.SetAlignment(nStringAlignment)

   ' // Use the second StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 260 = "This text was formatted by a second StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 150, 200, @stringFormat2, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 150, 200)

END SUB
' ========================================================================================
```

# <a name="GetDigitSubstitutionLanguage"></a>GetDigitSubstitutionLanguage

Gets the language that corresponds with the digits that are to be substituted for Western European digits.

```
FUNCTION GetDigitSubstitutionLanguage () AS LANGID
```

#### Return value

This method returns a 16-bit value that forms a National Language Support (NLS) language identifier. This identifier indicates the language that corresponds with the substitution digits. For example, if this **StringFormat** object uses Arabic substitution digits, then this method will return a value that indicates an Arabic language. An NLS language identifier is constructed by the MAKELANGID macro declared in Winnt.bi.

# <a name="GetDigitSubstitutionMethod"></a>GetDigitSubstitutionMethod

Gets an element of the **StringDigitSubstitute** enumeration that indicates the digit substitution method that is used by this **StringFormat** object.

```
FUNCTION GetDigitSubstitutionMethod () AS StringDigitSubstitute
```

#### Return value

The digit substitution method replaces, in a string, Western European digits with digits that correspond to a user's locale or language.

# <a name="GetFormatFlags"></a>GetFormatFlags

Gets the string format flags for this StringFormat object.

```
FUNCTION GetFormatFlags () AS INT_
```

#### Return value

This method returns a value that indicates which string format flags are set for this **StringFormat** object. This value can be any combination (the result of a bitwise OR applied to two or more elements) of elements of the **StringFormatFlag**s enumeration.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the string's format flags, and
' then gets the 32-bit value that contains the format flags and stores it in a variable.
' The code then creates another StringFormat object and uses the stored format flags value
' to set the format flags of the second StringFormat object. Next, the code uses the second
' StringFormat object to draw a formatted string . The code also draws the string's layout
' rectangle.
' ========================================================================================
SUB Example_GetFormatFlags (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set its format flags
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetFormatFlags(StringFormatFlagsDirectionVertical OR StringFormatFlagsNoFitBlackBox)

   ' // Get the format flags from the StringFormat object.
   DIM flags AS LONG = stringFormat.GetFormatFlags

   ' // Create a second StringFormat object with the same flags.
   DIM stringFormat2 AS CGpStringFormat
   stringFormat2.SetFormatFlags(flags)

      ' // Use the second StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "This text is vertical because of a format flag."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 150, 200, @stringFormat2, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 150, 200)

END SUB
' ========================================================================================
```
