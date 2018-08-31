# CGpStringFormat Class

The **CGpStringFormat** class encapsulates text layout information (such as alignment, orientation, tab stops, and clipping) and display manipulations (such as trimming, font substitution for characters that are not supported by the requested font, and digit substitution for languages that do not use Western European digits). A StringFormat object can be passed to the **DrawString** methods to format a string.

**Inherits from**: CGpBase.
**Imclude file**: CGpStringFormat.inc.

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
