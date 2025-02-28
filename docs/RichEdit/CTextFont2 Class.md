# CTextFont2 Class

Class that wraps all the methods of the **ITextFont2** interface.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTOR](#constructor) | Called when a class variable is created. |
| [DESTRUCTOR](#destructor) | Called automatically when a class variable goes out of scope or is destroyed. |

### Inherited ITextFont Properties

| Name       | Description |
| ---------- | ----------- |
| [AllCaps](#allcaps) | Gets/sets whether the characters are all uppercase. |
| [Animation](#animation) | Gets/sets the animation type. |
| [BackColor](#backColor) | Gets/sets the text background (highlight) color. |
| [Bold](#bold) | Gets/sets whether the characters are bold. |
| [CanChange](#canchange) | Determines whether the font can be changed. |
| [Duplicate](#duplicate) | Gets/sets a duplicate of this text font object. |
| [Emboss](#emboss) | Gets/sets whether characters are embossed. |
| [Engrave](#engrave) | Gets/sets whether characters are displayed as imprinted characters. |
| [ForeColor](#forecolor) | Gets/sts the foreground, or text, color. |
| [Hidden](#hidden) | Gets/sets whether characters are hidden. |
| [IsEqual](#isequal) | Determines whether this text font object has the same properties as the specified text font object. |
| [Italic](#italic) | Gets/sets whether characters are in italics. |
| [Kerning](#kerning) | Gets/sets the minimum font size at which kerning occurs. |
| [LanguageID](#languageid) | Gets/sets the language ID or language code identifier (LCID). |
| [Name](#name) | Gets/sets the font name. |
| [Outline](#outline) | Gets/sets whether characters are displayed as outlined characters. |
| [Position](#position) | Gets/sets the amount that characters are offset vertically relative to the baseline. |
| [Protected](#protected) | Gets/sets whether characters are protected against attempts to modify them. |
| [Reset](#reset) | Resets the character formatting to the specified values. |
| [Shadow](#shadow) | Gets/sets whether characters are displayed as shadowed characters. |
| [Size](#size) | Gets/sets the font size. |
| [SmallCaps](#smallcaps) | Gets/sets whether characters are in small capital letters. |
| [Spacing](#spacing) | Gets/sets the amount of horizontal spacing between characters. |
| [StrikeThrough](#strikethrough) | Gets whether characters are displayed with a horizontal line through the center. |
| [Style](#style) | Gets/sets the character style handle of the characters in a range. |
| [Subscript](#subscript) | Gets/sets whether characters are displayed as subscript. |
| [Superscript](#superscript) | Gets/sets whether characters are displayed as superscript. |
| [Underline](#underline) | Gets/sets the type of underlining for the characters in a range. |
| [Weight](#weight) | Gets/sets the font weight for the characters in a range. |

### ITextFont2 Interface

In the Text Object Model (TOM), applications access text-range attributes by using a pair of dual interfaces, **ITextFont** and **ITextPara**.

The **ITextFont2** interface extends **ITextFont**, providing the programming equivalent of the Microsoft Word format-font dialog.

| Name       | Description |
| ---------- | ----------- |
| [AutoLigatures](#autoligatures) | Gets/sets whether support for automatic ligatures is active. |
| [AutospaceAlpha](#autospacealpha) | Gets/sets the East Asian "autospace alphabetics" state. |
| [AutospaceNumeric](#autospacenumeric) | Gets/sets the East Asian "autospace numeric" state. |
| [AutospaceParens](#autospaceparens) | Gets/sets the East Asian "autospace parentheses" state. |
| [CharRep](#charrep) | Gets/sets the character repertoire (writing system). |
| [CompressionMode](#compressionmode) | Gets/sets the East Asian compression mode. |
| [Cookie](#cookie) | Gets/sets the client cookie. |
| [Count](#count) | Gets the count of extra properties in this character formatting collection. |
| [DoubleStrike](#doublestrike) | Gets/sets whether characters are displayed with double horizontal lines through the center. |
| [LinkType](#linktype) | Gets the link type. |
| [MathZone](#mathzone) | Gets/sets whether a math zone is active. |
| [ModWidthPairs](#modwidthpairs) | Gets/sets whether "decrease widths on pairs" is active. |
| [ModWidthSpace](#modwidthspace) | Gets/sets whether "increase width of whitespace" is active. |
| [OldNumbers](#oldnumbers) | Gets/sets whether old-style numbers are active. |
| [Overlapping](#overlapping) | Gets/sets whether overlapping text is active. |
| [PositionSubSuper](#positionsubsuper) | Gets/sets the subscript or superscript position relative to the baseline. |
| [Scaling](#scaling) | Gets/sets the font horizontal scaling percentage. |
| [SpaceExtension](#spaceextension) | Gets/sets the East Asian space extension value. |


| Name       | Description |
| ---------- | ----------- |
| [GetUnderlinePositionMode](#GetUnderlinePositionMode) | Gets the underline position mode. |
| [SetUnderlinePositionMode](#SetUnderlinePositionMode) | Sets the underline position mode. |
| [GetEffects](#GetEffects) | Gets the character format effects. |
| [GetEffects2](#GetEffects2) | Gets the additional character format effects. |
| [GetProperty](#GetProperty) | Gets the value of the specified property. |
| [GetPropertyInfo](#GetPropertyInfo) | Gets the property type and value of the specified extra property. |
| [SetEffects](#SetEffects) | Sets the character format effects. |
| [SetEffects2](#SetEffects2) | Sets the additional character format effects. |
| [SetProperty](#SetProperty) | Sets the value of the specified property. |

### Methods inherited from CTOMBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#getlastresult) | Returns the last result code |
| [SetResult](#setresult) | Sets the last result code. |
| [GetErrorInfo](#geterrorinfo) | Returns a description of the last result code. |

## <a name="constructor"></a>CONSTRUCTOR

Called when a **CTextFont2** class variable is created.

```
CONSTRUCTOR (BYVAL pTextFont2 AS ITextFont2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```
| Parameter | Description |
| --------- | ----------- |
| *pTextFont2* | A **ITextFont2** interface pointer. |
| *fAddRef* | Optional. **TRUE** to increment the reference count of the passed **ITextFont2** interface pointer; otherwise, **FALSE**. Default is **FALSE**. |

#### Return value

A pointer to the new instance of the class.

---

## <a name="destructor"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

```
DESTRUCTOR CTextFont2
```
---

## <a name="getlastresult"></a>GetLastResult

Returns the last result code

```
FUNCTION GetLastResult () AS HRESULT
```
---

## <a name="setresult"></a>SetResult

Sets the last result code.

```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Result* | The **HRESULT** error code returned by the methods. |

---

## <a name="geterrorinfo"></a>GetErrorInfo

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS CWSTR
```
---

## <a name="duplicate"></a>Duplicate

Gets/sets a duplicate of this range object. In this implementation of the class, **Duplicate** and **Duplicate2** are the same method.

```
(GET) PROPERTY Duplicate () AS ITextFont2 PTR
(SET) FUNCTION Duplicate (BYVAL pFont AS ITextFont2 PTR)
```
```
(GET) PROPERTY Duplicate2 () AS ITextFont2 PTR
(SET) FUNCTION Duplicate2 (BYVAL pFont AS ITextFont2 PTR)
```
```
FUNCTION GetDuplicate () AS ITextFont2 PTR
FUNCTION SetDuplicate (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
```
```
FUNCTION GetDuplicate2 () AS ITextFont2 PTR
FUNCTION SetDuplicate2 (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *pFont* | The text font object to apply to this font object. |

#### Return value

The duplicated text font object.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_INVALIDARG** | Invalid argument. |
| **E_OUTOFMEMORY** | Memory could not be allocated for the new object. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

Values with the **tomUndefined** attribute have no effect.

---

## <a name="canchange"></a>CanChange

Determines whether the font can be changed.

```
FUNCTION CanChange () AS LONG
```

#### Return value

A variable that is **tomTrue** if the font can be changed or **tomFalse** if it cannot be changed.

#### Result code

If the font can change, the method returns **S_OK**. If the method fails, it returns **S_FALSE**.

#### Remarks

This method returns **tomTrue** only if the font can be changed. That is, no part of an associated range is protected and an associated document is not read-only. If this **ITextFont** object is a duplicate, no protection rules apply.

---

## <a name="isequal"></a>IsEqual

Determines whether this text font object has the same properties as the specified text font object.

```
FUNCTION IsEqual (BYVAL pFont AS ITextFont2 PTR) AS LONG
FUNCTION IsEqual2 (BYVAL pFont AS ITextFont2 PTR) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *pFont* | The text font object to compare against. |

#### Return value

It returnss **tomTrue** if the font objects have the same properties or **tomFalse** if they do not.

#### Remarks

The text font objects are equal only if *pFont* belongs to the same Text Object Model (TOM) object as the current font object. The **IsEqual** method ignores entries for which either font object has an **tomUndefined**.

---

## <a name="Reset"></a>Reset

Resets the character formatting to the specified values.

```
FUNCTION Reset (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The kind of reset. This parameter can be a combination of the following values. |

| Value | Meaning |
| ----- | ------- |
| **tomDefault** | Set to the document default character format if this font object is attached to a range; otherwise, set the defaults to the basic TOM engine defaults. |
| **tomUndefined** | Sets all properties to undefined values. This value is valid only for a duplicate (clone) font object. |
| **tomApplyLater** | Allow property values to be set, but don’t apply them to the attached range yet. |
| **tomApplyNow** | Apply the current properties to attached range. |
| **tomCacheParms** | Do not update the current font with the attached range properties. |
| **tomTrackParms** | Update the current font with the attached range properties. |
| **tomApplyTmp** | Apply temporary formatting. |
| **tomDisableSmartFont** | Do not apply smart fonts. |
| **tomEnableSmartFont** | Do apply smart fonts. |
| **tomUsePoints** | Use points for floating-point measurements. |
| **tomUseTwips** | Use twips for floating-point measurements. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| **S_FALSE** | Protected from change. |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

Calling **Reset** with **tomUndefined** sets all properties to undefined values. Thus, applying the font object to a range changes nothing. This applies to a font object that is obtained by the **GetDuplicate** method.

---

## <a name="style"></a>Style

Gets/sets the character style handle of the characters in a range.

```
(GET) PROPERTY Style () AS LONG
(SET) PROPERTY Style (BYVAL Value AS LONG)
```
```
FUNCTION GetStyle () AS LONG
FUNCTION SetStyle (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new character style handle. |

#### Return value

(GET) The character style handle.

(SET) If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error code.

| Return code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

The Text Object Model (TOM) version 1.0 does not specify the meanings of the style handles. The meanings depend on other facilities of the text system that implements TOM.

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

---

## <a name="allcaps"></a>AllCaps

Gets/sets whether the characters are all uppercase.

```
(GET) FUNCTION GetAllCaps () AS LONG
(SET) FUNCTION SetAllCaps (BYVAL Value AS LONG) AS HRESULT
```
```
FUNCTION GetAllCaps () AS LONG
FUNCTION SetAllCaps (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are all uppercase. |
| **tomFalse** | Characters are not all uppercase. |
| **tomUndefined** | The AllCaps property is undefined. |

#### Return value

A **tomBool** value that can be one of the above.
	
#### Return code

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error code.

| Return code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

This property corresponds to the **CFE_ALLCAPS** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

---

## <a name="animation"></a>Animation

Gets/sets the animation type.

```
(GET) PROPERTY Animation () AS LONG
(SET) PROPERTY Animation (BYVAL Value AS LONG)
```
```
FUNCTION GetAnimation () AS LONG
FUNCTION SetAnimation (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

The animation type. It can be one of the following values.

| Animation type | Value | Meaning |
| -------------- | ----- | ------- |
| **tomNoAnimation** | 0 | Do not apply text animation. |
| **tomLasVegasLights** | 1 | Text is bordered by marquee lights that blink between the colors red, yellow, green, and blue. |
| **tomBlinkingBackground** | 2 | Text has a black background that blinks on and off. |
| **tomSparkleText** | 3 | Text is overlaid with multicolored stars that blink on and off at regular intervals. |
| **tomMarchingBlackAnts** | 4 | Text is surrounded by a black dashed-line border. The border is animated so that the individual dashes appear to move clockwise around the text. |
| **tomMarchingRedAnts** | 5 | Text is surrounded by a red dashed-line border that is animated to appear to move clockwise around the text. |
| **tomShimmer** | 6 | Text is alternately blurred and unblurred at regular intervals, to give the appearance of shimmering. |
| **tomWipeDown** | 7 | Text appears gradually from the top down. |
| **tomWipeRight** | 8 | Text appears gradually from the bottom up. |

#### Return value

One of the animation types listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

---

## <a name="backcolor"></a>BackColor

Gets/sets the text background (highlight) color.

```
PROPERTY BackColor () AS LONG
PROPERTY BackColor (BYVAL Value AS LONG)
```
```
FUNCTION GetBackColor () AS LONG
FUNCTION SetBackColor (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new background color. It can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| A **COLORREF** value | An RGB color. |
| A value returned by **PALETTEINDEX** | A palette index. |
| **tomUndefined** | No change. |
| **tomAutoColor** | Use the default background color. |

If **Value** contains an RGB color, generate the **COLORREF** by using the [RGB](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-rgb) macro (BGR function in FreeBasic).

#### Return value

The text background color. It can be one of the values listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="bold"></a>Bold

Gets/sets whether the characters are bold.

```
(GET) Bold () AS LONG
(SET) Bold (BYVAL Value AS LONG)
```
```
FUNCTION GetBold () AS LONG
FUNCTION SetBold (BYVAL Value AS LONG) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are bold. |
| **tomFalse** | Characters are not bold. |
| **tomToggle** | Toggle the state of the Bold property. |
| **tomUndefined** | The Bold property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

# <a name="emboss"></a>Emboss

Gets/sets whether the characters are embossed.

```
(GET) Bold () AS LONG
(SET) Bold (BYVAL Value AS LONG)
```
```
FUNCTION GetEmboss () AS LONG
FUNCTION SetEmboss (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are embossed. |
| **tomFalse** | Characters are not embossed. |
| **tomToggle** | Toggle the state of the Emboss property. |
| **tomUndefined** | The Emboss property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

This property corresponds to the **CFE_EMBOSS** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

---

## <a name="forecolor"></a>ForeColor

Gets/sets the foreground, or text, color.

```
(GET) PROPERTY ForeColor () AS LONG
(SET) PROPERTY ForeColor (BYVAL Value AS LONG)
```
```
FUNCTION GetForeColor () AS LONG
FUNCTION SetForeColor (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new foreground color. It can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| A **COLORREF** value | An RGB color. |
| A value returned by **PALETTEINDEX** | A palette index. |
| **tomUndefined** | No change. |
| **tomAutoColor** | Use the default background color. |

If **Value** contains an RGB color, generate the **COLORREF** by using the [RGB](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-rgb) macro (BGR function in FreeBasic).

#### Return value

The foreground color. It can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="hidden"></a>Hidden

Gets/sets whether characters are hidden.

```
(GET) Hidden () AS LONG
(SET) Hidden (BYVAL Value AS LONG)
```
```
FUNCTION GetHidden () AS LONG
FUNCTION SetHidden (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are hidden. |
| **tomFalse** | Characters are not hidden. |
| **tomToggle** | Toggle the state of the Hidden property. |
| **tomUndefined** | The Hidden property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

This property corresponds to the **CFE_HIDDEN** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

## <a name="engrave"></a>Engrave

Gets/sets whether characters are displayed as imprinted characters.

```
(GET) PROPERTY Engrave () AS LONG
(SET) PROPERTY Engrave (BYVAL Value AS LONG)
```
```
FUNCTION GetEngrave () AS LONG
FUNCTION SetEngrave (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are imprinted. |
| **tomFalse** | Characters are not imprinted. |
| **tomToggle** | Toggle the state of the Engrave property. |
| **tomUndefined** | The Engrave property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

This property corresponds to the **CFE_IMPRINT** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

---

## <a name="italic"></a>Italic

Gets/sets whether characters are in italics.

```
(GET) PROPERTY Italic () AS LONG
(SET) PROPERTY Italic (BYVAL Value AS LONG)
```
```
FUNCTION GetItalic () AS LONG
FUNCTION SetItalic (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are in italics. |
| **tomFalse** | Characters are not in italics. |
| **tomToggle** | Toggle the state of the Italic property. |
| **tomUndefined** | The Italic property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="kerning"></a>Kerning

Gets/sets the minimum font size at which kerning occurs.

```
(GET) PROPERTY Kerning () AS SINGLE
(SET) PROPERTY Kerning (BYVAL Value AS SINGLE)
```
```
FUNCTION GetKerning () AS SINGLE
SetKerning (BYVAL Value AS SINGLE) As HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | The new value of the minimum kerning size, in floating-point points. |

If the value is zero, kerning is turned off. Positive values turn on pair kerning for font sizes greater than this kerning value. For example, the value 1 turns on kerning for all legible sizes, whereas 16 turns on kerning only for font sizes of 16 points and larger.

#### Return value

The minimum font size at which kerning occurs, in floating-point points.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="languageid"></a>LanguageID

Gets/sets the language ID or language code identifier (LCID).

```
(GET) PROPERTY LanguageID () AS LONG
(SET) PROPERTY LanguageID (BYVAL Value AS LONG)
```
```
FUNCTION GetLanguageID () AS LONG
SetLanguageID (BYVAL Value AS LONG) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | The new value of the minimum kerning size, in floating-point points. |

The new language identifier. The low word contains the language identifier. The high word is either zero or it contains the high word of the locale identifier LCID. For more information, see [Locale Identifiers](https://learn.microsoft.com/en-us/windows/win32/intl/locale-identifiers).

#### Return value

The language ID or LCID. The low word contains the language identifier. The high word is either zero or it contains the high word of the LCID. To retrieve the language identifier, mask out the high word. For more information, see [Locale Identifiers](https://learn.microsoft.com/en-us/windows/win32/intl/locale-identifiers).

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

If the high nibble of *Value* is **tomCharset**, set the charrep from the charset in the low byte and the pitch and family from the next byte. See also **SetCharRep**.

If the high nibble of *Value* is **tomCharRepFromLcid**, set the *charrep* from the LCID and set the LCID as well. See **GetLanguageID** for more information.

To set the BCP-47 language tag, such as "en-US", call **SetText2** and set the **tomLanguageTag** and *bstr* with the language tag.

## <a name="name"></a>Name

Gets/sets the font name.

```
(GET) PROPERTY Name () AS CBSTR
(SET) PROPERTY Name (BYVAL fontName AS AFX_BSTR)
```
```
FUNCTION GetName () AS CBSTR
FUNCTION SetName (BYVAL fontName AS BSTR) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *fontName* | The new font name. |

#### Return value

The font name.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="outline"></a>Outline

Gets/sets whether characters are displayed as outlined characters.

```
(GET) PROPERTY Outline () AS LONG
(SET) PROPERTY Outline (BYVAL Value AS LONG)
```
```
FUNCTION GetOutline () AS LONG
FUNCTION SetOutline (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are outlined. |
| **tomFalse** | Characters are not outlined. |
| **tomToggle** | Toggle the state of the Outline property. |
| **tomUndefined** | The Outline property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="position"></a>Position

Gets/sets the amount that characters are offset vertically relative to the baseline.

```
(GET) PROPERTY Position () AS SINGLE
(SET) PROPERTY Position (BYVAL Value AS SINGLE)
```
```
FUNCTION GetPosition () AS SINGLE
FUNCTION SetPosition (BYVAL Value AS SINGLE) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new amount of vertical offset, in floating-point points. |

#### Return value

The amount of vertical offset, in floating-point points.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

Displayed text typically has a zero value for this property. Positive values raise the text, and negative values lower it.

---

## <a name="protected"></a>Protected

Gets/sets whether characters are protected against attempts to modify them.

```
(GET) PROPERTY Protected_ () AS LONG
(SET) PROPERTY Protected_ (BYVAL Value AS LONG) AS HRESULT
```
```
FUNCTION GetProtected () AS LONG
FUNCTION SetProtected (BYVAL Value AS LONG) AS HRESULT
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are protected. |
| **tomFalse** | Characters are not protected. |
| **tomUndefined** | The Protected property is undefined. |

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are protected. |
| **tomFalse** | Characters are not protected. |
| **tomToggle** | Toggle the state of the Protected property. |
| **tomUndefined** | The Protected property is undefined. |

#### Remarks

In general, Text Object Model (TOM) methods that attempt to change the formatting or content of a range fail with **E_ACCESSDENIED** if any part of that range is protected, or if the document is read only. To make a change in protected text, the TOM client should attempt to turn off the protection of the text to be modified. The owner of the document may permit this to happen. For example in rich edit controls, attempts to change protected text result in an **EN_PROTECTED** notification code to the creator of the document, who then can refuse or grant permission for the change. The creator is the client that created a windowed rich edit control through the **CreateWindowEx** function or the **ITextHost** object that called the **CreateTextServices** function to create a windowless rich edit control.

---

## <a name="shadow"></a>Shadow

Gets/sets whether characters are displayed as shadowed characters.

```
(GET) PROPERTY Shadow () AS LONG
(SET) PROPERTY Shadow (BYVAL Value AS LONG)
```
```
FUNCTION GetShadow () AS LONG
FUNCTION SetShadow (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are shadowed. |
| **tomFalse** | Characters are not shadowed. |
| **tomToggle** | Toggle the state of the Shadow property. |
| **tomUndefined** | The Shadow property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed below.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

This property corresponds to the **CFE_SHADOW** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

---

## <a name="size"></a>Size

Gets/sets the font size.

```
(GET) PROPERTY Size () AS SINGLE
(SET) PROPERTY Size (BYVAL Value AS SINGLE)
```
```
FUNCTION GetSize () AS SINGLE
FUNCTION SetSize (BYVAL Value AS SINGLE) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new font size, in floating-point points. |

#### Return value

The font size, in floating-point points.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="smallcaps"></a>SmallCaps

Gets/sets whether characters are in small capital letters.

```
(GET) PROPERTY SmallCaps () AS LONG
(SET) PROPERTY SmallCaps (BYVAL Value AS LONG)
```
```
FUNCTION GetSmallCaps () AS LONG
FUNCTION SetSmallCaps (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are in small capital letters. |
| **tomFalse** | Characters are not in small capital letters. |
| **tomToggle** | Toggle the state of the SmallCaps property. |
| **tomUndefined** | The SmallCaps property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="spacing"></a>Spacing

Gets/sets the amount of horizontal spacing between characters.

```
(GET) PROPERTY Spacing () AS SINGLE
(SET) PROPERTY Spacing (BYVAL Value AS SINGLE)
```
```
FUNCTION GetSpacing () AS SINGLE
FUNCTION SetSpacing (BYVAL Value AS SINGLE) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | The new amount of horizontal spacing between characters, in floating-point points. |

#### Return value

The amount of horizontal spacing between characters, in floating-point points.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

Displayed text typically has an intercharacter spacing value of zero. Positive values expand the spacing, and negative values compress it.

---

## <a name="strikethrough"></a>StrikeThrough

Gets/sets whether characters are displayed with a horizontal line through the center.

```
(GET) PROPERTY StrikeThrough () AS LONG
(SET) PROPERTY StrikeThrough (BYVAL Value AS LONG)
```
```
FUNCTION GetStrikeThrough () AS LONG
FUNCTION SetStrikeThrough (BYVAL Value AS LONG) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters have a horizontal line through the center. |
| **tomFalse** | Characters do not have a horizontal line through the center. |
| **tomToggle** | Toggle the state of the StrikeThrough property. |
| **tomUndefined** | The StrikeThrough property is undefined. |

#### Return value

A **tomBool** value that can be one of the following.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

This property corresponds to the **CFE_STRIKEOUT** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

## <a name="subscript"></a>Subscript

Gets/sets whether characters are displayed as subscript.

```
(GET) PROPERTY Subscript () AS LONG
(SET) PROPERTY Subscript (BYVAL Value AS LONG)
```
```
FUNCTION GetSubscript () AS LONG
SetSubscript (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed as subscript. |
| **tomFalse** | Characters are not displayed as subscript. |
| **tomToggle** | Toggle the state of the Subscript property. |
| **tomUndefined** | The Subscript property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

This property corresponds to the **CFE_SUBSCRIPT** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

## <a name="superscript"></a>Superscript

Gets/sets whether characters are displayed as superscript.

```
(GET) PROPERTY Superscript () AS LONG
(SET) PROPERTY Superscript (BYVAL Value AS LONG)
```
```
FUNCTION GetSuperscript () AS LONG
FUNCTION SetSuperscript (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed as superscript. |
| **tomFalse** | Characters are not displayed as superscript. |
| **tomToggle** | Toggle the state of the Superscript property. |
| **tomUndefined** | The Superscript property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed as superscript. |
| **tomFalse** | Characters are not displayed as superscript. |
| **tomUndefined** | The Superscript property is undefined. |

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

This property corresponds to the **CFE_SUPERSCRIPT** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

---

## <a name="underline"></a>Underline

Gets/sets the type of underlining for the characters in a range.

```
(GET) PROPERTY Underline () AS LONG
(SET) PROPERTY Underline (BYVAL Value AS LONG)
```
```
FUNCTION GetUnderline () AS LONG
FUNCTION SetUnderline (BYVAL Value AS LONG) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *Value* | The type of underlining. It can be one of the following values. |

| Constant | Value | Meaning |
| -------- | ----- | ------- |
| **tomNone** | 0 | No underlining. |
| **tomSingle** | 1 | Single underline. |
| **tomWords** | 2 | Underline words only. |
| **tomDouble** | 3 | Double underline. |
| **tomDash** | 5 | Dash underline. |
| **tomDashDot** | 6 | Dash dot underline. |
| **tomDashDotDot** | 7 | Dash dot dot underline. |
| **tomWave** | 8 | Wave underline. |
| **tomThick** | 9 | Thick underline. |
| **tomHair** | 10 | Hair underline. |
| **tomDoubleWave** | 11 | Double wave underline. |
| **tomHeavyWave** | 12 | Heavy wave underline. |
| **tomLongDash** | 13 | Long dash underline. |
| **tomThickDash** | 14 | Thick dash underline. |
| **tomThickDashDot** | 15 | Thick dash dot underline. |
| **tomThickDashDotDot** | 16 | Thick dash dot dot underline. |
| **tomThickDotted** | 17 | Thick dotted underline. |
| **tomThickLongDash** | 18 | Thick long dash underline. |

#### Return value

The type of underlining. It can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

---

## <a name="weight"></a>Weight

Gets/sets the font weight for the characters in a range.

```
(GET) PROPERTY Weight () AS LONG
(SET) PROPERTY Weight (BYVAL Value AS LONG)
```
```
FUNCTION GetWeight () AS LONG
FUNCTION SetWeight (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new font weight. The Bold property is a binary version of the Weight property that sets the weight to **FW_BOLD**. The font weight exists in the [LOGFONT](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-logfontw) structure and the [IFont](https://learn.microsoft.com/en-us/windows/win32/api/ocidl/nn-ocidl-ifont) interface. Windows defines the following degrees of font weight. |

| Font weight | Value |
| ----------- | ----- |
| **FW_DONTCARE** | 0 |
| **FW_THIN** | 100 |
| **FW_EXTRALIGHT** | 200 |
| **FW_LIGHT** | 300 |
| **FW_NORMAL** | 400 |
| **FW_MEDIUM** | 500 |
| **FW_SEMIBOLD** | 600 |
| **FW_BOLD** | 700 |
| **FW_EXTRABOLD** | 800 |
| **FW_HEAVY** | 900 |

#### Return value

The font weight. The **Bold** property is a binary version of the **Weight** property that sets the weight to **FW_BOLD**. The font weight exists in the [LOGFONT](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-logfontw) structure and the [IFont](https://learn.microsoft.com/en-us/windows/win32/api/ocidl/nn-ocidl-ifont) interface. Windows defines the following degrees of font weight (see table above).

#### Result code

If the method succeeds, it returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

## <a name="Count"></a>Count

Gets the count of extra properties in this character formatting collection.

```
FUNCTION Count () AS LONG
FUNCTION GetCount () AS LONG
```
#### Return value

The count of extra properties in this collection.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

## <a name="autoligatures"></a>AutoLigatures

Gets/sets whether support for automatic ligatures is active.

```
(GET) PROPERTY AutoLigatures () AS LONG
(SET) PROPERTY AutoLigatures (BYVAL Value AS LONG)
```
```
FUNCTION GetAutoLigatures () AS LONG
FUNCTION SetAutoLigatures (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Automatic ligature support is active. |
| **tomFalse** | Automatic ligature support is not active. |
| **tomToggle** | Toggle the AutoLigatures property. |
| **tomUndefined** | The AutoLigatures property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

## <a name="autospacealpha"></a>AutospaceAlpha

Gets/sets the East Asian "autospace alphabetics" state.

```
(GET) PROPERTY AutospaceAlpha () AS LONG
(SET) PROPERTY AutospaceAlpha (BYVAL Value AS LONG)
```
```
FUNCTION GetAutospaceAlpha () AS LONG
FUNCTION SetAutospaceAlpha (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Use East Asian autospace alphabetics. |
| **tomFalse** | Do not use East Asian autospace alphabetics. |
| **tomToggle** | Toggle the AutospaceAlpha property. |
| **tomUndefined** | The AutospaceAlpha property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="autospacenumeric"></a>AutospaceNumeric

Gets/sets the East Asian "autospace numeric" state.

```
(GET) PROPERTY AutospaceNumeric () AS LONG
(SET) PROPERTY AutospaceNumeric (BYVAL Value AS LONG)
```
```
FUNCTION GetAutospaceNumeric () AS LONG
FUNCTION SetAutospaceNumeric (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Use East Asian autospace numerics. |
| **tomFalse** | Do not use East Asian autospace numerics. |
| **tomToggle** | Toggle the AutospaceNumeric property. |
| **tomUndefined** | The AutospaceNumeric property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="autospaceparens"></a>AutospaceParens

Gets/sets the East Asian "autospace parentheses" state.

```
(GET) PROPERTY AutospaceParens () AS LONG
(SET) PROPERTY AutospaceParens (BYVAL Value AS LONG)
```
```
FUNCTION GetAutospaceParens () AS LONG
FUNCTION SetAutospaceParens (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Use East Asian autospace parentheses. |
| **tomFalse** | Do not use East Asian autospace parentheses. |
| **tomToggle** | Toggle the AutospaceParens property. |
| **tomUndefined** | The AutospaceParens property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="charrep"></a>CharRep

Gets/sets the character repertoire (writing system).

```
(GET) PROPERTY CharRep () AS LONG
(SET) PROPERTY CharRep (BYVAL Value AS LONG)
```
```
FUNCTION GetCharRep () AS LONG
FUNCTION SetCharRep (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new character repertoire. |

| Value | Meaning |
| ----- | ------- |
| **tomAboriginal** | Aboriginal |
| **tomAnsi** | Latin 1 |
| **tomArabic** | Arabic |
| **tomArmenian** | Armenian |
| **tomBaltic** | From Latin 1 and 2 |
| **tomBengali** | Bangla (formerly Bengali) |
| **tomBIG5** | Traditional Chinese |
| **tomBraille** | Braille |
| **tomCherokee** | Cherokee |
| **tomCyrillic** | Cyrillic |
| **tomDefaultCharRep** | Default character repertoire |
| **tomDevanagari** | Devanagari |
| **tomEastEurope** | From Latin 1 and 2 |
| **tomEmoji** | Emoji |
| **tomEthiopic** | Ethiopic |
| **tomGB2312** | Simplified Chinese |
| **tomGeorgian** | Georgian |
| **tomGreek** | Greek |
| **tomGujarati** | Gujarati |
| **tomGurmukhi** | Gurmukhi |
| **tomHangul** | Hangul |
| **tomHebrew** | Hebrew |
| **tomJamo** | Jamo |
| **tomKannada** | Kannada |
| **tomKayahli** | Kayah Li |
| **tomKharoshthi** | Kharoshthi |
| **tomKhmer** | Khmer |
| **tomLao** | Lao |
| **tomLimbu** | Limbu |
| **tomMac** | Main Macintosh character repertoire |
| **tomMalayalam** | Malayalam |
| **tomMongolian** | Mongolian |
| **tomMyanmar** | Myanmar |
| **tomNewTaiLu** | TaiLue |
| **tomOEM** | OEM character set (original PC) |
| **tomOriya** | Odia (formerly Oriya) |
| **tomPC437** | PC437 character set (DOS) |
| **tomRunic** | Runic |
| **tomShiftJIS** | Japanese |
| **tomSinhala** | Sinhala |
| **tomSylotinagr** | Syloti Nagri |
| **tomSymbol** | Symbol character set (not Unicode) |
| **tomSyriac** | Syriac |
| **tomTaiLe** | TaiLe |
| **tomTamil** | Tamil |
| **tomTelugu** | Telugu |
| **tomThaana** | Thaana |
| **tomThai** | Thai |
| **tomTibetan** | Tibetan |
| **tomTurkish** | Turkish (Latin 1 + dotless i, ...) |
| **tomVietnamese** | Latin 1 with some combining marks |
| **tomUsymbol** | Unicode symbol |
| **tomYi** | Yi |

#### Return value

The character repertoire. It can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="compressionmode"></a>CompressionMode

Gets/sets the East Asian compression mode.

```
(GET) PROPERTY CompressionMode () AS LONG
(SET) PROPERTY CompressionMode (BYVAL Value AS LONG)
```
```
FUNCTION GetCompressionMode () AS LONG
FUNCTION SetCompressionMode (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The compression mode, which can be one of these values: |

| Value | Meaning |
| ----- | ------- |
| **tomCompressNone** (default) | No compression. |
| **tomCompressPunctuation** | Compress punctuation. |
| **tomCompressPunctuationAndKana** | Compress punctuation and kana. |

#### Return value

The compression mode, which can be one of the ones listed above:

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="Cookie"></a>Cookie

Gets/sets the client cookie.

```
(GET) PROPERTY Cookie () AS LONG
(SET) PROPERTY Cookie (BYVAL Value AS LONG)
```
```
FUNCTION GetCookie () AS LONG
FUNCTION SetCookie (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The client cookie. |

#### Return value

The client cookie.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

This value is purely for the use of the client and has no meaning to the Text Object Model (TOM) engine. There are exceptions where different values correspond to different character format runs.

---

## <a name="doublestrike"></a>DoubleStrike

Gets/sets whether characters are displayed with double horizontal lines through the center.

```
(GET) PROPERTY DoubleStrike () AS LONG
(SET) PROPERTY DoubleStrike (BYVAL Value AS LONG)
```
```
FUNCTION GetDoubleStrike () AS LONG
FUNCTION SetDoubleStrike (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed with double horizontal lines. |
| **tomFalse** | Characters are not displayed with double horizontal lines. |
| **tomToggle** | Toggle the DoubleStrike property. |
| **tomUndefined** | The DoubleStrike property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="linktype"></a>LinkType

Gets the link type.

```
FUNCTION LinkType () AS LONG
FUNCTION GetLinkType () AS LONG
```
#### Return value

The link type. It can be one of the following values.

| Value | Meaning |
| ----- | ------- |
| **tomNoLink** | Not a link. |
| **tomClientLink** | The URL only; that is, no friendly name. |
| **tomFriendlyLinkName** | The name of friendly name link. |
| **tomFriendlyLinkAddress** | The URL of a friendly name link. |
| **tomAutoLinkURL** | The URL of an automatic link. |
| **tomAutoLinkEmail** | An automatic link to an email address. |
| **tomAutoLinkPhone** | An automatic link to a phone number. |
| **tomAutoLinkPath** | An automatic link to a storage location. |

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="mathzone"></a>MathZone

Gets/sets whether a math zone is active.

```
(GET) PROPERTY MathZone () AS LONG
(SET) PROPERTY MathZone (BYVAL Value AS LONG)
```
```
FUNCTION GetMathZone () AS LONG
FUNCTION SetMathZone (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | A math zone is active. |
| **tomFalse** | A math zone is not active. |
| **tomToggle** | Toggle the MathZone property. |
| **tomUndefined** | The MathZone property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="modwidthpairs"></a>ModWidthPairs

Gets/sets whether "decrease widths on pairs" is active.

```
(GET) PROPERTY ModWidthPairs () AS LONG
(SET) PROPERTY ModWidthPairs (BYVAL Value AS LONG)
```
```
FUNCTION GetModWidthPairs () AS LONG
FUNCTION SetModWidthPairs (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Decrease widths on pairs is active. |
| **tomFalse** | Decrease widths on pairs is not active. |
| **tomToggle** | Toggle the ModWidthPairs property. |
| **tomUndefined** | The ModWidthPairs property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="modwidthspace"></a>ModWidthSpace

Gets/sets whether "increase width of whitespace" is active.

```
(GET) PROPERTY ModWidthSpace () AS LONG
(SET) PROPERTY ModWidthSpace (BYVAL Value AS LONG)
```
```
FUNCTION GetModWidthSpace () AS LONG
FUNCTION SetModWidthSpace (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Increase width of whitespace is active. |
| **tomFalse** | Increase width of whitespace is not active. |
| **tomToggle** | Toggle the ModWidthSpace property. |
| **tomUndefined** | The ModWidthSpace property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="oldnumbers"></a>OldNumbers

Gets/sets whether old-style numbers are active.

```
(GET) PROPERTY OldNumbers () AS LONG
(SET) PROPERTY OldNumbers (BYVAL Value AS LONG)
```
```
FUNCTION GetOldNumbers () AS LONG
FUNCTION SetOldNumbers (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Old-style numbers are active. |
| **tomFalse** | Old-style numbers are not active. |
| **tomToggle** | Toggle the OldNumbers property. |
| **tomUndefined** | The OldNumbers property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="overlapping"></a>Overlapping

Gets/sets whether overlapping text is active.

```
(GET) PROPERTY Overlapping () AS LONG
(SET) PROPERTY Overlapping (BYVAL Value AS LONG)
```
```
FUNCTION GetOverlapping () AS LONG
FUNCTION SetOverlapping (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Overlapping text is active. |
| **tomFalse** | Overlapping text is not active. |
| **tomToggle** | Toggle the Overlapping property. |
| **tomUndefined** | The Overlapping property is undefined. |

#### Return value

A **tomBool** value that can be one of the ones listed above.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="positionsubsuper"></a>PositionSubSuper

Gets/sets the subscript or superscript position relative to the baseline.

```
(GET) PROPERTY PositionSubSuper () AS LONG
(SET) PROPERTY PositionSubSuper (BYVAL Value AS LONG)
```
```
FUNCTION GetPositionSubSuper () AS LONG
FUNCTION SetPositionSubSuper (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new subscript or superscript position. |

#### Return value

The subscript or superscript position relative to the baseline.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="scaling"></a>Scaling

Gets/sets the font horizontal scaling percentage.

```
(GET) PROPERTY Scaling () AS LONG
(SET) PROPERTY Scaling (BYVAL Value AS LONG)
```
```
FUNCTION GetScaling () AS LONG
FUNCTION SetScaling (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The scaling percentage. Values from 0 through 255 are valid. For example, a value of 200 doubles the widths of characters while retaining the same height. A value of 0 has the same effect as a value of 100; that is, it turns scaling off. |

#### Return value

The scaling percentage.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The font horizontal scaling percentage can range from 200, which doubles the widths of characters, to 0, where no scaling is performed. When the percentage is increased the height does not change.

---

## <a name="spaceextension"></a>SpaceExtension

Gets/sets the East Asian space extension value.

```
(GET) PROPERTY SpaceExtension () AS LONG
(SET) PROPERTY SpaceExtension (BYVAL Value AS LONG)
```
```
FUNCTION GetSpaceExtension () AS LONG
FUNCTION SetSpaceExtension (BYVAL Value AS LONG) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new space extension, in floating-points. |

#### Return value

The space extension, in floating-point points.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

---

## <a name="GetUnderlinePositionMode"></a>GetUnderlinePositionMode

Gets the underline position mode.

```
FUNCTION GetUnderlinePositionMode () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetUnderlinePositionMode(m_pTextFont2, @Value))
   RETURN Value
END FUNCTION
```
#### Return value

The underline position mode. It can be one of the following values.

| Value | Meaning |
| ----- | ------- |
| **tomUnderlinePositionAuto** | Automatically set the underline position. |
| **tomUnderlinePositionBelow** | Render underline below text. |
| **tomUnderlinePositionAbove** | Render underline above text. |

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetUnderlinePositionMode"></a>SetUnderlinePositionMode

Sets the underline position mode.

```
FUNCTION SetUnderlinePositionMode (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetUnderlinePositionMode(m_pTextFont2, Value))
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new underline position mode. It can be one of the following values. |

| Value | Meaning |
| ----- | ------- |
| **tomUnderlinePositionAuto** (the default) | Automatically set the underline position. |
| **tomUnderlinePositionBelow** | Render underline below text. |
| **tomUnderlinePositionAbove** | Render underline above text. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetEffects"></a>GetEffects

Gets the character format effects.

```
FUNCTION GetEffects (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->GetEffects(m_pTextFont2, pValue, pMask))
   RETURN m_Result
END FUNCTION
```

#### Parameters

| Parameter | Description |
| --------- | ----------- |
| *Value* | A combination of the following character format values. |

| Value | Meaning |
| ----- | ------- |
| **tomAllCaps** | All caps |
| **tomBold** | Boldface |
| **tomDisabled** | Disabled |
| **tomEmboss** | Emboss |
| **tomHidden** | Hidden |
| **tomImprint** | Imprint |
| **tomInlineObjectStart** | The start delimiter of an inline object |
| **tomItalic** | Italic |
| **tomLink** | Hyperlink |
| **tomLinkProtected** | The link is protected (friendly name link). |
| **tomMathZone** | Math zone |
| **tomMathZoneDisplay** | Display math zone |
| **tomMathZoneNoBuildUp** | Don't build up operator |
| **tomMathZoneOrdinary** | Math zone ordinary text. |
| **tomOutline** | Outline |
| **tomProtected** | Protected |
| **tomRevised** | Revised |
| **tomShadow** | Shadow |
| **tomSmallCaps** | Small caps |
| **tomStrikeout** | Strikeout |
| **tomUnderline** | Underline |

If the **tomInlineObjectStart** flag is set, you might want to call **GetInlineObject** for more inline object properties.

| Parameter | Description |
| --------- | ----------- |
| *pMask* | The differences in these flags over the range. A value of zero indicates that the properties are the same over the range. For an insertion point, this value is always zero. |

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetEffects2"></a>GetEffects2

Gets the additional character format effects.

```
FUNCTION GetEffects2 (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->GetEffects2(m_pTextFont2, pValue, pMask))
   RETURN m_Result
END FUNCTION
```

#### Parameters

| Parameter | Description |
| --------- | ----------- |
| *Value* | A combination of the following character format flags. |

| Value | Meaning |
| ----- | ------- |
| **tomAutoSpaceAlpha** | Use East Asian auto spacing between alphabetics. |
| **tomAutoSpaceNumeric** | Use East Asian auto spacing for digits. |
| **tomAutoSpaceParens** | Use East Asian automatic spacing for parentheses or brackets. |
| **tomDoublestrike** | Double strikeout. |
| **tomEmbeddedFont** | Embedded font (CLIP_EMBEDDED). |
| **tomModWidthPairs** | Use East Asian character-pair-width modification. |
| **tomModWidthSpace** | Use East Asian space-width modification. |
| **tomOverlapping** | Run has overlapping text. |

| Parameter | Description |
| --------- | ----------- |
| *pMask* | The differences in these flags over the range. Zero values indicate that the properties are the same over the range. For an insertion point, this value is always zero. |

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetProperty"></a>GetProperty

Gets the value of the specified property.

```
FUNCTION GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetProperty(m_pTextFont2, nType, @Value))
   RETURN Value
END FUNCTION
```

#### Parameters

| Parameter | Description |
| --------- | ----------- |
| *nType* | The property ID of the value to return. |

#### Return value

The property value.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetPropertyInfo"></a>GetPropertyInfo

Gets the property type and value of the specified extra property.

```
FUNCTION GetPropertyInfo (BYVAL Index AS LONG, BYVAL pType AS LONG PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->GetPropertyInfo(m_pTextFont2, Index, pType, pValue))
   RETURN m_Result
END FUNCTION
```

#### Parameters

| Parameter | Description |
| --------- | ----------- |
| *Index* | The collection index of the extra property. |
| *pType* | The property ID. |
| *pType* | The property value. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetEffects"></a>SetEffects

Sets the character format effects.

```
FUNCTION SetEffects (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetEffects(m_pTextFont2, Value, Mask))
   RETURN m_Result
END FUNCTION
```

#### Parameters

| Parameter | Description |
| --------- | ----------- |
| *Value* | A combination of the following character format values. |

| Value | Meaning |
| ----- | ------- |
| **tomAllCaps** | All caps |
| **tomBold** | Boldface |
| **tomDisabled** | Disabled |
| **tomEmboss** | Emboss |
| **tomHidden** | Hidden |
| **tomImprint** | Imprint |
| **tomInlineObjectStart** | The start delimiter of an inline object |
| **tomItalic** | Italic |
| **tomLink** | Hyperlink |
| **tomLinkProtected** | The link is protected (friendly name link). |
| **tomMathZone** | Math zone |
| **tomMathZoneDisplay** | Display math zone |
| **tomMathZoneNoBuildUp** | Don't build up operator |
| **tomMathZoneOrdinary** | Math zone ordinary text. |
| **tomOutline** | Outline |
| **tomProtected** | Protected |
| **tomRevised** | Revised |
| **tomShadow** | Shadow |
| **tomSmallCaps** | Small caps |
| **tomStrikeout** | Strikeout |
| **tomUnderline** | Underline |

| Parameter | Description |
| --------- | ----------- |
| *pMask* | The desired mask, which can be a combination of the Value flags. Only effects with the corresponding mask flag set are modified. |

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

Only effects with the corresponding mask flag set are modified.

# <a name="SetEffects2"></a>SetEffects2

Sets the additional character format effects.

```
FUNCTION SetEffects2 (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetEffects2(m_pTextFont2, Value, Mask))
   RETURN m_Result
END FUNCTION
```

#### Parameters

| Parameter | Description |
| --------- | ----------- |
| *Value* | A combination of the following character format values. |

| Value | Meaning |
| ----- | ------- |
| **tomAutoSpaceAlpha** | Use East Asian auto spacing between alphabetics. |
| **tomAutoSpaceNumeric** | Use East Asian auto spacing for digits. |
| **tomAutoSpaceParens** | Use East Asian automatic spacing for parentheses or brackets. |
| **tomDoublestrike** | Double strikeout. |
| **tomEmbeddedFont** | Embedded font (CLIP_EMBEDDED). |
| **tomModWidthPairs** | Use East Asian character-pair-width modification. |
| **tomModWidthSpace** | Use East Asian space-width modification. |
| **tomOverlapping** | Run has overlapping text. |

| Parameter | Description |
| --------- | ----------- |
| *pMask* | The desired mask, which can be a combination of the Value flags. Only effects with the corresponding mask flag set are modified. |

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

Only effects with the corresponding mask flag set are modified.

# <a name="SetProperty"></a>SetProperty

Sets the value of the specified property.

```
FUNCTION SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetProperty(m_pTextFont2, nType, Value))
   RETURN m_Result
END FUNCTION
```

#### Parameters

| Parameter | Description |
| --------- | ----------- |
| *nType* | The ID of the property value to set. |
| *Value* | The new property value. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.
