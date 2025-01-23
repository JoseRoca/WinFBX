# CTextFont2 Class

Class that wraps all the methods of the **ITextFont2** interface.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTORS](#CONSTRUCTORS) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [LET](#LET) | Assignment operator. |
| [CAST](#CAST) | Cast operator. |
| [TextDocumentPtr](#TextDocumentPtr) | Returns a pointer to the underlying **ITextFont2** interface. |
| [Attach](#Attach) | Attaches an **ITextFont2** interface pointer to the class. |
| [Detach](#Detach) | Detaches the underlying **ITextFont2** interface pointer from the class. |

### ITextFont Interface

Text Object Model (TOM) rich text-range attributes are accessed through a pair of dual interfaces, **ITextFont** and **ITextPara**.

The **ITextFont** interface inherits from the **IDispatch** interface. **ITextFont** also has these methods.

| Name       | Description |
| ---------- | ----------- |
| [GetDuplicate](#GetDuplicate) | Gets a duplicate of this text font object. |
| [SetDuplicate](#SetDuplicate) | Sets the character formatting by copying another text font object. |
| [CanChange](#CanChange) | Determines whether the font can be changed. |
| [IsEqual](#IsEqual) | Determines whether this text font object has the same properties as the specified text font object. |
| [Reset](#Reset) | Resets the character formatting to the specified values. |
| [GetStyle](#GetStyle) | Gets the character style handle of the characters in a range. |
| [SetStyle](#SetStyle) | Sets the character style handle of the characters in a range. |
| [GetAllCaps](#GetAllCaps) | Gets whether the characters are all uppercase. |
| [SetAllCaps](#SetAllCaps) | Sets whether the characters are all uppercase. |
| [GetAnimation](#GetAnimation) | Gets the animation type. |
| [SetAnimation](#SetAnimation) | Sets the animation type. |
| [GetBackColor](#GetBackColor) | Gets the text background (highlight) color. |
| [SetBackColor](#SetBackColor) | Sets the background color. |
| [GetBold](#GetBold) | Gets whether the characters are bold. |
| [SetBold](#SetBold) | Sets whether characters are bold. |
| [GetEmboss](#GetEmboss) | Gets whether characters are embossed. |
| [SetEmboss](#SetEmboss) | Sets whether characters are embossed. |
| [GetForeColor](#GetForeColor) | Gets the foreground, or text, color. |
| [SetForeColor](#SetForeColor) | Sets the foreground (text) color. |
| [GetHidden](#GetHidden) | Gets whether characters are hidden. |
| [SetHidden](#SetHidden) | Sets whether characters are hidden. |
| [GetEngrave](#GetEngrave) | Gets whether characters are displayed as imprinted characters. |
| [SetEngrave](#SetEngrave) | Sets whether characters are displayed as imprinted characters. |
| [GetItalic](#GetItalic) | Gets whether characters are in italics. |
| [SetItalic](#SetItalic) | Sets whether characters are in italics. |
| [GetKerning](#GetKerning) | Gets the minimum font size at which kerning occurs. |
| [SetKerning](#SetKerning) | Sets the minimum font size at which kerning occurs. |
| [GetLanguageID](#GetLanguageID) | Gets the language ID or language code identifier (LCID). |
| [SetLanguageID](#SetLanguageID) | Sets the language ID or language code identifier (LCID). |
| [GetName](#GetName) | Gets the font name. |
| [SetName](#SetName) | Sets the font name. |
| [GetOutline](#GetOutline) | Gets whether characters are displayed as outlined characters. |
| [SetOutline](#SetOutline) | Sets whether characters are displayed as outlined characters. |
| [GetPosition](#GetPosition) | Gets the amount that characters are offset vertically relative to the baseline. |
| [SetPosition](#SetPosition) | Sets the amount that characters are offset vertically relative to the baseline. |
| [GetProtected](#GetProtected) | Gets whether characters are protected against attempts to modify them. |
| [SetProtected](#SetProtected) | Sets whether characters are protected against attempts to modify them. |
| [GetShadow](#GetShadow) | Gets whether characters are displayed as shadowed characters. |
| [SetShadow](#SetShadow) | Sets whether characters are displayed as shadowed characters. |
| [GetSize](#GetSize) | Gets the font size. |
| [SetSize](#SetSize) | Sets the font size. |
| [GetSmallCaps](#GetSmallCaps) | Gets whether characters are in small capital letters. |
| [SetSmallCaps](#SetSmallCaps) | Sets whether characters are in small capital letters. |
| [GetSpacing](#GetSpacing) | Gets the amount of horizontal spacing between characters. |
| [SetSpacing](#SetSpacing) | Sets the amount of horizontal spacing between characters. |
| [GetStrikeThrough](#GetStrikeThrough) | Gets whether characters are displayed with a horizontal line through the center. |
| [SetStrikeThrough](#SetStrikeThrough) | Sets whether characters are displayed with a horizontal line through the center. |
| [GetSubscript](#GetSubscript) | Gets whether characters are displayed as subscript. |
| [SetSubscript](#SetSubscript) | Sets whether characters are displayed as subscript. |
| [GetSuperscript](#GetSuperscript) | Gets whether characters are displayed as superscript. |
| [SetSuperscript](#SetSuperscript) | Sets whether characters are displayed as superscript. |
| [GetUnderline](#GetUnderline) | Gets the type of underlining for the characters in a range. |
| [SetUnderline](#SetUnderline) | Sets the type of underlining for the characters in a range. |
| [GetWeight](#GetWeight) | Gets the font weight for the characters in a range. |
| [SetWeight](#SetWeight) | Sets the font weight for the characters in a range. |

### ITextFont2 Interface

In the Text Object Model (TOM), applications access text-range attributes by using a pair of dual interfaces, **ITextFont** and **ITextPara**.

The **ITextFont2** interface extends **ITextFont**, providing the programming equivalent of the Microsoft Word format-font dialog.

| Name       | Description |
| ---------- | ----------- |
| [GetCount](#GetCount) | Gets the count of extra properties in this character formatting collection. |
| [GetAutoLigatures](#GetAutoLigatures) | Gets whether support for automatic ligatures is active. |
| [SetAutoLigatures](#SetAutoLigatures) | Sets whether support for automatic ligatures is active. |
| [GetAutospaceAlpha](#GetAutospaceAlpha) | Gets the East Asian "autospace alphabetics" state. |
| [SetAutospaceAlpha](#SetAutospaceAlpha) | Sets the East Asian "autospace alphabetics" state. |
| [GetAutospaceNumeric](#GetAutospaceNumeric) | Gets the East Asian "autospace numeric" state. |
| [SetAutospaceNumeric](#SetAutospaceNumeric) | Sets the East Asian "autospace numeric" state. |
| [GetAutospaceParens](#GetAutospaceParens) | Gets the East Asian "autospace parentheses" state. |
| [SetAutospaceParens](#SetAutospaceParens) | Sets the East Asian "autospace parentheses" state. |
| [GetCharRep](#GetCharRep) | Gets the character repertoire (writing system). |
| [SetCharRep](#SetCharRep) | Sets the character repertoire (writing system). |
| [GetCompressionMode](#GetCompressionMode) | Gets the East Asian compression mode. |
| [SetCompressionMode](#SetCompressionMode) | Sets the East Asian compression mode. |
| [GetCookie](#GetCookie) | Gets the client cookie. |
| [SetCookie](#SetCookie) | Sets the client cookie. |
| [GetDoubleStrike](#GetDoubleStrike) | Gets whether characters are displayed with double horizontal lines through the center. |
| [SetDoubleStrike](#SetDoubleStrike) | Sets whether characters are displayed with double horizontal lines through the center. |
| [GetDuplicate2](#GetDuplicate) | Gets a duplicate of this character format object. |
| [SetDuplicate2](#SetDuplicate) | Sets the properties of this object by copying the properties of another text font object. |
| [GetLinkType](#GetLinkType) | Gets the link type. |
| [GetMathZone](#GetMathZone) | Gets whether a math zone is active. |
| [SetMathZone](#SetMathZone) | Sets whether a math zone is active. |
| [GetModWidthPairs](#GetModWidthPairs) | Gets whether "decrease widths on pairs" is active. |
| [SetModWidthPairs](#SetModWidthPairs) | Sets whether "decrease widths on pairs" is active. |
| [GetModWidthSpace](#GetModWidthSpace) | Gets whether "increase width of whitespace" is active. |
| [SetModWidthSpace](#SetModWidthSpace) | Sets whether "increase width of whitespace" is active. |
| [GetOldNumbers](#GetOldNumbers) | Gets whether old-style numbers are active. |
| [SetOldNumbers](#SetOldNumbers) | Sets whether old-style numbers are active. |
| [GetOverlapping](#GetOverlapping) | Gets whether overlapping text is active. |
| [SetOverlapping](#SetOverlapping) | Sets whether overlapping text is active. |
| [GetPositionSubSuper](#GetPositionSubSuper) | Gets the subscript or superscript position relative to the baseline. |
| [SetPositionSubSuper](#SetPositionSubSuper) | Sets the position of a subscript or superscript relative to the baseline, as a percentage of the font height. |
| [GetScaling](#GetScaling) | Gets the font horizontal scaling percentage. |
| [SetScaling](#SetScaling) | Sets the font horizontal scaling percentage. |
| [GetSpaceExtension](#GetSpaceExtension) | Gets the East Asian space extension value. |
| [SetSpaceExtension](#SetSpaceExtension) | Sets the East Asian space extension value. |
| [GetUnderlinePositionMode](#GetUnderlinePositionMode) | Gets the underline position mode. |
| [SetUnderlinePositionMode](#SetUnderlinePositionMode) | Sets the underline position mode. |
| [GetEffects](#GetEffects) | Gets the character format effects. |
| [GetEffects2](#GetEffects2) | Gets the additional character format effects. |
| [GetProperty](#GetProperty) | Gets the value of the specified property. |
| [GetPropertyInfo](#GetPropertyInfo) | Gets the property type and value of the specified extra property. |
| [IsEqual2](#IsEqual2) | The text font object to compare against. |
| [SetEffects](#SetEffects) | Sets the character format effects. |
| [SetEffects2](#SetEffects2) | Sets the additional character format effects. |
| [SetProperty](#SetProperty) | Sets the value of the specified property. |

### Methods inherited from CTOMBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#GetLastResult) | Returns the last result code |
| [SetResult](#SetResult) | Sets the last result code. |
| [GetErrorInfo](#GetErrorInfo) | Returns a description of the last result code. |

# <a name="CONSTRUCTORS"></a>CONSTRUCTORS

Called when a **CTextFont2** class variable is created.

```
DECLARE CONSTRUCTOR
DECLARE CONSTRUCTOR (BYVAL pTextFont2 AS ITextFont2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

## CONSTRUCTOR (Empty)

Can be used, for example, when we have an **pTextFont2** interface pointer returned by a function and we want to attach it to a new instance of the **CTextFont2** class.

```
DIM DIM pCTextFont2 AS CTextFont2
pCTextFont2.Attach(pTextFont2)
```
## CONSTRUCTOR (ITextFont2 PTR)

```
CONSTRUCTOR CTextFont2 (BYVAL pTextFont2 AS ITextFont2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Store the pointer of ITextFont2 interface
   IF pTextFont2 THEN
      IF fAddRef THEN pTextFont2->lpvtbl->AddRef(pTextFont2)
   END IF
   m_pTextFont2 = pTextFont2
END CONSTRUCTOR
```

| Parameter | Description |
| --------- | ----------- |
| *pTextFont2* | An **ITextFont2** interface pointer. |
| *fAddRef* | Optional. **TRUE** to increment the reference count of the passed **ITextFont2** interface pointer; otherwise, **FALSE**. Default is **FALSE**. |

#### Return value

A pointer to the new instance of the class.

# <a name="DESTRUCTOR"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

```
DESTRUCTOR CTextFont2
   ' // Release the ITextFont2 interface
   IF m_pTextFont2 THEN m_pTextFont2->lpvtbl->Release(m_pTextFont2)
END DESTRUCTOR
```

# <a name="LET"></a>LET

Assignment operator. The assigned pointer must be an "addrefed" one.

```
OPERATOR CTextFont2.LET (BYVAL pTextFont2 AS ITextFont2 PTR)
   m_Result = 0
   IF pTextFont2 = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextFont2 THEN m_pTextFont2->lpvtbl->Release(m_pTextFont2)
   ' // Attach the passed interface pointer to the class
   m_pTextFont2 = pTextFont2
END OPERATOR
```

# <a name="CAST"></a>CAST

Cast operator.

```
OPERATOR CTextFont2.CAST () AS ITextFont2 PTR
   m_Result = 0
   OPERATOR = m_pTextFont2
END OPERATOR
```

# <a name="TextFontPtr"></a>TextFontPtr

Returns a pointer to the underlying **ITextFont2** interface

```
FUNCTION CTextFont2.TextFont2Ptr () AS ITextFont2 PTR
   m_Result = 0
   RETURN m_pTextFont2
END FUNCTION
```

# <a name="Attach"></a>Attach

Attaches an **ITextFont2** interface pointer to the class.

```
PRIVATE FUNCTION CTextFont2.Attach (BYVAL pTextFont2 AS ITextFont2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextFont2 = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextFont2 THEN m_Result = m_pTextFont2->lpvtbl->Release(m_pTextFont2)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextFont2->lpvtbl->AddRef(pTextFont2)
   m_pTextFont2 = pTextFont2
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pTextFont2* | The **ITextFont2** interface pointer to attach. |
| *fAddRef* | **TRUE** to increment the reference count of te object; otherwise, **FALSE**. Default is **FALSE**. |

# <a name="GetLastResult"></a>GetLastResult

Returns the last result code

```
FUNCTION CTOMBase.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
```

# <a name="SetResult"></a>SetResult

Sets the last result code.

```
FUNCTION CTOMBase.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Result* | The **HRESULT** error code returned by the methods. |

# <a name="GetErrorInfo"></a>GetErrorInfo

Returns a description of the last result code.

```
FUNCTION CTOMBase.GetErrorInfo () AS CWSTR
   IF SUCCEEDED(m_Result) THEN RETURN "Success"
   DIM s AS CWSTR = "Error &h" & HEX(m_Result, 8)
   SELECT CASE m_Result
      CASE E_POINTER : s += ": E_POINTER - Null pointer"
      CASE S_OK : s += ": S_OK - Success"
      CASE S_FALSE : s += ": S_FALSE - Failure"
      CASE E_NOTIMPL : s += ": E_NOTIMPL - Not implemented."
      CASE E_INVALIDARG : s += ": E_INVALIDARG - Invalid argument"
      CASE E_OUTOFMEMORY : s += ": E_OUTOFMEMORY - Insufficient memory"
'      CASE E_FILENOTFOUND : s += "E_FILENOTFOUND - File not found"
      CASE &h80070002 : s += "E_FILENOTFOUND - File not found"
      CASE E_ACCESSDENIED : s += "E_ACCESSDENIED - Access denied"
      CASE E_FAIL : s += ": E_FAIL - Access denied"
      CASE NOERROR : s += ": NOERROR - Success" '' (same as S_OK)
      CASE CO_E_RELEASED:  : s += ": CO_E_RELEASED: - The object has been released"
      CASE ELSE
         s += "Unknown error"
   END SELECT
   RETURN s
END FUNCTION
```

# <a name="GetDuplicate"></a>GetDuplicate

Gets a duplicate of this range object. In this implementation of the class, **GetDuplicate** and **GetDuplicate2** are the same method.

```
FUNCTION CTextFont2.GetDuplicate () AS ITextFont2 PTR
   DIM pFont AS ITextFont2 PTR
   this.SetResult(m_pTextFont2->lpvtbl->GetDuplicate(m_pTextFont2, @pFont))
   FUNCTION = pFont
END FUNCTION
```
```
FUNCTION CTextFont2.GetDuplicate2 () AS ITextFont2 PTR
   DIM pFont AS ITextFont2 PTR
   this.SetResult(m_pTextFont2->lpvtbl->GetDuplicate2(m_pTextFont2, @pFont))
   FUNCTION = pFont
END FUNCTION
```

#### Return value

The duplicate text font object.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Result code | Description |
| ----------- | ----------- |
| **E_OUTOFMEMORY** | Memory could not be allocated for the new object. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetDuplicate"></a>SetDuplicate

Sets the character formatting by copying another text font object. In this implementation of the class, **SetDuplicate** and **SetDuplicate2** are the same method.

```
FUNCTION CTextFont2.SetDuplicate (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetDuplicate(m_pTextFont2, pFont))
   FUNCTION = m_Result
END FUNCTION
```
```
FUNCTION CTextFont2.SetDuplicate2 (BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetDuplicate2(m_pTextFont2, pFont))
   FUNCTION = m_Result
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *pFont* | The text font object to apply to this font object. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

Values with the **tomUndefined** attribute have no effect.

# <a name="CanChange"></a>CanChange

Determines whether the font can be changed.

```
FUNCTION CTextFont2.CanChange () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->CanChange(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

A variable that is **tomTrue** if the font can be changed or **tomFalse** if it cannot be changed.

#### Result code

If the font can change, the method returns **S_OK**. If the method fails, it returns **S_FALSE**.

#### Remarks

This method returns **tomTrue** only if the font can be changed. That is, no part of an associated range is protected and an associated document is not read-only. If this **ITextFont** object is a duplicate, no protection rules apply.

# <a name="IsEqual"></a>IsEqual

Determines whether this text font object has the same properties as the specified text font object.

```
FUNCTION CTextFont2.IsEqual (BYVAL pFont AS ITextFont2 PTR) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->IsEqual(m_pTextFont2, pFont, @Value))
   FUNCTION = Value
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pFont* | The text font object to compare against. |

#### Return value

A variable that is **tomTrue** if the font objects have the same properties or **tomFalse** if they do not.

#### Result code

If the text font objects have the same properties, the method succeeds and returns **S_OK**. If the text font objects do not have the same properties, the method fails and returns **S_FALSE**.

#### Remarks

The text font objects are equal only if *pFont* belongs to the same Text Object Model (TOM) object as the current font object. The **IsEqual** method ignores entries for which either font object has an **tomUndefined**.

# <a name="Reset"></a>Reset

Resets the character formatting to the specified values.

```
FUNCTION CTextFont2.Reset (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->Reset(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

If the method succeeds, it returns **S_OK^^. If the method fails, it returns one of the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| **S_FALSE** | Protected from change. |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

Calling **Reset** with **tomUndefined** sets all properties to undefined values. Thus, applying the font object to a range changes nothing. This applies to a font object that is obtained by the **GetDuplicate** method.

# <a name="GetStyle"></a>GetStyle

Gets the character style handle of the characters in a range.

```
FUNCTION CTextFont2.GetStyle () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetStyle(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The character style handle.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error code.

| Return code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

The Text Object Model (TOM) version 1.0 does not specify the meanings of the style handles. The meanings depend on other facilities of the text system that implements TOM.

# <a name="SetStyle"></a>SetStyle

Sets the character style handle of the characters in a range.

```
FUNCTION CTextFont2.SetStyle (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetStyle(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new character style handle. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

The Text Object Model (TOM) version 1.0 does not specify the meanings of the style handles. The meanings depend on other facilities of the text system that implements TOM.

# <a name="GetAllCaps"></a>GetAllCaps

Gets whether the characters are all uppercase.

```
FUNCTION CTextFont2.GetAllCaps () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAllCaps(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are all uppercase. |
| **tomFalse** | Characters are not all uppercase. |
| **tomUndefined** | The AllCaps property is undefined. |
	
#### Return code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns one of the following COM error code.

| Return code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the **CFE_ALLCAPS** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

# <a name="SetAllCaps"></a>SetAllCaps

Sets whether the characters are all uppercase.

```
FUNCTION CTextFont2.SetAllCaps (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAllCaps(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are all uppercase. |
| **tomFalse** | Characters are not all uppercase. |
| **tomToggle** | Toggle the state of the AllCaps property. |
| **tomUndefined** | The AllCaps property is undefined. |

#### Return value

If the method succeeds, it returns **S_OK**. If the method fails, it returns one of the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetAnimation"></a>GetAnimation

Gets the animation type.

```
FUNCTION CTextFont2.GetAnimation () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAnimation(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

One of the following animation types.

| Constant | Value | Meaning |
| -------- | ----- | ------- |
| **tomNoAnimation** | 0 | Do not apply text animation. |
| **tomLasVegasLights** | 1 | Text is bordered by marquee lights that blink between the colors red, yellow, green, and blue. |
| **tomBlinkingBackground** | 2 | Text has a black background that blinks on and off. |
| **tomSparkleText** | 3 | Text is overlaid with multicolored stars that blink on and off at regular intervals. |
| **tomMarchingBlackAnts** | 4 | Text is surrounded by a black dashed-line border. The border is animated so that the individual dashes appear to move clockwise around the text. |
| **tomMarchingRedAnts** | 5 | Text is surrounded by a red dashed-line border that is animated to appear to move clockwise around the text. |
| **tomShimmer** | 6 | Text is alternately blurred and unblurred at regular intervals, to give the appearance of shimmering. |
| **tomWipeDown** | 7 | Text appears gradually from the top down. |
| **tomWipeRight** | 8 | Text appears gradually from the bottom up. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error codes.

| Return code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetAnimation"></a>SetAnimation

Sets the animation type.

```
FUNCTION CTextFont2.SetAnimation (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAnimation(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

The animation type. It can be one of the following values.

| Abimation type | Value | Meaning |
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

If the method succeeds, it returns **S_OK**. If the method fails, it returns an **HRESULT** COM error code.

# <a name="GetBackColor"></a>GetBackColor

Gets the text background (highlight) color.

```
FUNCTION CTextFont2.GetBackColor () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetBackColor(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The text background color. It can be one of the following values.

| Value | Meaning |
| ----- | ------- |
| A **COLORREF** value | The high-order byte is zero, and the three low-order bytes specify an RGB color. |
| A value returned by **PALETTEINDEX** | The high-order byte is 1, and the **LOWORD** specifies the index of a logical-color palette entry. |
| **tomAutocolor** (-9999997) | Indicates the range uses the default system background color. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetBackColor"></a>SetBackColor

Sets the background color.

```
FUNCTION CTextFont2.SetBackColor (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetBackColor(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetBold"></a>GetBold

Gets whether the characters are bold.

```
FUNCTION CTextFont2.GetBold () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetBold(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are bold. |
| **tomFalse** | Characters are not bold. |
| **tomUndefined** | The Bold property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetBold"></a>SetBold

Sets whether characters are bold.

```
FUNCTION CTextFont2.SetBold (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetBold(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetEmboss"></a>GetEmboss

Gets whether the characters are embossed.

```
FUNCTION CTextFont2.GetEmboss () AS LONG
   DIM Value AS LONG
   IF m_pTextFont2 = NULL THEN m_Result = E_POINTER: RETURN Value
   this.SetResult(m_pTextFont2->lpvtbl->GetEmboss(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are embossed. |
| **tomFalse** | Characters are not embossed. |
| **tomUndefined** | The Emboss property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the **CFE_EMBOSS** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

# <a name="SetEmboss"></a>SetEmboss

Sets whether characters are embossed.

```
FUNCTION CTextFont2.SetEmboss (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetEmboss(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetForeColor"></a>GetForeColor

Gets the foreground, or text, color.

```
FUNCTION CTextFont2.GetForeColor () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetForeColor(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The foreground color. It can be one of the following values.

| Value | Meaning |
| ----- | ------- |
| A **COLORREF** value | The high-order byte is zero, and the three low-order bytes specify an RGB color. |
| A value returned by **PALETTEINDEX** | The high-order byte is 1, and the **LOWORD** specifies the index of a logical-color palette entry. |
| **tomAutocolor** (-9999997) | Indicates the range uses the default system foreground color. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetForeColor"></a>SetForeColor

Sets the foreground (text) color.

```
FUNCTION CTextFont2.SetForeColor (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetForeColor(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetHidden"></a>GetHidden

Gets whether characters are hidden.

```
FUNCTION CTextFont2.GetHidden () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetHidden(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are hidden. |
| **tomFalse** | Characters are not hidden. |
| **tomToggle** | Toggle the state of the Hidden property. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the **CFE_HIDDEN** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

# <a name="SetHidden"></a>SetHidden

Sets whether characters are hidden.

```
FUNCTION CTextFont2.SetHidden (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetHidden(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetEngrave"></a>GetEngrave

Gets whether characters are displayed as imprinted characters.

```
FUNCTION CTextFont2.GetEngrave () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetEngrave(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed as imprinted characters. |
| **tomFalse** | Characters are not displayed as imprinted characters. |
| **tomUndefined** | The Engrave property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the **CFE_IMPRINT** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

# <a name="SetEngrave"></a>SetEngrave

Sets whether characters are displayed as imprinted characters.

```
FUNCTION CTextFont2.SetEngrave (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetEngrave(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetItalic"></a>GetItalic

Gets whether characters are in italics.

```
FUNCTION CTextFont2.GetItalic () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetItalic(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are in italics. |
| **tomFalse** | Characters are not in italics. |
| **tomUndefined** | The Italic property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetItalic"></a>SetItalic

Sets whether characters are in italics.

```
FUNCTION CTextFont2.SetItalic (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetItalic(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetKerning"></a>GetKerning

Gets the minimum font size at which kerning occurs.

```
FUNCTION CTextFont2.GetKerning () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextFont2->lpvtbl->GetKerning(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The minimum font size at which kerning occurs, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetKerning"></a>SetKerning

Sets the minimum font size at which kerning occurs.

```
FUNCTION CTextFont2.SetKerning (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetKerning(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new value of the minimum kerning size, in floating-point points. |

If the value is zero, kerning is turned off. Positive values turn on pair kerning for font sizes greater than this kerning value. For example, the value 1 turns on kerning for all legible sizes, whereas 16 turns on kerning only for font sizes of 16 points and larger.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetLanguageID"></a>GetLanguageID

Gets the language ID or language code identifier (LCID).

```
FUNCTION CTextFont2.GetLanguageID () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetLanguageID(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The language ID or LCID. The low word contains the language identifier. The high word is either zero or it contains the high word of the LCID. To retrieve the language identifier, mask out the high word. For more information, see [Locale Identifiers](https://learn.microsoft.com/en-us/windows/win32/intl/locale-identifiers).

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetLanguageID"></a>SetLanguageID

Sets the language ID or language code identifier (LCID).

```
FUNCTION CTextFont2.SetLanguageID (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetLanguageID(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new value of the minimum kerning size, in floating-point points. |

The new language identifier. The low word contains the language identifier. The high word is either zero or it contains the high word of the locale identifier LCID. For more information, see [Locale Identifiers](https://learn.microsoft.com/en-us/windows/win32/intl/locale-identifiers).

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

# <a name="GetName"></a>GetName

Gets the font name.

```
FUNCTION CTextFont2.GetName () AS CBSTR
   DIM pName AS AFX_BSTR
   this.SetResult(m_pTextFont2->lpvtbl->GetName(m_pTextFont2, @pName))
   RETURN pName
END FUNCTION
```
#### Return value

The font name.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_OUTOFMEMORY** | Could not allocate memory for string. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetName"></a>SetName

Sets the new font name.

```
FUNCTION CTextFont2.SetName (BYVAL bstr AS BSTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetName(m_pTextFont2, bstr))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *bstr* | The new font name. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetOutline"></a>GetOutline

Gets whether characters are displayed as outlined characters.

```
FUNCTION CTextFont2.GetOutline () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetOutline(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed as outlined characters. |
| **tomFalse** | Characters are not displayed as outlined characters. |
| **tomUndefined** | The Outline property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetOutline"></a>SetOutline

Sets whether characters are displayed as outlined characters.

```
FUNCTION CTextFont2.SetOutline (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetOutline(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetPosition"></a>GetPosition

Gets the amount that characters are offset vertically relative to the baseline.

```
FUNCTION CTextFont2.GetPosition () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextFont2->lpvtbl->GetPosition(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The amount of vertical offset, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

Displayed text typically has a zero value for this property. Positive values raise the text, and negative values lower it.

# <a name="SetPosition"></a>SetPosition

Sets the amount that characters are offset vertically relative to the baseline.

```
FUNCTION CTextFont2.SetPosition (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetPosition(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new amount of vertical offset, in floating-point points. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

Displayed text typically has a zero value for this property. Positive values raise the text, and negative values lower it.

# <a name="GetProtected"></a>GetProtected

Gets whether characters are protected against attempts to modify them.

```
FUNCTION CTextFont2.GetProtected () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetProtected(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are protected. |
| **tomFalse** | Characters are not protected. |
| **tomUndefined** | The Protected property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

In general, Text Object Model (TOM) methods that attempt to change the formatting or content of a range fail with **E_ACCESSDENIED** if any part of that range is protected, or if the document is read only. To make a change in protected text, the TOM client should attempt to turn off the protection of the text to be modified. The owner of the document may permit this to happen. For example in rich edit controls, attempts to change protected text result in an **EN_PROTECTED** notification code to the creator of the document, who then can refuse or grant permission for the change. The creator is the client that created a windowed rich edit control through the **CreateWindowEx** function or the **ITextHost** object that called the **CreateTextServices** function to create a windowless rich edit control.

# <a name="SetProtected"></a>SetProtected

Sets whether characters are protected against attempts to modify them.

```
FUNCTION CTextFont2.SetProtected (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetProtected(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are protected. |
| **tomFalse** | Characters are not protected. |
| **tomToggle** | Toggle the state of the Protected property. |
| **tomUndefined** | The Protected property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetShadow"></a>GetShadow

Gets whether characters are displayed as shadowed characters.

```
FUNCTION CTextFont2.GetShadow () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetShadow(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed as shadowed characters. |
| **tomFalse** | Characters are not displayed as shadowed characters. |
| **tomUndefined** | The Shadow property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the **CFE_SHADOW** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

# <a name="SetShadow"></a>SetShadow

Sets whether characters are displayed as shadowed characters.

```
FUNCTION CTextFont2.SetShadow (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetShadow(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetSize"></a>GetSize

Gets the font size.

```
FUNCTION CTextFont2.GetSize () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextFont2->lpvtbl->GetSize(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The font size, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetSize"></a>SetSize

Sets the font size.

```
FUNCTION CTextFont2.SetSize (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSize(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new font size, in floating-point points. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetSmallCaps"></a>GetSmallCaps

Gets whether characters are in small capital letters.

```
FUNCTION CTextFont2.GetSmallCaps () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetSmallCaps(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are in small capital letters. |
| **tomFalse** | Characters are not in small capital letters. |
| **tomUndefined** | The SmallCaps property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetSmallCaps"></a>SetSmallCaps

Sets whether characters are in small capital letters.

```
FUNCTION CTextFont2.SetSmallCaps (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSmallCaps(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetSpacing"></a>GetSpacing

Gets the amount of horizontal spacing between characters.

```
FUNCTION CTextFont2.GetSpacing () AS SINGLE
   DIM Value AS SINGLE
   this.SetResult(m_pTextFont2->lpvtbl->GetSpacing(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The amount of horizontal spacing between characters, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetSpacing"></a>SetSpacing

Sets the amount of horizontal spacing between characters.

```
FUNCTION CTextFont2.SetSpacing (BYVAL Value AS SINGLE) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSpacing(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new amount of horizontal spacing between characters, in floating-point points. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

#### Remarks

Displayed text typically has an intercharacter spacing value of zero. Positive values expand the spacing, and negative values compress it.

# <a name="GetStrikeThrough"></a>GetStrikeThrough

Gets whether characters are displayed with a horizontal line through the center.

```
FUNCTION CTextFont2.GetStrikeThrough () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetStrikeThrough(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed with a horizontal line through the center. |
| **tomFalse** | Characters are not displayed with a horizontal line through the center. |
| **tomUndefined** | The StrikeThrough property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the **CFE_STRIKEOUT** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

# <a name="SetStrikeThrough"></a>SetStrikeThrough

Sets whether characters are displayed with a horizontal line through the center.

```
FUNCTION CTextFont2.SetStrikeThrough (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetStrikeThrough(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetSubscript"></a>GetSubscript

Gets whether characters are displayed as subscript.

```
FUNCTION CTextFont2.GetSubscript () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetSubscript(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed as subscript. |
| **tomFalse** | Characters are not displayed as subscript. |
| **tomUndefined** | The Subscript property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the **CFE_SUBSCRIPT** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

# <a name="SetSubscript"></a>SetSubscript

Sets whether characters are displayed as subscript.

```
FUNCTION CTextFont2.SetSubscript (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSubscript(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetSuperscript"></a>GetSuperscript

Gets whether characters are displayed as superscript.

```
FUNCTION CTextFont2.GetSuperscript () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetSuperscript(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed as superscript. |
| **tomFalse** | Characters are not displayed as superscript. |
| **tomUndefined** | The Superscript property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

#### Remarks

This property corresponds to the **CFE_SUPERSCRIPT** effect described in the [CHARFORMAT2](https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1) structure.

# <a name="SetSuperscript"></a>SetSuperscript

Sets whether characters are displayed as superscript.

```
FUNCTION CTextFont2.SetSuperscript (BYVAL Value AS LONG) AS HRESULT
   IF m_pTextFont2 = NULL THEN m_Result = E_POINTER: RETURN m_Result
   this.SetResult(m_pTextFont2->lpvtbl->SetSuperscript(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetUnderline"></a>GetUnderline

Gets the type of underlining for the characters in a range.

```
FUNCTION CTextFont2.GetUnderline () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetUnderline(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The type of underlining. It can be one of the following values.

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

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetUnderline"></a>SetUnderline

Sets thevtype of underlining for the characters in a range.

```
FUNCTION CTextFont2.SetUnderline (BYVAL Value AS LONG) AS HRESULT
   IF m_pTextFont2 = NULL THEN m_Result = E_POINTER: RETURN m_Result
   this.SetResult(m_pTextFont2->lpvtbl->SetUnderline(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetWeight"></a>GetWeight

Gets the font weight for the characters in a range.

```
FUNCTION CTextFont2.GetWeight () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetWeight(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The font weight. The Bold property is a binary version of the Weight property that sets the weight to **FW_BOLD**. The font weight exists in the [LOGFONT](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-logfontw) structure and the [IFont](https://learn.microsoft.com/en-us/windows/win32/api/ocidl/nn-ocidl-ifont) interface. Windows defines the following degrees of font weight.

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

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |

# <a name="SetWeight"></a>SetWeight

Sets the font weight for the characters in a range.

```
FUNCTION CTextFont2.SetWeight (BYVAL Value AS LONG) AS HRESULT
   IF m_pTextFont2 = NULL THEN m_Result = E_POINTER: RETURN m_Result
   this.SetResult(m_pTextFont2->lpvtbl->SetWeight(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, **GetLastResult** returns **S_OK**. If the method fails, it returns the following COM error code.

| Result code | Description |
| ----------- | ----------- |
| **E_INVALIDARG** | Invalid argument. |
| **CO_E_RELEASED** | The font object is attached to a range that has been deleted. |
| **E_ACCESSDENIED** | Write access is denied. |
| **E_OUTOFMEMORY** | Insufficient memory. |

# <a name="GetCount"></a>GetCount

Gets the count of extra properties in this character formatting collection.

```
FUNCTION CTextFont2.GetCount () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetCount(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The count of extra properties in this collection.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetAutoLigatures"></a>GetAutoLigatures

Gets whether support for automatic ligatures is active.

```
FUNCTION CTextFont2.GetAutoLigatures () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAutoLigatures(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Automatic ligature support is active. |
| **tomFalse** | Automatic ligature support is not active. |
| **tomUndefined** | The AutoLigatures property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetAutoLigatures"></a>SetAutoLigatures

Sets whether support for automatic ligatures is active.

```
FUNCTION CTextFont2.SetAutoLigatures (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAutoLigatures(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an ^^HRESULT** error code.

# <a name="GetAutospaceAlpha"></a>GetAutospaceAlpha

Gets the East Asian "autospace alphabetics" state.

```
FUNCTION CTextFont2.GetAutospaceAlpha () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAutospaceAlpha(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | East Asian autospace alphabetics. |
| **tomFalse** | Do not use East Asian autospace alphabetics. |
| **tomUndefined** | The AutospaceAlpha property is undefined |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetAutospaceAlpha"></a>SetAutospaceAlpha

Sets the East Asian "autospace alpha" state.

```
FUNCTION CTextFont2.SetAutospaceAlpha (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAutospaceAlpha(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an ^^HRESULT** error code.

# <a name="GetAutospaceNumeric"></a>GetAutospaceNumeric

Gets the East Asian "autospace numeric" state.

```
FUNCTION CTextFont2.GetAutospaceNumeric () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetAutospaceNumeric(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Use East Asian autospace numerics. |
| **tomFalse** | Do not use East Asian autospace numerics. |
| **tomUndefined** | The AutospaceNumeric property is undefined |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetAutospaceNumeric"></a>SetAutospaceNumeric

Sets the East Asian "autospace numeric" state.

```
FUNCTION CTextFont2.SetAutospaceNumeric (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAutospaceNumeric(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an ^^HRESULT** error code.

# <a name="GetAutospaceParens"></a>GetAutospaceParens

Gets the East Asian "autospace parentheses" state.

```
FUNCTION CTextFont2.GetAutospaceParens () AS LONG
   DIM Value AS LONG
   IF m_pTextFont2 = NULL THEN m_Result = E_POINTER: RETURN Value
   this.SetResult(m_pTextFont2->lpvtbl->GetAutospaceParens(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Use East Asian autospace parentheses. |
| **tomFalse** | Do not use East Asian autospace parentheses. |
| **tomUndefined** | The AutospaceParens property is undefined |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetAutospaceParens"></a>SetAutospaceParens

Sets the East Asian "autospace numeric" state.

```
FUNCTION CTextFont2.SetAutospaceParens (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetAutospaceParens(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an ^^HRESULT** error code.

# <a name="GetCharRep"></a>GetCharRep

Gets the character repertoire (writing system).

```
FUNCTION CTextFont2.GetCharRep () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetCharRep(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The character repertoire. It can be one of the following values.

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

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetCharRep"></a>SetCharRep

Sets the character repertoire (writing system).

```
FUNCTION CTextFont2.SetCharRep (BYVAL Value AS LONG) AS HRESULT
   IF m_pTextFont2 = NULL THEN m_Result = E_POINTER: RETURN m_Result
   this.SetResult(m_pTextFont2->lpvtbl->SetCharRep(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an ^^HRESULT** error code.

# <a name="GetCompressionMode"></a>GetCompressionMode

Gets the East Asian compression mode.

```
FUNCTION CTextFont2.GetCompressionMode () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetCompressionMode(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The compression mode, which can be one of these values:

| Value | Meaning |
| ----- | ------- |
| **tomCompressNone** (default) | No compression. |
| **tomCompressPunctuation** | Compress punctuation. |
| **tomCompressPunctuationAndKana** | Compress punctuation and kana. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetCompressionMode"></a>SetCompressionMode

Sets the East Asian compression mode.

```
FUNCTION CTextFont2.SetCompressionMode (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetCompressionMode(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The compression mode, which can be one of these values: |

| Value | Meaning |
| ----- | ------- |
| **tomCompressNone** (default) | No compression. |
| **tomCompressPunctuation** | Compress punctuation. |
| **tomCompressPunctuationAndKana** | Compress punctuation and kana. |

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an ^^HRESULT** error code.

# <a name="GetCookie"></a>GetCookie

Gets the client cookie.

```
FUNCTION CTextFont2.GetCookie () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetCookie(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The client cookie.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

This value is purely for the use of the client and has no meaning to the Text Object Model (TOM) engine. There are exceptions where different values correspond to different character format runs.

# <a name="SetCookie"></a>SetCookie

Sets the client cookie.

```
FUNCTION CTextFont2.SetCookie (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetCookie(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The client cookie. |

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

This value is purely for the use of the client. It has no meaning to the Text Object Model (TOM) engine except that different values correspond to different character format runs.

# <a name="GetDoubleStrike"></a>GetDoubleStrike

Gets whether characters are displayed with double horizontal lines through the center.

```
FUNCTION CTextFont2.GetDoubleStrike () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetDoubleStrike(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Characters are displayed with double horizontal lines. |
| **tomFalse** | Characters are not displayed with double horizontal lines. |
| **tomUndefined** | The DoubleStrike property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetDoubleStrike"></a>SetDoubleStrike

Sets whether characters are displayed with double horizontal lines through the center.

```
FUNCTION CTextFont2.SetDoubleStrike (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetDoubleStrike(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetLinkType"></a>GetLinkType

Gets the link type.

```
FUNCTION CTextFont2.GetLinkType () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetLinkType(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
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

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetMathZone"></a>GetMathZone

Gets whether a math zone is active.

```
FUNCTION CTextFont2.GetMathZone () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetMathZone(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | A math zone is active. |
| **tomFalse** | A math zone is not active. |
| **tomUndefined** | The MathZone property is undefined. |
	
#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetMathZone"></a>SetMathZone

Sets whether a math zone is active.

```
FUNCTION CTextFont2.SetMathZone (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetMathZone(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetModWidthPairs"></a>GetModWidthPairs

Gets whether "decrease widths on pairs" is active.

```
FUNCTION CTextFont2.GetModWidthPairs () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetModWidthPairs(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Decrease widths on pairs is active. |
| **tomFalse** | Decrease widths on pairs is not active. |
| **tomUndefined** | The ModWidthPairs property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetModWidthPairs"></a>SetModWidthPairs

Sets whether "decrease widths on pairs" is active.

```
FUNCTION CTextFont2.SetModWidthPairs (BYVAL Value AS LONG) AS HRESULT
   IF m_pTextFont2 = NULL THEN m_Result = E_POINTER: RETURN m_Result
   this.SetResult(m_pTextFont2->lpvtbl->SetModWidthPairs(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetModWidthSpace"></a>GetModWidthSpace

Gets whether "increase width of whitespace" is active.

```
FUNCTION CTextFont2.GetModWidthSpace () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetModWidthSpace(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Increase width of whitespace is active. |
| **tomFalse** | Increase width of whitespace is not active. |
| **tomUndefined** | The ModWidthSpace property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetModWidthSpace"></a>SetModWidthSpace

Sets whether "increase width of whitespace" is active.

```
FUNCTION CTextFont2.SetModWidthSpace (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetModWidthSpace(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetOldNumbers"></a>GetOldNumbers

Gets whether old-style numbers are active.

```
FUNCTION CTextFont2.GetOldNumbers () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetOldNumbers(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTIONN
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Old-style numbers are active. |
| **tomFalse** | Old-style numbers are not active. |
| **tomUndefined** | The OldNumbers property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetOldNumbers"></a>SetOldNumbers

Sets whether old-style numbers are active.

```
FUNCTION CTextFont2.SetOldNumbers (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetOldNumbers(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetOverlapping"></a>GetOverlapping

Gets whether overlapping text is active.

```
FUNCTION CTextFont2.GetOverlapping () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetOverlapping(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

A **tomBool** value that can be one of the following.

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Overlapping text is active. |
| **tomFalse** | Overlapping text is not active. |
| **tomUndefined** | The Overlapping property is undefined. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetOverlapping"></a>SetOverlapping

Sets whether overlapping text is active.

```
FUNCTION CTextFont2.SetOverlapping (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetOverlapping(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
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

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetPositionSubSuper"></a>GetPositionSubSuper

Gets the subscript or superscript position relative to the baseline.

```
FUNCTION CTextFont2.GetPositionSubSuper () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetPositionSubSuper(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The subscript or superscript position relative to the baseline.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetPositionSubSuper"></a>SetPositionSubSuper

Sets the position of a subscript or superscript relative to the baseline, as a percentage of the font height.

```
FUNCTION CTextFont2.SetPositionSubSuper (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetPositionSubSuper(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new subscript or superscript position. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetScaling"></a>GetScaling

Gets the font horizontal scaling percentage.

```
FUNCTION CTextFont2.GetScaling () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetScaling(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The scaling percentage.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The font horizontal scaling percentage can range from 200, which doubles the widths of characters, to 0, where no scaling is performed. When the percentage is increased the height does not change.

# <a name="SetScaling"></a>SetScaling

Sets the font horizontal scaling percentage.

```
FUNCTION CTextFont2.SetScaling (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetScaling(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The scaling percentage. Values from 0 through 255 are valid. For example, a value of 200 doubles the widths of characters while retaining the same height. A value of 0 has the same effect as a value of 100; that is, it turns scaling off. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetSpaceExtension"></a>GetSpaceExtension

Gets the East Asian space extension value.

```
FUNCTION CTextFont2.GetSpaceExtension () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetSpaceExtension(m_pTextFont2, @Value))
   FUNCTION = Value
END FUNCTION
```
#### Return value

The space extension, in floating-point points.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetSpaceExtension"></a>SetSpaceExtension

Sets the East Asian space extension value.

```
FUNCTION CTextFont2.SetSpaceExtension (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetSpaceExtension(m_pTextFont2, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new space extension, in floating-points. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetUnderlinePositionMode"></a>GetUnderlinePositionMode

Gets the underline position mode.

```
FUNCTION CTextFont2.GetUnderlinePositionMode () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetUnderlinePositionMode(m_pTextFont2, @Value))
   FUNCTION = Value
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

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetUnderlinePositionMode"></a>SetUnderlinePositionMode

Sets the underline position mode.

```
FUNCTION CTextFont2.SetUnderlinePositionMode (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetUnderlinePositionMode(m_pTextFont2, Value))
   FUNCTION = m_Result
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
FUNCTION CTextFont2.GetEffects (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->GetEffects(m_pTextFont2, pValue, pMask))
   FUNCTION = m_Result
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

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetEffects2"></a>GetEffects2

Gets the additional character format effects.

```
FUNCTION CTextFont2.GetEffects2 (BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->GetEffects2(m_pTextFont2, pValue, pMask))
   FUNCTION = m_Result
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

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetProperty"></a>GetProperty

Gets the value of the specified property.

```
FUNCTION CTextFont2.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->GetProperty(m_pTextFont2, nType, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Parameters

| Parameter | Description |
| --------- | ----------- |
| *nType* | The property ID of the value to return. |

#### Return value

The property value.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetPropertyInfo"></a>GetPropertyInfo

Gets the property type and value of the specified extra property.

```
FUNCTION CTextFont2.GetPropertyInfo (BYVAL Index AS LONG, BYVAL pType AS LONG PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->GetPropertyInfo(m_pTextFont2, Index, pType, pValue))
   FUNCTION = m_Result
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

# <a name="IsEqual2"></a>IsEqual2

Determines whether this text font object has the same properties as the specified text font object.

```
FUNCTION CTextFont2.IsEqual2 (BYVAL pFont AS ITextFont2 PTR) AS LONG
   DIM B AS LONG
   this.SetResult(m_pTextFont2->lpvtbl->IsEqual2(m_pTextFont2, pFont, @B))
   FUNCTION = B
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pFont* | The text font object to compare against. |

#### Return value

A variable that is **tomTrue** if the font objects have the same properties or **tomFalse** if they do not.

#### Result code

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="SetEffects"></a>SetEffects

Sets the character format effects.

```
FUNCTION CTextFont2.SetEffects (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetEffects(m_pTextFont2, Value, Mask))
   FUNCTION = m_Result
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

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

Only effects with the corresponding mask flag set are modified.

# <a name="SetEffects2"></a>SetEffects2

Sets the additional character format effects.

```
FUNCTION CTextFont2.SetEffects2 (BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetEffects2(m_pTextFont2, Value, Mask))
   FUNCTION = m_Result
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

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

Only effects with the corresponding mask flag set are modified.

# <a name="SetProperty"></a>SetProperty

Sets the value of the specified property.

```
FUNCTION CTextFont2.SetProperty (BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextFont2->lpvtbl->SetProperty(m_pTextFont2, nType, Value))
   FUNCTION = m_Result
END FUNCTION
```

#### Parameters

| Parameter | Description |
| --------- | ----------- |
| *nType* | The ID of the property value to set. |
| *Value* | The new property value. |

#### Return value

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.
