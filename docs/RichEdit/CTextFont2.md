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
| [GetSize](#GetSize) |  |
| [SetSize](#SetSize) |  |
| [GetSmallCaps](#GetSmallCaps) |  |
| [SetSmallCaps](#SetSmallCaps) |  |
| [GetSpacing](#GetSpacing) |  |
| [SetSpacing](#SetSpacing) |  |
| [GetStrikeThrough](#GetStrikeThrough) |  |
| [SetStrikeThrough](#SetStrikeThrough) |  |
| [GetSubscript](#GetSubscript) |  |
| [SetSubscript](#SetSubscript) |  |
| [GetSuperscript](#GetSuperscript) |  |
| [SetSuperscript](#SetSuperscript) |  |
| [GetUnderline](#GetUnderline) |  |
| [SetUnderline](#SetUnderline) |  |
| [GetWeight](#GetWeight) |  |
| [SetWeight](#SetWeight) |  |

### ITextFont2 Interface

In the Text Object Model (TOM), applications access text-range attributes by using a pair of dual interfaces, **ITextFont** and **ITextPara**.

The **ITextFont2** interface extends **ITextFont**, providing the programming equivalent of the Microsoft Word format-font dialog.

| Name       | Description |
| ---------- | ----------- |
| [GetCount](#GetCount) |  |
| [GetAutoLigatures](#GetAutoLigatures) |  |
| [SetAutoLigatures](#SetAutoLigatures) |  |
| [GetAutospaceAlpha](#GetAutospaceAlpha) |  |
| [SetAutospaceAlpha](#SetAutospaceAlpha) |  |
| [GetAutospaceNumeric](#GetAutospaceNumeric) |  |
| [SetAutospaceNumeric](#SetAutospaceNumeric) |  |
| [GetAutospaceParens](#GetAutospaceParens) |  |
| [SetAutospaceParens](#SetAutospaceParens) |  |
| [GetCharRep](#GetCharRep) |  |
| [SetCharRep](#SetCharRep) |  |
| [GetCompressionMode](#GetCompressionMode) |  |
| [SetCompressionMode](#SetCompressionMode) |  |
| [GetCookie](#GetCookie) |  |
| [SetCookie](#SetCookie) |  |
| [GetDoubleStrike](#GetDoubleStrike) |  |
| [SetDoubleStrike](#SetDoubleStrike) |  |
| [GetDuplicate2](#GetDuplicate) | Gets a duplicate of this character format object. |
| [SetDuplicate2](#SetDuplicate) | Sets the properties of this object by copying the properties of another text font object. |
| [GetLinkType](#GetLinkType) |  |
| [GetMathZone](#GetMathZone) |  |
| [SetMathZone](#SetMathZone) |  |
| [GetModWidthPairs](#GetModWidthPairs) |  |
| [SetModWidthPairs](#SetModWidthPairs) |  |
| [GetModWidthSpace](#GetModWidthSpace) |  |
| [SetModWidthSpace](#SetModWidthSpace) |  |
| [GetOldNumbers](#GetOldNumbers) |  |
| [SetOldNumbers](#SetOldNumbers) |  |
| [GetOverlapping](#GetOverlapping) |  |
| [SetOverlapping](#SetOverlapping) |  |
| [GetPositionSubSuper](#GetPositionSubSuper) |  |
| [SetPositionSubSuper](#SetPositionSubSuper) |  |
| [GetScaling](#GetScaling) |  |
| [SetScaling](#SetScaling) |  |
| [GetSpaceExtension](#GetSpaceExtension) |  |
| [SetSpaceExtension](#SetSpaceExtension) |  |
| [GetUnderlinePositionMode](#GetUnderlinePositionMode) |  |
| [SetUnderlinePositionMode](#SetUnderlinePositionMode) |  |
| [GetEffects](#GetEffects) |  |
| [GetEffects2](#GetEffects2) |  |
| [GetProperty](#GetProperty) |  |
| [GetPropertyInfo](#GetPropertyInfo) |  |
| [IsEqual2](#IsEqual2) |  |
| [SetEffects](#SetEffects) |  |
| [SetEffects2](#SetEffects2) |  |
| [SetProperty](#SetProperty) |  |

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

