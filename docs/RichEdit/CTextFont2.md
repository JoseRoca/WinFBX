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
| [SetDuplicate](#SetDuplicate) |  |
| [CanChange](#CanChange) |  |
| [IsEqual](#IsEqual) |  |
| [Reset](#Reset) |  |
| [GetStyle](#GetStyle) |  |
| [SetStyle](#SetStyle) |  |
| [GetAllCaps](#GetAllCaps) |  |
| [SetAllCaps](#SetAllCaps) |  |
| [GetAnimation](#GetAnimation) |  |
| [SetAnimation](#SetAnimation) |  |
| [GetBackColor](#GetBackColor) |  |
| [SetBackColor](#SetBackColor) |  |
| [GetBold](#GetBold) |  |
| [SetBold](#SetBold) |  |
| [GetEmboss](#GetEmboss) |  |
| [SetEmboss](#SetEmboss) |  |
| [GetForeColor](#GetForeColor) |  |
| [SetForeColor](#SetForeColor) |  |
| [GetHidden](#GetHidden) |  |
| [SetHidden](#SetHidden) |  |
| [GetEngrave](#GetEngrave) |  |
| [SetEngrave](#SetEngrave) |  |
| [GetItalic](#GetItalic) |  |
| [SetItalic](#SetItalic) |  |
| [GetKerning](#GetKerning) |  |
| [SetKerning](#SetKerning) |  |
| [GetLanguageID](#GetLanguageID) |  |
| [SetLanguageID](#SetLanguageID) |  |
| [GetName](#GetName) |  |
| [SetName](#SetName) |  |
| [GetOutline](#GetOutline) |  |
| [SetOutline](#SetOutline) |  |
| [GetPosition](#GetPosition) |  |
| [SetPosition](#SetPosition) |  |
| [GetProtected](#GetProtected) |  |
| [SetProtected](#SetProtected) |  |
| [GetShadow](#GetShadow) |  |
| [SetShadow](#SetShadow) |  |
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
| [GetDuplicate2](#GetDuplicate2) |  |
| [SetDuplicate2](#SetDuplicate2) |  |
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

Called when a **CTextRange2** class variable is created.

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
## CONSTRUCTOR (ITextRange2 PTR)

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

# <a name="TextRangePtr"></a>TextRangePtr

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

