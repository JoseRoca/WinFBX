# CTextPara2 Class

Class that wraps all the methods of the **ITextPara** and **ITextPara2** interfaces.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTORS](#CONSTRUCTORS) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [LET](#LET) | Assignment operator. |
| [CAST](#CAST) | Cast operator. |
| [TextDocumentPtr](#TextDocumentPtr) | Returns a pointer to the underlying **ITextPara2** interface. |
| [Attach](#Attach) | Attaches an **ITextPara2** interface pointer to the class. |
| [Detach](#Detach) | Detaches the underlying **ITextPara2** interface pointer from the class. |

### ITextPara Interface

Text Object Model (TOM) rich text-range attributes are accessed through a pair of dual interfaces, **ITextFont** and **ITextPara**.

The **ITextPara** interface inherits from the **IDispatch** interface. **ITextPara** also has these types of members:

| Name       | Description |
| ---------- | ----------- |
| [GetDuplicate](#GetDuplicate) |  |
| [SetDuplicate](#SetDuplicate) |  |
| [CanChange](#CanChange) |  |
| [IsEqual](#IsEqual) |  |
| [Reset](#Reset) |  |
| [GetStyle](#GetStyle) |  |
| [SetStyle](#SetStyle) |  |
| [GetAlignment](#GetAlignment) |  |
| [SetAlignment](#SetAlignment) |  |
| [GetHyphenation](#GetHyphenation) |  |
| [SetHyphenation](#SetHyphenation) |  |
| [GetFirstLineIndent](#GetFirstLineIndent) |  |
| [GetKeepTogether](#GetKeepTogether) |  |
| [SetKeepTogether](#SetKeepTogether) |  |
| [SetKeepWithNext](#SetKeepWithNext) |  |
| [GetLeftIndent](#GetLeftIndent) |  |
| [GetLineSpacing](#GetLineSpacing) |  |
| [GetLineSpacingRule](#GetLineSpacingRule) |  |
| [GetListAlignment](#GetListAlignment) |  |
| [SetListAlignment](#SetListAlignment) |  |
| [GetListLevelIndex](#GetListLevelIndex) |  |
| [SetListLevelIndex](#SetListLevelIndex) |  |
| [GetListStart](#GetListStart) |  |
| [SetListStart](#SetListStart) |  |
| [GetListTab](#GetListTab) |  |
| [SetListTab](#SetListTab) |  |
| [GetListType](#GetListType) |  |
| [SetListType](#SetListType) |  |
| [GetNoLineNumber](#GetNoLineNumber) |  |
| [SetNoLineNumber](#SetNoLineNumber) |  |
| [GetPageBreakBefore](#GetPageBreakBefore) |  |
| [SetPageBreakBefore](#SetPageBreakBefore) |  |
| [GetRightIndent](#GetRightIndent) |  |
| [SetRightIndent](#SetRightIndent) |  |
| [SetIndents](#SetIndents) |  |
| [SetLineSpacing](#SetLineSpacing) |  |
| [GetSpaceAfter](#GetSpaceAfter) |  |
| [SetSpaceAfter](#SetSpaceAfter) |  |
| [GetSpaceBefore](#GetSpaceBefore) |  |
| [SetSpaceBefore](#SetSpaceBefore) |  |
| [GetWidowControl](#GetWidowControl) |  |
| [SetWidowControl](#SetWidowControl) |  |
| [GetTabCount](#GetTabCount) |  |
| [AddTab](#AddTab) |  |
| [ClearAllTabs](#ClearAllTabs) |  |
| [DeleteTab](#DeleteTab) |  |
| [GetTab](#GetTab) |  |

### ITextPara2 Interface

The **ITextPara2** interface extends ITextPara, providing the equivalent of the Microsoft Word format-paragraph dialog.

The **ITextPara2** interface has these methods.

| Name       | Description |
| ---------- | ----------- |
| [GetBorders](#GetBorders) |  |
| [GetDuplicate2](#GetDuplicate2) |  |
| [SetDuplicate2](#SetDuplicate2) |  |
| [GetFontAlignment](#GetFontAlignment) |  |
| [SetFontAlignment](#SetFontAlignment) |  |
| [GetHangingPunctuation](#GetHangingPunctuation) |  |
| [SetHangingPunctuation](#SetHangingPunctuation) |  |
| [GetSnapToGrid](#GetSnapToGrid) |  |
| [SetSnapToGrid](#SetSnapToGrid) |  |
| [GetTrimPunctuationAtStart](#GetTrimPunctuationAtStart) |  |
| [SetTrimPunctuationAtStart](#SetTrimPunctuationAtStart) |  |
| [GetEffects](#GetEffects) |  |
| [GetProperty](#GetProperty) |  |
| [IsEqual2](#IsEqual2) |  |
| [SetEffects](#SetEffects) |  |
| [SetProperty](#SetProperty) |  |

### Methods inherited from CTOMBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#GetLastResult) | Returns the last result code |
| [SetResult](#SetResult) | Sets the last result code. |
| [GetErrorInfo](#GetErrorInfo) | Returns a description of the last result code. |

# <a name="CONSTRUCTORS"></a>CONSTRUCTORS

Called when a **CTextPara2** class variable is created.

```
DECLARE CONSTRUCTOR
DECLARE CONSTRUCTOR (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

## CONSTRUCTOR (Empty)

Can be used, for example, when we have an **ITextPara2** interface pointer returned by a function and we want to attach it to a new instance of the **CTextPara2** class.

```
DIM DIM pCTextPara2 AS CTextPara2
pCTextPara2.Attach(pTextPara2)
```
## CONSTRUCTOR (ITextPara2 PTR)

```
CONSTRUCTOR CTextPara2 (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   IF pTextPara2 THEN
      IF fAddRef THEN pTextPara2->lpvtbl->AddRef(pTextPara2)
   End IF
   m_pTextPara2 = pTextPara2
END CONSTRUCTOR
```

| Parameter | Description |
| --------- | ----------- |
| *pTextPara2* | An **ITextPara2** interface pointer. |
| *fAddRef* | Optional. **TRUE** to increment the reference count of the passed **ITextPara2** interface pointer; otherwise, **FALSE**. Default is **FALSE**. |

#### Return value

A pointer to the new instance of the class.

# <a name="DESTRUCTOR"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

```
DESTRUCTOR CTextPara2
   ' // Release the ITextPara2 interface
   IF m_pTextPara2 THEN m_pTextPara2->lpvtbl->Release(m_pTextPara2)
END DESTRUCTOR
```

# <a name="LET"></a>LET

Assignment operator. The assigned pointer must be an "addrefed" one.

```
OPERATOR CTextPara2.LET (BYVAL pTextPara2 AS ITextPara2 PTR)
   m_Result = 0
   IF pTextPara2 = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextPara2 THEN m_pTextPara2->lpvtbl->Release(m_pTextPara2)
   ' // Attach the passed interface pointer to the class
   m_pTextPara2 = pTextPara2
END OPERATOR
```

# <a name="CAST"></a>CAST

Cast operator.

```
OPERATOR CTextPara2.CAST () AS ITextPara2 PTR
   m_Result = 0
   OPERATOR = m_pTextPara2
END OPERATOR
```

# <a name="TextParaPtr"></a>TextParaPtr

Returns a pointer to the underlying **ITextPara2** interface

```
FUNCTION CTextPara2.TextParaPtr () AS ITextPara2 PTR
   m_Result = 0
   RETURN m_pTextPara2
END FUNCTION
```

# <a name="Attach"></a>Attach

Attaches an **ITextPara2** interface pointer to the class.

```
FUNCTION CTextPara2.Attach (BYVAL pTextPara2 AS ITextPara2 PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextPara2 = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextPara2 THEN m_Result = m_pTextPara2->lpvtbl->Release(m_pTextPara2)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextPara2->lpvtbl->AddRef(pTextPara2)
   m_pTextPara2 = pTextPara2
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pTextRange2* | The **ITextPara2** interface pointer to attach. |
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
