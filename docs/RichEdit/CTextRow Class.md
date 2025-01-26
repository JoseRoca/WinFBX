# CTextRow Class

Class that wraps all the methods of the **ITextRow** interface.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTORS](#CONSTRUCTORS) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [LET](#LET) | Assignment operator. |
| [CAST](#CAST) | Cast operator. |
| [TextDocumentPtr](#TextDocumentPtr) | Returns a pointer to the underlying **ITextRow** interface. |
| [Attach](#Attach) | Attaches an **ITextRow** interface pointer to the class. |
| [Detach](#Detach) | Detaches the underlying **ITextRow** interface pointer from the class. |

### ITextRow Interface

The **ITextRow** interface provides methods to insert one or more identical table rows, and to retrieve and change table row properties. To insert nonidentical rows, call **Insert** for each different row configuration.

The **ITextRow** interface inherits from the **IDispatch** interface. **ITextRow** also has these types of members:

| Name       | Description |
| ---------- | ----------- |
| [GetAlignment](#GetAlignment) |  |
| [SetAlignment](#SetAlignment) |  |
| [GetCellCount](#GetCellCount) |  |
| [SetCellCount](#SetCellCount) |  |
| [GetCellCountCache](#GetCellCountCache) |  |
| [SetCellCountCache](#SetCellCountCache) |  |
| [GetCellIndex](#GetCellIndex) |  |
| [SetCellIndex](#SetCellIndex) |  |
| [GetCellMargin](#GetCellMargin) |  |
| [SetCellMargin](#SetCellMargin) |  |
| [GetHeight](#GetHeight) |  |
| [SetHeight](#SetHeight) |  |
| [GetIndent](#GetIndent) |  |
| [SetIndent](#SetIndent) |  |
| [GetKeepTogether](#GetKeepTogether) |  |
| [SetKeepTogether](#SetKeepTogether) |  |
| [GetKeepWithNext](#GetKeepWithNext) |  |
| [SetKeepWithNext](#SetKeepWithNext) |  |
| [GetNestLevel](#GetNestLevel) |  |
| [GetRTL](#GetRTL) |  |
| [SetRTL](#SetRTL) |  |
| [GetCellAlignment](#GetCellAlignment) |  |
| [SetCellAlignment](#SetCellAlignment) |  |
| [GetCellColorBack](#GetCellColorBack) |  |
| [SetCellColorBack](#SetCellColorBack) |  |
| [GetCellColorFore](#GetCellColorFore) |  |
| [SetCellColorFore](#SetCellColorFore) |  |
| [GetCellMergeFlags](#GetCellMergeFlags) |  |
| [SetCellMergeFlags](#SetCellMergeFlags) |  |
| [GetCellShading](#GetCellShading) |  |
| [SetCellShading](#SetCellShading) |  |
| [GetCellVerticalText](#GetCellVerticalText) |  |
| [SetCellVerticalText](#SetCellVerticalText) |  |
| [GetCellWidth](#GetCellWidth) |  |
| [SetCellWidth](#SetCellWidth) |  |
| [GetCellBorderColors](#GetCellBorderColors) |  |
| [GetCellBorderWidths](#GetCellBorderWidths) |  |
| [SetCellBorderColors](#SetCellBorderColors) |  |
| [SetCellBorderWidths](#SetCellBorderWidths) |  |
| [Apply](#Apply) |  |
| [GetProperty](#GetProperty) |  |
| [Insert](#Insert) |  |
| [IsEqual](#IsEqual) |  |
| [Reset](#Reset) |  |
| [SetProperty](#SetProperty) |  |

### Methods inherited from CTOMBase Class

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#GetLastResult) | Returns the last result code |
| [SetResult](#SetResult) | Sets the last result code. |
| [GetErrorInfo](#GetErrorInfo) | Returns a description of the last result code. |


# <a name="CONSTRUCTORS"></a>CONSTRUCTORS

Called when a **CTextRow** class variable is created.

```
DECLARE CONSTRUCTOR
DECLARE CONSTRUCTOR (BYVAL pTextRow AS ITextRow PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

## CONSTRUCTOR (Empty)

Can be used, for example, when we have an **ITextRow** interface pointer returned by a function and we want to attach it to a new instance of the **CTextRow** class.

```
DIM DIM pCTextRow AS CTextRow
pCTextRow.Attach(pTextRow)
```
## CONSTRUCTOR (ITextRange2 PTR)

```
CONSTRUCTOR CTextRow (BYVAL pTextRow AS ITextRow PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Store the pointer of ITextRow interface
   IF pTextRow THEN
      IF fAddRef THEN pTextRow->lpvtbl->AddRef(pTextRow)
   END IF
   m_pTextRow = pTextRow
END CONSTRUCTOR
```

| Parameter | Description |
| --------- | ----------- |
| *pTextRow* | An **ITextRow** interface pointer. |
| *fAddRef* | Optional. **TRUE** to increment the reference count of the passed **ITextRange2** interface pointer; otherwise, **FALSE**. Default is **FALSE**. |

#### Return value

A pointer to the new instance of the class.

# <a name="DESTRUCTOR"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

```
DESTRUCTOR CTextRow
   ' // Release the ITextRow interface
   IF m_pTextRow THEN m_pTextRRow->lpvtbl->Release(m_pTextRow)
END DESTRUCTOR
```

# <a name="LET"></a>LET

Assignment operator. The assigned pointer must be an "addrefed" one.

```
OPERATOR CTextRow.LET (BYVAL pTextRow AS ITextRow PTR)
   m_Result = 0
   IF pTextRow = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF pTextRow THEN pTextRow->lpvtbl->Release(m_pTextRow)
   ' // Attach the passed interface pointer to the class
   m_pTextRow = pTextRow
END OPERATOR
```

# <a name="CAST"></a>CAST

Cast operator.

```
OPERATOR CTextRow.CAST () AS ITextRow PTR
   m_Result = 0
   OPERATOR = m_pTextRow
END OPERATOR
```

# <a name="TextRangePtr"></a>TextRangePtr

Returns a pointer to the underlying **ITextRange2** interface

```
FUNCTION CTextRow.TextRangePtr () AS ITextRow PTR
   m_Result = 0
   RETURN m_pTextRow
END FUNCTION
```

# <a name="Attach"></a>Attach

Attaches an **ITextRow** interface pointer to the class.

```
FUNCTION CTextRow.Attach (BYVAL pTextRow AS ITextRow PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pTextRow = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextRow THEN m_Result = m_pTextRow->lpvtbl->Release(m_pTextRow)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextRow->lpvtbl->AddRef(pTextRow)
   m_pTextRow = pTextRow
   RETURN m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *pTextRow* | The **ITextRow** interface pointer to attach. |
| *fAddRef* | **TRUE** to increment the reference count of te object; otherwise, **FALSE**. Default is **FALSE**. |

# <a name="Detach"></a>Detach

Detaches the interface pointer from the class

```
FUNCTION CTextRow.Detach () AS ITextRow PTR
   m_Result = 0
   DIM pTextRow AS ITextRow PTR = m_pTextRow
   m_pTextRow = NULL
   RETURN pTextRow
END FUNCTION
```

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

