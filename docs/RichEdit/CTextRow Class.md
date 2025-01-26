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
| [GetAlignment](#GetAlignment) | Gets the horizontal alignment of a row. |
| [SetAlignment](#SetAlignment) | Sets the horizontal alignment of a row. |
| [GetCellCount](#GetCellCount) | Gets the count of cells in a row. |
| [SetCellCount](#SetCellCount) | Sets the count of cells in a row. |
| [GetCellCountCache](#GetCellCountCache) | Gets the count of cells cached for this row. |
| [SetCellCountCache](#SetCellCountCache) | Sets the count of cells cached for this row. |
| [GetCellIndex](#GetCellIndex) | Gets the index of the active cell to get or set parameters for. |
| [SetCellIndex](#SetCellIndex) | Sets the index of the active cell to get or set parameters for. |
| [GetCellMargin](#GetCellMargin) | Gets the cell margin of this row. |
| [SetCellMargin](#SetCellMargin) | Sets the cell margin of this row. |
| [GetHeight](#GetHeight) | Gets the height of the row. |
| [SetHeight](#SetHeight) | Sets the height of the row. |
| [GetIndent](#GetIndent) | Gets the indent of this row. |
| [SetIndent](#SetIndent) | Sets the indent of this row. |
| [GetKeepTogether](#GetKeepTogether) | Gets whether this row is allowed to be broken across pages. |
| [SetKeepTogether](#SetKeepTogether) | Sets whether this row is allowed to be broken across pages. |
| [GetKeepWithNext](#GetKeepWithNext) | Gets whether this row should appear on the same page as the row that follows it. |
| [SetKeepWithNext](#SetKeepWithNext) | Sets whether this row should appear on the same page as the row that follows it. |
| [GetNestLevel](#GetNestLevel) | Gets the nest level of a table. |
| [GetRTL](#GetRTL) | Gets whether this row has right-to-left orientation. |
| [SetRTL](#SetRTL) | Sets whether this row has right-to-left orientation. |
| [GetCellAlignment](#GetCellAlignment) | Gets the vertical alignment of the active cell. |
| [SetCellAlignment](#SetCellAlignment) | Sets the vertical alignment of the active cell. |
| [GetCellColorBack](#GetCellColorBack) | Gets the background color of the active cell. |
| [SetCellColorBack](#SetCellColorBack) | Sets the background color of the active cell. |
| [GetCellColorFore](#GetCellColorFore) | Gets the foreground color of the active cell. |
| [SetCellColorFore](#SetCellColorFore) | Sets the foreground color of the active cell. |
| [GetCellMergeFlags](#GetCellMergeFlags) | Gets the merge flags of the active cell. |
| [SetCellMergeFlags](#SetCellMergeFlags) | Sets the merge flags of the active cell. |
| [GetCellShading](#GetCellShading) | Gets the shading of the active cell. |
| [SetCellShading](#SetCellShading) | Sets the shading of the active cell. |
| [GetCellVerticalText](#GetCellVerticalText) | Gets the vertical-text setting of the active cell. |
| [SetCellVerticalText](#SetCellVerticalText) | Sets the vertical-text setting of the active cell. |
| [GetCellWidth](#GetCellWidth) | Gets the width of the active cell. |
| [SetCellWidth](#SetCellWidth) | Sets the width of the active cell. |
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

# <a name="GetAlignment"></a>GetAlignment

Gets the horizontal alignment of a row.

```
FUNCTION CTextRow.GetAlignment () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetAlignment(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The horizontal alignment. It can be one of the following values.

| Constant | Value | Meaning |
| -------- | ----- | ------- |
| **tomAlignLeft** | 0 | Text aligns with the left margin. |
| **tomAlignCenter** | 1 | Text is centered between the margins. |
| **tomAlignRight** | 2 | Text aligns with the right margin. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetAlignment"></a>SetAlignment

Sets the horizontal alignment of a row.

```
FUNCTION CTextRow.SetAlignment (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetAlignment(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The new horizontal alignment. It can be one of the following values. |

| Constant | Value | Meaning |
| -------- | ----- | ------- |
| **tomAlignLeft** | 0 | Text aligns with the left margin. |
| **tomAlignCenter** | 1 | Text is centered between the margins. |
| **tomAlignRight** | 2 | Text aligns with the right margin. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetCellCount"></a>GetCellCount

Gets the count of cells in this row.

```
FUNCTION CTextRow.GetCellCount () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellCount(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The cell count for this row.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetCellCount"></a>SetCellCount

Sets the count of cells in a row.

```
FUNCTION CTextRow.SetCellCount (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellCount(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The row cell count. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetCellCountCache"></a>GetCellCountCache

Gets the count of cells cached for this row.

```
FUNCTION CTextRow.GetCellCountCache () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellCountCache(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The cached cell count.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetCellCountCache"></a>SetCellCountCache

Sets the count of cells cached for a row.

```
FUNCTION CTextRow.SetCellCountCache (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellCountCache(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The row cell count. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

If all cells are identical, properties need to be cached only for the cell with index 0. If the cached count is less than the cell count, the cell parameters for index CellCountCache – 1 are used for cells with larger indices.


# <a name="GetCellIndex"></a>GetCellIndex

Gets the index of the active cell to get or set parameters for.

```
FUNCTION CTextRow.GetCellIndex () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellIndex(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The cell index.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.



# <a name="SetCellIndex"></a>SetCellIndex

Sets the index of the active cell.

```
FUNCTION CTextRow.SetCellIndex (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellIndex(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The cell index. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

You can get or set parameters for an active cell.

If the cell index is greater than the cell count, and the cell index is less that 63 (the maximum cell count), then the cell count is increased to cell index + 1.


# <a name="GetCellMargin"></a>GetCellMargin

ets the cell margin of this row.

```
FUNCTION CTextRow.GetCellMargin () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellMargin(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The cell margin.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetCellMargin"></a>SetCellMargin

Sets the cell margin of a row.

```
FUNCTION CTextRow.SetCellMargin (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellMargin(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The cell margin. The cell margin is used for all cells in the row and is typically about 108 twips or 0.075 inches. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetHeight"></a>GetHeight

Gets the height of the row.

```
FUNCTION CTextRow.GetHeight () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetHeight(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The row height. A value of 0 indicates autoheight.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.



# <a name="SetHeight"></a>SetHeight

Sets the cell margin of a row.

```
FUNCTION CTextRow.SetHeight (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetHeight(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The row height. A value of 0 indicates autoheight. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetIndent"></a>GetIndent

Gets the indent of this row.

```
FUNCTION CTextRow.GetIndent () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetIndent(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The indent of the row.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetIndent"></a>SetIndent

Sets the indent of a row.

```
FUNCTION CTextRow.SetIndent (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetIndent(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The row indent. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetKeepTogether"></a>GetKeepTogether

Gets whether this row is allowed to be broken across pages.

```
FUNCTION CTextRow.GetKeepTogether () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetKeepTogether(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

A **tomBool** value that indicates whether this row can be broken across pages.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetKeepTogether"></a>SetKeepTogether

Sets whether this row is allowed to be broken across pages.

```
FUNCTION CTextRow.SetKeepTogether (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetKeepTogether(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that indicates whether this row can be broken across pages. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetKeepWithNext"></a>GetKeepWithNext

Gets whether this row is allowed to be broken across pages.

```
FUNCTION CTextRow.GetKeepWithNext () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetKeepWithNext(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

A **tomBool** value that indicates whether this row can be broken across pages.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

# <a name="GetNestLevel"></a>GetNestLevel

Gets the nest level of a table.

```
FUNCTION CTextRow.GetNestLevel () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetNestLevel(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The nest level. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

####Remarks

The nest level of the table is identified by the associated **ITextRange2** object. If there is only a single table, the nest level is 1. If there is no table, the nest level is 0.


# <a name="GetRTL"></a>GetRTL

Gets whether this row has right-to-left orientation.

```
FUNCTION CTextRow.GetRTL () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetRTL(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

A **tomBool** value that indicates whether this row has right-to-left orientation.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetRTL"></a>SetRTL

Sets whether this row has right-to-left orientation.

```
FUNCTION CTextRow.SetRTL (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetRTL(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | A **tomBool** value that can be one of the following. |

| Value | Meaning |
| ----- | ------- |
| **tomTrue** | Right-to-left orientation. |
| **tomFalse** | Left-to-right orientation. |
| **tomToggle** | Toggles the orientation. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetCellAlignment"></a>GetCellAlignment

Gets the vertical alignment of the active cell.

```
FUNCTION CTextRow.GetCellAlignment () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellAlignment(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The vertical alignment.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetCellAlignment"></a>SetCellAlignment

Sets the vertical alignment of the active cell.

```
FUNCTION CTextRow.SetCellAlignment (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellAlignment(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The vertical alignment. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetCellColorBack"></a>GetCellColorBack

Gets the background color of the active cell.

```
FUNCTION CTextRow.GetCellColorBack () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellColorBack(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The background color.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetCellColorBack"></a>SetCellColorBack

Sets the background color of the active cell.

```
FUNCTION CTextRow.SetCellColorBack (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellColorBack(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The background color. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks
See **GetCellShading** to see how the background color is used together with the foreground color.


# <a name="GetCellColorFore"></a>GetCellColorFore

Gets the foreground color of the active cell.

```
FUNCTION CTextRow.GetCellColorFore () AS LONG
   DIM Value AS LONG
   IF m_pTextRow = NULL THEN m_Result = E_POINTER : RETURN Value
   this.SetResult(m_pTextRow->lpvtbl->GetCellColorFore(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

Gets the foreground color of the active cell.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetCellColorFore"></a>SetCellColorFore

Sets the foreground color of the active cell.

```
FUNCTION CTextRow.SetCellColorFore (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellColorFore(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The foreground color. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

See **GetCellShading** to see how the foreground color is used together with the background color.


# <a name="GetCellMergeFlags"></a>GetCellMergeFlags

Gets the merge flags of the active cell.

```
FUNCTION CTextRow.GetCellMergeFlags () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellMergeFlags(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The merge flag of the active cell. The flag can be one of the following:

| Flag | Meaning |
| --------- | ----------- |
| **tomHContCell** | Any cell except the start in a horizontally merged cell set. |
| **tomHStartCell** | The top cell in vertically merged cell set. |
| **tomVLowCell** | Any cell except the top cell in a vertically merged cell set. |
| **tomVTopCell** | The top cell in vertically merged cell set. |

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetCellMergeFlags"></a>SetCellMergeFlags

Sets the merge flags of the active cell.

```
FUNCTION CTextRow.SetCellMergeFlags (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellMergeFlags(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The merge flag. It can be one of these values. |

| Flag | Meaning |
| --------- | ----------- |
| **tomHContCell** | Any cell except the start in a horizontally merged cell set. |
| **tomHStartCell** | The top cell in vertically merged cell set. |
| **tomVLowCell** | Any cell except the top cell in a vertically merged cell set. |
| **tomVTopCell** | The top cell in vertically merged cell set. |

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetCellShading"></a>GetCellShading

Gets the shading of the active cell.

```
FUNCTION CTextRow.GetCellShading () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellShading(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The shading of the active cell. See Remarks.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The shading is given in hundredths of a percent, so full shading is given by the value 10000. The shading percentage determines the mix of the cell foreground and background colors to be used for the cell background. A shading of 0 uses the cell background color alone. A shading of 10000 (100%) uses the foreground color alone. Values in between mix the foreground and background colors, weighting the background with (10000 – CellShading)/1000 and the foreground with CellShading/1000. These ratios are applied to the red, green, and blue channels independently of one another.


# <a name="SetCellShading"></a>SetCellShading

Sets the shading of the active cell.

```
FUNCTION CTextRow.SetCellShading (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellShading(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The shading of the active cell. |

#### Remarks

The shading is given in hundredths of a percent, so full shading is given by the value 10000. The shading percentage determines the mix of the cell foreground and background colors to be used for the cell background. A shading of 0 uses the cell background color alone. A shading of 10000 (100%) uses the foreground color alone. Values in between mix the foreground and background colors, weighting the background with (10000 – CellShading)/1000 and the foreground with CellShading/1000. These ratios are applied to the red, green, and blue channels independently of one another.

#### Return value

If the method succeeds, it returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="GetCellVerticalText"></a>GetCellVerticalText

Gets the vertical-text setting of the active cell.

This property is not currently implemented.

```
FUNCTION CTextRow.GetCellVerticalText () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellVerticalText(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The vertical-text setting.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.


# <a name="SetCellVerticalText"></a>SetCellVerticalText

Sets the vertical-text setting of the active cell.

This property is not currently implemented.

```
FUNCTION CTextRow.SetCellVerticalText (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellVerticalText(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The vertical-text setting. |


# <a name="GetCellWidth"></a>GetCellWidth

Gets the width of the active cell.

```
FUNCTION CTextRow.GetCellWidth () AS LONG
   DIM Value AS LONG
   this.SetResult(m_pTextRow->lpvtbl->GetCellWidth(m_pTextRow, @Value))
   FUNCTION = Value
END FUNCTION
```

#### Return value

The width of the active cell, in twips.

#### Result code

If the method succeeds, **GetLastResult** returns **NOERROR**. Otherwise, it returns an **HRESULT** error code.

#### Remarks

The width of the cell, and/or the entire row, must be less than 22 inches (1440 x 22).


# <a name="SetCellWidth"></a>SetCellWidth

Sets the active cell width in twips.

```
FUNCTION CTextRow.SetCellWidth (BYVAL Value AS LONG) AS HRESULT
   this.SetResult(m_pTextRow->lpvtbl->SetCellWidth(m_pTextRow, Value))
   FUNCTION = m_Result
END FUNCTION
```

| Parameter | Description |
| --------- | ----------- |
| *Value* | The width of the active cell. |

#### Remarks

The total width of the entire row must be less than 22 inches, or 1440×22.
