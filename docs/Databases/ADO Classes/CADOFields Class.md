# CADOFields Class

Contains all the **Field** objects of a **Recordset** or **Record** object.

**Include file**: CADOFields.inc (include CADODB.inc)

A **Recordset** object has a **Fields** collection made up of **Field** objects. Each **Field** object corresponds to a column in the **Recordset**. You can populate the **Fields** collection before opening the **Recordset** by calling the **Refresh** method on the collection.

See the **Field** object topic for a more detailed explanation of how to use **Field** objects.

The **Fields** collection has an **Append** method, which provisionally creates and adds a Field object to the collection, and an Update method, which finalizes any additions or deletions.

A **Record** object has two special fields that can be indexed with **FieldEnum** constants. One constant accesses a field containing the default stream for the **Record**, and the other accesses a field containing the absolute URL string for the **Record**.

Certain providers (for example, the Microsoft OLE DB Provider for Internet Publishing) may populate the **Fields** collection with a subset of available fields for the **Record** or **Recordset**. Other fields will not be added to the collection until they are first referenced by name or indexed by your code.

If you attempt to reference a nonexistent field by name, a new **Field** object will be appended to the **Fields** collection with a **Status** of **adFieldPendingInsert**. When you call Update, ADO will create a new field in your data source if allowed by your provider.

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Append](#Append) | Appends an object to a collection. |
| [Attach](#Attach1) | Attaches a reference to a **Fields** object to the class. |
| [CancelUpdate](#CancelUpdate) | Cancels any changes made to the **Fields** collection of a **Record** object before calling the **Update** method. |
| [Count](#Count) | Retrieves the number of objects in the **Fields** collection. |
| [Delete_](#Delete_) | Deletes an object from the **Fields** collection. |
| [Item](#Item) | Indicates a specific member of the **Fields** collection, by name or ordinal number. |
| [Refresh](#Refresh) | Refreshes the contents of the **Fields** collection. |
| [Resync](#Resync) | Resynchronizes the contents of the **Fields** collection. |
| [Update](#Update) | Saves any changes you make to the current **Fields** collection of a **Record** object. |

# CADOField Class

Represents a column of data with a common data type.

**Include file**: CADOField.inc (include CADODB.inc)

Each **Field** object corresponds to a column in the **Recordset**. You use the **Value** property of **Field** objects to set or return data for the current record. Depending on the functionality the provider exposes, some collections, methods, or properties of a **Field** object may not be available.

With the collections, methods, and properties of a **Field** object, you can do the following:

* Return the name of a field with the **Name** property.
* View or change the data in the field with the **Value** property. **Value** is the default property of the **Field** object.
* Return the basic characteristics of a field with the **Type_**, **Precision**, and **NumericScale** properties.
* Return the declared size of a field with the **DefinedSize** property.
* Return the actual size of the data in a given field with the **ActualSize** property.
* Determine what types of functionality are supported for a given field with the **Attributes** property and **Properties** collection.
* Manipulate the values of fields containing long binary or long character data with the **AppendChunk** and **GetChunk** methods.
* If the provider supports batch updates, resolve discrepancies in field values during batch updating with the **OriginalValue** and **UnderlyingValue** properties.

All of the metadata properties (**Name**, **Type_**, **DefinedSize**, **Precision**, and **NumericScale**) are available before opening the **Field** object's **Recordset**. Setting them at that time is useful for dynamically constructing forms.

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [ActualSize](#ActualSize) | Indicates the actual length of a field's value. |
| [AppendChunk](#AppendChunk) | Appends data to a large text or binary data **Field**. |
| [Attach](#Attach2) | Attaches a reference to a **Field** object to the class. |
| [Attributes](#Attributes) | Indicates one or more characteristics of a field. |
| [DataFormat](#DataFormat) | Links the current **Field** object to a data-bound control. |
| [DefinedSize](#DefinedSize) | Indicates the data capacity of a field. |
| [GetChunk](#GetChunk) | Returns all, or a portion, of the contents of a large text or binary data field. |
| [Name](#Name) | Returns the name of the field. |
| [NumericScale](#NumericScale) | Sets or returns a Byte value that indicates the number of decimal places to which numeric values will be resolved. |
| [OriginalValue](#OriginalValue) | Indicates the value of a field that existed in the record before any changes were made. |
| [Precision](#Precision) | Sets or returns a Byte value that indicates the maximum number of digits used to represent values. |
| [Properties](#Properties) | Returns a reference to the Properties collection. |
| [Status](#Status) | Indicates the status of a field. |
| [Type_](#Type_) | Sets or returns a **DataTypeEnum** value. |
| [UnderlyingValue](#UnderlyingValue) | Indicates a field's current value in the database. |
| [Value](#Value) | Sets or returns a Variant value that indicates the value of the object |

# <a name="Append"></a>Append

Appends an object to a collection. A new **Field** object may be created before it is appended to the collection.

```
FUNCTION Append (BYREF cbsName AS CBSTR, BYVAL nType AS DataTypeEnum, _
   BYVAL DefinedSize AS ADO_LONGPTR = 0, BYVAL Attrib AS FieldAttributeEnum = 0, _
   BYREF cvFieldValue AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsName* | An string value that contains the name of the new **Field** object, and must not be the same name as any other object in fields. |
| *nType* | A **DataTypeEnum** value, whose default value is **adEmpty**, that specifies the data type of the new field. The following data types are not supported by ADO, and should not be used when appending new fields to a Recordset: **adIDispatch**, **adIUnknown**, **adVariant**. |
| *DefinedSize* | Optional. An ADO_LONGPTR value that represents the defined size, in characters or bytes, of the new field. The default value for this parameter is derived from Type. Fields with a **DefinedSize** greater than 255 bytes, and treated as variable length columns. (The default **DefinedSize** is unspecified.) |
| *Attrib* | Optional. A **FieldAttributeEnum** value, whose default value is **adFldDefault**, that specifies attributes for the new field. If this value is not specified, the field will contain attributes derived from **Type_**. |
| *FieldValue* | Optional. A Vaiant that represents the value for the new field. If not specified, then the field is appended with a null value. |

#### Return value

S_OK (0) or an HESULT code.

#### DataTypeEnum

Specifies the data type of a **Field**, **Parameter**, or **Property**. The corresponding OLE DB type indicator is shown in parentheses in the description column of the following table.

| Constant   | Description |
| ---------- | ----------- |
| **AdArray** | A flag value, always combined with another data type constant, that indicates an array of that other data type. (Does not apply to ADOX.) |
| **adBigInt** | Indicates an eight-byte signed integer (DBTYPE_I8). |
| **adBinary** | Indicates a binary value (DBTYPE_BYTES). |
| **adBoolean** | Indicates a boolean value (DBTYPE_BOOL). |
| **adBSTR** | Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR). |
| **adChapter** | Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER). |
| **adChar** | Indicates a string value (DBTYPE_STR). |
| **adCurrency** | Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000. |
| **adDate** | Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day. |
| **adDBDate** | Indicates a date value (yyyymmdd) (DBTYPE_DBDATE). |
| **adDBTime** | Indicates a time value (hhmmss) (DBTYPE_DBTIME). |
| **adDBTimeStamp** | Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP). |
| **adDecimal** | Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL). |
| **adDouble** | Indicates a double-precision floating-point value (DBTYPE_R8). |
| **adEmpty** | Specifies no value (DBTYPE_EMPTY). |
| **adError** | Indicates a 32-bit error code (DBTYPE_ERROR). |
| **adFileTime** | Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME). |
| **adGUID** | Indicates a globally unique identifier (GUID) (DBTYPE_GUID). |
| **adIDispatch** | Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adInteger** | Indicates a four-byte signed integer (DBTYPE_I4). |
| **adIUnknown** | Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adLongVarBinary** | Indicates a long binary value. |
| **adLongVarChar** | Indicates a long string value. |
| **adLongVarWChar** | Indicates a long null-terminated Unicode string value. |
| **adNumeric** | Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC). |
| **adPropVariant** | Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT). |
| **adSingle** | Indicates a single-precision floating-point value (DBTYPE_R4). |
| **adSmallInt** | Indicates a two-byte signed integer (DBTYPE_I2). |
| **adTinyInt** | Indicates a one-byte signed integer (DBTYPE_I1). |
| **adUnsignedBigInt** | Indicates an eight-byte unsigned integer (DBTYPE_UI8). |
| **adUnsignedInt** | Indicates a four-byte unsigned integer (DBTYPE_UI4). |
| **adUnsignedSmallInt** | Indicates a two-byte unsigned integer (DBTYPE_UI2). |
| **adUnsignedTinyInt** | Indicates a one-byte unsigned integer (DBTYPE_UI1). |
| **adVarBinary** | Indicates a binary value. |
| **adVarChar** | Indicates a string value. |
| **adVariant** | Indicates an Automation Variant (DBTYPE_VARIANT). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adVarNumeric** | Indicates a numeric value. |
| **adVarWChar** | Indicates a null-terminated Unicode character string. |
| **adWChar** | Indicates a null-terminated Unicode character string (DBTYPE_WSTR). |

#### Remarks

The **FieldValue** parameter is only valid when adding a **Field** object to a **Record** object, not to a **Recordset** object. With a **Record** object, you may append fields and provide values at the same time. With a **Recordset** object, you must create fields while the **Recordset** is closed, then open the **Recordset** and assign values to the fields.

**Notes**: For new **Field** objects that have been appended to the **Fields** collection of a **Record** object, the Value property must be set before any other **Field** properties can be specified. First, a specific value for the **Value** property must have been assigned and **Update** on the **Fields** collection called. Then, other properties such as Type or **Attributes** can be accessed.

**Field** objects of the following data types (**DataTypeEnum**) cannot be appended to the **Fields** collection and will cause an error to occur: **adArray**, **adChapter**, **adEmpty**, **adPropVariant**, and **adUserDefined**. Also, the following data types are not supported by ADO: **adIDispatch**, **adIUnknown**, and **adIVariant**. For these types, no error will occur when appended, but usage can produce unpredictable results including memory leaks.

# <a name="Attach"></a>Attach

Attaches a reference to an ADO **Fields** object to the class.

```
SUB Attach (BYVAL pFields AS Afx_ADOFields PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFields* | A reference to an ADO **Fields** object. |
| *fAddRef* | TRUE = increase the reference count; FALSE = don't increase the reference count. |

# <a name="CancelUpdate"></a>CancelUpdate

Cancels any changes made to the **Fields** collection of a **Record** object before calling the **Update** method.

```
FUNCTION CancelUpdate () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

The **CancelUpdate** method cancels any pending insertions or deletions of Field objects, and cancels pending updates of existing fields and restores them to their original values. The **Status** property of all fields in the Fields collection is set to **adFieldOK**.

# <a name="Count"></a>Count

Retrieves the number of objects in the **Fields** collection.

```
PROPERTY Count () AS LONG
```

#### Return value

The number of objects in the **Fields** collection.

#### Remarks

Because numbering for members of a collection begins with zero, you should always code loops starting with the zero member and ending with the value of the **Count** property minus 1.

If the **Count** property is zero, there are no objects in the collection.

# <a name="CoDelete_unt"></a>Delete_

Deletes an object from the **Fields** collection.

```
FUNCTION Delete_ (BYREF cvIndex AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvIndex* | A Variant that designates the **Field** object to delete. This parameter can be the name of the **Field** object or the ordinal position of the **Field** object itself. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Calling the **Delete_** method on an open **Recordset** causes a run-time error.

# <a name="Item"></a>Item

Indicates a specific member of the **Fields** collection, by name or ordinal number.

```
PROPERTY Item (BYREF cvIndex AS CVAR) AS Afx_ADOField PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvIndex* | A Variant expression that evaluates either to the name or to the ordinal number of an object in a collection. |

#### Return value

An ADO **Fields** object reference.

#### Remarks

If **Item** cannot find an object in the collection corresponding to the *Index* argument, an error occurs.

# <a name="Refresh"></a>Refresh

Refreshes the contents of the **Fields** collection.

```
FUNCTION Refresh () AS HRESULT
```
#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Using the **Refresh** method on the **Fields** collection has no visible effect. To retrieve changes from the underlying database structure, you must use either the **Requery** method or, if the **Recordset** object does not support bookmarks, the **MoveFirst** method.

# <a name="Resync"></a>Resync

Resynchronizes the contents of the **Fields** collection.

```
FUNCTION Resync (BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *ResyncValues* | Optional. A **ResyncEnum** value that specifies whether underlying values are overwritten. The default value is **adResyncAllValues**. |

#### ResyncEnum

Specifies whether underlying values are overwritten by a call to **Resync**.

| Constant   | Description |
| ---------- | ----------- |
| **adResyncAllValues** | Default. Overwrites data, and pending updates are canceled. |
| **adResyncUnderlyingValues** | Does not overwrite data, and pending updates are not canceled. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Resync** method to resynchronize the values of the **Fields** collection of a **Record** object with the underlying data source. The **Count** property is not affected by this method.

If **ResyncValues** is set to **adResyncAllValues** (the default value), then the **UnderlyingValue**, **Value**, and **OriginalValue** properties of **Field** objects in the collection are synchronized. If **ResyncValues** is set to **adResyncUnderlyingValues**, only the **UnderlyingValue** property is synchronized.

The value of the **Status** property for each **Field** object at the time of the call also affects the behavior of **Resync**. For **Field** objects with **Status** values of **adFieldPendingUnknown** or **adFieldPendingInsert**, **Resync** has no effect. For **Status** values of **adFieldPendingChange** or **adFieldPendingDelete**, **Resync** synchronizes data values for fields that still exist at the data source.

**Resync** will not modify **Status** values of **Field** objects unless an error occurs when **Resync** is called. For example, if the field no longer exists, the provider will return an appropriate Status value for the **Field** object, such as **adFieldDoesNotExist**. Returned **Status** values may be logically combined within the value of the **Status** property.
