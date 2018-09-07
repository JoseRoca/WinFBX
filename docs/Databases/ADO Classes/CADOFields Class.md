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
| [Resync](#Resync) | Resynchronizes the contents of the **Fields** collection. |
| [Refresh](#Refresh) | Refreshes the contents of the **Fields** collection. |
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
