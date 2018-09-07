# CADORecordset Class

Represents the entire set of records from a base table or the results of an executed command. At any time, the **Recordset** object refers to only a single record within the set as the current record.

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [AbsolutePage](#AbsolutePage) | Sets or returns a Long value from 1 to the number of pages in the **Recordset** object (**PageCount**), or returns one of the **PositionEnum** values. |
| [AbsolutePosition](#AbsolutePosition) | A value from 1 to the number of records in the **Recordset** object (**RecordCount**). |
| [ActiveCommand](#ActiveCommand) | Indicates the **Command** object that created the associated **Recordset** object. |
| [ActiveConnection](#ActiveConnection) | Sets or returns a BSTR value that contains a definition for a connection if the connection is closed, or a Variant containing the current **Connection** object if the connection is open. Default is a null object reference. |
| [AddNew](#AddNew) | Creates a new record for an updatable **Recordset** object. |
| [Attach](#Attach) | Attaches a recordset to the class. |
| [BOF](#BOF) | Indicates that the current record position is before the first record in a **Recordset** object. |
| [Bookmark](#Bookmark) | Sets or returns a Variant expression that evaluates to a valid bookmark. |
| [CacheSize](#CacheSize) | Sets or returns a Long value that must be greater than 0. Default is 1. |
| [Cancel](#Cancel) | Cancels execution of a pending, asynchronous method call. |
| [CancelBatch](#CancelBatch) | Cancels a pending batch update. |
| [CancelUpdate](#CancelUpdate) | Cancels any changes made to the current or new row of a **Recordset** object before calling the **Update** method. |
| [Clone](#Clone) | Creates a duplicate **Recordset** object from an existing **Recordset** object. Optionally, specifies that the clone be read-only. |
| [Close](#Close) | Closes a **Recordset** object and any dependent objects. |
| [Collect](#Collect) | Sets or returns a Variant value that indicates the value of the object. |
| [CompareBookmarks](#CompareBookmarks) | Compares two bookmarks and returns an indication of their relative values. |
| [CursorLocation](#CursorLocation) | Indicates the location of the cursor service. |
| [CursorType](#CursorType) | Sets or returns a **CursorTypeEnum** value. The default value is **adOpenForwardOnly**. |
| [DataMember](#DataMember) | Indicates the name of the data member that will be retrieved from the object referenced by the **DataSource** property. |
| [DataSource](#DataSource) | Indicates an object that contains data to be represented as a **Recordset** object. |
| [Delete_](#Delete_) | Deletes the current record or a group of records. |
| [EditMode](#EditMode) | Indicates the editing status of the current record. |
| [EOF](#EOF) | Indicates that the current record position is after the last record in a **Recordset** object. |
| [Fields](#Fields) | Gets a reference to the **Fields** collection of a **Record** object. |
| [Filter](#Filter) | Indicates a filter for data in a **Recordset**. |
| [Find](#Find) | Searches a **Recordset** for the row that satisfies the specified criteria. |
| [GetErrorInfo](#GetErrorInfo) | Returns information about ADO errors. |
| [GetRows](#GetRows) | Retrieves multiple records of a **Recordset** object into an array. |
| [GetString](#GetString) | Returns the **Recordset** as a string. |
| [Index](#Index) | Sets or returns a string value, which is the name of the index. |
| [LockType](#LockType) | Sets or returns the lock type, a Long value that must be greater than 0. Default is 1. |
| [MarshalOptions](#MarshalOptions) | Indicates which records are to be marshaled back to the server. |
| [MaxRecords](#MaxRecords) | Sets or returns a Long value that indicates the maximum number of records to return. Default is zero (no limit). |
| [Move](#Move) | Moves the position of the current record in a **Recordset** object. |
| [MoveFirst](#MoveFirst) | Moves to the first record in a specified **Recordset** object and makes that record the current record. |
| [MoveLast](#MoveLast) | Moves to the last record in a specified **Recordset** object and makes that record the current record. |
| [MoveNext](#MoveNext) | Moves to the next record in a specified **Recordset** object and makes that record the current record. |
| [MovePrevious](#MovePrevious) | Moves to the previous record in a specified **Recordset** object and makes that record the current record. |
| [NextRecordset](#NextRecordset) | Moves to the previous record in a specified **Recordset** object and makes that record the current record. |
| [Open](#Open) | Opens a connection to a data source. |
| [PageCount](#PageCount) | Indicates how many pages of data the **Recordset** object contains. |
| [PageSize](#PageSize) | Sets or returns a Long value that indicates how many records are on a page. The default is 10. |
| [Properties](#Properties) | Returns a reference to the **Properties** collection. |
| [RecordCount](#RecordCount) | Indicates the number of records in a Recordset object. |
| [Requery](#Requery) | Updates the data in a **Recordset** object by re-executing the query on which the object is based. |
| [Resync](#Resync) | Refreshes the data in the current **Recordset** object from the underlying database. |
| [Save](#Save) | Saves the **Recordset** in a file or **Stream** object. |
| [Seek](#Seek) | Searches the index of a **Recordset** to quickly locate the row that matches the specified values, and changes the current row position to that row. |
| [Sort](#Sort) | Sets or returns a string value that indicates the field names in the **Recordset** on which to sort. |
| [Source](#Source) | Indicates the data source for a **Recordset** object. |
| [State](#State) | Indicates for whether the state of the **Recordset** object is open or closed. |
| [Status](#Status) | Indicates for whether the state of the **Recordset** object is open or closed. |
| [StayInSync](#StayInSync) | Indicates, in a hierarchical **Recordset** object, whether the reference to the underlying child records (that is, the chapter) changes when the parent row position changes. |
| [Supports](#Supports) | Determines whether a specified **Recordset** object supports a particular type of functionality. |
| [Update](#Update) | Saves any changes you make to the current row of a **Recordset** object. |
| [UpdateBatch](#UpdateBatch) | Writes all pending batch updates to disk. |


#### Remarks

You use **Recordset** objects to manipulate data from a provider. When you use ADO, you manipulate data almost entirely using **Recordset** objects. All **Recordset** objects consist of records (rows) and fields (columns). Depending on the functionality supported by the provider, some **Recordset** methods or properties may not be available.

There are four different cursor types defined in ADO:

* **Dynamic cursor** -- allows you to view additions, changes, and deletions by other users; allows all types of movement through the Recordset that doesn't rely on bookmarks; and allows bookmarks if the provider supports them.

* **Keyset cursor** -- behaves like a dynamic cursor, except that it prevents you from seeing records that other users add, and prevents access to records that other users delete. Data changes by other users will still be visible. It always supports bookmarks and therefore allows all types of movement through the **Recordset**.

* **Static cursor** -- provides a static copy of a set of records for you to use to find data or generate reports; always allows bookmarks and therefore allows all types of movement through the Recordset. Additions, changes, or deletions by other users will not be visible. This is the only type of cursor allowed when you open a client-side **Recordset** object.

* **Forward-only cursor** -- allows you to only scroll forward through the **Recordset++. Additions, changes, or deletions by other users will not be visible. This improves performance in situations where you need to make only a single pass through a **Recordset**.

Set the **CursorType** property prior to opening the **Recordset** to choose the cursor type, or pass a **CursorType** argument with the **Open** method. Some providers don't support all cursor types. Check the documentation for the provider. If you don't specify a cursor type, ADO opens a forward-only cursor by default.

If the **CursorLocation** property is set to **adUseClient** to open a **Recordset**, the **UnderlyingValue** property on **Field** objects is not available in the returned Recordset object. When used with some providers (such as the Microsoft ODBC Provider for OLE DB in conjunction with Microsoft SQL Server), you can create **Recordset** objects independently of a previously defined **Connection** object by passing a connection string with the **Open** method. ADO still creates a **Connection** object, but it doesn't assign that object to an object variable. However, if you are opening multiple Recordset objects over the same connection, you should explicitly create and open a **Connection** object; this assigns the **Connection** object to an object variable. If you do not use this object variable when opening your **Recordset** objects, ADO creates a new **Connection** object for each new **Recordset**, even if you pass the same connection string.

You can create as many **Recordset** objects as needed.

When you open a **Recordset**, the current record is positioned to the first record (if any) and the **BOF** and **EOF** properties are set to False. If there are no records, the **BOF** and **EOF** property settings are True.

You can use the **MoveFirst**, **MoveLast**, **MoveNext**, and **MovePrevious** methods; the **Move** method; and the **AbsolutePosition**, **AbsolutePage**, and **Filter** properties to reposition the current record, assuming the provider supports the relevant functionality. Forward-only **Recordset** objects support only the **MoveNext** method. When you use the **Move** methods to visit each record (or enumerate the **Recordset**), you can use the **BOF** and **EOF** properties to determine if you've moved beyond the beginning or end of the **Recordset**.

Before using any functionality of a **Recordset** object, you must call the **Supports** method on the object to verify that the functionality is supported or available. You must not use the functionality when the **Supports** method returns false. For example, you can use the **MovePrevious** method only if **Recordset.Supports**(*adMovePrevious*) returns true. Otherwise, you will get an error, because the **Recordset** object might have been closed and the functionality rendered unavailable on the instance. If a feature you are interested in is not supported, **Supports** will return false as well. In this case, you should avoid calling the corresponding property or method on the **Recordset** object.

Recordset objects can support two types of updating: immediate and batched. In immediate updating, all changes to data are written immediately to the underlying data source once you call the Update method. You can also pass arrays of values as parameters with the AddNew and Update methods and simultaneously update several fields in a record.

If a provider supports batch updating, you can have the provider cache changes to more than one record and then transmit them in a single call to the database with the UpdateBatch method. This applies to changes made with the AddNew, Update, and Delete methods. After you call the UpdateBatch method, you can use the Status property to check for any data conflicts in order to resolve them.

**Note**: To execute a query without using a **Command** object, pass a query string to the **Open** method of a **Recordset** object. However, a **Command** object is required when you want to persist the command text and re-execute it, or use query parameters.

The **Mode** property governs access permissions.

A **Recordset** object has a **Fields** collection made up of **Field** objects. Each **Field** object corresponds to a column in the **Recordset**. You can populate the **Fields** collection before opening the **Recordset** by calling the **Refresh** method on the collection.

The **Fields** collection has an **Append** method, which provisionally creates and adds a **Field** object to the collection, and an Update method, which finalizes any additions or deletions.

Certain providers (for example, the Microsoft OLE DB Provider for Internet Publishing) may populate the **Fields** collection with a subset of available fields for the **Record** or **Recordset**. Other fields will not be added to the collection until they are first referenced by name or indexed by your code.

If you attempt to reference a nonexistent field by name, a new **Field** object will be appended to the **Fields** collection with a **Status** of **adFieldPendingInsert**. When you call Update, ADO will create a new field in your data source if allowed by your provider.

Each **Field** object corresponds to a column in the **Recordset**. You use the **Value** property of **Field** objects to set or return data for the current record. Depending on the functionality the provider exposes, some collections, methods, or properties of a **Field** object may not be available.

With the collections, methods, and properties of a **Field** object, you can do the following:

* Return the name of a field with the **Name** property.
* View or change the data in the field with the **Value** property. **Value** is the default property of the **Field** object.
* Return the basic characteristics of a field with the **Type_**, **Precision**, and **NumericScale** properties.
* Return the declared size of a field with the **DefinedSize** property.
* Return the actual size of the data in a given field with the **ActualSize** property.
* Determine what types of functionality are supported for a given field with the **Attributes** property and **Properties** collection.
* Manipulate the values of fields containing long binary or long character data with the **AppendChunk** and **GetChunk** methods.
If the provider supports batch updates, resolve discrepancies in field values during batch updating with the **OriginalValue** and **UnderlyingValue** properties.

All of the metadata properties (**Name**, **Type_**, **DefinedSize**, **Precision**, and **NumericScale**) are available before opening the **Field** object's **Recordset**. Setting them at that time is useful for dynamically constructing forms.

When a **Recordset** object is passed across processes, only the rowset values are marshalled, and the properties of the **Recordset** object are ignored. During unmarshalling, the rowset is unpacked into a newly created **Recordset** object, which also sets its properties to the default values.

# <a name="AbsolutePage"></a>AbsolutePage

Sets or returns a Long value from 1 to the number of pages in the **Recordset** object (**PageCount**), or returns one of the **PositionEnum** values.

```
PROPERTY AbsolutePage () AS PositionEnum
PROPERTY AbsolutePage (BYVAL Page AS PositionEnum)
```

#### PositionEnum

Specifies the current position of the record pointer within a Recordset.

| Constant   | Description |
| ---------- | ----------- |
| **adPosBOF** | Indicates that the current record pointer is at **BOF** (that is, the **BOF** property is True). |
| **adPosEOF** | Indicates that the current record pointer is at **EOF** (that is, the **EOF** property is True). |
| **adPosUnknown** | Indicates that the **Recordset** is empty, the current position is unknown, or the provider does not support the **AbsolutePage** or **AbsolutePosition** property. |

#### Return value

The page number.

#### Remarks

This property can be used to identify the page number on which the current record is located. It uses the **PageSize** property to logically divide the total rowset count of the **Recordset** object into a series of pages, each of which has the number of records equal to **PageSize** (except for the last page, which may have fewer records). The provider must support the appropriate functionality for this property to be available.

When getting or setting the **AbsolutePage** property, ADO uses the **AbsolutePosition** property and the **PageSize** property together as follows:

* To get the **AbsolutePage**, ADO first retrieves the **AbsolutePosition**, and then divides it by the **PageSize**.
* To set the **AbsolutePage**, ADO moves the **AbsolutePosition** as follows: it multiplies the **PageSize** by the new **AbsolutePage** value and then adds 1 to the value. As a result, the current position in the **Recordset** after successfully setting **AbsolutePage** is, the first record in that page.

Like the **AbsolutePosition** property, **AbsolutePage** is 1-based and equals 1 when the current record is the first record in the **Recordset**. Set this property to move to the first record of a particular page. Obtain the total number of pages from the **PageCount** property.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Set the cursor location
DIM pRecordset AS CAdoRecordset
pRecordset.CursorLocation = adUseClient

' // Open the recordset
DIM cvSource AS CVAR = "SELECT * FROM Publishers"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Display five records at a time
DIM nPageSize AS LONG = 5
pRecordset.PageSize = nPageSize
' // Retrieve the number of pages
DIM nPageCount AS LONG = pRecordset.PageCount

' // Parse the recordset
FOR i AS LONG = 1 TO nPageCount
   ' // Set the cursor at the beginning of the page
   pRecordset.AbsolutePage = i
   FOR x AS LONG = 1 TO nPageSize
      ' // Get the content of the "Name" column
      DIM cvRes AS CVAR = pRecordset.Collect("Name")
      PRINT cvRes
      ' // Fetch the next row
      pRecordset.MoveNext
      IF pRecordset.EOF THEN EXIT FOR
   NEXT
   PRINT
NEXT
```

# <a name="AbsolutePosition"></a>AbsolutePosition

Sets or returns a Long value from 1 to the number of records in the **Recordset** object (**RecordCount**), or returns one of the **PositionEnum** values.

```
PROPERTY AbsolutePosition () AS PositionEnum
PROPERTY AbsolutePosition (BYVAL Position AS PositionEnum)
```

#### PositionEnum

Specifies the current position of the record pointer within a Recordset.

| Constant   | Description |
| ---------- | ----------- |
| **adPosBOF** | Indicates that the current record pointer is at **BOF** (that is, the **BOF** property is True). |
| **adPosEOF** | Indicates that the current record pointer is at **EOF** (that is, the **EOF** property is True). |
| **adPosUnknown** | Indicates that the **Recordset** is empty, the current position is unknown, or the provider does not support the **AbsolutePage** or **AbsolutePosition** property. |

#### Return value

The absolute position.

#### Remarks

In order to set the **AbsolutePosition** property, ADO requires that the OLE DB provider you are using implement the **IRowsetLocate** interface.

Accessing the **AbsolutePosition** property of a **Recordset** that was opened with either a forward-only or dynamic cursor raises the error **adErrFeatureNotAvailable**. With other cursor types, the correct position will be returned as long as the provider supports the **IRowsetScroll** interface. If the provider does not support the **IRowsetScroll** interface, the property is set to **adPosUnknown**. See the documentation for your provider to determine whether it supports **IRowsetScroll**.

Use the **AbsolutePosition** property to move to a record based on its ordinal position in the **Recordset** object, or to determine the ordinal position of the current record. The provider must support the appropriate functionality for this property to be available.

Like the **AbsolutePage** property, **AbsolutePosition** is 1-based and equals 1 when the current record is the first record in the **Recordset**. You can obtain the total number of records in the **Recordset** object from the **RecordCount** property.

When you set the **AbsolutePosition** property, even if it is to a record in the current cache, ADO reloads the cache with a new group of records starting with the record you specified. The **CacheSize** property determines the size of this group.

**Note**: You should not use the **AbsolutePosition** property as a surrogate record number. The position of a given record changes when you delete a preceding record. There is also no assurance that a given record will have the same **AbsolutePosition** if the **Recordset** object is requeried or reopened. Bookmarks are still the recommended way of retaining and returning to a given position and are the only way of positioning across all types of **Recordset** objects.

#### Example

This example demonstrates how the **AbsolutePosition** property can track the progress of a loop that enumerates all the records of a **Recordset**. It uses the **CursorLocation** property to enable the **AbsolutePosition** property by setting the cursor to a client cursor.

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Set the cursor location
DIM pRecordset AS CAdoRecordset
pRecordset.CursorLocation = adUseClient

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTable)

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents of the "City" and "Name" columns
   DIM cvRes AS CVAR = pRecordset.Collect("Name")
   PRINT "Position: "; pRecordset.AbsolutePosition; " "; cvRes

   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```
