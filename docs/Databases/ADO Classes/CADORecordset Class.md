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

# <a name="ActiveCommand"></a>ActiveCommand

Indicates the **Command** object that created the associated **Recordset** object.

```
PROPERTY ActiveCommand () AS Afx_ADOCommand PTR
```

#### Return value

A **Command** object reference.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' ========================================================================================
' The ShowActiveCommand routine is given only a Recordset object, yet it must print the
' command text and parameter that created the Recordset. This can be done because the
' Recordset object's ActiveCommand property yields the associated Command object.
' The Command object's CommandText property yields the parameterized command that was
' substituted for the command's parameter placeholder ("?").
' ========================================================================================
SUB ShowActiveCommand (BYREF pConnection AS CAdoConnection, BYREF pRecordset AS CAdoRecordset)

   DIM pCommand AS CAdoCommand = pRecordset.ActiveCommand
   DIM cbsCommandText AS CBSTR = pCommand.CommandText
   DIM pParameters AS CAdoParameters = pCommand.Parameters
   DIM pParameter AS CAdoParameter = pParameters.Item("Name")
   DIM cvValue AS CVAR = pParameter.Value
   PRINT "Command text: "; cbsCommandText
   PRINT "Parameter: "; cvValue

   IF pRecordset.BOF THEN
      PRINT "Name = '"; cvValue; "'not found"
   ELSE
      DIM cvAuthor AS CVAR = pRecordset.Collect("Author")
      DIM cvID AS CVAR = pRecordset.Collect("Author")
      PRINT "Name= "; cvAuthor; ","; cvID
   END IF

END SUB
' ========================================================================================

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' Set the ADOCommand active connection
DIM pCommand AS CAdoCommand
pCommand.ActiveConnection = pConnection

' //Set the command text
pCommand.CommandText = "SELECT * FROM Authors WHERE Author = ?"
' // Create the parameter
DIM pParameter AS CADOParameter = pCommand.CreateParameter("Name", adChar, adParamInput, 255, "Bard, Dick")
' // Add the parameter to the collection
DIM pParameters AS CAdoParameters = pCommand.Parameters
pParameters.Append(pParameter)

' // Create the recordset by executing the command string
DIM pRecordset AS CAdoRecordset = pCommand.Execute
' // Display the results
ShowActiveCommand pConnection, pRecordset
```

# <a name="ActiveConnection"></a>ActiveConnection

Sets or returns an string value that contains a definition for a connection if the connection is closed, or a Variant containing the current **Connection** object if the connection is open. Default is a null object reference.

```
PROPERTY ActiveConnection (BYREF vConn AS CVAR)
PROPERTY ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
PROPERTY ActiveConnection (BYREF pconn AS CAdoConnection)
PROPERTY ActiveConnection () AS Afx_ADOConnection
```

| Parameter  | Description |
| ---------- | ----------- |
| *vConn* | An string value that contains a definition for a connection if the connection is closed, or a Variant containing the current **Connection** object if the connection is open. |
| *pconn* | A reference to the **Connection** object or to the **CAdoConnection** class. |

#### Return value

A reference to the **Afx_ADOConnection** object. You must release it calling the **Release** method when no longer needed.

#### Remarks

Use the **ActiveConnection** property to determine the **Connection** object over which the specified **Recordset** will be opened.

For open **Recordset** objects or for **Recordset** objects whose **Source** property is set to a valid **Command** object, the **ActiveConnection** property is read-only. Otherwise, it is read/write.

You can set this property to a valid **Connection** object or to a valid connection string. In this case, the provider creates a new **Connection** object using this definition and opens the connection. Additionally, the provider may set this property to the new **Connection** object to give you a way to access the **Connection** object for extended error information or to execute other commands.

If you use the **ActiveConnection** argument of the **Open** method to open a **Recordset** object, the **ActiveConnection** property will inherit the value of the argument.

If you set the **Source** property of the **Recordset** object to a valid **Command** object variable, the **ActiveConnection** property of the **Recordset** inherits the setting of the **Command** object's **ActiveConnection** property.

#### Remote Data Service Usage

When used on a client-side **Recordset** object, this property can be set only to a connection string or to NULL.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Recordset object
DIM pRecordset AS CAdoRecordset
' // Set the active connection
pRecordset.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
' // Open the recordset
DIM cvSource AS CVAR = "SELECT TOP 20 * FROM Authors ORDER BY Author"
DIM hr AS HRESULT = pRecordset.Open(cvSource)
```

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Recordset object
DIM pRecordset AS CAdoRecordset
' // Open the connection
DIM ccConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr
' // Set the active connection
pRecordset.ActiveConnection = pConnection
' // Open the recordset
DIM cvSource AS CVAR = "SELECT TOP 20 * FROM Authors ORDER BY Author"
DIM hr AS HRESULT = pRecordset.Open(cvSource)
```

#### Disconnected recordset

A disconnected recordset is one of the powerful features of ADO wherein the connection is removed from a populated recordset. This recordset can be manipulated and again connected to the database for updating. Remote Data Services (RDS) uses this feature to send recordsets through either HTTP or Distributed Component Object Model (DCOM) protocols to a remote computer, however, you are not limited to using Remote Data Service (RDS) to generate a disconnected recordset.

We can manipulate ADO directly to disconnect a recordset without using either RDS Server or Client side components.

This technique is demonstrated below and is accomplished by setting the **ActiveConnection** property.

One of the primary requisites for a recordset to become a disconnected recordset is that it should use client side cursors. That is, the **CursorLocation** should be initialized to **adUseClient**.

Disconnecting a recordset can be done by setting the **ActiveConnection** property to NULL.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection PTR = NEW CAdoConnection
pConnection->Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
' // Setting the cursor location to client side is important to get a disconnected recordset
pRecordset.CursorLocation = adUseClient
' // Open the recordset
DIM cvSource AS CVAR = "SELECT * FROM Authors"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Disconnect the recordset by setting its active connection to null.
' // Casting to Afx_ADOConnection PTR is needed to get the correct overloaded method called;
' // otherwise, the CVAR version will be called and it will fail.
pRecordset.ActiveConnection = cast(Afx_ADOConnection PTR, NULL)

' // Close and release the connection
Delete pConnection

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
   DIM cvRes AS CVAR = pRecordset.Collect("Author")
   PRINT cvRes
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="AddNew"></a>AddNew

Creates a new record for an updatable **Recordset** object.

```
FUNCTION AddNew ( _
   BYVAL vFieldList AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
   BYVAL vValues AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND) _
) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vFieldList* | Optional. A single name, or an array of names or ordinal positions of the fields in the new record. |
| *vValues* | Optional. A single value, or an array of values for the fields in the new record. If *vFieldlist* is an array, *vValues* must also be an array with the same number of members; otherwise, an error occurs. The order of field names must match the order of field values in each array. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **AddNew** method to create and initialize a new record. Use the **Supports** method with **adAddNew** (a **CursorOptionEnum** value) to verify whether you can add records to the current **Recordset** object.

After you call the **AddNew** method, the new record becomes the current record and remains current after you call the **Update** method. Since the new record is appended to the **Recordset**, a call to **MoveNext** following the **Update** will move past the end of the **Recordset**, making **EOF** True. If the **Recordset** object does not support bookmarks, you may not be able to access the new record once you move to another record. Depending on your cursor type, you may need to call the **Requery** method to make the new record accessible.

If you call **AddNew** while editing the current record or while adding a new record, ADO calls the Update method to save any changes and then creates the new record.

The behavior of the **AddNew** method depends on the updating mode of the **Recordset** object and whether you pass the *vFieldlist* and *vValues* arguments.

In immediate update mode (in which the provider writes changes to the underlying data source once you call the **Update** method), calling the **AddNew** method without arguments sets the **EditMode** property to **adEditAdd** (an **EditModeEnum** value). The provider caches any field value changes locally. Calling the **Update** method posts the new record to the database and resets the **EditMode** property to **adEditNone** (an **EditModeEnum** value). If you pass the *vFieldlist* and *vValues* arguments, ADO immediately posts the new record to the database (no Update call is necessary); the **EditMode** property value does not change (**adEditNone**).

In batch update mode (in which the provider caches multiple changes and writes them to the underlying data source only when you call the **UpdateBatch** method), calling the **AddNew** method without arguments sets the **EditMode** property to **adEditAdd**. The provider caches any field value changes locally. Calling the Update method adds the new record to the current **Recordset**, but the provider does not post the changes to the underlying database, or reset the **EditMode** to **adEditNone**, until you call the **UpdateBatch** method. If you pass the *vFieldlist* and *vValues* arguments, ADO sends the new record to the provider for storage in a cache and sets the **EditMode** to **adEditAdd**; you need to call the **UpdateBatch** method to post the new record to the underlying database.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Recordset object
DIM pRecordset AS CAdoRecordset

' // Open the connection
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTableDirect)

' // Add a new record
pRecordset.AddNew
   pRecordset.Collect("PubID") = CLNG(10000)
   pRecordset.Collect("Name") = "Wile E. Coyote"
   pRecordset.Collect("Company Name") = "Warner Brothers Studios"
   pRecordset.Collect("Address") = "4000 Warner Boulevard"
   pRecordset.Collect("City") = "Burbank, CA. 91522"
pRecordset.Update
```

# <a name="Attach"></a>Attach

Attaches a recordset to the class.

```
FUNCTION Attach (BYVAL pRecordset AS AFX_ADORecordset PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pRecordset* | A reference to an **Afx_ADORecordset** object. |
| *fAddRef* | TRUE = increase the reference count; FALSE = don't increase the reference count. |

#### Return value

S_OK or an HRESULT code.

# <a name="BOF"></a>BOF

Indicates that the current record position is before the first record in a **Recordset** object.

```
PROPERTY BOF () AS BOOLEAN
```

#### Return value

TRUE if the current record position is before the first record; FALSE, otherwise.

# <a name="EOF"></a>EOF

Indicates that the current record position is after the last record in a **Recordset** object.

```
PROPERTY EOF () AS BOOLEAN
```

#### Return value

TRUE if the current record position is after the last record; FALSE, otherwise.

# <a name="Bookmark"></a>Bookmark

Sets or returns a Variant expression that evaluates to a valid bookmark.

```
PROPERTY Bookmark () AS CVAR
PROPERTY Bookmark (BYREF cvBookmark AS CVAR)
```

#### Return value

The bookmark.

#### Remarks

Use the **Bookmark** property to save the position of the current record and return to that record at any time. Bookmarks are available only in **Recordset** objects that support bookmark functionality.

When you open a **Recordset** object, each of its records has a unique bookmark. To save the bookmark for the current record, assign the value of the **Bookmark** property to a variable. To quickly return to that record at any time after moving to a different record, set the **Recordset** object's **Bookmark** property to the value of that variable.

The user may not be able to view the value of the bookmark. Also, users should not expect bookmarks to be directly comparable—two bookmarks that refer to the same record may have different values.

If you use the **Clone** method to create a copy of a **Recordset** object, the **Bookmark** property settings for the original and the duplicate **Recordset** objects are identical and you can use them interchangeably. However, you cannot use bookmarks from different **Recordset** objects interchangeably, even if they were created from the same source or command.

#### Remote Data Service Usage

When used on a client-side Recordset object, the **Bookmark** property is always available.


# <a name="CacheSize"></a>CacheSize

Sets or returns a Long value that must be greater than 0. Default is 1.

```
PROPERTY CacheSize () AS LONG
PROPERTY CacheSize (BYVAL size AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *size* | A value that must be greater than 0. |

#### Return value

The cache size.

#### Remarks

Use the **CacheSize** property to control how many records to retrieve at one time into local memory from the provider. For example, if the **CacheSize** is 10, after first opening the **Recordset** object, the provider retrieves the first 10 records into local memory. As you move through the **Recordset** object, the provider returns the data from the local memory buffer. As soon as you move past the last record in the cache, the provider retrieves the next 10 records from the data source into the cache.

**Note**: **CacheSize** is based on the Maximum Open Rows provider-specific property (in the **Properties** collection of the **Recordset** object). You cannot set **CacheSize** to a value greater than Maximum Open Rows. To modify the number of rows which can be opened by the provider, set Maximum Open Rows.

The value of **CacheSize** can be adjusted during the life of the **Recordset** object, but changing this value only affects the number of records in the cache after subsequent retrievals from the data source. Changing the property value alone will not change the current contents of the cache.

If there are fewer records to retrieve than **CacheSize** specifies, the provider returns the remaining records and no error occurs.

A **CacheSize** setting of zero is not allowed and returns an error.

Records retrieved from the cache don't reflect concurrent changes that other users made to the source data. To force an update of all the cached data, use the **Resync** method.

If **CacheSize** is set to a value greater than one, the navigation methods (**Move**, **MoveFirst**, **MoveLast**, **MoveNext**, and **MovePrevious**) may result in navigation to a deleted record, if deletion occurs after the records were retrieved. After the initial fetch, subsequent deletions will not be reflected in your data cache until you attempt to access a data value from a deleted row. However, setting **CacheSize** to one eliminates this issue since deleted rows cannot be fetched.

# <a name="Cancel"></a>Cancel

Cancels execution of a pending, asynchronous method call.

```
FUNCTION Cancel () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Cancel** method to terminate execution of an asynchronous method call (that is, a method invoked with the **adAsyncConnect**, **adAsyncExecute**, or **adAsyncFetch** option).

For a **Recordset** object, the last asynchronous call to the **Open** method is terminated.

# <a name="CancelBatch"></a>CancelBatch

Cancels a pending batch update.

```
FUNCTION CancelBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *AffectRecords* | Optional. An **AffectEnum** value that indicates how many records the **CancelBatch** method will affect. |

#### AffectEnum

Specifies which records are affected by an operation.

| Constant   | Description |
| ---------- | ----------- |
| **adAffectAll** | If there is not a Filter applied to the **Recordset**, affects all records. If the **Filter** property is set to a string criteria (such as "Author='Smith'"), then the operation affects visible records in the current chapter. If the **Filter** property is set to a member of the **FilterGroupEnum** or an array of Bookmarks, then the operation will affect all rows of the **Recordset**. |
| **adAffectAllChapters** | Affects all records in all sibling chapters of the **Recordset**, including those not visible via any **Filter** that is currently applied. |
| **adAffectCurrent** | Affects only the current record. |
| **adAffectGroup** | Affects only records that satisfy the current **Filter** property setting. You must set the **Filter** property to a **FilterGroupEnum** value or an array of Bookmarks to use this option. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **CancelBatch** method to cancel any pending updates in a **Recordset** in batch update mode. If the **Recordset** is in immediate update mode, calling **CancelBatch** without **adAffectCurrent** generates an error.

If you are editing the current record or are adding a new record when you call **CancelBatch**, ADO first calls the **CancelUpdate** method to cancel any cached changes. After that, all pending changes in the **Recordset** are canceled.

It's possible that the current record will be indeterminable after a **CancelBatch** call, especially if you were in the process of adding a new record. For this reason, it is prudent to set the current record position to a known location in the **Recordset** after the **CancelBatch** call. For example, call the **MoveFirst** method.

If the attempt to cancel the pending updates fails because of a conflict with the underlying data (for example, a record has been deleted by another user), the provider returns warnings to the Errors collection but does not halt program execution. A run-time error occurs only if there are conflicts on all the requested records. Use the **Filter** property (**adFilterAffectedRecords**) and the **Status** property to locate records with conflicts.

# <a name="CancelUpdate"></a>CancelUpdate

Cancels any changes made to the current or new row of a **Recordset** object before calling the **Update** method.

```
FUNCTION CancelUpdate () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **CancelUpdate** method to cancel any changes made to the current row or to discard a newly added row. You cannot cancel changes to the current row or a new row after you call the Update method, unless the changes are either part of a transaction that you can roll back with the **RollbackTrans** method, or part of a batch update. In the case of a batch update, you can cancel the **Update** with the **CancelUpdate** or **CancelBatch** method.

If you are adding a new row when you call the **CancelUpdate** method, the current row becomes the row that was current before the **AddNew** call.

If you are in edit mode and want to move off the current record (for example, with **Move**, **NextRecordset**, or **Close**), you can use **CancelUpdate** to cancel any pending changes. You may need to do this if the update cannot successfully be posted to the data source (for example, an attempted delete that fails due to referential integrity violations will leave the **Recordset** in edit mode after a call to **Delete_**).

# <a name="Clone"></a>Clone

Creates a duplicate **Recordset** object from an existing **Recordset object**. Optionally, specifies that the clone be read-only.

```
FUNCTION Clone (BYVAL nLockType AS LockTypeEnum = adLockUnspecified) AS Afx_AdoRecordset PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nLockType* | Optional. A **LockTypeEnum** value that specifies either the lock type of the original Recordset, or a read-only Recordset. Valid values are **adLockUnspecified** or **adLockReadOnly**. |

#### LockTypeEnum

Specifies the type of lock placed on records during editing.

| Constant   | Description |
| ---------- | ----------- |
| **adLockBatchOptimistic** | Indicates optimistic batch updates. Required for batch update mode. |
| **adLockOptimistic** | Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the Update method. |
| **adLockPessimistic** | Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing. |
| **adLockReadOnly** | Indicates read-only records. You cannot alter the data. |
| **adLockUnspecified** | Does not specify a type of lock. For clones, the clone is created with the same lock type as the original. |

#### Return value

An **Afx_ADORecordset** object reference.

#### Remarks

Use the **Clone** method to create multiple, duplicate **Recordset** objects, particularly if you want to maintain more than one current record in a given set of records. Using the **Clone** method is more efficient than creating and opening a new Recordset object with the same definition as the original.

The **Filter** property of the original **Recordset**, if any, will not be applied to the clone. Set the **Filter** property of the new **Recordset** in order to filter the results. The simplest way to copy any existing **Filter** value is to assign it directly, like this: 

The current record of a newly created clone is set to the first record.

Changes you make to one **Recordset** object are visible in all of its clones regardless of cursor type. However, after you execute **Requery** on the original **Recordset**, the clones will no longer be synchronized to the original.

Closing the original **Recordset** does not close its copies, nor does closing a copy close the original or any of the other copies.

You can only clone a **Recordset** object that supports bookmarks. Bookmark values are interchangeable; that is, a bookmark reference from one **Recordset** object refers to the same record in any of its clones.

Some **Recordset** events that are triggered will also fire in all **Recordset** clones. However, because the current record can differ between cloned **Recordsets**, the events may not be valid for the clone. For example, if you change a value of a field, a **WillChangeField** event will occur in the changed **Recordset** and in all clones. The **Fields** parameter of the **WillChangeField** event of a cloned **Recordset** (where the change was not made) will simply refer to the fields of the current record of the clone, which may be a different record than the current record of the original **Recordset** where the change occurred.

The following table provided a full listing of all **Recordset** events and indicates whether they are valid and triggered for any recordset clones generated using the **Clone** method.

| Event      | Triggered in clones? |
| ---------- | -------------------- |
| **EndOfRecordset** | No |
| **FetchComplete** | No |
| **FetchProgress** | No |
| **FieldChangeComplete** | Yes |
| **MoveComplete** | No |
| **RecordChangeComplete** | Yes |
| **RecordsetChangeComplete** | No |
| **WillChangeField** | Yes |
| **WillChangeRecord** | Yes |
| **WillChangeRecordset** | No |
| **WillMove** | No |
