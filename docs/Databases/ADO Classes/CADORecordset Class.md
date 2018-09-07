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

If the attempt to cancel the pending updates fails because of a conflict with the underlying data (for example, a record has been deleted by another user), the provider returns warnings to the **Errors** collection but does not halt program execution. A run-time error occurs only if there are conflicts on all the requested records. Use the **Filter** property (**adFilterAffectedRecords**) and the **Status** property to locate records with conflicts.

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
| *nLockType* | Optional. A **LockTypeEnum** value that specifies either the lock type of the original **Recordset**, or a read-only Recordset. Valid values are **adLockUnspecified** or **adLockReadOnly**. |

#### LockTypeEnum

Specifies the type of lock placed on records during editing.

| Constant   | Description |
| ---------- | ----------- |
| **adLockBatchOptimistic** | Indicates optimistic batch updates. Required for batch update mode. |
| **adLockOptimistic** | Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the **Update** method. |
| **adLockPessimistic** | Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing. |
| **adLockReadOnly** | Indicates read-only records. You cannot alter the data. |
| **adLockUnspecified** | Does not specify a type of lock. For clones, the clone is created with the same lock type as the original. |

#### Return value

An **Afx_ADORecordset** object reference.

#### Remarks

Use the **Clone** method to create multiple, duplicate **Recordset** objects, particularly if you want to maintain more than one current record in a given set of records. Using the **Clone** method is more efficient than creating and opening a new **Recordset** object with the same definition as the original.

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

# <a name="Close"></a>Close

Closes a **Recordset** object and any dependent objects.

```
FUNCTION Close () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Close** method to close a **Recordset** to free any associated system resources. Closing an object does not remove it from memory; you can change its property settings and open it again later. To completely eliminate an object from memory, release the connection calling the **Release** method.

# <a name="Collect"></a>Collect

Sets or returns a Variant value that indicates the value of the object

The ADO Recorset object exposes a hidden member: the **Collect** property. This property is functionally similar to the **Field**'s **Value** property, but it doesn't need a reference (explicit or implicit) to the **Field** object. You can pass either a numeric index or a field's name to this property.

```
PROPERTY Collect (BYREF cvIndex AS CVAR) AS CVAR
PROPERTY Collect (BYREF cvIndex AS CVAR, BYREF cvValue AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvIndex* | The zero-based ordinal number of the field or the name of the field. |
| *cvValue* | The value to assign to the field. |

#### Return value

The value of the field.

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

# <a name="CompareBookmarks"></a>CompareBookmarks

Compares two bookmarks and returns an indication of their relative values.

```
FUNCTION CompareBookmarks (BYREF cvBookmark1 AS CVAR, BYREF cvBookmark2 AS CVAR) AS CompareEnum
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvBookmark1* | The bookmark of the first row. |
| *cvBookmark2* | The bookmark of the second row. |

#### Return value

A **CompareEnum** value.

#### CompareEnum

Specifies the relative position of two records represented by their bookmarks.

| Constant   | Description |
| ---------- | ----------- |
| **adCompareEqual** | Indicates that the bookmarks are equal. |
| **adCompareGreaterThan** | Indicates that the first bookmark is after the second. |
| **adCompareLessThan** | Indicates that the first bookmark is before the second. |
| **adCompareNotComparable** | Indicates that the bookmarks cannot be compared. |
| **adCompareNotEqual** | Indicates that the bookmarks are not equal and not ordered. |

#### Remarks

The bookmarks must apply to the same **Recordset** object, or a **Recordset** object and its clone. You cannot reliably compare bookmarks from different **Recordset** objects, even if they were created from the same source or command. Nor can you compare bookmarks for a **Recordset** object whose underlying provider does not support comparisons.

A bookmark uniquely identifies a row in a **Recordset** object. Use the current row's **Bookmark** property to obtain its bookmark.

Because the data type of a bookmark is provider specific, ADO exposes it as a Variant. For example, SQL Server bookmarks are of type DBTYPE_R8 (Double). ADO would expose this type as a Variant with a subtype of Double.

When comparing bookmarks, ADO does not attempt any type of coercion. The values are simply passed to the provider where the compare occurs. If bookmarks passed to the **CompareBookmarks** method are stored in variables of differing types, it can generate the type mismatch error, "Arguments are of the wrong type, are out of the acceptable range, or are in conflict with each other."

A bookmark that is not valid or incorrectly formed will cause an error.

# <a name="CursorLocation"></a>CursorLocation

Indicates the location of the cursor service.

```
PROPERTY CursorLocation () AS CursorLocationEnum
PROPERTY CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lCursorLoc* | One of the **CursorLocationEnum** values |

#### Return value

A **CursorLocationEnum** value.

#### CursorLocationEnum

Specifies the location of the cursor service.

| Constant   | Description |
| ---------- | ----------- |
| **adUseClient** | Uses client-side cursors supplied by a local cursor library. Local cursor services often will allow many features that driver-supplied cursors may not, so using this setting may provide an advantage with respect to features that will be enabled. For backward compatibility, the synonym **adUseClientBatch** is also supported. |
| **adUseNone** | Does not use cursor services. (This constant is obsolete and appears solely for the sake of backward compatibility.) |
| **adUseServer** | Default. Uses data-provider or driver-supplied cursors. These cursors are sometimes very flexible and allow for additional sensitivity to changes others make to the data source. However, some features of the Microsoft Cursor Service for OLE DB (such as disassociated Recordset objects) cannot be simulated with server-side cursors and these features will be unavailable with this setting. |

# <a name="CursorType"></a>CursorType

Sets or returns a **CursorTypeEnum** value. The default value is **adOpenForwardOnly**.

```
PROPERTY CursorType () AS CursorTypeEnum
PROPERTY CursorType (BYVAL lCursorType AS CursorTypeEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lCursorType* | One of the **CursorTypeEnum** values. |

#### Return value

A **CursorTypeEnum** value.

#### CursorTypeEnum

Specifies the type of cursor used in a **Recordset** object.

| Constant   | Description |
| ---------- | ----------- |
| **adOpenDynamic** | Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the **Recordset** are allowed, except for bookmarks, if the provider doesn't support them. |
| **adOpenForwardOnly** | Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a **Recordset**. |
| **adOpenKeyset** | Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible. |
| **adOpenStatic** | Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible. |
| **adOpenUnspecified** | Does not specify the type of cursor. |

# <a name="DataMember"></a>DataMember

Indicates the name of the data member that will be retrieved from the object referenced by the **DataSource** property.

```
PROPERTY DataMember () AS CBSTR
PROPERTY DataMember (BYVAL cbsDataMember AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsDataMember* | Name of the data member. Not case sensitive. |

#### Return value

The name of the data member.

#### Remarks

This property is used to create data-bound controls with the Data Environment. The Data Environment maintains collections of data (data sources) containing named objects (data members) that will be represented as a **Recordset** object.

The **DataMember** and **DataSource** properties must be used in conjunction.

The **DataMember** property determines which object specified by the **DataSource** property will be represented as a **Recordset** object. The **Recordset** object must be closed before this property is set. An error is generated if the **DataMember** property isn't set before the **DataSource** property, or if the **DataMember** name isn't recognized by the object specified in the **DataSource** property.

# <a name="DataSource"></a>DataSource

Indicates an object that contains data to be represented as a **Recordset** object.

```
PROPERTY DataSource () AS IUnknown PTR
PROPERTY DataSource (BYVAL punkDataSource AS IUnknown PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *punkDataSource* | An object reference to the data source. |

#### Return value

An object reference to the data source.

#### Remarks

This property is used to create data-bound controls with the Data Environment. The Data Environment maintains collections of data (data sources) containing named objects (data members) that will be represented as a **Recordset** object.

The **DataMember** and **DataSource** properties must be used in conjunction.

The object referenced must implement the **IDataSource** interface and must contain an **IRowset** interface.


# <a name="Delete_"></a>Delete_

Deletes the current record or a group of records.

```
FUNCTION Delete_ (BYVAL AffectRecords AS AffectEnum = adAffectCurrent) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *AffectRecords* | Optional. An **AffectEnum** value that determines how many records the **Delete_** method will affect. The default value is **adAffectCurrent**. Note: **adAffectAll** and **adAffectAllChapters** are not valid arguments to Delete_. |

#### AffectEnum

Specifies which records are affected by a **Delete_** operation.

| Constant   | Description |
| ---------- | ----------- |
| **adAffectCurrent** | Affects only the current record. |
| **adAffectGroup** | Affects only records that satisfy the current **Filter** property setting. You must set the **Filter** property to a **FilterGroupEnum** value or an array of Bookmarks to use this option. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Using the **Delete_** method marks the current record or a group of records in a **Recordset** object for deletion. If the **Recordset** object doesn't allow record deletion, an error occurs. If you are in immediate update mode, deletions occur in the database immediately. If a record cannot be successfully deleted (due to database integrity violations, for example), the record will remain in edit mode after the call to **Update**. This means that you must cancel the update with **CancelUpdate** before moving off the current record (for example, with **Close**, **Move**, or **NextRecordset**).

If you are in batch update mode, the records are marked for deletion from the cache and the actual deletion happens when you call the **UpdateBatch** method. (Use the **Filter** property to view the deleted records.)

Retrieving field values from the deleted record generates an error. After deleting the current record, the deleted record remains current until you move to a different record. Once you move away from the deleted record, it is no longer accessible.

If you nest deletions in a transaction, you can recover deleted records with the RollbackTrans method. If you are in batch update mode, you can cancel a pending deletion or group of pending deletions with the **CancelBatch** method.

If the attempt to delete records fails because of a conflict with the underlying data (for example, a record has already been deleted by another user), the provider returns warnings to the **Errors** collection but does not halt program execution. A run-time error occurs only if there are conflicts on all the requested records.

If the Unique Table dynamic property is set, and the **Recordset** is the result of executing a JOIN operation on multiple tables, then the **Delete_** method will only delete rows from the table named in the Unique Table property.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Publishers WHERE PubID=10000"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

DIM cvRes AS CVAR = pRecordset.Collect("PubID")
IF cvRes.ValInt = 10000 THEN
   pRecordset.Delete_ adAffectCurrent
   PRINT "Record deleted"
ELSE
   PRINT "Record not found"
END IF
```

# <a name="EditMode"></a>EditMode

Indicates the editing status of the current record.

```
PROPERTY EditMode () AS LONG
```

#### EditModeEnum

Specifies the editing status of a record.

| Constant   | Description |
| ---------- | ----------- |
| **adEditNone** | Indicates that no editing operation is in progress. |
| **adEditInProgress** | Indicates that data in the current record has been modified but not saved. |
| **adEditAdd** | Indicates that the **AddNew** method has been called, and the current record in the copy buffer is a new record that has not been saved in the database. |
| **adEditDelete** | Indicates that the current record has been deleted. |

# <a name="Fields"></a>Fields

Gets a reference to the **Fields** collection of a **Record** object.

```
PROPERTY Fields () AS Afx_ADOFields PTR
```

#### Return value

An **Afx_ADOFields** object reference.

# <a name="Filter"></a>Filter

Indicates a filter for data in a **Recordset**.

```
PROPERTY Filter () AS CVAR
PROPERTY Filter (BYVAL cvCriteria AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvCriteria* | Can be one of the following values:<br><ul><li>Criteria string — a string made up of one or more individual clauses concatenated with AND or OR operators.</li><li>Array of bookmarks — an array of unique bookmark values that point to records in the **Recordset** object.</li></li><li>A **FilterGroupEnum** value.</li></ul> |

#### Return value

The filter criteria value.

#### Remarks

Use the **Filter** property to selectively screen out records in a **Recordset** object. The filtered **Recordset** becomes the current cursor. Other properties that return values based on the current cursor are affected, such as **AbsolutePosition**, **AbsolutePage**, **RecordCount**, and **PageCount**. This is because setting the **Filter** property to a specific value will move the current record to the first record that satisfies the new value.

The criteria string is made up of clauses in the form FieldName-Operator-Value (for example, "LastName = 'Smith'"). You can create compound clauses by concatenating individual clauses with AND (for example, "LastName = 'Smith' AND FirstName = 'John'") or OR (for example, "LastName = 'Smith' OR LastName = 'Jones'"). Use the following guidelines for criteria strings:

* FieldName must be a valid field name from the Recordset. If the field name contains spaces, you must enclose the name in square brackets.
* Operator must be one of the following: <, >, <=, >=, <>, =, or LIKE.
* Value is the value with which you will compare the field values (for example, 'Smith', #8/24/95#, 12.345, or $50.00). Use single quotes with strings and pound signs (#) with dates. For numbers, you can use decimal points, dollar signs, and scientific notation. If Operator is LIKE, Value can use wildcards. Only the asterisk (\*) and percent sign (%) wild cards are allowed, and they must be the last character in the string. Value cannot be null.

    * **Note**: To include single quotation marks (') in the filter Value, use two single quotation marks to represent one. For example, to filter on O'Malley, the criteria string should be "col1 = 'O''Malley'". To include single quotation marks at both the beginning and the end of the filter value, enclose the string with pound signs (#). For example, to filter on '1', the criteria string should be "col1 = #'1'#".

* There is no precedence between AND and OR. Clauses can be grouped within parentheses. However, you cannot group clauses joined by an OR and then join the group to another clause with an AND, like this:

    (LastName = 'Smith' OR LastName = 'Jones') AND FirstName = 'John'

    Instead, you would construct this filter as

    (LastName = 'Smith' AND FirstName = 'John') OR (LastName = 'Jones' AND FirstName = 'John')

* In a LIKE clause, you can use a wildcard at the beginning and end of the pattern (for example, LastName Like '*mit*'), or only at the end of the pattern (for example, LastName Like 'Smit*').

The filter constants make it easier to resolve individual record conflicts during batch update mode by allowing you to view, for example, only those records that were affected during the last **UpdateBatch** method call.

Setting the **Filter** property itself may fail because of a conflict with the underlying data (for example, a record has already been deleted by another user). In such a case, the provider returns warnings to the Errors collection but does not halt program execution. A run-time error occurs only if there are conflicts on all the requested records. Use the **Status** property to locate records with conflicts.

Setting the **Filter** property to a zero-length string ("") has the same effect as using the **adFilterNone** constant.

Whenever the **Filter** property is set, the current record position moves to the first record in the filtered subset of records in the Recordset. Similarly, when the Filter property is cleared, the current record position moves to the first record in the **Recordset**.

When a **Recordset** is filtered based on a field of some variant type (e.g., sql_variant), an error (DISP_E_TYPEMISMATCH or 80020005) will result if the subtypes of the field and filter values used in the criteria string do not match. For example, if a **Recordset** object (*lpRecordset*) contains a column (C) of the sql_variant type and a field of this column has been assigned a value of 1 of the I4 type, setting the criteria string of **Recordset** **Filter** "C='A'" on the field will produce the error at run time. However, **Recordset** **Filter** "C=2" applied on the same field will not produce any error although the field will be filtered out of the current record set.

See the **Bookmark** property for an explanation of bookmark values from which you can build an array to use with the **Filter** property.

Only Filters in the form of Criteria Strings (e.g. OrderDate > '12/31/1999') affect the contents of a persisted **Recordset**. Filters created with an Array of Bookmarks or using a value from the **FilterGroupEnum** will not affect the contents of the persisted Recordset. These rules apply to Recordsets created with either client-side or server-side cursors.

**Note**: When you apply the **adFilterPendingRecords** flag to a filtered and modified **Recordset** in the batch update mode, the resultant **Recordset** is empty if the filtering was based on the key field of a single-keyed table and the modification was made on the key field values. The resultant **Recordset** will be non-empty if one of the following is true:

* The filtering was based on non-key fields in a single-keyed table.
* The filtering was based on any fields in a multiple-keyed table.
* Modifications were made on non-key fields in a single-keyed table.
* Modifications were made on any fields in a multiple-keyed table.

The following table summarizes the effects of **adFilterPendingRecords** in different combinations of filtering and modifications. The left column shows the possible modifications; modifications can be made on any of the non-keyed fields, on the key field in a single-keyed table, or on any of the key fields in a multiple-keyed table. The top row shows the filtering criterion; filtering can be based on any of the non-keyed fields, the key field in a single-keyed table, or any of the key fields in a multiple-keyed table. The intersecting cells show the results: + means that applying adFilterPendingRecords results in a non-empty **Recordset**; - means an empty **Recordset**.

|     | Non Keys    | Single Key  | Multiple Keys |
| --- | ----------- | ----------- | ------------- |
| Non Keys | + | + | + |
| Single Key | + | - | N/A |
| Multiple Key | + | N/A | + |

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
DIM cvSource AS CVAR = "Publishers"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTableDirect)

' // Set the Filter property
pRecordset.Filter = "City = 'New York'"

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents of the "City" and "Name" columns
   DIM cvRes1 AS CVAR = pRecordset.Collect("City")
   DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
   PRINT cvRes1 & " " & cvRes2
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="Find"></a>Find

Searches a **Recordset** for the row that satisfies the specified criteria. Optionally, the direction of the search, starting row, and offset from the starting row may be specified. If the criteria is met, the current row position is set on the found record; otherwise, the position is set to the end (or start) of the **Recordset**.

**Note**: The **Find** method is a single column operation only because the OLE DB specification defines **IRowsetFind** this way.


```
FUNCTION Find (BYREF cbsCriteria AS CBSTR, _
   BYREF cvStart AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
   BYVAL SkipRecords AS LONG = 0, BYVAL SearchDirection AS SearchDirectionEnum = adSearchForward) AS HRESULT
```
```
FUNCTION Find ( BYREF cbsCriteria AS CBSTR, BYVAL SkipRecords AS LONG = 0, _
   BYVAL SearchDirection AS SearchDirectionEnum = adSearchForward) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsCriteria* | A string value that contains a statement specifying the column name, comparison operator, and value to use in the search. |
| *cvStart* | Optional. A bookmark that functions as the starting position for the search. |
| *SkipRecords* | Optional. A Long value, whose default value is zero, that specifies the row offset from the current row or start bookmark to begin the search. By default, the search will start on the current row. |
| *SearchDirection* | Optional. A **SearchDirectionEnum** value that specifies whether the search should begin on the current row or the next available row in the direction of the search. An unsuccessful search stops at the end of the **Recordset** if the value is **adSearchForward**. An unsuccessful search stops at the start of the **Recordset** if the value is **adSearchBackward**. |

#### SearchDirectionEnum

Specifies the direction of a record search within a Recordset.

| Constant   | Description |
| ---------- | ----------- |
| **adSearchBackward** | Searches backward, stopping at the beginning of the **Recordset**. If a match is not found, the record pointer is positioned at **BOF**. |
| **adSearchForward** | Searches forward, stopping at the end of the **Recordset**. If a match is not found, the record pointer is positioned at **EOF**. |


#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Only a single-column name may be specified in criteria. This method does not support multi-column searches.

The comparison operator in **Criteria** may be ">" (greater than), "<" (less than), "=" (equal), ">=" (greater than or equal), "<=" (less than or equal), "<>" (not equal), or "like" (pattern matching).

The value in **Criteria** may be a string, floating-point number, or date. String values are delimited with single quotes or "#" (number sign) marks (for example, "state = 'WA'" or "state = #WA#"). Date values are delimited with "#" (number sign) marks (for example, "start_date > #7/22/97#"). These values can contain hours, minutes, and seconds to indicate time stamps, but should not contain milliseconds or errors will occur.

If the comparison operator is "like", the string value may contain an asterisk (*) to find one or more occurrences of any character or substring. For example, "state like 'M*'" matches Maine and Massachusetts. You can also use leading and trailing asterisks to find a substring contained within the values. For example, "state like '*as*'" matches Alaska, Arkansas, and Massachusetts.

Asterisks can be used only at the end of a criteria string, or together at both the beginning and end of a criteria string, as shown above. You cannot use the asterisk as a leading wildcard ('*str'), or embedded wildcard ('s*r'). This will cause an error.

**Note**: An error will occur if a current row position is not set before calling **Find**. Any method that sets row position, such as **MoveFirst**, should be called before calling **Find**.

**Note**: If you call the **Find** method on a recordset, and the current position in the recordset is at the last record or end of file (EOF), you will not find anything. You need to call the MoveFirst method to set the current position/cursor to the beginning of the recordset.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Publishers ORDER By PubID"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

pRecordset.Find "PubID = #70#"
DIM cvRes1 AS CVAR = pRecordset.Collect("PubID")
DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
PRINT cvRes1 & " " & cvRes2
```
