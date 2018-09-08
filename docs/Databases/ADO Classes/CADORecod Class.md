# CADORecord Class

Represents a row from a **Recordset** or the data provider, or an object returned by a semi-structured data provider, such as a file or directory.

**Include file**: CADORecord.inc (include CADODB.inc)

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [ActiveConnection](#ActiveConnection) | Indicates to which **Connection** object the specified **Record** object currently belongs. |
| [Attach](#Attach) | Attaches a reference to an ADO **Record** object to the class. |
| [Cancel](#Cancel) | Cancels execution of a pending, asynchronous method call. |
| [Close](#Close) | Closes a **Record** object and any dependent objects. |
| [CopyRecord](#CopyRecord) | Copies a entity represented by a Record to another location. |
| [DeleteRecord](#DeleteRecord) | Deletes a the entity represented by a **Record**. |
| [Fields](#Fields) | Gets a reference to the Fields collection of a **Record** object. |
| [GetChildren](#GetChildren) | Returns a **Recordset** whose rows represent the children of a collection **Record**. |
| [GetErrorInfo](#GetErrorInfo) | Returns information about ADO errors. |
| [Mode](#Mode) | Sets or returns a **ConnectModeEnum** value. |
| [MoveRecord](#MoveRecord) | Moves a entity represented by a **Record** to another location. |
| [Open](#Open) | Opens an existing **Record** object, or creates a new item represented by the **Record** (such as a file or directory). |
| [ParentURL](#ParentURL) | Indicates an absolute URL string that points to the parent **Record** of the current **Record** object. |
| [Properties](#Properties) | Returns a reference to the **Properties** collection. |
| [RecordType](#RecordType) | Indicates the type of **Record** object. |
| [Source](#Source) | Indicates the data source or object represented by the **Record**. |
| [State](#State) | Indicates for all applicable objects whether the state of the object is open or closed. |

#### Remarks

A **Record** object represents one row of data, and has some conceptual similarities with a one-row **Recordset**. Depending upon the capabilities of your provider, **Record** objects may be returned directly from your provider instead of a one-row **Recordset**, for example when an SQL query that selects only one row is executed. Or, a **Record** object can be obtained directly from a **Recordset** object. Or, a **Record** can be returned directly from a provider to semi-structured data, such as the Microsoft Exchange OLE DB provider.

You can view the fields associated with the **Record** object by way of the **Fields** collection on the **Record** object. ADO allows object-valued columns including **Recordset**, **SafeArray*, and scalar values in the **Fields** collection of **Record** objects.

If the **Record** object represents a row in a **Recordset**, then it is possible to return to that original **Recordset** with the **Source** property.

The **Record** object can also be used by semi-structured data providers such as the Microsoft OLE DB Provider for Internet Publishing, to model tree-structured namespaces. Each node in the tree is a **Record** object with associated columns. The columns can represent the attributes of that node and other relevant information. The **Record** object can represent both a leaf node and a non-leaf node in the tree structure. Non-leaf nodes have other nodes as their contents while leaf nodes do not have such contents. Leaf nodes typically contain binary streams of data while non-leaf nodes may also have a default binary stream associated with them. **Properties** on the **Record** object identify the type of node.

The **Record** object also represents an alternative way for navigating hierarchically organized data. A **Record** object may be created to represent the root of a specific sub-tree in a large tree structure and new **Record** objects may be opened to represent child nodes.

A resource (for example, a file or directory) can be uniquely identified by an absolute URL. A **Connection** object is implicitly created and set to the **Record** object when the **Record** is opened with an absolute URL. A **Connection** object may explicitly be set to the **Record** object via the **ActiveConnection** property. The files and directories accessible via the **Connection** object define the context in which **Record** operations may occur.

Data modification and navigation methods on the **Record** object also accept a relative URL, which locates a resource using an absolute URL or the **Connection** object context as a starting point.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

A **Connection** object is associated with each **Record** object. Therefore, **Record** object operations can be part of a transaction by invoking **Connection** object transaction methods.

The **Record** object does not support ADO events, and therefore will not respond to notifications.

With the methods and properties of a **Record** object, you can do the following:

* Set or return the associated Connection object with the ActiveConnection property.
* Indicate access permissions with the Mode property.
* Return the URL of the directory, if any, that contains the resource represented by the Record with the ParentURL property.
* Indicate the absolute URL, relative URL, or Recordset from which the Record is derived with the Source property.
* Indicate the current status of the Record with the State property.
* Indicate the type of Record—simple, collection, or structured document—with the RecordType property.
* Halt execution of an asynchronous operation with the Cancel method.
* Disassociate the Record from a data source with the Close method.
* Copy the file or directory represented by a Record to another location with the CopyRecord method.
* Delete the file, or directory and subdirectories, represented by a Record with the DeleteRecord method.
* Open a Recordset containing rows that represent the subdirectories and files of the entity represented by the Record with the GetChildren method.
* Move (rename) the file, or directory and subdirectories, represented by a Record to another location with the MoveRecord method.
* Associate the Record with an existing data source, or create a new file or directory with the Open method.

#### Fields

A **Record** object has two special fields that can be indexed with FieldEnum constants. One constant accesses a field containing the default stream for the **Record**, and the other accesses a field containing the absolute URL string for the **Record**.

Certain providers (for example, the Microsoft OLE DB Provider for Internet Publishing) may populate the **Fields** collection with a subset of available fields for the **Record** or **Recordset**. Other fields will not be added to the collection until they are first referenced by name or indexed by your code.

Each **Field** object corresponds to a column in the **Recordset**. You use the **Value** property of **Field** objects to set or return data for the current record. Depending on the functionality the provider exposes, some collections, methods, or properties of a **Field** object may not be available.

With the collections, methods, and properties of a Field object, you can do the following:

* Return the name of a field with the **Name** property.
* View or change the data in the field with the **Value** property. **Valu**e is the default property of the **Field** object.
* Return the basic characteristics of a field with the **Type_**, **Precision**, and **NumericScale** properties.
* Return the declared size of a field with the **DefinedSize** property.
* Return the actual size of the data in a given field with the **ActualSize** property.
* Determine what types of functionality are supported for a given field with the **Attributes** property and **Properties** collection.
* Manipulate the values of fields containing long binary or long character data with the **AppendChunk** and **GetChunk** methods.
* If the provider supports batch updates, resolve discrepancies in field values during batch updating with the **OriginalValue** and **UnderlyingValue** properties.

All of the metadata properties (**Name**, **Type_**, **DefinedSize**, **Precision**, and **NumericScale**) are available before opening the **Field** object's **Recordset**. Setting them at that time is useful for dynamically constructing forms.

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

Sets or returns a string value that contains a definition for a connection if the connection is closed, or a Variant containing the current **Connection** object if the connection is open. Default is a null object reference. See the **ConnectionString** property in the documentation for the **CADOConnection** class.

This property is read/write when the **Record** object is closed, and may contain a connection string or reference to an open **Connection** object. This property is read-only when the **Record** object is open, and contains a reference to an open **Connection** object.

A **Connection** object is created implicitly when the **Record** object is opened from a URL. Open the **Record** with an existing, open **Connection** object by assigning the Connection object to this property, or using the **Connection** object as a parameter in the **Open** method call. If the **Record** is opened from an existing **Record** or **Recordset**, then it is automatically associated with that **Record** or **Recordset** object's **Connection** object.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

# <a name="Attach"></a>Attach

Attaches a record to the class.

```
FUNCTION Attach (BYVAL pRecordset AS Afx_ADORecord PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pRecordset* | A reference to an **Afx_ADORecord** object. |
| *fAddRef* | TRUE = increase the reference count; FALSE = don't increase the reference count. |

# <a name="Cancel"></a>Cancel

Cancels execution of a pending, asynchronous method call.

```
FUNCTION Cancel () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Cancel** method to terminate execution of an asynchronous method call (that is, a method invoked with the **adAsyncConnect**, **adAsyncExecute**, or **adAsyncFetch** option).

For a **Record** object, the last asynchronous call to the **CopyRecord**, **DeleteRecord**, **MoveRecord** or **Open** methods is terminated.
