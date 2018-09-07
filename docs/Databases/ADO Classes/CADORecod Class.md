# CADORecord Class

Represents a row from a **Recordset** or the data provider, or an object returned by a semi-structured data provider, such as a file or directory.

**Include file**: CADORecord.inc (include CADODB.inc)

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
