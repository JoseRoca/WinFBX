# CADOConnection Class

Represents an open connection to a data source.

A **Connection** object represents a unique session with a data source. In the case of a client/server database system, it may be equivalent to an actual network connection to the server. Depending on the functionality supported by the provider, some collections, methods, or properties of a **Connection** object may not be available.

With the collections, methods, and properties of a **Connection** object, you can do the following:

* Configure the connection before opening it with the **ConnectionString**, **ConnectionTimeout**, and **Mode** properties. **ConnectionString** is the default property of the **Connection** object.
* Set the **CursorLocation** property to client to invoke the Microsoft Cursor Service for OLE DB, which supports batch updates.
* Set the default database for the connection with the **DefaultDatabase** property.
* Set the level of isolation for the transactions opened on the connection with the **IsolationLevel** property.
* Specify an OLE DB provider with the **Provider** property.
* Establish, and later break, the physical connection to the data source with the **Open** and **Close** methods.
* Execute a command on the connection with the **Execute** method and configure the execution with the **CommandTimeout** property.
* Manage transactions on the open connection, including nested transactions if the provider supports them, with the **BeginTrans**, **CommitTrans**, and **RollbackTrans** methods and the **Attributes** property.
* Examine errors returned from the data source with the **Errors** collection.
* Read the version from the ADO implementation used with the **Version** property.
* Obtain schema information about your database with the **OpenSchema** method.

You can create **Connection** objects independently of any other previously defined object.

To execute a query without using a **Command** object, pass a query string to the **Execute** method of a **Connection** object. However, a **Command** object is required when you want to persist the command text and re-execute it, or use query parameters.

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Attributes](#Attributes) | Indicates one or more characteristics of an object. |
| [BeginTrans](#BeginTrans) | Begins a new transaction. |
| [Cancel](#Cancel) | Cancels execution of a pending, asynchronous method call. |
| [Close](#Close) | Closes a **Connection** object and any dependent objects. |
| [CommandTimeout](#CommandTimeout) | Sets or returns a Long value that indicates, in seconds, how long to wait for a command to execute. Default is 30. |
| [CommitTrans](#CommitTrans) | Saves any changes and ends the current transaction. It may also start a new transaction. |
| [ConnectionString](#ConnectionString) | Indicates the information used to establish a connection to a data source. |
| [ConnectionTimeout](#ConnectionTimeout) | Indicates how long to wait while establishing a connection before terminating the attempt and generating an error. |
| [CursorLocation](#CursorLocation) | Indicates the location of the cursor service. |
| [DefaultDatabase](#DefaultDatabase) | Indicates the default database for a **Connection** object. |
| [Execute](#Execute) | Executes the specified query, SQL statement, stored procedure, or provider-specific text. |
| [GetErrorInfo](#GetErrorInfo) | Returns information about ADO errors. |
| [IsolationLevel](#IsolationLevel) | Indicates the level of isolation for a **Connection** object. |
| [Mode](#Mode) | Indicates the available permissions for modifying data in a **Connection** object. |
| [Open](#Open) | Opens a connection to a data source. |
| [OpenSchema](#OpenSchema) | Obtains database schema information from the provider. |
| [Properties](#Properties) | Returns a reference to the **Properties** collection. |
| [Provider](#Provider) | Indicates the name of the provider for a **Connection** object. |
| [RollbackTrans](#RollbackTrans) | Saves any changes and ends the current transaction. It may also start a new transaction. |
| [State](#State) | Indicates if a **Connection** is open or closed. |
| [Version](#Version) | Indicates the ADO version number. |
