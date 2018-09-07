# CADOCommand Class

Defines a specific command that you intend to execute against a data source.

**Include file**: CADOCommand.inc (include CADODB.inc)

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [ActiveConnection](#ActiveConnection) | Determines the **Connection** object over which the specified **Command** object will execute. |
| [Cancel](#Cancel) | Cancels execution of a pending, asynchronous method call. |
| [CommandStream](#CommandStream) | Sets or returns the stream used as the input for a **Command** object. |
| [CommandText](#CommandText) | Sets or returns an string value that contains a provider command, such as an SQL statement, a table name, a relative URL, or a stored procedure call. |
| [CommandTimeout](#CommandTimeout) | Sets or returns a Long value that indicates, in seconds, how long to wait for a command to execute. |
| [CommandType](#CommandType) | Sets or returns one or more **CommandTypeEnum** values. |
| [CreateParameter](#CreateParameter) | Creates a new **Paramete**r object with the specified properties. |
| [Dialect](#Dialect) | Indicates the dialect of the **CommandText** or **CommandStream** properties. |
| [Execute](#Execute) | Executes the query, SQL statement, or stored procedure specified in the **CommandText** or **CommandStream** property. |
| [GetErrorInfo](#GetErrorInfo) | Returns information about ADO errors. |
| [Name](#Name) | Sets or returns an string value that indicates the name of a **Command** object. |
| [NamedParameters](#NamedParameters) | Indicates whether parameter names should be passed to the provider. |
| [Parameters](#Parameters) | Returns a reference to the **Parameters** collection. |
| [Prepared](#Prepared) | Sets or returns a Boolean value that, if set to True, indicates that the command should be prepared. |
| [Properties](#Properties) | Returns a reference to the **Properties** collection. |
| [State](#State) | Indicates if the **Command** object is open or closed. |

#### Remarks

Use a **Command** object to query a database and return records in a **Recordset** object, to execute a bulk operation, or to manipulate the structure of a database. Depending on the functionality of the provider, some **Command** collections, methods, or properties may generate an error when referenced.

With the collections, methods, and properties of a Command object, you can do the following:

* Define the executable text of the command (for example, an SQL statement) with the **CommandText** property. Alternatively, for command or query structures other than simple strings (for example, XML template queries) define the command with the **CommandStream** property.
* Optionally, indicate the command dialect used in the **CommandText** or **CommandStream** with the **Dialect** property.
* Define parameterized queries or stored-procedure arguments with **Parameter** objects and the **Parameters** collection.
* Indicate whether parameter names should be passed to the provider with the **NamedParameters** property.
* Execute a command and return a **Recordset** object if appropriate with the **Execute** method.
* Specify the type of command with the **CommandType** property prior to execution to optimize performance.
* Control whether the provider saves a prepared (or compiled) version of the command prior to execution with the **Prepared** property.
* Set the number of seconds that a provider will wait for a command to execute with the **CommandTimeout** property.
* Associate an open connection with a **Command** object by setting its **ActiveConnection** property.
* Set the Name property to identify the **Command** object as a method on the associated **Connection** object.
* Pass a **Command** object to the **Source** property of a **Recordset** in order to obtain data.
* Access provider-specific attributes with the **Properties** collection.

**Note**: To execute a query without using a **Command** object, pass a query string to the Execute method of a **Connection** object or to the **Open** method of a **Recordset** object. However, a **Command** object is required when you want to persist the command text and re-execute it, or use query parameters.

To create a **Command** object independently of a previously defined **Connection** object, set its **ActiveConnection** property to a valid connection string. ADO still creates a **Connection** object, but it doesn't assign that object to an object variable. However, if you are associating multiple **Command** objects with the same connection, you should explicitly create and open a **Connection** object; this assigns the **Connection** object to an object variable. Make sure the **Connection** object was opened successfully before you assign it to the **Command** object's **ActiveConnection** property, because assigning a closed **Connection** object causes an error. If you do not set the **Command** object's **ActiveConnection** property to this object variable, ADO creates a new **Connection** object for each **Command** object, even if you use the same connection string.

To execute a **Command**, simply call it by its **Name** property on the associated **Connection** object. The **Command** must have its **ActiveConnection** property set to the **Connection** object. If the **Command** has parameters, pass their values as arguments to the method.

If two or more **Command** objects are executed on the same connection and either **Command** object is a stored procedure with output parameters, an error occurs. To execute each **Command** object, use separate connections or disconnect all other **Command** objects from the connection.

# CADOParameters Class

Contains all the **Parameter** objects of a **Command** object.

**Include file**: CADOParameters.inc (include CADODB.inc)

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Append](#Append) | Appends an object to the **Parameters** collection. |
| [Count](#Count) | Retrieves the number of objects of the **Parameters** collection. |
| [Delete_](#Delete_) | Deletes an object from the **Parameters** collection. |
| [Item](#Item) | Indicates a specific member of the **Parameters** collection, by name or ordinal number. |
| [Refresh](#Refresh) | Refreshes the contents of the **Parameters** collection. |

#### Remarks

A **Command** object has a **Parameters** collection made up of **Parameter** objects.

Using the **Refresh** method on a **Command** object's **Parameters** collection retrieves provider parameter information for the stored procedure or parameterized query specified in the **Command** object. Some providers do not support stored procedure calls or parameterized queries; calling the **Refresh** method on the **Parameters** collection when using such a provider will return an error.

If you have not defined your own **Parameter** objects and you access the **Parameters** collection before calling the **Refresh** method, ADO will automatically call the method and populate the collection for you.

You can minimize calls to the provider to improve performance if you know the properties of the parameters associated with the stored procedure or parameterized query you wish to call. Use the **CreateParameter** method to create **Parameter** objects with the appropriate property settings and use the Append method to add them to the **Parameters** collection. This lets you set and return parameter values without having to call the provider for the parameter information. If you are writing to a provider that does not supply parameter information, you must manually populate the **Parameters** collection using this method to be able to use parameters at all. Use the **Delete_** method to remove **Parameter** objects from the **Parameters** collection if necessary.

The objects in the **Parameters** collection of a **Recordset** go out of scope (therefore becoming unavailable) when the **Recordset** is closed.
