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

# <a name="Attributes"></a>Attributes

Indicates one or more characteristics of an object.

```
PROPERTY Attributes () AS LONG
PROPERTY Attributes (BYVAL lAttr AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lAttr* | The value can be the sum of one or more **XactAttributeEnum** values. The default is zero (0). |

#### Return value

#### XactAttributeEnum

Specifies the transaction attributes of a Connection object.

| Constant   | Description |
| ---------- | ----------- |
| **adXactAbortRetaining** | Performs retaining aborts—that is, calling **RollbackTrans** automatically starts a new transaction. Not all providers support this. |
| **adXactCommitRetaining** | Performs retaining commits—that is, calling **CommitTrans** automatically starts a new transaction. Not all providers support this. |

#### Remarks

Use the **Attributes** property to set or return characteristics of **Connection** objects.

When you set multiple attributes, you can sum the appropriate constants. If you set the property value to a sum including incompatible constants, an error occurs.

#### Remote Data Service Usage

This property is not available on a client-side **Connection** object.


# <a name="BeginTrans"></a>BeginTrans

Begins a new transaction.

```
FUNCTION BeginTrans () AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *lAttr* | The value can be the sum of one or more **XactAttributeEnum** values. The default is zero (0). |

#### Return value

The nesting level of the transaction.

#### Remarks

Use **BeginTrans**, **CommitTrans** and **RollbackTrans** with a **Connection** object when you want to save or cancel a series of changes made to the source data as a single unit. For example, to transfer money between accounts, you subtract an amount from one and add the same amount to the other. If either update fails, the accounts no longer balance. Making these changes within an open transaction ensures that either all or none of the changes go through.

Not all providers support transactions. Verify that the provider-defined property "Transaction DDL" appears in the **Connectio** object's **Properties** collection, indicating that the provider supports transactions. If the provider does not support transactions, calling one of these methods will return an error.

After you call the **BeginTrans** method, the provider will no longer instantaneously commit changes you make until you call **CommitTrans** or **RollbackTrans** to end the transaction.

For providers that support nested transactions, calling the **BeginTrans** method within an open transaction starts a new, nested transaction. The return value indicates the level of nesting: a return value of "1" indicates you have opened a top-level transaction (that is, the transaction is not nested within another transaction), "2" indicates that you have opened a second-level transaction (a transaction nested within a top-level transaction), and so forth. Calling **CommitTrans** or **RollbackTrans** affects only the most recently opened transaction; you must close or roll back the current transaction before you can resolve any higher-level transactions.

Calling the **CommitTrans** method saves changes made within an open transaction on the connection and ends the transaction. Calling the **RollbackTrans** method reverses any changes made within an open transaction and ends the transaction. Calling either method when there is no open transaction generates an error.

Depending on the **Connection** object's **Attributes** property, calling either the **CommitTrans** or **RollbackTrans** methods may automatically start a new transaction. If the **Attributes** property is set to **adXactCommitRetaining**, the provider automatically starts a new transaction after a **CommitTrans** call. If the **Attributes** property is set to **adXactAbortRetaining**, the provider automatically starts a new transaction after a **RollbackTrans** call.

#### Remote Data Service

The **BeginTrans**, **CommitTrans**, and **RollbackTrans** methods are not available on a client-side **Connection** object.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Authors"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Begin a transaction
pConnection.BeginTrans

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
   DIM cvRes AS CVAR = pRecordset.Collect("Year Born")
   IF cvRes.ValInt = 1947 THEN pRecordset.Collect("Year Born") = 1900
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
' // Commit the transaction
'pConnection.CommitTrans
' // Rollback the transaction because this is a demo
pConnection.RollbackTrans
```

# <a name="Cancel"></a>Cancel

Cancels execution of a pending, asynchronous method call.

```
FUNCTION Cancel () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Cancel** method to terminate execution of an asynchronous method call (that is, a method invoked with the **adAsyncConnect**, **adAsyncExecute**, or **adAsyncFetch** option).

For a **Connection** object, the last asynchronous call to the **Execute** or **Open** methods is terminated.

# <a name="Close"></a>Close

Closes a **Connection** object and any dependent objects.

```
FUNCTION Close () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Close** method to close a **Connection** to free any associated system resources. Closing an object does not remove it from memory; you can change its property settings and open it again later. To completely eliminate an object from memory, release the connection calling the **Release** method of the interface.

#### Usage example

```
IF pConnection.State = adStateOpen THEN pConnection.Close
```

# <a name="CommandTimeout"></a>CommandTimeout

Sets or returns a Long value that indicates, in seconds, how long to wait for a command to execute. Default is 30.

```
PROPERTY CommandTimeout () AS LONG
PROPERTY CommandTimeout (BYVAL lTimeout AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lTimeout* | Time to wait for a command to execute. |

#### Return value

The timeout value.

#### Remarks

Use the **CommandTimeout** property on a **Connection** object to allow the cancellation of an **Execute** method call, due to delays from network traffic or heavy server use. If the interval set in the **CommandTimeout** property elapses before the command completes execution, an error occurs and ADO cancels the command. If you set the property to zero, ADO will wait indefinitely until the execution is complete. Make sure the provider and data source to which you are writing code support the **CommandTimeout** functionality.

The **CommandTimeout** setting on a **Connection** object has no effect on the **CommandTimeout** setting on a **Command** object on the same **Connection**; that is, the **Command** object's **CommandTimeout** property does not inherit the value of the **Connection** object's **CommandTimeout** value.

On a **Connection** object, the **CommandTimeout** property remains read/write after the **Connection** is opened.

#### Examples

```
' // Sets the timeout to 35 seconds
pConnection.CommandTimeout = 35
```
```
' // Gets the timeout value
DIM lTimeout AS LONG = pConnection.CommandTimeout
```

# <a name="CommitTrans"></a>CommitTrans

Saves any changes and ends the current transaction. It may also start a new transaction.

```
FUNCTION CommitTrans () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use **BeginTrans**, **CommitTrans** and **RollbackTrans** with a **Connection** object when you want to save or cancel a series of changes made to the source data as a single unit. For example, to transfer money between accounts, you subtract an amount from one and add the same amount to the other. If either update fails, the accounts no longer balance. Making these changes within an open transaction ensures that either all or none of the changes go through.

Not all providers support transactions. Verify that the provider-defined property "Transaction DDL" appears in the **Connection** object's **Properties** collection, indicating that the provider supports transactions. If the provider does not support transactions, calling one of these methods will return an error.

After you call the **BeginTrans method**, the provider will no longer instantaneously commit changes you make until you call **CommitTrans** or **RollbackTrans** to end the transaction.

For providers that support nested transactions, calling the **BeginTrans** method within an open transaction starts a new, nested transaction. The return value indicates the level of nesting: a return value of "1" indicates you have opened a top-level transaction (that is, the transaction is not nested within another transaction), "2" indicates that you have opened a second-level transaction (a transaction nested within a top-level transaction), and so forth. Calling **CommitTrans** or **RollbackTrans** affects only the most recently opened transaction; you must close or roll back the current transaction before you can resolve any higher-level transactions.

Calling the **CommitTrans** method saves changes made within an open transaction on the connection and ends the transaction. Calling the **RollbackTrans** method reverses any changes made within an open transaction and ends the transaction. Calling either method when there is no open transaction generates an error.

Depending on the **Connection** object's **Attributes** property, calling either the **CommitTrans** or **RollbackTrans** methods may automatically start a new transaction. If the **Attributes** property is set to **adXactCommitRetaining**, the provider automatically starts a new transaction after a **CommitTrans** call. If the **Attributes** property is set to **adXactAbortRetaining**, the provider automatically starts a new transaction after a **RollbackTrans** call.

#### Remote Data Service

The **BeginTrans**, **CommitTrans**, and **RollbackTrans** methods are not available on a client-side **Connection** object.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Authors"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Begin a transaction
pConnection.BeginTrans

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
   DIM cvRes AS CVAR = pRecordset.Collect("Year Born")
   IF cvRes.ValInr = 1947 THEN pRecordset.Collect("Year Born") = 1900
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
' // Commit the transaction
'pConnection.CommitTrans
' // Rollback the transaction because this is a demo
pConnection.RollbackTrans
```

# <a name="ConnectionString"></a>ConnectionString

Indicates the information used to establish a connection to a data source.

```
PROPERTY ConnectionString () AS CBSTR
PROPERTY ConnectionString (BYREF cbsConConStr AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsConConStr* | The connection string. |

#### Return value

The connection string.

#### Remarks

Use the **ConnectionString** property to specify a data source by passing a detailed connection string containing a series of argument = value statements separated by semicolons.

ADO supports five arguments for the **ConnectionString** property; any other arguments pass directly to the provider without any processing by ADO. The arguments ADO supports are as follows.

| Argument   | Description |
| ---------- | ----------- |
| **Provider=** | Specifies the name of a provider to use for the connection. |
| **File Name=** | Specifies the name of a provider-specific file (for example, a persisted data source object) containing preset connection information. |
| **Remote Provider=** | Specifies the name of a provider to use when opening a client-side connection. (Remote Data Service only.) |
| **Remote Server=** | Specifies the path name of the server to use when opening a client-side connection. (Remote Data Service only.) |
| **URL=** | Specifies the connection string as an absolute URL identifying a resource, such as a file or directory. |

After you set the **ConnectionString** property and open the **Connection** object, the provider may alter the contents of the property, for example, by mapping the ADO-defined argument names to their provider equivalents.

The **ConnectionString** property automatically inherits the value used for the **ConnectionString** argument of the Open method, so you can override the current **ConnectionString** property during the **Open** method call.

Because the **File Name** argument causes ADO to load the associated provider, you cannot pass both the **Provider** and **File Name** arguments.

The **ConnectionString** property is read/write when the connection is closed and read-only when it is open.

Duplicates of an argument in the **ConnectionString** property are ignored. The last instance of any argument is used.

#### Remote Data Service Usage

When used on a client-side **Connection** object, the **ConnectionString** property can include only the **Remote Provider** and **Remote Server** parameters.

#### Usage examples

```
pConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open
```
```
DIM cbsConStr AS CBSTR = pConnection.ConnectionString
```

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cbsSource AS CBSTR = "SELECT * FROM Authors"
pRecordset.Open(cbsSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

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

# <a name="ConnectionTimeout"></a>ConnectionTimeout

Indicates how long to wait while establishing a connection before terminating the attempt and generating an error.

```
PROPERTY ConnectionTimeout () AS LONG
PROPERTY ConnectionTimeout (BYVAL lTimeout AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lTimeout* | Time to wait for a command to execute. |

#### Return value

The timeout value.

#### Remarks

Use the **ConnectionTimeout** property on a **Connection** object if delays from network traffic or heavy server use make it necessary to abandon a connection attempt. If the time from the **ConnectionTimeout** property setting elapses prior to the opening of the connection, an error occurs and ADO cancels the attempt. If you set the property to zero, ADO will wait indefinitely until the connection is opened. Make sure the provider to which you are writing code supports the **ConnectionTimeout** functionality.

The **ConnectionTimeout** property is read/write when the connection is closed and read-only when it is open.

On a **Connection** object, the **ConnectionTimeout** property remains read/write after the **Connection** is opened.