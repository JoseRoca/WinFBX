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

