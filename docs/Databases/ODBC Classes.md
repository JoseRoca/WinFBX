# ODBC Classes

**CODBC** is a wrapper class on top of ODBC. 

The main purpose of the class if to ease the use of the ODBC API, that is notoriously difficult to use.

**Include files**: CODBC.inc, CODBCDb.inc, CODBCStmt.inc.

You create an instance of the class using:

```
' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
```

You create an statement object creating an instance of the CODBCStmt object passing the connection object pointer as the parameter:

' // Allocate an statement object
DIM pStmt AS COdbcStmt = pDbc

When the class is destroyed, its Destructor method closes the cursor and the connection and frees the allocated connection handle. Therefore, it is not needed to explicitly close the database and free handles. The class keeps track of the number of active connections and frees the environment handle if there are no active connections.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = pDbc

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Parse the result set
DIM cwsOutput AS CWSTR
DO
   ' // Fetch the record
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ' // Get the values of the columns and display them
   cwsOutput = ""
   cwsOutput += pStmt.GetData(1) & " "
   cwsOutput += pStmt.GetData(2) & " "
   cwsOutput += pStmt.GetData(3)
   PRINT cwsOutput
LOOP
```

#### Example using the pointer notation

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC PTR = NEW CODBC(wszConStr)
IF pDbc = NULL THEN END

' // Allocate an statement object
DIM pStmt AS COdbcStmt PTR = NEW COdbcStmt(pDbc)
IF pStmt = NULL THEN END

' // Generate a result sets
pStmt->ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Parse the result set
DIM cwsOutput AS CWSTR
DO
   ' // Fetch the record
   IF pStmt->Fetch = FALSE THEN EXIT DO
   ' // Get the values of the columns and display them
   cwsOutput = ""
   cwsOutput += pStmt->GetLongVarcharData(1) & " "
   cwsOutput += pStmt->GetData(2) & " "
   cwsOutput += pStmt->GetData(3)
   PRINT cwsOutput
LOOP

' // Delete the statement object
Delete pStmt

' // Delete the connection object
Delete pDbc
```

# CODBCBase Class

Base class for all the ODBC classes. Implements some common methods that all the other interfaces inherit. You don't have to instantiate this class.

**Include file**: CODBC.inc

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Error](#Error) | Returns TRUE if there has been an error; FALSE, otherwise. |
| [GetCPMatch](#GetCPMatch) | Returns a 32-bit SQLUINTEGER value that determines how a connection is chosen from a connection pool. |
| [GetDataSources](#GetDataSources) | Lists available DSN / Drivers installed. |
| [GetDrivers](#GetDrivers) | Lists driver descriptions and driver attribute keywords. |
| [GetEnvAttr](#GetEnvAttr) | Returns the current setting of an environment attribute. |
| [GetErrorInfo](#GetErrorInfo) | Returns a verbose description of the last error(s). |
| [GetLastResult](#GetLastResult) | Returns the last result code. |
| [GetOutputNTS](#GetOutputNTS) | Returns a 32-bit integer that determines how the driver returns string data. |
| [GetSqlState](#GetSqlState) | Returns the SqlState for the specified handle. |
| [ODBCVersion](#ODBCVersion) | Returns a 32-bit integer that determines whether certain functionality exhibits ODBC 2.x behavior or ODBC 3.x behavior. |
| [SetCPMatch](#SetCPMatch) | Returns a 32-bit SQLUINTEGER value that determines how a connection is chosen from a connection pool. |
| [SetEnvAttr](#SetEnvAttr) | Sets attributes that govern aspects of environments. |
| [SetErrorProc](#SetErrorProc) | Sets the address of an application defined error callback. |
| [SetOutputNTS](#SetOutputNTS) | Returns a 32-bit integer that determines how the driver returns string data. |

# CODBC Class

Implements methods to create and manage connection objects. Inherits from COdbcBase.

**Include fole**: CODBCDbc.inc

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsDb) | Requests a commit operation for all active operations on all statements associated with an environment. |
| [CommitTran](#CommitTran) | Requests a commit operation for all active operations on all statements associated with an environment. |
| [EnvHandle](#EnvHandle) | Returns the environment handle. |
| [Functions](#Functions) | Returns information about whether a driver supports a specific ODBC function. |
| [GetConnectAttr](#GetConnectAttr) | Returns the current setting of a connection attribute. |
| [GetDiagField](#GetDiagField) | Returns the current value of a field of a record of the diagnostic data structure (associated with an environment handle) that contains error, warning, and status information. |
| [GetDiagRec](#GetDiagRec) | Returns the current values of multiple fields of a diagnostic record that contains error, warning, and status information. |
| [GetErrorInfo](#GetErrorInfo) | Returns a verbose description of the last errors. |
| [GetInfo](#GetInfo) | Returns general information about the driver and data source associated with a connection. |
| [GetSqlState](#GetSqlState) | Returns the SqlState for the connection handle. |
| [Handle](#Handle) | Returns the connection handle. |
| [NativeSql](#NativeSql) | Returns the SQL string as modified by the driver. NativeSql does not execute the SQL statement. |
| [RollbackTran](#RollbackTran) | RollbackTran requests a rollback operation for all active operations on all statements associated with an environment. |
| [SetConnectAttr](#SetConnectAttr) | Sets attributes that govern aspects of connections. |
| [Supports](#Supports) | Returns information about whether a driver supports a specific ODBC function. It is an alias for **Functions**. |

# CODBCStmt Class

Implements methods to create and manage statement objects. Inherits from COdbcBase.

**Include fole**: CODBCStmt.inc

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsStmt) | Allocates an statement handle. |
| [AddRecord](#AddRecord) | Adds a record to the database. |
| [BindCol](#BindCol) | Binds application data buffers to columns in the result set. |
| [BindParameter](#BindParameter) | Binds a buffer to a parameter marker in an SQL statement. |
| [BulkOperations](#BulkOperations) | Performs bulk insertions and bulk bookmark operations. |
| [Cancel](#Cancel) | Cancels the processing on a statement. |
| [CloseCursor](#CloseCursor) | Closes a cursor that has been opened on a statement and discards pending results. |
| [ColAttribute](#ColAttribute) | Returns descriptor information for a column in a result set. |
| [ColAutoUniqueValue](#ColAutoUniqueValue) | Checks if the column is an autoincrementing column or not. |
| [ColBaseColumnName](#ColBaseColumnName) | Returns the base column name for the result set column. |
| [ColBaseTableName](#ColBaseTableName) | Returns the name of the base table that contains the column. |
| [ColCaseSensitive](#ColCaseSensitive) | Checks if the column is treated as case-sensitive for collations and comparisons. |
| [ColCatalogName](#ColCatalogName) | Returns the catalog of the table that contains the column. |
| [ColConciseType](#ColConciseType) | Returns the concise data type. |
| [ColCount](#ColCount) | Returns the number of columns available in the result set. |
| [ColDisplaySize](#ColDisplaySize) | Returns the maximum number of characters required to display data from the column. |
| [ColFixedPrecScale](#ColFixedPrecScale) | Checks if column has a fixed precision and nonzero scale that are data source-specific. |
| [ColIsNull](#ColIsNull) | Checks if the column is null. |
| [ColLabel](#ColLabel) | Returns the column label or title. |
| [ColLength](#ColLength) | Returns the maximum or actual character length of a character string or binary data type. |
| [ColLiteralPrefix](#ColLiteralPrefix) | Returns the character or characters that the driver recognizes as a prefix for a literal of this data type. |
| [ColLiteralSuffix](#ColLiteralSuffix) | Returns the character or characters that the driver recognizes as a suffix for a literal of this data type. |
| [ColLocalTypeName](#ColLocalTypeName) | Returns the localized (native language) name for the data type that may be different from the regular name of the data type. |
| [ColName](#ColName) | Returns the column alias, if it applies. |
| [ColNullable](#ColNullable) | Checks if the column can have NULL values. |
| [ColNumPrecRadix](#ColNumPrecRadix) | Returns the column numeric precision radix. |
| [ColOctetLength](#ColOctetLength) | The length, in bytes, of a character string or binary data type. |
| [ColPrecision](#ColPrecision) | A numeric value that for a numeric data type denotes the applicable precision. |
| [ColScale](#ColScale) | A numeric value that is the applicable scale for a numeric data type. |
| [ColSchemaName](#ColSchemaName) | The schema of the table that contains the column. |
| [ColSearchable](#ColSearchable) | Indicates if the column data type is searchable. |
| [ColTableName](#ColTableName) | The name of the table that contains the column. |
| [ColType](#ColType) | A numeric value that specifies the SQL data type. |
| [ColTypeName](#ColTypeName) | Data source-dependent data type name. |
| [ColUnnamed](#ColUnnamed) | SQL_NAMED or SQL_UNNAMED. If the SQL_DESC_NAME field of the IRD contains a column alias or a column name, SQL_NAMED is returned. |
| [ColUnsigned](#ColUnsigned) | SQL_TRUE if the column is unsigned (or not numeric). SQL_FALSE if the column is signed. |
| [ColUpdatable](#ColUpdatable) | SQL_TRUE if the column is updatable; SQL_FALSE otherwise. |
| [DbcHandle](#DbcHandle) | Returns the connection handle. |
| [DeleteByBookmark](#DeleteByBookmark) | Deletes a set of rows where each row is identified by a bookmark. |
| [DeleteRecord](#DeleteRecord) | Deletes the sepcified row of data. |
| [DescribeCol](#DescribeCol) | Returns the result descriptor for one column in the result set. |
| [DescribeParam](#DescribeParam) | Returns the description of a parameter marker associated with a prepared SQL statement. |
| [ExecDirect](#ExecDirect) | Executes a preparable statement. |
| [Execute](#Execute) | Executes a prepared statement. |
| [ExtendedFetch](#ExtendedFetch) | Fetches the specified rowset of data from the result set and returns data for all bound columns. |
| [Fetch](#Fetch) | Fetches the next rowset of data from the result set and returns data for all bound columns. |
| [FetchByBookmark](#FetchByBookmark) | Fetches a set of rows where each row is identified by a bookmark. |
| [FetchScroll](#FetchScroll) | Fetches the specified rowset of data from the result set and returns data for all bound columns. |
| [GetColumnPrivileges](#GetColumnPrivileges) | Returns a list of columns and associated privileges for the specified table. |
| [GetColumns](#GetColumns) | Returns the list of column names in specified tables. |
| [GetCursorConcurrency](#GetCursorConcurrency) | Gets a SQLUINTEGER value that specifies the cursor concurrency. |
| [GetCursorKeysetSize](#GetCursorKeysetSize) | Gets a SQLUINTEGER value that specifies the number of rows in the keyset-driven cursor. |
| [GetCursorName](#GetCursorName) | Returns the cursor name associated with a specified statement. |
| [GetCursorScrollability](#GetCursorScrollability) | Gets a SQLUINTEGER value that specifies the scrollability type. |
| [GetCursorSensitivity](#GetCursorSensitivity) | Gets a SQLUINTEGER value that specifies whether cursors on the statement handle made to a result set by another cursor. |
| [GetCursorType](#GetCursorType) | Gets SQLUINTEGER value that specifies the cursor type. |
| [GetData](#GetData) | Retrieves data for a single column in the result set. It can be called multiple times to retrieve variable-length data in parts. |
| [GetDiagField](#GetDiagField) | Returns the current value of a field of a record of the diagnostic data structure (associated with an statement handle) that contains error, warning, and status information. |
| [GetDiagRec](#GetDiagRec) | Returns the current values of multiple fields of a diagnostic record that contains error, warning, and status information. |
| [GetErrorInfo](#GetErrorInfo) | Returns a verbose description of the last errors. |
| [GetForeignKeys](#GetForeignKeys) | Returns list of foreign keys of the table. |
| [GetImpParamDesc](#GetImpParamDesc) | Returns the handle of the IPD. |
| [GetImpParamDescField](#GetImpParamDescField) | Returns the current setting or value of a single field of a descriptor record. |
| [GetImpParamDescFieldName](#GetImpParamDescFieldName) | Returns the name of a single field of a descriptor record. |
| [GetImpParamDescFieldNullable](#GetImpParamDescFieldNullable) | Returns the nullable value (TRUE or FALSE) of a single field of a descriptor record. |
| [GetImpParamDescFieldOctetLength](#GetImpParamDescFieldOctetLength) | Returns the octet length of a single field of a descriptor record. |
| [GetImpParamDescFieldPrecision](#GetImpParamDescFieldPrecision) | Returns the precision value of a single field of a descriptor record. |
| [GetImpParamDescFieldScale](#GetImpParamDescFieldScale) | Returns the scale value of a single field of a descriptor record. |
| [GetImpParamDescFieldType](#GetImpParamDescFieldType) | Returns the type of a single field of a descriptor record. |
| [GetImpParamDescRec](#GetImpParamDescRec) | Returns the current settings or values of multiple fields of a descriptor record. |
| [GetStmtImpRowDesc](#GetStmtImpRowDesc) | Returns the handle to the IRD. |
| [GetImpRowDescField](#GetImpRowDescField) | Returns the current setting or value of a single field of a descriptor record. |
| [GetImpRowDescFieldName](#GetImpRowDescFieldName) | Returns the name of a single field of a descriptor record. |
| [GetImpRowDescFieldNullable](#GetImpRowDescFieldNullable) | Returns the nullable value (TRUE or FALSE) of a single field of a descriptor record. |
| [GetImpRowDescFieldOctetLength](#GetImpRowDescFieldOctetLength) | Returns the octet length of a single field of a descriptor record. |
| [GetImpRowDescFieldPrecision](#GetImpRowDescFieldPrecision) | Returns the precision value of a single field of a descriptor record. |
| [GetImpRowDescFieldScale](#GetImpRowDescFieldScale) | Returns the scale value of a single field of a descriptor record. |
| [GetImpRowDescFieldType](#GetImpRowDescFieldType) | Returns the type of a single field of a descriptor record. |
| [GetImpRowDescRec](#GetImpRowDescRec) | Returns the current settings or values of multiple fields of a descriptor record. |
| [GetLongVarcharData](#GetLongVarcharData) | Retrieves long variable char data from the specified column of the result set. |
| [GetPrimaryKeys](#GetPrimaryKeys) | Returns the column names that make up the primary key for a table. |
| [GetProcedureColumns](#GetProcedureColumns) | Returns the list of input and output parameters, as well as the columns that make up the result set for the specified procedures. |
| [GetProcedures](#GetProcedures) | Returns the list of procedure names stored in a specific data source. |
| [GetSpecialColumns](#GetSpecialColumns) | Retrieves information about columns within a specified table. |
| [GetSqlState](#GetSqlState) | Returns the SqlState for the statement handle. |
| [GetStatistics](#GetStatistics) | Retrieves a list of statistics about a single table and the indexes associated with the table. |
| [GetStmtAppParamDesc](#GetStmtAppParamDesc) | Gets the handle to the APD for subsequent calls to **Execute** and **ExecDirect** on the statement handle. |
| [GetStmtAppRowDesc](#GetStmtAppRowDesc) | Gets the handle to the ARD for subsequent fetches on the statement handle. |
| [GetStmtAsyncEnable](#GetStmtAsyncEnable) | Gets an SQLUINTEGER value that specifies whether a function called with the specified statement is executed asynchronously. |
| [GetStmtAttr](#GetStmtAttr) | Returns the current setting of a statement attribute. |
| [GetStmtFetchBookmarkPtr](#GetStmtFetchBookmarkPtr) | Gets a pointer that points to a binary bookmark value. |
| [GetStmtImpParamDesc](#GetStmtImpParamDesc) | Gets the handle to the IPD. |
| [GetStmtMaxLength](#GetStmtMaxLength) | Gets an SQLUINTEGER value that specifies the maximum amount of data that the driver returns from a character or binary column. |
| [GetStmtMaxRows](#GetStmtMaxRows) | Gets an SQLUINTEGER value corresponding to the maximum number of rows to return to the application for a SELECT statement. |
| [GetStmtNoScan](#GetStmtNoScan) | Gets an SQLUINTEGER value that indicates whether the driver should scan SQL strings for escape sequences. |
| [GetStmtParamBindOffsetPtr](#GetStmtParamBindOffsetPtr) | Gets an SQLUINTEGER value that points to an offset added to pointers to change binding of dynamic parameters. |
| [GetStmtParamBindType](#GetStmtParamBindType) | Gets an SQLUINTEGER value that indicates the binding orientation to be used for dynamic parameters. |
| [GetStmtParamOperationPtr](#GetStmtParamOperationPtr) | Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a parameter during execution of an SQL statement. |
| [GetStmtParamsetSize](#GetStmtParamsetSize) | Gets an SQLUINTEGER value that specifies the number of values for each parameter. |
| [GetStmtParamsProcessedPtr](#GetStmtParamsProcessedPtr) | Gets an SQLUINTEGER record field that points to a buffer in which to return the number of sets of parameters that have been processed, including error sets. |
| [GetStmtParamStatusPtr](#GetStmtParamStatusPtr) | Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing status information for each row of parameter values after a  call to **Execute** or **ExecDirect**. |
| [GetStmtQueryTimeout](#GetStmtQueryTimeout) | Gets an SQLUINTEGER value corresponding to the number of seconds to wait for an SQL statement to execute before returning to the application. |
| [GetStmtRetrieveData](#GetStmtRetrieveData) | Gets an SQLUINTEGER value specifying the data retrieval mode. |
| [GetStmtRowArraySize](#GetStmtRowArraySize) | Gets an SQLUINTEGER value that specifies the number of rows returned by each call to **Fetch** or **FetchScroll**. |
| [GetStmtRowBindOffsetPtr](#GetStmtRowBindOffsetPtr) | Gets an SQLUINTEGER value that points to an offset added to pointers to change binding of column data. |
| [GetStmtRowBindType](#GetStmtRowBindType) | Gets an SQLUINTEGER value that sets the binding orientation to be used when **Fetch** or **FetchScroll** is called on the associated statement. |
| [GetStmtRowNumber](#GetStmtRowNumber) | Returns an SQLUINTEGER value that is the number of the current row in the entire result set. |
| [GetStmtRowOperationPtr](#GetStmtRowOperationPtr) | Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a row during a bulk operation using **SetPos**. |
| [GetStmtRowsFetchedPtr](#GetStmtRowsFetchedPtr) | Gets an SQLUINTEGER value that points to a buffer in which to return the number of rows fetched. |
| [GetStmtRowStatusPtr](#GetStmtRowStatusPtr) | Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing row status values after a call to **Fetch** or **FetchScroll**. |
| [GetStmtSimulateCursor](#GetStmtSimulateCursor) | Gets an SQLUINTEGER value that specifies whether drivers that simulate positioned update and delete statements guarantee that such statements affect only one single row. |
| [GetStmtUseBookmarks](#GetStmtUseBookmarks) | Gets an SQLUINTEGER value that specifies whether an application will use bookmarks with a cursor. |
| [GetTablePrivileges](#GetTablePrivileges) | Returns a list of tables and the privileges associated with each table. |
| [GetTables](#GetTables) | Returns the list of table, catalog, or schema names, and table types, stored in a specific data source. |
| [GetTypeInfo](#GetTypeInfo) | Returns information about data types supported by the data source. |
| [Handle](#Handle) | Returns the statement handle. |
| [LockRecord](#LockRecord) | Sets the cursor position in a rowset and locks the record. |
| [MoreResults](#MoreResults) | Determines whether more results are available on a statement containing SELECT, UPDATE, INSERT, or DELETE statements and, if so, initializes processing for those results. |
| [Move](#Move) | Moves the cursor forward or backward the specified number of rowsets. |
| [MoveFirst](#MoveFirst) | Fetches the first rowset of data from the result set and returns data for all bound columns. |
| [MoveLast](#MoveLast) | Fetches the last rowset of data from the result set and returns data for all bound columns. |
| [MoveNext](#MoveNext) | Fetches the next rowset of data from the result set and returns data for all bound columns. |
| [MovePrevious](#MovePrevious) | Fetches the previous rowset of data from the result set and returns data for all bound columns. |
| [NumParams](#NumParams) | Returns the number of parameters in an SQL statement. |
| [NumResultCols](#NumResultCols) | Returns the number of columns in a result set. |
| [ParamData](#ParamData) | Used in conjunction with **PutData** to supply parameter data at statement execution time. |
| [Prepare](#Prepare) | Prepares an SQL string for execution. |
| [PutData](#PutData) | Allows an application to send data for a parameter or column to the driver at statement execution time. |
| [RecordCount](#RecordCount) | Gets the number of records in the result set. |
| [RefreshRecord](#RefreshRecord) | Sets the cursor position in a rowset and allows to refresh data in the rowset. |
| [ResetParams](#ResetParams) | Releases all parameter buffers set by **BindParameter** for the given statement handle. |
| [RowCount](#RowCount) | Returns the number of rows affected by update, insert or delete statements. |
| [SetAbsolutePosition](#SetAbsolutePosition) | Fetches the rowset starting at the specified row. |
| [SetCursorConcurrency](#SetCursorConcurrency) | Sets a SQLUINTEGER value that specifies the cursor concurrency. |
| [SetCursorKeysetSize](#SetCursorKeysetSize) | Sets a SQLUINTEGER value that specifies the number of rows in the keyset-driven cursor. |
| [SetCursorKeysetSize](#SetCursorKeysetSize) | Sets a SQLUINTEGER value that specifies the number of rows in the keyset-driven cursor. |
| [SetCursorName](#SetCursorName) | Sets the cursor name associated with a specified statement. |
| [SetCursorScrollability](#SetCursorScrollability) | Sets a SQLUINTEGER value that specifies the scrollability type. |
| [SetCursorSensitivity](#SetCursorSensitivity) | Sets a SQLUINTEGER value that specifies whether cursors on the statement handle made to a result set by another cursor. |
| [SetCursorType](#SetCursorType) | Sets a SQLUINTEGER value that specifies the cursor type. |
| [SetDynamicCursor](#SetDynamicCursor) | Specifies a dynamic cursor. |
| [SetKeysetDrivenCursor](#SetKeysetDrivenCursor) | Specifies a keyset driven cursor. |
| [SetLockConcurrency](#SetLockConcurrency) | Cursor uses the lowest level of locking sufficient to ensure that the row can be updated. |
| [SetMultiuserKeysetCursor](#SetMultiuserKeysetCursor) | Creates a multiuser keyset cursor. |
| [SetOptimisticConcurrency](#SetOptimisticConcurrency) | Cursor uses optimistic concurrency control, comparing values. |
| [SetPos](#SetPos) | Fetches the rowset rowset from the start of the current rowset. |
| [SetPosition](#SetPosition) | Sets the cursor position in a rowset. |
| [SetReadOnlyConcurrency](#SetReadOnlyConcurrency) | Cursor is read-only. No updates are allowed. |
| [SetRelativePosition](#SetRelativePosition) | Fetches the rowset from from the start of the current rowset. |
| [SetRowVerConcurrency](#SetRowVerConcurrency) | Cursor uses optimistic concurrency control, comparing row versions such as SQLBase ROWID or Sybase TIMESTAMP. |
| [SetStaticCursor](#SetStaticCursor) | Specifies an static cursor. |
| [SetStmtAppParamDesc](#SetStmtAppParamDesc) | Sets the handle to the APD for subsequent calls to **Execute** and **ExecDirect** on the statement handle. |
| [SetStmtAppRowDesc](#SetStmtAppRowDesc) | Sets the handle to the ARD for subsequent fetches on the statement handle. |
| [SetStmtAttr](#SetStmtAttr) | Sets attributes related to a statement. |
| [SetStmtFetchBookmarkPtr](#SetStmtFetchBookmarkPtr) | Sets a pointer that points to a binary bookmark value. |
| [SetStmtMaxLength](#SetStmtMaxLength) | Sets an SQLUINTEGER value that specifies the maximum amount of data that the driver returns from a character or binary column. |
| [SetStmtMaxRows](#SetStmtMaxRows) | Sets an SQLUINTEGER value corresponding to the maximum number of rows to return to the application for a SELECT statement. |
| [SetStmtNoScan](#SetStmtNoScan) | Sets an SQLUINTEGER value that indicates whether the driver should scan SQL strings for escape sequences. |
| [SetStmtParamBindOffsetPtr](#SetStmtParamBindOffsetPtr) | Sets an SQLUINTEGER value that points to an offset added to pointers to change binding of dynamic parameters. |
| [SetStmtParamBindType](#SetStmtParamBindType) | Sets an SQLUINTEGER value that indicates the binding orientation to be used for dynamic parameters. |
| [SetStmtParamOperationPtr](#SetStmtParamOperationPtr) | Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a parameter during execution of an SQL statement. |
| [SetStmtParamsetSize](#SetStmtParamsetSize) | Sets an SQLUINTEGER value that specifies the number of values for each parameter. |
| [SetStmtParamsProcessedPtr](#SetStmtParamsProcessedPtr) | Sets an SQLUINTEGER record field that points to a buffer in which to return the number of sets of parameters that have been processed, including error sets. |
| [SetStmtParamStatusPtr](#SetStmtParamStatusPtr) | Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing status information for each row of parameter values after a  call to Execute or ExecDirect. This field is required only if PARAMSET_SIZE is greater than 1. |
| [SetStmtQueryTimeout](#SetStmtParamStatusPtr) | Sets an SQLUINTEGER value corresponding to the number of seconds to wait for an SQL statement to execute before returning to the application. |
| [SetStmtRetrieveData](#SetStmtRetrieveData) | Sets an SQLUINTEGER value specifying the data retrieval mode. |
| [SetStmtRowArraySize](#SetStmtRowArraySize) | Sets an SQLUINTEGER value that specifies the number of rows returned by each call to **Fetch** or **FetchScroll**. |
| [SetStmtRowBindOffsetPtr](#SetStmtRowBindOffsetPtr) | Sets an SQLUINTEGER value that points to an offset added to pointers to change binding of column data. |
| [SetStmtRowBindType](#SetStmtRowBindType) | Sets an SQLUINTEGER value that sets the binding orientation to be used when **Fetch** or **FetchScroll** is called on the associated statement. |
| [SetStmtRowOperationPtr](#SetStmtRowOperationPtr) | Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a row during a bulk operation using **SetPos**. |
| [SetStmtRowsFetchedPtr](#SetStmtRowsFetchedPtr) | Sets an SQLUINTEGER value that points to a buffer in which to return the number of rows fetched. |
| [SetStmtRowStatusPtr](#SetStmtRowStatusPtr) | Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing row status values after a call to **Fetch** or **FetchScroll**. |
| [SetStmtSimulateCursor](#SetStmtSimulateCursor) | Sets an SQLUINTEGER value that specifies whether drivers that simulate positioned update and delete statements guarantee that such statements affect only one single row. |
| [SetStmtUseBookmarks](#SetStmtUseBookmarks) | Sets an SQLUINTEGER value that specifies whether an application will use bookmarks with a cursor. |
| [SetStmtAsyncEnable](#SetStmtAsyncEnable) | Sets an SQLUINTEGER value that specifies whether a function called with the specified statement is executed asynchronously. |
| [UnbindCol](#UnbindCol) | Unbinds the specified column buffer bound by **BindCol** for the given statement handle. |
| [UnbindColumns](#UnbindColumns) | Unbinds all column buffers bound by **BindCol** for the given statement handle. |
| [UnlockRecord](#UnlockRecord) | Sets the cursor position in a rowset and unlocks the record. |
| [UpdateByBookmark](#UpdateByBookmark) | Updates a set of rows where each row is identified by a bookmark. |
| [UpdateRecord](#UpdateRecord) | Updates a record. |

# <a name="ConstructorsDb"></a>Constructors (CODBC)

Allocates a connection handle and, if needed, an environment handle, and opens the database.

```
CONSTRUCTOR CODBC (BYREF wszConnectionString AS WSTRING, BYVAL nODbcVersion AS SQLINTEGER = SQL_OV_ODBC3, _
   BYVAL ConnectionPoolingAttr AS SQLUINTEGER = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszConnectionString* | The connection string. |
| *nOdbcVersion* | Optional. ODBC version number: SQL_OV_ODBC2, SQL_OV_ODBC3 or SQL_OV_ODBC3_80. |
| *ConnectionPoolingAttr* | Optional. SQL_CP_ONE_PER_DRIVER or SQL_CP_ONE_PER_HENV. |

Establishes connections to a driver and a data source. The connection handle references storage of all information about the connection to the data source, including status, transaction state, and error information. 

```
CONSTRUCTOR CODBC (BYREF wszServerName AS WSTRING, BYREF wszUserName AS WSTRING, _
   BYREF wszAuthentication AS WSTRING, BYVAL nODbcVersion AS SQLINTEGER = SQL_OV_ODBC3, _
   BYVAL ConnectionPoolingAttr AS SQLUINTEGER = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszServerName* | Data source name. The data might be located on the same computer as the program, or on another computer somewhere on a network. |
| *wszUserName* | User identifier. |
| *wszAuthentication* | Authentication string (typically the password). |
| *nOdbcVersion* | Optional. ODBC version number: SQL_OV_ODBC2, SQL_OV_ODBC3 or SQL_OV_ODBC3_80. |
| *ConnectionPoolingAttr* | Optional. SQL_CP_ONE_PER_DRIVER or SQL_CP_ONE_PER_HENV. |

#### GetLastResult

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, SQL_INVALID_HANDLE.

# <a name="Error"></a>Error (CODBC)

Returns TRUE if there has been an error; FALSE, otherwise.

```
FUNCTION Error () AS BOOLEAN
```

#### Return value

Returns TRUE if the last result code is SQL_ERROR or SQL_INVALID_HANDLE.

# <a name="GetCPMatch"></a>GetCPMatch (CODBC)

Returns a 32-bit SQLUINTEGER value that determines how a connection is chosen from a connection pool.

```
FUNCTION GetCPMatch () AS SQLUINTEGER
```

#### Return value

The current value of the attribute.

#### Result code (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Remarks

The Driver Manager determines which connection is reused from the pool and attempts to match the connection options in the call and the connection attributes set by the application to the keywords and connection attributes of the connections in the pool. The value of this attribute determines the level of precision of the matching criteria. The following values are used to set the value of this attribute:

**SQL_CP_STRICT_MATCH**

Only connections that exactly match the connection options in the call and the connection attributes set by the application are reused. This is the default.

**SQL_CP_RELAXED_MATCH**

Connections with matching connection string keywords can be used. Keywords must match, but not all connection attributes must  match.

# <a name="GetDataSources"></a>GetDataSources (CODBC)

Lists available DSN / Drivers installed.

```
FUNCTION GetDataSources (BYVAL Direction AS SQLUSMALLINT, BYREF cwsServerName AS CWSTR, _
   BYREF cwsDescription AS CWSTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Direction* | Determines which data source the Driver Manager returns information on. Can be:<br>SQL_FETCH_NEXT (to fetch the next data source name in the list), SQL_FETCH_FIRST (to fetch from the beginning of the list), SQL_FETCH_FIRST_USER (to fetch the first user DSN), or SQL_FETCH_FIRST_SYSTEM (to fetch the first system DSN).<br>When *Direction* is set to SQL_FETCH_FIRST, subsequent calls to **GetDataSources** with *Direction* set to SQL_FETCH_NEXT return both user and system DSNs. When *Direction* is set to SQL_FETCH_FIRST_USER, all subsequent calls to **GetDataSources** with *Direction* set to SQL_FETCH_NEXT return only user DSNs. When *Direction* is set to SQL_FETCH_FIRST_SYSTEM, all subsequent calls to **GetDataSources** with *Direction* set to SQL_FETCH_NEXT return only system DSNs. |
| *cwsServerName* | An string variable in which to return the data source name. |
| *cwsDescription* | An string variable in which to return the description. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **GetDataSources** returns either SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling the **SqlState** property.

# <a name="GetDrivers"></a>GetDrivers (CODBC)

Lists driver descriptions and driver attribute keywords. This function is implemented only by the Driver Manager.

```
FUNCTION GetDrivers (BYVAL Direction AS SQLUSMALLINT, BYREF cwsDriverDesc AS CWSTR, _
   BYREF cwsDriverAttributes AS CWSTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Direction* | Determines whether the Driver Manager fetches the next driver description in the list (SQL_FETCH_NEXT) or whether the search starts from the beginning of the list (SQL_FETCH_FIRST). |
| *cwsDriverDesc* | A CWSTR variable in which to return the driver description. |
| *cwsDriverAttributes* | A CWSTR variable in which to return the list of driver attribute value pairs (see "Comments"). |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **GetDrivers** returns either SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling the **SqlState** property.

#### Comments

**GetDrivers** returns the driver description in the *cwsDriverDesc* variable. It returns additional information about the driver in the *cwsDriverAttributes* buffer as a list of keyword-value pairs. All keywords listed in the system information for drivers will be returned for all drivers, except for **CreateDSN**, which is used to prompt creation of data sources and therefore is optional. Each pair is terminated with a semicolon.

Driver attribute keywords are added from the system information when the driver is installed.

An application can call GetDrivers multiple times to retrieve all driver descriptions. The Driver Manager retrieves this information from the system information. When there are no more driver descriptions, **GetDrivers** returns SQL_NO_DATA. If **GetDrivers** is called with SQL_FETCH_NEXT immediately after it returns SQL_NO_DATA, it returns the first driver description.

If SQL_FETCH_NEXT is passed to **GetDrivers** the very first time it is called, **GetDrivers** returns the first data source name.

Because **GetDrivers** is implemented in the Driver Manager, it is supported for all drivers regardless of a particular driver's standards compliance.

# <a name="GetEnvAttr"></a>GetEnvAttr (CODBC)

Returns the current setting of an environment attribute.

```
FUNCTION GetEnvAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
   BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Attribute* | Attribute to retrieve. |
| *ValuePtr* |  Pointer to a buffer in which to return the current value of the attribute specified by **Attribute**. |
| *BufferLength* | If *ValuePtr* points to a character string, this argument should be the length of *ValuePtr*. If *ValuePtr* is an integer, *BufferLength* is ignored. If the attribute value is not a character string, *BufferLength* is unused. |
| *StringLength* | A pointer to a buffer in which to return the total number of bytes (excluding the null-termination character) available to return in *ValuePtr*. If *ValuePtr* is a null pointer, no length is returned. If the attribute value is a character string and the number of bytes available to return is greater than or equal to *BufferLength*, the data in *ValuePtr* is truncated to *BufferLength* minus the length of a null-termination character and is null-terminated by the driver. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetErrorInfo"></a>GetErrorInfo (CODBC)

Returns a verbose description of the last error(s).

```
FUNCTION GetErrorInfo (BYVAL HandleType AS SQLSMALLINT, BYVAL Handle AS SQLHANDLE, _
   BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *HandleType* | The handle type:<br>SQL_HANDLE_DBC (connection handle)<br>SQL_HANDLE_STMT (statement handle) |
| *Handle* | The handle value. |
| *iErrorCode* | Optional. The error code returned by **GetLastResult**. |

#### Return value

The description of the error or errors.

# <a name="GetLastResult"></a>GetLastResult (CODBC)

Returns the last result code.

```
FUNCTION GetLastResult () AS SQLRETURN
```

#### Return value

The result code returned by the last executed ODBC method.

GetDiagRec or GetDiagField returns SQLSTATE values as defined by X/Open Data Management: Structured Query Language (SQL), Version 2 (March 1995). SQLSTATE values are strings that contain five characters. The following table lists SQLSTATE values that a driver can return for GetDiagRec.

The character string value returned for an SQLSTATE consists of a two-character class value followed by a three-character subclass value. A class value of "01" indicates a warning and is accompanied by a return code of SQL_SUCCESS_WITH_INFO. Class values other than "01," except for the class "IM," indicate an error and are accompanied by a return value of SQL_ERROR. The class "IM" is specific to warnings and errors that derive from the implementation of ODBC itself. The subclass value "000" in any class indicates that there is no subclass for that SQLSTATE. The assignment of class and subclass values is defined by SQL-92.

Note Although successful execution of a function is normally indicated by a return value of SQL_SUCCESS, the SQLSTATE 00000 also indicates success.

# <a name="GetOutputNTS"></a>GetOutputNTS (CODBC)

Returns a 32-bit integer that determines how the driver returns string data. If SQL_TRUE, the driver returns string data null-terminated. If SQL_FALSE, the driver does not return string data null-terminated. This attribute defaults to SQL_TRUE. A call to SetEnvAttr to set it to SQL_TRUE returns SQL_SUCCESS. A call to SetEnvAttr to set it to SQL_FALSE returns SQL_ERROR and SQLSTATE HYC00.

**Note**: Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION GetOutputNTS () AS SQLUINTEGER
```

#### Return value

The current value of the attribute.

#### Result code (GetLastStatus)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="GetSqlState"></a>GetSqlState (CODBC)

Returns the SqlState for the specified handle.

**Note**: Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION GetSqlState (BYVAL HandleType AS SQLSMALLINT, BYVAL Handle AS SQLHANDLE) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *HandleType* | The handle type:<br>SQL_HANDLE_DBC (connection handle)<br>SQL_HANDLE_STMT (statement handle) |
| *Handle* | The handle value. |

#### Result code (GetLastStatus)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Return value

The SqlState.

#### ODBC Error Codes

**GetDiagRec** or **GetDiagField** returns SQLSTATE values as defined by X/Open Data Management: Structured Query Language (SQL), Version 2 (March 1995). SQLSTATE values are strings that contain five characters. The following table lists SQLSTATE values that a driver can return for GetDiagRec.

The character string value returned for an SQLSTATE consists of a two-character class value followed by a three-character subclass value. A class value of "01" indicates a warning and is accompanied by a return code of SQL_SUCCESS_WITH_INFO. Class values other than "01," except for the class "IM," indicate an error and are accompanied by a return value of SQL_ERROR. The class "IM" is specific to warnings and errors that derive from the implementation of ODBC itself. The subclass value "000" in any class indicates that there is no subclass for that SQLSTATE. The assignment of class and subclass values is defined by SQL-92.

**Note**: Although successful execution of a function is normally indicated by a return value of SQL_SUCCESS, the SQLSTATE 00000 also indicates success.

| SQLSTATE | Error |
| -------- | ----- |
| 01000 | General warning |
| 01001 | Cursor operation conflict |
| 01002 | Disconnect error |
| 01003 | NULL value eliminated in set function |
| 01004 | String data, right-truncated |
| 01006 | Privilege not revoked |
| 01007 | Privilege not granted |
| 01S00 | Invalid connection string attribute |
| 01S01 | Error in row |
| 01S02 | Option value changed |
| 01S06 | Attempt to fetch before the result set returned the first rowset |
| 01S08 | Error saving File DSN |
| 01S09 | Invalid keyword |
| 07001 | Wrong number of parameters |
| 07002 | COUNT field incorrect |
| 07005 | Prepared statement not a cursor-specification |
| 07006 | Restricted data type attribute violation |
| 07009 | Invalid descriptor index |
| 07S01 | Invalid use of default parameter |
| 08001 | Client unable to establish connection |
| 08002 | Connection name in use |
| 08003 | Connection does not exist |
| 08004 | Server rejected the connection |
| 08007 | Connection failure during transaction |
| 08S01 | Communication link failure |
| 21S01 | Insert value list does not match column list |
| 21S02 | Degree of derived table does not match column list |
| 22001 | String data, right-truncated |
| 22002 | Indicator variable required but not supplied |
| 22003 | Numeric value out of range |
| 22007 | Invalid datetime format |
| 22008 | Datetime field overflow |
| 22012 | Division by zero |
| 22015 | Interval field overflow |
| 22018 | Invalid character value for cast specification |
| 22019 | Invalid escape character |
| 22025 | Invalid escape sequence |
| 22026 | String data, length mismatch |
| 23000 | Integrity constraint violation |
| 24000 | Invalid cursor state |
| 25000 | Invalid transaction state |
| 25S01 | Transaction state |
| 25S02 | Transaction is still active |
| 25S03 | Transaction is rolled back |
| 28000 | Invalid authorization specification |
| 34000 | Invalid cursor name |
| 3C000 | Duplicate cursor name |
| 3D000 | Invalid catalog name |
| 3F000 | Invalid schema name |
| 40001 | Serialization failure |
| 40002 | Integrity constraint violation |
| 40003 | Statement completion unknown |
| 42000 | Syntax error or access violation |
| 42S01 | Base table or view already exists |
| 42S02 | Base table or view not found |
| 42S11 | Index already exists |
| 42S12 | Index not found |
| 42S21 | Column already exists |
| 42S22 | Column not found |
| 44000 | WITH CHECK OPTION violation |
| HY000 | General error |
| HY001 | Memory allocation error |
| HY003 | Invalid application buffer type |
| HY004 | Invalid SQL data type |
| HY007 | Associated statement is not prepared |
| HY008 | Operation canceled |
| HY009 | Invalid use of null pointer |
| HY010 | Function sequence error |
| HY011 | Attribute cannot be set now |
| HY012 | Invalid transaction operation code |
| HY013 | Memory management error |
| HY014 | Limit on the number of handles exceeded |
| HY015 | No cursor name available |
| HY016 | Cannot modify an implementation row descriptor |
| HY017 | Invalid use of an automatically allocated descriptor handle |
| HY018 | Server declined cancel request |
| HY019 | Non-character and non-binary data sent in pieces |
| HY020 | Attempt to concatenate a null value |
| HY021 | Inconsistent descriptor information |
| HY024 | Invalid attribute value |
| HY090 | Invalid string or buffer length |
| HY091 | Invalid descriptor field identifier |
| HY092 | Invalid attribute/option identifier |
| HY095 | Function type out of range |
| HY096 | Invalid information type |
| HY097 | Column type out of range |
| HY098 | Scope type out of range |
| HY099 | Nullable type out of range |
| HY100 | Uniqueness option type out of range |
| HY101 | Accuracy option type out of range |
| HY103 | Invalid retrieval code |
| HY104 | Invalid precision or scale value |
| HY105 | Invalid parameter type |
| HY106 | Fetch type out of range |
| HY107 | Row value out of range |
| HY109 | Invalid cursor position |
| HY110 | Invalid driver completion |
| HY111 | Invalid bookmark value |
| HYC00 | Optional feature not implemented |
| HYT00 | Timeout expired |
| HYT01 | Connection timeout expired |
| IM001 | Driver does not support this function |
| IM002 | Data source name not found and no default driver specified |
| IM003 | Specified driver could not be loaded |
| IM004 | Driver's SQLAllocHandle on SQL_HANDLE_ENV failed |
| IM005 | Driver's SQLAllocHandle on SQL_HANDLE_DBC failed |
| IM006 | Driver's SQLSetConnectAttr failed |
| IM007 | No data source or driver specified; dialog prohibited |
| IM008 | Dialog failed |
| IM009 | Unable to load translation DLL |
| IM010 | Data source name too long |
| IM011 | Driver name too long |
| IM012 | DRIVER keyword syntax error |
| IM013 | Trace file error |
| IM014 | Invalid name of File DSN |
| IM015 | Corrupt file data source |


# <a name="ODBCVersion"></a>ODBCVersion (CODBC)

Returns a 32-bit integer that determines whether certain functionality exhibits ODBC 2.x behavior or ODBC 3.x behavior.

```
FUNCTION ODBCVersion () AS SQLUINTEGER
```

#### Return value

* **SQL_OV_ODBC3_80** = The Driver Manager and driver exhibit the following ODBC 3.8 behavior:

   The driver returns and expects ODBC 3.x codes for date, time, and timestamp.
   The driver returns ODBC 3.x SQLSTATE codes when Error, GetDiagField, or GetDiagRec is called.
   The CatalogName argument in a call to SQLTables accepts a search pattern.
   The Driver Manager supports C data type extensibility. For more information about C data type extensibility, see C Data Types in ODBC.

* **SQL_OV_ODBC3** = The Driver Manager and driver exhibit the following ODBC 3.x behavior:

   The driver returns and expects ODBC 3.x codes for date, time, and timestamp.
   The driver returns ODBC 3.x SQLSTATE codes when Error,  GetDiagField, or GetDiagRec is called.
   The CatalogName argument in a call to SqlTables accepts a search pattern.

* **SQL_OV_ODBC2** = The Driver Manager and driver exhibit the following ODBC 2.x behavior. This is especially useful for an ODBC 2.x application working with an ODBC 3.x driver.

   The driver returns and expects ODBC 2.x codes for date, time, and timestamp.
   The driver returns ODBC 2.x SQLSTATE codes when Error, GetDiagField, or GetDiagRec is called.
   The CatalogName argument in a call to SqlTables does not accept a search pattern.

An application must set this environment attribute before calling any function that has an SQLHENV argument, or the call will return SQLSTATE HY010 (Function sequence error). It is driver-specific whether or not additional behaviors exist for these environmental flags.

#### Remarks

To set the ODBC version, use the optional parameters of the **CODBC** class constructors.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetCPMatch"></a>SetCPMatch (CODBC)

Sets a 32-bit SQLUINTEGER value that determines how a connection is chosen from a connection pool.

```
FUNCTION SetCPMatch (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

The Driver Manager determines which connection is reused from the pool and attempts to match the connection options in the call and the connection attributes set by the application to the keywords and connection  attributes of the connections in the pool. The value of this attribute determines the level of precision of the matching criteria. The following values are used to set the value of this attribute:

**SQL_CP_STRICT_MATCH**

Only connections that exactly match the connection options in the call and the connection attributes set by the application are reused. This is the default.

**SQL_CP_RELAXED_MATCH**

Connections with matching connection string keywords can be used. Keywords must match, but not all connection attributes must  match.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetEnvAttr"></a>SetEnvAttr (CODBC)

Sets attributes that govern aspects of environments.

```
FUNCTION SetEnvAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
   BYVAL StringLength AS SQLINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Attribute* | Attribute to set. |
| *ValuePtr* | Pointer to the value to be associated with *Attribute*. Depending on the value of *Attribute*, *ValuePtr* will be a 32-bit integer value or point to a null-terminated character string. |
| *StringLength* | If *ValuePtr* points to a character string or a binary buffer, this argument should be the length of *ValuePtr*. If *ValuePtr* is an integer, *StringLength* is ignored. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetErrorProc"></a>SetErrorProc (CODBC)

Sets the address of an application defined error callback.

```
SUB SetErrorProc (BYVAL pProc AS ANY PTR, BYVAL reportWarnings AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pProc* | Address of the application defined callback. |
| *reportWarnings* | Optional. Report also warnings. |

#### Example of an application defined callback:

```
SUB ODBC_ErrorCallback (BYVAL nResult AS SQLRETURN, BYREF wszSrc AS WSTRING, BYREF wszErrorMsg AS WSTRING)
   PRINT "Error: " & STR(nResult) & " - Source: " & wszSrc
   IF LEN(wszErrorMsg) THEN PRINT wszErrorMsg
END SUB

pDbc.SetErrorProc(@ODBC_ErrorCallback)    ' // Sets the error callback for the connection object
pStmt.SetErrorProc(@ODBC_ErrorCallback)   ' // Sets the error callback for the statement object
```

# <a name="SetOutputNTS"></a>SetOutputNTS (CODBC)

Returns a 32-bit integer that determines how the driver returns string data. If SQL_TRUE, the driver returns string data null-terminated. If SQL_FALSE, the driver does not return string data null-terminated. This attribute defaults to SQL_TRUE. A call to SetEnvAttr to set it to SQL_TRUE returns SQL_SUCCESS. A call to SetEnvAttr to set it to SQL_FALSE returns SQL_ERROR and SQLSTATE HYC00.

**Note**: Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION SetOutputNTS (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. SQL_TRUE or SQL_FALSE. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

