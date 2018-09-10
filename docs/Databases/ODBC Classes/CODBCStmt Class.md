# CODBCStmt Class

Implements methods to create and manage statement objects. Inherits from CODBCBase.

**Include file**: CODBCStmt.inc

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Allocates an statement handle. |
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
| [GetImpRowDesc](#GetImpRowDesc) | Returns the handle to the IRD. |
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

# <a name="Constructors"></a>Constructors

Allocates an statement handle. A statement handle provides access to statement information, such as error messages, the cursor name, and status information for SQL statement processing.

```
CONSTRUCTOR COdbcStmt (BYVAL pDbc AS COdbc PTR)
CONSTRUCTOR COdbcStmt (BYREF pCDbc AS COdbc)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pDbc* | Pointer to a connection object. |
| *pCDbc* | Reference to an instance of the CODBC class. |

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_INVALID_HANDLE, or SQL_ERROR.

### Example

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

# <a name="AddRecord"></a>AddRecord

Adds a record to the database.

```
FUNCTION AddRecord () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Unable to create the statement handle" : SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor
' // Bind the columns
DIM AS LONG lAuId, cbAuId
pStmt.BindCol(1, @lAuId, @cbAuId)
DIM wszAuthor AS WSTRING * 256, cbAuthor AS LONG
pStmt.BindCol(2, @wszAuthor, SIZEOF(wszAuthor), @cbAuthor)
DIM iYearBorn AS SHORT, cbYearBorn AS LONG
pStmt.BindCol(3, @iYearBorn, @cbYearBorn)

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Fill the values of the bounded application variables and its sizes
lAuId     = 999               : cbAuID     = SIZEOF(lAuId)
wszAuthor = "Edgar Allan Poe" : cbAuthor   = LEN(wszAuthor)
iYearBorn = 1809              : cbYearBorn = SIZEOF(iYearBorn)
' // Add the record
pStmt.AddRecord
IF pStmt.Error = FALSE THEN PRINT "Record added" ELSE PRINT pStmt.GetErrorInfo
```

# <a name="BindCol"></a>BindCol

Binds application data buffers to columns in the result set. 

```
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetType AS SQLSMALLINT, _
   BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, _
   BYVAL StrLen_or_Ind AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS BYTE PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS UBYTE PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SHORT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS USHORT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS LONG PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ULONG PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SINGLE PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS DOUBLE PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS LONGINT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ULONGINT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ZSTRING PTR, _
   BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS WSTRING PTR, _
   BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS DATE_STRUCT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION  BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS TIME_STRUCT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS TIMESTAMP_STRUCT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ANY PTR, _
   BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindColToBit (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SHORT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindColToNumeric (BYVAL ColNumber AS SQLUSMALLINT, BYREF TargetValue AS ZSTRING, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindColToDecimal (BYVAL ColNumber AS SQLUSMALLINT, BYREF TargetValue AS ZSTRING, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNumber* | Number of the result set column to bind. Columns are numbered in increasing column order starting at 0, where column 0 is the bookmark column. If bookmarks are not used — that is, the SQL_ATTR_USE_BOOKMARKS statement attribute is set to SQL_UB_OFF — then column numbers start at 1. |
| *TargetValue* | Pointer to the data buffer to bind to the column. **Fetch** and **FetchScroll** return data in this buffer. **BulkOperations** returns data in this buffer when *Operation* is SQL_FETCH_BY_BOOKMARK; it retrieves data from this buffer when *Operation* is SQL_ADD or SQL_UPDATE_BY_BOOKMARK. **SetPos** returns data in this buffer when *Operation* is SQL_REFRESH; it retrieves data from this buffer when *Operation* is SQL_UPDATE.<br>If *TargetValue* is a null pointer, the driver unbinds the data buffer for the column. An application can unbind all columns by calling **UnbindColumns** or **FreeStmt** with the SQL_UNBIND option. An application can unbind the data buffer for a column but still have a length/indicator buffer bound for the column, if the *TargetValue* argument in the call to **BindCol** is a null pointer but the *StrLen_or_IndPtr* argument is a valid value. |
| *StrLen_or_IndPtr* | Pointer to the length/indicator buffer to bind to the column. **Fetch** and **FetchScroll** return a value in this buffer. **BulkOperations** retrieves a value from this buffer when *Operation* is SQL_ADD, SQL_UPDATE_BY_BOOKMARK, or SQL_DELETE_BY_BOOKMARK. **BulkOperations** returns a value in this buffer when *Operation* is SQL_FETCH_BY_BOOKMARK. **SetPos** returns a value in this buffer when *Operation* is SQL_REFRESH; it retrieves a value from this buffer when *Operation* is SQL_UPDATE.<br><br>**Fetch**, **FetchScroll**, **BulkOperations**, and **SetPos** can return the following values in the length/indicator buffer:<br><ul><li>The length of the data available to return</li><li>SQL_NO_TOTAL</li><li>SQL_NULL_DATA</li></ul>The application can place the following values in the length/indicator buffer for use with **BulkOperations** or **SetPos**:<ul><li>The length of the data being sent</li><li>SQL_NTS</li><li>SQL_NULL_DATA</li><li>SQL_DATA_AT_EXEC</li><li>The result of the SQL_LEN_DATA_AT_EXEC macro</li><li>SQL_COLUMN_IGNORE</li></ul>If the indicator buffer and the length buffer are separate buffers, the indicator buffer can return only SQL_NULL_DATA, while the length buffer can return all other values.<br><br>If *StrLen_or_IndPtr* is a null pointer, no length or indicator value is used. This is an error when fetching data and the data is NULL.  |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_INVALID_HANDLE, or SQL_ERROR.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Unable to create the statement handle" : SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor
' // Bind the columns
DIM AS LONG lAuId, cbAuId
pStmt.BindCol(1, @lAuId, @cbAuId)
DIM wszAuthor AS WSTRING * 260, cbAuthor AS LONG
pStmt.BindCol(2, @wszAuthor, SIZEOF(wszAuthor), @cbAuthor)
DIM iYearBorn AS SHORT, cbYearBorn AS LONG
pStmt.BindCol(3, @iYearBorn, @cbYearBorn)

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

# <a name="BindParameter"></a>BindParameter

Binds a buffer to a parameter marker in an SQL statement.

```
FUNCTION BindParameter (BYVAL ParameterNumber AS SQLUSMALLINT, BYVAL InputOutputType AS SQLSMALLINT, _
   BYVAL ValueType AS SQLSMALLINT, BYVAL ParameterType AS SQLSMALLINT, BYVAL ColumnSize AS SQLULEN, _
   BYVAL DecimalDigits AS SQLSMALLINT, BYVAL ParameterValuePtr AS SQLPOINTER, _
   BYVAL BufferLength AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ParameterNumber* | Parameter number, ordered sequentially in increasing parameter order, starting at 1. |
| *InputOutputType* | The type of the parameter. |
| *ValueType* | The C data type of the parameter. |
| *ParameterType* | The SQL data type of the parameter. |
| *ColumnSize* | The size of the column or expression of the corresponding parameter marker. |
| *DecimalDigits* | The decimal digits of the column or expression of the corresponding parameter marker. |
| *ParameterValuePtr* | A pointer to a buffer for the parameter's data. |
| *BufferLength* | Length of the ParameterValuePtr buffer in bytes. |
| *StrLen_or_IndPtr* | A pointer to a buffer for the parameter's length. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_INVALID_HANDLE, or SQL_ERROR.

# <a name="BulkOperations"></a>BulkOperations

Performs bulk insertions and bulk bookmark operations, including update, delete, and fetch by bookmark. 

```
FUNCTION BulkOperations (BYVAL Operation AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Operation* |Operation to perform: SQL_ADD, SQL_UPDATE_BY_BOOKMARK, SQL_DELETE_BY_BOOKMARK, SQL_FETCH_BY_BOOKMARK |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="Cancel"></a>Cancel

Cancels the processing on a statement.

```
FUNCTION Cancel () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="CloseCursor"></a>CloseCursor

Closes a cursor that has been opened on a statement and discards pending results. Returns SQLSTATE 24000 (Invalid cursor state) if no cursor is open in the statement.

```
FUNCTION CloseCursor () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="ColAttribute"></a>ColAttribute

Returns descriptor information for a column in a result set. Descriptor information is returned as a character string, a 32-bit descriptor-dependent value, or an integer value.

```
FUNCTION ColAttribute (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL FieldIdentifier AS SQLUSMALLINT, _
   BYVAL CharacterAttribute AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, _
   BYVAL StringLength AS SQLSMALLINT PTR, BYVAL NumericAttribute AS ANY PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColumnNumber* | The number of the record in the IRD from which the field value is to be retrieved. This argument corresponds to the column number of result data, ordered sequentially in increasing column order, starting at 1. Columns can be described in any order. Column 0 can be specified in this argument, but all values except SQL_DESC_TYPE and SQL_DESC_OCTET_LENGTH will return undefined values. | 
| *FieldIdentifier* | The field in row *ColumnNumber* of the IRD that is to be returned. |
| *CharacterAttribute* | Pointer to a buffer in which to return the value in the *FieldIdentifier* field of the *ColumnNumber* row of the IRD, if the field is a character string. Otherwise, the field is unused. |
| *BufferLength* | If *FieldIdentifier* is an ODBC-defined field and *CharacterAttribute* points to a character string or binary buffer, this argument should be the length of *CharacterAttribute*. If *FieldIdentifier* is an ODBC-defined field and *CharacterAttribute* is an integer, this field is ignored. If *FieldIdentifier* is a driver-defined field, the application indicates the nature of the field to the Driver Manager by setting the *BufferLength* argument. *BufferLength* can have the following values:<br><br><ul><li>If *CharacterAttribute* is a pointer to a pointer, *BufferLength* should have the value SQL_IS_POINTER.</li><li>If *CharacterAttribute* is a pointer to a character string, the *BufferLength* is the length of the buffer.</li><li>If *CharacterAttribute* is a pointer to a binary buffer, the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *BufferLength*. This places a negative value in *BufferLength*.</li><li>If *CharacterAttribute* is a pointer to a fixed-length data type, *BufferLength* must be one of the following: SQL_IS_INTEGER, SQL_IS_UNINTEGER, SQL_SMALLINT, or %SQLUSMALLINT.</li></ul> |
| *StringLength* | Pointer to a buffer in which to return the total number of bytes (excluding the null-termination byte for character data) available to return in *CharacterAttribute*.<br>For character data, if the number of bytes available to return is greater than or equal to *BufferLength*, the descriptor information in *CharacterAttribute* is truncated to *BufferLength* minus the length of a null-termination character and is null-terminated by the driver.<br>For all other types of data, the value of *BufferLength* is ignored and the driver assumes the size of *CharacterAttribute* is 32 bits or 64 bits. |
| *NumericAttribute* | Pointer to an integer buffer in which to return the value in the *FieldIdentifier* field of the *ColumnNumber* row of the IRD, if the field is a numeric descriptor type, such as SQL_DESC_COLUMN_LENGTH. Otherwise, the field is unused. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Unable to create the statement handle" : SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor
' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Retrieve the number of columns
DIM numCols AS SHORT = pStmt.NumResultCols
PRINT "Number of columns:" & STR(numCols)
IF numCols = 0 THEN PRINT "There are no columns": SLEEP : END
' // Retrieve the names of the fields (columns)
FOR idx AS LONG = 1 TO numCols
   PRINT "Field #" & STR(idx) & " name: " & pStmt.ColName(idx)
NEXT
' // Parse the result set
DO
   ' // Fetch the record
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ' // Get the values of the columns and display them
   FOR idx AS LONG = 1 TO numCols
      PRINT pStmt.GetData(idx)
   NEXT
LOOP
```

# <a name="ColAutoUniqueValue"></a>ColAutoUniqueValue

Returns SQL_TRUE if the column is an autoincrementing column, SQL_FALSE if the column is not an autoincrementing column or is not numeric. This field is valid for numeric data type columns only. An application can insert values into a row containing an autoincrement number, but typically cannot update values in the column. When an insert is made into an autoincrement column, a unique value is inserted into the column at insert time. The increment is not defined, but is data source-specific. An application should not assume that an autoincrement column starts at any particular point or increments by any particular value.

```
FUNCTION ColAutoUniqueValue (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColBaseColumnName"></a>ColBaseColumnName

Returns the base column name for the result set column. If a base column name does not exist (as in the case of columns that are expressions), then this variable contains an empty string. This information is returned from the SQL_DESC_BASE_COLUMN_NAME record field of the IRD, which is a read-only field.

```
FUNCTION ColBaseColumnName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The base column name for the result set column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColBaseTableName"></a>ColBaseTableName

Returns the name of the base table that contains the column. If the base table name cannot be defined or is not applicable, then this variable contains an empty string. This information is returned from the SQL_DESC_BASE_TABLE_NAME record field of the IRD, which is a read-only field.

```
FUNCTION ColBaseTableName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The name of the base table that contains the column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColCaseSensitive"></a>ColCaseSensitive

Returns SQL_TRUE if the column is treated as case-sensitive for collations and comparisons. Returns SQL_FALSE if the column is not treated as case-sensitive for collations and comparisons or is noncharacter.

```
FUNCTION ColCaseSensitive (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColCatalogName"></a>ColCatalogName

Returns the catalog of the table that contains the column. The returned value is implementation-defined if the column is an expression or if the column is part of a view. If the data source does not support catalogs or the catalog name cannot be determined, an empty string is returned. This VARCHAR record field is not limited to 128 characters.

```
FUNCTION ColCatalogName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The catalog of the table that contains the column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColConciseType"></a>ColConciseType

Returns the concise data type.

For the datetime and interval data types, this field returns the concise data type; for example, SQL_TYPE_TIME or SQL_INTERVAL_YEAR. This information is returned from the SQL_DESC_CONCISE_TYPE record field of the IRD.

```
FUNCTION ColConciseType (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The concise data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColCount"></a>ColCount

Returns the number of columns available in the result set.

```
FUNCTION ColCount () AS SQLINTEGER
```

#### Return value

The number of columns available in the result set. It returns 0 if there are no columns in the result set.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColDisplaySize"></a>ColDisplaySize

Returns the maximum number of characters required to display data from the column.

```
FUNCTION ColDisplaySize (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The maximum number of characters required to display data from the column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColFixedPrecScale"></a>ColFixedPrecScale

Returns SQL_TRUE if the column has a fixed precision and nonzero scale that are data source-specific. Returns SQL_FALSE if the column does not have a fixed precision and nonzero scale that are data source-specific.

```
FUNCTION ColFixedPrecScale (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColIsNull"></a>ColIsNull

Checks if the column is null. Should not be used with a column that is currently binded to a variable or buffer or it will return an error. The binding functions already return an indicator in the last parameter.

```
FUNCTION ColIsNull (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColLabel"></a>ColLabel

Returns the column label or title. For example, a column named EmpName might be labeled Employee Name or might be labeled with an alias. If a column does not have a label, the column name is returned. If the column is unlabeled and unnamed, an empty string is returned.

```
FUNCTION ColLabel (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The column label or title.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColLength"></a>ColLength

Returns a numeric value that is either the maximum or actual character length of a character string or binary data type. It is the maximum character length for a fixed-length data type, or the actual character length for a variable-length data type. Its value always excludes the null-termination byte that ends the character string. This information is returned from the SQL_DESC_LENGTH record field of the IRD.

```
FUNCTION ColLength (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The maximum or actual character length of a character string or binary data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColLiteralPrefix"></a>ColLiteralPrefix

This VARCHAR(128) record field contains the character or characters that the driver recognizes as a prefix for a literal of this data type. This field contains an empty string for a data type for which a literal prefix is not applicable.

```
FUNCTION ColLiteralPrefix (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The character or characters that the driver recognizes as a prefix for a literal of this data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColLiteralSuffix"></a>ColLiteralSuffix

This VARCHAR(128) record field contains the character or characters that the driver recognizes as a suffix for a literal of this data type. This field contains an empty string for a data type for which a literal suffix is not applicable.

```
FUNCTION ColLiteralSuffix (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The character or characters that the driver recognizes as a suffix for a literal of this data type. 

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColLocalTypeName"></a>ColLocalTypeName

This VARCHAR(128) record field contains any localized (native language) name for the data type that may be different from the regular name of the data type. If there is no localized name, then an empty string is returned. This field is for display purposes only. The character set of the string is locale-dependent and is typically the default character set of the server.

```
FUNCTION ColLocalTypeName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The localized (native language) name for the data type that may be different from the regular name of the data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColName"></a>ColName

The column alias, if it applies. If the column alias does not apply, the column name is returned. In either case, SQL_DESC_UNNAMED is set to SQL_NAMED. If there is no column name or a column alias, an empty string is returned and SQL_DESC_UNNAMED is set to SQL_UNNAMED. This information is returned from the SQL_DESC_NAME record field of the IRD.

```
FUNCTION ColName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The column alias or name.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColNullable"></a>ColNullable

Returns SQL_NULLABLE if the column can have NULL values; SQL_NO_NULLS if the column does not have NULL values; or SQL_NULLABLE_UNKNOWN if it is not known whether the column accepts NULL values. This information is returned from the SQL_DESC_NULLABLE record field of the IRD.

```
FUNCTION ColNullable (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

SQL_NULLABLE, SQL_NO_NULLS or SQL_NULLABLE_UNKNOWN.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColNumPrecRadix"></a>ColNumPrecRadix

If the data type in the SQL_DESC_TYPE field is an approximate numeric data type, this SQLINTEGER field contains a value of 2 because the SQL_DESC_PRECISION field contains the number of bits. If the data type in the SQL_DESC_TYPE field is an exact numeric data type, this field contains a value of 10 because the SQL_DESC_PRECISION field contains the number of decimal digits. This field is set to 0 for all non-numeric data types.

```
FUNCTION ColNumPrecRadix (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The column numeric precision radix.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColOctetLength"></a>ColOctetLength

The length, in bytes, of a character string or binary data type. For fixed-length character or binary types, this is the actual length in bytes. For variable-length character or binary types, this is the maximum length in bytes. This value includes the null terminator. This information is returned from the SQL_DESC_OCTET_LENGTH record field of the IRD.

```
FUNCTION ColOctetLength (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The length, in bytes, of a character string or binary data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColPrecision"></a>ColPrecision

A numeric value that for a numeric data type denotes the applicable precision. For data types SQL_TYPE_TIME, SQL_TYPE_TIMESTAMP, and all the interval data types that represent a time interval, its value is the applicable precision of the fractional seconds component. This information is returned from the SQL_DESC_PRECISION record field of the IRD.

```
FUNCTION ColPrecision (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The column applicable precision.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColScale"></a>ColScale

A numeric value that is the applicable scale for a numeric data type. For DECIMAL and NUMERIC data types, this is the defined scale. It is undefined for all other data types. This information is returned from the SCALE record field of the IRD.

```
FUNCTION ColScale (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The column applicable scale for a numeric data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColSchemaName"></a>ColSchemaName

The schema of the table that contains the column. The returned value is implementation-defined if the column is an expression or if the column is part of a view. If the data source does not support schemas or the schema name cannot be determined, an empty string is returned. This VARCHAR record field is not limited to 128 characters.

```
FUNCTION ColSchemaName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The schema of the table that contains the column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColSearchable"></a>ColSearchable

Indicates if the column data type is searchable.

```
FUNCTION ColSearchable (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

SQL_PRED_NONE if the column cannot be used in a WHERE clause. (This is the same as the SQL_UNSEARCHABLE value in ODBC 2.x.)

SQL_PRED_CHAR if the column can be used in a WHERE clause but only with the LIKE predicate. (This is the same as the SQL_LIKE_ONLY value in ODBC 2.x.)

SQL_PRED_BASIC if the column can be used in a WHERE clause with all the comparison operators except LIKE. (This is the same as the SQL_EXCEPT_LIKE value in ODBC 2.x.)

SQL_PRED_SEARCHABLE if the column can be used in a WHERE clause with any comparison operator.

Columns of type SQL_LONGVARCHAR and SQL_LONGVARBINARY usually return SQL_PRED_CHAR.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColTableName"></a>ColTableName

The name of the table that contains the column. The returned value is implementation-defined if the column is an expression or if the column is part of a view. If the table name cannot be determined, an empty string is returned.

```
FUNCTION ColTableName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The name of the table that contains the column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColType"></a>ColType

A numeric value that specifies the SQL data type. When colNum is equal to 0, SQL_BINARY is returned for variable-length bookmarks and SQL_INTEGER is returned for fixed-length bookmarks. For the datetime and interval data types, this field returns the verbose data type: SQL_DATETIME or SQL_INTERVAL. This information is returned from the SQL_DESC_TYPE record field of the IRD.

```
FUNCTION ColType (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

A numeric value that specifies the SQL data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColTypeName"></a>ColTypeName

Data source-dependent data type name; for example, "CHAR", "VARCHAR", "MONEY", "LONG VARBINARY", or "CHAR ( ) FOR BIT DATA". If the type is unknown, an empty string is returned.

```
FUNCTION ColTypeName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

Data source-dependent data type name.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColUnnamed"></a>ColUnnamed

SQL_NAMED or SQL_UNNAMED. If the SQL_DESC_NAME field of the IRD contains a column alias or a column name, SQL_NAMED is returned. If there is no column name or column alias, SQL_UNNAMED is returned. This information is returned from the SQL_DESC_UNNAMED record field of the IRD.

```
FUNCTION ColUnnamed (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

SQL_NAMED or SQL_UNNAMED.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColUnsigned"></a>ColUnsigned

SQL_TRUE if the column is unsigned (or not numeric). SQL_FALSE if the column is signed.

```
FUNCTION ColUnsigned (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColUpdatable"></a>ColUpdatable

SQL_TRUE if the column is updatable; SQL_FALSE otherwise.

```
FUNCTION ColUpdatable (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="DbcHandle"></a>DbcHandle

Returns the connection handle.

```
FUNCTION DbcHandle () AS SQLHANDLE
```

# <a name="DeleteByBookmark"></a>DeleteByBookmark

Deletes a set of rows where each row is identified by a bookmark.

```
FUNCTION DeleteByBookmark () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="DeleteRecord"></a>DeleteRecord

The driver positions the cursor on the row specified by RowNumber and deletes the underlying row of data. It changes the corresponding element of the row status array to SQL_ROW_DELETED. After the row has been deleted, the following are not valid for the row: positioned update and delete statements, calls to **GetData** and calls to **SetPos** with *Operation* set to anything except SQL_POSITION. For drivers that support packing, the row is deleted from the cursor when new data is retrieved from the data source. Whether the row remains visible depends on the cursor type. For example, deleted rows are visible to static and keyset-driven cursors but invisible to dynamic cursors. The row operation array pointed to by the SQL_ATTR_ROW_OPERATION_PTR statement attribute can be used to indicate that a row in the current rowset should be ignored during a bulk delete.

```
FUNCTION DeleteRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RowNumber* | Row number inside the rowset. Note: *RowNumber* is the row number inside the rowset (if it is a single row rowset, RowNumber must be always 1). |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor

' // Bind the columns
DIM AS LONG lAuId, cbAuId
pStmt.BindCol(1, @lAuId, @cbAuId)
DIM wszAuthor AS WSTRING * 256, cbAuthor AS LONG
pStmt.BindCol(2, @wszAuthor, 256, @cbAuthor)
DIM iYearBorn AS SHORT, cbYearBorn AS LONG
pStmt.BindCol(3, @iYearBorn, @cbYearBorn)

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors WHERE Au_Id=999")

' // Fetch the record
pstmt.Fetch

' // Fill the values of the binded application variables and its sizes
cbAuID = SQL_COLUMN_IGNORE
wszAuthor = "Félix Lope de Vega Carpio"
cbAuthor = LEN(wszAuthor) * 2   ' Unicode uses 2 bytes per character
iYearBorn = 1562
cbYearBorn = SIZEOF(iYearBorn)

' // Delete the record
pStmt.DeleteRecord
IF pStmt.Error = FALSE THEN PRINT "Record deleted" ELSE PRINT pStmt.GetErrorInfo
```

# <a name="DescribeCol"></a>DescribeCol

Returns the result descriptor — column name, type, column size, decimal digits, and nullability — for one column in the result set. This information also is available in the fields of the IRD. 

```
FUNCTION DescribeCol (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL pwszColumnName AS WSTRING PTR, _
   BYVAL BufferLength AS SQLSMALLINT, BYVAL NameLength AS SQLSMALLINT PTR, _
   BYVAL DataType AS SQLSMALLINT PTR, BYVAL ColumnSizePtr AS SQLULEN PTR, _
   BYVAL DecimalDigits AS SQLSMALLINT PTR, BYVAL Nullable AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNumber* | Column number of result data, ordered sequentially in increasing column order, starting at 1. The *ColNumber* argument can also be set to 0 to describe the bookmark column. |
| *ColName* | Pointer to a null-terminated buffer in which to return the column name. This value is read from the SQL_DESC_NAME field of the IRD. If the column is unnamed or the column name cannot be determined, the driver returns an empty string. |
| *DataType* | Pointer to a buffer in which to return the SQL data type of the column. This value is read from the SQL_DESC_CONCISE_TYPE field of the IRD. This will be one of the values in SQL Data Types, or a driver-specific SQL data type. If the data type cannot be determined, the driver returns SQL_UNKNOWN_TYPE.<br>In ODBC 3.x, SQL_TYPE_DATE, SQL_TYPE_TIME, or SQL_TYPE_TIMESTAMP is returned in DataType for date, time, or timestamp data, respectively; in ODBC 2.x, SQL_DATE, SQL_TIME, or SQL_TIMESTAMP is returned. The Driver Manager performs the required mappings when an ODBC 2.x application is working with an ODBC 3.x driver or when an ODBC 3.x application is working with an ODBC 2.x driver.<br>When ColNumber is equal to 0 (for a bookmark column), SQL_BINARY is returned in DataType for variable-length bookmarks. (SQL_INTEGER is returned if bookmarks are used by an ODBC 3.x application working with an ODBC 2.x driver or by an ODBC 2.x application working with an ODBC 3.x driver.) |
| *ColSize* | Pointer to a buffer in which to return the size (in characters) of the column on the data source. If the column size cannot be determined, the driver returns 0. |
| *DecimalDigits* | Pointer to a buffer in which to return the number of decimal digits of the column on the data source. If the number of decimal digits cannot be determined or is not applicable, the driver returns 0. |
| *Nullable* | Pointer to a buffer in which to return a value that indicates whether the column allows NULL values. This value is read from the SQL_DESC_NULLABLE field of the IRD. The value is one of the following:<br>SQL_NO_NULLS: The column does not allow NULL values.<br>SQL_NULLABLE: The column allows NULL values.<br>QL_NULLABLE_UNKNOWN: The driver cannot determine if the column allows NULL values.  |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor

' // Generate a result set
pStmt.ExecDirect ("SELECT TOP 20 * FROM Authors ORDER BY Author")

' -------------------------------------------------------------------------------------
' Use DescribeCol to retrieve information about column 2
' -------------------------------------------------------------------------------------
DIM wszColName AS WSTRING * 260
DIM iNameLength AS SHORT
DIM iDataType AS SHORT
DIM dwColumnSize AS DWORD
DIM iDecimalDigits AS SHORT
DIM iNullable AS SHORT

pStmt.DescribeCol(2, @wszColName, 260, @iNameLength, @iDataType, @dwColumnSize, @iDecimalDigits, @iNullable)

? "Column name: " & wszColName
? "Name length: " & STR(iNameLength)
? "Data type: " & STR(iDataType)
? "Column size: " & STR(dwColumnSize)
? "Decimal digits: " & STR(iDecimalDigits)
? "Nullable: " & STR(iNullable) & " - " & IIF(iNullable, "TRUE", "FALSE")
```

# <a name="DescribeParam"></a>DescribeParam

Returns the description of a parameter marker associated with a prepared SQL statement. This information is also available in the fields of the IPD.

```
FUNCTION DescribeParam (BYVAL ParameterNumber AS SQLUSMALLINT, BYVAL DataType AS SQLSMALLINT PTR, _
   BYVAL ParameterSize AS SQLULEN PTR, BYVAL DecimalDigits AS SQLSMALLINT PTR, _
   BYVAL Nullable AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ParameterNumber* | Parameter marker number ordered sequentially in increasing parameter order, starting at 1. |
| *DataType* | Pointer to a buffer in which to return the SQL data type of the parameter. This value is read from the SQL_DESC_CONCISE_TYPE record field of the IPD.<br>In ODBC 3.x, SQL_TYPE_DATE, SQL_TYPE_TIME, or SQL_TYPE_TIMESTAMP will be returned in DataType for date, time, or timestamp data, respectively; in ODBC 2.x, SQL_DATE, SQL_TIME, or SQL_TIMESTAMP will be returned. The Driver Manager performs the required mappings when an ODBC 2.x application is working with an ODBC 3.x driver or when an ODBC 3.x application is working with an ODBC 2.x driver.<br>When *ColumnNumber* is equal to 0 (for a bookmark column), SQL_BINARY is returned in *DataType* for variable-length bookmarks. (SQL_INTEGER is returned if bookmarks are used by an ODBC 3.x application working with an ODBC 2.x driver or by an ODBC 2.x application working with an ODBC 3.x driver.) |
| *ParameterSize* | Pointer to a buffer in which to return the size of the column or expression of the corresponding parameter marker as defined by the data source. |
| *DecimalDigits* | Pointer to a buffer in which to return the number of decimal digits of the column or expression of the corresponding parameter as defined by the data source. |
| *Nullable* | Pointer to a buffer in which to return a value that indicates whether the parameter allows NULL values. This value is read from the SQL_DESC_NULLABLE field of the IPD. One of the following:<br>SQL_NO_NULLS: The parameter does not allow NULL values (this is the default value).<br>SQL_NULLABLE: The parameter allows NULL values.<br>SQL_NULLABLE_UNKNOWN: The driver cannot determine if the parameter allows NULL values. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ExecDirect"></a>ExecDirect

Executes a preparable statement, using the current values of the parameter marker variables if any parameters exist in the statement. **ExecDirect** is the fastest way to submit an SQL statement for one-time execution.

```
FUNCTION ExecDirect (BYREF wszSqlStr AS WSTRING) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSqlStr* | SQL statement to be executed. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

# <a name="Execute"></a>Execute

Executes a prepared statement, using the current values of the parameter marker variables if any parameter markers exist in the statement.

```
FUNCTION Execute () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

# <a name="ExtendedFetch"></a>ExtendedFetch

Fetches the specified rowset of data from the result set and returns data for all bound columns. Rowsets can be specified at an absolute or relative position or by bookmark.

**Note**: In ODBC 3.x, **ExtendedFetch** has been replaced by FetchScroll. ODBC 3.x applications should not call **ExtendedFetch**; instead they should call **FetchScroll**. The Driver Manager maps **FetchScroll** to **ExtendedFetch** when working with an ODBC 2.x driver. 

```
FUNCTION ExtendedFetch (BYVAL FetchOrientation AS SQLUSMALLINT, BYVAL FetchOffset AS SQLLEN, _
   BYVAL RowCountPtr AS ANY PTR, BYVAL RowStatusArray AS SQLUSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *FetchOrientation* | Type of fetch. This is the same as *FetchOrientation* in **FetchScroll**. |
| *FetchOffset* | Number of the row to fetch. This is the same as *FetchOffset* in **FetchScroll**, with one exception. When *FetchOrientation* is SQL_FETCH_BOOKMARK, *FetchOffset* is a fixed-length bookmark, not an offset from a bookmark. In other words, **ExtendedFetch** retrieves the bookmark from this argument, not the SQL_ATTR_FETCH_BOOKMARK_PTR statement attribute. It does not support variable-length bookmarks and does not support fetching a rowset at an offset (other than 0) from a bookmark. |
| *RowCountPtr* | Pointer to a buffer in which to return the number of rows actually fetched. This buffer is used in the same manner as the buffer specified by the SQL_ATTR_ROWS_FETCHED_PTR statement attribute. This buffer is used only by ExtendedFetch. It is not used by **Fetch** or **FetchScroll**. |
| *RowStatusArray* | Pointer to an array in which to return the status of each row. This array is used in the same manner as the array specified by the SQL_ATTR_ROW_STATUS_PTR statement attribute.<br>However, the address of this array is not stored in the SQL_DESC_STATUS_ARRAY_PTR field in the IRD. Furthermore, this array is used only by **ExtendedFetch** and by **BulkOperations** with an *Operation* of SQL_ADD or **SetPos** when it is called after **ExtendedFetch**. It is not used by **Fetch** or **FetchScroll**, and it is not used by **BulkOperations** or **SetPos** when they are called after **Fetch** or **FetchScroll**. It is also not used when **BulkOperations** with an *Operation* of SQL_ADD is called before any fetch function is called. In other words, it is used only in statement state S7. It is not used in statement states S5 or S6.<br>Applications should provide a valid pointer in the **RowStatusArray* argument; if not, the behavior of **ExtendedFetch** and the behavior of calls to **BulkOperations** or **SetPos** after a cursor has been positioned by **ExtendedFetch** are undefined. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="Fetch"></a>Fetch

Fetches the next rowset of data from the result set and returns data for all bound columns. 

```
FUNCTION Fetch () AS BOOLEAN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="FetchByBookmark"></a>FetchByBookmark

Fetches a set of rows where each row is identified by a bookmark.

```
FUNCTION FetchByBookmark () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="FetchScroll"></a>FetchScroll

Fetches the specified rowset of data from the result set and returns data for all bound columns. Rowsets can be specified at an absolute or relative position or by bookmark.

When working with an ODBC 2.x driver, the Driver Manager maps this function to **ExtendedFetch**. 

```
FUNCTION FetchScroll (BYVAL FetchOrientation AS SQLSMALLINT, BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *FetchOrientation* | Type of fetch: SQL_FETCH_NEXT, SQL_FETCH_PRIOR, SQL_FETCH_FIRST, SQL_FETCH_LAST, SQL_FETCH_ABSOLUTE, SQL_FETCH_RELATIVE, SQL_FETCH_BOOKMARK |
| *FetchOffset* | Number of the row to fetch. The interpretation of this argument depends on the value of the *FetchOrientation* argument. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetColumnPrivileges"></a>GetColumnPrivileges

Returns a list of columns and associated privileges for the specified table. The driver returns the information as a result set on the specified statement handle.

```
FUNCTION GetColumnPrivileges (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
   BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, _
   BYREF wszTableName AS WSTRING, BYVAL TableNameLength AS SQLSMALLINT, _
   BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCatalogName* | Catalog name. If a driver supports names for some catalogs but not for others,such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those catalogs that do not have names. *wszCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | Schema name. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have schemas. *wszSchemaName* cannot contain a string search pattern. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszTableName* | Table name. This argument cannot be a null pointer. *wszTableName* cannot contain a string search pattern. |
| *TableNameLength* | Length of *wszTableName*. |
| *wszColumnName* | String search pattern for column names. |
| *ColumnNameLength* | Length of *wszColumnName*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **ColumnPrivileges** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value may be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetColumns"></a>GetColumns

Returns the list of column names in specified tables. The driver returns this information as a result set on the specified statement handle.

```
FUNCTION GetColumns (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
   BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, _
   BYREF wszTableName AS WSTRING, BYVAL TableNameLength AS SQLSMALLINT, _
   BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCatalogName* | Catalog name. If a driver supports names for some catalogs but not for others,such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those catalogs that do not have names. *wszCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | String search pattern for schema names. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") indicates those tables that do not have schemas. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszTableName* | String search pattern for table names. |
| *TableNameLength* | Length of *wszTableName*. |
| *wszColumnName* | String search pattern for column names. |
| *ColumnNameLength* | Length of *wszColumnName*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **Columns** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetCursorConcurrency"></a>GetCursorConcurrency

Gets a SQLUINTEGER value that specifies the cursor concurrency.

```
FUNCTION GetCursorConcurrency () AS SQLUINTEGER
```

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetCursorKeysetSize"></a>GetCursorKeysetSize

Gets a SQLUINTEGER value that specifies the number of rows in the keyset-driven cursor.

```
FUNCTION GetCursorKeysetSize () AS SQLUINTEGER
```

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="GetCursorName"></a>GetCursorName

Returns the cursor name associated with a specified statement.

```
FUNCTION GetCursorName () AS CWSTR
```

#### Remarks

For efficient processing, the cursor name should not include any leading or trailing spaces in the cursor name, and if the cursor name includes a delimited identifier, the delimiter should be positioned as the first character in the cursor name.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetCursorScrollability"></a>GetCursorScrollability

Gets a SQLUINTEGER value that specifies the scrollability type. Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION GetCursorScrollability () AS SQLUINTEGER
```

#### Return value

The scrollability type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetCursorSensitivity"></a>GetCursorSensitivity

Gets a SQLUINTEGER value that specifies whether cursors on the statement handle made to a result set by another cursor. Setting this attribute affects subsequent calls to ExecDirect and SQLExecute. An application can read back the value of this attribute to obtain its initial state or its state  as most recently set by the application.

**Note**: Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION GetCursorSensitivity () AS SQLUINTEGER
```

#### Return value

The type of cursor sensitivity.

**SQL_UNSPECIFIED** = It is unspecified what the cursor type is and whether  cursors on the statement handle make visible the changes made to a result set by another cursor. Cursors on the statement handle may make visible none, some, or all such changes. This is the default.

**SQL_INSENSITIVE** = All cursors on the statement handle show the result set without reflecting any changes made to it by any other cursor, which has a concurrency that is read-only.

**SQL_SENSITIVE** = All cursors on the statement handle make visible all changes made to a result set by another cursor.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetCursorType"></a>GetCursorType

Gets SQLUINTEGER value that specifies the cursor type.

**Note**: Microsoft Access Driver changes SQL_CURSOR_DYNAMIC to SQL_CURSOR_KEYSET_DRIVEN.

```
FUNCTION GetCursorType () AS SQLUINTEGER
```

#### Return value

The type of cursor.

**SQL_CURSOR_FORWARD_ONLY** = The cursor only scrolls forward.

**SQL_CURSOR_STATIC** = The data in the result set is static.

**SQL_CURSOR_KEYSET_DRIVEN** = The driver saves and uses only the keys for the number of rows specified in the SQL_ATTR_KEYSET_SIZE statement attribute.

**SQL_CURSOR_DYNAMIC** = The driver saves and uses only the keys for the rows in the rowset.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="GetData"></a>GetData

Retrieves data for a single column in the result set. It can be called multiple times to retrieve variable-length data in parts.

```
FUNCTION GetData (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL TargetType AS SQLSMALLINT, _
   BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, _
   BYVAL StrLen_or_Ind AS SQLLEN PTR) AS SQLRETURN
FUNCTION GetData (BYREF ColumnName AS WSTRING, BYVAL TargetType AS SQLSMALLINT, _
   BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, _
   BYVAL StrLen_or_Ind AS SQLLEN PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColumnNumber* | Number or name of the column for which to return data. Result set columns are numbered in increasing column order starting at 1. The bookmark column is column number 0; this can be specified only if bookmarks are enabled. |
| *TargetType* | The type identifier of the C data type of the *TargetValuePtr* buffer. If *TargetType* is SQL_ARD_TYPE, the driver uses the type identifier specified in the SQL_DESC_CONCISE_TYPE field of the ARD. If it is SQL_C_DEFAULT, the driver selects the default C data type based on the SQL data type of the source. |
| *TargetValue* | Pointer to the buffer in which to return the data. |
| *BufferLength* | Length of the *TargetValue* buffer in bytes.<br>The driver uses *BufferLength* to avoid writing past the end of the TargetValue buffer when returning variable-length data, such as character or binary data. Note that the driver counts the null-termination character when returning character data to *TargetValue*. *TargetValue* must therefore contain space for the null-termination character, or the driver will truncate the data.<br>When the driver returns fixed-length data, such as an integer or a date structure, the driver ignores *BufferLength* and assumes the buffer is large enough to hold the data. It is therefore important for the application to allocate a large enough buffer for fixed-length data or the driver will write past the end of the buffer.<br>**GetData** returns SQLSTATE HY090 (Invalid string or buffer length) when *BufferLength* is less than 0 but not when *BufferLength* is 0.<br>If *TargetValue* is set to a null pointer,*BufferLength* is ignored by the driver.  |
| *StrLen_or_Ind* | Pointer to the buffer in which to return the length or indicator value. If this is a null pointer, no length or indicator value is returned. This returns an error when the data being fetched is NULL.<br>**GetData** can return the following values in the length/indicator buffer:<br>The length of the data available to return: SQL_NO_TOTAL, SQL_NULL_DATA. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

### Overloaded Functions

Retrieves data for a single column in the result set. It can be called multiple times to retrieve variable-length data in parts.

```
FUNCTION GetData (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL maxChars AS LONG = 256 * 2) AS CWSTR
FUNCTION GetData (BYREF ColumnName AS WSTRING, BYVAL maxChars AS LONG = 256 * 2) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColumnNumber* | Number or name of the column for which to return data. Result set columns are numbered in increasing column order starting at 1. The bookmark column is column number 0; this can be specified only if bookmarks are enabled. |
| *BufferLength* | Length of the *TargetValue* buffer in bytes.<br>The driver uses *BufferLength* to avoid writing past the end of the *TargetValue* buffer when returning variable-length data, such as character or binary data. Note that the driver counts the null-termination character when returning character data to *TargetValue*. *TargetValue* must therefore contain space for the null-termination character, or the driver will truncate the data.<br>When the driver returns fixed-length data, such as an integer or a date structure, the driver ignores *BufferLength* and assumes the buffer is large enough to hold the data. It is therefore important for the application to allocate a large enough buffer for fixed-length data or the driver will write past the end of the buffer.<br>**GetData** returns SQLSTATE HY090 (Invalid string or buffer length) when *BufferLength* is less than 0 but not when *BufferLength* is 0. If *TargetValue* is set to a null pointer, *BufferLength* is ignored by the driver.  |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetDiagField"></a>GetDiagField

Returns the current value of a field of a record of the diagnostic data structure (associated with an statement handle) that contains error, warning, and status information.

```
FUNCTION GetDiagField (BYVAL RecNumber AS SQLSMALLINT, BYVAL DiagIdentifier AS SQLSMALLINT, _
   BYVAL DiagInfo AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, _
   BYVAL StringLength AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the status record from which the application seeks information. Status records are numbered from 1. If the *DiagIdentifier* argument indicates any field of the diagnostics header, *RecNumber* is ignored. If not, it should be greater than 0. |
| *DiagIdentifier* | Indicates the field of the diagnostic whose value is to be returned. |
| *DiagInfo* | Pointer to a buffer in which to return the diagnostic information. The data type depends on the value of *DiagIdentifier*. |
| *BufferLength* | If *DiagIdentifier* is an ODBC-defined diagnostic and DiagInfoPtr points to a character string or a binary buffer, this argument should be the length of *DiagInfo*. If *DiagIdentifier* is an ODBC-defined field and *DiagInfo* is an integer, *BufferLength* is ignored.<br>If *DiagIdentifier* is a driver-defined field, the application indicates the nature of the field to the Driver Manager by setting the *BufferLength* argument. *BufferLength* can have the following values:<ul><li>If *DiagInfo* is a pointer to a character string, then *BufferLength* is the length of the string or SQL_NTS.</li><li>If *DiagInfo* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *BufferLength*. This places a negative value in *BufferLength*.</li><li>If *DiagInfoPtr* is a pointer to a value other than a character string or binary string, then *BufferLength* should have the value SQL_IS_POINTER.</li><li>If *DiagInfoPtr* is contains a fixed-length data type, then *BufferLength* is SQL_IS_INTEGER, SQL_IS_UINTEGER, SQL_IS_SMALLINT, or SQL_IS_USMALLINT, as appropriate.</li></ul> |
| *StringLength* | Pointer to a buffer in which to return the total number of bytes (excluding the number of bytes required for the null-termination character) available to return in *DiagInfo*, for character data. If the number of bytes available to return is greater than or equal to *BufferLength*, the text in *DiagInfo* is truncated to *BufferLength* minus the length of a null-termination character. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_INVALID_HANDLE, or SQL_NO_DATA.

# <a name="GetDiagRec"></a>GetDiagRec

Returns the current values of multiple fields of a diagnostic record that contains error, warning, and status information. Unlike **GetDiagField**, which returns one diagnostic field per call, **GetDiagRec** returns several commonly used fields of a diagnostic record, including the SQLSTATE, the native error code, and the diagnostic message text.

```
FUNCTION GetDiagRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL Sqlstate AS WSTRING PTR, _
   BYVAL NativeError AS SQLINTEGER PTR, BYVAL MessageText AS WSTRING PTR, _
   BYVAL BufferLength AS SQLSMALLINT, BYVAL TextLength AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the status record from which the application seeks information. Status records are numbered from 1. |
| *SQLState* | Pointer to a buffer in which to return a five-character SQLSTATE code pertaining to the diagnostic record RecNumber. The first two characters indicate the class; the next three indicate the subclass. This information is contained in the SQL_DIAG_SQLSTATE diagnostic field. |
| *NativeError* | Pointer to a buffer in which to return the native error code, specific to the data source. This information is contained in the SQL_DIAG_NATIVE diagnostic field. |
| *MessageText* | Pointer to a buffer in which to return the diagnostic message text string. This information is contained in the SQL_DIAG_MESSAGE_TEXT diagnostic field. |
| *BufferLength* | Length of the *MessageText* buffer in characters. There is no maximum length of the diagnostic message text. |
| *TextLength* | Pointer to a buffer in which to return the total number of bytes (excluding the number of bytes required for the null-termination character) available to return in *MessageText*. If the number of bytes available to return is greater than *BufferLength*, the diagnostic message text in *MessageText* is truncated to *BufferLength* minus the length of a null-termination character. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetErrorInfo"></a>GetErrorInfo

Returns a verbose description of the last error(s).

```
FUNCTION GetErrorInfo (BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *iErrorCode* | Optional. The error code returned by **GetLastResult**. |

# <a name="GetForeignKeys"></a>GetForeignKeys

ForeignKeys can return:

* A list of foreign keys in the specified table (columns in the specified table that refer to primary keys in other tables).
* A list of foreign keys in other tables that refer to the primary key in the specified table.

```
FUNCTION GetForeignKeys (BYREF wszPkCatalogName AS WSTRING, BYVAL pkCatalogNameLength AS SQLSMALLINT, _
   BYREF wszPkSchemaName AS WSTRING, BYVAL pkSchemaNameLength AS SQLSMALLINT, _
   BYREF wszPkTableName AS WSTRING, BYVAL pkTableNameLength AS SQLSMALLINT, _
   BYREF wszFkCatalogName AS WSTRING, BYVAL FkCatalogNameLength AS SQLSMALLINT, _
   BYREF wszFkSchemaName AS WSTRING, BYVAL FkSchemaNameLength AS SQLSMALLINT, _
   BYREF wszFkTableName AS WSTRING, BYVAL fkTableNameLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPkCatalogName* | Primary key table catalog name. If a driver supports catalogs for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have catalogs. *wszPkCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszPkCatalogName*. |
| *wszPkSchemaName* | Primary key table schema name. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have schemas. *wszPKSchemaName* cannot contain a string search pattern. |
| *PkSchemaNameLength* | Length of *wszPkSchemaName*. |
| *wszPkTableName* | Primary key table name. *wszPkTableName* cannot contain a string search pattern. |
| *PkTableNameLength* | Length of *wszPkTableName*. |
| *wszFkCatalogName* | Foreign key table catalog name. If a driver supports catalogs for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have catalogs. *wszFkCatalogName* cannot contain a string search pattern. |
| *FkCatalogNameLength* | Length of *wszFkCatalogName*. |
| *wszFkSchemaName* | Foreign key table schema name. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have schemas. *wszFkSchemaName* cannot contain a string search pattern. |
| *FkSchemaNameLength* | Length of *wszFkSchemaName*. |
| *wszFkTableName* | Foreign key table name. *wszFkTableName* cannot contain a string search pattern. |
| *FkTableNameLength* | Length of *wszFkTableName*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **ForeignKeys** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetImpParamDesc"></a>GetImpParamDesc

Returns the handle to the IPD (Implicitily Parameter Descriptor). The value of this attribute is the descriptor allocated when the statement was initially allocated. The application cannot set this attribute. This attribute can be retrieved by a call to **GetStmtAttr** but not set by a call to **SetStmtAttr**.

```
FUNCTION GetImpParamDesc () AS SQLUINTEGER
```

#### Return value

The handle of IPD.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetImpParamDescField"></a>GetImpParamDescField

Returns the current setting or value of a single field of a descriptor record. The field returned describe the name, data type, and storage of column or parameter data.

```
FUNCTION GetImpParamDescField (BYVAL RecNumber AS SQLSMALLINT, BYVAL FieldIdentifier AS SQLSMALLINT, _
   BYVAL ValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLINTEGER, _
   BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. If *RecNumber* is less that or equal to SQL_DESC_COUNT but the row does not contain data for a column or parameter, a call to **GetImpRowDesc** will return the default values of the fields. |
| *FieldIdentifier* | Indicates the field of the descriptor whose value is to be returned: SQL_DESC_NAME, SQL_DESC_TYPE, SQL_DESC_OCTET_LENGTH, SQL_DESC_PRECISION, SQL_DESC_SCALE, SQL_DESC_NULLABLE. |
| *ValuePtr* | Pointer to a buffer in which to return the descriptor information. The data type depends on the value of *FieldIdentifier*. |
| *BufferLength* | If *FieldIdentifier* is an ODBC-defined field and *ValuePtr* points to a character string or a binary buffer, this argument should be the length of *ValuePtr*. If *FieldIdentifier* is an ODBC-defined field and *ValuePtr* is an integer, *BufferLength* is ignored.<br><br>If *FieldIdentifier* is a driver-defined field, the application indicates the nature of the field to the Driver Manager by setting the *BufferLength* argument. *BufferLength* can have the following values:<br><br><ul><li>If *ValuePtr* is a pointer to a character string, then *BufferLength* is the length of the string or SQL_NTS.</li><li>If *ValuePtr* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *BufferLength*. This places a negative value in *BufferLength*.</li><li>If *ValuePtr* is a pointer to a value other than a character string or binary string, then *BufferLength* should have the value SQL_IS_POINTER.</li><li>If *ValuePtr* is contains a fixed-length data type, then *BufferLength* is either SQL_IS_INTEGER, SQL_IS_UINTEGER, SQL_IS_SMALLINT, or SQL_IS_USMALLINT, as appropriate.</li>/ul> |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldName"></a>GetImpParamDescFieldName

Returns the name of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldName (BYVAL RecNumber AS SQLSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The name of the field.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if RecNumber is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldNullable"></a>GetImpParamDescFieldNullable

Returns the nullable value (TRUE or FALSE) of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldNullable (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

TRUE or FALSE.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldOctetLength"></a>GetImpParamDescFieldOctetLength

Returns the octet length of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldOctetLength (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The octet length.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldPrecision"></a>GetImpParamDescFieldPrecision

Returns the precision value of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldPrecision (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The precision value.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldScale"></a>GetImpParamDescFieldScale

Returns the scale value of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldScale (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The scale value.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldType"></a>GetImpParamDescFieldType

Returns the type of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldType (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The type of the field.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescRec"></a>GetImpParamDescRec

Returns the current settings or values of multiple fields of a descriptor record. The fields returned describe the name, data type, and storage of column or parameter data.

```
FUNCTION GetImpParamDescRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL pwszName AS WSTRING PTR, _
   BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR, _
   BYVAL TypePtr AS SQLSMALLINT PTR, BYVAL SubTypePtr AS SQLSMALLINT PTR, _
   BYVAL LengthPtr AS SQLLEN PTR, BYVAL PrecisionPtr AS SQLSMALLINT PTR, _
   BYVAL ScalePtr AS SQLSMALLINT PTR, BYVAL NullablePtr AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less than or equal to the value of SQL_DESC_COUNT. If RecNumber is less than or equal to SQL_DESC_COUNT but the row does not contain data for a column or parameter, a call to **GetDescRec** will return the default values of the fields.ates the descriptor record from which the application seeks information. Descriptor records are numbered.  |
| *pwszName* | A pointer to a buffer in which to return the SQL_DESC_NAME field for the descriptor record. |
| *TypePtr* | A pointer to a buffer in which to return the value of the SQL_DESC_TYPE field for the descriptor record. |
| *SubTypePtr* | For records whose type is SQL_DATETIME or SQL_INTERVAL, this is a pointer to a buffer in which to return the value of the SQL_DESC_DATETIME_INTERVAL_CODE field. |
| *LengthPtr* | A pointer to a buffer in which to return the value of the SQL_DESC_OCTET_LENGTH field for the descriptor record. |
| *PrecisionPtr* | A pointer to a buffer in which to return the value of the SQL_DESC_PRECISION field for the descriptor record. |
| *ScalePtr* | A pointer to a buffer in which to return the value of the SQL_DESC_SCALE field for the descriptor record. |
| *NullablePtr* | A pointer to a buffer in which to return the value of the SQL_DESC_NULLABLE field for the descriptor record. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpRowDesc"></a>GetImpRowDesc

Returns the handle to the IRD. The value of this attribute is the descriptor  allocated when the statement was initially allocated. The application  cannot set this attribute.

This attribute can be retrieved by a call to **GetStmtAttr** but not set by a call to **SetStmtAttr**.

```
FUNCTION GetImpRowDesc () AS SQLUINTEGER
```

#### Return value

The handle of the IRD.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetImpRowDescField"></a>GetImpRowDescField

Returns the current setting or value of a single field of a descriptor record. The field returned describe the name, data type, and storage of column or parameter data.

```
FUNCTION GetImpRowDescField (BYVAL RecNumber AS SQLSMALLINT, BYVAL FieldIdentifier AS SQLSMALLINT, _
   BYVAL ValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLINTEGER, _
   BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The RecNumber argument must be less or equal to the value of SQL_DESC_COUNT. If RecNumber is less that or equal to SQL_DESC_COUNT but the row does not contain data for a column or parameter, a call to GetImpRowDesc will return the default values of the fields.  |
| *FieldIdentifier* | Indicates the field of the descriptor whose value is to be returned: SQL_DESC_NAME, SQL_DESC_TYPE, SQL_DESC_OCTET_LENGTH, SQL_DESC_PRECISION, SQL_DESC_SCALE, SQL_DESC_NULLABLE. |
| *ValuePtr* | Pointer to a buffer in which to return the descriptor information. The data type depends on the value of *FieldIdentifier*. |
| *BufferLength* | If *FieldIdentifier* is an ODBC-defined field and ValuePtr points to a character string or a binary buffer, this argument should be the length of *ValuePtr*. If *FieldIdentifier* is an ODBC-defined field and ValuePtr is an integer, *BufferLength* is ignored.<br>If *FieldIdentifier* is a driver-defined field, the application indicates the nature of the field to the Driver Manager by setting the BufferLength argument. *BufferLength* can have the following values:<br><br><ul><li>If *ValuePtr* is a pointer to a character string, then  BufferLengthis the length of the string or SQL_NTS.</li><li>If *ValuePtr* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *BufferLength*. This places a negative value in *BufferLength*.</li><li>If *ValuePtr* is a pointer to a value other than a character string or binary string, then *BufferLength* should have the value SQL_IS_POINTER.</li><li>If *ValuePtr* is contains a fixed-length data type, then *BufferLength* is either SQL_IS_INTEGER, SQL_IS_UINTEGER, SQL_IS_SMALLINT, or SQL_IS_USMALLINT, as appropriate.</li></ul> |

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if RecNumber is greater than the current number of descriptor records.
