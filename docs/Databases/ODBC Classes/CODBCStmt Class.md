# CODBCStmt Class

Implements methods to create and manage statement objects. Inherits from CODBCBase.

**Include file**: CODBCStmt.inc

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

