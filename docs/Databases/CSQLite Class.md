# CSQLite Class

SQLite is an embedded SQL database engine. Unlike most other SQL databases, SQLite does not have a separate server process. SQLite reads and writes directly to ordinary disk files. A complete SQL database with multiple tables, indices, triggers, and views, is contained in a single disk file. The database file format is cross-platform - you can freely copy a database between 32-bit and 64-bit systems or between big-endian and little-endian architectures. These features make SQLite a popular choice as an Application File Format.

**Include file**: CSQLite.inc

### Constructor

Creates a new instance of the **CSQlite** base class and loads the SQLite3 DLL.

```
CONSTRUCTOR CSQLite (BYREF wszDllPath AS WSTRING = "sqlite3.dll")
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszDllPath* | The path of the SQLite3 DLL. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [CompileOptionUsed](#CompileOptionUsed) | Determines whether the specified option was defined at compile time. |
| [Complete](#Complete) | Determines if the currently entered text seems to form a complete SQL statement. |
| [EnableSharedCache](#EnableSharedCache) | Enables or disables the sharing of the database cache and schema data structures between connections to the same database. |
| [ErrStr](#ErrStr) | Returns English-language text that describes the result code. |
| [Free](#Free) | Releases memory previously allocated by **Malloc** or **Realloc**. |
| [GetCompileOption](#GetCompileOption) | Allows iterating over the list of options that were defined at compile time by returning the N-th compile time option string. |
| [GetLastResult](#GetLastResult) | Returns the last result code. |
| [Malloc](#Malloc) | Returns a pointer to a block of memory at least N bytes in length, where N is the parameter. |
| [MemoryHighwater](#MemoryHighwater) | Returns the maximum value of **MemoryUsed** since the high-water mark was last reset. |
| [MemorySize](#MemorySize) | Returns the size of that memory allocation in bytes. |
| [MemoryUsed](#MemoryUsed) | Returns the number of bytes of memory currently outstanding (malloced but not freed). |
| [Randomness](#Randomness) | Pseudo-random number generator. |
| [Realloc](#Realloc) | Attempts to resize a prior memory allocation to be at least the specfified number of bytes. |
| [ReleaseMemory](#ReleaseMemory) | Attempts to free the specified number of bytes of heap memory by deallocating non-essential memory allocations held by the database library. |
| [Sleep](#Sleep) | Causes the current thread to suspend execution for at least a number of milliseconds specified in its parameter. |
| [SoftHeapLimit64](#SoftHeapLimit64) | Sets and/or queries the soft limit on the amount of heap memory that may be allocated by SQLite. |
| [SourceID](#SourceID) | Returns the SQLite3 source identifier. |
| [Status](#Status) | Retrieves runtime status information about the performance of SQLite, and optionally to reset various highwater marks. |
| [StrGlob](#StrGlob) | The **StrGlob** method returns zero if string *szStr* matches the glob pattern *szGlob*, and it returns non-zero if string *szStr* does not match the glob pattern *szGlob*. This function is case sensitive. |
| [ThreadSafe](#ThreadSafe) | Returns zero if and only if SQLite was compiled with mutexing code omitted due to the SQLITE_THREADSAFE compile-time option being set to 0. |
| [Version](#Version) | Returns the SQLite3 version. |
| [VersionNumber](#VersionNumber) | Returns the SQLite3 version number. |

# CSQLiteDb Class

Inherits from CSQLite.

### Constructor

Opens an SQLite database file as specified by the filename argument.

```
CONSTRUCTOR CSQLiteDb (BYREF wszPath AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The name of the database. |

#### Remarks

The database is opened for reading and writing, and is created if it does not already exist. If the filename is ":memory:", then a private, temporary in-memory database is created for the connection. This in-memory database will vanish when the database connection is closed. If the filename is an empty string, then a private, temporary on-disk database will be created. This private database will be automatically deleted as soon as the database connection is closed. If URI filename interpretation is enabled, and the filename argument begins with "file:", then the filename is interpreted as a URI.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Changes](#Changes) | Returns the number of database rows that were changed or inserted or deleted by the most recently completed SQL statement on the current database connection. |
| [CloseDb](#CloseDb) | Closes the database. |
| [ErrCode](#ErrCode) | Returns the numeric result code for the most recent failed sqlite3 call associated with a database connection. |
| [ErrMsg](#ErrMsg) | Returns English-language text that describes the error. |
| [Exec](#Exec) | Convenience wrapper for **Prepare** and **Step_**. |
| [ExtendedErrCode](#ExtendedErrCode) | Gets the extended error code associated with this dabatase connection. |
| [ExtendedResultCodes](#ExtendedResultCodes) | Enables or disables the extended result codes feature of SQLite. |
| [hDbc](#hDbc) | Gets/sets the database handle. |
| [Interrupt](#Interrupt) | This function causes any pending database operation to abort and return at its earliest opportunity. |
| [LastInsertRowId](#LastInsertRowId) | Returns the rowid of the most recent successful INSERT into the database from the database connection in the first argument. |
| [Limit](#Limit) | This function allows the size of various constructs to be limited on a connection by connection basis. |
| [OpenBlob](#OpenBlob) | Opens a handle to a BLOB. |
| [OpenDb](#OpenDb) | Opens an SQLite database file as specified by the filename argument. |
| [Prepare](#Prepare) | Creates a new prepared statement object. |
| [ProgressHandler](#ProgressHandler) | The **ProgressHandler** method causes a callback function to be invoked periodically during long running calls to **Step_** and **GetRow** for a database connection. |
| [ReleaseMemory](#ReleaseMemory) | Attempts to free as much heap memory as possible from the specified database connection. |
| [Status](#Status) | Retrieves runtime status information about a single database connection. |
| [TotalChanges](#TotalChanges) | This function returns the number of row changes caused by INSERT, UPDATE or DELETE statements since the database connection was opened. |
| [UnlockNotify](#UnlockNotify) | Registers a callback that SQLite will invoke when the connection currently holding the required lock relinquishes it. |

# CSQLiteStmt Class

Wraps a sqlite3_stmt pointer returned by the Prepare method of the CSQLite class.

Inherits from CSQLite.

```
CONSTRUCTOR CSQLiteStmt (BYVAL pStmt AS sqlite3_stmt PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pStmt* | The sqlite3_stmt pointer. |

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [BindBlob](#BindBlob) | Binds a blob with the statement. |
| [BindDouble](#BindDouble) | Binds a double value with the statement. |
| [BindLong](#BindLong) | Binds a long value with the statement. |
| [BindLongInt](#BindLongInt) | Binds a longint value with the statement. |
| [BindNull](#BindNull) | Binds a null value with the statement. |
| [BindParameterCount](#BindParameterCount) | This method can be used to find the number of SQL parameters in a prepared statement. |
| [BindParameterIndex](#BindParameterCount) | Returns the index of an SQL parameter given its name. |
| [BindParameterName](#BindParameterName) | Returns the name of the N-th SQL parameter in the prepared statement P. |
| [BindText](#BindText) | Binds a text value with the statement. |
| [BindZeroBlob](#BindZeroBlob) | Binds a BLOB that is filled with zeroes. |
| [Busy](#Busy) | Returns true if the prepared statement has been stepped at least once using **Step_** but has not run to completion and/or has not been reset using **Reset**. |
| [ClearBindings](#ClearBindings) | Sets all the parameters in the compiled SQL statement to NULL. |
| [ColumnBlob](#ColumnBlob) | Returns information about a single column of the current result row of a query. |
| [ColumnBytes](#ColumnBytes) | Returns the number of bytes of the column value. |
| [ColumnCount](#ColumnCount) | Returns the number of columns in the result set returned by the prepared statement. |
| [ColumnDatabaseName](#ColumnDatabaseName) | Returns the database name that is the origin of a particular result column in SELECT statement. |
| [ColumnDeclaredType](#ColumnDeclaredType) | Returns the declared data type of a query result. |
| [ColumnDouble](#ColumnDouble) | Returns the column value as a double. |
| [ColumnLong](#ColumnLong) | Returns the column value as a long. |
| [ColumnLongInt](#ColumnLongInt) | Returns the column value as a quad. |
| [ColumnName](#ColumnName) | Returns the name assigned to a particular column in the result set of a SELECT statement. |
| [ColumnOriginName](#ColumnOriginName) | Returns the column name that is the origin of a particular result column in SELECT statement. |
| [ColumnTableName](#ColumnTableName) | Returns the table name that is the origin of a particular result column in SELECT statement. |
| [ColumnText](#ColumnText) | Returns the column value as a UTF-16 string. |
| [ColumnType](#ColumnType) | Returns the column type. |
| [DataCount](#DataCount) | Returns the number of columns in the result set returned by the prepared statement. |
| [DbHandle](#DbHandle) | Returns the database connection handle to which a prepared statement belongs. |
| [Finalize](#Finalize) | Deletes a prepared statement. |
| [GetRow](#GetRow) | After a prepared statement has been prepared using either Prepare this method must be called one or more times to evaluate the statement. **GetRow** is an alias for **Step_**. |
| [hStmt](#hStmt) | Gets/sets the connection handle. |
| [IsColumnNull](#IsColumnNull) | Returns true is the column value is null or false otherwise. |
| [ReadOnly](#ReadOnly) | Rreturns true if and only if the prepared statement makes no direct changes to the content of the database file. |
| [Reset](#Reset) | Resets a prepared statement object back to its initial state, ready to be re-executed. |
| [Sql](#Sql) | Retrieve a saved copy of the original SQL text used to create a prepared statement if that statement was compiled using **Prepare**. |
| [Step_](#Step_) | After a prepared statement has been prepared using **Prepare** this method must be called one or more times to evaluate the statement. |

# <a name="CompileOptionUsed"></a>CompileOptionUsed

Returns FALSE or TRUE indicating whether the specified option was defined at compile time. The SQLITE_ prefix may be omitted from the option name passed to **CompileOptionUsed**.

```
FUNCTION CompileOptionUsed (BYREF szOptName AS ZSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *szOptName* | The option name. |

### Usage example

```
pSql.CompileOptionUsed("SQLITE_ENABLE_DBSTAT_VTAB")
```

# <a name="Complete"></a>Complete

Determines if the currently entered text seems to form a complete SQL statement or if additional input is needed before sending the text into SQLite for parsing. 

```
FUNCTION Complete (BYREF wszSql AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSql* | The SQL string to check. |

#### Return value

Returns 0 if the statement is incomplete or 1 if the input string appears to be a complete SQL statement. If a memory allocation fails, then SQLITE_NOMEM is returned.

#### Remarks

This method is useful during command-line input to determine if the currently entered text seems to form a complete SQL statement or if additional input is needed before sending the text into SQLite for parsing. This method returns 1 if the input string appears to be a complete SQL statement. A statement is judged to be complete if it ends with a semicolon token and is not a prefix of a well-formed CREATE TRIGGER statement. Semicolons that are embedded within string literals or quoted identifier names or comments are not independent tokens (they are part of the token in which they are embedded) and thus do not count as a statement terminator. Whitespace and comments that follow the final semicolon are ignored.

This method does not parse the SQL statements thus will not detect syntactically incorrect SQL.

# <a name="EnableSharedCache"></a>EnableSharedCache

Enables or disables the sharing of the database cache and schema data structures between connections to the same database. Sharing is enabled if the argument is true and disabled if the argument is false.. 

```
FUNCTION EnableSharedCache (BYVAL bSharing AS BOOLEAN) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *bSharing* | TRUE or FALSE. |

#### Return value

Returns SQLITE_OK if shared cache was enabled or disabled successfully. An error code is returned otherwise.

# <a name="ErrStr"></a>ErrStr

Returns English-language text that describes the result code.

```
FUNCTION ErrStr (BYVAL nErrorCode AS LONG) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *nErrorCode* | The result code. |

#### Return value

A string containing a description of the error.

# <a name="Free"></a>Free

Releases memory previously allocated by Malloc or Realloc.

```
SUB Free (BYVAL pMem AS ANY PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMem* | The pointer returned by Malloc or Realloc. |

# <a name="GetCompileOption"></a>GetCompileOption

Allows iterating over the list of options that were defined at compile time by returning the N-th compile time option string. If nOption is out of range, **GetCompileOption** returns a NULL pointer. The SQLITE_ prefix is omitted from any strings returned by **GetCompileOption**.

```
FUNCTION GetCompileOption (BYVAL nOption AS LONG) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *nOption* | The option number. |

#### Uage example

```
pSql.GetCompileOption(0)
```

# <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS LONG
```

#### Return value

The result code returned by the last executed method.

#### Remarks

The last result code is not global, but bound to each object, and its purpose is to check the failure or success of the last SQLite operation. To get descriptive information about SQLite errors call **ErrStr**.

# <a name="Malloc"></a>Malloc

Returns a pointer to a block of memory at least N bytes in length, where N is the parameter. If Malloc is unable to obtain sufficient free memory, it returns a NULL pointer. If the parameter nBytes to Malloc is zero or negative then returns a NULL pointer.

```
FUNCTION Malloc (BYVAL nBytes AS LONG) AS ANY PTR
FUNCTION Malloc64 (BYVAL nBytes AS sqlite3_uint64) AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nBytes* | The number of bytes to allocate. |

#### Return value

Pointer to the allocated memory.

#### Remarks

The memory returned by **Malloc** and **Realloc** is always aligned to at least an 8 byte boundary, or to a 4 byte boundary if the SQLITE_4_BYTE_ALIGNED_MALLOC compile-time option is used.

In SQLite version 3.5.0 and 3.5.1, it was possible to define the SQLITE_OMIT_MEMORY_ALLOCATION which would cause the built-in implementation of these functions to be omitted. That capability is no longer provided. Only built-in memory allocators can be used.

Prior to SQLite version 3.7.10, the Windows OS interface layer called the system **malloc** and **free** directly when converting filenames between the UTF-8 encoding used by SQLite and whatever filename encoding is used by the particular Windows installation. Memory allocation errors were detected, but they were reported back as SQLITE_CANTOPEN or SQLITE_IOERR rather than SQLITE_NOMEM.

The pointer arguments to **Free** and **Realloc** must be either NULL or else pointers obtained from a prior invocation of Malloc or **Realloc** that have not yet been released.

The application must not read or write any part of a block of memory after it has been released using **Free** or **Realloc**. 

# <a name="MemoryHighwater"></a>MemoryHighwater

Returns the maximum value of **MemoryUsed** since the high-water mark was last reset.

```
FUNCTION MemoryHighwater (BYVAL resetFlag AS BOOLEAN) AS sqlite3_uint64
```

| Parameter  | Description |
| ---------- | ----------- |
| *resetFlag* | TRUE or FALSE. Reset the high-water mark. |

#### Return value

The number of bytes of memory currently outstanding (malloced but not freed) since the high-water mark was last reset.

#### Remarks

The value returned by **MemoryHighwater** include any overhead added by SQLite in its implementation of **Malloc**, but not overhead added by the any underlying system library functions that **Malloc** may call.

# <a name="MemorySize"></a>MemorySize

Returns the size of that memory allocation in bytes.

```
FUNCTION MemorySize (BYVAL pMem AS ANY PTR) AS sqlite3_uint64
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMem* | Pointer to the memory allocation. |

#### Return value

If *pMem* is a memory allocation previously obtained from **Malloc**, **Malloc64**, **Realloc** or **Realloc64**, then **MemorySize** returns the size of that memory allocation in bytes. The value returned by **MemorySize** might be larger than the number of bytes requested when pMem was allocated. If *pMem* is a NULL pointer then **MemorySize** returns zero. If *pMem* points to something that is not the beginning of memory allocation, or if it points to a formerly valid memory allocation that has now been freed, then the behavior of **MemorySize** is undefined and possibly harmful.

# <a name="MemoryUsed"></a>MemoryUsed

Returns the number of bytes of memory currently outstanding (malloced but not freed).

```
FUNCTION MemoryUsed () AS sqlite3_uint64
```

#### Return value

The number of bytes of memory currently outstanding (malloced but not freed).

#### Remarks

The value returned by **MemoryUsed** include any overhead added by SQLite in its implementation of **Malloc**, but not overhead added by the any underlying system library functions that **Malloc** may call.

# <a name="Randomness"></a>Randomness

Pseudo-random number generator.

```
SUB Randomness (BYVAL nBytes AS LONG, BYVAL pbuffer AS ANY PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nBytes* | The number of bytes of randomness to return. |
| *pbuffer* | Pointer where to store the random bytes. It can be null. |

#### Remarks

SQLite contains a high-quality pseudo-random number generator (PRNG) used to select random ROWIDs when inserting new records into a table that already uses the largest possible ROWID. The PRNG is also used for the build-in random() and randomblob() SQL functions. This method allows applications to access the same PRNG for other purposes.

# <a name="Realloc"></a>Realloc

Attempts to resize a prior memory allocation to be at least *nBytes* bytes. *pMem* is the memory allocation to be resized.

```
FUNCTION Realloc (BYVAL pMem AS ANY PTR, BYVAL nBytes AS LONG) AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMem* | Pointer to memory allocated by **Malloc**. |
| *nBytes* | The number of bytes to reallocate. |

#### Return value

Pointer to the reallocated memory.

#### Remarks

The memory returned by **Malloc** and **Realloc** is always aligned to at least an 8 byte boundary, or to a 4 byte boundary if the SQLITE_4_BYTE_ALIGNED_MALLOC compile-time option is used.

In SQLite version 3.5.0 and 3.5.1, it was possible to define the SQLITE_OMIT_MEMORY_ALLOCATION which would cause the built-in implementation of these functions to be omitted. That capability is no longer provided. Only built-in memory allocators can be used.

Prior to SQLite version 3.7.10, the Windows OS interface layer called the system **malloc** and **free** directly when converting filenames between the UTF-8 encoding used by SQLite and whatever filename encoding is used by the particular Windows installation. Memory allocation errors were detected, but they were reported back as SQLITE_CANTOPEN or SQLITE_IOERR rather than SQLITE_NOMEM.

The pointer arguments to **Free** and **Realloc** must be either NULL or else pointers obtained from a prior invocation of **Malloc** or **Realloc** that have not yet been released.

The application must not read or write any part of a block of memory after it has been released using **Free** or **Realloc**. 

# <a name="ReleaseMemory"></a>ReleaseMemory

Attempts to free the specified number of bytes of heap memory by deallocating non-essential memory allocations held by the database library. Memory used to cache database pages to improve performance is an example of non-essential memory. **ReleaseMemory** returns the number of bytes actually freed, which might be more or less than the amount requested. The **ReleaseMemory** method is a no-op returning zero if SQLite is not compiled with SQLITE_ENABLE_MEMORY_MANAGEMENT.

```
FUNCTION ReleaseMemory (BYVAL nBytes AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nBytes* | The number of bytes to free. |

#### Return value

The number of bytes freed.

# <a name="Sleep"></a>Sleep

Causes the current thread to suspend execution for at least a number of milliseconds specified in its parameter.

```
FUNCTION Sleep (BYVAL ms AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *ms* | The number of milliseconds to sleep. |

#### Return value

The number of milliseconds of sleep actually requested from the operating system is returned.

#### Remarks

If the operating system does not support sleep requests with millisecond time resolution, then the time will be rounded up to the nearest second. The number of milliseconds of sleep actually requested from the operating system is returned.


# <a name="SoftHeapLimit64"></a>SoftHeapLimit64

Sets and/or queries the soft limit on the amount of heap memory that may be allocated by SQLite.

```
FUNCTION SoftHeapLimit64 (BYVAL nBytes AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nBytes* | The number of bytes. If *nBytes* is negative then no change is made to the soft heap limit. Hence, the current size of the soft heap limit can be determined by invoking **SoftHeapLimit64** with a negative argument. If *nBytes* is zero then the soft heap limit is disabled. |

#### Return value

The size of the soft heap limit prior to the call, or negative in the case of an error.

#### Remarks

SQLite strives to keep heap memory utilization below the soft heap limit by reducing the number of pages held in the page cache as heap memory usages approaches the limit. The soft heap limit is "soft" because even though SQLite strives to stay below the limit, it will exceed the limit rather than generate an SQLITE_NOMEM error. In other words, the soft heap limit is advisory only.

The circumstances under which SQLite will enforce the soft heap limit may change in future releases of SQLite.

# <a name="SourceID"></a>SourceID

Returns the SQLite3 source identifier.

```
FUNCTION SourceID () AS STRING
```

#### Return value

A string containing the SQLite3 identifier, e.g. "2012-06-11 02:05:22 f5b5a13f7394dc143aa136f1d4faba6839eaa6dc".

# <a name="Status"></a>Status

Retrieves runtime status information about the performance of SQLite, and optionally to reset various highwater marks.

```
FUNCTION Status (BYVAL op AS LONG, BYREF pCurrent AS LONG, BYREF pHighwater AS LONG, _
   BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
FUNCTION Status64 (BYVAL op AS LONG, BYREF pCurrent AS sqlite3_int64, BYREF pHighwater AS sqlite3_int64, _
   BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *op* | An integer code for a specific satus parameter. |
| *pCurrent* | Pointer to a variable that receives the current value of the parameter. |
| *pHighWater* | Pointer to a variable that receives the highest recorded value of the parameter. |
| *resetFlag* | TRUE or FALSE. If true, then the counter is reset to zero after this interface call returns. |

#### Return value

Returns SQLITE_OK (0) on success and a non-zero error code on failure.

#### Remarks

This method is used to retrieve runtime status information about the performance of SQLite, and optionally to reset various highwater marks. The first argument is an integer code for the specific parameter to measure. Recognized integer codes are of the form SQLITE_STATUS_.... The current value of the parameter is returned into *pCurrent*. The highest recorded value is returned in *pHighwater*. If the *resetFlag* is true, then the highest record value is reset after *pHighwater* is written. Some parameters do not record the highest value. For those parameters nothing is written into *pHighwater* and the *resetFlag* is ignored. Other parameters record only the highwater mark and not the current value. For these latter parameters nothing is written into *pCurrent*.

This function is threadsafe but is not atomic. This function can be called while other threads are running the same or different SQLite interfaces. However the values returned in pCurrent and pHighwater reflect the status of SQLite at different points in time and it is possible that another thread might change the parameter in between the times when *pCurrent* and *pHighwater* are written.

# <a name="StrGlob"></a>StrGlob

The *StrGlob* method returns zero if string *szStr* matches the glob pattern *szGlob*, and it returns non-zero if string *szStr* does not match the glob pattern *szGlob*. This function is case sensitive.

```
FUNCTION StrGlob (BYREF szGlob AS ZSTRING, BYREF szStr AS ZSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *szGlob* | The glob pattern to compare. |
| *szStr* | The string to compare. |

# <a name="ThreadSafe"></a>ThreadSafe

Returns zero if and only if SQLite was compiled with mutexing code omitted due to the SQLITE_THREADSAFE compile-time option being set to 0.

```
FUNCTION ThreadSafe () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

#### Remarks

SQLite can be compiled with or without mutexes. When the SQLITE_THREADSAFE C preprocessor macro is 1 or 2, mutexes are enabled and SQLite is threadsafe. When the SQLITE_THREADSAFE macro is 0, the mutexes are omitted. Without the mutexes, it is not safe to use SQLite concurrently from more than one thread.

Enabling mutexes incurs a measurable performance penalty. So if speed is of utmost importance, it makes sense to disable the mutexes. But for maximum safety, mutexes should be enabled. The default behavior is for mutexes to be enabled.

This interface can be used by an application to make sure that the version of SQLite that it is linking against was compiled with the desired setting of the SQLITE_THREADSAFE macro.

# <a name="Version"></a>Version

Returns the SQLite3 version.

```
FUNCTION Version () AS STRING
```

#### Return value

A string containing the SQLite3 version, e.g. "3.7.13".

# <a name="VersionNumber"></a>VersionNumber

Returns the SQLite3 version number.

```
FUNCTION VersionNumber () AS LONG
```

#### Return value

A long integer containing the SQLite3 version number, e.g. 3007013.

# <a name="Changes"></a>Changes

This method returns the number of database rows that were changed or inserted or deleted by the most recently completed SQL statement on the current database connection.

```
FUNCTION Changes () AS LONG
```

#### Return value

The number of changes.

#### ks

Only changes that are directly specified by the INSERT, UPDATE, or DELETE statement are counted. Auxiliary changes caused by triggers or foreign key actions are not counted. Use the **Changes** method to find the total number of changes including changes caused by triggers and foreign key actions.

Changes to a view that are simulated by an INSTEAD OF trigger are not counted. Only real table changes are counted.

A "row change" is a change to a single row of a single table caused by an INSERT, DELETE, or UPDATE statement. Rows that are changed as side effects of REPLACE constraint resolution, rollback, ABORT processing, DROP TABLE, or by any other mechanisms do not count as direct row changes.

A "trigger context" is a scope of execution that begins and ends with the script of a trigger. Most SQL statements are evaluated outside of any trigger. This is the "top level" trigger context. If a trigger fires from the top level, a new trigger context is entered for the duration of that one trigger. Subtriggers create subcontexts for their duration.

Calling **Exec** or **Step_** recursively does not create a new trigger context.

This method returns the number of direct row changes in the most recent INSERT, UPDATE, or DELETE statement within the same trigger context.

Thus, when called from the top level, this function returns the number of changes in the most recent INSERT, UPDATE, or DELETE that also occurred at the top level. Within the body of a trigger, the **Changes** method can be called to find the number of changes in the most recently completed INSERT, UPDATE, or DELETE statement within the body of the same trigger. However, the number returned does not include changes caused by subtriggers since those have their own context.

If a separate thread makes changes on the same database connection while **Changes** is running then the value returned is unpredictable and not meaningful. 

# <a name="CloseDb"></a>CloseDb

Closes the database.

```
FUNCTION CloseDb () AS LONG
```

#### Return value

SQLITE_OK if the sqlite3 object is successfully destroyed and all associated resources are deallocated.

#### Remarks

Applications must finalize all prepared statements and close all BLOB handles associated with the sqlite3 object prior to attempting to close the object. If **CloseDb** is called on a database connection that still has outstanding prepared statements or BLOB handles, then it returns SQLITE_BUSY.

If **CloseDb** is invoked while a transaction is open, the transaction is automatically rolled back.

# <a name="ErrCode"></a>ErrCode

Returns the numeric result code for the most recent failed sqlite3 call associated with a database connection. If a prior API call failed but the most recent API call succeeded, the return value from **ErrCode** is undefined.

```
FUNCTION ErrCode () AS LONG
```

#### Return value

The error code.

#### Remarks

When the serialized threading mode is in use, it might be the case that a second error occurs on a separate thread in between the time of the first error and the call to these functions. When that happens, the second error will be reported since these functions always report the most recent result. To avoid this, each thread can obtain exclusive use of the database connection by invoking **sqlite3_mutex_enter**(*sqlite3_db_mutex(D)*) before beginning to use D and invoking **sqlite3_mutex_leave**(*sqlite3_db_mutex(D)*) after all calls to the functions listed here are completed.

If a function fails with SQLITE_MISUSE, that means the function was invoked incorrectly by the application. In that case, the error code and message may or may not be set. 

# <a name="ErrMsg"></a>ErrMsg

Returns English-language text that describes the error.

```
FUNCTION ErrMsg () AS STRING
```

#### Return value

A string containing a description of the error.

#### Remarks

When the serialized threading mode is in use, it might be the case that a second error occurs on a separate thread in between the time of the first error and the call to these functions. When that happens, the second error will be reported since these functions always report the most recent result. To avoid this, each thread can obtain exclusive use of the database connection by invoking **sqlite3_mutex_enter**(*sqlite3_db_mutex(D)*) before beginning to use D and invoking **sqlite3_mutex_leave**(*sqlite3_db_mutex(D)*) after all calls to the functions listed here are completed.

If a function fails with SQLITE_MISUSE, that means the function was invoked incorrectly by the application. In that case, the error code and message may or may not be set. 

# <a name="Exec"></a>Exec

Convenience wrapper for **Prepare** and **Step_**. To be used with queries that don't return result sets, such CREATE, UPDATE and INSERT.

```
FUNCTION Exec (BYREF wszSql AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSql* | The SQL statement. |

#### Return value

SQLITE_BUSY means that the database engine was unable to acquire the database locks it needs to do its job. If the statement is a COMMIT or occurs outside of an explicit transaction, then you can retry the statement. If the statement is not a COMMIT and occurs within an explicit transaction then you should rollback the transaction before continuing.

SQLITE_DONE means that the statement has finished executing successfully. **Step_** should not be called again on this virtual machine without first calling Reset to reset the virtual machine back to its initial state.

SQLITE_ERROR means that a run-time error (such as a constraint violation) has occurred. Step should not be called again on the cirtual machine. More information may be found by calling **ErrMsg**.

# <a name="ExtendedErrCode"></a>ExtendedErrCode

Gets the extended error code associated with this dabatase connection.

```
FUNCTION ExtendedErrCode () AS LONG
```

# <a name="ExtendedResultCodes"></a>ExtendedResultCodes

Enables or disables the extended result codes feature of SQLite. The extended result codes are disabled by default for historical compatibility.

```
FUNCTION ExtendedResultCodes (BYVAL onoff AS BOOLEAN) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *onoff* | TRUE to enable; FALSE to disable. |

#### Return value

SQLITE_OK (0) or an error code.

# <a name="hDbc"></a>hDbc

Gets/sets the database handle.

```
PROPERTY hDbc () AS sqlite3 PTR
PROPERTY hDbc (pDbc AS sqlite3 PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hDbc* | The database handle. |

#### Return value

The database handle.

# <a name="Interrupt"></a>Interrupt

This function causes any pending database operation to abort and return at its earliest opportunity. This function is typically called in response to a user action such as pressing "Cancel" or Ctrl-C where the user wants a long query operation to halt immediately.

```
SUB Interrupt
```

#### Remarks

It is safe to call this function from a thread different from the thread that is currently running the database operation. But it is not safe to call this function with a database connection that is closed or might close before **Interrupt** returns.

If an SQL operation is very nearly finished at the time when **Interrupt** is called, then it might not have an opportunity to be interrupted and might continue to completion.

An SQL operation that is interrupted will return SQLITE_INTERRUPT. If the interrupted SQL operation is an INSERT, UPDATE, or DELETE that is inside an explicit transaction, then the entire transaction will be rolled back automatically.

The **Interrupt** call is in effect until all currently running SQL statements on database connection complete. Any new SQL statements that are started after the **Interrupt** call and before the running statements reaches zero are interrupted as if they had been running prior to the **Interrupt** call. New SQL statements that are started after the running statement count reaches zero are not effected by the **Interrupt**. A call to **Interrupt** that occurs when there are no running SQL statements is a no-op and has no effect on SQL statements that are started after the **Interrupt** call returns.

If the database connection closes while **Interrupt** is running then bad things will likely happen.

# <a name="LastInsertRowId"></a>LastInsertRowId

Returns the rowid of the most recent successful INSERT into the database from the database connection in the first argument.

```
FUNCTION LastInsertRowId () AS sqlite3_int64
```

#### Return value

The rowid of the most recent successful INSERT into the database from the database connection in the first argument.

#### Remarks

Each entry in an SQLite table has a unique 64-bit signed integer key called the "rowid". The rowid is always available as an undeclared column named ROWID, OID, or _ROWID_ as long as those names are not also used by explicitly declared columns. If the table has a column of type INTEGER PRIMARY KEY then that column is another alias for the rowid.

This function returns the rowid of the most recent successful INSERT into the database from the database connection in the first argument. As of SQLite version 3.7.7, this functions records the last insert rowid of both ordinary tables and virtual tables. If no successful INSERTs have ever occurred on that database connection, zero is returned.

If an INSERT occurs within a trigger or within a virtual table method, then this function will return the rowid of the inserted row as long as the trigger or virtual table method is running. But once the trigger or virtual table method ends, the value returned by this function reverts to what it was before the trigger or virtual table method began.

An INSERT that fails due to a constraint violation is not a successful INSERT and does not change the value returned by this function. Thus INSERT OR FAIL, INSERT OR IGNORE, INSERT OR ROLLBACK, and INSERT OR ABORT make no changes to the return value of this function when their insertion fails. When INSERT OR REPLACE encounters a constraint violation, it does not fail. The INSERT continues to completion after deleting rows that caused the constraint problem so INSERT OR REPLACE will always change the return value of this interface.

For the purposes of this function, an INSERT is considered to be successful even if it is subsequently rolled back.

This function is accessible to SQL statements via the **last_insert_rowid** SQL function.

If a separate thread performs a new INSERT on the same database connection while the **LastInsertRowid** function is running and thus changes the last insert rowid, then the value returned by **LastInsertRowid** is unpredictable and might not equal either the old or the new last insert rowid. 

# <a name="Limit"></a>Limit

This function allows the size of various constructs to be limited on a connection by connection basis.

```
FUNCTION Limit (BYVAL id AS LONG, BYVAL newVal AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *id* | One of the limit categories that define a class of constructs to be size limited. |
| *newVal* | The the new limit for that construct. |

#### Return value

The prior value of the limit.

#### Remarks

If the new limit is a negative number, the limit is unchanged. For each limit category SQLITE_LIMIT_NAME there is a hard upper bound set at compile-time by a C preprocessor macro called SQLITE_MAX_NAME. (The "_LIMIT_" in the name is changed to "_MAX_".) Attempts to increase a limit above its hard upper bound are silently truncated to the hard upper bound.

Regardless of whether or not the limit was changed, the **Limit** interface returns the prior value of the limit. Hence, to find the current value of a limit without changing it, simply invoke this interface with the third parameter set to -1.

Run-time limits are intended for use in applications that manage both their own internal database and also databases that are controlled by untrusted external sources. An example application might be a web browser that has its own databases for storing history and separate databases controlled by JavaScript applications downloaded off the Internet. The internal databases can be given the large, default limits. Databases managed by external sources can be given much smaller limits designed to prevent a denial of service attack.

New run-time limit categories may be added in future releases.
