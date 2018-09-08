# CSQLite Class

SQLite is an embedded SQL database engine. Unlike most other SQL databases, SQLite does not have a separate server process. SQLite reads and writes directly to ordinary disk files. A complete SQL database with multiple tables, indices, triggers, and views, is contained in a single disk file. The database file format is cross-platform - you can freely copy a database between 32-bit and 64-bit systems or between big-endian and little-endian architectures. These features make SQLite a popular choice as an Application File Format.

**Include file**: CSQLite.inc

# Constructor

Creates a new instance of the **CSQlite** base class and loads the SQLite3 DLL.

```
CONSTRUCTOR CSQLite (BYREF wszDllPath AS WSTRING = "sqlite3.dll")
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszDllPath* | The path of the SQLite3 DLL. |

### Methods and Properties

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
| [VersionNumber](#Version) | Returns the SQLite3 version number. |

# CSQLiteDb Class

Inherits from CSQLite.

# Constructor

Opens an SQLite database file as specified by the filename argument.

```
CONSTRUCTOR CSQLiteDb (BYREF wszPath AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The name of the database. |

#### Remarks

The database is opened for reading and writing, and is created if it does not already exist. If the filename is ":memory:", then a private, temporary in-memory database is created for the connection. This in-memory database will vanish when the database connection is closed. If the filename is an empty string, then a private, temporary on-disk database will be created. This private database will be automatically deleted as soon as the database connection is closed. If URI filename interpretation is enabled, and the filename argument begins with "file:", then the filename is interpreted as a URI.

### Methods and Properties

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
