# CMemStream Class

Creates a memory stream, allowing read, write and seek operations. The stream is thread-safe as of Windows 8. On earlier systems, the stream is not thread-safe. Cloning is supported as of Windows 8.

**Include file**: CStream.inc

### Constructor

```
CONSTRUCTOR CMemStream (BYVAL pInit AS CONST BYTE PTR, BYVAL cbInit AS UINT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pInit* | A pointer to a buffer of size *cbInit*. The contents of this buffer are used to set the initial contents of the memory stream. If this parameter is NULL, the returned memory stream does not have any initial content. |
| *cbInit* | The number of bytes in the buffer pointed to by *pInit*. If *pInit* is set to NULL, cbInit must be zero. |

### Operator CAST

Returns a pointer to the underlying **IStream** interface of the stream object.

```
OPERATOR CAST () AS IStream PTR
```

# CMemRexrStream Class

Creates a memory text stream, allowing read, write and seek operations. The stream is thread-safe as of Windows 8. On earlier systems, the stream is not thread-safe. Cloning is supported as of Windows 8.

**Include file**: CStream.inc

### Constructor

```
CONSTRUCTOR CMemStream
CONSTRUCTOR CMemStream (BYVAL pwszText AS CONST WSTRING PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszText* | Creates a memory text stream and initializes it with the content of a string. If **pwszText** is null, it creates a Memory text stream without content. |

### Operator CAST

Returns a pointer to the underlying **IStream** interface of the stream object.

```
OPERATOR CAST () AS IStream PTR
```

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Read](#Read2) | Reads a specified number of characters from the stream into memory, starting at the current seek pointer. |
| [Write](#Write2) | Writes a specified number of bytes into the stream starting at the current seek pointer. |
| [Append](#Append2) | Appends a string at the end of the stream. |
| [Seek](#Seek2) | Changes the seek pointer to a new location. The new location is relative to either the beginning of the stream, the end of the stream, or the current seek pointer. |
| [GetSeekPosition](#GetSeekPosition2) | Returns the seek position. |
| [ResetSeekPosition](#ResetSeekPosition2) | Sets the seek position at the beginning of the stream. |
| [SeekAtEndOfFile](#SeekAtEndOfFile2) | Sets the seek position at the end of the stream. |
| [SeekAtEndOfStream](#SeekAtEndOfStream2) | Sets the seek position at the end of the stream. |
| [GetSize](#GetSize2) | Returns the size of the stream. |
| [SetSize](#SetSize2) | Changes the size of the stream. |
| [Clone](#Clone2) | Creates a new stream with its own seek pointer that references the same bytes as the original stream. |
| [GetLastResult](#GetLastResult2) | Returns the last result code. |
| [GetErrorInfo](#GetErrorInfo2) | Returns a description of the last result code. |

# <a name="Read1"></a>Read

Reads a specified number of characters from the stream into memory, starting at the current seek pointer.

```
FUNCTION Read (BYVAL numChars AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *numChars* | The number of characters to read from the stream. Pass -1 to read all the characters from the current seek position. |

#### Return value

CWSTR. The characters read.

# <a name="Write2"></a>Write

Writes a string at the current seek position.

```
FUNCTION Write (BYREF wszText AS CONST WSTRING) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | The string to write. |

#### Return value

The number of characters actually written.

# <a name="Append2"></a>Append

Appends a string at the end of the stream.

```
FUNCTION Append (BYREF wszText AS CONST WSTRING) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | The string to append. |

#### Return value

The number of characters actually written.

# <a name="Seek2"></a>Seek

Sets the seek position as an absolute position from the start of the stream.

```
FUNCTION Seek (Seek (BYVAL nPos AS ULONGINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPos* | The new seek position (from 1 to the end of the stream). |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetSeekPosition2"></a>GetSeekPosition

Returns the current seek position.

```
FUNCTION GetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The current seek position.

# <a name="ResetSeekPosition2"></a>ResetSeekPosition

Sets the seek position at the beginning of the stream.

```
FUNCTION ResetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

# <a name="SeekAtEndOfFile2"></a>SeekAtEndOfFile

Sets the seek position at the end of the stream.

```
FUNCTION SeekAtEndOfFile () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

# <a name="SeekAtEndOfStream2"></a>SeekAtEndOfFile

Sets the seek position at the end of the stream.

```
FUNCTION SeekAtEndOfStream () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

# <a name="GetSize2"></a>GetSize

Returns the size of the stream in characters.

```
FUNCTION GetSize () AS ULONGINT
```

#### Return value

ULONGINT. The size of the stream in bytes.

# <a name="SetSize2"></a>SetSize

Changes the size of the stream object.

```
FUNCTION SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *libNewSize* | ULONGINT. Specifies the new size, in characters, of the stream. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="Clone2"></a>Clone

Creates a new stream with its own seek pointer that references the same bytes as the original stream. The **Clone** method creates a new stream for accessing the same bytes but using a separate seek pointer. The new stream sees the same data as the source-stream. Changes written to one stream are immediately visible in the other. Range locking is shared between the streams. The initial setting of the seek pointer in the cloned stream instance is the same as the current setting of the seek pointer in the original stream at the time of the clone operation.

```
FUNCTION Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *ppstm* | When successful, pointer to the location of an **IStream** pointer to the new stream. If an error occurs, this parameter is NULL. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetLastResult2"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetErrorInfo2"></a>GetErrorInfo

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS CWSTR
```

#### Return value

CWSTR. A description of the last result code. If the result code is S_OK (0), it returns "Success"; otherwise, it returns the hexadecimal value of the error code and a description such "Seek error", "Write fault", "Read fault" or "Invalid argument".
