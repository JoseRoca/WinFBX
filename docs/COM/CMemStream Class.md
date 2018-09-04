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
| *cbInit* | The number of bytes in the buffer pointed to by *pInit*. If *pInit* is set to NULL, *cbInit* must be zero. |

### Operator CAST

Returns a pointer to the underlying **IStream** interface of the stream object.

```
OPERATOR CAST () AS IStream PTR
```

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Read](#Read1) | Reads a specified number of characters from the stream into memory, starting at the current seek pointer. |
| [Write](#Write1) | Writes a specified number of bytes into the stream starting at the current seek pointer. |
| [Seek](#Seek1) | Changes the seek pointer to a new location. The new location is relative to either the beginning of the stream, the end of the stream, or the current seek pointer. |
| [GetSeekPosition](#GetSeekPosition1) | Returns the seek position. |
| [ResetSeekPosition](#ResetSeekPosition1) | Sets the seek position at the beginning of the stream. |
| [SeekAtEndOfStream](#SeekAtEndOfStream1) | Sets the seek position at the end of the stream. |
| [CopyTo](#CopyTo) | Copies a specified number of bytes from the current seek pointer in the stream to the current seek pointer in another stream. |
| [GetSize](#GetSize1) | Returns the size of the stream. |
| [SetSize](#SetSize1) | Changes the size of the stream. |
| [Clone](#Clone1) | Creates a new stream with its own seek pointer that references the same bytes as the original stream. |
| [GetLastResult](#GetLastResult1) | Returns the last result code. |
| [GetErrorInfo](#GetErrorInfo1) | Returns a description of the last result code. |

# CMemTextStream Class

Creates a memory text stream, allowing read, write and seek operations. The stream is thread-safe as of Windows 8. On earlier systems, the stream is not thread-safe. Cloning is supported as of Windows 8.

**Include file**: CStream.inc

### Constructor (CMemStream)

```
CONSTRUCTOR CMemTextStream
CONSTRUCTOR CMemTextStream (BYVAL pwszText AS CONST WSTRING PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszText* | Creates a memory text stream and initializes it with the content of a string. If this parameter is NULL, the returned memory stream does not have any initial content. |

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

# <a name="Read1"></a>Read (CMemStream)

Reads a specified number of bytes from the stream into memory, starting at the current seek pointer.

```
FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbRead AS ULONG PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer which the stream data is read into. |
| *cb* | The number of bytes of data to read from the stream. |
| *pcbRead* | A pointer to a ULONG variable that receives the actual number of bytes read from the stream. The number of bytes read may be zero. |

#### Return value

S_OK (0) or an HRESULT code.

```
FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer which the stream data is read into. |
| *cb* | The number of bytes of data to read from the stream. |

ULONG. The number characters read.

# <a name="Write1"></a>Write (CMemStream)

Writes a specified number of bytes into the stream starting at the current seek pointer.

```
FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbWritten AS ULONG PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer that contains the data that is to be written to the stream. A valid pointer must be provided for this parameter even when cb is zero. |
| *cb* | The number of bytes of data to attempt to write into the stream. This value can be zero. |
| *pcbWritten* | A pointer to a ULONG variable where this method writes the actual number of bytes written to the stream. The caller can set this pointer to NULL, in which case this method does not provide the actual number of bytes written. |

#### Return value

S_OK (0) or an HRESULT code.

```
FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer that contains the data that is to be written to the stream. A valid pointer must be provided for this parameter even when cb is zero. |
| *cb* | The number of bytes of data to attempt to write into the stream. This value can be zero. |

#### Return value

The number of characters actually written.

# <a name="Seek1"></a>Seek (CMemStream)

Changes the seek pointer to a new location. The new location is relative to either the beginning of the stream, the end of the stream, or the current seek pointer.

```
FUNCTION Seek (BYVAL dlibMove AS ULONGINT, BYVAL dwOrigin AS DWORD, BYVAL plibNewPosition AS ULONGINT PTR = NULL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dlibMove* | The displacement to be added to the location indicated by the *dwOrigin* parameter. If *dwOrigin* is STREAM_SEEK_SET, this is interpreted as an unsigned value rather than a signed value. |
| *dwOrigin* | The origin for the displacement specified in dlibMove. The origin can be the beginning of the file (STREAM_SEEK_SET), the current seek pointer (STREAM_SEEK_CUR), or the end of the file (STREAM_SEEK_END). For more information about values, see the STREAM_SEEK enumeration. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetSeekPosition1"></a>GetSeekPosition (CMemStream)

Returns the current seek position.

```
FUNCTION GetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The current seek position.

# <a name="ResetSeekPosition1"></a>ResetSeekPosition (CMemStream)

Sets the seek position at the beginning of the stream.

```
FUNCTION ResetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

# <a name="SeekAtEndOfStream1"></a>SeekAtEndOfStream (CMemStream)
Sets the seek position at the end of the stream.

```
FUNCTION SeekAtEndOfStream () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

# <a name="GetSize1"></a>GetSize (CMemStream)

Returns the size of the stream in bytes.

```
FUNCTION GetSize () AS ULONGINT
```

#### Return value

ULONGINT. The size of the stream in bytes.

# <a name="SetSize1"></a>SetSize (CMemStream)

Changes the size of the stream object.

```
FUNCTION SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *libNewSize* | ULONGINT. Specifies the new size, in bytes, of the stream. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="CopyTo"></a>CopyTo (CMemStream)

Copies a specified number of bytes from the current seek pointer in the stream to the current seek pointer in another stream.

```
FUNCTION CopyTo (BYVAL pstm AS IStream PTR, _
   BYVAL cb AS ULONGINT, _
   BYVAL pcbRead AS ULONGINT PTR = NULL, _
   BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstm* | A pointer to the destination stream. The stream pointed to by *pstm* can be a new stream or a clone of the source stream. |
| *cb* | The number of bytes of data to attempt to copy into the stream. |
| *pcbRead* | A pointer to the location where this method writes the actual number of bytes read from the source. You can set this pointer to NULL. In this case, this method does not provide the actual number of bytes read. |
| *pcbWritten* | A pointer to the location where this method writes the actual number of bytes written to the destination. You can set this pointer to NULL. In this case, this method does not provide the actual number of bytes written. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="Clone1"></a>Clone (CMemStream)

Creates a new stream with its own seek pointer that references the same bytes as the original stream. The **Clone** method creates a new stream for accessing the same bytes but using a separate seek pointer. The new stream sees the same data as the source-stream. Changes written to one stream are immediately visible in the other. Range locking is shared between the streams. The initial setting of the seek pointer in the cloned stream instance is the same as the current setting of the seek pointer in the original stream at the time of the clone operation.

```
FUNCTION Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *ppstm* | When successful, pointer to the location of an **IStream** pointer to the new stream. If an error occurs, this parameter is NULL. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetLastResult1"></a>GetLastResult (CMemStream)

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetErrorInfo1"></a>GetErrorInfo (CMemStream)

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS CWSTR
```

#### Return value

CWSTR. A description of the last result code. If the result code is S_OK (0), it returns "Success"; otherwise, it returns the hexadecimal value of the error code and a description such "Seek error", "Write fault", "Read fault" or "Invalid argument".

# <a name="Read2"></a>Read (CMemTextStream)

Reads a specified number of characters from the stream into memory, starting at the current seek pointer.

```
FUNCTION Read (BYVAL numChars AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *numChars* | The number of characters to read from the stream. Pass -1 to read all the characters from the current seek position. |

#### Return value

CWSTR. The characters read.

# <a name="Write2"></a>Write (CMemTextStream)

Writes a string at the current seek position.

```
FUNCTION Write (BYREF wszText AS CONST WSTRING) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | The string to write. |

#### Return value

The number of characters actually written.

# <a name="Append2"></a>Append (CMemTextStream)

Appends a string at the end of the stream.

```
FUNCTION Append (BYREF wszText AS CONST WSTRING) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | The string to append. |

#### Return value

The number of characters actually written.

# <a name="Seek2"></a>Seek (CMemTextStream)

Sets the seek position as an absolute position from the start of the stream.

```
FUNCTION Seek (Seek (BYVAL nPos AS ULONGINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPos* | The new seek position (from 1 to the end of the stream). |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetSeekPosition2"></a>GetSeekPosition (CMemTextStream)

Returns the current seek position.

```
FUNCTION GetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The current seek position.

# <a name="ResetSeekPosition2"></a>ResetSeekPosition (CMemTextStream)

Sets the seek position at the beginning of the stream.

```
FUNCTION ResetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

# <a name="SeekAtEndOfStream2"></a>SeekAtEndOfStream (CMemTextStream)

Sets the seek position at the end of the stream.

```
FUNCTION SeekAtEndOfStream () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

# <a name="GetSize2"></a>GetSize (CMemTextStream)

Returns the size of the stream in characters.

```
FUNCTION GetSize () AS ULONGINT
```

#### Return value

ULONGINT. The size of the stream in bytes.

# <a name="SetSize2"></a>SetSize (CMemTextStream)

Changes the size of the stream object.

```
FUNCTION SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *libNewSize* | ULONGINT. Specifies the new size, in characters, of the stream. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="Clone2"></a>Clone (CMemTextStream)

Creates a new stream with its own seek pointer that references the same bytes as the original stream. The **Clone** method creates a new stream for accessing the same bytes but using a separate seek pointer. The new stream sees the same data as the source-stream. Changes written to one stream are immediately visible in the other. Range locking is shared between the streams. The initial setting of the seek pointer in the cloned stream instance is the same as the current setting of the seek pointer in the original stream at the time of the clone operation.

```
FUNCTION Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *ppstm* | When successful, pointer to the location of an **IStream** pointer to the new stream. If an error occurs, this parameter is NULL. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetLastResult2"></a>GetLastResult (CMemTextStream)

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetErrorInfo2"></a>GetErrorInfo (CMemTextStream)

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS CWSTR
```

#### Return value

CWSTR. A description of the last result code. If the result code is S_OK (0), it returns "Success"; otherwise, it returns the hexadecimal value of the error code and a description such "Seek error", "Write fault", "Read fault" or "Invalid argument".
