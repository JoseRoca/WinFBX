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

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Read](#Read1) | Reads a specified number of bytes from the stream into memory, starting at the current seek pointer. |
| [Write](#Write1) | Writes a specified number of bytes into the stream starting at the current seek pointer. |
| [Seek](#Seek1) | Changes the seek pointer to a new location. The new location is relative to either the beginning of the stream, the end of the stream, or the current seek pointer. |
| [GetSeekPosition](#GetSeekPosition1) | Returns the seek position. |
| [ResetSeekPosition](#ResetSeekPosition1) | Sets the seek position at the beginning of the stream. |
| [SeekAtEndOfFile](#SeekAtEndOfFile1) | Sets the seek position at the end of the stream. |
| [SeekAtEndOfStream](#SeekAtEndOfStream1) | Sets the seek position at the end of the stream. |
| [GetSize](#GetSize1) | Returns the size of the stream. |
| [SetSize](#SetSize1) | Changes the size of the stream. |
| [CopyTo](#CopyTo1) | Copies a specified number of bytes from the current seek pointer in the stream to the current seek pointer in another stream. |
| [Clone](#Clone1) | Creates a new stream with its own seek pointer that references the same bytes as the original stream. |
| [GetLastResult](#GetLastResult1) | Returns the last result code. |
| [GetErrorInfo](#GetErrorInfo1) | Returns a description of the last result code. |

# <a name="Read1"></a>Read

Reads a specified number of bytes from the stream into memory, starting at the current seek pointer.

```
FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbRead AS ULONG PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer which the stream data is read into. |
| *cb* | The number of bytes of data to read from the stream. |
| *pcbRead* | A pointer to a ULONG variable that receives the actual number of bytes read from the stream. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

```
FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer which the stream data is read into. |
| *cb* | The number of bytes of data to read from the stream. |

#### Return value

ULONG. The actual number of bytes read from the stream. Note: The number of bytes read may be zero.

# <a name="Write1"></a>Write

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

HRESULT. S_OK (0) on success, or an error code on failure.

```
FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer that contains the data that is to be written to the stream. A valid pointer must be provided for this parameter even when cb is zero. |
| *cb* | The number of bytes of data to attempt to write into the stream. This value can be zero. |

#### Return value

ULONG. The actual number of bytes written to the stream. Note: The number of bytes read may be zero.

# <a name="Seek1"></a>Seek

Changes the seek pointer to a new location. The new location is relative to either the beginning of the stream, the end of the stream, or the current seek pointer.

```
FUNCTION Seek (BYVAL dlibMove AS ULONGINT, _
   BYVAL dwOrigin AS DWORD, _
   BYVAL plibNewPosition AS ULONGINT PTR = NULL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dlibMove* | ULONGINT. The displacement to be added to the location indicated by the dwOrigin parameter. If *dwOrigin* is **STREAM_SEEK_SET**, this is interpreted as an unsigned value rather than a signed value. |
| *dwOrigin* | DWORD. The origin for the displacement specified in *dlibMove*. The origin can be the beginning of the file (**STREAM_SEEK_SET**), the current seek pointer (**STREAM_SEEK_CUR**), or the end of the file (**STREAM_SEEK_END**).<br>**STREAM_SEEK_SET** : The new seek pointer is an offset relative to the beginning of the stream. In this case, the *dlibMove* parameter is the new seek position relative to the beginning of the stream.<br>**STREAM_SEEK_CUR** : The new seek pointer is an offset relative to the current seek pointer location. In this case, the *dlibMove* parameter is the signed displacement from the current seek position.<br>**STREAM_SEEK_END** : The new seek pointer is an offset relative to the end of the stream. In this case, the *dlibMove* parameter is the new seek position relative to the end of the stream. |
| *plibNewPosition* | ULONGINT PTR. A pointer to the location where this method writes the value of the new seek pointer from the beginning of the stream. You can set this pointer to NULL. In this case, this method does not provide the new seek pointer. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetSeekPosition1"></a>GetSeekPosition

Returns the current seek position.

```
FUNCTION GetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The current seek position.

# <a name="ResetSeekPosition1"></a>ResetSeekPosition

Sets the seek position at the beginning of the stream.

```
FUNCTION ResetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

# <a name="SeekAtEndOfFile1"></a>SeekAtEndOfFile

Sets the seek position at the end of the stream.

```
FUNCTION SeekAtEndOfFile () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

# <a name="SeekAtEndOfStream1"></a>SeekAtEndOfFile

Sets the seek position at the end of the stream.

```
FUNCTION SeekAtEndOfStream () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

# <a name="GetSize1"></a>GetSize

Returns the size of the stream in bytes.

```
FUNCTION GetSize () AS ULONGINT
```

#### Return value

ULONGINT. The size of the stream in bytes.

# <a name="SetSize1"></a>SetSize

Changes the size of the stream object.

```
FUNCTION SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *libNewSize* | ULONGINT. Specifies the new size, in bytes, of the stream. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="CopyTo1"></a>CopyTo

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

# <a name="Clone1"></a>Clone

Creates a new stream with its own seek pointer that references the same bytes as the original stream. The **Clone** method creates a new stream for accessing the same bytes but using a separate seek pointer. The new stream sees the same data as the source-stream. Changes written to one stream are immediately visible in the other. Range locking is shared between the streams. The initial setting of the seek pointer in the cloned stream instance is the same as the current setting of the seek pointer in the original stream at the time of the clone operation.

```
FUNCTION Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *ppstm* | When successful, pointer to the location of an **IStream** pointer to the new stream. If an error occurs, this parameter is NULL. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetLastResult1"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetErrorInfo1"></a>GetErrorInfo

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS CWSTR
```

#### Return value

CWSTR. A description of the last result code. If the result code is S_OK (0), it returns "Success"; otherwise, it returns the hexadecimal value of the error code and a description such "File not found", "Seek error", "Write fault", "Read fault" or "Share violation".
