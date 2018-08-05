' ########################################################################################
' Microsoft Windows
' File: CStream.inc
' Contents: Classes to work with streams.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2018 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "Afx/CWSTR.inc"

NAMESPACE Afx

' ========================================================================================
' Opens or creates a file and retrieves a stream to read or write to that file.
' ========================================================================================
PRIVATE FUNCTION AfxSHCreateStreamOnFileEx (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD, _
BYVAL dwAttributes AS DWORD, BYVAL fCreate AS WINBOOL, BYVAL ppStream AS IStream PTR PTR) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("shlwapi.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pShCreateStreamOnFileEx AS FUNCTION (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD, BYVAL dwAttributes AS DWORD, _
       BYVAL fCreate AS WINBOOL, BYVAL pstmTemplate AS IStream PTR, BYVAL ppStream AS IStream PTR PTR) AS HRESULT
   pShCreateStreamOnFileEx = DyLibSymbol(pLib, "SHCreateStreamOnFileEx")
   IF pShCreateStreamOnFileEx = NULL THEN EXIT FUNCTION
   DIM hr AS HRESULT = pShCreateStreamOnFileEx(pwszFile, grfMode, dwAttributes, fCreate, NULL, ppStream)
   DyLibFree(pLib)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' CFileStream class
' ========================================================================================
TYPE CFileStream

Public:
   m_Result AS HRESULT
   m_pStream AS IStream PTR

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD = STGM_READ, _
           BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS WINBOOL = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pstm AS IStream PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE OPERATOR LET (BYVAL pstm AS IStream PTR)
   DECLARE OPERATOR CAST () AS IStream PTR
   DECLARE FUNCTION Attach (BYVAL pstm AS IStream PTR) AS HRESULT
   DECLARE FUNCTION Detach () AS IStream PTR
   DECLARE FUNCTION Open (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD = STGM_READ, _
           BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS WINBOOL = FALSE) AS HRESULT
   DECLARE SUB Close
   DECLARE FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbRead AS ULONG PTR) AS HRESULT
   DECLARE FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   DECLARE FUNCTION ReadTextA (BYVAL numChars AS LONG) AS STRING
   DECLARE FUNCTION ReadTextW (BYVAL numChars AS LONG) AS CWSTR
   DECLARE FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbWritten AS ULONG PTR) AS HRESULT
   DECLARE FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   DECLARE FUNCTION WriteTextA (BYREF strText AS STRING) AS ULONG
   DECLARE FUNCTION WriteTextW (BYREF wszText AS WSTRING) AS ULONG
   DECLARE FUNCTION Seek (BYVAL dlibMove AS ULONGINT, BYVAL dwOrigin AS DWORD, BYVAL plibNewPosition AS ULONGINT PTR = NULL) AS HRESULT
   DECLARE FUNCTION GetSeekPosition () AS ULONGINT
   DECLARE FUNCTION ResetSeekPosition () AS ULONGINT
   DECLARE FUNCTION SeekAtEndOfFile () AS ULONGINT
   DECLARE FUNCTION SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
   DECLARE FUNCTION GetSize () AS ULONGINT
   DECLARE FUNCTION CopyTo (BYVAL pstm AS IStream PTR, BYVAL cb AS ULONGINT, _
           BYVAL pcbRead AS ULONGINT PTR = NULL, BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
   DECLARE FUNCTION Commit (BYVAL grfCommitFlags AS DWORD) AS HRESULT
   DECLARE FUNCTION Revert () AS HRESULT
   DECLARE FUNCTION LockRegion (BYVAL libOffset AS ULONGINT, BYVAL cb AS ULONGINT, BYVAL dwLockType AS DWORD) AS HRESULT
   DECLARE FUNCTION UnlockRegion (BYVAL libOffset AS ULONGINT, BYVAL cb AS ULONGINT, BYVAL dwLockType AS DWORD) AS HRESULT
   DECLARE FUNCTION Stat (BYVAL pstatstg AS STATSTG PTR, BYVAL grfStatFlag AS DWORD) AS HRESULT
   DECLARE FUNCTION Stat (BYVAL grfStatFlag AS DWORD) AS STATSTG
   DECLARE FUNCTION Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
   DECLARE FUNCTION Clone () AS IStream PTR
   DECLARE FUNCTION GetErrorInfo () AS STRING

END TYPE
' ========================================================================================

' ========================================================================================
' CFileStream constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CFileStream
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Opens or creates a file and retrieves a stream to read or write to that file.
' Parameters:
' - pwszFile : A pointer to a unicode null-terminated string that specifies the file name.
' - grfMode : One or more STGM values that are used to specify the file access mode and how
'   the object that exposes the stream is created and deleted.
'   STGM constants: See https://docs.microsoft.com/en-us/windows/desktop/stg/stgm-constants
' - dwAttributes : One or more flag values that specify file attributes in the case that a
'   new file is created. For a complete list of possible values, see the dwFlagsAndAttributes
'   parameter of the CreateFile function.
'   See: https://docs.microsoft.com/en-us/windows/desktop/api/fileapi/nf-fileapi-createfilea
' - fCreate : A BOOL value that helps specify, in conjunction with grfMode, how existing
'   files should be treated when creating the stream. See Remarks for details.
'   https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-shcreatestreamonfileex
' ========================================================================================
PRIVATE CONSTRUCTOR CFileStream (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD = STGM_READ, _
BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS WINBOOL = FALSE)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   this.SetResult(AfxSHCreateStreamOnFileEx(pwszFile, grfMode, dwAttributes, fCreate, @m_pStream))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes an instance of the class from an existing IStream object and attaches it.
' ========================================================================================
PRIVATE CONSTRUCTOR CFileStream (BYVAL pstm AS IStream PTR)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   IF pstm = NULL THEN EXIT CONSTRUCTOR
   m_pStream = pstm
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CFileStream destructor
' ========================================================================================
PRIVATE DESTRUCTOR CFileStream
   ' // Release the IStrean interface
   IF m_pStream THEN m_pStream->lpvtbl->Release(m_pStream)
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CFileStream.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CFileStream.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assignment operator
' ========================================================================================
PRIVATE OPERATOR CFileStream.LET (BYVAL pstm AS IStream PTR)
   m_Result = 0
   IF pstm = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the stream object
   IF m_pStream THEN m_pStream->lpvtbl->Release(m_pStream)
   ' // Attach the passed stream object to the class
   m_pStream = pstm
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operator
' ========================================================================================
PRIVATE OPERATOR CFileStream.CAST () AS IStream PTR
   m_Result = 0
   OPERATOR = m_pStream
END OPERATOR
' ========================================================================================

' ========================================================================================
' Attaches the passed stream object to the class
' ========================================================================================
PRIVATE FUNCTION CFileStream.Attach (BYVAL pstm AS IStream PTR) AS HRESULT
   m_Result = 0
   IF pstm = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the stream object
   IF m_pStream THEN m_pStream->lpvtbl->Release(m_pStream)
   ' // Attach the passed stream object to the class
   m_pStream = pstm
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the stream object from the class
' ========================================================================================
PRIVATE FUNCTION CFileStream.Detach () AS IStream PTR
   m_Result = 0
   DIM pstm AS IStream PTR = m_pStream
   m_pStream = NULL
   RETURN pstm
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens or creates a file and retrieves a stream to read or write to that file.
' Parameters:
' - pwszFile : A pointer to a unicode null-terminated string that specifies the file name.
' - grfMode : One or more STGM values that are used to specify the file access mode and how
'   the object that exposes the stream is created and deleted.
'   STGM constants: See https://docs.microsoft.com/en-us/windows/desktop/stg/stgm-constants
' - dwAttributes : One or more flag values that specify file attributes in the case that a
'   new file is created. For a complete list of possible values, see the dwFlagsAndAttributes
'   parameter of the CreateFile function.
'   See: https://docs.microsoft.com/en-us/windows/desktop/api/fileapi/nf-fileapi-createfilea
' - fCreate : A BOOL value that helps specify, in conjunction with grfMode, how existing
'   files should be treated when creating the stream. See Remarks for details.
'   https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-shcreatestreamonfileex
' ========================================================================================
PRIVATE FUNCTION CFileStream.Open (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD = STGM_READ, _
BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS WINBOOL = FALSE) AS HRESULT
   IF m_pStream <> NULL THEN m_pStream->lpvtbl->Release(m_pStream) : m_pStream = NULL
   RETURN this.SetResult(AfxSHCreateStreamOnFileEx(pwszFile, grfMode, dwAttributes, fCreate, @m_pStream))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the stream object.
' ========================================================================================
PRIVATE SUB CFileStream.Close
   ' // Release the IStrean interface
   IF m_pStream THEN m_pStream->lpvtbl->Release(m_pStream)
   m_Result = 0
END SUB
' ========================================================================================

' ========================================================================================
' Reads a specified number of bytes from the stream object into memory, starting at the
' current seek pointer.
' - pv : A pointer to the buffer which the stream data is read into.
' - cb : The number of bytes of data to read from the stream object.
' - pcbRead: A pointer to a ULONG variable that receives the actual number of bytes read
'   from the stream object. Note: The number of bytes read may be zero.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbRead AS ULONG PTR) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Read(m_pStream, pv, cb, pcbRead))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER)
   DIM cbRead AS ULONG
   this.SetResult(m_pStream->lpvtbl->Read(m_pStream, pv, cb, @cbRead))
   RETURN cbRead
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads a specified number of characters from the stream object into memory, starting at the
' current seek pointer.
' - numChars : The number of characters to read from the stream object.
'   Pass -1 to read all the characters from the current seek position.
' ========================================================================================
PRIVATE FUNCTION CFileStream.ReadTextA (BYVAL numChars AS LONG) AS STRING
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN ""
   DIM cbRead AS ULONG
   IF numChars < 0 THEN numChars = this.GetSize - this.GetSeekPosition
   DIM strText AS STRING = SPACE(numChars)
   this.SetResult(m_pStream->lpvtbl->Read(m_pStream, STRPTR(strText), numChars, @cbRead))
   IF cbRead < LEN(strText) THEN strText = LEFT(strText, cbRead)
   RETURN strText
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.ReadTextW (BYVAL numChars AS LONG) AS CWSTR
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN ""
   DIM cbRead AS ULONG
   IF numChars < 0 THEN numChars = (this.GetSize \ 2) - this.GetSeekPosition * 2
   DIM cwsText AS CWSTR = WSPACE(numChars)
   this.SetResult(m_pStream->lpvtbl->Read(m_pStream, *cwsText, numChars * 2, @cbRead))
   IF cbRead < LEN(cwsText) THEN cwsText = LEFT(**cwsText, cbRead)
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a specified number of bytes into the stream object starting at the current seek pointer.
' - pv: A pointer to the buffer that contains the data that is to be written to the stream.
'   A valid pointer must be provided for this parameter even when cb is zero.
' - cb : The number of bytes of data to attempt to write into the stream. This value can be zero.
' - pcbWritten : A pointer to a ULONG variable where this method writes the actual number
'   of bytes written to the stream object. The caller can set this pointer to NULL, in which
'   case this method does not provide the actual number of bytes written.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbWritten AS ULONG PTR) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   IF pv = NULL THEN RETURN this.SetResult(E_INVALIDARG)
   RETURN this.SetResult(m_pStream->lpvtbl->Write(m_pStream, pv, cb, pcbWritten))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER)
   IF pv = NULL THEN this.SetResult(E_INVALIDARG)
   DIM cbWritten AS ULONG
   this.SetResult(m_pStream->lpvtbl->Write(m_pStream, pv, cb, @cbWritten))
   RETURN cbWritten
END FUNCTION
' ========================================================================================
' ========================================================================================
' Writes a string at the current seek position.
' ========================================================================================
PRIVATE FUNCTION CFileStream.WriteTextA (BYREF strText AS STRING) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM cbWritten AS ULONG
   this.SetResult(m_pStream->lpvtbl->Write(m_pStream, STRPTR(strText), LEN(strText), @cbWritten))
   RETURN cbWritten
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.WriteTextW (BYREF wszText AS WSTRING) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM cbWritten AS ULONG
   this.SetResult(m_pStream->lpvtbl->Write(m_pStream, STRPTR(wszText), LEN(wszText) * 2, @cbWritten))
   RETURN cbWritten \ 2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the seek pointer to a new location. The new location is relative to either the
' beginning of the stream, the end of the stream, or the current seek pointer.
' - dlibMove : The displacement to be added to the location indicated by the dwOrigin
'   parameter. If dwOrigin is STREAM_SEEK_SET, this is interpreted as an unsigned value
'   rather than a signed value.
' - dwOrigin : The origin for the displacement specified in dlibMove. The origin can be the
'   beginning of the file (STREAM_SEEK_SET), the current seek pointer (STREAM_SEEK_CUR), or
'   the end of the file (STREAM_SEEK_END). For more information about values, see the
'   STREAM_SEEK enumeration.
' - plibNewPosition : A pointer to the location where this method writes the value of the
'   new seek pointer from the beginning of the stream. You can set this pointer to NULL.
'   In this case, this method does not provide the new seek pointer.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Seek (BYVAL dlibMove AS ULONGINT, BYVAL dwOrigin AS DWORD, BYVAL plibNewPosition AS ULONGINT PTR = NULL) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM lidlibMove AS LARGE_INTEGER : lidlibMove.QuadPart = dlibMove
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, lidlibMove, dwOrigin, @libNewPosition))
   IF plibNewPosition THEN *plibNewPosition = libNewPosition.QuadPart
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the seek position.
' ========================================================================================
PRIVATE FUNCTION CFileStream.GetSeekPosition () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_CUR, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the seek position at the beginning of the stream.
' ========================================================================================
PRIVATE FUNCTION CFileStream.ResetSeekPosition () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_SET, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the seek position at the end of the stream.
' ========================================================================================
PRIVATE FUNCTION CFileStream.SeekAtEndOfFile () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_END, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the size of the stream object.
' libNewSize : Specifies the new size, in bytes, of the stream.
' ========================================================================================
PRIVATE FUNCTION CFileStream.SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulilibNewSize AS ULARGE_INTEGER : ulilibNewSize.QuadPart = libNewSize
   RETURN this.SetResult(m_pStream->lpvtbl->SetSize(m_pStream, ulilibNewSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the size of the stream.
' ========================================================================================
PRIVATE FUNCTION CFileStream.GetSize () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM tstatstg AS STATSTG
   this.SetResult(m_pStream->lpvtbl->Stat(m_pStream, @tstatstg, STATFLAG_NONAME))
   RETURN tstatstg.cbSize.QuadPart
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies a specified number of bytes from the current seek pointer in the stream to the
' current seek pointer in another stream.
' - pstm : A pointer to the destination stream. The stream pointed to by pstm can be a new
'   stream or a clone of the source stream.
' - pcbRead : A pointer to the location where this method writes the actual number of bytes
'   read from the source. You can set this pointer to NULL. In this case, this method does
'   not provide the actual number of bytes read.
' - pcbWritten : A pointer to the location where this method writes the actual number of
'   bytes written to the destination. You can set this pointer to NULL. In this case, this
'   method does not provide the actual number of bytes written.
' ========================================================================================
PRIVATE FUNCTION CFileStream.CopyTo (BYVAL pstm AS IStream PTR, BYVAL cb AS ULONGINT, _
BYVAL pcbRead AS ULONGINT PTR = NULL, BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulicb AS ULARGE_INTEGER: ulicb.QuadPart = cb
   DIM ulicbRead AS ULARGE_INTEGER, ulicbWritten AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->CopyTo(m_pStream, pstm, ulicb, @ulicbRead, @ulicbWritten))
   IF pcbRead THEN *pcbRead = ulicbRead.QuadPart
   IF pcbWritten THEN *pcbWritten = ulicbWritten.QuadPart
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The Commit method ensures that any changes made to a stream object open in transacted
' mode are reflected in the parent storage.
' - grfCommitFlags : Controls how the changes for the stream object are committed.
'   See the STGC enumeration for a definition of these values.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Commit (BYVAL grfCommitFlags AS DWORD) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Commit(m_pStream, grfCommitFlags))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The Revert method discards all changes that have been made to a transacted stream since
' the last IStream.Commit call. On streams open in direct mode this method has no effect.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Revert () AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Revert(m_pStream))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Restricts access to a specified range of bytes in the stream.
' - libOffset : Integer that specifies the byte offset for the beginning of the range.
' - cb : Integer that specifies the length of the range, in bytes, to be restricted.
' - dwLockType : Specifies the restrictions being requested on accessing the range.
'   - LOCK_WRITE: If this lock is granted, the specified range of bytes can be opened and
'     read any number of times, but writing to the locked range is prohibited except for the
'     owner that was granted this lock.
'   - LOCK_EXCLUSIVE : If this lock is granted, writing to the specified range of bytes is
'     prohibited except by the owner that was granted this lock.
'   - LOCK_ONLYONCE: If this lock is granted, no other LOCK_ONLYONCE lock can be obtained
'     on the range. Usually this lock type is an alias for some other lock type. Thus,
'     specific implementations can have additional behavior associated with this lock type.
' ========================================================================================
PRIVATE FUNCTION CFileStream.LockRegion (BYVAL libOffset AS ULONGINT, BYVAL cb AS ULONGINT, BYVAL dwLockType AS DWORD) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulilibOffset AS ULARGE_INTEGER : ulilibOffset.QuadPart = libOffset
   DIM ulicb AS ULARGE_INTEGER : ulicb.QuadPart = cb
   RETURN this.SetResult(m_pStream->lpvtbl->LockRegion(m_pStream, ulilibOffset, ulicb, dwLockType))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes the access restriction on a range of bytes previously restricted with IStream.LockRegion.
' - libOffset : Integer that specifies the byte offset for the beginning of the range.
' - cb : Integer that specifies the length of the range, in bytes, to be restricted.
' - dwLockType : Specifies the access restrictions previously placed on the range.
' Remarks:
'   IStream.UnlockRegion unlocks a region previously locked with the IStream.LockRegion
'   method. Locked regions must later be explicitly unlocked by calling IStream.UnlockRegion
'   with exactly the same values for the libOffset, cb, and dwLockType parameters. The
'   region must be unlocked before the stream is released. Two adjacent regions cannot be
'   locked separately and then unlocked with a single unlock call.
' ========================================================================================
PRIVATE FUNCTION CFileStream.UnlockRegion (BYVAL libOffset AS ULONGINT, BYVAL cb AS ULONGINT, BYVAL dwLockType AS DWORD) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulilibOffset AS ULARGE_INTEGER : ulilibOffset.QuadPart = libOffset
   DIM ulicb AS ULARGE_INTEGER : ulicb.QuadPart = cb
   RETURN this.SetResult(m_pStream->lpvtbl->UnlockRegion(m_pStream, ulilibOffset, ulicb, dwLockType))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the STATSTG structure for this stream.
' - pstatstg : Pointer to a STATSTG structure where this method places information about
'   this stream object.
' - grfStatFlag : Specifies that this method does not return some of the members in the
'   STATSTG structure, thus saving a memory allocation operation. Values are taken from
'   the STATFLAG enumeration.
'   - STATFLAG_DEFAULT : Requests that the statistics include the pwcsName member of the
'     STATSTG structure.
'   - STATFLAG_NONAME : Requests that the statistics not include the pwcsName member of
'     the STATSTG structure. If the name is omitted, there is no need for the ILockBytes.Stat,
'     IStorage.Stat, and IStream.Stat methods methods to allocate and free memory for the
'     string value of the name, therefore the method reduces time and resources used in an
'     allocation and free operation.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Stat (BYVAL pstatstg AS STATSTG PTR, BYVAL grfStatFlag AS DWORD) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Stat(m_pStream, pstatstg, grfStatFlag))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.Stat (BYVAL grfStatFlag AS DWORD) AS STATSTG
   DIM tstatstg AS STATSTG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN tstatstg
   this.SetResult(m_pStream->lpvtbl->Stat(m_pStream, @tstatstg, grfStatFlag))
   RETURN tstatstg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new stream object with its own seek pointer that references the same bytes as
' the original stream.
' - ppstm : When successful, pointer to the location of an IStream pointer to the new stream
'   object. If an error occurs, this parameter is NULL.
' Remarks:
'    The Clone method creates a new stream object for accessing the same bytes but using a
'    separate seek pointer. The new stream object sees the same data as the source-stream
'    object. Changes written to one object are immediately visible in the other. Range
'    locking is shared between the stream objects.
'    The initial setting of the seek pointer in the cloned stream instance is the same as
'    the current setting of the seek pointer in the original stream at the time of the
'    clone operation.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Clone(m_pStream, ppstm))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.Clone () AS IStream PTR
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN NULL
   DIM ppstm AS ISTREAM PTR
   this.SetResult(m_pStream->lpvtbl->Clone(m_pStream, @ppstm))
   RETURN ppstm
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last status code.
' ========================================================================================
PRIVATE FUNCTION CFileStream.GetErrorInfo () AS STRING
   IF SUCCEEDED(m_Result) THEN RETURN ""
   DIM s AS STRING = "Error &h" & HEX(m_Result, 8)
   SELECT CASE m_Result
      CASE E_POINTER : s += ": Null pointer"
      CASE E_INVALIDARG : s += ": Null pointer"
      CASE STG_E_INVALIDFUNCTION : s += ": Unable to perform requested operation"
      CASE STG_E_FILENOTFOUND : s += ": File not found"
      CASE STG_E_PATHNOTFOUND : s += ": Path not found"
      CASE STG_E_TOOMANYOPENFILES : s += ": Too many open files"
      CASE STG_E_ACCESSDENIED : s += ": Access denid"
      CASE STG_E_INVALIDHANDLE : s += ": Invalid handle"
      CASE STG_E_INSUFFICIENTMEMORY : s += ": Insufficient memory"
      CASE STG_E_INVALIDPOINTER : s += ": Invalid pointer"
      CASE STG_E_NOMOREFILES : s += ": No more files"
      CASE STG_E_DISKISWRITEPROTECTED : s += ": Disk write protected"
      CASE STG_E_SEEKERROR : s += ": Seek error"
      CASE STG_E_WRITEFAULT : s += ": Write fault"
      CASE STG_E_READFAULT : s += ": Read fault"
      CASE STG_E_SHAREVIOLATION : s += ": Share violation"
      CASE STG_E_LOCKVIOLATION : s += ": Lock violation"
      CASE STG_E_FILEALREADYEXISTS : s += ": File already exists"
      CASE STG_E_INVALIDPARAMETER : s += ": Invalid parameter"
      CASE STG_E_MEDIUMFULL : s += ": Medium full"
      CASE STG_E_PROPSETMISMATCHED : s += ": Property set mismatched"
      CASE STG_E_ABNORMALAPIEXIT : s += ": Abnormal API call exit"
      CASE STG_E_INVALIDHEADER : s += ": Invalid header"
      CASE STG_E_INVALIDNAME : s += ": Invalid name"
      CASE STG_E_UNKNOWN : s += ": Unknown error"
      CASE STG_E_UNIMPLEMENTEDFUNCTION : s += ": Unimplemented function"
      CASE STG_E_INVALIDFLAG : s += ": Invalid flag"
      CASE STG_E_INUSE : s += ": Object is busy"
      CASE STG_E_NOTCURRENT : s += ": The storage has been changed since the last commit"
      CASE STG_E_REVERTED : s += ": The object has ceased to exist"
      CASE STG_E_CANTSAVE : s += ": Can't save"
      CASE STG_E_OLDFORMAT : s += ": The compound file was produced with an incompatible version of storage"
      CASE STG_E_OLDDLL : s += ": The compound file was produced with a newer version of storage"
      CASE STG_E_SHAREREQUIRED : s += ": Share.exe or equivalent is required for operation"
      CASE STG_E_NOTFILEBASEDSTORAGE : s += ": Illegal operation called on non-file based storage"
      CASE STG_E_EXTANTMARSHALLINGS : s += ": Illegal operation called on object with extant marshallings"
      CASE STG_E_DOCFILECORRUPT : s += ": The docfile has been corrupted"
      CASE STG_E_BADBASEADDRESS : s += ": OLE32.DLL has been loaded at the wrong address"
      CASE STG_E_DOCFILETOOLARGE : s += ": The compound file is too large for the current implementation"
      CASE STG_E_NOTSIMPLEFORMAT : s += ": The compound file was not created with the STGM_SIMPLE flag"
      CASE STG_E_INCOMPLETE : s += ": The file download was aborted abnormally. The file is incomplete."
      CASE STG_E_TERMINATED : s += ": The file download has been terminated"
      CASE STG_S_CONVERTED : s += ": The underlying file was converted to compound file format"
      CASE STG_S_BLOCK : s += ": The storage operation should block until more data is available"
      CASE STG_S_RETRYNOW : s += ": The storage operation should retry immediately"
      CASE STG_S_MONITORING : s += ": The notified event sink will not influence the storage operation"
      CASE STG_S_MULTIPLEOPENS : s += ": Multiple opens prevent consolidated. (commit succeeded)."
      CASE STG_S_CONSOLIDATIONFAILED : s += ": Consolidation of the storage file failed. (commit succeeded)."
      CASE STG_S_CANNOTCONSOLIDATE : s += ": Consolidation of the storage file is inappropriate. (commit succeeded)."
      CASE STG_E_STATUS_COPY_PROTECTION_FAILURE : s += ": Generic Copy Protection Error"
      CASE STG_E_CSS_AUTHENTICATION_FAILURE : s += ": Copy Protection Error - DVD CSS Authentication failed."
      CASE STG_E_CSS_KEY_NOT_PRESENT : s += ": Copy Protection Error - The given sector does not have a valid CSS key"
      CASE STG_E_CSS_KEY_NOT_ESTABLISHED : s += ": Copy Protection Error - DVD session key not established"
      CASE STG_E_CSS_SCRAMBLED_SECTOR : s += ": Copy Protection Error - The read failed because the sector is encrypted"
      CASE STG_E_CSS_REGION_MISMATCH : s += ": Copy Protection Error - The current DVD's region does not correspond to the region setting of the drive"
      CASE STG_E_RESETS_EXHAUSTED : s += ": Copy Protection Error - The drive's region setting may be permanent or the number of user resets has been exhausted"
   END SELECT
   RETURN s
END FUNCTION
' ========================================================================================

END NAMESPACE
