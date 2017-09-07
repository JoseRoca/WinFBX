' ########################################################################################
' TypeLib Browser
' File: TLB_ENUMTLBS.INC
' Contents: TypeLib Browser typelibs enumeration
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' ========================================================================================
' Searches for the win32 subkey
' ========================================================================================
FUNCTION TLB_RegSearchWin32 (BYVAL pwszKey AS WSTRING PTR) AS CWSTR

   IF pwszKey = NULL THEN RETURN ""

   ' // Recursively searches for the win directory
   DIM hr AS LONG, hKey AS HKEY, dwIdx AS DWORD, ft AS FILETIME
   DIM wszKeyName AS WSTRING * MAX_PATH, wszClass AS WSTRING * MAX_PATH
   DIM cchName AS DWORD = MAX_PATH, cchClass AS DWORD = MAX_PATH
   DIM wszKey AS WSTRING * MAX_PATH = *pwszKey
   DO
      wszKeyName = "" : cchName = MAX_PATH : wszClass = "" : cchClass = MAX_PATH
      hr = RegOpenKeyExW (HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
      IF hr <> ERROR_SUCCESS THEN RETURN ""
      IF hKey = NULL THEN RETURN ""
      hr = RegEnumKeyExW(hKey, dwIdx, @wszKeyName, @cchName, 0, @wszClass, @cchClass, @ft)
      IF hr <> S_OK THEN EXIT DO
      IF UCASE(wszKeyName) = "WIN32" THEN EXIT DO
      dwIdx += 1
   LOOP WHILE hr = S_OK

   ' // Closes the registry and returns the result
   RegCloseKey hKey
   IF hr <> S_OK OR LEN(wszKeyName) = 0 THEN RETURN ""
   RETURN wszKey

END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches for the win64 subkey
' ========================================================================================
FUNCTION TLB_RegSearchWin64 (BYVAL pwszKey AS WSTRING PTR) AS CWSTR

   IF pwszKey = NULL THEN RETURN ""

   ' // Recursively searches for the win directory
   DIM hr AS LONG, hKey AS HKEY, dwIdx AS DWORD, ft AS FILETIME
   DIM wszKeyName AS WSTRING * MAX_PATH, wszClass AS WSTRING * MAX_PATH
   DIM cchName AS DWORD = MAX_PATH, cchClass AS DWORD = MAX_PATH
   DIM wszKey AS WSTRING * MAX_PATH = *pwszKey
   DO
      wszKeyName = "" : cchName = MAX_PATH : wszClass = "" : cchClass = MAX_PATH
      hr = RegOpenKeyExW (HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
      IF hr <> ERROR_SUCCESS THEN RETURN ""
      IF hKey = NULL THEN RETURN ""
      hr = RegEnumKeyExW(hKey, dwIdx, @wszKeyName, @cchName, 0, @wszClass, @cchClass, @ft)
      IF hr <> S_OK THEN EXIT DO
      IF UCASE(wszKeyName) = "WIN64" THEN EXIT DO
      dwIdx += 1
   LOOP WHILE hr = S_OK

   ' // Closes the registry and returns the result
   RegCloseKey hKey
   IF hr <> S_OK OR LEN(wszKeyName) = 0 THEN RETURN ""
   RETURN wszKey

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the path of the typelib.
' ========================================================================================
FUNCTION TLB_RegEnumDirectory (BYVAL pwszKey AS WSTRING PTR) AS CWSTR

   IF pwszKey = NULL THEN RETURN ""

   ' // Searches the HKEY_CLASSES_ROOT\TypeLib\<LIBID> node.
   DIM hKey AS HKEY, wszKey AS WSTRING * MAX_PATH = *pwszKey
   DIM hr AS LONG = RegOpenKeyExW(HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN RETURN ""
   IF hKey = 0 THEN RETURN ""
   DIM dwIdx AS DWORD, wszKeyName AS WSTRING * MAX_PATH, wszClass AS WSTRING * MAX_PATH, ft AS FILETIME
   DIM cchName AS DWORD = MAX_PATH, cchClass AS DWORD = MAX_PATH, wszSubkey AS WSTRING * MAX_PATH
   DO
      wszKeyName = "" : cchName = MAX_PATH : wszClass = "" : cchClass = MAX_PATH
      hr = RegEnumKeyExW(hKey, dwIdx, @wszKeyName, @cchName, 0, @wszClass, @cchClass, @ft)
      IF hr <> S_OK THEN EXIT DO
#ifdef __FB_64BIT__
      wszSubkey = TLB_RegSearchWin64(wszKey & "\" & wszKeyName)
      IF LEN(wszSubkey) THEN wszKey = wszSubkey & "\" & "win64"
#else
      wszSubkey = TLB_RegSearchWin32(wszKey & "\" & wszKeyName)
      IF LEN(wszSubkey) THEN wszKey = wszSubkey & "\" & "win32"
#endif
#ifdef __FB_64BIT__
      ' // Not all the typelibs have separate entries in the win64 subkey
      ' // See: https://msdn.microsoft.com/en-us/library/windows/desktop/ms724072(v=vs.85).aspx
      ' // If not found in the win64 subkey search in the win32 subkey
      IF LEN(wszSubkey) = 0 THEN wszSubkey = TLB_RegSearchWin32(wszKey & "\" & wszKeyName)
      IF LEN(wszSubkey) THEN wszKey = wszSubkey & "\" & "win32"
#endif
      IF LEN(wszSubkey) THEN EXIT DO
      dwIdx += 1
   LOOP
   RegCloseKey hKey
   IF hr <> S_OK OR LEN(wszSubkey) = 0 THEN RETURN ""

   hKey = NULL
   hr = RegOpenKeyExW(HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN RETURN ""
   DIM keyType AS DWORD
   DIM wszValueName AS WSTRING * MAX_PATH
   DIM wszKeyValue  AS WSTRING * MAX_PATH
   DIM cValueName AS DWORD = MAX_PATH
   DIM cbData AS DWORD = MAX_PATH
   dwIdx = 0
   hr = RegEnumValueW(hKey, dwIdx, @wszValueName, @cValueName, NULL, @keyType, cast(BYTE PTR, @wszKeyValue), @cbData)

   ' // Closes the registry and returns the value
   RegCloseKey hKey
   RETURN wszKeyValue

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the different versions of the typelib.
' Parameter
' - hListView = Handle of the list view.
' - pwszLibId = Library guid.
' Return value: TRUE or FALSE.
' ========================================================================================
FUNCTION TLB_RegEnumVersions (BYVAL hListView AS HWND, BYVAL pwszLibId AS WSTRING PTR) AS BOOLEAN

   IF hListView = NULL OR pwszLibId = NULL THEN EXIT FUNCTION

   ' // Searches the HKEY_CLASSES_ROOT\TypeLib\<LIBID> node.
   DIM hKey AS HKEY
   DIM wszKey AS WSTRING * MAX_PATH = "TypeLib\" & *pwszLibId
   DIM hr AS LONG = RegOpenKeyExW(HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN RETURN FALSE
   IF hKey = NULL THEN RETURN FALSE

   ' // Opens the subtrees of the different versions of the TyleLib library
   DIM dwIdx AS DWORD, wszKeyName AS WSTRING * MAX_PATH, wszClass AS WSTRING * MAX_PATH, ft AS FILETIME
   DIM cchName AS DWORD = MAX_PATH, cchClass AS DWORD = MAX_PATH
   DO
      wszKeyName = "" : cchName = MAX_PATH : wszClass = "" : cchClass = MAX_PATH
      DIM hr AS LONG = RegEnumKeyExW(hKey, dwIdx, @wszKeyName, @cchName, 0, @wszClass, @cchClass, @ft)
      IF hr <> ERROR_SUCCESS THEN RETURN FALSE
      ' // Gets the default value
      DIM hVerKey AS HKEY, wszSubKey AS WSTRING * MAX_PATH = wszKey & "\" & wszKeyName
      hr = RegOpenKeyExW(HKEY_CLASSES_ROOT, @wszSubKey, 0, KEY_READ, @hVerKey)
      IF hr <> ERROR_SUCCESS THEN RETURN FALSE
      DIM wszVer AS WSTRING * MAX_PATH = wszKeyName
      ' // Enumerate the entries until the default one, if any, is found.
      ' // RegEnumValue returns values in any order. This includes unnamed values.
      ' // A key may have one unnamed value. An unnamed value is displayed as (Default)
      ' // in Regedit.exe. If an unnamed value doesn't exist under a given key, it is
      ' // displayed as (value not set) in Regedit.exe.
      ' // Only existing unnamed values can be enumerated. If an unnamed value is enumerated, the
      ' // RegEnumValue function sets wszValueName to an empty string ("") and it sets cValueName to 0.
      DIM verIdx AS DWORD, cValueName AS DWORD, cbData AS DWORD, keyType AS DWORD
      DIM wszValueName AS WSTRING * 16383, wszKeyValue AS WSTRING * MAX_PATH
      DO
         cValueName = 16383 : cbData = MAX_PATH * 2 : wszValueName = "" : wszKeyValue = ""
         hr = RegEnumValueW(hVerKey, verIdx, @wszValueName, @cValueName, NULL, @keyType, cast(BYTE PTR, @wszKeyValue), @cbData)
         IF LEN(wszValueName) = 0 THEN EXIT DO   ' // This is the default, unnamed value
         IF hr <> ERROR_SUCCESS THEN EXIT DO
         ' // Increment the index of the value to be retrieved.
         verIdx += 1
      LOOP
      ' // Closes the handle of the registry key
      RegCloseKey hVerKey
      DIM wszDesc AS WSTRING * MAX_PATH
      ' // If wszValueName is an empty string, assume that the value name is the key value.
      IF LEN(wszValueName) = 0 THEN wszDesc = wszKeyValue ELSE wszDesc = wszValueName
      ' // Increment the index of the subkey
      dwIdx += 1
      ' // Find the path of the type library
      DIM wszPath AS WSTRING * MAX_PATH = TLB_RegEnumDirectory(wszKey & "\" & wszKeyName)
      ' // If there is not path, skip the typelib because we can't retrieve it
      IF LEN(wszPath) = 0 THEN CONTINUE DO
      ' // Remove double quotes (if any)
      wszPath = AfxStrRemove(wszPath, """")
      ' // Convert short paths to long paths
      ' // Dont use it with all files or these ending with version numbers
      ' // (a \ and a number) will we skipped.
      IF INSTR(wszPath, "%") THEN
         DIM wszDest AS WSTRING * MAX_PATH, cbLen AS DWORD
         cbLen = ExpandEnvironmentStringsW(@wszPath, @wszDest, MAX_PATH)
         IF cbLen THEN wszPath = wszDest
      END IF
      IF INSTR(wszPath, "~") <> 0 THEN
         DIM cbPath AS DWORD = LEN(wszPath)
         cbPath = GetLongPathNameW(wszPath, wszPath, cbPath)
      END IF
      DIM pathPos AS LONG = INSTRREV(wszPath, "\")
      DIM wszFile AS WSTRING * MAX_PATH = MID(wszPath, pathPos + 1)
      ' // Some have an added backslah and a number
      IF INSTR(wszFile, ".") = 0 THEN
         DIM wszTmpPath AS WSTRING * MAX_PATH = LEFT(wszPath, pathpos - 1)
         pathPos = INSTRREV(wszTmpPath, "\", LEN(wszFile) - 3)
         wszFile = MID(wszPath, pathPos + 1)
      END IF
      IF LEN(wszFile) = 0 THEN CONTINUE DO  ' // Skip files without a full path
      ' // Skip .OCA files
      DIM wszTemp AS WSTRING * MAX_PATH = wszPath
      IF MID(wszTemp, LEN(wszTemp) - 2, 1) = "\" THEN wszTemp = LEFT(wszTemp, LEN(wszTemp) - 2)
      IF MID(wszTemp, LEN(wszTemp) - 3, 1) = "\" THEN wszTemp = LEFT(wszTemp, LEN(wszTemp) - 3)
      ' // .OCA files are created by Visual Basic (we don't need them)
      IF UCASE(RIGHT(wszTemp, 4)) = ".OCA" THEN wszPath = ""
      IF LEN(wszPath) THEN
         ' // If the description is empty, use the file name
         IF wszDesc = "" THEN wszDesc = "[" & wszFile & "]"
         ' // Add the version number
         wszDesc += " (Ver " & wszVer & ")"
         ' // Add the items to the list view
         DIM lItemIdx AS LONG = ListView_AddItem(hListView, 0, 0, @wszFile)
         ListView_SetItemText(hListView, lItemIdx, 1, @wszDesc)
         ListView_SetItemText(hListView, lItemIdx, 2, @wszPath)
         ListView_SetItemText(hListView, lItemIdx, 3, pwszLibId)
      END IF
   LOOP

   ' // Closes the registry key
   RegCloseKey hKey

END FUNCTION
' ========================================================================================

' ========================================================================================
' Enumerates all the typelibs.
' Parameter
' - hListView = Handle of the list view.
' Return value: TRUE or FALSE.
' ========================================================================================
FUNCTION TLB_RegEnumTypeLibs (BYVAL hListView AS HWND) AS BOOLEAN

   IF hListView = NULL THEN RETURN FALSE

   ' // Opens the HKEY_CLASSES_ROOT\TypeLib subtree
   DIM hKey AS HKEY
   DIM wszKey AS WSTRING * MAX_PATH = "TypeLib"
   DIM hr AS LONG = RegOpenKeyExW(HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN RETURN FALSE
   IF hKey = NULL THEN RETURN FALSE

   ' // Parses all the TypeLib subtree and gets the CLSIDs of all the TypeLibs
   DIM dwIdx AS DWORD, wszKeyName AS WSTRING * MAX_PATH, wszClass AS WSTRING * MAX_PATH, ft AS FILETIME
   DIM cchName AS DWORD = MAX_PATH, cchClass AS DWORD = MAX_PATH
   DO
      wszKeyName = "" : cchName = MAX_PATH : wszClass = "" : cchClass = MAX_PATH
      hr = RegEnumKeyExW(hKey, dwIdx, @wszKeyName, @cchName, 0, @wszClass, @cchClass, @ft)
      IF hr <> ERROR_SUCCESS THEN EXIT DO
      TLB_RegEnumVersions hListView, @wszKeyName
      dwIdx += 1
   LOOP
   ' // Closes the registry
   RegCloseKey hKey

   RETURN TRUE

END FUNCTION
' ========================================================================================
