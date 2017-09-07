'#CONSOLE ON
#include once "windows.bi"
#include once "Afx/CWmiDisp.inc"
using Afx

' // Connect to WMI using a moniker
' // Note: $ is used to avoid the pedantic warning of the compiler about escape characters
DIM pServices AS CWmiServices = $"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"
IF pServices.ServicesPtr = NULL THEN END

' // Execute a query
DIM hr AS HRESULT = pServices.ExecQuery("SELECT * FROM Win32_BIOS")
IF hr <> S_OK THEN PRINT AfxWmiGetErrorCodeText(hr) : SLEEP : END

' // Get the number of objects retrieved
DIM nCount AS LONG = pServices.ObjectsCount
print "Count: ", nCount
IF nCount = 0 THEN PRINT "No objects found" : SLEEP : END

' // Enumerate the objects using the standard IEnumVARIANT enumerator (NextObject method)
' // and retrieve the properties using the CDispInvoke class.
DIM pDispServ AS CDispInvoke = pServices.NextObject
IF pDispServ.DispPtr THEN
   PRINT "BIOS version: : "; pDispServ.Get("BIOSVersion").ToStr
   PRINT "BIOS characteristics:"; pDispServ.Get("BIOSCharacteristics").ToStr
   PRINT "Build number: "; pDispServ.Get("BuildNumber").ToStr
   PRINT "Caption: "; pDispServ.Get("Caption").ToStr
   PRINT "Current language: "; pDispServ.Get("CurrentLanguage").ToStr
   PRINT "Description: "; pDispServ.Get("Description").ToStr
   PRINT "Identification code: "; pDispServ.Get("IdentificationCode").ToStr
   PRINT "Installable languages: "; pDispServ.Get("InstallableLanguages").ToStr
   PRINT "Install date: "; pDispServ.Get("InstallDate").ToStr
   PRINT "Language edition: "; pDispServ.Get("LanguageEdition").ToStr
   PRINT "List of languages: "; pDispServ.Get("ListOfLanguages").ToStr
   PRINT "Manufacturer: "; pDispServ.Get("Manufacturer").ToStr
   PRINT "Other target OS: "; pDispServ.Get("OtherTargetOS").ToStr
   PRINT "Primary BIOS: "; pDispServ.Get("PrimaryBIOS").ToStr
   PRINT "Release date: "; pServices.WmiDateToStr(pDispServ.Get("ReleaseDate"), "dd-MM-yyyy")
   PRINT "Serial number: "; pDispServ.Get("SerialNumber").ToStr
   PRINT "SMBIOS BIOS version: "; pDispServ.Get("SMBIOSBIOSVersion").ToStr
   PRINT "SMBIOS major version: "; pDispServ.Get("SMBIOSMajorVersion").ToStr
   PRINT "SMBIOS minor version: "; pDispServ.Get("SMBIOSMinorVersion").ToStr
   PRINT "SMBIOS present: "; pDispServ.Get("SMBIOSPresent").ToStr
   PRINT "Software element ID: "; pDispServ.Get("SoftwareElementID").ToStr
   PRINT "Software element state: "; pDispServ.Get("SoftwareElementState").ToStr
   PRINT "Target operating system: "; pDispServ.Get("TargetOperatingSystem").ToStr
   PRINT "Version: "; pDispServ.Get("Version").ToStr
END IF

PRINT
PRINT "Press any key..."
SLEEP
