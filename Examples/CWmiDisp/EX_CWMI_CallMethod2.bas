'#CONSOLE ON
#include once "windows.bi"
#include once "Afx/CWmiDisp.inc"
using Afx

' // Connect to WMI using a moniker
DIM pServices AS CWmiServices = ( _
   $"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2:Win32_Process")
IF pServices.ServicesPtr = NULL THEN END
 
' // Assign the WMI services object pointer to CDispInvoke
' // CWmiServices.ServicesObj returns an AddRefed pointer, whereas CWmiServices.ServicesPtr not.
DIM pDispServices AS CDispInvoke = CDispInvoke(pServices.ServicesObj)

' // Note: Although the WMI documentation says that this OUT parameter is an UInt32,
' // it only works if I use "LONG".
DIM ProcessId AS LONG
pDispServices.Invoke("Create", 4, CVAR("notepad.exe"), , , CVAR(@ProcessId, "LONG"))
PRINT "Process id: ", ProcessId

PRINT
PRINT "Press any key..."
SLEEP
