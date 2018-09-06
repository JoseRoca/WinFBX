# CADODB

Microsoft ActiveX Data Objects (ADO) enable your client applications to access and manipulate data from a variety of sources through an OLE DB provider. Its primary benefits are ease of use, high speed, low memory overhead, and a small disk footprint. ADO supports key features for building client/server and Web-based applications. 

CADODB are a collection of classes to allow to work with ADO using the FreeBasic compilers.

The **CAdoBase** object, from which the other ADO classes inherit, initializes and uninitializes the COM library and implements two methods, **GetLastResult** and **SetResult** used by the derive classes to store the result codes of ADO calls.

**Folder**: Afx/CADODB<br>
**Files**: CADODB.inc, CAdoCommand.inc, CAdoConnection.inc, CAdoErrors.inc, CAdoFields.inc, CAdoParameters.inc, CAdoProperties.inc, CAdoRecord.inc, CAdoRecordset.inc, CAdoStream.inc.

| Name       | Description |
| ---------- | ----------- |
| [CAdoBase Class](#CAdoBase) | Base class for all the other ADO classes. |
| [ADO Object Model](#ADOObjectModel) | ADO objects and their collections. |
| [ADO Identifiers](#ADOIdentifiers) | PROGIDs, CLSIDs and IIDs. |
| [ADO Enumerations](#ADOEnums) | ADO enumerations. |
| [ADO Dynamic Properties](#ADODynProps) | ADO dynamic properties. |
| [ADO Errors](#ADOErrors) | ADO errors. |
| [ADO IUnknown Interfaces](#ADOIUnknownInterfaces) | ADO IUnknown interfaces. |
| [ADO Events](#ADOEvents) | ADO events. |

# <a name="ADOObjectModel"></a>ADO Object Model

ADO objects and their collections:

```
Connection
  |_ Errors     - Error
  |_ Properties - Property

Command
  |_ Parameters - Parameter
  |_ Properties - Property

Recordset
  |_ Fields - Field
  |_ Properties - Property

Record
  |_ Fields - Field

Stream
```

# <a name="CAdoBase"></a>CAdoBse Class

The CAdoBase class, from which the other ADO classes inherit, initializes and uninitializes the COM library and implements two methods, **GetLastResult** and **SetResult** used by the derived classes to store the result codes of ADO calls.

# <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

# <a name="SetResult"></a>SetResult

Sets the last result code.

```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
```

#### Return value

The result code returned by the last executed method.

# <a name="ADOIdentifiers"></a>ADO Identifiers

```
' ========================================================================================
' ProgIDs (Program identifiers)
' ========================================================================================

' CLSID = {00000507-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Command60 = "ADODB.Command.6.0"
' CLSID = {00000514-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Connection60 = "ADODB.Connection.6.0"
' CLSID = {0000050B-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Parameter60 = "ADODB.Parameter.6.0"
' CLSID = {00000560-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Record60 = "ADODB.Record.6.0"
' CLSID = {00000535-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Recordset60 = "ADODB.Recordset.6.0"
' CLSID = {00000566-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Stream60 = "ADODB.Stream.6.0"

' ========================================================================================
' Version independent ProgIDs
' ========================================================================================

' CLSID = {00000507-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Command = "ADODB.Command"
' CLSID = {00000514-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Connection = "ADODB.Connection"
' CLSID = {0000050B-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Parameter = "ADODB.Parameter"
' CLSID = {00000560-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Record = "ADODB.Record"
' CLSID = {00000535-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Recordset = "ADODB.Recordset"
' CLSID = {00000566-0000-0010-8000-00AA006D2EA4}
CONST PROGID_Stream = "ADODB.Stream"

' ========================================================================================
' ClsIDs (Class identifiers)
' ========================================================================================

CONST CLSID_Command = "{00000507-0000-0010-8000-00AA006D2EA4}"
CONST CLSID_Connection = "{00000514-0000-0010-8000-00AA006D2EA4}"
CONST CLSID_Parameter = "{0000050B-0000-0010-8000-00AA006D2EA4}"
CONST CLSID_Record = "{00000560-0000-0010-8000-00AA006D2EA4}"
CONST CLSID_Recordset = "{00000535-0000-0010-8000-00AA006D2EA4}"
CONST CLSID_Stream = "{00000566-0000-0010-8000-00AA006D2EA4}"

' ========================================================================================
' IIDs (Interface identifiers)
' ========================================================================================

CONST IID_ADO = "{00000534-0000-0010-8000-00AA006D2EA4}"
CONST IID_Collection = "{00000512-0000-0010-8000-00AA006D2EA4}"
CONST IID_Command = "{986761E8-7269-4890-AA65-AD7C03697A6D}"
CONST IID_Connection = "{00001550-0000-0010-8000-00AA006D2EA4}"
CONST IID_DynaCollection = "{00000513-0000-0010-8000-00AA006D2EA4}"
CONST IID_Parameter = "{0000150C-0000-0010-8000-00AA006D2EA4}"
CONST IID_Record = "{00001562-0000-0010-8000-00AA006D2EA4}"
CONST IID_Recordset = "{00001556-0000-0010-8000-00AA006D2EA4}"
CONST IID_Stream = "{00001565-0000-0010-8000-00AA006D2EA4}"
CONST IID_ADOCommandConstruction = "{00000517-0000-0010-8000-00AA006D2EA4}"
CONST IID_ADOConnectionConstruction = "{00000551-0000-0010-8000-00AA006D2EA4}"
CONST IID_ADORecordConstruction = "{00000567-0000-0010-8000-00AA006D2EA4}"
CONST IID_ADORecordsetConstruction = "{00000283-0000-0010-8000-00AA006D2EA4}"
CONST IID_ADOStreamConstruction = "{00000568-0000-0010-8000-00AA006D2EA4}"
CONST IID_ConnectionEvents = "{00001400-0000-0010-8000-00AA006D2EA4}"
CONST IID_ConnectionEventsVt = "{00001402-0000-0010-8000-00AA006D2EA4}"
CONST IID_Error = "{00000500-0000-0010-8000-00AA006D2EA4}"
CONST IID_Errors = "{00000501-0000-0010-8000-00AA006D2EA4}"
CONST IID_Field = "{00001569-0000-0010-8000-00AA006D2EA4}"
CONST IID_Fields = "{00001564-0000-0010-8000-00AA006D2EA4}"
CONST IID_Fields_Deprecated = "{00000564-0000-0010-8000-00AA006D2EA4}"
CONST IID_Parameters = "{0000150D-0000-0010-8000-00AA006D2EA4}"
CONST IID_Properties = "{00000504-0000-0010-8000-00AA006D2EA4}"
CONST IID_Property = "{00000503-0000-0010-8000-00AA006D2EA4}"
CONST IID_RecordsetEvents = "{00001266-0000-0010-8000-00AA006D2EA4}"
CONST IID_RecordsetEventsVt = "{00001403-0000-0010-8000-00AA006D2EA4}"
```
