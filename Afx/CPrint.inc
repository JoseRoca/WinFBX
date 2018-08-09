' ########################################################################################
' Microsoft Windows
' File: CPrint.inc
' Contents: Class for printing
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2018 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#INCLUDE ONCE "win/winspool.bi"
#INCLUDE ONCE "win/shellapi.bi"
#INCLUDE ONCE "win/commdlg.bi"
#INCLUDE ONCE "Afx/AfxGdiplus.inc"
#INCLUDE ONCE "Afx/CWSTR.inc"

NAMESPACE Afx

' ========================================================================================
' CPrint class
' ========================================================================================
TYPE CPrint

Public:
   m_hDC AS HDC
   m_wszPrinterName AS WSTRING * 260

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION ChoosePrinter (BYVAL hwndOwner AS HWND = NULL) AS BOOLEAN
   DECLARE FUNCTION AttachPrinter (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
   DECLARE FUNCTION PageSetup (BYVAL hwndOwner AS HWND = NULL) AS BOOLEAN
   DECLARE FUNCTION GetDC () AS .HDC
   DECLARE FUNCTION GetPrinterName () AS CWSTR
   DECLARE FUNCTION GetPPIX () AS LONG
   DECLARE FUNCTION GetPPIY () AS LONG
   DECLARE FUNCTION GetHorizontalResolution () AS LONG
   DECLARE FUNCTION GetVerticalResolution () AS LONG
   DECLARE FUNCTION GetDocumentProperties (BYVAL dmField AS DWORD) AS LONG
   DECLARE FUNCTION SetPrinterInfo (BYVAL dmField AS DWORD, BYVAL nValue AS LONG) AS BOOLEAN
   DECLARE PROPERTY Collate () AS BOOLEAN
   DECLARE PROPERTY Collate (BYVAL nMode AS LONG)
   DECLARE PROPERTY ColorMode () AS LONG
   DECLARE PROPERTY ColorMode (BYVAL nMode AS LONG)
   DECLARE PROPERTY Copies () AS LONG
   DECLARE PROPERTY Copies (BYVAL nCopies AS LONG)
   DECLARE FUNCTION Duplex () AS BOOLEAN
   DECLARE PROPERTY DuplexMode () AS LONG
   DECLARE PROPERTY DuplexMode (BYVAL nDuplexMode AS LONG)
   DECLARE PROPERTY Orientation () AS LONG
   DECLARE PROPERTY Orientation (BYVAL nOrientation AS LONG)
   DECLARE PROPERTY PaperSize () AS LONG
   DECLARE PROPERTY PaperSize (BYVAL nSize AS LONG)
   DECLARE FUNCTION GetPaperNames () AS CWSTR
   DECLARE PROPERTY Quality () AS LONG
   DECLARE PROPERTY Quality (BYVAL nMode AS LONG)
   DECLARE FUNCTION GetTrayNames () AS CWSTR
   DECLARE PROPERTY Scale () AS LONG
   DECLARE PROPERTY Scale (BYVAL nScale AS LONG)
   DECLARE PROPERTY Tray () AS LONG
   DECLARE PROPERTY Tray (BYVAL nTray AS LONG)
   DECLARE FUNCTION PrintBitmap (BYREF wszDocName AS WSTRING, BYVAL hbmp AS HBITMAP, BYVAL bStretch AS BOOLEAN = FALSE, _
                    BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' CPrint constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CPrint
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CPrint destructor
' ========================================================================================
PRIVATE DESTRUCTOR CPrint
   IF m_hDC then DeleteDC m_hDC
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Displays the printer dialog to select a printer.
' If the user clicks the OK button, the return value is true.
' If the user canceled or closed the Print dialog box or an error occurred, the return value is false.
' ========================================================================================
PRIVATE FUNCTION CPrint.ChoosePrinter (BYVAL hwndOwner AS HWND = NULL) AS BOOLEAN
   DIM pd AS PRINTDLGW
   pd.lStructSize = SIZEOF(pd)
   pd.hwndOwner = hwndOwner
   pd.flags = PD_RETURNDC OR PD_HIDEPRINTTOFILE OR PD_DISABLEPRINTTOFILE OR PD_NOSELECTION OR PD_NOPAGENUMS
   IF PrintDlgW(@pd) THEN
      IF pd.hDevMode THEN
         DIM pdm AS DEVMODEW PTR = GlobalLock(pd.hDevMode)
         IF pdm THEN
            m_wszPrinterName = pdm->dmDeviceName
            GlobalUnlock(pd.hDevMode)
         END IF
         GlobalFree(pd.hDevMode)
      END IF
      IF pd.hDevNames THEN GlobalFree(pd.hDevNames)
      IF m_hDC THEN DeleteDC m_hDC
      m_hDC = pd.hDC
      RETURN TRUE
   END IF
   RETURN FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches the specified printer to the class. Returns true of false.
' Example:
' DIM pPrint AS CPrint
' pPrint.AttachPrinter("OKI DATA CORP B410")
' ========================================================================================
PRIVATE FUNCTION CPrint.AttachPrinter (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
   ' // Create a device context for the printer
   DIM _hDC AS .HDC = CreateDCW(NULL, wszPrinterName, NULL, NULL)
   IF _hDC THEN
      m_wszPrinterName = wszPrinterName
      IF m_hDC THEN DeleteDC m_hDC
      m_hDC = _hDC
      RETURN TRUE
   END IF
   RETURN FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a Page Setup dialog box that enables the user to specify the attributes of a
' printed page. These attributes include the paper size and source, the page orientation
' (portrait or landscape), and the width of the page margins.
' ========================================================================================
PRIVATE FUNCTION CPrint.PageSetup (BYVAL hwndOwner AS HWND = NULL) AS BOOLEAN
   DIM psd AS PAGESETUPDLGW
   psd.lStructSize = SIZEOF(PAGESETUPDLGW)
   psd.hwndOwner = hwndOwner
   IF PageSetupDlgW(@psd) THEN RETURN TRUE
   RETURN FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the device context of the attached printer
' ========================================================================================
PRIVATE FUNCTION CPrint.GetDC () AS .HDC
   RETURN m_hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the name of the attached printer
' ========================================================================================
PRIVATE FUNCTION CPrint.GetPrinterName () AS CWSTR
   RETURN m_wszPrinterName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of pixels per inch of the specified host printer page (horizontal resolution).
' ========================================================================================
PRIVATE FUNCTION CPrint.GetPPIX () AS LONG
   IF m_hDC THEN RETURN GetDeviceCaps(m_hDC, LOGPIXELSX)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retuens the number of pixels per inch of the specified host printer page (vertical resolution).
' ========================================================================================
PRIVATE FUNCTION CPrint.GetPPIY () AS LONG
   IF m_hDC THEN RETURN GetDeviceCaps(m_hDC, LOGPIXELSY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Width, in pixels, of the printable area of the page.
' ========================================================================================
PRIVATE FUNCTION CPrint.GetHorizontalResolution () AS LONG
   IF m_hDC THEN RETURN GetDeviceCaps(m_hDC, HORZRES)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Width, in pixels, of the printable area of the page.
' ========================================================================================
PRIVATE FUNCTION CPrint.GetVerticalResolution () AS LONG
   IF m_hDC THEN RETURN GetDeviceCaps(m_hDC, VERTRES)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves printer initialization information.
' ========================================================================================
PRIVATE FUNCTION CPrint.GetDocumentProperties (BYVAL dmField AS DWORD) AS LONG
   ' // Start by opening the printer
   DIM hPrinter AS HANDLE
   IF OpenPrinterW(m_wszPrinterName, @hPrinter, NULL) = FALSE THEN RETURN 0
   ' // Allocate a buffer of the correct size
   DIM cbNeeded AS DWORD = DocumentPropertiesW(NULL, hPrinter, m_wszPrinterName, NULL, NULL, 0)
   DIM bufferDoc AS STRING = SPACE(cbNeeded)
   ' // Retrieve the printer configuration data
   DIM nRet AS LONG = DocumentPropertiesW(NULL, hPrinter, m_wszPrinterName, _
      cast(DEVMODEW PTR, STRPTR(bufferDoc)), NULL, DM_OUT_BUFFER)
   IF nRet = IDOK THEN
      ' // Cast it to a DEVMODEW structure
      DIM pDevMode AS DEVMODEW PTR = cast(DEVMODEW PTR, STRPTR(bufferDoc))
      ' // Finished with the printer
      ClosePrinter(hPrinter)
      ' // Return the requested value
      SELECT CASE dmField
         CASE DM_COLLATE       : RETURN pDevMode->dmCollate
         CASE DM_COPIES        : RETURN pDevMode->dmCopies
         CASE DM_ORIENTATION   : RETURN pDevMode->dmOrientation
         CASE DM_PAPERSIZE     : RETURN pDevMode->dmPaperSize
         CASE DM_PRINTQUALITY  : RETURN pDevMode->dmPrintQuality
         CASE DM_SCALE         : RETURN pDevMode->dmScale
         CASE DM_DEFAULTSOURCE : RETURN pDevMode->dmDefaultSource
         CASE DM_PAPERLENGTH   : RETURN pDevMode->dmPaperLength
         CASE DM_PAPERWIDTH    : RETURN pDevMode->dmPaperWidth
         CASE DM_DUPLEX        : RETURN pDevMode->dmDuplex
      END SELECT
   END IF
   RETURN nRet
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets data for a specified printer.
' - wszPrinterName : The name of the printer.
' - dmField : Field to change
' - nValue  : New value
' ========================================================================================
PRIVATE FUNCTION CPrint.SetPrinterInfo (BYVAL dmField AS DWORD, BYVAL nValue AS LONG) AS BOOLEAN

   ' // Start by opening the printer
   DIM hPrinter AS HANDLE
   DIM pd AS PRINTER_DEFAULTSW
   pd.DesiredAccess = PRINTER_ALL_ACCESS
   IF OpenPrinterW(m_wszPrinterName, @hPrinter, @pd) = FALSE THEN RETURN FALSE

   ' // The first GetPrinterW cll tells you how big the buffer should be in
   ' // order to hold all of PRINTER_INFO_2. Note that this should fail with
   ' // ERROR_INSUFFICIENT_BUFFER. If GetPrinterW fails for any other reason
   ' // or cbNeeded isn't set for some reason, then there is a problem...
   DIM dwLevel AS DWORD = 2, cbNeeded AS DWORD
   DIM nRet AS LONG = GetPrinterW(hPrinter, dwLevel, NULL, 0, @cbNeeded)
   IF nRet = 0 AND GetLastError <> ERROR_INSUFFICIENT_BUFFER THEN
      ClosePrinter(hPrinter)
      RETURN FALSE
   END IF

   ' // Allocate enough space for PRINTER_INFO_2...
   DIM bufferPrn AS STRING
   bufferPrn = SPACE(cbNeeded)
   ' // The second GetPrinterW fills in all the current settings, so all you
   ' // need to do is modify what you're interested in...
   nRet = GetPrinterW(hPrinter, dwLevel, CAST(BYTE PTR, STRPTR(bufferPrn)), cbNeeded, @cbNeeded)
   IF nRet = 0 THEN
      ClosePrinter(hPrinter)
      RETURN FALSE
   END IF

   ' // If GetPrinterW didn't fill in the DEVMODE, try to get it by calling DocumentProperties...
   DIM pi2 AS PRINTER_INFO_2W PTR
   pi2 = CAST(PRINTER_INFO_2W PTR, STRPTR(bufferPrn))
   DIM bufferDoc AS STRING
   DIM pDevMode AS DEVMODEW PTR
   IF pi2->pDevMode = NULL THEN
      ' // Allocate a buffer of the correct size
      cbNeeded = DocumentPropertiesW(NULL, hPrinter, m_wszPrinterName, NULL, NULL, 0)
      bufferDoc = SPACE(cbNeeded)
      ' // Retrieve the printer configuration data
      nRet = DocumentPropertiesW(NULL, hPrinter, m_wszPrinterName, _
         cast(DEVMODEW PTR, STRPTR(bufferDoc)), NULL, DM_OUT_BUFFER)
      IF nRet <> IDOK THEN
         ClosePrinter(hPrinter)
         RETURN FALSE
      END IF
      ' // Cast it to a DEVMODEW structure
      pDevMode = cast(DEVMODEW PTR, STRPTR(bufferDoc))
      pi2->pDevMode = pDevMode
   END IF

   ' // Specify exactly what we are attempting to change...
   SELECT CASE dmField
      CASE DM_COLLATE       : pi2->pDevMode->dmCollate = nValue
      CASE DM_COLOR         : pi2->pDevMode->dmColor = nValue
      CASE DM_COPIES        : pi2->pDevMode->dmCopies = nValue
      CASE DM_DUPLEX        : pi2->pDevMode->dmDuplex = nValue
      CASE DM_ORIENTATION   : pi2->pDevMode->dmOrientation = nValue
      CASE DM_PAPERSIZE     : pi2->pDevMode->dmPaperSize = nValue
      CASE DM_PRINTQUALITY  : pi2->pDevMode->dmPrintQuality = nValue
      CASE DM_DEFAULTSOURCE : pi2->pDevMode->dmDefaultSource = nValue
      CASE DM_SCALE         : pi2->pDevMode->dmScale = nValue
   END SELECT

   ' // Do not attempt to set security descriptor...
   pi2->pSecurityDescriptor = NULL

   ' // Make sure the driver-dependent part of devmode is updated...
   nRet = DocumentPropertiesW(NULL, hPrinter, m_wszPrinterName, _
             pi2->pDevMode, pi2->pDevMode, DM_IN_BUFFER OR DM_OUT_BUFFER)
   IF nRet <> IDOK THEN
      ClosePrinter(hPrinter)
      RETURN FALSE
   END IF

   ' // Update printer information...
   nRet = SetPrinterW(hPrinter, dwLevel, cast(BYTE PTR, pi2), 0)
   IF nRet = 0 THEN
      ' // The driver doesn't support, or it is unable to make the change...
      ClosePrinter(hPrinter)
      EXIT FUNCTION
   END IF

   ' // Finished with the printer
   ClosePrinter(hPrinter)

   IF nRet <> IDOK THEN RETURN FALSE
   RETURN TRUE

END FUNCTION
' ========================================================================================

' ========================================================================================
' If the printer supports collating, the return value is TRUE; otherwise, the return value is FALSE.
' If TRUE, the pages that are printed should be collated. To collate is to print out the
' entire document before printing the next copy, as opposed to printing out each page of
' the document the required number of times.
' ========================================================================================
PRIVATE PROPERTY CPrint.Collate () AS BOOLEAN
   RETURN DeviceCapabilitiesW(m_wszPrinterName, NULL, DC_COLLATE, NULL, NULL)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Specifies whether collation should be used when printing multiple copies.
' The following are the possible values:
'   DMCOLLATE_TRUE = 1
'   DMCOLLATE_FALSE = 0
' Remarks: Not all printers can collate.
' ========================================================================================
PRIVATE PROPERTY CPrint.Collate (BYVAL nMode AS LONG)
   this.SetPrinterInfo(DM_COLLATE, nMode)
END PROPERTY
' ========================================================================================

' ========================================================================================
' If the printer supports color printing, the return value is TRUE; otherwise, the return value is FALSE.
' Some color printers have the capability to print using true black instead of a combination
' of cyan, magenta, and yellow (CMY). This usually creates darker and sharper text for
' documents. This option is only useful for color printers that support true black printing.
' ========================================================================================
PRIVATE PROPERTY CPrint.ColorMode () AS LONG
   RETURN DeviceCapabilitiesW(m_wszPrinterName, NULL, DC_COLORDEVICE, NULL, NULL)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Switches between color and monochrome on color printers.
' The following are the possible values:
'   DMCOLOR_MONOCHROME = 1
'   DMCOLOR_COLOR = 2
' ========================================================================================
PRIVATE PROPERTY CPrint.ColorMode (BYVAL nMode AS LONG)
   this.SetPrinterInfo(DM_COLOR, nMode)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the number of copies to print if the device supports multiple-page copies.
' ========================================================================================
PRIVATE PROPERTY CPrint.Copies () AS LONG
   RETURN this.GetDocumentProperties(DM_COPIES)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Selects the number of copies to print if the device supports multiple-page copies.
' ========================================================================================
PRIVATE PROPERTY CPrint.Copies (BYVAL nCopies AS LONG)
   this.SetPrinterInfo(DM_COPIES, nCopies)
END PROPERTY
' ========================================================================================

' ========================================================================================
' If the printer supports duplex printing, the return value is TRUE; otherwise, the return value is FALSE.
' ========================================================================================
PRIVATE FUNCTION CPrint.Duplex () AS BOOLEAN
   RETURN DeviceCapabilitiesW(m_wszPrinterName, NULL, DC_DUPLEX, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' If the printer supports duplex printing, returns the current duplex mode
' DMDUP_SIMPLEX = 1 (Single sided printing)
' DMDUP_VERTICAL = 2 (Page flipped on the vertical edge)
' DMDUP_HORIZONTAL = 3 (Page flipped on the horizontal edge)
' ========================================================================================
PRIVATE PROPERTY CPrint.DuplexMode () AS LONG
'   RETURN DeviceCapabilitiesW(m_wszPrinterName, NULL, DC_DUPLEX, NULL, NULL)
   RETURN this.GetDocumentProperties(DM_DUPLEX)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the printer duplex mode
' DMDUP_SIMPLEX = 1 (Single sided printing)
' DMDUP_VERTICAL = 2 (Page flipped on the vertical edge)
' DMDUP_HORIZONTAL = 3 (Page flipped on the horizontal edge)
' ========================================================================================
PRIVATE PROPERTY CPrint.DuplexMode (BYVAL nDuplexMode AS LONG)
   this.SetPrinterInfo(DM_DUPLEX, nDuplexMode)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the printer orientation.
' The return value can be one of the following:
' - DMORIENT_PORTRAIT = Portrait
' - DMORIENT_LANDSCAPE = Landscape
' ========================================================================================
PRIVATE PROPERTY CPrint.Orientation () AS LONG
   RETURN this.GetDocumentProperties(DM_ORIENTATION)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the printer orientation.
' DMORIENT_PORTRAIT = Portrait
' DMORIENT_LANDSCAPE = Landscape
' ========================================================================================
PRIVATE PROPERTY CPrint.Orientation (BYVAL nOrientation AS LONG)
   this.SetPrinterInfo(DM_ORIENTATION, nOrientation)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the paper size for which the printer is currently configured.
' DMPAPER_LETTER = 1; DMPAPER_FIRST = DMPAPER_LETTER; DMPAPER_LETTERSMALL = 2; DMPAPER_TABLOID = 3;
' DMPAPER_LEDGER = 4; DMPAPER_LEGAL = 5; DMPAPER_STATEMENT = 6; DMPAPER_EXECUTIVE = 7; DMPAPER_A3 = 8;
' DMPAPER_A4 = 9; DMPAPER_A4SMALL = 10; DMPAPER_A5 = 11; DMPAPER_B4 = 12; DMPAPER_B5 = 13;
' DMPAPER_FOLIO = 14; DMPAPER_QUARTO = 15; DMPAPER_10X14 = 16; DMPAPER_11X17 = 17; DMPAPER_NOTE = 18;
' DMPAPER_ENV_9 = 19; DMPAPER_ENV_10 = 20; DMPAPER_ENV_11 = 21; DMPAPER_ENV_12 = 22; DMPAPER_ENV_14 = 23;
' DMPAPER_CSHEET = 24; DMPAPER_DSHEET = 25; DMPAPER_ESHEET = 26; DMPAPER_ENV_DL = 27; DMPAPER_ENV_C5 = 28;
' DMPAPER_ENV_C3 = 29; DMPAPER_ENV_C4 = 30; DMPAPER_ENV_C6 = 31; DMPAPER_ENV_C65 = 32; DMPAPER_ENV_B4 = 33;
' DMPAPER_ENV_B5 = 34; DMPAPER_ENV_B6 = 35; DMPAPER_ENV_ITALY = 36; DMPAPER_ENV_MONARCH = 37;
' DMPAPER_ENV_PERSONAL = 38; DMPAPER_FANFOLD_US = 39; DMPAPER_FANFOLD_STD_GERMAN = 40;
' DMPAPER_FANFOLD_LGL_GERMAN = 41; DMPAPER_ISO_B4 = 42; DMPAPER_JAPANESE_POSTCARD = 43; DMPAPER_9X11 = 44;
' DMPAPER_10X11 = 45; DMPAPER_15X11 = 46; DMPAPER_ENV_INVITE = 47; DMPAPER_RESERVED_48 = 48;
' DMPAPER_RESERVED_49 = 49; DMPAPER_LETTER_EXTRA = 50; DMPAPER_LEGAL_EXTRA = 51; DMPAPER_TABLOID_EXTRA = 52;
' DMPAPER_A4_EXTRA = 53; DMPAPER_LETTER_TRANSVERSE = 54; DMPAPER_A4_TRANSVERSE = 55;
' DMPAPER_LETTER_EXTRA_TRANSVERSE = 56; DMPAPER_A_PLUS = 57; DMPAPER_B_PLUS = 58; DMPAPER_LETTER_PLUS = 59;
' DMPAPER_A4_PLUS = 60; DMPAPER_A5_TRANSVERSE = 61; DMPAPER_B5_TRANSVERSE = 62; DMPAPER_A3_EXTRA = 63;
' DMPAPER_A5_EXTRA = 64; DMPAPER_B5_EXTRA = 65; DMPAPER_A2 = 66; DMPAPER_A3_TRANSVERSE = 67;
' DMPAPER_A3_EXTRA_TRANSVERSE = 68; DMPAPER_DBL_JAPANESE_POSTCARD = 69; DMPAPER_A6 = 70;
' DMPAPER_JENV_KAKU2 = 71; DMPAPER_JENV_KAKU3 = 72; DMPAPER_JENV_CHOU3 = 73; DMPAPER_JENV_CHOU4 = 74;
' DMPAPER_LETTER_ROTATED = 75; DMPAPER_A3_ROTATED = 76; DMPAPER_A4_ROTATED = 77; DMPAPER_A5_ROTATED = 78;
' DMPAPER_B4_JIS_ROTATED = 79; DMPAPER_B5_JIS_ROTATED = 80; DMPAPER_JAPANESE_POSTCARD_ROTATED = 81;
' DMPAPER_DBL_JAPANESE_POSTCARD_ROTATED = 82; DMPAPER_A6_ROTATED = 83; DMPAPER_JENV_KAKU2_ROTATED = 84;
' DMPAPER_JENV_KAKU3_ROTATED = 85; DMPAPER_JENV_CHOU3_ROTATED = 86; DMPAPER_JENV_CHOU4_ROTATED = 87;
' DMPAPER_B6_JIS = 88; DMPAPER_B6_JIS_ROTATED = 89; DMPAPER_12X11 = 90; DMPAPER_JENV_YOU4 = 91;
' DMPAPER_JENV_YOU4_ROTATED = 92; DMPAPER_P16K = 93; DMPAPER_P32K = 94; DMPAPER_P32KBIG = 95;
' DMPAPER_PENV_1 = 96; DMPAPER_PENV_2 = 97; DMPAPER_PENV_3 = 98; DMPAPER_PENV_4 = 99; DMPAPER_PENV_5 = 100;
' DMPAPER_PENV_6 = 101; DMPAPER_PENV_7 = 102; DMPAPER_PENV_8 = 103; DMPAPER_PENV_9 = 104;
' DMPAPER_PENV_10 = 105; DMPAPER_P16K_ROTATED = 106; DMPAPER_P32K_ROTATED = 107; DMPAPER_P32KBIG_ROTATED = 108;
' DMPAPER_PENV_1_ROTATED = 109; DMPAPER_PENV_2_ROTATED = 110; DMPAPER_PENV_3_ROTATED = 111;
' DMPAPER_PENV_4_ROTATED = 112; DMPAPER_PENV_5_ROTATED = 113; DMPAPER_PENV_6_ROTATED = 114;
' DMPAPER_PENV_7_ROTATED = 115; DMPAPER_PENV_8_ROTATED = 116; DMPAPER_PENV_9_ROTATED = 117;
' DMPAPER_PENV_10_ROTATED = 118; DMPAPER_LAST = DMPAPER_PENV_10_ROTATED; DMPAPER_USER = 256
' ========================================================================================
PRIVATE PROPERTY CPrint.PaperSize () AS LONG
   RETURN this.GetDocumentProperties(DM_PAPERSIZE)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the printer paper size.
' ========================================================================================
PRIVATE PROPERTY CPrint.PaperSize (BYVAL nSize AS LONG)
   this.SetPrinterInfo(DM_PAPERSIZE, nSize)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns a list of supported paper names (for example, Letter or Legal).
' The names are separated by a carriage return and a line feed characters.
' ========================================================================================
PRIVATE FUNCTION CPrint.GetPaperNames () AS CWSTR
   DIM wstrNames AS CWSTR, wszNames(ANY) AS WSTRING * 64
   DIM r AS LONG = DeviceCapabilitiesW(m_wszPrinterName, NULL, DC_PAPERNAMES, NULL, NULL)
   IF r = -1 THEN RETURN ""
   REDIM wszNames(r - 1) AS WSTRING * 64
   r = DeviceCapabilitiesW(m_wszPrinterName, NULL, DC_PAPERNAMES, @wszNames(0), NULL)
   IF r < 1 THEN RETURN ""
   DIM i AS LONG
   FOR i = 0 TO r - 1
      wstrNames += wszNames(i)
      IF i < r - 1 THEN wstrNames += CHR(13, 10)
   NEXT
   RETURN wstrNames
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the printer print quality mode.
' There are four predefined device-independent values:
' - DMRES_DRAFT  = Draft (-1)
' - DMRES_LOW    = Low (-2)
' - DMRES_MEDIUM = Medium (-3)
' - DMRES_HIGH   = High (-4)
' If a positive value is returned, it specifies the number of pixels per inch (PPI) and is
' therefore device dependent.
' ========================================================================================
PRIVATE PROPERTY CPrint.Quality () AS LONG
   RETURN this.GetDocumentProperties(DM_PRINTQUALITY)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Specifies the printer print quality mode.
' There are four predefined device-independent values:
' - DMRES_DRAFT  = Draft (-1)
' - DMRES_LOW    = Low (-2)
' - DMRES_MEDIUM = Medium (-3)
' - DMRES_HIGH   = High (-4)
' If a positive value is specified, it represents the number of pixels per inch (PPI) for
' the x resolution.
' ========================================================================================
PRIVATE PROPERTY CPrint.Quality (BYVAL nMode AS LONG)
   this.SetPrinterInfo(DM_PRINTQUALITY, nMode)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the factor by which the printed output is to be scaled. The apparent page size
' is scaled from the physical page size by a factor of dmScale /100. For example, a
' letter-sized page with a dmScale value of 50 would contain as much data as a page of
' 17- by 22-inches because the output text and graphics would be half their original
' height and width.
' ========================================================================================
PRIVATE PROPERTY CPrint.Scale () AS LONG
   RETURN this.GetDocumentProperties(DM_SCALE)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Specifies the factor by which the printed output is to be scaled. The apparent page size
' is scaled from the physical page size by a factor of dmScale /100. For example, a
' letter-sized page with a dmScale value of 50 would contain as much data as a page of
' 17- by 22-inches because the output text and graphics would be half their original
' height and width.
' ========================================================================================
PRIVATE PROPERTY CPrint.Scale (BYVAL nScale AS LONG)
   this.SetPrinterInfo(DM_SCALE, nScale)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the paper source.
' Returns one of the following values: DMBIN_UPPER = 1; DMBIN_LOWER = 2; DMBIN_MIDDLE = 3;
' DMBIN_MANUAL = 4; DMBIN_ENVELOPE = 5; DMBIN_ENVMANUAL = 6; DMBIN_AUTO = 7;
' DMBIN_TRACTOR = 8; DMBIN_SMALLFMT = 9; DMBIN_LARGEFMT = 10; DMBIN_LARGECAPACITY = 11;
' DMBIN_CASSETTE = 14; DMBIN_FORMSOURCE = 15
' ========================================================================================
PRIVATE PROPERTY CPrint.Tray () AS LONG
   RETURN this.GetDocumentProperties(DM_DEFAULTSOURCE)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the paper source. Can be one of the following values, or it can be a device-specific
' value greater than or equal to DMBIN_USER.
' DMBIN_UPPER = 1; DMBIN_LOWER = 2; DMBIN_MIDDLE = 3; DMBIN_MANUAL = 4; DMBIN_ENVELOPE = 5;
' DMBIN_ENVMANUAL = 6; DMBIN_AUTO = 7;  DMBIN_TRACTOR = 8; DMBIN_SMALLFMT = 9;
' DMBIN_LARGEFMT = 10; DMBIN_LARGECAPACITY = 11; DMBIN_CASSETTE = 14; DMBIN_FORMSOURCE = 15
' ========================================================================================
PRIVATE PROPERTY CPrint.Tray (BYVAL nTray AS LONG)
   this.SetPrinterInfo(DM_DEFAULTSOURCE, nTray)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Prints a Windows bitmap in the selected printer.
' Parameters:
' - wszDocName = Document name.
' - hbmp = Handle to the bitmap.
' - bStretch = Stretch the image.
' - nStretchMode = Stretching mode. Default value = InterpolationModeHighQualityBicubic.
'   InterpolationModeLowQuality = 1
'   InterpolationModeHighQuality = 2
'   InterpolationModeBilinear = 3
'   InterpolationModeBicubic = 4
'   InterpolationModeNearestNeighbor = 5
'   InterpolationModeHighQualityBilinear = 6
'   InterpolationModeHighQualityBicubic = 7
' Return value: Returns TRUE if the bitmap has been printed successfully, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CPrint.PrintBitmap ( _
   BYREF wszDocName AS WSTRING, _
   BYVAL hbmp AS HBITMAP, _
   BYVAL bStretch AS BOOLEAN = FALSE, _
   BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic _
   ) AS BOOLEAN

   IF m_hDC = NULL THEN RETURN FALSE
   DIM _wszDocName AS WSTRING * 260 = wszDocName

   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = NULL THEN EXIT FUNCTION

   DIM pBitmap AS GpBitmap PTR

   ' // Note: Using a DO... LOOP, with an EXIT DO before LOOP, to avoid the use of GOTO.
   DO
      DIM pGraphics AS GpGraphics PTR
      GdipCreateFromHDC(m_hDC, @pGraphics)
      IF pGraphics = NULL THEN EXIT DO
      ' // Create a Bitmap object from an HBITMAP
      GdipCreateBitmapFromHBITMAP(hbmp, NULL, @pBitmap)
      IF pBitmap = NULL THEN EXIT DO
      ' // Stretching
      DIM cx AS SINGLE, cy AS SINGLE
      IF bStretch THEN GdipSetInterpolationMode(pGraphics, nStretchMode)
      ' // Get the DPIs of the printer
      DIM dpiX AS SINGLE, dpiY AS SINGLE
      GdipGetDpiX(pGraphics, @dpiX)
      GdipGetDpiY(pGraphics, @dpiY)
      ' // Calculate the width and height according to the DPIs of the printer
      cx = GetDeviceCaps(m_hDC, HORZRES) / (dpiX / 100)
      cy = GetDeviceCaps(m_hDC, VERTRES) / (dpiY / 100)
      ' // Print the bitmap
      DIM di AS DOCINFOW
      di.cbSize = SIZEOF(DOCINFOW)
      di.lpszDocName = @_wszDocName
      DIM hr AS LONG = StartDocW(m_hDC, @di)
      IF hr <= 0 THEN EXIT DO
      DIM nCopies AS LONG = this.Copies
      IF nCopies < 1 THEN nCopies = 1
      FOR i AS LONG = 1 TO nCopies
         IF StartPage(m_hDC) THEN
            ' // Draw the image
            IF bStretch THEN
               GdipDrawImageRect(pGraphics, CAST(GpImage PTR, pBitmap), 0, 0, cx, cy)
            ELSE
               GdipDrawImage(pGraphics, CAST(GpImage PTR, pBitmap), 0, 0)
            END IF
            EndPage(m_hDC)
         END IF
      NEXT
      EndDoc(m_hDC)
      FUNCTION = TRUE
      EXIT DO
   LOOP

   IF pBitmap THEN GdipDisposeImage(cast(GpImage PTR, pBitmap))
   ' // Shutdown Gdiplus
   GdiplusShutdown token

END FUNCTION
' ========================================================================================

END NAMESPACE
