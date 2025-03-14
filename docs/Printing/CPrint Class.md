# CPrint Class

Allows to attach/choose a printer and get/set its setting values.

**Include file**: CPrint.inc

### Methods

| Name       | Description |
| ---------- | ----------- |
| [AttachPrinter](#attachprinter) | Attaches the specified printer to the class. |
| [ChoosePrinter](#chooseprinter) | Displays the printer dialog to select a printer. |
| [PageSetup](#pagesetup) | Displays a Page Setup dialog box that enables the user to specify the attributes of a printed page. |
| [GetPrinterName](#getprintername) | Returns the name of the attached printer. |
| [SetDefaultPrinter](#setdefaultprinter) | Sets the default printer. |
| [GetDefaultPrinter](#getdefaultprinter) | Returns the name of the default printer. |
| [GetDefaultPrinterDriver](#getdefaultprinterdriver) | Returns the name of the default printer driver. |
| [GetDefaultPrinterPort](#getdefaultprinterport) | Returns the name of the default printer port. |
| [EnumPrinterNames](#EnumPrinterNames) | Returns a list with the available printers, print servers, domains, or print providers. |
| [GetDC](#getdc) | Returns the handle of the device context of the attached printer. |
| [GetPPIX](#getppix) | Returns the number of pixels per inch of the page (horizontal resolution). |
| [GetPPIY](#getppiy) | Returns the number of pixels per inch of the page (vertical resolution). |
| [GetHorizontalUnits](#gethorizontalunits) | Returns the width, in world units, of the printable area of the page. |
| [GetVerticalUnits](#Getverticalunits) | Returns the height, in world units, of the printable area of the page. |
| [GetHorizontalResolution](#Gethorizontalresolution) | Returns the width, in pixels, of the printable area of the page. |
| [GetVerticalResolution](#Getverticalresolution) | Returns the height, in pixels, of the printable area of the page. |
| [GetMarginsPixels](#getmarginspixels) | Retrieves the margins (the non-printable area) of the printer page, in pixels. |
| [GetMarginsUnits](#getmarginsunits) | Retrieves the margins (the non-printable area) of the printer page, in world units. |
| [PixelsToUnitsX](#pixelstounitsx) | Converts pixels to world units. (horizontal resolution). |
| [PixelsToUnitsY](#pixelstounitsy) | Converts pixels to world units. (vertical resolution). |
| [UnitsToPixelsX](#unitstopixelsx) | Converts world units to pixels (horizontal resolution). |
| [UnitsToPixelsY](#unitstopixelsy) | Converts world units to pixels (vertical resolution). |
| [PixelsToPointsX](#pixelstopointsx) | Converts pixels to point size (1/72 of an inch) (horizontal resolution). |
| [PixelsToPointsY](#pixelstopointsy) | Converts pixels to point size (1/72 of an inch) (vertical resolution). |
| [PointsToPixelsX](#pointstopixelsx) | Converts a point size (1/72 of an inch) to pixels (horizontal resolution). |
| [PointsToPixelsY](#pointstopixelsy) | Converts a point size (1/72 of an inch) to pixels (vertical resolution). |
| [GetPaperNames](#getpapernames) | Returns a list of supported paper names. |
| [PrintBitmap](#printbitmap) | Prints a Windows bitmap in the attached printer. |
| [PrintBitmapToFile](#printbitmaptofile) | Prints a Windows bitmap in the specified file. |

### Properties

| Name       | Description |
| ---------- | ----------- |
| [Collate](#collate) | Returns true if the printer supports collating mode; False, otherwise. |
| [CollateMode](#collatemode) | Gets/sets the collate setting value. |
| [Color](#color) | Returns true if the printer supports color mode; False, otherwise. |
| [ColorMode](#colormode) | Switches between color and monochrome on color printers. |
| [Copies](#copies) | Gets/sets the number of copies to print if the device supports multiple-page copies. |
| [Duplex](#duplex) | Checks if the printer supports duplex printing. |
| [DuplexMode](#duplexmode)  | Gets/sets the current duplex mode. |
| [Orientation](#orientation) | Gets/sets the printer orientation. |
| [PaperLength](#paperlength) | Gets/sets the paper length in units of 1/10 of a millimeter. |
| [PaperSize](#papersize) | Gets/sets the printer paper size. |
| [PaperWidth](#paperwidth) | Gets/sets the paper width in units of 1/10 of a millimeter. |
| [Quality](#quality) | Gets/sets the printer print quality mode. |
| [Scale](#scale) | Gets/sets the factor by which the printed output is to be scaled. |
| [SetPaperSize](#setpapersize) | Sets the paper length and width in units of 1/10 of a millimeter. |
| [Tray](#tray) | Specifies the paper source. |

---

# <a name="attachprinter"></a>AttachPrinter

Creates a device context (DC) for the specified printer and attaches it to the class.

```
FUNCTION AttachPrinter (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The name of the printer to attach (as shown in the Devices and Printers applet of the Control Panel). |

#### Return value

BOOLEAN. True or false.

#### Example

The following console example creates an instance of the CPrint class to attach the "Microsoft Print to PDF" driver, retrieves a handle to its device context and uses the WinFBX GDI+ classes to draw the contents of the page.

```
'#CONSOLE ON
#define UNICODE
#INCLUDE ONCE "Afx/CPrint.inc"
#INCLUDE ONCE "Afx/CGdiplus/CGdiplus.inc"
USING Afx

DIM pPrint AS CPrint
pPrint.AttachPrinter("Microsoft Print to PDF")
DIM hdcPrint AS HDC = pPrint.GetDC
DIM docInfo AS DOCINFOW
docInfo.cbSize = SIZEOF(DOCINFOW)
DIM wszDocName AS WSTRING * 260 = "GdiplusPrint" 
docInfo.lpszDocName = VARPTR(wszDocName)
StartDocW(hdcPrint, @docInfo)
StartPage(hdcPrint)
SCOPE
   DIM graphics AS CGpGraphics = hdcPrint
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   graphics.DrawLine(@pen, 50, 50, 350, 550)
   graphics.DrawRectangle(@pen, 50, 50, 300, 500)
   graphics.DrawEllipse(@pen, 50, 50, 300, 500)
END SCOPE
EndPage(hdcPrint)
EndDoc(hdcPrint)

PRINT "Press any key to end..."
SLEEP
```

---

# <a name="chooseprinter"></a>ChoosePrinter

Displays a Print Dialog Box to select a printer.

```
FUNCTION ChoosePrinter (BYVAL hwndOwner AS HWND = NULL) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwndOwner* | Optional. A handle to the window that owns the dialog box. This member can be any valid window handle, or it can be NULL if the dialog box has no owner. |

#### Return value

BOOLEAN. If the user clicks the OK button, the return value is true. If the user canceled or closed the Print dialog box or an error occurred, the return value is false.

#### Example

```
DIM pPrint AS CPrint
pPrint.ChoosePrinter
```

---

# <a name="pagesetup"></a>PageSetup

Displays a Page Setup dialog box that enables the user to specify the attributes of a printed page. These attributes include the paper size and source, the page orientation (portrait or landscape), and the width of the page margins.

```
FUNCTION PageSetup (BYVAL hwndOwner AS HWND = NULL) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwndOwner* | Optional. A handle to the window that owns the dialog box. This member can be any valid window handle, or it can be NULL if the dialog box has no owner. |

#### Return value

BOOLEAN. If the user clicks the OK button, the return value is true. If the user canceled or closed the Page Setup dialog box or an error occurred, the return value is false.

#### Example

```
DIM pPrint AS CPrint
pPrint.PageSetup
```

---

# <a name="getprintername"></a>GetPrinterName

Returns the name of the attached printer.

```
FUNCTION GetPrinterName () AS CWSTR
```

#### Return value

CWSTR. The name of the attached printer. If there is not a printer attached, it returns an empty string.

---

# <a name="setdefaultprinter"></a>SetDefaultPrinter

Sets the printer name of the default printer for the current user on the local computer.

```
FUNCTION SetDefaultPrinter (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | A pointer to a null-terminated string containing the default printer name. For a remote printer connection, the name format is \**\\server\\**printername. For a local printer, the name format is printername. If this parameter is **NULL** or an empty string, that is, "", **SetDefaultPrinter** will select a default printer from one of the installed printers. If a default printer already exists, calling **SetDefaultPrinter** with a **NULL** or an empty string in this parameter might change the default printer. |

#### Return value

BOOLEAN. True or false.

#### Remarks

When using this method, you must specify a valid printer, driver, and port. If they are invalid, the APIs do not fail but the result is not defined. This could cause other programs to set the printer back to the previous valid printer. You can use **EnumPrinters** to retrieve the printer name, driver name, and port name of all available printers.

 #### Note

This is a blocking or synchronous function and might not return immediately. How quickly this function returns depends on run-time factors such as network status, print server configuration, and printer driver implementation factors that are difficult to predict when writing an application. Calling this function from a thread that manages interaction with the user interface could make the application appear to be unresponsive.

# <a name="getdefaultprinter"></a>GetDefaultPrinter

Returns the name of the default printer.

```
FUNCTION GetDefaultPrinter () AS CWSTR
```

# <a name="getdefaultprinterdriver"></a>GetDefaultPrinterDriver

Returns the name of the default printer driver.

```
FUNCTION GetDefaultPrinterDriver () AS CWSTR
```

---

# <a name="getdefaultprinterport"></a>GetDefaultPrinterPort

Returns the name of the default printer port.

```
FUNCTION GetDefaultPrinterPort () AS CWSTR
```

---

# <a name="enumprinternames"></a>EnumPrinterNames

Returns a list with the available printers, print servers, domains, or print providers.
Names are separated with a carriage return and a line feed characters.

```
FUNCTION EnumPrinterNames () AS CWSTR
```

---

# <a name="getdc"></a>GetDC

Returns the handle of the attached printer.

```
FUNCTION GetDC () AS HDC
```

#### Return value

HDC. The handle of the device context of the attached printer. If there is not a printer attached, it returns NULL.

---

# <a name="getppix"></a>GetPPIX

Returns the number of pixels per inch of the specified host printer page (horizontal resolution).

```
FUNCTION GetPPIX () AS LONG
```

#### Return value

LONG. The number of pixels per inch of the specified host printer page. If there is not a printer attached, it returns 0.

---

# <a name="getppiy"></a>GetPPIY

Returns the number of pixels per inch of the specified host printer page (vertical resolution).

```
FUNCTION GetPPIY () AS LONG
```

#### Return value

LONG. The number of pixels per inch of the specified host printer page. If there is not a printer attached, it returns 0.

---

# <a name="gethorizontalunits"></a>GetHorizontalUnits

Returns the width, in world units, of the printable area of the page.

```
FUNCTION GetHorizontalUnits () AS LONG
```

#### Return value

LONG. The width, in world units, of the printable area of the page. If there is not a printer attached, it returns 0.

---

# <a name="getverticalunits"></a>GetVerticalUnits

Returns the height, in world units, of the printable area of the page.

```
FUNCTION GetVerticalUnits () AS LONG
```

#### Return value

LONG. The height, in world units, of the printable area of the page. If there is not a printer attached, it returns 0.

---

# <a name="gethorizontalresolution"></a>GetHorizontalResolution

Returns the width, in pixels, of the printable area of the page.

```
FUNCTION GetHorizontalResolution () AS LONG
```

---

# <a name="getverticalresolution"></a>GetVerticalResolution

Returns the height, in pixels, of the printable area of the page.

```
FUNCTION GetVerticalResolution () AS LONG
```

---

# <a name="getmarginspixels"></a>GetMarginsPixels

Retrieves the margins (the non-printable area) of the printer page, in pixels.

```
SUB GetMarginsPixels (BYREF nLeft AS LONG, BYREF nTop AS LONG, BYREF nRight AS LONG, BYREF nBottom AS LONG)
FUNCTION GetMarginsPixels () AS RECT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nLeft* | Left margin. |
| *nTop* | Top margin. |
| *nRight* | Right margin. |
| *nBottom* | Bottom margin. |

The overloaded function returns a **RECT** structure.

#### Return value

LONG. The height, in pixels, of the printable area of the page. If there is not a printer attached, it returns 0.

---

# <a name="getmarginsunits"></a>GetMarginsUnits

Retrieves the margins (the non-printable area) of the printer page, in world units.

```
SUB GetMarginsUnits (BYREF nLeft AS LONG, BYREF nTop AS LONG, BYREF nRight AS LONG, BYREF nBottom AS LONG)
FUNCTION GetMarginsUnits () AS RECT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nLeft* | Left margin. |
| *nTop* | Top margin. |
| *nRight* | Right margin. |
| *nBottom* | Bottom margin. |

The overloaded function returns a **RECT** structure.

---

# <a name="pixelstounitsx"></a>PixelsToUnitsX

Converts pixels to world units (horizontal resolution).

```
FUNCTION PixelsToUnitsX (BYVAL pix AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pix* | The number of pixels to convert. |

# <a name="pixelstounitsy"></a>PixelsToUnitsY

Converts pixels to world units (vertical resolution).

```
FUNCTION PixelsToUnitsY (BYVAL pix AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pix* | The number of pixels to convert. |

---

# <a name="unitstopixelsx"></a>UnitsToPixelsX

Converts world units to pixels (horizontal resolution).

```
FUNCTION UnitsToPixelsX (BYVAL units AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *units* | The number of world units to convert. |

---

# <a name="unitstopixelsy"></a>UnitsToPixelsY

Converts world units to pixels (vertical resolution).

```
FUNCTION UnitsToPixelsY (BYVAL units AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *units* | The number of world units to convert. |

---

# <a name="pixelstopointsx"></a>PixelsToPointsX

Converts pixels to point size (1/72 of an inch) according to the PPI of the printer (horizontal resolution).

```
FUNCTION PixelsToPointsX (BYVAL pix AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pix* | The number of pixels to convert. |

#### Return value

LONG. The number of points. If there is not a printer attached, it returns 0.

---

# <a name="pixelstopointsy"></a>PixelsToPointsY

Converts pixels to point size (1/72 of an inch) according to the PPI of the printer (vertical resolution).

```
FUNCTION PixelsToPointsY (BYVAL pix AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pix* | The number of pixels to convert. |

#### Return value

LONG. The number of points. If there is not a printer attached, it returns 0.

---

# <a name="pointstopixelsx"></a>PointsToPixelsX

Converts a point size (1/72 of an inch) to pixels according to the PPI of the printer (horizontal resolution).

```
FUNCTION PointsToPixelsX (BYVAL pts AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | The number of points to convert. |

#### Return value

LONG. The number of pixels. If there is not a printer attached, it returns 0.

---

# <a name="pointstopixelsy"></a>PointsToPixelsY

Converts a point size (1/72 of an inch) to pixels according to the PPI of the printer (vertical resolution).

```
FUNCTION PointsToPixelsY (BYVAL pts AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | The number of points to convert. |

#### Return value

LONG. The number of pixels. If there is not a printer attached, it returns 0.

---

# <a name="getpapernames"></a>GetPaperNames

Returns a list of supported paper names (for example, Letter or Legal).

```
FUNCTION GetPaperNames () AS CWSTR
```

#### Return value

CWSTR. A list of supported paper names on success, or an empty string on failure. The names are separated by a carriage return and a line feed characters.

---

# <a name="printbitmap"></a>PrintBitmap

Prints a Windows bitmap to the attached printer.

```
FUNCTION PrintBitmap ( _
   BYREF wszDocName AS WSTRING, _
   BYVAL hbmp AS HBITMAP, _
   BYVAL bStretch AS BOOLEAN = FALSE, _
   BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic _
) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszDocName* | WSTRING. The document name. |
| *hbmp* | HBITMAP. Handle to the bitmap. |
| *bStretch* | BOOLEAN. Optional. True to strech the image or false. Defaults to False. |
| *nStretchMode* | LONG. Optional. Stretching mode. Defaults to InterpolationModeHighQualityBicubic.<br>Predefined constants:<br>*InterpolationModeLowQuality* = Specifies a low-quality mode.<br>*InterpolationModeHighQuality* = Specifies a high-quality mode.<br>*InterpolationModeBilinear* = Specifies bilinear interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 50 percent of its original size.<br>*InterpolationModeBicubic* = Specifies bicubic interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 25 percent of its original size.<br>*InterpolationModeNearestNeighbor* = Specifies nearest-neighbor interpolation.<br>*InterpolationModeHighQualityBilinear* = Specifies high-quality, bilinear interpolation. Prefiltering is performed to ensure high-quality shrinking.<br>*InterpolationModeHighQualityBicubic* = Specifies high-quality, bicubic interpolation. Prefiltering is performed to ensure high-quality shrinking. This mode produces the highest quality transformed images. |

#### Return value

BOOLEAN. Returns TRUE if the bitmap has been printed successfully, or FALSE otherwise.

---

# <a name="printbitmaptofile"></a>PrintBitmapToFile

Prints a Windows bitmap to the specified output file.

```
FUNCTION PrintBitmapToFile ( _
   BYREF wszDocName AS WSTRING, _
   BYREF wszOutputFileName AS WSTRING, _
   BYVAL hbmp AS HBITMAP, _
   BYVAL bStretch AS BOOLEAN = FALSE, _
   BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic _
) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszDocName* | WSTRING. The document name. |
| *wszOutputFileName* | WSTRING. The output file name. |
| *hbmp* | HBITMAP. Handle to the bitmap. |
| *bStretch* | BOOLEAN. Optional. True to strech the image or false. Defaults to False. |
| *nStretchMode* | LONG. Optional. Stretching mode. Defaults to InterpolationModeHighQualityBicubic.<br>Predefined constants:<br>*InterpolationModeLowQuality* = Specifies a low-quality mode.<br>*InterpolationModeHighQuality* = Specifies a high-quality mode.<br>*InterpolationModeBilinear* = Specifies bilinear interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 50 percent of its original size.<br>*InterpolationModeBicubic* = Specifies bicubic interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 25 percent of its original size.<br>*InterpolationModeNearestNeighbor* = Specifies nearest-neighbor interpolation.<br>*InterpolationModeHighQualityBilinear* = Specifies high-quality, bilinear interpolation. Prefiltering is performed to ensure high-quality shrinking.<br>*InterpolationModeHighQualityBicubic* = Specifies high-quality, bicubic interpolation. Prefiltering is performed to ensure high-quality shrinking. This mode produces the highest quality transformed images. |

#### Return value

BOOLEAN. Returns TRUE if the bitmap has been printed successfully, or FALSE otherwise.

---

# <a name="collate"></a>Collate

If the printer supports collating, the return value is TRUE; otherwise, the return value is FALSE.

```
PROPERTY Collate () AS BOOLEAN
```

---

# <a name="collatemode"></a>CollateMode

Gets/sets the printer collating mode.

```
PROPERTY CollateMode () AS LONG
PROPERTY CollateMode (BYVAL nMode AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMode* | LONG. The collating mode. Possible values: DMCOLLATE_TRUE, DMCOLLATE_FALSE. |

#### Return value

BOOLEAN. If the printer supports collating, the return value is TRUE; otherwise, the return value is FALSE. If TRUE, the pages that are printed should be collated. To collate is to print out the entire document before printing the next copy, as opposed to printing out each page of the document the required number of times.

---

# <a name="color"></a>Color

If the printer supports color mode, the return value is TRUE; otherwise, the return value is FALSE.

```
PROPERTY Color () AS BOOLEAN
```

---

# <a name="colormode"></a>ColorMode

Switches between color and monochrome on color printers. Some color printers have the capability to print using true black instead of a combination of cyan, magenta, and yellow (CMY). This usually creates darker and sharper text for documents. This option is only useful for color printers that support true black printing.

```
PROPERTY ColorMode () AS LONG
PROPERTY ColorMode (BYVAL nMode AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMode* | LONG. The color mode. Possible values: DMCOLOR_MONOCHROME, DMCOLOR_COLOR. |

#### Return value

LONG. Returns the printer color mode: DMCOLOR_MONOCHROME or DMCOLOR_COLOR.

---

# <a name="copies"></a>Copies

Gets/sets the number of copies to print if the device supports multiple-page copies.

```
PROPERTY Copies () AS BOOLEAN
PROPERTY Copies (BYVAL nCopies AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCopies* | LONG. The number of copies to print if the device supports multiple-page copies. |

#### Return value

LONG. The number of copies.

---

# <a name="duplex"></a>Duplex

Checks if the printer supports duplex printing.

```
PROPERTY Duplex () AS BOOLEAN
```

#### Return value

BOOLEAN. If the printer supports duplex printing, the return value is TRUE; otherwise, the return value is FALSE.

---

# <a name="duplexmode"></a>DuplexMode

Gets/sets the current duplex mode.

```
PROPERTY DuplexMode () AS LONG
PROPERTY DuplexMode (BYVAL nMode AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMode* | LONG. The duplex mode.<br>*DMDUP_SIMPLEX* = Single sided printing.<br>*DMDUP_VERTICAL* = Page flipped on the vertical edge.<br>*DMDUP_HORIZONTAL* = Page flipped on the horizontal edge. |

#### Return value

LONG. If the printer supports duplex printing, returns the current duplex mode

---

# <a name="orientation"></a>Orientation

Gets/sets the printer orientation.

```
PROPERTY Orientation () AS LONG
PROPERTY Orientation (BYVAL nOrientation AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nOrientation* | LONG. The printer orientation.<br>*DMORIENT_PORTRAIT* = Portrait.<br>*DMORIENT_LANDSCAPE* = Landscape. |

#### Return value

LONG. The printer orientation.

---

# <a name="paperlength"></a>PaperLength

Gets/sets the printer paper length in units of 1/10 of a millimeter. This value overrides the length of the paper specified by the **PaperSize** property, and is used if the paper is of a custom size, or if the device is a dot matrix printer, which can print a page of arbitrary length. If you use this property, you must also use the **PaperWidth** property to set the print paper width. Alternatively, you can use the **SetPageSize** function to set both the length and width with only one call.

```
PROPERTY PaperLength () AS LONG
PROPERTY PaperLength (BYVAL nLength AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nLength* | LONG. The printer paper length. |

#### Return value

LONG. The printer paper length.

---

# <a name="papersize"></a>PaperSize

Gets/sets the printer paper size.

```
PROPERTY PaperSize () AS LONG
PROPERTY PaperSize (BYVAL nSize AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nSize* | LONG. Specifies the printer paper size, with DMPAPER_LETTER, DMPAPER_LEGAL, DMPAPER_A3, and DMPAPER_A4 being the most typical. Note that the paper size types cannot be combined with one another.<br>For a list of paper sizes see [Paper Sizes](https://docs.microsoft.com/en-us/windows/desktop/intl/paper-sizes). |

#### Return value

LONG. The printer paper size.

---

# <a name="setpapersize"></a>SetPaperSize

Sets the printer paper length and width.

```
FUNCTION SetPaperSize (BYVAL nLength AS LONG, BYVAL nWidth AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *nLength* | LONG. The length of the printer paper size. |
| *nWidth* | LONG. The width of the printer paper size. |

#### Return value

BOOLEAN. Returns True on success or False on failure.

---

# <a name="paperwidth"></a>PaperWidth

Gets/sets the printer paper width in units of 1/10 of a millimeter. This value overrides the length of the paper specified by the *PaperSize* property, and is used if the paper is of a custom size, or if the device is a dot matrix printer, which can print a page of arbitrary length. It must be used if the **PaperLength** is also used. Alternatively, you can use the **SetPageSize** function to set both the length and width with only one call.

```
PROPERTY PaperWidth () AS LONG
PROPERTY PaperWidth (BYVAL nWidth AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | LONG. The printer paper length. |

#### Return value

LONG. The printer paper width.

---

# <a name="quality"></a>Quality

Gets/sets the printer print quality mode.

```
PROPERTY Quality () AS LONG
PROPERTY Quality (BYVAL nMode AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMode* | LONG. The printer print quality mode.<br>There are four predefined device-independent values:<br>*DMRES_DRAFT* (-1) = Draft.<br>*DMRES_LOW* (-2) = Low<br>*DMRES_MEDIUM* (-3) = Medium<br>*DMRES_HIGH* (-4) = High<br>If a positive value is specified, it represents the number of pixels per inch (PPI) for the x resolution. |

#### Return value

LONG. The printer print quality mode.

---

# <a name="scale"></a>Scale

Gets/sets the factor by which the printed output is to be scaled.

```
PROPERTY Scale () AS LONG
PROPERTY Scale (BYVAL nScale AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nScale* | LONG. The factor by which the printed output is to be scaled. The apparent page size is scaled from the physical page size by a factor of nScale / 100. For example, a letter-sized page with a nScale value of 50 would contain as much data as a page of 17- by 22-inches because the output text and graphics would be half their original height and width. |

#### Return value

LONG. The printer scaling factor.

---

# <a name="tray"></a>Tray

Gets/sets the paper source.

```
PROPERTY Tray () AS LONG
PROPERTY Tray (BYVAL nTray AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTray* | LONG. The paper source. Can be one of the following values, or it can be a device-specific value greater than or equal to DMBIN_USER (256).<br>*DMBIN_UPPER* = Select the upper paper bin. This value is also used for the paper source for printers that only have one paper bin.<br>*DMBIN_LOWER* = Select the lower bin.<br>*DMBIN_MIDDLE* = Select the middle paper bin.<br>*DMBIN_MANUAL* = Manually select the paper bin.<br>*DMBIN_ENVELOPE* = Select the auto envelope bin.<br>*DMBIN_ENVMANUAL* = Select the manual envelope bin.<br>*DMBIN_AUTO* = Auto-select the bin.<br>*DMBIN_TRACTOR* = Select the bin with the tractor paper.<br>*DMBIN_SMALLFMT* = Select the bin with the smaller paper format.<br>*DMBIN_LARGEFMT* = Select the bin with the larger paper format.<br>*DMBIN_LARGECAPACITY* = Select the bin with large capacity.<br>*DMBIN_CASSETTE* = Select the cassette bin.<br>*DMBIN_FORMSOURCE* = Select the bin with the required form. |

#### Return value

LONG. The paper source.

---
