' ########################################################################################
' Microsoft Windows
' File: CDialog.inc
' Contents: Dialog box
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/windowsx.bi"
#include once "win/richedit.bi"
#include once "Afx/AfxWin.inc"
USING Afx

' // For compatibility with the Power Basic syntax
#define CBHNDL(wp, lp) cast(HWND, lp)
#define CBCTL(wp, lp) LOWORD(wp)
#define CBCTLMSG(wp, lp) HIWORD(wp)
#define CBNMCODE(wp, lp) LOWORD(wp) 
#define CBNMHWND(wp, lp) cast(HWND, lp)
#define CBNMID(wp, lp) HIWORD(wp)
#define CBNMHDR(wp, lp) cast(NMHDR PTR, lp)

NAMESPACE Afx

' ========================================================================================
' CDialog class
' ========================================================================================
TYPE CDialog

   Private:
      m_hDlg  AS HWND                         ' // Dialog handle
      m_hInstance AS HINSTANCE                ' // Instance handle
      m_hRichEditLib AS HMODULE               ' // Rich Edit module handle
      m_ptSize AS LONG                        ' // Font point size
      m_fontStyle AS BYTE                     ' // Font style
      m_charset AS BYTE                       ' // Font character set
      m_fontName AS WSTRING * 32              ' // Font face name
      m_IsModal AS BOOLEAN                    ' // Dialog is modal
      m_DlgRetVal AS INT_PTR                  ' // To simulate EndDialog
      m_hDefPushButton AS HWND                ' // Handle of the default push button
      DIM m_rgUserData(0 TO 99) AS LONG_PTR   ' // User data

   Public:
      DECLARE CONSTRUCTOR (BYREF fontName AS WSTRING = "Segoe UI", BYVAL ptSize AS LONG = 9, _
         BYVAL fontStyle AS BYTE = 0, BYVAL charset AS BYTE = DEFAULT_CHARSET) 
      DECLARE DESTRUCTOR 
      DECLARE FUNCTION DialogTemplate (BYREF Title AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
         BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0) AS STRING
      DECLARE FUNCTION DialogNew (BYVAL hParent AS HWND = HWND_DESKTOP, BYREF wszTitle AS WSTRING = "", _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
         BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0) AS HWND
      DECLARE FUNCTION DialogShowModeless (BYVAL pCallback AS ANY PTR = NULL) AS INT_PTR
      DECLARE FUNCTION DialogShowModal (BYVAL pCallback AS ANY PTR = NULL) AS INT_PTR
      DECLARE FUNCTION DialogDoEvents (BYVAL dwMilliseconds AS DWORD = 1) AS BOOLEAN
      DECLARE FUNCTION DialogEnd (BYVAL nResult AS LONG = 0) AS INT_PTR
      DECLARE FUNCTION DialogEndResult () AS INT_PTR
      DECLARE FUNCTION ControlAdd (BYREF wszClassName AS WSTRING, BYVAL cId AS LONG_PTR = 0, _
         BYREF wszTitle AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
         BYVAL dwStyle AS LONG = -1, BYVAL dwExStyle AS LONG = -1, BYVAL lpPAram AS LONG_PTR = 0, _
         BYVAL pWndProc AS SUBCLASSPROC = NULL, BYVAL uIdSubclass AS UINT_PTR = &HFFFFFFFF, _
         BYVAL dwRefData AS DWORD_PTR = NULL) AS HWND
      DECLARE PROPERTY hDialog () AS HWND
      DECLARE PROPERTY hDialog (BYVAL hDlg AS HWND)
      DECLARE FUNCTION IsModal () AS BOOLEAN
      DECLARE FUNCTION FontNew (BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, _
         BYVAL lWeight AS LONG = 0, BYVAL bItalic AS UBYTE = FALSE, BYVAL bUnderline AS UBYTE = FALSE, _
         BYVAL bStrikeOut AS UBYTE = FALSE, BYVAL bCharSet AS UBYTE = DEFAULT_CHARSET) AS HFONT
      DECLARE SUB FontEnd (BYVAL hFont AS HFONT)
      DECLARE FUNCTION GetDefId () AS INT_PTR
      DECLARE FUNCTION SetDefId (BYVAL cId AS LONG) AS BOOLEAN
      DECLARE FUNCTION DialogReposition () AS BOOLEAN
      DECLARE FUNCTION DialogShowState (BYVAL showState AS LONG) AS BOOLEAN
      DECLARE FUNCTION DialogMaximize () AS BOOLEAN
      DECLARE FUNCTION DialogMinimize () AS BOOLEAN
      DECLARE FUNCTION DialogHide () AS BOOLEAN
      DECLARE FUNCTION DialogNormalize () AS BOOLEAN
      DECLARE FUNCTION DialogDisable () AS BOOLEAN
      DECLARE FUNCTION DialogEnable () AS BOOLEAN
      DECLARE FUNCTION DialogRedraw () AS BOOLEAN
      DECLARE FUNCTION DialogPost (BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
      DECLARE FUNCTION DialogSend (BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
      DECLARE SUB DialogSetIconEx (BYVAL hIconBig AS HICON, BYVAL hIconSmall AS HICON)
      DECLARE FUNCTION DialogUnitsToPixels (BYVAL dluX AS LONG, BYVAL dluY AS LONG, BYREF pixelX AS LONG, BYREF pixelY AS LONG) AS BOOLEAN
      DECLARE FUNCTION PixelsToDialogUnits (BYVAL pixelX AS LONG, BYVAL pixelY AS LONG, BYREF dluX AS LONG, BYREF dluY AS LONG) AS BOOLEAN
      DECLARE FUNCTION DialogGetSize (BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
      DECLARE FUNCTION DialogGetWidth () AS LONG
      DECLARE FUNCTION DialogGetHeight () AS LONG
      DECLARE FUNCTION DialogSetSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
      DECLARE FUNCTION DialogGetClient (BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
      DECLARE FUNCTION DialogGetClientWidth () AS LONG
      DECLARE FUNCTION DialogGetClientHeight () AS LONG
      DECLARE FUNCTION DialogSetClient (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEaN
      DECLARE FUNCTION DialogGetLoc (BYREF nLeft AS LONG, BYREF nTop AS LONG) AS BOOLEAN
      DECLARE FUNCTION DialogSetLoc (BYVAL nLeft AS LONG, BYVAL nTop AS LONG) AS BOOLEAN
      DECLARE FUNCTION DialogGetText () AS CWSTR
      DECLARE FUNCTION DialogSetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
      DECLARE FUNCTION DialogGetUser (BYVAL idx AS LONG) AS LONG_PTR
      DECLARE SUB DialogSetUser (BYVAL idx AS LONG, BYVAL newValue AS LONG_PTR)
      DECLARE FUNCTION DialogGetFont () AS HFONT
      DECLARE FUNCTION DialogGetFontFaceName () AS CWSTR
      DECLARE FUNCTION DialogGetFontPointSize () AS LONG

      ' // Dialog procedure callback
      DECLARE FUNCTION CBGetDlgMsgResult () AS LONG_PTR
      DECLARE FUNCTION CBSetDlgMsgResult (BYVAL result AS LONG_PTR) AS LONG_PTR

      ' To do: DialogStabilize, DialogNonStable, DialogSetColor, DialogSetIcon (from resource)
      ' DialogStabilize: Disabling the close menu and processing the SC_CLOSE message in the
      ' WM_SYSCOMMAND message has the same effect.

      DECLARE FUNCTION MenuAttach (BYVAL hMenu AS HMENU) AS BOOLEAN
      DECLARE FUNCTION MenuDestroy () AS BOOLEAN
      DECLARE FUNCTION SysMenuRemoveCloseOption () AS BOOLEAN
      DECLARE FUNCTION SysMenuRestoreCloseOption () AS BOOLEAN
      ' To do: Add more menu methods

      ' // Control methods
      DECLARE FUNCTION ControlDisable (BYVAL cId AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlEnable (BYVAL cId AS LONG) AS BOOLEAN
      DECLARE SUB ControlSetFocus (BYVAL cId AS LONG)
      DECLARE FUNCTION ControlGetLoc (BYVAL cId AS LONG, BYREF nLeft AS LONG, BYREF nTop AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlSetLoc (BYVAL cId AS LONG, BYVAL nLeft AS LONG, BYVAL nTop AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlGetText (BYVAL cId AS LONG) AS CWSTR
      DECLARE FUNCTION ControlSetText (BYVAL cId AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
      DECLARE FUNCTION ControlHandle (BYVAL cId AS LONG) AS HWND
      DECLARE FUNCTION ControlHide (BYVAL cId AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlNormalize (BYVAL cId AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlKill (BYVAL cId AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlRedraw (BYVAL cId AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlShowState (BYVAL cId AS LONG, BYVAL showState AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlPost (BYVAL cId AS LONG, BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
      DECLARE FUNCTION ControlSend (BYVAL cId AS LONG, BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
      DECLARE FUNCTION ControlGetSize (BYVAL cId AS LONG, BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlGetWidth (BYVAL cId AS LONG) AS LONG
      DECLARE FUNCTION ControlGetHeight (BYVAL cId AS LONG) AS LONG
      DECLARE FUNCTION ControlSetSize (BYVAL cId AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlGetClient (BYVAL cId AS LONG, BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlGetClientWidth (BYVAL cId AS LONG) AS LONG
      DECLARE FUNCTION ControlGetClientHeight (BYVAL cId AS LONG) AS LONG
      DECLARE FUNCTION ControlSetClient (BYVAL cId AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
      DECLARE FUNCTION ControlGetFont (BYVAL cId AS LONG) AS HFONT
      DECLARE FUNCTION ControlGetFontFaceName (BYVAL cId AS LONG) AS CWSTR
      DECLARE FUNCTION ControlGetFontPointSize (BYVAL cId AS LONG) AS LONG
      DECLARE SUB ControlSetFont (BYVAL cId AS LONG, BYVAL hFont AS HFONT, BYVAL fRedraw AS BOOLEAN = TRUE)

Private:
      DECLARE STATIC FUNCTION CDialogProc (BYVAL hDlg AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' * Dialog constructor.
' ========================================================================================
PRIVATE CONSTRUCTOR CDialog (BYREF fontName AS WSTRING = "Segoe UI", BYVAL ptSize AS LONG = 9, BYVAL fontStyle AS BYTE = 0, BYVAL charset AS BYTE = DEFAULT_CHARSET) 
   ' // Store the passed parameters
   m_fontName = fontName
   m_ptSize = ptSize
   m_fontStyle = fontStyle
   m_charset = charset
   ' // Instance handle
   m_hInstance = GetModuleHandleW(NULL)
   ' // Initialize the common controls library
   DIM icc AS INITCOMMONCONTROLSEX
   icc.dwSize = SIZEOF(icc)
   icc.dwICC  = ICC_NATIVEFNTCTL_CLASS OR ICC_COOL_CLASSES OR ICC_BAR_CLASSES OR _
                ICC_TAB_CLASSES OR ICC_USEREX_CLASSES OR ICC_WIN95_CLASSES OR _
                ICC_STANDARD_CLASSES OR ICC_ANIMATE_CLASS OR ICC_DATE_CLASSES OR _
                ICC_HOTKEY_CLASS OR ICC_INTERNET_CLASSES OR ICC_LISTVIEW_CLASSES OR _
                ICC_PAGESCROLLER_CLASS OR ICC_PROGRESS_CLASS OR ICC_TREEVIEW_CLASSES OR _
                ICC_UPDOWN_CLASS
   InitCommonControlsEx(@icc)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CDialog
   IF m_hRichEditLib THEN FreeLibrary m_hRichEditLib
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' * Creates a dialog template in memory.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogTemplate (BYREF wszTitle AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
   BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0) AS STRING
   OutputDebugStringW("Dialog Template ")

   ' // Default styles. PB also includes DS_3DLOOK, but this style is obsolete since the system automatically
   ' // applies the three-dimensional look to dialog boxes created by applications.
   IF dwStyle = 0 THEN dwStyle = WS_VISIBLE OR DS_CENTER OR DS_MODALFRAME OR DS_NOFAILCREATE OR DS_SETFONT OR WS_OVERLAPPEDWINDOW _
      OR WS_CAPTION OR WS_CLIPSIBLINGS OR WS_DLGFRAME OR WS_POPUP OR WS_SYSMENU
   ' // Default extended styles.
   IF dwExStyle = 0 THEN dwExStyle = WS_EX_LEFT OR WS_EX_LTRREADING OR WS_EX_RIGHTSCROLLBAR

   ' // We need raw unicode in the strings
   DIM sTitle AS STRING = SPACE(LEN(wszTitle) * 2)
   memcpy STRPTR(sTitle), VARPTR(wszTitle), LEN(wszTitle) * 2
   DIM sFontName AS STRING = SPACE(LEN(m_FontName) * 2)
   memcpy STRPTR(sFontName), VARPTR(m_FontName), LEN(m_FontName) * 2

   DIM sTemplateEx AS STRING = _
      MKShort(1) + _                   ' // Version is always 1
      MKShort(&hFFFF) + _              ' // Use extended dialog box template signature
      MKL(0) + _                       ' // Help context identifier
      MKL(dwExStyle) + _               ' // Window extended style
      MKL(dwStyle OR DS_SETFONT) + _   ' // Window style
      MKShort(0) + _                   ' // Number of controls
      MKShort(x) + _                   ' // x coordinate
      MKShort(y) + _                   ' // y coordinate
      MKShort(cx) + _                  ' // Width of the window
      MKShort(cy) + _                  ' // Height of the window
      MKShort(0) + _                   ' // Menu
      MKShort(0) + _                   ' // Dialog class Name
      sTitle & CHR(0) & CHR(0) + _     ' // Dialog caption (raw unicode with double nulls)
      MKShort(m_ptSize) + _            ' // Font point size
      MKShort(FW_NORMAL) + _           ' // Font weight
      CHR(m_fontStyle) + _             ' // Font Italic flag
      CHR(m_charset) + _               ' // Font charset
      sFontName & CHR(0) & CHR(0)      ' // Font face name (raw unicode with double nulls)

   FUNCTION = sTemplateEx
   
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a control to the dialog.
' Note: Although the control identifier is a LONG, we must declare it here as a LONG_PTR
' because we have to cast it in CreateWindowEx to the member HMENU, which is an handle
' (4 bytes in 32-bit, 8 bytes in 54-bit).
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlAdd ( _
   BYREF wszClassName AS WSTRING, _                       ' // Class name
   BYVAL cId AS LONG_PTR = 0, _                           ' // Control identifier
   BYREF wszTitle AS WSTRING = "", _                      ' // Control caption
   BYVAL x AS LONG = 0, _                                 ' // Horizontal position
   BYVAL y AS LONG = 0, _                                 ' // Vertical position
   BYVAL nWidth AS LONG = 0, _                            ' // Control width
   BYVAL nHeight AS LONG = 0, _                           ' // Control height
   BYVAL dwStyle AS LONG = -1, _                          ' // Control style
   BYVAL dwExStyle AS LONG = -1, _                        ' // Extended style
   BYVAL lpParam AS LONG_PTR = 0, _                       ' // Pointer to custom data
   BYVAL pWndProc AS SUBCLASSPROC = NULL, _               ' // Address of the window callback procedure
   BYVAL uIdSubclass AS UINT_PTR = &HFFFFFFFF, _          ' // The subclass ID
   BYVAL dwRefData as DWORD_PTR = NULL _                  ' // Pointer to reference data
   ) AS HWND                                              ' // Control handle

   DIM hCtl AS HWND
   IF LEN(wszClassName) = 0 THEN EXIT FUNCTION
   DIM bSetFont AS LONG = CTRUE

   ' // Window styles
   DIM wsClassName AS WSTRING * 260
   wsClassName = wszClassName
   SELECT CASE UCASE(wsClassName)
      CASE "BUTTON"
         ' Adds a button to the window
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_CENTER OR BS_VCENTER
         IF dwStyle = BS_FLAT THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_CENTER OR BS_VCENTER OR BS_FLAT
         IF dwStyle = BS_DEFPUSHBUTTON THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_CENTER OR BS_VCENTER OR BS_DEFPUSHBUTTON
         IF dwStyle = BS_OWNERDRAW THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_OWNERDRAW
         #if _WIN32_WINNT = &h0602
         IF dwStyle = BS_SPLITBUTTON THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_CENTER OR BS_VCENTER OR BS_SPLITBUTTON
         IF dwStyle = BS_DEFSPLITBUTTON THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_CENTER OR BS_VCENTER OR BS_DEFSPLITBUTTON
         #endif
      CASE "CUSTOMBUTTON", "OWNERDRAWBUTTON"
         ' Adds an ownerdraw button to the window.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_OWNERDRAW
      CASE "RADIOBUTTON", "OPTION"
         ' Adds a radio button to the window.
         ' Note: In PowerBASIC this control is called "Option".
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER
         IF dwStyle = WS_GROUP THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER OR WS_GROUP
      CASE "CHECKBOX"
         ' Adds a checkbox to the window.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTOCHECKBOX OR BS_LEFT OR BS_VCENTER
      CASE "CHECK3STATE"
         ' Adds a 3 state checkbox to the window.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTO3STATE OR BS_LEFT OR BS_VCENTER
      CASE "LABEL"
         ' Adds a label to the window.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR SS_LEFT OR WS_GROUP OR SS_NOTIFY
      CASE "BITMAPLABEL"
         ' Adds an image label to the window.
         ' You must delete the bitmap before the application ends.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR SS_BITMAP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "ICONLABEL"
         ' Adds an image label to the window.
         ' You must delete the icon before the application ends.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR SS_ICON
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "BITMAPBUTTON"
         ' Adds an image button to the window.
         ' You must delete the bitmap before the application ends.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_BITMAP
      CASE "ICONBUTTON"
         ' Adds an image button to the window.
         ' You must delete the icon before the application ends.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_ICON
      CASE "CUSTOMLABEL"
         ' Adds an ownerdraw label to the window.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR SS_OWNERDRAW
         bSetFont = FALSE
      CASE "FRAME", "FRAMEWINDOW"
         ' Adds a frame to the window.
         ' Note: This is not the same that PowerBASIC DDT's Frame control, that in fact is a Group Box.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPSIBLINGS OR WS_GROUP OR SS_BLACKFRAME
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "GROUPBOX"
         ' Adds a group box to the window.
         ' Note: This is the same that DDT's frame control.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPSIBLINGS OR WS_GROUP OR BS_GROUPBOX
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
      CASE "LINE"
         ' Adds an horizontal line to the window
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR SS_ETCHEDFRAME
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "EDIT", "TEXTBOX"
         ' Adds an edit control to the window.
         wsClassName = "Edit"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR ES_AUTOHSCROLL
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "EDITMULTILINE", "MULTILINETEXTBOX"
         ' Adds an edit control to the window.
         wsClassName = "Edit"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR WS_VSCROLL OR ES_LEFT OR ES_AUTOHSCROLL OR ES_MULTILINE OR ES_NOHIDESEL OR ES_WANTRETURN
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "COMBOBOX"
         ' Adds a combo box to the window.
         IF dwStyle = -1 THEN dwStyle = WS_CHILD OR WS_VISIBLE OR WS_VSCROLL OR WS_BORDER OR WS_TABSTOP OR CBS_DROPDOWN OR CBS_HASSTRINGS OR CBS_SORT
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "COMBOBOXEX", "COMBOBOXEX32"
         ' Adds a combo box ex to the window.
         wsClassName = "ComboBoxEx32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR CBS_DROPDOWNLIST
      CASE "LISTBOX"
         ' Adds a list box to the window.
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_HSCROLL OR WS_VSCROLL OR WS_BORDER OR WS_TABSTOP OR LBS_STANDARD OR LBS_HASSTRINGS OR LBS_SORT OR LBS_NOTIFY
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "PROGRESSBAR", "MSCTLS_PROGRESS32"
         ' Adds a progress bar to the window.
         wsClassName = "msctls_progress32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE
         bSetFont = FALSE
      CASE "HEADER", "SYSHEADER32"
         ' Adds an header control to the window.
         wsClassName = "SysHeader32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR CCS_TOP OR HDS_HORZ OR HDS_BUTTONS
      CASE "TREEVIEW", "SYSTREEVIEW32"
         ' Adds a tree view control to the window.
         wsClassName = "SysTreeView32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR TVS_HASBUTTONS OR TVS_HASLINES OR TVS_LINESATROOT OR TVS_SHOWSELALWAYS
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "LISTVIEW", "SYSLISTVIEW32"
         ' Adds a list view control to the window.
         wsClassName = "SysListView32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPCHILDREN OR WS_TABSTOP OR LVS_REPORT OR LVS_SHOWSELALWAYS OR LVS_SHAREIMAGELISTS OR LVS_AUTOARRANGE OR LVS_EDITLABELS OR LVS_ALIGNTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "TOOLBAR", "TOOLBARWINDOW32"
         ' Adds a toolbar control to the window.
         wsClassName = "ToolbarWindow32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR CCS_TOP OR WS_BORDER OR TBSTYLE_FLAT OR TBSTYLE_TOOLTIPS
      CASE "REBAR", "REBARWINDOW32"
         ' Adds a rebar control to the window.
         wsClassName = "ReBarWindow32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_BORDER OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR CCS_NODIVIDER OR RBS_VARHEIGHT OR RBS_BANDBORDERS
      CASE "DATETIMEPICKER", "SYSDATETIMEPICK32"
         ' Adds a date time picker control to the window.
         wsClassName = "SysDateTimePick32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR DTS_SHORTDATEFORMAT
      CASE "MONTHCALENDAR", "MONTHCAL", "SYSMONTHCAL32"
         ' Adds a month calendar control to the window.
         wsClassName = "SysMonthCal32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "IPADDRESS", "SYSIPADDRESS32"
         ' Adds an IPAddress control to the window.
         wsClassName = "SysIPAddress32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "HOTKEY", "MSCTLS_HOTKEY32"
         ' Adds an hotkey control to the window.
         wsClassName = "msctls_hotkey32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "ANIMATE", "ANIMATION", "SYSANIMATE32"
         ' Adds an animation control to the window.
         wsClassName = "SysAnimate32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR ACS_TRANSPARENT
      CASE "SYSLINK"
         ' Adds a SysLink control to the window.
         ' Note: The SysLink control is defined in the ComCtl32.dll version 6 and requires a manifest
         ' or directive that specifies that version 6 of the DLL should be used if it is available.
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         bSetFont = FALSE
      CASE "PAGER", "SYSPAGER"
         ' Adds a Pager control to the window.
         wsClassName = "SysPager"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR PGS_HORZ
         bSetFont = FALSE
      CASE "TAB", "TABCONTROL", "SYSTABCONTROL32"
         ' Adds a Tab control to the window.
         wsClassName = "SysTabControl32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR WS_TABSTOP OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR TCS_TABS OR TCS_SINGLELINE OR TCS_RAGGEDRIGHT
         IF dwExStyle = -1 THEN dwExStyle = 0
         dwExStyle =  dwExStyle OR WS_EX_CONTROLPARENT
      CASE "STATUSBAR", "MSCTLS_STATUSBAR32"
         ' Adds a StatusBar control to the window.
         wsClassName = "msctls_statusbar32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR CCS_BOTTOM OR SBARS_SIZEGRIP
      CASE "SIZEBAR", "SIZEBOX", "SIZEGRIP"
         ' Adds a size box to the window.
         wsClassName = "Scrollbar"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR SBS_SIZEGRIP OR SBS_SIZEBOXBOTTOMRIGHTALIGN
         bSetFont = FALSE
         nWidth = GetSystemMetrics(SM_CXVSCROLL)
         nHeight = GetSystemMetrics(SM_CYHSCROLL)
         DIM rcClient AS RECT
         GetClientRect(m_hDlg, @rcClient)
         x = rcClient.Right - nWidth
         y = rcClient.Bottom - nHeight
         bSetFont = FALSE
      CASE "HSCROLLBAR"
         ' Adds an horizontal scroll bar to the window.
         wsClassName = "Scrollbar"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR SBS_HORZ
         bSetFont = FALSE
      CASE "VSCROLLBAR"
         ' Adds a vertical scroll bar to the window.
         wsClassName = "Scrollbar"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR SBS_VERT
         bSetFont = FALSE
      CASE "TRACKBAR", "MSCTLS_TRACKBAR32", "SLIDER"
         wsClassName = "msctls_trackbar32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR TBS_AUTOTICKS OR TBS_HORZ OR TBS_BOTTOM OR TBS_TOOLTIPS
         bSetFont = FALSE
      CASE "UPDOWN", "MSCTLS_UPDOWN32"
         wsClassName = "msctls_updown32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR UDS_WRAP OR UDS_ARROWKEYS OR UDS_ALIGNRIGHT OR UDS_SETBUDDYINT
         bSetFont = FALSE
      CASE "RICHEDIT", "RICHEDIT50W"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR WS_HSCROLL OR WS_VSCROLL OR ES_AUTOHSCROLL OR ES_AUTOVSCROLL OR ES_MULTILINE OR ES_WANTRETURN OR ES_NOHIDESEL OR ES_SAVESEL
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
         wsClassName = "RichEdit50W"
         IF m_hRichEditLib = 0 THEN m_hRichEditLib = CAST(HMODULE, LoadLibraryW("MSFTEDIT.DLL"))
   END SELECT
   ' // Don't allow negative values for the styles
   IF dwStyle = -1 THEN dwStyle = 0
   IF dwExStyle = -1 THEN dwExStyle = 0
   ' // Make sure that the control has the WS_CHILD style
   dwStyle = dwStyle OR WS_CHILD

   ' // Create the control
   DIM rc AS RECT
   rc.Left = x : rc.Right  = x + nWidth
   rc.Top  = y : rc.Bottom = y + nHeight
   MapDialogRect m_hDlg, @rc
   hCtl = CreateWindowExW (dwExStyle, wsClassName, wszTitle, dwStyle, rc.Left, rc.Top, rc.Right - rc.Left, _
          rc.Bottom - rc.Top, m_hDlg, CAST(HMENU, cID), m_hInstance, CAST(LPVOID, lpParam))
   IF hCtl = NULL THEN EXIT FUNCTION
   ' // Set the font
   IF bSetFont THEN SendMessageW(hCtl, WM_SETFONT, SendMessageW(m_hDlg, WM_GETFONT, 0, 0), 0)
   SELECT CASE UCASE(wsClassName)
      CASE "BUTTON"
         IF (dwStyle AND BS_DEFPUSHBUTTON) = BS_DEFPUSHBUTTON THEN
            ' // Removes the default button because 32768 is not a valid control ID.
            ' // Needed if the user has set the BS_DEFBUTTON style to more of one button.
            ' // We don't want to have more than one default button.
            SendMessageW(m_hDlg, DM_SETDEFID, 32768, 0)
            ' // Sets the specified button as the default
            SendMessageW(m_hDlg, DM_SETDEFID, cId, 0)
            ' // DEfault push button
            m_hDefPushButton = hCtl
         END IF
      CASE "LISTBOX"
         ' // Adjust the height of the control so that the integral height
         ' // is based on the new font rather than the default SYSTEM_FONT
         SetWindowPos hCtl, NULL, x, y, nWidth, nHeight, SWP_NOZORDER
      CASE "DATETIMEPICKER", "SYSDATETIMEPICK32"
         ' // Sets the font to be used by the date and time picker control's child month calendar control.
         SendMessageW(hCtl, DTM_SETMCFONT, SendMessageW(m_hDlg, WM_GETFONT, 0, 0), CTRUE)
      CASE "PROGRESSBAR", "MSCTLS_PROGRESS32"
         ' // Set the default range
         .SendMessageW hCtl, PBM_SETRANGE32, 0, 100
         ' // Set the default initial value
         .SendMessageW hCtl, PBM_SETPOS, 0, 0
      CASE "TRACKBAR", "MSCTLS_TRACKBAR32"
         ' // Set the default range values
         .SendMessageW hCtl, TBM_SETRANGEMIN, CTRUE, 0
         .SendMessageW hCtl, TBM_SETRANGEMAX, CTRUE, 100
         ' // Set the default page size
         .SendMessageW hCtl, TBM_SETPAGESIZE, 0, 10
      CASE "UPDOWN", "MSCTLS_UPDOWN32"
         ' // Set the default base
         .SendMessageW hCtl, UDM_SETBASE, 10, 0
         ' // Set the default range values
         .SendMessageW hCtl, UDM_SETRANGE32, 100, 0
         ' // Set the default initial value
         .SendMessageW hCtl, UDM_SETPOS32, 0, 0
         ' // Correct for Windows using a default size for the updown control
         SetWindowPos hCtl, NULL, x, y, nWidth, nHeight, SWP_NOZORDER
      CASE "HSCROLLBAR", "VSCROLLBAR"
         ' // Initialize the scroll bar with default values
         DIM tsi AS SCROLLINFO
         tsi.cbSize = SIZEOF(tsi)
         tsi.fMask  = SIF_PAGE OR SIF_POS OR SIF_RANGE
         tsi.nMin   = 0
         tsi.nMax   = 100
         tsi.nPage  = 0
         tsi.nPos   = 0
         .SetScrollInfo hCtl, SB_CTL, @tsi, CTRUE
      CASE "TOOLBAR", "TOOLBARWINDOW32"
         ' // Set the button size
         DIM AS LONG nButtonWidth, nButtonHeight
         nButtonWidth = LOWORD(.SendMessageW(hCtl, TB_GETBUTTONSIZE, 0, 0))
         nButtonHeight = HIWORD(.SendMessageW(hCtl, TB_GETBUTTONSIZE, 0, 0))
         .SendMessageW hCtl, TB_SETBUTTONSIZE, 0, MAKELONG(nButtonWidth, nButtonHeight)
         ' // Send this message for backward compatibility
         .SendMessageW hCtl, TB_BUTTONSTRUCTSIZE, SIZEOF(TBBUTTON), 0
      CASE "BITMAPLABEL"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = .LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = .LoadImageW(m_hInstance, wszTitle, IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN .SendMessageW(hCtl, STM_SETIMAGE, IMAGE_BITMAP, CAST(LPARAM, hImage))
      CASE "ICONLABEL"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = .LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = .LoadImageW(m_hInstance, wszTitle, IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN .SendMessageW(hCtl, STM_SETIMAGE, IMAGE_ICON, CAST(LPARAM, hImage))
      CASE "BITMAPBUTTON"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = .LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = .LoadImageW(m_hInstance, wszTitle, IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN .SendMessageW(hCtl, BM_SETIMAGE, IMAGE_BITMAP, CAST(LPARAM, hImage))
      CASE "ICONBUTTON"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = .LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = .LoadImageW(m_hInstance, wszTitle, IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN .SendMessageW(hCtl, BM_SETIMAGE, IMAGE_ICON, CAST(LPARAM, hImage))
   END SELECT
   ' // Subclass the control if pWndProc is not null
   IF pWndProc <> NULL THEN
      IF uIdSubclass = &HFFFFFFFF THEN
         .SetPropW(hCtl, "OLDWNDPROC", CAST(HANDLE, .SetWindowLongPtrW(hCtl, GWLP_WNDPROC, CAST(LONG_PTR, pWndProc))))
      ELSE
'         .SetWindowSubclass hCtl, CAST(SUBCLASSPROC, pWndProc), uIdSubclass, dwRefData
         .SetWindowSubclass hCtl, pWndProc, uIdSubclass, dwRefData
      END IF
   END IF
   FUNCTION = hCtl
END FUNCTION
' =====================================================================================

' ========================================================================================
' Gets the handle of the dialog
' ========================================================================================
PRIVATE PROPERTY CDialog.hDialog () AS HWND
   RETURN m_hDlg
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the handle of the dialog
' ========================================================================================
PRIVATE PROPERTY CDialog.hDialog (BYVAL hDlg AS HWND)
   IF IsWindow(hDlg) THEN m_hDlg = hDlg
END PROPERTY
' ========================================================================================
' ========================================================================================
' Returns true if the dialog is modal
' ========================================================================================
PRIVATE FUNCTION CDialog.IsModal () AS BOOLEAN
   RETURN m_IsModal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Create a modeless dialog in memory
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogNew (BYVAL hParent AS HWND = HWND_DESKTOP, BYREF wszTitle AS WSTRING = "", _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
   BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0) AS HWND

   ' // Default styles
   IF dwStyle THEN dwStyle = dwStyle OR DS_3DLOOK OR DS_NOFAILCREATE OR DS_SETFONT
   IF dwStyle = 0 THEN dwStyle = DS_3DLOOK OR DS_SETFONT OR DS_MODALFRAME OR DS_NOFAILCREATE OR WS_BORDER OR _
                                 WS_CLIPSIBLINGS OR WS_DLGFRAME OR WS_POPUP
   IF dwExStyle = 0 THEN dwExstyle = WS_EX_LEFT OR WS_EX_LTRREADING OR WS_EX_RIGHTSCROLLBAR
   ' // Remove the WS_VISIBLE style to avoid partial images of the dialog to appear before it is completed.
   ' // The dialog will be sidplayed and activated in the ShowDialogModal or ShowDialogModeless methods with SHowWindow.
   dwStyle = dwStyle AND (NOT WS_VISIBLE)
   DIM sTemplate AS STRING = this.DialogTemplate(wszTitle, x, y, cx, cy, dwStyle, dwExStyle)
   m_hDlg = CreateDialogIndirectParamW(m_hInstance, cast(LPCDLGTEMPLATEW, STRPTR(sTemplate)), _
            hParent, NULL, cast(LPARAM, @this))
   RETURN m_hDlg

END FUNCTION
' ========================================================================================

' ========================================================================================
' Internal dialog callback procedure used if the user does not specify a callback
' in the ShowDialogModal or ShowDialogModeless methods.
' ========================================================================================
PRIVATE FUNCTION CDialog.CDialogProc (BYVAL hDlg AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_INITDIALOG
         OutputDebugStringW("CDialog.CDialogProc - WM_INITDIALOG " & WSTR(hDlg) & " - " & WSTR(wParam) & " - " & WSTR(lParam))

      CASE WM_CLOSE
         OutputDebugStringW("CDialog.CDialogProc - WM_CLOSE")
         DestroyWindow hDlg

      CASE WM_DESTROY
         OutputDebugStringW("CDialog.CDialogProc - WM_DESTROY")

   END SELECT

   FUNCTION = 0

END FUNCTION
' ========================================================================================

' ========================================================================================
' * Shows the dialog as modeless.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogShowModeless (BYVAL pCallback AS ANY PTR = NULL) AS INT_PTR
   m_IsModal = FALSE
   ' // Display the dialog
   ShowWindow m_hDlg, SW_SHOW
   ' // Set the dialog callback
   IF pCallback = NULL THEN pCallback = PROCPTR(CDialog.CDialogProc)
   SetWindowLongPtrW(m_hDlg, DWLP_DLGPROC, cast(LPARAM, pCallBack))
   ' // Send a WM_INITDIALOG message
   DIM hDefButton AS HWND = m_hDefPushButton
   '// IF the user has not set the BS_DEFPUSHBUTTON style to a button, pass the dialog handle
   IF hDefButton = NULL THEN hDefButton = m_hDlg
   SendMessage(m_hDlg, WM_INITDIALOG, cast(WPARAM, hDefButton), cast(LPARAM, @this))
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Shows the dialog as modal.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogShowModal (BYVAL pCallback AS ANY PTR = NULL) AS INT_PTR

   ' // Flag the dialog as modal
   m_IsModal = TRUE
   ' // Display the dialog
   ShowWindow m_hDlg, SW_SHOW
   ' // Set the dialog callback
   IF pCallback = NULL THEN pCallback = PROCPTR(CDialog.CDialogProc)
   SetWindowLongPtrW(m_hDlg, DWLP_DLGPROC, cast(LPARAM, pCallBack))
   ' // Send a WM_INITDIALOG message
   SendMessage(m_hDlg, WM_INITDIALOG, cast(WPARAM, m_hDlg), cast(LPARAM, @this))
   ' // Disable the parent window
   DIM hParent AS HWND = GetParent(m_hDlg)
   IF hParent THEN EnableWindow(hParent, FALSE)
   ' // Message handler loop
   DO
      this.DialogDoEvents
   LOOP WHILE IsWindow(m_hDlg)
   ' // Enable the parent window
   IF hParent THEN EnableWindow(hParent, TRUE)

   RETURN S_OK

END FUNCTION
' ========================================================================================

' ========================================================================================
' * Ends the dialog
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogEnd (BYVAL nResult AS LONG = 0) AS INT_PTR
   OutputDebugStringW("CDialog.DialogEnd - WM_DESTROY")
   m_DlgRetVal = nResult
   DestroyWindow(m_hDlg)
   RETURN m_DlgRetVal
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value passed to DialogEnd
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogEndResult () AS INT_PTR
   OutputDebugStringW("CDialog.DialogEndResult")
   RETURN m_DlgRetVal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Message pump for modeless dialogs
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogDoEvents (BYVAL dwMilliseconds AS DWORD = 1) AS BOOLEAN
   DIM uMsg AS tagMsg
   IF PeekMessageW(@uMsg, NULL, 0, 0, PM_REMOVE) THEN
      IF IsDialogMessageW(GetActiveWindow, @uMsg) = 0 THEN
         TranslateMessage @uMsg
         DispatchMessage @uMsg
      END IF
      RETURN TRUE
   ELSE
      Sleep_(dwMilliseconds)
      RETURN FALSE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the identifier of the default push button control for a dialog box.
' ========================================================================================
PRIVATE FUNCTION CDialog.GetDefId () AS INT_PTR
   RETURN LOWORD(SendMessageW(m_hDlg, DM_GETDEFID, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Changes the identifier of the default push button for a dialog box.
' ========================================================================================
PRIVATE FUNCTION CDialog.SetDefId (BYVAL cId AS LONG) AS BOOLEAN
   ' // Removes the default button because 32768 is not a valid control ID.
   SendMessageW(m_hDlg, DM_SETDEFID, 32768, 0)
   ' // Sets the specified button as the default
   RETURN SendMessageW(m_hDlg, DM_SETDEFID, cId, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Repositions a top-level dialog box so that it fits within the desktop area. 
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogReposition () AS BOOLEAN
   RETURN SendMessageW(m_hDlg, DM_REPOSITION, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Change the visible state of a dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogShowState (BYVAL showState AS LONG) AS BOOLEAN
   RETURN ShowWindow(m_hDlg, showState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Maximizes the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogMaximize () AS BOOLEAN
   RETURN ShowWindow(m_hDlg, SW_MAXIMIZE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Minimizes the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogMinimize () AS BOOLEAN
   RETURN ShowWindow(m_hDlg, SW_MINIMIZE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Hides the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogHide () AS BOOLEAN
   RETURN ShowWindow(m_hDlg, SW_HIDE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Makes the dialog visible at its normal size and position.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogNormalize () AS BOOLEAN
   RETURN ShowWindow(m_hDlg, SW_SHOWNORMAL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Disables the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogDisable () AS BOOLEAN
   RETURN EnableWindow(m_hDlg, FALSE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogEnable () AS BOOLEAN
   RETURN EnableWindow(m_hDlg, CTRUE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Redraws the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogRedraw () AS BOOLEAN
   RETURN RedrawWindow(m_hDlg, NULL, NULL, RDW_INVALIDATE OR RDW_UPDATENOW)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Places (posts) a message in the message queue associated with the thread that created the
' specified dialog and returns without waiting for the thread to process the message.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogPost (BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
   RETURN PostMessageW(m_hDlg, Msg, cast(WPARAM, wPrm), cast(LPARAM, lPrm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sends the specified message to a window or windows. The SendMessage function calls the
' window procedure for the dialog and does not return until the dialog procedure has
' processed the message.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogSend (BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
   RETURN SendMessageW(m_hDlg, Msg, cast(WPARAM, wPrm), cast(LPARAM, lPrm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Converts the specified dialog box units to screen units (pixels).
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogUnitsToPixels (BYVAL dluX AS LONG, BYVAL dluY AS LONG, BYREF pixelX AS LONG, BYREF pixelY AS LONG) AS BOOLEAN
   IF VARPTR(pixelX) = 0 OR VARPTR(pixelY) = 0 THEN RETURN FALSE
   DIM rc AS RECT
   rc.Left = dluX : rc.Top = dluY
   DIM bResult AS BOOLEAN = MapDialogRect(m_hDlg, @rc)
   IF bResult = FALSE THEN RETURN FALSE
   IF VARPTR(pixelX) THEN pixelX = rc.Left
   IF VARPTR(pixelY) THEN pixelY = rc.Top
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Converts the specied screen units (pixels) to dialog box units.
' ========================================================================================
PRIVATE FUNCTION CDialog.PixelsToDialogUnits (BYVAL pixelX AS LONG, BYVAL pixelY AS LONG, BYREF dluX AS LONG, BYREF dluY AS LONG) AS BOOLEAN
   IF VARPTR(pixelX) = 0 OR VARPTR(pixelY) = 0 THEN RETURN FALSE
   DIM rc AS RECT
   rc.Left = pixelX
   rc.Top = pixelY
   DIM bResult AS BOOLEAN = MapDialogRect(m_hDlg, @rc)
   IF bResult = FALSE THEN RETURN FALSE
   DIM rx AS SINGLE
   rx = rc.Left / pixelX
   DIM ry AS SINGLE
   ry = rc.Top / pixelY
   IF VARPTR(dluX) THEN dluX = pixelX / rx
   IF VARPTR(dluY) THEN dluY = pixelY / ry
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the width and height of the dialog in dialog units.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetSize (BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
   IF VARPTR(nWidth) = 0 OR VARPTR(nHeight) = 0 THEN RETURN FALSE
   ' // Get the dimensions of the window
   DIM rc AS RECT
   GetWindowRect(m_hDlg, @rc)
   ' // Return the width and height in dialog units
   RETURN this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the width of the dialog in dialog units.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetWidth () AS LONG
   ' // Get the dimensions of the window
   DIM rc AS RECT
   FUNCTION = GetWindowRect(m_hDlg, @rc)
   ' // Return the width in dialog units
   DIM AS LONG nWidth, nHeight
   this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
   RETURN nWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the height of the dialog in dialog units.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetHeight () AS LONG
   ' // Get the dimensions of the window
   DIM rc AS RECT
   FUNCTION = GetWindowRect(m_hDlg, @rc)
   ' // Return the width in dialog units
   DIM AS LONG nWidth, nHeight
   FUNCTION = this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
   RETURN nHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the size of the dialog in dialog units.
' - nWidth = The new width of the window.
' - nHeight = The new height of the window.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogSetSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
   this.DialogUnitsToPixels(nWidth, nHeight, nWidth, nHeight)
   RETURN SetWindowPos(m_hDlg, NULL, 0, 0, nWidth, nHeight, SWP_NOZORDER OR SWP_NOMOVE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the coordinates of the dialog's client area. The client coordinates specify the
' upper-left and lower-right corners of the client area. Because client coordinates are
' relative to the upper-left corner of a window's client area, the coordinates of the
' upper-left corner are (0,0).
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetClient (BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
   IF VARPTR(nWidth) = 0 OR VARPTR(nHeight) = 0 THEN RETURN FALSE
   ' // Get the dimensions of the window
   DIM rc AS RECT
   GetClientRect(m_hDlg, @rc)
   ' // Return the width and height values in dialog units
   RETURN this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the width of the the dialog's client area in dialog units.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetClientWidth () AS LONG
   ' // Get the dimensions of the window
   DIM rc AS RECT
   FUNCTION = GetClientRect(m_hDlg, @rc)
   ' // Return the width in dialog units
   DIM AS LONG nWidth, nHeight
   this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
   RETURN nWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the height of the dialog's client area in dialog units.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetClientHeight () AS LONG
   ' // Get the dimensions of the window
   DIM rc AS RECT
   FUNCTION = GetClientRect(m_hDlg, @rc)
   ' // Return the width in dialog units
   DIM AS LONG nWidth, nHeight
   this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
   RETURN nHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Adjusts the bounding rectangle of the dialog based on the desired size of the client area.
' Parameters:
' - nWidth = The new width of the client area of the dialog.
' - nHeight = The new height of the client area of the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogSetClient (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
   ' // Convert dialog units to pixels
   this.DialogUnitsToPixels(nWidth, nHeight, nWidth, nHeight)
   ' // Convert the client rectangle to a window rectangle.
   ' // The AdjustWindowRectEx function cannot take menu wrapping into account
   ' // because it doesn't know which menu we are using.
   DIM rc AS RECT, rcTemp  AS RECT
   SetRect(@rc, 0, 0, nWidth, nHeight)
   DIM hMenu AS HANDLE = GetMenu(m_hDlg)
   DIM dwStyle AS DWORD = GetWindowLongPtrW(m_hDlg, GWL_STYLE)
   AdjustWindowRectEx(@rc, dwStyle, (hMenu <> NULL), GetWindowLongPtrW(m_hDlg, GWL_EXSTYLE))
   ' // If there is a menu, we need to check how much wrapping occurs when we set
   ' // the window to the width specified by AdjustWindowRectEX and an infinite
   ' // amount of height. An infinite height allows us to see every single menu wrap.
   IF hMenu <> NULL THEN
      rcTemp = rc
      rcTemp.Bottom = &H7FFF   ' // "Infinite" height
      SendMessageW(m_hDlg, WM_NCCALCSIZE, 0, CAST(LPARAM, @rcTemp))
      ' // Adjust our previous calculation to compensate for menu wrapping.
      rc.Bottom = rc.Bottom + rcTemp.Top
   END IF
   ' // The AdjustWindowRectEx function does not take the WS_VSCROLL or WS_HSCROLL
   ' // styles into account. To account for the scroll bars, we need to call the
   ' // GetSystemMetrics function with SM_CXVSCROLL or SM_CYHSCROLL.
   IF (dwStyle AND WS_HSCROLL) = WS_HSCROLL THEN
      rc.Bottom = rc.Bottom + GetSystemMetrics(SM_CYHSCROLL)
   END IF
   IF (dwStyle AND WS_VSCROLL) = WS_VSCROLL THEN
      rc.Right = rc.Right + GetSystemMetrics(SM_CXVSCROLL)
   END IF
   DIM cx AS LONG = rc.Right - rc.Left
   DIM cy AS LONG = rc.Bottom - rc.Top
   RETURN SetWindowPos(m_hDlg, NULL, 0, 0, cx, cy, SWP_NOZORDER OR SWP_NOMOVE OR SWP_NOACTIVATE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the handle of the font used by the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetFont () AS HFONT
   RETURN CAST(HFONT, SendMessageW(m_hDlg, WM_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the face name of the font used by the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetFontFaceName () AS CWSTR
   ' // Get the handle of the font used by the dialog
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(m_hDlg, WM_GETFONT, 0, 0))
   IF hCurFont = NULL THEN RETURN ""   
   ' // Get the LOGFONTW structure
   DIM lfw AS LOGFONTW
   IF hCurFont THEN GetObjectW(hCurFont, SIZEOF(lfw), @lfw)
   DIM hDc AS HDC = GetDC(m_hDlg)
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(m_hDlg, hDC)
   RETURN lfw.lfFaceName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the point size of the font used by the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetFontPointSize () AS LONG
   ' // Get the handle of the font used by the dialog
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(m_hDlg, WM_GETFONT, 0, 0))
   IF hCurFont = NULL THEN RETURN 0
   ' // Get the LOGFONTW structure
   DIM lfw AS LOGFONTW
   IF hCurFont THEN GetObjectW(hCurFont, SIZEOF(lfw), @lfw)
   DIM hDc AS HDC = GetDC(m_hDlg)
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(m_hDlg, hDC)
   DIM ptSize AS LONG = MulDiv(lfw.lfHeight, 72, cyPixelsPerInch)
   IF ptSize < 0 THEN ptSize = -ptSize
   RETURN ptSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Attaches a menu to the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.MenuAttach (BYVAL hMenu AS HMENU) AS BOOLEAN
   RETURN SetMenu(m_hDlg, hMenu)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Destroys the main menu from the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.MenuDestroy () AS BOOLEAN
   ' // Get the menu
   DIM hMenu AS HMENU = GetMenu(m_hDlg)
   IF hMenu = NULL THEN RETURN FALSE
   ' // Detach the menu
   SetMenu(m_hDlg, NULL)
   ' // Destroy the menu
   IF DestroyMenu(hMenu) THEN
      ' // Redraw the menu
      DrawMenuBar m_hDlg
      RETURN TRUE
   ELSE RETURN FALSE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the text in a dialog or window caption.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetText () AS CWSTR
   DIM nLen AS LONG = SendMessageW(m_hDlg, WM_GETTEXTLENGTH, 0, 0)
   DIM cwsText AS CWSTR = SPACE(nLen + 1)
   SendMessageW(m_hDlg, WM_GETTEXT, nLen + 1, cast(LPARAM, *cwsText))
   RETURN LEFT(**cwsText, LEN(cwsText) - 1)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the text in a dialog or window caption.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogSetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   RETURN SendMessageW(m_hDlg, WM_SETTEXT, 0, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the location of the top left corner of the window in dialog units.
' The location is relative to the upper-left corner of the client area in the parent window.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogGetLoc (BYREF nLeft AS LONG, BYREF nTop AS LONG) AS BOOLEAN
   IF VARPTR(nLeft) = 0 OR VARPTR(nTop) = 0 THEN RETURN FALSE
   DIM rc AS RECT
   ' // Get the dimensions of the window
   GetWindowRect(m_hDlg, @rc)
   ' // Convert the coordinates to be relative to the parent
   MapWindowPoints(HWND_DESKTOP, GetParent(m_hDlg), CAST(POINT PTR, @rc), 2)
   ' // Return the left and top values
   RETURN this.PixelsToDialogUnits(rc.Left, rc.Top, nLeft, nTop)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the location of the top left corner of the window in dialog units.
' The location is relative to the upper-left corner of the client area in the parent window.
' Return value: TRUE or FALSE. To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogSetLoc (BYVAL nLeft AS LONG, BYVAL nTop AS LONG) AS BOOLEAN
   DialogUnitsToPixels(nLeft, nTop, nLeft, nTop)
   RETURN SetWindowPos(m_hDlg, 0, nLeft, nTop, 0, 0, SWP_NOSIZE OR SWP_NOZORDER)
END FUNCTION
' ========================================================================================

' =====================================================================================
' * Retrieves a value from the user data area of a CDialog.
' idx = The index number of the user data value to retrieve, in the range 0 to 99 inclusive.
' =====================================================================================
PRIVATE FUNCTION CDialog.DialogGetUser (BYVAL idx AS LONG) AS LONG_PTR
   IF idx < 0 OR idx > 99 THEN RETURN 0
   RETURN m_rgUserData(idx)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets a value in the user data area of a CWindow.
' idx = The index number of the user data value to retrieve, in the range 0 to 99 inclusive.
' =====================================================================================
PRIVATE SUB CDialog.DialogSetUser (BYVAL idx AS LONG, BYVAL newValue AS LONG_PTR)
   IF idx > -1 AND idx < 100 THEN m_rgUserData(idx) = newValue
END SUB
' =====================================================================================

' =====================================================================================
' * Sets the big and small icons.
' =====================================================================================
PRIVATE SUB CDialog.DialogSetIconEx (BYVAL hIconBig AS HICON, BYVAL hIconSmall AS HICON)
   SendMessageW(m_hDlg, WM_SETICON, ICON_BIG, cast(LPARAM, hIconBig))
   SendMessageW(m_hDlg, WM_SETICON, ICON_SMALL, cast(LPARAM, hIconSmall))
END SUB
' =====================================================================================

' =====================================================================================
' * Removes the system menu close option and disables Alt+F4 and the X button.
' =====================================================================================
PRIVATE FUNCTION CDialog.SysMenuRemoveCloseOption () AS BOOLEAN
   ' // Get the system menu handle
   DIM hMenu AS HMENU = GetSystemMenu(m_hDlg, FALSE)
   IF hMenu = NULL THEN RETURN FALSE
   ' // Get the number of menu items
   DIM cbItems AS LONG = GetMenuItemCount(hMenu)
   IF cbItems = 0 THEN RETURN FALSE
   ' // Remove the close menu item
   IF RemoveMenu(hMenu, cbItems - 1, MF_REMOVE OR MF_BYPOSITION) = 0 THEN RETURN FALSE
   ' // Remove the separator line
   IF RemoveMenu(hMenu, cbItems - 2, MF_REMOVE OR MF_BYPOSITION) = 0 THEN RETURN FALSE
   ' // Redraw the menu (this refreshes the caption bar, dimming the X button)
   DrawMenuBar(m_hDlg)
   RETURN TRUE
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Restores the system menu close option and enables Alt+F4 and the X button.
' =====================================================================================
PRIVATE FUNCTION CDialog.SysMenuRestoreCloseOption () AS BOOLEAN
   IF GetSystemMenu(m_hDlg, TRUE) <> NULL THEN RETURN FALSE
   DrawMenuBar(m_hDlg)
   RETURN TRUE
END FUNCTION
' =====================================================================================

' ========================================================================================
' * Sets the focus in the specified control of a dialog box.
' See https://devblogs.microsoft.com/oldnewthing/20040802-00/?p=38283
' ========================================================================================
PRIVATE SUB CDialog.ControlSetFocus (BYVAL cId AS LONG)
   SendMessageW(m_hDlg, WM_NEXTDLGCTL, cast(WPARAM, GetDlgItem(m_hDlg, cId)), CTRUE)
END SUB
' ========================================================================================

' ========================================================================================
' * Disables the control.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlDisable (BYVAL cId AS LONG) AS BOOLEAN
   RETURN EnableWindow(GetDlgItem(m_hDlg, cId), FALSE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Enables the control.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlEnable (BYVAL cId AS LONG) AS BOOLEAN
   RETURN EnableWindow(GetDlgItem(m_hDlg, cId), CTRUE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the location of the top left corner of the window in dialog units.
' The location is relative to the upper-left corner of the client area in the parent window.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetLoc (BYVAL cId AS LONG, BYREF nLeft AS LONG, BYREF nTop AS LONG) AS BOOLEAN
   IF VARPTR(nLeft) = 0 OR VARPTR(nTop) = 0 THEN RETURN FALSE
   DIM rc AS RECT
   ' // Get the dimensions of the window
   DIM hCtl AS HWND = GetDlgItem(m_hDlg, cID)
   GetWindowRect(hCtl, @rc)
   ' // Convert the coordinates to be relative to the parent
   MapWindowPoints(HWND_DESKTOP, GetParent(hCtl), CAST(POINT PTR, @rc), 2)
   ' // Return the left and top values
   RETURN this.PixelsToDialogUnits(rc.Left, rc.Top, nLeft, nTop)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the location of the top left corner of the window in dialog units.
' The location is relative to the upper-left corner of the client area in the parent window.
' Return value: TRUE or FALSE. To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlSetLoc (BYVAL cId AS LONG, BYVAL nLeft AS LONG, BYVAL nTop AS LONG) AS BOOLEAN
   DialogUnitsToPixels(nLeft, nTop, nLeft, nTop)
   RETURN SetWindowPos(GetDlgItem(m_hDlg, cId), 0, nLeft, nTop, 0, 0, SWP_NOSIZE OR SWP_NOZORDER)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the text in a control or control caption.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetText (BYVAL cId AS LONG) AS CWSTR
   DIM hCtl AS HWND = GetDlgItem(m_hDlg, cID)
   DIM nLen AS LONG = SendMessageW(hCtl, WM_GETTEXTLENGTH, 0, 0)
   DIM cwsText AS CWSTR = SPACE(nLen + 1)
   SendMessageW(hCtl, WM_GETTEXT, nLen + 1, cast(LPARAM, *cwsText))
   RETURN LEFT(**cwsText, LEN(cwsText) - 1)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the text in a control or control caption.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlSetText (BYVAL cId AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   RETURN SendMessageW(GetDlgItem(m_hDlg, cID), WM_SETTEXT, 0, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a window handle for the specified control ID.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlHandle (BYVAL cId AS LONG) AS HWND
   RETURN GetDlgItem(m_hDlg, cID)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Hides the specified control.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlHide (BYVAL cId AS LONG) AS BOOLEAN
   RETURN ShowWindow(GetDlgItem(m_hDlg, cId), SW_HIDE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Makes visible the specified control.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlNormalize (BYVAL cId AS LONG) AS BOOLEAN
   RETURN ShowWindow(GetDlgItem(m_hDlg, cId), SW_SHOW)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Destroys the specified window.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlKill (BYVAL cId AS LONG) AS BOOLEAN
   RETURN DestroyWindow(GetDlgItem(m_hDlg, cId))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Change the visible state of a control.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlShowState (BYVAL cId AS LONG, BYVAL showState AS LONG) AS BOOLEAN
   RETURN ShowWindow(GetDlgItem(m_hDlg, cId), showState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Redraws the specified control.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlRedraw (BYVAL cId AS LONG) AS BOOLEAN
   RETURN RedrawWindow(GetDlgItem(m_hDlg, cId), NULL, NULL, RDW_INVALIDATE OR RDW_UPDATENOW)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a logical font.
' Parameters:
' - wszFaceName = The typeface name.
' - lPointSize = The point size.
' - DPI = Dots per inch to calculate scaling. Default value = 96 (no scaling). If you pass -1
'   and the application is DPI aware, the DPI value used by the operating system will be used.
' - lWeight = The weight of the font in the range 0 through 1000. For example, 400 is normal
'      and 700 is bold. If this value is zero, a default weight is used.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
' - bItalic = Italic flag. CTRUE or FALSE
' - bUnderline = Underline flag. CTRUE or FALSE
' - bStrikeOut = StrikeOut flag. CTRUE or FALSE
' - bCharset = Charset.
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: The handle of the font or NULL on failure.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' Usage examples:
'   hFont = FontNew("MS Sans Serif", 8, FW_NORMAL)
'   hFont = FontNew("Courier New", 10, FW_BOLD)
'   hFont = FontNew("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
' ========================================================================================
PRIVATE FUNCTION CDialog.FontNew (BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, _
   BYVAL lWeight AS LONG = 0, BYVAL bItalic AS UBYTE = FALSE, BYVAL bUnderline AS UBYTE = FALSE, _
   BYVAL bStrikeOut AS UBYTE = FALSE, BYVAL bCharSet AS UBYTE = DEFAULT_CHARSET) AS HFONT

   DIM tlfw AS LOGFONTW
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)

   ' // Font scaling
   DIM DPI AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   IF DPI > 0 THEN lPointSize = (lPointSize * DPI) \ GetDeviceCaps(hDC, LOGPIXELSY)

   tlfw.lfHeight         = -MulDiv(lPointSize, .GetDeviceCaps(hDC, LOGPIXELSY), 72)  ' logical font height
   tlfw.lfWidth          =  0                                                        ' average character width
   tlfw.lfEscapement     =  0                                                        ' escapement
   tlfw.lfOrientation    =  0                                                        ' orientation angles
   tlfw.lfWeight         =  lWeight                                                  ' font weight
   tlfw.lfItalic         =  bItalic                                                  ' italic(CTRUE/FALSE)
   tlfw.lfUnderline      =  bUnderline                                               ' underline(CTRUE/FALSE)
   tlfw.lfStrikeOut      =  bStrikeOut                                               ' strikeout(CTRUE/FALSE)
   tlfw.lfCharSet        =  bCharset                                                 ' character set
   tlfw.lfOutPrecision   =  OUT_TT_PRECIS                                            ' output precision
   tlfw.lfClipPrecision  =  CLIP_DEFAULT_PRECIS                                      ' clipping precision
   tlfw.lfQuality        =  DEFAULT_QUALITY                                          ' output quality
   tlfw.lfPitchAndFamily =  FF_DONTCARE                                              ' pitch and family
   tlfw.lfFaceName       =  wszFaceName                                              ' typeface name

   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = CreateFontIndirectW(@tlfw)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Destroys a font when it is no longer needed.
' ========================================================================================
PRIVATE SUB CDialog.FontEnd (BYVAL hFont AS HFONT)
   DeleteObject hFont
END SUB
' ========================================================================================

' ========================================================================================
' Gets the handle of the font used by the control.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetFont (BYVAL cId AS LONG) AS HFONT
   RETURN CAST(HFONT, SendMessageW(GetDLgItem(m_hDlg, cId), WM_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font that a control is to use when drawing text.
' ========================================================================================
PRIVATE SUB CDialog.ControlSetFont (BYVAL cId AS LONG, BYVAL hFont AS HFONT, BYVAL fRedraw AS BOOLEAN = TRUE)
   DIM hCtl AS HWND = GetDlgItem(m_hDlg, cId)
   DIM hDC AS HDC = GetDC(hCtl)
   DeleteObject(SelectObject(hDC, CAST(HGDIOBJ, hFont)))
   ReleaseDC(hCtl, hDC)
   SendMessageW(hCtl, WM_SETFONT, CAST(WPARAM, hFont), fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the face name of the font used by the control.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetFontFaceName (BYVAL cId AS LONG) AS CWSTR
   ' // Get the handle of the font used by the dialog
   DIM hCtl AS HWND = GetDlgItem(m_hDlg, cId)
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hCtl, WM_GETFONT, 0, 0))
   ' // Get the LOGFONTW structure
   DIM lfw AS LOGFONTW
   IF hCurFont THEN GetObjectW(hCurFont, SIZEOF(lfw), @lfw)
   DIM hDc AS HDC = GetDC(hCtl)
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(hCtl, hDC)
   RETURN lfw.lfFaceName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the point size of the font used by the control.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetFontPointSize (BYVAL cId AS LONG) AS LONG
   ' // Get the handle of the font used by the dialog
   DIM hCtl AS HWND = GetDlgItem(m_hDlg, cId)
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hCtl, WM_GETFONT, 0, 0))
   ' // Get the LOGFONTW structure
   DIM lfw AS LOGFONTW
   IF hCurFont THEN GetObjectW(hCurFont, SIZEOF(lfw), @lfw)
   DIM hDc AS HDC = GetDC(hCtl)
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(hCtl, hDC)
   DIM ptSize AS LONG = MulDiv(lfw.lfHeight, 72, cyPixelsPerInch)
   IF ptSize < 0 THEN ptSize = -ptSize
   RETURN ptSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets/gets the return value of a message processed in the dialog box procedure.
' Return value: TRUE or FALSE.
' Must be set just before exiting the callback message (not after calling other functions)
' See https://forum.powerbasic.com/forum/user-to-user-discussions/programming/19130-info-dwl_msgresult
' Messages that may need the use of this method:
' WM_CTLCOLORMSGBOX, WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX, WM_CTLCOLORBTN, WM_CTLCOLORDLG
' WM_CTLCOLORSCROLLBAR, WM_CTLCOLORSTATIC, WM_COMPAREITEM, WM_VKEYTOITEM, WM_CHARTOITEM
' WM_QUERYDRAGICON, WM_INITDIALOG.
' ========================================================================================
PRIVATE FUNCTION CDialog.CBSetDlgMsgResult(BYVAL result AS LONG_PTR) AS LONG_PTR
   RETURN SetWindowLongPtrW(m_hDlg, DWLP_MSGRESULT, result)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDialog.CBGetDlgMsgResult () AS LONG_PTR
   RETURN GetWindowLongPtrW(m_hDlg, DWLP_MSGRESULT)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Places (posts) a message in the message queue associated with the thread that created the
' specified window and returns without waiting for the thread to process the message.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlPost (BYVAL cId AS LONG, BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
   RETURN PostMessageW(GetDlgItem(m_hDlg, cId), Msg, cast(WPARAM, wPrm), cast(LPARAM, lPrm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sends the specified message to a window or windows. The SendMessage function calls the
' window procedure for the specified window and does not return until the window procedure
' has processed the message.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlSend (BYVAL cId AS LONG, BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
   RETURN SendDlgItemMessageW(m_hDlg, cId, Msg, cast(WPARAM, wPrm), cast(LPARAM, lPrm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the width and height of the control in dialog units.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetSize (BYVAL cId AS LONG, BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
   IF VARPTR(nWidth) = 0 OR VARPTR(nHeight) = 0 THEN RETURN FALSE
   ' // Get the dimensions of the window
   DIM rc AS RECT
   GetWindowRect(GetDlgItem(m_hDlg, cId), @rc)
   ' // Return the width and height in dialog units
   RETURN this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the width of the window in dialog units.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetWidth (BYVAL cId AS LONG) AS LONG
   ' // Get the dimensions of the window
   DIM rc AS RECT
   FUNCTION = GetWindowRect(GetDLgItem(m_hDlg, cId), @rc)
   ' // Return the width in dialog units
   DIM AS LONG nWidth, nHeight
   this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
   RETURN nWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the height of the window in dialog units.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetHeight (BYVAL cId AS LONG) AS LONG
   ' // Get the dimensions of the window
   DIM rc AS RECT
   FUNCTION = GetWindowRect(GetDlgItem(m_hDlg, cId), @rc)
   ' // Return the width in dialog units
   DIM AS LONG nWidth, nHeight
   FUNCTION = this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
   RETURN nHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the size in dialog units of the specified window.
' - nWidth = The new width of the window.
' - nHeight = The new height of the window.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlSetSize (BYVAL cId AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
   this.DialogUnitsToPixels(nWidth, nHeight, nWidth, nHeight)
   RETURN SetWindowPos(GetDlgItem(m_hDlg, cId), NULL, 0, 0, nWidth, nHeight, SWP_NOZORDER OR SWP_NOMOVE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the coordinates of the control's client area. The client coordinates specify the
' upper-left and lower-right corners of the client area. Because client coordinates are
' relative to the upper-left corner of a window's client area, the coordinates of the
' upper-left corner are (0,0).
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetClient (BYVAL cId AS LONG, BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
   IF VARPTR(nWidth) = 0 OR VARPTR(nHeight) = 0 THEN RETURN FALSE
   ' // Get the dimensions of the window
   DIM rc AS RECT
   GetClientRect(GetDlgItem(m_hDlg, cId), @rc)
   ' // Return the width and height values in dialog units
   RETURN this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the width of the the control's client area in dialog units.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetClientWidth (BYVAL cId AS LONG) AS LONG
   ' // Get the dimensions of the window
   DIM rc AS RECT
   FUNCTION = GetClientRect(GetDlgItem(m_hDlg, cId), @rc)
   ' // Return the width in dialog units
   DIM AS LONG nWidth, nHeight
   this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
   RETURN nWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the height of the control's client area in dialog units.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlGetClientHeight (BYVAL cId AS LONG) AS LONG
   ' // Get the dimensions of the window
   DIM rc AS RECT
   FUNCTION = GetClientRect(GetDlgItem(m_hDlg, cId), @rc)
   ' // Return the width in dialog units
   DIM AS LONG nWidth, nHeight
   this.PixelsToDialogUnits(rc.Right - rc.Left, rc.Bottom - rc.Top, nWidth, nHeight)
   RETURN nHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Adjusts the bounding rectangle of the dialog based on the desired size of the client area.
' Parameters:
' - nWidth = The new width of the client area of the dialog.
' - nHeight = The new height of the client area of the dialog.
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlSetClient (BYVAL cId AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
   ' // Convert dialog units to pixels
   this.DialogUnitsToPixels(nWidth, nHeight, nWidth, nHeight)
   RETURN SetWindowPos(GetDlgItem(m_hDlg, cId), NULL, 0, 0, nWidth, nHeight, SWP_NOZORDER OR SWP_NOMOVE)
END FUNCTION
' ========================================================================================

END NAMESPACE
