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
#include once "Afx/AfxWin.inc"
#include once "Afx/AfxDDT.inc"
#include once "win/richedit.bi"
USING Afx

' // Some defines to add PowerBasic compatibility. More to be added.
#define ControlDisable(hDlg, id) EnableWindowGet(DlgItem(hDlg, id), FALSE)
#define ControlEnable(hDlg, id) EnableWindowGet(DlgItem(hDlg, id), TRUE)
#define ControlHandle(hDlg, id) GetDlgItem(hDlg, id)
#define ControlKill(hDlg, id) DestroyWindow(GetDlgItem(hDlg, id))
#define ControlSetFocus(hDlg, id) SetFocus(GetDlgItem(hDlg, id))
#define ControlShowState(hDlg, id, showstate) ShowWindow(GetDlgItem(hDlg, id), showstate)
#define IsWin(hDlg) IsWindow(hDlg)
#define MenuAttach(hMenu, hDlg) SetMenu(hDlg, hMenu)
' Iterate (DO | LOOP | FOR) Continue (DO | FOR | WHILE)
#define _DQ_ CHR(34)

NAMESPACE Afx

' ========================================================================================
' CDialog class
' ========================================================================================
TYPE CDialog

   Private:
      m_hDlg  AS HWND                ' // Dialog handle
      m_hInstance AS HINSTANCE       ' // Instance handle
      m_hRichEditLib AS HMODULE      ' // Rich Edit moudle handle
      m_ptSize AS LONG               ' // Font point size
      m_fontStyle AS BYTE            ' // Font style
      m_charset AS BYTE              ' // Font character set
      m_fontName AS WSTRING * 32     ' // Font face name
      m_IsVisible AS BOOLEAN         ' // The dialog has the WS_VISIBLE style
      m_IsModal AS BOOLEAN           ' // Dialog is modal
      m_DlgRetVal AS INT_PTR         ' // To simulate EndDialog

   Public:
      DECLARE CONSTRUCTOR (BYREF fontName AS WSTRING = "Segoe UI", BYVAL ptSize AS LONG = 9, _
         BYVAL fontStyle AS BYTE = 0, BYVAL charset AS BYTE = DEFAULT_CHARSET) 
      DECLARE DESTRUCTOR 
      DECLARE FUNCTION DialogTemplate (BYREF Title AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
         BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0) AS STRING
      DECLARE FUNCTION DialogNew (BYVAL hParent AS HWND = HWND_DESKTOP, BYREF wszTitle AS WSTRING = "", _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
         BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0) AS HWND
      DECLARE FUNCTION DialogShowModeless (BYVAL hDlg AS HWND, BYVAL pCallback AS ANY PTR = NULL) AS INT_PTR
      DECLARE FUNCTION DialogShowModal (BYVAL hDlg AS HWND, BYVAL pCallback AS ANY PTR = NULL) AS INT_PTR
      DECLARE FUNCTION DialogDoEvents (BYVAL dwMilliseconds AS DWORD = 1) AS BOOLEAN
      DECLARE FUNCTION DialogEnd (BYVAL nResult AS LONG = 0) AS INT_PTR
      DECLARE FUNCTION DialogEndResult () AS INT_PTR
      DECLARE FUNCTION ControlAdd (BYREF wszClassName AS WSTRING, BYVAL hDlg AS HWND = NULL, BYVAL cId AS LONG_PTR = 0, _
         BYREF wszTitle AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
         BYVAL dwStyle AS LONG = -1, BYVAL dwExStyle AS LONG = -1, BYVAL lpPAram AS LONG_PTR = 0, _
         BYVAL pWndProc AS SUBCLASSPROC = NULL, BYVAL uIdSubclass AS UINT_PTR = &HFFFFFFFF, _
         BYVAL dwRefData AS DWORD_PTR = NULL) AS HWND
      DECLARE PROPERTY hDialog () AS HWND
      DECLARE PROPERTY hDialog (BYVAL hDlg AS HWND)
      DECLARE FUNCTION IsModal () AS BOOLEAN
      DECLARE SUB SetDialogFocus (BYVAL hDlg AS HWND, BYVAL hCtl AS HWND)
      DECLARE FUNCTION GetDefId (BYVAL hDlg AS HWND) AS INT_PTR
      DECLARE FUNCTION SetDefId (BYVAL hDlg AS HWND, BYVAL cId AS INT_PTR) AS BOOLEAN

Private:
      DECLARE STATIC FUNCTION CDialogProc (BYVAL hDlg AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' Modal dialog constructor
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
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CDialog
   IF m_hRichEditLib THEN FreeLibrary m_hRichEditLib
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' New dialog
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
' Adds a control to the dialog
' ========================================================================================
PRIVATE FUNCTION CDialog.ControlAdd ( _
   BYREF wszClassName AS WSTRING, _                       ' // Class name
   BYVAL hDlg AS HWND = NULL, _                           ' // Handle of the dialog
   BYVAL cID AS LONG_PTR = 0, _                           ' // Control identifier
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
   IF hDlg = NULL THEN hDlg = m_hDlg
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
         GetClientRect(hDlg, @rcClient)
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
   MapDialogRect hDlg, @rc
   hCtl = CreateWindowExW (dwExStyle, wsClassName, wszTitle, dwStyle, rc.Left, rc.Top, rc.Right - rc.Left, _
          rc.Bottom - rc.Top, hDlg, CAST(HMENU, cID), m_hInstance, CAST(LPVOID, lpParam))
   IF hCtl = NULL THEN EXIT FUNCTION
   ' // Set the font
   IF bSetFont THEN SendMessageW(hCtl, WM_SETFONT, SendMessageW(hDlg, WM_GETFONT, 0, 0), 0)
   SELECT CASE UCASE(wsClassName)
      CASE "LISTBOX"
         ' // Adjust the height of the control so that the integral height
         ' // is based on the new font rather than the default SYSTEM_FONT
         SetWindowPos hCtl, NULL, x, y, nWidth, nHeight, SWP_NOZORDER
      CASE "DATETIMEPICKER", "SYSDATETIMEPICK32"
         ' // Sets the font to be used by the date and time picker control's child month calendar control.
         SendMessageW(hCtl, DTM_SETMCFONT, SendMessageW(hDlg, WM_GETFONT, 0, 0), CTRUE)
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
' Create a modeless dialog in memory
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogNew (BYVAL hParent AS HWND = HWND_DESKTOP, BYREF wszTitle AS WSTRING = "", _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
   BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0) AS HWND

   ' // If it has the WS_VISIBLE style
   IF (dwStyle AND WS_VISIBLE) = WS_VISIBLE THEN
      ' // Remove the WS_VISIBLE style
      m_IsVisible = TRUE
      dwStyle = dwStyle AND (NOT WS_VISIBLE)
   END IF
   ' // Default styles
   IF dwStyle THEN dwStyle = dwStyle OR DS_3DLOOK OR DS_NOFAILCREATE OR DS_SETFONT
   IF dwStyle = 0 THEN dwStyle = DS_3DLOOK OR DS_SETFONT OR DS_MODALFRAME OR DS_NOFAILCREATE OR WS_BORDER OR _
                                 WS_CLIPSIBLINGS OR WS_DLGFRAME OR WS_POPUP
   IF dwExStyle = 0 THEN dwExstyle = WS_EX_LEFT OR WS_EX_LTRREADING OR WS_EX_RIGHTSCROLLBAR
   DIM sTemplate AS STRING = this.DialogTemplate(wszTitle, x, y, cx, cy, dwStyle, dwExStyle)
   m_hDlg = CreateDialogIndirectParamW(m_hInstance, cast(LPCDLGTEMPLATEW, STRPTR(sTemplate)), _
            hParent, NULL, cast(LPARAM, @this))
   RETURN m_hDlg

END FUNCTION
' ========================================================================================
' ========================================================================================
' Default dialog callback procedure
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
' Show the modeless dialog
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogShowModeless (BYVAL hDlg AS HWND, BYVAL pCallback AS ANY PTR = NULL) AS INT_PTR
   m_IsModal = FALSE
   ' // If it had de WS_VISIBLE style...
   IF m_IsVisible THEN
      ' // ... restore it.
      DIM dwStyle AS DWORD = GetWindowLongPtrW(hDlg, GWL_STYLE)
      SetWindowLongPtrW(hDlg, GWL_STYLE, dwStyle OR WS_VISIBLE)
      ' // ... and redraw the dialog
      InvalidateRect hDlg, NULL, CTRUE
      UpdateWindow hDlg
      IF GetMenu(hdlg) THEN DrawMenuBar hDlg
   END IF
   ' // Set the dialog callback
   IF pCallback = NULL THEN pCallback = PROCPTR(CDialog.CDialogProc)
   SetWindowLongPtrW(hDlg, DWLP_DLGPROC, cast(LPARAM, pCallBack))
   ' // Send a WM_INITDIALOG message
   SendMessage(hDlg, WM_INITDIALOG, cast(WPARAM, hDlg), cast(LPARAM, @this))
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Show the modal dialog
' ========================================================================================
PRIVATE FUNCTION CDialog.DialogShowModal (BYVAL hDlg AS HWND, BYVAL pCallback AS ANY PTR = NULL) AS INT_PTR

   ' // Flag the dialog as modal
   m_IsModal = TRUE
   ' // Set the dialog callback
   IF pCallback = NULL THEN pCallback = PROCPTR(CDialog.CDialogProc)
   SetWindowLongPtrW(hDlg, DWLP_DLGPROC, cast(LPARAM, pCallBack))
   ' // If it had de WS_VISIBLE style...
   IF m_IsVisible THEN
      ' // ... restore it.
      DIM dwStyle AS DWORD = GetWindowLongPtrW(hDlg, GWL_STYLE)
      SetWindowLongPtrW(hDlg, GWL_STYLE, dwStyle OR WS_VISIBLE)
      ' // ... and redraw the dialog
      ' SetWindowPos(hDlg, NULL, 0, 0, 0, 0, SWP_SHOWWINDOW OR SWP_NOSIZE OR SWP_NOMOVE OR SWP_NOACTIVATE OR SWP_NOZORDER)
      InvalidateRect hDlg, NULL, FALSE
      UpdateWindow hDlg
      IF GetMenu(hdlg) THEN DrawMenuBar hDlg
   END IF
   ' // Send a WM_INITDIALOG message
   SendMessage(hDlg, WM_INITDIALOG, cast(WPARAM, hDlg), cast(LPARAM, @this))
   ' // Disable the parent window
   DIM hParent AS HWND = GetParent(hDlg)
   IF hParent THEN EnableWindow(hParent, FALSE)
   ' // Message handler loop
   DO
      this.DialogDoEvents
   LOOP WHILE IsWindow(hDlg)
   ' // Enable the parent window
   IF hParent THEN EnableWindow(hParent, FALSE)

   RETURN S_OK

END FUNCTION
' ========================================================================================

' ========================================================================================
' Ends the dialog
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
' Message pump for modeless dialogs
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
' Sets the focus in the specified control of a dialog box.
' See https://devblogs.microsoft.com/oldnewthing/20040802-00/?p=38283
' ========================================================================================
PRIVATE SUB CDialog.SetDialogFocus (BYVAL hDlg AS HWND, BYVAL hCtl AS HWND)
   SendMessageW(hDlg, WM_NEXTDLGCTL, cast(WPARAM, hCtl), CTRUE)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the identifier of the default push button control for a dialog box.
' ========================================================================================
PRIVATE FUNCTION CDialog.GetDefId (BYVAL hDlg AS HWND) AS INT_PTR
   RETURN LOWORD(SendMessageW(hDlg, DM_GETDEFID, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the identifier of the default push button for a dialog box.
' ========================================================================================
PRIVATE FUNCTION CDialog.SetDefId (BYVAL hDlg AS HWND, BYVAL cId AS INT_PTR) AS BOOLEAN
   ' // Removes the default button because 32768 is not a valid control ID.
   PostMessageW(hDlg, DM_SETDEFID, 32768, 0)
   ' // Sets the specified button as the default
   RETURN PostMessageW(hDlg, DM_SETDEFID, cId, 0)
END FUNCTION
' ========================================================================================

END NAMESPACE
