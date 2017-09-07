' ########################################################################################
' Microsoft Windows
' File: FBWinSpy.bas
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca, with the collaboration of Pierre Bellisle.
' Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "FBWSpyProcs.inc"
USING Afx

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Constants
CONST FBWS_CLASSNAME   = "FBWinSpy"
CONST FBWS_PROGNAME    = "FBWinSpy"
CONST FBWS_VERSION     = "1.00 Beta - Build 9"
CONST FBWS_CAPTION     = "FBWinSpy"
CONST FBWS_COPYRIGHT   = "© 2017 by José Roca"
CONST FBWS_MAILADDRESS = "JRoca@com.it-berater.org"
CONST FBWS_INIFILENAME = "FBWinSpy.ini"

ENUM
   IDC_LABEL = 101
   IDC_TAB
   IDC_TREEVIEW
   IDC_VIEWCODE
END ENUM

' // Handles
ENUM FBWS_USERDATA
   FBWS_HTAB, FBWS_HLABEL, FBWS_HTREEVIEW, FBWS_HCODEVIEW, FBWS_HCODEVIEWFONT
END ENUM

' // Globals
DIM SHARED AS HTREEITEM __hRootNode, __hMainWindowNode, __hMainWindowChildrenNode
DIM SHARED AS HTREEITEM __hMenuNode, __hSubNode, __hSubNodeChildren, __hControlSubNodeChildren

' // Forward declarations
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION FBWS_CreateTabControl (BYVAL pWindow AS CWindow PTR) AS HWND
DECLARE FUNCTION FBWS_TabPage1_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION FBWS_TabPage2_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION FBWS_Edit_CodeView_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE SUB FBWS_GetMainWindowInfo (BYVAL hwndMain AS HWND, BYVAL hTreeView AS HWND)
DECLARE SUB FBWS_GetWindowInfo (BYVAL hwnd AS HWND, BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM)
DECLARE SUB FBWS_GetChildWindowsInfo (BYVAL hwnd AS HWND, BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM = NULL, BYVAL bGetWindowInfo AS BOOLEAN = FALSE)
DECLARE FUNCTION FBWS_GetControlsInfo (BYVAL hwnd AS HWND, BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM, BYVAL bExpandNode AS BOOLEAN = FALSE) AS BOOLEAN
DECLARE SUB FBWS_GetControlInfo (BYVAL hwnd AS HWND, BYVAL hTreeView AS HWND, BYVAL hSubNode AS HTREEITEM)
DECLARE SUB FBWS_GetMenu (BYVAL hwndMain AS HWND, BYVAL hTreeView AS HWND)
DECLARE SUB FBWS_GetSubMenuItems (BYVAL hSubMenu AS HMENU, BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM)
DECLARE SUB FBWS_GetControlChildWindowsInfo (BYVAL hwnd AS HWND, BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM = NULL, BYVAL bGetWindowInfo AS BOOLEAN = FALSE)
DECLARE FUNCTION FBWS_ControlEnumChildProc (BYVAL hwnd AS HWND, BYVAL lParam AS LPARAM) AS LONG
DECLARE FUNCTION FBWS_EnumChildProc (BYVAL hwnd AS HWND, BYVAL lParam AS LPARAM) AS LONG

' ========================================================================================
' WindowFromPointEx procedure
' ========================================================================================
Function WindowFromPointEx (BYVAL CursorPosDesktop AS POINT) AS HANDLE
   ' Thanks to J. Brown @ http://read.pudn.com/downloads14/sourcecode/windows/system/55084/winspy+SCR/scr/WindowFromPointEx.c__.htm

   ' - J. Brown says...
   ' - The problem:
   ' WindowFromPoint API is not very good. It cannot cope with odd window arrangements,
   ' i.e. a group-box in a dialog may contain a few check-boxes. These check-boxes are
   ' not children of the groupbox, but are at the same "level" in the window-hierachy.
   ' WindowFromPoint will just return the first available window it finds which encompasses
   ' the mouse (i.e. the group-box), but will NOT be able to detect the contents.
   ' - Solution:
   ' We use WindowFromPoint to start us off, and then step back one level (i.e. from the
   ' parent of what WindowFromPoint returned).
   ' Once we have this window, we enumerate ALL children of this window ourselves, and
   ' find the one that best fits under the mouse - the smallest window that fits, in fact.
   ' I've tested this on alot of dIfferent apps, and it seems to work flawlessly - in fact,
   ' I havn't found a situation yet that this method doesn't work on.....we'll see!

   DIM hFromPoint  AS HANDLE
   DIM hParent     AS HANDLE
   DIM hControlTry AS HANDLE
   DIM CtrlRect    AS RECT
   DIM Area        AS LONG
   DIM AreaPrev    AS LONG

   hFromPoint = WindowFromPoint(CursorPosDesktop)
   FUNCTION   = hFromPoint
   hParent    = GetParent(hFromPoint)
   IF hParent THEN
      ' // Get first child window
      hControlTry = GetWindow(hParent, GW_CHILD)
      ' // Enumerate all child windows
      DO WHILE hControlTry                               ' Enumerate all child including ourselve
         GetWindowRect(hControlTry, @CtrlRect)           ' Get child's rect
         IF PtInRect(@CtrlRect, CursorPosDesktop) THEN   ' Is mouse on control
            Area = (CtrlRect.Right - CtrlRect.Left) * (CtrlRect.Bottom - CtrlRect.Top)   ' Calculate area
            IF AreaPrev THEN           ' This is another control under same mouse position, like FRAME or BUTTON
               IF Area < AreaPrev THEN ' If it's smaller Then we want it
                  IF IsWindowVisible(hControlTry) THEN
                     FUNCTION = hControlTry
                  END IF
               END IF
            END IF
            AreaPrev = Area ' Save value for next loop, if any
         END IF
         hControlTry = GetWindow(hControlTry, GW_HWNDNEXT)
      LOOP
   END IF

END FUNCTION
' ========================================================================================

' ========================================================================================
' Get process DPI awareness (undocumented)
' Untested. Requires Windows 8.1+
' ========================================================================================
FUNCTION GetProcessDpiAwarenessInternal (BYVAL hProcess AS HANDLE, BYVAL pValue AS ULONG PTR) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pGetProcessDpiAwarenessInternal AS FUNCTION (BYVAL hProcess AS HANDLE, BYVAL pValue AS ULONG PTR) AS LONG
   pGetProcessDpiAwarenessInternal = DyLibSymbol(pLib, "GetProcessDpiAwarenessInternal")
   IF pGetProcessDpiAwarenessInternal THEN FUNCTION = pGetProcessDpiAwarenessInternal(hProcess, pValue)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the bounding rectangle of a part in a status window.
' Parameters:
' - hStatusBar = Handle of the status bar control. It can belong to another application.
' - nPart = Zero-based index of the part whose bounding rectangle is to be retrieved.
' Return value: A RECT structure with the bounding rectangle.
' ========================================================================================
PRIVATE FUNCTION FBWS_GetStatusBarRect (BYVAL hStatusBar AS HWND, BYVAL nPart AS LONG) AS RECT
   DIM idProc AS DWORD, hProcess AS HANDLE, pSbRect AS ANY PTR, rc AS RECT, cbBytesRead AS SIZE_T
   GetWindowThreadProcessId(hStatusBar, @idProc)
   IF idProc = 0 THEN EXIT FUNCTION
   hProcess = OpenProcess(PROCESS_VM_OPERATION OR PROCESS_VM_READ OR PROCESS_VM_WRITE OR _
              PROCESS_QUERY_INFORMATION, FALSE, idProc)
   IF hProcess = NULL THEN EXIT FUNCTION
   pSbRect = VirtualAllocEx(hProcess, NULL, SIZEOF(RECT), MEM_COMMIT, PAGE_READWRITE)
   IF pSbRect THEN
      SendMessageW(hStatusBar, SB_GETRECT, nPart, cast(LPARAM, pSbRect))
      ReadProcessMemory(hProcess, pSbRect, @rc, SIZEOF(rc), @cbBytesRead)
      VirtualFreeEx(hProcess, pSbRect, 0, MEM_RELEASE)
   END IF
   CloseHandle(hProcess)
   FUNCTION = rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the bounding rectangle of a part in a status window.
' Parameters:
' - hRebarBar = Handle of the rebar control. It can belong to another application.
' - nBand = Zero-based index of the band for which the borders will be retrieved.
' Return value: A RECT structure with the band border.
'   If the rebar control has the RBS_BANDBORDERS style, each member of this structure will
'   receive the number of pixels, on the corresponding side of the band, that constitute
'   the border. If the rebar control does not have the RBS_BANDBORDERS style, only the left
'   member of this structure receives valid information.
' ========================================================================================
PRIVATE FUNCTION FBWS_GetRebarBandRect (BYVAL hRebar AS HWND, BYVAL nBand AS LONG) AS RECT
   DIM idProc AS DWORD, hProcess AS HANDLE, pRbRect AS ANY PTR, rc AS RECT, cbBytesRead AS SIZE_T
   GetWindowThreadProcessId(hRebar, @idProc)
   IF idProc = 0 THEN EXIT FUNCTION
   hProcess = OpenProcess(PROCESS_VM_OPERATION OR PROCESS_VM_READ OR PROCESS_VM_WRITE OR _
              PROCESS_QUERY_INFORMATION, FALSE, idProc)
   IF hProcess = NULL THEN EXIT FUNCTION
   pRbRect = VirtualAllocEx(hProcess, NULL, SIZEOF(RECT), MEM_COMMIT, PAGE_READWRITE)
   IF pRbRect THEN
      SendMessageW(hRebar, RB_GETBANDBORDERS, nBand, cast(LPARAM, pRbRect))
      ReadProcessMemory(hProcess, pRbRect, @rc, SIZEOF(rc), @cbBytesRead)
      VirtualFreeEx(hProcess, pRbRect, 0, MEM_RELEASE)
   END IF
   CloseHandle(hProcess)
   FUNCTION = rc
END FUNCTION
' ========================================================================================

' // Must use a buffer of 32 bytes because of the differences
' // between 32 and 64 bit for the TBBUTTON structure
type FBWS_TBBUTTON
	iBitmap AS LONG
	idCommand AS LONG
	fsState AS UBYTE
	fsStyle AS UBYTE
   filler (0 TO 19) AS BYTE
END TYPE

' ========================================================================================
' Retrieves information about the toolbar buttons
' ========================================================================================
PRIVATE FUNCTION FBWS_GetToolbarButton (BYVAL hToolBar AS HWND, BYVAL nButton AS LONG) AS FBWS_TBBUTTON
   DIM idProc AS DWORD, hProcess AS HANDLE, ptbbutton AS ANY PTR, cbBytesRead AS SIZE_T
   DIM afx_tbb AS FBWS_TBBUTTON
   GetWindowThreadProcessId(hToolBar, @idProc)
   IF idProc = 0 THEN EXIT FUNCTION
   hProcess = OpenProcess(PROCESS_VM_OPERATION OR PROCESS_VM_READ OR PROCESS_VM_WRITE OR _
              PROCESS_QUERY_INFORMATION, FALSE, idProc)
   IF hProcess = NULL THEN EXIT FUNCTION
   ptbbutton = VirtualAllocEx(hProcess, NULL, SIZEOF(FBWS_TBBUTTON), MEM_COMMIT, PAGE_READWRITE)
   IF ptbbutton THEN
      IF SendMessageW(hToolBar, TB_GETBUTTON, CAST(WPARAM, nButton), CAST(LPARAM, ptbbutton)) THEN
         ReadProcessMemory(hProcess, ptbbutton, @afx_tbb, SIZEOF(afx_tbb), @cbBytesRead)
      END IF
      VirtualFreeEx(hProcess, ptbbutton, 0, MEM_RELEASE)
   END IF
   CloseHandle(hProcess)
   FUNCTION = afx_tbb
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the specified band of the Rebar control
' ========================================================================================
PRIVATE FUNCTION FBWS_GetRebarBandInfo (BYVAL hRebar AS HWND, BYVAL nBand AS LONG) AS REBARBANDINFOW
   DIM idProc AS DWORD, hProcess AS HANDLE, pInfo AS ANY PTR, cbBytesRead AS SIZE_T
   DIM rbbi AS REBARBANDINFOW
   IF AfxWindowsVersion >= 600 AND AfxComCtlVersion >= 600 THEN
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   ELSE
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   END IF
   rbbi.fMask = RBBIM_BACKGROUND OR RBBIM_CHILD OR RBBIM_CHILDSIZE OR RBBIM_COLORS OR RBBIM_ID OR _
                RBBIM_HEADERSIZE OR RBBIM_IDEALSIZE OR RBBIM_IMAGE OR RBBIM_SIZE OR RBBIM_STYLE OR RBBIM_TEXT
   GetWindowThreadProcessId(hRebar, @idProc)
   IF idProc = 0 THEN EXIT FUNCTION
   hProcess = OpenProcess(PROCESS_VM_OPERATION OR PROCESS_VM_READ OR PROCESS_VM_WRITE OR _
              PROCESS_QUERY_INFORMATION, FALSE, idProc)
   IF hProcess = NULL THEN EXIT FUNCTION
   pInfo = VirtualAllocEx(hProcess, NULL, SIZEOF(REBARBANDINFOW), MEM_COMMIT, PAGE_READWRITE)
   IF pInfo THEN
      WriteProcessMemory(hProcess, pInfo, @rbbi, SIZEOF(REBARBANDINFOW), NULL)
      IF SendMessageW(hRebar, RB_GETBANDINFOW, CAST(WPARAM, nBand), CAST(LPARAM, pInfo)) THEN
         ReadProcessMemory(hProcess, pInfo, @rbbi, SIZEOF(rbbi), @cbBytesRead)
      END IF
      VirtualFreeEx(hProcess, pInfo, 0, MEM_RELEASE)
   END IF
   CloseHandle(hProcess)
   FUNCTION = rbbi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest file
   AfxSetProcessDPIAware

   ' // Creates the main window
   DIM pWindow AS CWindow = FBWS_CLASSNAME
   pWindow.Create(NULL, FBWS_CAPTION & " - " & FBWS_VERSION, @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 410)
   ' // Centers the window
   pWindow.Center
   ' // Sets the icons
   pWindow.BigIcon   = LoadImage(hInstance, MAKEINTRESOURCE(101), IMAGE_ICON, 48, 48, LR_SHARED)
   pWindow.SmallIcon = LoadImage(hInstance, MAKEINTRESOURCE(100), IMAGE_ICON, 32, 32, LR_SHARED)

   ' // Adds a label
   DIM wszLabelText AS WSTRING * 260 = "Click here to start" & CHR(13, 10) & _
       "Move cursor to target" & CHR(13, 10) & "Release mouse on target"
   DIM hLabel AS HWND = pWindow.AddControl("Label", , IDC_LABEL, wszLabelText, 10, 10, 150, 55, _
       WS_CHILD OR WS_VISIBLE OR SS_NOTIFY OR SS_CENTER)
   pWindow.UserData(FBWS_HLABEL) = cast(LONG_PTR, hLabel)

   ' // Adds a button
   DIM hb AS HWND = pWindow.AddControl("Button", , IDCANCEL, "&Close")

   ' // Create the Tab control
   FBWS_CreateTabControl(@pWindow)


   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC bStartDrag AS BOOLEAN
   DIM pt AS POINT, _hwnd AS HWND, cID AS LONG, wszClassName AS WSTRING * 260

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDC_LABEL
               IF GET_WM_COMMAND_CMD(wParam, lParam) = STN_CLICKED THEN
                  SetCapture hwnd
                  bStartDrag = TRUE
                  SetCursor(LoadCursor(NULL, IDC_SIZEALL))
                  SendMessageW hwnd, WM_MOUSEMOVE, 0, 0   ' // Trigger an initial mouse move
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_NOTIFY

         DIM nPage AS DWORD              ' // Page number
         DIM pTabPage AS CTabPage PTR    ' // Tab page object reference
         DIM tci AS TCITEMW              ' // TCITEMW structure

         DIM ptnmhdr AS NMHDR PTR   ' // Information about a notification message
         ptnmhdr = CAST(NMHDR PTR, lParam)
         SELECT CASE ptnmhdr->idFrom
            CASE IDC_TAB
               SELECT CASE ptnmhdr->code
                  CASE TCN_SELCHANGE
                     ' // Show the selected page
                     pTabPage = AfxCTabPagePtr(ptnmhdr->hwndFrom, -1)
                     IF pTabPage THEN
                        ShowWindow pTabPage->hTabPage, SW_SHOW
                        SetFocus pTabPage->hWindow
                     END IF
                  CASE TCN_SELCHANGING
                     ' // Hide the current page
                     pTabPage = AfxCTabPagePtr(ptnmhdr->hwndFrom, -1)
                     IF pTabPage THEN ShowWindow pTabPage->hTabPage, SW_HIDE
               END SELECT
         END SELECT

      CASE WM_LBUTTONUP
         IF bStartDrag = FALSE THEN EXIT FUNCTION
         GetCursorPos(@pt)
         ReleaseCapture
         SetCursor(LoadCursor(NULL, IDC_ARROW))
         bStartDrag = FALSE
         ' // Retrieve the handle of main window
'         DIM hwndFromPoint AS HWND = WindowFromPoint(pt)
         DIM hwndFromPoint AS HWND = WindowFromPointEx(pt)
         DIM hwndMain AS HWND = GetAncestor(hwndFromPoint, GA_ROOTOWNER)
         ' // Treeview handle
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
         DIM hTreeView AS HWND = cast(HWND, pWindow->UserData(FBWS_HTREEVIEW))
         IF hTreeView = NULL THEN EXIT FUNCTION
         ' // Delete all the items in the tree view
         TreeView_DeleteAllItems(hTreeView)
         ' // Retrieve the path of the application that created the window
         DIM wszExePath AS WSTRING * 260
         wszExePath = AfxGetPathFromWindowHandle(hwndMain)
         IF LEN(wszExePath) THEN
            __hRootNode = TreeView_AddItem(hTreeView, 0, TVI_ROOT, AfxGetFileName(wszExePath))
            TreeView_AddItem hTreeView, __hRootNode, NULL, "Path: " & wszExePath
            ' // Retrieve the binary type
            DIM nType AS DWORD
            IF GetBinaryTypeW(@wszExePath, @nType) THEN
               DIM wszType AS WSTRING * 260
               SELECT CASE nType
                  CASE SCS_32BIT_BINARY : wszType = "Windows 32 bit"
                  CASE SCS_64BIT_BINARY : wszType = "Windows 64 bit"
               END SELECT
            TreeView_AddItem hTreeView, __hRootNode, NULL, "Binary type: " & wszType
            END IF
         END IF
         ' // Expands the node
         TreeView_Expand(hTreeView, __hRootNode, TVE_EXPAND)
         ' // Retrieves info about the selected window
         DIM hwndForm AS HWND = AfxGetFormHandle(hwndFromPoint)
         IF hwndFromPoint = hwndMain THEN
            ' // Retrieve the menu
            FBWS_GetMenu(hwndMain, hTreeView)
            ' // Main window node
            __hMainWindowNode = TreeView_AddItem(hTreeView, TVI_ROOT, NULL, "Main window " & _
               IIF(IsWindowUnicode(hwndMain), "(Unicode)", "(Ansi)"))
            ' // Retrieve information about the main window
            FBWS_GetMainWindowInfo(hwndMain, hTreeView)
             ' // Expands the node
            TreeView_Expand(hTreeView, __hMainWindowNode, TVE_EXPAND)
         ELSEIF hwndFromPoint = hwndForm THEN
            ' // Secondary window (popup dialog)
            FBWS_GetChildWindowsInfo(hwndForm, hTreeView, NULL, TRUE)
            TreeView_Expand(hTreeView, __hSubNode, TVE_EXPAND)
         ELSE
            ' // Window control (with or without children)
            FBWS_GetControlsInfo(hwndFromPoint, hTreeView, __hSubNode, TRUE)
         END IF
         ' // Select the root node
         IF __hRootNode THEN TreeView_SelectSetFirstVisible(hTreeView, __hRootNode)
         IF __hRootNode = NULL THEN
            IF __hMainWindowNode THEN TreeView_SelectSetFirstVisible(hTreeView, __hMainWindowNode)
         END IF
         EXIT FUNCTION

      CASE WM_GETMINMAXINFO
         ' // Set the pointer to the address of the MINMAXINFO structure
         DIM ptmmi AS MINMAXINFO PTR = CAST(MINMAXINFO PTR, lParam)
         ' // Set the minimum and maximum sizes that can be produced by dragging the borders of the window
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
         IF pWindow THEN
            ptmmi->ptMinTrackSize.x = 250 * pWindow->rxRatio
            ptmmi->ptMinTrackSize.y = 250 * pWindow->ryRatio
         END IF
         EXIT FUNCTION

      CASE WM_SIZE
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            DIM hTab AS HWND = GetDlgItem(hwnd, IDC_TAB)
            ' // Resize the controls
            IF pWindow THEN
               DIM hLabel AS HWND = GetDlgItem(hwnd, IDC_LABEL)
               DIM LabelHeight AS LONG = AfxGetWindowHeight(hLabel) / pWindow->rxRatio
               pWindow->MoveWindow(hTab, 10, LabelHeight + 20, pWindow->ClientWidth - 20, _
                                   pWindow->ClientHeight - LabelHeight - 70, CTRUE)
               DIM hButton AS HWND = GetDlgItem(hwnd, IDCANCEL)
               pWindow->MoveWindow(hButton, pWindow->ClientWidth - 86, pWindow->ClientHeight - 35, 75, 23, CTRUE)
            END IF
            ' // Resize the tab pages
            AfxResizeTabPages(hTab)
         END IF

      CASE WM_CTLCOLORSTATIC
         IF GetDlgCtrlID(CAST(HWND, lParam)) = IDC_LABEL THEN
            ' // Set the background and text colors
            SetBkColor CAST(HDC, wParam), GetSysColor(COLOR_INFOBK)
            SetTextColor CAST(HDC, wParam), GetSysCOlor(COLOR_INFOTEXT)
            ' // Return handle of brush used to paint background
            FUNCTION = CAST(LRESULT, GetSysColorBrush(COLOR_INFOBK))
            EXIT FUNCTION
         END IF

    	CASE WM_DESTROY
         ' // Cleanup
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
         DIM hFont AS HFONT
         IF pWindow THEN hFont = cast(.HFONT, pWindow->UserData(FBWS_HCODEVIEWFONT))
         IF hFont THEN DeleteObject(hFont)
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ============================================================================

' ========================================================================================
' Retrieve the submenu items
' ========================================================================================
SUB FBWS_GetSubMenuItems (BYVAL hSubMenu AS HMENU, BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM)

   DIM nItems AS LONG = GetMenuItemCount(hSubMenu)
   FOR x AS LONG = 0 TO nItems - 1
      DIM mii AS MENUITEMINFOW
      mii.cbSize = SIZEOF(mii)
      mii.fMask = MIIM_CHECKMARKS OR MIIM_DATA OR MIIM_ID OR MIIM_STATE OR MIIM_SUBMENU OR MIIM_TYPE
      DIM wszText AS WSTRING * MAX_PATH
      mii.cch = SIZEOF(wszText)
      mii.dwTypeData = @wszText
      IF GetMenuItemInfoW(hSubMenu, x, 1, @mii) THEN
         IF mii.hSubMenu = NULL THEN
            DIM wszType AS WSTRING * 260 = FBWS_GetMenuItemType(mii.fType)
            DIM wszState AS WSTRING * 260 = FBWS_GetMenuItemState(mii.fState)
            IF wszState = "" THEN wszState = "MF_ENABLED"
            wszText = AfxStrReplace(wszText, CHR(9), " <TAB> ")
            IF LEN(wszText) = 0 AND wszType = "MF_SEPARATOR" THEN wszText = "(Separator)"
            DIM hSubNode AS HTREEITEM = TreeView_AddItem(hTreeView, hNode, NULL, wszText)
            IF LEN(wszType) THEN TreeView_AddItem(hTreeView, hSubNode, NULL, "Type: " & wszType)
            IF wszState <> "MF_SEPARATOR" THEN TreeView_AddItem(hTreeView, hSubNode, NULL, "State: " & wszState)
            ' // Expands the node
'            TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
         ELSE
            ' // A submenu with a nested menu
            DIM wszState AS WSTRING * 260 = FBWS_GetMenuItemState(mii.fState)
            IF wszState = "" THEN wszState = "MF_ENABLED"
            DIM hSubNode AS HTREEITEM = TreeView_AddItem(hTreeView, hNode, NULL, wszText)
            TreeView_AddItem(hTreeView, hSubNode, NULL, "State: " & wszState)
            ' // Recurse with the new submenu handle
            FBWS_GetSubMenuItems (mii.hSubMenu, hTreeView, hSubNode)
         END IF
         ' // Expands the node
         TreeView_Expand(hTreeView, hNode, TVE_EXPAND)
      END IF
   NEXT

END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the menu
' ========================================================================================
SUB FBWS_GetMenu (BYVAL hwndMain AS HWND, BYVAL hTreeView AS HWND)

   DIM hMenu AS HMENU = GetMenu(hwndMain)
   IF hMenu THEN
      DIM hMenuNode AS HTREEITEM = TreeView_AddItem(hTreeView, TVI_ROOT, NULL, "Menu")
      DIM nPos AS LONG
      DO
         DIM hSubMenu AS HMENU = GetSubMenu(hMenu, nPos)
         IF hSubMenu = NULL THEN EXIT DO
         IF hSubMenu <> GetSystemMenu(hwndMain, 0) THEN
            DIM mii AS MENUITEMINFOW
            mii.cbSize = SIZEOF(mii)
            mii.fMask = MIIM_CHECKMARKS OR MIIM_DATA OR MIIM_ID OR MIIM_STATE OR MIIM_SUBMENU OR MIIM_TYPE
            DIM wszText AS WSTRING * MAX_PATH
            mii.cch = SIZEOF(wszText)
            mii.dwTypeData = @wszText
            IF GetMenuItemInfoW(hMenu, nPos, 1, @mii) THEN
               IF LEN(wszText) THEN
                  DIM hMenuSubNode AS HTREEITEM = TreeView_AddItem(hTreeView, hMenuNode, NULL, wszText)
                  TreeView_AddItem(hTreeView, hMenuSubNode, NULL, "State: " & FBWS_GetMenuItemState(mii.fState))
                  DIM nItems AS LONG = GetMenuItemCount(hSubMenu)
                  DIM hNode AS HTREEITEM = TreeView_AddItem(hTreeView, hMenuSubNode, NULL, "Items: " & WSTR(nItems))
                  IF nItems THEN
                     FBWS_GetSubMenuItems (hSubMenu, hTreeView, hNode)
                  END IF
                  ' // Expands the node
                  TreeView_Expand(hTreeView, hMenuSubNode, TVE_EXPAND)
               END IF
            END IF
         END IF
         nPos += 1
      LOOP
      ' // Expands the node
'      TreeView_Expand(hTreeView, hMenuNode, TVE_EXPAND)
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' Creates the tab control
' ========================================================================================
FUNCTION FBWS_CreateTabControl (BYVAL pWindow AS CWindow PTR) AS HWND

   IF pWindow = NULL THEN EXIT FUNCTION
   ' // Add a tab control
   DIM hTab AS .HWND = pWindow->AddControl("Tab", , IDC_TAB, "", 0, 0, 0, 0, _
      WS_CHILD OR WS_VISIBLE OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR WS_GROUP OR WS_TABSTOP OR _
      TCS_FORCELABELLEFT OR TCS_TABS OR TCS_FIXEDWIDTH OR TCS_TOOLTIPS)
   pWindow->UserData(FBWS_HTAB) = cast(LONG_PTR, hTab)

   DIM dwExStyle AS DWORD
   IF AfxWindowsVersion > 500 THEN dwExStyle = WS_EX_COMPOSITED   ' Avoids flicker - Must be Windows XP or superior
   ' // Create the first tab page
   DIM pTabPage1 AS CTabPage PTR = NEW CTabPage
   pTabPage1->DPI = pWindow->DPI
   pTabPage1->InsertPage(hTab, 0, "Windows", , @FBWS_TabPage1_WndProc, , dwExStyle)
   ' // Create the second tab page
   DIM pTabPage2 AS CTabPage PTR = NEW CTabPage
   pTabPage2->DPI = pWindow->DPI
   pTabPage2->InsertPage(hTab, 1, "Code", , @FBWS_TabPage2_WndProc, , dwExStyle)

   ' // Display the first tab page
   ShowWindow pTabPage1->hTabPage, SW_SHOW
   ' // Set the focus to the first tab
   SendMessageW hTab, TCM_SETCURFOCUS, 0, 0

   FUNCTION = hTab

END FUNCTION
' ========================================================================================

' ========================================================================================
' Tab page 1 window procedure
' ========================================================================================
FUNCTION FBWS_TabPage1_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_CREATE
         ' // Get a pointer to the CWindow class
         DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
         ' // Get a pointer to the TabPage class
         DIM pTabPage AS CTabPage PTR = AfxCTabPagePtr(GetParent(hwnd), 0)
         ' // Add a tree view to the page
         DIM dwExStyle AS DWORD
         IF AfxWindowsVersion > 500 THEN dwExStyle = WS_EX_COMPOSITED   ' Avoids flicker - Must be Windows XP or superior
         DIM hTreeView AS HWND = pTabPage->AddControl("TreeView", hwnd, IDC_TREEVIEW, "", _
            0, 0, pTabPage->ControlClientWidth(hwnd), _
            pTabPage->ControlClientHeight(hwnd), , dwExStyle)
         pWindow->UserData(FBWS_HTREEVIEW) = cast(LONG_PTR, hTreeView)
         EXIT FUNCTION

      CASE WM_SIZE
         ' // If the window isn't minimized, resize it
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Get a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
            DIM rc AS RECT = AfxGetWindowClientRect(hwnd)
            DIM hTreeView AS HWND = cast(HWND, pWindow->UserData(FBWS_HTREEVIEW))
            IF hTreeView THEN MoveWindow hTreeView, 0, 0, rc.Right, rc.Bottom - rc.Top, CTRUE
         END IF

      CASE WM_SYSCOLORCHANGE
         ' // Forward this message to common controls so that they will
         ' // be properly updated when the user changes the color settings.
         DIM hTreeView AS HWND = GetDlgItem(hwnd, IDC_TREEVIEW)
         SendMessageW hTreeView, WM_SYSCOLORCHANGE, wParam, lParam

      CASE WM_SETFOCUS
         ' // Get a pointer to the CWindow class
         DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
         DIM hTreeView AS HWND = cast(HWND, pWindow->UserData(FBWS_HTREEVIEW))
         SetFocus hTreeView
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Tab page 2 window procedure
' ========================================================================================
FUNCTION FBWS_TabPage2_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM hBrush AS HBRUSH, rc AS RECT, tlb AS LOGBRUSH

   SELECT CASE uMsg

      CASE WM_CREATE

         ' // Get a pointer to the CWindow class
         DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
         ' // Get a pointer to the TabPage class
         DIM pTabPage AS CTabPage PTR = AfxCTabPagePtr(GetParent(hwnd), 1)
         ' // Add a subclassed edit control to the page
         DIM hCodeView AS HWND = pTabPage->AddControl("Edit", hwnd, IDC_VIEWCODE, "", _
            0, 0, pTabPage->ControlClientWidth(hwnd), pTabPage->ControlClientHeight(hwnd), _
            WS_CHILD OR WS_VISIBLE OR WS_HSCROLL OR WS_VSCROLL OR WS_BORDER OR WS_TABSTOP OR ES_LEFT OR _
            ES_AUTOHSCROLL OR ES_AUTOVSCROLL OR ES_MULTILINE OR ES_NOHIDESEL OR ES_WANTRETURN OR ES_SAVESEL, _
            WS_EX_CLIENTEDGE, , @FBWS_Edit_CodeView_WndProc)
         pWindow->UserData(FBWS_HCODEVIEW) = cast(LONG_PTR, hCodeView)
         Edit_LimitText(hCodeView, 31457279)
         ' // Create a font for the Code view edit control
         DIM hFont AS HFONT
         hFont = pTabPage->CreateFont("Lucida Console", 10, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET)
         ' // Set the font
         AfxSetWindowFont hCodeView, hFont
         pWindow->UserData(FBWS_HCODEVIEWFONT) = cast(LONG_PTR, hFont)
         EXIT FUNCTION

      CASE WM_SIZE
         ' // If the window isn't minimized, resize it
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Get a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
            DIM rc AS RECT = AfxGetWindowClientRect(hwnd)
            DIM hCodeView AS HWND = cast(HWND, pWindow->UserData(FBWS_HCODEVIEW))
            IF hCodeView THEN MoveWindow hCodeView, rc.Left, rc.Top, rc.Right, rc.Bottom - rc.Top, CTRUE
         END IF

      CASE WM_SETFOCUS
         ' // Get a pointer to the CWindow class
         DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
         DIM hCodeView AS HWND = cast(HWND, pWindow->UserData(FBWS_HCODEVIEW))
         SetFocus hCodeView
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Code text box window procedure
' ========================================================================================
FUNCTION FBWS_Edit_CodeView_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_KEYDOWN
'         ' // If Ctrl+A or Ctrl+E pressed, select all the text
'         IF wParam = VK_A OR wParam = VK_E THEN
'            IF GetAsyncKeyState(VK_CONTROL) THEN
'               PostMessage hwnd, EM_SETSEL, 0, -1
'            END IF
'         END IF
'         ' // Eat the Escape key to avoid the page being destroyed
         IF wParam = VK_ESCAPE THEN EXIT FUNCTION

      CASE WM_CHAR
         SELECT CASE wParam
            CASE 1   ' // Ctrl+A
               SendMessage(hWnd, EM_SETSEL, 0, - 1) 'Select everything
               EXIT FUNCTION
         END SELECT

      CASE WM_DESTROY
         ' // REQUIRED: Remove control subclassing
         SetWindowLongPtrW hwnd, GWLP_WNDPROC, CAST(LONG_PTR, RemovePropW(hwnd, "OLDWNDPROC"))

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = CallWindowProcW(GetPropW(hwnd, "OLDWNDPROC"), hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information of child controls
' ========================================================================================
SUB FBWS_GetControlInfo (BYVAL hwnd AS HWND, BYVAL hTreeView AS HWND, BYVAL hSubNode AS HTREEITEM)

   IF hTreeView = NULL THEN RETURN
   IF hSubNode = NULL THEN RETURN
   DIM wszClassName AS WSTRING * 260 = AfxGetWindowClassName(hwnd)

   TreeView_AddItem(hTreeView, hSubNode, NULL, "Handle: &h" & HEX(hwnd, 8))
   TreeView_AddItem(hTreeView, hSubNode, NULL, "ID: " & WSTR(GetDlgCtrlID(hwnd)))
   IF UCASE(wszClassName) = "STATIC" OR UCASE(wszClassName) = "BUTTON" OR wszClassName = "#32770" THEN
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Caption: " & AfxGetWindowText(hwnd))
   END IF
   TreeView_AddItem(hTreeView, hSubNode, NULL, "Width: " & WSTR(AfxGetWindowWidth(hwnd)))
   TreeView_AddItem(hTreeView, hSubNode, NULL, "Height: " & WSTR(AfxGetWindowHeight(hwnd)))
   DIM rc AS RECT = AfxGetWindowRect(hwnd)
   TreeView_AddItem(hTreeView, hSubNode, NULL, "Window rect: " & _
      "(" & WSTR(rc.Left) & ", " & WSTR(rc.Top) & ") - (" & WSTR(rc.Right) & ", " & _
      WSTR(rc.Bottom) & ") - " & WSTR(rc.Right - rc.Left) & "x" & WSTR(rc.Bottom - rc.Top))
   DIM rcClient AS RECT = AfxGetWindowClientRect(hwnd)
   TreeView_AddItem hTreeView, hSubNode, NULL, "Client rect: " & _
      "(" & WSTR(rcClient.Left) & ", " & WSTR(rcClient.Top) & ") - (" & WSTR(rcClient.Right) & ", " & _
      WSTR(rcClient.Bottom) & ") - " & WSTR(rcClient.Right - rcClient.Left) & "x" & WSTR(rcClient.Bottom - rcClient.Top)
   TreeView_AddItem(hTreeView, hSubNode, NULL, "Window styles: &h" & HEX(AfxGetWindowStyle(hwnd), 8))
   TreeView_AddItem(hTreeView, hSubNode, NULL, "Verbose window styles: " & FBWS_GetWindowStyles(hwnd, AfxGetWindowStyle(hwnd)))
   IF AfxGetWindowExStyle(hwnd) THEN
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Window extended styles: &h" & HEX(AfxGetWindowExStyle(hwnd), 8))
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Window verbose extended styles: " & FBWS_GetWindowExStyles(AfxGetWindowExStyle(hwnd)))
   END IF
   SELECT CASE UCASE(wszClassName)
      CASE "COMBOBOXEX32"
         IF ComboBoxEx_GetExtendedStyle(hwnd) THEN
            TreeView_AddItem(hTreeView, hSubNode, NULL, "Control extended styles: &h" & HEX(ComboBoxEx_GetExtendedStyle(hwnd), 8))
            TreeView_AddItem(hTreeView, hSubNode, NULL, "Control verbose extended styles: " & FBWS_GetComboBoxExExStyles(ComboBoxEx_GetExtendedStyle(hwnd)))
         END IF
      CASE "TOOLBARWINDOW32"
         IF Toolbar_GetExtendedStyle(hwnd) THEN
            TreeView_AddItem(hTreeView, hSubNode, NULL, "Control extended styles: &h" & HEX(Toolbar_GetExtendedStyle(hwnd), 8))
            TreeView_AddItem(hTreeView, hSubNode, NULL, "Control verbose extended styles: " & FBWS_GetToolbarExStyles(Toolbar_GetExtendedStyle(hwnd)))
         END IF
      CASE "SYSTABCONTROL32"
         IF TabCtrl_GetExtendedStyle(hwnd) THEN
            TreeView_AddItem(hTreeView, hSubNode, NULL, "Control extended styles: &h" & HEX(TabCtrl_GetExtendedStyle(hwnd), 8))
            TreeView_AddItem(hTreeView, hSubNode, NULL, "Control verbose extended styles: " & FBWS_GetSysTabExStyles(TabCtrl_GetExtendedStyle(hwnd)))
         END IF
      CASE "SYSTREEVIEW32"
         IF TreeView_GetExtendedStyle(hwnd) THEN
            TreeView_AddItem(hTreeView, hSubNode, NULL, "Control extended styles: &h" & HEX(TreeView_GetExtendedStyle(hwnd), 8))
            TreeView_AddItem(hTreeView, hSubNode, NULL, "Control verbose extended styles: " & FBWS_GetTreeViewExStyles(TreeView_GetExtendedStyle(hwnd)))
         END IF
      CASE "SYSLISTVIEW32"
         IF ListView_GetExtendedListViewStyle(hwnd) THEN
            TreeView_AddItem(hTreeView, hSubNode, NULL, "Control extended styles: &h" & HEX(ListView_GetExtendedListViewStyle(hwnd), 8))
            TreeView_AddItem(hTreeView, hSubNode, NULL, "Control verbose extended styles: " & FBWS_GetListViewExStyles(ListView_GetExtendedListViewStyle(hwnd)))
         END IF
      CASE "REBARWINDOW32"
         ' This control has not extended styles yet
   END SELECT

   TreeView_AddItem(hTreeView, hSubNode, NULL, "Parent handle: &h" & HEX(GetParent(hwnd), 8))
   DIM wsz AS WSTRING * MAX_PATH = AfxGetWindowClassName(GetParent(hwnd))
   TreeView_AddItem(hTreeView, hSubNode, NULL, "Parent class name: " & wsz)
   DIM lfw AS LOGFONTW
   IF FBWS_GetFont(hwnd, @lfw) THEN
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Font mame: " & lfw.lfFaceName)
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Font point size: " & AfxGetFontPointSize(lfw.lfHeight))
'      TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
   END IF

   IF UCASE(wszClassName) = "TOOLBARWINDOW32" THEN
      DIM nButtons AS LONG = Toolbar_ButtonCount(hwnd)
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Buttons: " & WSTR(nButtons))
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Button width: " & WSTR(Toolbar_GetButtonWidth(hwnd)))
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Button height: " & WSTR(Toolbar_GetButtonHeight(hwnd)))
      DIM tbb AS FBWS_TBBUTTON
      FOR i AS LONG = 0 TO nButtons - 1
         tbb = FBWS_GetToolBarButton(hwnd, i)
         DIM hSubNode2 AS HTREEITEM = TreeView_AddItem(hTreeView, hSubNode, NULL, "Button " & WSTR(i))
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "idCommand = " & WSTR(tbb.idCommand))
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "iBitmap = " & WSTR(tbb.iBitmap))
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "fsState = " & WSTR(tbb.fsState))
         IF tbb.fsState THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Verbose state flags = " & FBWS_GetToolBarButtonState(tbb.fsState))
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "fsStyle = " & WSTR(tbb.fsStyle))
         IF tbb.fsStyle THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Verbose styles = " & FBWS_GetToolbarStyles(tbb.fsStyle))
      NEXT
   END IF

   IF UCASE(wszClassName) = "MSCTLS_STATUSBAR32" THEN
      DIM nParts AS LONG = StatusBar_GetPartsCount(hwnd)
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Parts: " & WSTR(nParts))
      DIM rc AS RECT
      FOR i AS LONG = 0 TO nParts - 1
         rc = FBWS_GetStatusBarRect(hwnd, i)
         DIM hSubNode2 AS HTREEITEM = TreeView_AddItem(hTreeView, hSubNode, NULL, "Part " & WSTR(i))
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "Left = " & WSTR(rc.Left))
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "Right = " & WSTR(rc.Right))
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "Top = " & WSTR(rc.Top))
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "Bottom = " & WSTR(rc.Bottom))
      NEXT
   END IF

   IF UCASE(wszClassName) = "MSCTLS_PROGRESS32" THEN
      TreeView_AddItem(hTreeView, hSubNode, NULL, "High limit: " & WSTR(ProgressBar_GetHighLimit(hwnd)))
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Low limit: " & WSTR(ProgressBar_GetLowLimit(hwnd)))
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Bar color: &h" & HEX(ProgressBar_GetBarColor(hwnd), 8))
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Background color: &h" & HEX(ProgressBar_GetBkColor(hwnd), 8))
   END IF

   IF UCASE(wszClassName) = "REBARWINDOW32" THEN
      DIM hSubNode2 AS HTREEITEM, hSubNode3 AS HTREEITEM
      TreeView_AddItem(hTreeView, hSubNode, NULL, "Bar height = " & WSTR(Rebar_GetBarHeight(hwnd)))
      hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, "Row count = " & WSTR(Rebar_GetRowCount(hwnd)))
      FOR i AS LONG = 0 TO Rebar_GetRowCount(hwnd) - 1
         hSubNode3 = TreeView_AddItem(hTreeView, hSubNode2, NULL, "Row " & WSTR(i))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "Height = " & WSTR(Rebar_GetRowHeight(hwnd, i)))
      NEXT
      hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, "Band count = " & WSTR(Rebar_GetBandCount(hwnd)))
      FOR i AS LONG = 0 TO Rebar_GetBandCount(hwnd) - 1
         hSubNode3 = TreeView_AddItem(hTreeView, hSubNode2, NULL, "Band " & WSTR(i))
         DIM rc AS RECT = FBWS_GetRebarBandRect(hwnd, i)
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "Left = " & WSTR(rc.Left))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "Right = " & WSTR(rc.Right))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "Top = " & WSTR(rc.Top))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "Bottom = " & WSTR(rc.Bottom))
         DIM rbbi AS REBARBANDINFOW = FBWS_GetRebarBandInfo(hwnd, i)
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "Styles = &h" & HEX(rbbi.fStyle, 8))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "Verbose styles = " & FBWS_GetRebarBandStyles(rbbi.fStyle))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "hwndChild = &h" & HEX(rbbi.hwndChild, 8))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "cxMinChild = " & WSTR(rbbi.cxMinChild))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "cyMinChild = " & WSTR(rbbi.cyMinChild))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "cx = " & WSTR(rbbi.cx))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "wID = " & WSTR(rbbi.wID))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "cyChild = " & WSTR(rbbi.cyChild))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "cyMaxChild = " & WSTR(rbbi.cyMaxChild))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "cyIntegral = " & WSTR(rbbi.cyIntegral))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "cxIdeal = " & WSTR(rbbi.cxIdeal))
         TreeView_AddItem(hTreeView, hSubNode3, NULL, "cxHeader = " & WSTR(rbbi.cxHeader))
      ' TODO: Margins - need a manifest
      NEXT
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' Retrieves child windows of a control
' ========================================================================================
FUNCTION FBWS_ControlEnumChildProc (BYVAL hwnd AS HWND, BYVAL lParam AS LPARAM) AS LONG
   DIM hTreeView AS HWND = cast(HWND, lParam)
   DIM wszClassName AS WSTRING * 260 = AfxGetWindowClassName(hwnd)
   IF UCASE(wszClassName) = "EDIT" AND UCASE(AfxGetWindowClassName(GetParent(hwnd))) = "COMBOBOX" THEN RETURN CTRUE
   IF hTreeView THEN
      DIM hSubNode AS HTREEITEM = TreeView_AddItem(hTreeView, __hControlSubNodeChildren, NULL, wszClassName)
      FBWS_GetControlInfo(hwnd, hTreeView, hSubNode)
   END IF
   FUNCTION = CTRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the child windows
' ========================================================================================
SUB FBWS_GetControlChildWindowsInfo (BYVAL hwnd AS HWND, BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM = NULL, BYVAL bGetWindowInfo AS BOOLEAN = FALSE)
   IF hwnd = NULL THEN RETURN
   DIM wszClassName AS WSTRING * 260
'   IF bGetWindowInfo THEN FBWS_GetWindowInfo(hwnd, hTreeView, hNode)
   __hControlSubNodeChildren = TreeView_AddItem(hTreeView, hNode, NULL, "Child windows")
   EnumChildWindows(hwnd, CAST(WNDENUMPROC, @FBWS_ControlEnumChildProc), CAST(LPARAM, hTreeView))
   IF TreeView_GetChild(hTreeView, __hControlSubNodeChildren) = NULL THEN TreeView_DeleteItem(hTreeView, __hControlSubNodeChildren)
'   TreeView_Expand(hTreeView, __hControlSubNodeChildren, TVE_EXPAND)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves information of child controls
' ========================================================================================
FUNCTION FBWS_GetControlsInfo (BYVAL hwnd AS HWND, BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM, BYVAL bExpandNode AS BOOLEAN = FALSE) AS BOOLEAN
   IF hTreeView = NULL THEN RETURN FALSE
   DIM wszClassName AS WSTRING * 260 = AfxGetWindowClassName(hwnd)
   IF UCASE(wszClassName) = "EDIT" AND UCASE(AfxGetWindowClassName(GetParent(hwnd))) = "COMBOBOX" THEN RETURN TRUE
   DIM hSubNode AS HTREEITEM = TreeView_AddItem(hTreeView, hNode, NULL, wszClassName)
   IF hSubNode = NULL THEN RETURN FALSE
   FBWS_GetControlInfo(hwnd, hTreeView, hSubNode)
   FBWS_GetControlChildWindowsInfo(hwnd, hTreeView, hSubNode)
   IF bExpandNode THEN TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
   FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the class names of the main window's child windows
' ========================================================================================
FUNCTION FBWS_EnumMainWindowChildrenProc (BYVAL hwnd AS HWND, BYVAL lParam AS LPARAM) AS LONG
   DIM hTreeView AS HWND = cast(HWND, lParam)
   IF GetAncestor(hwnd, GA_ROOTOWNER) = GetAncestor(hwnd, GA_PARENT) THEN
      IF hTreeView THEN FBWS_GetControlsInfo(hWnd, hTreeview, __hMainWindowChildrenNode)
   END IF
   FUNCTION = CTRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the main window or popup dialog
' ========================================================================================
SUB FBWS_GetWindowInfo (BYVAL hwnd AS HWND, BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM)

   DIM wszClassName AS WSTRING * 260
   TreeView_AddItem hTreeView, hNode, NULL, "Handle: &h" & HEX(hwnd, 8)
   TreeView_AddItem hTreeView, hNode, NULL, "Caption: " & AfxGetWindowText(hwnd)
   wszClassName = AfxGetWindowClassName(hwnd)
   DIM wszText AS WSTRING * 260
   IF wszClassName = "#32770" THEN wszText = " (Dialog)"
   IF wszClassName = "#32768" THEN wszText = " (Menu)"
   IF wszClassName = "#32769" THEN wszText = " (Desktop window)"
   IF wszClassName = "#32771" THEN wszText = " (Task-switch window)"
   IF wszClassName = "#32772" THEN wszText = " (Icon title)"
   IF LEFT(wszClassName, 14) = "FBWindowClass:" THEN wszText = " (CWindow class)"
   TreeView_AddItem hTreeView, hNode, NULL, "Control ID: &h" & HEX(GetWindowLongPtr(hwnd, GWLP_ID), 8)
   TreeView_AddItem hTreeView, hNode, NULL, "Class name: " & wszClassName
   IF LEN(wszText) THEN TreeView_AddItem hTreeView, hNode, NULL, "Verbose class name: " & wszClassName & wszText
   TreeView_AddItem hTreeView, hNode, NULL, "Class atom: &h" & HEX(GetClassLongPtr(hwnd, GCW_ATOM), 4)
   TreeView_AddItem hTreeView, hNode, NULL, "Extra class bytes: " & WSTR(GetClassLongPtr(hwnd, GCL_CBCLSEXTRA))
   TreeView_AddItem hTreeView, hNode, NULL, "Window class bytes: " & WSTR(GetClassLongPtr(hwnd, GCL_CBWNDEXTRA))
   DIM rc AS RECT = AfxGetWindowRect(hwnd)
   TreeView_AddItem hTreeView, hNode, NULL, "Window rect: " & _
      "(" & WSTR(rc.Left) & ", " & WSTR(rc.Top) & ") - (" & WSTR(rc.Right) & ", " & _
      WSTR(rc.Bottom) & ") - " & WSTR(rc.Right - rc.Left) & "x" & WSTR(rc.Bottom - rc.Top)
   DIM rcClient AS RECT = AfxGetWindowClientRect(hwnd)
   DIM x AS LONG = rc.Left
   DIM y AS LONG = rc.Top
   MapWindowPoints(hwnd, 0, cast(LPPOINT, @rcCLient), 2)
   x = rcClient.left - x
   y = rcClient.top - y
   OffsetRect(@rcClient, -rcClient.left, -rcClient.top)
   OffsetRect(@rcClient, x, y)

' // Get the monitor that the window is currently displayed on
DIM hMonitor AS HMONITOR = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)

' // Get the logical width and height of the monitor.
DIM miex AS MONITORINFOEXW
miex.cbSize = sizeof(miex)
GetMonitorInfoW(hMonitor, CAST(LPMONITORINFO, @miex))
DIM cxLogical AS LONG = (miex.rcMonitor.right  - miex.rcMonitor.left)
DIM cyLogical AS LONG = (miex.rcMonitor.bottom - miex.rcMonitor.top)

' // Get the physical width and height of the monitor.
DIM dm AS DEVMODEW
dm.dmSize = sizeof(dm)
dm.dmDriverExtra = 0
EnumDisplaySettingsW(miex.szDevice, ENUM_CURRENT_SETTINGS, @dm)
DIM cxPhysical AS LONG = dm.dmPelsWidth
DIM cyPhysical as LONG = dm.dmPelsHeight

' // Calculate the scaling factor.
DIM horzScale AS DOUBLE = cxPhysical / cxLogical
DIM vertScale AS DOUBLE = cyPhysical / cyLogical


   TreeView_AddItem hTreeView, hNode, NULL, "Client rect: " & _
      "(" & WSTR(rcClient.Left) & ", " & WSTR(rcClient.Top) & ") - (" & WSTR(rcClient.Right) & ", " & _
      WSTR(rcClient.Bottom) & ") - " & WSTR(rcClient.Right - rcClient.Left) & "x" & WSTR(rcClient.Bottom - rcClient.Top)
   TreeView_AddItem hTreeView, hNode, NULL, "Window width: " & WSTR(rc.Right - rc.Left)
   TreeView_AddItem hTreeView, hNode, NULL, "Window height: " & WSTR(rc.Bottom - rc.Top)
   TreeView_AddItem hTreeView, hNode, NULL, "Client width: " & WSTR(rcClient.Right - rcClient.Left)
   TreeView_AddItem hTreeView, hNode, NULL, "Client height: " & WSTR(rcClient.Bottom - rcClient.Top)
   TreeView_AddItem hTreeView, hNode, NULL, "Scale: " & WSTR(cxLogical)
   TreeView_AddItem hTreeView, hNode, NULL, "Window styles: &h" & HEX(AfxGetWindowStyle(hwnd), 8)
   TreeView_AddItem hTreeView, hNode, NULL, "Verbose window styles: " & FBWS_GetWindowStyles(hwnd, AfxGetWindowStyle(hwnd))
   TreeView_AddItem hTreeView, hNode, NULL, "Extended styles: &h" & HEX(AfxGetWindowExStyle(hwnd), 8)
   TreeView_AddItem hTreeView, hNode, NULL, "Verbose extended styles: " & FBWS_GetWindowExStyles(AfxGetWindowExStyle(hwnd))
   TreeView_AddItem hTreeView, hNode, NULL, "Window procedure: &h" & HEX(GetWindowLongPtr(hwnd, GWLP_WNDPROC), 8)
   TreeView_AddItem hTreeView, hNode, NULL, "Instance handle: &h" & HEX(GetWindowLongPtr(hwnd, GWLP_HINSTANCE), 8)
   TreeView_AddItem hTreeView, hNode, NULL, "User data: &h" & HEX(GetWindowLongPtr(hwnd, GWLP_USERDATA), 8)
   TreeView_AddItem hTreeView, hNode, NULL, "Menu handle: &h" & HEX(GetClassLongPtr(hwnd, GCLP_MENUNAME), 8)
   TreeView_AddItem hTreeView, hNode, NULL, "Icon handle: &h" & HEX(GetClassLongPtr(hwnd, GCLP_HICON), 8)
   TreeView_AddItem hTreeView, hNode, NULL, "Small icon handle: &h" & HEX(GetClassLongPtr(hwnd, GCLP_HICONSM), 8)
   TreeView_AddItem hTreeView, hNode, NULL, "Cursor type: " & FBWS_GetCursorType(cast(HCURSOR, GetClassLongPtr(hwnd, GCLP_HCURSOR)))
   TreeView_AddItem hTreeView, hNode, NULL, "Background brush: " & FBWS_GetBrushType(cast(HBRUSH, GetClassLongPtr(hwnd, GCLP_HBRBACKGROUND)))
   DIM AS DWORD dwProcessId, dwThreadId
   dwThreadId = GetWindowThreadProcessId(hwnd, @dwProcessId)
   TreeView_AddItem hTreeView, hNode, NULL, "Process ID: &h" & HEX(dwProcessId, 8) & " (" & WSTR(dwProcessId) & ")"
   TreeView_AddItem hTreeView, hNode, NULL, "Thread ID: &h" & HEX(dwThreadId, 8) & " (" & WSTR(dwThreadId) & ")"
   ' // Get the font, if available
   DIM lfw AS LOGFONTW
   IF FBWS_GetFont(hwnd, @lfw) THEN
      TreeView_AddItem(hTreeView, hNode, NULL, "Font mame: " & lfw.lfFaceName)
      TreeView_AddItem(hTreeView, hNode, NULL, "Font point size: " & AfxGetFontPointSize(lfw.lfHeight))
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' Retrieves information about the main window
' ========================================================================================
SUB FBWS_GetMainWindowInfo (BYVAL hwndMain AS HWND, BYVAL hTreeView AS HWND)
   IF hwndMain = NULL THEN RETURN
   FBWS_GetWindowInfo(hwndMain, hTreeView, __hMainWindowNode)
   ' // Enumerate child controls
   __hMainWindowChildrenNode = TreeView_AddItem(hTreeView, __hMainWindowNode, NULL, "Child windows")
   EnumChildWindows(hwndMain, CAST(WNDENUMPROC, @FBWS_EnumMainWindowChildrenProc), CAST(LPARAM, hTreeView))
   IF TreeView_GetChild(hTreeView, __hMainWindowChildrenNode) = NULL THEN TreeView_DeleteItem(hTreeView, __hMainWindowChildrenNode)
   TreeView_Expand(hTreeView, __hMainWindowChildrenNode, TVE_EXPAND)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves child windows
' ========================================================================================
FUNCTION FBWS_EnumChildProc (BYVAL hwnd AS HWND, BYVAL lParam AS LPARAM) AS LONG
   DIM hTreeView AS HWND = cast(HWND, lParam)
   IF hTreeView THEN FBWS_GetControlsInfo(hwnd, hTreeview, __hSubNodeChildren)
   FUNCTION = CTRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the child windows
' ========================================================================================
SUB FBWS_GetChildWindowsInfo (BYVAL hwnd AS HWND, BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM = NULL, BYVAL bGetWindowInfo AS BOOLEAN = FALSE)
   IF hwnd = NULL THEN RETURN
   DIM wszClassName AS WSTRING * 260
   wszClassName = AfxGetWindowClassName(hwnd)
   IF hNode = NULL THEN
      __hSubNode = TreeView_AddItem(hTreeView, TVI_ROOT, NULL, wszClassName)
   ELSE
      __hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, wszClassName)
   END IF
   IF bGetWindowInfo THEN FBWS_GetWindowInfo(hwnd, hTreeView, __hSubNode)
   __hSubNodeChildren = TreeView_AddItem(hTreeView, __hSubNode, NULL, "Child windows")
   EnumChildWindows(hwnd, CAST(WNDENUMPROC, @FBWS_EnumChildProc), CAST(LPARAM, hTreeView))
   IF TreeView_GetChild(hTreeView, __hSubNodeChildren) = NULL THEN TreeView_DeleteItem(hTreeView, __hSubNodeChildren)
   TreeView_Expand(hTreeView, __hSubNodeChildren, TVE_EXPAND)
END SUB
' ========================================================================================
