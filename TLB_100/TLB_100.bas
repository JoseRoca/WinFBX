' ########################################################################################
' Microsoft Windows
' Contents: TypeLib Browser 1.00, Release candidate 09
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. All rights reserved. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ###################################e#####################################################

#define UNICODE
#define _WIN32_WINNT &h0502   ' // for compatibility with XP
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxCtl.inc"
#INCLUDE ONCE "Afx/AfxMenu.inc"
#INCLUDE ONCE "Afx/AfxGdiPlus.inc"
#INCLUDE ONCE "Afx/AfxCOM.inc"
USING Afx

' // Constants
CONST CLASSNAME    = "TLB_100"
CONST PROGNAME     = "TLB_100"
CONST TLBVERSION   = "1.00 Beta - Build 10"
CONST TLBCAPTION   = "TypeLib Browser for Free Basic 1.00"
CONST TLBCOPYRIGHT = "© 2016-2017 by José Roca"
CONST MAILADDRESS  = "JRoca@com.it-berater.org"
CONST INIFILENAME  = "TLB_100.ini"

' // Handles
ENUM AFX_USERDATA
   AFX_HMENU, AFX_HMENUFILE, AFX_HMENUHELP
   AFX_HTOOLBAR, AFX_HTAB, AFX_HLISTVIEW, AFX_HTREEVIEW, AFX_HSTATUSBAR
   AFX_HCODEVIEW, AFX_HCODEVIEWFONT, AFX_HEDITNAMESPACE, AFX_HOPTIONS
END ENUM

' // Forward declarations
DECLARE FUNCTION TLB_MsgBox (BYVAL hwnd AS HWND, BYVAL pwszMessage AS WSTRING PTR, BYVAL dwStyle AS DWORD = 0, BYVAL pwszCaption AS WSTRING PTR = NULL) AS LONG

#INCLUDE ONCE "TLB_TypeLib.bi"
#INCLUDE ONCE "TLB_Procs.inc"
#INCLUDE ONCE "TLB_ParseLib.inc"
#INCLUDE ONCE "TLB_EnumTlbs.inc"
#INCLUDE ONCE "TLB_Controls.inc"

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Initialize the COM library
   OleInitialize NULL

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest file
'   AfxSetProcessDPIAware

   ' // Create the main window
   DIM pWindow AS CWindow = CLASSNAME
   ' // Retrieve the size of the working area
   DIM rc AS RECT = pWindow.GetWorkArea
   DIM hwndMain AS HWND = pWindow.Create(NULL, TLBCAPTION & " " & TLBCOPYRIGHT, _
      @WndProc, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top)
   ' // Change the class style to avoid flicker
   pWindow.ClassStyle = CS_DBLCLKS
   ' // Set the client size
   pWindow.SetClientSize 750, 450
   ' // Centers the window
   pWindow.Center
   ' // Sets the icons
   pWindow.BigIcon   = LoadImage(hInstance, MAKEINTRESOURCE(101), IMAGE_ICON, 48, 48, LR_SHARED)
   pWindow.SmallIcon = LoadImage(hInstance, MAKEINTRESOURCE(100), IMAGE_ICON, 32, 32, LR_SHARED)

   ' // Create the toolbar
   TLB_CreateToolbar(@pWindow)
   ' // Create the status barw
   TLB_CreateStatusBar(@pWindow)
   ' // Create the tab control
   TLB_CreateTabControl(@pWindow)

   ' // Read saved window's state and placement
   TLB_SetWindowPlacement(hwndMain)
   ' // Set the ListView columns order
   DIM hListView AS HWND = cast(HWND, pWindow.UserData(AFX_HLISTVIEW))
   TLB_SetColumnsOrderArray(hListView)
   ' // Set the ListView columns width
   TLB_SetColumnsWidth(hListView)

   ' // Post a message to enumerate the type libraries
   PostMessage pWindow.hWindow, WM_USER + 999, 0, 0

   ' // Dispatch Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

   ' // Uninitialize the COM library
   OleUninitialize

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE LOWORD(wParam)

            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF HIWORD(wParam) = BN_CLICKED THEN
                  DIM r AS LONG = TLB_MsgBox(hwnd, "Are you sure you want to exit?", _
                           MB_YESNO, "Exit " & PROGNAME)
                  IF r = IDNO THEN EXIT FUNCTION
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF

            ' // Open a typelib loaded from disk
            CASE IDM_OPEN
               DIM wszFile AS WSTRING * 260 = "*.ocx;*dll;*exe"
               DIM wszInitialDir AS STRING * 260 = CURDIR
               DIM wszFilter AS WSTRING * 260 = "Object files (*.ocx,*.dll,*.exe)|*.ocx;*.dll;*.exe|" & "Type libraries (*.tlb;*.olb)|*.tlb;*.olb|" &  "All Files (*.*)|*.*|"
               DIM dwFlags AS DWORD = OFN_EXPLORER OR OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY
               DIM cwsPath AS CWSTR = AfxOpenFileDialog(hwnd, "", wszFile, wszInitialDir, wszFilter, "ocx", @dwFlags, NULL)
               IF RIGHT(**cwsPath, 1) ="," THEN cwsPath = LEFT(**cwsPath, LEN(cwsPath) - 1)
               IF LEN(cwsPath) THEN
                  ' // Load and parse the type library
                  DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
                  DIM pParseLib AS CParseTypeLib PTR = NEW CParseTypeLib(pWindow)
                  SetCursor LoadCursor(NULL, IDC_WAIT)
                  IF pParseLib->LoadTypeLibEx(cwsPath) = S_OK THEN
                     ' // Add the item to the list view
                     DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
                     DIM hListView AS HWND = cast(HWND, pWindow->UserData(AFX_HLISTVIEW))
                     DIM wszFile AS WSTRING * MAX_PATH = AfxGetFileName(cwsPath)
                     DIM lItemIdx AS LONG = ListView_AddItem(hListView, 0, 0, @wszFile)
                     DIM wszDesc AS WSTRING * MAX_PATH = pParseLib->m_LibHelpString & " (Ver " & WSTR(pParseLib->m_LibMajorVersion) & "." & WSTR(pParseLib->m_LibMinorVersion) & ")"
                     DIM nItem AS LONG = ListView_GetSelectionMark(hListView)
                     ListView_UnselectItem(hListView, nItem)
                     ListView_SetItemText(hListView, lItemIdx, 1, @wszDesc)
                     ListView_SetItemText(hListView, lItemIdx, 2, cwsPath)
                     ListView_SetItemText(hListView, lItemIdx, 3, @pParseLib->m_LibGuid)
                     ListView_SetItemState(hListView, lItemIdx, LVIS_SELECTED OR LVIS_FOCUSED, &h000F)
                     ListView_SetSelectionMark(hListView, lItemIdx)
                     ListView_EnsureVisible(hListView, lItemIdx, CTRUE)
                  END IF
                  SetCursor LoadCursor(NULL, IDC_ARROW)
                  Delete pParseLib
               END IF
               ' // Set the focus in the tree view
               DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
               DIM hTreeView AS HWND = cast(HWND, pWindow->UserData(AFX_HTREEVIEW))
               SetFocus hTreeView
               EXIT FUNCTION

            ' // Save generated code to disk
            CASE IDM_SAVE
               DIM wszFile AS WSTRING * 260 = "*.bi"
               DIM wszInitialDir AS STRING * 260 = CURDIR
               DIM wszFilter AS WSTRING * 260 = "FB code files (*.bas,*.inc,*.bi)|*.bas;*.inc;*.bi|" & "Text files (*.txt)|*.txt|" & "All Files (*.*)|*.*|"
               DIM dwFlags AS DWORD = OFN_EXPLORER OR OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY OR OFN_OVERWRITEPROMPT
               DIM cwsFile AS CWSTR = AfxSaveFileDialog(hwnd, "", wszFile, wszInitialDir, wszFilter, "bi", @dwFlags)
               IF LEN(cwsFile) THEN
                  DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
                  DIM hEdit AS HWND = cast(HWND, pWindow->UserData(AFX_HCODEVIEW))
                  DIM nLen AS LONG = GetWindowTextLengthW(hEdit)
                  IF nLen THEN
                     DIM cwsText AS CWSTR = SPACE(nLen + 1)
                     IF GetWindowText(hEdit, *cwsText, nLen + 1) THEN
                        DIM fn AS LONG = FREEFILE
                        OPEN cwsFile FOR OUTPUT AS #fn
                        IF ERR = 0 THEN
                           PRINT #fn, cwsText
                           CLOSE #fn
                        END IF
                     END IF
                  END IF
               END IF
               EXIT FUNCTION

            CASE IDM_VIEW
               DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
               DIM hToolbar AS HWND = cast(HWND, pWindow->UserData(AFX_HTOOLBAR))
               DIM tbbi AS TBBUTTONINFOW
               tbbi.cbSize = sizeof(TBBUTTONINFOW)
               tbbi.dwMask = TBIF_TEXT
               DIM wszTitle AS WSTRING * MAX_PATH
               tbbi.pszText = @wszTitle
               tbbi.cchText = MAX_PATH
               IF SendMessageW(hToolBar, TB_GETBUTTONINFOW, IDM_VIEW, cast(LPARAM, @tbbi)) <> -1 THEN
                  DIM wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
                  IF wszTitle = "VBAuto" THEN wszTitle = "VTable" ELSE wszTitle = "VBAuto"
                  IF wszTitle = "VBAuto" THEN
                     WritePrivateProfileStringW ("Browser options", "UseAutomationView", "1", @wszIniFileName)
                  ELSE
                     WritePrivateProfileStringW ("Browser options", "UseAutomationView", "0", @wszIniFileName)
                  END IF
                  tbbi.pszText = @wszTitle
                  SendMessageW(hToolBar, TB_SETBUTTONINFOW, IDM_VIEW, cast(LPARAM, @tbbi))
                  ' // Notify to the listview that the view has changed
                  DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
                  DIM hListView AS HWND
                  IF pWindow THEN hListView = cast(HWND, pWindow->UserData(AFX_HLISTVIEW))
                  DIM hdr AS NMHDR
                  hdr.hwndFrom = hListView
                  hdr.idFrom = IDC_LISTVIEW
                  hdr.code = NM_DBLCLK
                  SetCursor LoadCursor(NULL, IDC_WAIT)
                  SendMessageW GetParent(hListView), WM_NOTIFY, IDC_LISTVIEW, CAST(.LPARAM, @hdr)
                  SetCursor LoadCursor(NULL, IDC_ARROW)
               END IF
               EXIT FUNCTION

            ' // Expand all the nodes of the treeview
            CASE IDM_EXPAND
               ' // Disables redraw to minimize flicker
               DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
               DIM hTreeView AS HWND = cast(HWND, pWindow->UserData(AFX_HTREEVIEW))
               SendMessageW(hTreeView, WM_SETREDRAW, FALSE, 0)
               ' // Exapands the items
               SetCursor LoadCursor(NULL, IDC_WAIT)
               TreeView_ExpandAllItems(hTreeView)
               IF TreeView_SelectSetFirstVisible(hTreeView, TreeView_GetSelection(hTreeView)) = NULL THEN
                  TreeView_SelectSetFirstVisible(hTreeView, TreeView_GetRoot(hTreeView))
               END IF
               SetCursor LoadCursor(NULL, IDC_ARROW)
               ' // Enables redraw and repaints the control
               SendMessageW(hTreeView, WM_SETREDRAW, TRUE, 0)
               EXIT FUNCTION

            ' // Collapse all the nodes of the treeview
            CASE IDM_COLLAPSE
               ' // Disables redraw to minimize flicker
               DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
               DIM hTreeView AS HWND = cast(HWND, pWindow->UserData(AFX_HTREEVIEW))
               SendMessageW(hTreeView, WM_SETREDRAW, FALSE, 0)
               ' // Collapses all the items
               TreeView_CollapseAllChildItems(hTreeView, TreeView_GetChild(hTreeView, TreeView_GetRoot(hTreeView)))
               ' // Enables redraw and repaints the control
               SendMessageW(hTreeView, WM_SETREDRAW, TRUE, 0)
               EXIT FUNCTION

            ' // Restore the size of the main window
'            CASE IDM_RESTOREWSIZE
'               DIM rc AS RECT
'               SystemParametersInfo SPI_GETWORKAREA, 0, @rc, 0
'               MoveWindow hwnd, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, CTRUE
'               AfxCenterWindow hwnd
'               EXIT FUNCTION

            ' // End the application
            CASE IDM_EXIT
               SendMessage hwnd, WM_CLOSE, 0, 0
               EXIT FUNCTION

            ' // Display the about box
            CASE IDM_ABOUT
               DIM hFocus AS HWND = GetFocus
               TLB_MsgBox hwnd, "TypeLib Browser version " & TLBVERSION & CHR(13, 10) & _
                  "Object browser and code generator for FreeBASIC™" & CHR(13, 10) & _
                  "Copyright " & TLBCOPYRIGHT & CHR(13, 10) & "All rights reserved", _
                  MB_OK, "About TypeLib Browser"
               SetFocus hFocus
               EXIT FUNCTION

            CASE IDM_EVENTS
               DIM hFocus AS HWND = GetFocus
               TLB_MsgBox(hwnd, "Not yet implemented", MB_OK, "Events" & PROGNAME)
               SetFocus hFocus
               EXIT FUNCTION

         END SELECT

      CASE WM_USER + 998
         ' // Sets the focus in the control specified by wParam
         SetFocus cast(HWND, wParam)
         EXIT FUNCTION

      CASE WM_USER + 999
         ' // Disables redraw to minimize flicker
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
         DIM hListView AS HWND = cast(HWND, pWindow->UserData(AFX_HLISTVIEW))
         SendMessageW(hListView, WM_SETREDRAW, FALSE, 0)
         ' // Delete all items
         ListView_DeleteAllItems(hListView)
         ' // Enumerate the type libraries
         TLB_RegEnumTypeLibs(hListView)
         ' // Enables redraw and repaints the control
         SendMessageW(hListView, WM_SETREDRAW, TRUE, 0)
         ' // Sort the items
         DIM ColumnToSort AS LONG, SortOrder AS LONG
         TLB_GetSortOptions(hListView, @ColumnToSort, @SortOrder)
         DIM hTab AS HWND = cast(HWND, pWindow->UserData(AFX_HTAB))
         DIM pTabPage AS CTabPage PTR = AfxCTabPagePtr(hTab, 0)
         IF pTabPage THEN
            DIM nmlvi AS NMLISTVIEW
            nmlvi.hdr.hwndFrom = hListView
            nmlvi.hdr.idFrom = IDC_LISTVIEW
            nmlvi.hdr.code = LVN_COLUMNCLICK
            nmlvi.iSubItem = ColumnToSort
            nmlvi.lParam = SortOrder
            DIM hHeader AS HWND = cast(HWND, SendMessageW(hListView, LVM_GETHEADER, 0, 0))
'            SendMessageW(hHeader, HDM_SETFOCUSEDITEM, 0, ColumnToSort)   ' // does not work with XP
            DIM item AS HDITEM
            item.mask = HDI_FORMAT OR HDI_TEXT
            Header_GetItem(hHeader, ColumnToSort, @item)
            ' // Reverse the direction because TLB_ListviewSortItems will also reverse it
            IF SortOrder = 0 THEN item.fmt = HDF_SORTDOWN ELSE item.fmt = HDF_SORTUP
            item.mask = HDI_FORMAT
            item.fmt OR= HDF_STRING
            SendMessageW hHeader, HDM_SETITEM, ColumnToSort, cast(LPARAM, @item)
            TLB_ListviewSortItems(hListView, @nmlvi, ColumnToSort)
         END IF
         ' // Display info in the status bar
         DIM hStatusBar AS HWND = cast(HWND, pWindow->UserData(AFX_HSTATUSBAR))
         DIM wszText AS WSTRING * 260 = STR(ListView_GetItemCount(hListView)) & " TypeLibs"
         StatusBar_SetText hStatusBar, 0, @wszText
         wszText = "Double click the wanted item to retrieve information"
         StatusBar_SetText hStatusBar, 3, @wszText
         ' // Set the focus in the list view
         SetFocus hListView
         ' // Restore ListView selected item
         TLB_SetSelectedItem(hListView)
         EXIT FUNCTION

      CASE WM_GETMINMAXINFO
         ' Set the pointer to the address of the MINMAXINFO structure
         DIM ptmmi AS MINMAXINFO PTR = CAST(MINMAXINFO PTR, lParam)
         ' Set the minimum and maximum sizes that can be produced by dragging the borders of the window
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
         IF pWindow THEN
            ptmmi->ptMinTrackSize.x = 460 * pWindow->rxRatio
            ptmmi->ptMinTrackSize.y = 320 * pWindow->ryRatio
         END IF
         EXIT FUNCTION

      CASE WM_SIZE
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Resize the toolbar
            DIM hToolBar AS HWND = GetDlgItem(hwnd, IDC_TOOLBAR)
            SendMessage hToolBar, uMsg, wParam, lParam
            ' // Resize the status bar
            DIM hStatusBar AS HWND = GetDlgItem(hwnd, IDC_STATUSBAR)
            SendMessage hStatusBar, uMsg, wParam, lParam
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            ' // Resize the tab control
            DIM hTab AS HWND = GetDlgItem(hwnd, IDC_TAB)
            DIM ToolBarHeight AS DWORD = pWindow->ControlHeight(hToolBar)
            DIM StatusBarHeight AS DWORD = pWindow->ControlHeight(hStatusBar)
            IF pWindow THEN pWindow->MoveWindow(hTab, 0, ToolBarHeight, pWindow->ClientWidth, _
               pWindow->ClientHeight - ToolBarHeight - StatusBarHeight, CTRUE)
            ' // Resize the tab pages
            AfxResizeTabPages hTab
         END IF

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

    	CASE WM_DESTROY
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
         DIM hFont AS HFONT
         IF pWindow THEN hFont = cast(.HFONT, pWindow->UserData(AFX_HCODEVIEWFONT))
         IF hFont THEN DeleteObject(hFont)
         ' // Save ListView columns order
         DIM hListView AS HWND = cast(HWND, pWindow->UserData(AFX_HLISTVIEW))
         TLB_SaveColumnsOrderArray(hListView)
         ' // Save ListView columns width
         TLB_SaveColumnsWidth(hListView)
         ' // Save ListView selected item
         TLB_SaveSelectedItem(hListView)
         ' // Destroy the tab pages
         AfxDestroyAllTabPages(GetDlgItem(hwnd, IDC_TAB))
         ' // Destroy the toolbar imagelist
         ImageList_Destroy CAST(HIMAGELIST, SendMessageW(GetDlgItem(hwnd, IDC_TOOLBAR), TB_SETIMAGELIST, 0, 0))
         ' // Save the window placement
         TLB_SaveWindowPlacement(hwnd)
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
