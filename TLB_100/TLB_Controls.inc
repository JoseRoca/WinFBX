' ########################################################################################
' TypeLib Browser
' File: TLB_CONTROLS.INC
' Contents: TypeLib Browser controls
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

ENUM TLB_CONTROLS_IDS
   IDC_TOOLBAR = 1001
   IDC_STATUSBAR, IDC_EDITNAMESPACE, IDC_TAB, IDC_LISTVIEW, IDC_TREEVIEW, IDC_VIEWCODE, IDC_USEAUTOMATIONVIEW
END ENUM

' // Menu and toolbar identifiers
ENUM TLB_MENU_IDS
   IDM_FILE = 28000
   IDM_OPEN, IDM_SAVE, IDM_EXIT, IDM_VIEW, IDM_EVENTS, IDM_EXPAND, IDM_COLLAPSE, IDM_ABOUT
END ENUM

' // Forward declarations
DECLARE FUNCTION TabPage1_WndProc(BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION TabPage2_WndProc(BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION TabPage3_WndProc(BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION TLB_Edit_CodeView_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION TLB_ListviewSortItems (BYVAL hListView AS HWND, BYVAL pnmlvi AS NMLISTVIEW PTR, BYVAL ColumnToSort AS LONG) AS LONG

' ========================================================================================
' Create the status bar
' ========================================================================================
FUNCTION TLB_CreateStatusBar (BYVAL pWindow AS CWindow PTR) AS HWND

   ' // Create the status bar
   DIM hStatusBar AS HWND = pWindow->AddControl("Statusbar", , IDC_STATUSBAR)
   pWindow->UserData(AFX_HSTATUSBAR) = cast(LONG_PTR, hStatusBar)
   ' // Set the parts
   DIM rgWidths(3) AS LONG = {90 * pWindow->rxRatio, 160 * pWindow->rxRatio, 320 * pWindow->rxRatio, -1}
   StatusBar_SetParts(hStatusBar, 4, @rgWidths(0))
   StatusBar_SetText hStatusBar, 0, "0 TypeLibs"
   StatusBar_SetText hStatusBar, 1, "Namespace"
   ' // Update the size of the statusbar
   SendMessage hStatusBar, WM_SIZE, 0, 0

   ' // Create the edit control for the prefix
   DIM hEditNamespace AS HWND = pWindow->AddControl("Edit", hStatusBar, IDC_EDITNAMESPACE, "", 161, 1, 158, 17)
   pWindow->UserData(AFX_HEDITNAMESPACE) = cast(LONG_PTR, hEditNamespace)
   Edit_LimitText(hEditNamespace, 255)
   EnableWindow(hEditNamespace, FALSE)

   FUNCTION = hStatusBar

END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates the toolbar
' ========================================================================================
FUNCTION TLB_CreateToolBar (BYVAL pWindow AS CWindow PTR) AS HWND

   ' // Create the ToolBar Window.
   DIM hToolBar AS HWND = pWindow->AddControl("Toolbar", , IDC_TOOLBAR, "", 0, 0, 0, 0, _
      WS_CHILD OR WS_VISIBLE OR TBSTYLE_TOOLTIPS OR TBSTYLE_FLAT OR TBSTYLE_LIST, WS_EX_STATICEDGE)
   pWindow->UserData(AFX_HTOOLBAR) = cast(LONG_PTR, hToolBar)
   ' // Set the extended style
   ToolBar_SetExtendedStyle hToolbar, TBSTYLE_EX_DRAWDDARROWS

   ' // Module instance handle
   DIM hInst AS HINSTANCE = GetModuleHandle(NULL)
   ' // Create an image list for the toolbar
   DIM hImageList AS HIMAGELIST
   DIM cx AS LONG = 24 * pWindow->DPI \ 96
   hImageList = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 7, 0)
   IF hImageList THEN
      AfxGdipAddIconFromRes(hImageList, hInst, "IDI_TLB_OPEN_48")
      AfxGdipAddIconFromRes(hImageList, hInst, "IDI_TLB_EXPAND_48")
      AfxGdipAddIconFromRes(hImageList, hInst, "IDI_TLB_COLLAPSE_48")
      AfxGdipAddIconFromRes(hImageList, hInst, "IDI_TLB_EVENTS_48")
      AfxGdipAddIconFromRes(hImageList, hInst, "IDI_TLB_SAVE_48")
      AfxGdipAddIconFromRes(hImageList, hInst, "IDI_TLB_VIEW_48")
      AfxGdipAddIconFromRes(hImageList, hInst, "IDI_TLB_INFO_48")
      AfxGdipAddIconFromRes(hImageList, hInst, "IDI_TLB_EXIT_48")
   END IF

   SendMessageW hToolBar, TB_SETIMAGELIST, 0, CAST(LPARAM, hImageList)

   ' // Add buttons to the toolbar
   Toolbar_AddButton hToolBar, 0, IDM_OPEN,     0, 0, 0, "Open"
   Toolbar_AddButton hToolBar, 1, IDM_EXPAND,   0, 0, 0, "Expand"
   Toolbar_AddButton hToolBar, 2, IDM_COLLAPSE, 0, 0, 0, "Collapse"
   Toolbar_AddButton hToolBar, 3, IDM_EVENTS,   0, 0, 0, "Events"
   Toolbar_AddButton hToolBar, 4, IDM_SAVE,     0, 0, 0, "Save"
   ' // Get the use automation view option
   DIM wszBuffer AS WSTRING * MAX_PATH
   DIM wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   GetPrivateProfileStringW("Browser options", "UseAutomationView", NULL, @wszBuffer, MAX_PATH, @wszIniFileName)
   IF VAL(wszBuffer) = 0 THEN
      Toolbar_AddButton hToolBar, 5, IDM_VIEW,     0, 0, 0, "VTable"
   ELSE
      Toolbar_AddButton hToolBar, 5, IDM_VIEW,     0, 0, 0, "VBAuto"
   END IF
   Toolbar_AddButton hToolBar, 6, IDM_ABOUT,    0, 0, 0, "About"
   Toolbar_AddButton hToolBar, 7, IDM_EXIT,     0, 0, 0, "Exit"

   ' // Size the toolbar
   SendMessageW hToolBar, TB_AUTOSIZE, 0, 0

   ' // Return the handle of the toolbar
   FUNCTION = hToolbar

END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates the tab control
' ========================================================================================
FUNCTION TLB_CreateTabControl (BYVAL pWindow AS CWindow PTR) AS HWND

   ' // Get the height of the toolbar
   DIM ToolBarHeight AS LONG =  pWindow->ControlHeight(GetDlgItem(pWindow->hWindow, IDC_TOOLBAR))
   ' // Get the height of the statusbar
   DIM StatusBarHeight AS LONG =  pWindow->ControlHeight(GetDlgItem(pWindow->hWindow, IDC_STATUSBAR))

   ' // Add a tab control
   DIM hTab AS .HWND = pWindow->AddControl("Tab", , IDC_TAB, "", _
      0, ToolBarHeight, pWindow->ClientWidth, pWindow->ClientHeight - ToolBarHeight - StatusBarHeight, _
      WS_CHILD OR WS_VISIBLE OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR WS_GROUP OR WS_TABSTOP OR _
      TCS_FORCELABELLEFT OR TCS_TABS OR TCS_FIXEDWIDTH OR TCS_TOOLTIPS)
   pWindow->UserData(AFX_HTAB) = cast(LONG_PTR, hTab)

   DIM dwExStyle AS DWORD
   IF AfxWindowsVersion > 500 THEN dwExStyle = WS_EX_COMPOSITED   ' Avoids flicker - Must be Windows XP or superior
   ' // Don't use WS_EX_COMPOSITED with the dialog that hosts the list view
   ' // Create the first tab page
   DIM pTabPage1 AS CTabPage PTR = NEW CTabPage
   pTabPage1->DPI = pWindow->DPI
   pTabPage1->InsertPage(hTab, 0, "TypeLibs", , @TabPage1_WndProc)
   ' // Create the second tab page
   DIM pTabPage2 AS CTabPage PTR = NEW CTabPage
   pTabPage2->DPI = pWindow->DPI
   pTabPage2->InsertPage(hTab, 1, "Code", , @TabPage2_WndProc, , dwExStyle)

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
FUNCTION TabPage1_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_CREATE

         ' // Get a pointer to the CWindow class
         DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
         ' // Get a pointer to the TabPage class
         DIM pTabPage AS CTabPage PTR = AfxCTabPagePtr(GetParent(hwnd), 0)

         ' // *** Add a list view to the first page ***
         DIM hListView AS HWND = pTabPage->AddControl("ListView", hwnd, IDC_LISTVIEW, "", _
            0, 0, pTabPage->ControlClientWidth(hwnd) / 2, pTabPage->ControlClientHeight(hwnd), _
            WS_BORDER OR WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR LVS_REPORT OR LVS_SINGLESEL OR LVS_SHOWSELALWAYS OR LVS_NOSORTHEADER)
         pWindow->UserData(AFX_HLISTVIEW) = cast(LONG_PTR, hListView)
         ' // Set the extended class styles for the control
         ' // Note: LVS_EX_GRIDLINES doesn't work properly with XP.
         ' // See: http://support.microsoft.com/kb/813791
         DIM dwExStyle AS DWORD
         IF AfxWindowsVersion < 600 THEN
            dwExStyle = LVS_EX_HEADERDRAGDROP OR LVS_EX_FULLROWSELECT OR LVS_EX_INFOTIP OR LVS_EX_DOUBLEBUFFER
         ELSE
            dwExStyle = LVS_EX_GRIDLINES OR LVS_EX_HEADERDRAGDROP OR LVS_EX_FULLROWSELECT OR LVS_EX_INFOTIP OR LVS_EX_DOUBLEBUFFER
         END IF
         ListView_SetExtendedListViewStyle(hListView, dwExStyle)

         ' // Remove the buttons style from the header control to give it a flat appearance
         ' // BUG: ListView_GetHeader() has a "hwnd" macro parameter, but also wants to use
         ' // the HWND type in the macro body.
'         DIM hwndHdr AS HWND = ListView_GetHeader(hListView)   ' there is a bug in ListView_GetHeader and other macros
         DIM hwndHdr AS HWND = cast(HWND, SendMessageW(hListView, LVM_GETHEADER, 0, 0))
         SetWindowLongPtr hwndHdr, GWL_STYLE, GetWindowLong(hwndHdr, GWL_STYLE) XOR HDS_BUTTONS
         SetWindowPos hListView, NULL, 0, 0, 0, 0, SWP_NOZORDER OR SWP_NOMOVE OR SWP_NOSIZE OR SWP_DRAWFRAME

         ' // Add the columns
         ListView_AddColumn hListView, 0, "File Name", 110 * pTabPage->rxRatio, LVCFMT_LEFT
         ListView_AddColumn hListView, 1, "Description", 245 * pTabPage->rxRatio, LVCFMT_LEFT
         ListView_AddColumn hListView, 2, "Full Path", 500 * pTabPage->rxRatio, LVCFMT_LEFT
         ListView_AddColumn hListView, 3, "CLSID", 270 * pTabPage->rxRatio, LVCFMT_LEFT

         ' // *** Add a tree view to the page ***
         dwExStyle = 0
         IF AfxWindowsVersion > 500 THEN dwExStyle = WS_EX_COMPOSITED   ' Avoids flicker - Must be Windows XP or superior
         DIM hTreeView AS HWND = pTabPage->AddControl("TreeView", hwnd, IDC_TREEVIEW, "", _
            pTabPage->ControlClientWidth(hwnd) / 2, 0, pTabPage->ControlClientWidth(hwnd) / 2, _
            pTabPage->ControlClientHeight(hwnd), , dwExStyle)
         pWindow->UserData(AFX_HTREEVIEW) = cast(LONG_PTR, hTreeView)

         EXIT FUNCTION

      CASE WM_SIZE
         ' // If the window isn't minimized, resize it
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Get a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
            DIM rc AS RECT
            GetClientRect hwnd, @rc
            DIM hListView AS HWND = cast(HWND, pWindow->UserData(AFX_HLISTVIEW))
            IF hListView THEN
               MoveWindow hListView, rc.Left, rc.Top, (rc.Right / 2), rc.Bottom - rc.Top, CTRUE
            END IF
            DIM hTreeView AS HWND = cast(HWND, pWindow->UserData(AFX_HTREEVIEW))
            IF hTreeView THEN
               MoveWindow hTreeView, (rc.Right / 2), rc.Top, (rc.Right / 2), rc.Bottom - rc.Top, CTRUE
            END IF
         END IF

      CASE WM_NOTIFY
         DIM pnmhdr AS NMHDR PTR = cast(NMHDR PTR, lParam)
         SELECT CASE pnmhdr->idFrom
            CASE IDC_LISTVIEW
               SELECT CASE pnmhdr->code
                  CASE LVN_COLUMNCLICK
                     DIM pnmlvi AS NMLISTVIEW PTR = cast(NMLISTVIEW PTR, lParam)
                     ' // Sort items by the clicked column
                     DIM ColumnToSort AS LONG = pnmlvi->iSubItem
                     TLB_ListviewSortItems pnmhdr->hwndFrom, pnmlvi, ColumnToSort
                     EXIT FUNCTION
                  CASE NM_DBLCLK
                     ' // Get a pointer to the CWindow class
                     DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
                     ' // Retrieve the path of the type library
                     DIM nItem AS LONG = ListView_GetSelectionMark(pnmhdr->hwndFrom)
                     IF nItem > - 1 THEN
                        DIM wszPath AS WSTRING * MAX_PATH
                        ListView_GetItemText(pnmhdr->hwndFrom, nItem, 2, @wszPath, SIZEOF(wszPath))
                        ' // Load and parse the type library
                        DIM pParseLib AS CParseTypeLib PTR = NEW CParseTypeLib(pWindow)
                        SetCursor LoadCursor(NULL, IDC_WAIT)
                        pParseLib->LoadTypeLibEx(wszPath)
                        SetCursor LoadCursor(NULL, IDC_ARROW)
                        Delete pParseLib
                     END IF
                     EXIT FUNCTION
               END SELECT
         END SELECT

      CASE WM_SYSCOLORCHANGE
         ' // Forward this message to common controls so that they will
         ' // be properly updated when the user changes the color settings.
         DIM hListView AS HWND = GetDlgItem(hwnd, IDC_LISTVIEW)
         SendMessageW hListView, WM_SYSCOLORCHANGE, wParam, lParam
         DIM hTreeView AS HWND = GetDlgItem(hwnd, IDC_TREEVIEW)
         SendMessageW hTreeView, WM_SYSCOLORCHANGE, wParam, lParam

      CASE WM_SETFOCUS
         ' // Get a pointer to the CWindow class
         DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
         DIM hListView AS HWND = cast(HWND, pWindow->UserData(AFX_HLISTVIEW))
         SetFocus hListView
         ' // Display info in the status bar
         DIM hStatusBar AS HWND = cast(HWND, pWindow->UserData(AFX_HSTATUSBAR))
         DIM wszText AS WSTRING * MAX_PATH = "Double click the wanted item to retrieve information"
         StatusBar_SetText hStatusBar, 3, @wszText
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Tab page 2 window procedure
' ========================================================================================
FUNCTION TabPage2_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

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
            WS_EX_CLIENTEDGE, , @TLB_Edit_CodeView_WndProc)
         pWindow->UserData(AFX_HCODEVIEW) = cast(LONG_PTR, hCodeView)
         Edit_LimitText(hCodeView, 31457279)
         ' // Create a font for the Code view edit control
         DIM hFont AS HFONT
         hFont = pTabPage->CreateFont("Lucida Console", 10, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET)
         ' // Set the font
         AfxSetWindowFont hCodeView, hFont
         pWindow->UserData(AFX_HCODEVIEWFONT) = cast(LONG_PTR, hFont)
         EXIT FUNCTION

      CASE WM_SIZE
         ' // If the window isn't minimized, resize it
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Get a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
            DIM rc AS RECT
            GetClientRect hwnd, @rc
            DIM hCodeView AS HWND = cast(HWND, pWindow->UserData(AFX_HCODEVIEW))
            IF hCodeView THEN
               MoveWindow hCodeView, rc.Left, rc.Top, rc.Right, rc.Bottom - rc.Top, CTRUE
            END IF
         END IF

      CASE WM_SETFOCUS
         ' // Get a pointer to the CWindow class
         DIM pWindow AS CWindow PTR = AfxCWindowOwnerPtr(hwnd)
         DIM hCodeView AS HWND = cast(HWND, pWindow->UserData(AFX_HCODEVIEW))
         SetFocus hCodeView
         ' // Display info in the status bar
         DIM hStatusBar AS HWND = cast(HWND, pWindow->UserData(AFX_HSTATUSBAR))
         DIM wszText AS WSTRING * MAX_PATH = "Ctrl+A = Select all | Ctrl+C = Copy to clipboard | Ctrl+V = Paste from clipboard"
         StatusBar_SetText hStatusBar, 3, @wszText
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Code text box window procedure
' ========================================================================================
FUNCTION TLB_Edit_CodeView_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_KEYDOWN
         ' // If Ctrl+A or Ctrl+E pressed, select all the text
         IF wParam = VK_A OR wParam = VK_E THEN
            IF GetAsyncKeyState(VK_CONTROL) THEN
               PostMessage hwnd, EM_SETSEL, 0, -1
            END IF
         END IF
         ' // Eat the Escape key to avoid the page being destroyed
         IF wParam = VK_ESCAPE THEN EXIT FUNCTION

      CASE WM_DESTROY
         ' // REQUIRED: Remove control subclassing
         SetWindowLongPtrW hwnd, GWLP_WNDPROC, CAST(LONG_PTR, RemovePropW(hwnd, "OLDWNDPROC"))

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = CallWindowProcW(GetPropW(hwnd, "OLDWNDPROC"), hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Listview_SortItems callback procedure
' ========================================================================================
FUNCTION TLB_ListviewSortItemsProc (BYVAL lParam1 AS LPARAM, BYVAL lParam2 AS LPARAM, BYVAL lParamSort AS LPARAM) AS LONG

   DIM pnmlvi AS NMLISTVIEW PTR = cast(NMLISTVIEW PTR, lParamSort)
   IF pnmlvi THEN
      DIM wszText1 AS WSTRING * MAX_PATH, wszText2 AS WSTRING * MAX_PATH
      ListView_GetItemText(pnmlvi->hdr.hwndFrom, lParam1, pnmlvi->iSubItem, @wszText1, MAX_PATH)
      ListView_GetItemText(pnmlvi->hdr.hwndFrom, lParam2, pnmlvi->iSubItem, @wszText2, MAX_PATH)
      IF pnmlvi->lParam = -1 THEN
         ' // Descending order
         FUNCTION = lstrcmpW(@wszText2, @wszText1)
      ELSE
         ' // Ascending order
         FUNCTION = lstrcmpW(@wszText1, @wszText2)
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sort the items of the listview
' ========================================================================================
FUNCTION TLB_ListviewSortItems (BYVAL hListView AS HWND, BYVAL pnmlvi AS NMLISTVIEW PTR, BYVAL ColumnToSort AS LONG) AS LONG
   ' // BUG in ListView_GetHeader
'   DIM hHeader AS HWND = ListView_GetHeader(hListView)
   DIM hHeader AS HWND = cast(HWND, SendMessageW(hListView, LVM_GETHEADER, 0, 0))
   DIM item AS HDITEM
   item.mask = HDI_FORMAT OR HDI_TEXT
   Header_GetItem(hHeader, ColumnToSort, @item)
   ' // Check the direction of the sort arrow and fill the
   ' // lParam value to -1 if we want sorting by descending order
   IF (item.fmt AND HDF_SORTUP) = HDF_SORTUP THEN
      item.fmt XOR= HDF_SORTUP
      item.fmt OR= HDF_SORTDOWN
      pnmlvi->lParam = -1   ' // Descending order
   ELSEIF (item.fmt AND HDF_SORTDOWN) = HDF_SORTUP THEN
      item.fmt XOR= HDF_SORTDOWN
      item.fmt OR= HDF_SORTUP
   ELSE
      item.fmt XOR= HDF_SORTDOWN
      item.fmt OR= HDF_SORTUP
   END IF
   item.mask = HDI_FORMAT
   item.fmt OR= HDF_STRING
   SendMessageW hHeader, HDM_SETITEM, ColumnToSort, cast(LPARAM, @item)
   DIM nCount AS LONG = ListView_GetItemCount(hListView)
   ' // Remove sorting arrows from the other columns
   FOR i AS LONG = 0 TO nCount - 1
      IF i <> ColumnToSort THEN
         item.fmt = HDF_STRING
         SendMessageW hHeader, HDM_SETITEM, i, cast(LPARAM, @item)
      END IF
   NEXT
   TLB_SaveColumnToSort(hListView, ColumnToSort)
   TLB_SaveSortOrder(hListView, pnmlvi->lParam)
   RETURN ListView_SortItemsEx(hListView, @TLB_ListviewSortItemsProc, pnmlvi)
END FUNCTION
' ========================================================================================
