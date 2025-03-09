#define UNICODE
#define UNICODE
#include once "Afx/CDialog.inc"
USING Afx

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Menu identifiers
ENUM
   IDM_NEW = 1001     ' New file
   IDM_OPEN           ' Open file...
   IDM_SAVE           ' Save file
   IDM_SAVEAS         ' Save file as...
   IDM_EXIT           ' Exit

   IDM_UNDO = 2001    ' Undo
   IDM_CUT            ' Cut
   IDM_COPY           ' Copy
   IDM_PASTE          ' Paste

   IDM_TILEH = 3001   ' Tile hosizontal
   IDM_TILEV          ' Tile vertical
   IDM_CASCADE        ' Cascade
   IDM_ARRANGE        ' Arrange icons
   IDM_CLOSE          ' Close
   IDM_CLOSEALL       ' Close all
END ENUM

' ========================================================================================
' Build the menu
' ========================================================================================
FUNCTION BuildMenu () AS HMENU

   DIM hMenu AS HMENU
   DIM hPopUpMenu AS HMENU

   hMenu = CreateMenu
   hPopUpMenu = CreatePopUpMenu
      AppendMenuW hMenu, MF_POPUP OR MF_ENABLED, CAST(UINT_PTR, hPopUpMenu), "&File"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_NEW, "&New" & CHR(9) & "Ctrl+N"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_OPEN, "&Open..." & CHR(9) & "Ctrl+O"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_SAVE, "&Save" & CHR(9) & "Ctrl+S"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_SAVEAS, "Save &As..."
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_EXIT, "E&xit" & CHR(9) & "Alt+F4"
   hPopUpMenu = CreatePopUpMenu
      AppendMenuW hMenu, MF_POPUP OR MF_ENABLED, CAST(UINT_PTR, hPopUpMenu), "&Edit"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_UNDO, "&Undo" & CHR(9) & "Ctrl+Z"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_CUT, "Cu&t" & CHR(9) & "Ctrl+X"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_COPY, "&Copy" & CHR(9) & "Ctrl+C"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_PASTE, "&Paste" & CHR(9) & "Ctrl+V"
   hPopUpMenu = CreatePopUpMenu
      AppendMenuW hMenu, MF_POPUP OR MF_ENABLED, CAST(UINT_PTR, hPopUpMenu), "&Window"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_TILEH, "&Tile Horizontal"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_TILEV, "Tile &Vertical"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_CASCADE, "Ca&scade"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_ARRANGE, "&Arrange &Icons"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_CLOSE, "&Close" & CHR(9) & "Ctrl+F4"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_CLOSEALL, "Close &All"
   FUNCTION = hMenu

END FUNCTION
' ========================================================================================

' ========================================================================================
' Dialog callback procedure
' ========================================================================================
FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR
   
   ' // Pointer to the dialog class
   STATIC pDlg AS CDialog PTR
   
   SELECT CASE uMsg

      CASE WM_INITDIALOG
         OutputDebugStringW("DlgProc - WM_INITDIALOG " & WSTR(wParam) & " - " & WSTR(lParam))
         ' // You can add controls here this way:
         ' // Get a pointer to the CDialog class
          pDlg = cast(CDialog PTR, lParam)
'         ' // Add a menu
'         DIM hMenu AS HMENU = BuildMenu
'         SetMenu hDlg, hMenu
'         ' // Add controls to the dialog
'         pDlg->ControlAdd "Button", hDlg, 101,   " Just a simple question ", 5, 5, 160, 55
'         pDlg->ControlAdd "Static", hDlg, 102,   "What's your name ?", 10, 20, 80, 10
'         pDlg->ControlAdd "Edit",   hDlg, 103,   "", 90, 19, 70, 12
'         pDlg->ControlAdd "Button", hDlg, IDOK, "&Ok", 60, 40, 50, 12

      CASE WM_COMMAND
         OutputDebugStringW("DlgProc - WM_COMMAND " & WSTR(wParam) & " - " & WSTR(lParam))
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hDlg, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDOK
               DIM cws AS CWSTR = AfxGetWindowText(GetDlgItem(hDlg, 103))
               IF LEN(cws) THEN MessageBoxW(0, "Your name is " & cws, "Answer", MB_ICONINFORMATION OR MB_TASKMODAL)
               IF LEN(cws) = 0 THEN MessageBoxW(0, "Hmm ...", "No answer", MB_ICONEXCLAMATION OR MB_TASKMODAL)
         END SELECT

      CASE WM_CLOSE
         OutputDebugStringW("DlgProc - WM_CLOSE")
         ' // End the application
         OutputDebugStringW("DlgProc - WM_CLOSE - pDlg = " & WSTR(pDlg))
         IF pDlg THEN pDlg->DialogEnd(1)
         ' // You can use DestroyWindow instead
         ' DestroyWindow hDlg

      CASE WM_DESTROY
         OutputDebugStringW("DlgProc - WM_DESTROY")

   END SELECT

   FUNCTION = 0

END FUNCTION

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   AfxSetProcessDPIAware

   DIM pDlg AS CDialog = CDialog("Segoe UI", 9)
   DIM dwStyle AS LONG = WS_VISIBLE OR WS_OVERLAPPEDWINDOW OR DS_MODALFRAME OR DS_CENTER
   DIM hDlg AS HWND = pDlg.DialogNew(0, "Dialog New", 50, 50, 170, 75, dwStyle)

   ' // Add a menu
   DIM hMenu AS HMENU = BuildMenu
'   SetMenu hDlg, hMenu
   MenuAttach(hMenu, hDlg)

   ' // Add controls to the dialog
   pDlg.ControlAdd "GroupBox", hDlg, 101, "Just a simple question", 5, 5, 160, 55
   pDlg.ControlAdd "Label", hDlg, 102, "What's your name ?", 10, 20, 80, 10
   pDlg.ControlAdd "Edit",   hDlg, 103, "", 90, 19, 70, 12
   pDlg.ControlAdd "Button", hDlg, IDOK, "&Ok", 60, 40, 50, 12

   ' // Set the focus in the edit control
   ControlSetFocus(hDlg, 103)

   ' // Display and activate the dialog as modal
   pDlg.DialogShowModal(hDlg, @DlgProc)

'   ' // Display and activate the dialog as modeless
'   pDlg.DialogShowModeless(hDlg, @DlgProc)

'   ' // Message handler loop
'   DO
'      pDlg.DialogDoEvents
'   LOOP WHILE IsWindow(hDlg)


   ' // You can use a message pump instead
'   DIM uMsg AS tagMsg
'   WHILE GetMessage(@uMsg, NULL, 0, 0)
'      IF IsDialogMessage(hDlg, @uMsg) = 0 THEN
'         TranslateMessage @uMsg
'         DispatchMessage @uMsg
'      END IF
'   WEND
'   FUNCTION = uMsg.wParam

   FUNCTION = pDlg.DialogEndResult

END FUNCTION
' ========================================================================================

