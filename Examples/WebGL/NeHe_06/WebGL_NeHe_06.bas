' ########################################################################################
' Microsoft Windows
' File: WebGL_NeHe_06.bas
' Contents: WebBrowser - WebGL - NeHe Lesson 6
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 Jos� Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/CWebBrowser/CWebBrowser.inc"
USING Afx

CONST IDC_WEBBROWSER = 1001

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest file
   AfxSetProcessDPIAware

   ' // Creates the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   DIM hwndMain AS HWND = pWindow.Create(NULL, "WebBrowser - WebGL - Nehe Lesson 6", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(500, 350)
   ' // Centers the window
   pWindow.Center

   ' // Add a WebBrowser control
   DIM pwb AS CWebBrowser = CWebBrowser(@pWindow, IDC_WEBBROWSER, 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Set the IDocHostUIHandler interface
   pwb.SetUIHandler

   ' // Navigate to the path
   DIM wszPath AS WSTRING * MAX_PATH = ExePath & "\index.html"
   pwb.Navigate(wszPath)
   ' // Processes pending Windows messages to allow the page to load
   ' // Needed if the message pump isn't running
   AfxPumpMessages
   ' // Set the focus in the hosted html page
   pwb.SetFocus

   ' // Note: Instead of pWindow.DoeEvents, we need to use a custom message pump
   ' // to be able to forward the keyboard messages to the WebBrowser control.
   ' // Otherwise, the web page will not respond to them.

   ' // Display the window
   ShowWindow(hWndMain, nCmdShow)
   UpdateWindow(hWndMain)

   ' // Dispatch Windows messages
   DIM uMsg AS MSG
   WHILE (GetMessageW(@uMsg, NULL, 0, 0) <> FALSE)
      IF AfxForwardMessage(GetFocus, @uMsg) = FALSE THEN
         IF IsDialogMessageW(hWndMain, @uMsg) = 0 THEN
            TranslateMessage(@uMsg)
            DispatchMessageW(@uMsg)
         END IF
      END IF
   WEND
   FUNCTION = uMsg.wParam

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_SIZE
print "xxxxxxx", GetFOcus
         ' // Optional resizing code
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            ' // Move the position of the control
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDC_WEBBROWSER), _
               0, 0, pWindow->ClientWidth, pWindow->ClientHeight, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
