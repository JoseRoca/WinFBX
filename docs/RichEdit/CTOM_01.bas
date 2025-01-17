' ########################################################################################
' Microsoft Windows
' File: CTOM_01.bas
' Contents: CWindow Rich Edit example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxRichEdit.inc"
#define _CTOM_DEBUG_ 1
#INCLUDE ONCE "Afx/CTOM.inc"
#INCLUDE ONCE "Afx/AfxRichEditTOM.inc"
USING Afx

CONST IDC_RICHEDIT = 1001
CONST IDC_TEST = 1002

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
   AfxSetProcessDPIAware

   DIM pWindow AS CWindow
   pWindow.Create(NULL, "CWindow with a Rich Edit control", @WndProc)
   pWindow.SetClientSize(500, 320)
   pWindow.Center

   ' // Add a rich edit control without coordinates (it will be resized in WM_SIZE)
   DIM hRichEdit AS HWND = pWindow.AddControl("RichEdit", , IDC_RICHEDIT, "RichEdit box test")
   SetFocus hRichEdit

   ' // Add a button without coordinates (it will be resized in WM_SIZE)
   pWindow.AddControl("Button", , IDCANCEL, "&Close", 0, 0, 75, 23)
   pWindow.AddControl("Button", , IDC_TEST, "&Test", 0, 0, 75, 23)

   ' // Dispatch Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window callback procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            ' // If ESC key pressed, close the application sending an WM_CLOSE message
            CASE IDCANCEL
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDC_TEST
               DIM hRichEdit AS HWND = GetDlgItem(hwnd, IDC_RICHEDIT)
'               AfxMsg AfxRichEditTOM_GetText(hRichEdit, 3, -1)
'               AfxMsg AfxRichEditTOM_ChangeCase(hRichEdit, tomToggleCase)
'               DIM chrRange AS CHARRANGE = TYPE<CHARRANGE>(3, 12)
'               RichEdit_ExSetSel(hRichEdit, @chrRange)
'               AfxMsg STR(AfxRichEditTOM_GetCch(hRichEdit, 3, 10))
'               AfxMsg RichEdit_GetRtfText(hRichEdit)

               SCOPE
                  DIM pCTextDocument2 AS CTextDocument2 = hRichEdit
                  DIM pCRange2 AS CTextRange2 = pCTextDocument2.Range2(3, 8)
                  IF pCRange2 THEN pcRange2.SetChar(ASC("X"))
               END SCOPE

         END SELECT

      CASE WM_SIZE
         ' // If the window isn't minimized, resize it
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Resize the controls
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            IF pWindow THEN
               pWindow->MoveWindow GetDlgItem(hwnd, IDC_RICHEDIT), 100, 50, pWindow->ClientWidth - 200, pWindow->ClientHeight - 150, CTRUE
               pWindow->MoveWindow GetDlgItem(hwnd, IDCANCEL), pWindow->ClientWidth - 95, pWindow->ClientHeight - 35, 75, 23, CTRUE
               pWindow->MoveWindow GetDlgItem(hwnd, IDC_TEST), pWindow->ClientWidth - 195, pWindow->ClientHeight - 35, 75, 23, CTRUE
            END IF
         END IF

    	CASE WM_DESTROY

         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
