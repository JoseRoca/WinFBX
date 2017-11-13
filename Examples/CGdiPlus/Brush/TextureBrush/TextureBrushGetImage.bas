' ########################################################################################
' Microsoft Windows
' File: TextureBrushGetImage.bas
' Contents: GDI+ - TextureBrushGetImage example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 Jos� Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CGdiPlus/CGdiPlus.inc"
#INCLUDE ONCE "Afx/CGraphCtx.inc"
USING Afx

CONST IDC_GRCTX = 1001

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' The following example creates a texture brush and uses it to fill an ellipse. The code
' then gets the brush's image and draws it.
' ========================================================================================
SUB Example_GetImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a texture brush, and use it to fill an ellipse.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage
   graphics.FillEllipse(@textureBrush, 0, 0, 200, 200)

   ' // Get the brush's image, and draw that image
   DIM pImage2 AS CGpImage
   textureBrush.GetImage(@pImage2)
   graphics.DrawImage(@pImage2, 10, 150)

END SUB
' ========================================================================================

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

   ' // Create the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   pWindow.Create(NULL, "GDI+ TextureBrushGetImage", @WndProc)
   ' //Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(430, 250)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear BGR(255, 255, 255)
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_GetImage(hdc)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

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

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
