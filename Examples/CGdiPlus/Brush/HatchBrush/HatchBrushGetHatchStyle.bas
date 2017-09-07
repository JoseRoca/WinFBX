' ########################################################################################
' Microsoft Windows
' File: HatchBrushGetHatchStyle.bas
' Contents: GDI+ - HatchBrushGetHatchStyle example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 Jos� Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
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
' The following example sets up two hatch styles: horiz and current (initialized to
' HatchStyleDiagonalCross). A rectangle that uses horiz as the hatch style is painted.
' Then the HatchBrush.GetHatchStyle method is used to get the current hatch style of the
' brush (which at the time is HatchStyleHorizontal). The address of the current HatchStyle
' object (initialized to HatchStyleDiagonalCross) is passed as the return point for the
' call to GetHatchStyle. When the rectangle is painted again, notice that the hatch style
' is again HatchStyleHorizontal (not HatchStyleDiagonalCross). This shows that the call to
' HatchBrush.GetHatchStyle was successful.
' ========================================================================================
SUB Example_HatchBrushGetHatchStyle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Set colors
   DIM blue AS ARGB = GDIP_ARGB(255, 0, 0, 255)          ' // foreground
   DIM turquoise AS ARGB = GDIP_ARGB(255, 0, 255, 255)   ' // background

   ' // Set hatch styles
   DIM horiz AS HatchStyle = HatchStyleHorizontal
   DIM current AS HatchStyle = HatchStyleDiagonalCross

   ' // Set and then draw the first hatch style.
   DIM brush AS CGpHatchBrush = CGpHatchBrush(horiz, blue, turquoise)
   graphics.FillRectangle(@brush, 20, 20, 100, 50)

   ' // Get the current hatch style of the brush.
   current = brush.GetHatchStyle

   ' // Get the current background color of the brush.
   brush.GetBackgroundColor(@current)

   ' // Draw the rectangle again using the current hatch style.
   DIM brush2 AS CGpHatchBrush = CGpHatchBrush(current, blue, turquoise)
   graphics.FillRectangle(@brush2, 130, 20, 100, 50)

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
   pWindow.Create(NULL, "GDI+ HatchBrushGetBackgroundColor", @WndProc)
   ' // Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 250)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear BGR(255, 255, 255)
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_HatchBrushGetHatchStyle(hdc)

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
