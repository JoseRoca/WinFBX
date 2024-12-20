' ########################################################################################
' Microsoft Windows
' File: GetDataSize.bas
' Contents: GDI+ - GetDataSize example
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
' The following example creates a region from a path and then gets the data that describes
' the region.
' ========================================================================================
SUB Example_GetDataSize (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a path
   DIM pathRegion AS CGpRegion = @path

   ' // Get the region data.
   DIM bufferSize AS UINT
   DIM sizeFilled AS UINT
   DIM pData AS BYTE PTR

   bufferSize = pathRegion.GetDataSize
   pData = Allocate(bufferSize)
   pathRegion.GetData(pData, bufferSize, @sizeFilled)

   ' // Inspect or use the region data.
   ' ...
   Delete pData

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
   pWindow.Create(NULL, "GDI+ GetDataSize", @WndProc)
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
   Example_GetDataSize(hdc)

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
