# CMaskedEdit Class
The `CMaskedEdit` class supports a masked edit control, which validates user input against a mask and displays the validated results according to a template.

**Include file**: CMaskedEdit.inc
  
### Constructors

|Name|Description|  
|----------|-----------------|  
|[Constructors](#Constructors)|Creates an instance of the class|  
  
### Public Methods  
  
|Name|Description|  
|----------|-----------------|  
|[Create](#create)|Creates an instance of the control|  
|[DisableMask](#disablemask)|Disables validating user input.|  
|[EnableGetMaskedCharsOnly](#enablegetmaskedcharsonly)|Specifies whether the `GetWindowText` method retrieves only masked characters.|  
|[EnableMask](#enablemask)|Initializes the masked edit control.|  
|[EnableSetMaskedCharsOnly](#enablesetmaskedcharsonly)|Specifies whether the text is validated against only masked characters, or against the whole mask.|  
|[GetWindowText](#getwindowtext)|Retrieves validated text from the masked edit control.|  
|[hWindow](#hWindow)|Gets the control window handle. |
|[SetValidChars](#setvalidchars)|Specifies a string of valid characters that the user can enter.|  
|[SetWindowText](#setwindowtext)|Displays a prompt in the masked edit control.|  

##  <a name="Constructors"></a>Constructors

```
CONSTRUCTOR CMaskedEdit
CONSTRUCTOR CMaskedEdit (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR,  _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = -1, BYVAL dwExStyle AS DWORD = -1)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWindow* | Pointer to the parent `CWindow`class. |
| *cID* | The control identifier, an integer value used to notify its parent about events. The application determines the control identifier; it must be unique for all controls with the same parent window. |
| *X* | The x-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *y* |  The initial y-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *nWidth* | The width of the control. |
| *nHeight* | The height of the control. |
| *dwStyle* | The window styles of the control being created.<br>Default styles: WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR ES_AUTOHSCROLL. |
| *dwExStyle* | The extended window styles of the control being created.<br> Default extended style: WS_EX_CLIENTEDGE |

### Remarks  
 Perform the following steps to use the `CMaskedEdit` control in your application:  
  
 1. Embed a `CMaskedEdit` object into your window class.  
  
 2. Call the [EnableMask](#enablemask) method to specify the mask.  
  
 3. Call the [SetValidChars](#setvalidchars) method to specify the list of valid characters.  
  
 4. Call the [SetWindowText](#setwindowtext) method to specify the default text for the masked edit control.  
  
 5. Call the [GetWindowText](#getwindowtext) method to retrieve the validated text.  
  
 If you do not call one or more methods to initialize the mask, valid characters, and default text, the masked edit control behaves just as the standard edit control behaves.  
  
### Example  
 The following example demonstrates how to set up a mask (for example a phone number) by using the `EnableMask` method to create the mask for the masked edit control, and `SetWindowText` method to display a prompt in the masked edit control.

```
DIM pMakedEdit AS CMaskedEdit = CMaskedEdit(@pWindow, IDC_MASKED, 10, 30, 280, 23)
pMakedEdit.EnableMask(" ddd  ddd dddd", "(___) ___-____", "_")
pMakedEdit.SetWindowText("(123) 123-1212")
```

### Full example

```
' ########################################################################################
' Microsoft Windows
' Contents: CWindow masked edit control example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2018 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/CMaskedEdit.inc"
USING Afx

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST IDC_MASKED = 1001

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
   pWindow.Create(NULL, "Masked edit control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(300, 150)
   ' // Centers the window
   pWindow.Center

   ' // Adds a button
   pWindow.AddControl("Button", , IDCANCEL, "&Close", 350, 250, 75, 23)

   ' // Add a masked edit control
   DIM pMaskedEdit AS CMaskedEdit = CMaskedEdit(@pWindow, IDC_MASKED, 10, 30, 280, 23)
   SetFocus pMaskedEdit.hWindow
   pMaskedEdit.EnableMask(" ddd  ddd dddd", "(___) ___-____", "_")
   pMaskedEdit.SetWindowText("(123) 123-1212")

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_CREATE
         ' // If you want to create controls in WM_CREATE instead of in WinMain, you can
         ' // retrieve a pointer of the class with AfxCWindowPtr(lParam). Use hwnd as the
         ' // handle of the window instead of pWindow->hWindow or omitting this parameter
         ' // because CWindow doesn't know the value of this handle until the WM_CREATE
         ' // message has been processed.
         ' DIM pWindow AS CWindow PTR = AfxCWindowPtr(lParam)
         ' IF pWindow THEN pWindow->AddControl("Button", hwnd, IDCANCEL, "&Close", 350, 250, 75, 23)
         ' // An alternative is to pass the value of the main window handle to CWindow with
         ' DIM pWindow AS CWindow PTR = AfxCWindowPtr(lParam)
         ' pWindow->hWindow = hwnd
         ' IF pWindow THEN pWindow->AddControl("Button", , IDCANCEL, "&Close", 350, 250, 75, 23)
         EXIT FUNCTION

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
         ' // Optional resizing code
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            ' // Move the position of the button
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDCANCEL), _
               pWindow->ClientWidth - 120, pWindow->ClientHeight - 50, 75, 23, CTRUE
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
```

##  <a name="create"></a>Create

```
FUNCTION CMaskedEdit.Create (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR,  _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = -1, BYVAL dwExStyle AS DWORD = -1) AS HWND
```

### Parameters  

Same parameters that the `Constructor`.

```
DIM pMakedEdit AS CMaskedEdit
pMaskEdit.Create(CMaskedEdit(@pWindow, IDC_MASKED, 10, 30, 280, 23)
pMakedEdit.EnableMask(" ddd  ddd dddd", "(___) ___-____", "_")
pMakedEdit.SetWindowText("(123) 123-1212")
```

##  <a name="disablemask"></a>DisableMask  
 Disables validating user input.  
  
```  
SUB DisableMask
```  
  
### Remarks  
 If user input validation is disabled, the masked edit control behaves like the standard edit control.  
  
##  <a name="enablegetmaskedcharsonly"></a>EnableGetMaskedCharsOnly  
 Specifies whether the `GetWindowText` method retrieves only masked characters.  
  
```  
SUB EnableGetMaskedCharsOnly (BYVAL bEnable AS BOOLEAN = TRUE)
```  
  
| Parameter  | Description |
| ---------- | ----------- |
| *bEnable* | TRUE to specify that the [GetWindowText](#getwindowtext) method retrieve only masked characters; FALSE to specify that the method retrieve the whole text. The default value is TRUE.   |
  
### Remarks  
 Use this method to enable retrieving masked characters. Then create a masked edit control that corresponds to the telephone number, such as (425) 555-0187. If you call the `GetWindowText` method, it returns "4255550187". If you disable retrieving masked characters, the `GetWindowText` method returns the text that is displayed in the edit control, for example "(425) 555-0187".  
  
##  <a name="enablemask"></a>EnableMask  
 Initializes the masked edit control.  
  
```  
SUB EnableMask (BYVAL lpszMask AS WSTRING PTR, BYVAL lpszInputTemplate AS WSTRING PTR, _
   BYREF chMaskInputTemplate AS CWSTR = "_", BYVAL lpszValid AS WSTRING PTR = NULL)
```  
  
| Parameter  | Description |
| ---------- | ----------- |
| *lpszMask* | A mask string that specifies the type of character that can appear at each position in the user input. The length of the *lpszInputTemplate* and *lpszMask* parameter strings must be the same. See the Remarks section for more detail about mask characters. |
| *lpszInputTemplate* | A mask template string that specifies the literal characters that can appear at each position in the user input. Use the underscore character ('\_') as a character placeholder. The length of the *lpszInputTemplate* and *lpszMask* parameter strings must be the same.  |
| *chMaskInputTemplate* | A default character that the framework substitutes for each invalid character in the user input. The default value of this parameter is underscore ('\_'). |
| *lpszValid* | A string that contains a set of valid characters. NULL indicates that all characters are valid. The default value of this parameter is NULL.  |
  
### Remarks  
 Use this method to create the mask for the masked edit control.
  
 The following table list the default mask characters:  
  
|Mask Character|Definition|  
|--------------------|----------------|  
|D|Digit.|  
|d|Digit or space.|  
|+|Plus ('+'), minus ('-'), or space.|  
|C|Alphabetic character.|  
|c|Alphabetic character or space.|  
|A|Alphanumeric character.|  
|a|Alphanumeric character or space.|  
|*|A printable character.|  
  
##  <a name="enablesetmaskedcharsonly"></a>EnableSetMaskedCharsOnly  
 Specifies whether the text is validated against only the masked characters, or against the whole mask.  
  
```  
SUB EnableSetMaskedCharsOnly (BYVAL bEnable AS BOOLEAN = TRUE)
```  
  
| Parameter  | Description |
| ---------- | ----------- |
| *bEnable* | TRUE to validate the user input against only masked characters; FALSE to validate against the whole mask. The default value is TRUE. |
  
##  <a name="getwindowtext"></a>GetWindowText  
 Retrieves validated text from the masked edit control.  
  
```  
FUNCTION GetWindowText () AS CWSTR
```  
  
### Return Value  
 The text from the masked edit control.

## <a name="hWindow"></a>hWindow

Gets the control window handle.

```
PROPERTY hWindow () AS HWND
```

#### Return value

The window handle.

##  <a name="setvalidchars"></a>SetValidChars  
 Specifies a string of valid characters that the user can enter.  
  
```  
SUB SetValidChars (BYVAL lpszValid AS WSTRING PTR)
```  
  
| Parameter  | Description |
| ---------- | ----------- |
| *lpszValid* | A string that contains the set of valid input characters. NULL means that all characters are valid. The default value of this parameter is NULL. |
  
### Remarks  
 Use this method to define a list of valid characters. If an input character is not in this list, masked edit control will not accept it.  
  
 The following code example accepts only hexadecimal numbers.  
 
```
m_wndMaskEdit.EnableMask(" AAAA"), _   ' // Mask string
("0x____"), _   ' // Template string
('_')   ' // The default character that replaces the backspace character
m_wndMaskEdit.SetValidChars("1234567890ABCDEFabcdef")   ' // Valid string characters
m_wndMaskEdit.SetWindowText("0x01AF")
```
  
##  <a name="setwindowtext"></a>SetWindowText  
 Displays a prompt in the masked edit control.  
  
```  
FUNCTION SetWindowText (BYREF cwsText AS CWSTR) AS BOOLEAN
```  
  
| Parameter  | Description |
| ---------- | ----------- |
| *cwsText* | Points to a string that will be used as a prompt. |
