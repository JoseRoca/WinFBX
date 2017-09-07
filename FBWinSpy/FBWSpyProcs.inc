' ========================================================================================
' Get window styles
' ========================================================================================
FUNCTION FBWS_GetWindowStyles (BYVAL hwnd AS HWND, BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h80000000) = &h80000000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_POPUP" ELSE strStyle = "WS_POPUP"
   END IF
   IF (dwStyle AND &h40000000) = &h40000000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_CHILD" ELSE strStyle = "WS_CHILD"
   END IF
   IF (dwStyle AND &h20000000) = &h20000000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_MINIMIZE" ELSE strStyle = "WS_MINIMIZE"
   END IF
   IF (dwStyle AND &h10000000) = &h10000000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_VISIBLE" ELSE strStyle = "WS_VISIBLE"
   END IF
   IF (dwStyle AND &h08000000) = &h08000000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_DISABLED" ELSE strStyle = "WS_DISABLED"
   END IF
   IF (dwStyle AND &h04000000) = &h04000000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_CLIPSIBLINGS" ELSE strStyle = "WS_CLIPSIBLINGS"
   END IF
   IF (dwStyle AND &h02000000) = &h02000000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_CLIPCHILDREN" ELSE strStyle = "WS_CLIPCHILDREN"
   END IF
   IF (dwStyle AND &h01000000) = &h01000000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_MAXIMIZE" ELSE strStyle = "WS_MAXIMIZE"
   END IF
   IF (dwStyle AND &h00C00000) = &h00C00000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_CAPTION" ELSE strStyle = "WS_CAPTION"
   END IF
   IF (dwStyle AND &h00800000) = &h00800000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_BORDER" ELSE strStyle = "WS_BORDER"
   END IF
   IF (dwStyle AND &h00400000) = &h00400000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_DLGFRAME" ELSE strStyle = "WS_DLGFRAME"
   END IF
   IF (dwStyle AND &h00200000) = &h00200000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_VSCROLL" ELSE strStyle = "WS_VSCROLL"
   END IF
   IF (dwStyle AND &h00100000) = &h00100000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_HSCROLL" ELSE strStyle = "WS_HSCROLL"
   END IF
   IF (dwStyle AND &h00080000) = &h00080000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_SYSMENU" ELSE strStyle = "WS_SYSMENU"
   END IF
   IF (dwStyle AND &h00040000) = &h00040000 THEN
      IF LEN(strStyle) THEN strStyle += " OR WS_THICKFRAME" ELSE strStyle = "WS_THICKFRAME"
   END IF
   IF AfxGetFormHandle(hwnd) = hwnd THEN
      IF (dwStyle AND &h00020000) = &h00020000 THEN
         IF LEN(strStyle) THEN strStyle += " OR WS_MINIMIZEBOX" ELSE strStyle = "WS_MINIMIZEBOX"
      END IF
      IF (dwStyle AND &h00010000) = &h00010000 THEN
         IF LEN(strStyle) THEN strStyle += " OR WS_MAXIMIZEBOX" ELSE strStyle = "WS_MAXIMIZEBOX"
      END IF
      IF (dwStyle AND &h01) = &h01 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_ABSALIGN" ELSE strStyle = "DS_ABSALIGN"
      END IF
      IF (dwStyle AND &h02) = &h02 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_SYSMODAL" ELSE strStyle = "DS_SYSMODAL"
      END IF
      IF (dwStyle AND &h20) = &h20 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_LOCALEDIT" ELSE strStyle = "DS_LOCALEDIT"
      END IF
      IF (dwStyle AND &h40) = &h40 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_SETFONT" ELSE strStyle = "DS_SETFONT"
      END IF
      IF (dwStyle AND &h80) = &h80 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_MODALFRAME" ELSE strStyle = "DS_MODALFRAME"
      END IF
      IF (dwStyle AND &h100) = &h100 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_NOIDLEMSG" ELSE strStyle = "DS_NOIDLEMSG"
      END IF
      IF (dwStyle AND &h200) = &h200 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_SETFOREGROUND" ELSE strStyle = "DS_SETFOREGROUND"
      END IF
      IF (dwStyle AND &h0004) = &h0004 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_3DLOOK" ELSE strStyle = "DS_3DLOOK"
      END IF
      IF (dwStyle AND &h0008) = &h0008 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_FIXEDSYS" ELSE strStyle = "DS_FIXEDSYS"
      END IF
      IF (dwStyle AND &h0010) = &h0010 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_NOFAILCREATE" ELSE strStyle = "DS_NOFAILCREATE"
      END IF
      IF (dwStyle AND &h0400) = &h0400 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_CONTROL" ELSE strStyle = "DS_CONTROL"
      END IF
      IF (dwStyle AND &h0800) = &h0800 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_CENTER" ELSE strStyle = "DS_CENTER"
      END IF
      IF (dwStyle AND &h1000) = &h1000 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_CENTERMOUSE" ELSE strStyle = "DS_CENTERMOUSE"
      END IF
      IF (dwStyle AND &h2000) = &h2000 THEN
         IF LEN(strStyle) THEN strStyle += " OR DS_CONTEXTHELP" ELSE strStyle = "DS_CONTEXTHELP"
      END IF
   ELSE
      IF (dwStyle AND &h00020000) = &h00020000 THEN
         IF LEN(strStyle) THEN strStyle += " OR WS_GROUP" ELSE strStyle = "WS_GROUP"
      END IF
      IF (dwStyle AND &h00010000) = &h00010000 THEN
         IF LEN(strStyle) THEN strStyle += " OR WS_TABSTOP" ELSE strStyle = "WS_TABSTOP"
      END IF
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Get extended window styles
' ========================================================================================
FUNCTION FBWS_GetWindowExStyles (BYVAL dwExStyle AS DWORD) AS STRING
   DIM strExStyle AS STRING
   IF (dwExStyle AND &h00000001) = &h00000001 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_DLGMODALFRAME" ELSE strExStyle = "WS_EX_DLGMODALFRAME"
   END IF
   IF (dwExStyle AND &h00000004) = &h00000004 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_NOPARENTNOTIFY" ELSE strExStyle = "WS_EX_NOPARENTNOTIFY"
   END IF
   IF (dwExStyle AND &h00000008) = &h00000008 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_TOPMOST" ELSE strExStyle = "WS_EX_TOPMOST"
   END IF
   IF (dwExStyle AND &h00000010) = &h00000010 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_ACCEPTFILES" ELSE strExStyle = "WS_EX_ACCEPTFILES"
   END IF
   IF (dwExStyle AND &h00000020) = &h00000020 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_TRANSPARENT" ELSE strExStyle = "WS_EX_TRANSPARENT"
   END IF
   IF (dwExStyle AND &h00000040) = &h00000040 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_MDICHILD" ELSE strExStyle = "WS_EX_MDICHILD"
   END IF
   IF (dwExStyle AND &h00000080) = &h00000080 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_TOOLWINDOW" ELSE strExStyle = "WS_EX_TOOLWINDOW"
   END IF
   IF (dwExStyle AND &h00000100) = &h00000100 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_WINDOWEDGE" ELSE strExStyle = "WS_EX_WINDOWEDGE"
   END IF
   IF (dwExStyle AND &h00000200) = &h00000200 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_CLIENTEDGE" ELSE strExStyle = "WS_EX_CLIENTEDGE"
   END IF
   IF (dwExStyle AND &h00000400) = &h00000400 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_CONTEXTHELP" ELSE strExStyle = "WS_EX_CONTEXTHELP"
   END IF
   IF (dwExStyle AND &h00001000) = &h00001000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_RIGHT" ELSE strExStyle = "WS_EX_RIGHT"
   END IF
   IF (dwExStyle AND &h00002000) = &h00002000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_RTLREADING" ELSE strExStyle = "WS_EX_RTLREADING"
   END IF
   IF (dwExStyle AND &h00004000) = &h00004000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_LEFTSCROLLBAR" ELSE strExStyle = "WS_EX_LEFTSCROLLBAR"
   END IF
   IF (dwExStyle AND &h00010000) = &h00010000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_CONTROLPARENT" ELSE strExStyle = "WS_EX_CONTROLPARENT"
   END IF
   IF (dwExStyle AND &h00020000) = &h00020000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_STATICEDGE" ELSE strExStyle = "WS_EX_STATICEDGE"
   END IF
   IF (dwExStyle AND &h00040000) = &h00040000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_APPWINDOW" ELSE strExStyle = "WS_EX_APPWINDOW"
   END IF
   IF (dwExStyle AND &h00080000) = &h00080000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_LAYERED" ELSE strExStyle = "WS_EX_LAYERED"
   END IF
   IF (dwExStyle AND &h00100000) = &h00100000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_NOINHERITLAYOUT" ELSE strExStyle = "WS_EX_NOINHERITLAYOUT"
   END IF
   IF (dwExStyle AND &h00200000) = &h00200000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_NOREDIRECTIONBITMAP" ELSE strExStyle = "WS_EX_NOREDIRECTIONBITMAP"
   END IF
   IF (dwExStyle AND &h00400000) = &h00400000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_LAYOUTRTL" ELSE strExStyle = "WS_EX_LAYOUTRTL"
   END IF
   IF (dwExStyle AND &h02000000) = &h02000000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_COMPOSITED" ELSE strExStyle = "WS_EX_COMPOSITED"
   END IF
   IF (dwExStyle AND &h08000000) = &h08000000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR WS_EX_NOACTIVATE" ELSE strExStyle = "WS_EX_NOACTIVATE"
   END IF
   FUNCTION = strExStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns an string with the menu type
' ========================================================================================
FUNCTION FBWS_GetMenuItemType (BYVAL fType AS UINT) AS STRING
   DIM strType AS STRING
   IF (fType AND &h00000800) = &h00000800 THEN RETURN "MF_SEPARATOR"
   IF (fType AND &h00004000) = &H00004000 THEN
      IF LEN(strType) THEN strType += " OR "
      strType += "MFT_RIGHTJUSTIFY"
      fType XOR= &h00004000
   END IF
   IF (fType AND &h00002000) = &H00002000 THEN
      IF LEN(strType) THEN strType += " OR "
      strType += "MFT_RIGHTORDER"
      fType XOR= &h00002000
   END IF
   IF (fType AND &h00000200) = &h00000200 THEN
      IF LEN(strType) THEN strType += " OR "
      strType += "MFT_RADIOCHECK"
      fType XOR= &h00000200
   END IF
   IF (fType AND &h00000100) = &h00000100 THEN
      IF LEN(strType) THEN strType += " OR "
      strType += "MFT_OWNERDRAW"
      fType XOR= &h00000100
   END IF
   IF (fType AND &h00000040) = &h00000040 THEN
      IF LEN(strType) THEN strType += " OR "
      strType += "MFT_MENUBREAK"
      fType XOR= &h00000040
   END IF
   IF (fType AND &h00000020) = &h00000020 THEN
      IF LEN(strType) THEN strType += " OR "
      strType += "MFT_MENUBARBREAK"
      fType XOR= &h00000020
   END IF
   IF (fType AND &h00000004) = &h00000004 THEN
      IF LEN(strType) THEN strType += " OR "
      strType += "MFT_BITMAP"
      fType XOR= &h00000004
   END IF
   RETURN strType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns an string with the menu state
' ========================================================================================
FUNCTION FBWS_GetMenuItemState (BYVAL fState AS UINT) AS STRING
   DIM strState AS STRING
   IF (fState AND &h1) = &h1 THEN RETURN "MF_GRAYED"
   IF (fState AND &h2) = &h2 THEN
      IF LEN(strState) THEN strState + = " OR "
      strState += "MF_DISABLED"
   END IF
   IF (fState AND &h8) = &h8 THEN
      IF LEN(strState) THEN strState + = " OR "
      strState += "MF_CHECKED"
   END IF
   IF strState = "" THEN strState = "MF_ENABLED"
   RETURN strState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieve the submenu items
' ========================================================================================
FUNCTION FBWS_BuildMenuGetSubMenuItems (BYVAL hSubMenu AS HMENU, BYVAL bRecurse AS BOOLEAN = FALSE) AS CWSTR
   DIM cwsText AS CWSTR
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
            DIM strType AS STRING = FBWS_GetMenuItemType(mii.fType)
            DIM strState AS STRING = FBWS_GetMenuItemState(mii.fState)
            IF strState = "" THEN strState = "MF_ENABLED"
            IF strType = "MF_SEPARATOR" THEN
               strState = "MF_SEPARATOR"
            ELSE
               IF LEN(strType) THEN strState += " OR " & strType
            END IF
            DIM strEscape AS STRING = "$"
            IF INSTR(wszText, ":\") = 0 AND INSTR(wszText, ":/") = 0 THEN strEscape = ""
            wszText = AfxStrReplace(wszText, CHR(9), CHR(34) & " & CHR(9) & " & CHR(34))
            DIM strHandle AS STRING = "hPopUpMenu"
            IF bRecurse THEN strHandle = "hSubPopUpMenu"
            cwsText += "      AppendMenuW "& strHandle & ", " & strState & ", " & WSTR(mii.wID) & ", " & strEscape & CHR(34) & wszText & CHR(34) & CHR(13, 10)
         ELSE
            DIM strState AS STRING = FBWS_GetMenuItemState(mii.fState)
            IF strState = "" THEN strState = "MF_ENABLED"
            cwsText += "   ' ------------------------------------------------------" & CHR(13, 10)
            cwsText += "   hSubPopUpMenu = CreatePopUpMenu" & CHR(13, 10)
            cwsText += "      AppendMenuW hPopUpMenu, MF_POPUP OR " & strState & ", CAST(UINT_PTR, hSubPopUpMenu), " & CHR(34) & wszText & CHR(34) & CHR(13, 10)
            cwsText += FBWS_BuildMenuGetSubMenuItems(mii.hSubMenu, TRUE)
            cwsText += "   ' ------------------------------------------------------" & CHR(13, 10)
         END IF
      END IF
   NEXT
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the menu of the application, if any.
' ========================================================================================
FUNCTION FBWS_BuildMenu (BYVAL hwnd AS HWND) AS CWSTR

   IF hwnd = NULL THEN RETURN ""
   DIM _hMenu AS HMENU = GetMenu(hwnd)
   IF _hMenu = NULL THEN RETURN ""

   DIM nPos AS LONG, _hSubMenu AS HMENU, _hSubMenu2 AS HMENU
   DIM cwsType AS CWSTR, cwsState AS CWSTR, cwsText AS CWSTR

   cwsText  = "' ========================================================================================" & CHR(13, 10)
   cwsText += "' Builds the menu" & CHR(13, 10)
   cwsText += "' ========================================================================================" & CHR(13, 10)
   cwsText += "FUNCTION BuildMenu () AS HMENU" & CHR(13, 10)
   cwsText += "   DIM hMenu AS HMENU, hPopUpMenu AS HMENU, hSubPopUpMenu AS HMENU" & CHR(13, 10)
   cwsText += "   hMenu = CreateMenu" & CHR(13, 10)
   DO
      cwsType = "" : cwsState = ""
      _hSubMenu = GetSubMenu(_hMenu, nPos)
      IF _hSubMenu = NULL THEN EXIT DO
      IF _hSubMenu <> GetSystemMenu(hwnd, 0) THEN
         DIM mii AS MENUITEMINFOW
         mii.cbSize = SIZEOF(mii)
         mii.fMask = MIIM_CHECKMARKS OR MIIM_DATA OR MIIM_ID OR MIIM_STATE OR MIIM_SUBMENU OR MIIM_TYPE
         DIM wszText AS WSTRING * MAX_PATH
         mii.cch = SIZEOF(wszText)
         mii.dwTypeData = @wszText
         IF GetMenuItemInfoW(_hMenu, nPos, 1, @mii) THEN
            IF LEN(wszText) THEN
               DIM strState AS STRING = FBWS_GetMenuItemState(mii.fState)
               IF strState = "" THEN strState = "MF_ENABLED"
               cwsText += "   hPopUpMenu = CreatePopUpMenu" & CHR(13, 10)
               cwsText += "      AppendMenuW hMenu, MF_POPUP OR " & strState & ", CAST(UINT_PTR, hPopUpMenu), " & CHR(34) & wszText & CHR(34) & CHR(13, 10)
            END IF
         END IF
         cwsText += FBWS_BuildMenuGetSubMenuItems(_hSubMenu)
      END IF
      nPos += 1
   LOOP
   cwsText += "   RETURN hMenu" & CHR(13, 10)
   cwsText += "END FUNCTION" & CHR(13, 10)
   cwsText += "' ========================================================================================" & CHR(13, 10)
'   RichEdit_SetText(GetDlgItem(hwnd, IDC_RICHEDIT), **cwsText)
   RETURN cwsText

END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - Static control.
' ========================================================================================
FUNCTION FBWS_GetStaticStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h00000001) = &h00000001 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_CENTER" ELSE strStyle = "SS_CENTER"
   END IF
   IF (dwStyle AND &h00000002) = &h00000002 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_RIGHT" ELSE strStyle = "SS_RIGHT"
   END IF
   IF (dwStyle AND &h00000003) = &h00000003 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_ICON" ELSE strStyle = "SS_ICON"
   END IF
   IF (dwStyle AND &h00000004) = &h00000004 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_BLACKRECT" ELSE strStyle = "SS_BLACKRECT"
   END IF
   IF (dwStyle AND &h00000005) = &h00000005 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_GRAYRECT" ELSE strStyle = "SS_GRAYRECT"
   END IF
   IF (dwStyle AND &h00000006) = &h00000006 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_WHITERECT" ELSE strStyle = "SS_WHITERECT"
   END IF
   IF (dwStyle AND &h00000007) = &h00000007 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_BLACKFRAME" ELSE strStyle = "SS_BLACKFRAME"
   END IF
   IF (dwStyle AND &h00000008) = &h00000008 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_GRAYFRAME" ELSE strStyle = "SS_GRAYFRAME"
   END IF
   IF (dwStyle AND &h00000009) = &h00000009 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_WHITEFRAME" ELSE strStyle = "SS_WHITEFRAME"
   END IF
   IF (dwStyle AND &h0000000A) = &h0000000A THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_USERITEM" ELSE strStyle = "SS_USERITEM"
   END IF
   IF (dwStyle AND &h0000000B) = &h0000000B THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_SIMPLE" ELSE strStyle = "SS_SIMPLE"
   END IF
   IF (dwStyle AND &h0000000C) = &h0000000C THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_LEFTNOWORDWRAP" ELSE strStyle = "SS_LEFTNOWORDWRAP"
   END IF
   IF (dwStyle AND &h0000000D) = &h0000000D THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_OWNERDRAW" ELSE strStyle = "SS_OWNERDRAW"
   END IF
   IF (dwStyle AND &h0000000E) = &h0000000E THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_BITMAP" ELSE strStyle = "SS_BITMAP"
   END IF
   IF (dwStyle AND &h0000000F) = &h0000000F THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_ENHMETAFILE" ELSE strStyle = "SS_ENHMETAFILE"
   END IF
   IF (dwStyle AND &h00000010) = &h00000010 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_ETCHEDHORZ" ELSE strStyle = "SS_ETCHEDHORZ"
   END IF
   IF (dwStyle AND &h00000011) = &h00000011 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_ETCHEDVERT" ELSE strStyle = "SS_ETCHEDVERT"
   END IF
   IF (dwStyle AND &h00000012) = &h00000012 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_ETCHEDFRAME" ELSE strStyle = "SS_ETCHEDFRAME"
   END IF
   IF (dwStyle AND &h0000001F) = &h0000001F THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_TYPEMASK" ELSE strStyle = "SS_TYPEMASK"
   END IF
   IF (dwStyle AND &h00000080) = &h00000080 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_NOPREFIX" ELSE strStyle = "SS_NOPREFIX"
   END IF
   IF (dwStyle AND &h00000100) = &h00000100 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_NOTIFY" ELSE strStyle = "SS_NOTIFY"
   END IF
   IF (dwStyle AND &h00000200) = &h00000200 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_CENTERIMAGE" ELSE strStyle = "SS_CENTERIMAGE"
   END IF
   IF (dwStyle AND &h00000400) = &h00000400 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_RIGHTJUST" ELSE strStyle = "SS_RIGHTJUST"
   END IF
   IF (dwStyle AND &h00000800) = &h00000800 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_REALSIZEIMAGE" ELSE strStyle = "SS_REALSIZEIMAGE"
   END IF
   IF (dwStyle AND &h00001000) = &h00001000 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_SUNKEN" ELSE strStyle = "SS_SUNKEN"
   END IF
   IF (dwStyle AND &h00004000) = &h00004000 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_ENDELLIPSIS" ELSE strStyle = "SS_ENDELLIPSIS"
   END IF
   IF (dwStyle AND &h00008000) = &h00008000 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_PATHELLIPSIS" ELSE strStyle = "SS_PATHELLIPSIS"
   END IF
   IF (dwStyle AND &h0000C000) = &h0000C000 THEN
      IF LEN(strStyle) THEN strStyle += " OR SS_WORDELLIPSIS" ELSE strStyle = "SS_WORDELLIPSIS"
   END IF
   FUNCTION = strstyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - Button control.
' ========================================================================================
FUNCTION FBWS_GetButtonStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h00000001) = &h00000001 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_DEFPUSHBUTTON" ELSE strStyle = "BS_DEFPUSHBUTTON"
   END IF
   IF (dwStyle AND &h00000002) = &h00000002 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_CHECKBOX" ELSE strStyle = "BS_CHECKBOX"
   END IF
   IF (dwStyle AND &h00000003) = &h00000003 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_AUTOCHECKBOX" ELSE strStyle = "BS_AUTOCHECKBOX"
   END IF
   IF (dwStyle AND &h00000004) = &h00000004 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_RADIOBUTTON" ELSE strStyle = "BS_RADIOBUTTON"
   END IF
   IF (dwStyle AND &h00000005) = &h00000005 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_3STATE" ELSE strStyle = "BS_3STATE"
   END IF
   IF (dwStyle AND &h00000006) = &h00000006 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_AUTO3STATE" ELSE strStyle = "BS_AUTO3STATE"
   END IF
   IF (dwStyle AND &h00000007) = &h00000007 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_GROUPBOX" ELSE strStyle = "BS_GROUPBOX"
   END IF
   IF (dwStyle AND &h00000008) = &h00000008 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_USERBUTTON" ELSE strStyle = "BS_USERBUTTON"
   END IF
   IF (dwStyle AND &h00000009) = &h00000009 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_AUTORADIOBUTTON" ELSE strStyle = "BS_AUTORADIOBUTTON"
   END IF
   IF (dwStyle AND &h0000000A) = &h0000000A THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_PUSHBOX" ELSE strStyle = "BS_PUSHBOX"
   END IF
   IF (dwStyle AND &h0000000B) = &h0000000B THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_OWNERDRAW" ELSE strStyle = "BS_OWNERDRAW"
   END IF
   IF (dwStyle AND &h0000000F) = &h0000000F THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_TYPEMASK" ELSE strStyle = "BS_TYPEMASK"
   END IF
   IF (dwStyle AND &h00000020) = &h00000020 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_LEFTTEXT" ELSE strStyle = "BS_LEFTTEXT"
   END IF
   IF (dwStyle AND &h00000040) = &h00000040 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_ICON" ELSE strStyle = "BS_ICON"
   END IF
   IF (dwStyle AND &h00000080) = &h00000080 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_BITMAP" ELSE strStyle = "BS_BITMAP"
   END IF
   IF (dwStyle AND &h00000100) = &h00000100 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_LEFT" ELSE strStyle = "BS_LEFT"
   END IF
   IF (dwStyle AND &h00000200) = &h00000200 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_RIGHT" ELSE strStyle = "BS_RIGHT"
   END IF
   IF (dwStyle AND &h00000300) = &h00000300 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_CENTER" ELSE strStyle = "BS_CENTER"
   END IF
   IF (dwStyle AND &h00000400) = &h00000400 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_TOP" ELSE strStyle = "BS_TOP"
   END IF
   IF (dwStyle AND &h00000800) = &h00000800 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_BOTTOM" ELSE strStyle = "BS_BOTTOM"
   END IF
   IF (dwStyle AND &h00000C00) = &h00000C00 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_VCENTER" ELSE strStyle = "BS_VCENTER"
   END IF
   IF (dwStyle AND &h00001000) = &h00001000 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_PUSHLIKE" ELSE strStyle = "BS_PUSHLIKE"
   END IF
   IF (dwStyle AND &h00002000) = &h00002000 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_MULTILINE" ELSE strStyle = "BS_MULTILINE"
   END IF
   IF (dwStyle AND &h00004000) = &h00004000 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_NOTIFY" ELSE strStyle = "BS_NOTIFY"
   END IF
   IF (dwStyle AND &h00008000) = &h00008000 THEN
      IF LEN(strStyle) THEN strStyle += " OR BS_FLAT" ELSE strStyle = "BS_FLAT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - Toolbar control.
' ========================================================================================
FUNCTION FBWS_GetToolbarStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_SEP" ELSE strStyle = "TBSTYLE_SEP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_CHECK" ELSE strStyle = "TBSTYLE_CHECK"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_GROUP" ELSE strStyle = "TBSTYLE_GROUP"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_DROPDOWN" ELSE strStyle = "TBSTYLE_DROPDOWN"
   END IF
   IF (dwStyle AND &h10) = &h10 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_AUTOSIZE" ELSE strStyle = "TBSTYLE_AUTOSIZE"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_NOPREFIX" ELSE strStyle = "TBSTYLE_NOPREFIX"
   END IF
   IF (dwStyle AND &h100) = &h100 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_TOOLTIPS" ELSE strStyle = "TBSTYLE_TOOLTIPS"
   END IF
   IF (dwStyle AND &h200) = &h200 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_WRAPABLE" ELSE strStyle = "TBSTYLE_WRAPABLE"
   END IF
   IF (dwStyle AND &h400) = &h400 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_ALTDRAG" ELSE strStyle = "TBSTYLE_ALTDRAG"
   END IF
   IF (dwStyle AND &h800) = &h800 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_FLAT" ELSE strStyle = "TBSTYLE_FLAT"
   END IF
   IF (dwStyle AND &h1000) = &h1000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_LIST" ELSE strStyle = "TBSTYLE_LIST"
   END IF
   IF (dwStyle AND &h2000) = &h2000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_CUSTOMERASE" ELSE strStyle = "TBSTYLE_CUSTOMERASE"
   END IF
   IF (dwStyle AND &h4000) = &h4000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_REGISTERDROP" ELSE strStyle = "TBSTYLE_REGISTERDROP"
   END IF
   IF (dwStyle AND &h8000) = &h8000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBSTYLE_TRANSPARENT" ELSE strStyle = "TBSTYLE_TRANSPARENT"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR BTNS_SHOWTEXT" ELSE strStyle = "BTNS_SHOWTEXT"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR BTNS_WHOLEDROPDOWN" ELSE strStyle = "BTNS_WHOLEDROPDOWN"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric extended styles to constants - Toolbar control.
' ========================================================================================
FUNCTION FBWS_GetToolbarExStyles (BYVAL dwExStyle AS DWORD) AS STRING
   DIM strExStyle AS STRING
   IF (dwExStyle AND &h1) = &h1 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR TBSTYLE_EX_DRAWDDARROWS" ELSE strExStyle = "TBSTYLE_EX_DRAWDDARROWS"
   END IF
   IF (dwExStyle AND &h2) = &h2 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR TBSTYLE_EX_MULTICOLUMN" ELSE strExStyle = "TBSTYLE_EX_MULTICOLUMN"
   END IF
   IF (dwExStyle AND &h4) = &h4 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR TBSTYLE_EX_VERTICAL" ELSE strExStyle = "TBSTYLE_EX_VERTICAL"
   END IF
   IF (dwExStyle AND &h8) = &h8 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR TBSTYLE_EX_MIXEDBUTTONS" ELSE strExStyle = "TBSTYLE_EX_MIXEDBUTTONS"
   END IF
   IF (dwExStyle AND &h10) = &h10 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR TBSTYLE_EX_HIDECLIPPEDBUTTONS" ELSE strExStyle = "TBSTYLE_EX_HIDECLIPPEDBUTTONS"
   END IF
   IF (dwExStyle AND &h80) = &h80 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR TBSTYLE_EX_DOUBLEBUFFER" ELSE strExStyle = "TBSTYLE_EX_DOUBLEBUFFER"
   END IF
   FUNCTION = strExStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - ComboBox control.
' ========================================================================================
FUNCTION FBWS_GetComboBoxStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h0001) = &h0001 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_SIMPLE" ELSE strStyle = "CBS_SIMPLE"
   END IF
   IF (dwStyle AND &h0002) = &h0002 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_DROPDOWN" ELSE strStyle = "CBS_DROPDOWN"
   END IF
   IF (dwStyle AND &h0003) = &h0003 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_DROPDOWNLIST" ELSE strStyle = "CBS_DROPDOWNLIST"
   END IF
   IF (dwStyle AND &h0010) = &h0010 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_OWNERDRAWFIXED" ELSE strStyle = "CBS_OWNERDRAWFIXED"
   END IF
   IF (dwStyle AND &h0020) = &h0020 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_OWNERDRAWVARIABLE" ELSE strStyle = "CBS_OWNERDRAWVARIABLE"
   END IF
   IF (dwStyle AND &h0040) = &h0040 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_AUTOHSCROLL" ELSE strStyle = "CBS_AUTOHSCROLL"
   END IF
   IF (dwStyle AND &h0080) = &h0080 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_OEMCONVERT" ELSE strStyle = "CBS_OEMCONVERT"
   END IF
   IF (dwStyle AND &h0100) = &h0100 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_SORT" ELSE strStyle = "CBS_SORT"
   END IF
   IF (dwStyle AND &h0200) = &h0200 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_HASSTRINGS" ELSE strStyle = "CBS_HASSTRINGS"
   END IF
   IF (dwStyle AND &h0400) = &h0400 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_NOINTEGRALHEIGHT" ELSE strStyle = "CBS_NOINTEGRALHEIGHT"
   END IF
   IF (dwStyle AND &h0800) = &h0800 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_DISABLENOSCROLL" ELSE strStyle = "CBS_DISABLENOSCROLL"
   END IF
   IF (dwStyle AND &h2000) = &h2000 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_UPPERCASE" ELSE strStyle = "CBS_UPPERCASE"
   END IF
   IF (dwStyle AND &h4000) = &h4000 THEN
      IF LEN(strStyle) THEN strStyle += " OR CBS_LOWERCASE" ELSE strStyle = "CBS_LOWERCASE"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates extended numeric styles to constants - ComboBoxEx control.
' ========================================================================================
FUNCTION FBWS_GetComboBoxExExStyles (BYVAL dwExStyle AS DWORD) AS STRING
   DIM strExStyle AS STRING
   IF (dwExStyle AND &h1) = &h1 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR CBES_EX_NOEDITIMAGE" ELSE strExStyle = "CBES_EX_NOEDITIMAGE"
   END IF
   IF (dwExStyle AND &h2) = &h2 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR CBES_EX_NOEDITIMAGEINDENT" ELSE strExStyle = "CBES_EX_NOEDITIMAGEINDENT"
   END IF
   IF (dwExStyle AND &h4) = &h4 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR CBES_EX_PATHWORDBREAKPROC" ELSE strExStyle = "CBES_EX_PATHWORDBREAKPROC"
   END IF
   IF (dwExStyle AND &h8) = &h8 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR CBES_EX_NOSIZELIMIT" ELSE strExStyle = "CBES_EX_NOSIZELIMIT"
   END IF
   IF (dwExStyle AND &h10) = &h10 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR CBES_EX_CASESENSITIVE" ELSE strExStyle = "CBES_EX_CASESENSITIVE"
   END IF
   FUNCTION = strExStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - ListBox control.
' ========================================================================================
FUNCTION FBWS_GetListBoxStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h0001) = &h0001 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_NOTIFY" ELSE strStyle = "LBS_NOTIFY"
   END IF
   IF (dwStyle AND &h0002) = &h0002 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_SORT" ELSE strStyle = "LBS_SORT"
   END IF
   IF (dwStyle AND &h0004) = &h0004 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_NOREDRAW" ELSE strStyle = "LBS_NOREDRAW"
   END IF
   IF (dwStyle AND &h0008) = &h0008 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_MULTIPLESEL" ELSE strStyle = "LBS_MULTIPLESEL"
   END IF
   IF (dwStyle AND &h0010) = &h0010 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_OWNERDRAWFIXED" ELSE strStyle = "LBS_OWNERDRAWFIXED"
   END IF
   IF (dwStyle AND &h0020) = &h0020 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_OWNERDRAWVARIABLE" ELSE strStyle = "LBS_OWNERDRAWVARIABLE"
   END IF
   IF (dwStyle AND &h0040) = &h0040 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_HASSTRINGS" ELSE strStyle = "LBS_HASSTRINGS"
   END IF
   IF (dwStyle AND &h0080) = &h0080 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_USETABSTOPS" ELSE strStyle = "LBS_USETABSTOPS"
   END IF
   IF (dwStyle AND &h0100) = &h0100 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_NOINTEGRALHEIGHT" ELSE strStyle = "LBS_NOINTEGRALHEIGHT"
   END IF
   IF (dwStyle AND &h0200) = &h0200 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_MULTICOLUMN" ELSE strStyle = "LBS_MULTICOLUMN"
   END IF
   IF (dwStyle AND &h0400) = &h0400 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_WANTKEYBOARDINPUT" ELSE strStyle = "LBS_WANTKEYBOARDINPUT"
   END IF
   IF (dwStyle AND &h0800) = &h0800 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_EXTENDEDSEL" ELSE strStyle = "LBS_EXTENDEDSEL"
   END IF
   IF (dwStyle AND &h1000) = &h1000 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_DISABLENOSCROLL" ELSE strStyle = "LBS_DISABLENOSCROLL"
   END IF
   IF (dwStyle AND &h2000) = &h2000 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_NODATA" ELSE strStyle = "LBS_NODATA"
   END IF
   IF (dwStyle AND &h4000) = &h4000 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_NOSEL" ELSE strStyle = "LBS_NOSEL"
   END IF
   IF (dwStyle AND &h8000) = &h8000 THEN
      IF LEN(strStyle) THEN strStyle += " OR LBS_COMBOBOX" ELSE strStyle = "LBS_COMBOBOX"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - Rebar control.
' ========================================================================================
FUNCTION FBWS_GetRebarStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h100) = &h100 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBS_TOOLTIPS" ELSE strStyle = "RBS_TOOLTIPS"
   END IF
   IF (dwStyle AND &h200) = &h200 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBS_VARHEIGHT" ELSE strStyle = "RBS_VARHEIGHT"
   END IF
   IF (dwStyle AND &h400) = &h400 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBS_BANDBORDERS" ELSE strStyle = "RBS_BANDBORDERS"
   END IF
   IF (dwStyle AND &h800) = &h800 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBS_FIXEDORDER" ELSE strStyle = "RBS_FIXEDORDER"
   END IF
   IF (dwStyle AND &h1000) = &h1000 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBS_REGISTERDROP" ELSE strStyle = "RBS_REGISTERDROP"
   END IF
   IF (dwStyle AND &h2000) = &h2000 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBS_AUTOSIZE" ELSE strStyle = "RBS_AUTOSIZE"
   END IF
   IF (dwStyle AND &h4000) = &h4000 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBS_VERTICALGRIPPER" ELSE strStyle = "RBS_VERTICALGRIPPER"
   END IF
   IF (dwStyle AND &h8000) = &h8000 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBS_DBLCLKTOGGLE" ELSE strStyle = "RBS_DBLCLKTOGGLE"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - Rebar control.
' ========================================================================================
FUNCTION FBWS_GetRebarBandStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_BREAK" ELSE strStyle = "RBBS_BREAK"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_FIXEDSIZE" ELSE strStyle = "RBBS_FIXEDSIZE"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_CHILDEDGE" ELSE strStyle = "RBBS_CHILDEDGE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_HIDDEN" ELSE strStyle = "RBBS_HIDDEN"
   END IF
   IF (dwStyle AND &h10) = &h10 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_NOVERT" ELSE strStyle = "RBBS_NOVERT"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_FIXEDBMP" ELSE strStyle = "RBBS_FIXEDBMP"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_VARIABLEHEIGHT" ELSE strStyle = "RBBS_VARIABLEHEIGHT"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_GRIPPERALWAYS" ELSE strStyle = "RBBS_GRIPPERALWAYS"
   END IF
   IF (dwStyle AND &h100) = &h100 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_NOGRIPPER" ELSE strStyle = "RBBS_NOGRIPPER"
   END IF
   IF (dwStyle AND &h200) = &h200 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_USECHEVRON" ELSE strStyle = "RBBS_USECHEVRON"
   END IF
   IF (dwStyle AND &h400) = &h400 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_HIDETITLE" ELSE strStyle = "RBBS_HIDETITLE"
   END IF
   IF (dwStyle AND &h800) = &h800 THEN
      IF LEN(strStyle) THEN strStyle += " OR RBBS_TOPALIGN" ELSE strStyle = "RBBS_TOPALIGN"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - Edit control.
' ========================================================================================
FUNCTION FBWS_GetEditStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h0001) = &h0001 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_CENTER" ELSE strStyle = "ES_CENTER"
   END IF
   IF (dwStyle AND &h0002) = &h0002 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_RIGHT" ELSE strStyle = "ES_RIGHT"
   END IF
   IF (dwStyle AND &h0004) = &h0004 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_MULTILINE" ELSE strStyle = "ES_MULTILINE"
   END IF
   IF (dwStyle AND &h0008) = &h0008 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_UPPERCASE" ELSE strStyle = "ES_UPPERCASE"
   END IF
   IF (dwStyle AND &h0010) = &h0010 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_LOWERCASE" ELSE strStyle = "ES_LOWERCASE"
   END IF
   IF (dwStyle AND &h0020) = &h0020 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_PASSWORD" ELSE strStyle = "ES_PASSWORD"
   END IF
   IF (dwStyle AND &h0040) = &h0040 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_AUTOVSCROLL" ELSE strStyle = "ES_AUTOVSCROLL"
   END IF
   IF (dwStyle AND &h0080) = &h0080 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_AUTOHSCROLL" ELSE strStyle = "ES_AUTOHSCROLL"
   END IF
   IF (dwStyle AND &h0100) = &h0100 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_NOHIDESEL" ELSE strStyle = "ES_NOHIDESEL"
   END IF
   IF (dwStyle AND &h0400) = &h0400 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_OEMCONVERT" ELSE strStyle = "ES_OEMCONVERT"
   END IF
   IF (dwStyle AND &h0800) = &h0800 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_READONLY" ELSE strStyle = "ES_READONLY"
   END IF
   IF (dwStyle AND &h1000) = &h1000 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_WANTRETURN" ELSE strStyle = "ES_WANTRETURN"
   END IF
   IF (dwStyle AND &h2000) = &h2000 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_NUMBER" ELSE strStyle = "ES_NUMBER"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - RichEdit control.
' ========================================================================================
FUNCTION FBWS_GetRichEditStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h0001) = &h0001 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_CENTER" ELSE strStyle = "ES_CENTER"
   END IF
   IF (dwStyle AND &h0002) = &h0002 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_RIGHT" ELSE strStyle = "ES_RIGHT"
   END IF
   IF (dwStyle AND &h0004) = &h0004 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_MULTILINE" ELSE strStyle = "ES_MULTILINE"
   END IF
   IF (dwStyle AND &h0040) = &h0040 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_AUTOVSCROLL" ELSE strStyle = "ES_AUTOVSCROLL"
   END IF
   IF (dwStyle AND &h0080) = &h0080 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_AUTOHSCROLL" ELSE strStyle = "ES_AUTOHSCROLL"
   END IF
   IF (dwStyle AND &h0100) = &h0100 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_NOHIDESEL" ELSE strStyle = "ES_NOHIDESEL"
   END IF
   IF (dwStyle AND &h0800) = &h0800 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_READONLY" ELSE strStyle = "ES_READONLY"
   END IF
   IF (dwStyle AND &h1000) = &h1000 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_WANTRETURN" ELSE strStyle = "ES_WANTRETURN"
   END IF
   IF (dwStyle AND &h00008000) = &h00008000 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_SAVESEL" ELSE strStyle = "ES_SAVESEL"
   END IF
   IF (dwStyle AND &h00004000) = &h00004000 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_SUNKEN" ELSE strStyle = "ES_SUNKEN"
   END IF
   IF (dwStyle AND &h00002000) = &h00002000 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_DISABLENOSCROLL" ELSE strStyle = "ES_DISABLENOSCROLL"
   END IF
   IF (dwStyle AND &h01000000) = &h01000000 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_SELECTIONBAR" ELSE strStyle = "ES_SELECTIONBAR"
   END IF
   IF (dwStyle AND &h00000008) = &h00000008 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_NOOLEDRAGDROP" ELSE strStyle = "ES_NOOLEDRAGDROP"
   END IF
   IF (dwStyle AND &h00400000) = &h00400000 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_VERTICAL" ELSE strStyle = "ES_VERTICAL"
   END IF
   IF (dwStyle AND &h00080000) = &h00080000 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_NOIME" ELSE strStyle = "ES_NOIME"
   END IF
   IF (dwStyle AND &h00040000) = &h00040000 THEN
      IF LEN(strStyle) THEN strStyle += " OR ES_SELFIME" ELSE strStyle = "ES_SELFIME"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - ProgressBar control.
' ========================================================================================
FUNCTION FBWS_GetProgressBarStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR PBS_SMOOTH" ELSE strStyle = "PBS_SMOOTH"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR PBS_VERTICAL" ELSE strStyle = "PBS_VERTICAL"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - StatusBar control.
' ========================================================================================
FUNCTION FBWS_GetStatusBarStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h100) = &h100 THEN
      IF LEN(strStyle) THEN strStyle += " OR SBARS_SIZEGRIP" ELSE strStyle = "SBARS_SIZEGRIP"
   END IF
   IF (dwStyle AND &h800) = &h800 THEN
      IF LEN(strStyle) THEN strStyle += " OR SBARS_TOOLTIPS" ELSE strStyle = "SBARS_TOOLTIPS"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - TrackBar control.
' ========================================================================================
FUNCTION FBWS_GetTrackBarStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_AUTOTICKS" ELSE strStyle = "TBS_AUTOTICKS"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_VERT" ELSE strStyle = "TBS_VERT"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_TOP" ELSE strStyle = "TBS_TOP"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_BOTH" ELSE strStyle = "TBS_BOTH"
   END IF
   IF (dwStyle AND &h10) = &h10 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_NOTICKS" ELSE strStyle = "TBS_NOTICKS"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_ENABLESELRANGE" ELSE strStyle = "TBS_ENABLESELRANGE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_FIXEDLENGTH" ELSE strStyle = "TBS_FIXEDLENGTH"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_NOTHUMB" ELSE strStyle = "TBS_NOTHUMB"
   END IF
   IF (dwStyle AND &h100) = &h100 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_TOOLTIPS" ELSE strStyle = "TBS_TOOLTIPS"
   END IF
   IF (dwStyle AND &h400) = &h400 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_DOWNISLEFT" ELSE strStyle = "TBS_DOWNISLEFT"
   END IF
   IF (dwStyle AND &h800) = &h800 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_NOTIFYBEFOREMOVE" ELSE strStyle = "TBS_NOTIFYBEFOREMOVE"
   END IF
   IF (dwStyle AND &h1000) = &h1000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TBS_TRANSPARENTBKGND" ELSE strStyle = "TBS_TRANSPARENTBKGND"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - UpDown control.
' ========================================================================================
FUNCTION FBWS_GetUpDownStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR UDS_WRAP" ELSE strStyle = "UDS_WRAP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR UDS_SETBUDDYINT" ELSE strStyle = "UDS_SETBUDDYINT"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR UDS_ALIGNRIGHT" ELSE strStyle = "UDS_ALIGNRIGHT"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR UDS_ALIGNLEFT" ELSE strStyle = "UDS_ALIGNLEFT"
   END IF
   IF (dwStyle AND &h10) = &h10 THEN
      IF LEN(strStyle) THEN strStyle += " OR UDS_AUTOBUDDY" ELSE strStyle = "UDS_AUTOBUDDY"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR UDS_ARROWKEYS" ELSE strStyle = "UDS_ARROWKEYS"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR UDS_HORZ" ELSE strStyle = "UDS_HORZ"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR UDS_NOTHOUSANDS" ELSE strStyle = "UDS_NOTHOUSANDS"
   END IF
   IF (dwStyle AND &h100) = &h100 THEN
      IF LEN(strStyle) THEN strStyle += " OR UDS_HOTTRACK" ELSE strStyle = "UDS_HOTTRACK"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - ScrollBar control.
' ========================================================================================
FUNCTION FBWS_GetScrollBarStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h0001) = &h0001 THEN
      IF LEN(strStyle) THEN strStyle += " OR SBS_VERT" ELSE strStyle = "SBS_VERT"
   END IF
   IF (dwStyle AND &h0002) = &h0002 THEN
      IF LEN(strStyle) THEN strStyle += " OR SBS_TOPALIGN" ELSE strStyle = "SBS_TOPALIGN"
   END IF
   IF (dwStyle AND &h0004) = &h0004 THEN
      IF LEN(strStyle) THEN strStyle += " OR SBS_BOTTOMALIGN" ELSE strStyle = "SBS_BOTTOMALIGN"
   END IF
   IF (dwStyle AND &h0008) = &h0008 THEN
      IF LEN(strStyle) THEN strStyle += " OR SBS_SIZEBOX" ELSE strStyle = "SBS_SIZEBOX"
   END IF
   IF (dwStyle AND &h0010) = &h0010 THEN
      IF LEN(strStyle) THEN strStyle += " OR SBS_SIZEGRIP" ELSE strStyle = "SBS_SIZEGRIP"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - SysAnimate32 control.
' ========================================================================================
FUNCTION FBWS_GetSysAnimateStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR ACS_CENTER" ELSE strStyle = "ACS_CENTER"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR ACS_TRANSPARENT" ELSE strStyle = "ACS_TRANSPARENT"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR ACS_AUTOPLAY" ELSE strStyle = "ACS_AUTOPLAY"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR ACS_TIMER" ELSE strStyle = "ACS_TIMER"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - DateTimePicker control.
' ========================================================================================
FUNCTION FBWS_GetDateTimePickerStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR DTS_UPDOWN" ELSE strStyle = "DTS_UPDOWN"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR DTS_SHOWNONE" ELSE strStyle = "DTS_SHOWNONE"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR DTS_LONGDATEFORMAT" ELSE strStyle = "DTS_LONGDATEFORMAT"
   END IF
   IF (dwStyle AND &hc) = &hc THEN
      IF LEN(strStyle) THEN strStyle += " OR DTS_SHORTDATECENTURYFORMAT" ELSE strStyle = "DTS_SHORTDATECENTURYFORMAT"
   END IF
   IF (dwStyle AND &h9) = &h9 THEN
      IF LEN(strStyle) THEN strStyle += " OR DTS_TIMEFORMAT" ELSE strStyle = "DTS_TIMEFORMAT"
   END IF
   IF (dwStyle AND &h10) = &h10 THEN
      IF LEN(strStyle) THEN strStyle += " OR DTS_APPCANPARSE" ELSE strStyle = "DTS_APPCANPARSE"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR DTS_RIGHTALIGN" ELSE strStyle = "DTS_RIGHTALIGN"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - SysHeader32 control.
' ========================================================================================
FUNCTION FBWS_GetSysHeaderStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR HDS_BUTTONS" ELSE strStyle = "HDS_BUTTONS"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR HDS_HOTTRACK" ELSE strStyle = "HDS_HOTTRACK"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR HDS_HIDDEN" ELSE strStyle = "HDS_HIDDEN"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR HDS_DRAGDROP" ELSE strStyle = "HDS_DRAGDROP"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR HDS_FULLDRAG" ELSE strStyle = "HDS_FULLDRAG"
   END IF
   IF (dwStyle AND &h100) = &h100 THEN
      IF LEN(strStyle) THEN strStyle += " OR HDS_FILTERBAR" ELSE strStyle = "HDS_FILTERBAR"
   END IF
   IF (dwStyle AND &h200) = &h200 THEN
      IF LEN(strStyle) THEN strStyle += " OR HDS_FLAT" ELSE strStyle = "HDS_FLAT"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - SysIPAddress32 control.
' ========================================================================================
FUNCTION FBWS_GetIPAddressStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - SysMonthCal32 control.
' ========================================================================================
FUNCTION FBWS_GetMonthCalStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR MCS_DAYSTATE" ELSE strStyle = "MCS_DAYSTATE"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR MCS_MULTISELECT" ELSE strStyle = "MCS_MULTISELECT"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR MCS_WEEKNUMBERS" ELSE strStyle = "MCS_WEEKNUMBERS"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR MCS_NOTODAYCIRCLE" ELSE strStyle = "MCS_NOTODAYCIRCLE"
   END IF
   IF (dwStyle AND &h10) = &h10 THEN
      IF LEN(strStyle) THEN strStyle += " OR MCS_NOTODAY" ELSE strStyle = "MCS_NOTODAY"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - SysPager32 control.
' ========================================================================================
FUNCTION FBWS_GetSysPagerStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR PGS_HORZ" ELSE strStyle = "PGS_HORZ"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR PGS_AUTOSCROLL" ELSE strStyle = "PGS_AUTOSCROLL"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR PGS_DRAGNDROP" ELSE strStyle = "PGS_DRAGNDROP"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - SysTabControl32 control.
' ========================================================================================
FUNCTION FBWS_GetSysTabStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_SCROLLOPPOSITE" ELSE strStyle = "TCS_SCROLLOPPOSITE"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_BOTTOM" ELSE strStyle = "TCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_MULTISELECT" ELSE strStyle = "TCS_MULTISELECT"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_FLATBUTTONS" ELSE strStyle = "TCS_FLATBUTTONS"
   END IF
   IF (dwStyle AND &h10) = &h10 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_FORCEICONLEFT" ELSE strStyle = "TCS_FORCEICONLEFT"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_FORCELABELLEFT" ELSE strStyle = "TCS_FORCELABELLEFT"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_HOTTRACK" ELSE strStyle = "TCS_HOTTRACK"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_VERTICAL" ELSE strStyle = "TCS_VERTICAL"
   END IF
   IF (dwStyle AND &h100) = &h100 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_BUTTONS" ELSE strStyle = "TCS_BUTTONS"
   END IF
   IF (dwStyle AND &h200) = &h200 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_MULTILINE" ELSE strStyle = "TCS_MULTILINE"
   END IF
   IF (dwStyle AND &h400) = &h400 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_FIXEDWIDTH" ELSE strStyle = "TCS_FIXEDWIDTH"
   END IF
   IF (dwStyle AND &h800) = &h800 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_RAGGEDRIGHT" ELSE strStyle = "TCS_RAGGEDRIGHT"
   END IF
   IF (dwStyle AND &h1000) = &h1000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_FOCUSONBUTTONDOWN" ELSE strStyle = "TCS_FOCUSONBUTTONDOWN"
   END IF
   IF (dwStyle AND &h2000) = &h2000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_OWNERDRAWFIXED" ELSE strStyle = "TCS_OWNERDRAWFIXED"
   END IF
   IF (dwStyle AND &h4000) = &h4000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_TOOLTIPS" ELSE strStyle = "TCS_TOOLTIPS"
   END IF
   IF (dwStyle AND &h8000) = &h8000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TCS_FOCUSNEVER" ELSE strStyle = "TCS_FOCUSNEVER"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric extended styles to constants - SysTabControl32 control.
' ========================================================================================
FUNCTION FBWS_GetSysTabExStyles (BYVAL dwExStyle AS DWORD) AS STRING
   DIM strExStyle AS STRING
   IF (dwExStyle AND &h1) = &h1 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR TCS_EX_FLATSEPARATORS" ELSE strExStyle = "TCS_EX_FLATSEPARATORS"
   END IF
   IF (dwExStyle AND &h2) = &h2 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR TCS_EX_REGISTERDROP" ELSE strExStyle = "TCS_EX_REGISTERDROP"
   END IF
   FUNCTION = strExStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - TreeView control.
' ========================================================================================
FUNCTION FBWS_GetTreeViewStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_HASBUTTONS" ELSE strStyle = "TVS_HASBUTTONS"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_HASLINES" ELSE strStyle = "TVS_HASLINES"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_LINESATROOT" ELSE strStyle = "TVS_LINESATROOT"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_EDITLABELS" ELSE strStyle = "TVS_EDITLABELS"
   END IF
   IF (dwStyle AND &h10) = &h10 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_DISABLEDRAGDROP" ELSE strStyle = "TVS_DISABLEDRAGDROP"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_SHOWSELALWAYS" ELSE strStyle = "TVS_SHOWSELALWAYS"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_RTLREADING" ELSE strStyle = "TVS_RTLREADING"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_NOTOOLTIPS" ELSE strStyle = "TVS_NOTOOLTIPS"
   END IF
   IF (dwStyle AND &h100) = &h100 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_CHECKBOXES" ELSE strStyle = "TVS_CHECKBOXES"
   END IF
   IF (dwStyle AND &h200) = &h200 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_TRACKSELECT" ELSE strStyle = "TVS_TRACKSELECT"
   END IF
   IF (dwStyle AND &h400) = &h400 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_SINGLEEXPAND" ELSE strStyle = "TVS_SINGLEEXPAND"
   END IF
   IF (dwStyle AND &h800) = &h800 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_INFOTIP" ELSE strStyle = "TVS_INFOTIP"
   END IF
   IF (dwStyle AND &h1000) = &h1000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_FULLROWSELECT" ELSE strStyle = "TVS_FULLROWSELECT"
   END IF
   IF (dwStyle AND &h2000) = &h2000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_NOSCROLL" ELSE strStyle = "TVS_NOSCROLL"
   END IF
   IF (dwStyle AND &h4000) = &h4000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_NONEVENHEIGHT" ELSE strStyle = "TVS_NONEVENHEIGHT"
   END IF
   IF (dwStyle AND &h8000) = &h8000 THEN
      IF LEN(strStyle) THEN strStyle += " OR TVS_NOHSCROLL" ELSE strStyle = "TVS_NOHSCROLL"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric extended styles to constants - TreeView control.
' ========================================================================================
FUNCTION FBWS_GetTreeViewExStyles (BYVAL dwExStyle AS DWORD) AS STRING
   DIM strExStyle AS STRING
   IF (dwExStyle AND &h1) = &h1 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR TVS_EX_NOSINGLECOLLAPSE" ELSE strExStyle = "TVS_EX_NOSINGLECOLLAPSE"
   END IF
   FUNCTION = strExStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric styles to constants - ListView control.
' ========================================================================================
FUNCTION FBWS_GetListViewStyles (BYVAL dwStyle AS DWORD) AS STRING
   DIM strStyle AS STRING
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_REPORT" ELSE strStyle = "LVS_REPORT"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_SMALLICON" ELSE strStyle = "LVS_SMALLICON"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_LIST" ELSE strStyle = "LVS_LIST"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_SINGLESEL" ELSE strStyle = "LVS_SINGLESEL"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_SHOWSELALWAYS" ELSE strStyle = "LVS_SHOWSELALWAYS"
   END IF
   IF (dwStyle AND &h10) = &h10 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_SORTASCENDING" ELSE strStyle = "LVS_SORTASCENDING"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_SORTDESCENDING" ELSE strStyle = "LVS_SORTDESCENDING"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_SHAREIMAGELISTS" ELSE strStyle = "LVS_SHAREIMAGELISTS"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_NOLABELWRAP" ELSE strStyle = "LVS_NOLABELWRAP"
   END IF
   IF (dwStyle AND &h100) = &h100 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_AUTOARRANGE" ELSE strStyle = "LVS_AUTOARRANGE"
   END IF
   IF (dwStyle AND &h200) = &h200 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_EDITLABELS" ELSE strStyle = "LVS_EDITLABELS"
   END IF
   IF (dwStyle AND &h1000) = &h1000 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_OWNERDATA" ELSE strStyle = "LVS_OWNERDATA"
   END IF
   IF (dwStyle AND &h2000) = &h2000 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_NOSCROLL" ELSE strStyle = "LVS_NOSCROLL"
   END IF
   IF (dwStyle AND &hfc00) = &hfc00 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_TYPESTYLEMASK" ELSE strStyle = "LVS_TYPESTYLEMASK"
   END IF
   IF (dwStyle AND &h800) = &h800 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_ALIGNLEFT" ELSE strStyle = "LVS_ALIGNLEFT"
   END IF
   IF (dwStyle AND &hc00) = &hc00 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_ALIGNMASK" ELSE strStyle = "LVS_ALIGNMASK"
   END IF
   IF (dwStyle AND &h400) = &h400 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_OWNERDRAWFIXED" ELSE strStyle = "LVS_OWNERDRAWFIXED"
   END IF
   IF (dwStyle AND &h4000) = &h4000 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_NOCOLUMNHEADER" ELSE strStyle = "LVS_NOCOLUMNHEADER"
   END IF
   IF (dwStyle AND &h8000) = &h8000 THEN
      IF LEN(strStyle) THEN strStyle += " OR LVS_NOSORTHEADER" ELSE strStyle = "LVS_NOSORTHEADER"
   END IF
   ' // Common control styles
   IF (dwStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_TOP" ELSE strStyle = "CCS_TOP"
   END IF
   IF (dwStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOMOVEY" ELSE strStyle = "CCS_NOMOVEY"
   END IF
   IF (dwStyle AND &h3) = &h3 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_BOTTOM" ELSE strStyle = "CCS_BOTTOM"
   END IF
   IF (dwStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NORESIZE" ELSE strStyle = "CCS_NORESIZE"
   END IF
   IF (dwStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NOPARENTALIGN" ELSE strStyle = "CCS_NOPARENTALIGN"
   END IF
   IF (dwStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_ADJUSTABLE" ELSE strStyle = "CCS_ADJUSTABLE"
   END IF
   IF (dwStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_NODIVIDER" ELSE strStyle = "CCS_NODIVIDER"
   END IF
   IF (dwStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR CCS_VERT" ELSE strStyle = "CCS_VERT"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates numeric extended styles to constants - ListView control.
' ========================================================================================
FUNCTION FBWS_GetListViewExStyles (BYVAL dwExStyle AS DWORD) AS STRING
   DIM strExStyle AS STRING
   IF (dwExStyle AND &h1) = &h1 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_GRIDLINES" ELSE strExStyle = "LVS_EX_GRIDLINES"
   END IF
   IF (dwExStyle AND &h2) = &h2 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_SUBITEMIMAGES" ELSE strExStyle = "LVS_EX_SUBITEMIMAGES"
   END IF
   IF (dwExStyle AND &h4) = &h4 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_CHECKBOXES" ELSE strExStyle = "LVS_EX_CHECKBOXES"
   END IF
   IF (dwExStyle AND &h8) = &h8 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_TRACKSELECT" ELSE strExStyle = "LVS_EX_TRACKSELECT"
   END IF
   IF (dwExStyle AND &h10) = &h10 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_HEADERDRAGDROP" ELSE strExStyle = "LVS_EX_HEADERDRAGDROP"
   END IF
   IF (dwExStyle AND &h20) = &h20 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_FULLROWSELECT" ELSE strExStyle = "LVS_EX_FULLROWSELECT"
   END IF
   IF (dwExStyle AND &h40) = &h40 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_ONECLICKACTIVATE" ELSE strExStyle = "LVS_EX_ONECLICKACTIVATE"
   END IF
   IF (dwExStyle AND &h80) = &h80 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_TWOCLICKACTIVATE" ELSE strExStyle = "LVS_EX_TWOCLICKACTIVATE"
   END IF
   IF (dwExStyle AND &h100) = &h100 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_FLATSB" ELSE strExStyle = "LVS_EX_FLATSB"
   END IF
   IF (dwExStyle AND &h200) = &h200 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_REGIONAL" ELSE strExStyle = "LVS_EX_REGIONAL"
   END IF
   IF (dwExStyle AND &h400) = &h400 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_INFOTIP" ELSE strExStyle = "LVS_EX_INFOTIP"
   END IF
   IF (dwExStyle AND &h800) = &h800 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_UNDERLINEHOT" ELSE strExStyle = "LVS_EX_UNDERLINEHOT"
   END IF
   IF (dwExStyle AND &h1000) = &h1000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_UNDERLINECOLD" ELSE strExStyle = "LVS_EX_UNDERLINECOLD"
   END IF
   IF (dwExStyle AND &h2000) = &h2000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_MULTIWORKAREAS" ELSE strExStyle = "LVS_EX_MULTIWORKAREAS"
   END IF
   IF (dwExStyle AND &h4000) = &h4000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_LABELTIP" ELSE strExStyle = "LVS_EX_LABELTIP"
   END IF
   IF (dwExStyle AND &h8000) = &h8000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_BORDERSELECT" ELSE strExStyle = "LVS_EX_BORDERSELECT"
   END IF
   IF (dwExStyle AND &h10000) = &h10000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_DOUBLEBUFFER" ELSE strExStyle = "LVS_EX_DOUBLEBUFFER"
   END IF
   IF (dwExStyle AND &h20000) = &h20000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_HIDELABELS" ELSE strExStyle = "LVS_EX_HIDELABELS"
   END IF
   IF (dwExStyle AND &h40000) = &h40000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_SINGLEROW" ELSE strExStyle = "LVS_EX_SINGLEROW"
   END IF
   IF (dwExStyle AND &h80000) = &h80000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_SNAPTOGRID" ELSE strExStyle = "LVS_EX_SNAPTOGRID"
   END IF
   IF (dwExStyle AND &h100000) = &h100000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_SIMPLESELECT" ELSE strExStyle = "LVS_EX_SIMPLESELECT"
   END IF
   IF (dwExStyle AND &h200000) = &h200000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_JUSTIFYCOLUMNS" ELSE strExStyle = "LVS_EX_JUSTIFYCOLUMNS"
   END IF
   IF (dwExStyle AND &h400000) = &h400000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_TRANSPARENTBKGND" ELSE strExStyle = "LVS_EX_TRANSPARENTBKGND"
   END IF
   IF (dwExStyle AND &h800000) = &h800000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_TRANSPARENTSHADOWTEXT" ELSE strExStyle = "LVS_EX_TRANSPARENTSHADOWTEXT"
   END IF
   IF (dwExStyle AND &h1000000) = &h1000000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_AUTOAUTOARRANGE" ELSE strExStyle = "LVS_EX_AUTOAUTOARRANGE"
   END IF
   IF (dwExStyle AND &h2000000) = &h2000000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_HEADERINALLVIEWS" ELSE strExStyle = "LVS_EX_HEADERINALLVIEWS"
   END IF
   IF (dwExStyle AND &h8000000) = &h8000000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_AUTOCHECKSELECT" ELSE strExStyle = "LVS_EX_AUTOCHECKSELECT"
   END IF
   IF (dwExStyle AND &h10000000) = &h10000000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_AUTOSIZECOLUMNS" ELSE strExStyle = "LVS_EX_AUTOSIZECOLUMNS"
   END IF
   IF (dwExStyle AND &h40000000) = &h40000000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_COLUMNSNAPPOINTS" ELSE strExStyle = "LVS_EX_COLUMNSNAPPOINTS"
   END IF
   IF (dwExStyle AND &h80000000) = &h80000000 THEN
      IF LEN(strExStyle) THEN strExStyle += " OR LVS_EX_COLUMNOVERFLOW" ELSE strExStyle = "LVS_EX_COLUMNOVERFLOW"
   END IF
   FUNCTION = strExStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Get ToolBar button state
' ========================================================================================
FUNCTION FBWS_GetToolBarButtonState (BYVAL fsState AS UBYTE) AS STRING
   DIM strState AS STRING
   IF (fsState AND &h1) = &h1 THEN
      IF LEN(strState) THEN strState += " OR TBSTATE_CHECKED" ELSE strState = "TBSTATE_CHECKED"
   END IF
   IF (fsState AND &h2) = &h2 THEN
      IF LEN(strState) THEN strState += " OR TBSTATE_PRESSED" ELSE strState = "TBSTATE_PRESSED"
   END IF
   IF (fsState AND &h4) = &h4 THEN
      IF LEN(strState) THEN strState += " OR TBSTATE_ENABLED" ELSE strState = "TBSTATE_ENABLED"
   END IF
   IF (fsState AND &h8) = &h8 THEN
      IF LEN(strState) THEN strState += " OR TBSTATE_HIDDEN" ELSE strState = "TBSTATE_HIDDEN"
   END IF
   IF (fsState AND &h10) = &h10 THEN
      IF LEN(strState) THEN strState += " OR TBSTATE_INDETERMINATE" ELSE strState = "TBSTATE_INDETERMINATE"
   END IF
   IF (fsState AND &h20) = &h20 THEN
      IF LEN(strState) THEN strState += " OR TBSTATE_WRAP" ELSE strState = "TBSTATE_WRAP"
   END IF
   IF (fsState AND &h40) = &h40 THEN
      IF LEN(strState) THEN strState += " OR TBSTATE_ELLIPSES" ELSE strState = "TBSTATE_ELLIPSES"
   END IF
   IF (fsState AND &h80) = &h80 THEN
      IF LEN(strState) THEN strState += " OR TBSTATE_MARKED" ELSE strState = "TBSTATE_MARKED"
   END IF
   FUNCTION = strState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Get ToolBar button style
' ========================================================================================
FUNCTION FBWS_GetToolBarButtonStyle (BYVAL fsStyle AS UBYTE) AS STRING
   DIM strStyle AS STRING = "BTNS_BUTTON"
   IF (fsStyle AND &h1) = &h1 THEN
      IF LEN(strStyle) THEN strStyle += " OR BTNS_SEP" ELSE strStyle = "BTNS_SEP"
   END IF
   IF (fsStyle AND &h2) = &h2 THEN
      IF LEN(strStyle) THEN strStyle += " OR BTNS_CHECK" ELSE strStyle = "BTNS_CHECK"
   END IF
   IF (fsStyle AND &h4) = &h4 THEN
      IF LEN(strStyle) THEN strStyle += " OR BTNS_GROUP" ELSE strStyle = "BTNS_GROUP"
   END IF
   IF (fsStyle AND &h8) = &h8 THEN
      IF LEN(strStyle) THEN strStyle += " OR BTNS_DROPDOWN" ELSE strStyle = "BTNS_DROPDOWN"
   END IF
   IF (fsStyle AND &h10) = &h10 THEN
      IF LEN(strStyle) THEN strStyle += " OR BTNS_AUTOSIZE" ELSE strStyle = "BTNS_AUTOSIZE"
   END IF
   IF (fsStyle AND &h20) = &h20 THEN
      IF LEN(strStyle) THEN strStyle += " OR BTNS_NOPREFIX" ELSE strStyle = "BTNS_NOPREFIX"
   END IF
   IF (fsStyle AND &h40) = &h40 THEN
      IF LEN(strStyle) THEN strStyle += " OR BTNS_SHOWTEXT" ELSE strStyle = "TBSTYLE_TOOLTIPS"
   END IF
   IF (fsStyle AND &h80) = &h80 THEN
      IF LEN(strStyle) THEN strStyle += " OR BTNS_WHOLEDROPDOWN" ELSE strStyle = "BTNS_WHOLEDROPDOWN"
   END IF
   FUNCTION = strStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Get cursor type
' ========================================================================================
FUNCTION FBWS_GetCursorType (BYVAL hCursor AS HCURSOR) AS STRING
   IF hCursor = LoadCursor(NULL, IDC_ARROW) THEN RETURN "IDC_ARROW"
   IF hCursor = LoadCursor(NULL, IDC_IBEAM) THEN RETURN "IDC_IBEAM"
   IF hCursor = LoadCursor(NULL, IDC_WAIT) THEN RETURN "IDC_WAIT"
   IF hCursor = LoadCursor(NULL, IDC_CROSS) THEN RETURN "IDC_CROSS"
   IF hCursor = LoadCursor(NULL, IDC_UPARROW) THEN RETURN "IDC_UPARROW"
   IF hCursor = LoadCursor(NULL, IDC_SIZE) THEN RETURN "IDC_SIZE"
   IF hCursor = LoadCursor(NULL, IDC_ICON) THEN RETURN "IDC_ICON"
   IF hCursor = LoadCursor(NULL, IDC_SIZENWSE) THEN RETURN "IDC_SIZENWSE"
   IF hCursor = LoadCursor(NULL, IDC_SIZENWSE) THEN RETURN "IDC_SIZENWSE"
   IF hCursor = LoadCursor(NULL, IDC_SIZENESW) THEN RETURN "IDC_SIZENESW"
   IF hCursor = LoadCursor(NULL, IDC_SIZEWE) THEN RETURN "IDC_SIZEWE"
   IF hCursor = LoadCursor(NULL, IDC_SIZENS) THEN RETURN "IDC_SIZENS"
   IF hCursor = LoadCursor(NULL, IDC_SIZEALL) THEN RETURN "IDC_SIZEALL"
   IF hCursor = LoadCursor(NULL, IDC_NO) THEN RETURN "IDC_NO"
   IF hCursor = LoadCursor(NULL, IDC_HAND) THEN RETURN "IDC_HAND"
   IF hCursor = LoadCursor(NULL, IDC_APPSTARTING) THEN RETURN "IDC_APPSTARTING"
   IF hCursor = LoadCursor(NULL, IDC_HELP) THEN RETURN "IDC_HELP"
END FUNCTION
' ========================================================================================

' ========================================================================================
' Get brush type
' ========================================================================================
FUNCTION FBWS_GetBrushType (BYVAL hBrush AS HBRUSH) AS STRING
   SELECT CASE hBrush
      ' // First search by COLOR_xxx value
      CASE COLOR_WINDOW + 1 : RETURN "COLOR_WINDOW"
      CASE COLOR_WINDOWTEXT + 1 : RETURN "COLOR_WINDOWTEXT"
      CASE COLOR_WINDOWFRAME + 1 : RETURN "COLOR_WINDOWFRAME"
      CASE COLOR_3DDKSHADOW + 1 : RETURN "COLOR_3DDKSHADOW"
      CASE COLOR_3DFACE + 1 : RETURN "COLOR_3DFACE"
      CASE COLOR_3DHILIGHT + 1 : RETURN "COLOR_3DHILIGHT"
      CASE COLOR_3DLIGHT + 1 : RETURN "COLOR_3DLIGHT"
      CASE COLOR_3DSHADOW + 1 : RETURN "COLOR_3DSHADOW"
      CASE COLOR_ACTIVEBORDER + 1 : RETURN "COLOR_ACTIVEBORDER"
      CASE COLOR_ACTIVECAPTION + 1 : RETURN "COLOR_ACTIVECAPTION"
      CASE COLOR_APPWORKSPACE + 1 : RETURN "COLOR_APPWORKSPACE"
      CASE COLOR_DESKTOP + 1 : RETURN "COLOR_DESKTOP"
      CASE COLOR_BTNTEXT + 1 : RETURN "COLOR_BTNTEXT"
      CASE COLOR_CAPTIONTEXT + 1 :RETURN "COLOR_CAPTIONTEXT"
      CASE COLOR_GRAYTEXT + 1 : RETURN "COLOR_GRAYTEXT"
      CASE COLOR_HIGHLIGHT + 1 : RETURN "COLOR_HIGHLIGHT"
      CASE COLOR_HIGHLIGHTTEXT + 1 : RETURN "COLOR_HIGHLIGHTTEXT"
      CASE GetSysColor(COLOR_INACTIVEBORDER) : RETURN "COLOR_INACTIVEBORDER"
      CASE GetSysColor(COLOR_INACTIVECAPTION) : RETURN "COLOR_INACTIVECAPTION"
      CASE GetSysColor(COLOR_INACTIVECAPTIONTEXT) : RETURN "COLOR_INACTIVECAPTIONTEXT"
      CASE COLOR_INFOBK + 1 : RETURN "COLOR_INFOBK"
      CASE COLOR_INFOTEXT + 1 : RETURN "COLOR_INFOTEXT"
      CASE COLOR_MENU + 1 : RETURN "COLOR_MENU"
      CASE COLOR_MENUTEXT + 1 : RETURN "COLOR_MENUTEXT"
      CASE COLOR_SCROLLBAR + 1 : RETURN "COLOR_SCROLLBAR"
      ' // Search by handle value
      CASE GetStockObject(NULL_BRUSH) : RETURN "GetStockObject(NULL_BRUSH)"
      CASE GetStockObject(BLACK_BRUSH) : RETURN "GetStockObject(BLACK_BRUSH)"
      CASE GetStockObject(DKGRAY_BRUSH) : RETURN "GetStockObject(DKGRAY_BRUSH)"
      CASE GetStockObject(GRAY_BRUSH) : RETURN "GetStockObject(GRAY_BRUSH)"
      CASE GetStockObject(LTGRAY_BRUSH) : RETURN "GetStockObject(LTGRAY_BRUSH)"
      CASE GetStockObject(WHITE_BRUSH) : RETURN "GetStockObject(WHITE_BRUSH)"
      CASE GetSysColorBrush(COLOR_SCROLLBAR) : RETURN "GetSysColorBrush(COLOR_SCROLLBAR)"
      CASE GetSysColorBrush(COLOR_DESKTOP) : RETURN "GetSysColorBrush(COLOR_DESKTOP)"
      CASE GetSysColorBrush(COLOR_ACTIVECAPTION) : RETURN "GetSysColorBrush(COLOR_ACTIVECAPTION)"
      CASE GetSysColorBrush(COLOR_INACTIVECAPTION) : RETURN "GetSysColorBrush(COLOR_INACTIVECAPTION)"
      CASE GetSysColorBrush(COLOR_MENU) : RETURN "GetSysColorBrush(COLOR_MENU)"
      CASE GetSysColorBrush(COLOR_WINDOW) : RETURN "GetSysColorBrush(COLOR_WINDOW)"
      CASE GetSysColorBrush(COLOR_WINDOWFRAME) : RETURN "GetSysColorBrush(COLOR_WINDOWFRAME)"
      CASE GetSysColorBrush(COLOR_MENUTEXT) : RETURN "GetSysColorBrush(COLOR_MENUTEXT)"
      CASE GetSysColorBrush(COLOR_WINDOWTEXT) : RETURN "GetSysColorBrush(COLOR_WINDOWTEXT)"
      CASE GetSysColorBrush(COLOR_CAPTIONTEXT) : RETURN "GetSysColorBrush(COLOR_CAPTIONTEXT)"
      CASE GetSysColorBrush(COLOR_ACTIVEBORDER) : RETURN "GetSysColorBrush(COLOR_ACTIVEBORDER)"
      CASE GetSysColorBrush(COLOR_INACTIVEBORDER) : RETURN "GetSysColorBrush(COLOR_INACTIVEBORDER)"
      CASE GetSysColorBrush(COLOR_APPWORKSPACE) : RETURN "GetSysColorBrush(COLOR_APPWORKSPACE)"
      CASE GetSysColorBrush(COLOR_HIGHLIGHT) : RETURN "GetSysColorBrush(COLOR_HIGHLIGHT)"
      CASE GetSysColorBrush(COLOR_HIGHLIGHTTEXT) : RETURN "GetSysColorBrush(COLOR_HIGHLIGHTTEXT)"
      CASE GetSysColorBrush(COLOR_3DFACE) : RETURN "GetSysColorBrush(COLOR_3DFACE)"
      CASE GetSysColorBrush(COLOR_3DSHADOW) : RETURN "GetSysColorBrush(COLOR_3DSHADOW)"
      CASE GetSysColorBrush(COLOR_GRAYTEXT) : RETURN "GetSysColorBrush(COLOR_GRAYTEXT)"
      CASE GetSysColorBrush(COLOR_BTNTEXT) : RETURN "GetSysColorBrush(COLOR_BTNTEXT)"
      CASE GetSysColorBrush(COLOR_INACTIVECAPTIONTEXT) : RETURN "GetSysColorBrush(COLOR_INACTIVECAPTIONTEXT)"
      CASE GetSysColorBrush(COLOR_3DHIGHLIGHT) : RETURN "GetSysColorBrush(COLOR_3DHIGHLIGHT)"
      CASE GetSysColorBrush(COLOR_3DDKSHADOW) : RETURN "GetSysColorBrush(COLOR_3DDKSHADOW)"
      CASE GetSysColorBrush(COLOR_3DLIGHT) : RETURN "GetSysColorBrush(COLOR_3DLIGHT)"
      CASE GetSysColorBrush(COLOR_INFOTEXT) : RETURN "GetSysColorBrush(COLOR_INFOTEXT)"
      CASE GetSysColorBrush(COLOR_INFOBK) : RETURN "GetSysColorBrush(COLOR_INFOBK)"

'  CASE ELSE
'      RETURN "BGR(" + FORMAT$((iClr AND &h0000FF&)) + ", " + _           ' red component
'                      FORMAT$((iClr AND &h00FF00&) \ &h100&) + ", " + _  ' green component
'                      FORMAT$((iClr AND &hFF0000&) \ &h10000&) + ")"     ' blue component

   END SELECT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the font used by the specified window.
' Parameters:
' - hwnd = Handle to the window.
' - lplfw = Pointer to a LOGFONTW stucture.
' Return value:
'   If the function succeeds, and lplfw is a valid pointer, the return value is the number
'   of bytes stored into the buffer.
'   If the function succeeds, and lplfw is NULL, the return value is the number of bytes
'   required to hold the information the function would store into the buffer.
'   If the function fails, the return value is zero.
' Remarks:
'    WM_GETFONT works fine cross-process, provided that the window you're sending it to
'    responds to this message. This is only guaranteed for standard Windows controls. If the
'    standard Windows control is using the system font, it returns NULL.
'    The WM_GETFONT message does not return a font handle if the message is sent to a dialog
'    box created by the DialogBoxParam, DialogBoxIndirectParam, CreateDialogParam, or
'    CreateDialogIndirectParam function.
'    A window can use an arbitrary font for rendering, or even multiple fonts.
' ========================================================================================
FUNCTION FBWS_GetFont (BYVAL hwnd AS HWND, BYVAL lplfw AS LOGFONTW PTR) AS LONG
   IF IsWindow(hwnd) = FALSE THEN EXIT FUNCTION
   DIM hFont AS HFONT = CAST(HFONT, SendMessageW(hwnd, WM_GETFONT, 0, 0))
   IF hFont = 0 THEN EXIT FUNCTION
   FUNCTION = GetObjectW(hFont, SIZEOF(LOGFONTW), lplfw)
END FUNCTION
' ========================================================================================
