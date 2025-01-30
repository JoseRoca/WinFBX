# CRichEditCtrl Class

Provides the functionality of the rich edit control.

A "rich edit control" is a window in which the user can enter and edit text. The text can be assigned character and paragraph formatting, and can include embedded OLE objects. Rich edit controls provide a programming interface for formatting text. However, an application must implement any user interface components necessary to make formatting operations available to the user.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTOR](#CONSTRUCTOR) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [hRichEdit](#hRichEdit) | Returns the handle of the rich edit control. |

# <a name="CONSTRUCTOR"></a>CONSTRUCTOR

Creates an instance of the rich edit control.

```
CONSTRUCTOR CRichEditCtrl (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszTitle AS WSTRING = "", _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
```
| Parameter | Description |
| --------- | ----------- |
| *pWindow* | Pointer to the parent `CWindow`class. |
| *cID* | The control identifier, an integer value used to notify its parent about events. The application determines the control identifier; it must be unique for all controls with the same parent window. |
| *wszTitle* | Optional. The text of the control. |
| *x* | Optional. The x-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *y* | Optional. The initial y-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *nWidth* | Optional. The width of the control. |
| *nHeight* | Optional. The height of the control. |
| *dwStyle* | Optional. The window styles of the control being created.<br>Default styles: WS_VISIBLE OR WS_CHILD OR WS_TABSTOP OR ES_LEFT OR WS_HSCROLL OR WS_VSCROLL OR ES_AUTOHSCROLL OR ES_AUTOVSCROLL OR ES_MULTILINE OR ES_WANTRETURN OR ES_NOHIDESEL OR ES_SAVESEL |
| *dwExStyle* | Optional. The extended window styles of the control being created.<br>Default extended style: WS_EX_CLIENTEDGE |
| *lParam* | Optional. Pointer to custom data. |

# <a name="DESTRUCTOR"></a>DESTRUCTOR

Called automatically when a class variable goes out of scope or is destroyed.

# <a name="hRichEdit"></a>hRichEdit

Returns the handle of the rich edit control.

### CRichEditCtrl Properties

| Name       | Description |
| ---------- | ----------- |
| [AutoCorrectProc (Get)](#AutoCorrectProc1) | Gets a pointer to the application-defined AutoCorrectProc callback function. |
| [AutoCorrectProc (Set)](#AutoCorrectProc2) | Sets a pointer to the application-defined AutoCorrectProc callback function. |
| [AutoUrlDetect (Get)](#AutoUrlDetect1) | Indicates whether the auto URL detection is turned on in the rich edit control. |
| [AutoUrlDetect (Set)](#AutoUrlDetect2) | Enables or disables automatic detection of URLs by a rich edit control. |
| [BidiOptions (Get)](#BidiOptions1) | Indicates the current state of the bidirectional options in the rich edit control. |
| [BidiOptions (Set)](#BidiOptions2) | Sets the current state of the bidirectional options in the rich edit control. |
| [CharFormat (Get)](#CharFormat1) | Determines the current character formatting in a rich edit control. |
| [CharFormat (Set)](#CharFormat2) | Sets character formatting in a rich edit control. |
| [CTFModeBias (Get)](#CTFModeBias1) | Gets the Text Services Framework mode bias values for a Microsoft Rich Edit control. |
| [CTFModeBias (Set)](#CTFModeBias2) | Sets the Text Services Framework mode bias values for a Microsoft Rich Edit control. |
| [CTFOpenStatus (Get)](#CTFOpenStatus1) | Determines if the Text Services Framework (TSF) keyboard is open or closed. |
| [CTFOpenStatus (Set)](#CTFOpenStatus2) | Opens or closes the Text Services Framework (TSF) keyboard. |
| [EditStyle (Get)](#EditStyle1) | Retrieves the current edit style flags. |
| [EditStyle (Set)](#EditStyle2) | Sets the current edit style flags. |
| [EditStyleEx (Get)](#EditStyleEx1) | Returns the extended edit style flags. |
| [EditStyleEx (Set)](#EditStyleEx2) | Sets the current extended edit style flags. |
| [EllipsisMode (Get)](#EllipsisMode1) | Retrieves the current ellipsis mode. |
| [EllipsisMode (Set)](#EllipsisMode2) | Sets the current ellipsis mode. |
| [EventMask (Get)](#EventMask1) | Retrieves the event mask for a rich edit control. The event mask specifies which notification messages the control sends to its parent window. |
| [EventMask (Set)](#EventMask2) | Sets the event mask for a rich edit control. |
| [HyphenateInfo (Get)](#HyphenateInfo1) | Gets information about hyphenation for a Microsoft Rich Edit control. |
| [HyphenateInfo (Set)](#HyphenateInfo2) | Sets the way a Microsoft Rich Edit control does hyphenation.|
| [IMEColor (Get)](#IMEColor1) | Retrieves the Input Method Editor (IME) composition color. This message is available only in Asian-language versions of the operating system. |
| [IMEColor (Set)](#IMEColor2) | Sets the Input Method Editor (IME) composition color. |
| [IMEModeBias (Get)](#IMEModeBias1) | Gets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control. |
| [IMEModeBias (Set)](#IMEModeBias2) | Sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control. |
| [IMEOptions (Get)](#IMEOptions1) | Retrieves the current Input Method Editor (IME) options. This message is available only in Asian-language versions of the operating system. |
| [IMEOptions (Set)](#IMEOptions2) | Sets the Input Method Editor (IME) options. |
| [LangOptions (Get)](#LangOptions1) | Gets a rich edit control's option settings for Input Method Editor (IME) and Asian language support. |
| [LangOptions (Set)](#LangOptions2) | Sets options for Input Method Editor (IME) and Asian language support in a rich edit control. |
| [LimitText (Get)](#LimitText1) | Gets the current text limit for a rich edit control. |
| [LimitText (Set)](#LimitText2) | Sets the text limit of a rich edit control. The text limit is the maximum amount of text, in TCHARs, that the user can type into the edit control. |
| [Modify (Get)](#Modify1) | Gets the state of a rich edit control's modification flag. The flag indicates whether the contents of the rich edit control have been modified. |
| [Modify (Set)](#Modify2) | Sets or clears the modification flag for a rich edit control. The modification flag indicates whether the text within the rich edit control has been modified. |
| [Options (Get)](#Options1) | Retrieves the options for a rich edit control. |
| [Options (Set)](#Options2) | Sets the options for a rich edit control. |
| [PageRotate (Get)](#PageRotate1) | Deprecated. Gets the text layout for a Microsoft Rich Edit control. |
| [PageRotate (Set)](#PageRotate2) | Deprecated. Sets the text layout for a Microsoft Rich Edit control. |
| [ParaFormat (Get)](#ParaFormat1) | Retrieves the paragraph formatting of the current selection in a rich edit control. |
| [ParaFormat (Set)](#ParaFormat2) | Sets the paragraph formatting for the current selection in a rich edit control. |
| [PasswordChar (Get)](#PasswordChar1) | Gets the password character that a rich edit control displays when the user enters text. |
| [PasswordChar (Set)](#PasswordChar2) | Sets or removes the password character for a rich edit control. When a password character is set, that character is displayed in place of the characters typed by the user. |
| [Punctuation (Get)](#Punctuation1) | Gets the current punctuation characters for the rich edit control. |
| [Punctuation (Set)](#Punctuation2) | Sets the punctuation characters for a rich edit control. |
| [Rect (Get)](#Rect1) | Gets the formatting rectangle of a rich edit control. |
| [Rect (Set)](#Rect2) | Sets the formatting rectangle of a multiline rich edit control. |
| [RectNP (Set)](#RectNP) | Sets the formatting rectangle of a multiline rich edit control. |
| [ScrollPos (Get)](#ScrollPos1) | Obtains the current scroll position of the edit control. |
| [ScrollPos (Set)](#ScrollPos2) | Tells the rich edit control to scroll to a particular point. |
| [StoryType (Get)](#StoryType1) | Gets the story type. |
| [StoryType (Set)](#StoryType2) | Sets the story type. |
| [Text (Get)](#Text1) | Retrieves the text from a rich edit control. |
| [Text (Set)](#Text2) | Sets the text of an edit control. |
| [TextMode (Get)](#TextMode1) | Gets the current text mode and undo level of a rich edit control. |
| [TextMode (Set)](#TextMode2) | Sets the text mode or undo level of a rich edit control. |
| [TouchOptions (Get)](#TouchOptions1) | Retrieves the touch options that are associated with a rich edit control. |
| [TouchOptions (Set)](#TouchOptions2) | Sets the touch options associated with a rich edit control. |
| [TypographyOptions (Get)](#TypographyOptions1) | Returns the current state of the typography options of a rich edit control. |
| [TypographyOptions (Set)](#TypographyOptions2) | Sets the current state of the typography options of a rich edit control. |

### CRichEditCtrl Methods

| Name       | Description |
| ---------- | ----------- |
| [](#) |  |


