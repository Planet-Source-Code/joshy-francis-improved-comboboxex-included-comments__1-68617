Attribute VB_Name = "modComboBoxEx"
' ======================================================================================
' Name:     modComboBoxEx.bas
' Author:   Joshy Francis (joshylogicss@yahoo.co.in)
' Date:     3 March 2007
'
' Requires: None
'
' Copyright © 2000-2007 Joshy Francis
' --------------------------------------------------------------------------------------
'The implementation of ComboBoxEx in VB.All by API.
'you can freely use this code anywhere.But I wants you must include the copyright info
'All functions in this module written by me.
' --------------------------------------------------------------------------------------
'No updates.This is the first version.
'I Just included comments on every important lines.Sorry for my bad english.
'I developed this program by converting the C Documentation to VB and experiments with VB.
'You can improve this program by your experiments.I didn't done all parts of the
'ComboBoxEx.

Option Explicit

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_SYSMENU = &H80000
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wparam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Long, lParam As Any) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_CLASSDC = &H40
Public Const CS_DBLCLKS = &H8
Public Const CS_HREDRAW = &H2
Public Const CS_INSERTCHAR = &H2000
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_NOCLOSE = &H200
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOMOVECARET = &H4000
Public Const CS_OWNDC = &H20
Public Const CS_PARENTDC = &H80
Public Const CS_PUBLICCLASS = &H4000
Public Const CS_SAVEBITS = &H800
Public Const CS_VREDRAW = &H1

Public Enum ComCtlClasses
     ICC_LISTVIEW_CLASSES = &H1      ' listview, header
     ICC_TREEVIEW_CLASSES = &H2       ' treeview, tooltips
     ICC_BAR_CLASSES = &H4            ' toolbar, statusbar, trackbar, tooltips
     ICC_TAB_CLASSES = &H8            ' tab, tooltips
     ICC_UPDOWN_CLASS = &H10          ' updown
     ICC_PROGRESS_CLASS = &H20        ' progress
     ICC_HOTKEY_CLASS = &H40          ' hotkey
     ICC_ANIMATE_CLASS = &H80         ' animate
     ICC_WIN95_CLASSES = &HFF        '
     ICC_DATE_CLASSES = &H100         ' month picker, date picker, time picker, updown
     ICC_USEREX_CLASSES = &H200       ' comboex
     ICC_COOL_CLASSES = &H400         ' rebar (coolbar) control
    #If (WIN32_IE >= &H400) Then    '
         ICC_INTERNET_CLASSES &H800
         ICC_PAGESCROLLER_CLASS &H1000       ' page scroller
         ICC_NATIVEFNTCTL_CLASS &H2000       ' native font control
    #End If
End Enum
Public Type INITCOMMONCONTROLSEX
    dwSize As Long 'DWORD ;             // size of this structure
    dwICC As ComCtlClasses 'Long 'DWORD ;              // flags indicating which classes to be initialized
End Type '} INITCOMMONCONTROLSEX, *LPINITCOMMONCONTROLSEX;

Public Declare Function INITCOMMONCONTROLSEX Lib "COMCTL32.DLL" Alias "InitCommonControlsEx" (ICCClass As INITCOMMONCONTROLSEX) As Long 'Boolean
Public Type NMHDR
  hwndFrom As Long   ' Window handle of control sending message
  idFrom As Long        ' Identifier of control sending message
  code  As Long          ' Specifies the notification code
End Type
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long

Private Const WM_NOTIFY = &H4E
Private Const WM_DESTROY = &H2
Public Const WM_CREATE = &H1

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101

Public Const WM_SETREDRAW = &HB
Public Const WM_USER = &H400

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Const CCM_FIRST = &H2000                    ' Common control shared messages
Public Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Public Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
'
'Public Type NMHDR
'  hwndFrom As Long   ' Window handle of control sending message
'  idFrom As Long        ' Identifier of control sending message
'  code  As Long          ' Specifies the notification code
'End Type


''''''''''  ComboBoxEx ''''''''''''''''

'Public Const WC_COMBOBOXEXW         L"ComboBoxEx32"
Public Const WC_COMBOBOXEXA = "ComboBoxEx32"
'
'#ifdef UNICODE
'Public Const WC_COMBOBOXEX          WC_COMBOBOXEXW
'#Else
Public Const WC_COMBOBOXEX = WC_COMBOBOXEXA
'#End If

Public Const CBEIF_TEXT = &H1
Public Const CBEIF_IMAGE = &H2
Public Const CBEIF_SELECTEDIMAGE = &H4
Public Const CBEIF_OVERLAY = &H8
Public Const CBEIF_INDENT = &H10
Public Const CBEIF_LPARAM = &H20

Public Const CBEIF_DI_SETITEM = &H10000000

Public Type COMBOBOXEXITEMA 'tagCOMBOBOXEXITEMA
mask As Long '    UINT mask;
iItem   As Long '    int iItem;
pszText As String '    LPSTR pszText;
cchTextMax  As Long '    int cchTextMax;
iImage    As Long '   int iImage;
iSelectedImage    As Long '    int iSelectedImage;
iOverlay     As Long '  int iOverlay;
 iIndent    As Long ' int iIndent;
lParam    As Long '  LPARAM lParam;
'} COMBOBOXEXITEMA, *PCOMBOBOXEXITEMA;
'typedef COMBOBOXEXITEMA CONST *PCCOMBOEXITEMA;
End Type

'Public Type tagCOMBOBOXEXITEMW
'    UINT mask;
'    int iItem;
'    LPWSTR pszText;
'    int cchTextMax;
'    int iImage;
'    int iSelectedImage;
'    int iOverlay;
'    int iIndent;
'    LPARAM lParam;
'} COMBOBOXEXITEMW, *PCOMBOBOXEXITEMW;
'typedef COMBOBOXEXITEMW CONST *PCCOMBOEXITEMW;
'
'#ifdef UNICODE
'Public Const COMBOBOXEXITEM            COMBOBOXEXITEMW
'Public Const PCOMBOBOXEXITEM           PCOMBOBOXEXITEMW
'Public Const PCCOMBOBOXEXITEM          PCCOMBOBOXEXITEMW
'#Else
'Public Const COMBOBOXEXITEM = COMBOBOXEXITEMA
'Public Const PCOMBOBOXEXITEM = PCOMBOBOXEXITEMA
'Public Const PCCOMBOBOXEXITEM = PCCOMBOBOXEXITEMA
''#End If
Public Const CB_DELETESTRING = &H144

Public Const CBEM_INSERTITEMA = (WM_USER + 1)
Public Const CBEM_SETIMAGELIST = (WM_USER + 2)
Public Const CBEM_GETIMAGELIST = (WM_USER + 3)
Public Const CBEM_GETITEMA = (WM_USER + 4)
Public Const CBEM_SETITEMA = (WM_USER + 5)
Public Const CBEM_DELETEITEM = CB_DELETESTRING
Public Const CBEM_GETCOMBOCONTROL = (WM_USER + 6)
Public Const CBEM_GETEDITCONTROL = (WM_USER + 7)
'#if (_WIN32_IE >= = &h0400)
Public Const CBEM_SETEXSTYLE = (WM_USER + 8)         ' use  SETEXTENDEDSTYLE instead
Public Const CBEM_SETEXTENDEDSTYLE = (WM_USER + 14)    ' lparam == new style, wParam (optional) == mask
Public Const CBEM_GETEXSTYLE = (WM_USER + 9)        ' use GETEXTENDEDSTYLE instead
Public Const CBEM_GETEXTENDEDSTYLE = (WM_USER + 9)
Public Const CBEM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
Public Const CBEM_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT
''#Else
'Public Const CBEM_SETEXSTYLE = (WM_USER + 8)
'Public Const CBEM_GETEXSTYLE = (WM_USER + 9)
'#End If
Public Const CBEM_HASEDITCHANGED = (WM_USER + 10)
Public Const CBEM_INSERTITEMW = (WM_USER + 11)
Public Const CBEM_SETITEMW = (WM_USER + 12)
Public Const CBEM_GETITEMW = (WM_USER + 13)

'#ifdef UNICODE
'Public Const CBEM_INSERTITEM = CBEM_INSERTITEMW
'Public Const CBEM_SETITEM = CBEM_SETITEMW
'Public Const CBEM_GETITEM = CBEM_GETITEMW
'#Else
Public Const CBEM_INSERTITEM = CBEM_INSERTITEMA
Public Const CBEM_SETITEM = CBEM_SETITEMA
Public Const CBEM_GETITEM = CBEM_GETITEMA
'#End If

Public Const CBES_EX_NOEDITIMAGE = &H1
Public Const CBES_EX_NOEDITIMAGEINDENT = &H2
Public Const CBES_EX_PATHWORDBREAKPROC = &H4
'#if (_WIN32_IE >= = &h0400)
Public Const CBES_EX_NOSIZELIMIT = &H8
Public Const CBES_EX_CASESENSITIVE = &H10

Public Type NMCOMBOBOXEXA '{
hdr As Long '    NMHDR hdr;
ceItem As COMBOBOXEXITEMA '    COMBOBOXEXITEMA ceItem;
'} NMCOMBOBOXEXA, *PNMCOMBOBOXEXA;
End Type
'Public Type {
'    NMHDR hdr;
'    COMBOBOXEXITEMW ceItem;
'} NMCOMBOBOXEXW, *PNMCOMBOBOXEXW;
'
'#ifdef UNICODE
'Public Const NMCOMBOBOXEX            NMCOMBOBOXEXW
'Public Const PNMCOMBOBOXEX           PNMCOMBOBOXEXW
'Public Const CBEN_GETDISPINFO        CBEN_GETDISPINFOW
'#Else
'Public Const NMCOMBOBOXEX = NMCOMBOBOXEXA
'Public Const PNMCOMBOBOXEX = PNMCOMBOBOXEXA
'Public Const CBEN_GETDISPINFO = CBEN_GETDISPINFOA
'#End If

'#Else
Public Type NMCOMBOBOXEX '{
hdr As Long '    NMHDR hdr;
ceItem As COMBOBOXEXITEMA '    COMBOBOXEXITEM ;
'} NMCOMBOBOXEX, *PNMCOMBOBOXEX;
End Type
Public Const CBEN_FIRST = 0
Public Const CBEN_GETDISPINFO = (CBEN_FIRST - 0)

'#End If     ' _WIN32_IE >= = &h0400

'#if (_WIN32_IE >= = &h0400)
Public Const CBEN_GETDISPINFOA = (CBEN_FIRST - 0)
'#End If
Public Const CBEN_INSERTITEM = (CBEN_FIRST - 1)
Public Const CBEN_DELETEITEM = (CBEN_FIRST - 2)
Public Const CBEN_BEGINEDIT = (CBEN_FIRST - 4)
Public Const CBEN_ENDEDITA = (CBEN_FIRST - 5)
Public Const CBEN_ENDEDITW = (CBEN_FIRST - 6)

'#if (_WIN32_IE >= = &h0400)
Public Const CBEN_GETDISPINFOW = (CBEN_FIRST - 7)
'#End If

'#if (_WIN32_IE >= = &h0400)
Public Const CBEN_DRAGBEGINA = (CBEN_FIRST - 8)
Public Const CBEN_DRAGBEGINW = (CBEN_FIRST - 9)

'#ifdef UNICODE
'Public Const CBEN_DRAGBEGIN = CBEN_DRAGBEGINW
'#Else
Public Const CBEN_DRAGBEGIN = CBEN_DRAGBEGINA
'#End If

'#End If '(_WIN32_IE >= = &h0400)

' lParam specifies why the endedit is happening
'#ifdef UNICODE
'Public Const CBEN_ENDEDIT = CBEN_ENDEDITW
'#Else
Public Const CBEN_ENDEDIT = CBEN_ENDEDITA
'#End If

Public Const CBENF_KILLFOCUS = 1
Public Const CBENF_RETURN = 2
Public Const CBENF_ESCAPE = 3
Public Const CBENF_DROPDOWN = 4

Public Const CBEMAXSTRLEN = 260

'#if (_WIN32_IE >= = &h0400)
' CBEN_DRAGBEGIN sends this information ...
'
'Public Type =
'    NMHDR hdr;
'    int   iItemid;
'    WCHAR szText[CBEMAXSTRLEN];
'}NMCBEDRAGBEGINW, *LPNMCBEDRAGBEGINW, *PNMCBEDRAGBEGINW;

Public Type NMCBEDRAGBEGINA '{
hdr As Long '      NMHDR hdr;
iItemid  As Long '     int   iItemid;
szText As String * CBEMAXSTRLEN '    char szText[CBEMAXSTRLEN];
'}NMCBEDRAGBEGINA, *LPNMCBEDRAGBEGINA, *PNMCBEDRAGBEGINA;
End Type
'#ifdef UNICODE
'Public Const  NMCBEDRAGBEGIN NMCBEDRAGBEGINW
'Public Const  LPNMCBEDRAGBEGIN LPNMCBEDRAGBEGINW
'Public Const  PNMCBEDRAGBEGIN PNMCBEDRAGBEGINW
'#Else
'Public Const NMCBEDRAGBEGIN = NMCBEDRAGBEGINA
'Public Const LPNMCBEDRAGBEGIN = LPNMCBEDRAGBEGINA
'Public Const PNMCBEDRAGBEGIN = PNMCBEDRAGBEGINA
'#End If
'#End If     ' _WIN32_IE >= = &h0400

' CBEN_ENDEDIT sends this information...
' fChanged if the user actually did anything
' iNewSelection gives what would be the new selection unless the notify is failed
'                      iNewSelection may be CB_ERR if there's no match
'Public Type {
'        NMHDR hdr;
'        BOOL fChanged;
'        int iNewSelection;
'        WCHAR szText[CBEMAXSTRLEN];
'        int iWhy;
'} NMCBEENDEDITW, *LPNMCBEENDEDITW, *PNMCBEENDEDITW;

Public Type NMCBEENDEDITA '{
hdr As Long '        NMHDR hdr;
fChanged    As Boolean '        BOOL fChanged;
iNewSelection  As Long '       int iNewSelection;
szText As String * CBEMAXSTRLEN '     char szText[CBEMAXSTRLEN];
iWhy      As Long '    int iWhy;
'} NMCBEENDEDITA, *LPNMCBEENDEDITA,*PNMCBEENDEDITA;
End Type
'
'#ifdef UNICODE
'Public Const  NMCBEENDEDIT NMCBEENDEDITW
'Public Const  LPNMCBEENDEDIT LPNMCBEENDEDITW
'Public Const  PNMCBEENDEDIT PNMCBEENDEDITW
'#Else
'Public Const NMCBEENDEDIT = NMCBEENDEDITA
'Public Const LPNMCBEENDEDIT = LPNMCBEENDEDITA
'Public Const PNMCBEENDEDIT = PNMCBEENDEDITA
'#End If

'#End If

'#End If     ' _WIN32_IE >= = &h0300
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const MIN_COMBOCX = 113
Public Const MIN_COMBOCY = 42
'//====== COMMON CONTROL STYLES ================================================
Public Const CCS_TOP = &H1                      'L
Public Const CCS_NOMOVEY = &H2                  'L
Public Const CCS_BOTTOM = &H3                   'L
Public Const CCS_NORESIZE = &H4                 'L
Public Const CCS_NOPARENTALIGN = &H8            'L
Public Const CCS_ADJUSTABLE = &H20              'L
Public Const CCS_NODIVIDER = &H40               'L
#If WIN32_IE >= &H300 Then
Public Const CCS_VERT = &H80                    'L
Public Const CCS_LEFT = (CCS_VERT Or CCS_TOP)
Public Const CCS_RIGHT = (CCS_VERT Or CCS_BOTTOM)
Public Const CCS_NOMOVEX = (CCS_VERT Or CCS_NOMOVEY)
#End If
'/*
' * Combo Box styles
' */
Public Const CBS_SIMPLE = &H1
Public Const CBS_DROPDOWN = &H2
Public Const CBS_DROPDOWNLIST = &H3
Public Const CBS_OWNERDRAWFIXED = &H10
Public Const CBS_OWNERDRAWVARIABLE = &H20
Public Const CBS_AUTOHSCROLL = &H40
Public Const CBS_OEMCONVERT = &H80
Public Const CBS_SORT = &H100
Public Const CBS_HASSTRINGS = &H200
Public Const CBS_NOINTEGRALHEIGHT = &H400
Public Const CBS_DISABLENOSCROLL = &H800
#If (WINVER >= &H400) Then
Public Const CBS_UPPERCASE = &H2000
Public Const CBS_LOWERCASE = &H4000
#End If '/* WINVER >= =&h0400 */
Public Const CB_SETCURSEL = &H14E
Public Const CB_ADDSTRING = &H143
'Public Const CB_DELETESTRING = &H144
Public Const CB_DIR = &H145
Public Const CB_ERR = (-1)
Public Const CB_ERRSPACE = (-2)
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_GETDROPPEDCONTROLRECT = &H152
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_GETEDITSEL = &H140
Public Const CB_GETEXTENDEDUI = &H156
Public Const CB_GETITEMDATA = &H150
Public Const CB_GETITEMHEIGHT = &H154
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_GETLOCALE = &H15A
Public Const CB_INSERTSTRING = &H14A
Public Const CB_LIMITTEXT = &H141
Public Const CB_MSGMAX = &H15B
Public Const CB_OKAY = 0
Public Const CB_RESETCONTENT = &H14B
Public Const CB_SELECTSTRING = &H14D
'Public Const CB_SETCURSEL = &H14E
Public Const CB_SETEDITSEL = &H142
Public Const CB_SETEXTENDEDUI = &H155
Public Const CB_SETITEMDATA = &H151
Public Const CB_SETITEMHEIGHT = &H153
Public Const CB_SETLOCALE = &H159
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CBN_CLOSEUP = 8
Public Const CBN_DBLCLK = 2
Public Const CBN_DROPDOWN = 7
Public Const CBN_EDITCHANGE = 5
Public Const CBN_EDITUPDATE = 6
Public Const CBN_ERRSPACE = (-1)
Public Const CBN_KILLFOCUS = 4
Public Const CBN_SELCHANGE = 1
Public Const CBN_SELENDCANCEL = 10
Public Const CBN_SELENDOK = 9
Public Const CBN_SETFOCUS = 3
'Public Const CBS_AUTOHSCROLL = &H40&
'Public Const CBS_DISABLENOSCROLL = &H800&
'Public Const CBS_DROPDOWN = &H2&
'Public Const CBS_DROPDOWNLIST = &H3&
'Public Const CBS_HASSTRINGS = &H200&
'Public Const CBS_NOINTEGRALHEIGHT = &H400&
'Public Const CBS_OEMCONVERT = &H80&
'Public Const CBS_OWNERDRAWFIXED = &H10&
'Public Const CBS_OWNERDRAWVARIABLE = &H20&
'Public Const CBS_SIMPLE = &H1&
'Public Const CBS_SORT = &H100&
Public Const WM_COMMAND = &H111
'Global Window Handle & Windowprocedure Handle
Public Wnd As Long, OldProc As Long

Function CreateComboBoxEx(ByVal hwnd As Long) As Long
'This is the main function.It Creates a ComboBoxEx window and returns the window handle
'You can modify this function by including Coordinate Parameters,style etc.

Dim stl As Long, ExStl As Long, Inited As Long
Dim Ret As Long
    Dim IX As INITCOMMONCONTROLSEX
        IX.dwICC = ICC_USEREX_CLASSES  ' ComboBoxEx Classes
        IX.dwSize = Len(IX)
'If CommonControl library is not initialized the program does'nt work.
Inited = INITCOMMONCONTROLSEX(IX)
        If Inited <> 1 Then
            MsgBox "INITCOMMONCONTROLSEX Failed.", vbCritical
        End If
    stl = WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS 'Or WS_THICKFRAME 'Or WS_BORDER
'Here i gives the different styles for the ComboBoxEx.Select the styles you wish
    stl = stl Or CBS_AUTOHSCROLL
'    stl = stl Or CBS_DROPDOWNLIST
    stl = stl Or CBS_DROPDOWN
        Dim rc As RECT
            GetWindowRect hwnd, rc
                rc.Bottom = rc.Bottom - rc.Top
                rc.Right = rc.Right - rc.Left
 'Creates the ComboBoxEx Window.
'It is very fast and safe function.In VB the dynamic control creation is not possible when the controls are kept in DLLS that not compatible with VB.
'If the control not created no error will occur the return value will be zero else return value will be handle.
'This is the style of C & CPP programs. The concept of pointers is done in VB by this way.
                   
Wnd = CreateWindowEx(ExStl, WC_COMBOBOXEXA, ByVal " ", stl, 10, 10, rc.Right / 3, rc.Bottom / 2, hwnd, 0, App.hInstance, ByVal hwnd)
If Wnd = 0 Then
' Window is not created . Load Window class from the DLL. You can test this line by commenting the above line 'Inited = INITCOMMONCONTROLSEX(IX)...'
'The below technique is very useful to crteate other controls used by other Programs.
'        MsgBox "Registering Class"
        Dim Class As WNDCLASS
    With Class
        .cbClsextra = Len(Class)
        .hInstance = LoadLibrary("COMCTL32.DLL")
        .lpszClassName = WC_COMBOBOXEXA
        .style = CS_PUBLICCLASS
    End With
    Ret = RegisterClass(Class)
        If Ret = 0 Then
            'If Register class failed exit
            Exit Function
        End If
Wnd = CreateWindowEx(ExStl, WC_COMBOBOXEXA, ByVal " ", stl, 0, 0, rc.Right / 3, rc.Bottom / 2, hwnd, 0, App.hInstance, ByVal hwnd)
End If
If Wnd <> 0 Then
    CreateComboBoxEx = Wnd
        'An Imagelist is created for put the tab images quickly.
        'You can change the icon size to 48x48 or as you wish.
'***************** Used Bitmpas for Method 1 *************************************
    'Here creates the imagelist that will contain bitmaps
            hIml = ImageList_Create(16, 16, ILC_COLOR Or ILD_TRANSPARENT, 1, 0)
               Ret = ImageList_Add(hIml, frmComboBoxEx.Picture1.Picture.Handle, 0)
                Ret = ImageList_Add(hIml, frmComboBoxEx.Picture2.Picture.Handle, 0)
'************************************************************************************
'                           or
'***************** Used Icons for Method 2 *************************************
    'Here creates the imagelist that will contain icons
'            hIml = ImageList_Create(32, 32, ILC_COLOR Or ILD_TRANSPARENT, 1, 0)
'                Ret = ImageList_AddIcon(hIml, , frmComboBoxEx.Icon.Handle)
'                Ret = ImageList_AddIcon(hIml, , frmComboBoxEx.Picture3.Picture.Handle)
'************************************************************************************
        'Sets the Imagelist
            Ret = SendMessage(Wnd, CBEM_SETIMAGELIST, 0, ByVal hIml)

'            SendMessage Wnd, CBEM_SETUNICODEFORMAT, 0, ByVal 0 'Set to ANSI
'            SendMessage Wnd, CBEM_SETUNICODEFORMAT, 1, ByVal 0 'Set to Unicode
        'Adds some sample Items
                InsertItem "Item 1"
                InsertItem "Item 2", , 1, 1, 1
                InsertItem "Item 3", , , , 2
                InsertItem "Item 4", , 1, 1, 3
                'Selects Item index 1
                    SelItem 1
            'Sets the window procedure
    OldProc = SetWindowLong(Wnd, GWL_WNDPROC, AddressOf WndProc)
End If
End Function
Function InsertItem(ByVal str As String, Optional ByVal idx As Long = -1, _
    Optional ByVal img As Long = 0, Optional ByVal Selimg As Long = 0, _
    Optional ByVal Indent As Long = 0) As Long
'Inserts an item to the ComboBoxEx
        Dim cbi As COMBOBOXEXITEMA
            With cbi
                .mask = CBEIF_TEXT Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE Or CBEIF_INDENT
                .pszText = str
                .cchTextMax = Len(.pszText)
                .iItem = idx '-1
                .iImage = img '0 ' Image to display
                .iSelectedImage = Selimg '0 ' Image to display
                .iIndent = Indent ' Indent .
            End With
    InsertItem = SendMessage(Wnd, CBEM_INSERTITEM, 0, cbi)
End Function

Function SelItem(ByVal idx As Long) As Long
'Selects an item by Index
SelItem = SendMessage(Wnd, CB_SETCURSEL, idx, ByVal 0)
End Function
Function DelItem(ByVal idx As Long) As Long
'Delete an item by index
DelItem = SendMessage(Wnd, CBEM_DELETEITEM, idx, ByVal 0)
End Function
Function GetItem(cbi As COMBOBOXEXITEMA, Optional ByVal idx As Long = -1) As Long
'Gets an item's information
'    Dim cbi As COMBOBOXEXITEMA
        With cbi
            .mask = CBEIF_TEXT Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE Or CBEIF_INDENT
            .pszText = Space$(260)
            .cchTextMax = Len(.pszText)
            .iItem = idx '-1
            .iImage = 0 '0 ' Image to display
            .iSelectedImage = 0 '0 ' Image to display
            .iIndent = 0
        End With
    GetItem = SendMessage(Wnd, CBEM_GETITEMA, 0, cbi)
'Returns nonzero if successful, or zero otherwise.
'When the message is sent, the iItem and mask members of the structure must be set to indicate the index of the target item and the type of information to be retrieved. Other members are set as needed. For example, to get text, you must set the CBEIF_TEXT flag in mask, and assign a value to cchTextMax. Setting the iItem member to -1 will retrieve the item displayed in the edit control.
End Function
Function GetItemText(Optional ByVal idx As Long = -1) As String
'returns an item's text by index
    Dim cbi As COMBOBOXEXITEMA, Ret As Long
Ret = GetItem(cbi, idx)
If Ret Then
    GetItemText = Trim0(cbi.pszText)
End If
End Function
Function GetSelItem() As Long
'returns selected item index
GetSelItem = SendMessage(Wnd, CB_GETCURSEL, 0, ByVal 0)
End Function

Function SetItemText(ByVal str As String, Optional ByVal idx As Long = -1) As Long
'Change the item's text by giving index & newtext
    Dim cbi As COMBOBOXEXITEMA, Ret As Long
Ret = GetItem(cbi, idx)
If Ret Then
        cbi.pszText = str
        cbi.cchTextMax = Len(str)
    SetItemText = SendMessage(Wnd, CBEM_SETITEMA, 0, cbi)
'        Returns nonzero if successful, or zero otherwise
End If
End Function

Function DropDown() As Long
'Show the dropdown list
DropDown = SendMessage(Wnd, CB_SHOWDROPDOWN, 1, ByVal 0)
End Function
Function GetCount() As Long
'returns the itemcount
GetCount = SendMessage(Wnd, CB_GETCOUNT, 0, ByVal 0)
End Function
Function Clear() As Long
'clear the comboboxex control
Clear = SendMessage(Wnd, CB_RESETCONTENT, 0, ByVal 0)
End Function
Function GetEditCtl() As Long
'returns the editcontrl window handle of the comboboxex
'Usefule to get or set the typing text.
GetEditCtl = SendMessage(Wnd, CBEM_GETEDITCONTROL, 0, ByVal 0)
End Function
Function IsEditChanged() As Boolean
'returns if editcontrol is changed
IsEditChanged = SendMessage(Wnd, CBEM_HASEDITCHANGED, 0, ByVal 0) = 1
End Function

Function DestroyComboBoxEx() As Long
'Destroys the window
If Wnd <> 0 Then
    DestroyComboBoxEx = DestroyWindow(Wnd)
    If OldProc <> 0 Then
        SetWindowLong Wnd, GWL_WNDPROC, OldProc
    End If
        Wnd = 0
        OldProc = 0
End If
End Function
Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wparam As Long, ByVal lParam As Long) As Long
Dim nc As Long, str As String
Select Case msg
    Case WM_COMMAND
        nc = HiWord(wparam) 'Notify Code
            Select Case nc
                Case CBN_SELCHANGE
                    frmComboBoxEx.Caption = "SelChange " & GetSelItem
                Case Else
                    frmComboBoxEx.Caption = "else " & nc
'                     If IsEditChanged = True Then
'                                    str = Space$(260)
'                                GetWindowText GetEditCtl, str, Space$(260)
''                            frmComboBoxEx.Caption = "EditChange " & Trim0(str)
'                    End If
           End Select
    Case WM_NOTIFY
        Dim hdr As NMHDR
                    frmComboBoxEx.Caption = "WM_NOTIFY "
            CopyMemory hdr, ByVal lParam, Len(hdr)
        Select Case hdr.code
            Case CBEN_BEGINEDIT
                    frmComboBoxEx.Caption = "CBEN_BEGINEDIT "
        End Select

    Case Else
End Select
    WndProc = CallWindowProc(OldProc, hwnd, msg, wparam, lParam)
End Function
