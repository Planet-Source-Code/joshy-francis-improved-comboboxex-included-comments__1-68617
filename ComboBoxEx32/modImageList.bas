Attribute VB_Name = "modImageList"
' Some of the functions written by me and some of the functions i taken from PSC

Option Explicit
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'====== IMAGE APIS ===========================================================
'#ifndef NOIMAGEAPIS
Public Const CLR_NONE = vbBlack                '00FFFFFFFFL
Public Const CLR_DEFAULT = vbBlue        ' &hFF000000L
'struct _IMAGELIST;
'Public Type _IMAGELIST NEAR* HIMAGELIST;
'#if (_WIN32_IE >= = &h0300)
Public Type IMAGELISTDRAWPARAMS
           cbSize As Long
    hIml As Long '    HIMAGELIST  himl;
             i As Long
             hdcDst As Long
             x As Long
             y As Long
             cx As Long
             cy As Long
             xBitmap    As Long      ' x offest from the upperleft of bitmap
             yBitmap     As Long       ' y offset from the upperleft of bitmap
    rgbBk    As Long 'COLORREF    rgbBk;
    rgbFg  As Long 'COLORREF    rgbFg;
            fStyle As Long
           dwRop As Long
'} IMAGELISTDRAWPARAMS, FAR * LPIMAGELISTDRAWPARAMS;
'#End If     ' _WIN32_IE >= = &h0300
End Type
Public Const ILC_MASK = &H1
Public Const ILC_COLOR = &H0
Public Const ILC_COLORDDB = &HFE
Public Const ILC_COLOR4 = &H4
Public Const ILC_COLOR8 = &H8
Public Const ILC_COLOR16 = &H10
Public Const ILC_COLOR24 = &H18
Public Const ILC_COLOR32 = &H20
Public Const ILC_PALETTE = &H800                   ' (not implemented)
'All Byval
Public Declare Function ImageList_Create Lib "COMCTL32.DLL" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Public Declare Function ImageList_Destroy Lib "COMCTL32.DLL" (ByVal hIml As Long) As Boolean
Public Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" (ByVal hIml As Long) As Long
Public Declare Function ImageList_SetImageCount Lib "COMCTL32.DLL" (ByVal hIml As Long, uNewCount As Long) As Boolean
'#if (_WIN32_IE >= = &h0300)
'WINCOMMCTRLAPI BOOL        WINAPI ImageList_SetImageCount(HIMAGELIST himl, UINT uNewCount);
'#End If
Public Declare Function ImageList_Add Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Public Declare Function ImageList_ReplaceIcon Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Public Declare Function ImageList_SetBkColor Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal clrBk As Long) As Long
Public Declare Function ImageList_GetBkColor Lib "COMCTL32.DLL" (ByVal hIml As Long) As Long
Public Declare Function ImageList_SetOverlayImage Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal iImage As Long, ByVal iOverlay As Long) As Boolean
Public Declare Function ImageList_AddIcon Lib "COMCTL32.DLL" Alias "ImageList_ReplaceIcon" (ByVal hIml As Long, Optional ByVal i As Long = -1, Optional ByVal hIcon As Long) As Long
'Public Declare Function ImageList_ReplaceIcon Lib "COMCTL32.DLL" (himl As Long, Optional i As Long = -1, Optional hIcon As Long) As Long

Public Const ILD_NORMAL = &H0
Public Const ILD_TRANSPARENT = &H1
Public Const ILD_MASK = &H10
Public Const ILD_IMAGE = &H20
'#if (_WIN32_IE >= = &h0300)
'Public Const ILD_ROP = &H40
'#End If
Public Const ILD_BLEND25 = &H2
Public Const ILD_BLEND50 = &H4
Public Const ILD_OVERLAYMASK = &HF00
'Public Const INDEXTOOVERLAYMASK(i)   ((i) << 8)

Public Const ILD_SELECTED = ILD_BLEND50
Public Const ILD_FOCUS = ILD_BLEND25
Public Const ILD_BLEND = ILD_BLEND50
Public Const CLR_HILIGHT = CLR_DEFAULT
Public Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
'#ifdef _WIN32
Public Declare Function ImageList_Replace Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Public Declare Function ImageList_AddMasked Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Public Declare Function ImageList_DrawEx Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
'#if (_WIN32_IE >= = &h0300)
Public Declare Function ImageList_DrawIndirect Lib "COMCTL32.DLL" (ByRef pimldp As IMAGELISTDRAWPARAMS) As Long
'#End If
Public Declare Function ImageList_Remove Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long) As Long
Public Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal flags As Long) As Long
Public Declare Function ImageList_LoadImageA Lib "COMCTL32.DLL" (ByVal hi As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_LoadImageW Lib "COMCTL32.DLL" (ByVal hi As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long) As Long
'#ifdef UNICODE
'Public Const ImageList_LoadImage     ImageList_LoadImageW
'#Else
'Public Const ImageList_LoadImage     ImageList_LoadImageA
'#End If
'#if (_WIN32_IE >= = &h0300)
Public Const ILCF_MOVE = (&H0)
Public Const ILCF_SWAP = (&H1)
Public Declare Function ImageList_Copy Lib "COMCTL32.DLL" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long
'#End If
Public Declare Function ImageList_BeginDrag Lib "COMCTL32.DLL" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Public Declare Sub ImageList_EndDrag Lib "COMCTL32.DLL" ()
Public Declare Function ImageList_DragEnter Lib "COMCTL32.DLL" (ByVal hwndLock As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ImageList_DragLeave Lib "COMCTL32.DLL" (ByVal hwndLock As Long) As Long
Public Declare Function ImageList_DragMove Lib "COMCTL32.DLL" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ImageList_SetDragCursorImage Lib "COMCTL32.DLL" (ByVal hImlDrag As Long, ByVal iDrag As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Public Declare Function ImageList_DragShowNolock Lib "COMCTL32.DLL" (ByVal fShow As Long) As Long
Public Declare Function ImageList_GetDragImage Lib "COMCTL32.DLL" (ByRef ppt As Any, ByRef pptHotspot As Any) As Long
'Public Declare Function ImageList_RemoveAll Lib "COMCTL32.DLL" (himl As Long,   -1 As long) As Long
'Public Const     ImageList_RemoveAll(himl) ImageList_Remove(himl, -1)
'Public Const     ImageList_ExtractIcon(hi, himl, i) ImageList_GetIcon(himl, i, 0)
'Public Const     ImageList_LoadBitmap(hi, lpbmp, cx, cGrow, crMask) ImageList_LoadImage(hi, lpbmp, cx, cGrow, crMask, IMAGE_BITMAP, 0)
'#ifdef __IStream_INTERFACE_DEFINED__
Public Declare Function ImageList_Read Lib "COMCTL32.DLL" (ByRef pstm As Any) As Long
'WINCOMMCTRLAPI HIMAGELIST WINAPI ImageList_Read(LPSTREAM pstm);
Public Declare Function ImageList_Write Lib "COMCTL32.DLL" (ByVal hIml As Long, pstm As Any) As Long
'WINCOMMCTRLAPI BOOL       WINAPI ImageList_Write(HIMAGELIST himl, LPSTREAM pstm);
'#End If
Public Type IMAGEINFO
    hbmImage As Long '  HBITMAP ;
    hbmMask  As Long 'HBITMAP ;
         Unused1 As Long
         Unused2 As Long
        rcImage As RECT
'} IMAGEINFO, FAR *LPIMAGEINFO;
End Type
Public Declare Function ImageList_GetIconSize Lib "COMCTL32.DLL" (ByVal hIml As Long, ByRef cx As Long, ByRef cy As Long) As Long
'WINCOMMCTRLAPI BOOL        WINAPI ImageList_GetIconSize(HIMAGELIST himl, int FAR *cx, int FAR *cy);
Public Declare Function ImageList_SetIconSize Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal cx As Long, ByVal cy As Long) As Long
Public Declare Function ImageList_GetImageInfo Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByRef pImageInfo As Any) As Long
'WINCOMMCTRLAPI BOOL        WINAPI ImageList_GetImageInfo(HIMAGELIST himl, int i, IMAGEINFO FAR* pImageInfo);
Public Declare Function ImageList_Merge Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i1 As Long, ByVal himl2 As Long, ByVal i2 As Long, ByVal dx As Long, ByVal dy As Long) As Long
'#if (_WIN32_IE >= = &h0400)
Public Declare Function ImageList_Duplicate Lib "COMCTL32.DLL" (ByVal hIml As Long) As Long
'#End If
'Public Declare Function ShiftLeft Lib "MathDll.DLL" (ByVal a As Double, ByVal b As Double) As Double
Public Declare Function ShiftLeft Lib "Callback.dll" (ByVal a As Long, ByVal b As Long) As Long
Public Declare Function shl Lib "Callback.dll" Alias "ShiftLeft" (ByVal a As Long, ByVal b As Long) As Long
Public Declare Function ShiftRight Lib "Callback.dll" (ByVal a As Long, ByVal b As Long) As Long
Public Declare Function shr Lib "Callback.dll" Alias "ShiftRight" (ByVal a As Long, ByVal b As Long) As Long
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const VK_SHIFT = &H10
Public Const KEY_TOGGLED As Integer = &H1
Public Const KEY_DOWN As Integer = &H1000
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long
Public Type Size
        cx As Long
        cy As Long
End Type
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'Public Const WM_COMMAND = &H111
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Type TYPE_QWORD
    Value As Currency
End Type

Private Type TYPE_LOHIQWORD
    lLoDWord As Long
    lHiDWord As Long
End Type
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public hIml As Long, hImlSmall As Long, hImlState As Long, hImlStateSmall As Long
Public hImlDrag As Long
Public Icons As New Collection

Public Function CreateIML(Optional ByVal W As Long = 32, Optional ByVal H As Long = 32) As Long
    hIml = ImageList_Create(W, H, ILC_COLOR Or ILD_TRANSPARENT, 1, 0)
CreateIML = hIml
End Function
Public Function AddIcon(hIcon As Long) As Long
If hIml = 0 Then Exit Function
    AddIcon = ImageList_AddIcon(hIml, , hIcon)
End Function
Public Function ReplaceIcon(ByVal i As Long, ByVal hBmp As Long, Optional ByVal hBmpMask As Long = 0) As Long
If hIml = 0 Then Exit Function
    ReplaceIcon = ImageList_Replace(hIml, i, hBmp, hBmpMask)
End Function
Public Function DestroyIML() As Long
If hIml <> 0 Then
    DestroyIML = ImageList_Destroy(hIml)
        hIml = 0
End If
End Function
Function IsIconExists(ByVal Key As Variant) As Boolean
On Error GoTo E
    Dim T
    IsIconExists = False
        T = Icons.Item(Key)
    IsIconExists = True
Exit Function
E:
    IsIconExists = False
End Function
Function Trim0(ByVal str As String) As String
Dim i As Long
    i = InStr(str, Chr(0))
If i Then
    str = Left$(str, i - 1)
End If
    Trim0 = Trim$(str)
End Function
Function MakeWord(ByVal LowByte As Byte, HighByte As Byte) As Integer
'    HiByte = (iWord And &HFF00&) \ &H100
'    LoByte = iWord And &HFF
'MakeWord = (LowByte + 256) + (HighByte * 255) - (256 - HighByte)
MakeWord = (LowByte + &H100) + (HighByte * &HFF) - (&H100 - HighByte)
End Function
Function MakeDWord(ByVal lowWord As Integer, ByVal HighWord As Integer) As Long
'MakeDWord = (HighWord * 65536) + lowWord
MakeDWord = (HighWord * &H10000) + lowWord
End Function
Function HiByte(ByVal iWord As Integer) As Byte
    HiByte = (iWord And &HFF00&) \ &H100
End Function

Function LoByte(ByVal iWord As Integer) As Byte
    LoByte = iWord And &HFF
End Function
Function HiSWord(lDword As Single) As Integer
    HiSWord = (lDword And &HFFFF0000) \ &H10000
End Function

Function LoSWord(lDword As Single) As Integer
    If lDword And &H8000& Then
        LoSWord = lDword Or &HFFFF0000
    Else
        LoSWord = lDword And &HFFFF&
    End If
End Function

Function LoDWord(ByVal cQWord As Currency) As Long
    Dim QWord As TYPE_QWORD: Dim LoHiQword As TYPE_LOHIQWORD
    QWord.Value = cQWord / 10000
    LSet LoHiQword = QWord
    LoDWord = LoHiQword.lLoDWord
End Function

Function HiDWord(ByVal cQWord As Currency) As Long
    Dim QWord As TYPE_QWORD: Dim LoHiQword As TYPE_LOHIQWORD
    QWord.Value = cQWord / 10000
    LSet LoHiQword = QWord
    HiDWord = LoHiQword.lHiDWord
End Function

Function MakeQWord(ByVal lHiDWord As Long, ByVal lLoDWord As Long) As Currency
    Dim QWord As TYPE_QWORD: Dim LoHiQword As TYPE_LOHIQWORD
    LoHiQword.lHiDWord = lHiDWord: LoHiQword.lLoDWord = lLoDWord
    LSet QWord = LoHiQword: MakeQWord = QWord.Value * 10000
End Function


Function MAKEWPARAM(ByVal iLow As Integer, ByVal iHigh As Integer) As Long
MAKEWPARAM = (iLow + iHigh) * &H10000 + iLow
End Function
Function MAKELPARAM(ByVal cx As Integer, ByVal cy As Integer) As Long
'Print &HFFFF&
'65535
'Print &H8000&
'32768
'Print &H10000
'65536
'Print &HFFFF0000
'-65536
    MAKELPARAM = (cy) * &H10000 + cx
End Function
Function MAKELONG(ByVal a As Integer, ByVal b As Integer) As Long
'#define MAKELONG(a, b) \
'    ((LONG) (((WORD) (a)) | ((DWORD) ((WORD) (b))) << 16))
'Print &HFFFF&
'65535
'Print &H8000&
'32768
'Print &H10000
'65536
'Print &HFFFF0000
'-65536
'MAKELONG = a Or ShiftLeft((b), 16)
MAKELONG = b Or ShiftLeft(a, 16)
End Function
Function HiWord(lDword As Long) As Integer
    HiWord = (lDword And &HFFFF0000) \ &H10000
End Function

Function LoWord(lDword As Long) As Integer
    If lDword And &H8000& Then
        LoWord = lDword Or &HFFFF0000
    Else
        LoWord = lDword And &HFFFF&
    End If
End Function
'Function HiWord(lDword As Long) As Integer
'    HiWord = (lDword And &HFFFF0000) \ &H10000
'End Function
'
'Function LoWord(lDword As Long) As Integer
'    If lDword And &H8000& Then
'        LoWord = lDword Or &HFFFF0000
'    Else
'        LoWord = lDword And &HFFFF&
'    End If
'End Function

Function IsShiftDown() As Boolean
Dim Ret As Integer
' Const KEY_TOGGLED As Integer = &H1
' Const KEY_DOWN As Integer = &H1000
Ret = GetKeyState(VK_SHIFT)
IsShiftDown = Ret And KEY_DOWN
End Function

Public Function PtrToString(ByVal lParam As Long) As String
    Dim i As Long, b As Byte
    i = lParam
    CopyMemory b, ByVal i, 1
    Do While b <> 0
        PtrToString = PtrToString + Chr(b)
        i = i + 1
        CopyMemory b, ByVal i, 1
    Loop
End Function

Public Function PtrToStringW(ByVal lParam As Long) As String
    Dim i As Long, b As Integer
    i = lParam
    CopyMemory b, ByVal i, 2
    Do While b <> 0
        PtrToStringW = PtrToStringW + ChrW(b)
        i = i + 2
        CopyMemory b, ByVal i, 2
    Loop
End Function

Public Function CStrToString(bArray() As Byte) As String
    Dim i As Long
    On Error Resume Next
    i = -1
    i = LBound(bArray)
    If i < 0 Then Exit Function
    For i = LBound(bArray) To UBound(bArray)
        If bArray(i) = 0 Then Exit For
        CStrToString = CStrToString + Chr(bArray(i))
    Next i
End Function

Public Function StringToCStr(sData As String, bArray() As Byte, Optional ByVal Fixed As Boolean = True)
    Dim i As Long, _
        l As Long
    If (Not Fixed) Then
        ReDim bArray(0 To Len(sData))
    Else
        l = -1&
        l = LBound(bArray)
    End If
    For i = 1 To Len(sData)
        bArray(l) = Asc(Mid(sData, i, 1))
        l = l + 1
    Next i
    bArray(l) = 0&
End Function

Public Function CStrToStringW(iArray() As Integer) As String
    Dim i As Long
    On Error Resume Next
    i = -1
    i = LBound(iArray)
    If i < 0 Then Exit Function
    For i = LBound(iArray) To UBound(iArray)
        If iArray(i) = 0 Then Exit For
        CStrToStringW = CStrToStringW + ChrW(iArray(i))
    Next i
End Function

Public Function StringToCStrW(sData As String, iArray() As Integer, Optional ByVal Fixed As Boolean = True)
    Dim i As Long, _
        l As Long
    If (Not Fixed) Then
        ReDim iArray(0 To Len(sData))
    Else
        l = -1&
        l = LBound(iArray)
    End If
    For i = 1 To Len(sData)
        iArray(l) = Asc(Mid(sData, i, 1))
        l = l + 1
    Next i
    iArray(l) = 0&
End Function
'Public Function LBAddItem(ByVal LB As Long, ByVal ItemText As String, Optional ByVal InsertBefore As Long = -1) As Long
''    If (IsWindow(LB) = 0&) Then
''        LBAddItem = -1&
''        Exit Function
''    End If
'    Dim b() As Byte
'
'    StringToCStr ItemText, b, False
'    LBAddItem = SendMessage(LB, LB_INSERTSTRING, InsertBefore, b(0))
'End Function
Function GetTextWidthPix(ByVal hwnd As Long, ByVal str As String, _
    Optional cx As Long, Optional cy As Long) As Long
Dim sz As Size, Ret As Long, lHdc As Long
    lHdc = GetDC(hwnd)
Ret = GetTextExtentPoint32(lHdc, str, Len(str), sz)
    If Ret <> 0 Then
        cx = sz.cx
        cy = sz.cy
    End If
ReleaseDC hwnd, lHdc
    GetTextWidthPix = cx
End Function
Function GetTextHeightPix(ByVal hwnd As Long, ByVal str As String, _
    Optional cx As Long, Optional cy As Long) As Long
Dim sz As Size, Ret As Long, lHdc As Long
    lHdc = GetDC(hwnd)
Ret = GetTextExtentPoint32(lHdc, str, Len(str), sz)
    If Ret <> 0 Then
        cx = sz.cx
        cy = sz.cy
    End If
ReleaseDC hwnd, lHdc
    GetTextHeightPix = cy
End Function
'Function LShift(ByVal x As Long, ByVal y As Long) As Double
'If y >= 0 Then
'    LShift = x * (2 ^ y)
'Else
'    LShift = x * (2 ^ -y)
'End If
'End Function
'Function RShift(ByVal x As Long, ByVal y As Long) As Double
'If x >= 0 Then
'    RShift = (x / 2) ^ y
'Else
'    RShift = (Abs(x) ^ y) * (Abs(x) / y)
'End If
'End Function

Public Function INDEXTOOVERLAYMASK(iOverlay As Long) As Long
  ' #define INDEXTOOVERLAYMASK(i)   ((i) << 8)
  INDEXTOOVERLAYMASK = iOverlay * (2 ^ 8)
End Function
