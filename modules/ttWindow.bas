Attribute VB_Name = "ttWindow"
Option Explicit
'Initialization of New ClassNames
Public Const ICC_BAR_CLASSES = &H4      'toolbar, statusbar, trackbar, tooltips
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Type tagINITCOMMONCONTROLSEX
   dwSize As Long   ' size of this structure
   dwICC As Long    ' flags indicating which classes to be initialized.
End Type


' ToolTip Styles
Public Const TTS_ALWAYSTIP = &H1
Public Const TTS_NOPREFIX = &H2
Public Const TTS_BALLOON = &H40 ' comctl32.dll v5.8 require

Public Const CW_USEDEFAULT = &H80000000

Public Const WS_POPUP = &H80000000

Public Const WM_USER = &H400

' ToolTip Messages
Public Const TTM_SETDELAYTIME = (WM_USER + 3)
Public Const TTM_ADDTOOL = (WM_USER + 4)
Public Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Public Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Public Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)

Public Const TTDT_AUTOPOP = 2
Public Const TTDT_INITIAL = 3

Public Const TTF_IDISHWND = &H1
Public Const TTF_CENTERTIP = &H2
Public Const TTF_SUBCLASS = &H10



Public Type TOOLINFO
    cbSize      As Long
    uFlags      As Long
    hwnd        As Long
    uId         As Long
    cRect       As RECT
    hinst       As Long
    lpszText    As String
End Type

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public bCreated As Boolean, hTT As Long
Public hCreated() As Long

Public Sub CreateTTWindow(hParent As Long, Optional bBalloon As Boolean = False)
  Dim h As Long, lStyle As Long
  lStyle = TTS_NOPREFIX Or TTS_ALWAYSTIP
  If bBalloon Then lStyle = lStyle Or TTS_BALLOON
  hTT = CreateWindowEx(0, "tooltips_class32", 0, lStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, hParent, 0, App.hInstance, 0)
  If Not bCreated Then
     ReDim hCreated(0)
     bCreated = True
  Else
     ReDim Preserve hCreated(UBound(hCreated) + 1)
  End If
  hCreated(UBound(hCreated)) = hTT
End Sub

Public Sub SetToolTip(objTT As Object, sTipText As String, _
                      Optional BkColor As Long = &HEEFFFF, _
                      Optional TxtColor As Long = vbBlack, _
                      Optional MaxWidth As Long = 300, _
                      Optional DelayTime As Long = 500, _
                      Optional VisibleTime As Long = 2000, _
                      Optional bCenter As Boolean = False)
    Dim TI As TOOLINFO
    With TI
        GetClientRect objTT.hwnd, .cRect
        .hwnd = objTT.hwnd
        .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
        If bCenter Then
            .uFlags = .uFlags Or TTF_CENTERTIP
        End If
        .uId = objTT.hwnd
        .lpszText = sTipText
        .cbSize = Len(TI)
    End With
    SendMessageLong hTT, TTM_SETMAXTIPWIDTH, 0, MaxWidth
    SendMessageLong hTT, TTM_SETDELAYTIME, TTDT_INITIAL, DelayTime
    SendMessageLong hTT, TTM_SETDELAYTIME, TTDT_AUTOPOP, VisibleTime
    SendMessageLong hTT, TTM_SETTIPTEXTCOLOR, TxtColor, 0&
    SendMessageLong hTT, TTM_SETTIPBKCOLOR, BkColor, 0&
    SendMessage hTT, TTM_ADDTOOL, 0, TI
End Sub

Public Sub DestroyTT()
  If Not bCreated Then Exit Sub
  Dim i As Integer
  For i = 0 To UBound(hCreated)
      DestroyWindow hCreated(i)
  Next
End Sub

Public Function InitComctl32(dwFlags As Long) As Boolean
   Dim icc As tagINITCOMMONCONTROLSEX
   On Error GoTo Err_OldVersion
   icc.dwSize = Len(icc)
   icc.dwICC = dwFlags
   InitComctl32 = InitCommonControlsEx(icc)
   On Error GoTo 0
   Exit Function
Err_OldVersion:
   InitCommonControls
End Function

