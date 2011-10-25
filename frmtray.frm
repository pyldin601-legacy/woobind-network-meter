VERSION 5.00
Begin VB.Form frmOK 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2895
   ClientLeft      =   10005
   ClientTop       =   0
   ClientWidth     =   5385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmtray.frx":0000
   ScaleHeight     =   193
   ScaleMode       =   0  'User
   ScaleWidth      =   359
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00161515&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   180
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   20
      Top             =   420
      Width           =   3495
      Begin VB.Label lPeak 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "peak: 0 kB/s"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   60
         TabIndex        =   21
         Top             =   0
         Width           =   780
      End
      Begin VB.Label lPeak2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "peak: 0 kB/s"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   165
         Left            =   120
         TabIndex        =   22
         Top             =   60
         Width           =   780
      End
   End
   Begin VB.PictureBox picDOWN 
      Height          =   195
      Left            =   2700
      Picture         =   "frmtray.frx":32744
      ScaleHeight     =   135
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox imStat 
      BackColor       =   &H00161515&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      Picture         =   "frmtray.frx":33514
      ScaleHeight     =   255
      ScaleWidth      =   975
      TabIndex        =   14
      Top             =   495
      Width           =   975
   End
   Begin VB.PictureBox picON 
      Height          =   195
      Left            =   2700
      Picture         =   "frmtray.frx":342E4
      ScaleHeight     =   135
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   1260
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picOFF 
      Height          =   195
      Left            =   2700
      Picture         =   "frmtray.frx":350B4
      ScaleHeight     =   135
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lMe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Версия 2.1 сборка"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3480
      TabIndex        =   19
      Top             =   2940
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lUl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 K"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   4140
      TabIndex        =   18
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label lDl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 K"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   2640
      TabIndex        =   17
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Скорость (Dl/Ul):"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   2340
      Width           =   1530
   End
   Begin VB.Label mAll 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00 грн"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   4080
      TabIndex        =   11
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label mMonth 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00 грн"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4080
      TabIndex        =   10
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label mDay 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00 грн"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4140
      TabIndex        =   9
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label mSeans 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00 грн"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4140
      TabIndex        =   8
      Top             =   900
      Width           =   975
   End
   Begin VB.Label lAll 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 K"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   2700
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lMonth 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 K"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label lDay 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 K"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   2640
      TabIndex        =   5
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label lSeans 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 K"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   900
      Width           =   1275
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   272
      X2              =   272
      Y1              =   172
      Y2              =   60
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   184
      X2              =   184
      Y1              =   172
      Y2              =   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Остаток трафика/кредит:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2250
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Передано за месяц:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2325
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Передано за день:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1260
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "За текущее соединение:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   900
      Width           =   2205
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      X1              =   16
      X2              =   344
      Y1              =   148
      Y2              =   148
   End
   Begin VB.Menu mnuNMODE 
      Caption         =   "nMode"
      Visible         =   0   'False
      Begin VB.Menu mnuALEFT 
         Caption         =   "Прикрепить к левому углу"
      End
      Begin VB.Menu mnuARIGHT 
         Caption         =   "Прикрепить к правому углу"
      End
      Begin VB.Menu mnuACENTER 
         Caption         =   "Прикрепить по центру"
      End
      Begin VB.Menu setp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Больше не отображать это окно"
      End
      Begin VB.Menu sett 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAttach 
         Caption         =   "Закрепить"
      End
   End
End
Attribute VB_Name = "frmOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*                          Woobind Network Meter                          *
'***************************************************************************
'*   Copyright (C) 2007 by Roman Gemini                                    *
'*   networkmeter@ukr.net                                                  *
'*                                                                         *
'*   This program is free software; you can redistribute it and/or modify  *
'*   it under the terms of the GNU General Public License as published by  *
'*   the Free Software Foundation; either version 2 of the License, or     *
'*   (at your option) any later version.                                   *
'*                                                                         *
'*   This program is distributed in the hope that it will be useful,       *
'*   but WITHOUT ANY WARRANTY; without even the implied warranty of        *
'*   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the         *
'*   GNU General Public License for more details.                          *
'*                                                                         *
'*   You should have received a copy of the GNU General Public License     *
'*   along with this program; if not, write to the                         *
'*   Free Software Foundation, Inc.,                                       *
'*   59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.             *
'***************************************************************************

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const SW_SHOW = 5
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Dim mRECT As RECT
Dim RMR As Integer
Dim OldX As Long, ChangedX As Boolean

Enum FormStateConstants
  Hided = 0
  Showed = 1
  Hiding = 2
End Enum

Dim FormState As FormStateConstants
Dim SpeedArray(0 To 230) As Currency
Dim tmeMemory As Integer



Private Sub Form_DblClick()
      frmVelton.Visible = Not frmVelton.Visible
      If frmVelton.Visible = True Then frmVelton.SetFocus
End Sub

Sub DrawGraph(inValue)
    Dim ix As Integer, imax As Currency
    
    ' Смещение шкалы
    tmeMemory = (tmeMemory + 1) Mod 60
    
    ' Сдвигаю график влево
    For ix = 1 To 230
        SpeedArray(ix - 1) = SpeedArray(ix)
    Next ix
    
    ' Заношу новое значение
    Let SpeedArray(230) = inValue
    
    ' Определяю максимум
    imax = 1000
    For ix = 0 To 230
        If SpeedArray(ix) > imax Then imax = SpeedArray(ix)
    Next ix
    
    lPeak.Caption = "^ " & frmVelton.DataDevi(imax) + localize_do("WRD003A", "/с")
    
    ' Рисую таблицу
    pGraph.Cls
    
    For N = 0 To 300 Step 15
        If N Mod 60 = 0 Then
            pGraph.Line (N - tmeMemory, 0)-(N - tmeMemory, pGraph.ScaleHeight), RGB(60, 60, 60)
        ElseIf (N Mod 60 = 15) Or (N Mod 60 = 45) Then
            pGraph.Line (N - tmeMemory, 0)-(N - tmeMemory, pGraph.ScaleHeight), RGB(35, 35, 35)
        Else
            pGraph.Line (N - tmeMemory, 0)-(N - tmeMemory, pGraph.ScaleHeight), RGB(40, 40, 40)
        End If
    Next N
    
    pGraph.Line (0, 0)-(232, 0), RGB(60, 60, 60)
    pGraph.Line (0, 16)-(232, 16), RGB(60, 60, 60)
    pGraph.Line (0, 8)-(232, 8), RGB(60, 60, 60)
    
    pGraph.Line (0, 0)-(0, 17), RGB(60, 60, 60)
    pGraph.Line (pGraph.Width - 1, 0)-(pGraph.Width - 1, 17), RGB(60, 60, 60)
    
    ' Рисую график
    For ix = 0 To 230
        Select Case SpeedArray(ix)
        Case 0
            pGraph.Line (ix + 1, pGraph.ScaleHeight - 2)-(ix + 1, (pGraph.Height - 3) - (pGraph.Height - 3) / imax * SpeedArray(ix)), RGBBright(vbGreen, 150)
        Case imax
            pGraph.Line (ix + 1, pGraph.ScaleHeight - 2)-(ix + 1, (pGraph.Height - 3) - (pGraph.Height - 3) / imax * SpeedArray(ix)), RGBBright(vbGreen, 250)
        Case Else
            pGraph.Line (ix + 1, pGraph.ScaleHeight - 2)-(ix + 1, (pGraph.Height - 3) - (pGraph.Height - 3) / imax * SpeedArray(ix)), RGBBright(vbGreen, 200)
        End Select
    Next ix

    
End Sub

Sub InitOkForm()

ChangedX = False

mnuAttach.Checked = FWAlwaysVisible
Me.Top = -Me.Height + 25
FormState = Hided

OnTopForm frmOK, FWOnTop

Dim ChForm As RECT

ChForm = GetScreen

If Me.Left / 15 < ChForm.Left Then Me.Left = ChForm.Left * 15
If (Me.Left + Me.Width) / 15 > ChForm.Right Then Me.Left = (ChForm.Right * 15) - Me.Width

NormalWindowStyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
SetWindowLong Me.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
SetLayeredWindowAttributes Me.hWnd, 0, 5, LWA_ALPHA

End Sub

Private Sub Form_Load()

On Error Resume Next

lMe.Caption = "Версия " + GetVersion
LoadPos
DrawGraph 0

InitOkForm


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then OldX = x
If Button = 2 Then
 mnuALEFT.Caption = localize_do("FLOAT_WINDOW_01", "Прикрепить к левому краю")
 mnuARIGHT.Caption = localize_do("FLOAT_WINDOW_02", "Прикрепить к правому краю")
 mnuACENTER.Caption = localize_do("FLOAT_WINDOW_03", "Прикрепить по центру")
 mnuHide.Caption = localize_do("FLOAT_WINDOW_04", "Больше не отображать это окно")
 PopupMenu Me.mnuNMODE, , , , mnuHide
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim ML, MT

If Button = 1 Then

    Dim desktop As RECT: desktop = GetScreen
    ML = Me.Left + (x - OldX)
    If ML <= (desktop.Left * 15) + 120 Then ML = (1 + desktop.Left) * 15
    If ML > (desktop.Right * 15 - Me.Width - 120) Then ML = ((desktop.Right - 1) * 15) - Me.Width
    ChangedX = True
    Me.Left = ML

End If

If Button = 0 And FormState = Hided Then RMR = 1

End Sub

Sub LoadPos()

Dim IRP As String
IRP = Format(UniqMark, "000000000")
Me.Left = 15 * Val(GetSettingFake("Network Meter\" + IRP, "Window Position", "Left", Format(Screen.Width - Me.Width - 1500, "0")))

End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If Button = 1 Then
    Dim IRP As String
    IRP = def_complete_path(App.Path) + "app_data.ini"
    WritePrivateProfileString "Window Position", "Left", CStr(Fix(frmOK.Left / 15)), IRP
End If

End Sub

Private Sub lPeak_Change()
    lPeak2.Left = lPeak.Left + 1
    lPeak2.Top = lPeak.Top + 1
    lPeak2.Caption = lPeak.Caption
End Sub

Private Sub mnuACENTER_Click()

Dim UZ As RECT, B As Long

B = SystemParametersInfo(SPI_GETWORKAREA, 0, UZ, WM_SETTINGCHANGE)

Me.Move ((UZ.Right + UZ.Left) * 15 / 2) - (Me.Width / 2)
ChangedX = True
End Sub

Private Sub mnuALEFT_Click()

Dim UZ As RECT, B As Long

B = SystemParametersInfo(SPI_GETWORKAREA, 0, UZ, WM_SETTINGCHANGE)

Me.Move UZ.Left
ChangedX = True
End Sub

Private Sub mnuARIGHT_Click()

Dim UZ As RECT, B As Long

B = SystemParametersInfo(SPI_GETWORKAREA, 0, UZ, WM_SETTINGCHANGE)

Me.Move UZ.Right * 15 - Me.Width
ChangedX = True
End Sub

Sub tmrhyd_Timer()

On Error Resume Next

If mnuAttach.Checked = True And RMR = 0 Then RMR = 1

If ((IsMouseInWindow(Me.hWnd) And FormState = Hided) Or mnuAttach.Checked) And IIf(FWOnline, ProgConnected, True) Then
  If FWrmr = True Then
   If RMR < 4 And RMR > 0 Then RMR = RMR + 1
   If RMR = 4 Then ShowTray
  Else
   If RMR > 0 Then ShowTray
  End If
ElseIf Not IsMouseInWindow(Me.hWnd) And FormState = Showed And Not mnuAttach.Checked Then
   RMR = 0
   HideTray
End If


Dim UZ As RECT, B As Long

B = SystemParametersInfo(SPI_GETWORKAREA, 0, UZ, WM_SETTINGCHANGE)

If Not UZ.Right = mRECT.Right Or Not UZ.Left = mRECT.Left Or Not UZ.Top = mRECT.Top Or Not UZ.Bottom = mRECT.Bottom Then
   If FormState = Hided Then InitOkForm
End If

mRECT = UZ


End Sub



Sub HideTray()

FormState = Hiding

On Error Resume Next

For N = 0 To -Me.Height + 25 Step -100

    If FormState = Showed Then Exit For

    Let Me.Top = N
    Let y = 250 / (-Me.Height + 25) * N
    Call SetLayeredWindowAttributes(Me.hWnd, 0, 255 - y, LWA_ALPHA)
    Задержка 10

Next N

Me.Top = -Me.Height + 25

If ChangedX = True Then frmVelton.SaveToINI: ChangedX = False

FormState = Hided

End Sub

Sub ShowTray()

On Error Resume Next

Me.Top = 0
SetLayeredWindowAttributes Me.hWnd, 0, 250, LWA_ALPHA
FormState = Showed

End Sub

Private Sub mnuAttach_Click()
mnuAttach.Checked = Not mnuAttach.Checked
FWAlwaysVisible = mnuAttach.Checked
End Sub

Private Sub mnuHide_Click()

frmVelton.ShowBalloonParam "Woobind Network Meter", BalloonIcons.Information, localize_do("BALLOON02", "Всплывающее окно было отключено пользователем.%sДля того, чтобы вновь включить всплывающее окно, необходимо зайти в Опции программы."), vbCrLf
FloatWindow = False

frmVelton.SaveToINI
End Sub

Private Sub pGraph_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown Button, Shift, x, y
End Sub
Private Sub pGraph_Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseUp Button, Shift, x, y
End Sub
Private Sub pGraph_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub


Private Sub lPeak_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown Button, Shift, x, y
End Sub
Private Sub lPeak_Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseUp Button, Shift, x, y
End Sub
Private Sub lPeak_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub


