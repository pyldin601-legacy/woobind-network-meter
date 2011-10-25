VERSION 5.00
Begin VB.Form superMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3885
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   420
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3420
      Top             =   780
   End
   Begin VB.Image Image3 
      Height          =   15
      Index           =   3
      Left            =   780
      Picture         =   "superMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   2985
   End
   Begin VB.Label itmMnu 
      BackStyle       =   0  'Transparent
      Caption         =   "Посетить форум программы..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   960
      TabIndex        =   12
      Top             =   2460
      Width           =   2715
   End
   Begin VB.Label itmMnu 
      BackStyle       =   0  'Transparent
      Caption         =   "Проверка обновления"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   960
      TabIndex        =   11
      Top             =   2760
      Width           =   3075
   End
   Begin VB.Label lbCh 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   540
      TabIndex        =   9
      Top             =   3100
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   15
      Index           =   2
      Left            =   815
      Picture         =   "superMenu.frx":07C4
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3000
   End
   Begin VB.Image Image3 
      Height          =   15
      Index           =   0
      Left            =   810
      Picture         =   "superMenu.frx":0F88
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   2985
   End
   Begin VB.Image Image3 
      Height          =   15
      Index           =   1
      Left            =   815
      Picture         =   "superMenu.frx":174C
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2985
   End
   Begin VB.Label itmMnu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MONITORING CONNECTION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   900
      TabIndex        =   8
      Top             =   60
      Width           =   3090
   End
   Begin VB.Label itmMnu 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit programm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   960
      TabIndex        =   7
      Top             =   3600
      Width           =   2955
   End
   Begin VB.Label itmMnu 
      BackStyle       =   0  'Transparent
      Caption         =   "Start with Windows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   960
      TabIndex        =   6
      Top             =   3180
      Width           =   2715
   End
   Begin VB.Label itmMnu 
      BackStyle       =   0  'Transparent
      Caption         =   "About program"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   960
      TabIndex        =   5
      Top             =   2160
      Width           =   2940
   End
   Begin VB.Label itmMnu 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonuses..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   960
      TabIndex        =   4
      Top             =   1740
      Width           =   3060
   End
   Begin VB.Label itmMnu 
      BackStyle       =   0  'Transparent
      Caption         =   "Tariffs..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   3105
   End
   Begin VB.Label itmMnu 
      BackStyle       =   0  'Transparent
      Caption         =   "Options..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   2
      Top             =   1140
      Width           =   3135
   End
   Begin VB.Label itmMnu 
      BackStyle       =   0  'Transparent
      Caption         =   "Show report...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   3105
   End
   Begin VB.Label itmMnu 
      BackStyle       =   0  'Transparent
      Caption         =   "Show/Hide window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2955
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   465
      Picture         =   "superMenu.frx":1F10
      Top             =   450
      Width           =   3570
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   4140
      Left            =   0
      Picture         =   "superMenu.frx":5A10
      Top             =   -240
      Width           =   810
   End
End
Attribute VB_Name = "superMenu"
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

Private Const SW_SHOW = 5
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Dim Hiding As Boolean

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_SHOWWINDOW = &H40

Dim ShiftPressed As Boolean

Private Sub Check1_Click()
If Check1.Value > 0 Then lbCh.Visible = True Else lbCh.Visible = False
End Sub


Private Sub Form_Load()

On Error Resume Next

ShowAllForm

Dim N As Long


Me.Visible = False

End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim h As Integer
For h = 1 To itmMnu.Count - 1
  If GetDistanceToControl(Me, itmMnu(h)) <= 50 Then itmMnu_MouseMove h, Button, Shift, x, y: Exit For
  If itmMnu(h).Top - 50 < y And itmMnu(h).Top + itmMnu(h).Height + 50 > y And x >= Image2.Left Then itmMnu_MouseMove h, Button, Shift, x, y: Exit For
Next h

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim h As Integer
For h = 1 To itmMnu.Count - 1
  If GetDistanceToControl(Me, itmMnu(h)) <= 50 Then itmMnu_Click h: Exit For
  If itmMnu(h).Top - 50 < y And itmMnu(h).Top + itmMnu(h).Height + 50 > y And x >= Image2.Left Then itmMnu_Click h: Exit For
Next h
End Sub



Private Sub itmMnu_Click(index As Integer)

On Error Resume Next

If index = 0 Then Exit Sub

Select Case index

Case 1
 frmVelton.mnuShow_Click
 
Case 2
 frmVelton.mStatistic_Click

Case 3
 frmVelton.mnuSetup_Click
 
Case 4
 frmVelton.mnuTaxes_Click
 
Case 5
 frmVelton.mnuBonus_Click
 
Case 6
 frmVelton.mnuAbout_Click

Case 7
 frmVelton.mnuAutorun_Click

Case 8
 ShowSelected
 HideMe
 If NoExit Then
    '
 Else
    frmVelton.mExit_Click
 End If
 Exit Sub

Case 9
 ShowSelected
 HideMe
 DoEvents
 
 CheckUpdate True
    
 Exit Sub

Case 10
 RunWEB "http://woobind.org.ua/forum"
 
End Select

ShowSelected
HideMe


End Sub

Private Sub itmMnu_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If index > 0 And Not Image2.Top = itmMnu(index).Top - 40 Then Image2.Top = itmMnu(index).Top - 40
End Sub

Private Sub lbCh_Click()
itmMnu_Click 7
End Sub

Private Sub lbCh_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
itmMnu_MouseMove 7, Button, Shift, x, y
End Sub

Private Sub Timer1_Timer()
If DistanceToWindow(Me.hWnd) > 300 And Hiding = False Then HideWindow
End Sub


Sub HideMe()
On Error Resume Next

Dim FlyAway As Long: Let FlyAway = 1
Dim OldWX, OldWY

OldWX = Me.Left
OldWY = Me.Top

ReleaseCapture
Hiding = True
For N = 250 To 0 Step -20
SetLayeredWindowAttributes Me.hWnd, 0, N, LWA_ALPHA
Me.Move OldWX, OldWY - (250 - N)


Задержка 20
' FlyAway = FlyAway * 2
If Hiding = False Then Exit Sub
Next

Me.Visible = False
Timer1.Enabled = False

End Sub

Sub HideWindow()
On Error Resume Next

Dim FlyAway As Long: Let FlyAway = 1
Dim OldWX, OldWY

OldWX = Me.Left
OldWY = Me.Top

ReleaseCapture
Hiding = True

For N = 250 To 0 Step -20
SetLayeredWindowAttributes Me.hWnd, 0, N, LWA_ALPHA


Задержка 20
' FlyAway = FlyAway * 2
If Hiding = False Then Exit Sub
Next

Me.Visible = False
Timer1.Enabled = False

End Sub


Sub ShowAllForm()

Dim ret As Long
ret = CreateRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY)
SetWindowRgn Me.hWnd, ret, True


End Sub

Sub ShowSelected()

Dim ret As Long

Dim X1, X2, Y1, Y2


X1 = Image2.Left / Screen.TwipsPerPixelX
Y1 = Image2.Top / Screen.TwipsPerPixelY
X2 = (Image2.Left + Image2.Width) / Screen.TwipsPerPixelX
Y2 = (Image2.Top + Image2.Height) / Screen.TwipsPerPixelY

ret = CreateRectRgn(X1 + 1, Y1 + 1, X2 + 1, Y2 + 1)
SetWindowRgn Me.hWnd, ret, True



End Sub


Sub PositeMe()

On Error Resume Next
Dim P As POINTAPI, N As Long
Hiding = False
Image2.Top = -Image2.Height
N = GetCursorPos(P)

Me.MousePointer = 0

superMenu.Check1.Value = Bol2Int(CheckAutorun)

Dim ChForm As RECT

ChForm = GetScreen

Dim dX, dY

dY = (P.y - ChForm.Bottom) * tppY
dX = (P.x - ChForm.Right) * tppX

If P.y + (Me.Height / Screen.TwipsPerPixelY) <= ChForm.Bottom Then
    Me.Top = P.y * Screen.TwipsPerPixelY
Else
    Me.Top = (P.y * Screen.TwipsPerPixelY) - Me.Height - dY
End If

If P.x + Me.Width / Screen.TwipsPerPixelX <= ChForm.Right Then
    Me.Left = P.x * Screen.TwipsPerPixelX
Else
    Me.Left = P.x * Screen.TwipsPerPixelX - Me.Width - dX
End If

ShowAllForm


NormalWindowStyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
SetWindowLong Me.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_ALPHA

OnTopForm Me, True
Me.Visible = True
Putfocus Me.hWnd

If Err Then Me.Hide Else SetCapture Me.hWnd

Timer1.Enabled = True

End Sub
