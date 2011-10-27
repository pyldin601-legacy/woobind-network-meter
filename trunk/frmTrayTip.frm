VERSION 5.00
Begin VB.Form frmTrayTip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   126
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   258
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrE 
      Interval        =   100
      Left            =   3060
      Top             =   0
   End
   Begin VB.PictureBox skinbuff 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   0
      Picture         =   "frmTrayTip.frx":0000
      ScaleHeight     =   1185
      ScaleWidth      =   285
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lSpDn 
      BackStyle       =   0  'Transparent
      Caption         =   "123.00  байт/c"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   420
      TabIndex        =   10
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Label lSpUl 
      BackStyle       =   0  'Transparent
      Caption         =   "123.00  байт/c"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2340
      TabIndex        =   9
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frmTrayTip.frx":0264
      Top             =   1620
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2040
      Picture         =   "frmTrayTip.frx":05EE
      Top             =   1620
      Width           =   240
   End
   Begin VB.Label lDesc2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Поточна швидкість"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1380
      Width           =   1470
   End
   Begin VB.Label lMon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Потрачено:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label lCash 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "234.55 грн."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2580
      TabIndex        =   6
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label lR 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "123 456 789 900 байт"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1620
      TabIndex        =   5
      Top             =   840
      Width           =   2100
   End
   Begin VB.Label lS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "123 456 789 900 байт"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1635
      TabIndex        =   4
      Top             =   600
      Width           =   2100
   End
   Begin VB.Label lDesc1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Трафік за поточне з'єднання"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lRecv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Прийнято:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   840
      Width           =   795
   End
   Begin VB.Label lSent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Передано:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   600
      Width           =   810
   End
   Begin VB.Label lConnection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Назва з'єднання"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3750
   End
End
Attribute VB_Name = "frmTrayTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX, OldY
Dim Trys  As Integer


Sub ShowTip(strTip As String, x, y)

    Dim tmpReg As RECT
    Dim tx As POINTAPI
    Dim TP As RECT
    Call GetCursorPos(tx)
    
    OldX = tx.x
    OldY = tx.y
    x = OldX * tppX
    y = OldY * tppY
    tmpReg = GetWorkArea
    tmrE.Enabled = True
    
    TP = GetTaskbarPos
    
    If TP.Top > tmpReg.Bottom / 2 Then
        Me.Move TP.Right * Screen.TwipsPerPixelX - Me.Width - 30, TP.Top * tppY - Me.Height - 30
    Else
        Me.Move TP.Right * Screen.TwipsPerPixelX - Me.Width - 30, TP.Bottom * tppY + 30
    End If
    
    Trys = 5
    
End Sub



Sub tip_load_skin()
    
    Call StretchBlt(Me.hdc, 0, 0, 6, 6, skinbuff.hdc, 0, 0, 6, 6, vbSrcCopy)
    Call StretchBlt(Me.hdc, Me.ScaleWidth - 6, Me.ScaleHeight - 6, 6, 6, skinbuff.hdc, skinbuff.Width - 6, skinbuff.Height - 6, 6, 6, vbSrcCopy)
    Call StretchBlt(Me.hdc, 0, Me.ScaleHeight - 6, 6, 6, skinbuff.hdc, 0, skinbuff.Height - 6, 6, 6, vbSrcCopy)
    Call StretchBlt(Me.hdc, Me.ScaleWidth - 6, 0, 6, 6, skinbuff.hdc, skinbuff.Width - 6, 0, 6, 6, vbSrcCopy)
    Call StretchBlt(Me.hdc, 6, 0, Me.ScaleWidth - 12, 6, skinbuff.hdc, 6, 0, skinbuff.Width - 12, 6, vbSrcCopy)
    Call StretchBlt(Me.hdc, 0, 6, 6, Me.ScaleHeight - 12, skinbuff.hdc, 0, 6, 6, skinbuff.Height - 12, vbSrcCopy)
    Call StretchBlt(Me.hdc, Me.ScaleWidth - 6, 6, 6, Me.ScaleHeight - 12, skinbuff.hdc, skinbuff.Width - 6, 6, 6, skinbuff.Height - 12, vbSrcCopy)
    Call StretchBlt(Me.hdc, 6, Me.ScaleHeight - 6, Me.ScaleWidth - 12, 6, skinbuff.hdc, 6, skinbuff.Height - 6, skinbuff.Width - 12, 6, vbSrcCopy)
    Call StretchBlt(Me.hdc, 6, 6, Me.ScaleWidth - 12, Me.ScaleHeight - 12, skinbuff.hdc, 6, 6, skinbuff.Width - 12, skinbuff.Height - 12, vbSrcCopy)
    Me.Picture = Me.Image

End Sub

Private Sub Form_Load()

    SetFormTColorXP Me, vbMagenta, 255
    Call tip_load_skin
    
End Sub

Sub Rescan()

        lConnection.Caption = Left(filter_interface_name(localize_do(iph_interface, iph_interface)), 45) + " (" + NetworkText(NetworkStatus(0)) + ")"
        lSent.Caption = localize_do("WRD014", "Отправлено") & ":"
        lRecv.Caption = localize_do("WRD015", "Получено") + ":"
        lS.Caption = FormatEx(DataXmited, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
        lR.Caption = FormatEx(DataRcved, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
        lSpDn.Caption = frmVelton.DataDevi(CCur(tmpDL)) & localize_do("WRD003A", "/с")
        lSpUl.Caption = frmVelton.DataDevi(CCur(tmpUL)) & localize_do("WRD003A", "/с")
        lCash.Caption = FormatEx(TaxTax, "### ##0.00") + " " + TaxName
        lMon.Caption = localize_do("ADD005", "Начислено")
        lDesc1.Caption = localize_do("WNDA01", "За текущее соединение:")
        lDesc2.Caption = localize_do("WNDA05", "Скор. прием/передача")

End Sub

Private Sub tmrE_Timer()
    
   Dim tx As POINTAPI
    Call GetCursorPos(tx)
    
    If tx.x <> OldX Or tx.y <> OldY Then
        tmrE.Enabled = False
        Me.Hide
    Else
        If Trys > 0 Then Trys = Trys - 1
        If (Trys = 0 And Me.Visible = False And Not superMenu.Visible) Then
            'FillIn Me
            Rescan
            OnTopForm Me, True
        End If
    End If
    
End Sub
