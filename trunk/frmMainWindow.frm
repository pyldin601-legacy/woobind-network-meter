VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmVelton 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Woobind Network Meter hWnd"
   ClientHeight    =   5820
   ClientLeft      =   4500
   ClientTop       =   4785
   ClientWidth     =   9255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainWindow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainWindow.frx":0CCA
   ScaleHeight     =   5820
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox imgCheck 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   660
      Picture         =   "frmMainWindow.frx":B03FE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   69
      Top             =   300
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox imgPanel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   4
      Left            =   1980
      MouseIcon       =   "frmMainWindow.frx":B0788
      MousePointer    =   99  'Custom
      Picture         =   "frmMainWindow.frx":B0A92
      ScaleHeight     =   345
      ScaleWidth      =   1335
      TabIndex        =   66
      Top             =   5200
      Width           =   1335
   End
   Begin VB.PictureBox imgPanel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   0
      Left            =   480
      MouseIcon       =   "frmMainWindow.frx":B22EA
      MousePointer    =   99  'Custom
      Picture         =   "frmMainWindow.frx":B25F4
      ScaleHeight     =   345
      ScaleWidth      =   1335
      TabIndex        =   64
      Top             =   5200
      Width           =   1335
   End
   Begin MSComctlLib.ImageList imgNetLevel 
      Left            =   1140
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B3E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B43A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B4900
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B4E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B53B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B590E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B5E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B63C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B691C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B6E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B73D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B792A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B7E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainWindow.frx":B83DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   7
      Left            =   1980
      Picture         =   "frmMainWindow.frx":B8938
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   60
      Top             =   6660
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   6
      Left            =   1980
      Picture         =   "frmMainWindow.frx":B8CC2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   59
      Top             =   6180
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   5160
      Top             =   180
   End
   Begin VB.PictureBox clsD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   7080
      Picture         =   "frmMainWindow.frx":B904C
      ScaleHeight     =   270
      ScaleWidth      =   645
      TabIndex        =   57
      Top             =   300
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox clsD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   6360
      Picture         =   "frmMainWindow.frx":B99D8
      ScaleHeight     =   270
      ScaleWidth      =   645
      TabIndex        =   56
      Top             =   300
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox clsD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   5640
      Picture         =   "frmMainWindow.frx":BA364
      ScaleHeight     =   270
      ScaleWidth      =   645
      TabIndex        =   55
      Top             =   300
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   8380
      Picture         =   "frmMainWindow.frx":BACF0
      ScaleHeight     =   270
      ScaleWidth      =   645
      TabIndex        =   54
      ToolTipText     =   "Свернуть в трей"
      Top             =   0
      Width           =   645
   End
   Begin VB.PictureBox picTray 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4560
      Picture         =   "frmMainWindow.frx":BB67C
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   51
      Top             =   4620
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   480
      Picture         =   "frmMainWindow.frx":BC346
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   50
      Top             =   6180
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   960
      Picture         =   "frmMainWindow.frx":BC6D0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   49
      Top             =   6180
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   1440
      Picture         =   "frmMainWindow.frx":BCA5A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   48
      Top             =   6660
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   480
      Picture         =   "frmMainWindow.frx":BCDE4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   47
      Top             =   6660
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   960
      Picture         =   "frmMainWindow.frx":BD16E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   46
      Top             =   6660
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   1440
      Picture         =   "frmMainWindow.frx":BD4F8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   45
      Top             =   6180
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ограничение трафика"
      Height          =   795
      Left            =   4920
      TabIndex        =   40
      Top             =   4080
      Width           =   4035
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0 Б"
         Height          =   195
         Left            =   1860
         TabIndex        =   44
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblLimit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0 Б"
         Height          =   195
         Left            =   1860
         TabIndex        =   43
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Остаток в лимите:"
         Height          =   195
         Left            =   60
         TabIndex        =   42
         Top             =   480
         Width           =   1875
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ограничение:"
         Height          =   195
         Left            =   60
         TabIndex        =   41
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   235
      Left            =   6600
      Picture         =   "frmMainWindow.frx":BD882
      ScaleHeight     =   240
      ScaleWidth      =   975
      TabIndex        =   27
      Top             =   620
      Width           =   975
   End
   Begin VB.PictureBox picA 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00161515&
      BorderStyle     =   0  'None
      Height          =   235
      Left            =   7700
      Picture         =   "frmMainWindow.frx":BE652
      ScaleHeight     =   240
      ScaleWidth      =   975
      TabIndex        =   26
      Top             =   620
      Width           =   975
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Сведения"
      Height          =   2895
      Left            =   4920
      TabIndex        =   10
      Top             =   1080
      Width           =   4035
      Begin VB.Label lAwait 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   2160
         TabIndex        =   63
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ожидается к концу месяца:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   62
         Top             =   1560
         Width           =   2520
      End
      Begin VB.Label lblBonLeft 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         TabIndex        =   53
         Top             =   1920
         Width           =   1875
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ост. бонуса (Dl/Ul)"
         Height          =   195
         Left            =   240
         TabIndex        =   52
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   1920
         TabIndex        =   37
         Top             =   2220
         Width           =   1875
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Деньги в лимите:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   2220
         Width           =   1530
      End
      Begin VB.Label lbTrader 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   2040
         TabIndex        =   35
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Суточный расход:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1260
         Width           =   1620
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   3780
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Label lblFlag 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4680
         TabIndex        =   25
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cуточная скорость:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1020
         Width           =   1725
      End
      Begin VB.Label lblspeedavg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   2040
         TabIndex        =   23
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label lblcml 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 Мбит"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   3180
         TabIndex        =   22
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Скорость подключения:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lbUlSpd 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   2160
         TabIndex        =   14
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label lbDlSpd 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   1980
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Скорость приема:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Скорость передачи:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1785
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   300
      TabIndex        =   3
      Top             =   2820
      Width           =   4455
      Begin VB.Timer tmrAction 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   4020
         Picture         =   "frmMainWindow.frx":BF422
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1980
         Picture         =   "frmMainWindow.frx":BF7AC
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label lblMaxx2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2220
         TabIndex        =   61
         Top             =   1680
         Width           =   1755
      End
      Begin VB.Label lblMaxx 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Макс. замеченная скорость:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1380
         Width           =   2505
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1740
         TabIndex        =   33
         Top             =   780
         Width           =   2475
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Сумма байт:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lbSinceSumm 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Начислено:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label lbSinceRecev 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1860
         TabIndex        =   7
         Top             =   540
         Width           =   2355
      End
      Begin VB.Label lbSinceSend 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1860
         TabIndex        =   6
         Top             =   300
         Width           =   2355
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Принято байт:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   540
         Width           =   1245
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Передано байт:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Начисления"
      Height          =   1635
      Left            =   300
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
      Begin VB.Label mSeans 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00 грн"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   2760
         TabIndex        =   31
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label mDay 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00 грн"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   2820
         TabIndex        =   30
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label mMonth 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00 грн"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   2820
         TabIndex        =   29
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label mAll 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00 грн"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   2820
         TabIndex        =   28
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label wnd004 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "за все время:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Label lbXmitedAll 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   1620
         TabIndex        =   19
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label wnd003 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "за месяц:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   18
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lbXmitedMonth 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   1620
         TabIndex        =   17
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label wnd002 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "за неделю:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lbXmitedWeek 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   1620
         TabIndex        =   15
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lbXmitedToday 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   1620
         TabIndex        =   2
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label wnd001 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "за сегодня:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.PictureBox imgPanel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   2
      Left            =   4260
      MouseIcon       =   "frmMainWindow.frx":BFB36
      MousePointer    =   99  'Custom
      Picture         =   "frmMainWindow.frx":BFE40
      ScaleHeight     =   435
      ScaleWidth      =   1470
      TabIndex        =   65
      Top             =   5160
      Width           =   1470
   End
   Begin VB.PictureBox imgPanel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   8
      Left            =   5730
      MouseIcon       =   "frmMainWindow.frx":C200C
      MousePointer    =   99  'Custom
      Picture         =   "frmMainWindow.frx":C2316
      ScaleHeight     =   435
      ScaleWidth      =   1470
      TabIndex        =   68
      Top             =   5160
      Width           =   1470
   End
   Begin VB.PictureBox imgPanel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   6
      Left            =   7200
      MouseIcon       =   "frmMainWindow.frx":C44E2
      MousePointer    =   99  'Custom
      Picture         =   "frmMainWindow.frx":C47EC
      ScaleHeight     =   435
      ScaleWidth      =   1470
      TabIndex        =   67
      Top             =   5160
      Width           =   1470
   End
   Begin VB.Image Image4 
      Enabled         =   0   'False
      Height          =   435
      Left            =   8670
      Picture         =   "frmMainWindow.frx":C69B8
      Top             =   5160
      Width           =   210
   End
   Begin VB.Image Image3 
      Enabled         =   0   'False
      Height          =   435
      Left            =   4045
      Picture         =   "frmMainWindow.frx":C6EF8
      Top             =   5160
      Width           =   210
   End
   Begin ComctlLib.ImageList ipImages 
      Left            =   3420
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   98
      ImageHeight     =   29
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":C7438
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":C8C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":CA504
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":CC6DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":CE8B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":D011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":D1984
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":D3B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":D5D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":D7F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":DA0EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainWindow.frx":DB952
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgLevel 
      Height          =   210
      Left            =   6060
      Picture         =   "frmMainWindow.frx":DD1B8
      ToolTipText     =   "Network Ping Level"
      Top             =   630
      Width           =   450
   End
   Begin VB.Label lMe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   60
      TabIndex        =   58
      Top             =   90
      Width           =   8265
   End
End
Attribute VB_Name = "frmVelton"
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


Dim dayBegin As Integer
Dim ConnSpd As Long

Dim credit_blinker_status As Boolean
Dim mouse_old_x, mouse_old_y
Dim MultiCounter As Long

Dim iph_received_previous As Currency, iph_transmitted_previous As Currency
Dim iph_received_result As Currency, iph_transmitted_result As Currency
Dim iph_received_delta As Currency, iph_transmitted_delta As Currency
Dim bonus_received_delta As Currency, bonus_transmitted_delta As Currency

Dim index As Long
Dim AverageSpeed(1, 39) As Long
Dim AverageReceiving As Long
Dim AverageXmiting As Long

Dim ResetDate As String

Dim MassSpeed(5, 3) As Long
Dim RememberMe As Boolean

Dim LastSpeed As Currency
Dim WasSent As Currency

Dim ActiveLimit As Currency
Dim ActiveLeft As Currency

Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000

Dim Zond1 As Currency
Dim Zond2 As Currency

Dim previous_traffic_tax As Currency

Dim TimeLongBack As Long
Dim Blinker As Boolean


Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

' variables for delayed applications launch
Dim delay_enabled As Boolean
Dim delay_counter As Integer

Sub delay_launch_set_timeout(timeout As Integer)
  delay_counter = timeout
  delay_enabled = True
End Sub


Function iph_get_traffic(inB As Currency, outB As Currency) As Boolean

On Error Resume Next

Dim d, N, tmpIn As Currency, tmpOut As Currency, tmpSpd As Currency
Dim GetTraficTemp As Boolean
Dim RASRcved As Currency, RASXmited As Currency, RASSpeed As Currency

inB = 0: outB = 0: iph_get_traffic = False: ConnSpd = 0

Select Case iph_interface

Case ALL_OFF
    
    inB = 0: outB = 0
    GetTraficTemp = False
    
Case ALL_INT
  For N = 1 To m_objIpHelper.Interfaces.Count
    If m_objIpHelper.Interfaces(N).InterfaceType = 6 And m_objIpHelper.Interfaces(N).Speed < NO_SPEED Then
      If tmpSpd < m_objIpHelper.Interfaces(N).Speed Then tmpSpd = m_objIpHelper.Interfaces(N).Speed
      tmpIn = tmpIn + m_objIpHelper.Interfaces(N).OctetsReceived
      tmpOut = tmpOut + m_objIpHelper.Interfaces(N).OctetsSent
      
    End If
  Next N
  
  If GetDialUp(RASXmited, RASRcved, RASSpeed) Then
      tmpIn = tmpIn + RASRcved
      tmpOut = tmpOut + RASXmited
      If tmpSpd < RASSpeed Then tmpSpd = RASSpeed
  End If
  
  ConnSpd = tmpSpd
  If ConnSpd > 0 Then GetTraficTemp = True
  inB = tmpIn
  outB = tmpOut
  
Case RAS_INT
  GetTraficTemp = GetDialUp(tmpOut, tmpIn, tmpSpd)
  ConnSpd = tmpSpd
  inB = tmpIn
  outB = tmpOut
  
Case LAN_INT
  For N = 1 To m_objIpHelper.Interfaces.Count
    If m_objIpHelper.Interfaces(N).InterfaceType = 6 And m_objIpHelper.Interfaces(N).Speed < NO_SPEED Then
      If tmpSpd < m_objIpHelper.Interfaces(N).Speed Then tmpSpd = m_objIpHelper.Interfaces(N).Speed
      tmpIn = tmpIn + m_objIpHelper.Interfaces(N).OctetsReceived
      tmpOut = tmpOut + m_objIpHelper.Interfaces(N).OctetsSent
    End If
  Next N
  ConnSpd = tmpSpd
  If ConnSpd > 0 Then GetTraficTemp = True
  inB = tmpIn
  outB = tmpOut
    
Case Else
  d = GetIndexFrom2(iph_interface)
  If d > 0 Then
    If m_objIpHelper.Interfaces(d).Speed < NO_SPEED Then
      inB = m_objIpHelper.Interfaces(d).OctetsReceived
      outB = m_objIpHelper.Interfaces(d).OctetsSent
      GetTraficTemp = True
      ConnSpd = m_objIpHelper.Interfaces(d).Speed
    End If
  End If
End Select

iph_get_traffic = GetTraficTemp

End Function





Function GetIndexFrom2(inName As String) As Integer
Dim r
For r = 1 To m_objIpHelper.Interfaces.Count
  If Mid(m_objIpHelper.Interfaces(r).InterfaceDescription, 1, Len(m_objIpHelper.Interfaces(r).InterfaceDescription) - 1) = inName Then
    GetIndexFrom2 = r
    Exit For
  End If
Next r
End Function

Function Prevail(DataXmitedMonth, DataRcvedMonth) As Integer

If DataXmitedMonth > DataRcvedMonth Then Prevail = 2 Else Prevail = 1

End Function

Sub SaveFormPosition()

Dim MeLeft As Long, MeTop As Long
Dim IRP As String

IRP = def_complete_path(App.Path) + "app_data.ini"

MeLeft = Me.Left / 15: MeTop = Me.Top / 15

WritePrivateProfileString "Window Position", "X", CStr(MeLeft), IRP
WritePrivateProfileString "Window Position", "Y", CStr(MeTop), IRP

End Sub

Sub UpdateAverage()
Dim avgMoney As Currency, avgData As Currency, avgData2 As Currency
Dim avgIn As Currency, avgOut As Currency
Dim MaxDet As Long, avgByte As Currency, nz As Currency
avgData2 = CountTrafficPerMonth(Month(Now), tax_taxing_traffic, Prevail(DataXmitedMonth, DataRcvedMonth), avgMoney)
avgData = AverageCap(Month(Now), Day(Now), tax_taxing_traffic, Prevail(DataXmitedMonth, DataRcvedMonth), nz, Year(Now))

lblspeedavg.Caption = "~" + DataDevi(avgData2) + localize_do("WRD001", "/день")

Select Case tax_taxing_traffic
Case 0
    lAwait.Caption = "~" + DataDevi((DataRcvedMonth + DataXmitedMonth) - (DataRcvedToday + DataXmitedToday) + (avgData * (GetDaysInMonth(Month(Now)) - Day(Now) + 1)))
Case 1
    lAwait.Caption = "~" + DataDevi(DataRcvedMonth - DataRcvedToday + (avgData * (GetDaysInMonth(Month(Now)) - Day(Now) + 1)))
Case 2
    lAwait.Caption = "~" + DataDevi(DataXmitedMonth - DataXmitedToday + (avgData * (GetDaysInMonth(Month(Now)) - Day(Now) + 1)))
Case 3
 If DataRcvedMonth > DataXmitedMonth Then
    lAwait.Caption = "~" + DataDevi(DataRcvedMonth - DataRcvedToday + (avgData * (GetDaysInMonth(Month(Now)) - Day(Now) + 1)))
 Else
    lAwait.Caption = "~" + DataDevi(DataXmitedMonth - DataXmitedToday + (avgData * (GetDaysInMonth(Month(Now)) - Day(Now) + 1)))
 End If
End Select

lbTrader.Caption = "~" + FormatEx(avgMoney, "0.00") + " " + TaxName + localize_do("WRD001", "/день")

End Sub

Sub UpdateFluentWindow()
  
If Mid(ComboLinks, 2, 1) = "0" Then
  frmOK.lDay.Caption = DataDevi(DataXmitedToday + DataRcvedToday)
ElseIf Mid(ComboLinks, 2, 1) = "1" Then
  frmOK.lDay.Caption = DataDevi(DataRcvedToday)
ElseIf Mid(ComboLinks, 2, 1) = "2" Then
  frmOK.lDay.Caption = DataDevi(DataXmitedToday)
ElseIf Mid(ComboLinks, 2, 1) = "3" Then
 If DataXmitedMonth > DataRcvedMonth Then
  frmOK.lDay.Caption = DataDevi(DataXmitedToday)
 Else
  frmOK.lDay.Caption = DataDevi(DataRcvedToday)
 End If
End If


If Mid(ComboLinks, 3, 1) = "0" Then
  frmOK.lMonth.Caption = DataDevi(DataXmitedMonth + DataRcvedMonth)
  frmOK.lMonth.ToolTipText = FormatEx(DataXmitedMonth + DataRcvedMonth, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
ElseIf Mid(ComboLinks, 3, 1) = "1" Then
  frmOK.lMonth.Caption = DataDevi(DataRcvedMonth)
  frmOK.lMonth.ToolTipText = FormatEx(DataRcvedMonth, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
ElseIf Mid(ComboLinks, 3, 1) = "2" Then
  frmOK.lMonth.Caption = DataDevi(DataXmitedMonth)
  frmOK.lMonth.ToolTipText = FormatEx(DataXmitedMonth, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
ElseIf Mid(ComboLinks, 3, 1) = "3" Then
 If DataXmitedMonth > DataRcvedMonth Then
  frmOK.lMonth.Caption = DataDevi(DataXmitedMonth)
  frmOK.lMonth.ToolTipText = FormatEx(DataXmitedMonth, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
 Else
  frmOK.lMonth.Caption = DataDevi(DataRcvedMonth)
  frmOK.lMonth.ToolTipText = FormatEx(DataRcvedMonth, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
 End If
End If


If Mid(ComboLinks, 1, 1) = "0" Then
  frmOK.lSeans.Caption = DataDevi(DataXmited + DataRcved)
  frmOK.lSeans.ToolTipText = FormatEx(DataXmited + DataRcved, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
ElseIf Mid(ComboLinks, 1, 1) = "1" Then
  frmOK.lSeans.Caption = DataDevi(DataRcved)
  frmOK.lSeans.ToolTipText = FormatEx(DataRcved, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
ElseIf Mid(ComboLinks, 1, 1) = "2" Then
  frmOK.lSeans.Caption = DataDevi(DataXmited)
  frmOK.lSeans.ToolTipText = FormatEx(DataXmited, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
ElseIf Mid(ComboLinks, 1, 1) = "3" Then
 If DataXmitedMonth > DataRcvedMonth Then
  frmOK.lSeans.Caption = DataDevi(DataXmited)
  frmOK.lSeans.ToolTipText = FormatEx(DataXmited, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
 Else
  frmOK.lSeans.Caption = DataDevi(DataRcved)
  frmOK.lSeans.ToolTipText = FormatEx(DataRcved, "### ### ### ##0") + " " + localize_do("WRD007", "Байт")
 End If
End If

frmOK.lDay.ToolTipText = DataDeviTip32(DataXmitedToday, DataRcvedToday)
frmOK.lMonth.ToolTipText = DataDeviTip32(DataXmitedMonth, DataRcvedMonth)
frmOK.lSeans.ToolTipText = DataDeviTip32(DataXmited, DataRcved)

End Sub











Sub RefreshFace()

If MainWindowAlpha = True Then
  SetFormTColorXP Me, RGB(240, 0, 120), 255 - (255 / 100 * MainWindowAlphaLVL)
Else
  SetFormTColorXP Me, RGB(240, 0, 120), 255
End If

LimitName(0) = localize_do("WRD001", "/день")
LimitName(1) = localize_do("WRD002", "/нед.")
LimitName(2) = localize_do("WRD003", "/мес.")
LimitName(3) = ""

LimitNameA(0) = localize_do("WRD004", "Оставшийся на сегодня объем трафика составляет")
LimitNameA(1) = localize_do("WRD005", "Оставшийся на эту неделю объем трафика составляет")
LimitNameA(2) = localize_do("WRD006", "Оставшийся на этот месяц объем трафика составляет")
LimitNameA(3) = localize_do("WRD006a", "Оставшийся объем трафика составляет")

Frame7.Caption = localize_do("MSG001", "Начиная с") + " " + ResetDate

wnd001.Caption = localize_do("WND001", "за сегодня:")
wnd002.Caption = localize_do("WND002", "за неделю:")
wnd003.Caption = localize_do("WND003", "за месяц:")
wnd004.Caption = localize_do("WND004", "за все время:")
Frame1.Caption = localize_do("WND005", "Начисления")

Label21.Caption = localize_do("WND006", "Скорость приема:")
Label22.Caption = localize_do("WND007", "Скорость передачи:")
Label12.Caption = localize_do("WND008", "Суточная скорость:")
Label8.Caption = localize_do("WND009", "Суточный расход:")
Label1.Caption = localize_do("WNDX01", "Ожидается к концу месяца:")

Label9.Caption = localize_do("WND010", "Ост. бонуса (Dl/Ul)")
Label10.Caption = localize_do("WND011", "Остаток кредита:")
Label2.Caption = localize_do("WND012", "Скорость подключения:")
Frame6.Caption = localize_do("WND013", "Сведения")

Label13.Caption = localize_do("WND014", "Ограничение:")
Label14.Caption = localize_do("WND015", "Остаток:")
Frame2.Caption = localize_do("WND016", "Ограничение трафика")

Label17.Caption = localize_do("WND017", "Отправлено байт:")
Label18.Caption = localize_do("WND018", "Получено байт:")
Label4.Caption = localize_do("WND019", "Сумма байт:")
Label19.Caption = localize_do("WND020", "Начислено:")
Label11.Caption = localize_do("WND021", "Макс. замеченная скорость:")

Picture2.ToolTipText = localize_do("WND026", "Свернуть в трей")

frmOK.Label1.Caption = localize_do("WNDA01", "За текущее соединение:")
frmOK.Label2.Caption = localize_do("WNDA02", "За день:")
frmOK.Label3.Caption = localize_do("WNDA03", "За месяц:")
frmOK.Label5.Caption = localize_do("WNDA05", "Скор. прием/передача")

superMenu.itmMnu(1).Caption = localize_do("MNU001", "Показать/скрыть окно программы...")
superMenu.itmMnu(2).Caption = localize_do("MNU002", "Report")
superMenu.itmMnu(3).Caption = localize_do("MNU003", "Setup")
superMenu.itmMnu(4).Caption = localize_do("MNU004", "Taxes")
superMenu.itmMnu(5).Caption = localize_do("MNU005", "Bonus")
superMenu.itmMnu(6).Caption = localize_do("MNU006", "About")
superMenu.itmMnu(7).Caption = localize_do("MNU011", "Запускать вместе с Windows")
superMenu.itmMnu(8).Caption = localize_do("MNU007", "Exit")
superMenu.itmMnu(9).Caption = localize_do("MNU012", "Проверить обновление...")
superMenu.itmMnu(10).Caption = localize_do("MNU013", "Посетить форум программы...")

' scConfig.ToolTipText =
' scTaxes.ToolTipText =
 'scBonus.ToolTipText = localize_do("SC03A")
imgPanel(2).ToolTipText = localize_do("SC02A")
imgPanel(4).ToolTipText = localize_do("SC01A")
imgPanel(6).ToolTipText = localize_do("SC05A")
imgPanel(8).ToolTipText = localize_do("SC04A")

End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Me.Visible = False
End Sub

Private Sub Form_Load()
  
Set m_objIpHelper = New CIpHelper


On Error Resume Next

' Terminate programm by command
If Command = "/term" Then
Dim N As Long, A As Long
Me.Caption = "/term process"
Do
N = FindWindow(vbNullString, "Woobind Network Meter hWnd")
If N > 0 Then A = SendMessage(N, WM_CLOSE, 0, 0)
Loop While Not N = 0
End
End If
' End terminate the programm

' Check filename
If App.EXEName <> "wnmeter" Then End
' End check filename

Me.BackColor = RGB(240, 0, 120)
Me.Height = 5820

Dim e

For e = 0 To 5
   MassSpeed(e, 0) = 0
   MassSpeed(e, 1) = 0
Next e
  

' Load Subs
Call LoadlRecords
Call LoadConfiguration
Call LoadRulers
Call RefreshFace
Call RefreshPing
Call UpdateAverage
current_traffic_tax = taxes_matrix(Weekday(Now, vbMonday) - 1, Hour(Now)): previous_traffic_tax = current_traffic_tax


NetworkText(0) = "Offline"
NetworkText(1) = "Link Down"
NetworkText(2) = "Online"

host_is_alive = True
iph_reset_traffic_flag = True
NetworkStatus(0) = Offline
ProgConnected = False
Me.Visible = False
MultiCounter = -200

TrayAdd picTray, "Woobind Network Meter"


WasSent = DataRcvedToday + DataXmitedToday
TimeLongTemp = GetTickCount
TimeLongBack = TimeLongTemp

If iph_get_traffic(StartRc, StartXm) Then
   ProgConnected = True
Else
   ProgConnected = False
End If

tmrEE_Timer ' update tray icon

' CheckNewVersion

If FRun = True Then Dialog.Show 0, Me

End Sub

Sub CheckLimits()

If TipLimit Or TipPreLimit Then
    If CurrLimit <> WasLimit Then
        CheckLimitsX
    End If
End If

WasLimit = CurrLimit

End Sub

Sub CheckLimitsX()


 Select Case CurrLimit
 Case LimitStatus.None
 Case LimitStatus.Overload
   If TipLimit Then ShowBalloon "Woobind Network Meter", 3, localize_do("MSG002", "Внимание! Трафик исчерпан!") + vbCrLf + _
   localize_do("MSG003", "Перерасход") + ": " + DataDevi32(-ActiveLeft)
 Case LimitStatus.Redline
   If TipPreLimit Then ShowBalloon "Woobind Network Meter", 1, LimitNameA(DataLimitMode) + " " + DataDevi32(ActiveLeft)
 End Select


End Sub


Sub ShowBalloon(inTitle As String, inType As Long, inMessage As String)
' TrayBalloon picTray, inTitle, inType, inMessage
    PoP.ShowPopup "Woobind Network Meter", inMessage, inType, TipSound, ShowTop

End Sub

Sub ShowBalloonParam(inTitle As String, inType As Long, inMessage As String, inS As String)

inMessage = Replace(inMessage, "%s", inS, , , vbTextCompare)
PoP.ShowPopup "Woobind Network Meter", inMessage, inType, TipSound, ShowTop

End Sub


Sub LoadConfiguration()

' // SORTED OPTIONS
' // FRAME 0
' >> INTERFACES
PingNetwork = GetSettingFake("Network Meter\" + IRP, "Ping Settings", "PingNetwork", "0")
PingMode = GetSettingFake("Network Meter\" + IRP, "Ping Settings", "PingMode", "0")
PingManual = GetSettingFake("Network Meter\" + IRP, "Ping Settings", "PingManual", "127.0.0.1")
iph_interface = iph_interface_decode(GetSettingFake("Network Meter\" + IRP, "Configuration", "ActiveInterface", ALL_OFF))

' >> OPTIONS
EveryDayCheck = GetSettingFake("Network Meter\" + IRP, "Configuration", "CheckUpdate", True)
MainWindowAttach = GetSettingFake("Network Meter\" + IRP, "Configuration", "MainWindowAttach", True)
MainWindowAlpha = GetSettingFake("Network Meter\" + IRP, "Configuration", "MainWindowAlpha", False)
MainWindowAlphaLVL = GetSettingFake("Network Meter\" + IRP, "Configuration", "MainWindowAlphaLVL", 15)

' >> LANGUAGE
LanguageName = GetSettingFake("Network Meter\" + IRP, "Configuration", "LanguageName", "Russian")
CacheStrings def_complete_path(App.Path) & LanguageName & ".slf"

' // FRAME 1
' >> WINDOW
FWAlwaysVisible = GetSettingFake("Network Meter\" + IRP, "Configuration", "FWAlwaysVisible", False)
FloatWindow = GetSettingFake("Network Meter\" + IRP, "Configuration", "FloatWindow", True)
FWOnTop = GetSettingFake("", "Window Position", "FWOnTop", True)
FWrmr = GetSettingFake("", "Window Position", "FWrmr", False)
FWOnline = GetSettingFake("", "Window Position", "FWOnline", False)
FWAlwaysVisible = GetSettingFake("", "Window Position", "FWAlwaysVisible", False)
ComboLinks = GetSettingFake("Network Meter\" + IRP, "Configuration", "Combo", "000")


' // FRAME 2
' >> LIMIT
LimUse = GetSettingFake("Network Meter\" + IRP, "Limits", "LimUse", False)
DataLimit = GetSettingFake("Network Meter\" + IRP, "Limits", "DataLimit", "0")
DataLimitDivide = GetSettingFake("Network Meter\" + IRP, "Limits", "DataLimitDivide", "0")
DataLimitMode = GetSettingFake("Network Meter\" + IRP, "Limits", "DataLimitMode", "0")
DataLimitWay = GetSettingFake("Network Meter\" + IRP, "Limits", "DataLimitWay", "0")

' >> NOTIFY
TipLimit = GetSettingFake("Network Meter\" + IRP, "Limits", "TipLimit", "0")
TipPreLimit = GetSettingFake("Network Meter\" + IRP, "Limits", "TipPreLimit", False)
TipSound = GetSettingFake("Network Meter\" + IRP, "Limits", "TipSound", True)
ShowTop = GetSettingFake("Network Meter\" + IRP, "Limits", "ShowTop", False)
LimitLine = GetSettingFake("Network Meter\" + IRP, "Limits", "LimitLine", "10")


' // FRAME 3
' >> STATIC
StaticTariff = CCur(GetSettingFake("Network Meter\" + IRP, "Taxes", "StaticTariff", "0"))



' Load credit settings
CreditLeft = CCur(GetSettingFake("Network Meter\" + IRP, "Credit Settings", "CreditsLeft", "0"))

' Load Bonus
DataBonusRcved = Val(GetSettingFake("Network Meter\" + IRP, "Bonus", "DataBonusRcved", "0"))
DataBonusXmited = Val(GetSettingFake("Network Meter\" + IRP, "Bonus", "DataBonusXmited", "0"))
DataBonusBoth = Val(GetSettingFake("Network Meter\" + IRP, "Bonus", "DataBonusBoth", "0"))
DataBonusEnabled = def_any_to_bool(GetSettingFake("Network Meter\" + IRP, "Bonus", "DataBonusEnabled", False))
DataBonusMode = Val(GetSettingFake("Network Meter\" + IRP, "Bonus", "DataBonusMode", "0"))


' Load counter values
DataXmitedToday = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataXmitedToday", "0"))
DataRcvedToday = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataRcvedToday", "0"))
DataXmitedTemp = DataXmitedToday
DataRcvedTemp = DataRcvedToday

DataXmitedHour = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataXmitedHour", "0"))
DataRcvedHour = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataRcvedHour", "0"))
DataXmitedWeek = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataXmitedWeek", "0"))
DataRcvedWeek = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataRcvedWeek", "0"))
DataXmitedMonth = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataXmitedMonth", "0"))
DataRcvedMonth = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataRcvedMonth", "0"))
DataXmitedYear = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataXmitedYear", "0"))
DataRcvedYear = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataRcvedYear", "0"))
DataXmitedAll = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataXmitedAll", "0"))
DataRcvedAll = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataRcvedAll", "0"))
DataXmitedCount = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataXmitedCount", "0"))
DataRcvedCount = Val(GetSettingFake("Network Meter\" + IRP, "Counter Values", "DataRcvedCount", "0"))
ResetDate = GetSettingFake("Network Meter\" + IRP, "Counter Values", "CountingFrom", Format(Now, "dd.mm.yyyy hh:mm"))

' Load limits



' Load taxes
TaxHour = CCur(GetSettingFake("Network Meter\" + IRP, "Taxes", "TaxHour", "0"))
TaxToday = CCur(GetSettingFake("Network Meter\" + IRP, "Taxes", "TaxToday", "0"))
TaxWeek = CCur(GetSettingFake("Network Meter\" + IRP, "Taxes", "TaxWeek", "0"))
TaxMonth = CCur(GetSettingFake("Network Meter\" + IRP, "Taxes", "TaxMonth", "0"))
TaxYear = CCur(GetSettingFake("Network Meter\" + IRP, "Taxes", "TaxYear", "0"))
TaxAll = CCur(GetSettingFake("Network Meter\" + IRP, "Taxes", "TaxAll", "0"))
TaxCount = CCur(GetSettingFake("Network Meter\" + IRP, "Taxes", "TaxCount", "0"))
tax_taxing_traffic = Val(GetSettingFake("Network Meter\" + IRP, "Taxes", "CredMode", "0"))
TaxName = GetSettingFake("Network Meter\" + IRP, "Taxes", "TaxName", "грн.")
notify_tax_change = def_any_to_bool(GetSettingFake("Network Meter\" + IRP, "Taxes", "TaxChanged", 0))

' Load configuration
timeActive = Val(GetSettingFake("Network Meter\" + IRP, "Configuration", "timeActive", "0"))
MaxxSpeed(1) = CCur(GetSettingFake("Network Meter\" + IRP, "Configuration", "MaximumDetectedUp", "0"))
MaxxSpeed(2) = CCur(GetSettingFake("Network Meter\" + IRP, "Configuration", "MaximumDetectedDown", "0"))
dayBegin = Val(GetSettingFake("Network Meter\" + IRP, "Configuration", "CountBegin", "1"))
lOption = Val(GetSettingFake("Network Meter\" + IRP, "Configuration", "LastOption", "0"))

' Load fw settings
FloatNotify = def_any_to_bool(GetSettingFake("Network Meter\" + IRP, "Configuration", "FloatNotify", True))
FRun = def_any_to_bool(GetSettingFake("Network Meter\" + IRP, "Configuration", "First Run", True))

Use1024 = def_any_to_bool(GetSettingFake("Network Meter\" + IRP, "Configuration", "Use1024", False))

' Load ping settings



' Load window position
Me.Left = 15 * Val(GetSettingFake("", "Window Position", "X", ((Screen.Width / 2) - (Me.Width / 2)) / 15))
Me.Top = 15 * Val(GetSettingFake("", "Window Position", "Y", ((Screen.Height / 2) - (Me.Height / 2)) / 15))

' Load checkpoint
CurrYear = Val(GetSettingFake("Network Meter\" + IRP, "Save Checkpoint", "CurrentYear", "0"))
CurrMonth = Val(GetSettingFake("Network Meter\" + IRP, "Save Checkpoint", "CurrentMonth", "0"))
CurrWeek = Val(GetSettingFake("Network Meter\" + IRP, "Save Checkpoint", "CurrentWeek", "0"))
CurrDay = Val(GetSettingFake("Network Meter\" + IRP, "Save Checkpoint", "CurrentDay", "0"))
CurrentHour = GetSettingFake("Network Meter\" + IRP, "Save Checkpoint", "CurrentHour", Now)
stConnection = Val(GetSettingFake("Network Meter\" + IRP, "Save Checkpoint", "stConnection", "0"))

' Load Autostart
UseAutostart = def_any_to_bool(GetSettingFake("Network Meter\" + IRP, "Autostart", "UseAutostart", False))
UseAutostop = def_any_to_bool(GetSettingFake("Network Meter\" + IRP, "Autostart", "UseAutostop", False))
UseLinkDown = def_any_to_bool(GetSettingFake("Network Meter\" + IRP, "Autostart", "UseLinkDown", False))
UseAutoNotify = def_any_to_bool(GetSettingFake("Network Meter\" + IRP, "Autostart", "UseAutoNotify", False))


' Load Abonetic
With Abonetic
    .aI = Val(GetSettingFake("Network Meter\" + IRP, "Abonetic", "aI", "0"))
    .aO = Val(GetSettingFake("Network Meter\" + IRP, "Abonetic", "aO", "0"))
    .aIO = Val(GetSettingFake("Network Meter\" + IRP, "Abonetic", "aIO", "0"))
    .aMoney = Val(GetSettingFake("Network Meter\" + IRP, "Abonetic", "aMoney", "0"))
    .aPeriod = Val(GetSettingFake("Network Meter\" + IRP, "Abonetic", "aPeriod", "0"))
    .aLastSetted = Val(GetSettingFake("Network Meter\" + IRP, "Abonetic", "aLastSetted", "0"))
End With


End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
mouse_old_x = x
mouse_old_y = y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim ML, MT

If Button = 1 Then

 Dim desktop As RECT: desktop = GetScreen
 
 ML = Me.Left + (x - mouse_old_x)
 MT = Me.Top + (y - mouse_old_y)
 
 If MainWindowAttach Then
    If ML <= (desktop.Left * 15) + 120 Then ML = desktop.Left * 15
    If MT <= (desktop.Top * 15) + 120 Then MT = desktop.Top * 15
 
    If ML > (desktop.Right * 15 - Me.Width - 120) Then ML = (desktop.Right * 15) - Me.Width
    If MT > (desktop.Bottom * 15 - Me.Height - 120) Then MT = (desktop.Bottom * 15) - Me.Height
 End If

 Me.Left = ML
 Me.Top = MT

End If


End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then SaveFormPosition
If Button = 2 Then superMenu.PositeMe
End Sub


Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Me.Visible = False: Me.WindowState = vbNormal
End Sub


Private Sub Form_Terminate()
Form_Unload 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  
  SaveToINI
  SaveToINIcounts
  TrayRemove
  
  Fin

End Sub




Private Sub imgPanel_Click(index As Integer)
DoEvents
Select Case index
Case 0: mnuAbout_Click
Case 2: mStatistic_Click
Case 4: scReset
Case 6: mnuSetup_Click
Case 8: mnuTaxes_Click
End Select
End Sub

Private Sub imgPanel_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
imgPanel(index).Picture = ipImages.ListImages(index + 2).Picture

End Sub

Private Sub imgPanel_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case index
Case 0
    If imgPanel(0).Picture = ipImages.ListImages(1).Picture Then imgPanel(0).Picture = ipImages.ListImages(11).Picture
Case 4
    If imgPanel(4).Picture = ipImages.ListImages(5).Picture Then imgPanel(4).Picture = ipImages.ListImages(12).Picture
End Select
End Sub

Private Sub imgPanel_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
imgPanel(index).Picture = ipImages.ListImages(index + 1).Picture
End Sub

Private Sub lAbout_Click()
mnuAbout_Click
End Sub

Private Sub lblReset_Click()
scReset
End Sub

Private Sub lMe_DblClick()
Me.Visible = False
End Sub

Private Sub lMe_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_MouseDown Button, Shift, x, y
End Sub

Private Sub lMe_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_MouseMove Button, Shift, x, y

End Sub

Private Sub lMe_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_MouseUp Button, Shift, x, y
End Sub

Sub mExit_Click()
Unload Me
End Sub

Sub mReall_Click()
On Error Resume Next
Dim Q As Integer

DataXmitedYear = 0
DataXmitedToday = 0
DataXmitedMonth = 0
DataXmitedWeek = 0
DataXmitedAll = 0
DataXmitedCount = 0

TaxCount = 0
TaxToday = 0
TaxMonth = 0
TaxWeek = 0
TaxAll = 0

DataRcvedYear = 0
DataRcvedToday = 0
DataRcvedMonth = 0
DataRcvedWeek = 0
DataRcvedAll = 0
DataRcvedCount = 0

MaxxSpeed(1) = 0
MaxxSpeed(2) = 0
timeActive = 0
ResetDate = Format(Now, "dd.mm.yyyy hh:mm")

SaveToINI

End Sub

Sub mnuAbout_Click()
    OpenOptions 6
End Sub

Sub mnuAutorun_Click()

If NoExit Then Exit Sub

superMenu.Check1.Value = 1 - superMenu.Check1.Value

If superMenu.Check1.Value > 0 Then RegisterAutorun
If superMenu.Check1.Value = 0 Then UnRegisterAutorun

superMenu.Check1.Value = def_bool_to_int(CheckAutorun)

End Sub

Sub mnuBonus_Click()
    OpenOptions 4
End Sub

Sub mnuSetup_Click()
    OpenOptions 0
End Sub

Sub mnuShow_Click()
frmVelton.Visible = Not frmVelton.Visible
If frmVelton.Visible = True Then frmVelton.SetFocus

End Sub

Sub mnuTaxes_Click()
    OpenOptions 3
End Sub

Sub mStatistic_Click()
everyday.Show 0, Me
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim lMsg As Single, hFG As Long
lMsg = x

hFG = GetForegroundWindow


Select Case lMsg
    Case WM_LBUTTONUP
    Case WM_RBUTTONUP
    Case WM_BALLOONL
        Me.Visible = True
        Me.SetFocus
    Case WM_MOUSEMOVE
        If ProgConnected And Not superMenu.Visible Then frmTrayTip.ShowTip "Hello", 0, 0
    Case WM_LBUTTONDOWN
        Me.Visible = Not IsWindowOnScreen(Me.hWnd)
        If Me.Visible Then Me.SetFocus
    Case WM_RBUTTONDOWN
      superMenu.PositeMe
    Case Else
End Select
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then Picture2.Picture = clsD(2).Picture

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Picture2.Picture = clsD(0).Picture Then Picture2.Picture = clsD(1).Picture

End Sub


Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then Picture2.Picture = clsD(1).Picture
If Button = 1 Then Me.Visible = False


End Sub


Private Sub Timer1_Timer()

RefreshPing

If DataRcved > Zond1 Or DataXmited > Zond2 Then

SaveToINIcounts

End If

End Sub

Sub RefreshPing()

Dim PingString As String


 

If PingNetwork = True Then
    
    
    If PingMode = 0 Then
        PingString = "www.google.com"
    ElseIf PingMode = 1 Then
        PingString = PingManual
    End If
    
    If ProgConnected = False Then
        host_is_alive = True
        imgLevel.Picture = imgNetLevel.ListImages(1).Picture
        Exit Sub
    End If
    
    host_is_alive = Ping(PingString, PingLong)
    If Not imgLevel.Visible Then imgLevel.Visible = True
    
    If host_is_alive = False Then
      imgLevel.Picture = imgNetLevel.ListImages(1).Picture
    ElseIf PingLong > 1000 Then
      imgLevel.Picture = imgNetLevel.ListImages(2).Picture
    Else
      imgLevel.Picture = imgNetLevel.ListImages(14 - Fix(12 / Sqr(1000) * Sqr(PingLong))).Picture
    End If
    imgLevel.ToolTipText = "Ping: " + Format(PingLong, "0") + " мс."
Else
    host_is_alive = True
    If imgLevel.Visible Then imgLevel.Visible = False
End If

OldTempPing = host_is_alive

End Sub


Sub SaveToINI()

On Error Resume Next


Dim IRP As String
IRP = def_complete_path(App.Path) + "app_data.ini"

' // SORTED OPTIONS
' // FRAME 0
' >> INTERFACES
WritePrivateProfileString "Ping Settings", "PingNetwork", def_bool_to_str(PingNetwork), IRP
WritePrivateProfileString "Ping Settings", "PingMode", CStr(PingMode), IRP
WritePrivateProfileString "Ping Settings", "PingManual", CStr(PingManual), IRP
WritePrivateProfileString "Configuration", "ActiveInterface", iph_interface_encode(iph_interface), IRP

' >> OPTIONS
WritePrivateProfileString "Configuration", "CheckUpdate", def_bool_to_str(EveryDayCheck), IRP
WritePrivateProfileString "Configuration", "MainWindowAttach", def_bool_to_str(MainWindowAttach), IRP
WritePrivateProfileString "Configuration", "MainWindowAlpha", def_bool_to_str(MainWindowAlpha), IRP
WritePrivateProfileString "Configuration", "MainWindowAlphaLVL", CStr(MainWindowAlphaLVL), IRP

WritePrivateProfileString "Configuration", "Use1024", def_bool_to_str(Use1024), IRP

' >> LANGUAGE
WritePrivateProfileString "Configuration", "LanguageName", LanguageName, IRP

' // LIMITER
' >> LIMIT
WritePrivateProfileString "Limits", "LimUse", def_bool_to_str(LimUse), IRP
WritePrivateProfileString "Limits", "DataLimit", CStr(DataLimit), IRP
WritePrivateProfileString "Limits", "DataLimitDivide", CStr(DataLimitDivide), IRP
WritePrivateProfileString "Limits", "DataLimitMode", CStr(DataLimitMode), IRP
WritePrivateProfileString "Limits", "DataLimitWay", CStr(DataLimitWay), IRP

' >> NOTIFY
WritePrivateProfileString "Limits", "TipLimit", def_bool_to_str(TipLimit), IRP
WritePrivateProfileString "Limits", "LimitLine", CStr(LimitLine), IRP
WritePrivateProfileString "Limits", "TipPreLimit", def_bool_to_str(TipPreLimit), IRP
WritePrivateProfileString "Limits", "TipSound", def_bool_to_str(TipSound), IRP
WritePrivateProfileString "Limits", "ShowTop", def_bool_to_str(ShowTop), IRP


' // TARIFF
' >> STATIC
WritePrivateProfileString "Taxes", "StaticTariff", CStr(StaticTariff), IRP


'////
WritePrivateProfileString "Credit Settings", "CreditsLeft", CStr(CreditLeft), IRP

WritePrivateProfileString "Bonus", "DataBonusRcved", CStr(DataBonusRcved), IRP
WritePrivateProfileString "Bonus", "DataBonusXmited", CStr(DataBonusXmited), IRP
WritePrivateProfileString "Bonus", "DataBonusBoth", CStr(DataBonusBoth), IRP
WritePrivateProfileString "Bonus", "DataBonusEnabled", def_bool_to_str(DataBonusEnabled), IRP
WritePrivateProfileString "Bonus", "DataBonusMode", CStr(DataBonusMode), IRP

WritePrivateProfileString "Counter Values", "DataXmitedToday", CStr(DataXmitedToday), IRP
WritePrivateProfileString "Counter Values", "DataRcvedToday", CStr(DataRcvedToday), IRP
WritePrivateProfileString "Counter Values", "DataXmitedWeek", CStr(DataXmitedWeek), IRP
WritePrivateProfileString "Counter Values", "DataRcvedWeek", CStr(DataRcvedWeek), IRP
WritePrivateProfileString "Counter Values", "DataXmitedMonth", CStr(DataXmitedMonth), IRP
WritePrivateProfileString "Counter Values", "DataRcvedMonth", CStr(DataRcvedMonth), IRP
WritePrivateProfileString "Counter Values", "DataXmitedYear", CStr(DataXmitedYear), IRP
WritePrivateProfileString "Counter Values", "DataRcvedYear", CStr(DataRcvedYear), IRP
WritePrivateProfileString "Counter Values", "DataXmitedAll", CStr(DataXmitedAll), IRP
WritePrivateProfileString "Counter Values", "DataRcvedAll", CStr(DataRcvedAll), IRP
WritePrivateProfileString "Counter Values", "DataXmitedCount", CStr(DataXmitedCount), IRP
WritePrivateProfileString "Counter Values", "DataRcvedCount", CStr(DataRcvedCount), IRP
WritePrivateProfileString "Counter Values", "CountingFrom", CStr(ResetDate), IRP

WritePrivateProfileString "Taxes", "TaxToday", CStr(TaxToday), IRP
WritePrivateProfileString "Taxes", "TaxWeek", CStr(TaxWeek), IRP
WritePrivateProfileString "Taxes", "TaxMonth", CStr(TaxMonth), IRP
WritePrivateProfileString "Taxes", "TaxYear", CStr(TaxYear), IRP
WritePrivateProfileString "Taxes", "TaxAll", CStr(TaxAll), IRP
WritePrivateProfileString "Taxes", "TaxCount", CStr(TaxCount), IRP
WritePrivateProfileString "Taxes", "CredMode", CStr(tax_taxing_traffic), IRP
WritePrivateProfileString "Taxes", "TaxName", TaxName, IRP
WritePrivateProfileString "Taxes", "TaxChanged", def_bool_to_str(notify_tax_change), IRP


WritePrivateProfileString "Configuration", "timeActive", CStr(timeActive), IRP
WritePrivateProfileString "Configuration", "MaximumDetectedUp", CStr(MaxxSpeed(1)), IRP
WritePrivateProfileString "Configuration", "MaximumDetectedDown", CStr(MaxxSpeed(2)), IRP
WritePrivateProfileString "Configuration", "CountBegin", CStr(dayBegin), CStr(IRP)
WritePrivateProfileString "Configuration", "FloatWindow", def_bool_to_str(FloatWindow), IRP
WritePrivateProfileString "Configuration", "FloatNotify", def_bool_to_str(FloatNotify), IRP
WritePrivateProfileString "Configuration", "First Run", "False", IRP
WritePrivateProfileString "Configuration", "LastOption", CStr(lOption), IRP
WritePrivateProfileString "Configuration", "Combo", ComboLinks, IRP



WritePrivateProfileString "Window Position", "Left", CStr(Fix(frmOK.Left / 15)), IRP
WritePrivateProfileString "Window Position", "FWOnTop", def_bool_to_str(FWOnTop), IRP
WritePrivateProfileString "Window Position", "FWrmr", def_bool_to_str(FWrmr), IRP
WritePrivateProfileString "Window Position", "FWOnline", def_bool_to_str(FWOnline), IRP
WritePrivateProfileString "Window Position", "FWAlwaysVisible", def_bool_to_str(FWAlwaysVisible), IRP


WritePrivateProfileString "Save Checkpoint", "CurrentYear", CStr(CurrYear), IRP
WritePrivateProfileString "Save Checkpoint", "CurrentMonth", CStr(CurrMonth), IRP
WritePrivateProfileString "Save Checkpoint", "CurrentWeek", CStr(CurrWeek), IRP
WritePrivateProfileString "Save Checkpoint", "CurrentDay", CStr(CurrDay), IRP
WritePrivateProfileString "Save Checkpoint", "stConnection", CStr(stConnection), IRP

WritePrivateProfileString "Autostart", "UseAutoNotify", def_bool_to_str(UseAutoNotify), IRP
WritePrivateProfileString "Autostart", "UseAutostart", def_bool_to_str(UseAutostart), IRP
WritePrivateProfileString "Autostart", "UseAutostop", def_bool_to_str(UseAutostop), IRP
WritePrivateProfileString "Autostart", "UseLinkdown", def_bool_to_str(UseLinkDown), IRP


' Save Abonetic
With Abonetic
    WritePrivateProfileString "Abonetic", "aI", CStr(.aI), IRP
    WritePrivateProfileString "Abonetic", "aO", CStr(.aO), IRP
    WritePrivateProfileString "Abonetic", "aIO", CStr(.aIO), IRP
    WritePrivateProfileString "Abonetic", "aMoney", CStr(.aMoney), IRP
    WritePrivateProfileString "Abonetic", "aPeriod", CStr(.aPeriod), IRP
    WritePrivateProfileString "Abonetic", "aLastSetted", CStr(.aLastSetted), IRP
End With



End Sub

Sub SaveToINIcounts()

On Error Resume Next


Dim IRP As String
IRP = def_complete_path(App.Path) + "app_data.ini"

'////
WritePrivateProfileString "Credit Settings", "CreditsLeft", CStr(CreditLeft), IRP

WritePrivateProfileString "Bonus", "DataBonusRcved", CStr(DataBonusRcved), IRP
WritePrivateProfileString "Bonus", "DataBonusXmited", CStr(DataBonusXmited), IRP
WritePrivateProfileString "Bonus", "DataBonusBoth", CStr(DataBonusBoth), IRP
WritePrivateProfileString "Bonus", "DataBonusEnabled", def_bool_to_str(DataBonusEnabled), IRP
WritePrivateProfileString "Bonus", "DataBonusMode", CStr(DataBonusMode), IRP

WritePrivateProfileString "Counter Values", "DataXmitedToday", CStr(DataXmitedToday), IRP
WritePrivateProfileString "Counter Values", "DataRcvedToday", CStr(DataRcvedToday), IRP
WritePrivateProfileString "Counter Values", "DataXmitedWeek", CStr(DataXmitedWeek), IRP
WritePrivateProfileString "Counter Values", "DataRcvedWeek", CStr(DataRcvedWeek), IRP
WritePrivateProfileString "Counter Values", "DataXmitedMonth", CStr(DataXmitedMonth), IRP
WritePrivateProfileString "Counter Values", "DataRcvedMonth", CStr(DataRcvedMonth), IRP
WritePrivateProfileString "Counter Values", "DataXmitedYear", CStr(DataXmitedYear), IRP
WritePrivateProfileString "Counter Values", "DataRcvedYear", CStr(DataRcvedYear), IRP
WritePrivateProfileString "Counter Values", "DataXmitedAll", CStr(DataXmitedAll), IRP
WritePrivateProfileString "Counter Values", "DataRcvedAll", CStr(DataRcvedAll), IRP
WritePrivateProfileString "Counter Values", "DataXmitedCount", CStr(DataXmitedCount), IRP
WritePrivateProfileString "Counter Values", "DataRcvedCount", CStr(DataRcvedCount), IRP
WritePrivateProfileString "Counter Values", "DataXmitedHour", CStr(DataXmitedHour), IRP
WritePrivateProfileString "Counter Values", "DataRcvedHour", CStr(DataRcvedHour), IRP
WritePrivateProfileString "Counter Values", "CountingFrom", CStr(ResetDate), IRP

WritePrivateProfileString "Taxes", "TaxHour", CStr(TaxHour), IRP
WritePrivateProfileString "Taxes", "TaxToday", CStr(TaxToday), IRP
WritePrivateProfileString "Taxes", "TaxWeek", CStr(TaxWeek), IRP
WritePrivateProfileString "Taxes", "TaxMonth", CStr(TaxMonth), IRP
WritePrivateProfileString "Taxes", "TaxYear", CStr(TaxYear), IRP
WritePrivateProfileString "Taxes", "TaxAll", CStr(TaxAll), IRP
WritePrivateProfileString "Taxes", "TaxCount", CStr(TaxCount), IRP
WritePrivateProfileString "Taxes", "CredMode", CStr(tax_taxing_traffic), IRP
WritePrivateProfileString "Taxes", "TaxChanged", def_bool_to_str(notify_tax_change), IRP


WritePrivateProfileString "Configuration", "timeActive", CStr(timeActive), IRP
WritePrivateProfileString "Configuration", "MaximumDetectedUp", CStr(MaxxSpeed(1)), IRP
WritePrivateProfileString "Configuration", "MaximumDetectedDown", CStr(MaxxSpeed(2)), IRP
WritePrivateProfileString "Configuration", "CountBegin", CStr(dayBegin), CStr(IRP)

WritePrivateProfileString "Save Checkpoint", "CurrentYear", CStr(CurrYear), IRP
WritePrivateProfileString "Save Checkpoint", "CurrentMonth", CStr(CurrMonth), IRP
WritePrivateProfileString "Save Checkpoint", "CurrentWeek", CStr(CurrWeek), IRP
WritePrivateProfileString "Save Checkpoint", "CurrentDay", CStr(CurrDay), IRP
WritePrivateProfileString "Save Checkpoint", "CurrentHour", CStr(CurrentHour), IRP
WritePrivateProfileString "Save Checkpoint", "stConnection", CStr(stConnection), IRP


End Sub



Private Sub scReset()
On Error Resume Next

  DataXmitedCount = 0
  DataRcvedCount = 0
  TaxCount = 0
  timeActive = 0
  MaxxSpeed(1) = 0
  MaxxSpeed(2) = 0
  
  ResetDate = Format(Now, "dd.mm.yyyy hh:mm")
  Frame7.Caption = localize_do("MSG001", "Начиная с") + " " + ResetDate 'LOL FORM1MSG1
  
  tmrAction_Timer
End Sub

Private Sub Timer2_Timer()


' Псевдотаймер  0...60 сек
MultiCounter = MultiCounter + 200
MultiCounter = MultiCounter Mod 60000

' Последовательный запуск таймеров
If MultiCounter Mod 200 = 0 Then frmOK.tmrhyd_Timer: DoEvents
If MultiCounter Mod 1000 = 0 Then tmrEE_Timer: DoEvents
If MultiCounter = 0 Then Timer1_Timer: DoEvents
If MultiCounter Mod 10000 = 0 Then tmrAverage_Timer: DoEvents

' Обновленте изображения кнопки "свернуть" в стиле Vista
If IsMouseInWindow(Picture2.hWnd) = False And Not Picture2.Picture = clsD(2).Picture Then Picture2.Picture = clsD(0).Picture

If IsMouseInWindow(imgPanel(0).hWnd) = False And Not imgPanel(0).Picture = ipImages.ListImages(2).Picture Then imgPanel(0).Picture = ipImages.ListImages(1).Picture
If IsMouseInWindow(imgPanel(4).hWnd) = False And Not imgPanel(4).Picture = ipImages.ListImages(6).Picture Then imgPanel(4).Picture = ipImages.ListImages(5).Picture


End Sub


Sub tmrAction_Timer()
On Error Resume Next

Dim iph_traffic_result As Boolean
Dim TimeLong As Currency 'todo
Dim TimeLongTemp As Long 'todo
Dim WeekVariable



current_traffic_tax = taxes_matrix(Weekday(Now, vbMonday) - 1, Hour(Now))

credit_blinker_status = Not credit_blinker_status

superMenu.itmMnu(0).Caption = filter_interface_name(localize_do(iph_interface, iph_interface))
lMe.Caption = superMenu.itmMnu(0).Caption + " - Woobind Network Meter"

iph_received_result = 0: iph_transmitted_result = 0

iph_traffic_result = iph_get_traffic(iph_received_result, iph_transmitted_result)   ' very fat function :(

If iph_traffic_result = False Then iph_reset_traffic_flag = True

If iph_reset_traffic_flag Then
  iph_transmitted_previous = iph_transmitted_result
  iph_received_previous = iph_received_result
  iph_reset_traffic_flag = False
End If

iph_received_delta = iph_received_result - iph_received_previous
iph_transmitted_delta = iph_transmitted_result - iph_transmitted_previous
  
kill_sign iph_received_delta
kill_sign iph_transmitted_delta
  
bonus_received_delta = iph_received_delta
bonus_transmitted_delta = iph_transmitted_delta
  

If iph_traffic_result Then
  If Not current_traffic_tax = previous_traffic_tax And notify_tax_change Then ShowBalloonParam "Woobind Network Meter", 2, localize_do("BALLOON01", "Стоимость трафика была изменена на %s за мегабайт."), FormatEx(current_traffic_tax, "0.#####") + " " + TaxName
  previous_traffic_tax = current_traffic_tax
  ProgConnected = True
  If iph_received_delta > 0 Or iph_received_delta > 0 Then timeActive = timeActive + TimeLong
Else
  ProgConnected = False
  Zond1 = 0: Zond2 = 0
End If


' Network status update

If iph_traffic_result = True And host_is_alive = True Then
   Let NetworkStatus(0) = 2
ElseIf iph_traffic_result = True And host_is_alive = False Then
   Let NetworkStatus(0) = 1
ElseIf iph_traffic_result = False Then
   Let NetworkStatus(0) = 0
End If

' laucher delay appendix
If delay_enabled Then
  If Not delay_counter Then
    delay_enabled = False
    LaunchlRecords
  Else
    delay_counter = delay_counter - 1
  End If
End If

If Not NetworkStatus(0) = NetworkStatus(1) Then
    tmrAverage_Timer
    If FloatNotify = True Then
      If NetworkStatus(0) = Linkdown Then ShowBalloon "Woobind Network Meter", 1, localize_do("MSG005", "Сеть недоступна")
      If NetworkStatus(0) = Offline Then ShowBalloon "Woobind Network Meter", 1, localize_do("MSG006", "Соединение разорвано")
      If NetworkStatus(0) = Online Then ShowBalloon "Woobind Network Meter", 1, localize_do("MSG007", "Соединение установлено")
    End If
    
    If UseAutostart Then
        If UseLinkDown Then
            Select Case NetworkStatus(0)
            Case NS.Online
                If NetworkStatus(1) = Offline Or NetworkStatus(1) = Linkdown Then Call delay_launch_set_timeout(5)
            Case NS.Offline, NS.Linkdown
                If NetworkStatus(1) = Online Then TerminatelRecords
            End Select
        Else
            Select Case NetworkStatus(0)
            Case NS.Online, NS.Linkdown
                If NetworkStatus(1) = Offline Then Call delay_launch_set_timeout(5)
            Case NS.Offline
                If NetworkStatus(1) = Online Or NetworkStatus(1) = Linkdown Then TerminatelRecords
            End Select
        End If
    End If
    
    
    If NetworkStatus(0) = Online And EveryDayCheck Then CheckUpdate False
    If Not NetworkStatus(0) = Offline Then iph_reset_traffic_flag = True: CheckLimitsX
    If NetworkStatus(0) = Offline Then iph_reset_traffic_flag = True

End If

NetworkStatus(1) = NetworkStatus(0)

Select Case DataLimitDivide
Case 0:
    ActiveLimit = (DataLimit * IIf(Use1024, 1024 ^ 3, 1000 ^ 3))
Case 1:
    ActiveLimit = (DataLimit * IIf(Use1024, 1024 ^ 2, 1000 ^ 2))
Case 2:
    ActiveLimit = (DataLimit * IIf(Use1024, 1024, 1000))
Case Else
    ActiveLimit = (DataLimit)
End Select

If frmOK.Visible = Not FloatWindow Then frmOK.Visible = FloatWindow



If NetworkStatus(0) = Online Then
     frmOK.imStat.Picture = frmOK.picON.Picture
     picA.Picture = frmOK.picON.Picture
ElseIf NetworkStatus(0) = Linkdown Then
     frmOK.imStat.Picture = frmOK.picDOWN.Picture
     picA.Picture = frmOK.picDOWN.Picture
ElseIf NetworkStatus(0) = Offline Then
     frmOK.imStat.Picture = frmOK.picOFF.Picture
     picA.Picture = frmOK.picOFF.Picture
End If


If iph_traffic_result = True Then ' ------------------------------------------------------------

  
  lblcml.Caption = DataDeviSpeed32(CCur(ConnSpd)) ' LOL <----------------------------
  
  ' ////////////////////////// '
  ' // BONUS SECTION BEGIN  // '
  ' ////////////////////////// '
  If DataBonusEnabled = True Then
    Select Case DataBonusMode
    Case 0
     
     If DataBonusBoth > bonus_received_delta + bonus_transmitted_delta Then
       DataBonusBoth = DataBonusBoth - (bonus_received_delta + bonus_transmitted_delta)
       bonus_received_delta = 0: bonus_transmitted_delta = 0
     Else
       DataBonusBoth = DataBonusBoth - bonus_received_delta
       If DataBonusBoth < 0 Then
         bonus_received_delta = -DataBonusBoth: DataBonusBoth = 0
       Else
         bonus_received_delta = 0
       End If
       DataBonusBoth = DataBonusBoth - bonus_transmitted_delta
       If DataBonusBoth < 0 Then
         bonus_transmitted_delta = -DataBonusBoth: DataBonusBoth = 0
       Else
         bonus_transmitted_delta = 0
       End If
     End If
     
    Case 1
    
    If DataBonusRcved > bonus_received_delta Then
        DataBonusRcved = DataBonusRcved - bonus_received_delta: bonus_received_delta = 0
    Else
        DataBonusRcved = DataBonusRcved - bonus_received_delta: bonus_received_delta = -DataBonusRcved: DataBonusRcved = 0
    End If
    
    If DataBonusXmited > bonus_transmitted_delta Then
        DataBonusXmited = DataBonusXmited - bonus_transmitted_delta: bonus_transmitted_delta = 0
    Else
        DataBonusXmited = DataBonusXmited - bonus_transmitted_delta: bonus_transmitted_delta = -DataBonusXmited: DataBonusXmited = 0
    End If
    
    End Select
    
    Select Case Abonetic.aPeriod
    Case 0
    Case 1
        If Day(Now) <> Abonetic.aLastSetted Then
            DataBonusBoth = Abonetic.aIO
            DataBonusXmited = Abonetic.aO
            DataBonusRcved = Abonetic.aI
            Dialog.MinusMoney Abonetic.aMoney
            Abonetic.aLastSetted = Day(Now)
            SaveToINI
        End If
    Case 2
        If Month(Now) <> Abonetic.aLastSetted Then
            DataBonusBoth = Abonetic.aIO
            DataBonusXmited = Abonetic.aO
            DataBonusRcved = Abonetic.aI
            Dialog.MinusMoney Abonetic.aMoney
            Abonetic.aLastSetted = Month(Now)
            SaveToINI
        End If
    End Select
    
  End If
  ' ////////////////////////// '
  
  iph_received_previous = iph_received_result
  iph_transmitted_previous = iph_transmitted_result
  
Else ' -----------------------------------------------------------------------------

  lblcml.Caption = "-"
  ' stConnection = stConnection + 1
  DataRcved = 0
  DataXmited = 0
  TaxTax = 0

End If '----------------------------------------------------------------------------

  
RememberMe = iph_traffic_result
  
  If Day(Now) <> CurrDay Then
    DataXmitedToday = 0
    DataRcvedToday = 0
    TaxToday = 0
    CurrDay = Day(Now)
  End If
  
  If Hour(Now) <> Hour(CurrentHour) And Not NetworkStatus(0) = Offline Then
    SaveCurrentHourStatistic DataXmitedHour, DataRcvedHour, TaxHour, CurrentHour
    CurrentHour = Now
    DataXmitedHour = 0
    DataRcvedHour = 0
    TaxHour = 0
    SaveNewHourStatistic DataXmitedHour, DataRcvedHour, TaxHour, CurrentHour
  End If
    
  
  WeekVariable = GetWeek(Day(Now), Month(Now), Year(Now))
  If WeekVariable <> CurrWeek Then
    DataXmitedWeek = 0
    DataRcvedWeek = 0
    TaxWeek = 0
    CurrWeek = WeekVariable
  End If
  
  If Month(Now) <> CurrMonth Then
    dayBegin = Day(Now)
    DataXmitedMonth = 0
    DataRcvedMonth = 0
    TaxMonth = 0
    CurrMonth = Month(Now)
  End If
  
  
  If Year(Now) <> CurrYear Then
    DataXmitedYear = 0
    DataRcvedYear = 0
    TaxYear = 0
    CurrYear = Year(Now)
  End If
  
  
  
  Summ DataRcvedHour, iph_received_delta
  Summ DataXmitedHour, iph_transmitted_delta
  
  Summ DataRcvedAll, iph_received_delta
  Summ DataXmitedAll, iph_transmitted_delta
  
  Summ DataRcvedYear, iph_received_delta
  Summ DataXmitedYear, iph_transmitted_delta
  
  Summ DataRcvedCount, iph_received_delta
  Summ DataXmitedCount, iph_transmitted_delta
  
  Summ DataRcvedToday, iph_received_delta
  Summ DataXmitedToday, iph_transmitted_delta
  
  Summ DataRcved, iph_received_delta
  Summ DataXmited, iph_transmitted_delta
  
  Summ DataRcvedMonth, iph_received_delta
  Summ DataXmitedMonth, iph_transmitted_delta
  
  If ActiveLimit > 0 And LimUse = True Then
   lblLimit.Caption = DataDevi(ActiveLimit) + LimitName(DataLimitMode)
   
    Select Case DataLimitWay
    Case 0
       Select Case DataLimitMode
       Case 0
         ActiveLeft = ActiveLimit - (DataRcvedToday + DataXmitedToday)
       Case 1
         ActiveLeft = (ActiveLimit - (DataRcvedWeek + DataXmitedWeek))
       Case 2
         ActiveLeft = (ActiveLimit - (DataRcvedMonth + DataXmitedMonth))
       Case 3
         ActiveLeft = (ActiveLimit - (DataRcvedCount + DataXmitedCount))
       End Select
    Case 1
       Select Case DataLimitMode
       Case 0
         ActiveLeft = ActiveLimit - (DataRcvedToday)
       Case 1
         ActiveLeft = (ActiveLimit - (DataRcvedWeek))
       Case 2
         ActiveLeft = (ActiveLimit - (DataRcvedMonth))
       Case 3
         ActiveLeft = (ActiveLimit - (DataRcvedCount))
       End Select
    Case 2
       Select Case DataLimitMode
       Case 0
         ActiveLeft = ActiveLimit - (DataXmitedToday)
       Case 1
         ActiveLeft = (ActiveLimit - (DataXmitedWeek))
       Case 2
         ActiveLeft = (ActiveLimit - (DataXmitedMonth))
       Case 3
         ActiveLeft = (ActiveLimit - (DataXmitedCount))
       End Select
    End Select
  
       If ActiveLeft < 0 Then
           CurrLimit = LimitStatus.Overload
           lblLeft.ForeColor = RGB(255, 0, 0): lblLeft.Visible = credit_blinker_status
       ElseIf ActiveLeft < ActiveLimit / (100 / LimitLine) And TipPreLimit Then
           CurrLimit = LimitStatus.Redline
           lblLeft.ForeColor = RGB(200, 0, 0): lblLeft.Visible = credit_blinker_status
       Else
           CurrLimit = LimitStatus.OK
           lblLeft.ForeColor = RGB(0, 200, 0): lblLeft.Visible = True
       End If
  Else
  
   CurrLimit = 0 ' No limits
   lblLimit.Caption = DataDevi(0): lblLeft.Visible = True: lblLeft.ForeColor = vbBlack
   ActiveLeft = 0
   
  End If


  CheckLimits
  
  
lblLeft.Caption = DataDevi(ActiveLeft)
lblLeft.ToolTipText = DataDeviTip(ActiveLeft)

Select Case tax_taxing_traffic
Case 0:
  tax_current_delta = (current_traffic_tax / 1000000 * (bonus_received_delta + bonus_transmitted_delta))
Case 1:
  tax_current_delta = (current_traffic_tax / 1000000 * bonus_received_delta)
Case 2:
  tax_current_delta = (current_traffic_tax / 1000000 * bonus_transmitted_delta)
Case 3:
 If DataXmitedMonth >= DataRcvedMonth Then
  tax_current_delta = current_traffic_tax / 1000000 * (bonus_transmitted_delta)
 Else
  tax_current_delta = current_traffic_tax / 1000000 * (bonus_received_delta)
 End If
End Select

Summ CreditLeft, -tax_current_delta
Summ TaxHour, tax_current_delta
Summ TaxCount, tax_current_delta
Summ TaxToday, tax_current_delta
Summ TaxWeek, tax_current_delta
Summ TaxMonth, tax_current_delta
Summ TaxAll, tax_current_delta
Summ TaxYear, tax_current_delta
Summ TaxTax, tax_current_delta
 

If DataBonusEnabled = True Then
    Select Case DataBonusMode
    Case 0
        lblBonLeft.Caption = DataDevi32(DataBonusBoth)
        lblBonLeft.ToolTipText = DataDeviTip(DataBonusBoth)
    Case 1
        lblBonLeft.Caption = DataDevi(DataBonusXmited) + "/" + DataDevi(DataBonusRcved)
        lblBonLeft.ToolTipText = DataDeviTip(DataBonusXmited) + "/" + DataDeviTip(DataBonusRcved)
    End Select
Else
    lblBonLeft.Caption = "-"
    lblBonLeft.ToolTipText = ""
End If

  
  DataRcvedWeek = DataRcvedWeek + iph_received_delta
  DataXmitedWeek = DataXmitedWeek + iph_transmitted_delta

  ColorText lbXmitedAll, DataXmitedAll + DataRcvedAll, DataXmitedAll, DataRcvedAll
  ColorText lbXmitedToday, DataXmitedToday + DataRcvedToday, DataXmitedToday, DataRcvedToday
  ColorText lbXmitedWeek, DataXmitedWeek + DataRcvedWeek, DataXmitedWeek, DataRcvedWeek
  ColorText lbXmitedMonth, DataXmitedMonth + DataRcvedMonth, DataXmitedMonth, DataRcvedMonth
  
' Sleep 10

UpdateFluentWindow



  If ActiveLeft > 0 Then
   frmOK.Label4.Caption = localize_do("WNDA04", "Остаток (трафик в лимите)")
   frmOK.lAll.ForeColor = vbGreen
   frmOK.lAll.Caption = DataDevi(ActiveLeft)
   frmOK.lAll.ToolTipText = DataDeviTip(ActiveLeft)
  ElseIf ActiveLeft < 0 Then
   frmOK.Label4.Caption = localize_do("WNDA06", "Трафик за лимитом")
   frmOK.lAll.ForeColor = vbRed
   frmOK.lAll.Caption = DataDevi(ActiveLeft)
   frmOK.lAll.ToolTipText = DataDeviTip(ActiveLeft)
  ElseIf ActiveLeft = 0 Then
   frmOK.Label4.Caption = localize_do("WNDA04", "Остаток (трафик в лимите)")
   frmOK.lAll.ForeColor = vbWhite
   frmOK.lAll.Caption = "-"
   frmOK.lAll.ToolTipText = ""
  End If
  
  
  frmOK.mDay.Caption = FormatEx(TaxToday, "### ##0.00") + " " + TaxName
  frmOK.mMonth.Caption = FormatEx(TaxMonth, "### ##0.00") + " " + TaxName
  frmOK.mSeans.Caption = FormatEx(TaxTax, "### ##0.00") + " " + TaxName
  
    lblCredit.Caption = FormatEx(CreditLeft, "### ### ### ##0.00") + " " + TaxName
    frmOK.mAll.Caption = FormatEx(CreditLeft, "### ### ### ##0.00") + " " + TaxName
    If CreditLeft > 0 Then
      lblCredit.ForeColor = vbGreen - RGB(0, 140, 0)
      frmOK.mAll.ForeColor = vbGreen
    Else
      lblCredit.ForeColor = vbRed
      frmOK.mAll.ForeColor = vbRed
    End If
  
  mDay.Caption = FormatEx(TaxToday, "### ##0.00") + " " + TaxName
  mMonth.Caption = FormatEx(TaxMonth, "### ##0.00") + " " + TaxName
  mAll.Caption = FormatEx(TaxAll, "### ##0.00") + " " + TaxName
  mSeans.Caption = FormatEx(TaxWeek, "### ##0.00") + " " + TaxName
  
  lbSinceSumm.Caption = FormatEx(TaxCount, "### ##0.00") + " " + TaxName
      
  Dim e, DL, UL, DLx, ULx
  
  For e = 1 To 1 Step -1
    MassSpeed(e, 0) = MassSpeed(e - 1, 0)
    MassSpeed(e, 1) = MassSpeed(e - 1, 1)
  Next e
  
If Not TimeLongBack = 0 Then
    TimeLongTemp = GetTickCount
    If TimeLongTemp >= TimeLongBack Then TimeLong = 1000 / (TimeLongTemp - TimeLongBack) Else TimeLong = 1
    TimeLongBack = TimeLongTemp
Else
    TimeLongTemp = GetTickCount
    TimeLong = 1
    TimeLongBack = TimeLongTemp
End If
  
  MassSpeed(0, 0) = iph_received_delta * TimeLong
  MassSpeed(0, 1) = iph_transmitted_delta * TimeLong
  
  frmOK.DrawGraph IIf(MassSpeed(0, 0) > MassSpeed(0, 1), MassSpeed(0, 0), MassSpeed(0, 1))
  
  DL = 0
  UL = 0
  
  For e = 0 To 1
   DL = DL + MassSpeed(e, 0)
   UL = UL + MassSpeed(e, 1)
  Next e
  
  Dim tmpDLx As Currency, tmpULx As Currency
  
  tmpDL = DL / 2
  tmpUL = UL / 2
  
  lbDlSpd.Caption = IIf(tmpDL > 0, DataDevi32(tmpDL) + localize_do("WRD003A", "/с"), "-")
  lbUlSpd.Caption = IIf(tmpUL > 0, DataDevi32(tmpUL) + localize_do("WRD003A", "/с"), "-")
  frmOK.lDl.Caption = IIf(tmpDL > 0, DataDevi(tmpDL) + localize_do("WRD003A", "/с"), "-")
  frmOK.lUl.Caption = IIf(tmpUL > 0, DataDevi(tmpUL) + localize_do("WRD003A", "/с"), "-")
  
  lbDlSpd.ToolTipText = IIf(tmpDL > 0, DataDeviBit32(tmpDL * 8) + localize_do("WRD003A", "/с"), "-")
  lbUlSpd.ToolTipText = IIf(tmpUL > 0, DataDeviBit32(tmpUL * 8) + localize_do("WRD003A", "/с"), "-")
  frmOK.lDl.ToolTipText = lbDlSpd.ToolTipText
  frmOK.lUl.ToolTipText = lbUlSpd.ToolTipText
  
  
  
  Dim xparam As Currency, yparam As Currency, zparam As Currency, rparam As Currency, qparam As Currency
  
  If tmpDL > MaxxSpeed(2) Then MaxxSpeed(2) = tmpDL
  If tmpUL > MaxxSpeed(1) Then MaxxSpeed(1) = tmpUL
  lblMaxx.Caption = DataDevi32(MaxxSpeed(2)) + localize_do("WRD003A", "/с")
  lblMaxx2.Caption = DataDevi32(MaxxSpeed(1)) + localize_do("WRD003A", "/с")
  
  lblMaxx.ToolTipText = DataDeviBit32(MaxxSpeed(2) * 8) + localize_do("WRD003A", "/с")
  lblMaxx2.ToolTipText = DataDeviBit32(MaxxSpeed(1) * 8) + localize_do("WRD003A", "/с")
  
  ' Counters
  lbSinceSend.Caption = FormatEx(DataXmitedCount, "### ### ### ### ##0") + " (" + DataDevi(DataXmitedCount) + ")"
  lbSinceRecev.Caption = FormatEx(DataRcvedCount, "### ### ### ### ##0") + " (" + DataDevi(DataRcvedCount) + ")"
  Label6.Caption = FormatEx(DataRcvedCount + DataXmitedCount, "### ### ### ### ##0") + " (" + DataDevi(DataRcvedCount + DataXmitedCount) + ")"
  

End Sub

Function DataDevi(InLong As Currency) As String
Select Case Use1024
Case False
    If MMod(InLong) >= 1000000000000# Then DataDevi = FormatEx(InLong / (1000 ^ 4), "0.##") + " " + localize_do("WRD012", "Т") + localize_do("WRD008", "Б"): Exit Function
    If MMod(InLong) >= 1000000000 Then DataDevi = FormatEx(InLong / (1000 ^ 3), "0.##") + " " + localize_do("WRD011", "Г") + localize_do("WRD008", "Б"): Exit Function
    If MMod(InLong) >= 1000000 Then DataDevi = FormatEx(InLong / (1000 ^ 2), "0.##") + " " + localize_do("WRD010", "М") + localize_do("WRD008", "Б"): Exit Function
    If MMod(InLong) >= 1000 Then DataDevi = FormatEx(InLong / (1000 ^ 1), "0.#") + " " + localize_do("WRD009", "К") + localize_do("WRD008", "Б"): Exit Function
    If MMod(InLong) >= 0 Then DataDevi = FormatEx(InLong, "0") + " " + localize_do("WRD008", "Б"): Exit Function
Case True
    If MMod(InLong) >= 1000000000000# Then DataDevi = FormatEx(InLong / (1024 ^ 4), "0.##") + " " + localize_do("WRD012", "Т") + localize_do("WRD008", "Б"): Exit Function
    If MMod(InLong) >= 1000000000 Then DataDevi = FormatEx(InLong / (1024 ^ 3), "0.##") + " " + localize_do("WRD011", "Г") + localize_do("WRD008", "Б"): Exit Function
    If MMod(InLong) >= 1000000 Then DataDevi = FormatEx(InLong / (1024 ^ 2), "0.##") + " " + localize_do("WRD010", "М") + localize_do("WRD008", "Б"): Exit Function
    If MMod(InLong) >= 1000 Then DataDevi = FormatEx(InLong / (1024 ^ 1), "0.#") + " " + localize_do("WRD009", "К") + localize_do("WRD008", "Б"): Exit Function
    If MMod(InLong) >= 0 Then DataDevi = FormatEx(InLong, "0") + " " + localize_do("WRD008", "Б"): Exit Function
End Select
End Function

Function DataDeviBit(InLong As Currency) As String
Select Case Use1024
Case True
If MMod(InLong) >= 1000000000000# Then DataDeviBit = FormatEx(InLong / (1024 ^ 4), "0.##") + " " + localize_do("WRD012", "Т") + LCase(localize_do("WRD008", "Б")): Exit Function
If MMod(InLong) >= 1000000000 Then DataDeviBit = FormatEx(InLong / (1024 ^ 3), "0.##") + " " + localize_do("WRD011", "Г") + LCase(localize_do("WRD008", "Б")): Exit Function
If MMod(InLong) >= 1000000 Then DataDeviBit = FormatEx(InLong / (1024 ^ 2), "0.##") + " " + localize_do("WRD010", "М") + LCase(localize_do("WRD008", "Б")):  Exit Function
If MMod(InLong) >= 1000 Then DataDeviBit = FormatEx(InLong / (1024 ^ 1), "0.#") + " " + localize_do("WRD009", "К") + LCase(localize_do("WRD008", "Б")):  Exit Function
If MMod(InLong) >= 0 Then DataDeviBit = FormatEx(InLong, "0") + " " + LCase(localize_do("WRD008", "Б")): Exit Function
Case False
If MMod(InLong) >= 1000000000000# Then DataDeviBit = FormatEx(InLong / 1000000000000#, "0.##") + " " + localize_do("WRD012", "Т") + LCase(localize_do("WRD008", "Б")): Exit Function
If MMod(InLong) >= 1000000000 Then DataDeviBit = FormatEx(InLong / 1000000000, "0.##") + " " + localize_do("WRD011", "Г") + LCase(localize_do("WRD008", "Б")): Exit Function
If MMod(InLong) >= 1000000 Then DataDeviBit = FormatEx(InLong / 1000000, "0.##") + " " + localize_do("WRD010", "М") + LCase(localize_do("WRD008", "Б")):  Exit Function
If MMod(InLong) >= 1000 Then DataDeviBit = FormatEx(InLong / 1000, "0.#") + " " + localize_do("WRD009", "К") + LCase(localize_do("WRD008", "Б")):  Exit Function
If MMod(InLong) >= 0 Then DataDeviBit = FormatEx(InLong, "0") + " " + LCase(localize_do("WRD008", "Б")): Exit Function
End Select
End Function

Function DataDeviBit32(InLong As Currency) As String
Select Case Use1024
Case True
If MMod(InLong) >= 1000000000000# Then DataDeviBit32 = FormatEx(InLong / (1024 ^ 4), "0.##") + " " + localize_do("WRD012", "Т") + LCase(localize_do("WRD013", "Бит")): Exit Function
If MMod(InLong) >= 1000000000 Then DataDeviBit32 = FormatEx(InLong / (1024 ^ 3), "0.##") + " " + localize_do("WRD011", "Г") + LCase(localize_do("WRD013", "Бит")): Exit Function
If MMod(InLong) >= 1000000 Then DataDeviBit32 = FormatEx(InLong / (1024 ^ 2), "0.##") + " " + localize_do("WRD010", "М") + LCase(localize_do("WRD013", "Бит")):  Exit Function
If MMod(InLong) >= 1000 Then DataDeviBit32 = FormatEx(InLong / (1024 ^ 1), "0.#") + " " + localize_do("WRD009", "К") + LCase(localize_do("WRD013", "Бит")):  Exit Function
If MMod(InLong) >= 0 Then DataDeviBit32 = FormatEx(InLong, "0") + " " + LCase(localize_do("WRD008", "Б")): Exit Function
Case False
If MMod(InLong) >= 1000000000000# Then DataDeviBit32 = FormatEx(InLong / 1000000000000#, "0.##") + " " + localize_do("WRD012", "Т") + LCase(localize_do("WRD013", "Бит")): Exit Function
If MMod(InLong) >= 1000000000 Then DataDeviBit32 = FormatEx(InLong / 1000000000, "0.##") + " " + localize_do("WRD011", "Г") + LCase(localize_do("WRD013", "Бит")): Exit Function
If MMod(InLong) >= 1000000 Then DataDeviBit32 = FormatEx(InLong / 1000000, "0.##") + " " + localize_do("WRD010", "М") + LCase(localize_do("WRD013", "Бит")):  Exit Function
If MMod(InLong) >= 1000 Then DataDeviBit32 = FormatEx(InLong / 1000, "0.#") + " " + localize_do("WRD009", "К") + LCase(localize_do("WRD013", "Бит")):  Exit Function
If MMod(InLong) >= 0 Then DataDeviBit32 = FormatEx(InLong, "0") + " " + LCase(localize_do("WRD008", "Б")): Exit Function
End Select
End Function


Function DataDevi32(InLong As Currency) As String
Select Case Use1024
Case True
    If MMod(InLong) >= 1000000000000# Then DataDevi32 = FormatEx(InLong / (1024 ^ 4), "0.##") + " " + localize_do("WRD012", "Т") + localize_do("WRD008", "Байт"): Exit Function
    If MMod(InLong) >= 1000000000 Then DataDevi32 = FormatEx(InLong / (1024 ^ 3), "0.##") + " " + localize_do("WRD011", "Г") + localize_do("WRD008", "Байт"): Exit Function
    If MMod(InLong) >= 1000000 Then DataDevi32 = FormatEx(InLong / (1024 ^ 2), "0.##") + " " + localize_do("WRD010", "М") + localize_do("WRD008", "Байт"):  Exit Function
    If MMod(InLong) >= 1000 Then DataDevi32 = FormatEx(InLong / (1024 ^ 1), "0.#") + " " + localize_do("WRD009", "К") + localize_do("WRD008", "Байт"):  Exit Function
    If MMod(InLong) >= 0 Then DataDevi32 = FormatEx(InLong, "0") + " " + localize_do("WRD007", "Байт"): Exit Function
Case False
    If MMod(InLong) >= 1000000000000# Then DataDevi32 = FormatEx(InLong / 1000000000000#, "0.##") + " " + localize_do("WRD012", "Т") + localize_do("WRD008", "Байт"): Exit Function
    If MMod(InLong) >= 1000000000 Then DataDevi32 = FormatEx(InLong / 1000000000, "0.##") + " " + localize_do("WRD011", "Г") + localize_do("WRD008", "Байт"): Exit Function
    If MMod(InLong) >= 1000000 Then DataDevi32 = FormatEx(InLong / 1000000, "0.##") + " " + localize_do("WRD010", "М") + localize_do("WRD008", "Байт"):  Exit Function
    If MMod(InLong) >= 1000 Then DataDevi32 = FormatEx(InLong / 1000, "0.#") + " " + localize_do("WRD009", "К") + localize_do("WRD008", "Байт"):  Exit Function
    If MMod(InLong) >= 0 Then DataDevi32 = FormatEx(InLong, "0") + " " + localize_do("WRD007", "Байт"): Exit Function
End Select
End Function

Function DataDeviTip32(ddtRC As Currency, ddtXM As Currency) As String
DataDeviTip32 = localize_do("ADD003", "Получено") & ": " & FormatEx(ddtRC, "### ### ### ### ##0") + " " + localize_do("WRD007", "Байт") + " / " + _
              localize_do("ADD002", "Отправлено") & ": " & FormatEx(ddtXM, "### ### ### ### ##0") + " " + localize_do("WRD007", "Байт") + " / " + _
              localize_do("ADD004", "Сумма") & ": " & FormatEx(ddtRC + ddtXM, "### ### ### ### ##0") + " " + localize_do("WRD007", "Байт")

End Function


Function DataDeviTip(ddtSM As Currency) As String
DataDeviTip = FormatEx(ddtSM, "### ### ### ### ##0") + " " + localize_do("WRD007", "Байт")

End Function

Function DataDeviSpeed32(InLong As Currency) As String
    If MMod(InLong) >= 1000000000000# Then DataDeviSpeed32 = FormatEx(InLong / 1000000000000#, "0.##") + " " + localize_do("WRD012", "Т") + LCase(localize_do("WRD013", "бит")): Exit Function
    If MMod(InLong) >= 1000000000 Then DataDeviSpeed32 = FormatEx(InLong / 1000000000, "0.##") + " " + localize_do("WRD011", "Г") + LCase(localize_do("WRD013", "бит")): Exit Function
    If MMod(InLong) >= 1000000 Then DataDeviSpeed32 = FormatEx(InLong / 1000000, "0.##") + " " + localize_do("WRD010", "М") + LCase(localize_do("WRD013", "бит")):  Exit Function
    If MMod(InLong) >= 1000 Then DataDeviSpeed32 = FormatEx(InLong / 1000, "0.#") + " " + localize_do("WRD009", "К") + LCase(localize_do("WRD013", "бит")):  Exit Function
    If MMod(InLong) >= 0 Then DataDeviSpeed32 = FormatEx(InLong, "0") + " " + LCase(localize_do("WRD013", "бит")): Exit Function
End Function



Sub ColorText(inLabel As Label, inBytes As Currency, inXm As Currency, inRc As Currency)

Dim N As String

If inBytes >= 0 Then N = RGB(0, 0, 0)               ' 0 < 1,000,000
If inBytes >= 1000000 Then N = RGB(0, 150, 150)     ' 1,000,000 < 100,000,000
If inBytes >= 100000000 Then N = RGB(0, 0, 250)     ' 100,000,000 < 1,000,000,000
If inBytes >= 1000000000 Then N = RGB(200, 0, 200)    ' 1,000,000,000 < 5,000,000,000
If inBytes >= 5000000000# Then N = RGB(250, 100, 200)  ' 5,000,000,000

inLabel.ForeColor = N
inLabel.Caption = DataDevi(inBytes)
inLabel.ToolTipText = DataDeviTip32(inRc, inXm)


End Sub

Private Sub tmrAverage_Timer()

If ProgConnected = True Then
SaveTodayStatistish DataXmitedToday, DataRcvedToday, TaxToday
SaveWeekStatistish DataXmitedWeek, DataRcvedWeek, TaxWeek
SaveMonthStatistish DataXmitedMonth, DataRcvedMonth, TaxMonth
SaveYearStatistish DataXmitedYear, DataRcvedYear, TaxYear
SaveCurrentHourStatistic DataXmitedHour, DataRcvedHour, TaxHour, CurrentHour

UpdateAverage
End If

End Sub





Function GetConnType(IND As Integer) As String
Select Case IND
Case 6: GetConnType = "Ethernet"
Case 15: GetConnType = "FDDI"
Case 24: GetConnType = "Loopback"
Case 1: GetConnType = "Other"
Case 23: GetConnType = "PPP"
Case 28: GetConnType = "SLIP"
Case 9: GetConnType = "TokenRing"
End Select
End Function

Private Sub tmrEE_Timer()
On Error Resume Next
Dim orX As Currency, orR As Currency
Dim tmpPicture As String


If DataXmitedTemp < DataXmitedToday = 0 Then DataXmitedTemp = DataXmitedToday
If DataRcvedTemp < DataRcvedToday = 0 Then DataRcvedTemp = DataRcvedToday
Dim dHandle As Long

Blinker = Not Blinker

orX = (DataXmitedToday - DataXmitedTemp)
orR = (DataRcvedToday - DataRcvedTemp)


If ProgConnected = False Then
  dHandle = picStatus(0).Picture
Else
  If PingNetwork = True And host_is_alive = False Then
    dHandle = picStatus(5).Picture
  ElseIf MMod(orR - orX) <= MaxVal(orR, orX) / 1.8 And MaxVal(orR, orX) > 0 Then
    dHandle = picStatus(4).Picture ' Ul/Dl
  ElseIf orR > orX Then
    dHandle = picStatus(3).Picture ' Dl
  ElseIf orR < orX Then
    dHandle = picStatus(2).Picture ' Ul
  ElseIf PingNetwork = False Then
    dHandle = picStatus(1).Picture
  ElseIf PingNetwork = True And host_is_alive = True Then
    dHandle = picStatus(1).Picture
  End If
End If

If Blinker = True And CurrLimit = LimitStatus.Redline And TipPreLimit Then
    dHandle = picStatus(6).Picture ' Yellow Alert

ElseIf Blinker = True And CurrLimit = LimitStatus.Overload Then
    dHandle = picStatus(7).Picture ' Red Alert

End If


If ProgConnected = True Then TrayModify picTray, "", dHandle

If ProgConnected = False Then TrayModify picTray, Left(filter_interface_name(localize_do(iph_interface, iph_interface)), 45) + _
        " (" + NetworkText(NetworkStatus(0)) + ")", dHandle

DataRcvedTemp = DataRcvedToday
DataXmitedTemp = DataXmitedToday

End Sub


