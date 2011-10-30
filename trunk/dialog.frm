VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Dialog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Настройки"
   ClientHeight    =   10770
   ClientLeft      =   6105
   ClientTop       =   1575
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "dialog.frx":000C
   ScaleHeight     =   10770
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox opFrame 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   8
      Left            =   2940
      ScaleHeight     =   5415
      ScaleWidth      =   6855
      TabIndex        =   121
      Top             =   1140
      Width           =   6855
      Begin VB.CheckBox chkDelay 
         Caption         =   "с задержкой (сек):"
         Height          =   195
         Left            =   3840
         TabIndex        =   147
         Top             =   4020
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.TextBox txtDelay 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5880
         TabIndex        =   146
         Text            =   "0"
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "т"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6180
         TabIndex        =   139
         Top             =   2940
         Width           =   435
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "с"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6180
         TabIndex        =   138
         Top             =   2460
         Width           =   435
      End
      Begin VB.CheckBox chkStartNotify 
         Caption         =   "Уведомлять о запуске или закрытии программ"
         Height          =   195
         Left            =   300
         TabIndex        =   137
         Top             =   4920
         Width           =   6495
      End
      Begin VB.CheckBox prcLinkDown 
         Caption         =   "Понимать ""сеть недоступна"" как разрыв соединения (бета)"
         Height          =   195
         Left            =   300
         TabIndex        =   136
         Top             =   4620
         Width           =   6495
      End
      Begin VB.CheckBox prcStop 
         Caption         =   "Закрывать запущенные приложения при разрыве соединения (бета)"
         Height          =   195
         Left            =   300
         TabIndex        =   135
         Top             =   4320
         Width           =   6495
      End
      Begin VB.CheckBox prcEnable 
         Caption         =   "Использовать запуск приложений"
         Height          =   195
         Left            =   300
         TabIndex        =   134
         Top             =   4020
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin ComctlLib.ListView lstProcesses 
         Height          =   1995
         Left            =   180
         TabIndex        =   133
         Top             =   1320
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3519
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Программа"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Путь к программе"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.CommandButton prcRemove 
         Caption         =   "-"
         Height          =   435
         Left            =   6180
         TabIndex        =   131
         ToolTipText     =   "Удалить"
         Top             =   1860
         Width           =   435
      End
      Begin VB.CommandButton prcAdd 
         Caption         =   "+"
         Height          =   435
         Left            =   6180
         TabIndex        =   130
         ToolTipText     =   "Добавить"
         Top             =   1320
         Width           =   435
      End
      Begin VB.Image Image16 
         Height          =   15
         Left            =   240
         Picture         =   "dialog.frx":40FB
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   6360
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Опции"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   132
         Top             =   3480
         Width           =   780
      End
      Begin VB.Label Label27 
         Caption         =   $"dialog.frx":442E
         Height          =   675
         Left            =   180
         TabIndex        =   129
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Список запускаемых приложений"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   122
         Top             =   120
         Width           =   4260
      End
      Begin VB.Image Image15 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":44DB
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6360
      End
   End
   Begin VB.PictureBox opFrame 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Index           =   7
      Left            =   3300
      ScaleHeight     =   4935
      ScaleWidth      =   6855
      TabIndex        =   82
      Top             =   8220
      Width           =   6855
      Begin VB.TextBox txtVersion 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4635
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         Text            =   "dialog.frx":480E
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.PictureBox opFrame 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Index           =   6
      Left            =   2700
      ScaleHeight     =   5355
      ScaleWidth      =   6855
      TabIndex        =   76
      Top             =   8460
      Width           =   6855
      Begin VB.TextBox txtTitles 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   78
         Top             =   720
         Width           =   6375
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Woobind Software (C) 2007-2011"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   145
         Top             =   4920
         Width           =   2430
      End
      Begin VB.Label lblEMail 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Написать письмо автору"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3690
         MouseIcon       =   "dialog.frx":4814
         MousePointer    =   99  'Custom
         TabIndex        =   80
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label lblSiteVisit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Посетить сайт программы"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3660
         MouseIcon       =   "dialog.frx":4B1E
         MousePointer    =   99  'Custom
         TabIndex        =   79
         Top             =   4740
         Width           =   2235
      End
      Begin VB.Label lblAboutVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Woobind Network Meter   version 2.1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   77
         Top             =   120
         Width           =   4575
      End
      Begin VB.Image Image14 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":4E28
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6360
      End
   End
   Begin VB.PictureBox opFrame 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Index           =   5
      Left            =   2220
      ScaleHeight     =   5355
      ScaleWidth      =   6855
      TabIndex        =   74
      Top             =   8640
      Width           =   6855
      Begin VB.CheckBox cT1024 
         Caption         =   "Функция, угодная Богам"
         Height          =   195
         Left            =   300
         TabIndex        =   128
         Top             =   1200
         Width           =   6315
      End
      Begin VB.CommandButton cNotifyTest 
         Caption         =   "Тест"
         Height          =   375
         Left            =   5220
         TabIndex        =   119
         Top             =   4740
         Width           =   1455
      End
      Begin VB.CheckBox chkShowTop 
         Caption         =   "Показывать уведомления сверху"
         Height          =   195
         Left            =   300
         TabIndex        =   118
         Top             =   3360
         Width           =   3855
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "Использовать звуковые уведомления"
         Height          =   195
         Left            =   300
         TabIndex        =   117
         Top             =   3060
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Сброс"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5220
         TabIndex        =   113
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Настройка уведомлений"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   116
         Top             =   2520
         Width           =   3045
      End
      Begin VB.Image Image12 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":515B
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   6360
      End
      Begin VB.Label label99 
         Caption         =   "Для того, чтобы очистить показания всех счетчиков и отчеты по трафику, нажмите эту кнопку:"
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   300
         TabIndex        =   114
         Top             =   600
         Width           =   4455
      End
      Begin VB.Image Image13 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":548E
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6360
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Дополнительные возможности"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   75
         Top             =   120
         Width           =   3915
      End
   End
   Begin VB.PictureBox opFrame 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   4
      Left            =   1800
      ScaleHeight     =   5415
      ScaleWidth      =   6855
      TabIndex        =   92
      Top             =   8940
      Width           =   6855
      Begin VB.TextBox txtBonusSetted 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   125
         Text            =   "dialog.frx":57C1
         Top             =   3900
         Width           =   6255
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   6675
         TabIndex        =   93
         Top             =   660
         Width           =   6675
         Begin VB.CheckBox cNotNow 
            Caption         =   "Не применять к текущему дню/месяцу"
            Height          =   495
            Left            =   3180
            TabIndex        =   127
            Top             =   1440
            Width           =   3495
         End
         Begin VB.ComboBox cbDir 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   2100
            Width           =   1995
         End
         Begin VB.CommandButton cmdEnableBonus 
            Caption         =   "Применить"
            Height          =   375
            Left            =   4140
            TabIndex        =   111
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtBothBonus 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   4200
            TabIndex        =   101
            Text            =   "0"
            Top             =   900
            Width           =   855
         End
         Begin VB.OptionButton optMode 
            Caption         =   "Общий трафик"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   3180
            TabIndex        =   100
            Top             =   600
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.TextBox txtRcvBonus 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   99
            Text            =   "0"
            Top             =   1260
            Width           =   855
         End
         Begin VB.TextBox txtXmitBonus 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   98
            Text            =   "0"
            Top             =   900
            Width           =   855
         End
         Begin VB.OptionButton optMode 
            Caption         =   "Раздельный трафик"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   97
            Top             =   600
            Width           =   2475
         End
         Begin VB.CheckBox chkValBonus 
            Caption         =   "Стоимость:"
            Height          =   195
            Left            =   180
            TabIndex        =   96
            Top             =   1740
            Width           =   1335
         End
         Begin VB.TextBox txtMinus 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   95
            Text            =   "0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.CheckBox chkBonus 
            Caption         =   "Использовать оптовый трафик"
            Height          =   195
            Left            =   180
            TabIndex        =   94
            Top             =   60
            Value           =   2  'Grayed
            Width           =   3075
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Повторять:"
            Height          =   195
            Left            =   720
            TabIndex        =   126
            Top             =   2160
            Width           =   885
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Общий:"
            Height          =   195
            Left            =   3000
            TabIndex        =   108
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "МБайт"
            Height          =   195
            Left            =   5160
            TabIndex        =   107
            Top             =   960
            Width           =   540
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "МБайт"
            Height          =   195
            Left            =   2520
            TabIndex        =   106
            Top             =   1320
            Width           =   540
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "МБайт"
            Height          =   195
            Left            =   2520
            TabIndex        =   105
            Top             =   960
            Width           =   540
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Входящий:"
            Height          =   195
            Left            =   0
            TabIndex        =   104
            Top             =   1320
            Width           =   1350
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Исходящий:"
            Height          =   195
            Left            =   0
            TabIndex        =   103
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "грн."
            Height          =   195
            Left            =   2520
            TabIndex        =   102
            Top             =   1740
            Width           =   345
         End
      End
      Begin VB.Image Image10 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":5815
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   6360
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Текущее состояние"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   110
         Top             =   3420
         Width           =   2475
      End
      Begin VB.Image Image11 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":5B48
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6360
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Изменение/установка"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   109
         Top             =   120
         Width           =   2820
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Проверить обновление"
      Height          =   375
      Left            =   180
      TabIndex        =   86
      Top             =   6780
      Width           =   2535
   End
   Begin VB.PictureBox opFrame 
      BorderStyle     =   0  'None
      Height          =   5355
      Index           =   3
      Left            =   1320
      ScaleHeight     =   5355
      ScaleWidth      =   6795
      TabIndex        =   60
      Top             =   9300
      Width           =   6795
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   5160
         ScaleHeight     =   1155
         ScaleWidth      =   1335
         TabIndex        =   88
         Top             =   960
         Width           =   1335
         Begin VB.CommandButton Command3 
            Caption         =   "Добавить..."
            Height          =   375
            Left            =   0
            TabIndex        =   90
            Top             =   0
            Width           =   1275
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Удалить"
            Height          =   375
            Left            =   0
            TabIndex        =   89
            Top             =   420
            Width           =   1275
         End
         Begin VB.Label lblAch 
            Alignment       =   2  'Center
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   480
            TabIndex        =   112
            Top             =   900
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.ComboBox cmbDirect 
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   3660
         Width           =   1815
      End
      Begin VB.CheckBox chkBalonOK 
         Caption         =   "Уведомлять при изменении стоимости в связи со временем"
         Height          =   255
         Left            =   180
         TabIndex        =   72
         Top             =   4560
         Width           =   5955
      End
      Begin VB.ComboBox cmbCurrency 
         Height          =   315
         Left            =   5280
         TabIndex        =   70
         Top             =   4140
         Width           =   1215
      End
      Begin VB.TextBox txtCredits 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   68
         Text            =   "0"
         Top             =   4140
         Width           =   1215
      End
      Begin VB.TextBox txtStaticTariff 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4140
         TabIndex        =   66
         Text            =   "0"
         Top             =   2220
         Width           =   975
      End
      Begin ComctlLib.ListView lstTariffs 
         Height          =   1215
         Left            =   240
         TabIndex        =   65
         Top             =   960
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Время"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Дни недели"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Цена"
            Object.Width           =   707
         EndProperty
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Направление трафика"
         Height          =   195
         Left            =   2760
         TabIndex        =   71
         Top             =   3720
         Width           =   1710
      End
      Begin VB.Label Label22 
         Caption         =   "Денежные единицы"
         Height          =   195
         Left            =   3480
         TabIndex        =   69
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "Кредиты (остаток):"
         Height          =   195
         Left            =   180
         TabIndex        =   67
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Тариф по умолчанию. Стоимость 1 МБ данных:"
         Height          =   195
         Left            =   300
         TabIndex        =   64
         Top             =   2280
         Width           =   3600
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Динамический тариф"
         Height          =   195
         Left            =   300
         TabIndex        =   63
         Top             =   660
         Width           =   1620
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Дополнительно"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   62
         Top             =   3180
         Width           =   1980
      End
      Begin VB.Image Image9 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":5E7B
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   6360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Стоимость трафика"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   61
         Top             =   120
         Width           =   2430
      End
      Begin VB.Image Image8 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":61AE
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6360
      End
   End
   Begin VB.PictureBox opFrame 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Index           =   2
      Left            =   780
      ScaleHeight     =   4935
      ScaleWidth      =   6855
      TabIndex        =   6
      Top             =   9600
      Width           =   6855
      Begin VB.CheckBox chkTray 
         Caption         =   "Уведомлять если трафик исчерпан"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   3240
         Width           =   4815
      End
      Begin VB.CheckBox chkPreLimit 
         Caption         =   "Предварительное уведомление если трафика осталось менее..."
         Enabled         =   0   'False
         Height          =   435
         Left            =   240
         TabIndex        =   57
         Top             =   3600
         Width           =   4515
      End
      Begin VB.TextBox txtLimit 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   3300
         TabIndex        =   23
         Text            =   "0"
         Top             =   1200
         Width           =   1755
      End
      Begin VB.CheckBox chklimuse 
         Caption         =   "Использовать ограничитель трафика"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   3915
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   5715
         TabIndex        =   12
         Top             =   2100
         Width           =   5715
         Begin VB.OptionButton optLimit 
            Caption         =   "по счетчику"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   3900
            TabIndex        =   16
            Top             =   0
            Width           =   1755
         End
         Begin VB.OptionButton optLimit 
            Caption         =   "в день"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   1155
         End
         Begin VB.OptionButton optLimit 
            Caption         =   "в неделю"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   14
            Top             =   0
            Width           =   1395
         End
         Begin VB.OptionButton optLimit 
            Caption         =   "в месяц"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   2640
            TabIndex        =   13
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmbLine 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3660
         Width           =   1095
      End
      Begin VB.ComboBox cmbWay 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1620
         Width           =   2715
      End
      Begin VB.ComboBox cmbDividers 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.Image Image7 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":64E1
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   6360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Сообщения и уведомления"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   58
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Направление трафика"
         Height          =   195
         Left            =   600
         TabIndex        =   56
         Top             =   1680
         Width           =   1710
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Предельный объем трафика"
         Height          =   195
         Left            =   600
         TabIndex        =   55
         Top             =   1260
         Width           =   2175
      End
      Begin VB.Image Image6 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":6814
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ограничитель"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   54
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6120
         TabIndex        =   10
         Top             =   3720
         Width           =   180
      End
   End
   Begin VB.CommandButton btnApply 
      BackColor       =   &H00F0F0F0&
      Caption         =   "Применить"
      Height          =   375
      Left            =   8340
      TabIndex        =   53
      Top             =   6780
      Width           =   1455
   End
   Begin VB.PictureBox opFrame 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Index           =   1
      Left            =   360
      ScaleHeight     =   4875
      ScaleWidth      =   6795
      TabIndex        =   4
      Top             =   9900
      Width           =   6795
      Begin VB.CheckBox chkConOnly 
         Caption         =   "Показывать только при наличии соединения"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   115
         Top             =   2160
         Width           =   5415
      End
      Begin VB.CheckBox chkAlwaysVisible 
         Caption         =   "Закрепить всплывающее окно"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   1800
         Width           =   5415
      End
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   4140
         Width           =   1995
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   3780
         Width           =   1995
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   3420
         Width           =   1995
      End
      Begin VB.CheckBox chkFloatWindow 
         Caption         =   "Использовать всплывающее окно"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   5655
      End
      Begin VB.CheckBox chkSafe 
         Caption         =   "Защищать окно от случайного всплытия"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1080
         Width           =   5595
      End
      Begin VB.CheckBox chkOnTop 
         Caption         =   "Показывать поверх всех окон"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   5175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "За текущий месяц:"
         Height          =   195
         Left            =   540
         TabIndex        =   51
         Top             =   4200
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "За текущий день:"
         Height          =   195
         Left            =   540
         TabIndex        =   50
         Top             =   3840
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "За текущее соединение:"
         Height          =   195
         Left            =   540
         TabIndex        =   49
         Top             =   3480
         Width           =   1905
      End
      Begin VB.Image Image5 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":6B47
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   6360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Учитываемое направление трафика"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   45
         Top             =   2880
         Width           =   4560
      End
      Begin VB.Image Image4 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":6E7A
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Поведение всплывающего окна"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   44
         Top             =   120
         Width           =   4065
      End
   End
   Begin VB.PictureBox opFrame 
      BorderStyle     =   0  'None
      Height          =   5415
      Index           =   0
      Left            =   120
      ScaleHeight     =   5415
      ScaleWidth      =   6795
      TabIndex        =   2
      Top             =   10020
      Width           =   6795
      Begin VB.CommandButton cmGetSettingWindow 
         Caption         =   "Настройки"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5220
         TabIndex        =   140
         Top             =   1020
         Width           =   1335
      End
      Begin MSComctlLib.ImageCombo cmbInterfaces 
         Height          =   330
         Left            =   1380
         TabIndex        =   123
         Top             =   660
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "ImageList1"
      End
      Begin VB.CheckBox chkNotify 
         Caption         =   "Уведомлять при смене состояния подключения"
         Height          =   255
         Left            =   180
         TabIndex        =   84
         Top             =   1080
         Value           =   2  'Grayed
         Width           =   5595
      End
      Begin ComctlLib.Slider sliAlpha 
         Height          =   315
         Left            =   3060
         TabIndex        =   37
         Top             =   3900
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   327682
         Max             =   90
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin VB.CheckBox chkTransparent 
         Caption         =   "Прозрачность основного окна"
         Height          =   255
         Left            =   180
         TabIndex        =   43
         Top             =   3900
         Width           =   2895
      End
      Begin VB.CheckBox chkAttach 
         Caption         =   "Пристыковывать главное окно к краям экрана"
         Height          =   255
         Left            =   180
         TabIndex        =   42
         Top             =   3540
         Width           =   4695
      End
      Begin VB.CheckBox chkUpdate 
         Caption         =   "Проверять доступность новой версий программы"
         Height          =   255
         Left            =   180
         TabIndex        =   41
         Top             =   3180
         Width           =   4395
      End
      Begin VB.CheckBox chkAutorun 
         Caption         =   "Загружать программу вместе с Windows"
         Height          =   255
         Left            =   180
         TabIndex        =   40
         Top             =   2820
         Width           =   3615
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         Pattern         =   "*.slf"
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.ComboBox cmbFiles 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   4920
         Width           =   4815
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         ScaleHeight     =   375
         ScaleWidth      =   6195
         TabIndex        =   17
         Top             =   1740
         Width           =   6195
         Begin VB.TextBox pingAddr 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4080
            TabIndex        =   18
            Text            =   "192.168.0.1"
            Top             =   0
            Width           =   1995
         End
         Begin VB.OptionButton pingSel 
            Caption         =   "Св&ой:"
            Height          =   195
            Index           =   1
            Left            =   2700
            TabIndex        =   20
            Top             =   60
            Width           =   1215
         End
         Begin VB.OptionButton pingSel 
            Caption         =   "&Стандартный "
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   19
            Top             =   60
            Width           =   2535
         End
      End
      Begin VB.CheckBox chkPing 
         Caption         =   "&Проверять доступность сети используя домен/IP:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1440
         Value           =   1  'Checked
         Width           =   6075
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Язык интерфейса"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   4965
         Width           =   1350
      End
      Begin VB.Label lbAlphaLevel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10 %"
         Height          =   195
         Left            =   5700
         TabIndex        =   38
         Top             =   3900
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":71AD
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   6360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Выбор языка"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   36
         Top             =   4380
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Опции"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   33
         Top             =   2280
         Width           =   780
      End
      Begin VB.Image Image2 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":74E0
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   6360
      End
      Begin VB.Label TLInterface 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Соединение"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   32
         Top             =   120
         Width           =   1485
      End
      Begin VB.Image Image1 
         Height          =   15
         Left            =   180
         Picture         =   "dialog.frx":7813
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6355
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Интерфейс"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   1005
      End
   End
   Begin VB.PictureBox picMenu 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   180
      ScaleHeight     =   5475
      ScaleWidth      =   2535
      TabIndex        =   25
      Top             =   1140
      Width           =   2535
      Begin VB.Timer tmrVisible 
         Interval        =   100
         Left            =   2100
         Top             =   3780
      End
      Begin VB.Image mnuICO 
         Height          =   240
         Index           =   8
         Left            =   180
         Picture         =   "dialog.frx":7B46
         Top             =   4935
         Width           =   240
      End
      Begin VB.Label mnuITM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Запуск приложений"
         Height          =   255
         Index           =   8
         Left            =   540
         TabIndex        =   120
         Top             =   4935
         Width           =   1935
      End
      Begin VB.Image mnuICO 
         Height          =   240
         Index           =   7
         Left            =   180
         Picture         =   "dialog.frx":7F4D
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label mnuITM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "История версий"
         Height          =   255
         Index           =   7
         Left            =   540
         TabIndex        =   81
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Image mnuICO 
         Height          =   240
         Index           =   5
         Left            =   180
         Picture         =   "dialog.frx":81AE
         Top             =   3405
         Width           =   240
      End
      Begin VB.Label mnuITM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "О программе"
         Height          =   255
         Index           =   6
         Left            =   540
         TabIndex        =   73
         Top             =   3945
         Width           =   1935
      End
      Begin VB.Image mnuICO 
         Height          =   240
         Index           =   6
         Left            =   180
         Picture         =   "dialog.frx":8538
         Top             =   3930
         Width           =   240
      End
      Begin VB.Label mnuITM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Дополнительно"
         Height          =   255
         Index           =   5
         Left            =   540
         TabIndex        =   59
         Top             =   3420
         Width           =   1935
      End
      Begin VB.Image mnuICO 
         Height          =   240
         Index           =   4
         Left            =   180
         Picture         =   "dialog.frx":87A1
         Top             =   2870
         Width           =   240
      End
      Begin VB.Image mnuICO 
         Height          =   240
         Index           =   3
         Left            =   180
         Picture         =   "dialog.frx":8A1F
         Top             =   2325
         Width           =   240
      End
      Begin VB.Image mnuICO 
         Height          =   240
         Index           =   2
         Left            =   180
         Picture         =   "dialog.frx":8C66
         Top             =   1785
         Width           =   240
      End
      Begin VB.Image mnuICO 
         Height          =   240
         Index           =   1
         Left            =   180
         Picture         =   "dialog.frx":8EFB
         Top             =   1260
         Width           =   240
      End
      Begin VB.Image mnuICO 
         Height          =   240
         Index           =   0
         Left            =   180
         Picture         =   "dialog.frx":9186
         Top             =   735
         Width           =   240
      End
      Begin VB.Label l0 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Меню"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   2280
      End
      Begin VB.Label mnuITM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Оптовый трафик"
         Height          =   255
         Index           =   4
         Left            =   540
         TabIndex        =   30
         Top             =   2895
         Width           =   1935
      End
      Begin VB.Label mnuITM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Тарифы"
         Height          =   255
         Index           =   3
         Left            =   540
         TabIndex        =   29
         Top             =   2355
         Width           =   1935
      End
      Begin VB.Label mnuITM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ограничитель трафика"
         Height          =   255
         Index           =   2
         Left            =   540
         TabIndex        =   28
         Top             =   1815
         Width           =   1935
      End
      Begin VB.Label mnuITM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Всплывающее окно"
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   27
         Top             =   1275
         Width           =   1935
      End
      Begin VB.Label mnuITM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Основные настройки"
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   26
         Top             =   735
         Width           =   1935
      End
      Begin VB.Shape mnuSEL 
         BackColor       =   &H00FFE8D3&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00F3A67B&
         Height          =   495
         Left            =   0
         Top             =   600
         Width           =   2535
      End
      Begin VB.Shape mnuGHOST 
         BackColor       =   &H00FFF9F9&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFE8D3&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   6780
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "ОК"
      Default         =   -1  'True
      Height          =   375
      Left            =   5100
      TabIndex        =   0
      Top             =   6780
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dialog.frx":93F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dialog.frx":9747
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dialog.frx":9A9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dialog.frx":9DEF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5595
      Left            =   2880
      TabIndex        =   91
      Top             =   1020
      Width           =   6975
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Made in Ukraine"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3360
      TabIndex        =   144
      Top             =   6840
      Width           =   1140
   End
   Begin VB.Label lVer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.2.345"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   1020
      TabIndex        =   143
      Top             =   540
      Width           =   1140
   End
   Begin VB.Label mnuTCAP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Меню"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9050
      TabIndex        =   141
      Top             =   270
      Width           =   690
   End
   Begin VB.Image mnuTICO 
      Height          =   240
      Left            =   8580
      Picture         =   "dialog.frx":A143
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lREP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Меню"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9060
      TabIndex        =   142
      Top             =   285
      Width           =   690
   End
End
Attribute VB_Name = "Dialog"
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


Option Explicit

Dim ApplicationsChanged As Boolean

Function AbonPeriod(index As abPeriod) As String
    
    Select Case index
    Case 0
        AbonPeriod = LCase(localize_do("B324C0", "нет"))
    Case 1
        AbonPeriod = LCase(localize_do("B324C1", "ежедневно"))
    Case 2
        AbonPeriod = LCase(localize_do("B324C2", "ежемесячно"))
    Case Else
        AbonPeriod = LCase(localize_do("B324C3", "неизвестно"))
    End Select

End Function

Sub EnableBonusCtrls()
txtXmitBonus.Enabled = optMode(1).Value And def_any_to_bool(chkBonus.Value)
txtRcvBonus.Enabled = optMode(1).Value And def_any_to_bool(chkBonus.Value)
txtBothBonus.Enabled = optMode(0).Value And def_any_to_bool(chkBonus.Value)

If txtXmitBonus.Enabled Then txtXmitBonus.BackColor = vbWindowBackground Else txtXmitBonus.BackColor = vbButtonFace
If txtRcvBonus.Enabled Then txtRcvBonus.BackColor = vbWindowBackground Else txtRcvBonus.BackColor = vbButtonFace
If txtBothBonus.Enabled Then txtBothBonus.BackColor = vbWindowBackground Else txtBothBonus.BackColor = vbButtonFace

If Not txtXmitBonus.Enabled Then txtXmitBonus.Text = ""
If Not txtRcvBonus.Enabled Then txtRcvBonus.Text = ""
If Not txtBothBonus.Enabled Then txtBothBonus.Text = ""

txtMinus.Enabled = def_any_to_bool(chkValBonus.Value) And def_any_to_bool(chkBonus.Value)
If txtMinus.Enabled Then txtMinus.BackColor = vbWindowBackground Else txtMinus.BackColor = vbButtonFace


chkValBonus.Enabled = def_any_to_bool(chkBonus.Value)
optMode(0).Enabled = def_any_to_bool(chkBonus.Value)
optMode(1).Enabled = def_any_to_bool(chkBonus.Value)
cbDir.Enabled = def_any_to_bool(chkBonus.Value)

End Sub

Function FilterLanguage(inExpression As String, inLanguage As String) As String
    
    Dim strBeg As String, strEnd As String
    Dim LoadShift As Integer
    
    strBeg = InStr(inExpression, "<" & UCase(inLanguage) & ">")
    
    If strBeg > 0 Then
        LoadShift = Len("<" & UCase(inLanguage) & ">") + 2
        strEnd = InStr(inExpression, "</" & UCase(inLanguage) & ">")
        If strEnd > 0 Then
            FilterLanguage = Mid(inExpression, strBeg + LoadShift, strEnd - strBeg - LoadShift)
            Exit Function
        End If
    End If
    
    FilterLanguage = "Sorry, but I can't parse 'credits.txt' :("
    
End Function

Sub ResetBonus()
    DataBonusBoth = 0
    DataBonusRcved = 0
    DataBonusXmited = 0
    DataBonusEnabled = False
End Sub

Sub SetBonuses()

 Dim Kn: Kn = vbYes
 If (DataBonusBoth > 0 Or DataBonusRcved > 0 Or DataBonusXmited > 0) And DataBonusEnabled Then
   Kn = MsgBox(localize_do("M1", "Внимание! Оптовый трафик указанный ранее еще не был исчерпан! Вы уверены что хотите внести изменения?"), vbExclamation + vbYesNo, "Woobind Network Meter")
 End If
 
 If Kn = vbYes Then
 If chkBonus.Value = 0 Then ResetBonus: GoTo ResetBonusMark
    If chkValBonus.Value = 1 Then
     If Val(txtMinus.Text) > 0 Then
      If Not (cNotNow.Enabled And cNotNow.Value > 0) Then MinusMoney Val(txtMinus.Text)
     ElseIf chkValBonus.Value = 1 Then
      MsgBox localize_do("MESSAGE02", "Стоимость должна быть больше нуля!"), vbCritical, "Woobind Network Meter"
      Exit Sub
     End If
    End If

    If Not (cNotNow.Enabled And cNotNow.Value > 0) Then
        DataBonusBoth = Val(txtBothBonus.Text) * IIf(Use1024, 1024 ^ 2, 1000 ^ 2)
        DataBonusRcved = Val(txtRcvBonus.Text) * IIf(Use1024, 1024 ^ 2, 1000 ^ 2)
        DataBonusXmited = Val(txtXmitBonus.Text) * IIf(Use1024, 1024 ^ 2, 1000 ^ 2)
    End If
    
    If optMode(0).Value Then DataBonusMode = 0 Else DataBonusMode = 1
    
        DataBonusEnabled = True
    
            With Abonetic
                .aIO = Val(txtBothBonus.Text) * IIf(Use1024, 1024 ^ 2, 1000 ^ 2)
                .aI = Val(txtRcvBonus.Text) * IIf(Use1024, 1024 ^ 2, 1000 ^ 2)
                .aO = Val(txtXmitBonus.Text) * IIf(Use1024, 1024 ^ 2, 1000 ^ 2)
                .aMoney = IIf(chkValBonus > 0, Val(txtMinus.Text), 0)
                .aPeriod = cbDir.ListIndex
                Select Case .aPeriod
                Case 1
                    .aLastSetted = Day(Now)
                Case 2
                    .aLastSetted = Month(Now)
                End Select
            End With
    
    
ResetBonusMark:
    frmVelton.SaveToINI
    ReloadBonuses
 
 End If
 
End Sub


Sub MinusMoney(Value As Currency)
      CreditLeft = CreditLeft - Value
      TaxCount = TaxCount + Value
      TaxToday = TaxToday + Value
      TaxWeek = TaxWeek + Value
      TaxMonth = TaxMonth + Value
      TaxAll = TaxAll + Value
      TaxTax = TaxTax + Value
End Sub

Private Sub btnApply_Click()

Call SaveSettings          ' Сохранидь конфигурацио
Call SkinMe                  ' Обнофить SkinMe
Call RefreshInterface       ' Перерисавать интерфейсо
Call ReloadListsEx          ' Перечетадь списко

DoEvents                    ' Ничего не деладь
btnApply.Enabled = False    ' Без комментариеф

End Sub

Sub sC()
btnApply.Enabled = True
End Sub


Private Sub CancelButton_Click()
Unload Me
End Sub
Function GetIndexFrom(inName As String) As Integer
Dim r
For r = 1 To cmbInterfaces.ComboItems.Count
  If cmbInterfaces.ComboItems.Item(r).Key = inName Then GetIndexFrom = r: Exit Function
Next r
GetIndexFrom = -1
End Function





Private Sub cbDir_Click()
    cNotNow.Enabled = (cbDir.ListIndex > 0)
    If cbDir.ListIndex = 0 Then cNotNow.Value = 0
End Sub

Private Sub chkAlwaysVisible_Click()
sC
End Sub

Private Sub chkAttach_Click()
sC
End Sub

Private Sub chkAutorun_Click()
sC
End Sub


Private Sub chkBalonOK_Click()
sC
End Sub

Private Sub chkBonus_Click()
    EnableBonusCtrls
    Enablence
End Sub


Private Sub chkConOnly_Click()
sC
End Sub

Private Sub chkDelay_Click()
  sC
  Call CtrEnable(def_any_to_bool(chkDelay.Value), txtDelay)
End Sub

Private Sub chkFloatWindow_Click()
sC
chkSafe.Enabled = def_any_to_bool(chkFloatWindow.Value)
chkOnTop.Enabled = def_any_to_bool(chkFloatWindow.Value)
chkConOnly.Enabled = def_any_to_bool(chkFloatWindow.Value)
chkAlwaysVisible.Enabled = def_any_to_bool(chkFloatWindow.Value)

Combo1.Enabled = def_any_to_bool(chkFloatWindow.Value)
Combo2.Enabled = def_any_to_bool(chkFloatWindow.Value)
Combo3.Enabled = def_any_to_bool(chkFloatWindow.Value)

End Sub

Private Sub chklimuse_Click()
txtLimit.Enabled = def_any_to_bool(chklimuse.Value)
cmbDividers.Enabled = def_any_to_bool(chklimuse.Value)
cmbWay.Enabled = def_any_to_bool(chklimuse.Value)
optLimit(0).Enabled = def_any_to_bool(chklimuse.Value)
optLimit(1).Enabled = def_any_to_bool(chklimuse.Value)
optLimit(2).Enabled = def_any_to_bool(chklimuse.Value)
optLimit(3).Enabled = def_any_to_bool(chklimuse.Value)
cmbLine.Enabled = def_any_to_bool(chklimuse.Value)
chkTray.Enabled = def_any_to_bool(chklimuse.Value)
chkPreLimit.Enabled = def_any_to_bool(chklimuse.Value)
cmbLine.Enabled = def_any_to_bool(chklimuse.Value) And def_any_to_bool(chkPreLimit.Value)
sC
End Sub

Private Sub chkNotify_Click()
sC
End Sub

Private Sub chkOnTop_Click()
sC
End Sub

Private Sub chkPing_Click()
sC
If chkPing.Value = 0 Then
  pingSel(0).Enabled = False
  pingSel(1).Enabled = False
  pingAddr.Enabled = False
Else
  pingSel(0).Enabled = True
  pingSel(1).Enabled = True
  pingAddr.Enabled = True
End If
  
End Sub




Private Sub chkPreLimit_Click()
cmbLine.Enabled = def_any_to_bool(chklimuse.Value) And def_any_to_bool(chkPreLimit.Value)
sC
End Sub

Private Sub chkSafe_Click()
sC
End Sub

Private Sub chkShowTop_Click()
sC
End Sub

Private Sub chkSound_Click()
sC
End Sub

Private Sub chkStartNotify_Click()
  sC
End Sub

Private Sub chkTransparent_Click()
sliAlpha.Enabled = def_any_to_bool(chkTransparent.Value)
sC
End Sub


Private Sub chkTray_Click()
sC
End Sub

Private Sub chkUpdate_Click()
sC
End Sub

Private Sub chkValBonus_Click()
EnableBonusCtrls
End Sub


Private Sub cmbCurrency_Change()
sC
End Sub

Private Sub cmbCurrency_Click()
sC
End Sub

Private Sub cmbDirect_Click()
sC
End Sub

Private Sub cmbDividers_Click()
sC

End Sub

Private Sub cmbFiles_Click()
sC

End Sub

Private Sub cmbInterfaces_Change()
sC
End Sub

Private Sub cmbInterfaces_Click()
sC
End Sub

Private Sub cmbLine_Click()
sC

End Sub

Private Sub cmbLine_KeyPress(KeyAscii As Integer)


If InStr(1, "0123456789" & vbBack, Chr(KeyAscii)) > 0 Then

  Exit Sub
  
End If


SkipEnter:

KeyAscii = 0

End Sub

Private Sub cmbWay_Click()
sC

End Sub

Private Sub cmdEnableBonus_Click()

    SetBonuses
    
End Sub

Private Sub cmdMoveDown_Click()
    
    On Error Resume Next
    
    If IsNothing(lstProcesses.SelectedItem) = False Then
        If lstProcesses.SelectedItem.index < lstProcesses.ListItems.Count Then
            ExchangeWith lstProcesses.ListItems(lstProcesses.SelectedItem.index + 1), lstProcesses.SelectedItem
            Call sC
            ApplicationsChanged = True
        End If
    End If

End Sub

Private Sub cmdMoveUp_Click()

    On Error Resume Next
    
    If IsNothing(lstProcesses.SelectedItem) = False Then
        If lstProcesses.SelectedItem.index > 1 Then
            ExchangeWith lstProcesses.ListItems(lstProcesses.SelectedItem.index - 1), lstProcesses.SelectedItem
            Call sC
            ApplicationsChanged = True
        End If
    End If

End Sub


Sub ExchangeWith(inValue As ComctlLib.ListItem, outValue As ComctlLib.ListItem)

    Dim tmpStr As String
    
    tmpStr = outValue.Text
    outValue.Text = inValue.Text: inValue.Text = tmpStr
    
    tmpStr = outValue.Tag
    outValue.Tag = inValue.Tag: inValue.Tag = tmpStr
    
    tmpStr = outValue.SubItems(1)
    outValue.SubItems(1) = inValue.SubItems(1): inValue.SubItems(1) = tmpStr
    
    inValue.Selected = True
    
End Sub

Private Sub cNotifyTest_Click()
PoP.ShowPopup "Woobind Network Meter", localize_do("TEST01", "Это обычное сообщение\nКликните чтобы закрыть..."), 0, def_any_to_bool(chkSound.Value), def_any_to_bool(chkShowTop.Value)
PoP.ShowPopup "Woobind Network Meter", localize_do("TEST02", "Это простое уведомление"), 1, def_any_to_bool(chkSound.Value), def_any_to_bool(chkShowTop.Value)
PoP.ShowPopup "Woobind Network Meter", localize_do("TEST03", "Это уведомление повышеной важности"), 2, def_any_to_bool(chkSound.Value), def_any_to_bool(chkShowTop.Value)
PoP.ShowPopup "Woobind Network Meter", localize_do("TEST04", "А это сообщение высокой важности!\nКликните мышкой чтобы закрыть..."), 3, def_any_to_bool(chkSound.Value), def_any_to_bool(chkShowTop.Value)
End Sub

Private Sub Combo1_Click()
sC

End Sub

Private Sub Combo2_Click()
sC
End Sub

Private Sub Combo3_Click()
sC
End Sub

Private Sub Command1_Click()

On Error Resume Next

If MsgBox(localize_do("MSG010", "Вы действительно хотите очистить отчет и обнулить все счетчики?"), vbQuestion + vbYesNo, "Woobind Network Meter") = vbYes Then
   Kill def_complete_path(App.Path) + "app_data.rpt"
   frmVelton.mReall_Click
End If

End Sub

Function FCap(Expression)
    
    Dim N As String
    N = Expression
    Mid(N, 1, 1) = UCase(Mid(N, 1, 1))
    FCap = N
End Function

Sub SkinMe()

l0.Caption = localize_do("B324E4", "Меню")
' /// ABOUT
lblAboutVersion.Caption = "Woobind Network Meter   " + localize_do("B324E3", "версия") + vb32 + GetVersion
lVer.Caption = FCap(localize_do("B324E3", "версия")) + vb32 + GetVersion

Label18.Caption = localize_do("NOT01")
chkSound.Caption = localize_do("NOT02")
chkShowTop.Caption = localize_do("NOT03")
cNotifyTest.Caption = localize_do("NOT04")

mnuITM(0).Caption = localize_do("INIP1")
mnuITM(1).Caption = localize_do("INIP2")
mnuITM(2).Caption = localize_do("INIP3")
mnuITM(3).Caption = localize_do("INIP4")
mnuITM(4).Caption = localize_do("INIP5")
mnuITM(5).Caption = localize_do("INIP6")
mnuITM(6).Caption = localize_do("INIP7")
mnuITM(7).Caption = localize_do("INIP8")
mnuITM(8).Caption = localize_do("B326A0")

TLInterface.Caption = localize_do("INIP1A1")
Label3.Caption = localize_do("INIP1L2")
chkNotify.Caption = localize_do("INIP1L3")
chkPing.Caption = localize_do("INIP1L4")
pingSel(0).Caption = localize_do("INIP1L5")
pingSel(1).Caption = localize_do("INIP1L6")

Label1.Caption = localize_do("INIP1A2")
chkAutorun.Caption = localize_do("INIP1L7")
chkUpdate.Caption = localize_do("INIP1L8")
chkAttach.Caption = localize_do("INIP1L9")
chkTransparent.Caption = localize_do("INIP1L10")

Label2.Caption = localize_do("INIP1A3")
Label6.Caption = localize_do("INIP1L11")


Label7.Caption = localize_do("INIP2A1")
chkFloatWindow.Caption = localize_do("INIP2L1")
chkSafe.Caption = localize_do("INIP2L2")
chkOnTop.Caption = localize_do("INIP2L3")
chkAlwaysVisible.Caption = localize_do("INIP2L4")

Label12.Caption = localize_do("INIP2A2")
Label8.Caption = localize_do("INIP2L5")
Label9.Caption = localize_do("INIP2L6")
Label10.Caption = localize_do("INIP2L7")
chkConOnly.Caption = localize_do("INIP2L8")

Label11.Caption = localize_do("INIP3A1")
chklimuse.Caption = localize_do("INIP3L1")
Label13.Caption = localize_do("INIP3L2")
Label14.Caption = localize_do("INIP3L3")

Label4.Caption = localize_do("INIP3A2")
chkTray.Caption = localize_do("INIP3L4")
chkPreLimit.Caption = localize_do("INIP3L5")

optLimit(0).Caption = localize_do("INIP3L6", "на день")
optLimit(1).Caption = localize_do("INIP3L7", "на неделю")
optLimit(2).Caption = localize_do("INIP3L8", "на месяц")
optLimit(3).Caption = localize_do("INIP3L9", "по счетчику")

Label15.Caption = localize_do("INIP4A1")
Label17.Caption = localize_do("INIP4L1")
lstTariffs.ColumnHeaders(1).Text = localize_do("INIP4L2")
lstTariffs.ColumnHeaders(2).Text = localize_do("INIP4L3")
lstTariffs.ColumnHeaders(3).Text = localize_do("INIP4L4")

Command3.Caption = localize_do("INIP4L5")
Command5.Caption = localize_do("INIP4L6")
Label19.Caption = localize_do("INIP4L7")

Label16.Caption = localize_do("INIP4A2")
Label21.Caption = localize_do("INIP4L8")

Label20.Caption = localize_do("INIP4L9")
Label22.Caption = localize_do("INIP4L10")
chkBalonOK.Caption = localize_do("INIP4L12")


Label23.Caption = localize_do("INIP5A1")
chkBonus.Caption = localize_do("INIP5L0")
optMode(1).Caption = localize_do("INIP5L1")
optMode(0).Caption = localize_do("INIP5L2")

Label28.Caption = localize_do("INIP5L3")
Label29.Caption = localize_do("INIP5L4")
Label33.Caption = localize_do("INIP5L5")

chkValBonus.Caption = localize_do("INIP5L6")
Label34.Caption = localize_do("INIP5A2")


Label30.Caption = localize_do("WRD010", "М") & localize_do("WRD007", "Байт")
Label31.Caption = Label30.Caption
Label32.Caption = Label31.Caption

Label26.Caption = TaxName


Label25.Caption = localize_do("INIP6A1")
label99.Caption = localize_do("INIP6L1")

Command1.Caption = localize_do("INIB5")
OKButton.Caption = localize_do("INIB1")
CancelButton.Caption = localize_do("INIB2")
btnApply.Caption = localize_do("INIB3")
Command2.Caption = localize_do("INIB4")

Label26.Caption = TaxName
cNotNow.Caption = localize_do("B324A0", "Не применять к текущему дню/месяцу")
Label24.Caption = localize_do("B324A1", "Повторять:")

lblSiteVisit.Caption = localize_do("B324E1", "Посетить Web-сайт программы")
lblEMail.Caption = localize_do("B324E2", "Написать письмо автору")

Me.Caption = localize_do("INI")
cmGetSettingWindow.Caption = Me.Caption

cmdEnableBonus.Caption = btnApply.Caption
cT1024.Caption = localize_do("B325A0", "Двоичные приставки (1Кб = 1024 байта, 1 МБ = 1024 КБ и т.д.)")

Label44.Caption = localize_do("B326A1")
Label27.Caption = localize_do("B326A2")
Label35.Caption = localize_do("B326A3")
prcEnable.Caption = localize_do("B326A4")
prcStop.Caption = localize_do("B326A5")
prcLinkDown.Caption = localize_do("B326A6")
chkStartNotify.Caption = localize_do("B326A7")
prcAdd.ToolTipText = localize_do("B326A8")
prcRemove.ToolTipText = localize_do("B326A9")

lstProcesses.ColumnHeaders(1).Text = localize_do("B326AB")
lstProcesses.ColumnHeaders(2).Text = localize_do("B326AA")

Dim tmpData As String
tmpData = FilterLanguage(LoadFile(def_complete_path(App.Path) + "credits.txt"), LanguageName)

txtTitles.Text = IIf(tmpData > "", tmpData, "'credits.txt' not found!")

' autostart section
chkDelay.Caption = localize_do("B328A1", "width delay (sec)")


End Sub



Private Sub Command2_Click()
Call CheckUpdate(True)
End Sub

Private Sub Command3_Click()
' SetValuesDay 1
RulerAdded = False
frmAddRuler.Show vbModal, Me

If Not RulerAdded Then Exit Sub

AddRuler DaysSelected, Tariff, StartInterval, StopInterval
LoadRulers
ShowRulers
lblAch.Visible = ScanRulers

End Sub



Private Sub Command5_Click()
On Error Resume Next
If lstTariffs.SelectedItem.index > 0 Then
    DeleteRuler lstTariffs.SelectedItem.index
    LoadRulers
    ShowRulers
    lblAch.Visible = ScanRulers
End If

End Sub

Sub LoadConfig()
On Error Resume Next

Dim u As Long

' /// UPDATE PICTURES MENU
picMenu.Line (0, 0)-(0, picMenu.Height - 15), vbButtonShadow
picMenu.Line (0, 0)-(picMenu.Width - 15, 0), vbButtonShadow
picMenu.Line (picMenu.Width - 15, picMenu.Height - 15)-(picMenu.Width - 15, 0), vbButtonShadow
picMenu.Line (picMenu.Width - 15, picMenu.Height - 15)-(0, picMenu.Height - 15), vbButtonShadow

' mnuTPIC.Line (0, 0)-(0, mnuTPIC.Height - 15), vbButtonShadow
' mnuTPIC.Line (0, 0)-(mnuTPIC.Width - 15, 0), vbButtonShadow
' mnuTPIC.Line (mnuTPIC.Width - 15, mnuTPIC.Height - 15)-(mnuTPIC.Width - 15, 0), vbButtonShadow
' mnuTPIC.Line (mnuTPIC.Width - 15, mnuTPIC.Height - 15)-(0, mnuTPIC.Height - 15), vbButtonShadow


ReloadLists
ListProcesses

cmbLine.Clear
cmbLine.AddItem "90"
cmbLine.AddItem "80"
cmbLine.AddItem "70"
cmbLine.AddItem "60"
cmbLine.AddItem "50"
cmbLine.AddItem "40"
cmbLine.AddItem "30"
cmbLine.AddItem "20"
cmbLine.AddItem "10"
cmbLine.AddItem "5"


Call LoadConfig_Main
Call LoadConfig_Tariffs

' /// LOADING OPTIONS FOR FRAME 2 ///
' >> LIMITER
chklimuse.Value = Val(-LimUse)
txtLimit.Text = Format(DataLimit, "0")
optLimit(DataLimitMode).Value = True
cmbWay.ListIndex = DataLimitWay
cmbDividers.ListIndex = DataLimitDivide

' >> NOTIFY
chkTray.Value = Val(-TipLimit)
cmbLine.Text = Format(LimitLine, "0")
chkPreLimit.Value = Val(-TipPreLimit)
chkSound.Value = Val(-TipSound)
chkShowTop.Value = Val(-ShowTop)



' /// VERSION HISTORY
txtVersion.Text = IIf(LoadFile(def_complete_path(App.Path) + "vhist.txt") > "", LoadFile(def_complete_path(App.Path) + "vhist.txt"), "Файл истории не найден")


' //
txtStaticTariff.Text = FormatEx(StaticTariff, "### ### ##0.00#")

cmbCurrency.Text = TaxName

cmbDirect.ListIndex = tax_taxing_traffic
cT1024.Value = Val(-Use1024)

txtCredits.Text = FormatEx(CreditLeft, "### ### ##0.##")
txtCredits.Tag = ""

lblAch.Visible = ScanRulers

SkinMe


ShowRulers
chkBalonOK.Value = Val(-notify_tax_change)

If FloatNotify = True Then chkNotify.Value = 1 Else chkNotify.Value = 0

prcEnable.Value = Val(-UseAutostart)
prcStop.Value = Val(-UseAutostop)
prcLinkDown.Value = Val(-UseLinkDown)
chkStartNotify.Value = Val(-UseAutoNotify)
chkDelay.Value = Val(-use_auto_delay)
txtDelay.Text = CStr(use_auto_value)

Dim vld As Label
Dim accel As Integer
accel = 735

For Each vld In mnuITM
    vld.Top = accel
    mnuICO(vld.index).Top = accel
    accel = accel + 540
Next vld

End Sub

Sub LoadConfig_Main()
    
    ' /// LOADING OPTIONS FOR FRAME 0 ///
    ' >>INTERFACE
    Call RefreshInterface                                ' Update network interfaces
    chkPing.Value = Val(-use_ping_host)               ' Update ping option
    pingAddr.Text = PingManual
    pingSel(PingMode).Value = True

    ' >>OPTIONS
    chkAutorun.Value = def_bool_to_int(CheckAutorun)        ' Update autorun setting
    chkUpdate.Value = Val(-EveryDayCheck)           ' Update updating option
    chkAttach.Value = Val(-MainWindowAttach)
    chkTransparent.Value = Val(-MainWindowAlpha)
    sliAlpha.Value = MainWindowAlphaLVL
    sliAlpha.Enabled = MainWindowAlpha
  
    ' >>LANGUAGE
    cmbFiles.Clear
    File1.Path = App.Path
    File1.Refresh
    Dim tmpIndex As Integer
    For tmpIndex = 0 To File1.ListCount - 1
        cmbFiles.AddItem Left(File1.List(tmpIndex), Len(File1.List(tmpIndex)) - 4)
    Next tmpIndex
    cmbFiles.Text = LanguageName

End Sub

Sub LoadConfig_Tariffs()

    ' /// LOADING OPTIONS FOR FRAME 1 ///
    ' >> WINDOW
    chkFloatWindow.Value = Val(-FloatWindow)
    chkSafe.Value = Val(-FWrmr)
    chkOnTop.Value = Val(-FWOnTop)
    chkConOnly.Value = Val(-FWOnline)
    chkAlwaysVisible.Value = Val(-FWAlwaysVisible)

    ' >> TRAFFIC

    Combo1.ListIndex = Val(Mid(ComboLinks, 1, 1))
    Combo2.ListIndex = Val(Mid(ComboLinks, 2, 1))
    Combo3.ListIndex = Val(Mid(ComboLinks, 3, 1))


End Sub

Sub ReloadLists()
' /// UPDATE LISTS
cmbWay.Clear
cmbWay.AddItem localize_do("WRD0016", "Входящий/Исходящий")
cmbWay.AddItem localize_do("WRD0017", "Входящий")
cmbWay.AddItem localize_do("WRD0018", "Исходящий")
    
    Combo1.Clear
    Combo1.AddItem localize_do("WND059", "Оба направления")
    Combo1.AddItem localize_do("WND060", "Входящий")
    Combo1.AddItem localize_do("WND061", "Исходящий")
    Combo1.AddItem localize_do("WND071", "Преобладающий")

    Combo2.Clear
    Combo2.AddItem localize_do("WND059", "Оба направления")
    Combo2.AddItem localize_do("WND060", "Входящий")
    Combo2.AddItem localize_do("WND061", "Исходящий")
    Combo2.AddItem localize_do("WND071", "Преобладающий")

    Combo3.Clear
    Combo3.AddItem localize_do("WND059", "Оба направления")
    Combo3.AddItem localize_do("WND060", "Входящий")
    Combo3.AddItem localize_do("WND061", "Исходящий")
    Combo3.AddItem localize_do("WND071", "Преобладающий")

cmbDividers.Clear
cmbDividers.AddItem localize_do("WRD011", "Г") + localize_do("WRD008", "Б")
cmbDividers.AddItem localize_do("WRD010", "М") + localize_do("WRD008", "Б")
cmbDividers.AddItem localize_do("WRD009", "К") + localize_do("WRD008", "Б")

cmbDirect.Clear
cmbDirect.AddItem localize_do("WND059", "Оба направления")
cmbDirect.AddItem localize_do("WND060", "Входящий")
cmbDirect.AddItem localize_do("WND061", "Исходящий")
cmbDirect.AddItem localize_do("WND071", "Преобладающий")

cmbCurrency.Clear
cmbCurrency.AddItem "грн."
cmbCurrency.AddItem "руб."
cmbCurrency.AddItem "USD"

    cbDir.Clear
    cbDir.List(0) = localize_do("B324C0", "Нет повтора")
    cbDir.List(1) = localize_do("B324C1", "Ежедневно")
    cbDir.List(2) = localize_do("B324C2", "Ежемесячно")

End Sub

Sub ReloadListsEx()
' /// UPDATE LISTS
cmbWay.List(0) = localize_do("WRD0016", "Входящий/Исходящий")
cmbWay.List(1) = localize_do("WRD0017", "Входящий")
cmbWay.List(2) = localize_do("WRD0018", "Исходящий")
    
    Combo1.List(0) = localize_do("WND059", "Оба направления")
    Combo1.List(1) = localize_do("WND060", "Входящий")
    Combo1.List(2) = localize_do("WND061", "Исходящий")
    Combo1.List(3) = localize_do("WND071", "Преобладающий")

 
    Combo2.List(0) = localize_do("WND059", "Оба направления")
    Combo2.List(1) = localize_do("WND060", "Входящий")
    Combo2.List(2) = localize_do("WND061", "Исходящий")
    Combo2.List(3) = localize_do("WND071", "Преобладающий")


    Combo3.List(0) = localize_do("WND059", "Оба направления")
    Combo3.List(1) = localize_do("WND060", "Входящий")
    Combo3.List(2) = localize_do("WND061", "Исходящий")
    Combo3.List(3) = localize_do("WND071", "Преобладающий")


cmbDividers.List(0) = localize_do("WRD011", "Г") + localize_do("WRD008", "Б")
cmbDividers.List(1) = localize_do("WRD010", "М") + localize_do("WRD008", "Б")
cmbDividers.List(2) = localize_do("WRD009", "К") + localize_do("WRD008", "Б")


cmbDirect.List(0) = localize_do("WND059", "Оба направления")
cmbDirect.List(1) = localize_do("WND060", "Входящий")
cmbDirect.List(2) = localize_do("WND061", "Исходящий")
cmbDirect.List(3) = localize_do("WND071", "Преобладающий")

    cbDir.List(0) = localize_do("B324C0", "Нет повтора")
    cbDir.List(1) = localize_do("B324C1", "Ежедневно")
    cbDir.List(2) = localize_do("B324C2", "Ежемесячно")
    
End Sub




Private Sub cT1024_Click()
sC
End Sub

Private Sub Form_Load()
If NoExit Then Unload Me: Exit Sub

Set m_objIpHelper = New CIpHelper
Me.Height = 7700
LoadConfig
InitSettings
DoEvents
btnApply.Enabled = False

End Sub

Sub LogBonuses(inMode As Integer, inPeriod As abPeriod, iI As Currency, iO As Currency, iIO As Currency, iPay As Currency, li As Currency, lO As Currency, lIO As Currency)
    Dim uu As String
    uu = localize_do("B324B0", "Установлено:") & " "
    
    
    Select Case inMode
    Case 0
        uu = uu & FormatEx(iIO / 1000000, "### ##0.##") & vb32 & localize_do("B324B1", "Мбайт смешаного") & vb32 & localize_do("B324B4", "трафика")
    Case 1
        uu = uu & FormatEx(iI / 1000000, "### ##0.##") & vb32 & localize_do("B324B2", "МБайт входящего") & " & " & FormatEx(iO / 1000000, "### ##0.##") & vb32 & localize_do("B324B3", "МБайт исходящего") & vb32 & localize_do("B324B4", "трафика")
    End Select
    
    uu = uu & IIf(iPay > 0, vb32 & localize_do("B324B5", "за") & vb32 & Format(iPay, "0") & vb32 & TaxName, "") & vbCrLf
    
    If lO Or li Or lIO Then
        uu = uu & localize_do("B324D0", "Статус: активен") & vbCrLf
    Else
        uu = uu & localize_do("B324D1", "Статус: трафик исчерпан") & vbCrLf
    End If
    
    uu = uu & localize_do("B324B6", "Повтор:") & vb32 & AbonPeriod(inPeriod) & vbCrLf
    
    uu = uu & localize_do("B324B8", "Остаток:") & vb32
    
    Select Case inMode
    Case 0
        uu = uu & FormatEx(lIO / 1000000, "### ##0.##") & vb32 & localize_do("B324B1", "Мбайт смешаного") & vb32 & localize_do("B324B4", "трафика")
    Case 1
        uu = uu & FormatEx(li / 1000000, "### ##0.##") & vb32 & localize_do("B324B2", "МБайт входящего") & " & " & FormatEx(lO / 1000000, "### ##0.##") & vb32 & localize_do("B324B3", "МБайт исходящего") & vb32 & localize_do("B324B4", "трафика")
    End Select
    
    txtBonusSetted.Text = uu
    
End Sub

Sub ReloadBonuses()

On Error Resume Next

If DataBonusEnabled Then
    Call LogBonuses(DataBonusMode, Abonetic.aPeriod, Abonetic.aI, Abonetic.aO, _
        Abonetic.aIO, Abonetic.aMoney, DataBonusRcved, DataBonusXmited, DataBonusBoth)
Else
    txtBonusSetted.Text = localize_do("INIP5S1", "Отключен")
End If

txtXmitBonus.Text = ""
txtRcvBonus.Text = ""
txtBothBonus.Text = ""
txtMinus.Text = ""
cbDir.ListIndex = 0
chkBonus.Value = -CInt(DataBonusEnabled)
optMode(DataBonusMode).Value = True


Enablence

End Sub

Sub InitSettings()
'' init settings
mnuITM_Click 0
End Sub

Sub RefreshInterface()
On Error Resume Next


cmbInterfaces.ComboItems.Clear
cmbInterfaces.ComboItems.Add , ALL_OFF, localize_do(ALL_OFF, ALL_OFF), 3
cmbInterfaces.ComboItems.Add , ALL_INT, localize_do(ALL_INT, ALL_INT), 4
cmbInterfaces.ComboItems.Add , RAS_INT, localize_do(RAS_INT, RAS_INT), 1
cmbInterfaces.ComboItems.Add , LAN_INT, localize_do(LAN_INT, LAN_INT), 2

Dim tmpX As Integer
For tmpX = 1 To m_objIpHelper.Interfaces.Count
  If m_objIpHelper.Interfaces(tmpX).InterfaceType = 6 Then cmbInterfaces.ComboItems.Add , WithoutNull(m_objIpHelper.Interfaces(tmpX).InterfaceDescription), filter_interface_name(WithoutNull(m_objIpHelper.Interfaces(tmpX).InterfaceDescription)), 2
Next tmpX

Dim Ich As Integer
Ich = GetIndexFrom(iph_interface)

If Ich > -1 Then cmbInterfaces.ComboItems.Item(Ich).Selected = True Else cmbInterfaces.ComboItems.Item(0).Selected = True


End Sub


Private Sub lblEMail_Click()

RunWEB "mailto:networkmeter@ukr.net"

End Sub

Private Sub lblSiteVisit_Click()

RunWEB "http://woobind.org.ua"

End Sub





Private Sub lstProcesses_AfterLabelEdit(Cancel As Integer, NewString As String)

    ApplicationsChanged = True
    sC
    
End Sub

Private Sub lstProcesses_DblClick()

    Dim inZ As String
    On Error Resume Next
    
    If IsNothing(lstProcesses.SelectedItem) = False Then
        
        inZ = InputBox(localize_do("B326C0", "Строка запуска:"), , lstProcesses.SelectedItem.SubItems(1))
        If inZ <> "" Then
            lstProcesses.SelectedItem.SubItems(1) = inZ
            ApplicationsChanged = True
            sC
        End If
    
    End If
    
End Sub

Private Sub lstProcesses_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        If IsNothing(lstProcesses.SelectedItem) = False Then
            lstProcesses.StartLabelEdit
        End If
    End If

End Sub

Private Sub mnuICO_Click(index As Integer)
mnuSEL.Top = mnuITM(index).Top - 130
mnuSEL.Left = 30
mnuSEL.Width = picMenu.Width - 60
mnuTCAP.Caption = mnuITM(index).Caption
mnuTICO.Picture = mnuICO(index).Picture

End Sub

Private Sub mnuICO_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If mnuGHOST.Top <> mnuITM(index).Top - 130 Then mnuGHOST.Top = mnuITM(index).Top - 130
If mnuGHOST.Left <> 45 Then mnuGHOST.Left = 45
If mnuGHOST.Width <> picMenu.Width - 90 Then mnuGHOST.Width = picMenu.Width - 90
If mnuGHOST.Visible = False Then mnuGHOST.Visible = True
End Sub

Sub mnuITM_Click(index As Integer)
mnuSEL.Top = mnuITM(index).Top - 130
mnuSEL.Left = 45
mnuSEL.Width = picMenu.Width - 90
mnuTCAP.Caption = mnuITM(index).Caption
'mnuTICO.Picture = mnuICO(index).Picture

Dim undx, SIndex
For undx = 0 To opFrame.Count - 1
  opFrame(undx).Visible = False
Next undx

SIndex = index

opFrame(SIndex).Left = 2940
opFrame(SIndex).Top = 1140 ' 1140
opFrame(SIndex).Visible = True

If index = 4 Then ReloadBonuses


End Sub

Private Sub mnuITM_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If mnuGHOST.Top <> mnuITM(index).Top - 130 Then mnuGHOST.Top = mnuITM(index).Top - 130
If mnuGHOST.Left <> 45 Then mnuGHOST.Left = 45
If mnuGHOST.Width <> picMenu.Width - 90 Then mnuGHOST.Width = picMenu.Width - 90
If mnuGHOST.Visible = False Then mnuGHOST.Visible = True

End Sub


Private Sub mnuTCAP_Change()
    
    lREP.Caption = mnuTCAP.Caption

End Sub


Private Sub OKButton_Click()

If btnApply.Enabled Then SaveSettings
Unload Me

End Sub


Sub SaveSettings()
On Error Resume Next

If NoExit Then Exit Sub


If ApplicationsChanged Then
    Call UpdateProcess
    ApplicationsChanged = False
Else
    Call ListProcesses
End If

UseAutostart = def_any_to_bool(prcEnable.Value)
UseAutostop = def_any_to_bool(prcStop.Value)
UseLinkDown = def_any_to_bool(prcLinkDown.Value)
UseAutoNotify = def_any_to_bool(chkStartNotify.Value)
use_auto_delay = def_any_to_bool(chkDelay.Value)
use_auto_value = Format(txtDelay.Text, "0")

' // SORTED
' // FRAME 0
' >> INTERFACE
use_ping_host = def_any_to_bool(chkPing.Value)
If pingSel(0).Value = True Then PingMode = 0 Else PingMode = 1
PingManual = pingAddr.Text
If Not cmbInterfaces.Text = "" Then
   iph_interface = cmbInterfaces.SelectedItem.Key
End If

' >> OPTIONS
EveryDayCheck = def_any_to_bool(chkUpdate.Value)
MainWindowAttach = def_any_to_bool(chkAttach.Value)
MainWindowAlpha = def_any_to_bool(chkTransparent.Value)
MainWindowAlphaLVL = sliAlpha.Value

' >> LANGUAGE
Let LanguageName = cmbFiles.Text
CacheStrings def_complete_path(App.Path) & LanguageName & ".slf"

' // FRAME 1 //
' >> WINDOW
FloatWindow = def_any_to_bool(chkFloatWindow.Value)
FWrmr = def_any_to_bool(chkSafe.Value)
FWOnTop = def_any_to_bool(chkOnTop.Value)
FWOnline = def_any_to_bool(chkConOnly.Value)
FWAlwaysVisible = def_any_to_bool(chkAlwaysVisible.Value)

' >> TRAFFIC
Mid(ComboLinks, 1, 1) = Format(Combo1.ListIndex, "0")
Mid(ComboLinks, 2, 1) = Format(Combo2.ListIndex, "0")
Mid(ComboLinks, 3, 1) = Format(Combo3.ListIndex, "0")


StaticTariff = Val(txtStaticTariff.Text)

' ////


If txtCredits.Tag = "changed" Then
    CreditLeft = CCur(Val(txtCredits.Text))
End If

tax_taxing_traffic = cmbDirect.ListIndex
TaxName = cmbCurrency.Text


LimitLine = Val(cmbLine.Text)

DataLimit = Val(txtLimit.Text)
DataLimitDivide = cmbDividers.ListIndex
DataLimitWay = cmbWay.ListIndex


Dim TT As Integer: For TT = 0 To 3
  If optLimit(TT).Value = True Then DataLimitMode = TT: Exit For
Next TT



TipLimit = def_any_to_bool(chkTray.Value)
TipPreLimit = def_any_to_bool(chkPreLimit.Value)
TipSound = def_any_to_bool(chkSound.Value)
ShowTop = def_any_to_bool(chkShowTop.Value)
Use1024 = def_any_to_bool(cT1024.Value)


LimUse = def_any_to_bool(chklimuse.Value)

Call frmOK.InitOkForm

FloatNotify = def_any_to_bool(chkNotify.Value)
notify_tax_change = def_any_to_bool(chkBalonOK.Value)


frmVelton.RefreshPing
frmVelton.SaveToINI
' Call CleanLangCache
frmVelton.RefreshFace

LoadRulers
WasLimit = 0
iph_reset_traffic_flag = True

End Sub

Private Sub optLimit_Click(index As Integer)
sC
End Sub

Private Sub optMode_Click(index As Integer)
EnableBonusCtrls
End Sub

Private Sub pingAddr_Change()
pingSel(1).Value = True
sC
End Sub


Private Sub pingSel_Click(index As Integer)
sC
End Sub

Private Sub prcAdd_Click()

    On Error Resume Next

    Dim cdlg As New CommonDlg
    
    With cdlg
        .DefaultExt = "exe"
        .CancelError = True
        .hWndOwner = Me.hWnd
        .DialogTitle = "*.exe"
        .Filter = "Executable files|*.exe"
        .ShowOpen
        If Err Then Err.Clear: Exit Sub
    End With
    
    If FileExists(cdlg.FileName) Then
        AddProcess cdlg.FileName
        ApplicationsChanged = True
        sC
    End If

    
End Sub

Private Sub prcEnable_Click()

sC

Call CtrEnable(def_any_to_bool(prcEnable.Value), cmdMoveUp, cmdMoveDown, chkStartNotify, _
                                        prcRemove, prcAdd, lstProcesses, _
                                        prcLinkDown, prcStop, chkDelay)
Call CtrEnable(def_any_to_bool(chkDelay.Value) And def_any_to_bool(prcEnable.Value), txtDelay)

End Sub


Sub CtrEnable(inEnabled As Boolean, ParamArray inControls() As Variant)


    Dim z As Variant
    
    For Each z In inControls
        z.Enabled = inEnabled
    Next z
    
End Sub


Private Sub prcLinkDown_Click()
  sC
End Sub

Private Sub prcRemove_Click()
    
    On Error Resume Next
    If IsNothing(lstProcesses.SelectedItem) = False Then
        lstProcesses.ListItems.Remove (lstProcesses.SelectedItem.index)
        If Not Err Then sC: ApplicationsChanged = True
    End If
    
End Sub

Private Sub prcStop_Click()
  sC
End Sub

Private Sub sliAlpha_Change()
lbAlphaLevel.Caption = Format(sliAlpha.Value, "0") & " %"
sC
End Sub

Private Sub sliAlpha_Scroll()
sliAlpha_Change
End Sub

Private Sub Text1_Change()
Replace txtCredits.Text, ",", "."
sC
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "," Then KeyAscii = Asc(".")
If Chr(KeyAscii) = "0" Then Exit Sub
If Chr(KeyAscii) = "1" Then Exit Sub
If Chr(KeyAscii) = "2" Then Exit Sub
If Chr(KeyAscii) = "3" Then Exit Sub
If Chr(KeyAscii) = "4" Then Exit Sub
If Chr(KeyAscii) = "5" Then Exit Sub
If Chr(KeyAscii) = "6" Then Exit Sub
If Chr(KeyAscii) = "7" Then Exit Sub
If Chr(KeyAscii) = "8" Then Exit Sub
If Chr(KeyAscii) = "9" Then Exit Sub
If Chr(KeyAscii) = "." And InStr(1, txtCredits.Text, ".") = 0 Then Exit Sub
If Chr(KeyAscii) = vbBack Then Exit Sub

KeyAscii = 0
End Sub

Private Sub tmrVisible_Timer()
If IsMouseInWindow(picMenu.hWnd) = False Then If mnuGHOST.Visible = True Then mnuGHOST.Visible = False

End Sub

Private Sub txtBothBonus_Change()
Enablence

End Sub

Private Sub txtBothBonus_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "0" Then Exit Sub
If Chr(KeyAscii) = "1" Then Exit Sub
If Chr(KeyAscii) = "2" Then Exit Sub
If Chr(KeyAscii) = "3" Then Exit Sub
If Chr(KeyAscii) = "4" Then Exit Sub
If Chr(KeyAscii) = "5" Then Exit Sub
If Chr(KeyAscii) = "6" Then Exit Sub
If Chr(KeyAscii) = "7" Then Exit Sub
If Chr(KeyAscii) = "8" Then Exit Sub
If Chr(KeyAscii) = "9" Then Exit Sub
If Chr(KeyAscii) = vbBack Then Exit Sub

KeyAscii = 0

End Sub

Private Sub txtCredits_Change()
Replace txtCredits.Text, ",", "."
txtCredits.Tag = "changed"
sC
End Sub

Private Sub txtCredits_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "," Then KeyAscii = Asc(".")
If Chr(KeyAscii) = "0" Then Exit Sub
If Chr(KeyAscii) = "1" Then Exit Sub
If Chr(KeyAscii) = "2" Then Exit Sub
If Chr(KeyAscii) = "3" Then Exit Sub
If Chr(KeyAscii) = "4" Then Exit Sub
If Chr(KeyAscii) = "5" Then Exit Sub
If Chr(KeyAscii) = "6" Then Exit Sub
If Chr(KeyAscii) = "7" Then Exit Sub
If Chr(KeyAscii) = "8" Then Exit Sub
If Chr(KeyAscii) = "9" Then Exit Sub
If Chr(KeyAscii) = "." And InStr(1, txtCredits.Text, ".") = 0 Then Exit Sub
If Chr(KeyAscii) = "-" And InStr(1, txtCredits.Text, "-") = 0 Then Exit Sub
If Chr(KeyAscii) = vbBack Then Exit Sub

KeyAscii = 0
End Sub

Private Sub txtDelay_Change()
  sC
End Sub

Private Sub txtLimit_Change()

sC

If Val(txtLimit.Text) <= 0 Then txtLimit.Text = "0"

  
End Sub

Private Sub txtLimit_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) = "0" Then Exit Sub
If Chr(KeyAscii) = "1" Then Exit Sub
If Chr(KeyAscii) = "2" Then Exit Sub
If Chr(KeyAscii) = "3" Then Exit Sub
If Chr(KeyAscii) = "4" Then Exit Sub
If Chr(KeyAscii) = "5" Then Exit Sub
If Chr(KeyAscii) = "6" Then Exit Sub
If Chr(KeyAscii) = "7" Then Exit Sub
If Chr(KeyAscii) = "8" Then Exit Sub
If Chr(KeyAscii) = "9" Then Exit Sub
If Chr(KeyAscii) = vbBack Then Exit Sub


KeyAscii = 0

End Sub




Private Sub txtMinus_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "0" Then Exit Sub
If Chr(KeyAscii) = "1" Then Exit Sub
If Chr(KeyAscii) = "2" Then Exit Sub
If Chr(KeyAscii) = "3" Then Exit Sub
If Chr(KeyAscii) = "4" Then Exit Sub
If Chr(KeyAscii) = "5" Then Exit Sub
If Chr(KeyAscii) = "6" Then Exit Sub
If Chr(KeyAscii) = "7" Then Exit Sub
If Chr(KeyAscii) = "8" Then Exit Sub
If Chr(KeyAscii) = "9" Then Exit Sub
If Chr(KeyAscii) = vbBack Then Exit Sub

KeyAscii = 0

End Sub

Private Sub txtRcvBonus_Change()
Enablence

End Sub

Private Sub txtRcvBonus_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "0" Then Exit Sub
If Chr(KeyAscii) = "1" Then Exit Sub
If Chr(KeyAscii) = "2" Then Exit Sub
If Chr(KeyAscii) = "3" Then Exit Sub
If Chr(KeyAscii) = "4" Then Exit Sub
If Chr(KeyAscii) = "5" Then Exit Sub
If Chr(KeyAscii) = "6" Then Exit Sub
If Chr(KeyAscii) = "7" Then Exit Sub
If Chr(KeyAscii) = "8" Then Exit Sub
If Chr(KeyAscii) = "9" Then Exit Sub
If Chr(KeyAscii) = vbBack Then Exit Sub

KeyAscii = 0

End Sub

Private Sub txtStaticTariff_Change()
Replace txtStaticTariff.Text, ",", "."
sC
End Sub

Private Sub txtStaticTariff_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "," Then KeyAscii = Asc(".")
If Chr(KeyAscii) = "0" Then Exit Sub
If Chr(KeyAscii) = "1" Then Exit Sub
If Chr(KeyAscii) = "2" Then Exit Sub
If Chr(KeyAscii) = "3" Then Exit Sub
If Chr(KeyAscii) = "4" Then Exit Sub
If Chr(KeyAscii) = "5" Then Exit Sub
If Chr(KeyAscii) = "6" Then Exit Sub
If Chr(KeyAscii) = "7" Then Exit Sub
If Chr(KeyAscii) = "8" Then Exit Sub
If Chr(KeyAscii) = "9" Then Exit Sub
If Chr(KeyAscii) = "." And InStr(1, txtStaticTariff.Text, ".") = 0 Then Exit Sub
If Chr(KeyAscii) = vbBack Then Exit Sub

KeyAscii = 0
End Sub

Private Sub txtXmitBonus_Change()
Enablence
End Sub

Sub Enablence()

If txtXmitBonus.Text > "" Or txtRcvBonus.Text > "" Or txtBothBonus.Text > "" Or (Val(chkBonus.Value) = 0 And DataBonusEnabled = True) Then
  cmdEnableBonus.Enabled = True
Else
  cmdEnableBonus.Enabled = False
End If

End Sub


Private Sub txtXmitBonus_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "0" Then Exit Sub
If Chr(KeyAscii) = "1" Then Exit Sub
If Chr(KeyAscii) = "2" Then Exit Sub
If Chr(KeyAscii) = "3" Then Exit Sub
If Chr(KeyAscii) = "4" Then Exit Sub
If Chr(KeyAscii) = "5" Then Exit Sub
If Chr(KeyAscii) = "6" Then Exit Sub
If Chr(KeyAscii) = "7" Then Exit Sub
If Chr(KeyAscii) = "8" Then Exit Sub
If Chr(KeyAscii) = "9" Then Exit Sub
If Chr(KeyAscii) = vbBack Then Exit Sub

KeyAscii = 0
End Sub


Sub ListProcesses()
    
    On Error Resume Next
    Dim imax As Integer: Let imax = UBound(lRecord)
    Dim index As Integer
    Dim li As ComctlLib.ListItem
    
    lstProcesses.ListItems.Clear
    
    For index = 1 To imax
        Set li = lstProcesses.ListItems.Add()
        li.Text = Trim(lRecord(index).lR_Name)
        li.SubItems(1) = Trim(lRecord(index).lR_Path)
        li.Tag = lRecord(index).lR_PID
    Next index
    
    
End Sub

Sub AddProcess(pPath As String)
    
    ' On Error Resume Next
    Dim index As Integer
    Dim li As ComctlLib.ListItem
    
        Set li = lstProcesses.ListItems.Add()
        li.Text = Mid(def_get_file(pPath), 1, Len(def_get_file(pPath)) - 4)
        li.SubItems(1) = pPath
        li.Tag = 0
    

    
End Sub


Sub UpdateProcess()

    On Error Resume Next
    Dim index As Integer
    Dim li As ComctlLib.ListItem
    
    ReDim lRecord(1 To lstProcesses.ListItems.Count) As lRecord
    
    For index = 1 To lstProcesses.ListItems.Count
        lRecord(index).lR_Path = lstProcesses.ListItems(index).SubItems(1)
        lRecord(index).lR_Name = lstProcesses.ListItems(index).Text
        lRecord(index).lR_PID = lstProcesses.ListItems(index).Tag
    Next index
    
    SavelRecords
    
End Sub
