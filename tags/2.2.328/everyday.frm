VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form everyday 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Просмотр отчетов"
   ClientHeight    =   7200
   ClientLeft      =   3870
   ClientTop       =   -3150
   ClientWidth     =   8625
   ClipControls    =   0   'False
   Icon            =   "everyday.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Export -> csv"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   6660
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   240
      ScaleHeight     =   5895
      ScaleWidth      =   8115
      TabIndex        =   1
      Top             =   600
      Width           =   8115
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   7875
         TabIndex        =   10
         Top             =   5220
         Width           =   7875
         Begin VB.Label m4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Начислено"
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
            Left            =   6190
            TabIndex        =   20
            Top             =   360
            Width           =   810
         End
         Begin VB.Label m3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Сумма"
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
            Left            =   4990
            TabIndex        =   19
            Top             =   360
            Width           =   465
         End
         Begin VB.Label m2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Принято"
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
            Left            =   3810
            TabIndex        =   18
            Top             =   360
            Width           =   645
         End
         Begin VB.Label m1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Передано"
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
            Left            =   2620
            TabIndex        =   17
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Выделено (среднее)"
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
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1560
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Выделено (сумма)"
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
            Left            =   120
            TabIndex        =   15
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Передано"
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
            Left            =   2610
            TabIndex        =   14
            Top             =   60
            Width           =   750
         End
         Begin VB.Label l2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Принято"
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
            Left            =   3810
            TabIndex        =   13
            Top             =   60
            Width           =   645
         End
         Begin VB.Label l3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Сумма"
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
            Left            =   4995
            TabIndex        =   12
            Top             =   60
            Width           =   465
         End
         Begin VB.Label l4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Начислено"
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
            Left            =   6195
            TabIndex        =   11
            Top             =   60
            Width           =   810
         End
      End
      Begin MSComctlLib.ListView lstShow 
         Height          =   4695
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   8281
         SortKey         =   5
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Дата"
            Object.Width           =   4411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Передано"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Принято"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Сумма"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Начислено"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Системн."
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Дата"
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
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Начислено"
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
         Left            =   6420
         TabIndex        =   5
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Передано"
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
         Left            =   2760
         TabIndex        =   4
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Принято"
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
         Left            =   3960
         TabIndex        =   3
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Сумма"
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
         Left            =   5220
         TabIndex        =   2
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7140
      TabIndex        =   0
      Top             =   6660
      Width           =   1335
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6315
      Left            =   180
      TabIndex        =   8
      Top             =   240
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11139
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ежедневный"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Еженедельный"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ежемесячный"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ежегодный"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "За последние 48 часов"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnusep4 
      Caption         =   "@"
      Begin VB.Menu mnuClose 
         Caption         =   "Закрыть окно отчетов"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuRep0 
      Caption         =   "Отчеты"
      Begin VB.Menu mnuBuffCopy 
         Caption         =   "Копировать в буфер"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Выделить все"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearSelected 
         Caption         =   "Удалить выбранные строки"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Очистить все отчеты..."
      End
   End
   Begin VB.Menu mnuExp0 
      Caption         =   "Экспорт"
      Begin VB.Menu mnuExport 
         Caption         =   "Экспортировать всю вкладку"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuExp2 
         Caption         =   "Экспортировать выбранные строки"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "everyday"
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
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_SHOWWINDOW = &H40
Dim OldX, OldY

Dim UpdateMode As Integer


Sub ListAddValue(inText As String, inXm As Currency, inRc As Currency, inTax As Currency, inPos As String, Optional inBold As Integer = 0)

 Dim N As ListItem
 Set N = lstShow.ListItems.Add(, "itm" + Str(lstShow.ListItems.Count), inText)
 
 N.ListSubItems.Add 1, , frmVelton.DataDevi(inXm), , frmVelton.DataDeviTip(inXm)
 N.ListSubItems.Add 2, , frmVelton.DataDevi(inRc), , frmVelton.DataDeviTip(inRc)
 N.ListSubItems.Add 3, , frmVelton.DataDevi(inXm + inRc), , frmVelton.DataDeviTip(inXm + inRc)
 N.ListSubItems.Add 4, , FormatEx(inTax, "0.00") + " " + TaxName, , ""
 N.ListSubItems.Add 5, , inPos, , ""

 
 If inBold = 1 Then
    N.ForeColor = vbBlue
    N.ListSubItems(1).ForeColor = vbBlue
    N.ListSubItems(2).ForeColor = vbBlue
    N.ListSubItems(3).ForeColor = vbBlue
    N.ListSubItems(4).ForeColor = vbBlue
    N.ListSubItems(5).ForeColor = vbBlue
 ElseIf inBold = 2 Then
    N.ForeColor = vbRed
    N.ListSubItems(1).ForeColor = vbRed
    N.ListSubItems(2).ForeColor = vbRed
    N.ListSubItems(3).ForeColor = vbRed
    N.ListSubItems(4).ForeColor = vbRed
    N.ListSubItems(5).ForeColor = vbRed
 End If
    
    N.ListSubItems(1).Tag = Str(inXm)
    N.ListSubItems(2).Tag = Str(inRc)
    N.ListSubItems(3).Tag = Str(inXm + inRc)
    N.ListSubItems(4).Tag = Str(inTax)
 

End Sub

Sub CountSelected()

    Dim y As Integer
    Dim tmpRC As Currency, tmpXM As Currency, tmpSM As Currency, tmpMN As Currency
    Dim sL As Integer

    For y = 1 To lstShow.ListItems.Count
        If lstShow.ListItems(y).Selected = True Then
            sL = sL + 1
            tmpXM = tmpXM + Val(lstShow.ListItems(y).ListSubItems(1).Tag)
            tmpRC = tmpRC + Val(lstShow.ListItems(y).ListSubItems(2).Tag)
            tmpSM = tmpSM + Val(lstShow.ListItems(y).ListSubItems(3).Tag)
            tmpMN = tmpMN + Val(lstShow.ListItems(y).ListSubItems(4).Tag)
        End If
    Next y

    l1.Caption = frmVelton.DataDevi(tmpXM)
    l2.Caption = frmVelton.DataDevi(tmpRC)
    l3.Caption = frmVelton.DataDevi(tmpSM)
    l4.Caption = FormatEx(tmpMN, "### ##0.00") + " " + TaxName

    l1.ToolTipText = frmVelton.DataDeviTip(tmpXM)
    l2.ToolTipText = frmVelton.DataDeviTip(tmpRC)
    l3.ToolTipText = frmVelton.DataDeviTip(tmpSM)

    If sL = 0 Then
        mnuClearSelected.Enabled = False
        mnuExp2.Enabled = False
        mnuBuffCopy.Enabled = False
    Else
        mnuClearSelected.Enabled = True
        mnuExp2.Enabled = True
        mnuBuffCopy.Enabled = True
    End If
    If sL = 0 Then sL = 1
    
    m1.Caption = frmVelton.DataDevi(tmpXM / sL)
    m2.Caption = frmVelton.DataDevi(tmpRC / sL)
    m3.Caption = frmVelton.DataDevi(tmpSM / sL)
    m4.Caption = FormatEx(tmpMN / sL, "### ##0.00") + " " + TaxName

    m1.ToolTipText = frmVelton.DataDeviTip(tmpXM / sL)
    m2.ToolTipText = frmVelton.DataDeviTip(tmpRC / sL)
    m3.ToolTipText = frmVelton.DataDeviTip(tmpSM / sL)


End Sub

Sub UpdateDayMode()


On Error Resume Next

Dim gMaxDay As Integer
gMaxDay = GetMaxPerMonth

Dim IndexCount As Long, N As Long, z As String
i = FreeFile
lstShow.ListItems.Clear

Label1.Caption = localize_do("xDAY", "День")


Open def_complete_path(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
IndexCount = LOF(i) / Len(StatType)

For N = 1 To IndexCount
 Get #i, N, StatType
 If StatType.stType = ST_MODE.ST_DAILY Then
 z = Trim(StatType.stDATE)
 If Month(Now) = Month(Trim(StatType.stMeta)) And Year(Now) = Year(Trim(StatType.stMeta)) And gMaxDay = Day(Trim(StatType.stMeta)) Then
  ListAddValue Trim(StatType.stMeta), StatType.stXm, StatType.stRc, StatType.stTax, z, 2
 ElseIf Month(Now) = Month(Trim(StatType.stMeta)) And Year(Now) = Year(Trim(StatType.stMeta)) Then
  ListAddValue Trim(StatType.stMeta), StatType.stXm, StatType.stRc, StatType.stTax, z, 1
 Else
  ListAddValue Trim(StatType.stMeta), StatType.stXm, StatType.stRc, StatType.stTax, z, 0
 End If
 End If
Next N

Close #i



End Sub


Function GetMaxPerMonth() As Integer
On Error Resume Next

Dim IndexCount As Long, N As Long, z As String, zBuff As Currency, sBuff As Integer
i = FreeFile

Open def_complete_path(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
IndexCount = LOF(i) / Len(StatType)

For N = 1 To IndexCount
 Get #i, N, StatType
 If StatType.stType = ST_MODE.ST_DAILY Then
 z = Trim(StatType.stDATE)
 If Month(Now) = Month(Trim(StatType.stMeta)) And Year(Now) = Year(Trim(StatType.stMeta)) Then
  If StatType.stRc + StatType.stXm > zBuff Then zBuff = StatType.stRc + StatType.stXm: sBuff = Day(Trim(StatType.stMeta))
 End If
 End If
Next N


Close #i

GetMaxPerMonth = sBuff

End Function

Function GetMaxPerWeeks() As String
On Error Resume Next

Dim IndexCount As Long, N As Long, z As String, zBuff As Currency, sBuff As String
i = FreeFile

Open def_complete_path(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
IndexCount = LOF(i) / Len(StatType)

For N = 1 To IndexCount
 Get #i, N, StatType
 If StatType.stType = ST_MODE.ST_WEEKLY Then
 z = Trim(StatType.stDATE)
  If StatType.stRc + StatType.stXm > zBuff Then zBuff = StatType.stRc + StatType.stXm: sBuff = z
 End If
Next N


Close #i

GetMaxPerWeeks = sBuff

End Function

Function GetMaxPerMonths() As String
On Error Resume Next

Dim IndexCount As Long, N As Long, z As String, zBuff As Currency, sBuff As String
i = FreeFile

Open def_complete_path(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
IndexCount = LOF(i) / Len(StatType)

For N = 1 To IndexCount
 Get #i, N, StatType
 If StatType.stType = ST_MODE.ST_MONTHLY Then
 z = Trim(StatType.stDATE)
  If StatType.stRc + StatType.stXm > zBuff Then zBuff = StatType.stRc + StatType.stXm: sBuff = z
 End If
Next N


Close #i

GetMaxPerMonths = sBuff

End Function

Sub UpdateConnectMode()

On Error Resume Next

Dim IndexCount As Long, N As Long, z As String
i = FreeFile
lstShow.ListItems.Clear

Label1.Caption = "Подключение"


Open def_complete_path(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
IndexCount = LOF(i) / Len(StatType)

For N = 1 To IndexCount
 Get #i, N, StatType
 If StatType.stType = ST_MODE.ST_CONNECTION Then
     z = Trim(StatType.stDATE)
     ListAddValue Trim(StatType.stMeta), StatType.stXm, StatType.stRc, StatType.stTax, z
 End If
Next N

Close #i

End Sub

Sub UpdateWeekMode()


On Error Resume Next

Dim IndexCount As Long, N As Long, z As String, mxWeek As String

mxWeek = GetMaxPerWeeks

i = FreeFile
lstShow.ListItems.Clear
Label1.Caption = localize_do("xWEEK", "Неделя")

Open def_complete_path(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
IndexCount = LOF(i) / Len(StatType)

For N = 1 To IndexCount
 Get #i, N, StatType
 If StatType.stType = ST_MODE.ST_WEEKLY Then
 z = Trim(StatType.stDATE)
    If z = mxWeek Then
        ListAddValue Trim(StatType.stMeta), StatType.stXm, StatType.stRc, StatType.stTax, z, 2
    Else
        ListAddValue Trim(StatType.stMeta), StatType.stXm, StatType.stRc, StatType.stTax, z
    End If
 End If
Next N

Close #i

End Sub

Sub UpdateMonthMode()


On Error Resume Next

Dim IndexCount As Long, N As Long, z As String, mxMonth As String

mxMonth = GetMaxPerMonths

i = FreeFile
lstShow.ListItems.Clear
Label1.Caption = localize_do("xMONTH", "Месяц")


Open def_complete_path(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
IndexCount = LOF(i) / Len(StatType)

For N = 1 To IndexCount
 Get #i, N, StatType
 If StatType.stType = ST_MODE.ST_MONTHLY Then
 z = Trim(StatType.stDATE)
    If z = mxMonth Then
        ListAddValue Trim(StatType.stMeta), StatType.stXm, StatType.stRc, StatType.stTax, z, 2
    Else
        ListAddValue Trim(StatType.stMeta), StatType.stXm, StatType.stRc, StatType.stTax, z
    End If
 End If
Next N

Close #i



End Sub
Sub UpdateHourMode()


On Error Resume Next

Dim IndexCount As Long, N As Long, z As String, mxMonth As String

i = FreeFile
lstShow.ListItems.Clear
Label1.Caption = localize_do("B324E5", "Время")


Open def_complete_path(App.Path) + "app_data.h48" For Random As #i Len = Len(lDay)

For N = 1 To 48
    Get #i, N, lDay
    If lDay.idUsed = True Then
        z = Str(lDay.ldHour)
        ListAddValue Trim(lDay.ldDate), lDay.ldXmited, lDay.ldRcved, lDay.ldTax, z
    End If
Next N

Close #i



End Sub

Sub RemoveSelected(inTitles() As String)
    
    On Error Resume Next
    
    Dim IndexCount As Long, IndexCount2 As Long, N As Long, M As Long, L As Long
    Dim Scan As Boolean
    
    i = FreeFile
    Open def_complete_path(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
    IndexCount = LOF(i) / Len(StatType)
    
    j = FreeFile
    Open def_complete_path(App.Path) + "$$$app_data$$$.rpt" For Random As #j Len = Len(StatType)
    
    For N = 1 To IndexCount
        Get i, N, StatType
        Scan = False
        For M = 1 To UBound(inTitles) - 1
            If Trim(StatType.stMeta) = inTitles(M) Then Scan = True
        Next M
        If Not Scan Then
            L = L + 1
            Put j, L, StatType
        End If
    Next N
    
    Close i, j
    
    Kill def_complete_path(App.Path) + "app_data.rpt"
    Name def_complete_path(App.Path) + "$$$app_data$$$.rpt" As def_complete_path(App.Path) + "app_data.rpt"
    
End Sub

Sub RemoveSelectedH(inTitles() As String)
    
    On Error Resume Next
    
    Dim IndexCount As Long, IndexCount2 As Long, N As Long, M As Long, L As Long
    Dim Scan As Boolean
    
    i = FreeFile
    Open def_complete_path(App.Path) + "app_data.h48" For Random As #i Len = Len(lDay)
    IndexCount = LOF(i) / Len(lDay)
    
    j = FreeFile
    Open def_complete_path(App.Path) + "$$$app_data$$$.h48" For Random As #j Len = Len(lDay)
    
    For N = 1 To IndexCount
        Get i, N, lDay
        Scan = False
        For M = 1 To UBound(inTitles) - 1
            If Trim(lDay.ldDate) = inTitles(M) Then Scan = True
        Next M
        If Not Scan Or N = 1 Then
            L = L + 1
            Put j, L, lDay
        End If
    Next N
    
    Close i, j
    
    Kill def_complete_path(App.Path) + "app_data.h48"
    Name def_complete_path(App.Path) + "$$$app_data$$$.h48" As def_complete_path(App.Path) + "app_data.h48"
    
End Sub

Function GetTitles(z() As String)

    Dim y As Integer
    Dim zCount As Long
    ReDim z(1 To 1)
    
    For y = 1 To lstShow.ListItems.Count
        If lstShow.ListItems(y).Selected = True Then
            zCount = UBound(z)
            z(zCount) = lstShow.ListItems(y).Text
            ReDim Preserve z(1 To zCount + 1) As String
        End If
    Next y


End Function

'
'
Sub UpdateYearMode()


On Error Resume Next

Dim IndexCount As Long, N As Long, z As String
i = FreeFile
lstShow.ListItems.Clear
Label1.Caption = localize_do("xYEAR", "Год")


Open def_complete_path(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
IndexCount = LOF(i) / Len(StatType)

For N = 1 To IndexCount
 Get #i, N, StatType
 If StatType.stType = ST_MODE.ST_YEARLY Then
 z = Trim(StatType.stDATE)
 ListAddValue Trim(StatType.stMeta), StatType.stXm, StatType.stRc, StatType.stTax, z
 End If
Next N

Close #i



End Sub


Private Sub Command1_Click()
On Error Resume Next

HideMe
End Sub



Private Sub Command2_Click()

    ExportStatistic

End Sub


Sub ExportStatistic()

    On Error Resume Next

    Dim cdlg As New CommonDlg
    
    With cdlg
        .DefaultExt = "csv"
        .CancelError = True
        .hWndOwner = Me.hWnd
        .DialogTitle = localize_do("xSAVE", "Экспорт")
        .Filter = localize_do("B324M4", "Значения, разделенные запятой") & " (*.csv)|*.csv"
        .FileName = "Report " & Format(Now, "dd-mm-yyyy")
        .ShowSave
        If Err Then Err.Clear:   Exit Sub
    End With
    
    If FileExists(cdlg.FileName) Then If MsgBox(localize_do("xMSG01", "Overwrite?"), vbExclamation + vbYesNo, cdlg.DialogTitle) = vbNo Then Exit Sub

    u = FreeFile
    Open cdlg.FileName For Output As #u


    Print #u, "Woobind Network Meter"
    Print #u, Me.Caption
    Print #u, ""
    Print #u, Label1.Caption & ";" & Label2.Caption & ";" & Label3.Caption & ";" & Label4.Caption & ";" & Label5.Caption

    For j = 1 To lstShow.ListItems.Count
        Print #u, lstShow.ListItems(j).Text & ";" & lstShow.ListItems(j).SubItems(1) & ";" & lstShow.ListItems(j).SubItems(2) & ";" & lstShow.ListItems(j).SubItems(3) & ";" & lstShow.ListItems(j).SubItems(4)
    Next j
    Close #u

End Sub

Sub ExportStatistic2()

    On Error Resume Next

    Dim cdlg As New CommonDlg
    
    With cdlg
        .DefaultExt = "csv"
        .CancelError = True
        .hWndOwner = Me.hWnd
        .DialogTitle = localize_do("xSAVE", "Экспорт")
        .Filter = localize_do("B324M4", "Значения, разделенные запятой") & " (*.csv)|*.csv"
        .FileName = "Report " & Format(Now, "dd-mm-yyyy")
        .ShowSave
        If Err Then Err.Clear:   Exit Sub
    End With
    
    If FileExists(cdlg.FileName) Then If MsgBox(localize_do("xMSG01", "Overwrite?"), vbExclamation + vbYesNo, cdlg.DialogTitle) = vbNo Then Exit Sub

    u = FreeFile
    Open cdlg.FileName For Output As #u


    Print #u, "Woobind Network Meter"
    Print #u, Me.Caption
    Print #u, ""
    Print #u, Label1.Caption & ";" & Label2.Caption & ";" & Label3.Caption & ";" & Label4.Caption & ";" & Label5.Caption

    For j = 1 To lstShow.ListItems.Count
        If lstShow.ListItems(j).Selected = True Then Print #u, lstShow.ListItems(j).Text & ";" & lstShow.ListItems(j).SubItems(1) & ";" & lstShow.ListItems(j).SubItems(2) & ";" & lstShow.ListItems(j).SubItems(3) & ";" & lstShow.ListItems(j).SubItems(4)
    Next j
    Close #u

End Sub

Sub ExportStatisticBuffer()

    On Error Resume Next
    Dim tmpbuff As String


    For j = 1 To lstShow.ListItems.Count
        If lstShow.ListItems(j).Selected = True Then
            tmpbuff = tmpbuff & lstShow.ListItems(j).Text & vbTab & lstShow.ListItems(j).SubItems(1) & vbTab & lstShow.ListItems(j).SubItems(2) & vbTab & lstShow.ListItems(j).SubItems(3) & vbTab & lstShow.ListItems(j).SubItems(4) & vbCrLf
        End If
    Next j

    Clipboard.Clear
    Clipboard.SetText tmpbuff
    

End Sub

Private Sub Form_Load()

    On Error Resume Next
    Call RefreshLanguage

    UpdateMode = 1
    UpdateDayMode
    RefreshX

End Sub

Sub RefreshLanguage()

    Label1.Caption = localize_do("ADD001", "Период")
    Label2.Caption = localize_do("ADD002", "Отправлено")
    Label3.Caption = localize_do("ADD003", "Получено")
    Label4.Caption = localize_do("ADD004", "Сумма")
    Label5.Caption = localize_do("ADD005", "Начислено")
    Label6.Caption = localize_do("xSELECT", "Выделено")
    Label7.Caption = localize_do("xSELECT2", "Выделено (средн.)")
    Me.Caption = localize_do("FORM04", "Отчет по трафику")

    TabStrip1.Tabs(1).Caption = localize_do("TAB001", "DAILY")
    TabStrip1.Tabs(2).Caption = localize_do("TAB002", "WEEKLY")
    TabStrip1.Tabs(3).Caption = localize_do("TAB003", "MONTHLY")
    TabStrip1.Tabs(4).Caption = localize_do("TAB004", "YEAR")
    TabStrip1.Tabs(5).Caption = localize_do("B324E0", "За последние 48 часов")

    Command1.Caption = localize_do("BTN001", "Закрыть")
    Command2.Caption = localize_do("xBUT01", "Экспорт")

    mnuClearAll.Caption = localize_do("B324M0", "Очистить все отчеты...")
    mnuClearSelected.Caption = localize_do("B324M1", "Удалить выбранные строки")
    
    mnuExport.Caption = localize_do("B324M2", "Экспортировать всю вкладку...")
    mnuExp2.Caption = localize_do("B324M6", "Экспортировать виделенные строки...")
    
    mnuClose.Caption = localize_do("B324M3", "Закрыть окно отчетов")
    
    mnuExp0.Caption = localize_do("B324M7", "Экспорт")
    mnuRep0.Caption = localize_do("B324M8", "Отчеты")
    mnuBuffCopy.Caption = localize_do("B324M9", "Копировать в буфер")
    mnuSelAll.Caption = localize_do("B324MA", "Выбрать все")

End Sub

Private Sub optDAYS_Click()
On Error Resume Next

UpdateDayMode
End Sub

Private Sub optMONTHS_Click()

On Error Resume Next
UpdateMonthMode
End Sub

Private Sub optWEEKS_Click()
On Error Resume Next
UpdateWeekMode
End Sub


Sub RefreshX()

On Error Resume Next
If UpdateMode = 1 Then
  UpdateDayMode
End If

If UpdateMode = 2 Then
  UpdateWeekMode
End If

If UpdateMode = 3 Then
  UpdateMonthMode
End If

If UpdateMode = 4 Then
  UpdateYearMode
End If

If UpdateMode = 5 Then
  UpdateHourMode
End If

UnselectItems
Call frmVelton.UpdateAverage
CountSelected

End Sub

Sub UnselectItems()
For h = 1 To lstShow.ListItems.Count
 lstShow.ListItems(h).Selected = False
Next h
End Sub

Private Sub lstShow_Click()
Call CountSelected
End Sub

Private Sub lstShow_ItemClick(ByVal Item As MSComctlLib.ListItem)
Call CountSelected
End Sub


Private Sub lstShow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 Then PopupMenu mnuRep0

End Sub

Private Sub mnuBuffCopy_Click()
    
    Call ExportStatisticBuffer
    
End Sub

Private Sub mnuClearAll_Click()
    
    On Error Resume Next
    If MsgBox(localize_do("B324M5", "Вы уверены, что хотите очистить ВСЕ отчеты по трафику?"), vbQuestion + vbYesNo) = vbYes Then
        Kill def_complete_path(App.Path) & "app_data.h48"
        Kill def_complete_path(App.Path) & "app_data.rpt"
        RefreshX
    End If

End Sub

Private Sub mnuClearSelected_Click()

    Dim tmpv() As String
    GetTitles tmpv
    If UpdateMode <> 5 Then
        RemoveSelected tmpv
    Else
        RemoveSelectedH tmpv
    End If
    RefreshX
    
End Sub

Private Sub mnuClose_Click()

    HideMe
    
End Sub

Private Sub mnuExp2_Click()

    Call ExportStatistic2
    
End Sub

Private Sub mnuExport_Click()
    
    Call ExportStatistic
    
End Sub

Private Sub mnuSelAll_Click()

    On Error Resume Next
    Dim j

    For j = 1 To lstShow.ListItems.Count
        lstShow.ListItems(j).Selected = True
    Next j
    
    lstShow.SetFocus
    
End Sub

Private Sub TabStrip1_Click()
    
    On Error Resume Next
    UpdateMode = TabStrip1.SelectedItem.index
    RefreshX
    
End Sub

Sub HideMe()
    
    Unload Me

End Sub
