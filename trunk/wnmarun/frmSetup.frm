VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPostinstall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Автозагрузка Woobind Network Meter"
   ClientHeight    =   5475
   ClientLeft      =   5700
   ClientTop       =   4215
   ClientWidth     =   6795
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   4740
      TabIndex        =   11
      Top             =   4980
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4980
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      TabIndex        =   4
      Top             =   4980
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   0
      Picture         =   "frmSetup.frx":599A
      ScaleHeight     =   4695
      ScaleWidth      =   6735
      TabIndex        =   0
      Top             =   0
      Width           =   6795
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Проверять наличине новой версии"
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
         Left            =   2700
         TabIndex        =   12
         Top             =   4320
         Width           =   3735
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   2940
         ScaleHeight     =   1335
         ScaleWidth      =   3195
         TabIndex        =   5
         Top             =   2100
         Width           =   3195
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Да (рекомендуется)"
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
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   60
            Value           =   -1  'True
            Width           =   2235
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Нет"
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
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   2235
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Не вносить изменений"
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
            Left            =   120
            TabIndex        =   6
            Top             =   900
            Width           =   2355
         End
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Примечание. Включить/выключить автозагрузку можно из меню программы."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2700
         TabIndex        =   3
         Top             =   3480
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Хотите, чтобы программа Woobind Network Meter запускалась каждый раз при загрузке Windows?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2700
         TabIndex        =   2
         Top             =   1440
         Width           =   3795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Woobind Network Autostart"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   4035
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Язык:"
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
      Left            =   600
      TabIndex        =   10
      Top             =   5040
      Width           =   450
   End
End
Attribute VB_Name = "frmPostinstall"
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



Private Sub Combo1_Click()
LoadLang Combo1.ListIndex
End Sub

Private Sub Command2_Click()

If Option1(0).Value = True Then RegisterAutorun
If Option1(1).Value = True Then UnRegisterAutorun


On Error Resume Next


Dim IRP As String
IRP = def_complete_path(App.Path) + "app_data.ini"

WritePrivateProfileString "Configuration", "CheckUpdate", def_bool_to_str(def_any_to_bool(Check1.Value)), IRP
WritePrivateProfileString "Configuration", "LanguageName", Combo1.List(Combo1.ListIndex), IRP

End


End Sub

Private Sub Form_Load()

Dim tmpLng As String

InitCommonControls


Combo1.AddItem "Russian", 0
Combo1.AddItem "Ukrainian", 1
Combo1.AddItem "English", 2
tmpLng = GetIniRecord("LanguageName=", def_complete_path(App.Path) & "app_data.ini", "Russian")
Select Case tmpLng
Case "Russian": Combo1.ListIndex = 0
Case "Ukrainian": Combo1.ListIndex = 1
Case "English": Combo1.ListIndex = 2
Case Else: Combo1.ListIndex = 0
End Select

LoadLang Combo1.ListIndex

On Error Resume Next

Check1.Value = def_bool_to_int(GetSettingFake("Network Meter\" + IRP, "Configuration", "CheckUpdate", False))

End Sub

Sub LoadLang(index As Integer)
Select Case index
Case 0 'RUS
    Label2.Caption = "Хотите, чтобы программа Woobind Network Meter загружалась автоматически при загрузке Windows?"
    Option1(0).Caption = "Да (рекомендуется)"
    Option1(1).Caption = "Нет"
    Option1(2).Caption = "Не вносить изменений"
    Label3.Caption = "Примечание. Включить/выключить автозагрузку можно также из меню программы."
    Label1.Caption = "Woobind Network Meter Автозапуск"
    Me.Caption = "Автозагрузка Woobind Network Meter"
    Label4.Caption = "Язык"
    Check1.Caption = "Проверять наличине новой версии"
Case 1 'UKR
    Label2.Caption = "Бажаєте, щоб програма Woobind Network Meter завантажувалась автоматично разом з Windows?"
    Option1(0).Caption = "Так (рекомендується)"
    Option1(1).Caption = "Ні"
    Option1(2).Caption = "Не вносити змін"
    Label3.Caption = "Примітка. Включити/вимкнути автозавантаження можна також із меню програми."
    Label1.Caption = "Woobind Network Meter Автозавантаження"
    Me.Caption = "Автовантаження Woobind Network Meter"
    Label4.Caption = "Мова"
    Check1.Caption = "Перевіряти наявність нової версії"
Case 2 'ENG
    Label2.Caption = "Do you want, that the program Woobind Network Meter was started automatically when Windows starts?"
    Option1(0).Caption = "Yes (recommended)"
    Option1(1).Caption = "No"
    Option1(2).Caption = "Don't make changes"
    Label3.Caption = "Note. You can enable/disable autostart from menu of the program."
    Label1.Caption = "Woobind Network Meter Autostart"
    Me.Caption = "Woobind Network Meter autostart"
    Label4.Caption = "Language"
    Check1.Caption = "Check for new version after connect"
End Select

End Sub


