VERSION 5.00
Begin VB.Form frmAddRuler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Добавление правила"
   ClientHeight    =   4095
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7065
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddRuler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4920
      TabIndex        =   24
      Text            =   "0"
      Top             =   2940
      Width           =   975
   End
   Begin VB.ComboBox c_tarif 
      Height          =   315
      Index           =   0
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   900
      Width           =   1095
   End
   Begin VB.ComboBox c_tarif 
      Height          =   315
      Index           =   1
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   900
      Width           =   1095
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4920
      TabIndex        =   16
      Text            =   "0"
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1800
      ScaleHeight     =   735
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   1560
      Width           =   4875
      Begin VB.CheckBox cWeekend 
         Caption         =   "Выходные"
         Height          =   195
         Left            =   2820
         TabIndex        =   15
         Top             =   420
         Width           =   1575
      End
      Begin VB.CheckBox cWork 
         Caption         =   "Weekdays"
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   420
         Width           =   1335
      End
      Begin VB.CheckBox cAllD 
         Caption         =   "All"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   420
         Width           =   675
      End
      Begin VB.CheckBox cDay 
         Caption         =   "Sun"
         Height          =   195
         Index           =   6
         Left            =   4140
         TabIndex        =   12
         Top             =   60
         Width           =   675
      End
      Begin VB.CheckBox cDay 
         Caption         =   "Sat"
         Height          =   195
         Index           =   5
         Left            =   3480
         TabIndex        =   11
         Top             =   60
         Width           =   675
      End
      Begin VB.CheckBox cDay 
         Caption         =   "Fri"
         Height          =   195
         Index           =   4
         Left            =   2820
         TabIndex        =   10
         Top             =   60
         Width           =   675
      End
      Begin VB.CheckBox cDay 
         Caption         =   "Thu"
         Height          =   195
         Index           =   3
         Left            =   2100
         TabIndex        =   9
         Top             =   60
         Width           =   675
      End
      Begin VB.CheckBox cDay 
         Caption         =   "Wed"
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   8
         Top             =   60
         Width           =   675
      End
      Begin VB.CheckBox cDay 
         Caption         =   "Tue"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   7
         Top             =   60
         Width           =   675
      End
      Begin VB.CheckBox cDay 
         Caption         =   "Mon"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   3540
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4260
      TabIndex        =   0
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "исходящего трафика"
      Height          =   195
      Left            =   1800
      TabIndex        =   26
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label lblTax2 
      Caption         =   "fish"
      Height          =   195
      Left            =   6060
      TabIndex        =   25
      Top             =   3000
      Width           =   555
   End
   Begin VB.Label lblTax 
      Caption         =   "fish"
      Height          =   195
      Left            =   6060
      TabIndex        =   23
      Top             =   2580
      Width           =   555
   End
   Begin VB.Image Image8 
      Height          =   15
      Left            =   240
      Picture         =   "frmAddRuler.frx":038A
      Stretch         =   -1  'True
      Top             =   540
      Width           =   6360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add a new rule"
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
      TabIndex        =   22
      Top             =   180
      Width           =   1830
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cтоимость 1 МБ входящего трафика"
      Height          =   195
      Left            =   1500
      TabIndex        =   21
      Top             =   2580
      Width           =   3225
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tariff"
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
      Left            =   360
      TabIndex        =   20
      Top             =   2580
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   360
      TabIndex        =   19
      Top             =   1560
      Width           =   1470
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "to"
      Height          =   195
      Left            =   3645
      TabIndex        =   4
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "from"
      Height          =   195
      Left            =   1620
      TabIndex        =   3
      Top             =   960
      Width           =   330
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
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
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   420
   End
End
Attribute VB_Name = "frmAddRuler"
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

Private Sub cAllD_Click()
    cDay(0).Value = 1
    cDay(1).Value = 1
    cDay(2).Value = 1
    cDay(3).Value = 1
    cDay(4).Value = 1
    cDay(5).Value = 1
    cDay(6).Value = 1
    cAllD.Value = 0
End Sub



Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cDay_Click(index As Integer)
Dim tmp As Integer, tmp1 As Boolean
For tmp = 0 To 6
 If cDay(tmp).Value = 1 Then tmp1 = True
Next tmp

If tmp1 = False Then OKButton.Enabled = False Else OKButton.Enabled = True

End Sub

Private Sub cWeekend_Click()
    cDay(0).Value = 0
    cDay(1).Value = 0
    cDay(2).Value = 0
    cDay(3).Value = 0
    cDay(4).Value = 0
    cDay(5).Value = 1
    cDay(6).Value = 1
    cWeekend.Value = 0
End Sub

Private Sub cWork_Click()
    cDay(0).Value = 1
    cDay(1).Value = 1
    cDay(2).Value = 1
    cDay(3).Value = 1
    cDay(4).Value = 1
    cDay(5).Value = 0
    cDay(6).Value = 0
    cWork.Value = 0
End Sub

Private Sub Form_Load()

Dim N

For N = 1 To 7
    cDay(N - 1).Caption = Format(N + 1, "ddd")
Next N

Me.Caption = Label6.Caption
Label6.Caption = localize_do("RULL1")
Label15.Caption = localize_do("RULL2")
Label2.Caption = localize_do("RULL3")
Label3.Caption = localize_do("RULL4")
Label1.Caption = localize_do("RULL5")
Label4.Caption = localize_do("RULL6")
Label5.Caption = localize_do("RULL7")
Label8.Caption = localize_do("RULL11")

cAllD.Caption = localize_do("RULL8", "Все")
cWork.Caption = localize_do("RULL9", "Будни")
cWeekend.Caption = localize_do("RULL10", "Выходные")


lblTax.Caption = TaxName
lblTax2.Caption = TaxName

OKButton.Caption = localize_do("INIB1")
CancelButton.Caption = localize_do("INIB2")

Call LoadTime


End Sub

Sub LoadTime()
 
 Dim P As Integer
 
  For P = 0 To 23
   c_tarif(0).List(P) = Format(P, "0") + ":00"
   c_tarif(1).List(P) = Format(P, "0") + ":59"
  Next P
  
  c_tarif(0).ListIndex = 0
  c_tarif(1).ListIndex = 23

End Sub

Private Sub OKButton_Click()
Dim tmp As Integer, tmp1 As Boolean
For tmp = 0 To 6
 If cDay(tmp).Value = 1 Then tmp1 = True
Next tmp

If tmp1 = False Then Exit Sub

Dim A As Integer
RulerAdded = True
For A = 0 To 6
Mid(DaysSelected, A + 1, 1) = Format(cDay(A).Value, "0")
Next A
StartInterval = c_tarif(0).ListIndex
StopInterval = c_tarif(1).ListIndex
Tariff = Val(txtValue.Text)
Unload Me
End Sub

Private Sub txtValue_Change()
Replace txtValue.Text, ",", "."
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
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
If Chr(KeyAscii) = "." And InStr(1, txtValue.Text, ".") = 0 Then Exit Sub
If Chr(KeyAscii) = vbBack Then Exit Sub

KeyAscii = vbNull
End Sub
