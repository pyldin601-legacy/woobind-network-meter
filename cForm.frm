VERSION 5.00
Begin VB.Form cForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   181
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Image1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3720
      Picture         =   "cForm.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2820
      Top             =   180
   End
   Begin VB.Label lTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4020
      TabIndex        =   2
      Top             =   60
      Width           =   465
   End
   Begin VB.Label lText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Put some text here !"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label lTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Woobind Network Meter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   2040
   End
End
Attribute VB_Name = "cForm"
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


Dim FileToExecute As Long
Dim WithSound As Boolean
Dim IsTopOrder As Boolean
Dim UnderMouse As Boolean
Dim FatInterval As Integer
Dim FatMaximum As Integer

Dim MuPos As POINTAPI



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 And FileToExecute Then
        CallWindowProc FileToExecute, 0, 0, 0, 0
    End If

    Call HideForm(False)

End Sub


Private Sub lText_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_MouseUp(Button, Shift, x, y)
End Sub

Private Sub lTime_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_MouseUp(Button, Shift, x, y)
End Sub

Private Sub lTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_MouseUp(Button, Shift, x, y)
End Sub

Private Sub Timer1_Timer()
    If FatMaximum < FatInterval Then FatMaximum = FatInterval
    
    FatInterval = FatInterval - 1
    
    If FatInterval <= 0 Then Call HideForm(True)
    Me.Line (1, Me.ScaleHeight / FatMaximum * FatInterval)-(3, Me.ScaleHeight), RGB(200, 200, 200), BF
End Sub

Sub HideForm(inWait As Boolean)

    Timer1.Enabled = False
    

    If inWait Then
        Do
            DoEvents
        Loop Until Not IsMouseInWindow(Me.hWnd)
    End If

    If WithSound Then PlayWAVAsync (def_complete_path(App.Path) & "closed.wav")

    Dim i As Long, j As Long

    j = Me.Left

    mpwInstances2 = mpwInstances2 - 1
    If mpwInstances2 <= 0 Then mpwInstances = 0

    For i = 0 To Me.Width Step Screen.TwipsPerPixelX
        Me.Move j + i
    Next i

    Unload Me
    AlignForms IsTopOrder

End Sub

Function ShowedForms(Optional TotalHeight As Long)
    
    Dim i As Integer
    Dim j As Long
    
    For Each iForm In Forms
        If iForm.Tag = 911 Then _
            i = i + 1: _
            j = j + iForm.Height + 30
    Next
    
    ShowedForms = i
    TotalHeight = j
    
End Function

Sub PopupForm(bnTITLE As String, bnTEXT As String, bnSOUND As Boolean, bnTOP As Boolean, bnTYPE As Long, bnCALLBACK As Long)

    Dim EmColor As Long
    Dim BrColor As Integer
    Dim cRect As RECT
    Dim fRect As RECT
    Dim gHeight As Long

    cRect = GetScreen

    Let lTitle.Caption = Trim(bnTITLE)
    Let lText.Caption = Trim(bnTEXT)
    Let Me.Tag = 911

    Let FileToExecute = bnCALLBACK
    Let WithSound = bnSOUND
    Let IsTopOrder = bnTOP

    Me.Height = dsRes((lText.Height + lText.Top + 5) * Screen.TwipsPerPixelY, 4)
    Me.Width = dsRes(IIf(lText.Width > 256 - lText.Left - 10, lText.Width + lText.Left + 10, 256) * Screen.TwipsPerPixelX, 5 * Screen.TwipsPerPixelX)
    Call GetWindowRect(Me.hWnd, fRect)
    Me.Left = cRect.Right * Screen.TwipsPerPixelX

    BitBlt Me.hdc, Me.ScaleWidth - Image1.Width - 8, Me.ScaleHeight - Image1.Height - 4, 32, 32, Image1.hdc, 0, 0, vbSrcCopy

    Call ShowedForms(gHeight)

    lTime.Caption = Format(Now, "H:MM")
    lTime.Left = Me.ScaleWidth - (lTime.Width + 8)

    Select Case bnTOP
    Case True
        Me.Top = cRect.Top * Screen.TwipsPerPixelX + gHeight - Me.Height - 30
    Case False
        Me.Top = cRect.Bottom * Screen.TwipsPerPixelY - gHeight + 30
    End Select

    Me.Line (0, 0)-(Me.ScaleWidth, 0), RGBBright(vbWhite, 200)
    Me.Line (0, 0)-(0, Me.ScaleHeight), RGBBright(vbWhite, 200)
    Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight), RGBBright(vbWhite, 200)
    Me.Line (0, Me.ScaleHeight - 1)-(Me.ScaleWidth, Me.ScaleHeight - 1), RGBBright(vbWhite, 200)

    lText.ForeColor = Choose(bnTYPE + 1, vbBlack, vbBlue, RGBBright(vbGreen, 128), vbRed, vbBlack)

    Select Case bnTYPE
    Case 0
        If bnSOUND Then PlayWAVAsync (def_complete_path(App.Path) & "notify.wav")
        Me.Line (Me.ScaleWidth - 3, 0)-(Me.ScaleWidth, Me.ScaleHeight - 1), RGBBright(vbWhite, 200), BF

    Case 1
        If bnSOUND Then PlayWAVAsync (def_complete_path(App.Path) & "information.wav")
        Me.Line (Me.ScaleWidth - 3, 0)-(Me.ScaleWidth, Me.ScaleHeight - 1), vbBlue, BF

    Case 2
        If bnSOUND Then PlayWAVAsync (def_complete_path(App.Path) & "network.wav")
        Me.Line (Me.ScaleWidth - 3, 0)-(Me.ScaleWidth, Me.ScaleHeight - 1), RGBBright(vbGreen, 128), BF

    Case 3
        If bnSOUND Then PlayWAVAsync (def_complete_path(App.Path) & "limit.wav")
        Me.Line (Me.ScaleWidth - 3, 0)-(Me.ScaleWidth, Me.ScaleHeight - 1), vbRed, BF
        Me.Line (1, 1)-(3, Me.ScaleHeight), RGB(200, 200, 200), BF

    Case Else
        lText.ForeColor = vbBlack
        Me.Line (1, 1)-(3, Me.ScaleHeight), RGB(200, 200, 200), BF

    End Select

    Dim Px As Integer
    For Px = 7 To Me.ScaleWidth - 100
        Me.Line (Px, 20)-(Px, 21), RGBBright(vbWhite, 255 - 100 / Me.ScaleWidth * (Me.ScaleWidth - 100 - Px))
    Next Px

    OnTopForm Me, True
    SetFormAlphaXP Me, 230

    Dim i As Integer

    DoEvents

    For i = cRect.Right To cRect.Right - (Me.Width / Screen.TwipsPerPixelX) Step -5
        Me.Move (i - 0) * Screen.TwipsPerPixelX
    Next i


    Select Case bnTYPE
    Case 0
        FatInterval = 600
        Timer1.Enabled = True

    Case 1
        FatInterval = 40
        Timer1.Enabled = True

    Case 2
        FatInterval = 100
        Timer1.Enabled = True

    Case Else

    End Select
    


End Sub

Sub AlignForms(cTop As Boolean)
    
    Dim i As Integer
    Dim j As Long
    
    Dim CR As RECT
    Dim iForm As Form
    CR = GetScreen
    With Screen
      For Each iForm In Forms
        If Val(iForm.Tag) = 911 And cTop Then
            MoveFormEasy iForm, (CR.Top * .TwipsPerPixelY) + j
            j = j + iForm.Height + 30
        ElseIf Val(iForm.Tag) = 911 And Not cTop Then
            MoveFormEasy iForm, (CR.Bottom * .TwipsPerPixelY) - iForm.Height - j
            j = j + iForm.Height + 30
        End If
      Next
    End With
   
End Sub

Sub MoveFormEasy(ByRef inForm As Form, Top As Long)
    
    Dim L As Long
    
    L = inForm.Top
    
    For k = L To Top Step Screen.TwipsPerPixelX * IIf(Top > L, 1, -1) * 1
        inForm.Move inForm.Left, k + i
    Next k
    
    inForm.Move inForm.Left, Top
    
End Sub

