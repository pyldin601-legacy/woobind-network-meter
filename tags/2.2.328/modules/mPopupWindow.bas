Attribute VB_Name = "PoP"
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


Public Enum BalloonType
    btMessage = 0
    btNotify = 1
    btHighNitify = 2
    btRed = 3
End Enum

Type BalloonNotify
    bnTITLE As String
    bnTEXT As String
    bnTYPE As BalloonType
    bnSOUND As Boolean
    bnTOP As Boolean
    bnCALLBACK As Long
End Type

Global mpwInstances As Long
Global mpwInstances2 As Integer

Global mpwShows As Integer


Sub ShowPopup(pTitle As String, pText As String, pType As Long, Optional pSound As Boolean, Optional pTop As Boolean, Optional pCallback As Long = 0)

    Set NewForm = New cForm
    Load NewForm

    With NewForm
        Call .PopupForm(pTitle, Replace(pText, "\n", vbCrLf), pSound, pTop, pType, pCallback)
    End With

End Sub


