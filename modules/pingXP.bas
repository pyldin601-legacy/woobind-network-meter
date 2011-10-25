Attribute VB_Name = "pingXP"
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


Function Ping(inIP As String, outTime As Long) As Boolean

    strComputer = "."

    On Error Resume Next

    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    
    Dim A As Long, B As Long

    A = GetTickCount
    Set colPings = objWMIService.ExecQuery("Select * From Win32_PingStatus where Address = '" + inIP + "'")


    For Each objStatus In colPings
        If IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0 Then
            outTime = -1
            Ping = False
        Else
            outTime = GetTickCount - A
            Ping = True
        End If
    Next

End Function
