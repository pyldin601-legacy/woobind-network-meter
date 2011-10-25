Attribute VB_Name = "RAS_Statistics"
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


Declare Function RasEnumConnections Lib "rasapi32" Alias "RasEnumConnectionsA" (ByVal lprasconn As Long, ByVal lpcb As Long, ByVal lpcConnections As Long) As Long
Declare Function RasGetConnectionStatistics Lib "rasapi32" (ByVal hRasConn As Long, ByVal lpStatistics As Long) As Long

Type RASCONN
    dwSize As Long
    hRasConn As Long
    szEntryName(0 To 256) As Byte
    szDeviceType(0 To 16) As Byte
    szDeviceName(0 To 128) As Byte
    pad As Byte
End Type

Private Type RAS_STATS
    dwSize As Long
    dwBytesXmited As Long
    dwBytesRcved As Long
    dwFramesXmited As Long
    dwFramesRcved As Long
    dwCrcErr As Long
    dwTimeoutErr As Long
    dwAlignmentErr As Long
    dwHardwareOverrunErr As Long
    dwFramingErr As Long
    dwBufferOverrunErr As Long
    dwCompressionRatioIn As Long
    dwCompressionRatioOut As Long
    dwBps As Long
    dwConnectDuration As Long
End Type

Function GetDialUp(dwXmited As Currency, dwRcved As Currency, dwSpeed As Currency) As Boolean

    Dim conn As RASCONN
    Dim stat As RAS_STATS
    Dim y As Long, z As Long
    Dim NewVal As Long
        
    conn.dwSize = Len(conn)
    y = conn.dwSize
    
    If RasEnumConnections(VarPtr(conn), VarPtr(y), VarPtr(z)) = 0 Then
        stat.dwSize = Len(stat)
        If RasGetConnectionStatistics(conn.hRasConn, VarPtr(stat)) = 0 Then
            dwRcved = stat.dwBytesRcved
            dwXmited = stat.dwBytesXmited
            dwSpeed = stat.dwBps
            GetDialUp = True
        Else
            GetDialUp = False
        End If
    End If

End Function

