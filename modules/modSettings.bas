Attribute VB_Name = "modSettings"
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


Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpString As Any, _
ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String) As Long



Sub SaveSettingINI(Section As String, Key As String, Setting As String, FileName As String)
    
    WritePrivateProfileString Section, Key, Setting, FileName

End Sub

Function GetSettingINI(Section As String, Key As String, inDefault, FileName As String)

    Dim iVal As String
    iVal = String(256, 32)
    GetPrivateProfileString Section, Key, inDefault, iVal, Len(iVal), FileName
    GetSettingINI = Trim32(iVal)

End Function

Function ConfigFile()
    
    ConfigFile = def_complete_path(App.Path) + App.EXEName + ".cfg"

End Function

