Attribute VB_Name = "multilangEngine"
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

Dim langCache() As String
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)


Sub CacheStrings(FileName As String)

    Dim tmpStr As String
    tmpStr = LoadFile(FileName)
    langCache = Split(tmpStr, vbCrLf)

End Sub

Function localize_do(inMASK As String, Optional inDefault As String = "???") As String

    Dim index As Integer, tmpText As String
    Dim jndex As Integer

    For index = LBound(langCache) To UBound(langCache)
        If UCase$(Mid$(langCache(index), 1, Len(inMASK & "="))) = UCase(inMASK & "=") Then
            localize_do = Mid$(langCache(index), Len(inMASK & "=") + 1)
            tmpText = langCache(index)
            For jndex = index To 1 Step -1
                langCache(jndex) = langCache(jndex - 1)
            Next jndex
            langCache(0) = tmpText
            Exit Function
        End If
    Next index

    localize_do = inDefault

End Function

Function GetLanguage2(inMASK As String, Optional inDefault As String = "???") As String

    Dim index As Integer

    For index = LBound(langCache) To UBound(langCache)
        If UCase$(Right$(langCache(index), Len("=" & inMASK))) = UCase("=" & inMASK) Then
            GetLanguage2 = Mid$(langCache(index), 1, Len(langCache(index)) - Len("=" & inMASK))
            Exit Function
        End If
    Next index

    GetLanguage2 = inDefault

End Function

Private Function LoadFile(FileName As String) As String

    Dim f As Long, B() As Byte, IC As Long
    f = FreeFile
    If CheckFile(FileName) <> 1 Then Exit Function
    
    Open FileName For Binary As f
        IC = LOF(f)
        If Not IC = 0 Then
            ReDim B(1 To IC) As Byte
            Get #f, 1, B()
            LoadFile = String(IC, " ")
            CopyMemory ByVal LoadFile, B(1), IC
        End If
    Close f
    
End Function

Private Function CheckFile(Name As String) As Integer

    Dim S As Long
    S = GetFileAttributes(Name)
    If S = -1 Then CheckFile = 0: Exit Function
    If S And &H10 Then CheckFile = 2: Exit Function
    CheckFile = 1

End Function
