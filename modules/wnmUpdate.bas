Attribute VB_Name = "Automatic_Updater"
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

Enum ProgressState
    IsConnecting = 1
    IsDownloading = 2
    IsIdle = 3
    IsTrouble = 4
    IsCancel = 5
End Enum

Dim PState As ProgressState
Public ArrData As String
Public ArrCompleted As Boolean

Dim FileUpdate As String


Function WithoutNull(Expression As String) As String

    Dim x
    x = InStr(Expression, vbNullChar)
    Select Case x
    Case 0
        WithoutNull = Expression
    Case Is > 0
        WithoutNull = Mid(Expression, 1, x - 1)
    End Select
    
End Function

Sub CheckUpdate(isManual As Boolean)
  
    Dim tmpStr As String
    Dim AppBuild As Integer
    Dim AppVersion As String
    Dim AppDescr As String
    Dim Downloader
  
    TrayModify frmVelton.picTray, "Checking for updates...", frmVelton.imgCheck.Picture.handle
  
    ArrCompleted = False
    ArrData = ""
  
    Set Downloader = New clsKachalka
  
    tmpStr = Downloader.DownloadToString("http://woobind.org.ua/wnmeter.up0?my_build=" & Format(App.Revision, "0"))
  
    If Len(tmpStr) Then
        AppBuild = Val(Mid(tmpStr, 1, 5))
        AppVersion = Trim(Mid(tmpStr, 6, 18))
        FileUpdate = Trim(Mid(tmpStr, 24, 128))
        AppDescr = Trim(Mid(tmpStr, 152))
   
        If App.Revision < AppBuild Then
            If App.Revision + 1 < AppBuild Then FileUpdate = Replace(FileUpdate, "update", "install")
            PoP.ShowPopup "Woobind Network Meter", "Обнаружена новая версия программы!\n" & AppDescr & "\n___________\nЩелкните здесь левой кнопкой мыши чтобы начать загрузку\nили правой кнопкой чтобы закрыть это окно.", 0, TipSound, ShowTop, AddressOf UpdateProc
        Else
            If isManual Then PoP.ShowPopup "Woobind Network Meter", "У Вас самая новая версия программы.\nОбновление не требуется.", 1, TipSound, ShowTop
        End If
    Else
        If isManual Then PoP.ShowPopup "Woobind Network Meter", "Программа обновления Woobind Network Meter \nне смогла запросить информацию с сервера. \nВозмножно проблемы с сетью.", 3, TipSound, ShowTop
    End If
 
    Set Downloader = Nothing
 
End Sub


Sub UpdateProc()
    RunWEB FileUpdate
End Sub
