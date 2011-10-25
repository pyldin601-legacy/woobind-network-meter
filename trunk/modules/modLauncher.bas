Attribute VB_Name = "Autostart_Functions"
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

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function PostThreadMessage Lib "user32.dll" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_VM_READ = &H10

Public Const WM_QUIT = &H12
Public Const WM_SYSCOMMAND = &H112
Public Const SC_CLOSE = &HF060&


Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260

Public Type lRecord
    lR_Name As String * 32
    lR_Path As String * 256
    lR_PID As Double
End Type



Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Type EnumP
    epHwnd As Long
    epPID As Double
End Type

Public lRecord() As lRecord
Dim EnumP() As EnumP
Dim lR As lRecord


Function IsProcessOur(PID As Double, pName As String) As Boolean

    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    Dim hProcess As Long, FileName As String
    
    IsProcessOur = False
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
    Do While r
        If uProcess.th32ProcessID = PID Then
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
            FileName = Space(255)
            FileName = Left(FileName, GetModuleFileNameEx(hProcess, 0, FileName, 255))
            If FileName = pName Then
                IsProcessOur = True
                CloseHandle hProcess
                CloseHandle hSnapShot
                Exit Function
            End If
        End If
        r = Process32Next(hSnapShot, uProcess)
        ' DoEvents
    Loop
    
    CloseHandle hSnapShot
    
End Function



Function AreProcessExists(pName As String) As Double

    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    Dim hProcess As Long, FileName As String
    
    AreProcessExists = 0
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
    Do While r
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, uProcess.th32ProcessID)
        FileName = Space(255)
        FileName = Left(FileName, GetModuleFileNameEx(hProcess, 0, FileName, 255))
        If LCase(FileName) = LCase(pName) Then
            AreProcessExists = uProcess.th32ProcessID
            CloseHandle hSnapShot
            CloseHandle hProcess
            Exit Function
        End If
        r = Process32Next(hSnapShot, uProcess)
    Loop
    
    CloseHandle hSnapShot
    CloseHandle hProcess
End Function
 
 
Function Unnulled(inString As String) As String

    If InStr(inString, vbNullChar) Then
        Unnulled = Mid(inString, 1, InStr(inString, vbNullChar) - 1)
    Else
        Unnulled = inString
    End If
    
End Function

Sub LoadlRecords()

    On Error Resume Next

    Dim i, index, j As Integer

    i = FreeFile

    Open LowPath(App.Path) & "launch.dat" For Random As #i Len = Len(lR)

    j = LOF(i) / Len(lRecord(0))
    If j > 0 Then ReDim lRecord(1 To j) As lRecord

    For index = 1 To j
        Get #i, index, lRecord(index)
    Next index

    Close #i
    
End Sub


Sub SavelRecords()

    On Error Resume Next

    Dim i, index, j As Integer

    i = FreeFile

    Open LowPath(App.Path) & "launch.dat" For Output As #i: Close #i
    Open LowPath(App.Path) & "launch.dat" For Random As #i Len = Len(lR)

    j = UBound(lRecord)

    For index = 1 To j
        Put #i, index, lRecord(index)
    Next index

    Close #i
    
End Sub

Public Sub LaunchlRecords()

    On Error Resume Next
    Dim index As Integer
    Dim ttl, lnc As Integer
    Dim Pex As Double
    Dim ggg As Boolean
    Dim Launched As String
    Dim Skipped As String
    
    ttl = UBound(lRecord)
    
    For index = 1 To ttl
        Pex = 0
        Pex = AreProcessExists(GetProcName(Trim(lRecord(index).lR_Path)))
        If Pex = 0 Then
            lRecord(index).lR_PID = Shell(lRecord(index).lR_Path, vbNormalFocus)
            If lRecord(index).lR_PID > 0 Then
                lnc = lnc + 1
                Launched = Launched & vbCrLf & "   » " & Trim(lRecord(index).lR_Name)
            End If
        Else
            lRecord(index).lR_PID = Pex
            Skipped = Skipped & vbCrLf & "   » " & Trim(lRecord(index).lR_Name)
            ggg = True
        End If
    Next index
    
    If UseAutoNotify Then ShowPopup "Woobind Network Meter", localize_do("B326B0", "Автозапуск приложений.\nУспешно запущено") & Chr(32) & Format(lnc, "0") & Chr(32) & localize_do("B326B1", "из") & " " & Format(ttl, "0") & " " & localize_do("B326B2", "приложений") & IIf(Len(Launched) > 0, ":" & Launched, ".") & IIf(ggg = True, vbCrLf & localize_do("B326B3", "\nНекоторые уже были запущены ранее:") & Skipped, ""), 1, TipSound, ShowTop
    
    Call SavelRecords

End Sub


Public Sub TerminatelRecords()

    On Error Resume Next
    If UseAutostop = False Then Exit Sub
    
    Dim index As Integer
    Dim ttl, lnc As Integer
    Dim hproc As Long, ecode As Long
    Dim Launched As String

    Const PROCESS_ALL_ACCESS = &H1F0FFF

    ttl = UBound(lRecord)
    
    For index = 1 To ttl
        If lRecord(index).lR_PID > 0 And IsProcessOur(lRecord(index).lR_PID, GetProcName(Trim(lRecord(index).lR_Path))) Then
            Call KillWindowByPID(lRecord(index).lR_PID)
            hproc = OpenProcess(PROCESS_ALL_ACCESS, 0, lRecord(index).lR_PID)
            Call GetExitCodeProcess(hproc, ecode)
            If TerminateProcess(hproc, ecode) <> 0 Then
                lnc = lnc + 1
                Launched = Launched & vbCrLf & "   » " & Trim(lRecord(index).lR_Name)
            End If
            CloseHandle hproc
            lRecord(index).lR_PID = 0
        End If
        
    Next index
    
    If UseAutoNotify Then ShowPopup "Woobind Network Meter", localize_do("B326B4", "Автозакрытие приложений.\nУспешно закрыто") & Chr(32) & Format(lnc, "0") & Chr(32) & localize_do("B326B1", "из") & " " & Format(ttl, "0") & " " & localize_do("B326B2", "приложений") & IIf(Len(Launched) > 0, ":" & Launched, "."), 1, TipSound, ShowTop

    Call SavelRecords

End Sub

Function KillWindowByPID(PID As Double) As Long
    ReDim EnumP(1 To 1) As EnumP
    KillWindowByPID = 0
    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
    For N = 1 To UBound(EnumP)
        If EnumP(N).epPID = PID Then
            If PostMessage(EnumP(N).epHwnd, 2, 0, 0) > 0 Then KillWindowByPID = KillWindowByPID + 1
            DoEvents
        End If
    Next N
End Function

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    Dim tr
    GetWindowThreadProcessId hWnd, tr
    EnumP(UBound(EnumP)).epHwnd = hWnd
    EnumP(UBound(EnumP)).epPID = tr
    ReDim Preserve EnumP(1 To (UBound(EnumP) + 1)) As EnumP
    EnumWindowsProc = True
End Function

Function GetProcName(inPath As String) As String

    Dim FileName As String
    FileName = inPath
    
    If InStr(FileName, "/") Then
        FileName = Mid(FileName, 1, InStr(FileName, "/") - 2)
    End If

    GetProcName = FileName

End Function
