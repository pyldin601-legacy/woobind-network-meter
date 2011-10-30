Attribute VB_Name = "Special_Functions"
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

Enum abPeriod
    apNone = 0
    apDaily = 1
    apMonthly = 2
End Enum

Type ConnectionListEx
    CL_Name As String
    CL_Mask As String
End Type

Type Abonetic
    aMoney As Currency
    aPeriod As abPeriod
    aEnabled As Boolean
    aIO As Currency
    aI As Currency
    aO As Currency
    aLastSetted As Integer
End Type

Global Abonetic As Abonetic

Global ConnectionList(1 To 128) As ConnectionListEx


Public Enum LimitStatus
  None = 0
  Overload = 1
  Redline = 2
  OK = 3
End Enum

Global Freq2                            As Currency
Global Freq1                            As Currency

Public Const vb32 As String = " "

Const MIB_IF_TYPE_OTHER = 1
Const MIB_IF_TYPE_ETHERNET = 6
Const MIB_IF_TYPE_TOKENRING = 9
Const MIB_IF_TYPE_FDDI = 15
Const MIB_IF_TYPE_PPP = 23
Const MIB_IF_TYPE_LOOPBACK = 24
Const MIB_IF_TYPE_SLIP = 28

' /////////////////////////////////
' // BONUS DATA VARIABLE SECTION //
' /////////////////////////////////

Global DataBonusRcved                   As Currency
Global DataBonusXmited                  As Currency
Global DataBonusBoth                    As Currency
Global DataBonusEnabled                 As Boolean
Global DataBonusMode                    As Integer

' /////////////////////////////////
' //    INFORMATOR VARIABLES     //
' /////////////////////////////////

Public Enum NS
  Offline = 0
  Linkdown = 1
  Online = 2
End Enum

Global net_connection_status(0 To 1)            As NS


Global m_objIpHelper                    As CIpHelper

Global CurrLimit                        As Integer
Global WasLimit                         As Integer
Global StartXm                          As Currency
Global StartRc                          As Currency
Global ProgConnected                    As Boolean
Global TXTVer                           As String

' // Global settings
Global CreditLeft                       As Currency  ' How many money in credit
Global current_traffic_tax                            As Currency  ' Cost of 1 Mb traffic
Global FloatNotify                      As Boolean   ' Offline/Online notify

        Global CurrYear                As Long
        Global CurrMonth                As Long
        Global CurrDay                  As Long
        Global CurrWeek                 As Long
        Global stConnection             As Long
        
Global YoY                              As Integer

' // Global ping settings
Global host_is_alive                         As Boolean
Global OldTempPing As Boolean
Global NetworkText(2)                   As String
Global PingLong                         As Long

' // Other settings
Global TaxCount                         As Currency
Global TaxToday                         As Currency
Global TaxWeek                          As Currency
Global TaxMonth                         As Currency
Global TaxYear                          As Currency
Global TaxAll                           As Currency
Global TaxHour                          As Currency
Global TaxTax                           As Currency
Global tax_taxing_traffic                         As Integer
Global TaxName                          As String
Global notify_tax_change                       As Boolean
Global NextMonth                        As Boolean

Global DataXmited                       As Currency
Global DataRcved                        As Currency

Global DataRcvedCount                   As Currency
Global DataXmitedCount                  As Currency
Global DataXmitedToday                  As Currency
Global DataRcvedToday                   As Currency
Global DataXmitedWeek                   As Currency
Global DataRcvedWeek                    As Currency
Global DataXmitedMonth                  As Currency
Global DataRcvedMonth                   As Currency
Global DataXmitedYear                   As Currency
Global DataRcvedYear                    As Currency
Global DataXmitedAll                    As Currency
Global DataRcvedAll                     As Currency
Global DataXmitedBonus                  As Currency
Global DataRcvedBonus                   As Currency
Global DataXmitedHour                   As Currency
Global DataRcvedHour                    As Currency
Global timeActive                       As Currency
Global tmpDL As Currency, tmpUL As Currency


Global RegistededCopy                   As String * 32

Global MaxxSpeed(1 To 2)                As Currency

Global DataRcvedTemp                    As Currency
Global DataXmitedTemp                   As Currency

Global Use1024                          As Boolean


Global FRun                             As Boolean

' // Interface constants
Global Const NO_SPEED = 1073741824
Global Const ALL_OFF = "No monitoring"
Global Const RAS_INT = "Only DialUp connections"
Global Const LAN_INT = "Only LAN connections"
Global Const ALL_INT = "All connections"
Global Const ERR_INT = "Leave as Is"
Global Const PRX_INT = "Proxy Server Driver"

Global progPassword As String

Global iph_reset_traffic_flag                           As Boolean
Global LimitName(3)                     As String
Global LimitNameA(3)                    As String
Global svTime                           As Currency

Global lOption                          As Integer

Type LangTemp
  lngCODE                               As String
  lngSTRING                             As String
  lngNAME                               As String
End Type

' // SORTED OPTIONS
' // FRAME 0
' >> INTERFACES
Global use_ping_host                      As Boolean
Global PingMode                         As Integer
Global PingManual                       As String
Global iph_interface                             As String

' >> OPTIONS
Global EveryDayCheck                    As Boolean
Global MainWindowAttach                 As Boolean
Global MainWindowAlpha                  As Boolean
Global MainWindowAlphaLVL               As Integer

' >> LANGUAGE
Global LanguageName                     As String
Global LanguageTemp(1 To 128)           As LangTemp

' // FRAME 1
' >> WINDOW
Global FloatWindow                      As Boolean
Global FWAlwaysVisible                  As Boolean
Global FWOnTop                          As Boolean
Global FWrmr                            As Boolean
Global FWOnline                         As Boolean
Global ComboLinks                       As String * 3

' // FRAME 2
' >> LIMIT
Global LimUse                           As Boolean
Global DataLimit                        As Currency
Global DataLimitDivide                  As Integer
Global DataLimitMode                    As Integer
Global DataLimitWay                     As Integer

' >> NOTIFY
Global TipLimit                         As Boolean
Global LimitLine                        As Integer
Global TipPreLimit                      As Boolean
Global TipSound                         As Boolean
Global ShowTop                          As Boolean


' // GLOBAL RULER BUFFER
Global DaysSelected As String * 7
Global StartInterval As Integer
Global StopInterval As Integer
Global Tariff As Currency
Global RulerAdded As Boolean

' //FRAME 4
' //TARIFF
Global StaticTariff                     As Currency
Global NoExit As Boolean
Global CurrentHour                      As Date


' //AUTORUN
Global UseAutostart                     As Boolean
Global UseAutostop                      As Boolean
Global UseLinkDown                      As Boolean
Global UseAutoNotify                    As Boolean
Global use_auto_delay                   As Boolean
Global use_auto_value                   As Integer

Sub CleanRullerBuffer()
    DaysSelected = "00000000"
    StartInterval = "0"
    StopInterval = "0"
    Tariff = "0"
    RulerAdded = False
End Sub

Sub OpenOptions(ipot As Integer)

    On Error Resume Next
    Dialog.Show
    Dialog.mnuITM_Click ipot

End Sub



Function GetWeek(inDay As Integer, inMonth As Integer, inYear As Integer) As Long

    Dim N As Integer, M As Integer, i As Integer
    GetWeek = Val(Format(ReturnDate(inDay, inMonth, inYear), "ww", vbMonday, vbFirstFullWeek))

End Function

Function ReturnDate(inDay As Integer, inMonth As Integer, inYear As Integer) As Date
    
    ReturnDate = DateSerial(inYear, inMonth, inDay)
    
End Function


Function GetSettingFake(inAPP As String, inSection As String, inOPTION As String, inDefault)

    Dim IRP As String, iVal As String
    IRP = def_complete_path(App.Path) + "app_data.ini"

    iVal = String(256, 32)
    GetPrivateProfileString inSection, inOPTION, inDefault, iVal, Len(iVal), IRP
    GetSettingFake = def_cut_by_zero(iVal)

End Function

Sub RegisterAutorun()

    Dim nKey As String: nKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    Dim ret&, vz As Long
       
    SetKeyValue HKEY_LOCAL_MACHINE, nKey, "wnmeter", def_complete_path(App.Path) + "wnmeter.exe", REG_SZ

End Sub


Sub UnRegisterAutorun()

    Dim nKey As String: nKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    Dim ret&, vz As Long
    DeleteValue HKEY_LOCAL_MACHINE, nKey, "wnmeter"

End Sub

Function CheckAutorun() As Boolean

    On Error GoTo errores

    Dim nKey As String: nKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    Dim ret&, vz As String
       
    vz = QueryValue(HKEY_LOCAL_MACHINE, nKey, "wnmeter")
    If LCase(vz) = LCase(def_complete_path(App.Path) + "wnmeter.exe" & vbNullChar) Then CheckAutorun = True: Exit Function

errores:
    CheckAutorun = False

End Function


Sub Main()

    On Error Resume Next
    If App.PrevInstance Then End
    InitCommonControls
    App.TaskVisible = False
    If InStr(LCase(Command), "/noexit") > 0 Then NoExit = True
    Load frmVelton

End Sub



Function iph_interface_encode(inString As String) As String

    Dim u As Integer
    Dim V As String
    Dim W As String

    For u = 1 To Len(inString)
        V = Mid(inString, u, 1)
        
        Select Case Asc(V)
        Case Is <= 32
            W = W & "$#" & Format(Asc(V), "00")
        Case Else
            W = W & V
        End Select
    Next u
    
    iph_interface_encode = W
    
End Function

Function iph_interface_decode(inString As String) As String

    Dim u As Integer
    Dim V As String
    
    V = inString
    For u = 32 To 0 Step -1
        V = Replace(V, "$#" & Format(u, "00"), Chr(u))
    Next u
        
    iph_interface_decode = V
    
End Function


Sub Fin()

    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next frm
    End
    
End Sub


Function GetTaskbarPos() As RECT

    Dim i As Long
    
    i = FindWindow("Shell_TrayWnd", vbNullString)
    Call GetWindowRect(i, GetTaskbarPos)
   

End Function


' THE END
