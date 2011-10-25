Attribute VB_Name = "Report_Logger"
'***************************************************************************
'*                          Woobind Network Meter                          *
'*                             [ OPTIMIZED ]                               *
'***************************************************************************
'*   Copyright (C) 2008 by Roman Gemini                                    *
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

Type Statistic_Type
  stDATE                    As String * 16
  stXm                      As Currency
  stRc                      As Currency
  stTax                     As Currency
  stMeta                    As String * 32
  stType                    As Byte
End Type

Public Enum ST_MODE
  ST_DAILY = 1
  ST_WEEKLY = 2
  ST_MONTHLY = 3
  ST_YEARLY = 4
  ST_CONNECTION = 5
End Enum


Type Log_Day
  ldXmited As Currency
  ldRcved As Currency
  ldTax As Currency
  ldHour As Currency
  ldDate As String * 32
  idUsed As Boolean
End Type


Global LogDay(0 To 23)      As Log_Day
Global lDay                 As Log_Day
Global StatType             As Statistic_Type


Sub SaveTodayStatistish(inXm As Currency, inRc As Currency, inTax As Currency)

    Dim IndexCount As Long, N As Long, DayNow As String, xNow As Date
    On Error Resume Next

    i = FreeFile
    xNow = Now
    DayNow = Format(xNow, "yyyy.mm.dd")

    Open LowPath(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
        IndexCount = LOF(i) / Len(StatType)
        For N = 1 To IndexCount
            Get #i, N, StatType
            If Trim(StatType.stDATE) = DayNow And StatType.stType = ST_MODE.ST_DAILY Then
                StatType.stRc = inRc
                StatType.stXm = inXm
                StatType.stTax = inTax
                StatType.stMeta = Format(xNow, "dd.mm.yyyy")
                Put #i, N, StatType
                Close #i: Exit Sub
            End If
        Next N

    StatType.stType = ST_MODE.ST_DAILY
    StatType.stDATE = DayNow
    StatType.stRc = inRc
    StatType.stXm = inXm
    StatType.stTax = inTax
    StatType.stMeta = Format(xNow, "dd.mm.yyyy")
    Put #i, IndexCount + 1, StatType
    Close #i

End Sub

Sub SaveCurrentHourStatistic(inXm As Currency, inRc As Currency, inTax As Currency, inHour As Date)

    Dim IndexCount As Long, N As Long, DayNow As String, xNow As Date
    
    On Error Resume Next
    
    i = FreeFile
    xNow = inHour

    Open LowPath(App.Path) + "app_data.h48" For Random As #i Len = Len(lDay)
        IndexCount = LOF(i) / Len(lDay)
        lDay.ldRcved = inRc
        lDay.ldXmited = inXm
        lDay.ldTax = inTax
        lDay.ldDate = Format(xNow, "dd.mm.yyyy hh:00") & "-" & Format(xNow, "hh:59")
        lDay.idUsed = True
        lDay.ldHour = Year(xNow) * 8760 + Day(xNow) * 24 + Hour(xNow)
        Put #i, 1, lDay
     Close #i

End Sub

Sub SaveNewHourStatistic(inXm As Currency, inRc As Currency, inTax As Currency, inHour As Date)

    Dim IndexCount As Long, N As Long, DayNow As String, xNow As Date

    On Error Resume Next
    i = FreeFile
    xNow = inHour

    Open LowPath(App.Path) + "app_data.h48" For Random As #i Len = Len(lDay)
     
        For N = 1 To 48 Step 1
            Get #i, N, lDay
            Put #i, N, lDay
        Next N
     
        For N = 47 To 1 Step -1
            Get #i, N, lDay
            Put #i, N + 1, lDay
        Next N
     
        lDay.ldRcved = inRc
        lDay.ldXmited = inXm
        lDay.ldTax = inTax
        lDay.ldDate = Format(xNow, "dd.mm.yyyy hh:00-hh:59")
        lDay.idUsed = True
        lDay.ldHour = Year(xNow) * 8760 + Day(xNow) * 24 + Hour(xNow)
        Put #i, 1, lDay
     
    Close #i

End Sub

'Sub SaveConnectStatistish(inXm As Currency, inRc As Currency, inTax As Currency)
'
'On Error Resume Next
'
'Dim IndexCount As Long, N As Long, DayNow As String, xNow As Date
'i = FreeFile
'
'xNow = Now
'DayNow = "Connection " & Format(stConnection, "0")
'
'
'Open LowPath(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
'IndexCount = LOF(i) / Len(StatType)
'For N = 1 To IndexCount
' Get #i, N, StatType
'
'  If Trim(StatType.stDATE) = DayNow And StatType.stType = ST_MODE.ST_CONNECTION Then
'     StatType.stRc = inRc
'     StatType.stXm = inXm
'     StatType.stTax = inTax
'     StatType.stMeta = Format(xNow, "dd.mm.yyyy hh:mm:ss")
'     Put #i, N, StatType
'     Close #i: Exit Sub
'  End If
'Next N
'
'StatType.stType = ST_MODE.ST_CONNECTION
'StatType.stDATE = DayNow
'StatType.stRc = inRc
'StatType.stXm = inXm
'StatType.stTax = inTax
'StatType.stMeta = Format(xNow, "dd.mm.yyyy hh:mm:ss")
'Put #i, IndexCount + 1, StatType
'Close #i
'
'End Sub

Sub SaveWeekStatistish(inXm As Currency, inRc As Currency, inTax As Currency)

    On Error Resume Next

    Dim IndexCount As Long, N As Long, DayNow As Long, Meta As String, xNow As Date
    Dim TxA As Integer, TxB As Integer

    i = FreeFile
    xNow = Now
    
    DayNow = Year(xNow) & GetWeek(Day(xNow), Month(xNow), Year(xNow))
    Tx = GetLongFromData(Day(xNow), Month(xNow), Year(xNow))
    TxA = Tx - (Weekday(xNow, vbMonday) - 1)
    TxB = TxA + (Weekday(xNow, vbMonday) - 1)

    tra = Format(DateSerial(Year(xNow), 1, TxA), "dd mmm yyyy")
    trb = Format(DateSerial(Year(xNow), 1, TxB), "dd mmm yyyy")

    Open LowPath(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
        IndexCount = LOF(i) / Len(StatType)
        For N = 1 To IndexCount
            Get #i, N, StatType
            If Val(StatType.stDATE) = DayNow And StatType.stType = ST_MODE.ST_WEEKLY Then
                StatType.stMeta = tra + " - " + trb
                StatType.stRc = inRc
                StatType.stXm = inXm
                StatType.stTax = inTax
                Put #i, N, StatType
                Close #i: Exit Sub
            End If
        Next N

        StatType.stType = ST_MODE.ST_WEEKLY
        StatType.stMeta = tra + " - " + trb
        StatType.stDATE = Str(DayNow)
        StatType.stRc = inRc
        StatType.stXm = inXm
        StatType.stTax = inTax
        Put #i, IndexCount + 1, StatType
    Close #i

End Sub

Sub SaveMonthStatistish(inXm As Currency, inRc As Currency, inTax As Currency)

    On Error Resume Next

    Dim IndexCount As Long, N As Long, DayNow As String, xNow As Date

    i = FreeFile
    xNow = Now

    DayNow = Format(Year(xNow), "0000") & Format(Month(xNow), "00")

    Open LowPath(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
        IndexCount = LOF(i) / Len(StatType)
        For N = 1 To IndexCount
            Get #i, N, StatType
            If Trim(StatType.stDATE) = DayNow And StatType.stType = ST_MODE.ST_MONTHLY Then
                StatType.stRc = inRc
                StatType.stXm = inXm
                StatType.stTax = inTax
                StatType.stMeta = Format(xNow, "mmmm yyyy")
                Put #i, N, StatType
                Close #i: Exit Sub
            End If
        Next N

        StatType.stType = ST_MODE.ST_MONTHLY
        StatType.stDATE = DayNow
        StatType.stRc = inRc
        StatType.stXm = inXm
        StatType.stTax = inTax
        StatType.stMeta = Format(xNow, "mmmm yyyy")
        Put #i, IndexCount + 1, StatType
    Close #i

End Sub

Sub SaveYearStatistish(inXm As Currency, inRc As Currency, inTax As Currency)

    Dim IndexCount As Long, N As Long, DayNow As String, xNow As Date
    
    On Error Resume Next
    
    i = FreeFile
    xNow = Now
    DayNow = Format(Year(xNow), "0000")

    Open LowPath(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
        IndexCount = LOF(i) / Len(StatType)
        For N = 1 To IndexCount
            Get #i, N, StatType
            If Trim(StatType.stDATE) = DayNow And StatType.stType = ST_MODE.ST_YEARLY Then
                StatType.stRc = inRc
                StatType.stXm = inXm
                StatType.stTax = inTax
                StatType.stMeta = Format(xNow, "yyyy")
                Put #i, N, StatType
                Close #i: Exit Sub
            End If
        Next N

        StatType.stType = ST_MODE.ST_YEARLY
        StatType.stDATE = DayNow
        StatType.stRc = inRc
        StatType.stXm = inXm
        StatType.stTax = inTax
        StatType.stMeta = Format(xNow, "yyyy")
        Put #i, IndexCount + 1, StatType
    Close #i

End Sub


Function CountTrafficPerMonth(inMonth As Integer, inDirection As Integer, inOPTION As Integer, outMoney As Currency) As Currency

    Dim IndexCount As Long, N As Long, z As String
    Dim inSum As Currency, inMon As Currency
    Dim inCnt As Integer

    On Error Resume Next

    i = FreeFile
    inSum = 0: inMon = 0: inCnt = 0: lftDate = 0

    Open LowPath(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
        IndexCount = LOF(i) / Len(StatType)
        For N = 1 To IndexCount
            Get #i, N, StatType
            If StatType.stType = ST_MODE.ST_DAILY And Month(Trim(StatType.stDATE)) = inMonth Then
                inMon = inMon + StatType.stTax
                Select Case inDirection
                Case 0
                    inSum = inSum + (StatType.stRc + StatType.stXm)
                Case 1
                    inSum = inSum + (StatType.stRc)
                Case 2
                    inSum = inSum + (StatType.stXm)
                Case 3
                    If inOPTION = 1 Then inSum = inSum + (StatType.stRc)
                    If inOPTION = 2 Then inSum = inSum + (StatType.stXm)
                End Select
                inCnt = inCnt + 1
    
            End If
        Next N
    Close #i

    CountTrafficPerMonth = inSum / inCnt
    outMoney = inMon / inCnt

End Function
Function AverageCap(inMonth As Integer, inDay As Integer, inDirection As Integer, inOPTION As Integer, outMoney As Currency, inYear As Integer) As Currency

    Dim IndexCount As Long, N As Long, z As String
    Dim inSum As Currency, inMon As Currency
    Dim inCnt As Integer
    Dim xmTMP As Currency, rcTMP As Currency, mnTMP As Currency

    On Error Resume Next

    inCnt = 0
    i = FreeFile
    inSum = 0: inMon = 0: inCnt = 0: lftDate = 0

    Open LowPath(App.Path) + "app_data.rpt" For Random As #i Len = Len(StatType)
        IndexCount = LOF(i) / Len(StatType)
        For N = 1 To IndexCount
            Get #i, N, StatType
            If StatType.stType = ST_MODE.ST_DAILY And GetLongFromDataEx(Day(Trim(StatType.stDATE)), Month(Trim(StatType.stDATE)), Year(Trim(StatType.stDATE))) = GetLongFromDataEx(inDay, inMonth, inYear) Then
                rcTMP = StatType.stRc
                xmTMP = StatType.stXm
                mnTMP = StatType.stTax
      
            End If
 
            If StatType.stType = ST_MODE.ST_DAILY And GetLongFromDataEx(Day(Trim(StatType.stDATE)), Month(Trim(StatType.stDATE)), Year(Trim(StatType.stDATE))) < GetLongFromDataEx(inDay, inMonth, inYear) And GetLongFromDataEx(Day(Trim(StatType.stDATE)), Month(Trim(StatType.stDATE)), Year(Trim(StatType.stDATE))) >= GetLongFromDataEx(inDay, inMonth, inYear) - 4 Then
                inCnt = inCnt + 1
                inMon = inMon + StatType.stTax
                Select Case inDirection
                Case 0
                    inSum = inSum + (StatType.stRc + StatType.stXm)
                Case 1
                    inSum = inSum + (StatType.stRc)
                Case 2
                    inSum = inSum + (StatType.stXm)
                Case 3
                    If inOPTION = 1 Then inSum = inSum + (StatType.stRc)
                    If inOPTION = 2 Then inSum = inSum + (StatType.stXm)
                End Select
            End If
        Next N
    Close #i

    If inCnt > 0 Then inSum = inSum / inCnt: inMon = inMon / inCnt

    Select Case inDirection
    Case 0
        AverageCap = IIf(inCnt > 0, inSum, rcTMP + xmTMP)
    Case 1
        AverageCap = IIf(inCnt > 0, inSum, rcTMP)
    Case 2
        AverageCap = IIf(inCnt > 0, inSum, xmTMP)
    Case 3
        If inOPTION = 1 Then AverageCap = IIf(inCnt > 0, inSVS, rcTMP)
        If inOPTION = 2 Then AverageCap = IIf(inCnt > 0, inSVS, xmTMP)
    End Select

    outMoney = IIf(inCnt > 0, inMon, mnTMP)

End Function
