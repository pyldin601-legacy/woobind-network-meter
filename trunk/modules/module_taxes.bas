Attribute VB_Name = "Special_Functions_2"
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

Type RulerFormat
    rfDays As String * 7
    rfTimeStart As Integer
    rfTimeStop As Integer
    rfTariff As Currency
End Type

Global RFRecord As RulerFormat
Global taxes_matrix(7, 24) As Currency


Sub LoadTariff()

    On Error Resume Next
    Dim KTL As Currency

    Open LowPath(App.Path) + "dynamic.bin" For Random As #1 Len = Len(KTL)
        For i = 0 To 24 * 7
            Get #1, i + 1, KTL
            taxes_matrix(i Mod 7, Fix(i / 7)) = KTL
        Next i
    Close #1

End Sub

Sub SaveTariff()

    On Error Resume Next
    Dim KTL As Currency

    Open LowPath(App.Path) + "dynamic.bin" For Random As #1 Len = Len(KTL)
        For i = 0 To 24 * 7
            KTL = taxes_matrix(i Mod 7, Fix(i / 7))
            Put #1, i + 1, KTL
        Next i
    Close #1

End Sub


Sub LoadRulers()

    On Error Resume Next
    Dim c As Integer, d As Integer

    For c = 1 To 7: For d = 0 To 23
        taxes_matrix(c - 1, d) = StaticTariff
    Next d: Next c

    Open LowPath(App.Path) + "rulers.bin" For Random As #1 Len = Len(RFRecord)
        For i = 1 To LOF(1) / Len(RFRecord)
            Get #1, i, RFRecord
            For c = 1 To 7
                If Mid(RFRecord.rfDays, c, 1) = "1" Then
                    If RFRecord.rfTimeStart <= RFRecord.rfTimeStop Then
                        For d = RFRecord.rfTimeStart To RFRecord.rfTimeStop
                            taxes_matrix(c - 1, d) = RFRecord.rfTariff
                        Next d
                    Else
                        For d = RFRecord.rfTimeStart To 23
                            taxes_matrix(c - 1, d) = RFRecord.rfTariff
                        Next d
                        For d = 0 To RFRecord.rfTimeStop
                            taxes_matrix(c - 1, d) = RFRecord.rfTariff
                        Next d
                    End If
                End If
            Next c
        Next i
    Close #1

End Sub


Function ScanRulers() As Boolean

    On Error Resume Next
    ScanRulers = False
    Dim c As Integer, d As Integer, dts(6, 23) As Currency

    For c = 1 To 7: For d = 0 To 23
        dts(c - 1, d) = -20061985
    Next d: Next c

    Open LowPath(App.Path) + "rulers.bin" For Random As #1 Len = Len(RFRecord)
        For i = 1 To LOF(1) / Len(RFRecord)
            Get #1, i, RFRecord
            For c = 1 To 7
                If Mid(RFRecord.rfDays, c, 1) = "1" Then
                    If RFRecord.rfTimeStart <= RFRecord.rfTimeStop Then
                        For d = RFRecord.rfTimeStart To RFRecord.rfTimeStop
                            If dts(c - 1, d) = -20061985 Then dts(c - 1, d) = RFRecord.rfTariff Else ScanRulers = True: Exit Function
                        Next d
                    Else
                        For d = RFRecord.rfTimeStart To 23
                            If dts(c - 1, d) = -20061985 Then dts(c - 1, d) = RFRecord.rfTariff Else ScanRulers = True: Exit Function
                        Next d
                        For d = 0 To RFRecord.rfTimeStop
                            If dts(c - 1, d) = -20061985 Then dts(c - 1, d) = RFRecord.rfTariff Else ScanRulers = True: Exit Function
                        Next d
                    End If
                End If
            Next c
        Next i

    Close #1

End Function

Sub AddRuler(inDays As String, inCost As Currency, inStart As Integer, inStop As Integer)

    Open LowPath(App.Path) + "rulers.bin" For Random As #1 Len = Len(RFRecord)
        RFRecord.rfDays = inDays
        RFRecord.rfTariff = inCost
        RFRecord.rfTimeStart = inStart
        RFRecord.rfTimeStop = inStop
        Put #1, Fix(LOF(1) / Len(RFRecord)) + 1, RFRecord
    Close #1

End Sub

Sub ShowRulers()

    On Error Resume Next
    Dim c As Integer, d As Integer, e As String
    Dialog.lstTariffs.ListItems.Clear

    Open LowPath(App.Path) + "rulers.bin" For Random As #1 Len = Len(RFRecord)
        For i = 1 To Fix(LOF(1) / Len(RFRecord))
            e = ""
            Get #1, i, RFRecord
            For c = 1 To 7
                If Mid(RFRecord.rfDays, c, 1) = "1" Then e = e + Format(c + 1, "ddd") + " "
            Next c
            L = Dialog.lstTariffs.ListItems.Count
            Set f = Dialog.lstTariffs.ListItems.Add(L + 1, "L" + Format(i, "0"), Format(RFRecord.rfTimeStart, "0") & ":00-" & Format(RFRecord.rfTimeStop, "0") & ":59")
            Dialog.lstTariffs.ListItems(L + 1).SubItems(1) = e
            Dialog.lstTariffs.ListItems(L + 1).SubItems(2) = FormatEx(RFRecord.rfTariff, "0.00#")
        Next i
    Close #1

End Sub

Sub DeleteRuler(inIndex As Integer)

    On Error Resume Next
    Dim c As Integer, d As Integer, e As String
    Dialog.lstTariffs.ListItems.Clear

    Open LowPath(App.Path) + "rulers.bin" For Random As #1 Len = Len(RFRecord)
    Open LowPath(App.Path) + "rulerstmp.bin" For Random As #2 Len = Len(RFRecord)
            For i = 1 To Fix(LOF(1) / Len(RFRecord))
                Get #1, i, RFRecord
                If i <> inIndex Then
                    c = c + 1
                    Put #2, c, RFRecord
                End If
            Next i
    Close #2, #1

    Kill LowPath(App.Path) + "rulers.bin"
    Name LowPath(App.Path) + "rulerstmp.bin" As LowPath(App.Path) + "rulers.bin"

End Sub


