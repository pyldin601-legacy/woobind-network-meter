Attribute VB_Name = "formHWND"
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

Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_DESKTOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_SHOWWINDOW = &H40

Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long


Public Sub OnTopForm(wForm As Form, Вариант As Boolean)

    Select Case Вариант

    Case True
        SetWindowPos wForm.hWnd, -1, wForm.Left / 15, _
            wForm.Top / 15, wForm.Width / 15, _
            wForm.Height / 15, SWP_SHOWWINDOW Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER

    Case False
        SetWindowPos wForm.hWnd, -2, wForm.Left / 15, _
            wForm.Top / 15, wForm.Width / 15, _
            wForm.Height / 15, SWP_NOACTIVATE

    End Select

End Sub

Public Sub GoodTopForm(wForm As Form, Вариант As Boolean)

    Select Case Вариант

    Case True
        SetWindowPos wForm.hWnd, 0, wForm.Left / 15, _
            wForm.Top / 15, wForm.Width / 15, _
            wForm.Height / 15, SWP_SHOWWINDOW

    Case False
        SetWindowPos wForm.hWnd, -2, wForm.Left / 15, _
            wForm.Top / 15, wForm.Width / 15, _
            wForm.Height / 15, SWP_NOACTIVATE

    End Select

End Sub


Public Sub RaiseForm(wForm As Form)

    SetWindowPos wForm.hWnd, HWND_TOP, wForm.Left / 15, _
            wForm.Top / 15, wForm.Width / 15, _
            wForm.Height / 15, SWP_SHOWWINDOW Or SWP_NOACTIVATE


End Sub


Public Sub OnBottomForm(wForm As Form, Вариант As Boolean)

    Select Case Вариант

    Case True
        SetWindowPos wForm.hWnd, 1, wForm.Left / 15, _
            wForm.Top / 15, wForm.Width / 15, _
            wForm.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

    Case False
        SetWindowPos wForm.hWnd, -2, wForm.Left / 15, _
            wForm.Top / 15, wForm.Width / 15, _
            wForm.Height / 15, SWP_NOACTIVATE

    End Select

End Sub





