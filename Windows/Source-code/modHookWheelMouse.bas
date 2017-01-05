Attribute VB_Name = "modHooKWheelMouse"

'-----------------------------------------------------
'    Bitrate calculator by philippe734 - Inspired from the bitrate calculator of MeGUI in a stand alone version
'    Copyright 2010 philippe734
'    http://sourceforge.net/projects/bitratecalc/
'
'    Bitrate calculator by philippe734 is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Bitrate calculator by philippe734 is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program. If not, write to the
'    Free Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'-----------------------------------------------------

'---------------------------------------------------------------------------------------
' Module    : modHookWheelMouse
' Author    : philippe734
' Date      : 07/2012
' Purpose   : Set hook to use mouse wheel to change value of control in frmMain
'---------------------------------------------------------------------------------------

Option Explicit

Private Const HC_ACTION = 0
Private Const WH_MOUSE_LL = 14
Private Const WH_MOUSE = 7
Private Const WM_MOUSEWHEEL = &H20A

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private udtlParamStuct As MSLLHOOKSTRUCT

Public plHooking As Long

Public CtrlFocused As Control

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'

Private Function GetHookStruct(ByVal lParam As Long) As MSLLHOOKSTRUCT
    CopyMemory VarPtr(udtlParamStuct), lParam, LenB(udtlParamStuct)
    GetHookStruct = udtlParamStuct
End Function

Private Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim iSign As Integer

    On Error Resume Next

    If (nCode = HC_ACTION) Then
        If wParam = WM_MOUSEWHEEL Then

            MouseProc = True

            'iSign = IIf(GetHookStruct(lParam).mouseData > 0, 1, -1)
            iSign = IIf(GetHookStruct(lParam).dwExtraInfo > 0, 1, -1)

            Call CtrlHookedEvent(CtrlFocused, iSign)

        End If
        Exit Function
    End If

    MouseProc = CallNextHookEx(0&, nCode, wParam, ByVal lParam)

    On Error GoTo 0
End Function

Public Sub HookMouse()

    If plHooking < 1 Then

        plHooking = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, App.hInstance, GetCurrentThreadId)

        'Debug.Print Timer, "Hook ON"

    End If

End Sub

Public Sub UnHookMouse()
    If plHooking <> 0 Then
        UnhookWindowsHookEx plHooking
        plHooking = 0
        'Debug.Print Timer, "Hook OFF"
    End If
End Sub
