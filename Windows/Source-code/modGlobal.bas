Attribute VB_Name = "modGlobal"

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
' Module    : modGlobal
' Author    : philippe734
' Date      : 07/2012
' Purpose   : Calculation of the bitrate for encoding a video
'---------------------------------------------------------------------------------------


Option Explicit

' 1 mega-byte (Mb) = 8 mega-bits (MB)
' 8 MB constant =
Public Const cMB As Long = 8388608    ' = 1024*1024*8

' Length of the movie
Public iLength As Single

' Audio size only (MB)
Public iAudioSize As Single

' Video size only (MB)
Public iVideoSize As Single

' Intermediate value Bits/frame
Public iBitsPerFrame As Single
'

Public Sub CalcBitsPerPixelFrame()
    Dim iBitsPerPixelFrame As Single

    ' Bits/(pixel*frame)

    On Error Resume Next

    ' Bits/frame
    iBitsPerFrame = CSng(frmMain.txtBitrate) * 1000 / CSng(frmMain.comboFPS.Text)

    ' Bits/(pixel*frame)
    iBitsPerPixelFrame = iBitsPerFrame / (Val(frmMain.txtWidth) * Val(frmMain.txtHeight))
    frmMain.txtBitsPerPixel = IIf(iBitsPerPixelFrame > 0.001, Round(iBitsPerPixelFrame, 3), 0.001)
End Sub

Public Sub CalcFinalFileSize()
    Dim iFileSize As Single

    On Error Resume Next

    ' Video size only (MB)
    iVideoSize = (CSng(frmMain.txtBitrate) * iLength * 1000) / cMB

    ' Final file size (MB)
    iFileSize = iVideoSize + iAudioSize
    frmMain.txtFinalFileSize = IIf(iFileSize > 0, Round(iFileSize, 2), 0.01)
End Sub

Public Sub CheckResolution(ByVal iSource As Single, ByVal iSet As Single, bWidthAsSource As Boolean)
    Dim iRatio As Single

    On Error Resume Next

    iSource = CheckValueByEight(iSource)

    If frmMain.chkRatio.Value = vbChecked Then

        iRatio = GetRatioSelected

        If bWidthAsSource = True Then
            iSet = iSource / iRatio
        Else
            iSet = iSource * iRatio
        End If

        iSet = CheckValueByEight(iSet)

        If bWidthAsSource = True Then
            frmMain.txtHeight = iSet
        Else
            frmMain.txtWidth = iSet
        End If

    End If

    If bWidthAsSource = True Then
        frmMain.txtWidth = iSource
        iRatio = iSource / iSet
    Else
        frmMain.txtHeight = iSource
        iRatio = iSet / iSource
    End If

    frmMain.txtAspectRatio = Round(iRatio, 3)

End Sub

Public Sub CheckKeyPress(ByRef KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call Calculation
    Case Is < 32
        ' Control keys are OK
    Case 48 To 57
        ' This is a digit
    Case Else
        ' Reject any other key
        KeyAscii = 0
    End Select
End Sub

Public Sub CheckKeyPressDot(ByRef KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call Calculation
    Case 44, 46
        ' This comma or point
        KeyAscii = 44
    Case Is < 32
        ' Control keys are OK
    Case 48 To 57
        ' This is a digit
    Case Else
        ' Reject any other key
        KeyAscii = 0
    End Select
End Sub

Private Function IsolateString(ByVal str As String, ByVal sLeft As String, ByVal sRigth As String) As String
    ' Get string beetween left character and right character
    IsolateString = Mid(str, InStr(1, str, sLeft, vbTextCompare) + 1, (InStr(1, str, sRigth, vbTextCompare) - InStr(1, str, sLeft, vbTextCompare)) - 1)
End Function

Public Function GetRatioSelected() As Single
    Dim sAspect As String
    Dim sA As String
    Dim sB As String

    ' Get the ratio selected from the comboRatio

    sAspect = frmMain.comboRatio.List(frmMain.comboRatio.ListIndex)

    sA = IsolateString(sAspect, "(", ":")

    sB = IsolateString(sAspect, ":", ")")

    GetRatioSelected = Val(sA) / sB

End Function

Public Sub CtrlHookedEvent(ByRef ctr As Control, ByVal iSign As Integer)
    Dim bCalc As Boolean

    ' Hooking mouse wheel : change value of the control hooked

    On Error Resume Next

    With frmMain

        Select Case ctr.Name

        Case .txtHours.Name, .txtMin.Name, .txtSeconds.Name, .txtBitrate.Name
            ctr = ctr + 3 * iSign
            bCalc = True

        Case .txtFinalFileSize.Name
            ctr = ctr + 50 * iSign
            bCalc = True

        Case .txtWidth.Name
            ctr = ctr + 24 * iSign
            Call CheckResolution(frmMain.txtWidth, frmMain.txtHeight, True)
            bCalc = True

        Case .txtHeight.Name
            ctr = ctr + 24 * iSign
            Call CheckResolution(frmMain.txtHeight, frmMain.txtWidth, False)
            bCalc = True

        Case .txtAudioBitrate(0).Name
            ctr = ctr + 32 * iSign
            bCalc = True

        Case .txtBitsPerPixel.Name
            ctr = ctr + 0.001 * iSign
            bCalc = True

        Case .comboRatio.Name
            ctr.TopIndex = ctr.TopIndex - 1 * iSign
            bCalc = True

        End Select

    End With

    If bCalc = True Then Call Calculation

End Sub

Public Function GetSizeCombo() As String
    ' Extract size value from comboSizePreset

    On Error Resume Next
    GetSizeCombo = Mid(frmMain.comboSizePreset.List(frmMain.comboSizePreset.ListIndex), 1, InStr(1, frmMain.comboSizePreset.List(frmMain.comboSizePreset.ListIndex), " ", vbTextCompare) - 1)
End Function

Private Function CheckValueByEight(ByVal iValue As Single) As Single
    Dim xD As Single
    Dim xU As Single

    ' Check resolution multiple by 8

    xD = Int(iValue)
    Do Until (xD / 8) = Int(xD / 8)
        xD = xD - 1
    Loop

    xU = Int(iValue)
    Do Until (xU / 8) = Int(xU / 8)
        xU = xU + 1
    Loop

    If (iValue - xD) < (xU - iValue) Then
        CheckValueByEight = xD
    Else
        CheckValueByEight = xU
    End If
End Function

Public Sub Calculation()
    Dim k As Byte
    Dim opt As Byte
    Static bFlag As Boolean
    Dim iFileSize As Single
    Dim iVideoBitrate As Single
    Dim iBitsPerPixelFrame As Single

    If bFlag = True Then Exit Sub
    bFlag = True

    On Error Resume Next

    'Debug.Print Timer, "calculation"

    ' Length in seconds
    iLength = CSng(frmMain.txtHours * 3600 + frmMain.txtMin * 60 + frmMain.txtSeconds)

    frmMain.txtTotalFrames = iLength * CSng(frmMain.comboFPS.Text)

    ' Audio size only (MB)
    iAudioSize = 0
    For k = 0 To 2
        iAudioSize = iAudioSize + (iLength * frmMain.txtAudioBitrate(k) / 8 / 1000) / 1.024 / 1.024
        frmMain.txtAudioSize(k) = Round((iLength * frmMain.txtAudioBitrate(k) / 8 / 1000) / 1.024 / 1.024, 1) & " MB"
    Next k

    ' Get index of option selected calculate by
    For opt = 0 To frmMain.optVideo.Count - 1
        If frmMain.optVideo(opt).Value = True Then
            Exit For
        End If
    Next opt

    ' Calculate from ...
    Select Case opt

    Case 0    ' from Bitrate

        ' Limit bitrate > 0
        iVideoBitrate = Val(frmMain.txtBitrate)
        frmMain.txtBitrate = IIf(iVideoBitrate > 0, Round(iVideoBitrate, 0), 1)

        ' Bits/(pixel*frame)
        Call CalcBitsPerPixelFrame

        ' Final file size (MB)
        Call CalcFinalFileSize

    Case 1    ' from Bits/(pixel*frame)

        ' Limit Bits/(pixel*frame) > 0.001
        iBitsPerPixelFrame = CSng(frmMain.txtBitsPerPixel)
        frmMain.txtBitsPerPixel = IIf(iBitsPerPixelFrame > 0.001, Round(iBitsPerPixelFrame, 3), 0.001)

        ' Bits/Frame
        iBitsPerFrame = iBitsPerPixelFrame * Val(frmMain.txtWidth) * Val(frmMain.txtHeight)

        ' Bitrate video (kbps)
        iVideoBitrate = iBitsPerFrame * CSng(frmMain.comboFPS.Text) / 1000
        frmMain.txtBitrate = IIf(iVideoBitrate > 0, Round(iVideoBitrate, 0), 1)

        ' Final file size (MB)
        Call CalcFinalFileSize

    Case 2    ' from Final file size

        ' Limit file size > 0
        iFileSize = CSng(frmMain.txtFinalFileSize)
        frmMain.txtFinalFileSize = IIf(iFileSize > 0, Round(iFileSize, 2), 0.01)

        ' Video size only (MB)
        iVideoSize = CSng(frmMain.txtFinalFileSize) - iAudioSize

        ' Bitrate video (kbps)
        iVideoBitrate = (iVideoSize * cMB) / (iLength * 1000)
        frmMain.txtBitrate = IIf(iVideoBitrate > 0, Round(iVideoBitrate, 0), 1)

        ' Bits/(pixel*frame)
        Call CalcBitsPerPixelFrame

    End Select

    bFlag = False
End Sub

