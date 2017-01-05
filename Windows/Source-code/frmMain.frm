VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitrate calculator GPL"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRatio 
      Height          =   315
      Left            =   2880
      TabIndex        =   17
      ToolTipText     =   "Check to keep aspect ratio"
      Top             =   1560
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.ComboBox comboRatio 
      Height          =   315
      Left            =   3145
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame Frame 
      Caption         =   "Keep aspect ratio"
      Height          =   615
      Index           =   3
      Left            =   2760
      TabIndex        =   32
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtAudioBitrate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   33
      Text            =   "96"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtAudioSize 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   0
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "0"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtAudioBitrate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   35
      Text            =   "0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtAudioSize 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "0"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtAudioBitrate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   41
      Text            =   "0"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox txtAudioSize 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "0"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtTotalFrames 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0"
      Top             =   720
      Width           =   870
   End
   Begin VB.TextBox txtMin 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Text            =   "30"
      Top             =   720
      Width           =   555
   End
   Begin VB.TextBox txtSeconds 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Text            =   "0"
      Top             =   720
      Width           =   630
   End
   Begin VB.TextBox txtWidth 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Text            =   "1280"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtHeight 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Text            =   "720"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtHours 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Text            =   "1"
      Top             =   720
      Width           =   420
   End
   Begin VB.ComboBox comboFPS 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtHookStatut 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Hooking OFF"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtAspectRatio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "1.778"
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox comboSizePreset 
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   3480
      Width           =   1470
   End
   Begin VB.TextBox txtFinalFileSize 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   48
      Text            =   "1000"
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtBitrate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtBitsPerPixel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.OptionButton optVideo 
      Caption         =   "Video bitrate"
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   43
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton optVideo 
      Caption         =   "Bits/(pixel*frame)"
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   45
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton optVideo 
      Caption         =   "Final file size"
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   47
      Top             =   3120
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About..."
      Height          =   375
      Left            =   6120
      TabIndex        =   50
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   7560
      TabIndex        =   51
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame Frame 
      Caption         =   "Calculate from"
      Height          =   4095
      Index           =   2
      Left            =   5520
      TabIndex        =   37
      Top             =   120
      Width           =   3375
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "MB"
         Height          =   195
         Index           =   13
         Left            =   2640
         TabIndex        =   39
         Top             =   3045
         Width           =   240
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "kbps"
         Height          =   195
         Index           =   12
         Left            =   2470
         TabIndex        =   38
         Top             =   880
         Width           =   345
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Audio"
      Height          =   2655
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   5295
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   4200
         X2              =   360
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   4200
         X2              =   360
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "kbps"
         Height          =   195
         Index           =   16
         Left            =   2160
         TabIndex        =   31
         Top             =   2200
         Width           =   345
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "kbps"
         Height          =   195
         Index           =   15
         Left            =   2160
         TabIndex        =   30
         Top             =   1500
         Width           =   345
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "kbps"
         Height          =   195
         Index           =   14
         Left            =   2160
         TabIndex        =   29
         Top             =   780
         Width           =   345
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Track 3"
         Height          =   195
         Index           =   11
         Left            =   600
         TabIndex        =   28
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Index           =   10
         Left            =   3600
         TabIndex        =   27
         Top             =   1920
         Width           =   300
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Bitrate"
         Height          =   195
         Index           =   9
         Left            =   1560
         TabIndex        =   26
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Track 2"
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   25
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Index           =   7
         Left            =   3600
         TabIndex        =   24
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Bitrate"
         Height          =   195
         Index           =   6
         Left            =   1560
         TabIndex        =   23
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Track 1"
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   22
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Index           =   4
         Left            =   3600
         TabIndex        =   21
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Bitrate"
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   20
         Top             =   480
         Width           =   450
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Video"
      Height          =   2055
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Ratio"
         Height          =   195
         Index           =   18
         Left            =   1800
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   7
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total frames"
         Height          =   195
         Left            =   4080
         TabIndex        =   5
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Framerate"
         Height          =   195
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Seconds"
         Height          =   195
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Minutes"
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Hours"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.Label lblInfoMouseWheel 
      AutoSize        =   -1  'True
      Caption         =   "Use the mouse wheel on all elements ;-)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   5520
      TabIndex        =   52
      Top             =   4320
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
' Module    : frmMain
' Author    : philippe734
' Date      : 07/2012
' Purpose   : Calculation of the bitrate for encoding a video
'---------------------------------------------------------------------------------------


Option Explicit

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long    'DLL inclus dans tout les windows
'

Private Sub Form_Initialize()
    On Error Resume Next
    InitCommonControls
End Sub

Private Sub Form_Load()

    On Error Resume Next

    Debug.Print Timer, "Version "; App.Major & "." & App.Minor & "." & App.Revision

    comboFPS.AddItem "23,976"
    comboFPS.AddItem "24,0"
    comboFPS.AddItem "25,0"
    comboFPS.AddItem "29,97"
    comboFPS.AddItem "30,0"
    comboFPS.AddItem "50,0"
    comboFPS.AddItem "59,94"
    comboFPS.AddItem "60,0"

    comboSizePreset.AddItem "Custom "
    comboSizePreset.AddItem "350 "
    comboSizePreset.AddItem "700 (1 CD)"
    comboSizePreset.AddItem "1000 "
    comboSizePreset.AddItem "1400 (2 CDs)"
    comboSizePreset.AddItem "4480 (DVD)"
    comboSizePreset.AddItem "23450 (BD)"
    comboSizePreset.AddItem "46900 (BD-DL)"

    comboRatio.AddItem "1.00 (1:1)"
    comboRatio.AddItem "1.50 (3:2)"
    comboRatio.AddItem "1.33 (4:3)"
    comboRatio.AddItem "1.25 (5:4)"
    comboRatio.AddItem "0.83 (5:6)"
    comboRatio.AddItem "1.22 (11:9)"
    comboRatio.AddItem "1.56 (14:9)"
    comboRatio.AddItem "1.67 (15:9)"
    comboRatio.AddItem "1.78 (16:9)"
    comboRatio.AddItem "1.60 (16:10)"
    comboRatio.AddItem "1.89 (17:9)"
    comboRatio.AddItem "2.33 (2.33:1)"
    comboRatio.AddItem "2.21 (2.21:1)"
    comboRatio.AddItem "2.35 (2.35:1)"
    comboRatio.AddItem "2.40 (2.40:1)"

End Sub

Private Sub Form_Activate()

    On Error Resume Next

    comboFPS.ListIndex = 0

    comboSizePreset.ListIndex = 0

    comboRatio.ListIndex = 9

    Call Calculation

    ' Set hook on mouse wheel to change value of controls
    Call HookMouse
    txtHookStatut.Visible = IIf(plHooking > 0, False, True)

    txtHours.SetFocus
End Sub

Private Sub Form_GotFocus()
    Call HookMouse
End Sub

Private Sub Form_LostFocus()
    Call UnHookMouse
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call UnHookMouse
End Sub

Private Sub Form_Terminate()
    Call UnHookMouse
End Sub

Private Sub Frame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub optVideo_Click(Index As Integer)
    Select Case Index
    Case 0
        txtBitrate.Locked = False
        txtBitrate.BackColor = &H80000005
        txtBitsPerPixel.BackColor = &H8000000F
        txtFinalFileSize.BackColor = &H8000000F
        txtBitsPerPixel.Locked = True
        txtFinalFileSize.Locked = True
        comboSizePreset.Visible = False
        txtBitrate.SetFocus
    Case 1
        txtBitsPerPixel.Locked = False
        txtBitrate.BackColor = &H8000000F
        txtBitsPerPixel.BackColor = &H80000005
        txtFinalFileSize.BackColor = &H8000000F
        txtBitrate.Locked = True
        txtFinalFileSize.Locked = True
        comboSizePreset.Visible = False
        txtBitsPerPixel.SetFocus
    Case 2
        txtFinalFileSize.Locked = False
        comboSizePreset.Visible = True
        txtBitrate.BackColor = &H8000000F
        txtBitsPerPixel.BackColor = &H8000000F
        txtFinalFileSize.BackColor = &H80000005
        txtBitrate.Locked = True
        txtBitsPerPixel.Locked = True
        txtFinalFileSize.SetFocus
    End Select
End Sub

Private Sub cmdAbout_Click()
    ' Disable hooking
    Call UnHookMouse
    frmAbout.Show vbModal
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub comboRatio_Click()
    Call CheckResolution(frmMain.txtWidth, frmMain.txtHeight, True)
    Call Calculation
    txtWidth.SetFocus
End Sub

Private Sub comboRatio_GotFocus()
    On Error Resume Next
    Set CtrlFocused = comboRatio
End Sub

Private Sub comboFPS_Click()
    On Error Resume Next
    Call Calculation
End Sub

Private Sub comboFPS_GotFocus()
    On Error Resume Next
    Set CtrlFocused = comboFPS
End Sub

Private Sub chkRatio_Click()
    On Error Resume Next
    Set CtrlFocused = chkRatio
    If chkRatio.Value = vbChecked Then
        comboRatio.Locked = False
        comboRatio.BackColor = &H80000005
        Call CheckResolution(frmMain.txtWidth, frmMain.txtHeight, True)
        Call Calculation
    Else
        comboRatio.Locked = True
        comboRatio.BackColor = &H8000000F
    End If
End Sub

Private Sub txtAspectRatio_GotFocus()
    On Error Resume Next
    Set CtrlFocused = txtAspectRatio
End Sub

Private Sub txtAudioBitrate_Change(Index As Integer)
    On Error Resume Next
    If txtAudioBitrate(Index) < 0 Then
        txtAudioBitrate(Index) = 0
    End If
End Sub

Private Sub txtAudioBitrate_GotFocus(Index As Integer)
    On Error Resume Next
    Set CtrlFocused = txtAudioBitrate(Index)
    txtAudioBitrate(Index).SelStart = 0
    txtAudioBitrate(Index).SelLength = 9999
End Sub

Private Sub txtAudioBitrate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        txtAudioBitrate(Index) = txtAudioBitrate(Index) + 32
        Call Calculation
    Case vbKeyDown
        txtAudioBitrate(Index) = txtAudioBitrate(Index) - 32
        Call Calculation
    End Select
End Sub

Private Sub txtAudioBitrate_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckKeyPress(KeyAscii)
End Sub

Private Sub txtAudioSize_GotFocus(Index As Integer)
    On Error Resume Next
    Set CtrlFocused = txtAudioSize(Index)
End Sub

Private Sub comboSizePreset_GotFocus()
    On Error Resume Next
    Set CtrlFocused = comboSizePreset
End Sub

Private Sub comboSizePreset_Click()
    On Error Resume Next
    If (GetSizeCombo & " ") <> comboSizePreset.List(0) Then
        txtFinalFileSize = GetSizeCombo
        Call Calculation
        comboSizePreset.ListIndex = 0
    End If
    txtFinalFileSize.SetFocus
End Sub

Private Sub comboSizePreset_LostFocus()
    On Error Resume Next
End Sub

Private Sub txtBitrate_GotFocus()
    On Error Resume Next
    Set CtrlFocused = txtBitrate
    txtBitrate.SelStart = 0
    txtBitrate.SelLength = 9999
End Sub

Private Sub txtBitrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        txtBitrate = txtBitrate + 1
        Call Calculation
    Case vbKeyDown
        txtBitrate = txtBitrate - 1
        Call Calculation
    End Select
End Sub

Private Sub txtBitrate_KeyPress(KeyAscii As Integer)
    Call CheckKeyPress(KeyAscii)
End Sub

Private Sub txtBitrate_LostFocus()
    Call Calculation
End Sub

Private Sub txtAudioBitrate_LostFocus(Index As Integer)
    Call Calculation
End Sub

Private Sub txtBitsPerPixel_GotFocus()
    On Error Resume Next
    Set CtrlFocused = txtBitsPerPixel
    txtBitsPerPixel.SelStart = 0
    txtBitsPerPixel.SelLength = 9999
End Sub

Private Sub txtBitsPerPixel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        txtBitsPerPixel = txtBitsPerPixel + 0.001
        Call Calculation
    Case vbKeyDown
        txtBitsPerPixel = txtBitsPerPixel - 0.001
        Call Calculation
    End Select
End Sub

Private Sub txtBitsPerPixel_KeyPress(KeyAscii As Integer)
    Call CheckKeyPressDot(KeyAscii)
End Sub

Private Sub txtBitsPerPixel_LostFocus()
    Call Calculation
End Sub

Private Sub txtFinalFileSize_GotFocus()
    On Error Resume Next
    Set CtrlFocused = txtFinalFileSize
    txtFinalFileSize.SelStart = 0
    txtFinalFileSize.SelLength = 9999
End Sub

Private Sub txtFinalFileSize_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        txtFinalFileSize = txtFinalFileSize + 1
        Call Calculation
    Case vbKeyDown
        txtFinalFileSize = txtFinalFileSize - 1
        Call Calculation
    End Select
End Sub

Private Sub txtFinalFileSize_KeyPress(KeyAscii As Integer)
    Call CheckKeyPressDot(KeyAscii)
End Sub

Private Sub txtFinalFileSize_LostFocus()
    Call Calculation
End Sub

Private Sub txtHeight_GotFocus()
    On Error Resume Next
    Set CtrlFocused = txtHeight
    txtHeight.SelStart = 0
    txtHeight.SelLength = 9999
End Sub

Private Sub txtHeight_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        txtHeight = txtHeight + 8
        Call CheckResolution(frmMain.txtHeight, frmMain.txtWidth, False)
        Call Calculation
    Case vbKeyDown
        txtHeight = txtHeight - 8
        Call CheckResolution(frmMain.txtHeight, frmMain.txtWidth, False)
        Call Calculation
    End Select
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call CheckResolution(frmMain.txtHeight, frmMain.txtWidth, False)
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

Private Sub txtHeight_LostFocus()
    Call CheckResolution(frmMain.txtHeight, frmMain.txtWidth, False)
    Call Calculation
End Sub

Private Sub txtHours_Change()
    On Error Resume Next
    If CSng(txtHours) < 0 Then
        txtHours = 0
    End If
End Sub

Private Sub txtHours_GotFocus()
    On Error Resume Next
    txtHours.SelStart = 0
    txtHours.SelLength = 9999
    Set CtrlFocused = txtHours
End Sub

Private Sub txtHours_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        txtHours = txtHours + 1
        Call Calculation
    Case vbKeyDown
        txtHours = txtHours - 1
        Call Calculation
    End Select
End Sub

Private Sub txtHours_KeyPress(KeyAscii As Integer)
    Call CheckKeyPress(KeyAscii)
End Sub

Private Sub txtHours_LostFocus()
    Call Calculation
End Sub

Private Sub txtMin_Change()
    On Error Resume Next
    If CSng(txtMin) < 0 Then
        txtMin = 0
    ElseIf CSng(txtMin) > 59 Then
        txtMin = 59
    End If
End Sub

Private Sub txtMin_GotFocus()
    On Error Resume Next
    txtMin.SelStart = 0
    txtMin.SelLength = 9999
    Set CtrlFocused = txtMin
End Sub

Private Sub txtMin_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        txtMin = txtMin + 1
        Call Calculation
    Case vbKeyDown
        txtMin = txtMin - 1
        Call Calculation
    End Select
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
    Call CheckKeyPress(KeyAscii)
End Sub

Private Sub txtMin_LostFocus()
    Call Calculation
End Sub

Private Sub txtSeconds_Change()
    On Error Resume Next
    If CSng(txtSeconds) < 0 Then
        txtSeconds = 0
    ElseIf CSng(txtSeconds) > 59 Then
        txtSeconds = 59
    End If
End Sub

Private Sub txtSeconds_GotFocus()
    On Error Resume Next
    Set CtrlFocused = txtSeconds
    txtSeconds.SelStart = 0
    txtSeconds.SelLength = 9999
End Sub

Private Sub txtSeconds_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        txtSeconds = txtSeconds + 1
        Call Calculation
    Case vbKeyDown
        txtSeconds = txtSeconds - 1
        Call Calculation
    End Select
End Sub

Private Sub txtSeconds_KeyPress(KeyAscii As Integer)
    Call CheckKeyPress(KeyAscii)
End Sub

Private Sub txtSeconds_LostFocus()
    Call Calculation
End Sub

Private Sub txtTotalFrames_GotFocus()
    On Error Resume Next
    Set CtrlFocused = txtTotalFrames
    txtTotalFrames.SelStart = 0
    txtTotalFrames.SelLength = 9999
End Sub

Private Sub txtTotalFrames_KeyPress(KeyAscii As Integer)
    Call CheckKeyPress(KeyAscii)
End Sub

Private Sub txtTotalFrames_LostFocus()
    Call Calculation
End Sub

Private Sub txtWidth_GotFocus()
    On Error Resume Next
    Set CtrlFocused = txtWidth
    txtWidth.SelStart = 0
    txtWidth.SelLength = 9999
End Sub

Private Sub txtWidth_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        txtWidth = txtWidth + 8
        Call CheckResolution(frmMain.txtWidth, frmMain.txtHeight, True)
        Call Calculation
    Case vbKeyDown
        txtWidth = txtWidth - 8
        Call CheckResolution(frmMain.txtWidth, frmMain.txtHeight, True)
        Call Calculation
    End Select
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call CheckResolution(frmMain.txtWidth, frmMain.txtHeight, True)
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

Private Sub txtWidth_LostFocus()
    Call CheckResolution(frmMain.txtWidth, frmMain.txtHeight, True)
    Call Calculation
End Sub

