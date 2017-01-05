VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSiteWeb 
      Caption         =   "Web site"
      Height          =   300
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "http://sourceforge.net/projects/bitratecalc/"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Close"
      Height          =   300
      Left            =   3000
      TabIndex        =   1
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Free and Open Source GNU/GPL"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "©2012 philippe734"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1350
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":000C
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Line Line 
      X1              =   -480
      X2              =   4320
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "Donate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   1
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "http://vpnlifeguard.blogspot.com/p/faire-un-don.html"
      ToolTipText     =   "Paypal"
      Top             =   3000
      Width           =   3675
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "Source code and last version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   0
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Tag             =   "http://sourceforge.net/projects/bitratecalc/"
      ToolTipText     =   "http://sourceforge.net/projects/bitratecalc/"
      Top             =   3240
      Width           =   3795
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "App title"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmAbout"
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
' Module    : frmAbout
' Author    : philippe734
' Date      : 07/2012
' Purpose   : About the application
'---------------------------------------------------------------------------------------

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long    'DLL inclus dans tout les windows

Option Explicit

Private Sub cmdQuit_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdSiteWeb_Click()
    On Error Resume Next
    ShellExecute 0&, vbNullString, cmdSiteWeb.ToolTipText, vbNullString, vbNullString, vbNormalFocus
    DoEvents
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.lblTitle.Caption = "Bitrate calculator" & " - " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim item As Control

    On Error Resume Next
    For Each item In frmAbout.Controls
        If TypeOf item Is Label Then
            item.FontUnderline = False
        End If
    Next item
End Sub

Private Sub lblURL_Click(Index As Integer)
    On Error Resume Next
    ShellExecute 0&, vbNullString, lblURL(Index).Tag, vbNullString, vbNullString, vbNormalFocus
    DoEvents
End Sub

Private Sub lblURL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Me.lblURL(Index).FontUnderline = True
    MousePointerHand
End Sub

Private Function MousePointerHand()
    Const IDC_HAND As Long = 32649
    Dim iHandle As Long

    On Error Resume Next
    iHandle = LoadCursor(0, IDC_HAND)
    If (iHandle > 0) Then
        iHandle = SetCursor(iHandle)
    End If
End Function

