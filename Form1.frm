VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "clsWaveOut Test"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtLoops 
      Height          =   285
      Left            =   1950
      TabIndex        =   15
      Text            =   "0"
      Top             =   2850
      Width           =   990
   End
   Begin VB.CheckBox chkLoop 
      Caption         =   "Loop"
      Height          =   240
      Left            =   1050
      TabIndex        =   12
      Top             =   2550
      Width           =   1365
   End
   Begin VB.Timer tmrPos 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   75
      Top             =   1800
   End
   Begin MSComctlLib.Slider sldPos 
      Height          =   285
      Left            =   1050
      TabIndex        =   10
      Top             =   2250
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   503
      _Version        =   393216
      TickStyle       =   3
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   75
      Top             =   1275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3135
      TabIndex        =   7
      Top             =   1275
      Width           =   1065
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2025
      TabIndex        =   6
      Top             =   1275
      Width           =   1065
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   315
      Left            =   900
      TabIndex        =   5
      Top             =   1275
      Width           =   1065
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   3780
      TabIndex        =   4
      Top             =   750
      Width           =   390
   End
   Begin VB.TextBox txtWAV 
      Height          =   285
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   750
      Width           =   2790
   End
   Begin VB.ComboBox cboDevice 
      Height          =   315
      Left            =   990
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   180
      Width           =   3180
   End
   Begin MSComctlLib.Slider sldVol 
      Height          =   285
      Left            =   1050
      TabIndex        =   13
      Top             =   1875
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   503
      _Version        =   393216
      Max             =   65535
      SelStart        =   65535
      TickStyle       =   3
      Value           =   65535
   End
   Begin VB.Label Label1 
      Caption         =   "0 = unlimited"
      Height          =   240
      Left            =   3225
      TabIndex        =   16
      Top             =   2880
      Width           =   990
   End
   Begin VB.Label lblLoops 
      Caption         =   "Loops:"
      Height          =   240
      Left            =   1320
      TabIndex        =   14
      Top             =   2880
      Width           =   840
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "0:00/0:00"
      Height          =   195
      Left            =   3450
      TabIndex        =   11
      Top             =   2550
      Width           =   720
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      Caption         =   "Position:"
      Height          =   195
      Left            =   330
      TabIndex        =   9
      Top             =   2250
      Width           =   615
   End
   Begin VB.Label lblVol 
      AutoSize        =   -1  'True
      Caption         =   "Volume:"
      Height          =   195
      Left            =   323
      TabIndex        =   8
      Top             =   1875
      Width           =   570
   End
   Begin VB.Label lblWAV 
      AutoSize        =   -1  'True
      Caption         =   "WAV:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   810
      Width           =   405
   End
   Begin VB.Label lblDevice 
      AutoSize        =   -1  'True
      Caption         =   "Device:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clsPlayer    As WAVPlayer
Attribute clsPlayer.VB_VarHelpID = -1

Private blnDontMove             As Boolean

Private Function ShowDevices( _
) As Boolean

    Dim i           As Long
    Dim lngDevices  As Long

    lngDevices = clsPlayer.DeviceCount
    If lngDevices = 0 Then
        Debug.Print "No devices found!"
        Exit Function
    End If

    ' add the WaveOut devices (-1 is the Wave Mapper)
    For i = -1 To lngDevices - 1
        cboDevice.AddItem clsPlayer.DeviceName(i)
    Next

    cboDevice.ListIndex = 0

    ShowDevices = True
End Function

Private Sub cboDevice_Click()
    clsPlayer.SelectedDevice = cboDevice.ListIndex - 1
    sldVol.value = clsPlayer.Volume
End Sub

Private Sub chkLoop_Click()
    clsPlayer.PlaybackLoop = CBool(chkLoop.value = 1)
End Sub

Private Sub clsPlayer_EndOfStream()
    Debug.Print "End Of Stream!"
    clsPlayer_StatusChanged Status_Stopped
End Sub

Private Sub clsPlayer_NextLoop(ByVal LoopCount As Integer, StopPlaying As Boolean)
    Dim lngLoops As Long
    
    Debug.Print "Loop  " & LoopCount & " finished"
    
    lngLoops = Val(txtLoops.Text)
    If lngLoops > 0 Then
        StopPlaying = lngLoops = LoopCount
    End If
End Sub

' status of WaveOut has changed (play, pause, stop)
Private Sub clsPlayer_StatusChanged( _
    ByVal status As PlayerStatus _
)

    Select Case status
        Case Status_Pausing
            cmdPlay.Enabled = True
            cmdPause.Enabled = False
            cmdStop.Enabled = True
            tmrPos.Enabled = False
            cboDevice.Enabled = False
        Case Status_Playing
            cmdPlay.Enabled = False
            cmdPause.Enabled = True
            cmdStop.Enabled = True
            tmrPos.Enabled = True
            cboDevice.Enabled = False
        Case Status_Stopped
            cmdPlay.Enabled = True
            cmdPause.Enabled = False
            cmdStop.Enabled = False
            tmrPos.Enabled = False
            cboDevice.Enabled = True
    End Select
End Sub

Private Sub cmdBrowse_Click()
    Dim strExt  As String

    ' WAV Reader is part of a bigger project, where
    ' all decoders have the same interface, and can support
    ' multiple file formats. This is just copy'n'pasted.
    ' Of course WAV Reader has only support for WAV.

    With dlg
        .FileName = vbNullString
        .Filter = "WAV files (*.wav)|*.wav"
        .ShowOpen
    End With

    If dlg.FileName = vbNullString Then Exit Sub

    clsPlayer.FileClose
    clsPlayer.SelectedDevice = cboDevice.ListIndex - 1

    If Not clsPlayer.FileOpen(dlg.FileName) Then
        MsgBox "Couldn't open the file!", vbExclamation
        txtWAV.Text = ""
    Else
        txtWAV.Text = dlg.FileName
        sldPos.value = 0
        sldPos.Max = clsPlayer.Duration
        sldVol.value = clsPlayer.Volume
    End If
End Sub

Private Sub cmdPause_Click()
    If Not clsPlayer.PlaybackPause() Then
        MsgBox "Couldn't pause!", vbExclamation
    End If
End Sub

Private Sub cmdPlay_Click()
    Dim i   As Long

    If Not clsPlayer.PlaybackStart() Then
        MsgBox "Couldn't play!", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub cmdStop_Click()
    If Not clsPlayer.PlaybackStop() Then
        MsgBox "Couldn't stop!", vbExclamation
    End If
End Sub

Private Sub Form_Load()
    Set clsPlayer = New WAVPlayer

    If Not ShowDevices() Then
        MsgBox "No devices found!", vbExclamation
    End If
End Sub

Private Sub Form_Unload( _
    Cancel As Integer _
)

    Set clsPlayer = Nothing
End Sub

Private Sub sldVol_Change()
    clsPlayer.Volume = sldVol.value
End Sub

Private Sub sldVol_Scroll()
    clsPlayer.Volume = sldVol.value
End Sub

Private Function MS_to_Str( _
    ByVal ms As Long _
) As String

    Dim min As Long
    Dim sec As Long

    sec = ms / 1000
    min = sec \ 60
    sec = sec Mod 60

    MS_to_Str = min & ":" & format(sec, "00")
End Function

Private Sub sldPos_MouseDown( _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single _
)

    blnDontMove = True
End Sub

Private Sub sldPos_MouseUp( _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single _
)

    clsPlayer.Position = sldPos.value

    blnDontMove = False
End Sub

Private Sub tmrPos_Timer()
    If Not blnDontMove Then
        sldPos.value = clsPlayer.Position

        ' current position in minutes:seconds
        lblTime.Caption = MS_to_Str(clsPlayer.Position) & "/" & _
                          MS_to_Str(clsPlayer.Duration)
    Else
        ' if the user currently slides the position slider,
        ' show the stream time at the slider's current position.
        lblTime.Caption = MS_to_Str(sldPos.value) & "/" & _
                          MS_to_Str(clsPlayer.Duration)
    End If
End Sub
