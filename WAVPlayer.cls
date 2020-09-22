VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WAVPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_clsWaveOut     As WaveOut
Attribute m_clsWaveOut.VB_VarHelpID = -1
Private m_clsStream                 As StreamWAV

Public Enum PlayerStatus
    Status_Stopped
    Status_Pausing
    Status_Playing
End Enum

Public Event NextLoop(ByVal LoopCount As Integer, ByRef StopPlaying As Boolean)
Public Event StatusChanged(ByVal status As PlayerStatus)
Public Event EndOfStream()

Private Const LOOP_BUFFER_ID        As Long = 1

Private Const ERR_BASE              As Long = &HFF
Private Const ERR_OUT_OF_RANGE      As Long = ERR_BASE + 0
Private Const ERR_INVALID_OPERATION As Long = ERR_BASE + 1

Private m_blnFileOpen               As Boolean

Private m_intDeviceIndex            As Integer

Private m_intBuffers                As Integer
Private m_lngBufferSize             As Long

Private m_blnEndOfStream            As Boolean
Private m_blnLooping                As Boolean

Private m_intLoopCount              As Integer

Private m_udeStatus                 As PlayerStatus

Public Property Get DeviceCount() As Integer
    DeviceCount = m_clsWaveOut.GetNumDevs()
End Property

Public Property Get DeviceName(ByVal index As Integer) As String
    DeviceName = m_clsWaveOut.GetDevName(index)
End Property

Public Property Get SelectedDevice() As Integer
    SelectedDevice = m_intDeviceIndex
End Property

Public Property Get LoopCount() As Integer
    LoopCount = m_intLoopCount
End Property

Public Property Let SelectedDevice(ByVal value As Integer)
    If status = Status_Playing Then
        err.Raise ERR_INVALID_OPERATION, , "Stop playback before you change the device!"
    Else
        m_intDeviceIndex = value
                
        If m_blnFileOpen Then
            m_clsWaveOut.Deinit
            
            With m_clsStream
                If m_clsWaveOut.Init(.SamplesPerSecond, .Channels, .BitsPerSample, m_intDeviceIndex) Then
                    m_blnFileOpen = True
                Else
                    m_clsStream.StreamClose
                    m_blnFileOpen = False
                    err.Raise ERR_INVALID_OPERATION, , "Couldn't change device!"
                End If
            End With
        End If
    End If
End Property

Public Function PlaybackStart() As Boolean
    Dim i           As Long
    Dim intBufCnt   As Long

    If m_blnFileOpen Then

        For i = m_clsWaveOut.Buffers To Buffers - 1
            If Not PlayChunk() Then
                If Not m_blnEndOfStream Then
                    If m_clsWaveOut.Buffers <= 0 Then
                        m_clsWaveOut.ClearBufferQueue
                        PlaybackStart = False
                        Exit Function
                    End If
                End If
            End If
        Next
    
        If m_blnEndOfStream And m_clsWaveOut.Buffers = 0 Then
            m_clsStream.StreamSeek 0, SND_SEEK_SECONDS
            m_blnEndOfStream = False
        
            For i = m_clsWaveOut.Buffers To Buffers - 1
                If Not PlayChunk() Then
                    If Not m_blnEndOfStream Then
                        If m_clsWaveOut.Buffers <= 0 Then
                            m_clsWaveOut.ClearBufferQueue
                            PlaybackStart = False
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
    
        ' If the WAV file is so small that it will fit in
        ' 1 buffer there can be gaps when looping.
        ' Better buffer it a few times more.
        If m_blnLooping And m_clsWaveOut.Buffers = 1 Then
            If m_blnEndOfStream Then
                For i = m_clsWaveOut.Buffers To Buffers - 1
                    m_clsStream.StreamSeek 0, SND_SEEK_SECONDS
                    m_blnEndOfStream = False
                    PlayChunk
                Next
            End If
        End If
    
        If Not m_clsWaveOut.Play() Then
            m_clsWaveOut.ClearBufferQueue
            m_clsWaveOut.StopPlayback
            PlaybackStart = False
        Else
            PlaybackStart = True
            status = Status_Playing
        End If
    Else
        PlaybackStart = False
    End If
End Function

Public Function PlaybackPause() As Boolean
    PlaybackPause = m_clsWaveOut.Pause()
    status = Status_Pausing
End Function

Public Function PlaybackStop() As Boolean
    PlaybackStop = m_clsWaveOut.StopPlayback()
    m_clsStream.StreamSeek 0, SND_SEEK_SECONDS
    m_blnEndOfStream = False
    m_intLoopCount = 0

    status = Status_Stopped
End Function

Public Property Get PlaybackLoop() As Boolean
    PlaybackLoop = m_blnLooping
End Property

Public Property Let PlaybackLoop(ByVal value As Boolean)
    m_blnLooping = value
End Property

Public Property Get Volume() As Long
    Volume = m_clsWaveOut.VolumeForDev(m_intDeviceIndex)
End Property

Public Property Let Volume(ByVal value As Long)
    m_clsWaveOut.VolumeForDev(m_intDeviceIndex) = value
End Property

Public Property Get Buffers() As Integer
    Buffers = m_intBuffers
End Property

Public Property Let Buffers(ByVal value As Integer)
    If value < 2 Then
        err.Raise ERR_OUT_OF_RANGE, , "Invalid number of buffers!"
    Else
        m_intBuffers = value
    End If
End Property

Public Property Get BufferSize() As Long
    BufferSize = m_lngBufferSize
End Property

Public Property Let BufferSize(ByVal value As Long)
    value = value + (value Mod 4)
    
    If value <= 4 Then
        err.Raise ERR_OUT_OF_RANGE, , "Invalid buffer size!"
    Else
        m_lngBufferSize = value
    End If
End Property

Public Property Get Duration() As Long
    Duration = m_clsStream.Duration()
End Property

Public Property Get Position() As Long
    Position = m_clsWaveOut.Elapsed()
End Property

Public Property Let Position(ByVal value As Long)
    If Not m_blnFileOpen Then
        err.Raise ERR_INVALID_OPERATION, , "No file opened!"
    Else
        m_clsStream.StreamSeek value \ 1000, SND_SEEK_SECONDS
        m_clsWaveOut.SetElapsed value
        m_clsWaveOut.ClearBufferQueue
    End If
End Property

Public Property Get status() As PlayerStatus
    status = m_udeStatus
End Property

Private Property Let status(ByVal value As PlayerStatus)
    m_udeStatus = value
    RaiseEvent StatusChanged(m_udeStatus)
End Property

Public Function FileOpen(ByVal strFile As String) As Boolean
    If m_clsStream.StreamOpen(strFile) = SND_ERR_SUCCESS Then
        With m_clsStream
            If m_clsWaveOut.Init(.SamplesPerSecond, .Channels, .BitsPerSample, m_intDeviceIndex) Then
                m_blnFileOpen = True
                FileOpen = True
            Else
                m_clsStream.StreamClose
                m_blnFileOpen = False
                FileOpen = False
            End If
        End With
    Else
        m_blnFileOpen = False
        FileOpen = False
    End If
End Function

Public Function FileClose() As Boolean
    Dim blnReturn As Boolean

    PlaybackStop

    blnReturn = True
    blnReturn = blnReturn And (m_clsWaveOut.Deinit())
    blnReturn = blnReturn And (m_clsStream.StreamClose() = SND_ERR_SUCCESS)

    m_blnFileOpen = False
    m_blnEndOfStream = False
    m_intLoopCount = 0

    FileClose = blnReturn
End Function

Private Sub Class_Initialize()
    Set m_clsWaveOut = New WaveOut
    Set m_clsStream = New StreamWAV
    
    m_intBuffers = 10
    m_lngBufferSize = 10& * 1024&

    m_udeStatus = Status_Stopped
    m_blnEndOfStream = False
End Sub

Private Sub Class_Terminate()
    m_clsWaveOut.Deinit
    m_clsStream.StreamClose
    
    Set m_clsWaveOut = Nothing
    Set m_clsStream = Nothing
End Sub

Private Function PlayChunk() As Boolean
    Dim bt()        As Byte
    Dim lngRead     As Long
    Dim lngUserData As Long

    If m_blnEndOfStream Then
        PlayChunk = False
        Exit Function
    End If
    
    ReDim bt(m_lngBufferSize - 1) As Byte

    ' read a chunk of wave data (PCM)
    m_clsStream.StreamRead VarPtr(bt(0)), m_lngBufferSize, lngRead

    ' if the WAV Reader didn't return any samples,
    ' we reached the end of the file
    If lngRead = 0 Then
        m_blnEndOfStream = True
        PlayChunk = False
        Exit Function
    ElseIf lngRead < m_lngBufferSize Then
        If m_blnLooping Then
            ' read less than m_lngBufferSize probably means End Of Stream,
            ' so mark it as the last buffer of a loop
            lngUserData = LOOP_BUFFER_ID
        End If
    End If

    ' got the samples, add them to the WaveOut Buffer Queue
    PlayChunk = m_clsWaveOut.WriteSamples(VarPtr(bt(0)), lngRead, lngUserData)
End Function

Private Sub m_clsWaveOut_BufferDone(ByVal userdata As Long)
    Dim i       As Long
    Dim blnStop As Boolean
    
    ' a buffer in the WaveOut Buffer Queue got played,
    ' add a new one to it so the stream doesn't get interrupted
    
    ' If end of file (or stream) is reached and there
    ' are no more buffers in the WaveOut Buffer Queue,
    ' stop playback.
    If m_blnEndOfStream Then
        m_clsStream.StreamSeek 0, SND_SEEK_PERCENT

        If Not m_blnLooping Then
            If m_clsWaveOut.Buffers = 0 Then
                m_blnEndOfStream = False
                m_clsWaveOut.StopPlayback
                m_udeStatus = Status_Stopped
                RaiseEvent EndOfStream
                Exit Sub
            End If
        Else
            m_blnEndOfStream = False
        End If
    End If
    
    If userdata = LOOP_BUFFER_ID Then
        ' the buffer containing the last bits of the file
        ' will be marked with LOOP_BUFFER_ID, so we can identify
        ' end of stream when looping
        m_clsWaveOut.ResetElapsed
        m_intLoopCount = m_intLoopCount + 1
        
        RaiseEvent NextLoop(m_intLoopCount, blnStop)
        
        If blnStop Then
            m_clsWaveOut.StopPlayback
            m_intLoopCount = 0
            m_udeStatus = Status_Stopped
            RaiseEvent EndOfStream
            Exit Sub
        End If
    End If

    For i = m_clsWaveOut.Buffers To m_intBuffers - 1
        If Not PlayChunk() Then
            If Not m_blnEndOfStream Then
                ' must be a serious problem
                m_clsWaveOut.StopPlayback
            End If
        End If
    Next

    ' If the WAV file is so small that it will fit in
    ' 1 buffer there can be gaps when looping.
    ' Better buffer it a few times more.
    If m_blnLooping And m_clsWaveOut.Buffers = 1 Then
        If m_blnEndOfStream Then
            For i = m_clsWaveOut.Buffers To Buffers - 1
                m_clsStream.StreamSeek 0, SND_SEEK_SECONDS
                m_blnEndOfStream = False
                PlayChunk
            Next
        End If
    End If
End Sub
