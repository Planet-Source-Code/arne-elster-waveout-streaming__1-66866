VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamWAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**************************************'
' clsStreamWAV                         '
'                                      '
' Reads WAVs, with ACM support         '
' for compressed formats               '
'                                      '
' Will only return 16 bit samples!     '
' If a codec can't convert from 8 bit  '
' to 16 bit then conversion will fail! '
'                                      '
' Supported Chunks: data, fmt          '
'**************************************'


Private Const Extensions    As String = "WAV"
Private Const Description   As String = "Wave Audio"


Private Declare Function CreateFile Lib "kernel32.dll" _
Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any _
) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Any _
) As Long

Private Declare Function SetFilePointer Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lDistanceToMove As Long, _
    ByVal lpDistanceToMoveHigh As Long, _
    ByVal dwMoveMethod As Long _
) As Long

Private Declare Function GetFileSize Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpFileSizeHigh As Long _
) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long _
) As Long

Private Declare Sub ZeroMem Lib "kernel32" _
Alias "RtlZeroMemory" ( _
    pDst As Any, _
    ByVal dwLen As Long _
)

Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal cb As Long _
)

Private Declare Function IsBadReadPtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long

Private Declare Function IsBadWritePtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long

Private Declare Function acmStreamPrepareHeader Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwPrepare As Long _
) As Long

Private Declare Function acmStreamUnprepareHeader Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwUnprepare As Long _
) As Long

Private Declare Function acmStreamOpen Lib "msacm32" ( _
    phas As Long, _
    ByVal had As Long, _
    pwfxSrc As Any, _
    pwfxDst As Any, _
    ByVal pwfltr As Long, _
    ByVal dwCallback As Long, _
    ByVal dwInstance As Long, _
    ByVal fdwOpen As Long _
) As Long

Private Declare Function acmStreamSize Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal cbInput As Long, _
    pdwOutputBytes As Long, _
    ByVal fdwSize As Long _
) As Long

Private Declare Function acmStreamConvert Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwConvert As Long _
) As Long

Private Declare Function acmStreamReset Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal fdwReset As Long _
) As Long

Private Declare Function acmStreamClose Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal fdwClose As Long _
) As Long

Private Declare Function mmioClose Lib "winmm" ( _
    ByVal hmmio As Long, _
    ByVal uFlags As Long _
) As Long

Private Declare Function mmioDescend Lib "winmm" ( _
    ByVal hmmio As Long, _
    lpck As MMCKINFO, _
    lpckParent As MMCKINFO, _
    ByVal uFlags As Long _
) As Long

Private Declare Function mmioDescendParent Lib "winmm" _
Alias "mmioDescend" ( _
    ByVal hmmio As Long, _
    lpck As MMCKINFO, _
    ByVal x As Long, _
    ByVal uFlags As Long _
) As Long

Private Declare Function mmioOpen Lib "winmm" _
Alias "mmioOpenA" ( _
    ByVal szFileName As String, _
    lpmmioinfo As MMIOINFO, _
    ByVal dwOpenFlags As Long _
) As Long

Private Declare Function mmioSeek Lib "winmm" ( _
    ByVal hmmio As Long, _
    ByVal lOffset As Long, _
    ByVal iOrigin As Long _
) As Long

Private Declare Function mmioStringToFOURCC Lib "winmm" _
Alias "mmioStringToFOURCCA" ( _
    ByVal sz As String, _
    ByVal uFlags As Long _
) As Long

Private Type ACMSTREAMHEADER
    cbStruct                    As Long
    fdwStatus                   As Long
    dwUser                      As Long
    pbSrc                       As Long
    cbSrcLength                 As Long
    cbSrcLengthUsed             As Long
    dwSrcUser                   As Long
    pbDst                       As Long
    cbDstLength                 As Long
    cbDstLengthUsed             As Long
    dwDstUser                   As Long
    dwReservedDriver(9)         As Long
End Type

Private Type MMIOINFO
   dwFlags                      As Long
   fccIOProc                    As Long
   pIOProc                      As Long
   wErrorRet                    As Long
   htask                        As Long
   cchBuffer                    As Long
   pchBuffer                    As String
   pchNext                      As String
   pchEndRead                   As String
   pchEndWrite                  As String
   lBufOffset                   As Long
   lDiskOffset                  As Long
   adwInfo(4)                   As Long
   dwReserved1                  As Long
   dwReserved2                  As Long
   hmmio                        As Long
End Type

Private Type WAVE_FORMAT
    wFormatTag                  As Integer
    wChannels                   As Integer
    dwSampleRate                As Long
    dwBytesPerSec               As Long
    wBlockAlign                 As Integer
    wBitsPerSample              As Integer
End Type

Private Type MMCKINFO
   ckid                         As Long
   ckSize                       As Long
   fccType                      As Long
   dwDataOffset                 As Long
   dwFlags                      As Long
End Type

Private Type CHUNKINFO
    Start                       As Long
    Length                      As Long
End Type

Private Type MMWAVEFORMATEX
    wFormatTag                  As Integer
    nChannels                   As Integer
    nSamplesPerSec              As Long
    nAvgBytesPerSec             As Long
    nBlockAlign                 As Integer
    wBitsPerSample              As Integer
    cbSize                      As Integer
End Type

Private Type hFile
    handle                      As Long
    path                        As String
End Type

Private Enum HACMSTREAM
    INVALID_STREAM_HANDLE = 0
End Enum

Private Enum ACM_STREAMSIZEF
    ACM_STREAMSIZEF_DESTINATION = &H1
    ACM_STREAMSIZEF_SOURCE = &H0
    ACM_STREAMSIZEF_QUERYMASK = &HF
End Enum

Private Enum ACM_STREAMCONVERTF
    ACM_STREAMCONVERTF_BLOCKALIGN = &H4
    ACM_STREAMCONVERTF_START = &H10
    ACM_STREAMCONVERTF_END = &H20
End Enum

Private Enum FILE_OPEN_METHOD
    CREATE_NEW = 1
    CREATE_ALWAYS = 2
    OPEN_EXISTING = 3
    OPEN_ALWAYS = 4
End Enum

Private Enum FILE_SHARE_RIGHTS
    FILE_SHARE_READ = &H1
    FILE_SHARE_WRITE = &H2
End Enum

Private Enum FILE_ACCESS_RIGHTS
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
End Enum

Private Enum SEEK_METHOD
    FILE_BEGIN = 0
    FILE_CURRENT = 1
    FILE_END = 2
End Enum

Public Enum SND_RESULT
    SND_ERR_SUCCESS
    SND_ERR_INVALID_SOURCE
    SND_ERR_INVALID_OUTPUT
    SND_ERR_INTERNAL
    SND_ERR_OUT_OF_RANGE
    SND_ERR_END_OF_STREAM
    SND_ERR_INVALID_TAG
    SND_ERR_INVALID_PARAM
    SND_ERR_TOO_BIG
    SND_ERR_NEED_MORE
    SND_ERR_UNKNOWN
End Enum

Public Enum SND_SEEK_MODE
    SND_SEEK_PERCENT
    SND_SEEK_SECONDS
End Enum

Private Const MMIO_READ         As Long = &H0
Private Const MMIO_FINDCHUNK    As Long = &H10
Private Const MMIO_FINDRIFF     As Long = &H20

Private Const INVALID_HANDLE    As Long = -1

Private Const SEEK_CUR          As Long = 1

Private Const WAVE_FORMAT_PCM   As Long = 1


' Stream

Private hFWave                  As hFile
Private cnkData                 As CHUNKINFO
Private cnkInfo                 As CHUNKINFO
Private udtWFXIn                As MMWAVEFORMATEX
Private udtWFXOut               As MMWAVEFORMATEX
Private btWFX()                 As Byte

' ACM

Private Const OUTPUT_BUFFER_MS  As Long = 500

Private hStream                 As HACMSTREAM

Private btInput()               As Byte
Private intOutput()             As Integer

Private lngInputLen             As Long
Private lngOutputLen            As Long

Private lngPosInBuffer          As Long
Private lngBufferData           As Long

Private lngFilePositionMS       As Long

Private blnEndOfStream          As Boolean
Private blnFirst                As Boolean

Private lngKeepInBuffer         As Long

Private Sub Class_Initialize()
    hStream = INVALID_STREAM_HANDLE
    hFWave.handle = INVALID_HANDLE
End Sub

Private Sub Class_Terminate()
    StreamClose
End Sub

Public Property Get StreamDescription( _
) As String

    StreamDescription = Description
End Property

Public Property Get EndOfStream( _
) As Boolean

    If blnEndOfStream Then
        If lngPosInBuffer = lngBufferData Then
            EndOfStream = True
        End If
    End If
End Property

Public Function StreamExtensions( _
) As String()

    StreamExtensions = Split(Extensions, ";")
End Function

Public Function StreamOpen( _
    ByVal Source As String _
) As SND_RESULT

    If Not IsValidFile(Source) Then
        StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    StreamClose

    ' find WAV Chunks "data" and "fmt "
    cnkData = GetWavChunkPos(Source, "data")
    cnkInfo = GetWavChunkPos(Source, "fmt ")

    ' valid Chunks?
    If cnkData.Start = 0 Then
        StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If cnkInfo.Start = 0 Then
        StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If cnkInfo.Length < 16 Then
        StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    hFWave = FileOpen(Source, _
                      GENERIC_READ, _
                      FILE_SHARE_READ)

    If hFWave.handle = INVALID_HANDLE Then
        StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    ' shrink data chunks with ilegal length to file length
    If FileLength(hFWave) < (cnkData.Start + cnkData.Length) Then
        cnkData.Length = FileLength(hFWave) - cnkData.Start
    End If

    ' read info chunk
    ReDim btWFX(cnkInfo.Length - 1) As Byte
    FileSeek hFWave, cnkInfo.Start, FILE_BEGIN
    FileRead hFWave, VarPtr(btWFX(0)), cnkInfo.Length

    CpyMem udtWFXIn, btWFX(0), Len(udtWFXIn)

    ' seek to the beginning of the audio data
    FileSeek hFWave, cnkData.Start, FILE_BEGIN

    ' init the Audio Compression Manager
    If Not InitConversion(True) Then
        StreamOpen = SND_ERR_INTERNAL
        StreamClose
        Exit Function
    End If

    StreamOpen = SND_ERR_SUCCESS
End Function

Public Function StreamClose( _
) As SND_RESULT

    If hFWave.handle = INVALID_HANDLE Then
        StreamClose = SND_ERR_INVALID_SOURCE
    Else
        CloseConverter
        FileClose hFWave
        lngFilePositionMS = 0
        StreamClose = SND_ERR_SUCCESS
    End If
End Function

' StreamRead returns exactly as many bytes as wanted,
' as long as the end of the stream isn't reached
Public Function StreamRead( _
    ByVal buffer_ptr As Long, _
    ByVal buffer_len As Long, _
    ByRef buffer_read As Long _
) As SND_RESULT

    StreamRead = SND_ERR_SUCCESS

    buffer_read = 0

    Do While buffer_read < buffer_len

        ' PCM buffer empty
        If lngBufferData = 0 Then
            If Not FillBuffer Then
                StreamRead = SND_ERR_END_OF_STREAM
                Exit Do
            End If

        ' not enough data in the PCM buffer
        ElseIf (lngBufferData - lngPosInBuffer) < (buffer_len - buffer_read) Then
            If 0 < (lngBufferData - lngPosInBuffer) Then

                If 0 = IsBadReadPtr(ByVal VarPtr(intOutput(0)) + lngPosInBuffer, _
                                    lngBufferData - lngPosInBuffer) Then

                    If 0 = IsBadWritePtr(ByVal buffer_ptr + buffer_read, _
                                         lngBufferData - lngPosInBuffer) Then

                        CpyMem ByVal buffer_ptr + buffer_read, _
                               ByVal VarPtr(intOutput(0)) + lngPosInBuffer, _
                               lngBufferData - lngPosInBuffer

                    End If

                End If

                buffer_read = buffer_read + (lngBufferData - lngPosInBuffer)
            End If

            If Not FillBuffer Then
                StreamRead = SND_ERR_END_OF_STREAM
                Exit Do
            End If

        ' enough data in the PCM buffer
        Else
            If 0 = IsBadReadPtr(ByVal VarPtr(intOutput(0)) + lngPosInBuffer, _
                                buffer_len - buffer_read) Then

                If 0 = IsBadWritePtr(ByVal buffer_ptr + buffer_read, _
                                     buffer_len - buffer_read) Then

                    CpyMem ByVal buffer_ptr + buffer_read, _
                           ByVal VarPtr(intOutput(0)) + lngPosInBuffer, _
                           buffer_len - buffer_read

                End If

            End If

            lngPosInBuffer = lngPosInBuffer + (buffer_len - buffer_read)
            buffer_read = buffer_read + (buffer_len - buffer_read)

        End If

    Loop
    
    lngFilePositionMS = lngFilePositionMS + (buffer_read / udtWFXOut.nAvgBytesPerSec * 1000)
End Function

Public Function StreamSeek( _
    ByVal value As Long, _
    ByVal seek_mode As SND_SEEK_MODE _
) As SND_RESULT

    Dim lngBytes    As Long

    If hFWave.handle = INVALID_HANDLE Then
        StreamSeek = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    Select Case seek_mode

        Case SND_SEEK_PERCENT
            If value < 0 Or value > 99 Then
                StreamSeek = SND_ERR_OUT_OF_RANGE
                Exit Function
            End If

            lngBytes = value / 100 * cnkData.Length

        Case SND_SEEK_SECONDS
            If value < 0 Or value > (Duration / 1000) Then
                StreamSeek = SND_ERR_OUT_OF_RANGE
                Exit Function
            End If

            lngBytes = udtWFXIn.nAvgBytesPerSec * value

    End Select

    If value = 0 Then
        FileSeek hFWave, cnkData.Start, FILE_BEGIN
    Else
        lngBytes = AlignBytes(lngBytes) + cnkData.Start
        FileSeek hFWave, lngBytes, FILE_BEGIN
    End If

    ' reset ACM stream to clear buffers of codecs
    lngFilePositionMS = (FilePosition(hFWave) - cnkData.Start) / udtWFXIn.nAvgBytesPerSec * 1000
    ResetConverter
    
    StreamSeek = SND_ERR_SUCCESS
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Position( _
) As Long

    Position = lngFilePositionMS - OUTPUT_BUFFER_MS
End Property

Public Property Get Duration( _
) As Long

    Duration = (cnkData.Length) / udtWFXIn.nAvgBytesPerSec * 1000
End Property

Public Property Get BitsPerSample( _
) As Integer

    BitsPerSample = 16
End Property

Public Property Get BitsPerSecond( _
) As Long

    BitsPerSecond = udtWFXIn.nAvgBytesPerSec * 8
End Property

Public Property Get Channels( _
) As Integer

    Channels = udtWFXIn.nChannels
End Property

Public Property Get SamplesPerSecond( _
) As Long

    SamplesPerSecond = udtWFXIn.nSamplesPerSec
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Private Function InitConversion( _
    Optional Force16Bit As Boolean = False _
) As Boolean

    Dim mmr     As Long

    If hStream <> INVALID_STREAM_HANDLE Then
        CloseConverter
    End If

    udtWFXOut = udtWFXIn

    If udtWFXOut.wBitsPerSample < 8 Then
        udtWFXOut.wBitsPerSample = 8
    ElseIf udtWFXOut.wBitsPerSample > 8 Then
        udtWFXOut.wBitsPerSample = 16
    End If

    If Force16Bit Then
        udtWFXOut.wBitsPerSample = 16
    End If

    With udtWFXOut
        udtWFXOut = CreateWFX(.nSamplesPerSec, _
                              .nChannels, _
                              .wBitsPerSample)
    End With

    mmr = acmStreamOpen(hStream, 0, _
                        btWFX(0), udtWFXOut, _
                        0, 0, 0, 0)

    If mmr <> 0 Then
        If Force16Bit Then Exit Function

        If udtWFXOut.wBitsPerSample = 16 Then
            udtWFXOut.wBitsPerSample = 8
        Else
            udtWFXOut.wBitsPerSample = 16
        End If

        mmr = acmStreamOpen(hStream, 0, _
                            btWFX(0), udtWFXOut, _
                            0, 0, 0, 0)

        If mmr <> 0 Then Exit Function
    End If

    ' set size of output buffer
    lngOutputLen = OUTPUT_BUFFER_MS / 1000 * udtWFXOut.nAvgBytesPerSec

    ' needed size of input buffer to fill the output buffer
    mmr = acmStreamSize(hStream, _
                        lngOutputLen, _
                        lngInputLen, _
                        ACM_STREAMSIZEF_DESTINATION)

    If mmr <> 0 Then
        acmStreamClose hStream, 0
        hStream = 0
        Exit Function
    End If

    ReDim intOutput(lngOutputLen / 2 - 1) As Integer
    ReDim btInput(lngInputLen - 1) As Byte

    blnEndOfStream = False
    blnFirst = True

    lngKeepInBuffer = 0

    InitConversion = True
End Function

Private Function Convert( _
    ByVal lngInLen As Long, _
    ByVal lngOutLen As Long, _
    lngInUsed As Long, _
    lngOutUsed As Long, _
    Optional ByVal LastConversion As Boolean = False _
) As Boolean

    Dim lngFlags    As Long
    Dim udtHdr      As ACMSTREAMHEADER

    lngFlags = ACM_STREAMCONVERTF_BLOCKALIGN

    If blnFirst Then _
        lngFlags = lngFlags Or ACM_STREAMCONVERTF_START
    If LastConversion Then
        lngFlags = lngFlags Or ACM_STREAMCONVERTF_END
    End If

    With udtHdr
        .cbStruct = Len(udtHdr)
        .cbSrcLength = lngInLen
        .cbDstLength = lngOutLen
        .pbDst = VarPtr(intOutput(0))
        .pbSrc = VarPtr(btInput(0))
    End With

    acmStreamPrepareHeader hStream, udtHdr, 0

    blnFirst = False

    If 0 = acmStreamConvert(hStream, udtHdr, lngFlags) Then
        With udtHdr
            lngInUsed = .cbSrcLengthUsed
            lngOutUsed = .cbDstLengthUsed

            lngKeepInBuffer = .cbSrcLength - .cbSrcLengthUsed
        End With

        If lngKeepInBuffer > 0 Then
            ' codec didn't use all the input bytes,
            ' move them to the first index of the input buffer
            ' to decode them with the next call to convert()
            CpyMem btInput(0), _
                   btInput(lngInLen - lngKeepInBuffer), _
                   lngKeepInBuffer
        End If

        Convert = True
    End If

    acmStreamUnprepareHeader hStream, udtHdr, 0
End Function

Private Function FillBuffer( _
) As Boolean

    Dim lngRead     As Long
    Dim lngWritten  As Long
    Dim udeRet      As SND_RESULT

    If blnEndOfStream Then
        'If lngPosInBuffer >= lngBufferData Then
            lngBufferData = 0
            lngPosInBuffer = 0
            ZeroMem intOutput(0), lngOutputLen
            ZeroMem btInput(0), lngInputLen
            Exit Function
        'Else
        '    FillBuffer = True
        '    Exit Function
        'End If
    End If

    ' get data from WAV
    udeRet = ReadWAVData(VarPtr(btInput(lngKeepInBuffer)), _
                         lngInputLen - lngKeepInBuffer, _
                         lngRead)

    If udeRet <> SND_ERR_SUCCESS Then
        ' either read error or end of file
        blnEndOfStream = True
    End If

    Convert lngRead + lngKeepInBuffer, _
            lngOutputLen, _
            lngRead, lngWritten, _
            blnEndOfStream

    lngPosInBuffer = 0
    lngBufferData = lngWritten

    FillBuffer = True
End Function

Private Sub ResetConverter()
    If hStream = INVALID_STREAM_HANDLE Then
        Exit Sub
    End If

    CloseConverter

    acmStreamOpen hStream, 0, _
                  btWFX(0), udtWFXOut, _
                  0, 0, 0, 0

    lngOutputLen = OUTPUT_BUFFER_MS / 1000 * udtWFXOut.nAvgBytesPerSec

    acmStreamSize hStream, _
                  lngOutputLen, _
                  lngInputLen, _
                  ACM_STREAMSIZEF_DESTINATION

    ReDim intOutput(lngOutputLen / 2 - 1) As Integer
    ReDim btInput(lngInputLen - 1) As Byte

    blnEndOfStream = False
    blnFirst = True

    lngKeepInBuffer = 0
End Sub

Private Function CloseConverter( _
) As Boolean

    On Error Resume Next

    acmStreamClose hStream, 0
    hStream = INVALID_STREAM_HANDLE

    ZeroMem btInput(0), lngInputLen
    ZeroMem intOutput(0), lngOutputLen

    blnEndOfStream = False
    lngBufferData = 0
    lngPosInBuffer = 0
    lngInputLen = 0
    lngOutputLen = 0
    lngKeepInBuffer = 0

    CloseConverter = True
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Private Function ReadWAVData( _
    ByVal data_ptr As Long, _
    ByVal data_len As Long, _
    data_read As Long _
) As SND_RESULT

    ReadWAVData = SND_ERR_SUCCESS

    If hFWave.handle = INVALID_HANDLE Then
        ReadWAVData = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If FilePosition(hFWave) > (cnkData.Start + cnkData.Length) Then
        ' end of file reached
        ReadWAVData = SND_ERR_END_OF_STREAM
        data_read = 0
        Exit Function
    End If

    If FilePosition(hFWave) + data_len > (cnkData.Start + cnkData.Length) Then
        ' almost at the end of the file,
        ' but reached after this read
        data_len = (cnkData.Start + cnkData.Length) - FilePosition(hFWave)
        ReadWAVData = SND_ERR_END_OF_STREAM
    End If

    data_read = FileRead(hFWave, data_ptr, data_len)
End Function

' finds a chunk in a WAV container
Private Function GetWavChunkPos( _
    ByVal strFile As String, _
    ByVal strChunk As String _
) As CHUNKINFO

    Dim hMmioIn             As Long
    Dim lngRet              As Long
    Dim mmckinfoParentIn    As MMCKINFO
    Dim mmckinfoSubchunkIn  As MMCKINFO
    Dim mmioinf             As MMIOINFO

    hMmioIn = mmioOpen(strFile, mmioinf, MMIO_READ)
    If hMmioIn = 0 Then
        Exit Function
    End If

    mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
    lngRet = mmioDescendParent(hMmioIn, _
                               mmckinfoParentIn, _
                               0, _
                               MMIO_FINDRIFF)

    If Not (lngRet = 0) Then
        mmioClose hMmioIn, 0
        Exit Function
    End If

    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC(strChunk, 0)
    lngRet = mmioDescend(hMmioIn, _
                         mmckinfoSubchunkIn, _
                         mmckinfoParentIn, _
                         MMIO_FINDCHUNK)

    If Not (lngRet = 0) Then
        mmioClose hMmioIn, 0
        Exit Function
    End If

    GetWavChunkPos.Start = mmioSeek(hMmioIn, 0, SEEK_CUR)
    GetWavChunkPos.Length = mmckinfoSubchunkIn.ckSize

    mmioClose hMmioIn, 0
End Function

' when seeking in WAV files you need to align
' the position seeked to on the Block Align of
' the audio data
Private Function AlignBytes( _
    ByVal bytes As Long _
) As Long

    AlignBytes = bytes - (bytes Mod udtWFXIn.nBlockAlign)
End Function

Private Function CreateWFX( _
    sr As Long, _
    chs As Integer, _
    bps As Integer _
) As MMWAVEFORMATEX

    With CreateWFX
        .wFormatTag = WAVE_FORMAT_PCM
        .nChannels = chs
        .nSamplesPerSec = sr
        .wBitsPerSample = bps
        .nBlockAlign = chs * (bps / 8)
        .nAvgBytesPerSec = sr * .nBlockAlign
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''

Private Function IsValidFile( _
    ByVal strFile As String _
) As Boolean

    Dim hInp    As hFile

    hInp = FileOpen(strFile, _
                    GENERIC_READ, _
                    FILE_SHARE_READ)

    IsValidFile = hInp.handle <> INVALID_HANDLE
    FileClose hInp
End Function

Private Function FileOpen( _
    ByVal strFile As String, _
    Optional access As FILE_ACCESS_RIGHTS = GENERIC_READ Or GENERIC_WRITE, _
    Optional share As FILE_SHARE_RIGHTS = FILE_SHARE_READ Or FILE_SHARE_WRITE, _
    Optional method As FILE_OPEN_METHOD = OPEN_EXISTING _
) As hFile

    FileOpen.handle = CreateFile(strFile, _
                                 access, _
                                 share, _
                                 ByVal 0&, _
                                 method, _
                                 0, 0)

    FileOpen.path = strFile
End Function

Private Sub FileClose( _
    filehandle As hFile _
)

    CloseHandle filehandle.handle
    filehandle.handle = INVALID_HANDLE
    filehandle.path = vbNullString
End Sub

Private Function FileRead( _
    filehandle As hFile, _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As Long

    Dim dwRead  As Long
    Dim lngRet  As Long

    If filehandle.handle = INVALID_HANDLE Then
        Exit Function
    End If

    lngRet = ReadFile(filehandle.handle, _
                      ByVal ptr, _
                      bytes, _
                      dwRead, _
                      0&)

    If lngRet = 1 Then
        FileRead = dwRead
    Else
        FileRead = -1
    End If
End Function

Private Function FileWrite( _
    filehandle As hFile, _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As Long

    Dim dwWritten   As Long
    Dim lngRet      As Long

    If filehandle.handle = INVALID_HANDLE Then
        Exit Function
    End If

    lngRet = WriteFile(filehandle.handle, _
                       ByVal ptr, _
                       bytes, _
                       dwWritten, _
                       0&)

    If lngRet = 1 Then
        FileWrite = dwWritten
    Else
        FileWrite = -1
    End If
End Function

Private Function FileSeek( _
    filehandle As hFile, _
    ByVal bytes As Long, _
    ByVal method As SEEK_METHOD _
) As Long

    FileSeek = SetFilePointer(filehandle.handle, _
                              bytes, _
                              0, _
                              method)
End Function

Private Function FilePosition( _
    filehandle As hFile _
) As Long

    FilePosition = FileSeek(filehandle, _
                            0, _
                            FILE_CURRENT)
End Function

Private Function FileLength( _
    filehandle As hFile _
) As Long

    FileLength = GetFileSize(filehandle.handle, 0)
End Function

Private Function FileEnd( _
    filehandle As hFile _
) As Boolean

    FileEnd = FilePosition(filehandle) >= FileLength(filehandle)
End Function
