VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaveOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**************************************'
' clsWaveOut                           '
'                                      '
' play PCM data with the WaveOut API   '
'**************************************'


' Buffer in memory:
'
'          Pointer
'          ^
' +--------+------------+---------...
' |  hMem  |   WAVEHDR  | PCM
' +--------+------------+---------...
'  4 Bytes    32 Bytes    x Bytes


Private Const WAVE_FORMAT_PCM   As Long = 1

Private Const MMSYSERR_BASE     As Long = 0
Private Const WAVERR_BASE       As Long = 32

Private Const MAXPNAMELEN       As Long = 32
Private Const MAXERRORLENGTH    As Long = 256


Private Type WAVEHDR
    lpData              As Long
    dwBufferLength      As Long
    dwBytesRecorded     As Long
    dwUser              As Long
    dwFlags             As Long
    dwLoops             As Long
    lpNext              As Long
    Reserved            As Long
End Type

Private Type WAVEOUTCAPS
    wMid                As Integer
    wPid                As Integer
    vDriverVersion      As Long
    szPname(MAXPNAMELEN - 1) As Byte
    dwFormats           As Long
    wChannels           As Integer
    dwSupport           As Long
End Type

Private Type MMTIME
    wType               As Long
    u                   As Long
    x                   As Long
End Type

Private Type WAVEFORMATEX
    wFormatTag          As Integer
    nChannels           As Integer
    nSamplesPerSec      As Long
    nAvgBytesPerSec     As Long
    nBlockAlign         As Integer
    wBitsPerSample      As Integer
    cbSize              As Integer
End Type

Private Enum TIME_FORMAT
    TIME_MS = &H1
    TIME_SAMPLES = &H2
    TIME_BYTES = &H4
    TIME_SMPTE = &H8
    TIME_MIDI = &H10
    TIME_MIDI_TICKS = &H20
End Enum

Private Enum WAVEFORM_STATUS
    MM_WOM_OPEN = &H3BB
    MM_WOM_CLOSE = &H3BC
    MM_WOM_DONE = &H3BD
End Enum

Private Enum MMSYS_ERROR
    MMSYSERR_NOERROR = 0                         ' no error */
    MMSYSERR_ERROR = (MMSYSERR_BASE + 1)         ' unspecified error */
    MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)   ' device ID out of range */
    MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)    ' driver failed enable */
    MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)     ' device already allocated */
    MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)   ' device handle is invalid */
    MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)      ' no device driver present */
    MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)         ' memory allocation error */
    MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)  ' function isn't supported */
    MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)     ' error value out of range */
    MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10)    ' invalid flag passed */
    MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)   ' invalid parameter passed */
    MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12)   ' handle being used */
                                                 ' simultaneously on another */
                                                 ' thread (eg callback) */
    MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13) ' specified alias not found */
    MMSYSERR_BADDB = (MMSYSERR_BASE + 14)        ' bad registry database */
    MMSYSERR_KEYNOTFOUND = (MMSYSERR_BASE + 15)  ' registry key not found */
    MMSYSERR_READERROR = (MMSYSERR_BASE + 16)    ' registry read error */
    MMSYSERR_WRITEERROR = (MMSYSERR_BASE + 17)   ' registry write error */
    MMSYSERR_DELETEERROR = (MMSYSERR_BASE + 18)  ' registry delete error */
    MMSYSERR_VALNOTFOUND = (MMSYSERR_BASE + 19)  ' registry value not found */
    MMSYSERR_NODRIVERCB = (MMSYSERR_BASE + 20)   ' driver does not call DriverCallback */
    MMSYSERR_LASTERROR = (MMSYSERR_BASE + 20)    ' last error in range */
End Enum

Private Enum MMSYS_CALLBACK
    CALLBACK_TYPEMASK = &H70000
    CALLBACK_NULL = &H0
    CALLBACK_WINDOW = &H10000
    CALLBACK_TASK = &H20000
    CALLBACK_FUNCTION = &H30000
    CALLBACK_THREAD = (CALLBACK_TASK)
    CALLBACK_EVENT = &H50000
End Enum

Private Enum WAVE_ERR
    WAVERR_BADFORMAT = (WAVERR_BASE + 0)    ' unsupported wave format */
    WAVERR_STILLPLAYING = (WAVERR_BASE + 1) ' still something playing */
    WAVERR_UNPREPARED = (WAVERR_BASE + 2)   ' header not prepared */
    WAVERR_SYNC = (WAVERR_BASE + 3)         ' device is synchronous */
    WAVERR_LASTERROR = (WAVERR_BASE + 3)    ' last error in range */
End Enum

Private Enum WAVEOUT_FLAGS
     WAVE_FORMAT_QUERY = &H1
     WAVE_ALLOWSYNC = &H2
     WAVE_MAPPED = &H4
     WAVE_FORMAT_DIRECT = &H8
     WAVE_FORMAT_DIRECT_QUERY = (WAVE_FORMAT_QUERY Or WAVE_FORMAT_DIRECT)
End Enum

Private Enum WHDR_FLAGS
    WHDR_DONE = &H1
    WHDR_PREPARED = &H2
    WHDR_BEGINLOOP = &H4
    WHDR_ENDLOOP = &H8
    WHDR_INQUEUE = &H10
End Enum

Private Enum WAVECAPS_FLAGS
    WAVECAPS_PITCH = &H1
    WAVECAPS_PLAYBACKRATE = &H2
    WAVECAPS_VOLUME = &H4
    WAVECAPS_LRVOLUME = &H8
    WAVECAPS_SYNC = &H10
    WAVECAPS_SAMPLEACCURATE = &H20
    WAVECAPS_DIRECTSOUND = &H40
End Enum

Public Enum WO_STATUS
    WO_PLAYING = 0
    WO_PAUSING
    WO_STOPPED
End Enum

Private Enum MEM_ALLOC_FLAGS
    GHND = &H40
    GMEM_DDESHARE = &H2000
    GMEM_DISCARDABLE = &H100
    GMEM_FIXED = &H0
    GMEM_MOVEABLE = &H2
    GMEM_NOCOMPACT = &H10
    GMEM_NODISCARD = &H20
    GMEM_ZEROINIT = &H40
    GPTR = &H42
End Enum

'-Selfsub declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H38                                  'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data index of the shutdown flag
Private Const IDX_HWND      As Long = 2                                     'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9                                     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11                                    'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12                                    'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13                                    'Thunk data index of the User-defined callback parameter data index

Private z_ScMem             As Long                                         'Thunk base address
Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address collection

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Declare Function waveOutGetPosition Lib "winmm.dll" ( _
    ByVal hWaveOut As Long, _
    lpInfo As Any, _
    ByVal uSize As Long _
) As Long

Private Declare Function waveOutOpen Lib "winmm.dll" ( _
    hWaveOut As Long, _
    ByVal uDeviceID As Long, _
    format As Any, _
    ByVal dwCallback As Long, _
    ByVal dwInstance As Long, _
    ByVal dwFlags As Long _
) As Long

Private Declare Function waveOutPrepareHeader Lib "winmm.dll" ( _
    ByVal hWaveOut As Long, _
    lpWaveInHdr As Any, _
    ByVal uSize As Long _
) As Long

Private Declare Function waveOutReset Lib "winmm.dll" ( _
    ByVal hWaveOut As Long _
) As Long

Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" ( _
    ByVal hWaveOut As Long, _
    lpWaveInHdr As Any, _
    ByVal uSize As Long _
) As Long

Private Declare Function waveOutClose Lib "winmm.dll" ( _
    ByVal hWaveOut As Long _
) As Long

Private Declare Function waveOutGetDevCaps Lib "winmm.dll" _
Alias "waveOutGetDevCapsA" ( _
    ByVal uDeviceID As Long, _
    lpCaps As Any, _
    ByVal uSize As Long _
) As Long

Private Declare Function waveOutGetNumDevs Lib "winmm.dll" ( _
) As Long

Private Declare Function waveOutGetErrorText Lib "winmm.dll" _
Alias "waveInGetErrorTextA" ( _
    ByVal err As Long, _
    ByVal lpText As String, _
    ByVal uSize As Long _
) As Long

Private Declare Function waveOutWrite Lib "winmm.dll" ( _
    ByVal hWaveOut As Long, _
    lpWaveOutHdr As Any, _
    ByVal uSize As Long _
) As Long

Private Declare Function waveOutRestart Lib "winmm.dll" ( _
    ByVal hWaveOut As Long _
) As Long

Private Declare Function waveOutPause Lib "winmm.dll" ( _
    ByVal hWaveOut As Long _
) As Long

Private Declare Function waveOutGetID Lib "winmm.dll" ( _
    ByVal hWaveOut As Long, _
    ByRef puDeviceID As Long _
) As Long

Private Declare Function waveOutSetVolume Lib "winmm.dll" ( _
    ByVal hWaveOut As Long, _
    ByVal dwVolume As Long _
) As Long

Private Declare Function waveOutGetVolume Lib "winmm.dll" ( _
    ByVal hWaveOut As Long, _
    ByRef pdwVolume As Long _
) As Long

Private Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As Long, _
    ByVal dwBytes As Long _
) As Long

Private Declare Function GlobalLock Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Function GlobalFree Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Function GlobalUnlock Lib "kernel32" ( _
    ByVal hMem As Long _
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

Private Declare Function PostMessage Lib "user32" _
Alias "PostMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long

Private Declare Function CallWindowProc Lib "user32" _
Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long

Private Declare Function CreateWindowEx Lib "user32" _
Alias "CreateWindowExA" ( _
    ByVal dwExStyle As Long, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hWndParent As Long, _
    ByVal hMenu As Long, _
    ByVal hInstance As Long, _
    ByVal lpParam As Long _
) As Long

Private Declare Function DestroyWindow Lib "user32" ( _
    ByVal hwnd As Long _
) As Long


Public Event BufferDone(ByVal userdata As Long)
Public Event StatusChanged(ByVal status As WO_STATUS)


Private wfx_out     As WAVEFORMATEX

Private blnReady    As Boolean
Private udeStatus   As WO_STATUS

Private hWO         As Long
Private hFakeWnd    As Long
Private pOldWndProc As Long

Private lngBufCnt   As Long

Private lngAddTime  As Long

''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''

' number of buffers in the WaveOut Buffer Queue
Public Property Get Buffers( _
) As Long

    Buffers = lngBufCnt
End Property

Public Property Get VolumeForDev( _
    ByVal idx As Integer _
) As Long

    Dim lngVolume   As Long
    
    waveOutGetVolume idx, lngVolume
    
    VolumeForDev = (lngVolume And &HFFFF&)
End Property

Public Property Let VolumeForDev( _
    ByVal idx As Integer, _
    ByVal value As Long _
)

    Dim lngVol As Long

    value = (value And &HFFFF&)

    RtlMoveMemory VarPtr(lngVol) + 0, VarPtr(value), 2
    RtlMoveMemory VarPtr(lngVol) + 2, VarPtr(value), 2

    waveOutSetVolume idx, lngVol
End Property

Public Property Get Volume( _
) As Long

    Dim lngVolume   As Long

    waveOutGetVolume hWO, lngVolume

    ' one 16 bit word for one channel
    Volume = (lngVolume And &HFFFF&)
End Property

Public Property Let Volume( _
    ByVal lngVolume As Long _
)

    lngVolume = (lngVolume And &HFFFF&)

    RtlMoveMemory VarPtr(lngVolume) + 0, VarPtr(lngVolume), 2
    RtlMoveMemory VarPtr(lngVolume) + 2, VarPtr(lngVolume), 2

    waveOutSetVolume hWO, lngVolume
End Property

Public Property Get PlayStatus( _
) As WO_STATUS

    PlayStatus = udeStatus
End Property

Public Function GetNumDevs( _
) As Long

    GetNumDevs = waveOutGetNumDevs()
End Function

Public Function GetDevName( _
    ByVal device_id As Long _
) As String

    Dim strDevName  As String
    Dim lngNullChr  As Long
    Dim udtCaps     As WAVEOUTCAPS

    waveOutGetDevCaps device_id, _
                      udtCaps, _
                      Len(udtCaps)

    strDevName = StrConv(udtCaps.szPname, vbUnicode)

    lngNullChr = InStr(strDevName, Chr$(0))
    If lngNullChr < 1 Then lngNullChr = Len(strDevName)

    ' Device Name
    strDevName = Left$(strDevName, lngNullChr - 1) & " "
    ' Driver Version
    strDevName = strDevName & ((udtCaps.vDriverVersion And &HFF00) \ &H100) & "."
    strDevName = strDevName & ((udtCaps.vDriverVersion And &HFF) \ &H1)

    GetDevName = strDevName
End Function

Public Function Init( _
    ByVal samplerate As Long, _
    ByVal Channels As Integer, _
    ByVal BitsPerSample As Integer, _
    Optional ByVal device_id As Long = -1 _
) As Boolean

    Dim udeErr  As MMSYS_ERROR

    ' just to be on the safe side
    Deinit

    ' I don't think any other sample widths then 8/16/24
    ' are allowed for PCM described with a WaveFormatEx structure.
    If BitsPerSample < 8 Or BitsPerSample > 24 Then
        Exit Function
    End If

    If samplerate < 1 Or samplerate > 100000 Then
        Exit Function
    End If

    ' only support for mono/stereo
    ' (are there more channels defined for WaveFormatEx?)
    If Channels < 1 Or Channels > 2 Then
        Exit Function
    End If

    With wfx_out
        .wFormatTag = WAVE_FORMAT_PCM
        '
        .wBitsPerSample = BitsPerSample
        .nSamplesPerSec = samplerate
        .nChannels = Channels
        '
        .nBlockAlign = .nChannels * (.wBitsPerSample / 8)
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
    End With

    lngAddTime = 0

    ' create an invisible window for WaveOut callbacks
    hFakeWnd = CreateFakeWnd()
    If hFakeWnd = 0 Then
        blnReady = False
        Exit Function
    End If

    udeErr = waveOutOpen(hWO, _
                         device_id, _
                         wfx_out, _
                         hFakeWnd, _
                         0, _
                         CALLBACK_WINDOW)

    If udeErr = MMSYSERR_NOERROR Then
        HookFakeWindow

        waveOutPause hWO

        udeStatus = WO_STOPPED
        RaiseEvent StatusChanged(udeStatus)

        blnReady = True
        Init = True
    Else
        DestroyWindow hFakeWnd
        hFakeWnd = 0
        blnReady = False
    End If
End Function

Public Function Deinit( _
) As Boolean

    If Not blnReady Then Exit Function

    waveOutReset hWO
    waveOutClose hWO
    hWO = 0

    UnhookFakeWindow
    DestroyWindow hFakeWnd
    hFakeWnd = 0

    lngAddTime = 0

    blnReady = False

    udeStatus = WO_STOPPED
    RaiseEvent StatusChanged(udeStatus)

    Deinit = True
End Function

Public Function WriteSamples( _
    ByVal buf_ptr As Long, _
    ByVal buf_len As Long, _
    Optional ByVal userdata As Long _
) As Boolean

    Dim pMem    As Long
    Dim udeErr  As MMSYS_ERROR
    Dim udtHdr  As WAVEHDR

    If Not blnReady Then Exit Function
    If buf_len = 0 Then Exit Function

    ' data needs to be aligned on the block align
    If buf_len Mod wfx_out.nBlockAlign Then Exit Function

    ' allocate some memory for WAVEHDR + PCM data
    pMem = AllocBufferMem(buf_len + Len(udtHdr))
    If pMem = 0 Then Exit Function

    With udtHdr
        .dwUser = userdata
        .dwBufferLength = buf_len
        .lpData = pMem + Len(udtHdr)
    End With

    ' copy WAVEHDR to the allocated buffer
    CpyMem ByVal pMem, udtHdr, Len(udtHdr)

    ' copy PCM data to the buffer
    CpyMem ByVal pMem + Len(udtHdr), _
           ByVal buf_ptr, _
           buf_len

    udeErr = waveOutPrepareHeader(hWO, _
                                  ByVal pMem, _
                                  Len(udtHdr))

    udeErr = waveOutWrite(hWO, _
                          ByVal pMem, _
                          Len(udtHdr))

    lngBufCnt = lngBufCnt + 1

    WriteSamples = True
End Function

Public Function Pause( _
) As Boolean

    If Not blnReady Then Exit Function

    waveOutPause hWO
    udeStatus = WO_PAUSING
    RaiseEvent StatusChanged(udeStatus)

    Pause = True
End Function

Public Function Play( _
) As Boolean

    If Not blnReady Then Exit Function

    waveOutRestart hWO
    udeStatus = WO_PLAYING
    RaiseEvent StatusChanged(udeStatus)

    Play = True
End Function

Public Function StopPlayback( _
) As Boolean

    If Not blnReady Then Exit Function

    waveOutReset hWO
    waveOutPause hWO

    lngAddTime = 0
    lngBufCnt = 0
    udeStatus = WO_STOPPED
    RaiseEvent StatusChanged(udeStatus)

    StopPlayback = True
End Function

Public Function ClearBufferQueue( _
) As Boolean

    waveOutReset hWO

    ClearBufferQueue = True
End Function

Public Sub SetElapsed( _
    ByVal lngMS As Long _
)

    If Not blnReady Then Exit Sub

    lngAddTime = Elapsed() - (Elapsed() - lngMS)
End Sub

Public Sub ResetElapsed( _
)

    lngAddTime = -GetWaveOutElapsed
End Sub

Private Function GetWaveOutElapsed() As Long
    If Not blnReady Then Exit Function
    
    Dim udtTime As MMTIME
    
    ' the device mustn't support the selected
    ' time format, in fact waveOutGetPosition can
    ' return any time format
    udtTime.wType = TIME_BYTES

    waveOutGetPosition hWO, udtTime, Len(udtTime)

    Select Case udtTime.wType
        Case TIME_BYTES
            GetWaveOutElapsed = (udtTime.u / wfx_out.nAvgBytesPerSec) * 1000
        Case TIME_SAMPLES
            GetWaveOutElapsed = (udtTime.u / wfx_out.nSamplesPerSec) * 1000
        Case TIME_MS
            GetWaveOutElapsed = udtTime.u
    End Select
End Function

Public Function Elapsed( _
) As Long

    If Not blnReady Then Exit Function

    Elapsed = GetWaveOutElapsed + lngAddTime
End Function

''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''

Private Function AllocBufferMem( _
    ByVal bytes As Long _
) As Long

    Dim hMem    As Long
    Dim pMem    As Long

    ' bytes + 4 to save the memory handle (for FreeBufferMem())
    hMem = GlobalAlloc(GPTR, bytes + 4)
    If hMem = 0 Then Exit Function

    pMem = GlobalLock(hMem)
    If pMem = 0 Then Exit Function

    CpyMem ByVal pMem, hMem, 4

    ' return a pointer to the allocated memory,
    ' but hide the memory handle
    AllocBufferMem = pMem + 4
End Function

Private Sub FreeBufferMem( _
    ByVal buf_ptr As Long _
)

    Dim hMem    As Long

    ' memory handle should be at (pointer - 4)
    CpyMem hMem, ByVal buf_ptr - 4, 4

    GlobalUnlock hMem
    GlobalFree hMem
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''

Private Function CreateFakeWnd( _
) As Long

    CreateFakeWnd = CreateWindowEx(0, "static", _
                                   "clsWaveOut", 0, _
                                   0, 0, 0, 0, _
                                   0, 0, 0, 0)

End Function

Private Sub HookFakeWindow()
    sc_Subclass hFakeWnd, 0, 1, Me, True
    sc_AddMsg hFakeWnd, MM_WOM_DONE, MSG_AFTER
End Sub

Private Sub UnhookFakeWindow()
    sc_UnSubclass hFakeWnd
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    udeStatus = WO_STOPPED
End Sub

Private Sub Class_Terminate()
    Deinit
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''

'-SelfSub code------------------------------------------------------------------------------------
Private Function sc_Subclass( _
    ByVal lng_hWnd As Long, _
    Optional ByVal lParamUser As Long = 0, _
    Optional ByVal nOrdinal As Long = 1, _
    Optional ByVal oCallback As Object = Nothing, _
    Optional ByVal bIdeSafety As Boolean = True _
) As Boolean 'Subclass the specified window handle

'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************

    Const CODE_LEN      As Long = 260                                   'Thunk length in bytes
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))    'Bytes to allocate per thunk, data + code + msg tables
    Const PAGE_RWX      As Long = &H40&                                 'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&                               'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&                               'Release allocated memory flag
    Const IDX_EBMODE    As Long = 3                                     'Thunk data index of the EbMode function address
    Const IDX_CWP       As Long = 4                                     'Thunk data index of the CallWindowProc function address
    Const IDX_SWL       As Long = 5                                     'Thunk data index of the SetWindowsLong function address
    Const IDX_FREE      As Long = 6                                     'Thunk data index of the VirtualFree function address
    Const IDX_BADPTR    As Long = 7                                     'Thunk data index of the IsBadCodePtr function address
    Const IDX_OWNER     As Long = 8                                     'Thunk data index of the Owner object's vTable address
    Const IDX_CALLBACK  As Long = 10                                    'Thunk data index of the callback method address
    Const IDX_EBX       As Long = 16                                    'Thunk code patch index of the thunk data
    Const SUB_NAME      As String = "sc_Subclass"                       'This routine's name
      Dim nAddr         As Long
      Dim nID           As Long
      Dim nMyID         As Long
  
    If IsWindow(lng_hWnd) = 0 Then                                      'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If

    nMyID = GetCurrentProcessId                                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID                              'Get the process ID associated with the window handle
    If nID <> nMyID Then                                                'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
  
    If oCallback Is Nothing Then                                        'If the user hasn't specified the callback owner
        Set oCallback = Me                                              'Then it is me
    End If
  
    nAddr = zAddressOf(oCallback, nOrdinal)                             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                                   'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
    
    If z_Funk Is Nothing Then                                           'If this is the first time through, do the one-time initialization
        Set z_Funk = New Collection                                     'Create the hWnd/thunk-address collection
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960
        z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C
        z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&
        z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F
        z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B
        z_Sc(25) = &H60&: z_Sc(26) = &H10C261
        z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8
        z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&
        z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075
        z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF
        z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF
        z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775
        z_Sc(39) = &H58627404
        z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC
        z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D
        z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&
        z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D
        z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04
        z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985
        z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845
        z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28
        z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34
        z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF
        z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28
        z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073
        z_Sc(64) = &HC328&

        z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")            'Store CallWindowProc function address in the thunk data
        z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")             'Store the SetWindowLong function address in the thunk data
        z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")             'Store the VirtualFree function address in the thunk data
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")          'Store the IsBadCodePtr function address in the thunk data
    End If
  
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)            'Allocate executable memory

    If z_ScMem <> 0 Then                                                'Ensure the allocation succeeded
        On Error GoTo CatchDoubleSub                                    'Catch double subclassing
            z_Funk.Add z_ScMem, "h" & lng_hWnd                          'Add the hWnd/thunk-address to the collection
        On Error GoTo 0
  
        If bIdeSafety Then                                              'If the user wants IDE protection
            z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                'Store the EbMode function address in the thunk data
        End If
    
        z_Sc(IDX_EBX) = z_ScMem                                         'Patch the thunk data address
        z_Sc(IDX_HWND) = lng_hWnd                                       'Store the window handle in the thunk data
        z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                           'Store the address of the before table in the thunk data
        z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4) 'Store the address of the after table in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)                             'Store the callback owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                                      'Store the callback address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                                'Store the lParamUser callback parameter in the thunk data
    
        nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
        If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
            zError SUB_NAME, "SetWindowLong failed, error #" & err.LastDllError
            GoTo ReleaseMemory
        End If
        
        z_Sc(IDX_WNDPROC) = nAddr                                       'Store the original WndProc address in the thunk data
        RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                'Copy the thunk code/data to the allocated memory
        sc_Subclass = True                                              'Indicate success
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & err.LastDllError
    End If
  
    Exit Function                                                       'Exit sc_Subclass

CatchDoubleSub:
    zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE                                 'sc_Subclass has failed after memory allocation, so release the memory
    
End Function

'Terminate all subclassing
Private Sub sc_Terminate()

    Dim i As Long

    If Not (z_Funk Is Nothing) Then                                     'Ensure that subclassing has been started
        With z_Funk
            For i = .Count To 1 Step -1                                 'Loop through the collection of window handles in reverse order
                z_ScMem = .Item(i)                                      'Get the thunk address
                If IsBadCodePtr(z_ScMem) = 0 Then                       'Ensure that the thunk hasn't already released its memory
                    sc_UnSubclass zData(IDX_HWND)                       'UnSubclass
                End If
            Next i                                                      'Next member of the collection
        End With
        Set z_Funk = Nothing                                            'Destroy the hWnd/thunk-address collection
    End If
    
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass( _
    ByVal lng_hWnd As Long _
)

    If z_Funk Is Nothing Then                                           'Ensure that subclassing has been started
        zError "sc_UnSubclass", "Window handle isn't subclassed"
    Else
        If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                   'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN) = -1                                    'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE                            'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE                            'Delete all after messages
        End If
        z_Funk.Remove "h" & lng_hWnd                                    'Remove the specified window handle from the collection
    End If
    
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg( _
    ByVal lng_hWnd As Long, _
    ByVal uMsg As Long, _
    Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER _
)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                       'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                     'If the message is to be added to the before original WndProc table...
            zAddMsg uMsg, IDX_BTABLE                                    'Add the message to the before table
        End If
        If When And MSG_AFTER Then                                      'If message is to be added to the after original WndProc table...
            zAddMsg uMsg, IDX_ATABLE                                    'Add the message to the after table
        End If
    End If
  
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg( _
    ByVal lng_hWnd As Long, _
    ByVal uMsg As Long, _
    Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER _
)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                       'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                     'If the message is to be deleted from the before original WndProc table...
            zDelMsg uMsg, IDX_BTABLE                                    'Delete the message from the before table
        End If
        If When And MSG_AFTER Then                                      'If the message is to be deleted from the after original WndProc table...
            zDelMsg uMsg, IDX_ATABLE                                    'Delete the message from the after table
        End If
    End If
  
End Sub

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg( _
    ByVal uMsg As Long, _
    ByVal nTable As Long _
)

    Dim nCount As Long                                                  'Table entry count
    Dim nBase  As Long                                                  'Remember z_ScMem
    Dim i      As Long                                                  'Loop index

    nBase = z_ScMem                                                     'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                             'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                         'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                           'Set the table entry count to ALL_MESSAGES
    Else
        nCount = zData(0)                                               'Get the current table entry count
        If nCount >= MSG_ENTRIES Then                                   'Check for message table overflow
            zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
            GoTo Bail
        End If

        For i = 1 To nCount                                             'Loop through the table entries
            If zData(i) = 0 Then                                        'If the element is free...
                zData(i) = uMsg                                         'Use this element
                GoTo Bail                                               'Bail
            ElseIf zData(i) = uMsg Then                                 'If the message is already in the table...
                GoTo Bail                                               'Bail
            End If
        Next i                                                          'Next message table entry

        nCount = i                                                      'On drop through: i = nCount + 1, the new table entry count
        zData(nCount) = uMsg                                            'Store the message in the appended table entry
    End If

    zData(0) = nCount                                                   'Store the new table entry count
Bail:
    z_ScMem = nBase                                                     'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg( _
    ByVal uMsg As Long, _
    ByVal nTable As Long _
)

    Dim nCount As Long                                                  'Table entry count
    Dim nBase  As Long                                                  'Remember z_ScMem
    Dim i      As Long                                                  'Loop index

    nBase = z_ScMem                                                     'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                             'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                         'If ALL_MESSAGES are being deleted from the table...
        zData(0) = 0                                                    'Zero the table entry count
    Else
        nCount = zData(0)                                               'Get the table entry count
    
        For i = 1 To nCount                                             'Loop through the table entries
            If zData(i) = uMsg Then                                     'If the message is found...
                zData(i) = 0                                            'Null the msg value -- also frees the element for re-use
                GoTo Bail                                               'Bail
            End If
        Next i                                                          'Next message table entry
    
        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
    End If
  
Bail:
    z_ScMem = nBase                                                     'Restore the value of z_ScMem
End Sub

'Error handler
Private Sub zError( _
    ByVal sRoutine As String, _
    ByVal sMsg As String _
)

    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
    
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr( _
    ByVal sDLL As String, _
    ByVal sProc As String _
) As Long

    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)             'Get the specified procedure address
    Debug.Assert zFnAddr                                                'In the IDE, validate that the procedure address was located
    
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd( _
    ByVal lng_hWnd As Long _
) As Long

    If z_Funk Is Nothing Then                                           'Ensure that subclassing has been started
        zError "zMap_hWnd", "Subclassing hasn't been started"
    Else
        On Error GoTo Catch                                             'Catch unsubclassed window handles
        z_ScMem = z_Funk("h" & lng_hWnd)                                'Get the thunk address
        zMap_hWnd = z_ScMem
    End If
  
    Exit Function                                                       'Exit returning the thunk address

Catch:
    zError "zMap_hWnd", "Window handle isn't subclassed"
    
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf( _
    ByVal oCallback As Object, _
    ByVal nOrdinal As Long _
) As Long

    Dim bSub  As Byte                                                   'Value we expect to find pointed at by a vTable method entry
    Dim bVal  As Byte
    Dim nAddr As Long                                                   'Address of the vTable
    Dim i     As Long                                                   'Loop index
    Dim j     As Long                                                   'Loop limit
  
    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                   'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then                           'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                      'Probe for a Form method
            If Not zProbe(nAddr + &H7A4, i, bSub) Then                  'Probe for a UserControl method
                Exit Function                                           'Bail...
            End If
        End If
    End If
  
    i = i + 4                                                           'Bump to the next entry
    j = i + 1024                                                        'Set a reasonable limit, scan 256 vTable entries
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4                               'Get the address stored in this vTable entry
    
        If IsBadCodePtr(nAddr) Then                                     'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4     'Return the specified vTable entry address
            Exit Do                                                     'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                            'Get the byte pointed to by the vTable entry
        If bVal <> bSub Then                                            'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4     'Return the specified vTable entry address
            Exit Do                                                     'Bad method signature, quit loop
        End If
    
        i = i + 4                                                       'Next vTable entry
    Loop
    
End Function

'Probe at the specified start address for a method signature
Private Function zProbe( _
    ByVal nStart As Long, _
    ByRef nMethod As Long, _
    ByRef bSub As Byte _
) As Boolean

    Dim bVal    As Byte
    Dim nAddr   As Long
    Dim nLimit  As Long
    Dim nEntry  As Long
  
    nAddr = nStart                                                      'Start address
    nLimit = nAddr + 32                                                 'Probe eight entries
    Do While nAddr < nLimit                                             'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4                          'Get the vTable entry
    
        If nEntry <> 0 Then                                             'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1                       'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then                          'Check for a native or pcode method signature
                nMethod = nAddr                                         'Store the vTable entry
                bSub = bVal                                             'Store the found method signature
                zProbe = True                                           'Indicate success
                Exit Function                                           'Return
            End If
        End If
    
        nAddr = nAddr + 4                                               'Next vTable entry
    Loop
    
End Function

Private Property Get zData( _
    ByVal nIndex As Long _
) As Long

    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
    
End Property

Private Property Let zData( _
    ByVal nIndex As Long, _
    ByVal nValue As Long _
)

    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
    
End Property

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1( _
    ByVal bBefore As Boolean, _
    ByRef bHandled As Boolean, _
    ByRef lReturn As Long, _
    ByVal lng_hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long, _
    ByRef lParamUser As Long _
)

'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************

    Dim udtHdr  As WAVEHDR

    If uMsg = MM_WOM_DONE Then
        ' free buffer memory
        RtlMoveMemory VarPtr(udtHdr), ByVal lParam, Len(udtHdr)
        waveOutUnprepareHeader hWO, ByVal lParam, Len(udtHdr)
        FreeBufferMem lParam

        lngBufCnt = lngBufCnt - 1
        If lngBufCnt < 0 Then lngBufCnt = 0
        
        If PlayStatus = WO_PLAYING Then
            ' inform class owner about finished buffer
            RaiseEvent BufferDone(udtHdr.dwUser)
        End If
    End If

End Sub
