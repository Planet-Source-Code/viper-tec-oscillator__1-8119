Attribute VB_Name = "modSynth_DX7"
Public Declare Function osQueryPerformanceCounter Lib "kernel32.dll" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As Long
Public Declare Function osQueryPerformanceFrequency Lib "kernel32.dll" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As Long
Public freq As Currency, Count As Currency

Public DX7 As New DirectX7, DS As DirectSound, DSB(1) As DirectSoundBuffer
Public dsbd As DSBUFFERDESC, PCM As WAVEFORMATEX

Public O1SBuffer(360) As Byte, O2SBuffer(360) As Byte

Public Osc1Samp As Single, Osc2Samp As Single
Public Osc1FCutoff As Integer, Osc2FCutoff As Integer
Public Osc1Amp As Integer, Osc2Amp As Integer, O2F As Integer

Public AM_Speed As Integer, AM2_Speed As Integer, FM_Speed As Integer

Public fm As Long
Public i As Integer, f As Integer
Public n As Single, f2 As Single

Public Temp1 As Integer, Temp2 As Integer 'VU METER VARS
    
Public Const pi = 3.14159265358979

Function Init_DX7(Hwnd As Long) As Boolean: On Error GoTo InitErrorOut1
''Fill WaveFormat Structure
PCM.nFormatTag = WAVE_FORMAT_PCM
PCM.nChannels = 1
PCM.lSamplesPerSec = 11050
PCM.nBitsPerSample = 8
PCM.nBlockAlign = 1
PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign
PCM.nSize = 0
''Fill BufferDescription Structure
dsbd.lFlags = DSBCAPS_STATIC
dsbd.lBufferBytes = 360
''Create the DirectSound Object
Set DS = DX7.DirectSoundCreate("")
''Set the Cooperative Level
DS.SetCooperativeLevel Hwnd, DSSCL_NORMAL
''Create Buffers
On Error GoTo InitErrorOut2
Set DSB(0) = DS.CreateSoundBuffer(dsbd, PCM)
Set DSB(1) = DS.CreateSoundBuffer(dsbd, PCM)

Init_DX7 = True
Exit Function 'Function WAS successful!
InitErrorOut2:
Set DSB(0) = Nothing
Set DSB(1) = Nothing
Set DS = Nothing
InitErrorOut1:
Init_DX7 = False
End Function 'Function WAS NOT successful!

Sub Term_DX7() 'Clear the created DX7 Objects.
Set DSB(0) = Nothing
Set DSB(1) = Nothing
Set DS = Nothing
End Sub

Sub DSBWRITE(Num As Integer, ByRef Buffer() As Byte)
'This writes a given buffer (an array of bytes) to a given
'DirectSoundBuffer.
DSB(Num).WriteBuffer 0, 0, Buffer(0), DSBLOCK_ENTIREBUFFER
End Sub

Sub DrawVU(Value As Integer, PB As PictureBox)
''True VU meters with multiple inputs require a FFT...
'VB is a bit slow for that kind of complex algorhythm.
'A VU for each oscillator is the fast & easy way out.
If Value < 75 Then PB.Line (0, 0)-(Value, 0), vbGreen
If Value > 75 And Value <= 95 Then PB.Line (0, 0)-(Value, 0), vbYellow
If Value > 95 Then PB.Line (0, 0)-(Value, 0), vbRed
End Sub

Sub DrawPOINT(dI As Integer, dSamp As Single, PB As PictureBox)
'This plots a given sample. (offset = midpoint of PB's YAxis)
'If you give the offset a set number, in this case 125, the
'CPU cycles with be far less than if you had an equation find
'the midpoint each time.
PB.PSet (dI, dSamp + 125), vbGreen
End Sub

Public Function Timer() As Single
'Speed Up the standard Timer function.
osQueryPerformanceFrequency freq
osQueryPerformanceCounter Count
Let Timer = Count / freq
End Function
