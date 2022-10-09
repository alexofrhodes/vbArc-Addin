Attribute VB_Name = "F_Wav"
'Option Compare Database
Option Explicit
Rem @Subfolder Wav>FromVoice
Rem found at http://www.vbarchiv.net
Public Declare PtrSafe Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
Public Declare PtrSafe Function mciSendString Lib "winmm.dll" _
Alias "mciSendStringA" ( _
ByVal lpstrCommand As String, _
ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long
Public Enum BitsPerSec
    Bits16 = 16
    Bits8 = 8
End Enum

Public Enum SampelsPerSec
    Sampels8000 = 8000
    Sampels11025 = 11025
    Sampels12000 = 12000
    Sampels16000 = 16000
    Sampels22050 = 22050
    Sampels24000 = 24000
    Sampels32000 = 32000
    Sampels44100 = 44100
    Sampels48000 = 48000
End Enum

Public Enum Channels
    Mono = 1
    Stereo = 2
End Enum

Public Sub play(file As String)
    Dim wavefile
    wavefile = file
    Call sndPlaySound(wavefile, SND_ASYNC Or SND_FILENAME)
End Sub

Public Sub StartRecord(ByVal BPS As BitsPerSec, _
                       ByVal SPS As SampelsPerSec, ByVal Mode As Channels)
    Dim retStr As String
    Dim cBack As Long
    Dim BytesPerSec As Long
    retStr = Space$(128)
    BytesPerSec = (Mode * BPS * SPS) / 8
    mciSendString "open new type waveaudio alias capture", retStr, 128, cBack
    mciSendString "set capture time format milliseconds" & _
                  " bitspersample " & CStr(BPS) & _
                  " samplespersec " & CStr(SPS) & _
                  " channels " & CStr(Mode) & _
                  " bytespersec " & CStr(BytesPerSec) & _
                  " alignment 4", retStr, 128, cBack
    mciSendString "record capture", retStr, 128, cBack
End Sub

Public Sub SaveRecord(strFile)
    Dim retStr As String
    Dim TempName As String
    Dim cBack As Long
    Dim fs, f
    TempName = strFile        'Left$(strFile, 3) & "Temp.wav"
    retStr = Space$(128)
    mciSendString "stop capture", retStr, 128, cBack
    mciSendString "save capture " & TempName, retStr, 128, cBack
    mciSendString "close capture", retStr, 128, cBack
End Sub

Public Sub StartRecord_Click()
    '#INCLUDE StartRecord
    F_Wav.StartRecord Bits16, Sampels32000, Mono
End Sub

Public Sub EndRecord_Click()
    '#INCLUDE SaveRecord
    F_Wav.SaveRecord Environ("USERPROFILE") & "\Desktop\test.wav"
End Sub

Public Sub Play_Click()
    '#INCLUDE play
    F_Wav.play Environ("USERPROFILE") & "\Desktop\test.wav"
End Sub

Rem @Subfolder Wav>FromString
Sub TestStringToWavFile()
    'needs reference Microsoft Speech Object Library
    'run this to make a wav file from a text input
    '#INCLUDE StringToWavFile
    Dim sP As String, sFN As String, sStr As String, sFP As String
    'set parameter values - insert your own profile name first
    'paths
    sP = Environ("USERPROFILE") & "\Desktop\"        'for example
    sFN = "Mytest.wav"        'overwrites if file name same
    sFP = sP & sFN
    'string to use for the recording
    sStr = "This is a short test string to be spoken in a user's wave file."
    'make voice wav file from string
    StringToWavFile sStr, sFP
End Sub

Function StringToWavFile(sIn As String, sPath As String) As Boolean
    'needs reference Microsoft Speech Object Library
    'makes a spoken wav file from parameter text string
    'sPath parameter needs full path and file name to new wav file
    'If wave file does not initially exist it will be made
    'If wave file does initially exist it will be overwritten
    Dim fs As New SpFileStream
    Dim Voice As New SpVoice
    Dim audioType As Long
    'set the audio format
    fs.Format.Type = SAFT22kHz16BitMono
    'create wav file for writing without events
    fs.Open sPath, SSFMCreateForWrite, False
    'Set wav file stream as output for Voice object
    Set Voice.AudioOutputStream = fs
    'send output to default wav file "SimpTTS.wav" and wait till done
    Voice.Speak sIn, SVSFDefault
    'Close file
    fs.Close
    'wait
    Voice.WaitUntilDone (6000)
    'release object variables
    Set fs = Nothing
    Set Voice.AudioOutputStream = Nothing
    'transfers
    StringToWavFile = True
End Function


