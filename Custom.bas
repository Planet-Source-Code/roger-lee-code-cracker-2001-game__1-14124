Attribute VB_Name = "Custom"
Option Explicit

'Declarations for playing WAV files
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10


Public Function RndNum(Min As Long, Max As Long) As Long
Attribute RndNum.VB_Description = "Returns a random number between the min and max variable values.                                                                "

'Usage:   Number = RndNum(1,100)

'Number will now be equal to the random number generated
'and will be between 1 and 100

Dim X As Long

100
     Randomize
     X = Int(((Max + 1) * Rnd))

If X < Min Then GoTo 100

     RndNum = X

End Function




Public Function PlayWav(WavFile As String)
Attribute PlayWav.VB_Description = "Simplified function to play a non looping wave file"

'Usage:    PlayWav (c:\your.wav)

'The code used above is only used to play a
'non looping wav file

Dim Flags As Long
Dim X As Long

Flags = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(WavFile, Flags)

End Function

