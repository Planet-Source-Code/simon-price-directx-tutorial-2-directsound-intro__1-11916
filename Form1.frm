VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DX Tutorial #2 - by Simon Price"
   ClientHeight    =   1800
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   852
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   1092
   End
   Begin VB.CommandButton cmdLoop 
      Caption         =   "Play Sound Looping"
      Height          =   852
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1092
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Sound Once"
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''
'
'       DIRECTX TUTORIAL #2 BY SIMON PRICE
'
'     DIRECT SOUND - LOAD AND PLAY A .WAV FILE
'
''''''''''''''''''''''''''''''''''''''''''''''''''

' The main DirectX object
Private DX As New DirectX7
' The main Direct Sound Object
Private DSOUND As DirectSound
' This is the sound buffer object, it contains
' a piece of memory (buffer) which holds a sound
Private SoundBuffer As DirectSoundBuffer
' This is the description of the sound buffer object
Private SoundDesc As DSBUFFERDESC
' This defines the format of wave file (.wav) data
Private WavFormat As WAVEFORMATEX

' in the format load event we load our sound
Private Sub Form_Load()
' get DirectX to create the DirectSound object
Set DSOUND = DX.DirectSoundCreate("")
' set the cooperative level, here we use normal.
' if we set exclusive cooperative level, no
' other programs would be able to play sounds
DSOUND.SetCooperativeLevel hWnd, DSSCL_NORMAL
' the we load the .wav file into the memory which
' is stored in our soundbuffer object
Set SoundBuffer = DSOUND.CreateSoundBufferFromFile(App.Path & "\sound.wav", SoundDesc, WavFormat)
End Sub

' the play button plays the sound
Private Sub cmdPlay_Click()
'rewind the sound to the beginning
SoundBuffer.SetCurrentPosition 0
'play the sound
SoundBuffer.Play DSBPLAY_DEFAULT
End Sub

' the loop button plays the sound continuously
Private Sub cmdLoop_Click()
'rewind the sound to the beginning
SoundBuffer.SetCurrentPosition 0
'play the sound, passing the DSBPLAY_LOOPING
' flag to tell DirectSound to keep looping until
' you call the stop method on it
SoundBuffer.Play DSBPLAY_LOOPING
End Sub

' the stop button stops the sound from playing
Private Sub cmdStop_Click()
' stop the sound playing
SoundBuffer.Stop
End Sub

