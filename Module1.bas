Attribute VB_Name = "Module1"
Public iHeight As Integer
Public iWidth As Integer
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function sndPlaySound Lib "winmm.dll" Alias _
"sndPlaySoundA" (ByVal lpszSoundName As String, ByVal _
uFlags As Long) As Long

