Option Explicit
dim tts, fs, wsh, ts, strFile, strText, strOld, intLine, intSpace
Const ForReading = 1
   'Register a text-to-voice engine. Any engine!
   TTSEngine
   If tts Is Nothing Then
      InputBox "You don't have a voice! Please hop on the internet, download, and install the Microsoft speech engine from here:", "Text Reader", "http://activex.microsoft.com/activex/controls/sapi/spchapi.exe"
      WScript.Quit 1
   End If

   'Read the file one line at a time
   Do Until WScript.stdin.AtEndOfStream
      Say WScript.stdin.ReadLine
   Loop
   Set tts = Nothing

Sub Say(strText)
Dim blnIsSpeaking
   On Error Resume Next 'Ignore unsupported methods and arguments!
   Err.Clear
   tts.Speak strText
   If Err.Number <> 0 Then tts.Speak strText, 16
   'The value 16 is "READING" in SAPI 4, and "NOT_XML" in SAPI 5
   blnIsSpeaking = False
   blnIsSpeaking = tts.IsSpeaking 'Unsupported method will not change value of blnIsSpeaking
   Do While blnIsSpeaking
      WScript.Sleep 500
      blnIsSpeaking = tts.IsSpeaking
   Loop
End Sub

Sub Status(strMessage)
   If Lcase(Right(Wscript.FullName, 12)) = "\cscript.exe" Then
      Wscript.Echo strMessage
   End If
End Sub

Sub TTSEngine()
   On Error Resume Next
   Set tts = Nothing
   
   'C:\Program Files\Common Files\Microsoft Shared\Speech\sapi.dll
   'Microsoft Speech Object Library
   'SpeechLib.SpVoice
   'TypeName "SpVoice"
   Set tts = CreateObject("Sapi.SpVoice")
   If Not(tts Is Nothing) Then Exit Sub

   'C:\WINNT\speech\vcmd.exe
   'C:\WINNT\Speech\vtxtAuto.tlb
   'Voice Text Object Library
   'Voice Text 1.0 Type Library
   'VTxtAuto.VTxtAuto
   'TypeName "IVTxtAuto"

   Set tts = CreateObject("Speech.VoiceText")
   tts.Register "", " "
   tts.Enabled = True
   If Not(tts Is Nothing) Then Exit Sub
   
   Set tts = CreateObject("Speech.VoiceText.1")
   tts.Register "", " "
   tts.Enabled = True
   If Not(tts Is Nothing) Then Exit Sub
   
   'C:\WINNT\Speech\vtext.dll
   'Microsoft Voice Text
   'HTTSLib.TextToSpeech
   'TypeName "TextToSpeech"

   Set tts = CreateObject("TextToSpeech.TextToSpeech")
   tts.Enabled = True
   If Not(tts Is Nothing) Then Exit Sub
   
   Set tts = CreateObject("TextToSpeech.TextToSpeech.1")
   tts.Enabled = True
   If Not(tts Is Nothing) Then Exit Sub

   'C:\WINNT\Speech\Xvoice.dll
   'Microsoft Direct Speech Synthesis
   'ACTIVEVOICEPROJECTLib.DirectSS
   'TypeName "DirectSS"

   Set tts = CreateObject("ActiveVoice.ActiveVoice")
   tts.Enabled = True
   If Not(tts Is Nothing) Then Exit Sub

   Set tts = CreateObject("ActiveVoice.ActiveVoice.1")
   tts.Enabled = True
   If Not(tts Is Nothing) Then Exit Sub

   On Error Goto 0
End Sub