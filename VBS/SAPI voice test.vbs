Set objVoice = CreateObject("SAPI.SpVoice")
'Set objVoice.Voice = objVoice.GetVoices("Name=Microsoft Simplified Chinese").Item(0)
Objvoice. Rate = 2 'speed: – 10,10 0
Objvoice.volume = 100 'sound: 0100 100
Dim  englishVoice

Wscript.Echo "There are " & objVoice.GetVoices.Count & " languages installed in this machine!"
'List voices that are installed.
For Each strVoice in objVoice.GetVoices
    Wscript.Echo strVoice.GetDescription
    If InStr(strVoice.GetDescription, "English") <> 0 Then
      Set englishVoice = strVoice
   End If
Next

If Len(englishVoice.GetDescription) <> 0 Then
For Each voice in objVoice.GetVoices
 Set objVoice.Voice = englishVoice
 If InStr(voice.GetDescription, "English") <> 0 Then
    Wscript.Echo "Speaking in English!"
    objVoice.Speak "I speak English"
    objVoice.Speak "hello how are you"
 End If
 If InStr(voice.GetDescription, "French") <> 0 Then
    Wscript.Echo "Speaking in French!"
    objVoice.Speak "I speak French"
    Set objVoice.Voice = voice
    objVoice.Speak "Bonjour comment vas-tu"
 End If
 If InStr(voice.GetDescription, "Russian") <> 0 Then
    Wscript.Echo "Speaking in Russian!"
    objVoice.Speak "I speak Russian" 
    Set objVoice.Voice = voice  
    objVoice.Speak "Привет, как дела"
 End If
 If InStr(voice.GetDescription, "Spanish") <> 0 Then
    Wscript.Echo "Speaking in Spanish!" 
    objVoice.Speak "I speak Spanish" 
    Set objVoice.Voice = voice 
    objVoice.Speak "Hola como estas"
 End If
 If InStr(voice.GetDescription, "Chinese") <> 0 Then
    Wscript.Echo "Speaking in Chinese!" 
    objVoice.Speak "I speak Chinese" 
    Set objVoice.Voice = voice
    objVoice.Speak "你好吗。"
 End If
Next
 End If

