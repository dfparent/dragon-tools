' This script can execute a voice command without having access to a mic
' Just type in what you would say in the prompt and the script will mimic
' the recognition of that text

option explicit

dim dragonEngine
set dragonEngine = CreateObject("Dragon.DgnEngineControl")
dragonEngine.Register

dim response
response = InputBox("Type in what you would like to say:")
if response = "" then
	WScript.Quit
end if

dragonEngine.RecognitionMimic response

dragonEngine.UnRegister false

WScript.Quit

