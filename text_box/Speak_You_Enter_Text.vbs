do
var=inputbox ("Text in window","text at top of window","default text")
set speech = Wscript.CreateObject("SAPI.spVoice")
speech.speak var
loop