Dim message, sapi
message="Hibernation cancelled"
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message 