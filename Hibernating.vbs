Dim message, sapi
message="System will hibernate in ten second"
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message 