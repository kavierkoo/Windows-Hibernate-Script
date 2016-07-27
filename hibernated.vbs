Dim message, sapi
message="System will now hibernate"
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message 