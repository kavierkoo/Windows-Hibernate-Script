Dim message, sapi
message="System hibernated"
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message 