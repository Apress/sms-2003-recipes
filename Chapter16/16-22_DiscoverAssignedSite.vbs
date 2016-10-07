'Requires a restart of the SMS Agent Host Service
strComputerName = "2kPro"
Set objWMIService = GetObject("winmgmts://" & _
	strComputerName & "/root/ccm")
Set objSMSClient = objWMIService.ExecQuery _
	("Select * from SMS_Client")

for each objClient in objSMSClient
	objClient.EnableAutoAssignment = 1
	objClient.Put_ 0
next