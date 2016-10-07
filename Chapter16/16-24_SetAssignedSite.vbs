strComputer = "2kPro"
strSMSSite = "XXX"
Set objSMS = GetObject("winmgmts://" & strComputer & _
	"/root/ccm")
Set objSMSClient = objSMS.Get("SMS_Client")
set objParams = objSMSClient.Methods_("SetAssignedSite"). _
	inParameters.SpawnInstance_()
objParams.sSiteCode = strSMSSite
objSMS.ExecMethod _
	"SMS_Client", "SetAssignedSite", objParams




' sSiteCode = "XXX"

' sMachine = "."

' Set oCCMNamespace = GetObject("winmgmts:{impersonationLevel=impersonate}//" & sMachine & "/root/ccm")

' Set smsClient = oCCMNamespace.Get("SMS_Client")

' smsClient.SetAssignedSite sSiteCode, oParams
