strComputerName = "."
Set smsClient = GetObject("winmgmts://" & strComputerName & _
	"/root/ccm:SMS_Client")
Set result = smsClient.ExecMethod_("GetAssignedSite")
WScript.Echo "Client is currently assigned to site " & _
	result.sSiteCode
