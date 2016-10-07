strComputerName = "."
Set smsClient = GetObject("winmgmts://" & strComputerName & "/root/ccm:SMS_Client")

Set result = smsClient.ExecMethod_("GetAssignedSite")
WScript.Echo "Client is currently assigned to site " & result.sSiteCode

Set inParam = smsClient.Methods_.Item("SetAssignedSite").inParameters.SpawnInstance_()
inParam.sSiteCode = "LAB"
Set result = smsClient.ExecMethod_("SetAssignedSite", inParam)

Set result = smsClient.ExecMethod_("GetAssignedSite")
WScript.Echo "Client is now assigned to " & result.sSiteCode


