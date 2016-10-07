Set smsClient = CreateObject("Microsoft.SMS.Client")
WScript.Echo "Client is currently assigned to site " & _
	smsClient.GetAssignedSite
