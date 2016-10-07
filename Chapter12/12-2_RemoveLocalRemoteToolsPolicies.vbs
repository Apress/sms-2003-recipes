strComputer = "."
Set objWMIService = GetObject _
	("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strComputer & "\root\ccm\Policy\Machine\RequestedConfig")
Set colLocalPolicy = objWMIService.ExecQuery _
	("Select * from CCM_RemoteToolsConfig " & _
		"where policysource = 'local'")

for each objPolicy in colLocalPolicy
	objPolicy.Delete_
next

