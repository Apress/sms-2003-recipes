strComputer = "."
Set objSMS = GetObject("winmgmts://" & strComputer & _
	"/root/ccm/SoftMgmtAgent")
Set colER = objSMS.ExecQuery _
	("Select * from CCM_ExecutionRequest")
for each oER in colER
	wscript.echo oER.ProgramID & vbTAB & oER.State & vbTAB & _
		oER.ProcessID & vbTAB & oER.AdvertID & vbTAB & _
		oER.IsAdminContext
next