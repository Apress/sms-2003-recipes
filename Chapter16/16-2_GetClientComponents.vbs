strComputerName = "."
Set objSMS = GetObject("winmgmts://" & strComputerName & _
	 "/root/ccm")
Set objSMSComponents = objSMS.ExecQuery _
	("Select * from CCM_InstalledComponent")
for each objComponent in objSMSComponents
	wscript.echo objComponent.Name & vbTAB & objComponent.Version
next