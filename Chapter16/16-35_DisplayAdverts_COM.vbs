Set UI = CreateObject("UIResource.UIResourceMgr")
Set programList = UI.GetAvailableApplications

For each program in programList
	wscript.echo program.PackageID & vbTAB & program.ID & _
		vbTAB & program.Name & vbTAB & program.PackageName & _
		vbTAB & program.Version
Next
