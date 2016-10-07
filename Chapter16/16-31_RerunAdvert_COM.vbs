Set uiResource = CreateObject("UIResource.UIResourceMgr")
Set programList = uiResource.GetAvailableApplications
For each p in programList
    wscript.echo p.Name
    If p.Name = "SMS 2003 SDK V3" then
        uiResource.ExecuteProgram p.ID, p.PackageID, True
        Exit For
    End if
Next
