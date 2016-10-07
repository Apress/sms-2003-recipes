strSMSServer = <SMSServer>

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
    ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
    If Loc.ProviderForLocalSite = True Then
        Set objSMS = objLoc.ConnectServer(Loc.Machine, "root\sms\site_" & _
            Loc.SiteCode)
            strSMSSiteCode = Loc.Sitecode
    end if
Next

Set colPackages = objSMS.ExecQuery("Select * From SMS_Package where " & _
     "ImageFlags = 0 order by Name")     

For each objPackage in ColPackages
    wscript.echo objPackage.Name & " (" & objPackage.PackageID & ")"
    ListPrograms(objPackage.PackageID)
Next

Sub ListPrograms(strPackageID)
    Set colPrograms = objSMS.ExecQuery("Select * From SMS_Program " & _
        "WHERE PackageID='" & strPackageID & "' order by ProgramName")
    For Each objProgram In colPrograms
        wscript.echo vbTAB & objProgram.ProgramName 
    Next
End Sub
