strSMSServer = <SMSServer>

DISABLE_PROGRAM = 2^(12)

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

Set colPrograms = objSMS.ExecQuery _
    ("Select * From SMS_Program order by PackageID")
For Each objProgram In colPrograms
    if (objProgram.ProgramFlags and DISABLE_PROGRAM) then
        wscript.echo "Program Name: " & _
            objProgram.ProgramName & vbTAB & _
            "Package Name: " & _
            GetPackageName(objProgram.PackageID) & " (" & _
            objProgram.PackageID & ")"
    end if
Next

Function GetPackageName(strPckID)
    Set objSWbemLocator =  CreateObject _
        ("WbemScripting.SWbemLocator")
    Set objSMS = objSWbemLocator.ConnectServer _
        (strSMSServer, "root\sms\site_" & strSMSSiteCode )
    Set objPackage=objSMS.Get _
        ("SMS_Package.PackageID='" & _
        strPckID & "'")
    GetPackageName = objPackage.Name
End Function
