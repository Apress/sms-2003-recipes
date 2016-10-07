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
    Set objSWbemLocator =  CreateObject _
        ("WbemScripting.SWbemLocator")
    Set objSMS = objSWbemLocator.ConnectServer _
        (strSMSServer, "root\sms\site_" & strSMSSiteCode )
    Set colPrograms = objSMS.ExecQuery _
        ("Select * From SMS_Program " & _
        "WHERE PackageID='" & strPackageID & _
        "' order by ProgramName")
    For Each objProgram In colPrograms
        wscript.echo vbTAB & objProgram.ProgramName & vbTAB & _
            GetUserLogonRequirement(strPackageID, _
                 objProgram.ProgramName)
    Next
End Sub

Function GetUserLogonRequirement(strPackageID, strProgramName)
    USER_LOGGED_ON = 2^(14-1)
    WHETHER_OR_NOT_USER_LOGGED_ON = 2^(15-1)
    NO_USER_LOGGED_ON = 2^(17-1)
    
    Set objProgram=objSMS.Get _
        ("SMS_Program.PackageID='" & strPackageID & "',ProgramName='" & _
        strProgramName & "'")
    
    intProgramFlags = objProgram.ProgramFlags
    
    if (intProgramFlags and USER_LOGGED_ON) then
        strInfo = "Only when a user is logged on."
    elseif _
    (intProgramFlags and WHETHER_OR_NOT_USER_LOGGED_ON) then
        strInfo = "Whether or not a user is logged on."
    elseif (intProgramFlags and NO_USER_LOGGED_ON) then
        strInfo = "Only when no user is logged on."
    else

    end if
    GetUserLogonRequirement = strInfo
End Function
