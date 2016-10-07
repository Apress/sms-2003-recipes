strSMSServer = <SMSServer>

strWindowsVersion = "5.00" 'looking for Windows 2000 here
ANY_PLATFORM = 2^(27)

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
    if (objProgram.ProgramFlags and ANY_PLATFORM) then
        'program is set to "any platform" - don't check further
    else
        ListFilteredPrograms objProgram.ProgramName, objProgram.PackageID, _
        	strWindowsVersion
    end if
Next

Sub ListFilteredPrograms(strProgramName, strPackageID, strWinVer)
    Set objProgram=objSMS.Get _
        ("SMS_Program.PackageID='" & _
        strPackageID & "',ProgramName='" & _
        strProgramName & "'")
    for i = 0 to ubound(objProgram.SupportedOperatingSystems)
        if instr(objProgram.SupportedOperatingSystems(i). _
            MaxVersion, strWinVer) then
            blnFoundOne = true
        end if
    next
    if blnFoundOne then
        wscript.echo "Package Name:" & _
            GetPackageName(strPackageID) & vbTAB & _
            "Program Name:" & strProgramName
    else
        'this would capture all programs that had supported 
        'platforms configured that did not contain strWinVer
    end if
End Sub

Function GetPackageName(strPckID)
    Set objPackage=objSMS.Get _
        ("SMS_Package.PackageID='" & _
        strPckID & "'")
    GetPackageName = objPackage.Name
End Function


