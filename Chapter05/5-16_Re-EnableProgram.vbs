DISABLE_PROGRAM = 2^(12)

strSMSServer = <SMSServer>

strPackageID = "LAB0000C"
strProgramName = "NET Framework 1.1 SP1"

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

Set objProgram=objSMS.Get("SMS_Program.PackageID='" & _
    strPackageID & "',ProgramName='" & strProgramName & "'")

intProgramFlags = objProgram.ProgramFlags

if (intProgramFlags and DISABLE_PROGRAM) then
    wscript.echo "Re-enabling disabled program"
    intProgramFlags = intProgramFlags and not _
        DISABLE_PROGRAM  
	objProgram.ProgramFlags = intProgramFlags
	objProgram.Put_
else
    wscript.echo "DISABLE_PROGRAM flag not set"
end if

