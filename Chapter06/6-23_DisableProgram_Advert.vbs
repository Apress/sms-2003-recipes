strSMSServer = <SMSServer>

strPackageID = "LAB00006"
strProgramName = "Microsoft Updates Tool"
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

Set objProgram=objSMS.Get("SMS_Program.PackageID='" & _
    strPackageID & "',ProgramName='" & strProgramName & "'")

intProgramFlags = objProgram.ProgramFlags

if (intProgramFlags and DISABLE_PROGRAM) then
    wscript.echo "Disable Program Flag already set!"
else
    wscript.echo "Disabling program now."
    intProgramFlags = intProgramFlags or DISABLE_PROGRAM      
	objProgram.ProgramFlags = intProgramFlags
	objProgram.Put_
end if

