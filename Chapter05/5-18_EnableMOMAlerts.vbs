DISABLE_MOM_ALERT_WHILE_RUNNING = 2^(5)
GENERATE_MOM_ALERT_IF_FAILURE = 2^(6)

strSMSServer = <SMSServer>

strPackageID = "LAB0000A"
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

if (intProgramFlags and DISABLE_MOM_ALERT_WHILE_RUNNING) then
    wscript.echo "DISABLE_MOM_ALERT_WHILE_RUNNING flag " & _
        "already set!"
else
    wscript.echo "Setting DISABLE_MOM_ALERT_WHILE_RUNNING flag"
    intProgramFlags = intProgramFlags or _
        DISABLE_MOM_ALERT_WHILE_RUNNING
end if

if (intProgramFlags and GENERATE_MOM_ALERT_IF_FAILURE) then
    wscript.echo "GENERATE_MOM_ALERT_IF_FAILURE flag " & _
        "already set!"
else
    wscript.echo "Setting GENERATE_MOM_ALERT_IF_FAILURE flag"
    intProgramFlags = intProgramFlags or _
        GENERATE_MOM_ALERT_IF_FAILURE
end if
objProgram.ProgramFlags = intProgramFlags
objProgram.Put_

