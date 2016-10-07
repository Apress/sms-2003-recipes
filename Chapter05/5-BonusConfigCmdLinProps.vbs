RUN_MAXIMIZED = 2^(23)
RUN_MINIMIZED = 2^(22)
HIDE_WINDOW = 2^(24)

LOGOFF_USER = 2^(25)
SMS_RESTART = 2^(19)
PROGRAM_RESTART = 2^(18)

strSMSServer = <SMSServer>

strPackageID = "LAB0000C"
strProgramName = "sdktest2"
strCategory = "Developer"

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

objProgram.CommandLine = "msiexec.exe /q ALLUSERS=2 " & _
	"/m MSI4LYMT /i " & chr(34) & "SMSSDKSetup.msi" & chr(34)
intProgramFlags = objProgram.ProgramFlags
if (intProgramFlags and RUN_MAXIMIZED) then
	wscript.echo "RUN_MAXIMIZED already configured!"
else
	'ensure HIDE_WINDOW and RUN_MINIMIZED are not enabled.
	intProgramFlags = intProgramFlags and not HIDE_WINDOW
	intProgramFlags = intProgramFlags and not RUN_MINIMIZED
	wscript.echo "Configuring Program to RUN_MAXIMIZED"
	intProgramFlags = intProgramFlags or RUN_MAXIMIZED
end if

if (intProgramFlags and SMS_RESTART) then
	wscript.echo "SMS_RESTART already configured!"
else
	'ensure LOGOFF_USER and PROGRAM_RESTART are not enabled.
	intProgramFlags = intProgramFlags and not LOGOFF_USER
	intProgramFlags = intProgramFlags and not PROGRAM_RESTART
	wscript.echo "Configuring Program to SMS_RESTART"
	intProgramFlags = intProgramFlags or SMS_RESTART
end if
objProgram.ProgramFlags = intProgramFlags
objProgram.Description = strCategory  'description is actually the program "category"
objProgram.Put_



