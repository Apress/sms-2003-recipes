strSMSServer = <SMSServer>
strPackageID = "LAB0000A"
strProgramName = "NET Framework 1.1 SP1"
strProgramCMDLine = _
    "NDP1.1sp1-KB867460-X86.exe /I /Q /L:%temp%\NetFW1.1.sp1.log"

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
    
Set newProgram = objSMS.Get("SMS_Program").SpawnInstance_()
newProgram.PackageID = strPackageID
newProgram.ProgramName = strProgramName
newProgram.CommandLine = strProgramCMDLine
newProgram.Put_     
