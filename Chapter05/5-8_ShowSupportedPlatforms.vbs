strSMSServer = <SMSServer>

ANY_PLATFORM = 2^(27)  'sms doc incorrect

strPackageID = "LAB00086"
strProgramName = "Microsoft Updates"

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

'Check for "Any Platform in ProgramFlags first
if (objProgram.ProgramFlags and ANY_PLATFORM) then
    wscript.echo "This program is configured to run " & _
        "on any platform"
else
    for i = 0 to ubound(objProgram.SupportedOperatingSystems)
        strInfo = _
            objProgram.SupportedOperatingSystems(i).Name & vbTAB
        strInfo = strInfo & _
            objProgram.SupportedOperatingSystems(i).Platform & _
                vbTAB
        strInfo = strInfo & _
            objProgram.SupportedOperatingSystems(i). _
                MinVersion & vbTAB
        strInfo = strInfo & objProgram. _
            SupportedOperatingSystems(i).MaxVersion
        wscript.echo strInfo
    next
end if
