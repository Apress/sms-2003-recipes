strSMSServer = <SMSServer>
strPackageID = "LAB0000B"

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
    
Set colPrograms = objSMS.ExecQuery("Select * From SMS_Program " & _
    "WHERE PackageID='" & strPackageID & "'")
For Each objProgram In colPrograms
    wscript.echo "Deleting " & objProgram.ProgramName
    objProgram.Delete_
Next

