strSMSServer = <SMSServer>

strPackageID = "LAB000CC"

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

'Remove Package
Set objPackage = GetObject( "WinMgmts:\\" & strSMSServer & _
    "\root\SMS\site_" & strSMSSiteCode & _
    ":SMS_Package.PackageID='" & strPackageID & "'")
objPackage.Delete_
wscript.echo "Deleted " & strPackageID
