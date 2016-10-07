strSMSServer = <SMSServer>

strPackageID = "LAB000CE"
strPreferredAddressType = "SMS_LAN_SENDER"
intPriority = 1

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

Set objPackage = objSMS.Get("SMS_Package.PackageID='" & strPackageID & "'")
objPackage.PreferredAddressType = strPreferredAddressType
objPackage.Priority = intPriority
objPackage.Put_
