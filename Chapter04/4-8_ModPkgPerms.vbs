'This script grants the "Help Desk" Read, Modify, and Distribute
'Permissions to PackageID = LAB00007
strSMSServer = <SMSServer>

strHelpDesk="SMSVPC\Help Desk"
strPackageID="LAB00007"

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
wscript.echo objPackage.PackageID
Set objNewRight = objSWbemServices.Get _
    ("SMS_UserInstancePermissions").SpawnInstance_()
objNewRight.UserName = strHelpDesk
objNewRight.ObjectKey = 2 'package
objNewRight.InstanceKey = objPackage.PackageID
objNewRight.InstancePermissions = 11 
    '0000000000001011 (read, modify, distribute)
objNewRight.Put_
