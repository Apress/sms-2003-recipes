strSMSServer = <SMSServer>
strPackageID = "LAB00002"

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

Set colAdverts = objSMS.ExecQuery("Select * From SMS_Advertisement " & _
    "WHERE PackageID='" & strPackageID & "'")
For Each objAdvert In colAdverts
    wscript.echo "Deleting " & objAdvert.AdvertisementName
    objAdvert.Delete_
Next

