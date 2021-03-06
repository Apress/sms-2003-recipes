DOWNLOAD_FROM_LOCAL_DISPPOINT = 2^(4)
DOWNLOAD_FROM_REMOTE_DISPPOINT = 2^(6)

strSMSServer = <SMSServer>

strAdvertID = "LAB20016"

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

Set objAdvertisement = objSMS.Get _
    ("SMS_Advertisement.AdvertisementID='" & strAdvertID & "'")

objAdvertisement.RemoteClientFlags = _
    DOWNLOAD_FROM_LOCAL_DISPPOINT + _
        DOWNLOAD_FROM_REMOTE_DISPPOINT
objAdvertisement.Put_
