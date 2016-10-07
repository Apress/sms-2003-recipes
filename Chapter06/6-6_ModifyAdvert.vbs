strSMSServer = <SMSServer>

strAdvertID = "LAB20015"
strNewCollectionID = "LAB00011"

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

Set objAdvert=objSMS.Get _
    ("SMS_Advertisement.AdvertisementID='" & strAdvertID & "'")
objAdvert.CollectionID = strNewCollectionID 
objAdvert.Put_  
