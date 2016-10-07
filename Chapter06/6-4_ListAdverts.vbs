strSMSServer = <SMSServer>

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

Set colAdverts = objSMS.ExecQuery _
    ("Select * From SMS_Advertisement order by " & _
        " AdvertisementName")
For Each objAdvert In colAdverts
    wscript.echo objAdvert.AdvertisementName & vbTAB & _
        objAdvert.PresentTime & objAdvert.AssignedSchedule
Next
