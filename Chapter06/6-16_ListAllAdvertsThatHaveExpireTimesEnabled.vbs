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
    strExpireTime = GetExpirationDateTimeString _
        (objAdvert.AdvertisementID)
    if (not strExpireTime = "") then
        wscript.echo objAdvert.AdvertisementID & vbTAB & _
            objAdvert.AdvertisementName & vbTAB & strExpireTime
    end if
Next

Function GetExpirationDateTimeString(strAdvertID)
    Set objAdvert=objSMS.Get _
        ("SMS_Advertisement.AdvertisementID='" & _
            strAdvertID & "'")
    if objAdvert.ExpirationTimeEnabled = True then
        GetExpirationDateTimeString = _
            WMIDateStringToDate(objAdvert.ExpirationTime)
        if (objAdvert.ExpirationTimeIsGMT) Then
            GetExpirationDateTimeString = _
                GetExpirationDateTimeString & " GMT"
        end if
    end if
End Function

'Utility function to convert WMI Date string to a real date
Function WMIDateStringToDate(dtmInstallDate)
    '4/12/2005 3:46:04 AM = 20050412034604.000000-000
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & _
        "/" & Mid(dtmInstallDate, 7, 2) & "/" & _
        Left(dtmInstallDate, 4) & " " & _
        Mid (dtmInstallDate, 9, 2) & ":" & _
        Mid(dtmInstallDate, 11, 2) & ":" & _
        Mid(dtmInstallDate,13, 2))
End Function