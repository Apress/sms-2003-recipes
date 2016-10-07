ONUSERLOGON = 2^(9)

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

intAdvertFlags = objAdvertisement.AdvertFlags
if (intAdvertFlags and ONUSERLOGON) then
	wscript.echo "ONUSERLOGON flag " & _
		"already set!"
else
	wscript.echo "Setting ONUSERLOGON flag"
	intAdvertFlags = intAdvertFlags or _
		ONUSERLOGON
	objAdvertisement.AdvertFlags = intAdvertFlags
	objAdvertisement.Put_			
end if

