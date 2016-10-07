ONUSERLOGOFF = 2^(10)
NO_DISPLAY = 2^(25)

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
if (intAdvertFlags and ONUSERLOGOFF) then
	wscript.echo "ONUSERLOGOFF flag " & _
		"already set!"
else
	wscript.echo "Setting ONUSERLOGOFF flag"
	intAdvertFlags = intAdvertFlags or _
		ONUSERLOGOFF
end if
if (intAdvertFlags and NO_DISPLAY) then
	'allow users to run program independent of assignment
	wscript.echo "Clearing NO_DISPLAY flag"
	intAdvertFlags = intAdvertFlags and not _
		NO_DISPLAY
		
else
	'allow users to run program independent of assignment 
	' already enabled
	wscript.echo "NO_DISPLAY flag already cleared!"
end if
objAdvertisement.AdvertFlags = intAdvertFlags
objAdvertisement.Put_
