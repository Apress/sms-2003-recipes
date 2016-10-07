strSMSServer = <SMSServer>
strQueryName = "Basic Query"

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
    ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
    If Loc.ProviderForLocalSite = True Then
        Set objSMS = objLoc.ConnectServer(Loc.Machine, "root\sms\site_" & _
            Loc.SiteCode)
    end if
Next

Set newQuery = objSMS.Get("SMS_Query").SpawnInstance_()
newQuery.Name = strQueryName
newQuery.Expression = "select * from SMS_R_System"
newQuery.TargetClassName = "SMS_R_System"
newQuery.Put_
