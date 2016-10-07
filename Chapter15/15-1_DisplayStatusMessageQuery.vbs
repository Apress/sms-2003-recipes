strQueryID = "SMS577"
strSMSServer = <SMSServer>

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

Set objQuery = objSMS.Get("SMS_Query.QueryID='" & _
    strQueryID & "'")
wscript.echo objQuery.Expression
wscript.echo objQuery.GetObjectText_
wscript.echo objQuery.GetText_(2)