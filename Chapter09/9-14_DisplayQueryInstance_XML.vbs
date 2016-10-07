Const wbemObjectTextFormatWMIDTD20 = 2
strSMSServer = <SMSServer>
strQueryID = "SMS012"

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

Set objQuery = objSMS.Get("SMS_Query.QueryID='" & strQueryID & "'")

Set colQueryResults = objSMS.ExecQuery(objQuery.Expression)
For Each objResult In colQueryResults
    wscript.echo objResult.GetText_(wbemObjectTextFormatWMIDTD20)
Next
