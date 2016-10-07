strSMSServer = <SMSServer>
strQueryID = "LAB0001E"

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
    'use the following line to display all instance information
    wscript.echo objResult.GetObjectText_
Next
