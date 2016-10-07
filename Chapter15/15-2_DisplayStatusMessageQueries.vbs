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

Set colQueries = objSMS.ExecQuery _
    ("select * from SMS_Query where TargetClassName = " & _
    "'SMS_StatusMessage' order by Name")
for each objQuery in colQueries
    wscript.echo objQuery.Name & "(" & _
        objQuery.QueryID & ")" & vbTAB & objQuery.Expression
next
