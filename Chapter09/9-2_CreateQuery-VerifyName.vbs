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

if QueryDoesntExist(strQueryname) then
    Set newQuery = objSMS.Get("SMS_Query").SpawnInstance_()
    newQuery.Name = strQueryName
    newQuery.Expression = "select *  from  SMS_R_System"
    newQuery.TargetClassName = "SMS_R_System"
    newQuery.Put_
else
    wscript.echo "Query named " & chr(34) & strQueryname & _
        chr(34) & " already exists!"
end if

'function used to verify query name doesn't exist
Function QueryDoesntExist(strName)
    Set colQueries = objSMS.ExecQuery _
    ("select * from SMS_Query where TargetClassName <> '" & _
    "SMS_StatusMessage" & "' and Name = '" & strName & "'")
    if colQueries.Count > 0 then
        QueryDoesntExist = False
    else
        QueryDoesntExist = True
    end if
End Function
