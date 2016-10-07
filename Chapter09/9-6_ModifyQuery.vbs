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

strQueryID = "LAB0001D"

Set objQuery = objSMS.Get("SMS_Query.QueryID='" & _
    strQueryID & "'")
objQuery.Comments = "Use this query to display all systems " & _
    "in the SMS_R_System class"
objQuery.Put_
