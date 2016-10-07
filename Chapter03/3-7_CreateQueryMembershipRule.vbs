strSMSServer = <SMSServer>

strCollID = "LAB0000F"
strQuery = "select * from SMS_R_System inner join " & _
    "SMS_G_System_ADD_REMOVE_PROGRAMS on " & _
    "SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID = " & _
    "SMS_R_System.ResourceId where " & _
    "SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName = 'SMS View'"
strQueryName = "Systems that have SMSView Installed"

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

    Set instCollection = objSMS.Get _
    ("SMS_Collection.CollectionID='" & strCollID & "'")
Set clsQueryRule = objSMS.Get _
    ("SMS_CollectionRuleQuery")

'next we need to validate the query
ValidQuery = clsQueryRule.ValidateQuery(strQuery)

If ValidQuery Then
    Set instQueryRule = clsQueryRule.SpawnInstance_
    instQueryRule.QueryExpression = strQuery
    instQueryRule.RuleName = strQueryName
    instCollection.AddMembershipRule instQueryRule
End If
