strSMSServer = <SMSServer>

strCollID = "LAB0002B"

strQuery = "select SMS_R_System.ResourceID," & _
    "SMS_R_System.ResourceType,SMS_R_System.Name," & _
    "SMS_R_System.SMSUniqueIdentifier," & _
    "SMS_R_System.ResourceDomainORWorkgroup," & _
    "SMS_R_System.Client from SMS_R_System inner join " & _
    "SMS_G_System_SYSTEM on " &_
    "SMS_G_System_SYSTEM.ResourceID = " & _
    "SMS_R_System.ResourceId where " & _
    "SMS_G_System_SYSTEM.Name not in (select " & _
    "SMS_G_System_SYSTEM.Name from SMS_R_System inner " & _
    "join SMS_G_System_SoftwareFile on " & _
    "SMS_G_System_SoftwareFile.ResourceID = " & _
    "SMS_R_System.ResourceId inner join " & _
    "SMS_G_System_SYSTEM on " & _
    "SMS_G_System_SYSTEM.ResourceID =" & _
    "SMS_R_System.ResourceId where " & _
    "SMS_G_System_SoftwareFile.FileName = 'wuauclt.exe' " & _
    "and SMS_G_System_SoftwareFile.FilePath like '%system32\\%')"
strQueryName = "Systems That Need Windows Update Agent"

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
'make sure we have a valid query
ValidQuery = clsQueryRule.ValidateQuery(strQuery)
If ValidQuery Then
    Set instQueryRule = clsQueryRule.SpawnInstance_
    instQueryRule.QueryExpression = strQuery
    instQueryRule.RuleName = strQueryName
    instCollection.AddMembershipRule instQueryRule
End If
