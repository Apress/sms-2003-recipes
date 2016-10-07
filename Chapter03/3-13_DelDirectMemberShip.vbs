strSMSServer = <SMSServer>

strCollID = "LAB0002B"

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
    
Set objCollection=objSMS.Get("SMS_Collection='" & strCollID & "'" )
'Get the array of embedded SMS_CollectionRule objects.
RuleSet = objCollection.CollectionRules
For Each Rule In RuleSet
    if Rule.Path_.Class = "SMS_CollectionRuleDirect" then
        objCollection.DeleteMembershipRule Rule
    end if
Next
