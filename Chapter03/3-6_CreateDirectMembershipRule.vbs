strSMSServer = <SMSServer>

strCollID = "LAB0000F" 'ID of the collection
strComputerName = "2kPro" 'computer name to add

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

'obtain the ResourceID for strComputerName
Set colResourceIDs=objSMS.ExecQuery _
    ("SELECT ResourceId FROM SMS_R_System WHERE NetbiosName ='" & _
         strComputerName & "'")
For each insResource in colResourceIDs
        strNewResourceID = insResource.ResourceID
Next
'add the ResourceID to the collection
Set instColl = objSMS.Get _
    ("SMS_Collection.CollectionID=""" & strCollID & """")
Set instDirectRule = objSMS.Get _
    ("SMS_CollectionRuleDirect").SpawnInstance_ ()
instDirectRule.ResourceClassName = "SMS_R_System"
instDirectRule.ResourceID = strNewResourceID
instDirectRule.RuleName = strComputerName
instColl.AddMembershipRule instDirectRule
