strSMSServer = <SMSServer>
strHelpDesk="LAB\SMSHelpDesk" 'Domain\Group or username
strCollID = "LAB00159"  'ID of the collection

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
    
Set objNewRight = objSMS.Get _
  ("SMS_UserInstancePermissions").SpawnInstance_()
objNewRight.UserName = strHelpDesk
objNewRight.ObjectKey = 1 '1=collection
objNewRight.InstanceKey = strCollID
objNewRight.InstancePermissions = 1+2'grant Read and Modify
objNewRight.Put_    
