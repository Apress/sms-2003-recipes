strSMSServer = <SMSServer>

strNewFolderName = "Security Patches"

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
  ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
  If Loc.ProviderForLocalSite = True Then
    Set objSMS = objLoc.ConnectServer(Loc.Machine, "root\sms\site_" & _ 
        Loc.SiteCode)
    strSMSSiteCode = Loc.SiteCode
  end if
Next

Set newFolder = objSMS.Get("SMS_ObjectContainerNode") _
    .SpawnInstance_()
newFolder.Name = strNewFolderName
newFolder.ObjectType = 2  'package
newFolder.ParentContainerNodeId = 0
newFolder.Put_
