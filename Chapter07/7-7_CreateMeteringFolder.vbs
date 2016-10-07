strSMSServer = <SMSServer>

strNewFolderName = "InBit Software"

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
    
Set newFolder = objSMS.Get("SMS_ObjectContainerNode") _
    .SpawnInstance_()
newFolder.Name = strNewFolderName
newFolder.ObjectType = 9  'metering Folder
newFolder.ParentContainerNodeId = 0 'root folder
newFolder.Put_    

'Property value Object type 
'2 Package 
'3 Advertisement 
'7 Query 
'8 Report 
'9 Software Metering 
