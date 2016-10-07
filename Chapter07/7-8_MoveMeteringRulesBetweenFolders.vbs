strSMSServer = <SMSServer>

strMeteringIDs = "LAB00008,LAB00009"    
'to move more than one, separate with commas
intSourceFolder = 0 'Source Folder (root node in this case)
intDestFolder = 8 'Destination Folder
intObjectType = 9 '9=Metering Rule for type of object

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

Set objFolder = objSMS.Get("SMS_ObjectContainerItem")
arrMeteringIDs = split(strMeteringIDs, ",")
retval = objFolder.MoveMembers _
(arrMeteringIDs, intSourceFolder, intDestFolder , intObjectType)
