strSMSServer = <SMSServer>
strQueryIDs = "LAB0001E,LAB0001F,LAB0002E"
'to move more than one, separate with commas

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


intSourceFolder=0 'Source Folder (root node in this case)
intDestFolder=10 'Destination Folder
intObjectType=7 '7=Query for type of object

Set objFolder = objSMS.Get("SMS_ObjectContainerItem")
arrQueryIDs = split(strQueryIDs, ",")
retval = objFolder.MoveMembers _
(arrQueryIDs, intSourceFolder, intDestFolder , intObjectType)
wscript.echo retval
