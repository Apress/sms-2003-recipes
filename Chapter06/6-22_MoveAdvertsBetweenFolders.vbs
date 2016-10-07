strSMSServer = <SMSServer>

strSMSSiteCode = "LAB"
strAdvertIDs = "LAB20013,LAB20014"    

'to move more than one, separate with commas
intSourceFolder=0 'Source Folder (root node in this case)
intDestFolder=7 'Destination Folder
intObjectType=3 '3=Advertisement for type of object

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
    ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
    If Loc.ProviderForLocalSite = True Then
        Set objSMS = objLoc.ConnectServer(Loc.Machine, "root\sms\site_" & _
            Loc.SiteCode)
            strSMSSiteCode = Loc.Sitecode
    end if
Next

Set objFolder = objSMS.Get("SMS_ObjectContainerItem")
arrAdvertIDs = split(strAdvertIDs, ",")
retval = objFolder.MoveMembers _
(arrAdvertIDs, intSourceFolder, intDestFolder, intObjectType)
