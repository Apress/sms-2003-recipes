strSMSServer = <SMSServer>

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
'Use the "SecuirtyKey" ID as the Report ID
strReportIDs = "LAB00001,LAB00002,LAB00004,LAB00005,LAB00023"    
'to move more than one, separate with commas
intSourceFolder=0 'Source Folder (root node in this case)
intDestFolder=9 'Destination Folder
intObjectType=8 '8=Report for type of object

Set objFolder = objSMS.Get("SMS_ObjectContainerItem")
arrReportIDs = split(strReportIDs, ",")
retval = objFolder.MoveMembers _
(arrReportIDs, intSourceFolder, intDestFolder , intObjectType)
wscript.echo retval
