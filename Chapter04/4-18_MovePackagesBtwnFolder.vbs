strSMSServer = <SMSServer>
strSMSSiteCode = <SMSSiteCode>
strPackageIDs = "LAB00004,LAB00001"
'to move more than one, separate with commas
intSourceFolder=0 'Source Folder (root node in this case)
intDestFolder=4 'Destination Folder
intObjectType=2 '2=Package for type of object
Set loc = CreateObject("WbemScripting.SWbemLocator")
Set objSMS = loc.ConnectServer(SMSServer, "root\SMS\site_" & _
    SMSSiteCode)
Set objFolder = objSMS.Get("SMS_ObjectContainerItem")
arrPackageIDs = split(strPackageIDs, ",")
retval = objFolder.MoveMembers _
    (arrPackageIDs, intSourceFolder , intDestFolder , intObjectType)
