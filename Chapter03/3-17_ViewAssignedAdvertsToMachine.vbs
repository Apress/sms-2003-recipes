strSMSServer = <SMSServer>

strComputer = "Computer1"

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

'first, get the resource ID of the computer
intResourceID = GetResourceID(strComputer)

Set colAdverts = objSMS.ExecQuery _
("select * from SMS_ClientAdvertisementStatus where ResourceID = " & _
     intResourceID)
for each strAdvert in colAdverts 'enumerate all adverts for client
    wscript.echo GetAdvertisementName(strAdvert.AdvertisementID) & _ 
        "(" & strAdvert.AdvertisementID & ")" & vbTAB  & _
    strAdvert.LastStateName & vbTAB & strAdvert.LastStatusTime & _
    vbTAB & GetCollectionName(strAdvert.AdvertisementID)
next

'used to obtain the Advertisement Name
Function GetAdvertisementName(strAdvertID)
    Set instAdvert = objSMS.Get _
        ("SMS_Advertisement.AdvertisementID='" & strAdvertID & "'")
    GetAdvertisementName = instAdvert.AdvertisementName
End Function

'used to obtain the Collection Name
Function GetCollectionName(strAdvertID)
    'first, get advert based on advert ID
    Set instAdvert = objSMS.Get _
        ("SMS_Advertisement.AdvertisementID='" & strAdvertID & "'")
    'then, get collection name based on collectionID from advert
    Set instCollection = objSMS.Get _
        ("SMS_Collection.CollectionID='" & instAdvert.CollectionID & "'")
    GetCollectionName = instCollection.Name
End Function

'used to obtain the SMS resource ID
Function GetResourceID(strComputerName)
    Set colResourceIDs = objSMS.ExecQuery _
        ("select ResourceID from SMS_R_System where Name = '" & _
             strComputer & "'")
    for each objResID in colResourceIDs
        GetResourceID = objResID.ResourceID
    next
End Function
