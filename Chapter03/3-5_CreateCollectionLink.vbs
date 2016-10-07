strSMSServer = <SMSServer>

strParentColl = "LAB0000D" 'parent collection ID
strSubColl = "LAB0000E"  'collection ID to link

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

Set ColSvcs = objSMS.Get("SMS_Collection")
    ColSvcs.VerifyNoLoops "SMS_Collection.CollectionID=""" & _
    strParentColl & """", "SMS_Collection.CollectionID=""" & _
    strSubColl & """", Result

if Result = 0 then
    wscript.echo "Link would cause looping condition, exiting"
else
    Set objCol = objSMS.Get _
        ("SMS_CollectToSubCollect").SpawnInstance_()
    objCol.parentCollectionID = strParentColl
    objCol.subCollectionID = strSubColl
    objCol.Put_
    wscript.echo "Created Collection Link!"
end if