strSMSServer = <SMSServer>
strParentCollID = "COLLROOT" 'This creates the collection in 
'the collection root.  Replace COLLROOT with the CollectionID
'of an existing collection to make the new collection a child

strCollectionName = "Systems Without Windows Update Agent"
strCollectionComment = "This collection contains all systems " & _
    "that do not have the Windows Update Agent installed."
    
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
    
Set newCollection = objSMS.Get("SMS_Collection").SpawnInstance_()

newCollection.Name = strCollectionName
newCollection.OwnedByThisSite = True
newCollection.Comment = strCollectionComment
path=newCollection.Put_

'the following two lines are used to obtain the CollectionID
'of the collection we just created
Set Collection=objSMS.Get(path)
strCollID= Collection.CollectionID
'now we create a relationship betwen the new collection
'and it's parent.

Set newCollectionRelation = objSMS.Get _
    ( "SMS_CollectToSubCollect" ).SpawnInstance_()
newCollectionRelation.parentCollectionID = strParentCollID
newCollectionRelation.subCollectionID = strCollID
newCollectionRelation.Put_
