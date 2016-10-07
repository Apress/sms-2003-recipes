strSMSServer = <SMSServer>
strCollID = "LAB000FE"

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
        
ListAdverts strCollID, False
CheckParent strCollID
    
Sub ListAdverts(strCollID,blnSubCollect)
    If blnSubCollect then
        'This is to check all parent questions of the
        'collection in question. We can only look at
        'parent collections that are set to "Include
        'members of subcollections"
        Set colAdverts = objSMS.ExecQuery _
        ("select * from SMS_Advertisement where " & _
        "CollectionID = '" & strCollID & "' and " & _
        "IncludeSubCollection = 1")
    Else
        'This is for the first check, this will look for
        'advertisements assigned directly to the collection
        Set colAdverts = objSMS.ExecQuery _
         ("select * from SMS_Advertisement where " & _
         "CollectionID = '" & strCollID & "'")
    end if
    for each objAdvert in colAdverts
        wscript.echo "Advertisement: " & _
           objAdvert.AdvertisementName
        wscript.echo "   Collection: " & _
           GetCollectionName(objAdvert.CollectionID) & _
           " (" & objAdvert.CollectionID  & ")" & VbCRLF
    next
End Sub

Function GetCollectionName(strCollID)
    Set instColl = objSMS.Get _
    ("SMS_Collection.CollectionID=""" & strCollID & """")
    GetCollectionName = instColl.Name
End Function

Sub CheckParent(strCollID)
    Set colParents = objSMS.ExecQuery _
    ("select * from SMS_CollectToSubCollect where subCollectionID = '" & _
    strCollID & "'")
    for each objParent in colParents
        if not objParent.ParentCollectionID = "COLLROOT" then
            ListAdverts objParent.ParentCollectionID , True
            CheckParent objParent.ParentCollectionID
        end if
    next    
End Sub
