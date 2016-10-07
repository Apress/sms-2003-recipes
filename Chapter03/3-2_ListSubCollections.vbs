strSMSServer = <SMSServer>

'If you want to start from the root collection, set
'   strCollID = "COLLROOT"
strCollID = "LAB00080"

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

wscript.echo strCollID & vbTAB & GetCollectionName(strCollID)

DisplaySubCollections strCollID, 3

Sub DisplaySubCollections(strCollID, intSpace)

strWQL ="SELECT col.* FROM SMS_Collection as col " & _
        "INNER JOIN SMS_CollectToSubCollect as ctsc " & _
        "ON col.CollectionID = ctsc.subCollectionID " & _
        "WHERE ctsc.parentCollectionID='" & strcollID & "' " & _
        "ORDER by col.Name"
    
    Set colSubCollections = objSMS.ExecQuery(strWQL) 

    For each objSubCollection in colSubCollections
        wscript.echo space(intSpace) & objSubCollection.CollectionID & _
            vbTAB & objSubCollection.Name
        
        DisplaySubCollections strSubCollID, intSpace + 3
    Next
End Sub

Function GetCollectionName(strCollID)
    Set objCollection = objSMS.Get _
    ("SMS_Collection.CollectionID='" & strCollID & "'")
    GetCollectionName = objCollection.Name
End Function
