on error resume next
strSMSServer = <SMSServer>

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

Set colCollections = objSMS.ExecQuery _
("select * from SMS_Collection order by Name")
For each objCollection in colCollections
    wscript.echo objCollection.Name & vbTAB & _
    	GetRecurSchedule(objCollection.CollectionID)
next
        
Function GetRecurSchedule(strCollID)
	Set objCollection = objSMS.Get _
	    ("SMS_Collection.CollectionID='"  & strCollID & "'")
	if objCollection.RefreshType = 1 then
		GetRecurSchedule = "Manual Refresh"
	else	    
		for each objSched in objCollection.RefreshSchedule
			GetRecurSchedule = objSched.GetObjectText_
		next
	end if
End Function