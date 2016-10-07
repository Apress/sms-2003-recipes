strSMSServer = <SMSServer>
strADSite = "Cleveland"

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
    ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
    If Loc.ProviderForLocalSite = True Then
        Set objSMS = objLoc.ConnectServer(Loc.Machine, "root\sms\site_" & _
            Loc.SiteCode)
            strSMSSiteCode = Loc.SiteCode
    end if
Next
    
Set WbemContext=CreateObject("WbemScripting.SWbemNamedValueSet")
WbemContext.Add "SessionHandle", objSMS.ExecMethod _
    ("SMS_SiteControlFile", "GetSessionHandle").SessionHandle
objSMS.ExecMethod "SMS_SiteControlFile.Filetype=1,SiteCode='" _
    & strSMSSiteCode & "'", "Refresh", , , WbemContext
'retrieve boundary details
Set WbemInst = objSMS.Get _
    ("SMS_SCI_SiteAssignment.Filetype=2,Itemtype='Site Assignment'," _
        & "SiteCode='" & strSMSSiteCode & _
        "',ItemName='Site Assignment'", , WbemContext)
proparray1 = WbemInst.AssignDetails
proparray2 = WbemInst.AssignTypes

onemore = ubound(proparray1) + 1
redim preserve proparray1( onemore ) 'add one to size of array
redim preserve proparray2( onemore )
proparray1( onemore ) = strADSite
proparray2( onemore ) = "Active Directory site"
WbemInst.AssignDetails = proparray1
WbemInst.AssignTypes = proparray2
WbemInst.Put_ , WbemContext
objSMS.ExecMethod "SMS_SiteControlFile.Filetype=0,SiteCode=""" & _
    strSMSSiteCode & """", "Commit", , , WbemContext
objSMS.Get("SMS_SiteControlFile").ReleaseSessionHandle _
    WbemContext.Item("SessionHandle").Value
