strSMSServer = <SMSServer>

strCollID = "LAB00017"

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
    ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
    If Loc.ProviderForLocalSite = True Then

            strSMSSiteCode = Loc.Sitecode
    end if
Next

Set objCollection = GetObject( "WinMgmts:!\\" & strSMSServer & _
     "\root\SMS\site_" & strSMSSiteCode & _
    ":SMS_Collection.CollectionID='" & strCollID & "'")
objCollection.RequestRefresh False
