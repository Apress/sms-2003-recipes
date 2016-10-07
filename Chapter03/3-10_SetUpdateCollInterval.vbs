strSMSServer = <SMSServer>
strCollID = "LAB0002B"

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
    
Set Token = objSMS.Get("SMS_ST_RecurInterval")
Token.DaySpan = 1
Token.StartTime = "20051202103000.000000+***" 'wmi date-string
'If omitted, StartTime = Jan 1, 1990 - this shouldn't
'cause any issues
Set objCollection = objSMS.Get _
    ("SMS_Collection.CollectionID='"  & strCollID & "'")
objCollection.RefreshSchedule = Array(Token)
objCollection.RefreshType = 2  'Periodic refresh
objCollection.Put_
