Const STORAGE_DIRECT=2
strSMSServer = <SMSServer>
strPackageID = "LAB0000A"

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

Set Token = objSMS.Get("SMS_ST_RecurWeekly")
Token.Day = 5 'Thursday
Token.DayDuration = 0 'recur indefinitely
Token.ForNumberOfWeeks = 1 'recur every 1 week
Token.StartTime = "20061202103000.000000+***" 'wmi date-string
'If omitted, StartTime = Jan 1, 1990 - this shouldn't
'cause any issues
Set objPackage = objSMS.Get _
    ("SMS_Package.PackageID='"  & strPackageID & "'")
'Make sure package is set to "Always obtain files
' from source directory, then add schedule.
if objPackage.PkgSourceFlag=STORAGE_DIRECT then
    objPackage.RefreshSchedule = Array(Token)
    objPackage.Put_
end if
