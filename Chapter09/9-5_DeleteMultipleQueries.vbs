strSMSServer = <SMSServer>

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

'to delete more than one, separate with commas as shown below
strQueryIDs = "LAB00001,LAB00002,LAB00014"
arrQueryIDs = split(strQueryIDs, ",")

for each strQueryID in arrQueryIDs
    Set objQuery = objSMS.Get ("SMS_Query.QueryID='" & strQueryID & "'")
    objQuery.Delete_
next
