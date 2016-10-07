strSMSServer = <SMSServer>

strProductName = "WinZip"

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

Set colMTRRules = objSMS.ExecQuery _
    ("Select * From SMS_MeteredProductRule where " & _
        "ProductName = '" & strProductName & "'")
For Each objMTRRule In colMTRRules
    wscript.echo "Deleting " & objMTRRule.ProductName
    objMTRRule.Delete_
Next