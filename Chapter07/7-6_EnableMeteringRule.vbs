strSMSServer = <SMSServer>

strProductName = "Sol"

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

Set colMTRs = objSMS.ExecQuery _
    ("Select * From SMS_MeteredProductRule where " & _
        "ProductName = '" & strProductName & "'")
for each objMTRRule in colMTRS
    if not (objMTRRule.Enabled) then
        wscript.echo "Enabling Rule " & strProductName
        objMTRRule.Enabled = True
        objMTRRule.Put_
    else
        wscript.echo "Rule " & strProductName & " aready " & _
            "enabled!"
    end if
next