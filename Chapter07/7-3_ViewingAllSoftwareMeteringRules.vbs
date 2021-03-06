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

Set colMTRs = objSMS.ExecQuery _
    ("Select * From SMS_MeteredProductRule order by " & _
        " ProductName")
for each objMTRRule in colMTRS
    wscript.echo objMTRRule.ProductName & vbTAB & _
        objMTRRule.FileName & vbTAB & objMTRRule.FileVersion & _
        vbTAB & objMTRRule.SiteCode & vbTAB & _
        objMTRRule.ApplyToChildSites & vbTAB & objMTRRule.Enabled
next