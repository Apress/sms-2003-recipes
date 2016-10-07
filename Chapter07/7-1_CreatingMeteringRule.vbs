strSMSServer = <SMSServer>

strProdcutName = "WinZip"
strFileName = "WinZip32.exe"
strFileVersion = "*"
strLanguageID = 65535 '65535 = 'any language'

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

Set newMTRRule = _
    objSMS.Get("SMS_MeteredProductRule").SpawnInstance_()
    
newMTRRule.ProductName = strProdcutName
newMTRRule.FileName = strFileName
newMTRRule.FileVersion = strFileVersion
newMTRRule.LanguageID = strLanguageID
newMTRRule.SiteCode = ucase(strSMSSiteCode)
newMTRRule.ApplyToChildSites = TRUE
newMTRRule.Enabled = TRUE
newMTRRule.Put_

