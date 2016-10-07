strSMSServer = <SMSServer>

strPackageID = "LAB0000A"
StrSiteCode = "CLE"

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

Set DPs = objSMS.ExecQuery _
    ("Select * From SMS_DistributionPoint " & _
    "WHERE SiteCode='" & strSiteCode & _
    "' AND PackageID='" & strPackageID & "'")
For Each DP In DPs
    DP.RefreshNow = True
    DP.Put_
Next
