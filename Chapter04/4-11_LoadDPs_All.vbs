strSMSServer = <SMSServer>

PackageID = "LAB0000A"

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

Set AllDPs = objSMS.ExecQuery _
    ("Select * From SMS_SystemResourceList WHERE " & _
    "RoleName='SMS Distribution Point'")
For Each DP In AllDPs
    Wscript.echo DP.SiteCode & vbTAB & DP.ServerName
    Set Site = objSMS.Get("SMS_Site='" & DP.SiteCode & "'")
    Set newDP = objSMS.Get _
        ("SMS_DistributionPoint").SpawnInstance_()
    newDP.ServerNALPath = DP.NALPath
    newDP.PackageID = PackageID
    newDP.SiteCode = DP.SiteCode
    newDP.SiteName = Site.SiteName
    newDP.Put_
Next
