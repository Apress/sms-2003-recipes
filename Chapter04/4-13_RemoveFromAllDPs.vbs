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

Set AllDPs = objSMS.ExecQuery _
    ("Select * From SMS_DistributionPoint where " & _
    "PackageID = '" & strPackageID & "'")
for each DP in AllDPs
        wscript.echo "Removing " & DP.ServerNALPath
        DP.Delete_
next
