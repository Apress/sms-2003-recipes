strSMSServer = <SMSServer>'primary site server

strSMSSiteCodeforDP = <SMSDPSiteCode>'site code that
                        ' contains packages to remove
strServerSharePath = "\\sms-testgmr\smspkgd$" 
                            'DP path

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
    
Set objDPs = objSMS.ExecQuery("select * from " & _
    "SMS_DistributionPoint where SiteCode = '" & _
    strSMSSiteCodeforDP & "'")

For each objDP in objDPs 
    if instr(ucase(objDP.ServerNALPath), _
        ucase(strServerSharePath)) then
         wscript.echo "Removing " & objDP.PackageID & _
             vbTAB & objDP.ServerNALPath   
        objDP.Delete_ 
    end if
Next
