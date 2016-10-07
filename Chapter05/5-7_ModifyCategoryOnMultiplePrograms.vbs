strSMSServer = <SMSServer>

strCategory = "SMS Admin"

strPkgList = "LAB00002,LAB00006,LAB00003,LAB00005"
arrPackages = split(strPkgList,",")

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
             
For each strPackageID in arrPackages
    Set colPrograms = objSMS.ExecQuery _
        ("Select * From SMS_Program " & _
        "WHERE PackageID='" & strPackageID & "'")
    For Each objProgram In colPrograms
        wscript.echo "Modifying category for:" & _
            objProgram.ProgramName
        objProgram.Description = strCategory
        objProgram.Put_
    Next
Next