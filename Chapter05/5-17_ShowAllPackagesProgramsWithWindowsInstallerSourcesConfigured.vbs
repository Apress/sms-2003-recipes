strSMSServer = <SMSServer>

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

Set colPrograms = objSMS.ExecQuery _
    ("Select * From SMS_Program " & _
    "WHERE MSIProductID <> '' order by PackageID")

for each objProgram in colPrograms
    wscript.echo "Package Name: " & GetPackageName _
        (objProgram.PackageID) & "(" & objProgram.PackageID & ")"
    wscript.echo vbTAB & "Program Name: " & _
        objProgram.ProgramName
    wscript.echo vbTAB & vbTAB & "MSIFilePath MSIProductID: " & _
        objProgram.MSIFilePath & " " & objProgram.MSIProductID
next        

Function GetPackageName(strPckID)
    Set objPackage=objSMS.Get("SMS_Package.PackageID='" & strPckID & "'")
    GetPackageName = objPackage.Name
End Function
