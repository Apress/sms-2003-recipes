strSMSServer = <SMSServer>

strComputer = "2kPro"

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

strSQL = "SELECT sys.Name, colFil.FileName, " & _
    "colFil.FileSize, colFil.ModifiedDate, colFil.FilePath, " & _
    "colFil.LocalFilePath, colFil.CollectionDate " & _
    "FROM SMS_G_System_CollectedFile colFil INNER JOIN " & _
        "SMS_R_System sys ON " & _
         "colFil.ResourceID = sys.ResourceID " & _
    "WHERE sys.Name = '" & strComputer & "'"
         
Set colFiles = objSMS.ExecQuery(strSQL)
for each objFile in colFiles
    wscript.echo objFile.sys.Name & vbTAB & _
        objFile.colFil.FileName & vbTAB & _
        objFile.colFil.FileSize & vbTAB & _
        WMIDateStringToDate(objFile.colFil.ModifiedDate) & _
        vbTAB & objFile.colFil.FilePath & vbTAB & _
        objFile.colFil.LocalFilePath & vbTAB & _
        WMIDateStringToDate(objFile.colFil.CollectionDate)
next

Function WMIDateStringToDate(dtmInstallDate)
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & _
        "/" & Mid(dtmInstallDate, 7, 2) & "/" & _
        Left(dtmInstallDate, 4) & " " & _
        Mid (dtmInstallDate, 9, 2) & ":" & _
        Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
        13, 2))
End Function
