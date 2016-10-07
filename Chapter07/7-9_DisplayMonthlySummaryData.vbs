strSMSServer = <SMSServer>

strUserName = "ramseyg"
intDays = 180

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

strWQL = "SELECT SMS_MeteredFiles.ProductName, " & _
 "SMS_MeteredUser.FullName, SMS_MonthlyUsageSummary.UsageCount, " & _
 "SMS_MonthlyUsageSummary.TSUsageCount, " & _
 "SMS_MonthlyUsageSummary.LastUsage, SMS_R_System.Name " & _
 "FROM SMS_MonthlyUsageSummary INNER JOIN " & _
 "SMS_R_System ON SMS_MonthlyUsageSummary.ResourceID = " & _
 "SMS_R_System.ResourceID INNER JOIN " & _
 "SMS_MeteredUser ON SMS_MonthlyUsageSummary.MeteredUserID = " & _
 "SMS_MeteredUser.MeteredUserID INNER JOIN " & _
 "SMS_MeteredFiles ON SMS_MonthlyUsageSummary.FileID = " & _
 "SMS_MeteredFiles.MeteredFileID " & _
 "WHERE (SMS_MeteredUser.UserName = '" & strUserName & "') " & _
 "and datediff" & _
 "(day, SMS_MonthlyUsageSummary.LastUsage, getdate()) <= " & _
 intDays & "ORDER BY SMS_MonthlyUsageSummary.LastUsage"
 
wscript.echo "Product Name" & vbTAB & "Domain\UserName" & vbTAB & _
    "Usage Count" & vbTAB & "TSUsage Count" & vbTAB & _
    "ComputerName" & vbTAB & "LastUsageRecordedForMonth"
    
Set colMTRResults = objSMS.ExecQuery(strWQL)
for each objMTRResult in colMTRResults
    wscript.echo objMTRResult.SMS_MeteredFiles.ProductName & vbTAB & _
    objMTRResult.SMS_MeteredUser.FullName & vbTAB & _
    objMTRResult.SMS_MonthlyUsageSummary.UsageCount & vbTAB & _
    objMTRResult.SMS_MonthlyUsageSummary.TSUsageCount & vbTAB & _
    objMTRResult.SMS_R_System.Name & vbTAB & _
        WMIDateStringToDate(objMTRResult.SMS_MonthlyUsageSummary.LastUsage)
next

Function WMIDateStringToDate(dtmInstallDate)
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
        Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
            & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
                Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
                    13, 2))
End Function

