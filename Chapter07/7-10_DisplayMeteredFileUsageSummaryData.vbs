strSMSServer = <SMSServer>

intDays = 10
strProductName = "sol"

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
 "SMS_FileUsageSummary.DistinctUserCount, " & _
 "SMS_FileUsageSummary.IntervalStart, " & _
 "SMS_FileUsageSummary.IntervalWidth, " & _
 "SMS_FileUsageSummary.SiteCode " & _
 "FROM SMS_MeteredFiles INNER JOIN SMS_FileUsageSummary " & _
 "ON SMS_MeteredFiles.MeteredFileID = " & _
 "SMS_FileUsageSummary.FileID WHERE " & _
 "datediff(day,SMS_FileUsageSummary.IntervalStart, getdate()) " & _
 "<= " & intDays & " AND " & _
 "(SMS_FileUsageSummary.IntervalWidth = 15) AND " & _
 "(SMS_MeteredFiles.ProductName = '" & strProductName & "')"
                        
wscript.echo "Product Name" & vbTAB & "Metered Interval" & vbTAB & _
    "Distinct Users" & vbTAB & "Site Code"
Set colMTRResults = objSMS.ExecQuery(strWQL)
for each objMTRResult in colMTRResults
    wscript.echo objMTRResult.SMS_MeteredFiles.ProductName & _
        vbTAB & WMIDateStringToDate(objMTRResult. _
        SMS_FileUsageSummary.IntervalStart) & _
        vbTAB & objMTRResult.SMS_FileUsageSummary.DistinctUserCount & _
        vbTAB & objMTRResult.SMS_FileUsageSummary.SiteCode
    if intPeak < objMTRResult.SMS_FileUsageSummary.DistinctUserCount then
        intPeak = objMTRResult.SMS_FileUsageSummary.DistinctUserCount
        strPeakInfo = WMIDateStringToDate(objMTRResult. _
            SMS_FileUsageSummary.IntervalStart)  & vbTAB & _
            objMTRResult.SMS_FileUsageSummary.DistinctUserCount _
            & vbTAB & objMTRResult.SMS_FileUsageSummary.SiteCode _
            & vbCRLF
    elseif intPeak = _
        objMTRResult.SMS_FileUsageSummary.DistinctUserCount then
            strPeakInfo = strPeakInfo & WMIDateStringToDate(objMTRResult. _
            SMS_FileUsageSummary.IntervalStart)  & vbTAB & _
            objMTRResult.SMS_FileUsageSummary.DistinctUserCount _
            & vbTAB & objMTRResult.SMS_FileUsageSummary.SiteCode _
            & vbCRLF
    end if
next

wscript.echo vbCRLF & vbCRLF
wscript.echo "Peak concurrentusage over the past " & intDays & _
    " days of metering data:" & vbCRLF & strPeakInfo

Function WMIDateStringToDate(dtmInstallDate)
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
        Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
            & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
                Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
                    13, 2))
End Function
        
