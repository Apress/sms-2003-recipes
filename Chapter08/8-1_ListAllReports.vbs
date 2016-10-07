strSMSServer = <SMSServer>

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

Set colReports = objSMS.ExecQuery _
    ("Select * From SMS_Report order by " & _
        " ReportID")
    
For Each objReport In colReports
    DisplayRptInfo(objReport.ReportID)
Next

Sub DisplayRptInfo(intReportID)
    'SQLQuery is a lazy property, so we need to use the
    'Get method to retrieve the information
    Set objRpt = objSMS.Get("SMS_Report.ReportID=" & intReportID)   
    wscript.echo objRpt.ReportID & vbTAB & objRpt.Name & _
        vbTAB & objRpt.DrillThroughReportID & _
        objRpt.SecurityKey & vbTAB & objRpt.MachineDetail & _
        vbCRLF & vbCRLF & objRpt.SQLQuery & vbCRLF & vbCRLF
End Sub
