strSMSServer = <SMSServer>

strFileName = "Acrobat.exe"

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

strSQL = "select SMS_R_System.Name, " & _
    "SMS_G_System_SoftwareFile.FileName, " & _
    "SMS_G_System_SoftwareFile.FileDescription, " & _
    "SMS_G_System_SoftwareFile.FileVersion from " & _
    "SMS_R_System inner join SMS_G_System_SoftwareFile on " & _
    "SMS_G_System_SoftwareFile.ResourceID = " & _
    "SMS_R_System.ResourceId where " & _
    "SMS_G_System_SoftwareFile.FileName = '" & strFileName & "'"
    
Set colSystems = objSMS.ExecQuery(strSQL)

for each objSystem in colSystems
    wscript.echo objSystem.SMS_R_System.Name & vbTAB & _
    objSystem.SMS_G_System_SoftwareFile.FileName & vbTAB & _
    objSystem.SMS_G_System_SoftwareFile.FileDescription & vbTAB & _
    objSystem.SMS_G_System_SoftwareFile.FileVersion
next
