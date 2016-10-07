strSMSServer = <SMSServer>

strComputer = "Computer1"

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
    ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
    If Loc.ProviderForLocalSite = True Then
        Set objSMS = objLoc.ConnectServer(Loc.Machine, "root\sms\site_" & _
            Loc.SiteCode)
        strSMSSiteCode = Loc.SiteCode
    end if
Next

'get the resource ID of the computer
intResourceID = GetResourceID(strComputer)

'Remove ResourceID
Set objResource = GetObject( "WinMgmts:\\" & strSMSServer & _
    "\root\SMS\site_" & strSMSSiteCode & _
    ":SMS_R_System.ResourceID=" & cint(intResourceID))
objResource.Delete_
wscript.echo "Deleted " & strComputer & "(" & intResourceID & ")"

Function GetResourceID(strComputerName)
    Set colResourceIDs = objSMS.ExecQuery _
        ("select ResourceID from SMS_R_System where Name = '" & _
             strComputer & "'")
    for each objResID in colResourceIDs
        GetResourceID = objResID.ResourceID
    next
End Function
