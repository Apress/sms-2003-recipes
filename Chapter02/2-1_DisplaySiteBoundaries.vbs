Set objArgs = WScript.Arguments
strSMSServer = objArgs(0)
strSiteToDisplay = objArgs(1)

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

strInfo = "Site Boundary Informaiton for Site " & _
    strSiteToDisplay & vbCRLF
Set boundaries=objSMS.Get _
    ("SMS_SCI_SiteAssignment.SiteCode='" & _
    strSiteToDisplay & "',Filetype=1,ItemName='" & _
    "Site Assignment',ItemType='Site Assignment'")

For i=0 to ubound(boundaries.AssignDetails)
    strInfo = strInfo & space(len(descr)) & _
        boundaries.AssignTypes(i) & _
        ": " & boundaries.AssignDetails(i) & vbCRLF
Next
wscript.echo strInfo
