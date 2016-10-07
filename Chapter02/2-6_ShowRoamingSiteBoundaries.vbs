Set objArgs = WScript.Arguments
strSMSServer = objArgs(0)
strSiteToDisplay = ucase(objArgs(1))

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

strInfo = "Roaming Site Boundary Informaiton for Site " & _
    strSiteToDisplay & vbCRLF
        
Set boundaries=objSMS.Get("SMS_SCI_RoamingBoundary." & _
    "SiteCode='" & strSiteToDisplay & "',Filetype=2,ItemName='" & _
    "Roaming Boundary',ItemType='Roaming Boundary'")
if boundaries.IncludeSiteBoundary then
    strInfo = strInfo &  "Site Boundaries are included in the "& _
    "local roaming boundaries." & vbCRLF
else
    strInfo = strInfo &  "Site Boundaries are NOT included in the "& _
    "local roaming boundaries." & vbCRLF
end if
msgbox ubound(boundaries.Details)
For i=0 to ubound(boundaries.Details)
    if boundaries.Flags(i) = 1 then
        strBoundary = "Remote Boundary"
    else
        strBoundary = "Local Boundary"
    end if
    strInfo = strInfo & _
        boundaries.Types(i) & _
        ": " & boundaries.Details(i) & vbTAB & strBoundary & vbCRLF
Next
wscript.echo strInfo
