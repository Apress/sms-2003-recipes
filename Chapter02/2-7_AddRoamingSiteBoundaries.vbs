Set objArgs = WScript.Arguments
strSMSServer = objArgs(0)
strSiteToModify = objArgs(1)
strBoundaryType = ucase(objArgs(2))
strBoundaryData = objArgs(3)
strRemote = ucase(objArgs(4))
Dim inputArray1(1)
Dim inputArray2(1)
Dim inputArray3(1)

select case strBoundaryType
    case "SUBNET"
        strBoundaryType = "IP Subnets"
     case "AD"
         strBoundaryType = "AD Site Name"
     case "RANGE"
         strBoundaryType = "IP Ranges"
End Select

select case strRemote
    case "LOCAL"
        strRemote = 0
    case "REMOTE"
        strRemote = 1
end select
inputArray1(0)=strBoundaryData
inputArray2(0)=strRemote
inputArray3(0)=strBoundaryType


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
    
Set objSWbemContext = CreateObject _
    ("WbemScripting.SWbemNamedValueSet")
objSWbemContext.Add "SessionHandle", _
    objSMS.ExecMethod("SMS_SiteControlFile", _
    "GetSessionHandle").SessionHandle
objSMS.ExecMethod _
    "SMS_SiteControlFile.Filetype=1,Sitecode='" & _
    strSiteToModify & "'", "RefreshSCF", , , objSWbemContext
Set objSWbemInst = objSMS.Get _
    ("SMS_SCI_RoamingBoundary.Filetype=2,Itemtype='" & _
    "Roaming Boundary',Sitecode='" & strSiteToModify & _
    "',ItemName='Roaming Boundary'", , objSWbemContext)

'Retrieve the roaming boundary details.
proparray1 = objSWbemInst.Details
proparray2 = objSWbemInst.Flags
proparray3 = objSWbemInst.Types

if ubound(objSWbemInst.Details)=-1 then
    'There are no boundaries so create an array.
    bounds=0
    redim proparray1(0)
    redim proparray2(0)
    redim proparray3(0)
    proparray1(bounds)=inputArray1(0)
    proparray2(bounds)=inputArray2(0)
    proparray3(bounds)=inputArray3(0)
Else
    bounds=ubound (objSWbemInst.Details)+1
    'Increase array for new boundaries
    ReDim Preserve proparray1 (ubound (proparray1) + _
        ubound (inputArray1))
    ReDim Preserve proparray2 (ubound (proparray2) + _
        ubound (inputArray2))
    ReDim Preserve proparray3 (ubound (proparray3) + _
        ubound (inputArray3))
    for i= 0 to ubound(inputArray1)-1 'Add boundaries
        proparray1(bounds+i)=inputArray1(i)
        proparray2(bounds+i)=inputArray2(i)
        proparray3(bounds+i)=inputArray3(i)
    Next
End If

objSWbemInst.Details = proparray1
objSWbemInst.Flags = proparray2
objSWbemInst.Types = proparray3
objSWbemInst.Put_ , objSWbemContext

objSMS.ExecMethod _
    "SMS_SiteControlFile.Filetype=0,Sitecode=""" & _
    strSiteToModify & """", "Commit", , , objSWbemContext
objSMS.Get("SMS_SiteControlFile"). _
    ReleaseSessionHandle objSWbemContext.Item _
    ("SessionHandle").Value
