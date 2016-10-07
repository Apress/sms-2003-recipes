Set objArgs = WScript.Arguments
strSMSServer = objArgs(0)
strSiteToConfig= objArgs(1)
strBoundaryType = ucase(objArgs(2))
strBoundaryData = objArgs(3)
'creating 2 arrays of size one
Dim inputArray1(1)
Dim inputArray2(1)

select case strBoundaryType
    case "SUBNET"
        strBoundaryType = "IP Subnet"
     case "AD"
         strBoundaryType = "Active Directory site"
End Select

inputArray1(0) = trim(strBoundaryData)
inputArray2(0) = trim(strBoundaryType)

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
    ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
    If Loc.ProviderForLocalSite = True Then
        Set objSWbemServices = objLoc.ConnectServer _
            (Loc.Machine, "root\sms\site_" & Loc.SiteCode)
    end if
Next
    
    
Set objSWbemContext=CreateObject _
    ("WbemScripting.SWbemNamedValueSet")
objSWbemContext.Add "SessionHandle", _
    objSWbemServices.ExecMethod("SMS_SiteControlFile", _
    "GetSessionHandle").SessionHandle
objSWbemServices.ExecMethod _
    "SMS_SiteControlFile.Filetype=1,Sitecode='" & _
    strSiteToConfig & "'", "RefreshSCF", , , objSWbemContext
Set objSWbemInst = objSWbemServices.Get _
    ("SMS_SCI_SiteAssignment.Filetype=2,Itemtype='" & _
    "Site Assignment',Sitecode='" & strSiteToConfig & _
    "',ItemName='Site Assignment'", , objSWbemContext)

'Retrieve the boundary details.
proparray1 = objSWbemInst.AssignDetails
proparray2 = objSWbemInst.AssignTypes

if ubound(objSWbemInst.AssignDetails)=-1 then
    'There are no boundaries so create an array.
    bounds=0
    redim proparray1(0)
    redim proparray2(0)
    proparray1(bounds)=inputArray1(0)
    proparray2(bounds)=inputArray2(0)
Else
    bounds=ubound (objSWbemInst.AssignDetails)+1
    'Increase array for new boundaries
    ReDim Preserve proparray1 (ubound (proparray1) + _
        ubound (inputArray1))
    ReDim Preserve proparray2 (ubound (proparray2) + _
        ubound (inputArray2))
    for i= 0 to ubound(inputArray1)-1 'Add boundaries
        proparray1(bounds+i)=inputArray1(i)
        proparray2(bounds+i)=inputArray2(i)
    Next
End If

objSWbemInst.AssignDetails = proparray1
objSWbemInst.AssignTypes = proparray2
objSWbemInst.Put_ , objSWbemContext

objSWbemServices.ExecMethod _
    "SMS_SiteControlFile.Filetype=0,Sitecode=""" & _
    strSiteToConfig & """", "Commit", , , objSWbemContext
objSWbemServices.Get("SMS_SiteControlFile"). _
    ReleaseSessionHandle objSWbemContext.Item _
    ("SessionHandle").Value
