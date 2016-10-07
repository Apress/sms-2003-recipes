Const strSMSServer = <SMSServer>

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
    ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
    If Loc.ProviderForLocalSite = True Then
        Set objSMS = objLoc.ConnectServer(Loc.Machine, "root\sms\site_" & _
            Loc.SiteCode)
            strSMSSiteCode = Loc.Sitecode
    end if
Next

'for intPackageFolder=0 enumerate from the root package node
' to enumerate from a subfolder, replace the 0 with
' the proper ContainerNodeID
intPackageFolder = 0
intSpace = 0
'First, display packages in current folder (as defined
' by intPackageFolder
ListPkgsInFolder "", intPackageFolder, intSpace
'Then enumerate all subfolders of the current folder
DisplaySubFolders intPackageFolder, intSpace + 3

Sub DisplaySubFolders(intPackageFolder, intSpace)
        Set colItems = objSMS.ExecQuery _
            ("select * from SMS_ObjectContainerNode where " & _
            "ObjectType = 2 and " & _
            "ParentContainerNodeID = " & cint(intPackageFolder))
        For each objContainerItem in colItems
            ListPkgsInFolder objContainerItem.Name, _
                objContainerItem.ContainerNodeID,intSpace + 3
            DisplaySubFolders objContainerItem.ContainerNodeID, _
                intSpace + 3                
        Next
End Sub

Sub ListPkgsInFolder(strContainerName, intPkgFolder, intSpace)
    If intPkgFolder = 0 then   'we're looking at the root node here

        strSQL = "Select PackageID From SMS_Package where " & _
            "PackageID not in (select InstanceKey " & _
            "from SMS_ObjectContainerItem where ObjectType=2) " & _
            "and ImageFlags = 0"            
        'ImageFlags = 0 is used to make sure we don't pick up
        'and image packages from the Operating System Deployment
        'Feature Pack (OSD FP)                
        Set colPkgs = objSMS.ExecQuery(strSQL)
        wscript.echo space(intSpace) & "----" & vbCRLF & _
            space(intSpace) & "Root Collection" 
        For each objPkg in colPkgs
            wscript.echo objPkg.PackageID & vbTAB & _
                GetPackageName(objPkg.PackageID)
        Next
    Else
        Set colItems = objSMS.ExecQuery _
            ("select * from SMS_ObjectContainerItem where " & _
            "ObjectType = 2 and " & _ 
            "ContainerNodeID = " & cint(intPkgFolder))
        wscript.echo space(intSpace) & "----" & vbCRLF & _
            space(intSpace) &  strContainerName
        For each objContainerItem in colItems
            strPackageID = objContainerItem.InstanceKey
            wscript.echo space(intSpace) & strPackageID & vbTAB & _
                 GetPackageName(strPackageID)
        Next
    End If
End Sub

Function GetPackageName(strPckID)
    Set objPackage = GetObject( "WinMgmts:\\" & strSMSServer & _
        "\root\SMS\site_" & strSMSSiteCode & _
        ":SMS_Package.PackageID='" & strPckID & "'")
    GetPackageName =  objPackage.Name & vbTAB & objPackage.Version
End Function
