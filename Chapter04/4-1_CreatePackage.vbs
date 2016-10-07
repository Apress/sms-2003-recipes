strSMSServer = <SMSServer>

'Specify package attributes one time. From here on out
'we reference these variables.
pkgName = "Visual Studio .NET Framework 1.1 SP1"
pkgSource = "\\smsvpc\source\KB867460"
pkgDesc = "This Package Installs .NET 1.1 SP1"
pkgManufacturer = "Microsoft"

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

'Define package attributes such as name, source for the files, etc.
Set newPackage = objSMS.Get("SMS_Package").SpawnInstance_()
newPackage.Name = pkgName
newPackage.Description = pkgDesc
newPackage.Manufacturer = pkgManufacturer
newPackage.PkgSourceFlag = 2 
'2=direct, 1=no source 3=use compressed source
newPackage.PkgSourcePath = pkgSource
path=newPackage.Put_

'the following three lines are used to obtain the PackageID
'of the package we just created
Set Package=objSMS.Get(path)
PackageID= Package.PackageID
wscript.echo PackageID & " = " & pkgName
