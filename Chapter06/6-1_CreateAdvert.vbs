strSMSServer = <SMSServer>
advName = "Microsoft .NET Framework 1.1 SP1"
advCollection = "SMS000GS"
advPackageID = "LAB0000A"
advProgramName = "NET Framework 1.1 SP1"

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

Set newAdvert = objSMS.Get("SMS_Advertisement").SpawnInstance_()
newAdvert.AdvertisementName = advName
newAdvert.CollectionID = advCollection
newAdvert.PackageID = advPackageID
newAdvert.ProgramName = advProgramName
newAdvert.Put_
