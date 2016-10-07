CONST SINCE_ADVERTISED = "0001128000080008"

strSMSServer = <SMSServer>

strAdvertID = "LAB20014"

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

Set colAdvertStatus = objSMS.ExecQuery _
    ("Select * From SMS_AdvertisementStatusSummarizer where" & _
        " AdvertisementID = '" & strAdvertID & _
        "' and DisplaySchedule = '" & SINCE_ADVERTISED &  "'")
 for each objAS in colAdvertStatus
    wscript.echo objAS.SiteCode  & vbTAB & _
        objAS.AdvertisementsReceived & vbTAB & _
        objAS.AdvertisementsFailed & vbTAB & _
        objAS.ProgramsStarted & vbTAB & _
        objAS.ProgramsFailed & vbTAB & _ 
        objAS.ProgramsSucceeded & vbTAB & _
        objAS.ProgramsFailedMIF & vbTAB & _
        objAS.ProgramsSucceededMIF
 next       
            

'TABLE -- Tally Intervals -- from the SDK
' CONST SINCE_ADVERTISED = "0001128000080008"
' CONST SINCE12_00_AM = "0001128000100008"
' CONST SINCE06_00AM = "00C1128000100008"
' CONST SINCE12_00_PM = "0181128000100008"
' CONST SINCE06_00_PM = "0241128000100008"
' CONST SINCE_SUNDAY = "0001128000192000"
' CONST SINCE_MONDAY = "00011280001A2000"
' CONST SINCE_TUESDAY = "00011280001B2000"
' CONST SINCE_WEDNESDAY = "00011280001C2000"
' CONST SINCE_THURSDAY = "00011280001D2000"
' CONST SINCE_FRIDAY = "00011280001E2000"
' CONST SINCE_SATURDAY = "00011280001F2000"
' CONST SINCE_1ST_OF_MONTH = "000A470000284400"
' CONST SINCE_15TH_OF_MONTH = "000A4700002BC400"