strSMSServer = <SMSServer>

strAdvertID = "LAB20016"
'"occurs day 5 of every two months beginning now()"
'For advStartTime, 'Now()' is used to get the current
'date/time of the system. A properly fomatted date/time would
'just fine here also:  e.g., "12/02/2006 12:59 AM"
advStartTime = Now()

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
	
'''''''''''''Schedule Token Information		
'Create a new instance of SMS_ST_RecurInterval
Set advNewRecurToken = objSMS.Get("SMS_ST_RecurMonthlyByDate"). _
	SpawnInstance_()
advNewRecurToken.ForNumberOfMonths = 2
advNewRecurToken.MonthDay = 5 'set to zero for last day of month
advNewRecurToken.StartTime = ConvertToWMIDate(advStartTime)
'''''''''''''END Schedule Token Information

retval = AddSchedToken(advNewRecurToken, strAdvertID)

Function AddSchedToken(objSchedToken, strAdvertID)
	'enlarge the current schedule array by 1
	Set objAdvert=objSMS.Get _
		("SMS_Advertisement.AdvertisementID='" & strAdvertID & "'")
	advertArray = objAdvert.AssignedSchedule
	onemore = ubound(advertArray) +  1
	redim preserve advertArray(onemore)
	objAdvert.AssignedSchedule = advertArray
	'add the new assignment
	objAdvert.AssignedSchedule(onemore) = objSchedToken
	objadvert.AssignedScheduleEnabled = True
	AddSchedToken = objAdvert.Put_
End Function

Function ConvertToWMIDate(strDate)
	'Convert from a standard date time to wmi date
	'4/18/2005 11:30:00 AM = 2005041811300.000000+*** 
	strYear = year(strDate):strMonth = month(strDate)
	strDay = day(strDate):strHour = hour(strDate)
	strMinute = minute(strDate)
	'Pad single digits with leading zero
	if len(strmonth) = 1 then strMonth = "0" & strMonth
	if len(strDay) = 1 then strDay = "0" & strDay
	if len(strHour) = 1 then strHour = "0" & strHour
	if len(strMinute) = 1 then strMinute = "0" & strMinute
	ConvertToWMIDate = strYear & strMonth & strDay & strHour _
		& strMinute & "00.000000+***"
end function
