IMMEDIATE = 2^(5)
ONUSERLOGON = 2^(9)
ONUSERLOGOFF = 2^(10)

strSMSServer = <SMSServer>

strAdvertID = "LAB20015"
 
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

Set objAdvert=objSMS.Get _
    ("SMS_Advertisement.AdvertisementID='" & strAdvertID & "'")

wscript.echo "Mandatory Assignment(s) for " & chr(34) & _
    objAdvert.AdvertisementName & " (" & strAdvertID & ")" & _
        chr(34)
            
intAdvertFlags = objAdvert.AdvertFlags
if (intAdvertFlags and IMMEDIATE) then
    wscript.echo "Event Assignment to run IMMEDIATELY."
end if
if (intAdvertFlags and ONUSERLOGON) then
    wscript.echo "Event Assignment to run On User Logon."
end if
if (intAdvertFlags and ONUSERLOGOFF) then
    wscript.echo "Event Assignment to run On User Logoff."
end if

for each objSched in objAdvert.AssignedSchedule
    'for a 'quick and dirty', instead of the case statement
    'below, you could just display the information in the object
    'by using objSched.GetObjectText_
    'wscript.echo objSched.GetObjectText_
    wscript.echo objSched.Path_.Class
    select case objSched.Path_.Class
        case "SMS_ST_NonRecurring"
            strInfo = vbTAB & "Non-Recurring Assignment: "
            strInfo = strInfo & "Occurs at " & _
                WMIDateStringToDate(objSched.StartTime)
            if objSched.IsGMT then
                strInfo = strInfo & " GMT"
            end if
            
        case "SMS_ST_RecurInterval"
            strInfo = vbTAB & "Recurring Interval Assignment: "
            strInfo = strInfo & "Every " & objSched.DaySpan & _
                " days, " & objSched.MinuteSpan & " minutes, "
            strInfo = strInfo & "beginning on " & _
                WMIDateStringToDate(objSched.StartTime)
            if objSched.IsGMT then
                strInfo = strInfo & " GMT"
            end if      
        
        case "SMS_ST_RecurMonthlyByDate"
            strInfo = vbTAB & "Recurring Monthly By Date: "
            strInfo = strInfo & "Occurs on the " & _
                objSched.MonthDay & " day, every " & _
                objSched.ForNumberOfMonths & " months, "
            strInfo = strInfo & "beginning on " & _
                WMIDateStringToDate(objSched.StartTime)
            if objSched.IsGMT then
                strInfo = strInfo & " GMT"
            end if          
        
        case "SMS_ST_RecurMonthlyByWeekday"
            strInfo = vbTAB & "Recurring Monthly By Weekday: "
            strInfo = strInfo & "Occurs on the " & _
                objSched.Day & " day, every " & _
                objSched.ForNumberOfMonths & " months, " & _
                "for week order " & objSched.WeekOrder & ","
            strInfo = strInfo & "beginning on " & _
                WMIDateStringToDate(objSched.StartTime)
            if objSched.IsGMT then
                strInfo = strInfo & " GMT"
            end if              
        
        case "SMS_ST_RecurWeekly"
            strInfo = vbTAB & "Recurring Monthly By Weekday: "
            strInfo = strInfo & "Occurs on the " & _
                objSched.Day & " day, every " & _
                objSched.ForNumberOfWeeks & " weeks, " 
            strInfo = strInfo & "beginning on " & _
                WMIDateStringToDate(objSched.StartTime)
            if objSched.IsGMT then
                strInfo = strInfo & " GMT"
            end if              
                
    end select
    wscript.echo strInfo
next

Function WMIDateStringToDate(dtmInstallDate)
WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
    Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
        & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
            Mid(dtmInstallDate, 11, 2) & ":" & _
            Mid(dtmInstallDate, 13, 2))
End Function
