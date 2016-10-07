Function ConvertToWMIDate(strDate)
    strYear = year(strDate)
    strMonth = month(strDate)
    strDay = day(strDate)
    strHour = hour(strDate)
    strMinute = minute(strDate)
    'pad date appropriately
    if len(strmonth) = 1 then strMonth = "0" & strMonth
    if len(strDay) = 1 then strDay = "0" & strDay
    if len(strHour) = 1 then strHour = "0" & strHour
    if len(strMinute) = 1 then strMinute = "0" & strMinute
    ConvertToWMIDate = strYear & strMonth & strDay & strHour & _
        strMinute & "00.000000+***"
end function
