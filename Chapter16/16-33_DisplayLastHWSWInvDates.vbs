strComputerName = "."

'first, get the time zone bias
Set objWMIService = GetObject("winmgmts:\\" & strComputerName & _
	"\root\cimv2")
Set colTimeZone = objWMIService.ExecQuery _
	("Select * from Win32_TimeZone")
For Each objTimeZone in colTimeZone
    intBias = cint(objTimeZone.Bias )
Next

Set objSMS = GetObject("winmgmts://" & strComputerName & _
	 "/root/ccm/invagt")
Set colInvInfo = objSMS.ExecQuery _
	("Select * from InventoryActionStatus")

for each objInvInfo in colInvInfo

	select case objInvInfo.InventoryActionID
		case "{00000000-0000-0000-0000-000000000001}"
			strInv = "Hardware Inventory"
		case "{00000000-0000-0000-0000-000000000002}"
			strInv = "Software Inventory"
		case "{00000000-0000-0000-0000-000000000010}"
			strInv = "File Collection"
		case "{00000000-0000-0000-0000-000000000003}"
			strInv = "Discovery Data Record"
	end select
	wscript.echo strInv & vbTAB & _
		convDate(objInvInfo.LastCycleStartedDate, intBias) & _
		vbTAB & convDate(objInvInfo.LastReportDate, intBias)
next		

Function convDate(dtmInstallDate, intBias)
    convDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
    Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
    & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
    Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, 13, 2))
	convDate = DateAdd("N",intBias,convDate)           
End Function
