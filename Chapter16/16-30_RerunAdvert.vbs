strComputerName = "2KPro"
strAdvID = "LAB20001"
set objSMS = GetObject("winmgmts://" & _
	strComputerName & _
	"/root/ccm/policy/machine/actualconfig")
Set objScheds = objSMS.ExecQuery _
	("select * from CCM_Scheduler_ScheduledMessage")
For each objSched in objScheds
	'locate ScheduleMessageID that contains strAdvID
	If Instr(objSched.ScheduledMessageID, strAdvID) > 0 then
		strMsgID = objSched.ScheduledMessageID
		exit for
	End If
Next
'strMsgID now contains proper Advertisement

Set objSWDs = objSMS.ExecQuery _
("select * from CCM_SoftwareDistribution where " & _ 
	"ADV_AdvertisementID = '" & strAdvID & "'" )
for each objSWD in objSWDs
	strOrigBehavior = objSWD.ADV_RepeatRunBehavior
	'strOrigBehavior now contains original Repeat Behavior
	'Now temporarily set RepeatRunBehavior to RerunAlways
	objSWD.ADV_RepeatRunBehavior = "RerunAlways"
	objSWD.Put_ 0
Next

set objCCM = GetObject("winmgmts://" & strComputerName & _
	"/root/ccm")
Set objSMSClient = objCCM.Get("SMS_Client")
objSMSClient.TriggerSchedule strMsgID

'sleep for 5 seconds, for advert to start
wscript.sleep 5000

Set objScheds = objSMS.ExecQuery _
("select * from CCM_SoftwareDistribution where " & _
    "ADV_AdvertisementID = '" & strAdvID & "'" )
for each objSched in objScheds
	'Set RepeatRunBehavior back to original config
	objSched.ADV_RepeatRunBehavior = strOrigBehavior
	objSched.Put_ 0
Next