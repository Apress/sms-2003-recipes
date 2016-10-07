Const DDR = "{00000000-0000-0000-0000-000000000003}"
strComputer = "."

Set objCCM = GetObject("winmgmts://" & strComputer & "/root/ccm")
Set objClient = objCCM.Get("SMS_Client")
Set objSched = objClient.Methods_("TriggerSchedule"). _
	inParameters.SpawnInstance_()
objSched.sScheduleID = DDR
objCCM.ExecMethod "SMS_Client", "TriggerSchedule", objSched
