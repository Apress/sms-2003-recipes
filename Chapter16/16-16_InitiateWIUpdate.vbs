Const WIU = �{00000000-0000-0000-0000-000000000032}�
strComputer = "2kPro"
Set objCCM = GetObject("winmgmts://" & strComputer & "/root/ccm")
Set objClient = objCCM.Get("SMS_Client")
Set objSched = objClient.Methods_("TriggerSchedule"). _
	inParameters.SpawnInstance_()
objSched.sScheduleID = WIU
objCCM.ExecMethod "SMS_Client", "TriggerSchedule", objSched
