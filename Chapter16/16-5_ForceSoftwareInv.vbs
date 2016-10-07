Const SWINV = "{00000000-0000-0000-0000-000000000002}"
strComputer = "."

FullInv = Msgbox("Full Inventory?", vbYesNo)

If FullInv = vbYes then
	Set objSMS = GetObject("winmgmts://" & strComputer & _
		"/root/ccm/invagt")
	objSMS.Delete "InventoryActionStatus.InventoryActionID=" _
		& Chr(34) & SWINV & Chr(34)
End If

Set objCCM = GetObject("winmgmts://" & strComputer & "/root/ccm")
Set objClient = objCCM.Get("SMS_Client")
Set objSched = objClient.Methods_("TriggerSchedule"). _
	inParameters.SpawnInstance_()
objSched.sScheduleID = HWINV
objCCM.ExecMethod "SMS_Client", "TriggerSchedule", objSched
