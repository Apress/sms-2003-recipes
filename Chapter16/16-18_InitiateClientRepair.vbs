strComputer = "2kPro"
Set smsClient = GetObject("winmgmts://" & strComputer & _
 "/root/ccm:SMS_Client")
Set result = smsClient.ExecMethod_("RepairClient")
wscript.echo "Repair Initiated on " & strComputer
