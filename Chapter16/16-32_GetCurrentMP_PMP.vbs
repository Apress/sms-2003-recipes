strComputerName = "2kPro"
Set objSMS = GetObject("winmgmts://" & strComputerName & _
	 "/root/ccm")
Set objAuthority = objSMS.ExecQuery _
	("Select * from SMS_Authority")

For Each authority In objAuthority
    wscript.echo "SMS Site: " & _
    	Replace(authority.Name, "SMS:", "")
    wscript.echo "MP: " & authority.CurrentManagementPoint
Next

Set colProxyMPs = objSMS.ExecQuery _
	("Select * from SMS_MPProxyInformation")
For Each objProxyMP in colProxyMPs
	wscript.echo "Proxy MP: " & objProxyMP.Name
Next

