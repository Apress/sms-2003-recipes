Set  oCPAppletMgr=CreateObject("CPApplet.CPAppletMgr")
Set oClientComponents=oCPAppletMgr.GetClientComponents
For Each oClientComponent In oClientComponents
	strInfo = oClientComponent.DisplayName
     Select Case oClientComponent.State
     Case 0
     	strInfo = strInfo & vbTAB & "(Installed)"
     Case 1 
          strInfo = strInfo & vbTAB & "(Enabled)"
     Case 2
          strInfo = strInfo & vbTAB & "(Disabled)"
     End Select
     	strinfo = strInfo & vbTAB & oClientcomponent.Version
     wscript.echo strInfo
Next
