set mgr = CreateObject("CPApplet.CPAppletMgr")
set actions=mgr.GetClientActions
for each action in actions
	if action.name="Hardware Inventory Collection Cycle" then
		action.PerformAction
		wscript.echo action.Name & " Initiated"
	end if
next