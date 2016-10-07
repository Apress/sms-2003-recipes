strSMSServer = "SMSVPC"
strComputer = "2KPRO"

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
    ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
    If Loc.ProviderForLocalSite = True Then
        Set objSMS = objLoc.ConnectServer(Loc.Machine, "root\sms\site_" & _
            Loc.SiteCode)
    end if
Next
    
strWQL = "select arp.*, sys.Name from SMS_R_System sys " & _
    "inner join SMS_G_System_ADD_REMOVE_PROGRAMS arp on " & _
    "arp.ResourceID = sys.ResourceId where sys.Name = '" & _
    strComputer & "' order by arp.DisplayName"
    
Set colARPs = objSMS.ExecQuery(strWQL)

wscript.echo "Add/Remove Programs information for " & strComputer
for each objARP in colARPs
    wscript.echo objARP.arp.DisplayName & vbTAB & _
    objARP.arp.InstallDate & vbTAB & objARP.arp.Publisher
next


