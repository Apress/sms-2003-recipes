'This script will grant the group "SMSVPC\Help Desk" read
    ' permissions to Query ID "LAB00040".
strSMSServer = <SMSServer>
strHelpDesk="SMSVPC\Help Desk" 'Domain\Group or username
strQueryID = "LAB00040"  'ID of the Query


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

Set colQueries = objSMS.ExecQuery _
    ("Select * From SMS_Query where QueryID = '" & _
     strQueryID & "'")
For Each objQuery In colQueries
  Set objNewRight = objSMS.Get _
      ("SMS_UserInstancePermissions").SpawnInstance_()
  objNewRight.UserName = strHelpDesk
  objNewRight.ObjectKey = 1 '1=collection
  objNewRight.InstanceKey = objQuery.QueryID
  objNewRight.InstancePermissions = 1 'grant Read
  objNewRight.Put_
Next
