RUN_FROM_LOCAL_DISPPOINT = 2^(3)
DOWNLOAD_FROM_LOCAL_DISPPOINT = 2^(4)
DONT_RUN_NO_LOCAL_DISPPOINT = 2^(5)
DOWNLOAD_FROM_REMOTE_DISPPOINT = 2^(6)
RUN_FROM_REMOTE_DISPPOINT = 2^(7)

RunFromLocalDP_DownloadIfRemoteDP = _ 
    RUN_FROM_LOCAL_DISPPOINT + DOWNLOAD_FROM_REMOTE_DISPPOINT
DownloadFromLocal_DontRunNoLocal = _
    DOWNLOAD_FROM_LOCAL_DISPPOINT + DONT_RUN_NO_LOCAL_DISPPOINT
DownloadFromLocalDP_DownloadIfRemoteDP = _
    DOWNLOAD_FROM_LOCAL_DISPPOINT + _
        DOWNLOAD_FROM_REMOTE_DISPPOINT

        
strSMSServer = <SMSServer>

Set objLoc =  CreateObject("WbemScripting.SWbemLocator")
Set objSMS= objLoc.ConnectServer(strSMSServer, "root\sms")
Set Results = objSMS.ExecQuery _
    ("SELECT * From SMS_ProviderLocation WHERE ProviderForLocalSite = true")
For each Loc in Results
    If Loc.ProviderForLocalSite = True Then
        Set objSMS = objLoc.ConnectServer(Loc.Machine, "root\sms\site_" & _
            Loc.SiteCode)
            strSMSSiteCode = Loc.Sitecode
    end if
Next

Set colAdverts = objSMS.ExecQuery _
    ("Select * From SMS_Advertisement where " & _
    "RemoteClientFlags = " & _
    RunFromLocalDP_DownloadIfRemoteDP & _
    " or RemoteClientFlags = " & _
    DownloadFromLocal_DontRunNoLocal & _
    " or RemoteClientFlags = " & _
    DownloadFromLocalDP_DownloadIfRemoteDP & " order by " & _
    " AdvertisementName")
    
For Each objAdvert In colAdverts
    Select Case objAdvert.RemoteclientFlags
    
        Case RunFromLocalDP_DownloadIfRemoteDP
            wscript.echo objAdvert.AdvertisementName & vbTAB & _
                "(RunFromLocalDP_DownloadIfRemoteDP)"
                
        Case DownloadFromLocal_DontRunNoLocal
            wscript.echo objAdvert.AdvertisementName & vbTAB & _
                "(DownloadFromLocal_DontRunNoLocal)"                    
        
        Case DownloadFromLocalDP_DownloadIfRemoteDP
            wscript.echo objAdvert.AdvertisementName & vbTAB & _
                "(DownloadFromLocalDP_DownloadIfRemoteDP)"  
                
        Case Else
            'neither remote or local are configured to
            ' download from DP
    End Select  
Next

