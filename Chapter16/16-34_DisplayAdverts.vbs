strComputer = "."
Set objSMS = GetObject("winmgmts:\\" & strComputer & _
	"\root\ccm\Policy\Machine\ActualConfig")
Set colSW = objSMS.ExecQuery _
	("Select * from CCM_SoftwareDistribution")
wscript.echo colSW.count & " Advertisements"
for each oSW in colSW
	wscript.echo oSW.PRG_HistoryLocation & vbTAB & _
		oSW.ADV_AdvertisementID & vbTAB & oSW.PKG_PackageID & _
		vbTAB & oSW.PRG_ProgramID & vbTAB & _
		oSW.ADV_ActiveTime & vbTAB & oSW.ADV_ExpirationTime & _
		vbTAB & oSW.ADV_MandatoryAssignments & vbTAB & _
		oSW.PKG_Name & vbTAB & oSW.PRG_ProgramName & vbTAB & _
		oSW.PRG_PRF_AfterRunning & vbTAB & _
		oSW.PRG_CustomLogoffReturnCodes & vbTAB & _
		oSW.PRG_MaxDuration & vbTAB & _
		oSW.PRG_PRF_UserLogonRequirement
next