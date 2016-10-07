Const HKEY_LOCAL_MACHINE = &H80000002
' Only show adverts that have occurred within the last 90 days
Const intMaxDays = 90
strComputer = "."

'this is the base key path
strKeyPath = "SOFTWARE\Microsoft\SMS\Mobile Client\" & _
	"Software Distribution\Execution History\System"
'connect to the registry provider
Set oReg=GetObject _
	("winmgmts:{impersonationLevel=impersonate}!\\" & _
	strComputer & "\root\default:StdRegProv")
	
oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

'this first for-each loop is used to enumerate each package
'   key
For Each PackageID In arrSubKeys
	oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath & "\" & _
	PackageID, arrSubKeys2
	'the second for-each loop is used to enumerate each
	'  GUID key within a package key
	For Each GUID in arrSubKeys2
		SearchStrKeyPath = strKeyPath & "\" & PackageID & _
			"\" & GUID
		oReg.GetStringValue HKEY_LOCAL_MACHINE, _
			SearchStrKeyPath, "_ProgramID", strProgramID
		oReg.GetStringValue HKEY_LOCAL_MACHINE, _
			SearchStrKeyPath, "_RunStartTime", strRunStartTime
		oReg.GetStringValue HKEY_LOCAL_MACHINE, _
			SearchStrKeyPath, "_State", strState
		oReg.GetStringValue HKEY_LOCAL_MACHINE, _
			SearchStrKeyPath, "SuccessOrFailureCode", _
				strSuccessOrFailure
		oReg.GetStringValue HKEY_LOCAL_MACHINE, _
			SearchStrKeyPath, "SuccessOrFailureReason", _
				strSuccessOrFailureReason

		'Only display data for advert starts < 90 days					
		If not DateDiff _
			("d",strRunStartTime, now())	> intMaxDays Then				
		wscript.echo strProgramID & vbTAB & _
			strRunStartTime & vbTAB & strState & vbTAB & _
			strSuccessOrFailure & vbTAB & _
			strSuccessOrFailureReason
		End If
	Next
Next
	
