Const ForReading = 1
strSMSServer = <SMSServer>

strFileName = "dbcfg.ini"
strDataToCheck = "DBServername=MYPRODSERVER"
intCount = 0
Set objFSO = CreateObject("Scripting.FileSystemObject")

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

strSQL = "SELECT sys.Name, colFil.FileName, " & _
    "colFil.LocalFilePath, colFil.RevisionID " & _
    "FROM SMS_G_System_CollectedFile colFil INNER JOIN " & _
        "SMS_R_System sys ON " & _
         "colFil.ResourceID = sys.ResourceID " & _
    "WHERE colFil.FileName = '" & strFileName & "' " & _
    "ORDER BY colFil.CollectionDate"
         
Set colFiles = objSMS.ExecQuery(strSQL)

for each objFile in colFiles
    intCount = intCount + 1
    wscript.echo "Checking " & objFile.sys.Name & ". . ."
    Set objReadFile = objFSO.OpenTextFile _
        (objFile.colFil.LocalFilePath, ForReading)
    strContents = objReadFile.ReadAll
    if instr(ucase(strContents), ucase(strDataToCheck)) = 0 then
        strInfo = strInfo & objFile.sys.Name &  vbTAB & _
            objFile.colFil.RevisionID & vbCRLF
    end if
next

wscript.echo vbCRLF & intCount & " Files Checked" & vbCRLF
wscript.echo "The following computers do not have " & _
    strServerName & "in the file " & strFileName
wscript.echo strInfo
