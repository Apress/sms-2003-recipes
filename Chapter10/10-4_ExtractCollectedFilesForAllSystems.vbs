Const OverwriteExisting = TRUE

strSMSServer = <SMSServer>

strFileName = "dbcfg.ini"
strTargetPath = "\\smsvpc\fileanalysis"

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
    "WHERE colFil.FileName = '" & strFileName & "'"
         
Set colFiles = objSMS.ExecQuery(strSQL)

for each objFile in colFiles
    wscript.echo objFile.sys.Name & vbTAB & _
        objfile.colFil.FileName & vbTAB & _
        objFile.colFil.LocalFilePath
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CopyFile objFile.colFil.LocalFilePath , _
        strTargetPath & "\" & objFile.sys.Name & "__" & _
        objfile.colFil.RevisionID & "__" & _
        objfile.colFil.FileName, OverwriteExisting
next
