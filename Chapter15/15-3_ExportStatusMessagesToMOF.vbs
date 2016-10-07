Const ForWriting = 2
Set fso = CreateObject("Scripting.FileSystemObject")
strSMSServer = <SMSServer>
strNEWSMSServer = "SMSVPC"
strNewSMSSiteCode = "LAB"
strFileLoc = "C:\scripts\"

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

Set colStatQueries = objSMS.ExecQuery _
("select * from SMS_Query where TargetClassName = '" & _
"SMS_StatusMessage' and QueryID not like 'SMS%'")

for each objStatQuery in colStatQueries
    wscript.echo "Exporting " & objStatQuery.Name & vbTAB & _
        objStatQuery.QueryID
    Set fin = fso.OpenTextFile(strfileLoc & _
        objStatQuery.QueryID & ".MOF", ForWriting, True)
    fin.writeline "//********************************"
    fin.writeline "//Created by SMS Recipes Exporter"
    fin.writeline "//********************************"
    fin.writeline vbCRLF
    'only use the following line if planning to import MOF
    'from the command line
    fin.writeline "#pragma namespace(" & chr(34) & "\\\\" & _
        strNEWSMSServer & "\\root\\SMS\\site_" & _
        strnewSMSSiteCode & chr(34) & ")"
    fin.writeline vbCRLF
    fin.writeline "// **** Class : SMS_Query ****"
    fin.writeline "[SecurityVerbs(140551)]"
    for each strLine in split(objStatQuery.GetObjectText_, chr(10))
        if instr(strLine, "QueryID =") then
            fin.writeline(vbTAB & "QueryID = " & chr(34) & _
                chr(34)) & ";"
        else
            fin.writeline cstr(strLine)
        end if
    next    
    fin.writeline "// **** End ****"
    fin.close
next
