Const ForWriting = 2
Set fso = CreateObject("Scripting.FileSystemObject")

strSMSServer = <SMSServer>

strNEWSMSServer = "MOFSMSSERVER"
strNewSMSSiteCode = "MOFSMSSITECODE"

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
("select * from SMS_Query where TargetClassName <> '" & _
"SMS_StatusMessage' and QueryID not like 'SMS%'")

for each objQuery in colQueries
    wscript.echo "Exporting " & objQuery.Name & vbTAB & _
        objQuery.QueryID
    Set fout = fso.OpenTextFile("C:\Scripts\sms\Queries\mofs\" & _
        objQuery.QueryID & ".MOF", ForWriting, True)
    fout.writeline "//********************************"
    fout.writeline "//Created by SMS Recipes Exporter"
    fout.writeline "//********************************"
    fout.writeline vbCRLF
    'only use the following line if planning to import MOF
    'from the command line
    fout.writeline "#pragma namespace(" & chr(34) & "\\\\" & _
        strNEWSMSServer & "\\root\\SMS\\site_" & _
        strnewSMSSiteCode & chr(34) & ")"
    fout.writeline vbCRLF
    fout.writeline "// **** Class : SMS_Query ****"
    for each strLine in split(objQuery.GetObjectText_, chr(10))
        if instr(strLine, "QueryID =") then
            fout.writeline(vbTAB & "QueryID = " & chr(34) & _
                chr(34)) & ";"
        else
            fout.writeline cstr(strLine)
        end if
    next    
    fout.writeline "// **** End ****"
    fout.close
next
