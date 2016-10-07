'you could use a for-loop in batch to bring in every .mof in a directory: for /F %G in ('dir /b') do mofcomp %G
Const ForWriting = 2
Set fso = CreateObject("Scripting.FileSystemObject")
strSMSServer = "SMSVPC"
strNEWSMSServer = "MYPRODSMSSERVER"
strNewSMSSiteCode = "PRD"
strExportFolder = "C:\Scripts\Packages\"

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

Set colPkgs = objSMS.ExecQuery _
("select * from SMS_Package")

for each objPkg in colPkgs
'wscript.echo objPkg.GetObjectText_
'wscript.echo objPkg.GetText_(2) 'output to xml
    wscript.echo "Exporting " & objPkg.Name & vbTAB & _
        objPkg.PackageID
    Set fout = fso.OpenTextFile(strExportFolder & _
        objPkg.Name & " (" & objPkg.PackageID & ")" & ".MOF", ForWriting, True)
    fout.writeline "//********************************"
    fout.writeline "//Created by VBScript" & vbTAB & Now()
    fout.writeline "//********************************"
    fout.writeline vbCRLF
    'only use the following line if planning to import MOF
    'from the command line
    fout.writeline "#pragma namespace(" & chr(34) & "\\\\" & _
        strNEWSMSServer & "\\root\\SMS\\site_" & _
        strnewSMSSiteCode & chr(34) & ")"
        
    'Write the package info
    fout.writeline vbCRLF
    fout.writeline "// **** Class : SMS_Package ****"
    
    for each strLine in split(objPkg.GetObjectText_, chr(10))
        if instr(strLine, "PackageID =") then
            fout.writeline(vbTAB & "PackageID = " & chr(34) & _
                chr(34)) & ";"
        elseif instr(strLine, "instance of SMS_Package") then
            strLine = "instance of SMS_Package as $pID"
            fout.writeline cstr(strLine)
        elseif instr(strLine, "SourceDate") then
            strLine = "SourceDate = " & Chr(34) & Chr(34)
            fout.writeline cstr(strLine)
        elseif instr(strLine, "SourceSite") then
            strLine = "SourceSite = " & Chr(34) & Chr(34)
            fout.writeline cstr(strLine)            
        elseif instr(strLine, "SourceVersion") then
            strLine = "SourceVersion = " & Chr(34) & Chr(34)
            fout.writeline cstr(strLine)            
        else
            fout.writeline cstr(strLine)
        end if
    next    
    
    'now write program info
    fout.writeline vbCRLF
    fout.writeline "// **** Class : SMS_Program ****"
    
    Set colPrograms = objSMS.ExecQuery _
    ("select * from SMS_Program where PackageID = '" & _
        objPkg.PackageID & "'")

    for each objProgram in colPrograms
        for each strLine in split(objProgram.GetObjectText_, chr(10))
            if instr(strLine, "PackageID =") then
                fout.writeline(vbTAB & "PackageID = $pID;")
            elseif instr(strLine, "DependentProgram") then
                if len(strLine) = 23 then
                    fout.writeline strLine
                else
                    fout.writeline vbTAB & "//" & strLine
                    fout.writeline vbTAB & "DependentProgram = " & Chr(34) & Chr(34) & ";"
                end if
            else
                fout.writeline cstr(strLine)
            end if
        next    
    next
    fout.writeline "// **** End ****"
    fout.close
next


