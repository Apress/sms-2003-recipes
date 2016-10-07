strNewAccess = ucase("c:\program Files\microsoft " & _
    "office\office11\msaccess.EXE")

strComputer = "." 'connecting to local computer

Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colFiles = objWMIService. _
    ExecQuery("Select * from CIM_DataFile where Extension = 'lnk'")
    For Each objFile in colFiles
        CheckShortcut(objFile.Name)
    Next

Sub CheckShortcut(strName)
    wscript.echo strName
    'this actually looks inside the shortcut and modifies 
    '   "targetpath" as needed. This only affects the target, 
    '   not the arguments after the target.
    Set WshShell = wscript.CreateObject("WScript.Shell")
    'using create is a bit confusing.  If it already exists, 
    '   CreateShortcut edits instead of creates.
    Set oShellLink = WshShell.CreateShortcut(strName) 
    
    If InStr(1, UCase(oShellLink.TargetPath), _
        UCase("MSACCESS.EXE")) Then
        oShellLink.TargetPath = strNewAccess
        oShellLink.IconLocation = strNewAccess
    End If
     oShellLink.Save
end sub    
    