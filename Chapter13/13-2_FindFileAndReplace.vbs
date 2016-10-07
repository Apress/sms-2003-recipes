Const OverwriteExisting = TRUE
strComputer = "."
'Because of CIM_DataFile, split File Name and Extension into
' two variables
strFileName = "Foo"
strFileExt = "doc"
strNewFileName = "Foo2.doc"

'use this to obtain the current path to the VBScript
Set objFso= createobject("Scripting.FileSystemObject")
strScriptPath = objFso.GetParentFolderName(WScript.ScriptFullName)

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\cimv2")

'WQL query - notice we're only looking for files on C:
Set colFiles = objWMIService.ExecQuery _
	("Select * from CIM_Datafile where FileName = '" & _
	strFileName & "' and Extension = '" & strFileExt & _
	"' and drive = 'C:'")

For Each objFile in colFiles
	'rename file
    objFile.Rename(objFile.Name & ".bak")
	'copy the new file
    objFSO.CopyFile strScriptPath & "\" & strNewFileName,_
    	objFile.Drive & objFile.Path, OverwriteExisting   
Next
