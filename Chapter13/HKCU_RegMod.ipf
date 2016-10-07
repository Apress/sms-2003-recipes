Document Type: IPF
item: Global
  Version=6.0
  Flags=01000100
  Languages=0 0 65 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
  LanguagesList=English
  Default Language=2
  Copy Default=1
  Japanese Font Name=MS Gothic
  Japanese Font Size=9
  Start Gradient=0 0 255
  End Gradient=0 0 0
  Windows Flags=00010100000000000010110001011010
  Message Font=MS Sans Serif
  Font Size=8
  Disk Filename=SETUP
  Patch Flags=0000000000000001
  Patch Threshold=85
  Patch Memory=4000
  FTP Cluster Size=20
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
  Text=Install the File to be Run when the user logs in
end
item: Install File
  Source=c:\temp\CLAIM_HKCUMOD.EXE
  Destination=%SYSWIN%\CLAIM_HKCUMOD.exe
  Flags=0000000000000010
end
item: Remark
end
item: Remark
  Text=Add/Update Registry keys in ActiveSetup
end
item: Edit Registry
  Total Keys=1
  Key=SOFTWARE\Microsoft\Active Setup\Installed Components\CLAIM_HKCUMOD
  New Value=%SYSWIN%\CLAIM_HKCUMOD.EXE /s
  Value Name=StubPath
  Root=2
end
item: Edit Registry
  Total Keys=1
  Key=SOFTWARE\Microsoft\Active Setup\Installed Components\CLAIM_HKCUMOD
  New Value=1
  Value Name=Version
  Root=2
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
item: Remark
end
