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
  Text=File to find
end
item: Set Variable
  Variable=FINDFILE
  Value=Foo.doc
end
item: Set Variable
  Variable=FOUNDFILE
end
item: Remark
end
item: Remark
  Text=While Loop (perform loop at least once)
end
item: If/While Statement
  Variable=FOUNDFILE
  Flags=00110001
end
item: Remark
end
item: Search for File
  Variable=FOUNDFILE
  Pathname List=%FINDFILE%
  Flags=00000001
end
item: Remark
end
item: If/While Statement
  Variable=FOUNDFILE
  Flags=00000001
end
item: Rename File/Directory
  Old Pathname=%FOUNDFILE%\%FINDFILE%
  New Filename=%FINDFILE%.Bak
end
item: Install File
  Source=\\smsserver\soruces\FooDoc\Foo2.doc
  Destination=%FOUNDFILE%\Foo2.doc
  Flags=0000000000000010
end
item: End Block
end
item: Remark
end
item: End Block
end
item: Remark
end
item: Remark
end
