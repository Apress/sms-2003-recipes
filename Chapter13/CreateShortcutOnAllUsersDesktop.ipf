Document Type: IPF
item: Global
  Version=6.0
  Title English=Icon To All Users Desktop
  Flags=01000100
  Languages=0 0 65 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
  LanguagesList=English
  Default Language=2
  Copy Default=1
  Japanese Font Name=MS Gothic
  Japanese Font Size=9
  Start Gradient=0 0 255
  End Gradient=0 0 0
  Windows Flags=00010100000000000010110000011010
  Message Font=MS Sans Serif
  Font Size=8
  Disk Filename=SETUP
  Patch Flags=0000000000000001
  Patch Threshold=85
  Patch Memory=4000
  FTP Cluster Size=20
  Variable Name1=_SYS_
  Variable Default1=C:\WINDOWS\system32
  Variable Flags1=00001000
  Variable Name2=_SMSINSTL_
  Variable Default2=C:\Program Files\Microsoft SMS Installer
  Variable Flags2=00001000
end
item: Remark
  Text=Define variables for ease-of-use
end
item: Set Variable
  Variable=IEXPLORE
  Value=c:\Program Files\Internet Explorer\iexplore.exe
end
item: Set Variable
  Variable=URL
  Value=www.microsoft.com/technet
end
item: Set Variable
  Variable=URLCOMMENTS
  Value=Launch Microsoft Technet
end
item: Remark
end
item: Remark
  Text=Obtain the path to the All Users Desktop
end
item: Get Registry Key Value
  Variable=ALLUSERSDSKTP
  Key=Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders
  Value Name=Common Desktop
  Flags=00000100
end
item: Remark
end
item: Remark
  Text=Create the Icon (OverWrite if it already exists)
end
item: Create Shortcut
  Source English=%IEXPLORE%
  Destination English=%ALLUSERSDSKTP%\Launch TechNet.lnk
  Command Options English=%URL%
  Description English=%URLCOMMENTS%
  Key Type English=1536
  Flags=00000001
end
