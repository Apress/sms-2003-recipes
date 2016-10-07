Document Type: IPF
item: Global
  Version=6.0
  Title English=Create Shortcut on Each Users Quick Launch
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
  Variable Name1=_SYS_
  Variable Default1=C:\WINDOWS\system32
  Variable Flags1=00001000
  Variable Name2=_SMSINSTL_
  Variable Default2=C:\Program Files\Microsoft SMS Installer
  Variable Flags2=00001000
end
item: Set Variable
  Variable=QLAUNCH
  Value=Application Data\Microsoft\Internet Explorer\Quick Launch
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
  Text=Parse string ALLUSERSDSKTP to extract Profile Paths
end
item: Parse String
  Source=%ALLUSERSDSKTP%
  Pattern=\All Users
  Variable1=PROFILEPATH
  Variable2=TMPJUNK
  Flags=00000010
end
item: Remark
end
item: Remark
  Text=Get the TEMP environement variable
end
item: Get Environment Variable
  Variable=SYSTEMP
  Environment=TEMP
end
item: Remark
end
item: Remark
  Text=Install the icon file locally
end
item: Install File
  Source=\\MySMSServer\SourceFiles\GoogleLaunch\GoogleLocal.ico
  Destination=%SYSWIN%\GoogleLocal.ico
  Flags=0000000000000010
end
item: Remark
end
item: Remark
  Text=Execute Command to create %temp%\users.txt, and wait...
end
item: Execute Program
  Pathname=cmd.exe
  Command Line=/c dir /b /ad "%PROFILEPATH%">%SYSTEMP%\users.txt
  Flags=00000010
end
item: Remark
end
item: Remark
  Text=Read each line of users.txt, and use the username to comlete the path
end
item: Remark
  Text=   to the user's quick Launch
end
item: Read/Update Text File
  Variable=USERNAME
  Pathname=%SYSTEMP%\users.txt
end
item: If/While Statement
  Variable=USERNAME
  Value=Administrator
  Flags=00000101
end
item: If/While Statement
  Variable=USERNAME
  Value=All Users
  Flags=00000101
end
item: Create Shortcut
  Source English=c:\Program Files\Internet Explorer\iexplore.exe
  Destination English=%PROFILEPATH%\%USERNAME%\%QLaunch%\GoogleLocal.lnk
  Command Options English=http://www.google.com/local
  Icon Pathname English=%SYSWIN%\googlelocal.ico
  Description English=Launch Google Local!
  Key Type English=1536
  Flags=00000001
end
item: End Block
end
item: End Block
end
item: End Block
end
item: Remark
end
item: Delete File
  Pathname=%SYSTEMP%\users.txt
end
item: Remark
end
