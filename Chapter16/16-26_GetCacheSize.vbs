'Requires a restart of the SMS Agent Host Service
strComputer = "2kpro"
Set objSMS = _
	GetObject("winmgmts:{impersonationLevel=impersonate}!//" & _
	strComputer & "/root/ccm/SoftMgmtAgent")
Set objCacheConfig = objSMS.ExecQuery _
	("Select * from CacheConfig")
for each objCache in objCacheConfig
	wscript.echo objCache.Size
next
