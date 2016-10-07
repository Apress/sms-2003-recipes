Set oUIResource = CreateObject("UIResource.UIResourceMgr")
Set objCacheInfo = oUIResource.GetCacheInfo
wscript.echo objCacheInfo.TotalSize

