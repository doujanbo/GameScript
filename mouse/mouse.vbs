Set wshshell = CreateObject("wscript.shell")
wshshell.run "regsvr32  /n /i:user "".\wyhkm.dll"""

Set wyhkm=CreateObject("wyp.hkm")
ver = wyhkm.GetVersion()
MsgBox "ģ��汾��" & Hex(ver), 4096
