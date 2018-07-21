Set aWscript = CreateObject("Wscript.Shell")
msgbox Getobject("winmgmts:root\cimv2:win32_Processor='cpu0'").AddressWidth