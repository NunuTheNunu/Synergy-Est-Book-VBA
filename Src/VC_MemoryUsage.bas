Attribute VB_Name = "VC_MemoryUsage"


Function GetMemUsage()
' Returns the current Excel.Application
' memory usage in MB

Set objSWbemServices = GetObject("winmgmts:")

GetMemUsage = objSWbemServices.Get("Win32_Process.Handle='" & GetCurrentProcessId & "'").WorkingSetSize / 1024 / 1024

Set objSWbemServices = Nothing
MsgBox GetMemUsage
End Function

