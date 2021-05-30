' Как работать с процессом
Dim WindowStyle : WindowStyle = 0 ' 0 - фоновый режим, 1 - обычный режим, 2 - свернутый вид, 3 - развернутый вид

Dim KillWaitOnReturn : KillWaitOnReturn = true ' Не продолжать выполнение скрипта, пока процесс не будет убит

Dim oShell : Set oShell = CreateObject("WScript.Shell")

KillExtraProcesses("part_name_of_process") ' Убиваем все процессы, в названии которых присутствует передаваемое значение

Function KillExtraProcesses(ProcessName)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

    Set colItems = objWMIService.ExecQuery("Select * from Win32_Process")
    For Each objItem in colItems
        If InStr(objItem.Name, ProcessName) <> 0 Then
            Command = "taskkill /f /im " & objItem.Name
            oShell.Run Command, WindowStyle, KillWaitOnReturn
        End If
    Next
End Function
