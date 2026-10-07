' DMS 后台启动脚本
' 用于计划任务，无窗口启动 DMS 服务

Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\AI\QC\DMS"
WshShell.Run "cmd /c start_production.bat", 0, False
