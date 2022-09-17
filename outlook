'====================================================================
'Description: 删除Outlook 2013/2016缓存联系人
'
'====================================================================

'Close Outlook socially
WScript.Echo "关闭Outlook 点OK."

'Set WMI Service
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

'Close Outlook forcefully if it is still running
Set colProcessList = objWMIService.ExecQuery _
 ("Select * from Win32_Process Where Name = 'outlook.exe'")
For Each objProcess in colProcessList
  objProcess.Terminate()
Next

'Get OS version
Set colOperatingSystems = objWMIService.ExecQuery _
 ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
  VersionOS = objOperatingSystem.Version
Next

'Get major OS Kernel number
sOSKernelMajor = Left(VersionOS,(InStr(VersionOS,".")-1))

'Set Shell
Set oShell = CreateObject("WScript.Shell")

'Determine Forms Cache path
If sOSKernelMajor > 5 Then
   sFormsCachePath = oShell.ExpandEnvironmentStrings("%UserProfile%") _
   & "\AppData\Local\Microsoft\Outlook\Offline Address Books"
   End If

'Verify whether the Forms Cache exists and delete it
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(sFormsCachePath) Then
  WScript.Echo "删除缓存联系人."
  Const DeleteReadOnly = True
  objFSO.DeleteFolder(sFormsCachePath), DeleteReadOnly
  WScript.Echo "删除缓存联系人成功. " _
   & VbNewLine & "请重新打开你的Outlook 确认！."
Else
 WScript.Echo "不能找到缓存联系人！. " _
  & VbNewLine & "请重新打开你的Outlook 确认！"
End If

If sOSKernelMajor > 5 Then
   sRoamCachePath = oShell.ExpandEnvironmentStrings("%UserProfile%") _
   & "\AppData\Local\Microsoft\Outlook\RoamCache"
End If

'Verify whether the Forms Cache exists and delete it
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(sRoamCachePath) Then
  WScript.Echo "删除自动完成."
  
  objFSO.DeleteFolder(sRoamCachePath), DeleteReadOnly
  WScript.Echo "删除自动完成成功. " _
   & VbNewLine & "请重新打开你的Outlook 确认！."
Else
 WScript.Echo "不能找到自动完成！. " _
  & VbNewLine & "请重新打开你的Outlook 确认！"
End If

Set ws=WScript.CreateObject("WScript.Shell")
ws.Run "Outlook.exe /CleanAutoCompleteCache"
