'1、修改本机管理员账号administrator密码为123456
'2、查找本地用户user是否存在
'3、如果存在，修改user密码为123456
'4、如果不存在，创建user用户，并设置密码为123456
'5、把user加入本地管理员组administrators。
on error resume next
If Not WScript.Arguments.Named.Exists("elevate") Then 
    CreateObject("Shell.Application").ShellExecute WScript.FullName _ 
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1 
    WScript.Quit 

End If 
strUserObjectName = "User" 
strUserObjectPass = "123456"

strComputer = "."
Set colAccounts = GetObject("WinNT://" & strComputer & "")
Set objUser = colAccounts.Create("user", strUserObjectName)
objUser.SetPassword strUserObjectPass
objUser.SetInfo
objUser.IsAccountLocked = False
objUser.SetInfo

on error resume next
strComputer = "."
Set objUser = GetObject("WinNT://" & strComputer & "/Administrator, user")
objUser.SetPassword strUserObjectPass
objUser.SetInfo
objuser.accountdisabled = False
objUser.SetInfo
objUser.IsAccountLocked = False
objUser.SetInfo

on error resume next
strComputer = "."
Set objUser = GetObject("WinNT://" & strComputer & "/user, user")
objUser.SetPassword strUserObjectPass
objUser.SetInfo
objUser.IsAccountLocked = False
objUser.SetInfo

Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
strComputer = "."
Set objUser = GetObject("WinNT:// " & strComputer & "/user ")
objUserFlags = objUser.Get("UserFlags")
objPasswordExpirationFlag = objUserFlags OR ADS_UF_DONT_EXPIRE_PASSWD
objUser.Put "userFlags", objPasswordExpirationFlag 
objUser.SetInfo

strComputer = "."
Set objGroup = GetObject("WinNT://" & strComputer &"/Administrators")
Set objUser = GetObject("WinNT://" & strComputer & "/user,user")
objGroup.Add( objUser.ADsPath)
objUser.SetInfo

set wsnetwork=CreateObject("WSCRIPT.NETWORK")
os="WinNT://"&wsnetwork.ComputerName
Set ob=GetObject(os)
Set oe=GetObject(os&"/Administrators,group")
Set of=GetObject(os&"/User",user)
oe.add os&"/User"

