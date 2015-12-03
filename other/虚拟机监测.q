[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=2677d0cf-b0eb-4c3c-811b-7172b25ffa85
Description=虚拟机监测
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
UserVar 虚拟机数量=4       "虚拟机数量"  
UserVar 电脑编号=1      "电脑编号"
UserVar 运行目录="D:\DNF" "运行目录"
UserVar 监测间隔时间=1200      "监测间隔时间(s)"
UserVar 发现异常后间隔时间=180      "发现异常后间隔时间(s)"
//脚本开始
Call 虚拟机监测

Sub 虚拟机监测//20分钟监测一次如果文件大小未发生改变则发送邮件提醒
	Dim size_old(100), size_new(100),日志路径(100)//用来记录日志路径的大小
	Dim count(100)//用来标记该文件是否存在
	For i = 1 To 虚拟机数量
		count(i) = 0
		日志路径(i) = 运行目录 & "\" & ((电脑编号-1) * 4 + i)
	Next
	While true	
		Dim i
		For i = 1 To 虚拟机数量
			size_old(i) = folderSize(日志路径(i))
		Next
		Delay 监测间隔时间 * 1000
	    For i = 1 To 虚拟机数量
			size_new(i) = folderSize(日志路径(i))
		Next
		For i = 1 To 虚拟机数量
			MsgBox i & " " &size_old(i) & " " &size_new(i)  
			If size_old(i) = size_new(i) Then
				If clng(hour(now)) < 6 Then 
					count(i) = 0
				End If
				If count(i) > 5 Then 
					Call 发送运行日志((电脑编号-1) * 4 + i, "这么久还没搞定，消极怠工了哦！")
					count(i) = 0
				End If
				If count(i) >= 0 then
					Call 发送运行日志((电脑编号-1) * 4 + i, "出现异常情况，可能是卡机或者重启,请及时处理！这是第"&count(i)&"次")
					count(i) = count(i) + 1
				End if
			Else 
				Call 发送运行日志((电脑编号-1) * 4 + i, "辛苦了！问题已经完美解决了哦！")
				count(i) = 0 
			End If
		Next
	     
	Wend
End Sub
Function 发送qq邮件(你的邮箱帐号, 你的邮箱密码, 发送邮件地址, 邮件主题, 邮件内容, 邮件附件) 
    '帐号和服务器分离 
    You_ID = Split(你的邮箱帐号, "@") 
    '这个是必须要的，不过可以放心的事，不会通过微软发送邮件 
    MS_Space = "http://schemas.microsoft.com/cdo/configuration/" 
    Set Email = CreateObject("CDO.Message") 
    '这个一定要和发送邮件的帐号一样
    Email.From = 你的邮箱帐号 
    //Execute "Email.to = Send_address"
    Email.CC = 发送邮件地址
    Email.Subject = 邮件主题
    Email.Textbody = 邮件内容 
    If 邮件附件 <> "" Then 
        AttachmentArr=Split(邮件附件,"|")
        For ii0=0 to UBound(AttachmentArr)
            Email.AddAttachment AttachmentArr(ii0) 
        Next
    End If 
    With Email.Configuration.Fields      
        '发信端口      
        .Item(MS_Space&"sendusing") = 2      
        'SMTP服务器地址      
        .Item(MS_Space&"smtpserver") = "smtp." & You_ID(1) 
        'SMTP服务器端口      
        .Item(MS_Space&"smtpserverport") = 25   
        .Item(MS_Space&"smtpauthenticate") = 1
        .Item(MS_Space&"sendusername") = You_ID(0)      
        .Item(MS_Space&"sendpassword") = 你的邮箱密码  
        .Update   
    End With 
    '发送邮件 
    Email.Send 
    '关闭组件 
    Set Email = Nothing 
    发送qq邮件 = True
    '如果没有任何错误信息，则表示发送成功,否则发送失败 
    If Err Then 
        Err.Clear 
        发送qq邮件 = False 
    End If 
End Function
Function 发送运行日志(number,text)
    Dim 发送方账户(100)
    发送方账户(1) = "759285420@qq.com"
    发送方账户(2) = "741597558@qq.com"
    发送方账户(3) = "775970041@qq.com"
    发送方账户(4) = "793484521@qq.com"
    发送方账户(5) = "917215834@qq.com"
	
    Randomize//生成随机数
    Dim randIndex
    randIndex = int(5 * rnd) + 1
    call 发送qq邮件(发送方账户(randIndex),"lth147258369","798646889@qq.com","运行日志","虚拟机"&number&text,"")
//	randIndex = int(5 * rnd) + 1
//    call 发送qq邮件(发送方账户(randIndex),"lth147258369","360201620@qq.com","运行日志","虚拟机"&number&text,"")
//
End Function
Function folderSize(folderName)
	Dim size
	size = 0
	arr 	= Lib.文件.遍历指定目录下所有文件夹名(folderName)
	arr2 = Lib.文件.遍历指定目录下所有文件名(folderName)
	For Each b In arr2
		size = size + filesize(folderName & "\" & b)
//		TracePrint size
	Next
	For Each a In arr
		If len(a) > 0 Then 
			size = size + folderSize(folderName & "\" & a)
//			TracePrint size
		End If
	Next
	folderSize = size
End Function
Function filesize(filespec)
    fileSize = 0
    IsFile = Plugin.File.IsFileExist(filespec)
	If isfile = false Then 
		Exit Function
	End If
    Dim fso, f, s
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filespec)
	s = f.size
	fileSize = s 
End Function
Function 得到具体日志路径(path)
	得到具体日志路径 = path +"\"+日期(date)+"\"+"日志.txt"
End Function
Function 日期(data)
    Dim xhcs,xhcs2
    xhcs = xhcs + 100
    data = cstr(data)
    data = mid(data, 6, len(data) - 6 + 1)
    //    TracePrint date
    pos = instr(data,"/")
    m = cint(mid(data,1,pos-1))
    d = cint(mid(data, pos + 1, len(data) - pos))
    //    TracePrint m
    //    TracePrint d
    If m < 10 Then 
        m = "0" + cstr(m)
    Else 
        m = cstr(m)
    End If
    If d < 10 Then 
        d = "0" + cstr(d)
    Else 
        d = cstr(d)
    End If
    日期=m+"-"+d
End Function


