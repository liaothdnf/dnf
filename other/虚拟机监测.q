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
Description=��������
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
UserVar ���������=4       "���������"  
UserVar ���Ա��=1      "���Ա��"
UserVar ����Ŀ¼="D:\DNF" "����Ŀ¼"
UserVar �����ʱ��=1200      "�����ʱ��(s)"
UserVar �����쳣����ʱ��=180      "�����쳣����ʱ��(s)"
//�ű���ʼ
Call ��������

Sub ��������//20���Ӽ��һ������ļ���Сδ�����ı������ʼ�����
	Dim size_old(100), size_new(100),��־·��(100)//������¼��־·���Ĵ�С
	Dim count(100)//������Ǹ��ļ��Ƿ����
	For i = 1 To ���������
		count(i) = 0
		��־·��(i) = ����Ŀ¼ & "\" & ((���Ա��-1) * 4 + i)
	Next
	While true	
		Dim i
		For i = 1 To ���������
			size_old(i) = folderSize(��־·��(i))
		Next
		Delay �����ʱ�� * 1000
	    For i = 1 To ���������
			size_new(i) = folderSize(��־·��(i))
		Next
		For i = 1 To ���������
			MsgBox i & " " &size_old(i) & " " &size_new(i)  
			If size_old(i) = size_new(i) Then
				If clng(hour(now)) < 6 Then 
					count(i) = 0
				End If
				If count(i) > 5 Then 
					Call ����������־((���Ա��-1) * 4 + i, "��ô�û�û�㶨������������Ŷ��")
					count(i) = 0
				End If
				If count(i) >= 0 then
					Call ����������־((���Ա��-1) * 4 + i, "�����쳣����������ǿ�����������,�뼰ʱ�������ǵ�"&count(i)&"��")
					count(i) = count(i) + 1
				End if
			Else 
				Call ����������־((���Ա��-1) * 4 + i, "�����ˣ������Ѿ����������Ŷ��")
				count(i) = 0 
			End If
		Next
	     
	Wend
End Sub
Function ����qq�ʼ�(��������ʺ�, �����������, �����ʼ���ַ, �ʼ�����, �ʼ�����, �ʼ�����) 
    '�ʺźͷ��������� 
    You_ID = Split(��������ʺ�, "@") 
    '����Ǳ���Ҫ�ģ��������Է��ĵ��£�����ͨ��΢�����ʼ� 
    MS_Space = "http://schemas.microsoft.com/cdo/configuration/" 
    Set Email = CreateObject("CDO.Message") 
    '���һ��Ҫ�ͷ����ʼ����ʺ�һ��
    Email.From = ��������ʺ� 
    //Execute "Email.to = Send_address"
    Email.CC = �����ʼ���ַ
    Email.Subject = �ʼ�����
    Email.Textbody = �ʼ����� 
    If �ʼ����� <> "" Then 
        AttachmentArr=Split(�ʼ�����,"|")
        For ii0=0 to UBound(AttachmentArr)
            Email.AddAttachment AttachmentArr(ii0) 
        Next
    End If 
    With Email.Configuration.Fields      
        '���Ŷ˿�      
        .Item(MS_Space&"sendusing") = 2      
        'SMTP��������ַ      
        .Item(MS_Space&"smtpserver") = "smtp." & You_ID(1) 
        'SMTP�������˿�      
        .Item(MS_Space&"smtpserverport") = 25   
        .Item(MS_Space&"smtpauthenticate") = 1
        .Item(MS_Space&"sendusername") = You_ID(0)      
        .Item(MS_Space&"sendpassword") = �����������  
        .Update   
    End With 
    '�����ʼ� 
    Email.Send 
    '�ر���� 
    Set Email = Nothing 
    ����qq�ʼ� = True
    '���û���κδ�����Ϣ�����ʾ���ͳɹ�,������ʧ�� 
    If Err Then 
        Err.Clear 
        ����qq�ʼ� = False 
    End If 
End Function
Function ����������־(number,text)
    Dim ���ͷ��˻�(100)
    ���ͷ��˻�(1) = "759285420@qq.com"
    ���ͷ��˻�(2) = "741597558@qq.com"
    ���ͷ��˻�(3) = "775970041@qq.com"
    ���ͷ��˻�(4) = "793484521@qq.com"
    ���ͷ��˻�(5) = "917215834@qq.com"
	
    Randomize//���������
    Dim randIndex
    randIndex = int(5 * rnd) + 1
    call ����qq�ʼ�(���ͷ��˻�(randIndex),"lth147258369","798646889@qq.com","������־","�����"&number&text,"")
//	randIndex = int(5 * rnd) + 1
//    call ����qq�ʼ�(���ͷ��˻�(randIndex),"lth147258369","360201620@qq.com","������־","�����"&number&text,"")
//
End Function
Function folderSize(folderName)
	Dim size
	size = 0
	arr 	= Lib.�ļ�.����ָ��Ŀ¼�������ļ�����(folderName)
	arr2 = Lib.�ļ�.����ָ��Ŀ¼�������ļ���(folderName)
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
Function �õ�������־·��(path)
	�õ�������־·�� = path +"\"+����(date)+"\"+"��־.txt"
End Function
Function ����(data)
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
    ����=m+"-"+d
End Function


