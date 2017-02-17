# $language = "VBScript"
# $interface = "1.0"

dim objCurrentTab
dim strPrompt
Set re = New RegExp
re.Global = True  
re.Pattern = "(\d/\s\d+/\d+)\s+(\d+)\s+"

dim g_fso
set g_fso = CreateObject("Scripting.FileSystemObject")
Const ForWriting = 2 
Const ForAppending = 8

sub main()
	'��ȡ��ǰ�����������ʾ��
	set objCurrentTab = crt.GetScriptTab
	strPrompt = getstrPrompt(objCurrentTab)
	sn_number = crt.Dialog.Prompt("������Ҫ��ѯ��ONU��16λSN����", "������", "4857544374692E75", false)
	if len(sn_number) = 16 then
		delete_huawei_ONU(sn_number)
		msgbox "ɾ���ɹ���"
	else
		msgbox "SN����λ��������"
	end if
end sub

'sn-auth 4857544374692E75
'��#��ģʽ�� ɾ��onu �Լ���Ӧ��service,
Sub delete_huawei_ONU(sn_number)
	'----------���������� ont pon �� �������Ϣ----------
	'����ont pon �� �������Ϣ
	objCurrentTab.Screen.Synchronous = True
	objCurrentTab.Screen.Send "display ont info 0 all" &chr(13) 
	objCurrentTab.Screen.waitForString vbcr
			
	strCompleteOutput = ""
	Do
		strResult = objCurrentTab.Screen.ReadString("---- More ( Press 'Q' to break ) ----", strPrompt)
		strCompleteOutput = strCompleteOutput & strResult
		If objCurrentTab.Screen.MatchIndex = 1 Then crt.Screen.Send " "
		If objCurrentTab.Screen.MatchIndex = 2 Then Exit Do
	Loop
			
	objCurrentTab.Screen.Synchronous = false

	re.Pattern = "(\d/\s\d+/\d+)\s+(\d+)\s+" & sn_number
	'��������ʽ���gpon_num��ont_num
	If re.Test(strCompleteOutput) <> True Then
		MsgBox "�쳣����"
		crt.quit
	Else
		Set matches = re.Execute(strCompleteOutput)
		For Each match In matches
			gpon_num = match.SubMatches(0)
			ont_num = match.SubMatches(1)
			exit for
		Next
	End If
	'--------------�������ִ�в��Ҳ�ɾ��service-port ����------------
	'��ѯ service-port ��Ϣ
	'0/ 3/2
	str1 = Split(gpon_num, "/")
	objCurrentTab.Screen.Synchronous = True
	objCurrentTab.Screen.Send "display service-port board " & str1(0) &  "/" & str1(1) &" sort-by port" &chr(13) 
	objCurrentTab.Screen.waitForString vbcr
			
	strCompleteOutput = ""
	Do
		strResult = objCurrentTab.Screen.ReadString("---- More ( Press 'Q' to break ) ----", strPrompt)
		strCompleteOutput = strCompleteOutput & strResult
		If objCurrentTab.Screen.MatchIndex = 1 Then crt.Screen.Send " "
		If objCurrentTab.Screen.MatchIndex = 2 Then Exit Do
	Loop
	objCurrentTab.Screen.Send "config" & vbcr
	objCurrentTab.Screen.WaitForString "#"		
	'  1066 1100 stacking gpon 0/5 /8  3     2     vlan  1000       -    -    up
	re.Pattern = "\s+(\d+).*" & trim(str1(0)) &  "/" & trim(str1(1)) & "\s*/" & trim(str1(2)) & "\s+" & ont_num
	'��������ʽ���gpon_num��ont_num
	If re.Test(strCompleteOutput) <> True Then
		'MsgBox "�쳣����"
		'crt.quit
	Else
		Set matches = re.Execute(strCompleteOutput)
		For Each match In matches
			service_num = match.SubMatches(0)
			'ִ��ɾ�� service-port ����
			objCurrentTab.Screen.Send "undo service-port " & service_num & chr(13)
			objCurrentTab.Screen.waitForString "#"
		Next
	End If
	'----------�������ִ��ONU ETH�ڵ�native-vlan�޸ĳ�Ĭ�ϵ�VLAN 1����---------
	'��configģʽ�½���  gpon ,
	objCurrentTab.Screen.Send "interface gpon " & str1(0) & "/" & str1(1) & vbcr
	objCurrentTab.Screen.WaitForString "#"
	'ont port native-vlan 7 2 eth 1 vlan 1
	objCurrentTab.Screen.Send "ont port native-vlan " & str1(2) & " " & ont_num & " eth 1 vlan 1" & vbcr
	objCurrentTab.Screen.WaitForString "{ <cr>|priority<K> }:"
	objCurrentTab.Screen.Send vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "ont port native-vlan " & str1(2) & " " & ont_num & " eth 2 vlan 1" & vbcr
	objCurrentTab.Screen.WaitForString "{ <cr>|priority<K> }:"
	objCurrentTab.Screen.Send vbcr
	objCurrentTab.Screen.WaitForString "#"
	'----------ִ��ɾ��ont����---------
	'ont delete 7 2
	objCurrentTab.Screen.Send "ont delete " & str1(2) & " " & ont_num & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "quit" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "quit" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Synchronous = false
	'д��־�Ǳ���İ�������Ե�רҵ!
	Set logStream = g_fso.OpenTextFile(".\huawei_onu_config_log.txt", 8, True)
	logStream.WriteLine Now & ", Huawei-OLT: " & objCurrentTab.Caption & "(" & objCurrentTab.session.RemoteAddress & ")  delete an ONT : " & _
			gpon_num & ":" & ont_num &", sn_number: " & sn_number
	logStream.close
end sub

'�ú����Ĺ����ǻ�ȡ���������������ʾ��
Function getstrPrompt(objCurrentTab)

	objCurrentTab.activate
	
	if objCurrentTab.Session.Connected = True  then
		
			objCurrentTab.Screen.Send vbcrlf
			objCurrentTab.Screen.WaitForString vbcr

			Do 
			' Attempt to detect the command prompt heuristically... 
				Do 
					bCursorMoved = objCurrentTab.Screen.WaitForCursor(1)
				Loop Until bCursorMoved = False
			' Once the cursor has stopped moving for about a second, we'll 
			' assume it's safe to start interacting with the remote system. 
			' Get the shell prompt so that we can know what to look for when 
			' determining if the command is completed. Won't work if the prompt 
			' is dynamic (e.g., changes according to current working folder, etc.) 
				nRow = objCurrentTab.Screen.CurrentRow 
				strPrompt = objCurrentTab.screen.Get(nRow, 0, nRow, objCurrentTab.Screen.CurrentColumn - 1)
				' Loop until we actually see a line of text appear (the 
				' timeout for WaitForCursor above might not be enough 
				' for slower-responding hosts. 
				strPrompt = Trim(strPrompt)
				If strPrompt <> "" Then Exit Do
			Loop 
		
			getstrPrompt = strPrompt
		
	end if

End Function
