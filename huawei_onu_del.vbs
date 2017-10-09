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
	objCurrentTab.Screen.Send "display ont info by-sn " & sn_number & chr(13) 
	objCurrentTab.Screen.waitForString vbcr

	strResult = objCurrentTab.Screen.ReadString("---- More ( Press 'Q' to break ) ----")	
	objCurrentTab.Screen.Send "Q" & chr(13) 
	objCurrentTab.Screen.waitForString vbcr
	objCurrentTab.Screen.Synchronous = false
	' -----------------------------------------------------------------------------
	' F/S/P                   : 0/5/7
	' ONT-ID                  : 2
	' Control flag            : active
	' Run state               : online
	' Config state            : normal
	' Match state             : match
	' DBA type                : SR
	' ONT distance(m)         : 5
	' ONT battery state       : not support
	' Memory occupation       : 13%
	' CPU occupation          : 1%
	' Temperature             : 50(C)
	' Authentic type          : SN-auth
	' SN                      : 4857544374692E75 (HWTC-74692E75)
	' Management mode         : OMCI
	' Software work mode      : normal
	' Isolation state         : normal
	' Description             : 4857544374692E75
	' Last down cause         : -
	' Last up time            : 2017-10-09 21:03:44+08:00
	' Last down time          : -
	' Last dying gasp time    : -
                                   

	' DYGD_MA5680T#
	re.Pattern = "F/S/P\s+:\s(\d/\d+/\d+)[\s\S]{5,30}:\s(\d+)"
	'��������ʽ���gpon_num��ont_num
	If re.Test(strResult) <> True Then
		MsgBox "�쳣����"
		crt.quit
	Else
		Set matches = re.Execute(strResult)
		For Each match In matches
			gpon_num = match.SubMatches(0)
			ont_num = match.SubMatches(1)
			exit for
		Next
	End If
	'--------------�������ִ��ɾ��service-port ����------------
	objCurrentTab.Screen.Synchronous = True
	objCurrentTab.Screen.Send "config" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "undo service-port port " & gpon_num & " ont " & ont_num & chr(13)
	objCurrentTab.Screen.waitForString "{ <cr>|gemport<K> }:"
	objCurrentTab.Screen.Send vbcr
	objCurrentTab.Screen.WaitForString "Are you sure to release service virtual port(s)? (y/n)[n]:"
	objCurrentTab.Screen.Send "y" & chr(13)
	objCurrentTab.Screen.WaitForString "#"
	'----------�������ִ��ONU ETH�ڵ�native-vlan�޸ĳ�Ĭ�ϵ�VLAN 1����---------
	'��configģʽ�½���  gpon ,
	str1 = Split(gpon_num, "/")
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
