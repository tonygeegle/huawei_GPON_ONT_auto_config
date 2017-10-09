# $language = "VBScript"
# $interface = "1.0"
'���ű�������secureCRT 7.1 ���ϰ汾�������Ͱ汾û�н��в���
'���ű��������Զ����û�ΪGPON��ONU���Զ�����
'���ű��������� ɽ������������޹�˾�����ֹ�˾�� ���㳬
'Email:	newchoose@163.com	��ӭ����ѧϰ
'2016��12��3��

dim g_fso
set g_fso = CreateObject("Scripting.FileSystemObject")
Const ForWriting = 2 
Const ForAppending = 8

dim delayTime
delayTime = 1000 * 8

dim objCurrentTab
dim strPrompt
Set re = New RegExp
re.Global = True  
re.Pattern = "(\d/\s\d+/\d+)\s+(\d+)\s+"

Sub main()

	dim olt_name
	olt_name = ""
	
	set objCurrentTab = crt.GetScriptTab
	
	'��ȡ���vlan��Ϣ
	'�򿪵�ǰĿ¼�µ�huawei_olt_info.csv�ļ��������ļ�����ȫ����ȡ��strFileData����
	set objStream = g_fso.openTextFile(".\huawei_olt_info.csv", 1, false)
	strFileData = objStream.ReadAll
	objStream.close

	vLines = split(strFileData, vbcrlf)
	'������ȡ����ÿһ�У����Һ�������
	for i = 1 to UBound(vLines)
		oltInfo = split(vLines(i), ",")
		if inStr(objCurrentTab.Caption, oltInfo(1)) then
			'��ȡ�����ļ������vlan����Ϣ
			cvlan2 = oltInfo(4)
			cvlan1 = oltInfo(5)
			'��ȡOLT���ƺ�ip��Ϣ
			olt_name = oltInfo(0)
			olt_ip = oltInfo(1)
			username = oltInfo(2)
			passwd = oltInfo(3)
			exit for
		end if
	next
	
	if olt_name = "" then
		msgbox "û���ҵ���ǰZTE-OLT��������Ϣ��"
		exit sub
	end if
	
	'����һ�����ӣ���������
	objCurrentTab.Session.Disconnect
	objCurrentTab.session.Connect()
	'�����û��������룬��ȷ���Ƿ����#��ģʽ����д��־
	inputPasswd olt_ip, username, passwd
	'��ȡ��ǰ�����������ʾ��
	strPrompt = getstrPrompt(objCurrentTab)
	'strP = getstrPrompt(objCurrentTab)
	'strPrompt = Left(strP, len(strP) - 1)
	'msgbox strPrompt
	
	do	
		'������ӶϿ��򲻶ϵĳ�����������
		On Error Resume Next
	
		objCurrentTab.Screen.Synchronous = True
		objCurrentTab.Screen.Send "display ont autofind all" & vbcr
		objCurrentTab.Screen.waitForString vbcr
		'��ֹ���ֶ��ONTͬʱ����
		strResult = objCurrentTab.Screen.ReadString("---- More ( Press 'Q' to break ) ----", strPrompt)
		If objCurrentTab.Screen.MatchIndex = 1 Then crt.Screen.Send "Q" & chr(13)
		
		objCurrentTab.Screen.Synchronous = false
		REM ------------------------------------------------------------------------
		REM Number              : 1
		REM F/S/P               : 0/5/7
		REM Ont SN              : 4857544374692E75
		REM Password            : 0x00000000000000000000
		REM Loid                : 
		REM Checkcode           : 
		REM VenderID            : HWTC
		REM Ont Version         : E24.B
		REM Ont SoftwareVersion : V8R016C00S200
		REM Ont EquipmentID     : MA5671-G4
		REM Ont autofind time   : 2016-12-12 20:13:32+08:00
		REM ------------------------------------------------------------------------
		REM The number of GPON autofind ONT is 1
		if Instr(strResult, "Ont autofind time") then 
			re.Pattern = "F/S/P\s+:\s(\d/\d+/\d+)[\s\S]{5,30}:\s([0-9A-F]{16})"
			'��������ʽ���gpon_num��ont_num
			If re.Test(strResult) <> True Then
				MsgBox "�쳣����"
				crt.quit
			Else
				Set matches = re.Execute(strResult)
				For Each match In matches
					gpon_num = match.SubMatches(0)
					sn_number = match.SubMatches(1)
					exit for
				Next
			End If
			'����ONU�����б�Ҫ��������ȡ�������������͵���һ��Sub��OK��
			config_huawei_ONU gpon_num, sn_number, cvlan2, cvlan1
			
		end if
		
		nError = Err.Number 
		strErr = Err.Description 
		' Restore normal error handling so that any other errors in our 
		' script are not masked/ignored 
		On Error Goto 0 
		
		'���ִ������д��־�����ҳ�����������Զ��OLT
		If nError <> 0 Then
			Set logStream = g_fso.OpenTextFile(".\huawei_onu_config_log.txt", 8, True)
			logStream.writeLine Now & ", An Error happened on Huawei-OLT: " & olt_name & "(" & olt_ip & ") . Error: " & strErr
			objCurrentTab.Session.Disconnect
			logStream.writeLine Now & ", The session for Huawei-OLT: " & olt_name & "(" & olt_ip & ") was disconnected. Trying reConnect..."
			logStream.close
			objCurrentTab.session.Connect()
			'�����û��������룬��ȷ���Ƿ����#��ģʽ
			inputPasswd olt_ip, username, passwd
		end if
		'����ʱ������Ϣһ��
		crt.sleep delayTime
	loop 
	
End Sub

'�ù��̵������ǣ�����OLT���û��������루�û�����������ݵ����Լ������, ��ȷ���Ƿ����#��ģʽ
REM strResult = objCurrentTab.Screen.ReadString("---- More ( Press 'Q' to break ) ----", strPrompt)
		REM strCompleteOutput = strCompleteOutput & strResult
		REM If objCurrentTab.Screen.MatchIndex = 1 Then crt.Screen.Send " "
		REM If objCurrentTab.Screen.MatchIndex = 2 Then Exit Do

Sub inputPasswd(olt_ip, username, passwd)

	
	Set objCurrentTab = crt.GetScriptTab
	objCurrentTab.Screen.Synchronous = True
	
	objCurrentTab.Screen.WaitForString "name:"
	objCurrentTab.Screen.Send username & chr(13)
	objCurrentTab.Screen.WaitForString "assword:"
	objCurrentTab.Screen.Send passwd & chr(13)
	nn = objCurrentTab.Screen.WaitForStrings("---- More ( Press 'Q' to break ) ----", ">")
	if nn = 1 then objCurrentTab.Screen.Send " "
	objCurrentTab.Screen.Send "en" & chr(13)
	'�ж��Ƿ����#��ģʽ
	if objCurrentTab.Screen.WaitForString("#", 3) <> true then
		msgbox "û�н���#��ģʽ�������û��������������Ϣ������ִ��ʧ�ܣ�"
		crt.Quit
	end if
	
	'����־�ļ�, ���û�����½����ļ�
	Set logStream = g_fso.OpenTextFile(".\huawei_onu_config_log.txt", 8, True)
	logStream.writeLine Now & ", The Script has been running at Huawei-OLT: " & olt_name & "(" & olt_ip & ")"
	logStream.close
	objCurrentTab.Screen.Synchronous = false
End Sub

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


'�ù��̵Ĺ������ø����Ĳ�������ONU
'epon_num��Ϊָ��ONU���ڵ�PON�ڵ����
'sn_number��Ϊָ��ONU��SN��
'cvlan2��  Ϊ�����������vlan
'cvlan1:   Ϊ�㲥�����vlan
Sub config_huawei_ONU(epon_num, sn_number, cvlan2, cvlan1)
	
	str1 = Split(epon_num, "/")
	objCurrentTab.Screen.Synchronous = True
	objCurrentTab.Screen.Send "config" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "interface gpon " & str1(0) & "/" & str1(1) & vbcr
	objCurrentTab.Screen.WaitForString "#"
	'ont add 7 sn-auth 4857544374692E75 omci ont-lineprofile-name jn ont-srvprofile-name jn desc 4857544374692E75 
	objCurrentTab.Screen.Send "ont add " & str1(2) & " sn-auth " & sn_number & " omci ont-lineprofile-name jn ont-srvprofile-name jn desc " & sn_number & vbcr
	objCurrentTab.Screen.waitForString vbcr
	strResult = crt.Screen.ReadString(")#")
	'objCurrentTab.Screen.Synchronous = false
	'msgbox strResult
	'������ظ���onu
	if instr(strResult, "Failure: SN already exists") then 
		objCurrentTab.Screen.Send "quit" & vbcr
		objCurrentTab.Screen.WaitForString "#"
		objCurrentTab.Screen.Send "quit" & vbcr
		objCurrentTab.Screen.WaitForString "#"
		delete_huawei_ONU(sn_number)
		objCurrentTab.Screen.Synchronous = false
		exit sub
	end if
	'����Ĵ��빦���Ƿ��صĽ�����з���
	strLines = Split(strResult, vbcrlf)
	str2 = Split(strLines(1), ":")
	ontid = str2(2)
	'����ONU������Զ�������������͵㲥���ڲ�Vlan
	onu_vlan2 = 2000 + ontid
	onu_vlan1 = 1000 + ontid
	'objCurrentTab.Screen.Synchronous = True
	'�˵�config����
	objCurrentTab.Screen.Send "quit" & vbCr
	objCurrentTab.Screen.WaitForString "#"
	'service-port vlan 2200 gpon 0/5/7 ont 2 gemport 1 multi-service user-vlan 2000 tag-transform translate-and-add inner-vlan 2006
	objCurrentTab.Screen.Send "service-port vlan " & cvlan2 & " gpon " & epon_num & " ont " & ontid & " gemport 1 multi-service user-vlan 2000 tag-transform translate-and-add inner-vlan " & onu_vlan2 & vbcr
	objCurrentTab.Screen.WaitForString "{ <cr>|bundle<K>|inbound<K>|inner-priority<K>|rx-cttr<K> }:"
	objCurrentTab.Screen.Send vbCr
	objCurrentTab.Screen.WaitForString "#"
	'service-port vlan 1200 gpon 0/5/7 ont 2 gemport 2 multi-service user-vlan 1000 tag-transform translate-and-add inner-vlan 1006
    objCurrentTab.Screen.Send "service-port vlan " & cvlan1 & " gpon " & epon_num & " ont " & ontid & " gemport 2 multi-service user-vlan 1000 tag-transform translate-and-add inner-vlan " & onu_vlan1 & vbcr
	objCurrentTab.Screen.WaitForString "{ <cr>|bundle<K>|inbound<K>|inner-priority<K>|rx-cttr<K> }:"
	objCurrentTab.Screen.Send vbCr
	objCurrentTab.Screen.WaitForString "#"
	
	'�ص�pon config ����
	objCurrentTab.Screen.Send "interface gpon " & str1(0) & "/" & str1(1) & vbcr
	objCurrentTab.Screen.WaitForString "#"
	'ont port native-vlan 7 2 eth 1 vlan 2000
	objCurrentTab.Screen.Send "ont port native-vlan " & str1(2) & " " & ontid & " eth 1 vlan 2000" & vbcr
	objCurrentTab.Screen.WaitForString "<cr>|priority<K> }:"
	objCurrentTab.Screen.Send vbCr
	objCurrentTab.Screen.WaitForString "#"
	'ont port native-vlan 7 2 eth 1 vlan 2000
	objCurrentTab.Screen.Send "ont port native-vlan " & str1(2) & " " & ontid & " eth 2 vlan 1000" & vbcr
	objCurrentTab.Screen.WaitForString "<cr>|priority<K> }:"
	objCurrentTab.Screen.Send vbCr
	objCurrentTab.Screen.WaitForString "#"
	'�˵�config����
	objCurrentTab.Screen.Send "quit" & vbCr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "save configuration" & vbCr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "quit" & vbCr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Synchronous = false
	'д��־�Ǳ���İ�������Ե�רҵ!
	Set logStream = g_fso.OpenTextFile(".\huawei_onu_config_log.txt", 8, True)
	logStream.WriteLine Now & ", Huawei-OLT: " & objCurrentTab.Caption & "(" & objCurrentTab.session.RemoteAddress & ")  add an ONT : " & _
			epon_num & ":" & ontid &", sn_number: " & sn_number
	logStream.close
	'��ͣ25��ȴ����ñ������
	crt.sleep 1000 * 25

End Sub

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
