# $language = "VBScript"
# $interface = "1.0"
'本脚本适用于secureCRT 7.1 以上版本，其它低版本没有进行测试
'本脚本适用于自动配置华为GPON的ONU的自动配置
'本脚本的作者是 山东广电网络有限公司济宁分公司的 姬广超
'Email:	newchoose@163.com	欢迎交流学习
'2016年12月3日

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
	
	'获取外层vlan信息
	'打开当前目录下的huawei_olt_info.csv文件，并将文件内容全部读取到strFileData里面
	set objStream = g_fso.openTextFile(".\huawei_olt_info.csv", 1, false)
	strFileData = objStream.ReadAll
	objStream.close

	vLines = split(strFileData, vbcrlf)
	'分析读取到的每一行，并且忽略首行
	for i = 1 to UBound(vLines)
		oltInfo = split(vLines(i), ",")
		if inStr(objCurrentTab.Caption, oltInfo(1)) then
			'获取配置文件中外层vlan的信息
			cvlan2 = oltInfo(4)
			cvlan1 = oltInfo(5)
			'获取OLT名称和ip信息
			olt_name = oltInfo(0)
			olt_ip = oltInfo(1)
			username = oltInfo(2)
			passwd = oltInfo(3)
			exit for
		end if
	next
	
	if olt_name = "" then
		msgbox "没有找到当前ZTE-OLT的配置信息！"
		exit sub
	end if
	
	'重置一下连接，输入密码
	objCurrentTab.Session.Disconnect
	objCurrentTab.session.Connect()
	'输入用户名和密码，并确认是否进入#号模式，并写日志
	inputPasswd olt_ip, username, passwd
	'获取当前界面的命令提示符
	strPrompt = getstrPrompt(objCurrentTab)
	'strP = getstrPrompt(objCurrentTab)
	'strPrompt = Left(strP, len(strP) - 1)
	'msgbox strPrompt
	
	do	
		'如果连接断开则不断的尝试重新连接
		On Error Resume Next
	
		objCurrentTab.Screen.Synchronous = True
		objCurrentTab.Screen.Send "display ont autofind all" & vbcr
		objCurrentTab.Screen.waitForString vbcr
		'防止出现多个ONT同时上线
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
			'用正则表达式提出gpon_num和ont_num
			If re.Test(strResult) <> True Then
				MsgBox "异常错误！"
				crt.quit
			Else
				Set matches = re.Execute(strResult)
				For Each match In matches
					gpon_num = match.SubMatches(0)
					sn_number = match.SubMatches(1)
					exit for
				Next
			End If
			'配置ONU的所有必要参数都获取到啦，接下来就调用一个Sub就OK啦
			config_huawei_ONU gpon_num, sn_number, cvlan2, cvlan1
			
		end if
		
		nError = Err.Number 
		strErr = Err.Description 
		' Restore normal error handling so that any other errors in our 
		' script are not masked/ignored 
		On Error Goto 0 
		
		'发现错误进行写日志，并且尝试重新连接远程OLT
		If nError <> 0 Then
			Set logStream = g_fso.OpenTextFile(".\huawei_onu_config_log.txt", 8, True)
			logStream.writeLine Now & ", An Error happened on Huawei-OLT: " & olt_name & "(" & olt_ip & ") . Error: " & strErr
			objCurrentTab.Session.Disconnect
			logStream.writeLine Now & ", The session for Huawei-OLT: " & olt_name & "(" & olt_ip & ") was disconnected. Trying reConnect..."
			logStream.close
			objCurrentTab.session.Connect()
			'输入用户名和密码，并确认是否进入#号模式
			inputPasswd olt_ip, username, passwd
		end if
		'给定时间内休息一会
		crt.sleep delayTime
	loop 
	
End Sub

'该过程的作用是：输入OLT的用户名和密码（用户名和密码根据当地自己情况）, 并确认是否进入#号模式
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
	'判断是否进入#号模式
	if objCurrentTab.Screen.WaitForString("#", 3) <> true then
		msgbox "没有进入#号模式，请检查用户名和密码相关信息！程序执行失败！"
		crt.Quit
	end if
	
	'打开日志文件, 如果没有则新建该文件
	Set logStream = g_fso.OpenTextFile(".\huawei_onu_config_log.txt", 8, True)
	logStream.writeLine Now & ", The Script has been running at Huawei-OLT: " & olt_name & "(" & olt_ip & ")"
	logStream.close
	objCurrentTab.Screen.Synchronous = false
End Sub

'该函数的功能是获取给定界面的命令提示符
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


'该过程的功能是用给定的参数配置ONU
'epon_num：为指定ONU所在的PON口的序号
'sn_number：为指定ONU的SN号
'cvlan2：  为互联网的外层vlan
'cvlan1:   为点播的外层vlan
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
	'如果有重复的onu
	if instr(strResult, "Failure: SN already exists") then 
		objCurrentTab.Screen.Send "quit" & vbcr
		objCurrentTab.Screen.WaitForString "#"
		objCurrentTab.Screen.Send "quit" & vbcr
		objCurrentTab.Screen.WaitForString "#"
		delete_huawei_ONU(sn_number)
		objCurrentTab.Screen.Synchronous = false
		exit sub
	end if
	'下面的代码功能是返回的结果进行分行
	strLines = Split(strResult, vbcrlf)
	str2 = Split(strLines(1), ":")
	ontid = str2(2)
	'根据ONU的序号自动计算出互联网和点播的内层Vlan
	onu_vlan2 = 2000 + ontid
	onu_vlan1 = 1000 + ontid
	'objCurrentTab.Screen.Synchronous = True
	'退到config界面
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
	
	'回到pon config 界面
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
	'退到config界面
	objCurrentTab.Screen.Send "quit" & vbCr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "save configuration" & vbCr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "quit" & vbCr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Synchronous = false
	'写日志是必须的啊，这才显的专业!
	Set logStream = g_fso.OpenTextFile(".\huawei_onu_config_log.txt", 8, True)
	logStream.WriteLine Now & ", Huawei-OLT: " & objCurrentTab.Caption & "(" & objCurrentTab.session.RemoteAddress & ")  add an ONT : " & _
			epon_num & ":" & ontid &", sn_number: " & sn_number
	logStream.close
	'暂停25秒等待配置保存完毕
	crt.sleep 1000 * 25

End Sub

'sn-auth 4857544374692E75
'在#号模式下 删除onu 以及相应的service,
Sub delete_huawei_ONU(sn_number)
	'----------下面代码查找 ont pon 口 和序号信息----------
	'查找ont pon 口 和序号信息
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
	'用正则表达式提出gpon_num和ont_num
	If re.Test(strResult) <> True Then
		MsgBox "异常错误！"
		crt.quit
	Else
		Set matches = re.Execute(strResult)
		For Each match In matches
			gpon_num = match.SubMatches(0)
			ont_num = match.SubMatches(1)
			exit for
		Next
	End If
	'--------------下面代码执行删除service-port 命令------------
	objCurrentTab.Screen.Synchronous = True
	objCurrentTab.Screen.Send "config" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "undo service-port port " & gpon_num & " ont " & ont_num & chr(13)
	objCurrentTab.Screen.waitForString "{ <cr>|gemport<K> }:"
	objCurrentTab.Screen.Send vbcr
	objCurrentTab.Screen.WaitForString "Are you sure to release service virtual port(s)? (y/n)[n]:"
	objCurrentTab.Screen.Send "y" & chr(13)
	objCurrentTab.Screen.WaitForString "#"
	'----------下面代码执行ONU ETH口的native-vlan修改成默认的VLAN 1操作---------
	'在config模式下进入  gpon ,
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
	'----------执行删除ont操作---------
	'ont delete 7 2
	objCurrentTab.Screen.Send "ont delete " & str1(2) & " " & ont_num & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "quit" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "quit" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Synchronous = false
	'写日志是必须的啊，这才显的专业!
	Set logStream = g_fso.OpenTextFile(".\huawei_onu_config_log.txt", 8, True)
	logStream.WriteLine Now & ", Huawei-OLT: " & objCurrentTab.Caption & "(" & objCurrentTab.session.RemoteAddress & ")  delete an ONT : " & _
			gpon_num & ":" & ont_num &", sn_number: " & sn_number
	logStream.close
end sub
