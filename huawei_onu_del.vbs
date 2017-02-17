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
	'获取当前界面的命令提示符
	set objCurrentTab = crt.GetScriptTab
	strPrompt = getstrPrompt(objCurrentTab)
	sn_number = crt.Dialog.Prompt("请输入要查询的ONU的16位SN号码", "请输入", "4857544374692E75", false)
	if len(sn_number) = 16 then
		delete_huawei_ONU(sn_number)
		msgbox "删除成功！"
	else
		msgbox "SN号码位数不够！"
	end if
end sub

'sn-auth 4857544374692E75
'在#号模式下 删除onu 以及相应的service,
Sub delete_huawei_ONU(sn_number)
	'----------下面代码查找 ont pon 口 和序号信息----------
	'查找ont pon 口 和序号信息
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
	'用正则表达式提出gpon_num和ont_num
	If re.Test(strCompleteOutput) <> True Then
		MsgBox "异常错误！"
		crt.quit
	Else
		Set matches = re.Execute(strCompleteOutput)
		For Each match In matches
			gpon_num = match.SubMatches(0)
			ont_num = match.SubMatches(1)
			exit for
		Next
	End If
	'--------------下面代码执行查找并删除service-port 命令------------
	'查询 service-port 信息
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
	'用正则表达式提出gpon_num和ont_num
	If re.Test(strCompleteOutput) <> True Then
		'MsgBox "异常错误！"
		'crt.quit
	Else
		Set matches = re.Execute(strCompleteOutput)
		For Each match In matches
			service_num = match.SubMatches(0)
			'执行删除 service-port 命令
			objCurrentTab.Screen.Send "undo service-port " & service_num & chr(13)
			objCurrentTab.Screen.waitForString "#"
		Next
	End If
	'----------下面代码执行ONU ETH口的native-vlan修改成默认的VLAN 1操作---------
	'在config模式下进入  gpon ,
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
