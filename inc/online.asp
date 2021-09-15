<!--#include file="../conn.asp" -->
<!--#include file="const.asp"-->
<%
'=====================================================================
' 软件名称：恩池网站管理系统
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Response.Expires = 0
If Not IsNumeric(Request("id")) And Request("id")<>"" then
	Response.write"错误的系统参数!ID必须是数字"
	Response.End
End If
Dim rsOnline,strUsername,statuserid,stridentitys,remoteaddr,onlinemany
Dim Rs,SQL,Grades,strReferer,onlinemember,userid,BrowserType,CurrentStation
Application.Lock
If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = Empty Then
	remoteaddr = Request.ServerVariables("REMOTE_ADDR")
Else
	remoteaddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
End If
strReferer = enchiasp.Checkstr(Request.Querystring("Referer"))
If strReferer = Empty Then
	strReferer = "★直接输入或书签导入★"
Else
	strReferer = enchiasp.CheckStr(Left(strReferer,250))
End If
CurrentStation = enchiasp.Checkstr(Request.Querystring("stat"))
If enchiasp.membername = "" Then
	Grades = 0
	strUsername = "匿名用户"
	userid = 0
Else
	Grades = CInt(enchiasp.membergrade)
	strUsername = Trim(enchiasp.membername)
	userid = CLng(enchiasp.memberid)
End If
Set Rs=enchiasp.Execute("select GroupName from ECCMS_UserGroup where Grades = "& Grades)
	stridentitys = Rs("GroupName")
Rs.Close
Set Rs=Nothing
Set BrowserType=new SystemInfo_Cls
Call UserActiveOnline
Set BrowserType=Nothing
Application.UnLock
'---- 删除超时用户
Conn.Execute("delete from ECCMS_Online where DateDIff('s',lastTime,Now()) > "& CLng(enchiasp.ActionTime) &" * 60")
onlinemany = Conn.Execute("Select Count(*) from ECCMS_Online")(0)
onlinemember = Conn.Execute("Select Count(*) from ECCMS_Online where userid <> 0")(0)
If CInt(Request.Querystring("id")) = 1 And Trim(Request.Querystring("id")) <> "" Then
	Response.Write "document.writeln(" & chr(34) & ""& onlinemany &""& chr(34) & ");"
ElseIf CInt(Request.Querystring("id")) = 2 And Trim(Request.Querystring("id")) <> "" Then
	Response.Write "document.writeln(" & Chr(34) & ""& onlinemember &""& chr(34) & ");"
Else
	Response.Write "document.writeln(" & Chr(34) & chr(34) & ");"
End If

Sub UserActiveOnline()
	Dim UserSessionID,OnlineSQL
	UserSessionID = Session.sessionid
	SQL = "select * from [ECCMS_Online] where ip='" & remoteaddr & "' And username='" & strUsername & "' Or id=" & UserSessionID
	Set rsOnline = enchiasp.Execute(SQL)
	If rsOnline.BOF And rsOnline.EOF Then
		OnlineSQL = "Insert Into ECCMS_Online(id,username,identitys,station,ip,browser,startTime,lastTime,userid,strReferer) Values (" & UserSessionID & ",'" & strUsername & "','" & stridentitys & "','" & CurrentStation & "','" & remoteaddr & "','" & BrowserType.platform&"|"&BrowserType.Browser&BrowserType.version & "'," & NowString & "," & NowString & "," & userid & ",'" & strReferer & "')"
	Else
		OnlineSQL = "update ECCMS_Online Set ID=" & UserSessionID & ",username='" & strUsername & "',identitys='" & stridentitys & "',station='" & CurrentStation & "',lastTime=" & NowString & ",userid=" & userid & " Where ID = " & UserSessionID
	End If
	Conn.Execute(OnlineSQL)
	rsOnline.close
	Set rsOnline = Nothing
End Sub
CloseConn
Class SystemInfo_Cls
	Public Browser, version, platform, IsSearch
	Private Sub Class_Initialize()
		Dim Agent, Tmpstr
		IsSearch = False
		If Not IsEmpty(Session("SystemInfo_Cls")) Then
			Tmpstr = Split(Session("SystemInfo_Cls"), "|||")
			Browser = Tmpstr(0)
			version = Tmpstr(1)
			platform = Tmpstr(2)
			If Tmpstr(3) = "1" Then
				IsSearch = True
			End If
			Exit Sub
		End If
		Browser = "unknown"
		version = "unknown"
		platform = "unknown"
		Agent = Request.ServerVariables("HTTP_USER_AGENT")
		If Left(Agent, 7) = "Mozilla" Then '有此标识为浏览器
			Agent = Split(Agent, ";")
			If InStr(Agent(1), "MSIE") > 0 Then
				Browser = "Internet Explorer "
				version = Trim(Left(Replace(Agent(1), "MSIE", ""), 6))
			ElseIf InStr(Agent(4), "Netscape") > 0 Then
				Browser = "Netscape "
				Tmpstr = Split(Agent(4), "/")
				version = Tmpstr(UBound(Tmpstr))
			ElseIf InStr(Agent(4), "rv:") > 0 Then
				Browser = "Mozilla "
				Tmpstr = Split(Agent(4), ":")
				version = Tmpstr(UBound(Tmpstr))
				If InStr(version, ")") > 0 Then
					Tmpstr = Split(version, ")")
					version = Tmpstr(0)
				End If
			End If
			If InStr(Agent(2), "NT 5.2") > 0 Then
				platform = "Windows 2003"
			ElseIf InStr(Agent(2), "Windows CE") > 0 Then
				platform = "Windows CE"
			ElseIf InStr(Agent(2), "NT 5.1") > 0 Then
				platform = "Windows XP"
			ElseIf InStr(Agent(2), "NT 4.0") > 0 Then
				platform = "Windows NT"
			ElseIf InStr(Agent(2), "NT 5.0") > 0 Then
				platform = "Windows 2000"
			ElseIf InStr(Agent(2), "NT") > 0 Then
				platform = "Windows NT"
			ElseIf InStr(Agent(2), "9x") > 0 Then
				platform = "Windows ME"
			ElseIf InStr(Agent(2), "98") > 0 Then
				platform = "Windows 98"
			ElseIf InStr(Agent(2), "95") > 0 Then
				platform = "Windows 95"
			ElseIf InStr(Agent(2), "Win32") > 0 Then
				platform = "Win32"
			ElseIf InStr(Agent(2), "Linux") > 0 Then
				platform = "Linux"
			ElseIf InStr(Agent(2), "SunOS") > 0 Then
				platform = "SunOS"
			ElseIf InStr(Agent(2), "Mac") > 0 Then
				platform = "Mac"
			ElseIf UBound(Agent) > 2 Then
				If InStr(Agent(3), "NT 5.1") > 0 Then
					platform = "Windows XP"
				End If
				If InStr(Agent(3), "Linux") > 0 Then
					platform = "Linux"
				End If
			End If
			If InStr(Agent(2), "Windows") > 0 And platform = "unknown" Then
				platform = "Windows"
			End If
		ElseIf Left(Agent, 5) = "Opera" Then '有此标识为浏览器
			Agent = Split(Agent, "/")
			Browser = "Mozilla "
			Tmpstr = Split(Agent(1), " ")
			version = Tmpstr(0)
			If InStr(Agent(1), "NT 5.2") > 0 Then
				platform = "Windows 2003"
			ElseIf InStr(Agent(1), "Windows CE") > 0 Then
				platform = "Windows CE"
			ElseIf InStr(Agent(1), "NT 5.1") > 0 Then
				platform = "Windows XP"
			ElseIf InStr(Agent(1), "NT 4.0") > 0 Then
				platform = "Windows NT"
			ElseIf InStr(Agent(1), "NT 5.0") > 0 Then
				platform = "Windows 2000"
			ElseIf InStr(Agent(1), "NT") > 0 Then
				platform = "Windows NT"
			ElseIf InStr(Agent(1), "9x") > 0 Then
				platform = "Windows ME"
			ElseIf InStr(Agent(1), "98") > 0 Then
				platform = "Windows 98"
			ElseIf InStr(Agent(1), "95") > 0 Then
				platform = "Windows 95"
			ElseIf InStr(Agent(1), "Win32") > 0 Then
				platform = "Win32"
			ElseIf InStr(Agent(1), "Linux") > 0 Then
				platform = "Linux"
			ElseIf InStr(Agent(1), "SunOS") > 0 Then
				platform = "SunOS"
			ElseIf InStr(Agent(1), "Mac") > 0 Then
				platform = "Mac"
			ElseIf UBound(Agent) > 2 Then
				If InStr(Agent(3), "NT 5.1") > 0 Then
					platform = "Windows XP"
				End If
				If InStr(Agent(3), "Linux") > 0 Then
					platform = "Linux"
				End If
			End If
		Else
			'识别搜索引擎
			Dim botlist, i
			botlist = "Google,Isaac,Webdup,SurveyBot,Baiduspider,ia_archiver,P.Arthur,FAST-WebCrawler,Java,Microsoft-ATL-Native,TurnitinBot,WebGather,Sleipnir"
			botlist = Split(botlist, ",")
			For i = 0 To UBound(botlist)
				If InStr(Agent, botlist(i)) > 0 Then
					platform = botlist(i) & "搜索器"
					IsSearch = True
					Exit For
				End If
			Next
		End If
		If IsSearch Then
			Browser = ""
			version = ""
			Session("SystemInfo_Cls") = Browser & "|||" & version & "|||" & platform & "|||1"
		Else
			Session("SystemInfo_Cls") = Browser & "|||" & version & "|||" & platform & "|||0"
		End If
	End Sub
End Class
%>