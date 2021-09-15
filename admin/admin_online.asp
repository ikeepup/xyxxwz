<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
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
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th>在线人数信息管理</th>
</tr>
<tr>
	<td class=tablerow1>菜单导航：<a href='admin_online.asp'>管理首页</a> | 
	<a href='admin_online.asp?action=zone'>详细地址</a> | 
	<a href='admin_online.asp?action=refer'>访问来源</a> |
	<a href='admin_online.asp?action=delall' onclick="{if(confirm('您确定要删除所有在线人数吗?')){return true;}return false;}"><font color=blue>删除所有在线人数</font></a></td>
</tr>
<tr>
	<td class=tablerow2>当前位置：在线人数统计信息</td>
</tr>
</table>
<br>
<%
Dim Action,i,strClass,sFileName
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum

maxperpage = 30 '###每页显示数
If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
	Response.Write ("错误的系统参数!请输入整数")
	Response.End
End If
If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
	CurrentPage = CInt(Request("page"))
Else
	CurrentPage = 1
End If
If CInt(CurrentPage) = 0 Then CurrentPage = 1
TotalNumber = enchiasp.Execute("Select Count(ID) from ECCMS_Online")(0)
TotalPageNum = CInt(TotalNumber / maxperpage)  '得到总页数
If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum

Action = LCase(Request("action"))
If Not ChkAdmin("Online") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Select Case Trim(Action)
	Case "refer"
		Call OnlineReferer
	Case "zone"
		Call OnlineZone
	Case "del"
		Call DelOnline
	Case "delall"
		Call DelAllOnline
	Case Else
		Call MainPage
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub MainPage()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th width='5%' nowrap>选择</th>
	<th nowrap>用 户 名</th>
	<th nowrap>访 问 时 间</th>
	<th nowrap>活 动 时 间</th>
	<th nowrap>用 户 IP 地 址</th>
	<th nowrap>操 作 系 统</th>
	<th nowrap>浏 览 器</th>
</tr>
<%
	sFileName = "admin_online.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Online] order by startTime desc"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=7 class=TableRow2>当前无人在线！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td class=tablerow2 colspan=7><%Call showpage()%></td>
</tr>
<form name=selform method=post action=?action=del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
%>
<tr align=center>
	<td <%=strClass%>><input type=checkbox name=OnlineID value='<%=Rs("id")%>'></td>
	<td <%=strClass%>><%=Rs("username")%></td>
	<td <%=strClass%>><%=Rs("startTime")%></td>
	<td <%=strClass%>><%=Rs("lastTime")%></td>
	<td <%=strClass%>><%=Rs("ip")%></td>
	<td <%=strClass%>><%=usersysinfo(Rs("browser"), 0)%></td>
	<td <%=strClass%>><%=usersysinfo(Rs("browser"), 1)%></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class=tablerow1 colspan=7>
	<input class=Button type="button" name="chkall" value="全选" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="反选" onClick="ContraSel(this.form)">
	<input class=Button type="submit" name="Submit2" value="删除" onclick="{if(confirm('您确定要删除此在线人员吗?')){this.document.selform.submit();return true;}return false;}">
	<input class=Button type="button" name="Submit3" value="清空所有在线人数" onclick="{if(confirm('您确定要清空所有在线人数吗?')){location.href='admin_online.asp?action=delall';return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td class=tablerow2 colspan=7><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub OnlineReferer()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th width='5%' nowrap>选择</th>
	<th width='15%' nowrap>来访时间/IP</th>
	<th>访 问 来 源</th>
	<th>当 前 位 置</th>
</tr>
<%
	sFileName = "admin_online.asp?action=refer&"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Online] order by startTime desc"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=4 class=TableRow2>当前无人在线！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td class=tablerow2 colspan=4><%Call showpage()%></td>
</tr>
<form name=selform method=post action=?action=del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
%>
<tr>
	<td align=center <%=strClass%>><input type=checkbox name=OnlineID value='<%=Rs("id")%>'></td>
	<td align=center <%=strClass%> nowrap><%=Rs("startTime")%><br><%=Rs("ip")%></td>
	<td <%=strClass%>><a href='<%=Rs("strReferer")%>' target=_blank><%=Rs("strReferer")%></a></td>
	<td <%=strClass%>><a href='<%=Rs("station")%>' target=_blank><%=Rs("station")%></a></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class=tablerow1 colspan=4>
	<input class=Button type="button" name="chkall" value="全选" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="反选" onClick="ContraSel(this.form)">
	<input class=Button type="submit" name="Submit2" value="删除" onclick="{if(confirm('您确定要删除此在线人员吗?')){this.document.selform.submit();return true;}return false;}">
	<input class=Button type="button" name="Submit3" value="清空所有在线人数" onclick="{if(confirm('您确定要清空所有在线人数吗?')){location.href='admin_online.asp?action=delall';return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td class=tablerow2 colspan=7><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub OnlineZone()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th width='5%' nowrap>选择</th>
	<th nowrap>用 户 名</th>
	<th nowrap>用 户 组</th>
	<th nowrap>IP 地 址</th>
	<th nowrap>详 细 地 址</th>
	<th nowrap>操 作 系 统</th>
	<th nowrap>浏 览 器</th>
</tr>
<%
	sFileName = "admin_online.asp?action=zone&"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Online] order by startTime desc"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=7 class=TableRow2>当前无人在线！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td class=tablerow2 colspan=7><%Call showpage()%></td>
</tr>
<form name=selform method=post action=?action=del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
%>
<tr align=center>
	<td <%=strClass%>><input type=checkbox name=OnlineID value='<%=Rs("id")%>'></td>
	<td <%=strClass%>><%=Rs("username")%></td>
	<td <%=strClass%>><%=Rs("identitys")%></td>
	<td <%=strClass%>><%=Rs("ip")%></td>
	<td <%=strClass%>><%=GetAddress(Rs("ip"))%></td>
	<td <%=strClass%>><%=usersysinfo(Rs("browser"), 0)%></td>
	<td <%=strClass%>><%=usersysinfo(Rs("browser"), 1)%></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class=tablerow1 colspan=7>
	<input class=Button type="button" name="chkall" value="全选" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="反选" onClick="ContraSel(this.form)">
	<input class=Button type="submit" name="Submit2" value="删除" onclick="{if(confirm('您确定要删除此在线人员吗?')){this.document.selform.submit();return true;}return false;}">
	<input class=Button type="button" name="Submit3" value="清空所有在线人数" onclick="{if(confirm('您确定要清空所有在线人数吗?')){location.href='admin_online.asp?action=delall';return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td class=tablerow2 colspan=7><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub DelAllOnline()
	enchiasp.Execute("Delete From ECCMS_Online")
	Call OutputScript ("在线人数全部清除完成！","admin_online.asp")
End Sub

Private Sub DelOnline()
	Dim OnlineID
	If Request("OnlineID") <> "" Then
		OnlineID = Request("OnlineID")
		enchiasp.Execute("Delete From ECCMS_Online where ID in (" & OnlineID & ")")
		OutHintScript ("在线人数删除成功！")
	Else
		OutAlertScript("请选择正确的系统参数！")
	End If
End Sub
Private Sub showpage()
	Dim n
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action='" & sFileName & "'><tr><td align=center> " & vbCrLf
	If CurrentPage < 2 Then
		Response.Write "当前在线人数 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 人&nbsp;首 页&nbsp;上一页&nbsp;|&nbsp;"
	Else
		Response.Write "当前在线人数 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 人&nbsp;<a href=" & sFileName & "page=1>首 页</a>&nbsp;"
		Response.Write "<a href=" & sFileName & "page=" & CurrentPage - 1 & ">上一页</a>&nbsp;|&nbsp;"
	End If
	If n - CurrentPage < 1 Then
		Response.Write "下一页&nbsp;尾 页" & vbCrLf
	Else
		Response.Write "<a href=" & sFileName & "page=" & (CurrentPage + 1) & ">下一页</a>"
		Response.Write "&nbsp;<a href=" & sFileName & "page=" & n & ">尾 页</a>" & vbCrLf
	End If
	Response.Write "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
	Response.Write "&nbsp;转到："
	Response.Write "<input name=page size=3 value='" & CurrentPage & "'> <input class=Button type=submit name=Submit value='转到'>"
	Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub

Private Function usersysinfo(info, getinfo)
	Dim usersys
	usersys = Split(info, "|")
	usersysinfo = usersys(getinfo)
End Function

Public Function GetAddress(sip)
	If Len(sip) < 5 Then
		GetAddress = "未知"
		Exit Function
	End If
	On Error Resume Next
	Dim Wry,IPType
	Set Wry = New TQQWry
	If Not Wry.IsIp(sip) Then
		GetAddress = " 未知"
		Exit Function
	End If
	IPType = Wry.QQWry(sip)
	GetAddress = Wry.Country & " " & Wry.LocalStr
End Function

Class TQQWry
	' ============================================
	' 变量声名
	' ============================================
	Dim Country, LocalStr, Buf, OffSet
	Private StartIP, EndIP, CountryFlag
	Public QQWryFile
	Public FirstStartIP, LastStartIP, RecordCount
	Private Stream, EndIPOff
	' ============================================
	' 类模块初始化
	' ============================================
	Private Sub Class_Initialize
		On Error Resume Next
		Country 		= ""
		LocalStr 		= ""
		StartIP 		= 0
		EndIP 			= 0
		CountryFlag 	= 0 
		FirstStartIP 	= 0 
		LastStartIP 	= 0 
		EndIPOff 		= 0 
		QQWryFile = Server.MapPath("../DataBase/IPAddress.dat") 'QQ IP库路径，要转换成物理路径
	End Sub
	' ============================================
	' IP地址转换成整数
	' ============================================
	Function IPToInt(IP)
		Dim IPArray, i
		IPArray = Split(IP, ".", -1)
		FOr i = 0 to 3
			If Not IsNumeric(IPArray(i)) Then IPArray(i) = 0
			If CInt(IPArray(i)) < 0 Then IPArray(i) = Abs(CInt(IPArray(i)))
			If CInt(IPArray(i)) > 255 Then IPArray(i) = 255
		Next
		IPToInt = (CInt(IPArray(0))*256*256*256) + (CInt(IPArray(1))*256*256) + (CInt(IPArray(2))*256) + CInt(IPArray(3))
	End Function
	' ============================================
	' 整数逆转IP地址
	' ============================================
	Function IntToIP(IntValue)
		p4 = IntValue - Fix(IntValue/256)*256
		IntValue = (IntValue-p4)/256
		p3 = IntValue - Fix(IntValue/256)*256
		IntValue = (IntValue-p3)/256
		p2 = IntValue - Fix(IntValue/256)*256
		IntValue = (IntValue - p2)/256
		p1 = IntValue
		IntToIP = Cstr(p1) & "." & Cstr(p2) & "." & Cstr(p3) & "." & Cstr(p4)
	End Function
	' ============================================
	' 获取开始IP位置
	' ============================================
	Private Function GetStartIP(RecNo)
		OffSet = FirstStartIP + RecNo * 7
		Stream.Position = OffSet
		Buf = Stream.Read(7)
		
		EndIPOff = AscB(MidB(Buf, 5, 1)) + (AscB(MidB(Buf, 6, 1))*256) + (AscB(MidB(Buf, 7, 1))*256*256) 
		StartIP  = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256) + (AscB(MidB(Buf, 4, 1))*256*256*256)
		GetStartIP = StartIP
	End Function
	' ============================================
	' 获取结束IP位置
	' ============================================
	Private Function GetEndIP()
		Stream.Position = EndIPOff
		Buf = Stream.Read(5)
		EndIP = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256) + (AscB(MidB(Buf, 4, 1))*256*256*256) 
		CountryFlag = AscB(MidB(Buf, 5, 1))
		GetEndIP = EndIP
	End Function
	' ============================================
	' 获取地域信息，包含国家和和省市
	' ============================================
	Private Sub GetCountry(IP)
		If (CountryFlag = 1 Or CountryFlag = 2) Then
			Country = GetFlagStr(EndIPOff + 4)
			If CountryFlag = 1 Then
				LocalStr = GetFlagStr(Stream.Position)
				' 以下用来获取数据库版本信息
				If IP >= IPToInt("255.255.255.0") And IP <= IPToInt("255.255.255.255") Then
					LocalStr = GetFlagStr(EndIPOff + 21)
					Country = GetFlagStr(EndIPOff + 12)
				End If
			Else
				LocalStr = GetFlagStr(EndIPOff + 8)
			End If
		Else
			Country = GetFlagStr(EndIPOff + 4)
			LocalStr = GetFlagStr(Stream.Position)
		End If
		' 过滤数据库中的无用信息
		Country = Trim(Country)
		LocalStr = Trim(LocalStr)
		If InStr(Country, "CZ88.NET") Then Country = "GZ110.CN"
		If InStr(LocalStr, "CZ88.NET") Then LocalStr = "GZ110.CN"
	End Sub
	' ============================================
	' 获取IP地址标识符
	' ============================================
	Private Function GetFlagStr(OffSet)
		Dim Flag
		Flag = 0
		Do While (True)
			Stream.Position = OffSet
			Flag = AscB(Stream.Read(1))
			If(Flag = 1 Or Flag = 2 ) Then
				Buf = Stream.Read(3) 
				If (Flag = 2 ) Then
					CountryFlag = 2
					EndIPOff = OffSet - 4
				End If
				OffSet = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256)
			Else
				Exit Do
			End If
		Loop
		
		If (OffSet < 12 ) Then
			GetFlagStr = ""
		Else
			Stream.Position = OffSet
			GetFlagStr = GetStr() 
		End If
	End Function
	' ============================================
	' 获取字串信息
	' ============================================
	Private Function GetStr() 
		Dim c
		GetStr = ""
		Do While (True)
			c = AscB(Stream.Read(1))
			If (c = 0) Then Exit Do 
			
			'如果是双字节，就进行高字节在结合低字节合成一个字符
			If c > 127 Then
				If Stream.EOS Then Exit Do
				GetStr = GetStr & Chr(AscW(ChrB(AscB(Stream.Read(1))) & ChrB(C)))
			Else
				GetStr = GetStr & Chr(c)
			End If
		Loop 
	End Function
	' ============================================
	' 核心函数，执行IP搜索
	' ============================================
	Public Function QQWry(DotIP)
		Dim IP, nRet
		Dim RangB, RangE, RecNo
		
		IP = IPToInt (DotIP)
		
		Set Stream = CreateObject("ADodb.Stream")
		Stream.Mode = 3
		Stream.Type = 1
		Stream.Open
		Stream.LoadFromFile QQWryFile
		Stream.Position = 0
		Buf = Stream.Read(8)
		
		FirstStartIP = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256) + (AscB(MidB(Buf, 4, 1))*256*256*256)
		LastStartIP  = AscB(MidB(Buf, 5, 1)) + (AscB(MidB(Buf, 6, 1))*256) + (AscB(MidB(Buf, 7, 1))*256*256) + (AscB(MidB(Buf, 8, 1))*256*256*256)
		RecordCount = Int((LastStartIP - FirstStartIP)/7)
		' 在数据库中找不到任何IP地址
		If (RecordCount <= 1) Then
			Country = "未知"
			QQWry = 2
			Exit Function
		End If
		
		RangB = 0
		RangE = RecordCount
		
		Do While (RangB < (RangE - 1)) 
			RecNo = Int((RangB + RangE)/2) 
			Call GetStartIP (RecNo)
			If (IP = StartIP) Then
				RangB = RecNo
				Exit Do
			End If
			If (IP > StartIP) Then
				RangB = RecNo
			Else 
				RangE = RecNo
			End If
		Loop
		
		Call GetStartIP(RangB)
		Call GetEndIP()

		If (StartIP <= IP) And ( EndIP >= IP) Then
			' 没有找到
			nRet = 0
		Else
			' 正常
			nRet = 3
		End If
		Call GetCountry(IP)

		QQWry = nRet
	End Function
	' ============================================
	' 检查IP地址合法性
	' ============================================
	Public Function IsIp(IP)
		IsIp = True
		If IP = "" Then IsIp = False : Exit Function
		Dim Re
		Set Re = New RegExp
		Re.Pattern = "^(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])\.(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])\.(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])\.(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])$"
		Re.IgnoreCase = True
		Re.Global = True
		IsIp = Re.Test(IP)
		Set Re = Nothing
	End Function
	' ============================================
	' 类终结

	' ============================================
	Private Sub Class_Terminate
		On ErrOr Resume Next
		Stream.Close
		If Err Then Err.Clear
		Set Stream = Nothing
	End Sub
End Class 

%>