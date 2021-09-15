<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
'=====================================================================
' 软件名称：恩池网站管理系统--------下载服务器管理
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim Action, Flag, i, RsObj

ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID = 0 Then ChannelID = 2
Response.Write "<table width=""96%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">" & vbNewLine
Response.Write "<tr>" & vbNewLine
Response.Write "<th colspan=2>下载服务器管理" & vbNewLine
Response.Write "</th>" & vbNewLine
Response.Write "</tr>" & vbNewLine
Response.Write "<tr>" & vbNewLine
Response.Write "<td class=""TableRow1"" colspan=2>" & vbNewLine
Response.Write "<p><B>说明</B>：<BR>①、在这里您可以进行添加/删除下载服务器，请先添加服务器分类然后添加下载路径。<BR>" & vbNewLine & " "
Response.Write " ②、您可以添加多个下载路径，将会在下载信息页面显示。<BR>"
Response.Write " ③、建议添加后的服务器一级分类最好不要轻易删除，下载路径可以根据需要修改、删除或锁定。</p>"
Response.Write "</td>" & vbNewLine
Response.Write "</tr>" & vbNewLine
Response.Write "<tr>" & vbNewLine
Response.Write "<td class=""TableRow1"">" & vbNewLine
Response.Write "<B>操作选项</B></td>" & vbNewLine
Response.Write "<td class=""TableRow1""><a href=""admin_server.asp?ChannelID=" & ChannelID & """>服务器管理首页</a> | <a href=""admin_server.asp?action=add&amp;ChannelID=" & ChannelID & """>添加新的服务器</a>" & vbNewLine
Response.Write " | <a href=""admin_server.asp?action=serverorders&amp;ChannelID=" & ChannelID & """>服务器路径排序</a>" & vbNewLine
Response.Write "</td>" & vbNewLine
Response.Write "</tr>" & vbNewLine
Response.Write "</table>" & vbNewLine
Response.Write "<br>"


Flag = "DownServer" & ChannelID
Action = LCase(enchiasp.RemoveBadCharacters(Request("action")))
If Not ChkAdmin(Flag) Then
	Server.Transfer ("showerr.asp")
	Response.End
End If

Select Case Request("action")
	Case "add"
		Call sAdd
	Case "edit"
		Call sEdit
	Case "savenew"
		Call savenew
	Case "savedit"
		Call saveedit
	Case "del"
		Call DelDownPath
	Case "serverorders"
		Call serverorders
	Case "updateorders"
		Call updateorders
	Case "lock"
		Call isLock
	Case "free"
		Call FreeLock
	Case Else
		Call ShowMain
End Select
If FoundErr = True Then
	ReturnError (ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
'================================================
'过程名：ShowMain
'作  用：服务器管理首页
'================================================
Sub ShowMain()
	Response.Write " <table width=""96%"" class=""tableBorder"" cellspacing=""1"" cellpadding=""2"" align=center>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <th width=""35%""><strong>服务器分类</strong> </th>" & vbNewLine
	Response.Write " <th width=""35%""><strong>操 作</strong> </th>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	SQL = "SELECT * FROM ECCMS_DownServer WHERE ChannelID=" & ChannelID & " ORDER BY rootid,orders"
	Set Rs = CreateObject("ADODB.Recordset")
	Rs.Open SQL, Conn, 1, 1
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	Do While Not Rs.EOF
		Response.Write " <tr class=""TableRow1"">" & vbNewLine
		Response.Write " <td width=35% class=""TableRow1"">" & vbNewLine
		If Rs("isLock") = 1 Then
			Response.Write " <img src='images/locks.gif' border=0 align=absMiddle>"
		End If
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font>"
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;<font color=""#666666"">│</font>"
			Next
			Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font> "
		End If
		If Rs("parentid") = 0 Then Response.Write ("<b>[" & Rs("rootid") & "] ")
		Response.Write Rs("DownloadName")
		If Rs("child") > 0 Then Response.Write "(" & Rs("child") & ")"
		Response.Write " </td>" & vbNewLine
		Response.Write " <td width=65% class=""TableRow1"" align=right>"
		If Rs("depth") = 0 Then
			Response.Write "<a href=""admin_server.asp?action=add&editid="
			Response.Write Rs("downid")
			Response.Write "&amp;ChannelID=" & ChannelID & """>添加下载服务器路径</a>" & vbNewLine
		Else
			Response.Write "<a href=""admin_server.asp?action=lock&editid="
			Response.Write Rs("downid")
			Response.Write "&amp;ChannelID=" & ChannelID & """>锁定服务器</a>"
			Response.Write " | <a href=""admin_server.asp?action=free&editid="
			Response.Write Rs("downid")
			Response.Write "&amp;ChannelID=" & ChannelID & """>解除锁定</a>"
		End If
		Response.Write " | <a href=""admin_server.asp?action=edit&editid="
		Response.Write Rs("downid")
		Response.Write "&amp;ChannelID=" & ChannelID & """>服务器设置</a>" & vbNewLine
		Response.Write " |" & vbNewLine
		Response.Write " "
		If Rs("child") = 0 Then
			Response.Write " <a href=""admin_server.asp?action=del&editid="
			Response.Write Rs("downid")
			Response.Write "&amp;ChannelID=" & ChannelID & """ onclick=""{if(confirm('删除将包括该服务器的所有信息，确定删除吗?')){return true;}return false;}"">删除" & vbNewLine
			Response.Write " "
		Else
			Response.Write "<a href=""#"" onclick=""{if(confirm('该服务器含有下载路径，必须先删除其下载路径方能删除本服务器！')){return true;}return false;}"">" & vbNewLine
			Response.Write " 删除</a>" & vbNewLine
			Response.Write " "
		End If
		Response.Write " </td>" & vbNewLine
		Response.Write "</tr>" & vbNewLine
		Rs.MoveNext
		i = i + 1
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>" & vbNewLine
End Sub
'================================================
'过程名：sAdd
'作  用：添加服务器
'================================================
Sub sAdd()
	Dim ServerNum
	On Error Resume Next
	Set Rs = CreateObject("ADODB.Recordset")
	SQL = "SELECT MAX(downid) FROM ECCMS_DownServer"
	Rs.Open SQL, Conn, 1, 1
	If Rs.BOF And Rs.EOF Then
		ServerNum = 1
	Else
		ServerNum = Rs(0) + 1
	End If
	If IsNull(ServerNum) Then ServerNum = 1
	Rs.Close
	Response.Write "<form action =""admin_server.asp?action=savenew"" method=post>" & vbNewLine
	Response.Write "<input type=""hidden"" name=""newdownid"" value="""
	Response.Write ServerNum
	Response.Write """>" & vbNewLine
	Response.Write "<input type=""hidden"" name=ChannelID value="""
	Response.Write ChannelID
	Response.Write """>" & vbNewLine
	Response.Write " <table width=""96%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <th colspan=2>添加新的服务器</th>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write " <tr class=""TableRow1"">" & vbNewLine
	Response.Write " <td width=""30%"" class=""TableRow1"" height=30><U>服务器名称</U></td>" & vbNewLine
	Response.Write " <td width=""70%"" class=""TableRow1"">" & vbNewLine
	Response.Write " <input type=""text"" name=""DownloadName"" size=""60"">" & vbNewLine
	Response.Write "</td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=24 class=""TableRow1""><U>服务器路径</U></td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">" & vbNewLine
	Response.Write " <input type=""text"" name=""DownloadPath"" size=""60"">" & vbNewLine
	Response.Write "</td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=30 class=""TableRow1""><U>所属类别</U></td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">" & vbNewLine
	Response.Write " <select name=""servers"">" & vbNewLine
	Response.Write "<option value=""0"">做为服务器分类</option>" & vbNewLine
	SQL = "SELECT * FROM ECCMS_DownServer WHERE ChannelID=" & ChannelID & " And depth = 0 ORDER BY rootid"
	Rs.Open SQL, Conn, 1, 1
	Do While Not Rs.EOF
		Response.Write "<option value=""" & Rs("downid") & """ "
		If Len(Request("editid")) <> 0 And CLng(Request("editid")) = Rs("downid") Then Response.Write "selected"
		Response.Write ">"
		Response.Write Rs("DownloadName") & "</option>" & vbCrLf
		Rs.MoveNext
	Loop
	Rs.Close
	Response.Write "</select>"
	Response.Write "</td></tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=30 class=""TableRow1""><U>使用下载服务器的权限</U></td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">"
	Response.Write " <select name=""UserGroup"">" & vbNewLine
	Set RsObj = enchiasp.Execute("SELECT GroupName,Grades FROM ECCMS_UserGroup ORDER BY Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & """"
		If RsObj("Grades") = 0 Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.MoveNext
	Loop
	Set RsObj = Nothing
	Response.Write " </select> </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=24 class=""TableRow1""><U>下载所需点数</U></td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">" & vbNewLine
	Response.Write " <input type=""text"" name=""DownPoint"" size=""10"" onkeyup=if(isNaN(this.value))this.value='' value='0'>" & vbNewLine
	Response.Write "</td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=30 class=""TableRow1""><U>是否直接显示下载地址</U></td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">"
	Response.Write " <input type=radio name=isDisp value=""0"" checked> 否&nbsp;&nbsp;"
	Response.Write " <input type=radio name=isDisp value=""0""> 是"
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=24 class=""TableRow1"">&nbsp;</td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">" & vbNewLine
	Response.Write " <input type=""submit"" name=""Submit"" class=button value=""添加服务器"">" & vbNewLine
	Response.Write "</td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write "</table>" & vbNewLine
	Response.Write "</form>" & vbNewLine
	Set Rs = Nothing
End Sub
'================================================
'过程名：sEdit
'作  用：编辑服务器
'================================================
Sub sEdit()
	Dim Rs_e
	On Error Resume Next
	Set Rs = CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_DownServer WHERE downid=" & Request("editid")
	Set Rs_e = enchiasp.Execute(SQL)
	Response.Write "<form action =""admin_server.asp?action=savedit"" method=post>" & vbNewLine
	Response.Write "<input type=""hidden"" name=editid value="""
	Response.Write Request("editid")
	Response.Write """>" & vbNewLine
	Response.Write "<input type=""hidden"" name=ChannelID value="""
	Response.Write ChannelID
	Response.Write """>" & vbNewLine
	Response.Write " <table width=""96%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <th height=24 colspan=2>编辑服务器："
	Response.Write Rs_e("DownloadName")
	Response.Write "</th>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr class=""TableRow1"">" & vbNewLine
	Response.Write " <td width=""30%"" height=30 class=""TableRow1""><U>服务器名称</U></td>" & vbNewLine
	Response.Write " <td width=""70%"" class=""TableRow1"">" & vbNewLine
	Response.Write " <input type=""text"" name=""DownloadName"" size=""60"" value="""
	Response.Write Rs_e("DownloadName")
	Response.Write """>" & vbNewLine
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td class=""TableRow1""height=24><U>服务器路径</U><BR>" & vbNewLine
	Response.Write " 可以使用HTML代码</td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">" & vbNewLine
	Response.Write " <input type=""text"" name=""DownloadPath"" size=""60"" value="""
	Response.Write Rs_e("DownloadPath")
	Response.Write """>" & vbNewLine
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=30 class=""TableRow1""><U>所属类别</U></td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">" & vbNewLine
	Response.Write " <select name=""servers"">" & vbNewLine
	Response.Write " <option value=""0"">做为主服务器分类</option>" & vbNewLine
	Response.Write " "
	SQL = "SELECT * FROM ECCMS_DownServer WHERE ChannelID=" & ChannelID & " ORDER BY rootid,orders"
	Set Rs = enchiasp.Execute(SQL)
	Do While Not Rs.EOF
		Response.Write "<option value=""" & Rs("downid") & """ "
		If Rs_e("parentid") = Rs("downid") Then Response.Write "selected"
		Response.Write ">"
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├ "
		End If
		Response.Write Rs("DownloadName") & "</option>" & vbCrLf
		Rs.MoveNext
	Loop
	Rs.Close: Set Rs = Nothing
	Response.Write " </select> </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=30 class=""TableRow1""><U>使用下载服务器的权限</U></td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">"
	Response.Write " <select name=""UserGroup"">" & vbNewLine
	Set RsObj = enchiasp.Execute("SELECT GroupName,Grades FROM ECCMS_UserGroup ORDER BY Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & """"
		If Rs_e("UserGroup") = RsObj("Grades") Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.MoveNext
	Loop
	Set RsObj = Nothing
	Response.Write " </select> </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=24 class=""TableRow1""><U>下载所需点数</U></td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">" & vbNewLine
	Response.Write " <input type=""text"" name=""DownPoint"" size=""10"" onkeyup=if(isNaN(this.value))this.value='' value='"
	Response.Write Rs_e("DownPoint")
	Response.Write "'>" & vbNewLine
	Response.Write "</td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=30 class=""TableRow1""><U>是否直接显示下载地址</U></td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">"
	Response.Write " <input type=radio name=isDisp value=""0"""
	If Rs_e("IsDisp") = 0 Then Response.Write "  checked"
	Response.Write "> 否&nbsp;&nbsp;"
	Response.Write " <input type=radio name=isDisp value=""1"""
	If Rs_e("IsDisp") = 1 Then Response.Write "  checked"
	Response.Write "> 是"
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td class=""TableRow1""height=24>&nbsp;</td>" & vbNewLine
	Response.Write " <td class=""TableRow1"">" & vbNewLine
	Response.Write " <input type=""submit"" name=""Submit"" class=button value=""保存修改"">" & vbNewLine
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " </table>" & vbNewLine
	Response.Write "</form>" & vbNewLine
	Set Rs_e = Nothing
End Sub
'================================================
'过程名：savenew
'作  用：保存新的服务器
'================================================
Sub savenew()
	Dim downid,rootid,ParentID
	Dim depth,orders,Maxrootid
	Dim strParent,neworders
	Dim DownloadPath,Server_Url
	
	On Error Resume Next
	'保存添加服务器信息
	If Request.Form("DownloadName") = "" Then
		ErrMsg = ErrMsg + "<br>" + "<li>请输入服务器名称。"
		FoundErr = True
		Exit Sub
	End If
	If Request.Form("servers") = "" Then
		ErrMsg = ErrMsg + "<br>" + "<li>请选择服务器。"
		FoundErr = True
		Exit Sub
	End If
	If Request.Form("DownloadPath") = "" Then
		ErrMsg = ErrMsg + "<br>" + "<li>服务器路径不能为空。"
		FoundErr = True
		Exit Sub
	End If
	Server_Url = Replace(Request.Form("DownloadPath"), "\", "/")
	If Right(Server_Url, 1) <> "/" Then
		DownloadPath = Server_Url
	Else
		DownloadPath = Server_Url
	End If
	Set Rs = CreateObject("adodb.recordset")
	If Request.Form("servers") <> "0" Then
		SQL = "SELECT rootid,downid,depth,orders,strparent FROM ECCMS_DownServer WHERE downid=" & Request("servers")
		Rs.Open SQL, Conn, 1, 1
		rootid = Rs(0)
		ParentID = Rs(1)
		depth = Rs(2)
		orders = Rs(3)
		If depth + 1 > 2 Then
			ErrMsg = "<li>本系统限制最多只能有2级子服务器</li>"
			FoundErr = True
			Exit Sub
		End If
		strParent = Rs(4)
		Rs.Close
		neworders = orders
		SQL = "SELECT MAX(orders) FROM ECCMS_DownServer WHERE ParentID=" & Request("servers")
		Rs.Open SQL, Conn, 1, 1
		If Not (Rs.EOF And Rs.BOF) Then
			neworders = Rs(0)
		End If
		If IsNull(neworders) Then neworders = orders
		Rs.Close
		enchiasp.Execute ("UPDATE ECCMS_DownServer SET orders=orders+1 WHERE orders>" & CInt(neworders) & "")
	Else
		SQL = "SELECT MAX(rootid) FROM ECCMS_DownServer"
		Rs.Open SQL, Conn, 1, 1
		If Rs.BOF And Rs.EOF Then
			Maxrootid = 1
		Else
			Maxrootid = Rs(0) + 1
		End If
		If IsNull(Maxrootid) Then Maxrootid = 1
		Rs.Close
	End If
	If Maxrootid = 0 Then Maxrootid = 1
	
	SQL = "SELECT downid FROM ECCMS_DownServer WHERE downid=" & Request("newdownid")
	Rs.Open SQL, Conn, 1, 1
	If Not (Rs.EOF And Rs.BOF) Then
		ErrMsg = "<li>您不能指定和别的服务器一样的序号。</li>"
		FoundErr = True
		Exit Sub
	Else
		downid = CLng(Request("newdownid"))
	End If
	Rs.Close
	
	SQL = "SELECT * FROM ECCMS_DownServer"
	Rs.Open SQL, Conn, 1, 3
	Rs.AddNew
	If Request("servers") <> "0" Then
		Rs("depth") = depth + 1
		Rs("rootid") = rootid
		Rs("orders") = neworders + 1
		Rs("parentid") = Request.Form("servers")
		If strParent = "0" Then
			Rs("strparent") = Request.Form("servers")
		Else
			Rs("strparent") = strParent & "," & Request.Form("servers")
		End If
	Else
		Rs("depth") = 0
		Rs("rootid") = Maxrootid
		Rs("orders") = 0
		Rs("parentid") = 0
		Rs("strparent") = 0
	End If
	Rs("child") = 0
	Rs("downid") = Request.Form("newdownid")
	Rs("DownloadName") = Replace(enchiasp.ChkFormStr(Request.Form("DownloadName")), "|", "")
	Rs("DownloadPath") = Replace(DownloadPath, "|", "")
	Rs("isDisp") = Request.Form("isDisp")
	Rs("UserGroup") = Request.Form("UserGroup")
	Rs("ChannelID") = Request.Form("ChannelID")
	Rs("DownPoint") = CLng(Request.Form("DownPoint"))
	Rs("isLock") = 0
	Rs.Update
	Rs.Close
	If Request("servers") <> "0" Then
		If depth > 0 Then enchiasp.Execute ("update ECCMS_DownServer set child=child+1 where downid in (" & strParent & ")")
		enchiasp.Execute ("update ECCMS_DownServer set child=child+1 where downid=" & Request("servers"))
	End If
	SucMsg = "<li>服务器添加成功！</li>"
	Succeed (SucMsg)
	Set Rs = Nothing
End Sub
'================================================
'过程名：saveedit
'作  用：保存编辑
'================================================
Sub saveedit()
	Dim newdownid,Maxrootid,ParentID
	Dim depth,Child,strParent,rootid
	Dim iparentid,istrparent
	Dim trs,brs,mrs,k
	Dim nstrparent,mstrparent,ParentSql
	Dim boardcount,DownloadPath,Server_Url
	
	On Error Resume Next
	If CLng(Request("editid")) = CLng(Request("servers")) Then
		ErrMsg = "<li>所属服务器不能指定自己</li>"
		ReturnError (ErrMsg)
		Exit Sub
	End If
	Server_Url = Replace(Request.Form("DownloadPath"), "\", "/")
	If Right(Server_Url, 1) <> "/" Then
		DownloadPath = Server_Url
	Else
		DownloadPath = Server_Url
	End If
	Set Rs = CreateObject("adodb.recordset")
	SQL = "SELECT * FROM ECCMS_DownServer WHERE downid=" & CLng(Request("editid"))
	Rs.Open SQL, Conn, 1, 3
	newdownid = Rs("downid")
	ParentID = Rs("parentid")
	iparentid = Rs("parentid")
	strParent = Rs("strparent")
	depth = Rs("depth")
	Child = Rs("child")
	rootid = Rs("rootid")
	If ParentID = 0 Then
		If CLng(Request("servers")) <> 0 Then
			Set trs = enchiasp.Execute("select rootid from ECCMS_DownServer where downid=" & Request("servers"))
			If rootid = trs(0) Then
				ErrMsg = "<li>您不能指定该服务器的下属服务器作为所属服务器</li>"
				FoundErr = True
				Exit Sub
			End If
		End If
	Else
		Set trs = enchiasp.Execute("select downid from ECCMS_DownServer where strparent like '%" & strParent & "%' and downid=" & Request("servers"))
		If Not (trs.EOF And trs.BOF) Then
			ErrMsg = "<li>您不能指定该服务器的下属服务器作为所属服务器</li>"
			FoundErr = True
			Exit Sub
		End If
	End If
	If ParentID = 0 Then
		ParentID = Rs("downid")
		iparentid = 0
	End If
	Rs("DownloadName") = Replace(enchiasp.ChkFormStr(Request.Form("DownloadName")), "|", "")
	Rs("DownloadPath") = Replace(DownloadPath, "|", "")
	Rs("isDisp") = Request.Form("isDisp")
	Rs("UserGroup") = Request.Form("UserGroup")
	Rs("ChannelID") = Request.Form("ChannelID")
	Rs("DownPoint") = enchiasp.CheckNumeric(Request.Form("DownPoint"))
	Rs("isLock") = 0
	Rs.Update
	Rs.Close
	Set Rs = Nothing
	Set mrs = enchiasp.Execute("select max(rootid) from ECCMS_DownServer")
	Maxrootid = mrs(0) + 1
	If CLng(ParentID) <> CLng(Request("servers")) And Not (iparentid = 0 And CInt(Request("servers")) = 0) Then
		If iparentid > 0 And CInt(Request("servers")) = 0 Then
			enchiasp.Execute ("update ECCMS_DownServer set depth=0,orders=0,rootid=" & Maxrootid & ",parentid=0,strparent='0' where downid=" & newdownid)
			strParent = strParent & ","
			Set Rs = enchiasp.Execute("select count(*) from ECCMS_DownServer where strparent like '%" & strParent & "%'")
			boardcount = Rs(0)
			If IsNull(boardcount) Then
				boardcount = 1
			Else
				boardcount = boardcount + 1
			End If
			enchiasp.Execute ("update ECCMS_DownServer set child=child-" & boardcount & " where downid=" & iparentid)
			For i = 1 To depth
				Set Rs = enchiasp.Execute("select parentid from ECCMS_DownServer where downid=" & iparentid)
				If Not (Rs.EOF And Rs.BOF) Then
					iparentid = Rs(0)
					enchiasp.Execute ("update ECCMS_DownServer set child=child-" & boardcount & " where downid=" & iparentid)
				End If
			Next
			If Child > 0 Then
				i = 0
				Set Rs = enchiasp.Execute("select * from ECCMS_DownServer where strparent like '%" & strParent & "%'")
				Do While Not Rs.EOF
					i = i + 1
					mstrparent = Replace(Rs("strparent"), strParent, "")
					enchiasp.Execute ("update ECCMS_DownServer set depth=depth-" & depth & ",rootid=" & Maxrootid & ",strparent='" & mstrparent & "' where downid=" & Rs("downid"))
					Rs.MoveNext
				Loop
			End If
		ElseIf iparentid > 0 And CInt(Request("servers")) > 0 Then
			Set trs = enchiasp.Execute("select * from ECCMS_DownServer where downid=" & Request("servers"))
			strParent = strParent & ","
			Set Rs = enchiasp.Execute("select count(*) from ECCMS_DownServer where strparent like '%" & strParent & "%'")
			boardcount = Rs(0)
			If IsNull(boardcount) Then boardcount = 1
			enchiasp.Execute ("update ECCMS_DownServer set orders=orders + " & boardcount & " + 1 where rootid=" & trs("rootid") & " and orders>" & trs("orders") & "")
			enchiasp.Execute ("update ECCMS_DownServer set depth=" & trs("depth") & "+1,orders=" & trs("orders") & "+1,rootid=" & trs("rootid") & ",ParentID=" & Request("servers") & ",strparent='" & trs("strparent") & "," & trs("downid") & "' where downid=" & newdownid)
			i = 1
			SQL = "select * from ECCMS_DownServer where strparent like '%" & strParent & "%' order by orders"
			Set Rs = enchiasp.Execute(SQL)
			Do While Not Rs.EOF
				i = i + 1
				istrparent = trs("strparent") & "," & trs("downid") & "," & Replace(Rs("strparent"), strParent, "")
				enchiasp.Execute ("update ECCMS_DownServer set depth=depth+" & trs("depth") & "-" & depth & "+1,orders=" & trs("orders") & "+" & i & ",rootid=" & trs("rootid") & ",strparent='" & istrparent & "' where downid=" & Rs("downid"))
				Rs.MoveNext
			Loop
			ParentID = Request("servers")
			If rootid = trs("rootid") Then
				enchiasp.Execute ("update ECCMS_DownServer set child=child+" & i & " where (not ParentID=0) and downid=" & ParentID)
				For k = 1 To trs("depth")
					Set Rs = enchiasp.Execute("select parentid from ECCMS_DownServer where (not ParentID=0) and downid=" & ParentID)
					If Not (Rs.EOF And Rs.BOF) Then
						ParentID = Rs(0)
						enchiasp.Execute ("update ECCMS_DownServer set child=child+" & i & " where (not ParentID=0) and  downid=" & ParentID)
					End If
				Next
				enchiasp.Execute ("update ECCMS_DownServer set child=child-" & i & " where (not ParentID=0) and downid=" & iparentid)
				For k = 1 To depth
					Set Rs = enchiasp.Execute("select parentid from ECCMS_DownServer where (not ParentID=0) and downid=" & iparentid)
					If Not (Rs.EOF And Rs.BOF) Then
						iparentid = Rs(0)

						enchiasp.Execute ("update ECCMS_DownServer set child=child-" & i & " where (not ParentID=0) and  downid=" & iparentid)
					End If
				Next
			Else

				enchiasp.Execute ("update ECCMS_DownServer set child=child+" & i & " where downid=" & ParentID)
				For k = 1 To trs("depth")
					Set Rs = enchiasp.Execute("select parentid from ECCMS_DownServer where downid=" & ParentID)
					If Not (Rs.EOF And Rs.BOF) Then
						ParentID = Rs(0)
						enchiasp.Execute ("update ECCMS_DownServer set child=child+" & i & " where downid=" & ParentID)
					End If
				Next
				enchiasp.Execute ("update ECCMS_DownServer set child=child-" & i & " where downid=" & iparentid)
				For k = 1 To depth
					Set Rs = enchiasp.Execute("select parentid from ECCMS_DownServer where downid=" & iparentid)
					If Not (Rs.EOF And Rs.BOF) Then
						iparentid = Rs(0)
						enchiasp.Execute ("update ECCMS_DownServer set child=child-" & i & " where downid=" & iparentid)
					End If
				Next
			End If
		Else
			Set trs = enchiasp.Execute("select * from ECCMS_DownServer where downid=" & Request("servers"))
			Set Rs = enchiasp.Execute("select count(*) from ECCMS_DownServer where rootid=" & rootid)
			boardcount = Rs(0)
			ParentID = Request("servers")
			enchiasp.Execute ("update ECCMS_DownServer set child=child+" & boardcount & " where downid=" & ParentID)
			For k = 1 To trs("depth")
				Set Rs = enchiasp.Execute("select parentid from ECCMS_DownServer where downid=" & ParentID)
				If Not (Rs.EOF And Rs.BOF) Then
					ParentID = Rs(0)
					enchiasp.Execute ("update ECCMS_DownServer set child=child+" & boardcount & " where downid=" & ParentID)
				End If

			Next
			enchiasp.Execute ("update ECCMS_DownServer set orders=orders + " & boardcount & " + 1 where rootid=" & trs("rootid") & " and orders>" & trs("orders") & "")
			i = 0
			SQL = "select * from ECCMS_DownServer where rootid=" & rootid & " order by orders"
			Set Rs = enchiasp.Execute(SQL)
			Do While Not Rs.EOF
				i = i + 1
				If Rs("parentid") = 0 Then
					If trs("strparent") = "0" Then
						strParent = trs("downid")
					Else
						strParent = trs("strparent") & "," & trs("downid")
					End If
					enchiasp.Execute ("update ECCMS_DownServer set depth=depth+" & trs("depth") & "+1,orders=" & trs("orders") & "+" & i & ",rootid=" & trs("rootid") & ",strparent='" & strParent & "',parentid=" & Request("servers") & " where downid=" & Rs("downid"))
				Else
					If trs("strparent") = "0" Then
						strParent = trs("downid") & "," & Rs("strparent")
					Else
						strParent = trs("strparent") & "," & trs("downid") & "," & Rs("strparent")
					End If
					enchiasp.Execute ("update ECCMS_DownServer set depth=depth+" & trs("depth") & "+1,orders=" & trs("orders") & "+" & i & ",rootid=" & trs("rootid") & ",strparent='" & strParent & "' where downid=" & Rs("downid"))
				End If
				Rs.MoveNext
			Loop
		End If
	End If
	SucMsg = "<li>服务器修改成功！</li>"
	Succeed (SucMsg)
	Set Rs = Nothing
	Set mrs = Nothing
	Set trs = Nothing
End Sub
'================================================
'过程名：DelDownPath
'作  用：删除服务器
'================================================
Sub DelDownPath()
	Dim rsUsage
	
	On Error Resume Next
	Set Rs = enchiasp.Execute("select strparent,child,depth,rootid from ECCMS_DownServer where downid=" & Request("editid"))
	If Not (Rs.EOF And Rs.BOF) Then
		If Rs(1) > 0 Then
			ErrMsg = "该服务器含有下载路径，请删除其下载路径后再进行删除本服务器的操作"
			FoundErr = True
			Exit Sub
		End If
		If Rs("depth") = 0 Then
			Set rsUsage = enchiasp.Execute("SELECT downid FROM ECCMS_DownAddress WHERE downid=" & Rs("rootid"))
			If Not (rsUsage.EOF And rsUsage.BOF) Then
				ErrMsg = "该下载服务器正在使用中，不能删除!"
				FoundErr = True
				Exit Sub
			End If
			Set rsUsage = Nothing
		End If
		If Rs(2) > 0 Then
			enchiasp.Execute ("UPDATE ECCMS_DownServer SET child=child-1 WHERE downid in (" & Rs(0) & ")")
		End If
		SQL = "DELETE FROM ECCMS_DownServer WHERE downid=" & Request("editid")
		enchiasp.Execute (SQL)
	End If
	Set Rs = Nothing
	Succeed ("服务器删除成功！")
End Sub
'================================================
'过程名：isLock
'作  用：锁定服务器
'================================================
Sub isLock()

	enchiasp.Execute ("update ECCMS_DownServer set isLock = 1 where downid in (" & Request("editid") & ")")
	Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
'================================================
'过程名：FreeLock
'作  用：解除服务器锁定
'================================================
Sub FreeLock()
	enchiasp.Execute ("update ECCMS_DownServer set isLock = 0 where downid in (" & Request("editid") & ")")
	Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
'================================================
'过程名：serverorders
'作  用：服务器排序
'================================================
Sub serverorders()
	Dim trs
	Dim uporders
	Dim doorders
	
	Response.Write " <table width=""96%"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""tableBorder"" align=center>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <th colspan=2>服务器路径重新排序修改(请在相应服务器的排序表单内输入相应的排列序号)" & vbNewLine
	Response.Write " </th>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Set Rs = CreateObject("Adodb.recordset")
	SQL = "SELECT * FROM ECCMS_DownServer WHERE ChannelID=" & ChannelID & " ORDER BY RootID,orders"
	Rs.Open SQL, Conn, 1, 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "还没有相应的服务器。"
	Else
		Do While Not Rs.EOF
			Response.Write "<form action=admin_server.asp?action=updateorders method=post><tr><td width=""50%"" class=TableRow1>"
			If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font>"
			If Rs("depth") > 1 Then
				For i = 2 To Rs("depth")
					Response.Write "&nbsp;&nbsp;<font color=""#666666"">│</font>"
				Next
				Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font> "
			End If
			If Rs("parentid") = 0 Then Response.Write ("<b>")
			Response.Write Rs("DownloadName")
			If Rs("child") > 0 Then Response.Write "(" & Rs("child") & ")"
			Response.Write "</td><td width=""50%"" class=TableRow1>"
			If Rs("ParentID") > 0 Then
				Set trs = enchiasp.Execute("SELECT COUNT(*) FROM ECCMS_DownServer WHERE ParentID=" & Rs("ParentID") & " and orders<" & Rs("orders") & "")
				uporders = trs(0)
				If IsNull(uporders) Then uporders = 0
				Set trs = enchiasp.Execute("SELECT COUNT(*) FROM ECCMS_DownServer WHERE ParentID=" & Rs("ParentID") & " and orders>" & Rs("orders") & "")
				doorders = trs(0)
				If IsNull(doorders) Then doorders = 0
				If uporders > 0 Then
					Response.Write "<select name=uporders size=1><option value=0>↑</option>"
					For i = 1 To uporders
						Response.Write "<option value=" & i & ">↑" & i & "</option>"
					Next
					Response.Write "</select>"
				End If
				If doorders > 0 Then
					If uporders > 0 Then Response.Write "&nbsp;"
					Response.Write "<select name=doorders size=1><option value=0>↓</option>"
					For i = 1 To doorders
						Response.Write "<option value=" & i & ">↓" & i & "</option>"
					Next
					Response.Write "</select>"
				End If
				If doorders > 0 Or uporders > 0 Then
					Response.Write vbNewLine & "<input type=""hidden"" name=ChannelID value="""
					Response.Write ChannelID
					Response.Write """>" & vbNewLine
					Response.Write "<input type=hidden name=""editID"" value=""" & Rs("downid") & """>&nbsp;<input type=submit name=Submit class=button value='修 改'>"
				End If
			End If
			Response.Write "</td></tr></form>"
			uporders = 0
			doorders = 0
			Rs.MoveNext
		Loop
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>" & vbNewLine
End Sub
'================================================
'过程名：updateorders
'作  用：更新服务器排序
'================================================
Sub updateorders()
	Dim ParentID
	Dim orders
	Dim strParent
	Dim Child
	Dim uporders
	Dim doorders
	Dim oldorders
	Dim trs
	Dim ii
	If Not IsNumeric(Request("editID")) Then
		ReturnError ("非法的参数！")
		Exit Sub
	End If
	If Request("uporders") <> "" And Not CInt(Request("uporders")) = 0 Then
		If Not IsNumeric(Request("uporders")) Then
			ReturnError ("非法的参数！")
			Exit Sub
		ElseIf CInt(Request("uporders")) = 0 Then
			ReturnError ("请选择要提升的数字！")
			Exit Sub
		End If
		Set Rs = enchiasp.Execute("SELECT ParentID,orders,strparent,child FROM ECCMS_DownServer where downid=" & Request("editID"))
		ParentID = Rs(0)
		orders = Rs(1)
		strParent = Rs(2) & "," & Request("editID")
		Child = Rs(3)
		i = 0
		If Child > 0 Then
			Set Rs = enchiasp.Execute("SELECT COUNT(*) FROM ECCMS_DownServer WHERE strparent like '%" & strParent & "%'")
			oldorders = Rs(0)
		Else
			oldorders = 0
		End If
		Set Rs = enchiasp.Execute("SELECT downid,orders,child,strparent FROM ECCMS_DownServer WHERE ParentID=" & ParentID & " and orders<" & orders & " order by orders desc")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("uporders")) >= i Then
				If Rs(2) > 0 Then
					ii = 0
					Set trs = enchiasp.Execute("select downid,orders from ECCMS_DownServer where strparent like '%" & Rs(3) & "," & Rs(0) & "%' order by orders")
					If Not (trs.EOF And trs.BOF) Then
						Do While Not trs.EOF
							ii = ii + 1
							enchiasp.Execute ("update ECCMS_DownServer set orders=" & orders & "+" & oldorders & "+" & ii & " where downid=" & trs(0))
							trs.MoveNext
						Loop
					End If
				End If
				enchiasp.Execute ("update ECCMS_DownServer set orders=" & orders & "+" & oldorders & " where downid=" & Rs(0))
				If CInt(Request("uporders")) = i Then uporders = Rs(1)
			End If
			orders = Rs(1)
			Rs.MoveNext
		Loop
		enchiasp.Execute ("update ECCMS_DownServer set orders=" & uporders & " where downid=" & Request("editID"))
		If Child > 0 Then
			i = uporders
			Set Rs = enchiasp.Execute("select downid from ECCMS_DownServer where strparent like '%" & strParent & "%' order by orders")
			Do While Not Rs.EOF
				i = i + 1
				enchiasp.Execute ("update ECCMS_DownServer set orders=" & i & " where downid=" & Rs(0))
				Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
		Set trs = Nothing
	ElseIf Request("doorders") <> "" Then
		If Not IsNumeric(Request("doorders")) Then
			ReturnError ("非法的参数！")
			Exit Sub
		ElseIf CInt(Request("doorders")) = 0 Then
			ReturnError ("请选择要下降的数字！")
			Exit Sub
		End If
		Set Rs = enchiasp.Execute("select ParentID,orders,strparent,child from ECCMS_DownServer where downid=" & Request("editID"))
		ParentID = Rs(0)
		orders = Rs(1)
		strParent = Rs(2) & "," & Request("editID")
		Child = Rs(3)
		i = 0
		If Child > 0 Then
			Set Rs = enchiasp.Execute("select count(*) from ECCMS_DownServer where strparent like '%" & strParent & "%'")
			oldorders = Rs(0)
		Else
			oldorders = 0
		End If
		Set Rs = enchiasp.Execute("select downid,orders,child,strparent from ECCMS_DownServer where ParentID=" & ParentID & " and orders>" & orders & " order by orders")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("doorders")) >= i Then
				If Rs(2) > 0 Then
					ii = 0
					Set trs = enchiasp.Execute("select downid,orders from ECCMS_DownServer where strparent like '%" & Rs(3) & "," & Rs(0) & "%' order by orders")
					If Not (trs.EOF And trs.BOF) Then
						Do While Not trs.EOF
							ii = ii + 1
							enchiasp.Execute ("update ECCMS_DownServer set orders=" & orders & "+" & ii & " where downid=" & trs(0))
							trs.MoveNext
						Loop
					End If
				End If
				enchiasp.Execute ("update ECCMS_DownServer set orders=" & orders & " where downid=" & Rs(0))
				If CInt(Request("doorders")) = i Then doorders = Rs(1)
			End If
			orders = Rs(1)
			Rs.MoveNext
		Loop
		enchiasp.Execute ("UPDATE ECCMS_DownServer SET orders=" & doorders & " WHERE downid=" & Request("editID"))
		If Child > 0 Then
			i = doorders
			Set Rs = enchiasp.Execute("SELECT downid from ECCMS_DownServer WHERE strparent like '%" & strParent & "%' ORDER BY orders")
			Do While Not Rs.EOF
				i = i + 1
				enchiasp.Execute ("UPDATE ECCMS_DownServer SET orders=" & i & " WHERE downid=" & Rs(0))
				Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
		Set trs = Nothing
	End If
	Response.Redirect "admin_server.asp?action=serverorders&ChannelID=" & Request("ChannelID")
End Sub
%>