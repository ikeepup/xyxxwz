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
Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
Response.Write "	<tr>"
Response.Write "	  <th>留 言 管 理</th>"
Response.Write "	</tr>"
Response.Write "	<tr><form method=Post name=myform action='' onSubmit='return JugeQuery(this);'>"
Response.Write "	<td class=TableRow1>搜索："
Response.Write "	  <input name=keyword type=text size=30>"
Response.Write "	  条件："
Response.Write "	  <select name='field'>"
Response.Write "		<option value='1' selected>留言主题</option>"
Response.Write "		<option value='2'>留言作者</option>"
Response.Write "		<option value='0'>不限条件</option>"
Response.Write "	  </select> <input type=submit name=Submit value='开始查询' class=Button><br>"
Response.Write "	  </td></form>"
Response.Write "	</tr></form>"
Response.Write "	<tr>"
Response.Write "	  <td colspan=2 class=TableRow2><strong>操作选项：</strong> <a href='admin_book.asp'>所有留言</a> | "
Response.Write "	  <a href='?isAccept=1'>已审核留言</a> | "
Response.Write "	  <a href='?isAccept=0'>未审核留言</a>"
Response.Write "	  </td>"
Response.Write "	</tr>"
Response.Write "</table>"
Response.Write "<br>"
Dim Action,isAccept,i,guestid,replyid

ChannelID = 4
enchiasp.ReadChannel(ChannelID)
Action = LCase(Request("action"))
If Not ChkAdmin("GuestBook") Then
	Server.Transfer("showerr.asp")
	Response.End
End If

Select Case Trim(Action)
Case "del"
	Call DelGuestBook
Case "rdel"
	Call DelGuestReply
Case "accept"
	Call AcceptGuestBook
Case Else
	Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Sub showmain()
	Dim keyword,findword,foundsql,j
	Dim maxperpage,CurrentPage,Pcount,totalrec,totalnumber
	Dim strList,strName,strRowstyle
	
	maxperpage = 30		'--每页显示列表数
	If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("field")) = 1 Then
			foundsql = "WHERE title like '%" & keyword & "%'"
		ElseIf CInt(Request("field")) = 2 Then
			foundsql = "WHERE username like '%" & keyword & "%'"
		Else
			foundsql = "WHERE title like '%" & keyword & "%' Or username like '%" & keyword & "%'"
		End If
		strName = "查询结果"
		strList = "&keyword=" & keyword
	Else
		If Request("isAccept") <> "" Then
			isAccept = enchiasp.ChkNumeric(Request("isAccept"))
			foundsql = "WHERE isAccept=" & isAccept
			strList = "&isAccept=" & isAccept
			If isAccept = 0 Then
				strName = "未审核留言"
			Else
				strName = "已审核留言"
			End If
		Else
			foundsql = vbNullString
			strName = "所有留言"
			strList = vbNullString
		End If
	End If
	CurrentPage = enchiasp.ChkNumeric(Request("page"))
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>选择</th>"
	Response.Write "	  <th width='40%'>留 言 主 题</th>"
	Response.Write "	  <th width='15%' nowrap>作 者</th>"
	Response.Write "	  <th width='8%' nowrap>回 复</th>"
	Response.Write "	  <th width='15%' nowrap>更 新 时 间</th>"
	Response.Write "	  <th width='17%' nowrap>管 理 操 作</th>"
	Response.Write "	</tr>"
	'记录总数
	totalrec = enchiasp.Execute("SELECT COUNT(guestid) FROM ECCMS_GuestBook " & foundsql & "")(0)
	Pcount = CLng(totalrec / maxperpage)  '得到总页数
	If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_GuestBook " & foundsql & " ORDER BY isTop DESC,lastime DESC,guestid DESC"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = enchiasp.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=6 class=TableRow2>还没有找到任何留言！</td></tr>"
	Else
		Response.Write "	<tr>"
		Response.Write "	  <td colspan=""6"" class=""TableRow2"">"
		ShowListPage CurrentPage,Pcount,totalrec,maxperpage,strList,strName
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "<form name=selform method=post action="""">"
		Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
		Response.Write "<input type=hidden name=action value='del'>"
		i = 0
		If Pcount > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		j = totalrec - ((CurrentPage - 1) * maxperpage)
		Do While Not Rs.EOF And i < CLng(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			If (i mod 2) = 0 Then
				strRowstyle = "class=""TableRow1"""
			Else
				strRowstyle = "class=""TableRow2"""
			End If
			Response.Write "	<tr>"
			Response.Write "	  <td " & strRowstyle & " align=""center""><input type=checkbox name=guestid value=" & Rs("guestid") & "></td>"
			Response.Write "	  <td " & strRowstyle & " title=""点击此处查看所有留言信息""><a href='../" & enchiasp.ChannelDir & "showreply.asp?guestid=" & Rs("guestid") & "' target='_blank'>"
			Response.Write enchiasp.CheckTopic(Rs("title"))
			Response.Write "</a></td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & " align=""center"">"
			Response.Write enchiasp.CheckTopic(Rs("username"))
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & " align=""center"">"
			Response.Write Rs("ReplyNum")
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & " align=""center"" nowrap>"
			If Datediff("d",Rs("lastime"),Now()) = 0 Then
				Response.Write "<font color=""red"">" & Rs("lastime") & "</font>"
			Else
				Response.Write "<font color=""#808080"">" & Rs("lastime") & "</font>"
			End If
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & " align=""center"">"
			If Rs("isAccept") = 0 Then
				Response.Write "<a href=?action=Accept&isAccept=1&guestid="& Rs("guestid") &" onclick=""{if(confirm('确定要审核该留言吗?')){return true;}return false;}"" title='点击此处直接审核'>"
				Response.Write "<font color='red'>审 核</font>"
			Else
				Response.Write "<a href=?action=Accept&isAccept=0&guestid="& Rs("guestid") &" onclick=""{if(confirm('确定要取消审核吗?')){return true;}return false;}"" title='点击取消留言审核'>"
				Response.Write "<font color='blue'>已审核</font>"
			End If
			Response.Write "</a> | "
			Response.Write "<a href='../" & enchiasp.ChannelDir & "edit.asp?guestid=" & Rs("guestid") & "' target='_blank'>编辑</a> | "
			Response.Write "<a href=?action=del&ChannelID="& ChannelID &"&guestid="& Rs("guestid") &" onclick=""{if(confirm('留言删除后将不能恢复，您确定要删除该留言吗?')){return true;}return false;}"">删除</a>"
			Response.Write "</td>" & vbNewLine
			Rs.movenext
			i = i + 1
			j = j - 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
	<tr>
	  <td colspan="6" class="TableRow1">
	  <input class=Button type="button" name="chkall" value="全选" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="反选" onClick="ContraSel(this.form)">
	  <input class=Button type="submit" name="Submit2" value="删 除" onclick="return confirm('留言删除后将不能恢复\n您确定执行该操作吗?');">
	  </td>
	</tr>
	</form>
	<tr>
	  <td colspan="6" align="right" class="TableRow2"><%ShowListPage CurrentPage,Pcount,totalrec,maxperpage,strList,strName %></td>
	</tr>
</table>
<%
End Sub

Sub DelGuestBook()
	If Request("guestid") <> "" Then
		enchiasp.Execute("DELETE FROM ECCMS_GuestBook WHERE guestid in (" & Request("guestid") & ")")
		enchiasp.Execute("DELETE FROM ECCMS_GuestReply WHERE guestid in (" & Request("guestid") & ")")
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ID参数错误！</li>"
		Exit Sub
	End If
End Sub

Sub DelGuestReply()
	If enchiasp.ChkNumeric(Request("replyid")) > 0 Then
		replyid = CLng(Request("replyid"))
		guestid = CLng(Request("guestid"))
		If guestid > 0 Then
			enchiasp.Execute("DELETE FROM ECCMS_GuestReply WHERE id="& replyid)
			enchiasp.Execute ("UPDATE ECCMS_GuestBook SET ReplyNum=ReplyNum-1 WHERE guestid="& guestid)
			Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
		Else
			FoundErr = True
			ErrMsg = ErrMsg + "<li>ID参数错误！</li>"
			Exit Sub
		End If
	Else
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ID参数错误！</li>"
		Exit Sub
	End If
End Sub

Sub AcceptGuestBook()
	isAccept = enchiasp.ChkNumeric(Request("isAccept"))
	guestid = CLng(Request("guestid"))
	If guestid > 0 Then
		If isAccept = 0 Then
			enchiasp.Execute ("UPDATE ECCMS_GuestBook SET isAccept=0 WHERE guestid="& guestid)
		Else
			enchiasp.Execute ("UPDATE ECCMS_GuestBook SET isAccept=1 WHERE guestid="& guestid)
		End If
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ID参数错误！</li>"
		Exit Sub
	End If
End Sub

%>









