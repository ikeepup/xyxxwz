<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
'=====================================================================
' ������ƣ�������վ����ϵͳ
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж�������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
Response.Write "	<tr>"
Response.Write "	  <th>�� �� �� ��</th>"
Response.Write "	</tr>"
Response.Write "	<tr><form method=Post name=myform action='' onSubmit='return JugeQuery(this);'>"
Response.Write "	<td class=TableRow1>������"
Response.Write "	  <input name=keyword type=text size=30>"
Response.Write "	  ������"
Response.Write "	  <select name='field'>"
Response.Write "		<option value='1' selected>��������</option>"
Response.Write "		<option value='2'>��������</option>"
Response.Write "		<option value='0'>��������</option>"
Response.Write "	  </select> <input type=submit name=Submit value='��ʼ��ѯ' class=Button><br>"
Response.Write "	  </td></form>"
Response.Write "	</tr></form>"
Response.Write "	<tr>"
Response.Write "	  <td colspan=2 class=TableRow2><strong>����ѡ�</strong> <a href='admin_book.asp'>��������</a> | "
Response.Write "	  <a href='?isAccept=1'>���������</a> | "
Response.Write "	  <a href='?isAccept=0'>δ�������</a>"
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
	
	maxperpage = 30		'--ÿҳ��ʾ�б���
	If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("field")) = 1 Then
			foundsql = "WHERE title like '%" & keyword & "%'"
		ElseIf CInt(Request("field")) = 2 Then
			foundsql = "WHERE username like '%" & keyword & "%'"
		Else
			foundsql = "WHERE title like '%" & keyword & "%' Or username like '%" & keyword & "%'"
		End If
		strName = "��ѯ���"
		strList = "&keyword=" & keyword
	Else
		If Request("isAccept") <> "" Then
			isAccept = enchiasp.ChkNumeric(Request("isAccept"))
			foundsql = "WHERE isAccept=" & isAccept
			strList = "&isAccept=" & isAccept
			If isAccept = 0 Then
				strName = "δ�������"
			Else
				strName = "���������"
			End If
		Else
			foundsql = vbNullString
			strName = "��������"
			strList = vbNullString
		End If
	End If
	CurrentPage = enchiasp.ChkNumeric(Request("page"))
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>ѡ��</th>"
	Response.Write "	  <th width='40%'>�� �� �� ��</th>"
	Response.Write "	  <th width='15%' nowrap>�� ��</th>"
	Response.Write "	  <th width='8%' nowrap>�� ��</th>"
	Response.Write "	  <th width='15%' nowrap>�� �� ʱ ��</th>"
	Response.Write "	  <th width='17%' nowrap>�� �� �� ��</th>"
	Response.Write "	</tr>"
	'��¼����
	totalrec = enchiasp.Execute("SELECT COUNT(guestid) FROM ECCMS_GuestBook " & foundsql & "")(0)
	Pcount = CLng(totalrec / maxperpage)  '�õ���ҳ��
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
		Response.Write "<tr><td align=center colspan=6 class=TableRow2>��û���ҵ��κ����ԣ�</td></tr>"
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
			Response.Write "	  <td " & strRowstyle & " title=""����˴��鿴����������Ϣ""><a href='../" & enchiasp.ChannelDir & "showreply.asp?guestid=" & Rs("guestid") & "' target='_blank'>"
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
				Response.Write "<a href=?action=Accept&isAccept=1&guestid="& Rs("guestid") &" onclick=""{if(confirm('ȷ��Ҫ��˸�������?')){return true;}return false;}"" title='����˴�ֱ�����'>"
				Response.Write "<font color='red'>�� ��</font>"
			Else
				Response.Write "<a href=?action=Accept&isAccept=0&guestid="& Rs("guestid") &" onclick=""{if(confirm('ȷ��Ҫȡ�������?')){return true;}return false;}"" title='���ȡ���������'>"
				Response.Write "<font color='blue'>�����</font>"
			End If
			Response.Write "</a> | "
			Response.Write "<a href='../" & enchiasp.ChannelDir & "edit.asp?guestid=" & Rs("guestid") & "' target='_blank'>�༭</a> | "
			Response.Write "<a href=?action=del&ChannelID="& ChannelID &"&guestid="& Rs("guestid") &" onclick=""{if(confirm('����ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����������?')){return true;}return false;}"">ɾ��</a>"
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
	  <input class=Button type="button" name="chkall" value="ȫѡ" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="��ѡ" onClick="ContraSel(this.form)">
	  <input class=Button type="submit" name="Submit2" value="ɾ ��" onclick="return confirm('����ɾ���󽫲��ָܻ�\n��ȷ��ִ�иò�����?');">
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
		ErrMsg = ErrMsg + "<li>ID��������</li>"
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
			ErrMsg = ErrMsg + "<li>ID��������</li>"
			Exit Sub
		End If
	Else
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ID��������</li>"
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
		ErrMsg = ErrMsg + "<li>ID��������</li>"
		Exit Sub
	End If
End Sub

%>









