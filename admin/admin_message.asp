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
Dim Action
If Not ChkAdmin("SendMessage") Then
	Server.Transfer("showerr.asp")
	Response.End
End If

Action = LCase(Request("action"))
Select Case Trim(Action)
Case "save"
	Call SaveNewMessage
Case "del"
	Call DeleteMessage
Case "delall"
	Call DeleteAllMessage
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
	Dim RsObj
	Response.Write "<script language=JavaScript>" & vbNewLine
	Response.Write "var _maxCount = '64000';" & vbNewLine
	Response.Write "function doSubmit(){" & vbNewLine
	Response.Write "	if (document.myform.topic.value==''){" & vbNewLine
	Response.Write "		alert('���ű��ⲻ��Ϊ�գ�');" & vbNewLine
	Response.Write "		return false;" & vbNewLine
	Response.Write "	}" & vbNewLine
	Response.Write "	myform.content1.value = Composition.document.body.innerHTML; " & vbNewLine
	Response.Write "	MessageLength = Composition.document.body.innerHTML.length;" & vbNewLine
	Response.Write "	if(MessageLength < 2){" & vbNewLine
	Response.Write "		alert('�������ݲ���С��2���ַ���');" & vbNewLine
	Response.Write "		return false;" & vbNewLine
	Response.Write "	}" & vbNewLine
	Response.Write "	if(MessageLength > _maxCount){" & vbNewLine
	Response.Write "		alert('���ŵ����ݲ��ܳ���'+_maxCount+'���ַ���');" & vbNewLine
	Response.Write "		return false;" & vbNewLine
	Response.Write "	}" & vbNewLine
	Response.Write "	document.myform.Submit1.disabled = true;" & vbNewLine
	Response.Write "	document.myform.submit();" & vbNewLine
	Response.Write "}" & vbNewLine
	Response.Write "</script>" & vbNewLine
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=2> >>�û����Ź���<< </th>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=TableRow1 colspan=2>�����û����ţ�<b><font color=red>" & AllUsersmsnum & "</font></b> �� &nbsp;&nbsp;�����û����ţ�<b><font color=red>" & DayUsersmsnum & "</font></b> �� &nbsp;&nbsp;<a href=""?action=del"" onclick=""return confirm('��ȷ��Ҫɾ�������û�������?')"" class=showmeun>ɾ�������û�����</a></td>"
	Response.Write "	</tr>"
	Response.Write "	<form name=form1 method=post action='admin_message.asp?action=del'>"
	Response.Write "	<tr>"
	Response.Write "		<td class=TableRow2 colspan=2>&nbsp;&nbsp;<b>����ɾ��ĳ�û��Ķ��ţ�</b>"
	Response.Write "<input type=text name=username size=30> &nbsp;<input type=submit value="" �� �� "" class=button onclick=""return confirm('��ȷ��Ҫɾ�����û�������?')"">"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "	<form name=form2 method=post action='admin_message.asp?action=delall'>"
	Response.Write "	<tr>"
	Response.Write "		<td class=TableRow2 colspan=2><b>����ɾ��ָ�������ڶ��ţ�</b>"
	Response.Write "		<select name=delDate size=1>"
	Response.Write "			<option value=7>һ������ǰ</option>"
	Response.Write "			<option value=30>һ����ǰ</option>"
	Response.Write "			<option value=60>������ǰ</option>"
	Response.Write "			<option value=180>����ǰ</option>"
	Response.Write "			<option value=""all"">���ж���</option>"
	Response.Write "		</select>"
	Response.Write "		&nbsp;<input type=checkbox name=isread value='yes'>����δ����Ϣ"
	Response.Write "		&nbsp;<input type=submit name=Submit value="" �� �� "" class=button onclick=""return confirm('��ȷ��Ҫɾ���˶�����?')"">"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=2> >>����Ⱥ��<< </th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action='admin_message.asp?action=save'>"
	Response.Write "	<tr>"
	Response.Write "		<td class=TableRow1 align=right><b>���ű���:</b></td>"
	Response.Write "		<td class=TableRow1><input type=text name=topic maxlength=70 size=70 value=''></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=TableRow2 align=right><b>�ռ���:</b></td>"
	Response.Write "		<td class=TableRow2>"
	Response.Write "		<select name=UserGroup size='1'>"
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup order by Groupid")
	Do While Not RsObj.EOF
		Response.Write "	<option value=""" & RsObj("Grades") & """"
		If RsObj("Grades") = 0 Then Response.Write " selected"
		Response.Write ">"
		If RsObj("Grades") = 0 Then
			Response.Write "�����û�"
		Else
			Response.Write RsObj("GroupName")
		End If
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=TableRow1 align=right><b>��������:</b></td>"
	Response.Write "		<td class=TableRow1><textarea name='content1' id='content1' style='display:none'></textarea>"
	Response.Write "		<script Language=Javascript src=""../editor/editor1.js""></script>"
	Response.Write "		<input type=radio name=isshow value=1 checked>��ʾ���͹��� <input type=radio name=isshow value=0 > ����ʾ���͹��̣��ٶȽϿ죩</td>" & vbNewLine
	Response.Write "	</tr>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=TableRow2 colspan=2><input type=""button"" name=""Submit1"" value="" ���Ͷ��� "" onclick=""doSubmit();"" class=button></td>"
	Response.Write "	</tr><form>"
	Response.Write "</table>"
End Sub
Sub SaveNewMessage()
	Dim strTopic,strContent,sender
	Dim UserGrade,isshow,i,smsnum,userlist
	isshow = CInt(Request("isshow"))
	If Trim(Request("topic")) = "" Then
		ErrMsg = "<li>���ű��ⲻ��Ϊ�ա�</li>"
		FoundErr = True
		Exit Sub
	Else
		strTopic = enchiasp.CheckStr(Request("topic"))
	End If
	If Trim(Request("content1")) = "" Then
		ErrMsg = "<li>�������ݲ���Ϊ�ա�</li>"
		FoundErr = True
		Exit Sub
	Else
		strContent = enchiasp.CheckStr(Request("content1"))
	End If
	sender = enchiasp.SiteName
	UserGrade = CInt(Request("UserGroup"))
	If CInt(Request("UserGroup")) = 0 Then
		SQL = "INSERT INTO ECCMS_Message (sender,incept,title,content,flag,SendTime,isRead,delSend) VALUES ('"& sender &"','�����û�','"& strTopic &"','"& strContent &"',1,"& NowString &",0,0) "
		enchiasp.Execute(SQL)
		enchiasp.Execute ("UPDATE ECCMS_User SET usermsg=usermsg+1")
		Succeed("<li>���������û����ųɹ���</li>")
		Exit Sub
	Else
		SQL = "SELECT COUNT(userid) FROM [ECCMS_User] WHERE UserGrade="& UserGrade
		Set Rs = enchiasp.Execute(SQL)
		smsnum = Rs(0)
		Rs.Close
		SQL = "SELECT username FROM [ECCMS_User] WHERE UserGrade="& UserGrade &" ORDER BY userid DESC"
	End If
	Response.Write "<br><table width='400' align=center border=0 cellspacing=1 cellpadding=1>" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write "<td>���濪ʼ���Ͷ���Ϣ��Ԥ�Ʊ��η���" & smsnum & "���û���</td></tr>" & vbCrLf
	Response.Write "<tr> " & vbCrLf
	Response.Write "<td style=""border: 1px #384780 solid ;background-color: #FFFFFF;"">" & vbCrLf
	Response.Write "<table width=0 id=table2 name=table2 border=0 cellspacing=0 cellpadding=0>" & vbCrLf
	Response.Write "<tr height=12><td bgcolor=#36D91A>" & vbCrLf
	Response.Write "</td></tr></table></td></tr><tr> " & vbCrLf
	Response.Write "<td> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span></td></tr>" & vbCrLf
	Response.Write "</table><br>" & vbCrLf
	Response.Flush
	Set Rs = enchiasp.Execute(SQL)
	If Not (Rs.EOF And Rs.BOF) Then
		userlist=Rs.GetRows(-1)
		Set Rs = Nothing
		For i=0 to UBound(userlist,2)
			userlist(0,i)=enchiasp.CheckStr(userlist(0,i))
			If Response.IsClientConnected Then
				If isshow = 1 Then
					Response.Write "<script>" & vbCrLf
					Response.Write "table2.style.width=" & Fix((i / smsnum) * 400) & ";" & vbCrLf
					Response.Write "txt2.innerHTML=""" & FormatNumber(i / smsnum * 100, 2, -1) & "�����Ͷ��Ÿ�" & userlist(0,i) & "�ɹ���"";" & vbCrLf
					Response.Write "</script>" & vbCrLf
					Response.Flush
				End If
				SQL = "INSERT INTO ECCMS_Message (sender,incept,title,content,flag,SendTime,isRead,delSend) VALUES ('"& sender &"','"& userlist(0,i) &"','"& strTopic &"','"& strContent &"',0,"& NowString &",0,0) "
				enchiasp.Execute(SQL)
				enchiasp.Execute ("UPDATE ECCMS_User SET usermsg=usermsg+1 WHERE username='"& userlist(0,i) &"'")
			End If
		Next
		Response.Write "<script>table2.style.width=400;txt2.innerHTML=""100%���������..."";</script>"
		Response.Flush
	End If
	Succeed("<li>�����û�������ɣ����������������</li>")
End Sub

Sub DeleteMessage()
	If Trim(Request("username")) = "" Then
		ErrMsg = "<li>������Ҫ����ɾ�����û�����</li>"
		FoundErr = True
		Exit Sub
	End If
	SQL = "DELETE FROM ECCMS_Message WHERE Sender='" & enchiasp.CheckStr(Request("username")) & "'"
	enchiasp.Execute(SQL)
	enchiasp.Execute ("UPDATE ECCMS_User SET usermsg=0 WHERE username='"& enchiasp.CheckStr(Request("username")) &"'")
	Succeed("<li>ɾ���û���" & Request("username") & " �Ķ��ųɹ���</li>")
End Sub
Sub DeleteAllMessage()
	Dim selRead, summid,i
	If Request("isread") = "yes" Then
		selRead = " ORDER BY id"
	Else
		selRead = " And isRead = 1 ORDER BY id"
	End If
	Select Case Request("delDate")
	Case "all"
		If Request("isread") = "yes" Then
			enchiasp.Execute("DELETE FROM ECCMS_Message")
		Else
			enchiasp.Execute("DELETE FROM ECCMS_Message WHERE isRead > 0")
		End If
		Succeed("<li>ɾ�������û����ųɹ���</li>")
		Exit Sub
	Case 7
		If IsSqlDataBase = 1 Then
			SQL = "SELECT id FROM ECCMS_Message WHERE DATEDIFF(d, Sendtime, GetDate()) > 7 " & selRead
		Else
			SQL = "SELECT id FROM ECCMS_Message WHERE DATEDIFF('d', Sendtime, Now()) > 7 " & selRead
		End If
	Case 30
		If IsSqlDataBase = 1 Then
			SQL = "SELECT id FROM ECCMS_Message WHERE DATEDIFF(d, Sendtime, GetDate()) > 30 " & selRead
		Else
			SQL = "SELECT id FROM ECCMS_Message WHERE DATEDIFF('d', Sendtime, Now()) > 30 " & selRead
		End If
	Case 60
		If IsSqlDataBase = 1 Then
			SQL = "SELECT id FROM ECCMS_Message WHERE DATEDIFF(d, Sendtime, GetDate()) > 60 " & selRead
		Else
			SQL = "SELECT id FROM ECCMS_Message WHERE DATEDIFF('d', Sendtime, Now()) > 60 " & selRead
		End If
	Case 180
		If IsSqlDataBase = 1 Then
			SQL = "SELECT id FROM ECCMS_Message WHERE DATEDIFF(d, Sendtime, GetDate()) > 180 " & selRead
		Else
			SQL = "SELECT id FROM ECCMS_Message WHERE DATEDIFF('d', Sendtime, Now()) > 180 " & selRead
		End If
	End Select
	Set Rs = enchiasp.Execute(SQL)
	summid = 0
	If Not (Rs.EOF And Rs.BOF) Then
		SQL = Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		For i = 0 To Ubound(SQL,2)
			enchiasp.Execute("DELETE FROM ECCMS_Message WHERE id = " & SQL(0,i))
			summid = summid + 1
		Next
	End If
	Succeed("<li>��ɾ��" & summid & "���û����ųɹ����������Ĳ�����</li>")
	Exit Sub
End Sub
Function AllUsersmsnum()
	On Error Resume Next
	AllUsersmsnum = enchiasp.Execute("SELECT COUNT(id) FROM ECCMS_Message")(0)
End Function
Function DayUsersmsnum()
	On Error Resume Next
	If isSqlDataBase = 1 Then
		DayUsersmsnum = enchiasp.Execute("SELECT COUNT(id) FROM ECCMS_Message WHERE datediff(d,SendTime,GetDate())=0")(0)
	Else
		DayUsersmsnum = enchiasp.Execute("SELECT COUNT(id) FROM ECCMS_Message WHERE SendTime >= Date()")(0)
	End If
End Function
%>