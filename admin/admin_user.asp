<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/chkinput.asp"-->
<!--#include file="../api/cls_api.asp"-->
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
Dim i,ii,RsObj
Dim keyword,findword,strClass,sUserGroup,foundsql
Dim seluserid,UserGrade,UserGroupStr,UserPassWord,username
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum,userlock
Action = LCase(Request("action"))

Select Case Trim(Action)
	Case "save"
		If Not ChkAdmin("AddUser") Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
		Call SaveUser
	Case "modify"
		Call ModifyUser
	Case "add"
		If Not ChkAdmin("AddUser") Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
		Call AddUser
	Case "edit"
		Call EditUser
	Case "del"
		Call BatDelUser
	Case Else
		Call MainPage
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Sub PageTop()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th colspan=2>��Ա����</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr><form method=Post name=myform action=admin_user.asp onSubmit='return JugeQuery(this);'>"
	Response.Write "	  <td class=TableRow1>������"
	Response.Write "	  <input name=keyword type=text size=20>"
	Response.Write "	  ������"
	Response.Write "	  <select name=queryopt>"
	Response.Write "		<option value=1 selected>��Ա����</option>"
	Response.Write "		<option value=2>��ʵ����</option>"
	Response.Write "		<option value=3>�û��ǳ�</option>"
	Response.Write "	  </select> <input type=submit name=Submit value='��ʼ����' class=Button></td>"
	Response.Write "	  <td class=TableRow1>"
	Response.Write "	  </td></form>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td colspan=2 class=TableRow2><strong>����ѡ�</strong> <a href='admin_user.asp'>��Ա������ҳ</a> | "
	Response.Write "	  <a href='admin_user.asp?action=add'><font color=blue>��ӻ�Ա</font></a> | "
	Response.Write "	  <a href='admin_user.asp?lock=1'><font color=blue>�ȴ���֤�Ļ�Ա</font></a> "
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup where Grades <> 0 order by Groupid")
	Do While Not RsObj.EOF
		Response.Write " | <a href=admin_user.asp?UserGrade="
		Response.Write RsObj("Grades")
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</a>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
	Response.Write "	  </td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub
Sub MainPage()
	Call PageTop
	If Not ChkAdmin("AdminUser") Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	If Not IsEmpty(Request("seluserid")) Then
		seluserid = Request("seluserid")
		Select Case enchiasp.CheckStr(Request("act"))
			Case "ɾ���û�"
				Call BatDelUser
			Case "�����û�"
				Call NolockUser
			Case "�����û�"
				Call IslockUser
			Case "ת���û�"
				Call MoveUser
			Case Else
				Response.Write "��Ч������"
		End Select
	End If
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th width='5%' nowrap>ѡ��</th>
	<th width='20%' nowrap>�û���</th>
	<th width='10%' nowrap>�û����֤</th>
	<th width='10%' nowrap>��Ա����</th>
	<th width='5%' nowrap>����</th>
	<th width='5%' nowrap>�Ա�</th>
	<th width='20%' nowrap>����ѡ��</th>
	<th width='15%' nowrap>����½ʱ��</th>
	<th width='5%' nowrap>��½����</th>
	<th width='5%' nowrap>״̬</th>
</tr>
<%
	If Trim(Request("UserGrade")) <> "" Then
		SQL = "SELECT GroupName,Grades FROM [ECCMS_UserGroup] WHERE Grades=" & Request("UserGrade")
		Set Rs = enchiasp.Execute(SQL)
		If Rs.Bof And Rs.EOF Then
			Response.Write "Sorry��û���ҵ��κλ�Ա��������ѡ���˴����ϵͳ����!"
			Response.End
		Else
			sUserGroup = Rs("GroupName")
			UserGrade = Rs("Grades")
		End If
		Rs.Close
	Else
		sUserGroup = "ȫ����Ա"
		UserGrade = 0
	End If
	maxperpage = 20 '###ÿҳ��ʾ��
	
	If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
		Response.Write ("�����ϵͳ����!����������")
		Response.End
	End If
	If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
		CurrentPage = CInt(Request("page"))
	Else
		CurrentPage = 1
	End If
	userlock =0
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("queryopt")) = 1 Then
			findword = "where username like '%" & keyword & "%'"
		ElseIf CInt(Request("queryopt")) = 2 Then
			findword = "where TrueName like '%" & keyword & "%'"
		Else
			findword = "where nickname like '%" & keyword & "%'"
		End If
		foundsql = findword
		sUserGroup = "��ѯ��Ա"
	Else
		If Trim(Request("UserGrade")) <> "" Then
			foundsql = "where UserGrade = " & Request("UserGrade")
			UserGrade = Request("UserGrade")
		Else
			If Trim(Request("lock")) <> "" Then
				foundsql = "where userlock =1"
				userlock =1
			Else
				foundsql = ""
			End If
		End If
	End If
	TotalNumber = enchiasp.Execute("SELECT COUNT(userid) FROM ECCMS_User "& foundsql &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT userid,username,nickname,UserGrade,UserGroup,UserClass,UserLock,userpoint,usermoney,TrueName,UserSex,usermail,HomePage,oicq,JoinTime,ExpireTime,LastTime,userlogin FROM [ECCMS_User] "& foundsql &" ORDER BY JoinTime DESC ,userid DESC"
	If IsSqlDataBase=1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = enchiasp.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=10 class=TableRow1>��û���ҵ��κλ�Ա��Ϣ��</td></tr>"
	Else
		Rs.MoveFirst
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0

		Response.Write "<tr>"
		Response.Write "	<td colspan=10 class=tablerow2>"
		Call showpage()
		Response.Write "</td>"
		Response.Write "	<form name=selform method=post action="""">"
		Response.Write "</tr>"

		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			If (i mod 2) = 0 Then
				strClass = "class=TableRow1"
			Else
				strClass = "class=TableRow2"
			End If
			Response.Write "<tr align=center>"
			Response.Write "	<td " & strClass & "><input type=checkbox name=seluserid value='" & Rs("userid") & "'></td>"
			Response.Write "	<td " & strClass & ">"
			Response.Write "<a href='?action=edit&userid=" & Rs("userid") & "' title='�û��ǳƣ�" & Rs("nickname") & "'>"
			If Rs("UserGrade") = 999 Then
				Response.Write "<span class=style2>"
			Else
				Response.Write "<span>"
			End If
			Response.Write Rs("username")
			Response.Write "</span></a>"
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			Response.Write Rs("UserGroup")
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			If Rs("UserGrade") = 999 Then
				Response.Write "����Ա"
			Else
				If Rs("UserClass") = 0 Then
					Response.Write "�Ƶ��Ա"
				ElseIf Rs("UserClass") = 1 Then
					Response.Write "��ʱ��Ա"
				Else
					Response.Write "��ʱ����"
				End If
			End If
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			Response.Write "<a href='admin_mailist.asp?action=mail&useremail="
			Response.Write Rs("usermail")
			Response.Write "'><img src='images/email.gif' border=0 alt='���û����ʼ�'></a>"
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			Response.Write Rs("UserSex")
			Response.Write "	</td>"
			Response.Write "	<td nowrap " & strClass & ">"
			Response.Write "<a href='?action=edit&userid=" & Rs("userid") & "'>�༭</a> | "
			Response.Write "<a href='?action=del&userid=" & Rs("userid") & "'>ɾ��</a>"
			Response.Write "	</td>"
			Response.Write "	<td nowrap " & strClass & ">"
			If Rs("LastTime") >= Date Then
				Response.Write "<font color=red>"
				Response.Write Rs("LastTime")
				Response.Write "</font>"
			Else
				Response.Write Rs("LastTime")
			End If
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			Response.Write Rs("userlogin")
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			If Rs("UserLock") = 0 Then
				Response.Write "<font color=blue>��</font>"
			Else
				Response.Write "<font color=red>��</font>"
			End If
			Response.Write "	</td>"
			Response.Write "</tr>"
			Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td colspan=10 class=tablerow1>
	<input class=Button type=button name=chkall value='ȫѡ' onClick=CheckAll(this.form)><input class=Button type=button name=chksel value='��ѡ' onClick=ContraSel(this.form)>&nbsp;&nbsp;����ѡ�&nbsp;
	 <input class=button onClick="{if(confirm('ȷ��ɾ��ѡ�����û���?')){this.document.form.submit();return true;}return false;}" type=submit value='ɾ���û�' name=act> 
	 <input class=button onClick="{if(confirm('ȷ������ѡ�����û���?')){this.document.form.submit();return true;}return false;}" type=submit value='�����û�' name=act> 
	 <input class=button onClick="{if(confirm('ȷ������ѡ�����û���?')){this.document.form.submit();return true;}return false;}" type=submit value='�����û�' name=act> 
	 <input class=button onClick="{if(confirm('ȷ��ת��ѡ�����û���?')){this.document.form.submit();return true;}return false;}" type=submit value='ת���û�' name=act> �� 
	 <select name='sUserGrade'>
	 <option value=''>����ѡ���û����</option>
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup where Grades <> 0 order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & "," & RsObj("GroupName") & """>"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>
	</select>
	</td>
</tr></form>
<tr>
	<td colspan=10 class=tablerow1><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Sub AddUser()
	Call PageTop
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan="2">��ӻ�Ա</th>
</tr>
<form name=myform method=post action=?action=save>
<tr>
	<td width='30%' align=right class=tablerow1><strong>��½���ƣ�</strong></td>
	<td width='70%' class=tablerow1><input type=text name=username size=20 value=''></td>
</tr>
<tr>
	<td align=right class=tablerow2><strong>�û����룺</strong></td>
	<td class=tablerow2><input type=password name=password1 size=20></td>
</tr>
<tr>
	<td align=right class=tablerow1><strong>ȷ�����룺</strong></td>
	<td class=tablerow1><input type=password name=password2 size=20></td>
</tr>
<tr>
	<td align=right class=tablerow2><strong>�û��ǳƣ�</strong></td>
	<td class=tablerow2><input type=text name=nickname size=20 value=''></td>
</tr>
<tr>
	<td align=right class=tablerow1><strong>�û����䣺</strong></td>
	<td class=tablerow1><input type=text name=usermail size=30 value='<%=enchiasp.MasterMail%>'></td>
</tr>
<tr>
	<td align=right class=tablerow2><strong>�û��ձ�</strong></td>
	<td class=tablerow2><select name='UserSex'>
		<option value='��'>˧��</option>
		<option value='Ů'>��Ů</option>
	</select></td>
</tr>
<tr>
	<td align=right class=tablerow1><strong>�����û��飺</strong></td>
	<td class=tablerow1><select name='UserGrade'>
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup where Grades <> 0 order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & "," & RsObj("GroupName") & """"
		If RsObj("Grades") = 1 Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>
	</select></td>
</tr>
<tr>
	<td align=right class=tablerow2><strong>�û�������</strong></td>
	<td class=tablerow2><input type=text name=userpoint size=10 value='50'></td>
</tr>
<tr align=center>
	<td colspan=2 class=tablerow1>
	<input type=button name=Submit2 onclick="javascript:history.go(-1)" value='������һҳ' class=Button>
	<input type=Submit name=Submit1 value='����û�' class=Button></td>
</tr>
</form>
</table>

<%
End Sub

Sub EditUser()
	Call PageTop
	Dim userid,username
	userid = enchiasp.ChkNumeric(Request("userid"))
	username = Replace(Request("username"), "'", "")
	If userid = 0 Then
		SQL = "SELECT * FROM ECCMS_user WHERE username='" & username & "'"
	Else
		SQL = "SELECT * FROM ECCMS_user WHERE userid=" & userid
	End If
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κλ�Ա��������ѡ���˴����ϵͳ����!</li>"
		Exit Sub
	End If
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan=4>�鿴/�޸Ļ�Ա����</th>
</tr>
<form name=myform method=post action=?action=modify>
<input type=hidden name=userid value='<%=Rs("userid")%>'>
<tr>
	<td width='10%' class=tablerow1>��Ա����</td>
	<td width='40%' class=tablerow1><input type=text name=username size=20 value='<%=Rs("username")%>' disabled></td>
	<td width='10%' class=tablerow1>��ʵ����</td>
	<td width='40%' class=tablerow1><input type=text name=TrueName size=20 value='<%=Rs("TrueName")%>'></td>
</tr>
<tr>
	<td class=tablerow2>�û�����</td>
	<td class=tablerow2><input type=password name=password size=20> <font color=blue>������޸�����������</font></td>
	<td class=tablerow2>�û�����</td>
	<td class=tablerow2><input type=text name=usermail size=30 value='<%=Rs("usermail")%>'></td>
</tr>
<tr>
	<td class=tablerow1>��������</td>
	<td class=tablerow1><input type=text name=BuyCode size=20> <font color=blue>������޸�����������</font></td>
	<td class=tablerow1>�û�״̬</td>
	<td class=tablerow1>
	<input type=radio name=UserLock value='0'<%If Rs("UserLock") = 0 Then Response.Write " checked"%>> ����&nbsp;&nbsp;
	<input type=radio name=UserLock value='1'<%If Rs("UserLock") <> 0 Then Response.Write " checked"%>> ����&nbsp;&nbsp;
	</td>
</tr>
<tr>
	<td class=tablerow2>�û��ȼ�</td>
	<td class=tablerow2><select name='UserGrade'>
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup where Grades <> 0 order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & "," & RsObj("GroupName") & """"
		If RsObj("Grades") = Rs("UserGrade") Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>
	</select></td>
	<td class=tablerow2>��Ա����</td>
	<td class=tablerow2><select name='UserClass'>
		<option value='0'<%If Rs("UserClass") = 0 Then Response.Write " selected"%>>�Ƶ��Ա</option>
		<option value='1'<%If Rs("UserClass") = 1 Then Response.Write " selected"%>>��ʱ��Ա</option>
		<option value='999'<%If Rs("UserClass") = 999 Then Response.Write " selected"%>>���ڻ�Ա</option>
	</select></td>
</tr>
<tr>
	<td class=tablerow1>�û�����</td>
	<td class=tablerow1><input type=text name=userpoint size=10 value='<%=Rs("userpoint")%>'></td>
	<td class=tablerow1>�˻����</td>
	<td class=tablerow1><input type=text name=usermoney size=10 value='<%=Rs("usermoney")%>'> Ԫ</td>
</tr>
<tr>
	<td class=tablerow2 nowrap>�û�����ֵ</td>
	<td class=tablerow2><input type=text name=experience size=10 value='<%=Rs("experience")%>'></td>
	<td class=tablerow2 nowrap>�û�����ֵ</td>
	<td class=tablerow2><input type=text name=charm size=10 value='<%=Rs("charm")%>'></td>
</tr>
<tr>
	<td class=tablerow1>���֤����</td>
	<td class=tablerow1><input type=text name=UserIDCard size=35 value='<%=Rs("UserIDCard")%>'></td>
	<td class=tablerow1>�ձ�</td>
	<td class=tablerow1><select name='UserSex'>
		<option value='��'<%If Rs("UserSex") = "��" Then Response.Write " selected"%>>˧��</option>
		<option value='Ů'<%If Rs("UserSex") = "Ů" Then Response.Write " selected"%>>��Ů</option>
	</select></td>
</tr>
<tr>
	<td class=tablerow2>�û��绰</td>
	<td class=tablerow2><input type=text name=phone size=20 value='<%=Rs("phone")%>'></td>
	<td class=tablerow2>�û�QQ</td>
	<td class=tablerow2><input type=text name=oicq size=20 value='<%=Rs("oicq")%>'></td>
</tr>
<tr>
	<td class=tablerow1>��������</td>
	<td class=tablerow1><input type=text name=postcode size=20 value='<%=Rs("postcode")%>'></td>
	<td class=tablerow1>��ϵ��ַ</td>
	<td class=tablerow1><input type=text name=address size=45 value='<%=Rs("address")%>'></td>
</tr>
<tr>
	<td class=tablerow2>��������</td>
	<td class=tablerow2><input type=text name=question size=20 value='<%=Rs("question")%>'></td>
	<td class=tablerow2>�����</td>
	<td class=tablerow2><input type=text name=answer size=20> <font color=blue>������޸�����������</font></td>
</tr>
<tr>
	<td class=tablerow1 nowrap>����½ʱ��</td>
	<td class=tablerow1><input type=text name=LastTime size=30 value='<%=Rs("LastTime")%>'></td>
	<td class=tablerow1>����½IP</td>
	<td class=tablerow1><input type=text name=userlastip size=20 value='<%=Rs("userlastip")%>'></td>
</tr>
<tr>
	<td class=tablerow2>ע��ʱ��</td>
	<td class=tablerow2><input type=text name=JoinTime size=30 value='<%=Rs("JoinTime")%>'></td>
	<td class=tablerow2>����ʱ��</td>
	<td class=tablerow2><input type=text name=ExpireTime size=30 value='<%=Rs("ExpireTime")%>'></td>
</tr>
<tr>
	<td class=tablerow1>�û�ͼ��</td>
	<td class=tablerow1><input type=text name=UserFace size=30 value='<%=Rs("UserFace")%>'></td>
	<td class=tablerow1>��½����</td>
	<td class=tablerow1><input type=text name=userlogin size=10 value='<%=Rs("userlogin")%>'></td>
</tr>
<tr>
	<td class=tablerow1>���뱣��</td>
	<td class=tablerow1>
	<input type=radio name=Protect value='0'<%If Rs("Protect") = 0 Then Response.Write " checked"%>> δ����&nbsp;&nbsp;
	<input type=radio name=Protect value='1'<%If Rs("Protect") <> 0 Then Response.Write " checked"%>> ������&nbsp;&nbsp;</td>
	<td class=tablerow1>�û��ǳ�</td>
	<td class=tablerow1><input type=text name=nickname size=20 value='<%=Rs("nickname")%>'></td>
</tr>
<tr align=center>
	<td colspan=4 class=tablerow2>
	<input type=button name=Submit2 onclick="javascript:history.go(-1)" value='������һҳ' class=Button>
	<input type=Submit name=Submit1 value='ȷ���޸�' class=Button></td>
</tr></form>
</table>

<%
End Sub

Sub CheckSave()
	If Trim(Request.Form("usermail")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�û����䲻��Ϊ�գ�</li>"
	End If
	If IsValidEmail(Trim(Request.Form("usermail"))) = False Then
		ErrMsg = ErrMsg + "<li>����Email�д���</li>"
		FoundErr = True
	End If
	If Not IsNumeric(Request.Form("userpoint")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�û��������������֣�</li>"
	End If
	If Trim(Request.Form("nickname")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�û��ǳƲ���Ϊ�գ�</li>"
	End If
	If enchiasp.IsValidStr(Request("nickname")) = False Then
		ErrMsg = ErrMsg + "<li>�û��ǳ��к��зǷ��ַ���</li>"
		Founderr = True
	End If
	UserGroupStr = Split(Request.Form("UserGrade"), ",")
End Sub

Sub SaveUser()
	CheckSave
	Dim Password,Question,Answer
	Dim usersex,sex
	If Trim(Request.Form("username")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�û�������Ϊ�գ�</li>"
	End If
	If enchiasp.IsValidStr(Request("username")) = False Then
		ErrMsg = ErrMsg + "<li>�û����к��зǷ��ַ���</li>"
		Founderr = True
	Else
		username = enchiasp.CheckBadstr(Request("username"))
	End If
	If Trim(Request.Form("password1")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�û����벻��Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("password2")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ȷ�����벻��Ϊ�գ�</li>"
	End If
	If Request.Form("password1") <> Request.Form("password2") Then
		ErrMsg = ErrMsg + "<li>������������ȷ�����벻һ�¡�</li>"
		FoundErr = True
	End If
	If enchiasp.IsValidPassword(Request("password2")) = False Then
		ErrMsg = ErrMsg + "<li>�����к��зǷ��ַ���</li>"
		Founderr = True
	Else
		Password = Trim(Request.Form("password2"))
		UserPassWord =  md5(Password)
	End If
	If Trim(Request.Form("usersex")) = "" Then
		ErrMsg = ErrMsg + "<li>�����ձ���Ϊ�գ�</li>"
		Founderr = True
	Else
		usersex = enchiasp.CheckBadstr(Request.Form("usersex"))
	End If
	If usersex = "Ů" Then
		sex = 0
	Else
		sex = 1
	End If
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_User WHERE username = '" & username & "'")
	If Not (Rs.bof And Rs.EOF) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry�����û��Ѿ�����,�뻻һ���û������ԣ�</li>"
		Exit Sub
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_Admin WHERE username='" & UserName & "'")
	If Not (Rs.BOF And Rs.EOF) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry�����û��Ѿ�����,�뻻һ���û������ԣ�</li>"
		Exit Sub
	End If
	Rs.Close:Set Rs = Nothing
	Question = Trim(Request.Form("question"))
	Answer = Trim(Request.Form("answer"))
	If Question = "" Then Question = enchiasp.GetRandomCode
	If Answer = "" Then Answer = enchiasp.GetRandomCode
	'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	Dim API_enchiasp,API_SaveCookie,SysKey
	If API_Enable Then
		Set API_enchiasp = New API_Conformity
		API_enchiasp.NodeValue "action","reguser",0,False
		API_enchiasp.NodeValue "username",UserName,1,False
		Md5OLD = 1
		SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
		Md5OLD = 0
		API_enchiasp.NodeValue "syskey",SysKey,0,False
		API_enchiasp.NodeValue "password",Password,0,False
		API_enchiasp.NodeValue "email",enchiasp.CheckStr(Request.Form("usermail")),1,False
		API_enchiasp.NodeValue "question",Question,1,False
		API_enchiasp.NodeValue "answer",Answer,1,False
		API_enchiasp.NodeValue "gender",sex,0,False
		API_enchiasp.SendHttpData
		If API_enchiasp.Status = "1" Then
			Founderr = True
			ErrMsg =  ErrMsg & API_enchiasp.Message
			Exit Sub
		End If
		Set API_enchiasp = Nothing
	End If
	'-----------------------------------------------------------------
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_User WHERE (userid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("username") = username
		Rs("password") = UserPassWord
		Rs("nickname") = Trim(Request.Form("nickname"))
		Rs("UserGrade") = CInt(UserGroupStr(0))
		Rs("UserGroup") = Trim(UserGroupStr(1))
		Rs("UserClass") = 0
		Rs("UserLock") = 0
		Rs("UserFace") = "face/1.gif"
		Rs("userpoint") = Trim(Request.Form("userpoint"))
		Rs("usermoney") = 0
		Rs("savemoney") = 0
		Rs("prepaid") = 0
		Rs("experience") = 10
		Rs("charm") = 10
		Rs("TrueName") = Trim(Request.Form("username"))
		Rs("usersex") = enchiasp.CheckStr(Request.Form("usersex"))
		Rs("usermail") = enchiasp.CheckStr(Request.Form("usermail"))
		Rs("oicq") = ""
		Rs("question") = Question
		Rs("answer") = md5(Answer)
		Rs("JoinTime") = Now()
		Rs("ExpireTime") = Now()
		Rs("LastTime") = Now()
		Rs("Protect") = 0
		Rs("usermsg") = 0
		Rs("userlastip") = ""
		Rs("userlogin") = 0
		Rs("usersetting") = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
	Rs.update
	Rs.Close:Set Rs = Nothing
	Call RemoveCache
	Succeed("<li>��ϲ������ӻ�Ա[<font color=blue>" & Request("username") & "</font>]�ɹ���</li>")
End Sub

Sub ModifyUser()
	CheckSave
	Dim sex
	If Trim(Request.Form("usersex")) = "Ů" Then
		sex = 0
	Else
		sex = 1
	End If
	If enchiasp.IsValidPassword(Request("password")) = False And Trim(Request("password")) <> "" Then
		ErrMsg = ErrMsg + "<li>�����к��зǷ��ַ���</li>"
		Founderr = True
	End If
	If enchiasp.IsValidPassword(Request("BuyCode")) = False And Trim(Request("BuyCode")) <> "" Then
		ErrMsg = ErrMsg + "<li>���������к��зǷ��ַ���</li>"
		Founderr = True
	End If
	If enchiasp.IsValidPassword(Request("answer")) = False And Trim(Request("answer")) <> "" Then
		ErrMsg = ErrMsg + "<li>������к��зǷ��ַ���</li>"
		Founderr = True
	End If
	If Not IsDate(Request.Form("JoinTime")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ע��ʱ���������</li>"
	End If
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_User WHERE userid = " & CLng(Request("userid"))
	Rs.Open SQL,Conn,1,3
		'Rs("username") = Trim(Request.Form("username"))
		Rs("nickname") = Trim(Request.Form("nickname"))
		If Trim(Request.Form("password")) <> "" Then Rs("password") = md5(Request.Form("password"))
		If Trim(Request.Form("BuyCode")) <> "" Then Rs("BuyCode") = md5(Request.Form("BuyCode"))
		Rs("UserGrade") = CInt(UserGroupStr(0))
		Rs("UserGroup") = Trim(UserGroupStr(1))
		Rs("UserClass") = Trim(Request.Form("UserClass"))
		Rs("UserLock") = Trim(Request.Form("UserLock"))
		Rs("UserFace") = Trim(Request.Form("UserFace"))
		Rs("userpoint") = Trim(Request.Form("userpoint"))
		Rs("usermoney") = Trim(Request.Form("usermoney"))
		Rs("experience") = Trim(Request.Form("experience"))
		Rs("charm") = Trim(Request.Form("charm"))
		Rs("TrueName") = Trim(Request.Form("TrueName"))
		Rs("UserIDCard") = Trim(Request.Form("UserIDCard"))
		Rs("usersex") = Trim(Request.Form("usersex"))
		Rs("usermail") = Trim(Request.Form("usermail"))
		Rs("phone") = Trim(Request.Form("phone"))
		Rs("oicq") = Trim(Request.Form("oicq"))
		Rs("postcode") = Trim(Request.Form("postcode"))
		Rs("address") = Trim(Request.Form("address"))
		Rs("question") = Trim(Request.Form("question"))
		If Trim(Request.Form("answer")) <> "" Then Rs("answer") = md5(Request.Form("answer"))
		Rs("Protect") = Trim(Request.Form("Protect"))
		Rs("JoinTime") = Trim(Request.Form("JoinTime"))
		Rs("ExpireTime") = Trim(Request.Form("ExpireTime"))
		Rs("LastTime") = Trim(Request.Form("LastTime"))
		Rs("userlastip") = Trim(Request.Form("userlastip"))
		Rs("userlogin") = Trim(Request.Form("userlogin"))
	Rs.update
	username = Rs("username")
	Rs.Close:Set Rs = Nothing
	If Founderr = False Then
		'-----------------------------------------------------------------
		'ϵͳ����
		'-----------------------------------------------------------------
		Dim API_enchiasp,API_SaveCookie,SysKey
		If API_Enable Then
			Set API_enchiasp = New API_Conformity
			API_enchiasp.NodeValue "action","update",0,False
			API_enchiasp.NodeValue "username",UserName,1,False
			Md5OLD = 1
			SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
			Md5OLD = 0
			API_enchiasp.NodeValue "syskey",SysKey,0,False
			API_enchiasp.NodeValue "password",Trim(Request.form("password")),1,False
			API_enchiasp.NodeValue "answer",Trim(Request.Form("answer")),1,False
			API_enchiasp.NodeValue "question",Trim(Request.Form("question")),1,False
			API_enchiasp.NodeValue "email",Trim(Request.Form("usermail")),1,False
			API_enchiasp.NodeValue "gender",sex,0,False
			API_enchiasp.SendHttpData
			If API_enchiasp.Status = "1" Then
				ErrMsg = API_enchiasp.Message
			End If
			Set API_enchiasp = Nothing
		End If
		'-----------------------------------------------------------------
	End If
	Call RemoveCache
	Succeed("<li>��ϲ�����޸Ļ�Ա[<font color=blue>" & username & "</font>]�����ϳɹ���</li>" & ErrMsg)
End Sub

Sub BatDelUser()
	Dim AllUserID,AllUserName
	If Trim(Request("userid")) <> "" Then
		seluserid = Request("userid")
	End If
	If Len(seluserid) = 0 Then seluserid = "0"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT userid,username FROM [ECCMS_User] WHERE userid in (" & seluserid & ")"
	Rs.Open SQL,Conn,1,1
	If Not (Rs.Bof And Rs.EOF) Then
		Do While Not Rs.EOF
			AllUserID = AllUserID & Rs(0) & ","
			AllUserName = AllUserName & Rs(1) & ","
			enchiasp.Execute("UPDATE ECCMS_Message SET delsend=1 WHERE sender='"& enchiasp.CheckStr(Rs(1)) &"'")
			enchiasp.Execute("DELETE FROM ECCMS_Message WHERE flag=0 And incept='"& enchiasp.CheckStr(Rs(1)) &"'")
		Rs.movenext
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	If AllUserID <> "" Then
		If Right(AllUserID,1) = "," Then AllUserID = Left(AllUserID,Len(AllUserID)-1)
		If Right(AllUserName,1) = "," Then AllUserName = Left(AllUserName,Len(AllUserName)-1)
		enchiasp.Execute ("DELETE FROM ECCMS_User WHERE userid in (" & AllUserID & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_Favorite WHERE userid in (" & AllUserID & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_Friend WHERE userid in (" & AllUserID & ")")
	
		'-----------------------------------------------------------------
		'ϵͳ����
		'-----------------------------------------------------------------
		Dim API_enchiasp,API_SaveCookie,SysKey
		If API_Enable Then
			Set API_enchiasp = New API_Conformity
			API_enchiasp.NodeValue "action","delete",0,False
			API_enchiasp.NodeValue "username",AllUserName,1,False
			Md5OLD = 1
			SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
			Md5OLD = 0
			API_enchiasp.NodeValue "syskey",SysKey,0,False
			API_enchiasp.SendHttpData
			Set API_enchiasp = Nothing
		End If
		'-----------------------------------------------------------------
		OutHintScript ("����ɾ�������ɹ���")
	End If
	Call RemoveCache
	'OutHintScript ("����ɾ�������ɹ���")
End Sub

Sub IslockUser()
	enchiasp.Execute ("UPDATE ECCMS_User SET UserLock=1 WHERE userid in (" & seluserid & ")")
	'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	Dim API_enchiasp,API_SaveCookie,SysKey
	If API_Enable Then
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT userid,username FROM [ECCMS_User] WHERE userid in (" & seluserid & ")"
		Rs.Open SQL,Conn,1,1
		If Not (Rs.Bof And Rs.EOF) Then
			Do While Not Rs.EOF
				UserName = Rs(1)
				Set API_enchiasp = New API_Conformity
				API_enchiasp.NodeValue "action","lock",0,False
				API_enchiasp.NodeValue "username",UserName,1,False
				Md5OLD = 1
				SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
				Md5OLD = 0
				API_enchiasp.NodeValue "syskey",SysKey,0,False
				API_enchiasp.NodeValue "userstatus",1,0,False
				API_enchiasp.SendHttpData
				Set API_enchiasp = Nothing
			Rs.movenext
			Loop
		End If
		Rs.Close:Set Rs = Nothing
	End If
	'-----------------------------------------------------------------
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Sub NolockUser()
	enchiasp.Execute ("UPDATE ECCMS_User SET UserLock=0 WHERE userid in (" & seluserid & ")")
	'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	Dim API_enchiasp,API_SaveCookie,SysKey
	If API_Enable Then
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT userid,username FROM [ECCMS_User] WHERE userid in (" & seluserid & ")"
		Rs.Open SQL,Conn,1,1
		If Not (Rs.Bof And Rs.EOF) Then
			Do While Not Rs.EOF
				UserName = Rs(1)
				Set API_enchiasp = New API_Conformity
				API_enchiasp.NodeValue "action","lock",0,False
				API_enchiasp.NodeValue "username",UserName,1,False
				Md5OLD = 1
				SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
				Md5OLD = 0
				API_enchiasp.NodeValue "syskey",SysKey,0,False
				API_enchiasp.NodeValue "userstatus",0,0,False
				API_enchiasp.SendHttpData
				Set API_enchiasp = Nothing
			Rs.movenext
			Loop
		End If
		Rs.Close:Set Rs = Nothing
	End If
	'-----------------------------------------------------------------
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Sub MoveUser()
	If Request("sUserGrade") = "" Then
		OutAlertScript("��ѡ����ȷ��ϵͳ������")
		Exit Sub
	End If
	UserGroupStr = Split(Request("sUserGrade"), ",")
	enchiasp.Execute ("update ECCMS_User set UserGrade=" & CInt(UserGroupStr(0)) & ", UserGroup='" & UserGroupStr(1) & "' where userid in (" & seluserid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Sub showpage()
	Dim n
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action=?UserGrade=" & Request("UserGrade") & "&lock=" & Request("lock") & "><tr><td align=center> " & vbCrLf
	Response.Write "<font color='red'>" & sUserGroup & "</font> "
	If CurrentPage < 2 Then
		Response.Write "���л�Ա <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> λ&nbsp;�� ҳ&nbsp;��һҳ&nbsp;|&nbsp;"
	Else
		Response.Write "���л�Ա <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> λ&nbsp;<a href=?page=1&UserGrade=" & Request("UserGrade") & "&lock=" & Request("lock") & ">�� ҳ</a>&nbsp;"
		Response.Write "<a href=?page=" & CurrentPage - 1 & "&UserGrade=" & Request("UserGrade") & "&lock=" & Request("lock") & ">��һҳ</a>&nbsp;|&nbsp;"
	End If
	If n - CurrentPage < 1 Then
		Response.Write "��һҳ&nbsp;β ҳ" & vbCrLf
	Else
		Response.Write "<a href=?page=" & (CurrentPage + 1) & "&UserGrade=" & Request("UserGrade") & "&lock=" & Request("lock") & ">��һҳ</a>"
		Response.Write "&nbsp;<a href=?page=" & n & "&UserGrade=" & Request("UserGrade") & "&lock=" & Request("lock") & ">β ҳ</a>" & vbCrLf
	End If
	Response.Write "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
	Response.Write "&nbsp;ת����"
	Response.Write "<input name=page size=3 value='" & CurrentPage & "'> <input class=Button type=submit name=Submit value='ת��'>"
	Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub
Sub RemoveCache()
	enchiasp.DelCahe "RenewStatistics"
	enchiasp.DelCahe "TotalStatistics"
End Sub
%>












