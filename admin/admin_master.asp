<!--#include file="setup.asp" -->
<!--#include file="check.asp"-->
<%
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
Dim selAdminID
Dim i,Action,strClass
Admin_header
If Not ChkAdmin("999") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Response.Write "<table cellpadding=2 cellspacing=1 border=0 class=tableBorder align=center>" & vbCrLf
Response.Write " <tr>" & vbCrLf
Response.Write " <th height=22 colspan=6>����Ա����</th>" & vbCrLf
Response.Write " </tr>" & vbCrLf
Response.Write " <tr>" & vbCrLf
Response.Write " <td class=TableRow1> <b>����ѡ�</b> <a href=admin_master.asp>������ҳ</a> &nbsp;<a href=admin_master.asp?action=add>��ӹ���Ա</a>"
Response.Write " </td>" & vbCrLf
Response.Write " </tr>" & vbCrLf
Response.Write " </table><br>" & vbCrLf
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "renew"
	Call UpdateFlag
Case "del"
	Call del
Case "pasword"
	Call pasword
Case "newpass"
	Call newpass
Case "add"
	Call addadmin
Case "edit"
	Call userinfo
Case "savenew"
	Call savenew
Case "active"
	Call ActiveLock
Case Else
	Call userlist
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
private function instr2(ByVal old,byval str)
	'Adminflag, "Add" & strModules & ChanID
	on error resume next
	dim temp
	dim i
	dim tempb 
	tempb=0
	if old<>"" then
		
		temp=split(old,",")
		for i=0 to ubound(temp)
		
		if str=temp(i) then
				tempb=1
				exit for
			end if
		next 
		instr2=tempb
	else
		instr2=0
	end if
	
	
	
End Function
Private Sub userlist()
	Response.Write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th height=22 colspan=6>����Ա����(����û������в���)</th>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr align=center>" & vbCrLf
	Response.Write "<td height=22 class=TableTitle><B>�û���</B></td>" & vbCrLf
	Response.Write "<td class=TableTitle><B>�ϴε�½ʱ��</B></td>" & vbCrLf
	Response.Write "<td class=TableTitle><B>�ϴε�½IP</B></td>" & vbCrLf
	Response.Write "<td class=TableTitle><B>����</B></td>" & vbCrLf
	Response.Write "<td class=TableTitle><B>״̬</B></td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Set Rs = enchiasp.Execute("select * from ECCMS_Admin order by Logintime desc")
	i = 0
	Do While Not Rs.EOF
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
		Response.Write " <tr>" & vbCrLf
		Response.Write " <td " & strClass & "><a href=""?id="
		Response.Write Rs("id")
		Response.Write "&action=pasword"" title='����˴��޸Ĺ���Ա��Ϣ'>"
		Response.Write Rs("username")
		Response.Write "</a></td>" & vbCrLf
		Response.Write "<td align=center " & strClass & ">"
		Response.Write Rs("Logintime")
		Response.Write "</td>" & vbCrLf
		Response.Write "<td align=center " & strClass & ">"
		Response.Write Rs("Loginip")
		Response.Write "</td>" & vbCrLf
		Response.Write "<td align=center " & strClass & "><a href=""?action=Active&id=" & Rs("id") & "&lock="
		If Rs("isLock") = 0 Then
			Response.Write "1"" onclick=""return confirm('��ȷ��Ҫ�����˹���Ա��?')"">��������Ա</a> | "
		Else
			Response.Write "0"" onclick=""return confirm('��ȷҪ����˹���Ա��?')"">�������Ա</a> | "
		End If
		Response.Write "<a href=""?action=del&id="
		Response.Write Rs("id")
		Response.Write "&name="
		Response.Write Rs("username")
		Response.Write """ onclick=""return confirm('�˲�����ɾ���ù���Ա\n ��ȷ��ִ�д˲�����?')"">ɾ��</a>&nbsp;|&nbsp;<a href=""?id="
		Response.Write Rs("id")
		Response.Write "&action=edit"">�༭Ȩ��</a>" & vbCrLf
		
		response.write "| "
		Response.Write "<a href='?action=pasword&id="
		Response.Write Rs("id")
		response.write "'>"
		response.write "��������� "
		response.write "</a>"
		
		response.write "</td>"
		Response.Write "<td align=center " & strClass & ">"
		If Rs("isLock") = 0 Then
			Response.Write "����"
		Else
			Response.Write "<font color=red>����<font>"
		End If
		Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf
		Rs.movenext
		i = i + 1
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td colspan=""6"" align=center Class=TableRow1>" & vbCrLf
	Response.Write " <input class=""button"" type=button name=""Submit"" value=""��ӹ���Ա"" onClick=""self.location='admin_master.asp?action=add'"" >" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " </table>" & vbCrLf
End Sub

Private Sub del()
	If Trim(Request("id")) <> "" Then
		enchiasp.Execute ("delete from ECCMS_Admin where username<>'" & AdminName & "' And id=" & Request("id"))
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		ErrMsg = "<li>�����ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
End Sub

Private Sub ActiveLock()
	If Trim(Request("lock")) <> "" And Trim(Request("id")) <> "" Then
		enchiasp.Execute ("update ECCMS_Admin set isLock="&Request("lock")&" where username<>'" & AdminName & "' And id=" & Request("id"))
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		ErrMsg = "<li>�����ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
End Sub


Private Sub pasword()
	Dim oldpassword
	If Not ChkAdmin("9999") Then
		ErrMsg = "<li>��û�д˲���Ȩ��!</li><li>����ʲô��������ϵվ����</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("select * from ECCMS_Admin where id=" & Request("id"))
	oldpassword = Rs("password")
	Response.Write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write "<form action=""?action=newpass"" method=post>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th colspan=2 height=23>����Ա���Ϲ����������޸�" & vbCrLf
	Response.Write " </th>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>��̨��½���ƣ�</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=hidden name=""oldusername"" value="""
	Response.Write Rs("username")
	Response.Write """>" & vbCrLf
	Response.Write " <input type=text size=25 name=""username2"" value="""
	Response.Write Rs("username")
	Response.Write """>" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>��̨��½���룺</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=""password"" size=25 name=""password2"">"
	Response.Write " (������޸�����������)" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>����Ա����</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=AdminGrade value='0' "
	If Rs("AdminGrade") = 0 Then Response.Write " checked"
	Response.Write " > ��ͨ����Ա&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=AdminGrade value='999' "
	If Rs("AdminGrade") = 999 Then Response.Write " checked"
	Response.Write " > �߼�����Ա ��ӵ�����Ȩ�ޣ�" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td align=""right"" class=tablerow1>�Ƿ񼤻����Ա��</td>" & vbCrLf
	Response.Write " <td class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isLock value='1' "
	If Rs("isLock") = 1 Then Response.Write " checked"
	Response.Write " > ��&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=isLock value='0' "
	If Rs("isLock") = 0 Then Response.Write " checked"
	Response.Write " > ��" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td align=""right"" class=tablerow1>����һ������Ա��½��</td>" & vbCrLf
	Response.Write " <td class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isAloneLogin value='0' "
	If Rs("isAloneLogin") = 0 Then Response.Write " checked"
	Response.Write " > ��&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=isAloneLogin value='1' "
	If Rs("isAloneLogin") = 1 Then Response.Write " checked"
	Response.Write " > ��" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
		Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>�Ƿ������������</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isuseercima value='1' "
	If Rs("isuseercima") = 1 Then Response.Write " checked"
	Response.Write " > ��" & vbCrLf
	Response.Write " <input type=radio name=isuseercima value='0'"
	If Rs("isuseercima") = 0 Then Response.Write " checked"
	Response.Write " > ��" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>�������</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=jiafa value='1' "
	If Rs("jiafa") = 1 Then Response.Write " checked"
	Response.Write " > �ӷ�" & vbCrLf
	Response.Write " <input type=radio name=jiafa value='0' "
	If Rs("jiafa") = 0 Then Response.Write " checked"
	Response.Write " > �˷�" & vbCrLf
	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>��1�������λ�ã�</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name='weizhi1' value='"
	response.write rs("weizhi1")
	response.write "'>" & vbCrLf	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>��2�������λ�ã�</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name='weizhi2' value='"
	response.write rs("weizhi2")
	response.write "'>" & vbCrLf	
	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>������λ����λ�ã�</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name='jimaweizhi' value='"
	response.write rs("jimaweizhi")
	response.write "'>" & vbCrLf	
	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>�������˵����</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " �������������Ч��������Ա�����룬��ͬ�Ĺ���Ա�������ò�ͬ��������򣬶��������ڶ��λ�������ϰ���һ���Ĺ������ɡ����翪�����������������Ϊliuyunfan������Ϊ�ӷ�����֤��Ϊ2365��ȡ��1��λ�ú͵�3��λ�ã��������ĵ�4��λ�ã���ô��������Ϊliuy8unfan�����Ҫȡ������Ķ���������룬����[��������]���޸ġ�" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>��IP��</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name='useip' value='"
	response.write rs("useip")
	response.write "'>" & vbCrLf	
	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	
	
	Response.Write " <tr align=""center"">" & vbCrLf
	Response.Write " <td colspan=""2"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=hidden name=id value="""
	Response.Write Request("id")
	Response.Write """>" & vbCrLf
	Response.Write " <input type=button name=Submit4 onclick='javascript:history.go(-1)' value='������һҳ' class=Button> <input type=""submit"" name=""Submit"" value=""�� ��"" class=""button"">" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " </form>" & vbCrLf
	Response.Write " </table>" & vbCrLf
Rs.Close
Set Rs = Nothing
End Sub

Private Sub newpass()
	Dim passnw
	Dim usernw
	Dim aduser
	Dim oldpassword
	If Not ChkAdmin("9999") Then
		ErrMsg = "<li>��û�д˲���Ȩ��!</li><li>����ʲô��������ϵվ����</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("select * from ECCMS_Admin where id=" & Request("id"))
	oldpassword = Rs("password")
	If Request("username2") = "" Then
		ErrMsg = "<li>���������Ա���֡�<a href=?>�� <font color=red>����</font> ��</a>"
		Founderr = True
		Exit Sub
	Else
		usernw = Trim(Request("username2"))
	End If
	
	if Request.Form("isuseercima") = "1" then
		If Request.Form("weizhi1") = "" or Request.Form("weizhi2") = "" or Request.Form("jimaweizhi") = "" Then
			ErrMsg = "���������������������������ݣ�"
			Founderr = True
			Exit Sub
		else
			if not (IsNumeric(Trim(Request.Form("weizhi1"))) and IsNumeric(Trim(Request.Form("weizhi2"))) and IsNumeric(Trim(Request.Form("jimaweizhi")))) then
				FoundErr = True
				ErrMsg = ErrMsg + "<li>����������������ֻ���������֣�</li>"
				exit sub
			else
				if cint(Trim(Request.Form("weizhi1")))>4 or cint(Trim(Request.Form("weizhi2")))>4 then
					FoundErr = True
					ErrMsg = ErrMsg + "<li>���������������ֳ�����֤��ĳ���4λ�����޸ģ�</li>"
					exit sub
				end if
			end if

		End If
	end if


	If Request("password2") = "" Then
		passnw = "û���޸�"
	Else
		passnw = Request("password2")
	End If
	Set Rs = CreateObject("adodb.recordset")
	SQL = "select * from ECCMS_Admin where username='" & Trim(Request("oldusername")) & "'"
	Rs.Open SQL, conn, 1, 3
	If Not Rs.EOF And Not Rs.bof Then
		Rs("username") = usernw
		If Request("password2") <> "" Then Rs("password") = md5(Request("password2"))
		If CInt(Request.Form("AdminGrade")) = 999 Then
			Rs("status") = "�߼�����Ա"
		Else
			Rs("status") = "��ͨ����Ա"
		End If
		Rs("AdminGrade") = Request.Form("AdminGrade")
		Rs("isLock") = Request.Form("isLock")
		Rs("isAloneLogin") = Request.Form("isAloneLogin")
		rs("isuseercima")= Request.Form("isuseercima")
		rs("jiafa")= Request.Form("jiafa")
		rs("weizhi1")= Request.Form("weizhi1")
		rs("weizhi2")= Request.Form("weizhi2")
		rs("jimaweizhi")= Request.Form("jimaweizhi")
		'if Request.Form("useip")<>"" then
		rs("useip")= Request.Form("useip")
		'end if



		Succeed ("<li>����Ա���ϸ��³ɹ������ס������Ϣ��<br> ����Ա��" & Request("username2") & " <BR> ��   �룺" & passnw & "")
		Rs.update
	End If
	Rs.Close
	Set Rs = Nothing	
End Sub

Private Sub addadmin()
	If Not ChkAdmin("9999") Then
		ErrMsg = "<li>��û�д˲���Ȩ��!</li><li>����ʲô��������ϵվ����</li>"
		Founderr = True
		Exit Sub
	End If
	Response.Write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write "<form action=""?action=savenew"" method=post>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th colspan=2 height=23>����Ա��������ӹ���Ա" & vbCrLf
	Response.Write " </th>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>��̨��½���ƣ�</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name=""username2"">" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>��̨��½���룺</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=""password"" name=""password2"">" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>����Ա����</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=AdminGrade value='0' checked> ��ͨ����Ա&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=AdminGrade value='999'> �߼�����Ա ��ӵ�����Ȩ�ޣ�" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td align=""right"" class=tablerow1>����һ������Ա��½��</td>" & vbCrLf
	Response.Write " <td class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isAloneLogin value='0' checked> ��&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=isAloneLogin value='1'> ��" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>�Ƿ񼤻����Ա��</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isLock value='1' checked> ��&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=isLock value='0'> ��" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>�Ƿ������������</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isuseercima value='1' checked> ��&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=isuseercima value='0'> ��" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>�������</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=jiafa value='1' checked> �ӷ�&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=jiafa value='0'> �˷�" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>��1�������λ�ã�</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name=""weizhi1"">" & vbCrLf	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>��2�������λ�ã�</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name=""weizhi2"">" & vbCrLf	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>������λ����λ�ã�</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name=""jimaweizhi"">" & vbCrLf	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>�������˵����</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " �������������Ч��������Ա�����룬��ͬ�Ĺ���Ա�������ò�ͬ��������򣬶��������ڶ��λ�������ϰ���һ���Ĺ������ɡ����翪�����������������Ϊliuyunfan������Ϊ�ӷ�����֤��Ϊ2365��ȡ��1��λ�ú͵�3��λ�ã��������ĵ�4��λ�ã���ô��������Ϊliuy8unfan�����Ҫȡ������Ķ���������룬����[��������]���޸ġ�" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>�û�IP�󶨣�</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name=""useip"">" & vbCrLf	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	
	Response.Write " <tr align=""center"">" & vbCrLf
	Response.Write " <td colspan=""2"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=button name=Submit4 onclick='javascript:history.go(-1)' value='������һҳ' class=Button>  <input type=""submit"" name=""Submit"" value=""�� ��"" class=""button"">" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " </form>" & vbCrLf
	Response.Write " </table>" & vbCrLf
End Sub

Private Sub savenew()
	Dim adminuserid
	If Not ChkAdmin("9999") Then
		ErrMsg = "<li>��û�д˲���Ȩ��!</li><li>����ʲô��������ϵվ����</li>"
		Founderr = True
		Exit Sub
	End If
	If Request.Form("username2") = "" Then
		ErrMsg = "�������̨��½�û�����"
		Founderr = True
		Exit Sub
	Else
		adminuserid = Request.Form("username2")
	End If
	If Request.Form("password2") = "" Then
		ErrMsg = "�������̨��½���룡"
		Founderr = True
		Exit Sub
	End If
	if Request.Form("isuseercima") = "1" then
		If Request.Form("weizhi1") = "" or Request.Form("weizhi2") = "" or Request.Form("jimaweizhi") = "" Then
			ErrMsg = "���������������������������ݣ�"
			Founderr = True
			Exit Sub
		else
			if not (IsNumeric(Trim(Request.Form("weizhi1"))) and IsNumeric(Trim(Request.Form("weizhi2"))) and IsNumeric(Trim(Request.Form("jimaweizhi")))) then
				FoundErr = True
				ErrMsg = ErrMsg + "<li>����������������ֻ���������֣�</li>"
				exit sub
			else
				if cint(Trim(Request.Form("weizhi1")))>4 or cint(Trim(Request.Form("weizhi2")))>4 then
					FoundErr = True
					ErrMsg = ErrMsg + "<li>���������������ֳ�����֤��ĳ���4λ�����޸ģ�</li>"
					exit sub
				end if
			end if

		End If
	end if

	
	
	
	Set Rs = enchiasp.Execute("select username from ECCMS_Admin where username='" & Replace(Request.Form("username2"), "'", "") & "'")
	If Not (Rs.EOF And Rs.bof) Then
		ErrMsg = "��������û����Ѿ��ڹ����û��д��ڣ�"
		Founderr = True
		Exit Sub
	End If
	Set Rs = Server.CreateObject("adodb.recordset")
	SQL = "select * from ECCMS_Admin where (id is null)"
	Rs.open SQL,conn,1,3
	Rs.addnew
		Rs("username") = Replace(Request.Form("username2"), "'", "")
		If CInt(Request.Form("AdminGrade")) = 999 Then
			Rs("status") = "�߼�����Ա"
		Else
			Rs("status") = "��ͨ����Ա"
		End If
		Rs("password") = md5(Request.Form("password2"))
		Rs("isLock") = Request.Form("isLock")
		Rs("AdminGrade") = Request.Form("AdminGrade")
		Rs("Adminflag") = ",,,,,,,,,,,,,,,"
		Rs("LoginTime") = Now()
		Rs("Loginip") = enchiasp.GetUserIP
		Rs("RandomCode") = enchiasp.GetRandomCode
		Rs("isAloneLogin") = Request.Form("isAloneLogin")
		rs("isuseercima")= Request.Form("isuseercima")
		rs("jiafa")= Request.Form("jiafa")
		rs("weizhi1")= Request.Form("weizhi1")
		rs("weizhi2")= Request.Form("weizhi2")
		rs("jimaweizhi")= Request.Form("jimaweizhi")
		'if Request.Form("useip")<>"" then
		rs("useip")= Request.Form("useip")
		'end if


	Rs.update
	Rs.close:set Rs=Nothing
	Succeed ("�û�ID:" & adminuserid & " ��ӳɹ����뵽����Ա���������Ӧ��Ȩ�ޣ������޸��뷵�ع���Ա����")
End Sub

Private Sub userinfo()
	Dim Adminflag,rsChannel
	Dim ChanID,ModuleName,strModules
	Set Rs = enchiasp.Execute("SELECT id,Adminflag FROM ECCMS_Admin WHERE id=" & Request("id"))
	Adminflag = Rs("Adminflag")
	Rs.Close
	Set Rs = Nothing
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan=6>����ԱȨ�޹���(��ѡ����Ӧ��Ȩ�޷��������Ա)</th>
</tr>
<form name=myform method=post action=?action=renew>
<input type=hidden name=id value="<%=Request("id")%>">
<tr>
	<td class=tablerow2 colspan=6>&nbsp;<b>��������</b></td>
</tr>
<tr>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="SiteConfig" <%If InStr2(Adminflag, "SiteConfig") <> 0 Then Response.Write "checked"%>> ��������</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="Advertise" <%If InStr2(Adminflag, "Advertise") <> 0 Then Response.Write "checked"%>> ������</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="Channel" <%If InStr2(Adminflag, "Channel") <> 0 Then Response.Write "checked"%>> Ƶ������</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="Template" <%If InStr2(Adminflag, "Template") <> 0 Then Response.Write "checked"%>> ģ�����</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="TemplateLoad" <%If InStr2(Adminflag, "TemplateLoad") <> 0 Then Response.Write "checked"%>> ģ�嵼�롢����</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="Announce" <%If InStr2(Adminflag, "Announce") <> 0 Then Response.Write "checked"%>> �������</td>
</tr>
<tr>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminLog" <%If InStr2(Adminflag, "AdminLog") <> 0 Then Response.Write "checked"%>> ��־����</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="SendMessage" <%If InStr2(Adminflag, "SendMessage") <> 0 Then Response.Write "checked"%>> ������Ϣ</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="CreateIndex" <%If InStr2(Adminflag, "CreateIndex") <> 0 Then Response.Write "checked"%>> ������ҳ</td>
	<td class=tablerow1></td>
	<td class=tablerow1></td>
	<td class=tablerow1></td>
</tr>
<%
	Set rsChannel = enchiasp.Execute("SELECT ChannelID,ChannelName,modules,ModuleName FROM ECCMS_Channel WHERE StopChannel = 0 And ChannelID <> 4 And ChannelType < 2 Order By orders Asc")
	Do While Not rsChannel.EOF
	ChanID = rsChannel("ChannelID")
	Select Case rsChannel("modules")
		Case 1:strModules = "Article"
		Case 2:strModules = "Soft"
		Case 3:strModules = "Shop"
		Case 5:strModules = "Flash"
		Case 6:strModules = "yemian"
		Case 7:strModules = "job"
	Case Else
		strModules = "Article"
	End Select
%>
<tr>
	<td class=tablerow2 colspan=6>&nbsp;<b><%=rsChannel("ChannelName")%></b></td>

</tr>

<%
select case rsChannel("modules")
	case 6
		'��ҳ��ͼ��Ƶ��
		%>
		<tr>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="Add<%=strModules & ChanID%>" <%If InStr2(Adminflag, "Add" & strModules & ChanID) <> 0 Then Response.Write "checked"%>> �������</td> 
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="Admin<%=strModules & ChanID%>" <%If InStr2(Adminflag, "Admin" & strModules & ChanID) <> 0 Then Response.Write "checked"%>> ���ݹ���</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminClass<%=ChanID%>" <%If InStr2(Adminflag, "AdminClass" & ChanID) <> 0 Then Response.Write "checked"%>> ��Ŀ����</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminUpload<%=ChanID%>" <%If InStr2(Adminflag, "AdminUpload" & ChanID) <> 0 Then Response.Write "checked"%>> �ϴ��ļ�����</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminSelect<%=ChanID%>" <%If InStr2(Adminflag, "AdminSelect" & ChanID) <> 0 Then Response.Write "checked"%>> ѡ���ϴ��ļ�</td>
			<td class=tablerow1></td>
		</tr>
		<%
	
	case else
		%>
		<tr>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="Add<%=strModules & ChanID%>" <%If InStr2(Adminflag, "Add" & strModules & ChanID) <> 0 Then Response.Write "checked"%>> ���<%=rsChannel("ModuleName")%></td> 
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="Admin<%=strModules & ChanID%>" <%If InStr2(Adminflag, "Admin" & strModules & ChanID) <> 0 Then Response.Write "checked"%>> <%=rsChannel("ModuleName")%>����</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminClass<%=ChanID%>" <%If InStr2(Adminflag, "AdminClass" & ChanID) <> 0 Then Response.Write "checked"%>> <%=rsChannel("ModuleName")%>�������</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminUpload<%=ChanID%>" <%If InStr2(Adminflag, "AdminUpload" & ChanID) <> 0 Then Response.Write "checked"%>> �ϴ��ļ�����</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminSelect<%=ChanID%>" <%If InStr2(Adminflag, "AdminSelect" & ChanID) <> 0 Then Response.Write "checked"%>> ѡ���ϴ��ļ�</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminJsFile<%=ChanID%>" <%If InStr2(Adminflag, "AdminJsFile" & ChanID) <> 0 Then Response.Write "checked"%>> JS�ļ�����</td> 

		</tr>
		<tr>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="Auditing<%=ChanID%>" <%If InStr2(Adminflag, "Auditing" & ChanID) <> 0 Then Response.Write "checked"%>>  <%=rsChannel("ModuleName")%>��˹���</td>
			<td class=tablerow1><%If rsChannel("modules") = 2 Or rsChannel("modules") = 5 Then%><input type="checkbox" name="Adminflag" value="DownServer<%=ChanID%>" <%If InStr2(Adminflag, "DownServer" & ChanID) <> 0 Then Response.Write "checked"%>> ���ط���������<%End If%></td>
			<td class=tablerow1><%If rsChannel("modules") = 2 Then%><input type="checkbox" name="Adminflag" value="ErrorSoft<%=ChanID%>" <%If InStr2(Adminflag, "ErrorSoft" & ChanID) <> 0 Then Response.Write "checked"%>> �����������<%End If%></td>
			<td class=tablerow1></td> 			
			<td class=tablerow1></td>
			<td class=tablerow1></td>
		</tr>
		<%
end select
		rsChannel.movenext
	Loop
	Set rsChannel = Nothing
%>


<tr>
	<td class=tablerow2 colspan=6>&nbsp;<b>��������</b></td>
</tr>
<tr>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="Vote" <%If InStr2(Adminflag, "Vote") <> 0 Then Response.Write "checked"%>> ͶƱ����</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="FriendLink" <%If InStr2(Adminflag, "FriendLink") <> 0 Then Response.Write "checked"%>> �������ӹ���</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="UploadFile" <%If InStr2(Adminflag, "UploadFile") <> 0 Then Response.Write "checked"%>> �ϴ��ļ�</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="GuestBook" <%If InStr2(Adminflag, "GuestBook") <> 0 Then Response.Write "checked"%>> ���Թ���</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="rizhi" <%If InStr2(Adminflag, "rizhi") <> 0 Then Response.Write "checked"%>> ��־����</td>
	<td class=tablerow1></td>

</tr>
<tr>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian" <%If InStr2(Adminflag, "flashtupian") <> 0 Then Response.Write "checked"%>> ����ͼƬ�任</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian2" <%If InStr2(Adminflag, "flashtupian2") <> 0 Then Response.Write "checked"%>> �ഺ֮��ͼƬ�任</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian3" <%If InStr2(Adminflag, "flashtupian3") <> 0 Then Response.Write "checked"%>> ���Ľ���ͼƬ�任</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian4" <%If InStr2(Adminflag, "flashtupian4") <> 0 Then Response.Write "checked"%>> ְ��֮��ͼƬ�任</td>
<td class=tablerow1><input type="checkbox" name="Adminflag" value="gundong" <%If InStr2(Adminflag, "gundong") <> 0 Then Response.Write "checked"%>> ͼƬ���ҹ�������</td>
	<td class=tablerow1></td>

</tr>

<tr>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian5" <%If InStr2(Adminflag, "flashtupian5") <> 0 Then Response.Write "checked"%>>������ԴͼƬ�任</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian6" <%If InStr2(Adminflag, "flashtupian6") <> 0 Then Response.Write "checked"%>>�Ƽ�����ͼƬ�任</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminSelect0" <%If InStr2(Adminflag, "AdminSelect0") <> 0 Then Response.Write "checked"%>>��ͼƬ��ѡ���ļ�</td>
	<td class=tablerow1></td>
<td class=tablerow1></td>
	<td class=tablerow1></td>

</tr>




<tr>
	<td class=tablerow2 colspan=6 align=center><input type=button name=Submit4 onclick='javascript:history.go(-1)' value='������һҳ' class=Button> ����<input class=Button type=button name=chkall value='ȫѡ' onClick='CheckAll(this.form)'><input class=Button type=button name=chksel value='��ѡ' onClick='ContraSel(this.form)'>
	<input type="submit" name="Submit" value="���¹���ԱȨ��" class=button></td>
</tr>
</form>
</table>

<%
End Sub

Private Sub UpdateFlag()
	Set Rs = Server.CreateObject("adodb.recordset")
	SQL = "SELECT * FROM ECCMS_Admin WHERE id=" & Request("id")
	Rs.Open SQL, conn, 1, 3
	If Not (Rs.EOF And Rs.BOF) Then
		Rs("Adminflag") = Replace(Replace(Request("Adminflag"), "'", ""), " ", "")
		Rs.update
	End If
	Rs.Close
	Set Rs = Nothing
	Sucmsg = "<li>����Ա���³ɹ������ס������Ϣ��"
	Succeed (Sucmsg)
End Sub
%>