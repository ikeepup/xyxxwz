<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/email.asp"-->

<!--#include file="head.inc"-->
<script language="JavaScript">locationid.innerHTML = "�һ�����";</script>
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

Dim Founderr
Dim useremail
Dim topic
Dim mailbody
Dim sendmsg
Dim answer
Dim repassword
Dim username
Dim password
Dim rs
Dim sql
Dim Errmsg
Dim Sucmsg

Founderr = False
Set Rs = Server.CreateObject("adodb.recordset")
Response.Write "<br>"
Founderr = False
If Founderr Then
	Response.Write "<script>alert('" & Errmsg & "');history.go(-1)</script>"
Else
	If Trim(Request("action")) = "step1" Then
		Call step1
	ElseIf Trim(Request("action")) = "step2" Then
		Call step2
	ElseIf Trim(Request("action")) = "step3" Then
		Call step3
	Else
		Call main
	End If
	If Founderr Then Response.Write "<script>alert('" & Errmsg & "');history.go(-1)</script>"
End If

Private Sub step1()
	If Trim(Request("username")) = "" Then
		Founderr = True
		Errmsg = Errmsg + "�����������û�����"
		Exit Sub
	Else
		UserName = enchiasp.CheckBadstr(Request("username"))
	End If
	If CInt(enchiasp.IsCloseMail) = 1 Then
		Set Rs = enchiasp.Execute("Select * from [ECCMS_User] where username='" & UserName & "'")
	Else
		Set Rs = enchiasp.Execute("Select * from [ECCMS_User] where username='" & UserName & "'")
	End If
	If Rs.EOF And Rs.bof Then
		Founderr = True
		Errmsg = Errmsg + "��������û����������ڣ����������롣�������ڸ�ϵͳ��֧���ʼ����ͣ�ֻ��ͨ����ϵվ��������롣"
	Else
		If Rs(13) = "" Or IsNull(Rs(13)) Then
			Founderr = True
			Errmsg = Errmsg + "���û�û����д�������⼰�𰸣�ֻ����д���û����ܼ�����"
		Else
			Response.Write "<form action=""sendpass.asp?action=step2"" method=""post""> "
			Response.Write "<table cellpadding=4 cellspacing=1 align=center class=tableborder1>"
			Response.Write " <tr>"
			Response.Write " <th valign=middle colspan=2 align=center height=25>ȡ�����루�ڶ������ش����⣩</th></tr>"
			Response.Write " <tr>"
			Response.Write " <td valign=middle class=tablerow><b>�ʣ��⣺</b></td>"
			Response.Write " <td class=tablerow valign=middle>"
			Response.Write Rs("question")
			Response.Write "</td>"
			Response.Write " </tr>"
			Response.Write " <tr>"
			Response.Write " <td class=tablerow valign=middle><b>�𣠰���</b></td>"
			Response.Write " <td class=tablerow valign=middle><INPUT name=""answer"" type=text class=inputbody></td>"
			Response.Write " </tr>"
			Response.Write " <tr>"
			Response.Write " <td class=tablerow colspan=2>"
			Response.Write " <b>˵����</b>����д����ȷ������𰸡�</td></tr>"
			Response.Write " <tr>"
			Response.Write " <td class=tablerow valign=middle colspan=2 align=center><input type=submit name=""submit"" value="" ��һ�� "" class=button></td></tr>"
			Response.Write "</table>"
			Response.Write "<input type=hidden value="""
			Response.Write UserName
			Response.Write """ name=username>"
			Response.Write "</form>"
		End If
	End If
	Rs.Close
	Set Rs = Nothing
End Sub

Private Sub step2()
	If Trim(Request("username")) = "" Then
		Founderr = True
		Errmsg = Errmsg + "�����������û�����"
		Exit Sub
	Else
		UserName = enchiasp.CheckBadstr(Request("username"))
	End If
	If enchiasp.checkpost = False Then
		Errmsg = Errmsg + "���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ���ԡ�"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request("answer")) = "" Then
		Founderr = True
		Errmsg = Errmsg + "��������������𰸡�"
		Exit Sub
	Else
		Answer = md5(Trim(Request("answer")))
	End If
	Set Rs = enchiasp.Execute("select * from ECCMS_User where username = '" & UserName & "' and answer = '" & Answer & "'")
	If Rs.EOF And Rs.bof Then
		Founderr = True
		Errmsg = Errmsg + "�����������𰸲���ȷ�����������롣"
		Exit Sub
	Else
		Response.Write "<form action=""sendpass.asp?action=step3"" method=""post""> "
		Response.Write "<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>"
		Response.Write " <tr>"
		Response.Write " <th valign=middle colspan=2 height=25>ȡ�����루���������޸����룩</th></tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>�ʣ��⣺</b></td>"
		Response.Write " <td class=tablerow valign=middle>"
		Response.Write Rs("question")
		Response.Write "</td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>�𣠰���</b></td>"
		Response.Write " <td class=tablerow valign=middle>"
		Response.Write Request("answer")
		Response.Write "</td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>�����룺</b></td>"
		Response.Write " <td class=tablerow valign=middle><input type=password name=password class=inputbody></td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>ȷ�������룺</b></td>"
		Response.Write " <td class=tablerow valign=middle><input type=password name=repassword class=inputbody></td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow colspan=2>"
		Response.Write " <b>˵����</b>"
		Response.Write " "
		If CInt(enchiasp.IsCloseMail) = 0 Then
			Response.Write " ϵͳ�����Զ���һ���ʼ�����ע��ʱ��д�����䣬�ڴ��ʼ��е����뼤���ַ�����������뽫��ʽ���á�"
			Response.Write " "
		Else
			Response.Write " ����д���������룬����ס������д��Ϣ��"
		End If
		Response.Write " </td></tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle colspan=2 align=center><input type=submit name=""submit"" value=""��һ��"" class=button></td></tr>"
		Response.Write "</table>"
		Response.Write "<input type=hidden value="""
		Response.Write Request("answer")
		Response.Write """ name=answer>"
		Response.Write "<input type=hidden value="""
		Response.Write UserName
		Response.Write """ name=username>"
		Response.Write "</form>"
	End If
	Rs.Close
	Set Rs = Nothing
End Sub

Private Sub step3()
	If Trim(Request("username")) = "" Then
		Founderr = True
		Errmsg = Errmsg + "�����������û�����"
		Exit Sub
	Else
		UserName = enchiasp.CheckBadstr(Request("username"))
	End If
	If Trim(Request("answer")) = "" Then
		Founderr = True
		Errmsg = Errmsg + "��������������𰸡�"
		Exit Sub
	Else
		Answer = md5(Request("answer"))
	End If
	If Trim(Request("password")) = "" Or Len(Request("password")) > 15 Or Len(Request("password")) < 6 Then
		Founderr = True
		Errmsg = Errmsg + "����������������(���Ȳ��ܴ���15С��6)��"
		Exit Sub
	ElseIf Trim(Request("repassword")) = "" Then
		Founderr = True
		Errmsg = Errmsg + "���ٴ��������������롣"
		Exit Sub
	ElseIf Trim(Request("password")) <> Trim(Request("repassword")) Then
		Founderr = True
		Errmsg = Errmsg + "��������������ȷ�ϲ�һ������ȷ������д����Ϣ��"
		Exit Sub
	Else
		PassWord = md5(Request("password"))
	End If
	If Not IsObject(Conn) Then ConnectionDatabase
	Set Rs = Server.CreateObject("adodb.recordset")
	SQL = "select * from [ECCMS_User] where username='" & UserName & "' and answer='" & Answer & "'"
	Rs.Open SQL, conn, 1, 3
	If Rs.EOF And Rs.bof Then
		Founderr = True
		Errmsg = Errmsg + "�����������𰸲���ȷ�����������롣"
	Else
		If CInt(enchiasp.IsCloseMail) = 0 Then
			RePassword = Request.Form("password")
			Answer = Request.Form("answer")
			PassWord = Rs("password")
			UsereMail = Rs("usermail")
			Call sendusermail
			If SendMail = "OK" Then
				SendMsg = "ϵͳ�Ѿ�����һ���ʼ�����ע��ʱ��д�����䣬�ڴ��ʼ��е����뼤���ַ�����������뽫��ʽ���á�"
			Else
				Rs("password") = md5(RePassword)
				Rs.Update
				SendMsg = "����ϵͳ���󣬸������͵���������δ�ɹ������Ѿ��޸�����ɹ�����ʹ���������½ϵͳ��"
			End If
		Else
			Rs("password") = PassWord
			Rs.Update
		End If
		Response.Write "<form action=""login.asp"" method=""post""> "
		Response.Write "<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>"
		Response.Write " <tr>"
		Response.Write " <th valign=middle colspan=2>ȡ�����루���Ĳ����޸ĳɹ���</th></tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>�ʣ��⣺</b></td>"
		Response.Write " <td class=tablerow valign=middle>"
		Response.Write Rs("question")
		Response.Write "</td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>�𣠰���</b></td>"
		Response.Write " <td class=tablerow valign=middle>"
		Response.Write Request("answer")
		Response.Write "</td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>�����룺</b></td>"
		Response.Write " <td class=tablerow valign=middle>"
		Response.Write Request("password")
		Response.Write "</td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow colspan=2>"
		Response.Write " <b>˵����</b>"
		Response.Write " "
		If CInt(enchiasp.IsCloseMail) = 0 Then
			Response.Write SendMsg
			Response.Write " "
		Else
			Response.Write " ���ס���������벢ʹ��������<a href=login.asp>��½</a>��"
			Response.Write " "
		End If
		Response.Write "</td></tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle colspan=2 align=center><input type=submit name=""submit"" value=""�� ��"" class=button></td></tr>"
		Response.Write "</table>"
		Response.Write "</form>"
	End If
	Rs.Close
	Set Rs = Nothing
End Sub

Private Sub main()
	Response.Write "<form action=""sendpass.asp?action=step1"" method=""post""> "
	Response.Write "<table cellpadding=6 cellspacing=1 align=center class=tableborder1>"
	Response.Write " <tr>"
	Response.Write " <th valign=middle colspan=2>ȡ������</b>����һ�����û�����</th></tr>"
	Response.Write " <tr>"
	Response.Write " <td class=tablerow valign=middle>�����������û���</td>"
	Response.Write " <td class=tablerow valign=middle><INPUT name=""username"" type=text class=inputbody></td>"
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write " <td class=tablerow colspan=2>"
	Response.Write " <b>˵����</b>������ֻ���޸��������룬���ܶ�ԭ��������޸ģ���ȷ�����Ѿ���д���������⼰�𰸡�</td></tr>"
	Response.Write " <tr>"
	Response.Write " <td class=tablerow valign=middle colspan=2 align=center><input type=submit name=""submit"" value="" ��һ�� "" class=button></td></tr>"
	Response.Write "</table>"
	Response.Write "</form>"
End Sub

Private Sub sendusermail()
	On Error Resume Next
	Topic = "����" & enchiasp.SiteName & "��������Ϣ"
	MailBody = MailBody & "<style>A:visited { TEXT-DECORATION: none }"
	MailBody = MailBody & "A:active  { TEXT-DECORATION: none }"
	MailBody = MailBody & "A:hover   { TEXT-DECORATION: underline overline }"
	MailBody = MailBody & "A:link    { text-decoration: none;}"
	MailBody = MailBody & "A:visited { text-decoration: none;}"
	MailBody = MailBody & "A:active  { TEXT-DECORATION: none;}"
	MailBody = MailBody & "A:hover   { TEXT-DECORATION: underline overline}"
	MailBody = MailBody & "BODY   { FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
	MailBody = MailBody & "TD    { FONT-FAMILY: ����; FONT-SIZE: 9pt }</style>"
	MailBody = MailBody & "<TABLE border=0 width='95%' align=center><TBODY><TR>"
	MailBody = MailBody & "<TD valign=middle align=top>"
	MailBody = MailBody & "" & enchiasp.CheckStr(UserName) & "�����ã�<br><br>"
	MailBody = MailBody & "��ӭ��ʹ�ñ�ϵͳ�������������ܣ�<b>��������ϣ�������������룬�벻Ҫ�������ļ�������</b>��<br>"
	MailBody = MailBody & "<a href=http://" & Request.servervariables("server_name") & Replace(Request.servervariables("script_name"), "sendpass.asp", "") & "activepass.asp?username=" & enchiasp.CheckStr(UserName) & "&pass=" & PassWord & "&repass=" & RePassword & "&answer=" & Answer & ">http://" & Request.servervariables("server_name") & Replace(Request.servervariables("script_name"), "sendpass.asp", "") & "activepass.asp?username=" & enchiasp.CheckStr(UserName) & "&pass=" & PassWord & "&repass=" & RePassword & "&answer=" & Answer & "</a>"
	MailBody = MailBody & "<br><br>"
	MailBody = MailBody & "<center><font color=red>�ٴθ�л��ʹ�ñ�ϵͳ��������һ��������������ϼ�԰��</font>"
	MailBody = MailBody & "</TD></TR></TBODY></TABLE><br><hr width=95% size=1>"
	Select Case CInt(enchiasp.SendMailType)
		Case 0
			SendMsg = "����ϵͳ���󣬸������͵���������δ�ɹ��������ұߵ����ӽ��������뼤�<a href=http://" & Request.servervariables("server_name") & Replace(Request.servervariables("script_name"), "sendpass.asp", "") & "activepass.asp?username=" & enchiasp.CheckStr(UserName) & "&pass=" & PassWord & "&repass=" & RePassword & "><B>��������</B></a>"
		Case 1
			Call jmail(UsereMail, Topic, MailBody)
		Case 2
			Call Cdonts(UsereMail, Topic, MailBody)
		Case 3
			Call aspemail(UsereMail, Topic, MailBody)
		Case Else
			SendMsg = "����ϵͳ���󣬸������͵���������δ�ɹ��������ұߵ����ӽ��������뼤�<a href=http://" & Request.servervariables("server_name") & Replace(Request.servervariables("script_name"), "sendpass.asp", "") & "activepass.asp?username=" & enchiasp.CheckStr(UserName) & "&pass=" & PassWord & "&repass=" & RePassword & "><B>��������</B></a>"
	End Select
End Sub
%>
<!--#include file="foot.inc"-->