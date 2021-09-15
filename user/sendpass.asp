<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/email.asp"-->

<!--#include file="head.inc"-->
<script language="JavaScript">locationid.innerHTML = "找回密码";</script>
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
		Errmsg = Errmsg + "请输入您的用户名。"
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
		Errmsg = Errmsg + "您输入的用户名并不存在，请重新输入。或者由于该系统不支持邮件发送，只能通过联系站长获得密码。"
	Else
		If Rs(13) = "" Or IsNull(Rs(13)) Then
			Founderr = True
			Errmsg = Errmsg + "该用户没有填写密码问题及答案，只有填写的用户方能继续。"
		Else
			Response.Write "<form action=""sendpass.asp?action=step2"" method=""post""> "
			Response.Write "<table cellpadding=4 cellspacing=1 align=center class=tableborder1>"
			Response.Write " <tr>"
			Response.Write " <th valign=middle colspan=2 align=center height=25>取回密码（第二步：回答问题）</th></tr>"
			Response.Write " <tr>"
			Response.Write " <td valign=middle class=tablerow><b>问题：</b></td>"
			Response.Write " <td class=tablerow valign=middle>"
			Response.Write Rs("question")
			Response.Write "</td>"
			Response.Write " </tr>"
			Response.Write " <tr>"
			Response.Write " <td class=tablerow valign=middle><b>答案：</b></td>"
			Response.Write " <td class=tablerow valign=middle><INPUT name=""answer"" type=text class=inputbody></td>"
			Response.Write " </tr>"
			Response.Write " <tr>"
			Response.Write " <td class=tablerow colspan=2>"
			Response.Write " <b>说明：</b>请填写您正确的问题答案。</td></tr>"
			Response.Write " <tr>"
			Response.Write " <td class=tablerow valign=middle colspan=2 align=center><input type=submit name=""submit"" value="" 下一步 "" class=button></td></tr>"
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
		Errmsg = Errmsg + "请输入您的用户名。"
		Exit Sub
	Else
		UserName = enchiasp.CheckBadstr(Request("username"))
	End If
	If enchiasp.checkpost = False Then
		Errmsg = Errmsg + "您提交的数据不合法，请不要从外部提交发言。"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request("answer")) = "" Then
		Founderr = True
		Errmsg = Errmsg + "请输入您的问题答案。"
		Exit Sub
	Else
		Answer = md5(Trim(Request("answer")))
	End If
	Set Rs = enchiasp.Execute("select * from ECCMS_User where username = '" & UserName & "' and answer = '" & Answer & "'")
	If Rs.EOF And Rs.bof Then
		Founderr = True
		Errmsg = Errmsg + "您输入的问题答案不正确，请重新输入。"
		Exit Sub
	Else
		Response.Write "<form action=""sendpass.asp?action=step3"" method=""post""> "
		Response.Write "<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>"
		Response.Write " <tr>"
		Response.Write " <th valign=middle colspan=2 height=25>取回密码（第三步：修改密码）</th></tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>问题：</b></td>"
		Response.Write " <td class=tablerow valign=middle>"
		Response.Write Rs("question")
		Response.Write "</td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>答案：</b></td>"
		Response.Write " <td class=tablerow valign=middle>"
		Response.Write Request("answer")
		Response.Write "</td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>新密码：</b></td>"
		Response.Write " <td class=tablerow valign=middle><input type=password name=password class=inputbody></td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>确认新密码：</b></td>"
		Response.Write " <td class=tablerow valign=middle><input type=password name=repassword class=inputbody></td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow colspan=2>"
		Response.Write " <b>说明：</b>"
		Response.Write " "
		If CInt(enchiasp.IsCloseMail) = 0 Then
			Response.Write " 系统将会自动发一封邮件到您注册时填写的邮箱，在打开邮件中的密码激活地址后，您的新密码将正式启用。"
			Response.Write " "
		Else
			Response.Write " 请填写您的新密码，并记住您所填写信息。"
		End If
		Response.Write " </td></tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle colspan=2 align=center><input type=submit name=""submit"" value=""下一步"" class=button></td></tr>"
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
		Errmsg = Errmsg + "请输入您的用户名。"
		Exit Sub
	Else
		UserName = enchiasp.CheckBadstr(Request("username"))
	End If
	If Trim(Request("answer")) = "" Then
		Founderr = True
		Errmsg = Errmsg + "请输入您的问题答案。"
		Exit Sub
	Else
		Answer = md5(Request("answer"))
	End If
	If Trim(Request("password")) = "" Or Len(Request("password")) > 15 Or Len(Request("password")) < 6 Then
		Founderr = True
		Errmsg = Errmsg + "请输入您的新密码(长度不能大于15小于6)。"
		Exit Sub
	ElseIf Trim(Request("repassword")) = "" Then
		Founderr = True
		Errmsg = Errmsg + "请再次输入您的新密码。"
		Exit Sub
	ElseIf Trim(Request("password")) <> Trim(Request("repassword")) Then
		Founderr = True
		Errmsg = Errmsg + "您输入的新密码和确认不一样，请确认您填写的信息。"
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
		Errmsg = Errmsg + "您输入的问题答案不正确，请重新输入。"
	Else
		If CInt(enchiasp.IsCloseMail) = 0 Then
			RePassword = Request.Form("password")
			Answer = Request.Form("answer")
			PassWord = Rs("password")
			UsereMail = Rs("usermail")
			Call sendusermail
			If SendMail = "OK" Then
				SendMsg = "系统已经发送一封邮件到您注册时填写的邮箱，在打开邮件中的密码激活地址后，您的新密码将正式启用。"
			Else
				Rs("password") = md5(RePassword)
				Rs.Update
				SendMsg = "由于系统错误，给您发送的密码资料未成功。您已经修改密码成功，请使用新密码登陆系统。"
			End If
		Else
			Rs("password") = PassWord
			Rs.Update
		End If
		Response.Write "<form action=""login.asp"" method=""post""> "
		Response.Write "<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>"
		Response.Write " <tr>"
		Response.Write " <th valign=middle colspan=2>取回密码（第四步：修改成功）</th></tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>问题：</b></td>"
		Response.Write " <td class=tablerow valign=middle>"
		Response.Write Rs("question")
		Response.Write "</td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>答案：</b></td>"
		Response.Write " <td class=tablerow valign=middle>"
		Response.Write Request("answer")
		Response.Write "</td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle><b>新密码：</b></td>"
		Response.Write " <td class=tablerow valign=middle>"
		Response.Write Request("password")
		Response.Write "</td>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow colspan=2>"
		Response.Write " <b>说明：</b>"
		Response.Write " "
		If CInt(enchiasp.IsCloseMail) = 0 Then
			Response.Write SendMsg
			Response.Write " "
		Else
			Response.Write " 请记住您的新密码并使用新密码<a href=login.asp>登陆</a>。"
			Response.Write " "
		End If
		Response.Write "</td></tr>"
		Response.Write " <tr>"
		Response.Write " <td class=tablerow valign=middle colspan=2 align=center><input type=submit name=""submit"" value=""返 回"" class=button></td></tr>"
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
	Response.Write " <th valign=middle colspan=2>取回密码</b>（第一步：用户名）</th></tr>"
	Response.Write " <tr>"
	Response.Write " <td class=tablerow valign=middle>请输入您的用户名</td>"
	Response.Write " <td class=tablerow valign=middle><INPUT name=""username"" type=text class=inputbody></td>"
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write " <td class=tablerow colspan=2>"
	Response.Write " <b>说明：</b>本操作只能修改您的密码，不能对原密码进行修改，请确认您已经填写了密码问题及答案。</td></tr>"
	Response.Write " <tr>"
	Response.Write " <td class=tablerow valign=middle colspan=2 align=center><input type=submit name=""submit"" value="" 下一步 "" class=button></td></tr>"
	Response.Write "</table>"
	Response.Write "</form>"
End Sub

Private Sub sendusermail()
	On Error Resume Next
	Topic = "您在" & enchiasp.SiteName & "的密码信息"
	MailBody = MailBody & "<style>A:visited { TEXT-DECORATION: none }"
	MailBody = MailBody & "A:active  { TEXT-DECORATION: none }"
	MailBody = MailBody & "A:hover   { TEXT-DECORATION: underline overline }"
	MailBody = MailBody & "A:link    { text-decoration: none;}"
	MailBody = MailBody & "A:visited { text-decoration: none;}"
	MailBody = MailBody & "A:active  { TEXT-DECORATION: none;}"
	MailBody = MailBody & "A:hover   { TEXT-DECORATION: underline overline}"
	MailBody = MailBody & "BODY   { FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	MailBody = MailBody & "TD    { FONT-FAMILY: 宋体; FONT-SIZE: 9pt }</style>"
	MailBody = MailBody & "<TABLE border=0 width='95%' align=center><TBODY><TR>"
	MailBody = MailBody & "<TD valign=middle align=top>"
	MailBody = MailBody & "" & enchiasp.CheckStr(UserName) & "，您好：<br><br>"
	MailBody = MailBody & "欢迎您使用本系统的密码遗忘功能，<b>假如您不希望更改您的密码，请不要点击下面的激活连接</b>。<br>"
	MailBody = MailBody & "<a href=http://" & Request.servervariables("server_name") & Replace(Request.servervariables("script_name"), "sendpass.asp", "") & "activepass.asp?username=" & enchiasp.CheckStr(UserName) & "&pass=" & PassWord & "&repass=" & RePassword & "&answer=" & Answer & ">http://" & Request.servervariables("server_name") & Replace(Request.servervariables("script_name"), "sendpass.asp", "") & "activepass.asp?username=" & enchiasp.CheckStr(UserName) & "&pass=" & PassWord & "&repass=" & RePassword & "&answer=" & Answer & "</a>"
	MailBody = MailBody & "<br><br>"
	MailBody = MailBody & "<center><font color=red>再次感谢您使用本系统，让我们一起来建设这个网上家园！</font>"
	MailBody = MailBody & "</TD></TR></TBODY></TABLE><br><hr width=95% size=1>"
	Select Case CInt(enchiasp.SendMailType)
		Case 0
			SendMsg = "由于系统错误，给您发送的密码资料未成功。请点击右边的连接将您的密码激活：<a href=http://" & Request.servervariables("server_name") & Replace(Request.servervariables("script_name"), "sendpass.asp", "") & "activepass.asp?username=" & enchiasp.CheckStr(UserName) & "&pass=" & PassWord & "&repass=" & RePassword & "><B>激活密码</B></a>"
		Case 1
			Call jmail(UsereMail, Topic, MailBody)
		Case 2
			Call Cdonts(UsereMail, Topic, MailBody)
		Case 3
			Call aspemail(UsereMail, Topic, MailBody)
		Case Else
			SendMsg = "由于系统错误，给您发送的密码资料未成功。请点击右边的连接将您的密码激活：<a href=http://" & Request.servervariables("server_name") & Replace(Request.servervariables("script_name"), "sendpass.asp", "") & "activepass.asp?username=" & enchiasp.CheckStr(UserName) & "&pass=" & PassWord & "&repass=" & RePassword & "><B>激活密码</B></a>"
	End Select
End Sub
%>
<!--#include file="foot.inc"-->