<!--#include file="setup.asp" -->
<!--#include file="check.asp"-->
<!--#include file="../inc/chkinput.asp"-->
<!--#include file="../inc/email.asp" -->
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
Dim useremail, topic, mailbody, alluser, i,Action
i = 1
Action = LCase(Request("action"))
If Not ChkAdmin("MailList") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Set Rs = server.CreateObject ("adodb.recordset")
If Action = "send" Then
	Call send_mail()
ElseIf Request("action") = "sends" Then
	Call Send_Email()
ElseIf Request("action") = "mail" Then
	Call semail()
Else
	Call mail()
End If
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Sub mail()
	Response.Write "<form action=""admin_mailist.asp?action=send"" method=post>"& vbCrLf
	Response.Write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" width=""95%"" class=""tableBorder"" align=center>"& vbCrLf
	Response.Write "		<tr>"& vbCrLf
	Response.Write "		  <th colspan=""2"">系统邮件列表"& vbCrLf
	Response.Write "		  </th>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "		<tr>"& vbCrLf
	Response.Write "		  <td width=""15%"" class=TableRow1>注意事项：</td>"& vbCrLf
	Response.Write "		  <td class=TableRow1>在完整填写以下表单后点击发送，信息将发送到所有注册时完整填写了信箱的用户，邮件列表的使用将消耗大量的服务器资源，请慎重使用。</td>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "		<tr>"& vbCrLf
	Response.Write "		  <td class=TableRow1>邮件用户：</td>"& vbCrLf
	Response.Write "		  <td class=TableRow1><select name=Grade size=1>"& vbCrLf
	Set Rs = enchiasp.Execute("select * from ECCMS_UserGroup order by groupid")
	Rs.movenext
	Do While Not Rs.EOF
		Response.Write " <option value='" & Rs("Grades") & "'>" & Rs("GroupName") & "</option>"
		Rs.movenext
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write " <option value=''>所有用户</option>"
	Response.Write "		  </select>"& vbCrLf
	Response.Write "		  </td>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "		<tr>"& vbCrLf
	Response.Write "		  <td class=TableRow1>邮件标题：</td>"& vbCrLf
	Response.Write "		  <td class=TableRow1><input type=text name=topic size=60></td>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "		<tr>"& vbCrLf
	Response.Write "		  <td class=TableRow1>邮件内容：</td>"& vbCrLf
	Response.Write "		  <td class=TableRow1><textarea style=""width:100%;"" rows=10 name=""content""></textarea></td>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "		<tr>  <td class=TableRow1></td>"& vbCrLf
	Response.Write "		  <td height=20 class=TableRow1>"& vbCrLf
	Response.Write "		    &nbsp; <input type=""reset"" name=""Clear"" value=""清 除"" class=""button"">&nbsp; &nbsp; <input type=Submit value=""发送邮件"" name=Submit"" class=""button"">"& vbCrLf
	Response.Write "		  </td>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "	      </table>"& vbCrLf
	Response.Write "</form>"& vbCrLf
End Sub

Sub send_mail()
	If Request.Form("topic") = "" Then
		Errmsg = Errmsg + "<br>" + "<li>请输入邮件标题。"
		founderr = true
	Else
		topic = Request.Form("topic")
	End If
	If Request.Form("content") = "" Then
		Errmsg = Errmsg + "<br>" + "<li>请输入邮件内容。"
		founderr = true
	Else
		mailbody = Request.Form("content")
	End If
	If founderr = false Then
		On Error Resume Next
		If Len(Request.Form("Grade")) = 0 Then
			SQL = "select username,usermail from [ECCMS_User]"
		Else
			SQL = "select username,usermail from [ECCMS_User] where Grade = " & Request.Form("Grade")
		End If
		Rs.Open sql, conn, 1, 1
		If Not Rs.EOF And Not Rs.bof Then
			alluser = Rs.recordcount
			Do While Not Rs.EOF
				If Rs("usermail")<>"" Then
					useremail = Rs("usermail")
					If enchiasp.SendMailType = 0 Then
						errmsg = errmsg + "<br>" + "<li>本系统不支持发送邮件。"
						ReturnError(Errmsg)
						Exit Sub
					ElseIf enchiasp.SendMailType = 1 Then
						Call jmail(useremail, topic, mailbody)
					ElseIf enchiasp.SendMailType = 2 Then
						Call Cdonts(useremail, topic, mailbody)
					ElseIf enchiasp.SendMailType = 3 Then
						Call aspemail(useremail, topic, mailbody)
					End If
					i = i + 1
				End If
				Rs.movenext
			Loop
			If SendMail = "OK" Then
				Succeed("<li>成功发送"&i&"封邮件。")
			Else
				errmsg = errmsg + "<li>由于系统错误，邮件发送不成功。"
				ReturnError(Errmsg)
			End If
		End If
		Rs.Close
		Set Rs = Nothing
	End If
End Sub

Sub Send_Email()

	If Request("topic") = "" Then
		Errmsg = Errmsg + "<br>" + "<li>请输入邮件标题。"
		founderr = true
	Else
		topic = Request("topic")
	End If
	If Request("content") = "" Then
		Errmsg = Errmsg + "<br>" + "<li>请输入邮件内容。"
		founderr = true
	Else
		mailbody = Request("content")
	End If
	If Request("useremail") = "" Then
		Errmsg = Errmsg + "<br>" + "<li>请输入邮件地址。"
		founderr = true
		ElseIf IsValidEmail(Request("useremail")) = False Then
		Errmsg = Errmsg + "<br>" + "<li>你输入的Email有误，请重新输入。"
		founderr = true
	Else
		useremail = Request("useremail")
	End If
	If founderr = false Then
		If enchiasp.SendMailType = 0 Then
			errmsg = errmsg + "<br>" + "<li>本系统不支持发送邮件。"
			ReturnError(Errmsg)
			Exit Sub
		ElseIf enchiasp.SendMailType = 1 Then
			Call jmail(useremail, topic, mailbody)
		ElseIf enchiasp.SendMailType = 2 Then
			Call Cdonts(useremail, topic, mailbody)
		ElseIf enchiasp.SendMailType = 3 Then
			Call aspemail(useremail, topic, mailbody)
		End If
		If SendMail = "OK" Then
			Succeed("<li>你给 "&useremail&" 邮件成功发送。<li>主题："&topic&"")
		Else
			errmsg = errmsg + "<li>由于系统错误，邮件发送不成功。"
			ReturnError(Errmsg)
		End If
	End If
End Sub

Sub semail()
	Response.Write "<form action=""admin_mailist.asp?action=sends"" method=post>"& vbCrLf
	Response.Write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" width=""95%"" class=""tableBorder"" align=center>"& vbCrLf
	Response.Write "		<tr>"& vbCrLf
	Response.Write "		  <th colspan=""2"">系统邮件列表"& vbCrLf
	Response.Write "		  </th>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "		<tr>"& vbCrLf
	Response.Write "		  <td width=""15%"" class=TableRow1>电子邮件：</td>"& vbCrLf
	Response.Write "		  <td class=TableRow1><input type=text name=useremail value="""& vbCrLf
	Response.Write Request("useremail")
	Response.Write """ size=40></td>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "		<tr>"& vbCrLf
	Response.Write "		  <td class=TableRow1>邮件标题：</td>"& vbCrLf
	Response.Write "		  <td class=TableRow1><input type=text name=topic size=60></td>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "		<tr>"& vbCrLf
	Response.Write "		  <td class=TableRow1>邮件内容：</td>"& vbCrLf
	Response.Write "		  <td class=TableRow1><textarea style=""width:100%;"" rows=10 name=""content""></textarea></td>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "		<tr>  <td class=TableRow1></td>"& vbCrLf
	Response.Write "		  <td height=20 class=TableRow1>"& vbCrLf
	Response.Write "		    &nbsp; <input type=""reset"" name=""Clear"" value=""清 除"" class=""button"">&nbsp; &nbsp; <input type=Submit value=""发送邮件"" name=Submit"" class=""button"">"& vbCrLf
	Response.Write "		  </td>"& vbCrLf
	Response.Write "		</tr>"& vbCrLf
	Response.Write "	      </table>"& vbCrLf
	Response.Write "</form>"& vbCrLf
End Sub
%>
