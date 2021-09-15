<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->

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
Dim ID
Response.Write "<script language=""JavaScript"">" & vbCrLf
Response.Write "<!--" & vbCrLf
Response.Write "function CheckForm()" & vbCrLf
Response.Write "{" & vbCrLf
Response.Write "if (document.myform.password.value.length == 0)" & vbCrLf
Response.Write " {" & vbCrLf
Response.Write "alert(""请输入您的原始密码!"");" & vbCrLf
Response.Write "document.myform.password.focus();" & vbCrLf
Response.Write "return false;" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "if (document.myform.password1.value.length == 0)" & vbCrLf
Response.Write " {" & vbCrLf
Response.Write "alert(""请输入您的新密码!"");" & vbCrLf
Response.Write "document.myform.password1.focus();" & vbCrLf
Response.Write "return false;" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "if (document.myform.password2.value.length == 0)" & vbCrLf
Response.Write " {" & vbCrLf
Response.Write "alert(""请输入您的确认密码"");" & vbCrLf
Response.Write "document.myform.password2.focus();" & vbCrLf
Response.Write "return false;" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "return true;" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "//-->"
Response.Write "</script>" & vbCrLf
Admin_header
Dim Action
Action = LCase(Request("action"))
'If Not ChkAdmin("ChangePassword") Then
	'Server.Transfer("showerr.asp")
	'Response.End
'End If
Set Rs = Server.CreateObject("adodb.recordset")
Select Case Action
	Case "save"
		Call svaeadmin
	Case Else
		Call PassMain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Private Sub PassMain()
	Response.Write "<table border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class=""tableBorder"">" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th colspan=""2"">管理员名称及密码修改</th></tr>" & vbCrLf
	Response.Write "<form method=Post name=""myform"" action=""admin_password.asp?action=save"" onSubmit=""return CheckForm();"">" & vbCrLf
	Response.Write " <tr> " & vbCrLf
	Response.Write " <td width=""25%"" align=""right"" nowrap class=""tablerow2"">管理员名称：</td>" & vbCrLf
	Response.Write " <td class=""tablerow1"" width=""75%"">　<INPUT type=text size=25 name=username value="""
	Response.Write Session("AdminName")
	Response.Write """> * <font COLOR=#FF0000>不修改可以留空</font></td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr> " & vbCrLf
	Response.Write " <td align=""right"" nowrap class=""tablerow2"">原始密码：</td>" & vbCrLf
	Response.Write " <td class=""tablerow1"">　<INPUT type=password size=25 name=password></td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr> " & vbCrLf
	Response.Write " <td align=""right"" nowrap class=""tablerow2"">新密码：</td>" & vbCrLf
	Response.Write " <td class=""tablerow1"">　<INPUT type=password size=25 name=password1></td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr> " & vbCrLf
	Response.Write " <td align=""right"" nowrap class=""tablerow2"">确认新密码：</td>" & vbCrLf
	Response.Write " <td class=""tablerow1"">　<INPUT type=password size=25 name=password2></td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr> " & vbCrLf
	Response.Write " <td align=""center"" colspan=""2"" class=""tablerow1"">" & vbCrLf
	Response.Write "<INPUT type=hidden name=id value="""
	Response.Write Session("Adminid")
	Response.Write """>" & vbCrLf
	Response.Write "<input type=""submit"" name=""Submit"" class=button value=""确认修改"">　" & vbCrLf
	Response.Write "</td>" & vbCrLf
	Response.Write " </tr></form>" & vbCrLf
	Response.Write "</table><BR>" & vbCrLf
End Sub


Private Sub svaeadmin()
	Dim password
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	password = md5(Request.Form("password"))
	If enchiasp.checkpost = False Then
		ErrMsg = ErrMsg + "<li>您提交的数据不合法，请不要从外部提交注册。</li>"
		founderr = True
	End If
	If InStr(Request("username"), "=") > 0 Or InStr(Request("username"), "%") > 0 Or InStr(Request("username"), Chr(32)) > 0 Or InStr(Request("username"), "?") > 0 Or InStr(Request("username"), "&") > 0 Or InStr(Request("username"), ";") > 0 Or InStr(Request("username"), ",") > 0 Or InStr(Request("username"), "'") > 0 Or InStr(Request("username"), ",") > 0 Or InStr(Request("username"), Chr(34)) > 0 Or InStr(Request("username"), Chr(9)) > 0 Or InStr(Request("username"), "") > 0 Or InStr(Request("username"), "$") > 0 Then
		ErrMsg = ErrMsg + "<br>" + "<li>用户名中含有非法字符。</li>"
		founderr = True
	End If
	If InStr(Request("password1"), "=") > 0 Or InStr(Request("password1"), "+") > 0 Or InStr(Request("password1"), "&") > 0 Or InStr(Request("password1"), "'") > 0 Or InStr(Request("password1"), " ") > 0 Or InStr(Request("password1"), "%") > 0 Then
		ErrMsg = ErrMsg + "<li>密码中含有非法字符 </li>"
		founderr = True
	End If
	If Request.Form("password") = "" Then
		ErrMsg = ErrMsg + "<li>您还没有输入原始密码。<li>"
		founderr = True
	End If
	If Request.Form("password1") = "" And Request.Form("password2") = "" Then
		ErrMsg = ErrMsg + "<li>您的密码不能为空。</li>"
		founderr = True
	End If
	If Request.Form("password1") <> Request.Form("password2") Then
		ErrMsg = ErrMsg + "<li>您输入的密码和确认密码不一致。</li>"
		founderr = True
	End If
	Rs.Open "Select * from ECCMS_Admin where username = '" & Session("AdminName") & "' and id = " & Session("Adminid") & "", conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "Sorry！没有找到此用户信息信息。"
	Else
		If password <> Rs("password") Then
			ErrMsg = ErrMsg + "<li>您输入的原始密码错误。</i>"
			founderr = True
			Exit Sub
		End If
	End If
	Rs.Close
	If founderr = True Then Exit Sub
	If founderr = False Then
		SQL = "select * from ECCMS_Admin where id = " & Request("id")
		Rs.Open SQL, conn, 1, 3
		Rs("password") = md5(Request.Form("password1"))
		If Request.Form("username") <> "" Then
			Rs("username") = Request.Form("username")
		End If
		Rs.update
			Session("AdminPass") = Rs("password")
			Session("AdminName") = Rs("username")
		Rs.Close
		Set Rs = Nothing
		Succeed("<li>管理员修改成功!</li>")
	End If
End Sub

%>
