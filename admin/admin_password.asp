<!--#include file="setup.asp"-->
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
Dim ID
Response.Write "<script language=""JavaScript"">" & vbCrLf
Response.Write "<!--" & vbCrLf
Response.Write "function CheckForm()" & vbCrLf
Response.Write "{" & vbCrLf
Response.Write "if (document.myform.password.value.length == 0)" & vbCrLf
Response.Write " {" & vbCrLf
Response.Write "alert(""����������ԭʼ����!"");" & vbCrLf
Response.Write "document.myform.password.focus();" & vbCrLf
Response.Write "return false;" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "if (document.myform.password1.value.length == 0)" & vbCrLf
Response.Write " {" & vbCrLf
Response.Write "alert(""����������������!"");" & vbCrLf
Response.Write "document.myform.password1.focus();" & vbCrLf
Response.Write "return false;" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "if (document.myform.password2.value.length == 0)" & vbCrLf
Response.Write " {" & vbCrLf
Response.Write "alert(""����������ȷ������"");" & vbCrLf
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
	Response.Write " <th colspan=""2"">����Ա���Ƽ������޸�</th></tr>" & vbCrLf
	Response.Write "<form method=Post name=""myform"" action=""admin_password.asp?action=save"" onSubmit=""return CheckForm();"">" & vbCrLf
	Response.Write " <tr> " & vbCrLf
	Response.Write " <td width=""25%"" align=""right"" nowrap class=""tablerow2"">����Ա���ƣ�</td>" & vbCrLf
	Response.Write " <td class=""tablerow1"" width=""75%"">��<INPUT type=text size=25 name=username value="""
	Response.Write Session("AdminName")
	Response.Write """> * <font COLOR=#FF0000>���޸Ŀ�������</font></td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr> " & vbCrLf
	Response.Write " <td align=""right"" nowrap class=""tablerow2"">ԭʼ���룺</td>" & vbCrLf
	Response.Write " <td class=""tablerow1"">��<INPUT type=password size=25 name=password></td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr> " & vbCrLf
	Response.Write " <td align=""right"" nowrap class=""tablerow2"">�����룺</td>" & vbCrLf
	Response.Write " <td class=""tablerow1"">��<INPUT type=password size=25 name=password1></td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr> " & vbCrLf
	Response.Write " <td align=""right"" nowrap class=""tablerow2"">ȷ�������룺</td>" & vbCrLf
	Response.Write " <td class=""tablerow1"">��<INPUT type=password size=25 name=password2></td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr> " & vbCrLf
	Response.Write " <td align=""center"" colspan=""2"" class=""tablerow1"">" & vbCrLf
	Response.Write "<INPUT type=hidden name=id value="""
	Response.Write Session("Adminid")
	Response.Write """>" & vbCrLf
	Response.Write "<input type=""submit"" name=""Submit"" class=button value=""ȷ���޸�"">��" & vbCrLf
	Response.Write "</td>" & vbCrLf
	Response.Write " </tr></form>" & vbCrLf
	Response.Write "</table><BR>" & vbCrLf
End Sub


Private Sub svaeadmin()
	Dim password
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	password = md5(Request.Form("password"))
	If enchiasp.checkpost = False Then
		ErrMsg = ErrMsg + "<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύע�ᡣ</li>"
		founderr = True
	End If
	If InStr(Request("username"), "=") > 0 Or InStr(Request("username"), "%") > 0 Or InStr(Request("username"), Chr(32)) > 0 Or InStr(Request("username"), "?") > 0 Or InStr(Request("username"), "&") > 0 Or InStr(Request("username"), ";") > 0 Or InStr(Request("username"), ",") > 0 Or InStr(Request("username"), "'") > 0 Or InStr(Request("username"), ",") > 0 Or InStr(Request("username"), Chr(34)) > 0 Or InStr(Request("username"), Chr(9)) > 0 Or InStr(Request("username"), "��") > 0 Or InStr(Request("username"), "$") > 0 Then
		ErrMsg = ErrMsg + "<br>" + "<li>�û����к��зǷ��ַ���</li>"
		founderr = True
	End If
	If InStr(Request("password1"), "=") > 0 Or InStr(Request("password1"), "+") > 0 Or InStr(Request("password1"), "&") > 0 Or InStr(Request("password1"), "'") > 0 Or InStr(Request("password1"), " ") > 0 Or InStr(Request("password1"), "%") > 0 Then
		ErrMsg = ErrMsg + "<li>�����к��зǷ��ַ� </li>"
		founderr = True
	End If
	If Request.Form("password") = "" Then
		ErrMsg = ErrMsg + "<li>����û������ԭʼ���롣<li>"
		founderr = True
	End If
	If Request.Form("password1") = "" And Request.Form("password2") = "" Then
		ErrMsg = ErrMsg + "<li>�������벻��Ϊ�ա�</li>"
		founderr = True
	End If
	If Request.Form("password1") <> Request.Form("password2") Then
		ErrMsg = ErrMsg + "<li>������������ȷ�����벻һ�¡�</li>"
		founderr = True
	End If
	Rs.Open "Select * from ECCMS_Admin where username = '" & Session("AdminName") & "' and id = " & Session("Adminid") & "", conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "Sorry��û���ҵ����û���Ϣ��Ϣ��"
	Else
		If password <> Rs("password") Then
			ErrMsg = ErrMsg + "<li>�������ԭʼ�������</i>"
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
		Succeed("<li>����Ա�޸ĳɹ�!</li>")
	End If
End Sub

%>
