<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
<!--#include file="../api/cls_api.asp"-->
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
Call InnerLocation("�޸Ļ�Ա����")

If CInt(GroupSetting(0)) = 0 Then
	ErrMsg = ErrMsg + "<li>�Բ�����û���޸������Ȩ�ޣ�����ʲô��������ϵ����Ա��</li>"
	Founderr = True
ElseIf LCase(Request("action")) = "save" Then
	Call ChangePassword
Else
%>
<script language="JavaScript">
<!--
function CheckForm() {
	if (document.myform.password.value.length == 0) {
		alert("����������ԭʼ����!");
		document.myform.password.focus();
		return false;
	}
	if (document.myform.password1.value.length == 0) {
		alert("����������������!");
		document.myform.password1.focus();
		return false;
	}
	if (document.myform.codestr.value.length != 4) {
		alert("��֤����������!");
		document.myform.codestr.focus();
		return false;
	}
	if (document.myform.password2.value.length == 0) {
		alert("����������ȷ������");
		document.myform.password2.focus();
		return false;
	}
		return true;
}
//-->
</script>
<table cellspacing=0 align=center cellpadding=0 width="98%" border=0> 
	<tr>
		<td>
			<form method="post" name=myform action="?action=save" onsubmit="return CheckForm();">
			<table cellspacing=1 align=center cellpadding=2 bgcolor=#cccccc border=0 class=Usertableborder>
			<tr>
				<th colspan=2>�޸�����</th>
			</tr>
				<tr>
					<td align=right width="38%" class=Usertablerow1 height=20>�û�����</td>
					<td width="62%" class=Usertablerow1> <strong class=userfont1><%=enchiasp.membername%></strong>
						<input type=hidden name=username value="<%=enchiasp.membername%>"><input type=hidden name=userid value="<%=enchiasp.memberid%>"></td></tr>
				<tr>
					<td align=right class=Usertablerow2 height=20>ԭʼ����(<font color=#ff6600>*</font>)��</td>
					<td class=Usertablerow2> <input class=inputbody type=password size=20 name=password></td>
				</tr>
				<tr>
					<td align=right class=Usertablerow1 height=20>������(<font color=#ff6600>*</font>)��</td>
					<td class=Usertablerow1> <input class=inputbody type=password size=20 name=password1></td>
				</tr>
				<tr bgcolor=#ffffff>
					<td align=right class=Usertablerow2 height=20>&nbsp;ȷ��������(<font color=#ff6600>*</font>)��</td>
					<td class=Usertablerow2> <input type=password class=inputbody size=20 name=password2> </td>
				</tr>
				<tr>
					<td align=right class=Usertablerow1 height=20>�� ֤ �룺</td>
					<td class=Usertablerow1> <input class=inputbody type=text size=6 name=codestr maxlength="6">&nbsp;<img src="../inc/getcode.asp" alt="��֤��,�������?����ˢ����֤��" onclick="this.src='../inc/getcode.asp'"> <font color=#808080>��������֤��</font></td>
				</tr>
				<tr>
					<td align=middle class=Usertablerow2 height=25>&nbsp; </td>
					<td class=Usertablerow2 align=center><input type=submit value=" ȷ �� " name=submit class=button></td>
				</tr>
			</table></form>
		</td>
	</tr>
</table>
<table align=center cellspacing=3 cellpadding=0 width="98%" border=0>
	<tr>
		<td width=15></td>
		<td><strong class=userfont2>ע�����</strong></td></tr>
	<tr>
	<tr>
		<td><img height=10 src="images/sword03.gif" width=10 align=absMiddle></td>
		<td>�û�����Ϊ�����������ʺ���վ��Կ�ף������Ʊ��ܺá�</td></tr>
	<tr>
		<td><img height=10 src="images/sword03.gif" width=10 align=absMiddle></td>
		<td>������ð�������,��ĸ�ͷ��š�ֻ�����ֵ��������ױ�����,����ȫ��</td></tr>
	<tr>
		<td><img height=10 src="images/sword03.gif" width=10 align=absMiddle></td>
		<td>ֻ�о�������ȷ�����޸ĳɹ�!</td></tr>
	<tr>
		<td><img height=10 src="images/sword03.gif" width=10 align=absMiddle></td>
		<td>����<font color=#ff6600>*</font>���ű��</td>
	</tr>
</table>
<br style="overflow: hidden; line-height: 5px">
<%
End If
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
Sub ChangePassword()
	On Error Resume Next
	Dim Rs,SQL,username, password,userid,newPassWord
	password = md5(Request.Form("password"))
	username = enchiasp.CheckBadstr(MemberName)
	userid = CLng(memberid)
	If enchiasp.IsValidPassword(Request.Form("password1")) = False Then
		ErrMsg = ErrMsg + "<li>�����к��зǷ��ַ���</li>"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request.Form("username")) <> username Then
		ErrMsg = ErrMsg + "<li>�Ƿ�������</li>"
		Founderr = True
	End If
	If enchiasp.IsValidStr(username) = False Then
		ErrMsg = ErrMsg + "<li>�û��к��зǷ��ַ���</li>"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request.Form("password")) = "" Then
		ErrMsg = ErrMsg + "<li>����û������ԭʼ���룡</li>"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request.Form("password1")) = "" And Trim(Request.Form("password2")) = "" Then
		ErrMsg = ErrMsg + "<li>�������벻��Ϊ�գ�</li>"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request.Form("password1")) <> Trim(Request.Form("password2")) Then
		ErrMsg = ErrMsg + "<li>������������ȷ�����벻һ�£�</li>"
		Founderr = True
		Exit Sub
	End If
	If Not enchiasp.CodeIsTrue() Then
		ErrMsg = ErrMsg + "<meta http-equiv=""refresh"" content=""2;URL=changeinfo.asp""><li>��֤��У��ʧ�ܣ��뷵��ˢ��ҳ�����ԡ�������Զ�����</li>"
		Session("GetCode") = ""
		Founderr = True
		Exit Sub
	End If
	Session("GetCode") = ""
	newPassWord = md5(Trim(Request.Form("password1")))
	Set Rs = enchiasp.Execute("SELECT * FROM [ECCMS_User] WHERE username='" & username & "' And userid=" & userid)
	If Rs.bof And Rs.EOF Then
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ����û���Ϣ��Ϣ��</li>"
		Founderr = True
		Exit Sub
	Else
		If password <> Rs("password") Then
			ErrMsg = ErrMsg + "<li>�������ԭʼ�������</li>"
			Founderr = True
			Exit Sub
		End If
	End If
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
			API_enchiasp.NodeValue "password",Trim(Request.form("password1")),1,False
			API_enchiasp.SendHttpData
			If API_enchiasp.Status = "1" Then
				Founderr = True
				ErrMsg = API_enchiasp.Message
			End If
			Set API_enchiasp = Nothing
		End If
		'-----------------------------------------------------------------
	End If
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	SQL = "SELECT password FROM [ECCMS_user] WHERE username='" & username & "' and userid=" & userid
	Rs.Open SQL, Conn, 1, 3
	Rs("password") = newPassWord
	Rs.Update
	Rs.Close
	Set Rs = Nothing
	Response.Cookies(enchiasp.Cookies_Name)("password") = newPassWord
	Call Returnsuc("<li>��ϲ���������޸ĳɹ���</li><li>���ס���������룺<font color=red>" & Request.Form("password2") & "</font></li>")
End Sub
%>
<!--#include file="foot.inc"-->