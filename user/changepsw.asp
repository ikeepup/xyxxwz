<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
<!--#include file="../api/cls_api.asp"-->
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
Call InnerLocation("修改会员密码")

If CInt(GroupSetting(0)) = 0 Then
	ErrMsg = ErrMsg + "<li>对不起！您没有修改密码的权限，如有什么问题请联系管理员。</li>"
	Founderr = True
ElseIf LCase(Request("action")) = "save" Then
	Call ChangePassword
Else
%>
<script language="JavaScript">
<!--
function CheckForm() {
	if (document.myform.password.value.length == 0) {
		alert("请输入您的原始密码!");
		document.myform.password.focus();
		return false;
	}
	if (document.myform.password1.value.length == 0) {
		alert("请输入您的新密码!");
		document.myform.password1.focus();
		return false;
	}
	if (document.myform.codestr.value.length != 4) {
		alert("验证码输入有误!");
		document.myform.codestr.focus();
		return false;
	}
	if (document.myform.password2.value.length == 0) {
		alert("请输入您的确认密码");
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
				<th colspan=2>修改密码</th>
			</tr>
				<tr>
					<td align=right width="38%" class=Usertablerow1 height=20>用户名：</td>
					<td width="62%" class=Usertablerow1> <strong class=userfont1><%=enchiasp.membername%></strong>
						<input type=hidden name=username value="<%=enchiasp.membername%>"><input type=hidden name=userid value="<%=enchiasp.memberid%>"></td></tr>
				<tr>
					<td align=right class=Usertablerow2 height=20>原始密码(<font color=#ff6600>*</font>)：</td>
					<td class=Usertablerow2> <input class=inputbody type=password size=20 name=password></td>
				</tr>
				<tr>
					<td align=right class=Usertablerow1 height=20>新密码(<font color=#ff6600>*</font>)：</td>
					<td class=Usertablerow1> <input class=inputbody type=password size=20 name=password1></td>
				</tr>
				<tr bgcolor=#ffffff>
					<td align=right class=Usertablerow2 height=20>&nbsp;确认新密码(<font color=#ff6600>*</font>)：</td>
					<td class=Usertablerow2> <input type=password class=inputbody size=20 name=password2> </td>
				</tr>
				<tr>
					<td align=right class=Usertablerow1 height=20>验 证 码：</td>
					<td class=Usertablerow1> <input class=inputbody type=text size=6 name=codestr maxlength="6">&nbsp;<img src="../inc/getcode.asp" alt="验证码,看不清楚?请点击刷新验证码" onclick="this.src='../inc/getcode.asp'"> <font color=#808080>请输入验证码</font></td>
				</tr>
				<tr>
					<td align=middle class=Usertablerow2 height=25>&nbsp; </td>
					<td class=Usertablerow2 align=center><input type=submit value=" 确 认 " name=submit class=button></td>
				</tr>
			</table></form>
		</td>
	</tr>
</table>
<table align=center cellspacing=3 cellpadding=0 width="98%" border=0>
	<tr>
		<td width=15></td>
		<td><strong class=userfont2>注意事项：</strong></td></tr>
	<tr>
	<tr>
		<td><img height=10 src="images/sword03.gif" width=10 align=absMiddle></td>
		<td>用户密码为您管理您的帐号网站的钥匙，请妥善保管好。</td></tr>
	<tr>
		<td><img height=10 src="images/sword03.gif" width=10 align=absMiddle></td>
		<td>密码最好包括数字,字母和符号。只有数字的密码容易被猜破,不安全。</td></tr>
	<tr>
		<td><img height=10 src="images/sword03.gif" width=10 align=absMiddle></td>
		<td>只有旧密码正确才能修改成功!</td></tr>
	<tr>
		<td><img height=10 src="images/sword03.gif" width=10 align=absMiddle></td>
		<td>带“<font color=#ff6600>*</font>”号必填。</td>
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
		ErrMsg = ErrMsg + "<li>密码中含有非法字符！</li>"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request.Form("username")) <> username Then
		ErrMsg = ErrMsg + "<li>非法操作！</li>"
		Founderr = True
	End If
	If enchiasp.IsValidStr(username) = False Then
		ErrMsg = ErrMsg + "<li>用户中含有非法字符！</li>"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request.Form("password")) = "" Then
		ErrMsg = ErrMsg + "<li>您还没有输入原始密码！</li>"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request.Form("password1")) = "" And Trim(Request.Form("password2")) = "" Then
		ErrMsg = ErrMsg + "<li>您的密码不能为空！</li>"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request.Form("password1")) <> Trim(Request.Form("password2")) Then
		ErrMsg = ErrMsg + "<li>您输入的密码和确认密码不一致！</li>"
		Founderr = True
		Exit Sub
	End If
	If Not enchiasp.CodeIsTrue() Then
		ErrMsg = ErrMsg + "<meta http-equiv=""refresh"" content=""2;URL=changeinfo.asp""><li>验证码校验失败，请返回刷新页面再试。两秒后自动返回</li>"
		Session("GetCode") = ""
		Founderr = True
		Exit Sub
	End If
	Session("GetCode") = ""
	newPassWord = md5(Trim(Request.Form("password1")))
	Set Rs = enchiasp.Execute("SELECT * FROM [ECCMS_User] WHERE username='" & username & "' And userid=" & userid)
	If Rs.bof And Rs.EOF Then
		ErrMsg = ErrMsg + "<li>Sorry！没有找到此用户信息信息！</li>"
		Founderr = True
		Exit Sub
	Else
		If password <> Rs("password") Then
			ErrMsg = ErrMsg + "<li>您输入的原始密码错误！</li>"
			Founderr = True
			Exit Sub
		End If
	End If
	Rs.Close:Set Rs = Nothing
	If Founderr = False Then
		'-----------------------------------------------------------------
		'系统整合
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
	Call Returnsuc("<li>恭喜您！密码修改成功。</li><li>请记住您的新密码：<font color=red>" & Request.Form("password2") & "</font></li>")
End Sub
%>
<!--#include file="foot.inc"-->