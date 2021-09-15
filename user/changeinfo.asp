<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/chkinput.asp"-->
<!--#include file="../api/cls_api.asp"-->
<!--#include file="head.inc"-->
<%
'=====================================================================
' 软件名称：恩池网站管理系统---修改会员资料
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================

Call InnerLocation("修改会员资料")

Dim Rs,SQL
If CInt(GroupSetting(1)) = 0 Then
	ErrMsg = ErrMsg + "<li>对不起！您没有修改用户资料的权限，如有什么问题请联系管理员。</li>"
	Founderr = True
ElseIf LCase(Request("action")) = "save" Then
	Call ChangeUserInfo
Else
	Set Rs = enchiasp.Execute("SELECT * FROM [ECCMS_User] WHERE username='" & enchiasp.membername & "' And userid=" & enchiasp.memberid)
	If (Rs.bof And Rs.EOF) Then
		ErrMsg = ErrMsg + "<li>Sorry！错误的系统参数！</li>"
		Founderr = True
	Else
%>
<script language="JavaScript">
<!--
function checkForm() {
	if (document.myform.password.value.length == 0) {
		alert("请输入您的用户密码!");
		document.myform.password.focus();
		return false;
	}
	if (document.myform.nickname.value.length == 0) {
		alert("请输入您的用户昵称!");
		document.myform.nickname.focus();
		return false;
	}
	if (document.myform.codestr.value.length != 4) {
		alert("验证码输入有误!");
		document.myform.codestr.focus();
		return false;
	}
	if (document.myform.usermail.value.length == 0) {
		alert("请输入您的E-mail");
		document.myform.usermail.focus();
		return false;
	}
		return true;
}
//-->
</script>
<table cellspacing=1 align=center cellpadding=2 bgcolor=#cccccc border=0 class=Usertableborder>
	<tr>
		<th colspan=2>修改个人资料</th>
	</tr>
	<form method="post" name=myform action="?action=save" onsubmit="return checkForm();">
	<tr>
		<td align=right width="25%" class=Usertablerow1 height=20>用户名：</td>
		<td width="75%" class=Usertablerow1> <strong class=userfont1><%=enchiasp.membername%></strong>
			<input type=hidden name=username value="<%=Server.HTMLEncode(Rs("username"))%>"><input type=hidden name=userid value="<%=enchiasp.memberid%>"></td>
	</tr>
	<tr>
		<td align=right class=Usertablerow2 height=20>用户昵称(<span class=userfont1>*</span>)：</td>
		<td class=Usertablerow2> <input class=inputbody type=text size=20 name=nickname value="<%=enchiasp.HTMLEncodes(Rs("nickname"))%>" maxlength="15"></td>
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>真实姓名(<span class=userfont1>*</span>)：</td>
		<td class=Usertablerow1> <input class=inputbody type=text size=20 name=TrueName value="<%=enchiasp.HTMLEncodes(Rs("TrueName"))%>" maxlength="15"></td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align=right class=Usertablerow2 height=20>&nbsp;用户邮箱(<span class=userfont1>*</span>)：</td>
		<td class=Usertablerow2> <input type=text class=inputbody size=30 name=usermail value="<%=enchiasp.HTMLEncodes(Rs("usermail"))%>" maxlength="50"> <span class=userfont1>注意：</span><font color=#808080>请填写你常用的邮箱</font></td>
	
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>性别：</td>
		<td class=Usertablerow1> <input type=radio name=usersex value="男"<%If Trim(Rs("usersex")) = "男" Then Response.Write " checked"%>> 男&nbsp;&nbsp;    
			<input type=radio name=usersex value="女"<%If Trim(Rs("usersex")) = "女" Then Response.Write " checked"%>> 女&nbsp;&nbsp;    
			<input type=radio name=usersex value="女"<%If Trim(Rs("usersex")) = "保密" Then Response.Write " checked"%>> 保密</td>    
	</tr>
	<tr>
		<td align=right class=Usertablerow2 height=20>密码提示问题(<span class=userfont1>*</span>)：</td>
		<td class=Usertablerow2> <input class=inputbody type=text size=30 name=question value="<%=enchiasp.HTMLEncodes(Rs("question"))%>" maxlength="35"> <select onChange="question.value=this.value;")>    
			<option value="" selected>[请选择]</option>
			<option value="最喜欢的宠物？">最喜欢的宠物？</option>
			<option value="最喜爱的电影？">最喜爱的电影？</option>
			<option value="周年纪念日 [年/月/日]？">周年纪念日 
            [年/月/日]？</option>
			<option value="父亲的名字？">父亲的名字？</option>
			<option value="配偶的名字？">配偶的名字？</option>
			<option value="第一个孩子的爱称？">第一个孩子的爱称？</option>
			<option value="中学的校名？">中学的校名？</option>
			<option value="最尊敬的老师？">最尊敬的老师？</option>
			<option value="最喜欢的运行队？">最喜欢的运行队？</option>
			</select></td>
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>密码问题答案：</td>
		<td class=Usertablerow1> <input class=inputbody type=text size=30 name=answer maxlength="35"> <font color=#808080>忘记密码的提示问题答案，用于取回密码</font></td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align=right class=Usertablerow2 height=20>联系电话：</td>
		<td class=Usertablerow2> <input type=text class=inputbody size=20 name=phone value="<%=enchiasp.HTMLEncodes(Rs("phone"))%>" maxlength="20"> <font color=#808080>如：+86-27-85188888</font></td>    
	
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>你的OICQ：</td>
		<td class=Usertablerow1> <input class=inputbody type=text size=20 name=oicq value="<%=enchiasp.HTMLEncodes(Rs("oicq"))%>" maxlength="20"></td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align=right class=Usertablerow2 height=20>邮政编码：</td>
		<td class=Usertablerow2> <input type=text class=inputbody size=20 name=postcode value="<%=enchiasp.HTMLEncodes(Rs("postcode"))%>" maxlength="20"></td>
	
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>身份证：</td>
		<td class=Usertablerow1> <input class=inputbody type=text size=30 name=UserIDCard value="<%=enchiasp.HTMLEncodes(Rs("UserIDCard"))%>" maxlength="35"></td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align=right class=Usertablerow2 height=20>联系地址：</td>
		<td class=Usertablerow2> <input type=text class=inputbody size=50 name=address value="<%=enchiasp.HTMLEncodes(Rs("address"))%>" maxlength="50"></td>
	
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>交易密码：</td>
		<td class=Usertablerow1> <input class=inputbody type=text size=30 name=BuyCode maxlength="35"> <font color=#808080>站内支付所用的交易密码</font></td>
	</tr>
	<tr>
		<td align=right class=Usertablerow2 height=20>个人主页：</td>
		<td class=Usertablerow2> <input class=inputbody type=text size=30 name=HomePage value="<%=enchiasp.HTMLEncodes(Rs("HomePage"))%>" maxlength="35"> <font color=#808080>以“http://”开头</font></td>    
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>用户密码：</td>
		<td class=Usertablerow1> <input class=inputbody type=password size=30 name=password value="" maxlength="50"> <span class=userfont1>输入正确的密码才能修改用户资料</span></td>
	</tr>
	<tr>
		<td align=right class=Usertablerow2 height=20>验 证 码：</td>
		<td class=Usertablerow2> <input class=inputbody type=text size=6 name=codestr maxlength="6">&nbsp;<img src="../inc/getcode.asp" alt="验证码,看不清楚?请点击刷新验证码" onclick="this.src='../inc/getcode.asp'"> <font color=#808080>请输入验证码</font></td>    
	</tr>
	<tr>
		<td align=middle class=Usertablerow2 height=20>&nbsp; </td>
		<td class=Usertablerow2 align=center><input type=submit value=" 确 认 " name="submit" class="button"></td>
	</tr></form>
</table>
<%
	End If
	Rs.Close:Set Rs = Nothing
End If
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
Sub ChangeUserInfo()
	On Error Resume Next
	Dim username, password,userid
	Dim usersex,sex
	username = enchiasp.CheckBadstr(enchiasp.membername)
	userid = enchiasp.ChkNumeric(memberid)
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If enchiasp.IsValidPassword(Request.Form("answer")) = False And Trim(Request.Form("answer")) <> "" Then
		ErrMsg = ErrMsg + "<li>密码问题答案中含有非法字符！</li>"
		Founderr = True
	End If
	If Trim(Request.Form("username")) <> username Then
		ErrMsg = ErrMsg + "<li>非法操作！</li>"
		Founderr = True
	End If
	If Trim(Request.Form("password")) = "" Then
		ErrMsg = ErrMsg + "<li>请输入用户密码！</li>"
		Founderr = True
	Else
		password = md5(Request.Form("password"))
	End If
	If userid = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	
	If Trim(Request.Form("nickname")) = "" Then
		ErrMsg = ErrMsg + "<li>用户昵称不能为空！</li>"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request.Form("nickname")) = False Then
		ErrMsg = ErrMsg + "<li>用户昵称中含有非法字符！</li>"
		Founderr = True
	End If
	If Trim(Request.Form("TrueName")) = "" Then
		ErrMsg = ErrMsg + "<li>真实姓名不能为空！</li>"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request.Form("TrueName")) = False Then
		ErrMsg = ErrMsg + "<li>真实姓名中含有非法字符！</li>"
		Founderr = True
	End If
	If Trim(Request.Form("usermail")) = "" Then
		ErrMsg = ErrMsg + "<li>您的Email不能为空！</li>"
		Founderr = True
	End If
	If IsValidEmail(Request.Form("usermail")) = False Then
		ErrMsg = ErrMsg + "<li>您的Email有错误！</li>"
		Founderr = True
	End If
	If Not IsNumeric(Request.Form("oicq")) And Trim(Request.Form("oicq")) <> "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>QQ号码请用数字填写。</li>"
	End If
	If Trim(Request.Form("HomePage")) <> "" And Left(Request.Form("HomePage"),7) <> "http://" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>个人主页地址输入有误，请以“http://”开头。</li>"
	End If
	If Not enchiasp.CodeIsTrue() Then
		ErrMsg = ErrMsg + "<meta http-equiv=""refresh"" content=""2;URL=changeinfo.asp""><li>验证码校验失败，请返回刷新页面再试。两秒后自动返回</li>"
		Session("GetCode") = ""
		Founderr = True
		Exit Sub
	End If
	Session("GetCode") = ""
	If Trim(Request.Form("usersex")) = "" Then
		ErrMsg = ErrMsg + "<li>您的姓别不能为空！</li>"
		Founderr = True
	Else
		usersex = enchiasp.CheckBadstr(Request.Form("usersex"))
	End If
	If usersex = "女" Then
		sex = 0
	Else
		sex = 1
	End If
	
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	SQL = "SELECT * FROM [ECCMS_user] WHERE username='" & username & "' And userid=" & CLng(userid)
	Rs.Open SQL, Conn, 1, 3
	If Rs.bof And Rs.EOF Then
		ErrMsg = ErrMsg + "<li>Sorry！没有找到此用户信息信息！</li>"
		Founderr = True
		Exit Sub
	Else
		If password <> Rs("password") Then
			ErrMsg = ErrMsg + "<li>您输入的密码错误！</li>"
			Founderr = True
			Exit Sub
		End If
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
				API_enchiasp.NodeValue "password","",1,False
				API_enchiasp.NodeValue "answer",Request.Form("answer"),1,False
				API_enchiasp.NodeValue "question",Request.Form("question"),1,False
				API_enchiasp.NodeValue "email",Request.Form("usermail"),1,False
				API_enchiasp.NodeValue "gender",sex,0,False
				API_enchiasp.SendHttpData
				If API_enchiasp.Status = "1" Then
					Founderr = True
					ErrMsg = API_enchiasp.Message
					Exit Sub
				End If
				Set API_enchiasp = Nothing
			End If
			'-----------------------------------------------------------------
		End If
		Rs("nickname") = enchiasp.CheckBadstr(Request.Form("nickname"))
		Rs("TrueName") = enchiasp.CheckBadstr(Request.Form("TrueName"))
		Rs("usermail") = Trim(Request.Form("usermail"))
		If Trim(Request.Form("HomePage")) <> "" Then Rs("HomePage") = enchiasp.CheckBadstr(Trim(Request.Form("HomePage")))
		If Trim(Request.Form("usersex")) <> "" Then Rs("usersex") = usersex
		If Trim(Request.Form("question")) <> "" Then Rs("question") =enchiasp.CheckBadstr(Trim(Request.Form("question"))) 
		If Trim(Request.Form("answer")) <> "" Then Rs("answer") = md5(Trim(Request.Form("answer")))
		If Trim(Request.Form("phone")) <> "" Then Rs("phone") = enchiasp.CheckBadstr(Trim(Request.Form("phone")))
		If Trim(Request.Form("oicq")) <> "" Then Rs("oicq") = enchiasp.CheckBadstr(Trim(Request.Form("oicq")))
		If Trim(Request.Form("postcode")) <> "" Then Rs("postcode") = enchiasp.CheckBadstr(Trim(Request.Form("postcode")))
		If Trim(Request.Form("UserIDCard")) <> "" Then Rs("UserIDCard") = enchiasp.CheckBadstr(Trim(Request.Form("UserIDCard")))
		If Trim(Request.Form("address")) <> "" Then Rs("address") = enchiasp.CheckBadstr(Trim(Request.Form("address")))

		If Trim(Request.Form("BuyCode")) <> "" Then Rs("BuyCode") = md5(Trim(Request.Form("BuyCode")))
		Rs.Update
	End If
	Rs.Close
	Set Rs = Nothing
	Call Returnsuc("<li>恭喜您！用户资料修改成功。</li>")
End Sub
%>
<!--#include file="foot.inc"-->