<!--#include file="setup.asp"-->

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
Response.CacheControl = "no-cache"
Dim RefreshTime,GetCode
FoundErr = False
RefreshTime = 3 '设置防刷新时间
If DateDiff("s", Session("UserTime"), Now()) < RefreshTime Then
	Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT="&RefreshTime&"><br>本页面起用了防刷新机制，请不要在"&RefreshTime&"秒内连续刷新本页面<BR>正在打开页面，请稍后……"
	Response.End
End If
FoundErr = False
Select Case enchiasp.CheckStr(Request("action"))
	Case "logout" '退出系统
		Call logout()
	Case "login" '登陆系统
		Call chklogin()
	Case Else
		if enchiasp.ercilogin ="1" then
		'转向特殊登陆页面
			session("mypasskey")=enchiasp.mypasskey
	 		response.redirect "admin_loginx.asp" 
		 else
			 Call main()
		end if
		
End Select

If Founderr = True Then
	Session("UserTime") = Now()
	SaveLogInfo("非法登陆！")
	Response.Redirect("showerr.asp?action=error&message=" & Server.URLEncode(ErrMsg) & "")
End If
CloseConn

Sub main()
	
	If Session("AdminName") = "" Then
%>
<html>
<head>
<title>管理员登陆</title>
<meta http-equiv="Content-Type" content="text/html; chaRset=gb2312">
<link rel="stylesheet" href="images/admin.css" type="text/css">
</head>
<script language="javascript">
<!--//
function SetFocus()
{
if (document.myform.AdminName.value=="")
	document.myform.AdminName.focus();
else
	document.myform.AdminName.select();
}
function CheckForm()
{
	if(document.myform.AdminName.value=="")
	{
		alert("请输入您的用户名！");
		document.myform.AdminName.focus();
		return false;
	}
	if(document.myform.PassWord.value == "")
	{
		alert("请输入您的密码！");
		document.myform.PassWord.focus();
		return false;
	}
	if (document.myform.verifycode.value==""){
       alert ("请输入您的验证码！");
       document.myform.verifycode.focus();
       return(false);
    }
}
function CheckBrowser()
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("系统友情提示：\n    你使用的是Netscape浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  }
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("系统友情提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  }
}
//-->
</script>
<body topmargin="0" leftmargin="0" rightmargin="0">
<script language="JavaScript" src="keyboard.js" type="text/javascript"></script>
<div align="center"><BR>
  <p>　</p>  <p>　</p>
  <form name=myform method="post" action="admin_klogin.asp?action=login" target="_top" onSubmit="return CheckForm();">

<table border="0" cellpadding="0" cellspacing="0" width="642" background="images/adminlogin.jpg" height="425" align="center">
  <tr>
    <td width="100%" height="370">
      <table border="0" cellpadding="0" cellspacing="0" width="100%" height="401">
        <tr>
          <td width="50%" height="401"></td>
          <td width="50%" height="401">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
              <tr>
                <td width="100%"></td>
              </tr>
              <tr>
                <td width="100%"></td>
              </tr>
              <tr>
                <td width="100%"></td>
              </tr>
              <tr>
                <td width="100%">　
                  <p>　</p>
                  <p>　</p>
                  <p>　</p>
                  <p>　</p>
                  <p>　</td>
              </tr>
              <tr>
               <TD height=25 align=center><form name="form1" method="post" action="admin_klogin.asp?action=login" onsubmit="return login()">
用户名：<input type="text" name="AdminName"  style=width:150px autocomplete="off"  class="ycenchi"><input onclick="showkeyboard('Login.AdminName')" type="button" value="键盘" title="用软键盘输入密码，防止黑客软件记录键盘的录入信息" />
</td></tr>
<tr><td height=25 align=center>
密&nbsp; 码：<input type="password" name="Password"  style=width:150px  class="ycenchi"><input onclick="showkeyboard('Login.AdminName')" type="button" value="键盘" title="用软键盘输入密码，防止黑客软件记录键盘的录入信息" />     
</td></tr>
<tr><td height=25 align=center>
认证码：<input name="verifycode" type="text" size="12" maxlength="9" class="ycenchi">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="../inc/getcode.asp"  id="GetCodePic" align=absmiddle height=16 border=0></td></tr><tr><td height=25 align=center>
<input type="submit" name="Submit" value="登 录" class="adminbutton">&nbsp;<input type="reset" name="reset" value="清 空" class="adminbutton">
</TD>
              </tr>
              <tr>
                <td width="100%"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>


</form>
<script language="JavaScript">
<!--
CheckBrowser();
SetFocus();
-->
</script>
<p align=center></p>
</div>
<%
Else
	Response.Redirect "admin_kindex.asp"
End If
End Sub

Sub logout()
	'清除COOKIES中管理员身份的验证信息.
	Session.Abandon
	Session("AdminName") = ""
	Session("AdminPass") = ""
	Session("AdminGrade") = ""
	Session("AdminFlag") = ""
	Session("AdminStatus") = ""
	Session("AdminID") = ""
	Session("AdminRandomCode") = ""
	Response.Cookies(Admin_Cookies_Name) = ""
	Response.Redirect ("../")
End Sub

Sub chklogin()
	Dim adminname, password,RandomCode,mypass
	dim tempmima,ss,tt
	adminname = Trim(Replace(Request("adminname"), "'", ""))
	password = md5(Trim(Replace(Request("password"), "'", "")))
	mypass=Trim(Replace(Request("mypassword"), "'", ""))
	If enchiasp.CheckPost = False Then
		ErrMsg = ErrMsg + "您提交的数据不合法，请不要从外部提交登陆。"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request("adminname")) = False Then
		ErrMsg = ErrMsg + "<li>用户名中含有非法字符。</li>"
		Founderr = True
	End If
	If enchiasp.IsValidPassword(Request("password")) = False Then
		ErrMsg = ErrMsg + "<li>密码中含有非法字符。</li>"
		Founderr = True
	End If
	If Request("verifycode") = "" Then
		ErrMsg = ErrMsg + "<br>" + "<li>请返回输入确认码。</li>"
		Founderr = True
	ElseIf Session("getcode") = "9999" Then
		Session("getcode") = ""
		ErrMsg = ErrMsg + "<br>" + "<li>请不要重复提交，如需重新登陆请返回登陆页面。</li>"
		Founderr = True
	ElseIf CStr(Session("getcode"))<>CStr(Trim(Request("verifycode"))) Then
		ErrMsg = ErrMsg + "<br>" + "<li>您输入的验证码和系统产生的不一致，请重新输入。</li>"
		Founderr = True
	End If
	Session("getcode") = ""
	If adminname = "" Or password = "" Then
		Founderr = True
		ErrMsg = ErrMsg + "<br>" + "<li>请输入您的用户名或密码。</li>"
		Exit Sub
	End If
	
	
	

	
	If Founderr = True Then Exit Sub
	If Not IsObject(Conn) Then ConnectionDatabase
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Admin where password='" & password & "' And username='" & adminname & "'"
	Rs.Open SQL, Conn, 1, 3
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>您输入的用户名和密码不正确或者您不是系统管理员。！</li>"
		Exit Sub
	Else
		If password <> Rs("password") Then
			FoundErr = True
			ErrMsg = ErrMsg + "<br><li>用户名或密码错误！！！</li>"
			Exit Sub
		End If
		If Rs("isLock") <> 0 Or Rs("isLock") = "" Then
			Founderr = True
			ErrMsg = "<li>你的用户名已被锁定,你不能登陆！如要开通此帐号，请联系管理员。</li>"
			Exit Sub
		End If
		'检查是否开启二次密码开关
		if enchiasp.ercilogin="1" then
			if mypass="" then
				ErrMsg = ErrMsg + "<br>" + "<li>请不要非法尝试登陆，请确认你是否是系统管理员。系统已经记录下你的操作记录。</li>"
				Founderr = True
				exit sub
			else	
				if rs("isuseercima")=1 then
				
					'判断是加法规则还是乘法规则
					if rs("jiafa")=1 then
						'加法
						tempmima=cstr(cint(mid(Request("verifycode"),rs("weizhi1"),1))+cint(mid(Request("verifycode"),rs("weizhi2"),1)))
						ss=""
						ss=mid(enchiasp.mypass,1,cint(rs("jimaweizhi")))
						tempmima=ss+tempmima+mid(enchiasp.mypass,cint(rs("jimaweizhi"))+1)
						if mypass<>tempmima then
							ErrMsg = ErrMsg + "<br>" +  "<li>请不要非法尝试登陆，请确认你是否是系统管理员。系统已经记录下你的操作记录。</li>"
							Founderr = True
							exit sub
						end if
						
					elseif rs("jiafa")=0 then
						'乘法
						tempmima=cstr(cint(mid(Request("verifycode"),rs("weizhi1"),1))*cint(mid(Request("verifycode"),rs("weizhi2"),1)))
						ss=""
						ss=mid(enchiasp.mypass,1,cint(rs("jimaweizhi")))
						tempmima=ss+tempmima+mid(enchiasp.mypass,cint(rs("jimaweizhi"))+1)
						if mypass<>tempmima then
							ErrMsg = ErrMsg + "<br>" + "<li>请不要非法尝试登陆，请确认你是否是系统管理员。系统已经记录下你的操作记录。</li>"
							Founderr = True
							exit sub
						end if

					end if
					
				else
					'没有开启密码规则
					if mypass<>enchiasp.mypass then
						ErrMsg = ErrMsg + "<br>" + "<li>请不要非法尝试登陆，请确认你是否是系统管理员。系统已经记录下你的操作记录。</li>"
						Founderr = True
						exit sub
					end if
				end if
			end if
		end if
	
		
		
		
		
	End If
	RandomCode = enchiasp.GetRandomCode
	Rs("LoginTime") = Now()
	Rs("Loginip") = enchiasp.GetUserip
	Rs("RandomCode") = RandomCode
	Rs.Update
	If FoundErr = False Then
		Session("AdminName") = Rs("username")
		Session("AdminPass") = Rs("password")
		Session("AdminGrade") = Rs("AdminGrade")
		Session("Adminflag") = Rs("Adminflag")
		Session("AdminStatus") = Rs("Status")
		Session("AdminRandomCode") = RandomCode
		Session("AdminID") = Rs("id")
		Response.Cookies(Admin_Cookies_Name)("AdminName") = Rs("username")
		Response.Cookies(Admin_Cookies_Name)("AdminPass") = Rs("password")
		Response.Cookies(Admin_Cookies_Name)("AdminGrade") = Rs("AdminGrade")
		Response.Cookies(Admin_Cookies_Name)("Adminflag") = Rs("Adminflag")
		Response.Cookies(Admin_Cookies_Name)("AdminStatus") = Rs("Status")
		Response.Cookies(Admin_Cookies_Name)("RandomCode") = RandomCode
		Response.Cookies(Admin_Cookies_Name)("AdminID") = Rs("id")
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Redirect("admin_kindex.asp")
End Sub

Function GetCode1()
	Dim Test
	On Error Resume Next
	Set Test = Server.CreateObject("Adodb.Stream")
	Set Test = Nothing
	If Err Then
		Dim zNum
		Randomize Timer
		zNum = CInt(8999 * Rnd + 1000)
		Session("GetCode") = zNum
		getcode1 = Session("GetCode")
	Else
		getcode1 = "<img src=""../inc/getcode.asp"">"
	End If
End Function
%>
</body>
</html>
