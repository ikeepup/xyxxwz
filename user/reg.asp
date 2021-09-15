<!--#include file="config.asp"-->
<!--#include file="../inc/classmenu.asp"-->
<!--#include file="../inc/chkinput.asp"-->
<!--#include file="../inc/email.asp"-->
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
Dim HtmlContent,ChannelRootDir
Dim strRegItem,GetCode

ChannelRootDir = enchiasp.InstallDir & "user/"
enchiasp.LoadTemplates 9999, 5, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
'--频道目录
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}","用户注册")
HtmlContent = Replace(HtmlContent,"{$PageTitle}","用户注册")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)

If CInt(enchiasp.membergrade) > 0 Then Response.Redirect "index.asp"

If CInt(enchiasp.CheckUserReg) <> 1 Then
	ErrMsg = ErrMsg + enchiasp.HtmlSetting(1)
	Founderr = True
ElseIf enchiasp.CheckStr(Request("action")) = "agree" Then
	Call ApplyMember
ElseIf enchiasp.CheckStr(Request("action")) = "reg" Then
	Call RegNewMember
Else
	strRegItem = enchiasp.HtmlSetting(5)
	HtmlContent = Replace(HtmlContent,"{$UserManageContent}", enchiasp.HtmlSetting(3))
	HtmlContent = Replace(HtmlContent,"{$UserRegItem}", Server.HTMLEncode(strRegItem))
	HtmlContent = Replace(HtmlContent,"{$SiteName}", enchiasp.SiteName)
	Response.Write HtmlContent

End If
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If

Sub ApplyMember()
	If Trim(Request.Form("action")) <> "agree" Then
		ErrMsg = ErrMsg + "<li>错误的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	'是否开启邮件及密码发送
	If CInt(enchiasp.IsCloseMail) = 0 And CInt(enchiasp.SendRegMessage) = 1 Then
		If CInt(enchiasp.MailInformPass) = 1 Then
			HtmlContent = Replace(HtmlContent,"{$UserManageContent}", "{$youjian}{$mima}{$UserManageContent}")
		else
			HtmlContent = Replace(HtmlContent,"{$UserManageContent}", "{$youjian}{$UserManageContent}")
		end if
		Select Case CInt(enchiasp.SendMailType)
			Case 0
				HtmlContent = Replace(HtmlContent,"{$youjian}", "<li>系统不支持邮件发送功能，请记住您的注册信息。</li>")
			Case 1
				HtmlContent = Replace(HtmlContent,"{$youjian}", "<li>系统已经开邮件通知，请填写正确的邮箱地址，否则无法接收系统发送的注册相关信息。</li>")
			Case 2
				HtmlContent = Replace(HtmlContent,"{$youjian}", "<li>系统已经开启邮件通知，请填写正确的邮箱地址，否则无法接收系统发送的注册相关信息。</li>")

			Case 3
				HtmlContent = Replace(HtmlContent,"{$youjian}", "<li>系统已经开启邮件通知，请填写正确的邮箱地址，否则无法接收系统发送的邮件密码。</li>")
			Case Else
				HtmlContent = Replace(HtmlContent,"{$youjian}", "<li>系统不支持邮件发送功能，请记住您的注册信息。</li>")
		End Select
		If CInt(enchiasp.MailInformPass) = 1 Then
			HtmlContent = Replace(HtmlContent,"{$mima}", "<li>您的密码将由系统随机生成,您可以不设置密码，在注册后密码将发送到您注册的邮箱中，请注意查收。</li>")
		else
			HtmlContent = Replace(HtmlContent,"{$mima}", "")
		end if

	end if
	If CInt(enchiasp.MailInformPass) = 1 Then
		HtmlContent = Replace(HtmlContent,"{$UserManageContent}", enchiasp.HtmlSetting(20))
	else
		HtmlContent = Replace(HtmlContent,"{$UserManageContent}", enchiasp.HtmlSetting(4))
	end if
	HtmlContent = Replace(HtmlContent,"{$SiteName}", enchiasp.SiteName)
	Response.Write HtmlContent
	End Sub

Sub RegNewMember()
	Dim Rs,SQL
	Dim UserPassWord,strUserName,strGroupName,Password
	Dim rndnum,num1
	Dim Question,Answer,usersex,sex
	On Error Resume Next
	If enchiasp.CheckPost = False Then
		ErrMsg = ErrMsg + "<li>您提交的数据不合法，请不要从外部提交注册。</li>"
		FoundErr = True
	End If
	If Trim(Request.Form("username")) = "" Then
		ErrMsg = ErrMsg + "<li>登录账号不能为空！</li>"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request.Form("username")) = False Then
		ErrMsg = ErrMsg + "<li>登录账号中含有非法字符！</li>"
		Founderr = True
	Else
		strUserName = enchiasp.CheckBadstr(Trim(Request.Form("username")))
	End If
	If Trim(Request.Form("nickname")) = "" Then
		ErrMsg = ErrMsg + "<li>用户昵称不能为空！</li>"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request.Form("nickname")) = False Then
		ErrMsg = ErrMsg + "<li>用户昵称中含有非法字符！</li>"
		Founderr = True
	End If
	If CInt(enchiasp.MailInformPass) = 1 Then
	
	else
		If enchiasp.IsValidPassword(Request.Form("password1")) = False Then
			ErrMsg = ErrMsg + "<li>密码中含有非法字符！</li>"
			Founderr = True
		End If
		If Trim(Request.Form("password1")) <> Trim(Request.Form("password2")) Then
			ErrMsg = ErrMsg + "<li>您输入的密码和确认密码不一致！</li>"
			Founderr = True
		End If
	end if
	
	If IsValidEmail(Request.Form("usermail")) = False Then
		ErrMsg = ErrMsg + "<li>您的Email有错误！</li>"
		Founderr = True
	End If
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
	If Request("verifycode") = "" Then
		ErrMsg = ErrMsg + "<li>请返回输入验证码码。</li>"
		Founderr = True
	ElseIf Session("getcode") = "9999" Then
		Session("getcode") = ""
		ErrMsg = ErrMsg + "<li>请不要重复提交，如需重新登陆请返回登陆页面。</li>"
		Founderr = True
	ElseIf CStr(Session("getcode"))<>CStr(Trim(Request("verifycode"))) Then
		ErrMsg = ErrMsg + "<li>您输入的验证码和系统产生的不一致，请重新输入。</li>"
		Founderr = True
	End If
	Session("getcode") = ""
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_User WHERE username='" & strUserName & "'")
	If Not (Rs.BOF And Rs.EOF) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！此用户已经存在,请换一个用户名再试！</li>"
		Exit Sub
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_Admin WHERE username='" & strUserName & "'")
	If Not (Rs.BOF And Rs.EOF) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！此用户已经存在,请换一个用户名再试！</li>"
		Exit Sub
	End If
	Rs.Close:Set Rs = Nothing
	If CInt(enchiasp.ChkSameMail) = 1 Then
		Set Rs = enchiasp.Execute("SELECT userid FROM ECCMS_User WHERE usermail='" & enchiasp.CheckStr(Request("usermail")) & "'")
		If Not Rs.EOF Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>对不起！本系统已经限制一个邮箱只能注册一个账号。</li><li>此邮箱["&Request("usermail")&"]已经占用，请您换一个邮箱再注册吧。</li>"
		End If
		Rs.Close:Set Rs = Nothing
	End If
	If CInt(enchiasp.MailInformPass) = 1 Then
		Randomize
		Do While Len(rndnum) < 8
			num1 = CStr(Chr((57 - 48) * rnd + 48))
			rndnum = rndnum & num1
		loop
		UserPassWord = rndnum
	Else
		UserPassWord = Trim(Request.Form("password2"))
	End If
	Password = md5(UserPassWord)
	Question = Trim(Request.Form("question"))
	Answer = Trim(Request.Form("answer"))
	If Question = "" Then Question = enchiasp.GetRandomCode
	If Answer = "" Then Answer = enchiasp.GetRandomCode
	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	Dim API_enchiasp,API_SaveCookie,SysKey
	If API_Enable Then
		Set API_enchiasp = New API_Conformity
		API_enchiasp.NodeValue "action","reguser",0,False
		API_enchiasp.NodeValue "username",strUserName,1,False
		Md5OLD = 1
		SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
		Md5OLD = 0
		API_enchiasp.NodeValue "syskey",SysKey,0,False
		API_enchiasp.NodeValue "password",UserPassWord,0,False
		API_enchiasp.NodeValue "email",enchiasp.CheckStr(Request.Form("usermail")),1,False
		API_enchiasp.NodeValue "question",Question,1,False
		API_enchiasp.NodeValue "answer",Answer,1,False
		API_enchiasp.NodeValue "gender",sex,0,False
		API_enchiasp.SendHttpData
		If API_enchiasp.Status = "1" Then
			Founderr = True
			ErrMsg =  ErrMsg & API_enchiasp.Message
			Exit Sub
		Else
			API_SaveCookie = API_enchiasp.SetCookie(SysKey,strUserName,Password,1)
		End If
		Set API_enchiasp = Nothing
	End If
	'-----------------------------------------------------------------
	If Founderr = True Then Exit Sub
	Call PreventRefresh  '防刷新
	Set Rs = enchiasp.Execute("SELECT GroupName FROM ECCMS_UserGroup WHERE Groupid=3")
	If Rs.BOF And Rs.EOF Then
		strGroupName = "普通会员"
	Else
		strGroupName = enchiasp.CheckBadstr(Rs(0))
		If Len(strGroupName) = 0 Then strGroupName = "普通会员"
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_User where (userid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("username") = strUserName
		Rs("password") = Password
		Rs("nickname") = enchiasp.CheckBadstr(Request.Form("nickname"))
		Rs("UserGrade") = 1
		Rs("UserGroup") = strGroupName
		Rs("UserClass") = 0
		If CInt(enchiasp.AdminCheckReg) = 1 Then
			Rs("UserLock") = 1
		Else
			Rs("UserLock") = 0
		End If
		Rs("UserFace") = "face/1.gif"
		Rs("userpoint") = CLng(enchiasp.AddUserPoint)
		Rs("usermoney") = 0
		Rs("savemoney") = 0
		Rs("prepaid") = 0
		Rs("experience") = 10
		Rs("charm") = 10
		Rs("TrueName") = enchiasp.CheckBadstr(Request.Form("username"))
		Rs("usersex") = usersex
		Rs("usermail") = enchiasp.CheckStr(Request.Form("usermail"))
		Rs("oicq") = ""
		Rs("question") = Question
		Rs("answer") = md5(Answer)
		Rs("JoinTime") = Now()
		Rs("ExpireTime") = Now()
		Rs("LastTime") = Now()
		Rs("Protect") = 0
		Rs("usermsg") = 0
		Rs("userlastip") = enchiasp.GetUserip
		If CInt(enchiasp.AdminCheckReg) = 0 And CInt(enchiasp.MailInformPass) = 0 Then
			Rs("userlogin") = 1
		Else
			Rs("userlogin") = 0
		End If
		Rs("usersetting") = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
	Rs.update
	Rs.Close
	SQL = "SELECT userid,username,password,nickname,UserGrade,UserGroup,UserClass,UserLock,userlogin FROM ECCMS_user WHERE username = '" & enchiasp.CheckBadstr(Request.Form("username")) & "' ORDER BY userid DESC"
	Rs.Open SQL, Conn, 1, 3
	If Rs("UserLock") = 0 And CInt(enchiasp.MailInformPass) = 0 Then
		Response.Cookies(enchiasp.Cookies_Name)("userid") = Rs("userid")
		Response.Cookies(enchiasp.Cookies_Name)("username") = Rs("username")
		Response.Cookies(enchiasp.Cookies_Name)("password") = Rs("password")
		Response.Cookies(enchiasp.Cookies_Name)("nickname") = Rs("nickname")
		Response.Cookies(enchiasp.Cookies_Name)("UserGrade") = Rs("UserGrade")
		Response.Cookies(enchiasp.Cookies_Name)("UserGroup") = Rs("UserGroup")
		Response.Cookies(enchiasp.Cookies_Name)("UserClass") = Rs("UserClass")
		'-----------------------------------------------------------------
		'系统整合
		'-----------------------------------------------------------------
		If API_Enable Then
			Response.Write API_SaveCookie
			Response.Flush
		End If
		'-----------------------------------------------------------------
	End If
	Rs.Close
	Set Rs = Nothing
	'发送注册邮件
	Dim username,useremail,topic,mailbody,strMessage
	If CInt(enchiasp.IsCloseMail) = 0 And CInt(enchiasp.SendRegMessage) = 1 Then
		username = strUserName
		useremail = Trim(Request.Form("usermail"))
		topic = "您在 " & enchiasp.SiteName & " 的注册资料"
		mailbody = enchiasp.HtmlSetting(6)
		mailbody = Replace(mailbody,"{$SiteName}", enchiasp.SiteName, 1, -1, 1)
		mailbody = Replace(mailbody,"{$SiteUrl}", enchiasp.SiteUrl, 1, -1, 1)
		mailbody = Replace(mailbody,"{$UserName}", username, 1, -1, 1)
		mailbody = Replace(mailbody,"{$EmailTopic}", topic, 1, -1, 1)
		mailbody = Replace(mailbody,"{$PassWord}", UserPassWord, 1, -1, 1)
		Select Case CInt(enchiasp.SendMailType)
			Case 0
				strMessage = "<li>系统未开启邮件功能，请记住您的注册信息。</li>"
			Case 1
				Call Jmail(useremail, topic, mailbody)
			Case 2
				Call Cdonts(useremail, topic, mailbody)
			Case 3
				Call aspemail(useremail, topic, mailbody)
			Case Else
				strMessage = "<li>系统未开启邮件功能，请记住您的注册信息。</li>"
		End Select
		If SendMail = "OK" Then
			strMessage = "<li>您的注册信息已经发往您的邮箱，[" & Request("usermail") & "] 请注意查收。</li>"
		Else
			strMessage = "<li>由于系统错误，给您发送的注册资料未成功。</li>"
		End If
	End If
	If CInt(enchiasp.AdminCheckReg) = 1 Then
		strMessage = strMessage & "<li>请等待管理员认证……</li>"
	End If
	sucmsg = enchiasp.HtmlSetting(2)
	sucmsg = Replace(sucmsg, "{$UserName}", Request("username"))
	sucmsg = Replace(sucmsg, "{$Message}", strMessage)
	Call ReturnIndex(sucmsg)
	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	If API_Enable Then
		If API_ReguserUrl <> "0" Then
			Response.Write "<script language=JavaScript>"
			Response.Write "setTimeout(""window.location='"& API_ReguserUrl &"'"",1000);"
			Response.Write "</script>"
		End If
	End If
	'-----------------------------------------------------------------
End Sub

Sub ReturnIndex(message)
	Response.Write "<html><head><title>成功提示信息!</title><meta http-equiv=Content-Type content=text/html; charset=gb2312>" & vbCrLf
	Response.Write "<meta http-equiv=refresh content=3;url=index.asp>"
	Response.Write "<link href=user_style.css rel=stylesheet type=text/css></head><body><br /><br />" & vbCrLf
	Response.Write "<table width=460 border=0 align=center cellpadding=0 cellspacing=0>"
	Response.Write "<tr bgcolor='#3795d2'>"
	Response.Write "  <td height='25' valign='top' bgcolor='#3795d2'> <img src='images/user_msg.gif' width=69 height=20></td>"
	Response.Write "  <td align='right' valign='top'> <img src='images/user_login_02.gif' width=4 height=4></td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "  <td width=526 height=1 colspan=2 bgcolor=#f8f6f5></td>"
	Response.Write "</tr>"
	Response.Write "<tr bgcolor=#f8f6f5>"
	Response.Write "  <td width=355 style='padding-left: 10px;padding-top: 5px;'><b style=color:blue><span id=jump>3</span> 秒钟后系统将自动转到用户管理首页</b><br>" & message & "</td>"
	Response.Write "  <td> <img src='images/user_suc.gif' width=95 height=97></td>"
	Response.Write "</tr>"
	Response.Write "<tr bgcolor=#f8f6f5><td align=center colspan=2><a href=index.asp>返回上一页...</a></td></tr>"
	Response.Write "<tr bgcolor='#3795d2'>"
	Response.Write "  <td height='8' valign='bottom'> <img src='images/user_login_04.gif' width=4 height=4></td>"
	Response.Write "  <td align='right' valign='bottom'> <img src='images/user_login_05.gif' width=4 height=4></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<br /><br /></body></html>"
	Response.Write "<script>function countDown(secs){jump.innerText=secs;if(--secs>0)setTimeout(""countDown(""+secs+"")"",1000);}countDown(3);</script>"
End Sub
CloseConn
%>