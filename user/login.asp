<!--#include file="config.asp"-->
<!--#include file="../inc/classmenu.asp"-->
<!--#include file="../api/cls_api.asp"-->
<!--#include file="../inc/cls_public.asp"-->
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
Dim HtmlContent,ChannelRootDir
ChannelRootDir = enchiasp.InstallDir & "user/"
enchiasp.LoadTemplates 9999, 5, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
'--Ƶ��Ŀ¼
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","�û���¼")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = html.ReadAnnounceContent(HtmlContent, 0)
HtmlContent = HTML.ReadFriendLink(HtmlContent)

if enchiasp.membername<>"" then
	Response.Redirect ("./index.asp")

end if

'If CheckLogin Then
	'Response.Redirect ("./index.asp")
'End If

If LCase(Request("action")) = "login" Then
	Call MemberLogin
Else
	HtmlContent = Replace(HtmlContent,"{$UserManageContent}", enchiasp.HtmlSetting(7))
    HtmlContent = Replace(HtmlContent,"{$SiteName}", enchiasp.SiteName)
	HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
	HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
	
	Response.Write HtmlContent
	
End If
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
CloseConn

Sub MemberLogin()
	On Error Resume Next
	Dim Rs,SQL,username, password,usercookies,Group_Setting
	If LCase(Request("reset")) = "login" Then
		'����̳��½
		username = enchiasp.CheckBadstr(Request("username"))
		password =  md5(Request("password"))

	else
		If Trim(Request("username")) <> "" And Trim(Request("password")) <> "" Then
			username = enchiasp.CheckBadstr(Request("username"))
			password = md5(Request("password"))
		Else
			ErrMsg = ErrMsg + "<li>�û��������벻��Ϊ�գ�</li>"
			Founderr = True
			Exit Sub
		End If
	end if
	
	If enchiasp.IsValidStr(Request("username")) = False Then
		ErrMsg = ErrMsg + "<li>�û����к��зǷ��ַ���</li>"
		Founderr = True
	End If
	If enchiasp.IsValidPassword(Request("password")) = False Then
		ErrMsg = ErrMsg + "<li>�����к��зǷ��ַ���</li>"
		Founderr = True
	End If
	
	
	usercookies = enchiasp.ChkNumeric(request("CookieDate"))
	
'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	Dim API_enchiasp,API_SaveCookie,SysKey
	If API_Enable Then
		Set API_enchiasp = New API_Conformity
		API_enchiasp.NodeValue "action","login",0,False
		API_enchiasp.NodeValue "username",UserName,1,False
		Md5OLD = 1
		SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
		Md5OLD = 0
		API_enchiasp.NodeValue "syskey",SysKey,0,False
		API_enchiasp.NodeValue "password",Request("password"),0,False
		API_enchiasp.SendHttpData
		If API_enchiasp.Status = "1" Then
			Founderr = True
			ErrMsg =  API_enchiasp.Message
			Exit Sub
		Else
			API_SaveCookie = API_enchiasp.SetCookie(SysKey,UserName,Password,usercookies)
		End If
		Set API_enchiasp = Nothing
	End If
	'-----------------------------------------------------------------
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM [ECCMS_User] WHERE username='" & username & "'"
	Rs.Open SQL, Conn, 1, 3
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������û��������벻��ȷ������ϵ����Ա��</li>"
		Exit Sub
	Else
		If password <> Rs("password") Then
			FoundErr = True
			ErrMsg = ErrMsg + "<br><li>�û�����������󣡣���</li>"
			Exit Sub
		End If
		If Rs("UserLock") <> 0 Then
			Founderr = True
			ErrMsg = enchiasp.HtmlSetting(8)
			Exit Sub
		End If
		Response.Cookies(enchiasp.Cookies_Name)("LastTimeDate") = Rs("LastTime")
		Response.Cookies(enchiasp.Cookies_Name)("LastTimeIP") = Rs("userlastip")
		Response.Cookies(enchiasp.Cookies_Name)("LastTime") = Rs("LastTime")
		Group_Setting=Split(enchiasp.UserGroupSetting(Rs("UserGrade")), "|||")
		If Rs("userpoint") < 0 Then
			Rs("userpoint") = CLng(Group_Setting(25))
		Else
			Rs("userpoint") = Rs("userpoint") + CLng(Group_Setting(25))
		End If
		If Rs("experience") < 0 Then
			Rs("experience") = CLng(Group_Setting(32))
		Else
			Rs("experience") = Rs("experience") + CLng(Group_Setting(32))
		End If
		If Rs("charm") < 0 Then
			Rs("charm") = CLng(Group_Setting(33))
		Else
			Rs("charm") = Rs("charm") + CLng(Group_Setting(33))
		End If
		Rs("LastTime") = Now()
		Rs("userlastip") = enchiasp.GetUserip
		Rs("UserLogin") = Rs("UserLogin") + 1
		Rs.Update
		'If isnull(usercookies) Or usercookies="" Then usercookies=0
		Select Case usercookies
		Case 0
			Response.Cookies(enchiasp.Cookies_Name)("usercookies") = usercookies
		Case 1
			Response.Cookies(enchiasp.Cookies_Name).Expires=Date+1
			Response.Cookies(enchiasp.Cookies_Name)("usercookies") = usercookies
		Case 2
			Response.Cookies(enchiasp.Cookies_Name).Expires=Date+31
			Response.Cookies(enchiasp.Cookies_Name)("usercookies") = usercookies
		Case 3
			Response.Cookies(enchiasp.Cookies_Name).Expires=Date+365
			Response.Cookies(enchiasp.Cookies_Name)("usercookies") = usercookies
		End Select
		Response.Cookies(enchiasp.Cookies_Name).path = "/"
		Response.Cookies(enchiasp.Cookies_Name)("userid") = Rs("userid")
		Response.Cookies(enchiasp.Cookies_Name)("username") = Rs("username")
		Response.Cookies(enchiasp.Cookies_Name)("password") = Rs("password")
		Response.Cookies(enchiasp.Cookies_Name)("nickname") = Rs("nickname")
		Response.Cookies(enchiasp.Cookies_Name)("UserGrade") = Rs("UserGrade")
		Response.Cookies(enchiasp.Cookies_Name)("UserGroup") = Rs("UserGroup")
		Response.Cookies(enchiasp.Cookies_Name)("UserClass") = Rs("UserClass")
		Response.Cookies(enchiasp.Cookies_Name)("UserToday") = Rs("UserToday")
	End If
	Rs.Close
	Set Rs = Nothing
	'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	If API_Enable Then
		Response.Write API_SaveCookie
		Response.Flush
		If API_LoginUrl <> "0" Then
			Response.Write "<script language=JavaScript>"
			Response.Write "setTimeout(""window.location='"& API_LoginUrl &"'"",1000);"
			Response.Write "</script>"
			Response.End
		End If
	End If
	'-----------------------------------------------------------------
	'Response.Redirect("index.asp")
	Dim comeurlname,comeurl,Returnstr
	comeurl = Trim(Request("comeurl"))
	If Len(comeurl) = 0 Then
		comeurl = Request.ServerVariables("HTTP_REFERER")
	End If
	If instr(lcase(comeurl),"reg.asp")>0 Or instr(lcase(comeurl),"user/login.asp")>0 Or Trim(comeurl)="" Or (Not enchiasp.CheckPost) Then
		comeurlname=""
		comeurl="index.asp"
		Returnstr = "<span id=jump>3</span> ���Ӻ�ϵͳ���Զ����ؿ�������"
	Else
		comeurl=comeurl
		comeurlname="<li><a href="&comeurl&">"&comeurl&"</a></li>"
		Returnstr = "<span id=jump>3</span> ���Ӻ�ϵͳ���Զ�����"
	End If
	HtmlContent = Replace(HtmlContent,"{$UserManageContent}", enchiasp.HtmlSetting(9))
	HtmlContent = Replace(HtmlContent,"{$SiteName}", enchiasp.SiteName)
	HtmlContent = Replace(HtmlContent,"{$UserName}", Request("username"))
	HtmlContent = Replace(HtmlContent,"{$ComeUrl}", comeurl)
	HtmlContent = Replace(HtmlContent,"{$ComeUrlName}", comeurlname)
	HtmlContent = Replace(HtmlContent,"{$ReturnStr}", Returnstr)
	Response.Write HtmlContent
End Sub
%>