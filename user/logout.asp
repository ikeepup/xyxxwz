<!--#include file="../api/cls_api.asp"-->
<!--#include file="config.asp"-->
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
Dim UserName,Password
UserName = enchiasp.membername
Password = ""



'清除COOKIES中的验证信息.
Response.Cookies(enchiasp.Cookies_Name).path = "/"
Response.Cookies(enchiasp.Cookies_Name)("userid") = ""
Response.Cookies(enchiasp.Cookies_Name)("username") = ""
Response.Cookies(enchiasp.Cookies_Name)("password") = ""
Response.Cookies(enchiasp.Cookies_Name)("nickname") = ""
Response.Cookies(enchiasp.Cookies_Name)("UserGrade") = ""
Response.Cookies(enchiasp.Cookies_Name)("UserGroup") = ""
Response.Cookies(enchiasp.Cookies_Name)("UserClass") = ""
Response.Cookies(enchiasp.Cookies_Name)("UserToday") = ""
Response.Cookies(enchiasp.Cookies_Name)("usercookies") = ""
Response.Cookies(enchiasp.Cookies_Name)("LastTimeDate") = ""
Response.Cookies(enchiasp.Cookies_Name)("LastTimeIP") = ""
Response.Cookies(enchiasp.Cookies_Name)("LastTime") = ""
Response.Cookies(enchiasp.Cookies_Name) = ""
'-----------------------------------------------------------------
'系统整合
'-----------------------------------------------------------------
Dim API_enchiasp,API_SaveCookie,SysKey
If API_Enable Then
	Set API_enchiasp = New API_Conformity
	Md5OLD = 1
	SysKey = Md5(UserName & API_ConformKey)
	Md5OLD = 0
	API_SaveCookie = API_enchiasp.SetCookie(SysKey,UserName,Password,0)
	Set API_enchiasp = Nothing
	Response.Write API_SaveCookie
	If API_LogoutUrl <> "0" Then
		Response.Write "<script language=JavaScript>"
		Response.Write "setTimeout(""window.location='"& API_LogoutUrl &"'"",1000);"
		Response.Write "</script>"
	Else
		Response.Write "<script language=""javascript"">window.setInterval(""location.reload('../')"",1000);</script>"
	End If
Else
	Response.Redirect ("../")
End If
'-----------------------------------------------------------------

%>