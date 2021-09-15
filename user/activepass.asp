<!--#include file="../conn.asp"-->
<!-- #include file="../inc/const.asp" -->
<!--#include file="../api/cls_api.asp"-->
<%
'=====================================================================
' 软件名称：恩池网站管理系统---重新设置用户密码
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim username,password
Dim repassword,answer
Dim Rs, SQL
If Trim(Request("username")) = "" Or Trim(Request("pass")) = "" Or Trim(Request("repass")) = "" Or Trim(Request("answer")) = "" Then
	Response.Write "<script>alert('非法参数！');history.go(-1)</script>"
Else
	username = enchiasp.CheckBadstr(Request("username"))
	password = enchiasp.CheckBadstr(Request("pass"))
	repassword = MD5(enchiasp.Checkstr(Request("repass")))
	answer = MD5(Request("answer"))
	SQL = "SELECT password,UserGrade FROM [ECCMS_User] WHERE username='" & username & "' And password='" & password & "' And answer='" & answer & "'"
	Set Rs = Server.CreateObject("adodb.recordset")
	Rs.open SQL, Conn, 1, 3
	If Rs.EOF And Rs.bof Then
		Response.Write "<script>alert('您返回的用户资料不正确，非法操作！');history.go(-1)</script>"
	Else
		If Rs("UserGrade") = 999 Then
			Response.Write "<script>alert('非法操作！必须和管理员联系取回密码！');history.go(-1)</script>"
		Else
			Rs("password") = repassword
			Rs.Update
			'-----------------------------------------------------------------
			'系统整合
			'-----------------------------------------------------------------
			Dim API_enchiasp,API_SaveCookie,SysKey
			If API_Enable Then
				Set API_enchiasp = New API_Conformity
				API_enchiasp.NodeValue "action","reguser",0,False
				API_enchiasp.NodeValue "username",UserName,1,False
				Md5OLD = 1
				SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
				Md5OLD = 0
				API_enchiasp.NodeValue "syskey",SysKey,0,False
				API_enchiasp.NodeValue "password",Request("repass"),0,False
				API_enchiasp.SendHttpData
				Set API_enchiasp = Nothing
			End If
			'-----------------------------------------------------------------
			Response.Write "<script>alert('您的密码成功激活，请登陆!');location.replace('./')</script>"
		End If
	End If
	Rs.Close
	Set Rs = Nothing
End If
CloseConn
%>