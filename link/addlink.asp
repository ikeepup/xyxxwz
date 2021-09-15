<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/classmenu.asp"-->
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
Dim Rs, SQL, FoundErr,ErrMsg
Dim isLock,HtmlContent,ListContent
FoundErr = False

enchiasp.PreventInfuse

enchiasp.LoadTemplates 9999, 6, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","申请友情连接")

isLock = enchiasp.ChkNumeric(enchiasp.HtmlSetting(3))   '设置申请连接默认状态。0=正常显示，1=锁定

HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
ListContent = enchiasp.CutFixContent(HtmlContent, "<!--ListBegin", "ListEnd-->", 1)
HtmlContent = Replace(HtmlContent, ListContent, enchiasp.HtmlSetting(5))

If enchiasp.CheckStr(LCase(Request.Form("action"))) = "save" Then
	Call FriendLinkSave
Else
	If CInt(enchiasp.StopApplyLink) <> 0 Then
		Call OutAlertScript(enchiasp.HtmlSetting(6))
	Else
		Response.Write HtmlContent
	End If
End If

Sub FriendLinkSave()
	Call PreventRefresh
	If CInt(enchiasp.StopApplyLink) <> 0 Then
		Call OutAlertScript(enchiasp.HtmlSetting(6))
		Founderr = True
	End If
	If Trim(Request.Form("LinkName")) = "" Then
		ErrMsg = ErrMsg + "网站名称不能为空\n"
		Founderr = True
	End If
	If Trim(Request.Form("LinkUrl")) = "" Then
		ErrMsg = ErrMsg + "网站URL不能为空\n"
		Founderr = True
	End If
	If Trim(Request.Form("Readme")) = "" Then
		ErrMsg = ErrMsg + "网站简介不能为空\n"
		Founderr = True
	End If
	If Trim(Request.Form("password1")) = "" Then
		ErrMsg = ErrMsg + "管理密码不能为空\n"
		Founderr = True
	End If
	If Trim(Request.Form("password2")) = "" Then
		ErrMsg = ErrMsg + "确认管理密码不能为空\n"
		Founderr = True
	End If
	If enchiasp.IsValidPassword(Request("password2")) = False Then
		ErrMsg = ErrMsg + "管理密码中含有非法字符\n"
		Founderr = True
	End If
	If Trim(Request.Form("password1")) <> Trim(Request.Form("password2")) Then
		ErrMsg = ErrMsg + "管理密码和确认密码不一至，请重新输入管理密码\n"
		Founderr = True
	End If
	Set Rs = enchiasp.Execute("SELECT LinkID FROM ECCMS_Link WHERE LinkName='" & Replace(Request.Form("LinkName"), "'", "") & "' And LinkUrl='" & Replace(Request.Form("LinkUrl"), "'", "") & "'")
	If Not (Rs.EOF And Rs.BOF) Then
		ErrMsg = "您申请的友情连接已经存在！"
		Founderr = True
	End If
	Set Rs = Nothing
	If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		Exit Sub
	End If
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_Link WHERE (LinkID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("LinkName") = enchiasp.FormEncode(Request.Form("LinkName"),50)
		Rs("LinkUrl") = enchiasp.FormEncode(Request.Form("LinkUrl"),200)
		Rs("LogoUrl") = enchiasp.FormEncode(Request.Form("LogoUrl"),200)
		Rs("Readme") = enchiasp.FormEncode(Request.Form("Readme"),200)
		Rs("LinkTime") = Now()
		Rs("password") = md5(Request.Form("password2"))
		Rs("LinkHist") = 0
		Rs("isLogo") = enchiasp.ChkNumeric(Request.Form("isLogo"))
		Rs("isIndex") = 0
		Rs("isLock") = isLock
	Rs.update
	Rs.Close:Set Rs = Nothing
	Call OutputScript(enchiasp.HtmlSetting(7),"index.asp")
End Sub
CloseConn
%>