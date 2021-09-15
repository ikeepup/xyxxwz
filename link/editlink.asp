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
Dim Rs, SQL, FoundErr, ErrMsg, LinkID
Dim HtmlContent,ListContent,TempListContent
Dim strChecked

FoundErr = False
enchiasp.PreventInfuse

enchiasp.LoadTemplates 9999, 6, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","修改友情连接")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
TempListContent = enchiasp.CutFixContent(HtmlContent, "<!--ListBegin", "ListEnd-->", 1)

LinkID = enchiasp.ChkNumeric(Request("id"))
If LinkID = 0 Then
	Response.Write"错误的系统参数!"
	Response.End
End If
If enchiasp.CheckStr(LCase(Request.Form("action"))) = "modify" Then
	Call FriendLinkModify
Else
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_Link WHERE LinkID="& LinkID)
	If Rs.BOF And Rs.EOF Then
		Set Rs = Nothing
		Call OutAlertScript("错误的系统参数!")
	Else
		ListContent = enchiasp.HtmlSetting(13)
		ListContent = Replace(ListContent,"{$LinkID}", Rs("LinkID"))
		ListContent = Replace(ListContent,"{$LinkName}", enchiasp.HTMLEncode(Rs("LinkName")))
		ListContent = Replace(ListContent,"{$LinkUrl}", enchiasp.CheckTopic(Rs("LinkUrl")))
		ListContent = Replace(ListContent,"{$LogoUrl}", enchiasp.CheckTopic(Rs("LogoUrl")))
		ListContent = Replace(ListContent,"{$Readme}", enchiasp.HTMLEncode(Rs("Readme")))
		If Rs("isLogo") = 0 Then
			ListContent = Replace(ListContent,"{$CheckedA}", " checked")
			ListContent = Replace(ListContent,"{$CheckedB}", "")
		Else
			ListContent = Replace(ListContent,"{$CheckedA}", "")
			ListContent = Replace(ListContent,"{$CheckedB}", " checked")
		End If
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
		Response.Write HtmlContent
	End If
	Rs.Close:Set Rs = Nothing

End If
Sub FriendLinkModify()
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
	If Trim(Request.Form("password")) = "" Then
		ErrMsg = ErrMsg + "管理密码不能为空\n"
		Founderr = True
	End If
	If enchiasp.IsValidPassword(Request("password")) = False Then
		ErrMsg = ErrMsg + "管理密码中含有非法字符\n"
		Founderr = True
	End If
	Set Rs = enchiasp.Execute("SELECT password FROM ECCMS_Link WHERE LinkID="& LinkID)
	If Rs.BOF And Rs.EOF Then
		Set Rs = Nothing
		Call OutAlertScript("错误的系统参数!")
		Exit Sub
	Else
		If Not IsNull(Trim(Rs("password"))) And Trim(Rs("password")) <> "" Then
			If Rs("password") <> md5(Request.Form("password")) Then
				Set Rs = Nothing
				Call OutAlertScript(enchiasp.HtmlSetting(10))
				Exit Sub
			End If
		Else
			Set Rs = Nothing
			Call OutAlertScript(enchiasp.HtmlSetting(11))
			Exit Sub
		End If
	End If
	Set Rs = Nothing
	If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		Exit Sub
	End If
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_Link WHERE LinkID="& LinkID
	Rs.Open SQL,Conn,1,3
		Rs("LinkName") = enchiasp.FormEncode(Request.Form("LinkName"),50)
		Rs("LinkUrl") = enchiasp.FormEncode(Request.Form("LinkUrl"),200)
		Rs("LogoUrl") = enchiasp.FormEncode(Request.Form("LogoUrl"),200)
		Rs("Readme") = enchiasp.FormEncode(Request.Form("Readme"),200)
		Rs("isLogo") = enchiasp.ChkNumeric(Request.Form("isLogo"))
	Rs.update
	Rs.Close:Set Rs = Nothing
	Call OutputScript(enchiasp.HtmlSetting(14),"index.asp")
End Sub
CloseConn
%>
