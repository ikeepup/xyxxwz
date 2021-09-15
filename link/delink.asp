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
Dim Rs, LinkID
Dim HtmlContent,ListContent,TempListContent

enchiasp.PreventInfuse

enchiasp.LoadTemplates 9999, 6, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","删除友情连接")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
TempListContent = enchiasp.CutFixContent(HtmlContent, "<!--ListBegin", "ListEnd-->", 1)

LinkID = enchiasp.ChkNumeric(Request("id"))

If LinkID = 0 Then
	Response.Write"错误的系统参数!"
	Response.End
End If
If enchiasp.CheckStr(LCase(Request.Form("action"))) = "del" Then
	Call LinkDel
Else
	Set Rs = enchiasp.Execute("Select LinkID,LinkName,LinkUrl From ECCMS_Link where LinkID="& LinkID)
	If Rs.BOF And Rs.EOF Then
		Set Rs = Nothing
		Call OutAlertScript("错误的系统参数!")
	Else
		ListContent = enchiasp.HtmlSetting(8)
		ListContent = Replace(ListContent,"{$LinkID}", Rs("LinkID"))
		ListContent = Replace(ListContent,"{$LinkName}", enchiasp.HTMLEncode(Rs("LinkName")))
		ListContent = Replace(ListContent,"{$LinkUrl}", enchiasp.CheckTopic(Rs("LinkUrl")))
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
		Response.Write HtmlContent
	End If
	Rs.Close:Set Rs = Nothing
End If

Sub LinkDel()
	If Trim(Request.Form("password")) = "" Then
		Call OutAlertScript(enchiasp.HtmlSetting(9))
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("Select password From ECCMS_Link where LinkID="& LinkID)
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
	If enchiasp.CheckStr(LCase(Request.Form("action"))) = "del" Then
		enchiasp.Execute("Delete from ECCMS_Link where LinkID="& LinkID)
		Call OutputScript("友情连接删除成功！","index.asp")
	End If
End Sub
CloseConn
%>