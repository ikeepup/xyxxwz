<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/classmenu.asp"-->
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
Dim Rs, SQL, FoundErr, ErrMsg, LinkID
Dim HtmlContent,ListContent,TempListContent
Dim strChecked

FoundErr = False
enchiasp.PreventInfuse

enchiasp.LoadTemplates 9999, 6, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","�޸���������")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
TempListContent = enchiasp.CutFixContent(HtmlContent, "<!--ListBegin", "ListEnd-->", 1)

LinkID = enchiasp.ChkNumeric(Request("id"))
If LinkID = 0 Then
	Response.Write"�����ϵͳ����!"
	Response.End
End If
If enchiasp.CheckStr(LCase(Request.Form("action"))) = "modify" Then
	Call FriendLinkModify
Else
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_Link WHERE LinkID="& LinkID)
	If Rs.BOF And Rs.EOF Then
		Set Rs = Nothing
		Call OutAlertScript("�����ϵͳ����!")
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
		ErrMsg = ErrMsg + "��վ���Ʋ���Ϊ��\n"
		Founderr = True
	End If
	If Trim(Request.Form("LinkUrl")) = "" Then
		ErrMsg = ErrMsg + "��վURL����Ϊ��\n"
		Founderr = True
	End If
	If Trim(Request.Form("Readme")) = "" Then
		ErrMsg = ErrMsg + "��վ��鲻��Ϊ��\n"
		Founderr = True
	End If
	If Trim(Request.Form("password")) = "" Then
		ErrMsg = ErrMsg + "�������벻��Ϊ��\n"
		Founderr = True
	End If
	If enchiasp.IsValidPassword(Request("password")) = False Then
		ErrMsg = ErrMsg + "���������к��зǷ��ַ�\n"
		Founderr = True
	End If
	Set Rs = enchiasp.Execute("SELECT password FROM ECCMS_Link WHERE LinkID="& LinkID)
	If Rs.BOF And Rs.EOF Then
		Set Rs = Nothing
		Call OutAlertScript("�����ϵͳ����!")
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
