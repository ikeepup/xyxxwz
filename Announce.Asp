<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/classmenu.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/cls_public.asp"-->
<%
'=====================================================================
' 软件名称：恩池网站管理系统
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：94022511
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim HtmlContent
enchiasp.LoadTemplates 9999, 8, 0

HtmlContent = enchiasp.HtmlContent



HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
HtmlContent = Replace(HtmlContent,"{$PageTitle}",enchiasp.HtmlSetting(1))
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = HTML.ReadFriendLink(HtmlContent)
Call AnnounceContent

Response.Write HtmlContent
CloseConn
Sub AnnounceContent()
	Dim Rs,AnnounceID
	enchiasp.PreventInfuse
	If Not IsNumeric(Request("AnnounceID")) Then
		HtmlContent = Replace(HtmlContent,"{$Topic}", enchiasp.HtmlSetting(2))
		HtmlContent = Replace(HtmlContent,"{$PostTime}", vbNullString)
		HtmlContent = Replace(HtmlContent,"{$Hits}", vbNullString)
		HtmlContent = Replace(HtmlContent,"{$Writer}", vbNullString)
		HtmlContent = Replace(HtmlContent,"{$Content}", enchiasp.HtmlSetting(2))
		Exit Sub
	Else
		AnnounceID = CLng(Request("AnnounceID"))
	End If
	On Error Resume Next
	enchiasp.Execute ("UPDATE ECCMS_Announce SET hits=hits+1 WHERE AnnounceID="&AnnounceID)
	Set Rs = enchiasp.Execute("SELECT title,Content,PostTime,writer,hits FROM ECCMS_Announce WHERE AnnounceID="&AnnounceID)
	If Rs.BOF And Rs.EOF Then
		HtmlContent = Replace(HtmlContent,"{$Topic}", enchiasp.HtmlSetting(2))
		HtmlContent = Replace(HtmlContent,"{$PostTime}", vbNullString)
		HtmlContent = Replace(HtmlContent,"{$Hits}", vbNullString)
		HtmlContent = Replace(HtmlContent,"{$Writer}", vbNullString)
		HtmlContent = Replace(HtmlContent,"{$Content}", enchiasp.HtmlSetting(2))
		Set Rs = Nothing
	Else
		HtmlContent = Replace(HtmlContent,"{$Topic}", Rs("title"))
		HtmlContent = Replace(HtmlContent,"{$PostTime}", Rs("PostTime"))
		HtmlContent = Replace(HtmlContent,"{$Hits}", Rs("hits"))
		HtmlContent = Replace(HtmlContent,"{$Writer}", enchiasp.ChkNull(Rs("writer")))
		HtmlContent = Replace(HtmlContent,"{$Content}", Ubbcode(Rs("content")))
	End If
	Set Rs = Nothing
End Sub
%>