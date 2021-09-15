<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/cls_public.asp"-->
<!--#include file="../inc/NewsChannel.asp"-->
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
Dim ArticleID,sortid,showid,indexid
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID < 1 Then ChannelID = 1
ArticleID = enchiasp.ChkNumeric(Request("ArticleID"))
sortid = enchiasp.ChkNumeric(Request("sortid"))
showid = enchiasp.ChkNumeric(Request("showid"))
indexid = enchiasp.ChkNumeric(Request("indexid"))
enchicms.Channel = ChannelID
enchicms.ChannelMain
enchicms.ShowFlush = showid
If indexid > 0 Then
	enchicms.CreateArticleIndex
End If
If ArticleID > 0 Then
	enchicms.CreateArticleContent(ArticleID)
End If
Set enchicms = Nothing
CloseConn
%>