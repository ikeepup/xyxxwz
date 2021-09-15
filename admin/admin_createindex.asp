<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/cls_public.asp"-->
<%
Admin_header
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
If Not ChkAdmin("CreateIndex") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Dim ShowPage,HtmlContent,FilePath

If LCase(Request("act")) <> "ok" Then
	'Call CreateSoftIndex
	Call OutputScript("生成首页（HTML）完成！","?act=ok")
Else
	Call CreateSiteIndex
	Response.Write "<p align=center>生成首页（HTML）完成......<br><a href=" & FilePath & " target=_blank>"
	Response.Write Server.MapPath(FilePath)
	Response.Write "</a></p>"
End If
Admin_footer
Set HTML = Nothing
CloseConn
Public Sub CreateSiteIndex()
	On Error Resume Next
	HtmlContent = HTML.ShowIndex(True)
	FilePath = "../" & enchiasp.IndexName
	enchiasp.CreatedTextFile FilePath,HtmlContent
End Sub
%>