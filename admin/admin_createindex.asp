<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/cls_public.asp"-->
<%
Admin_header
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
If Not ChkAdmin("CreateIndex") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Dim ShowPage,HtmlContent,FilePath

If LCase(Request("act")) <> "ok" Then
	'Call CreateSoftIndex
	Call OutputScript("������ҳ��HTML����ɣ�","?act=ok")
Else
	Call CreateSiteIndex
	Response.Write "<p align=center>������ҳ��HTML�����......<br><a href=" & FilePath & " target=_blank>"
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