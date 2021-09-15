<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
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
Dim Rs, LinkID, LinkUrl
If Not IsNumeric(Request("id")) Then
	Response.Write"错误的系统参数!"
	Response.End
Else
	LinkID = enchiasp.ChkNumeric(Request.Querystring("id"))
End If
If Trim(Request("url")) <> "" Then
	LinkUrl = enchiasp.CheckStr(Trim(Request.Querystring("url")))
Else
	Response.Redirect("../")
End If
enchiasp.Execute ("update ECCMS_Link Set LinkHist = LinkHist + 1 where LinkID = "& LinkID)
if LinkUrl<>"" then
Response.Redirect(LinkUrl)
else
Response.Redirect("../")
end if
%>