<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
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
Dim Rs, LinkID, LinkUrl
If Not IsNumeric(Request("id")) Then
	Response.Write"�����ϵͳ����!"
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