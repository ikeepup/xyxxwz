<%@ LANGUAGE = VBScript CodePage = 936%>
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
'��������ͳ��ģ�飬û��ʵ������
Option Explicit
Response.Buffer = True
If Not IsNumeric(Request("id")) And Request("id")<>"" then
	Response.write"�����ϵͳ����!ID����������"
	Response.End
End If
Dim theurl,id,strStation
id = CInt(Request.Querystring("id"))
strStation = Left(Request.ServerVariables("HTTP_REFERER"),250)
theurl="http://" & Request.ServerVariables("HTTP_HOST") & finddir(Request.ServerVariables("url"))
%>document.write("<script>var url='<%=theurl%>';</script>")
_dwrite("<script language=javascript src="+url+"inc/online.asp?id=<%=id%>&stat=<%=strStation%>&Referer="+escape(document.referrer)+"></script>");
function _dwrite(string) {document.write(string);}
<%
Function finddir(filepath)
	finddir=""
	Dim i,abc
	for i=1 to len(filepath)
		if left(right(filepath,i),1)="/" or left(right(filepath,i),1)="\" then
	  		abc=i
	  		exit for
		end if
	next
	if abc <> 1 then
		finddir=left(filepath,len(filepath)-abc+1)
	end if
end Function
%>
