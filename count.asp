<%@ LANGUAGE = VBScript CodePage = 936%>
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
'在线人数统计模块，没有实际意义
Option Explicit
Response.Buffer = True
If Not IsNumeric(Request("id")) And Request("id")<>"" then
	Response.write"错误的系统参数!ID必须是数字"
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
