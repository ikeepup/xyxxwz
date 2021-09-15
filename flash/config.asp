<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/cls_public.asp"-->
<!--#include file="../inc/FlashChannel.asp"-->
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
Dim ChannelID,FoundErr
ChannelID = 5
FoundErr = False
enchicms.Channel = ChannelID
enchicms.MainChannel

'================================================
'过程名：Returnerr
'作  用：返回错误信息
'================================================
Sub Returnerr(message)
	Response.Write "<html><head><title>错误提示信息!</title><meta http-equiv=Content-Type content=text/html; charset=gb2312>" & vbCrLf
	Response.Write "<meta http-equiv=refresh content=10;url=./>"
	Response.Write "<style type=""text/css"">" & vbNewLine
	Response.Write "body {font-size: 12px;font-family: 宋体;}" & vbNewLine
	Response.Write "td {font-size: 12px; font-family: 宋体; line-height: 18px;table-layout:fixed;word-break:break-all}" & vbNewLine
	Response.Write "a {color: #555555; text-decoration: none}" & vbNewLine
	Response.Write "a:hover {color: #FF8C40; text-decoration: underline}" & vbNewLine
	Response.Write "th{ background-color: #3795D2;color: white;font-size: 12px;font-weight:bold;height: 25;}" & vbNewLine
	Response.Write ".TableRow1 {background-color:#F7F7F7;}" & vbNewLine
	Response.Write ".TableRow2 {background-color:#F0F0F0;}" & vbNewLine
	Response.Write ".TableBorder {border: 1px #3795D2 solid ; background-color: #FFFFFF;font: 12px;}" & vbNewLine
	Response.Write "</style>" & vbNewLine
	Response.Write "</head><body><br /><br />" & vbCrLf
	Response.Write "<table width=500 border=0 align=center cellpadding=0 cellspacing=0 class=TableBorder>"
	Response.Write "<tr>"
	Response.Write "  <th>出错啦~!</th>"
	Response.Write "</tr>"
	Response.Write "<tr height=50>"
	Response.Write "  <td valign='top' class=TableRow1 style='padding-left: 10px;padding-top: 5px;'><b style=color:blue><span id=jump>10</span> 秒钟后系统将自动返回首页</b><br>" & message & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr height=22><td align=center class=TableRow2><a href=./>返回首页...</a> | <a href=javascript:window.close()>关闭本窗口...</a></td></tr>"
	Response.Write "</table>"
	Response.Write "<br /><br /></body></html>"
	Response.Write "<script>function countDown(secs){jump.innerText=secs;if(--secs>0)setTimeout(""countDown(""+secs+"")"",1000);}countDown(10);</script>"
End Sub
%>