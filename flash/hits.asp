<!--#include file="config.asp" -->
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
Dim flashid,Rs,SQL
Dim AllHits,DayHits,WeekHits,MonthHits,HitsTime,hits
If Not IsNumeric(Request("flashid")) And Request("flashid") <> "" then
	Response.Write"错误的系统参数!ID必须是数字"
	Response.End
Else
	flashid = CLng(Request.querystring("flashid"))
End If
If Not IsObject(Conn) Then ConnectionDatabase
Set Rs = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT AllHits,DayHits,WeekHits,MonthHits,HitsTime FROM ECCMS_FlashList where flashid = "& flashid  
Rs.Open SQL,Conn,1,3
	hits = CLng(Rs("AllHits"))+1
	Rs("AllHits").Value = hits
	If DateDiff("Ww", Rs("HitsTime"), Now) <= 0 Then
		Rs("WeekHits").Value = Rs("WeekHits").Value + 1
	Else
		Rs("WeekHits").Value = 1
	End If
	If DateDiff("M", Rs("HitsTime"), Now) <= 0 Then
		Rs("MonthHits").Value = Rs("MonthHits").Value + 1
	Else
		Rs("MonthHits").Value = 1
	End If
	If DateDiff("D", Rs("HitsTime"), Now) <= 0 Then
		Rs("DayHits").Value = Rs("DayHits").Value + 1
	Else
		Rs("DayHits").Value = 1
		Rs("HitsTime").Value = Now
	End If
	Rs.Update
	AllHits = Rs("AllHits")
	DayHits = Rs("DayHits")
	WeekHits = Rs("WeekHits")
	MonthHits = Rs("MonthHits")
Rs.close
set Rs=nothing
Response.Write "document.write ("& Chr(34) &"本日："& DayHits &" 本周："& WeekHits &" 本月："& MonthHits &" 总计："& AllHits &" "& Chr(34) &");"
CloseConn
%>