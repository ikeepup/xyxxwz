<!--#include file="config.asp" -->
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
Dim flashid,Rs,SQL
Dim AllHits,DayHits,WeekHits,MonthHits,HitsTime,hits
If Not IsNumeric(Request("flashid")) And Request("flashid") <> "" then
	Response.Write"�����ϵͳ����!ID����������"
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
Response.Write "document.write ("& Chr(34) &"���գ�"& DayHits &" ���ܣ�"& WeekHits &" ���£�"& MonthHits &" �ܼƣ�"& AllHits &" "& Chr(34) &");"
CloseConn
%>