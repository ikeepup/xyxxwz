<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
<%
'=====================================================================
' �������ƣ�������վ����ϵͳ
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж��������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Call InnerLocation("���ʽ")

Dim Rs,i
Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>"
Response.Write "	<tr>"
Response.Write "		<th colspan=2>���ʽ</th>"
Response.Write "	</tr>"
Set Rs = enchiasp.Execute("SELECT modeid,modename,site,code,payee,url,readme FROM ECCMS_Paymode ORDER BY modeid")
If Rs.BOF And Rs.EOF Then
	Response.Write "<tr><td align=center colspan=2 class=UserTableRow1>û�и��ʽ��</td></tr>"
Else
	i = 0
	Do While Not Rs.EOF
		i = i + 1
%>
	<tr height=20>
		<td class=Usertablerow2 colspan=2><%=i%>��<a href="<%=Rs("url")%>" target=_blank><b><%=Rs("modename")%></b></a></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 width="20%" align=right>�������У�</td>
		<td class=Usertablerow1 width="80%"><%=Rs("site")%></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>�����ʺţ�</td>
		<td class=Usertablerow1><%=Rs("code")%></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>�� �� �ˣ�</td>
		<td class=Usertablerow1><%=Rs("payee")%></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>˵ ����</td>
		<td class=Usertablerow1>&nbsp;&nbsp;<%=Rs("readme")%></td>
	</tr>
<%
		Rs.movenext
	Loop
End If
Rs.Close:Set Rs = Nothing
Response.Write "</table>"
%>
<!--#include file="foot.inc"-->










