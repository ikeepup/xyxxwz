<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
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
Call InnerLocation("付款方式")

Dim Rs,i
Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>"
Response.Write "	<tr>"
Response.Write "		<th colspan=2>付款方式</th>"
Response.Write "	</tr>"
Set Rs = enchiasp.Execute("SELECT modeid,modename,site,code,payee,url,readme FROM ECCMS_Paymode ORDER BY modeid")
If Rs.BOF And Rs.EOF Then
	Response.Write "<tr><td align=center colspan=2 class=UserTableRow1>没有付款方式！</td></tr>"
Else
	i = 0
	Do While Not Rs.EOF
		i = i + 1
%>
	<tr height=20>
		<td class=Usertablerow2 colspan=2><%=i%>、<a href="<%=Rs("url")%>" target=_blank><b><%=Rs("modename")%></b></a></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 width="20%" align=right>开户银行：</td>
		<td class=Usertablerow1 width="80%"><%=Rs("site")%></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>银行帐号：</td>
		<td class=Usertablerow1><%=Rs("code")%></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>收 款 人：</td>
		<td class=Usertablerow1><%=Rs("payee")%></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>说 明：</td>
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











