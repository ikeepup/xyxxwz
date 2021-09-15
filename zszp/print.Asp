<!--#include file="config.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>

<head>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">

<%
if Request("action")="save" then
%>
<OBJECT id=ObjWB height=0 width=0 
classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2></OBJECT>

<%
else
%>
<script language="javascript">
		window.print();
</script>
<%
end if
%>
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
Dim ArticleID,Rs,SQL
Dim AllHits,DayHits,WeekHits,MonthHits,HitsTime,hits
If Not IsNumeric(Request("ArticleID")) And Request("ArticleID") <> "" then
	Response.Write"错误的系统参数!ID必须是数字"
	Response.End
Else
	ArticleID = CLng(Request.querystring("ArticleID"))
End If
If Not IsObject(Conn) Then ConnectionDatabase
Set Rs = Server.CreateObject("ADODB.Recordset")
SQL = "select * from ECCMS_Article where ArticleID = "& ArticleID  
Rs.Open SQL,Conn,1,3
%>

	<title><%=rs("title")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {
	font-size: 16px;
	line-height: 200%;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	margin-left: 0px;
	color: #000000;
	text-align: left;
	background-position: left;
}
.image 
{
	border:0px;
	cursor:hand;	
}
.f16 {
	font-family: 16;
	line-height: 200%;
	color: #000000;
}
.title
{
	font-size:25px;
	color:#FF0000;
	line-height:150%;
	font-weight:bold;
}
a:link {
	font-size: 16px;
	line-height: 200%;
	color: #333333;
	text-decoration: underline;
}
a:visited {
	font-size: 16px;
	line-height: 200%;
	color: #666666;
}
a:hover {
	font-size: 16px;
	line-height: 200%;
	color: #FF0000;
}
a:active {
	font-size: 16px;
	line-height: 200%;
	color: #990000;
}
-->
     </style>
	 <script language="javascript">
	 function myprint()
		{
			window.print();
		}
	 </script>
</head>
<body>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
											<td bgcolor="#FFFFFF">
										<!--有内容显示-->
										
										<table width="90%"  border="0" cellspacing="0" cellpadding="0" align="center">
										  <tr>
											<td align="center">
											<%
if Request("action")="save" then

else
%>


											&nbsp;&nbsp;&nbsp;&nbsp;           
						<a href="#" onClick="javascript:myprint();">打印</a><%
end if
%>
											</td>
										  </tr>
										</table>

										
											<TABLE width="98%" border="0" align="center" cellSpacing="1">
													<TBODY>
														<TR>
															<TD align="left">
																<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
																	<TBODY>
																	
																		<TR>
																			<TD align=left width="100%" bgcolor="#FFFFFF" style="font-size:16px; ">
																				<DIV align=center>
<TABLE cellSpacing=1 width=640 border=0>
<TBODY>
<TR>
<TD align=left width=636>
<TABLE height=162 cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD align=middle width="100%" height=100>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD style="FONT-SIZE: 30pt" align=middle width="100%" height=100>
<P style="MARGIN: 0px; LINE-HEIGHT: 100%"><b><font face="华文中宋" color="#ff0000">运城市公安局交警支队</font></b></P></TD></TR></TBODY></TABLE></TD></TR>
<TR>
<TD style="MARGIN-TOP: 2px; FONT-SIZE: 18pt; MARGIN-BOTTOM: 2px; LINE-HEIGHT: 150%" align=middle width="100%" height=31>
<TABLE style="MARGIN-TOP: 3px; FONT-SIZE: 18pt; MARGIN-BOTTOM: 3px; LINE-HEIGHT: 150%" width="100%" border=0>
<TBODY>
<TR>
<TD style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px; LINE-HEIGHT: 150%" align=middle width="100%">
<P style="MARGIN: 3px 0px">
<%
response.write rs("title")
%>
</P></TD></TR></TBODY></TABLE></TD></TR>
<TR>
<TD align=right width="100%" height=31>
<TABLE style="MARGIN-TOP: 1px; FONT-SIZE: 12pt; MARGIN-BOTTOM: 1px; LINE-HEIGHT: 200%" width="100%" border=0>
<TBODY>
<TR>
<TD align=left width="100%">
<%
response.write rs("content")
%>
</TD></TR>
<TR>
<TD align=right width="100%">
<%
response.write rs("writetime")
%></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>

<%

Rs.close
set Rs=nothing
CloseConn
%>
</TD></TR></TBODY></TABLE></DIV>
																			</TD>
																		</TR>
																	</TBODY>
																</TABLE>
															</TD>
														</TR>
													</TBODY>
											  </TABLE>
											  <!--有内容显示结束-->
											  
										  </td>
  </tr>
</table>
<%
if Request("action")="save" then
%>
 <script language="javascript">
	 document.all.ObjWB.ExecWB(4,1); 
	 </script>

<%
end if
%>


</body>
</html>
