<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>��������</title>
<meta http-equiv=Content-Type content=text/html;charset=gb2312>
<base target="main">
<script language="javascript">
<!--
var displayBar=true;
function switchBar(obj){
	if (displayBar)
	{
		parent.frame.cols="0,*";
		displayBar=false;
		obj.src="images/admin_logo_2.gif";
		obj.title="����߹���˵�";
	}
	else{
		parent.frame.cols="180,*";
		displayBar=true;
		obj.src="images/admin_logo_1.gif";
		obj.title="�ر���߹���˵�";
	}
}
//-->
</script>
<link href="style.css" type=text/css rel=stylesheet>
<style type=text/css>
a { color:#FFFFFF;text-decoration:none}
a:hover {color:#DBDBDB;text-decoration: underline}
td {color: #FFFFFF; font-family: "����";font-weight:bold;}
</style>
</head>
<body leftmargin="0" topmargin="0">
<table cellSpacing="0" cellPadding="0" align="center" width="100%" border="0">
	<tr>
		<td class="BodyTitle" height="28"><table cellSpacing="0" cellPadding="0" width="100%" border="0">
	<tr>
		<td align="left"><img src="images/admin_logo.gif" onclick="switchBar(this)" width="150" height="32" border=0 alt="�ر���߹���˵�" style="cursor:hand"></td>
		<td width="50%"  align=right><font color="FFFFAA">�������</font>&nbsp;&nbsp;      
		<a href="help.asp" target=blank>ϵͳ����</a>&nbsp;&nbsp;<a href="admin_label.asp" target=blank>��ǩһ��</a>&nbsp;&nbsp;<a href="admin_config.asp?action=reload" target=main>�ؽ�����</a>&nbsp;&nbsp;<a href=admin_password.asp target=main>�޸�����</a>&nbsp;&nbsp;</td>
		<td width="5%" align=right><A href=../ target=_blank><img src="images/i_home.gif" title="������ҳ" border=0></A>&nbsp;</TD>
	</tr>
		</table></td> </tr>
	<tr><td bgColor="#485161" height="1"></td></tr>
	<tr><td bgColor="#CDCDCD" height="1"></td></tr>
	<tr><td bgColor="#B5BCC7" height="1"></td></tr>
</table>
</body>
</html>
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
Call CloseConn
%>
