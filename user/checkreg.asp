<!--#include file="config.asp"-->
<!--#include file="../api/cls_api.asp"-->
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
Dim Rs
Dim CheckTitle,strUserName,strMessage,UserEmail
If enchiasp.IsValidStr(Request("username")) = False Then
	strMessage = "<li>用户中含有非法字符！</li>"
	CheckTitle = "出错啦！"
	Founderr = True
Else
	strUserName = enchiasp.CheckBadstr(Request("username"))
	UserEmail = enchiasp.Checkstr(Request("usermail"))
End If
If Founderr = False Then
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_User WHERE username = '" & strUserName & "'")
	If Not (Rs.bof And Rs.EOF) Then
		CheckTitle = "出错啦！"
		strMessage =  "<li>Sorry！此用户已经存在,请换一个用户名再试！</li>"
	Else
	
		CheckTitle = "检测通过！"
		strMessage =  "<li><font color=red><b>" & strUserName & "</b></font> 尚未被人使用，赶紧注册吧！</li>"
	End If
	Rs.Close:Set Rs = Nothing
	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	Dim API_enchiasp,API_SaveCookie,SysKey
	If API_Enable Then
		Set API_enchiasp = New API_Conformity
		API_enchiasp.NodeValue "action","checkname",0,False
		API_enchiasp.NodeValue "username",strUserName,1,False
		Md5OLD = 1
		SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
		Md5OLD = 0
		API_enchiasp.NodeValue "syskey",SysKey,0,False
		API_enchiasp.NodeValue "email",UserEmail,1,False
		API_enchiasp.SendHttpData
		If API_enchiasp.Status = "1" Then
			CheckTitle = "出错啦！"
			strMessage = API_enchiasp.Message
		End If
		Set API_enchiasp = Nothing
	End If
	'-----------------------------------------------------------------
End If
%>
<html>
<head>
<title>用户名检测</title>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link href=user_style.css rel=stylesheet type=text/css>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="320" border=0 align=center cellpadding=3 cellspacing=1 style="border: 1px #3795D2 solid ; background-color: #FFFFFF;font: 12px;">
<tr>
	<th style="background-color: #3795D2;color: white;font-size: 12px;font-weight:bold;height: 26;"><%=CheckTitle%></th>
</tr>
<tr>
	<td style="background-color:#F7F7F7;font-size: 12px;height: 20;color:blue">&nbsp;&nbsp;<span id="jump">15</span> 秒钟后系统将自动关闭本窗口</td>
</tr>
<tr>
	<td style="background-color:#F7F7F7;font-size: 12px;height: 42;"><%=strMessage%></td>
</tr>
<tr>
	<td align=center style="background-color:#F0F0F0;font-size: 12px;height: 20;">【<a href='javascript:onclick=window.close()'>关闭本窗口</a>】</td>
</tr>
</table>
<script language="JavaScript">
function countDown(secs){
	jump.innerText=secs;if(--secs>0)setTimeout("countDown("+secs+")",1000);
}
countDown(15);
setTimeout('window.close();', 15000);
</script>
</body>
</html>