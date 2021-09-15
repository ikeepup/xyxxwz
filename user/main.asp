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
Dim Rs,SQL
SQL = "select * from ECCMS_User where username='" & enchiasp.membername & "'"
Set Rs = enchiasp.Execute(SQL)
If Rs("usermsg") > 0 Then
	Response.Write "<bgsound src=""images/mail.wav"" border=0>"
End If
Call InnerLocation("会员管理首页 - 欢迎 <font color=red>" & enchiasp.membername & "</font> 登录控制面板")
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=Usertableborder>
<tr>
	<th colspan=2>用户控制面板 -- 首页</th>            
</tr>
<tr>
	<td width="50%" class=Usertablerow1><b class=userfont2>用户名称：</b><font color=red><%=enchiasp.membername%></font></td>
	<td width="50%" class=Usertablerow1><b class=userfont2>用户身份：</b><font color=red><%=enchiasp.membergroup%></font></td>
</tr>
<tr>
	<td class=Usertablerow2><b class=userfont2>用户昵称：</b><%=enchiasp.menbernickname%></td>
	<td class=Usertablerow2><b class=userfont2>真实姓名：</b><%=Rs("TrueName")%></td>
</tr>
<tr>
	<td class=Usertablerow1><b class=userfont2>账户余额：</b>￥<%=FormatNumber(Rs("usermoney"),,-1)%> 元</td>
	<td class=Usertablerow1><b class=userfont2>已经消费：</b>￥<%=FormatNumber(Rs("prepaid"),,-1)%> 元</td>
</tr>
<tr>
	<td class=Usertablerow2><b class=userfont2>可用点数：</b><%=Rs("userpoint")%> 点</td>
	<td class=Usertablerow2><b class=userfont2>用户经验：</b><%=Rs("experience")%></td>
</tr>
<tr>
	<td class=Usertablerow1><b class=userfont2>用户魅力：</b><%=Rs("charm")%></td>
	<td class=Usertablerow1><b class=userfont2>注册日期：</b><%=Rs("JoinTime")%></td>
</tr>
<tr>
	<td class=Usertablerow2><b class=userfont2>会员类型：</b><%If Rs("UserGrade") = 999 Then
				Response.Write "管理员"
			Else
				If Rs("UserClass") = 0 Then
					Response.Write "计点会员"
				ElseIf Rs("UserClass") = 1 Then
					Response.Write "计时会员"
				Else
					Response.Write "到期会员"
				End If
			End If%></td>
	<td class=Usertablerow2><b class=userfont2>到期时间：</b><%=Rs("ExpireTime")%></td>
</tr>
<tr>
	<td class=Usertablerow1><b class=userfont2>上次登录时间：</b><%=Request.Cookies("enchiasp_net")("LastTimeDate")%></td>
	<td class=Usertablerow1><b class=userfont2>上次登录IP：</b><%=Request.Cookies("enchiasp_net")("LastTimeIP")%></td>
</tr>
<tr>
	<td class=Usertablerow2><b class=userfont2>登录次数：</b><%=Rs("userlogin")%> 次</td>
	<td class=Usertablerow2>
<%
If Rs("UserClass") > 0 Then
	Response.Write "<b class=userfont2>友情提示：</b>"
	If DateDiff("D", CDate(Rs("ExpireTime")), Now()) < 0 Then
		Response.Write "您的账号使用时限还有 <font color=red><b>"
		Response.Write DateDiff("D", Now(), CDate(Rs("ExpireTime")))
		Response.Write "</b></font> 天"
	Else
		Response.Write "<font color=red>您的账号已过期,请联系管理员！</font>"
	End If
End If
%>
	
	</td>
</tr>
<tr>
	<td class=Usertablerow1><a href="usersms.asp">我的收件箱 
<%
If Rs("usermsg") > 0 Then
	Response.Write "(<b class=userfont1>" & Rs("usermsg") & "</b>)"
Else
	Response.Write "(<font color=""#808080"">" & Rs("usermsg") & "</font>)"
End If
%>
	</a></td>
	<td class=Usertablerow1><a href="changepsw.asp">修改密码</a> | <a href="changeinfo.asp">修改资料</a> | <a href="addmoney.asp">会员充值</a></td>            
</tr>

</table>
<%
Rs.Close:Set Rs = Nothing
%>
<!--#include file="foot.inc"-->