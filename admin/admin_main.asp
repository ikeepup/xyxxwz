<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
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
Dim theInstalledObjects(4)
theInstalledObjects(0) = "Persits.Jpeg"
theInstalledObjects(1) = "Scripting.FileSystemObject"
theInstalledObjects(2) = "adodb.connection"

theInstalledObjects(3) = "JMail.SMTPMail"
theInstalledObjects(4) = "CDONTS.NewMail"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<title>管理页面</title>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<link rel="stylesheet" href="style.css" type="text/css">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="5" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<table cellpadding="2" cellspacing="1" border="0" width="96%" class="tableBorder" align=center>
<tr>
<th colspan=2 height=25>系统信息统计</th>
</tr>
<tr>
<td class=BodyTitle colspan=2 height=25>系统信息统计　　
<%
On Error Resume Next
Response.Write "今日下载：<font color=red><b>"
SQL = "SELECT SUM(DayHits) FROM ECCMS_SoftList WHERE isAccept>0 And Datediff('d',HitsTime,Now())=0"
Set Rs = enchiasp.Execute(SQL)
If Rs.BOF And Rs.EOF Then
	Response.Write 0
Else
	Response.Write enchiasp.CheckNumeric(Rs(0))
End If
Set Rs = Nothing
Response.Write "</b></font> 次　"
Response.Write "今日浏览：<font color=red><b>"
SQL = "SELECT SUM(DayHits) FROM ECCMS_Article WHERE isAccept>0 And Datediff('d',HitsTime,Now())=0"
Set Rs = enchiasp.Execute(SQL)
If Rs.BOF And Rs.EOF Then
	Response.Write 0
Else
	Response.Write enchiasp.CheckNumeric(Rs(0))
End If
Set Rs = Nothing
Response.Write "</b></font> 次　"
Response.Write "今日观看：<font color=red><b>"
SQL = "SELECT SUM(DayHits) FROM ECCMS_FlashList WHERE isAccept>0 And Datediff('d',HitsTime,Now())=0"
Set Rs = enchiasp.Execute(SQL)
If Rs.BOF And Rs.EOF Then
	Response.Write 0
Else
	Response.Write enchiasp.CheckNumeric(Rs(0))
End If
Set Rs = Nothing
Response.Write "</b></font> 次　"
Response.Write "当前在线：<font color=red><b>"
SQL = "SELECT COUNT(id) FROM ECCMS_Online WHERE DateDIff('s',lastTime,Now()) < 20*60"
Set Rs = enchiasp.Execute(SQL)
If Rs.BOF And Rs.EOF Then
	Response.Write 0
Else
	Response.Write enchiasp.CheckNumeric(Rs(0))
End If
Set Rs = Nothing
Response.Write "</b></font> 人　"
%>
</td>
</tr>
<tr>
<td width="50%"  class="TableRow2" height=23>服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="TableRow1">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="50%" class="TableRow2" height=23>站点物理路径：<%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
<td width="50%" class="TableRow1">AspJpeg组件：
<%If Not IsObjInstalled(theInstalledObjects(0)) Then%><font color="#FF0066"><b>×</b></font><%else%><b>√</b><%end if%>
</td>

</td>

</tr>
<tr>
<td width="50%" class="TableRow2" height=23>FSO文本读写：<%If Not IsObjInstalled(theInstalledObjects(1)) Then%><font color="#FF0066"><b>×</b></font><%else%><b>√</b><%end if%></td>
<td width="50%" class="TableRow1">数据库使用：<%If Not IsObjInstalled(theInstalledObjects(2)) Then%><font color="#FF0066"><b>×</b></font><%else%><b style="color:blue"><%If IsSqlDataBase = 1 Then%>MS SQL<%else%>ACCESS<%end if%></b><%end if%></td>
</tr>
<tr>
<td width="50%" class="TableRow2" height=23>Jmail组件支持：<%If Not IsObjInstalled(theInstalledObjects(3)) Then%><font color="#FF0066"><b>×</b></font><%else%><b>√</b><%end if%></td>
<td width="50%" class="TableRow1">CDONTS组件支持：<%If Not IsObjInstalled(theInstalledObjects(4)) Then%><font color="#FF0066"><b>×</b></font><%else%><b>√</b><%end if%></td>
</tr>
</tr>
<tr><td colspan=2 class="TableRow1" height=25><B>快捷操作选项：</B> <a href=admin_config.asp>网站基本设置</a>&nbsp;
<a href=CleanCache.asp>重建系统缓存</a>&nbsp;
<a href=admin_user.asp>会员管理</a>&nbsp;
<a href=admin_online.asp>在线人数统计</a>&nbsp;
<a href=admin_template.asp>模板样式总管理</a></td>
</tr>
</table>
<BR>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
      <tr>
	<th colspan=2 height=25>网站管理系统说明</th>
	</tr>
	<tr>
	<td width="60" class="TableRow2" height=23>高级管理员</td>
	<td class="TableRow1" style="LINE-HEIGHT: 150%">拥有所有权限。<BR>第一次使用时请到<font color=Red>用户管理</font>-<a href=admin_password.asp target=main><font color=Red>管理员密码修改</font></A>重新设置管理密码</td>
      </tr>
      <tr>
	<td class="TableRow2" height=23>普通管理员</td>
	<td class="TableRow1" style="LINE-HEIGHT: 150%">需要高级管理员给予权限。<BR>注:<a href=admin_master.asp><font color=Red>在管理员管理－设置权限!</font></A><br>
	   </td>
      </tr>
      <tr>
	<td class="TableRow2" height=23>使用设置</td>
	<td class="TableRow1" style="LINE-HEIGHT: 150%">第一次使用<font color=Red>恩池软件</font>网站管理系统<BR>
	 点击左边管理导航菜单中的“<a href=admin_config.asp><font color=Red>常规设置-基本设置</A></font>”<BR>对你的网站信息和一些网站配置参数进行配置。
	   </td>
      </tr>
    </table>
<BR>
    <BR>

    <table cellpadding="3" cellspacing="1" border="0" width="96%" class="tableBorder" align=center>
      <tr>
	<th colspan=2 height=25>恩池软件网站管理系统开发</th>

      <tr>
	<td width="60" class="TableRow2" height=23>程序制作</td>
	<td class="TableRow1">运城市恩池软件科技开发有限公司</td>
      </tr>
      <tr>
	<td class="TableRow2" height=23>联系方式</td>
	<td class="TableRow1">E_mail：liuyunfan@163.com<br>QQ：21556923<br>电话：0359-8698845<br>
	</td>
      </tr>
      <tr>
	<td class="TableRow2" height=23>程序主页</td>
	<td class="TableRow1"><a href="http://www.enchi.com.cn/" target=_blank>www.enchi.com.cn</a>
	</td>
      </tr>
    </table>
<%
If CInt(enchiasp.VersionID) <> 0 Then
	enchiasp.Execute("UPDATE ECCMS_Config SET VersionID=0")
End If
Admin_footer
CloseConn
%>






















