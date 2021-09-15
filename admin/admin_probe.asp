<%@ Language="VBScript" %>
<%' Option Explicit %>
<script language=javascript>
 <!--
 var startTime,endTime;
 var d=new Date();
 startTime=d.getTime();
 //-->
 </script>
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
if session("AdminName")="" then
	Response.Redirect "admin_klogin.asp"
	Response.End
end if
'不使用输出缓冲区，直接将运行结果显示在客户端
'Response.Buffer = False
Dim starttime
starttime=timer()*1000
'声明待检测数组
Dim ObjTotest(26,4)

ObjTotest(0,0) = "MSWC.AdRotator"
ObjTotest(1,0) = "MSWC.BrowserType"
ObjTotest(2,0) = "MSWC.NextLink"
ObjTotest(3,0) = "MSWC.Tools"
ObjTotest(4,0) = "MSWC.Status"
ObjTotest(5,0) = "MSWC.Counters"
ObjTotest(6,0) = "IISSample.ContentRotator"
ObjTotest(7,0) = "IISSample.PageCounter"
ObjTotest(8,0) = "MSWC.PermissionChecker"
ObjTotest(9,0) = "Scripting.FileSystemObject"
	ObjTotest(9,1) = "(FSO 文本文件读写)"
ObjTotest(10,0) = "adodb.connection"
	ObjTotest(10,1) = "(ADO 数据对象)"
'ObjTotest(11,0) = "enchicmsCMS.SiteMainObject"
	'ObjTotest(11,1) = "(恩池网站管理组件)"	
ObjTotest(12,0) = "SoftArtisans.FileUp"
	ObjTotest(12,1) = "(SA-FileUp 文件上传)"
ObjTotest(13,0) = "SoftArtisans.FileManager"
	ObjTotest(13,1) = "(SoftArtisans 文件管理)"
ObjTotest(14,0) = "LyfUpload.UploadFile"
	ObjTotest(14,1) = "(刘云峰的文件上传组件)"
ObjTotest(15,0) = "Persits.Upload.1"
	ObjTotest(15,1) = "(ASPUpload 文件上传)"
ObjTotest(16,0) = "w3.upload"
	ObjTotest(16,1) = "(Dimac 文件上传)"

ObjTotest(17,0) = "JMail.SmtpMail"
	ObjTotest(17,1) = "(Dimac JMail 邮件收发) "
ObjTotest(18,0) = "CDONTS.NewMail"
	ObjTotest(18,1) = "(虚拟 SMTP 发信)"
ObjTotest(19,0) = "Persits.MailSender"
	ObjTotest(19,1) = "(ASPemail 发信)"
ObjTotest(20,0) = "SMTPsvg.Mailer"
	ObjTotest(20,1) = "(ASPmail 发信)"
ObjTotest(21,0) = "DkQmail.Qmail"
	ObjTotest(21,1) = "(dkQmail 发信)"
ObjTotest(22,0) = "Geocel.Mailer"
	ObjTotest(22,1) = "(Geocel 发信)"
ObjTotest(23,0) = "IISmail.Iismail.1"
	ObjTotest(23,1) = "(IISmail 发信)"
ObjTotest(24,0) = "SmtpMail.SmtpMail.1"
	ObjTotest(24,1) = "(SmtpMail 发信)"
	
ObjTotest(25,0) = "SoftArtisans.ImageGen"
	ObjTotest(25,1) = "(SA 的图像读写组件)"
ObjTotest(26,0) = "W3Image.Image"
	ObjTotest(26,1) = "(Dimac 的图像读写组件)"

public IsObj,VerObj,TestObj

'检查预查组件支持情况及版本

dim i
for i=0 to 26
	on error resume next
	IsObj=false
	VerObj=""
	'dim TestObj
	TestObj=""
	set TestObj=server.CreateObject(ObjTotest(i,0))
	If -2147221005 <> Err then		
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if
	ObjTotest(i,2)=IsObj
	ObjTotest(i,3)=VerObj
next

'检查组件是否被支持及组件版本的子程序
sub ObjTest(strObj)
	on error resume next
	IsObj=false
	VerObj=""
	TestObj=""
	set TestObj=server.CreateObject (strObj)
	If -2147221005 <> Err then		'
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if	
End sub
%>
<html>
<head>
<title>服务器的相关信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="90%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
    <tr> 
      <th width="27%" height=25  Class=TableTitle>服 务 器 信 息</th></tr>
    <tr>
      <td height=22 class="TableBody1">
<font class=fonts>你的服务器是否支持ASP</font>
<br>出现以下情况即表示您的空间不支持ASP：
<br>1、访问本文件时提示下载。
<br>2、访问本文件时看到类似“&lt;%@ Language="VBScript" %&gt;”的文字。
</td></tr></table><br>
	<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
          <tr><th colspan=2>■ 服务器的有关参数</th></tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;服务器名</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;服务器IP</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;服务器端口</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;服务器时间</td><td class="TableRow1">&nbsp;<%=now%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;IIS版本</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;脚本超时时间</td><td class="TableRow1">&nbsp;<%=Server.ScriptTimeout%> 秒</td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;本文件路径</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("PATH_TRANSLATED")%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;服务器CPU数量</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;服务器解译引擎</td><td class="TableRow1">&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;服务器操作系统</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("OS")%></td>
	  </tr>
	</table>
<br>
<%
Dim strClass
	strClass = Trim(Request.Form("classname"))
	If "" <> strClass then
	Response.Write "<br>您指定的组件的检查结果："
	Dim Verobj1
	ObjTest(strClass)
	  If Not IsObj then 
		Response.Write "<br><font color=red>很遗憾，该服务器不支持 " & strclass & " 组件！</font>"
	  Else
		if VerObj="" or isnull(VerObj) then 
			Verobj1="无法取得该组件版本"
		Else
			Verobj1="该组件版本是：" & VerObj
		End If
		Response.Write "<br><font class=fonts>恭喜！该服务器支持 " & strclass & " 组件。" & verobj1 & "</font>"
	  End If
	  Response.Write "<br>"
	end if
	%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=2>■ IIS自带的ASP组件</th></tr>
	<tr height=23 align=center><td width=70% class=TableTitle>组 件 名 称</td><td width=30% class=TableTitle>支持及版本</td></tr>
	<%For i=0 to 10%>
	<tr>
		<td height=20 class=TableRow2>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td class=TableRow1>&nbsp;<%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		Else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
	</tr>
	<%next%>
</table>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=2>■ 常见的文件上传和管理组件</th></tr>
	<tr height=23 align=center><td width=70% class=TableTitle>组 件 名 称</td><td width=30% class=TableTitle>支持及版本</td></tr>
	<%For i=12 to 16%>
	<tr height="18" class=backq>
		<td height=20 class=TableRow2>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td class=TableRow1>&nbsp;<%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		Else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
	</tr>
	<%next%>
</table>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=2>■ 常见的收发邮件组件</th></tr>
	<tr height=23 align=center><td width=70% class=TableTitle>组 件 名 称</td><td width=30% class=TableTitle>支持及版本</td></tr>
	<%For i=17 to 24%>
	<tr height="18" class=backq>
		<td height=20 class=TableRow2>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td class=TableRow1>&nbsp;<%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		Else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
	</tr>
	<%next%>
</table>
<BR>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=2>■ 图像处理组件</th></tr>
	<tr height=23 align=center><td width=70% class=TableTitle>组 件 名 称</td><td width=30% class=TableTitle>支持及版本</td></tr>
	<%For i=25 to 26%>
	<tr height="20">
		<td height=20 class=TableRow2>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td class=TableRow1>&nbsp;<%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		Else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
	</tr>
	<%next%>
</table>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th>■ 其他组件支持情况检测</th></tr>
<tr height=23 align=center><td width=70% class=TableTitle>
在下面的输入框中输入你要检测的组件的ProgId或ClassId。</td></tr>
<FORM action=<%=Request.ServerVariables("SCRIPT_NAME")%> method=post id=form1 name=form1>
	<tr height="23">
		<td align=center class=TableRow1 height=30><input class=kuang type=text value="" name="classname" size=40>
<INPUT type=submit value=" 确 定 " class=kuang id=submit1 name=submit1>
<INPUT type=reset value=" 重 填 " class=kuang id=reset1 name=reset1> 
</td>
	  </tr>
</FORM>
</table>

<%if ObjTest("Scripting.FileSystemObject") then

	set fsoobj=server.CreateObject("Scripting.FileSystemObject")

%>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=6>■ 服务器磁盘信息</th></tr>
  <tr height="20" align=center>
	<td width="100" class=TableTitle>盘符和磁盘类型</td>
	<td width="50" class=TableTitle>就绪</td>
	<td width="80" class=TableTitle>卷标</td>
	<td width="60" class=TableTitle>文件系统</td>
	<td width="80" class=TableTitle>可用空间</td>
	<td width="80" class=TableTitle>总空间</td>
  </tr>
<%

	' 测试磁盘信息的想法来自“COCOON ASP 探针”
	
	set drvObj=fsoobj.Drives
	for each d in drvObj
%>
  <tr height="18" align=center>
	<td class=TableRow2 align="right"><%=cdrivetype(d.DriveType) & " " & d.DriveLetter%>:</td>
<%
	if d.DriveLetter = "A" then	'为防止影响服务器，不检查软驱
		Response.Write "<td class=TableRow1></td><td class=TableRow1></td><td class=TableRow1></td><td class=TableRow1></td><td class=TableRow1></td>"
	else
%>
	<td class=TableRow1><%=cIsReady(d.isReady)%></td>
	<td class=TableRow1><%=d.VolumeName%></td>
	<td class=TableRow1><%=d.FileSystem%></td>
	<td align="right" class=TableRow1><%=cSize(d.FreeSpace)%></td>
	<td align="right" class=TableRow1><%=cSize(d.TotalSize)%></td>
<%
	end if
%>
  </tr>
<%
	next
%>
</td></tr>
</table>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=5>■ 当前文件夹信息 <%
	dPath = server.MapPath("./")
	set dDir = fsoObj.GetFolder(dPath)
	set dDrive = fsoObj.GetDrive(dDir.Drive)
%>
文件夹: <%=dPath%></th></tr>
  <tr height="23" align="center">
	<td width="75" class=TableTitle>已用空间</td>
	<td width="75" class=TableTitle>可用空间</td>
	<td width="75" class=TableTitle>文件夹数</td>
	<td width="75" class=TableTitle>文件数</td>
	<td width="150" class=TableTitle>创建时间</td>
  </tr>
  <tr height="20" align="center">
	<td class=TableRow1><%=cSize(dDir.Size)%></td>
	<td class=TableRow1><%=cSize(dDrive.AvailableSpace)%></td>
	<td class=TableRow1><%=dDir.SubFolders.Count%></td>
	<td class=TableRow1><%=dDir.Files.Count%></td>
	<td class=TableRow1><%=dDir.DateCreated%></td>
  </tr>
</td></tr>
</table>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=2>■ 磁盘文件操作速度测试</th></tr>
<tr height="20" align=center>
	<td colspan=2 class=TableRow1><%
	
	Response.Write "正在重复创建、写入和删除文本文件50次..."

	dim thetime3,tempfile,iserr

iserr=false
	t1=timer
	tempfile=server.MapPath("./") & "\aspchecktest.txt"
	for i=1 to 50
		Err.Clear

		set tempfileOBJ = FsoObj.CreateTextFile(tempfile,true)
		if Err <> 0 then
			Response.Write "创建文件错误！"
			iserr=true
			Err.Clear
			exit for
		end if
		tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
		if Err <> 0 then
			Response.Write "写入文件错误！"
			iserr=true
			Err.Clear
			exit for
		end if
		tempfileOBJ.close
		Set tempfileOBJ = FsoObj.GetFile(tempfile)
		tempfileOBJ.Delete 
		if Err <> 0 then
			Response.Write "删除文件错误！"
			iserr=true
			Err.Clear
			exit for
		end if
		set tempfileOBJ=nothing
	next
	t2=timer
if iserr <> true then
	thetime3=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！<font color=red>" & thetime3 & "毫秒</font>。"

%>
</td></tr>
  <tr height=18>
	<td height=20 class=TableRow2>&nbsp;<font color=red>您的服务器: <%=Request.ServerVariables("SERVER_NAME")%></font>&nbsp;</td><td class=TableRow1>&nbsp;<font color=red><%=thetime3%></font></td>
  </tr>
</table>
<%
end if
	
	set fsoobj=nothing

end if%>
<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=3>■ ASP脚本解释和运算速度测试</th></tr>
<tr height="20" align=center>
	<td colspan=3 class=TableRow1>
<%
	'因为只进行50万次计算，所以去掉了是否检测的选项而直接检测
	
	Response.Write "整数运算测试，正在进行50万次加法运算..."
	dim t1,t2,lsabc,thetime,thetime2
	t1=timer
	for i=1 to 500000
		lsabc= 1 + 1
	next
	t2=timer
	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！<font color=red>" & thetime & "毫秒</font>。<br>"


	Response.Write "浮点运算测试，正在进行20万次开方运算..."
	t1=timer
	for i=1 to 200000
		lsabc= 2^0.5
	next
	t2=timer
	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！<font color=red>" & thetime2 & "毫秒</font>。<br>"
%>
</td></tr>
<tr height=18>
	<td height=20 class=TableRow2>&nbsp;<font color=red>您的服务器: <%=Request.ServerVariables("SERVER_NAME")%></font>&nbsp;</td><td class=TableRow1 nowrap>&nbsp;<font color=red><%=thetime%></font></td><td  class=TableRow1 nowrap>&nbsp;<font color=red><%=thetime2%></font></td>
  </tr>
</table>
<BR>
<table border=0 width=90% align=center cellspacing=0 cellpadding=0>
<tr><td align=center style="LINE-HEIGHT: 150%">
页面装载时间：<script language=javascript>d=new Date();endTime=d.getTime();document.write((endTime-startTime)/1000);</script> 秒<BR> Copyright (c) 2002-2005 <a href="http://www.enchi.com.cn" target="_blank"><font color=#6C70AA><b>enchi<font color=#CC0000>.com</font></b></font></a>. All Rights Reserved .


</td></tr>
</table>
</BODY>
</HTML>

<%
function cdrivetype(tnum)
    Select Case tnum
        Case 0: cdrivetype = "未知"
        Case 1: cdrivetype = "可移动磁盘"
        Case 2: cdrivetype = "本地硬盘"
        Case 3: cdrivetype = "网络磁盘"
        Case 4: cdrivetype = "CD-ROM"
        Case 5: cdrivetype = "RAM 磁盘"
    End Select
end function

function cIsReady(trd)
    Select Case trd
		case true: cIsReady="<font class=fonts><b>√</b></font>"
		case false: cIsReady="<font color='red'><b>×</b></font>"
	End Select
end function

function cSize(tSize)
    if tSize>=1073741824 then
		cSize=int((tSize/1073741824)*1000)/1000 & " GB"
    elseif tSize>=1048576 then
    	cSize=int((tSize/1048576)*1000)/1000 & " MB"
    elseif tSize>=1024 then
		cSize=int((tSize/1024)*1000)/1000 & " KB"
	else
		cSize=tSize & "B"
	end if
end function
%>







