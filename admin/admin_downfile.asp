<!--#include file="setup.asp" -->
<!--#include file="check.asp"-->
<%
Server.ScriptTimeOut = 18000
Admin_header
Response.Write "<base target=""_self"">" & vbNewLine
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
Dim Action
If CInt(ChannelID) = 0 Then ChannelID = 2
Action = LCase(Request("action"))
If Action = "down" Then
	Call BeginDown
Else
	Call showmain
End If

If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
Public Sub showmain()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
	Response.Write " <tr>"
	Response.Write "   <th colspan=2>远程软件下载</th>"
	Response.Write " </tr>"
	Response.Write "<form name=myform method=post action=admin_downfile.asp?action=down>" & vbNewLine
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<td class=tablerow1>远程URL：</td>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=text name=FileUrl size=60></td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<td class=tablerow2>保存路径：</td>" & vbNewLine
	Response.Write "	<td class=tablerow2><input type=text name=FilePath size=60 value='" & enchiasp.InstallDir & enchiasp.ChannelDir & "UploadFile/'><br><b>注意：</b>请输入完整的路径和文件名</td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<td class=tablerow1>是否显示下载进度：</td>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=radio name=IsPross value='0'> 不显示&nbsp;&nbsp;<input type=radio name=IsPross value='1'> 速度&nbsp;&nbsp;"
	Response.Write "	<input type=radio name=IsPross value='2' checked> 进度&nbsp;&nbsp;<input type=radio name=IsPross value='3'> 速度+进度&nbsp;&nbsp; </td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<td class=tablerow1></td>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=button value=' 关闭窗口 ' onClick=""window.close();"" class=Button>&nbsp;&nbsp;<input type=submit value=' 开始下载 ' class=Button></td>" & vbNewLine
	Response.Write "</tr></form>" & vbNewLine
	Response.Write "</table>" & vbNewLine
End Sub

Public Sub BeginDown()
	Dim strFilePath,IsPross
	Dim FileUrl,FilePath,strFileName
	If Trim(Request.Form("FileUrl")) = "" Then
		Call AlertInform("远程URL不能为空！","admin_downfile.asp")
		Exit Sub
	Else
		FileUrl = Request.Form("FileUrl")
	End If
	If Trim(Request.Form("FilePath")) = "" Then
		Call AlertInform("保存路径不能为空！","admin_downfile.asp")
		Exit Sub
	Else
		FilePath = Trim(Request.Form("FilePath"))
	End If
	If Right(FilePath,1) = "/" Or Right(FilePath,1) = "\" Then
		Call AlertInform("请输入完整的路径和文件名称！","admin_downfile.asp")
		Exit Sub
	End If
	
	IsPross = enchiasp.ChkNumeric(Request.Form("IsPross"))
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
	Response.Write " <tr>"
	Response.Write "   <th><span id=txt1>正在采集，请稍候....</span></th>"
	Response.Write " </tr>"
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<td class=tablerow1 style=""line-height: 20px;"">" & vbNewLine
	
	If IsPross <> 0 Then
		If IsPross = 2 Or IsPross = 3 Then
			RevealProgress
		End If
		
		If IsPross = 1 Or IsPross = 3 Then
			Response.Write "<div><span id=Proess1 style=""color: #008800;""></span></div>"
		End If
	End If

	Response.Write "<br><strong><font color=blue>软件大小：</font></strong><span id=txt2 style=""font-size:9pt;color:red"">0</span>"
	Response.Write "</td></tr>" & vbNewLine
	Response.Write "<tr><td class=tablerow1><div align=center><a href='admin_downfile.asp' class=showmenu>停止下载任务</a>"
	'Response.Write "<a href='#'  onClick=""window.close();"" class=showmenu>停止下载任务</a>"
	Response.Write "</div></td></tr></table>" & vbNewLine
	Response.Flush
	Dim Myenchicms

	On Error Resume Next
	Set Myenchicms = CreateObject("Gatherer.SoftCollection")
	If Err Then
		Response.Write "<br /><br /><br /><div align=""center"" style=""font-size:18px;color:red"">非常遗憾，您的服务器不支持采集组件！</div>"
		Err.Clear
		Set Myenchiasp = Nothing
		Response.End
	End If
	Myenchicms.Connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ChkMapPath(DBPath)
	Myenchicms.ChannelPath = enchiasp.InstallDir & enchiasp.ChannelDir
	Myenchicms.DBConnstr = Connstr
	Myenchicms.SqlDataBase = IsSqlDataBase
	strFilePath = Myenchicms.FileDownload(FileUrl,FilePath,IsPross,200,vbNullString)
	If len(strFilePath) > 0 Then
		strFileName = Mid(strFilePath, InStrRev(strFilePath, "\") + 1)
		Response.Write "<br><br><div align=center><a href='#' onClick=""window.returnValue='" & strFileName & "|" & Myenchicms.SoftSize & "';window.close();"" class=showmenu>文件下载完成！点击关闭窗口将自动复制文件名</a></div>"
		Response.Write "<div align=center>保存文件路径：<input type=text name=sPath size=75 value='" & strFilePath & "'></div>"
		Response.Write "<div align=center><input type=button value=' 关闭窗口 ' onClick=""window.returnValue='" & strFileName & "|" & Myenchicms.SoftSize & "';window.close();"" class=Button></div><br><br>"
	Else
		Response.Write "<br><br><div align=center>文件下载失败!</div>"
		Response.Write "<div align=center><input type=button value=' 关闭窗口 ' onClick=""window.close();"" class=Button></div><br><br>"
	End If
	
	
	Set Myenchicms = Nothing
End Sub

Sub RevealProgress()
	Response.Write "<div><table width='400' align=left border=0 cellspacing=1 cellpadding=1>" & vbCrLf
	Response.Write "<tr> " & vbCrLf
	Response.Write "<td style=""border: 1px #384780 solid ;background-color: #FFFFFF;"">" & vbCrLf
	Response.Write "<table width=0 id=tablePros name=tablePros border=0 cellspacing=0 cellpadding=0>" & vbCrLf
	Response.Write "<tr height=12><td bgcolor='#0650D2'>" & vbCrLf
	Response.Write "</td></tr></table></td></tr>" & vbCrLf
	Response.Write "</table><br></div>" & vbCrLf
	Response.Flush
End Sub

%>
