<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/UploadCls.Asp"-->
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
Server.ScriptTimeOut = 18000
Dim UploadObject,AllowFileSize,AllowFileExt
Dim sUploadDir,SaveFileName,PathFileName,FileExtName
Dim sAction,sType,AutoRename
UploadObject = CInt(enchiasp.UploadClass)   '上传文件对象 --- 0=无组件上传,1=Aspupload3.0组件,2=SA-FileUp 4.0组件
if  CLng(enchiasp.MaxFileSize)>  CLng(enchiasp.UploadFileSize) then
	AllowFileSize =CLng(enchiasp.UploadFileSize)
else
	AllowFileSize = CLng(enchiasp.MaxFileSize)
end if

AllowFileExt = enchiasp.UpFileType
AllowFileExt = Replace(Replace(Replace(UCase(AllowFileExt), "ASP", ""), "ASPX", ""), "|", ",")
If enchiasp.CheckPost=False Then
	Call Returnerr(Postmsg)
	Response.End
End If
Select Case ChannelID
	Case 0
		sUploadDir = enchiasp.InstallDir & "UploadFile/"
	Case Else
		sUploadDir = enchiasp.InstallDir & enchiasp.ChannelDir & "UploadFile/"
End Select
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文件上传</title>
<LINK href="style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR></head>
<body leftMargin=0 topMargin=0 marginwidth=0 marginheight=0>
<table style="width:100%;height:100%" border="0" cellspacing="0" cellpadding="3" align=center>
<tr vAlign=top>
<td class=TableRow1>
<%
sAction = UCase(Trim(Request.QueryString("action")))
If sAction = "SAVE" Then
	If CInt(enchiasp.StopUpload) = 1 Then
		Response.Write ("<script>alert('对不起!本频道未开放上传功能!');history.go(-1)</script>")
		Response.End
	End If
	If CInt(GroupSetting(20)) <> 1 Then
		Response.Write ("<script>alert('对不起!您没有上传文件的权限');history.go(-1)</script>")
		Response.End
	End If
	If CLng(UserToday(1)) => CLng(GroupSetting(21)) Then
		Response.Write ("<script>alert('对不起!您每天只能上传" & GroupSetting(21) & "个文件。');history.go(-1)</script>")
		Response.End
	End If
	Select Case UploadObject
		Case 0,1,2,3
			Call UploadFile
		Case 999
			Response.Write ("<script>alert('本系统未开放上传功能!');history.go(-1)</script>")
			Response.End
		Case Else
			Response.Write ("<script>alert('本系统未开放上传功能!');history.go(-1)</script>")
			Response.End
	End Select
	Dim strUserToday
	strUserToday = UserToday(0) &","& UserToday(1)+1 &","& UserToday(2) &","& UserToday(3) &","& UserToday(4) &","& UserToday(5)
	UpdateUserToday(strUserToday)
	PathFileName = SaveFileName
	'Call OutScript(PathFileName)
%>
<script language=javascript>
parent.document.myform.filePath.value='<%=PathFileName%>';
</script>
<input type=text name=file1 size=60 value='<%=PathFileName%>'> <input type="button" name="Submit4" onclick="javascript:location.replace('<%=Request.ServerVariables("HTTP_REFERER")%>');" value="继续上传文件" class="Button"><br>
<font color=blue>请把地址复制到相应的输入框</font>
<%
Else
	Dim PostRanNum
	Randomize
	PostRanNum = Int(900*rnd)+1000
	Session("uploadfile") = Cstr(PostRanNum)
%>
<table border="0" cellspacing="0" cellpadding="0">
<form action='?action=save&ChannelID=<%=ChannelID%>' method=post name=myform enctype="multipart/form-data">
<tr>
<td align="center" noWrap>
<INPUT TYPE="hidden" name="uploadfile" value="<%=PostRanNum%>">
<input type="file" name="file1" size=45>
<input type="submit" name="Submit" value="开始上传">
<input type="checkbox" name="Rename" value='1'> 不自动更名
</td>
</tr><tr vAlign=top><TD colspan=4 class=TableRow1 valign=top>
允许上传的文件类型：<%=AllowFileExt%> 　大小：<font color=red><B>

<% 
if  CLng(enchiasp.MaxFileSize)>  CLng(enchiasp.UploadFileSize) then
	response.write Cstr(enchiasp.UploadFileSize)
else
	response.write Cstr(enchiasp.MaxFileSize)
end if


%>
</B></font>&nbsp;KB<br>
您今天还可以上传<font color=red><B><%=CLng(GroupSetting(21)) - CLng(UserToday(1)) %></B></font>个文件</td></td>
</tr></form></table>
<%
End If
%>
</td>
</tr></table>
</body>
</html>
<%
Sub UploadFile()
	Dim Upload,FilePath,sFilePath,FormName,File
	sFilePath = CreatePath(sUploadDir) '按日期生成目录
	FilePath = sUploadDir & sFilePath
	
	Set Upload = New UpFile_Cls
	Upload.UploadType = UploadObject				'设置上传组件类型
	Upload.UploadPath = FilePath					'设置上传路径
	Upload.MaxSize	= AllowFileSize					'单位 KB
	Upload.InceptMaxFile = 10					'每次上传文件个数上限
	Upload.InceptFileType	= AllowFileExt				'设置上传文件限制
	Upload.ChkSessionName	= "uploadfile"
	'执行上传
	Upload.SaveUpFile
	If Upload.ErrCodes<>0 Then
		Response.write ("<script>alert('错误："& Upload.Description & "');history.go(-1)</script>")
		Exit Sub
	End If
	If Upload.Count > 0 Then
		For Each FormName In Upload.UploadFiles
			Set File = Upload.UploadFiles(FormName)
			SaveFileName = FilePath & File.FileName
			Set File = Nothing
		Next
		Call OutFilesize(Upload.MaxSize)
	Else
		Response.Write ("<script>alert('sorry！请选择一个有效的上传文件！');history.go(-1)</script>")
		Exit Sub
	End If
	Set Upload = Nothing
End Sub





Sub OutFilesize(filesize)
	Response.Write "<script language=javascript>" & vbCrLf
	Response.Write "parent.document.myform.filesize.value='" & Round(filesize/1024,2) & "';" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub
%>



