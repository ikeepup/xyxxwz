<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/UploadCls.Asp"-->
<script language=JavaScript>
// 文件上传成功接口操作
function doInterfaceUpload(strValue){
	if (strValue=="") return;
	var objLinkUpload = parent.document.getElementsByName("UploadFileList")[0];
	if (objLinkUpload){
		if (objLinkUpload.value!=""){
			objLinkUpload.value = objLinkUpload.value + "|";
		}
		objLinkUpload.value = objLinkUpload.value + strValue;
		objLinkUpload.fireEvent("onchange");
	}
}
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
Server.ScriptTimeOut = 18000
Dim UploadObject,AllowFileSize,AllowFileExt
Dim sUploadDir,SaveFileName,PathFileName,url
Dim sAction,sType,SaveFilePath,UploadPath
UploadObject = CInt(enchiasp.UploadClass)   '上传文件对象 --- 0=无组件上传,1=Aspupload3.0组件,2=SA-FileUp 4.0组件
if  CLng(enchiasp.MaxFileSize)>  CLng(enchiasp.UploadFileSize) then
	AllowFileSize =CLng(enchiasp.UploadFileSize)
else
	AllowFileSize = CLng(enchiasp.MaxFileSize)
end if


AllowFileExt = enchiasp.UpFileType
AllowFileExt = Replace(Replace(Replace(UCase(AllowFileExt), "ASP", ""), "ASPX", ""), "|", ",")
url = Split(Request.ServerVariables("SERVER_PROTOCOL"), "/")(0) & "://" & Request.ServerVariables("HTTP_HOST")
sType = UCase(Request.QueryString("sType"))
If enchiasp.CheckPost=False Then
	Call Returnerr(Postmsg)
	Response.End
End If
Select Case ChannelID
	Case 0
		If stype = "AD" Then
			UploadPath = "adfile/UploadPic/"
			sUploadDir = enchiasp.InstallDir & UploadPath
		ElseIf stype = "LINK" Then
			UploadPath = "link/UploadPic/"
			sUploadDir = enchiasp.InstallDir & UploadPath
		Else
			UploadPath = "UploadFile/"
			sUploadDir = enchiasp.InstallDir & UploadPath
		End If
	Case Else
		UploadPath = "UploadPic/"
		sUploadDir = enchiasp.InstallDir & enchiasp.ChannelDir & UploadPath
End Select

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
	SaveFilePath = UploadPath & SaveFilePath
	If CInt(enchiasp.Modules) = 1 Then
		Call InputScript(enchiasp.InstallDir & enchiasp.ChannelDir & SaveFilePath)
	End If
	Call OutScript(SaveFilePath)
Else
	Call UploadMain
End If
Sub UploadFile()
	Dim Upload,FilePath,sFilePath,FormName,File,F_FileName
	Dim PreviewSetting,DrawInfo,Previewpath,strPreviewPath
	Dim PreviewName,F_Viewname,MakePreview
	'-- 是否生成缩略图片
	MakePreview = False
	Previewpath = enchiasp.InstallDir & enchiasp.ChannelDir
	strPreviewPath = "UploadPic/" & CreatePath(Previewpath & "UploadPic/")
	PreviewPath = Previewpath & strPreviewpath
	PreviewSetting = Split(enchiasp.PreviewSetting, ",")
	If CInt(PreviewSetting(2)) = 1 Then
		DrawInfo = PreviewSetting(5)
	ElseIf CInt(PreviewSetting(2)) = 2 Then
		DrawInfo = enchiasp.InstallDir & PreviewSetting(10)
	Else
		DrawInfo = ""
	End If
	If DrawInfo = "0" Then
		DrawInfo = ""
		PreviewSetting(2) = 0
	End If
	sFilePath = CreatePath(sUploadDir) '按日期生成目录
	FilePath = sUploadDir & sFilePath
	
	Set Upload = New UpFile_Cls
	Upload.UploadType = UploadObject				'设置上传组件类型
	Upload.UploadPath = FilePath					'设置上传路径
	Upload.MaxSize	= AllowFileSize					'单位 KB
	Upload.InceptMaxFile = 10					'每次上传文件个数上限
	Upload.InceptFileType	= AllowFileExt				'设置上传文件限制
	Upload.ChkSessionName	= "uploadPic"
	'预览图片设置
	Upload.MakePreview		= MakePreview
	Upload.PreviewType		= CInt(PreviewSetting(0))		'设置预览图片组件类型
	Upload.PreviewImageWidth	= CInt(PreviewSetting(3))		'设置预览图片宽度
	Upload.PreviewImageHeight	= CInt(PreviewSetting(4))		'设置预览图片高度
	Upload.DrawImageWidth		= CInt(PreviewSetting(13))		'设置水印图片或文字区域宽度
	Upload.DrawImageHeight		= CInt(PreviewSetting(14))		'设置水印图片或文字区域高度
	Upload.DrawGraph		= CCur(PreviewSetting(11))		'设置水印透明度
	Upload.DrawFontColor		= PreviewSetting(7)			'设置水印文字颜色
	Upload.DrawFontFamily		= PreviewSetting(8)			'设置水印文字字体格式
	Upload.DrawFontSize		= CInt(PreviewSetting(6))		'设置水印文字字体大小
	Upload.DrawFontBold		= CInt(PreviewSetting(9))		'设置水印文字是否粗体
	Upload.DrawInfo			= DrawInfo				'设置水印文字信息或图片信息
	Upload.DrawType			= CInt(PreviewSetting(2))		'0=不加载水印 ，1=加载水印文字，2=加载水印图片
	Upload.DrawXYType		= CInt(PreviewSetting(15))		'"0" =左上，"1"=左下,"2"=居中,"3"=右上,"4"=右下
	Upload.DrawSizeType		= CInt(PreviewSetting(1))		'"0"=固定缩小，"1"=等比例缩小
	If PreviewSetting(12)<>"" Or PreviewSetting(12)<>"0" Then
		Upload.TransitionColor	= PreviewSetting(12)			'透明度颜色设置
	End If
	'执行上传
	Upload.SaveUpFile
	If Upload.ErrCodes<>0 Then
		Response.write ("<script>alert('错误："& Upload.Description & "');history.go(-1)</script>")
		Exit Sub
	End If
	If Upload.Count > 0 Then
		For Each FormName In Upload.UploadFiles
			Set File = Upload.UploadFiles(FormName)
			SaveFilePath = sFilePath & File.FileName
			F_FileName = FilePath & File.FileName
			'创建预览及水印图片
			If Upload.PreviewType<>999 and File.FileType=1 then
				PreviewName = "p" & Replace(File.FileName,File.FileExt,"") & "jpg"
				F_Viewname = Previewpath & PreviewName
				'创建预览图片:Call CreateView(原始文件的路径,预览文件名及路径,原文件后缀)
				Upload.CreateView F_FileName,F_Viewname,File.FileExt
				If CBool(MakePreview) Then
					Call OutPreview(strPreviewPath & PreviewName)
				End If
			End If
			Set File = Nothing
		Next
	Else
		Call OutAlertScript("sorry！请选择一个有效的上传文件。")
		Exit Sub
	End If
	Set Upload = Nothing
End Sub

Sub UploadMain()
	Dim PostRanNum
	Randomize
	PostRanNum = Int(900*rnd)+1000
	Session("uploadPic") = Cstr(PostRanNum)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文件上传</title>
<LINK href="style.css" type=text/css rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></head>
<body leftMargin=0 topMargin=0 marginwidth=0 marginheight=0>
<table style="width:100%;height:100%" border="0" cellspacing="0" cellpadding="0" align=center>
<tr vAlign=top>
<td class=TableRow1>
<table border="0" cellspacing="0" cellpadding="0">
<form action='?action=save&ChannelID=<%=ChannelID%>&sType=<%=sType%>' method="post" name="myform" enctype="multipart/form-data">
<tr vAlign=top>
<td align=center noWrap valign=top>
<INPUT TYPE="hidden" name="uploadPic" value="<%=PostRanNum%>">
<input type="file" name="file1" size=50>
<input type="submit" name="Submit" value="开始上传">
<input type="hidden" name="Rename" value="0">
</td>
</tr><tr vAlign=top><TD colspan=4 class=TableRow1 valign=top>
允许上传的文件类型：<%=AllowFileExt%> <br>
允许上传的大小：<font color=red><B><%=CStr(enchiasp.UploadFileSize)%></B></font>&nbsp;KB
您今天还可以上传<font color=red><B><%=CLng(GroupSetting(21)) - CLng(UserToday(1)) %></B></font>个文件</td>
</tr></form></table></td>
</tr></table>
</body>
</html>
<%
End Sub

Private Sub OutScript(url)
	Response.Write "<script language=javascript>" & vbCrLf
	Response.Write "parent.document.myform.ImageUrl.value='" & url & "';" & vbCrLf
	If CInt(enchiasp.Modules) = 1 Then
		Response.Write "doInterfaceUpload('" & url & "')" & vbCrLf
	End If
	Response.Write "alert('文件上传成功!\n"&url&"');" & vbCrLf
	'Response.Write "history.go(-1);" & vbCrLf
	Response.Write "location.replace('" & Request.ServerVariables("HTTP_REFERER") & "')" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub

Sub OutPreview(url)
	Response.Write "<script language=javascript>" & vbCrLf
	Response.Write "parent.document.myform.ImageUrl1.value='" & url & "';" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub

Sub InputScript(url)
	Dim FileExtName
	FileExtName = Mid(url,InStrRev(url, ".")+1)
	FileExtName = LCase(FileExtName)
	Response.Write "<script>" & vbNewLine
	Response.Write "if ('" & FileExtName & "'=='gif' || '" & FileExtName & "'=='jpg' || '" & FileExtName & "'=='png' || '" & FileExtName & "'=='bmp'){" & vbNewLine
	Response.Write "img='<img src=""" & url & """>'" & vbNewLine
	Response.Write "}else{" & vbNewLine
	Response.Write "	img='<a target=""_blank"" href=""" & url & """>相关附件</a>'" & vbNewLine
	Response.Write "}" & vbNewLine
	Response.Write "document.oncontextmenu = new Function('return false')" & vbNewLine
	Response.Write "parent.IframeID.document.body.innerHTML+='\n'+img+''" & vbNewLine
	Response.Write "</script>" & vbNewLine
End Sub
%>