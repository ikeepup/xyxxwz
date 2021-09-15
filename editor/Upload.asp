<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/UploadCls.Asp"-->
<!--#include file="../inc/cls_down.Asp"-->
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
If enchiasp.CheckPost = False Then
	Call OutAlertScript("您提交的数据不合法，请不要从外部提交。")
End If
If Session("AdminName") = Empty Then Response.End
Dim sType
Dim sAllowExt, nAllowSize, sUploadDir, nUploadObject, sBaseUrl, sContentPath
Dim sFileExt, sOriginalFileName, sSaveFileName, sPathFileName, nFileNum
Dim ChannelID,SaveFilePath,UploadPath,strUploadDir


If Request("ChannelID") <> "" And Request("ChannelID") <> 0 and Request("ChannelID") <>"-1" Then
	ChannelID = CInt(Request("ChannelID"))
	enchiasp.ReadChannel(ChannelID)
Else
	if Request("ChannelID") ="-1" then
		ChannelID=-1
			
	else
		ChannelID = 0
	end if
End If
Call InitUpload()		' 初始化上传变量


Dim sAction
sAction = UCase(Trim(Request.QueryString("action")))

Select Case sAction
Case "REMOTE"
	Call LoadRemote()			' 远程自动获取
Case "SAVE"
	Call ShowForm()			' 显示上传表单
	Call DoSave()			' 存文件
Case Else
	Call ShowForm()			' 显示上传表单
End Select

CloseConn
Sub ShowForm() 
%>
<HTML>
<HEAD>
<TITLE>文件上传</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
body, a, table, div, span, td, th, input, select{font:9pt;font-family: "宋体", Verdana, Arial, Helvetica, sans-serif;}
body {padding:0px;margin:0px}
</style>
<script language="JavaScript" src="dialog/dialog.js"></script>
</head>
<body bgcolor=menu>
<form action="?action=save&type=<%=sType%>&ChannelID=<%=ChannelID%>" method=post name=myform enctype="multipart/form-data">
<input type=file name=uploadfile size=1 style="width:100%" onchange="originalfile.value=this.value">
<input type=hidden name=originalfile value="">
</form>
<script language=javascript>
var sAllowExt = "<%=sAllowExt%>";
// 检测上传表单
function CheckUploadForm() {
	if (!IsExt(document.myform.uploadfile.value,sAllowExt)){
		parent.UploadError("提示：\n\n请选择一个有效的文件，\n支持的文件格式有（"+sAllowExt+"）！");
		return false;
	}
	return true
}

// 提交事件加入检测表单
var oForm = document.myform ;
oForm.attachEvent("onsubmit", CheckUploadForm) ;
if (! oForm.submitUpload) oForm.submitUpload = new Array() ;
oForm.submitUpload[oForm.submitUpload.length] = CheckUploadForm ;
if (! oForm.originalSubmit) {
	oForm.originalSubmit = oForm.submit ;
	oForm.submit = function() {
		if (this.submitUpload) {
			for (var i = 0 ; i < this.submitUpload.length ; i++) {
				this.submitUpload[i]() ;
			}
		}
		this.originalSubmit() ;
	}
}

// 上传表单已装入完成
try {
	parent.UploadLoaded();
}
catch(e){
}
</script>
</body>
</html>
<% 
End Sub 
' 保存操作
Sub DoSave()
	If Session("AdminName") = "" Then
		Call OutScript("parent.UploadError('对不起！你还没有登陆不能上传文件。')")
		Response.End
	End If
	If Not enchiasp.CheckAdmin("UploadFile") Then
		Call OutScript("parent.UploadError('对不起!您没有上传文件的权限')")
		Response.End
	End If
	Select Case nUploadObject
		Case 0,1,2,3
			Call UploadFile 
		Case 999
			Call OutScript("parent.UploadError('对不起系统已经关闭上传文件功能！')")
			Response.End
		Case Else
			Call OutScript("parent.UploadError('对不起系统已经关闭上传文件功能！')")
			Response.End
	End Select
	
	if enchiasp.StopUpload="1" then
		Call OutScript("parent.UploadError('对不起系统已经关闭上传文件功能！')")
		Response.End

	end if
	Select Case sBaseUrl
		Case "0"
			sContentPath = sUploadDir
		Case "1"
			sContentPath = RelativePath2RootPath(sUploadDir)
		Case "2"
			sContentPath = RootPath2DomainPath(RelativePath2RootPath(sUploadDir))
	End Select
	sPathFileName = sContentPath & sSaveFileName
	SaveFilePath = UploadPath & strUploadDir & sSaveFileName
	Call OutScript("parent.UploadSaved('" & sPathFileName & "');var obj=parent.dialogArguments.dialogArguments;if (!obj) obj=parent.dialogArguments;try{obj.addUploadFile('" & sOriginalFileName & "', '" & sSaveFileName & "', '" & SaveFilePath & "');} catch(e){}")

End Sub

' 自动获取远程文件
Sub LoadRemote()
	Dim sContent, i,objFile
	strUploadDir = CreatePath(sUploadDir)
	sUploadDir = sUploadDir & strUploadDir
	For i = 1 To Request.form("enchicms_UploadText").Count 
		sContent = sContent & Request.form("enchicms_UploadText")(i) 
	Next
	If sAllowExt <> "" Then
		Set objFile = New Download_Cls
		objFile.RemoteDir = sUploadDir
		objFile.AllowMaxSize = nAllowSize
		objFile.AllowExtName = sAllowExt
		sContent = objFile.ChangeRemote(sContent)
		sOriginalFileName = objFile.RemoteFileName
		sSaveFileName = objFile.LocalFileName
		sPathFileName = objFile.LocalFilePath
		SaveFilePath = Replace(sPathFileName, enchiasp.InstallDir & enchiasp.ChannelDir, "",1,-1,1)
	End If

	Response.Write "<HTML><HEAD><TITLE>远程上传</TITLE><meta http-equiv='Content-Type' content='text/html; charset=gb2312'></head><body>" & _
		"<input type=hidden id=UploadText value=""" & inHTML(sContent) & """>" & _
		"</body></html>"

	Call OutScriptNoBack("parent.setHTML(UploadText.value);try{parent.addUploadFile('" & sOriginalFileName & "', '" & sSaveFileName & "', '" & SaveFilePath & "');} catch(e){} parent.remoteUploadOK();")
	Set objFile = Nothing
End Sub

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
	strUploadDir = sFilePath
	sUploadDir = sUploadDir & strUploadDir
	
	Set Upload = New UpFile_Cls
	Upload.UploadType = nUploadObject				'设置上传组件类型
	Upload.UploadPath = FilePath					'设置上传路径
	Upload.MaxSize	= nAllowSize					'单位 KB
	Upload.InceptMaxFile = 10					'每次上传文件个数上限
	Upload.InceptFileType	= Replace(sAllowExt, "|", ",")				'设置上传文件限制
	'Upload.ChkSessionName	= "uploadfile"
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
			sSaveFileName = File.FileName
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

' 输出客户端脚本
Sub OutScript(str)
	Response.Write "<script language=javascript>" & vbcrlf
	Response.Write str
	Response.Write ";history.back()" & vbcrlf
	Response.Write "</script>" & vbcrlf
End Sub
Sub OutScriptNoBack(str)
	Response.Write "<script language=javascript>" & str & "</script>" & vbcrlf
End Sub
' 初始化上传限制数据
Sub InitUpload()
	sType = UCase(Trim(Request.QueryString("type")))
	sBaseUrl = "1"        '路径模式 --- 0=相对路径,1=绝对根路径,2绝对全路径
	nUploadObject = CInt(enchiasp.UploadClass)   '上传文件对象 --- 0=无组件上传,1=恩池上传组件,2=刘云峰上传组件
	
	if  CLng(enchiasp.MaxFileSize)>  CLng(enchiasp.UploadFileSize) then
		nAllowSize =CLng(enchiasp.UploadFileSize)
	else
		nAllowSize = CLng(enchiasp.MaxFileSize)
	end if
	   '上传文件类型
	if  enchiasp.UpFileType="" then
		sAllowExt = enchiasp.UploadFileType
	else
		sAllowExt = enchiasp.UpFileType
	end if
 
	If ChannelID <> 0 Then
		if ChannelID=-1 then
			sUploadDir = enchiasp.InstallDir & "fengmian/"
		else
		sUploadDir = enchiasp.InstallDir & enchiasp.ChannelDir    '上传文件路径
		end if
	Else
		sUploadDir = enchiasp.InstallDir    '上传文件路径
	End If

	Select Case sType
		Case "REMOTE"     '远程文件设置
			UploadPath = "UploadPic/"
			sUploadDir = sUploadDir & UploadPath    '上传文件路径
			sAllowExt = "gif|jpg|bmp|png|jpge"           '上传文件类型
		Case "FILE"       '上传文件设置
			UploadPath = "UploadFile/"
			sUploadDir = sUploadDir & UploadPath    '上传文件路径
		Case "MEDIA"      '上传媒体设置
			UploadPath = "UploadFile/"
			sUploadDir = sUploadDir & UploadPath    '上传文件路径
		Case "FLASH"      '上传动画设置
			UploadPath = "UploadFile/"
			sUploadDir = sUploadDir & UploadPath    '上传文件路径
	Case Else         '上传图片设置
		UploadPath = "UploadPic/"
		sUploadDir = sUploadDir & UploadPath    '上传文件路径
	End Select
	' 任何情况下都不允许上传asp脚本文件
	sAllowExt = Replace(Replace(UCase(sAllowExt), "ASP", ""), "ASPX", "")
End Sub
'================================================
' 得到安全字符串,在查询中使用
'================================================
Function Get_SafeStr(str)
	Get_SafeStr = Replace(Replace(Replace(Trim(str), "'", ""), Chr(34), ""), ";", "")
End Function
'================================================
' 去除Html格式，用于从数据库中取出值填入输入框时
' 注意：value="?"这边一定要用双引号
'================================================
Function inHTML(str)
	Dim sTemp
	sTemp = str
	inHTML = ""
	If IsNull(sTemp) = True Then
		Exit Function
	End If
	sTemp = Replace(sTemp, "&", "&amp;")
	sTemp = Replace(sTemp, "<", "&lt;")
	sTemp = Replace(sTemp, ">", "&gt;")
	sTemp = Replace(sTemp, Chr(34), "&quot;")
	inHTML = sTemp
End Function
%>