<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/UploadCls.Asp"-->
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
Server.ScriptTimeOut = 18000
Dim UploadObject,AllowFileSize,AllowFileExt
Dim sUploadDir,SaveFileName,PathFileName,FileExtName
Dim sAction,sType
UploadObject = CInt(enchiasp.UploadClass)   '�ϴ��ļ����� --- 0=������ϴ�,1=Aspupload3.0���,2=SA-FileUp 4.0���
if  CLng(enchiasp.MaxFileSize)>  CLng(enchiasp.UploadFileSize) then
	AllowFileSize =CLng(enchiasp.UploadFileSize)
else
	AllowFileSize = CLng(enchiasp.MaxFileSize)
end if
if  enchiasp.UpFileType="" then
	AllowFileExt = enchiasp.UploadFileType
else
	AllowFileExt = enchiasp.UpFileType
end if

AllowFileExt = Replace(Replace(Replace(UCase(AllowFileExt), "ASP", ""), "ASPX", ""), "|", ",")

Select Case ChannelID
	Case 0
		sUploadDir = enchiasp.InstallDir & "UploadFile/"
	Case Else
		sUploadDir = enchiasp.InstallDir & enchiasp.ChannelDir & "UploadPic/"
End Select
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ļ��ϴ�</title>
<LINK href="style.css" type=text/css rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></head>
<body leftMargin=0 topMargin=0 marginwidth=0 marginheight=0>
<table style="width:100%;height:100%" border="0" cellspacing="0" cellpadding="3" align=center>
<tr vAlign=top>
<td class=TableRow1>
<%
sAction = UCase(Trim(Request.QueryString("action")))
If sAction = "SAVE" Then
	If Not ChkAdmin("UploadFile") Then
		Response.Write ("<script>alert('�Բ���!��û���ϴ��ļ���Ȩ��');history.go(-1)</script>")
		Response.End
	End If
	Select Case UploadObject
		Case 0,1,2,3
			Call UploadFile
		Case 999
			Response.Write ("<script>alert('��ϵͳδ�����ϴ�����!');history.go(-1)</script>")
			Response.End
		Case Else
			Response.Write ("<script>alert('��ϵͳδ�����ϴ�����!');history.go(-1)</script>")
			Response.End
	End Select
	PathFileName = SaveFileName
	'Call OutScript(PathFileName)
	FileExtName=mid(PathFileName,InStrRev(PathFileName, ".")+1)
	FileExtName=LCase(FileExtName)
%>
<script>
if ("<%=FileExtName%>"=="gif" || "<%=FileExtName%>"=="jpg" || "<%=FileExtName%>"=="png" || "<%=FileExtName%>"=="bmp"){
img="<img src=<%=PathFileName%>>"
}else{
img="<a target=_blank href=<%=PathFileName%>>��ظ���</a>"
}
document.oncontextmenu = new Function('return false')
parent.IframeID.document.body.innerHTML+='\n'+img+''
</script>
<input type=text name=file1 size=70 value='<%=PathFileName%>'> <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="�����ϴ��ļ�" class=Button>
<%
Else
%>
<table border="0" cellspacing="0" cellpadding="0">
<form action='?action=save&ChannelID=<%=ChannelID%>' method=post name=myform enctype="multipart/form-data">
<tr>
<td align=center noWrap>
<input type="file" name="file1" size=50>
<input type="submit" name="Submit" value="��ʼ�ϴ�">
</td>
</tr><tr vAlign=top><TD colspan=4 class=TableRow1 valign=top>
�����ϴ����ļ����ͣ�<%=AllowFileExt%> ����С��<font color=red><B><%=CStr(enchiasp.UploadFileSize)%></B></font>&nbsp;KB</td>
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
	Dim Upload,FilePath,sFilePath,FormName,File,F_FileName
	Dim PreviewSetting,DrawInfo,Previewpath,strPreviewPath
	Dim PreviewName,F_Viewname,MakePreview
	'-- �Ƿ���������ͼƬ
	MakePreview = True
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
	sFilePath = CreatePath(sUploadDir) '����������Ŀ¼
	FilePath = sUploadDir & sFilePath

	Set Upload = New UpFile_Cls
	Upload.UploadType = UploadObject				'�����ϴ��������
	Upload.UploadPath = FilePath					'�����ϴ�·��
	Upload.MaxSize	= AllowFileSize					'��λ KB
	Upload.InceptMaxFile = 10					'ÿ���ϴ��ļ���������
	Upload.InceptFileType	= AllowFileExt				'�����ϴ��ļ�����
	'Upload.ChkSessionName	= "uploadfile"
	'Ԥ��ͼƬ����
	Upload.MakePreview		= MakePreview
	Upload.PreviewType		= CInt(PreviewSetting(0))		'����Ԥ��ͼƬ�������
	Upload.PreviewImageWidth	= CInt(PreviewSetting(3))		'����Ԥ��ͼƬ���
	Upload.PreviewImageHeight	= CInt(PreviewSetting(4))		'����Ԥ��ͼƬ�߶�
	Upload.DrawImageWidth		= CInt(PreviewSetting(13))		'����ˮӡͼƬ������������
	Upload.DrawImageHeight		= CInt(PreviewSetting(14))		'����ˮӡͼƬ����������߶�
	Upload.DrawGraph		= CCur(PreviewSetting(11))		'����ˮӡ͸����
	Upload.DrawFontColor		= PreviewSetting(7)			'����ˮӡ������ɫ
	Upload.DrawFontFamily		= PreviewSetting(8)			'����ˮӡ���������ʽ
	Upload.DrawFontSize		= CInt(PreviewSetting(6))		'����ˮӡ���������С
	Upload.DrawFontBold		= CInt(PreviewSetting(9))		'����ˮӡ�����Ƿ����
	Upload.DrawInfo			= DrawInfo				'����ˮӡ������Ϣ��ͼƬ��Ϣ
	Upload.DrawType			= CInt(PreviewSetting(2))		'0=������ˮӡ ��1=����ˮӡ���֣�2=����ˮӡͼƬ
	Upload.DrawXYType		= CInt(PreviewSetting(15))		'"0" =���ϣ�"1"=����,"2"=����,"3"=����,"4"=����
	Upload.DrawSizeType		= CInt(PreviewSetting(1))		'"0"=�̶���С��"1"=�ȱ�����С
	If PreviewSetting(12)<>"" Or PreviewSetting(12)<>"0" Then
		Upload.TransitionColor	= PreviewSetting(12)			'͸������ɫ����
	End If
	'ִ���ϴ�
	Upload.SaveUpFile
	If Upload.ErrCodes<>0 Then
		Response.write ("<script>alert('����"& Upload.Description & "');history.go(-1)</script>")
		Exit Sub
	End If
	If Upload.Count > 0 Then
		For Each FormName In Upload.UploadFiles
			Set File = Upload.UploadFiles(FormName)
			F_FileName = FilePath & File.FileName
			SaveFileName = F_FileName
			'����Ԥ����ˮӡͼƬ
			If Upload.PreviewType<>999 and File.FileType=1 then
				PreviewName = "p" & Replace(File.FileName,File.FileExt,"") & "jpg"
				F_Viewname = Previewpath & PreviewName
				'����Ԥ��ͼƬ:Call CreateView(ԭʼ�ļ���·��,Ԥ���ļ�����·��,ԭ�ļ���׺)
				Upload.CreateView F_FileName,F_Viewname,File.FileExt
				If CBool(MakePreview) Then
					Call OutPreview(strPreviewPath & PreviewName)
				End If
			End If
			Set File = Nothing
		Next
	Else
		Call OutAlertScript("sorry����ѡ��һ����Ч���ϴ��ļ���")
		Exit Sub
	End If
	Set Upload = Nothing
End Sub

Sub OutScript(url)
	Response.Write "<script language=javascript>" & vbCrLf
	Response.Write "parent.document.myform.filePath.value='" & url & "';" & vbCrLf
	Response.Write "alert('�ļ��ϴ��ɹ�!\n"&url&"');" & vbCrLf
	Response.Write "history.go(-1);" & vbCrLf
	'Response.Write "location.replace('" & Request.ServerVariables("HTTP_REFERER") & "')" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub

Sub OutPreview(url)
	Response.Write "<script language=javascript>" & vbCrLf
	Response.Write "parent.document.myform.ImageUrl.value='" & url & "';" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub
%>