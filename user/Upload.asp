<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/UploadCls.Asp"-->
<script language=JavaScript>
// �ļ��ϴ��ɹ��ӿڲ���
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
Dim sUploadDir,SaveFileName,PathFileName,url
Dim sAction,sType,SaveFilePath,UploadPath
UploadObject = CInt(enchiasp.UploadClass)   '�ϴ��ļ����� --- 0=������ϴ�,1=Aspupload3.0���,2=SA-FileUp 4.0���
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
		Response.Write ("<script>alert('�Բ���!��Ƶ��δ�����ϴ�����!');history.go(-1)</script>")
		Response.End
	End If
	If CInt(GroupSetting(20)) <> 1 Then
		Response.Write ("<script>alert('�Բ���!��û���ϴ��ļ���Ȩ��');history.go(-1)</script>")
		Response.End
	End If
	If CLng(UserToday(1)) => CLng(GroupSetting(21)) Then
		Response.Write ("<script>alert('�Բ���!��ÿ��ֻ���ϴ�" & GroupSetting(21) & "���ļ���');history.go(-1)</script>")
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
	'-- �Ƿ���������ͼƬ
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
	sFilePath = CreatePath(sUploadDir) '����������Ŀ¼
	FilePath = sUploadDir & sFilePath
	
	Set Upload = New UpFile_Cls
	Upload.UploadType = UploadObject				'�����ϴ��������
	Upload.UploadPath = FilePath					'�����ϴ�·��
	Upload.MaxSize	= AllowFileSize					'��λ KB
	Upload.InceptMaxFile = 10					'ÿ���ϴ��ļ���������
	Upload.InceptFileType	= AllowFileExt				'�����ϴ��ļ�����
	Upload.ChkSessionName	= "uploadPic"
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
			SaveFilePath = sFilePath & File.FileName
			F_FileName = FilePath & File.FileName
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
<title>�ļ��ϴ�</title>
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
<input type="submit" name="Submit" value="��ʼ�ϴ�">
<input type="hidden" name="Rename" value="0">
</td>
</tr><tr vAlign=top><TD colspan=4 class=TableRow1 valign=top>
�����ϴ����ļ����ͣ�<%=AllowFileExt%> <br>
�����ϴ��Ĵ�С��<font color=red><B><%=CStr(enchiasp.UploadFileSize)%></B></font>&nbsp;KB
�����컹�����ϴ�<font color=red><B><%=CLng(GroupSetting(21)) - CLng(UserToday(1)) %></B></font>���ļ�</td>
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
	Response.Write "alert('�ļ��ϴ��ɹ�!\n"&url&"');" & vbCrLf
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
	Response.Write "	img='<a target=""_blank"" href=""" & url & """>��ظ���</a>'" & vbNewLine
	Response.Write "}" & vbNewLine
	Response.Write "document.oncontextmenu = new Function('return false')" & vbNewLine
	Response.Write "parent.IframeID.document.body.innerHTML+='\n'+img+''" & vbNewLine
	Response.Write "</script>" & vbNewLine
End Sub
%>