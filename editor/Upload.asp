<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/UploadCls.Asp"-->
<!--#include file="../inc/cls_down.Asp"-->
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
If enchiasp.CheckPost = False Then
	Call OutAlertScript("���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ��")
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
Call InitUpload()		' ��ʼ���ϴ�����


Dim sAction
sAction = UCase(Trim(Request.QueryString("action")))

Select Case sAction
Case "REMOTE"
	Call LoadRemote()			' Զ���Զ���ȡ
Case "SAVE"
	Call ShowForm()			' ��ʾ�ϴ���
	Call DoSave()			' ���ļ�
Case Else
	Call ShowForm()			' ��ʾ�ϴ���
End Select

CloseConn
Sub ShowForm() 
%>
<HTML>
<HEAD>
<TITLE>�ļ��ϴ�</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
body, a, table, div, span, td, th, input, select{font:9pt;font-family: "����", Verdana, Arial, Helvetica, sans-serif;}
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
// ����ϴ���
function CheckUploadForm() {
	if (!IsExt(document.myform.uploadfile.value,sAllowExt)){
		parent.UploadError("��ʾ��\n\n��ѡ��һ����Ч���ļ���\n֧�ֵ��ļ���ʽ�У�"+sAllowExt+"����");
		return false;
	}
	return true
}

// �ύ�¼��������
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

// �ϴ�����װ�����
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
' �������
Sub DoSave()
	If Session("AdminName") = "" Then
		Call OutScript("parent.UploadError('�Բ����㻹û�е�½�����ϴ��ļ���')")
		Response.End
	End If
	If Not enchiasp.CheckAdmin("UploadFile") Then
		Call OutScript("parent.UploadError('�Բ���!��û���ϴ��ļ���Ȩ��')")
		Response.End
	End If
	Select Case nUploadObject
		Case 0,1,2,3
			Call UploadFile 
		Case 999
			Call OutScript("parent.UploadError('�Բ���ϵͳ�Ѿ��ر��ϴ��ļ����ܣ�')")
			Response.End
		Case Else
			Call OutScript("parent.UploadError('�Բ���ϵͳ�Ѿ��ر��ϴ��ļ����ܣ�')")
			Response.End
	End Select
	
	if enchiasp.StopUpload="1" then
		Call OutScript("parent.UploadError('�Բ���ϵͳ�Ѿ��ر��ϴ��ļ����ܣ�')")
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

' �Զ���ȡԶ���ļ�
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

	Response.Write "<HTML><HEAD><TITLE>Զ���ϴ�</TITLE><meta http-equiv='Content-Type' content='text/html; charset=gb2312'></head><body>" & _
		"<input type=hidden id=UploadText value=""" & inHTML(sContent) & """>" & _
		"</body></html>"

	Call OutScriptNoBack("parent.setHTML(UploadText.value);try{parent.addUploadFile('" & sOriginalFileName & "', '" & sSaveFileName & "', '" & SaveFilePath & "');} catch(e){} parent.remoteUploadOK();")
	Set objFile = Nothing
End Sub

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
	strUploadDir = sFilePath
	sUploadDir = sUploadDir & strUploadDir
	
	Set Upload = New UpFile_Cls
	Upload.UploadType = nUploadObject				'�����ϴ��������
	Upload.UploadPath = FilePath					'�����ϴ�·��
	Upload.MaxSize	= nAllowSize					'��λ KB
	Upload.InceptMaxFile = 10					'ÿ���ϴ��ļ���������
	Upload.InceptFileType	= Replace(sAllowExt, "|", ",")				'�����ϴ��ļ�����
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
			sSaveFileName = File.FileName
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

' ����ͻ��˽ű�
Sub OutScript(str)
	Response.Write "<script language=javascript>" & vbcrlf
	Response.Write str
	Response.Write ";history.back()" & vbcrlf
	Response.Write "</script>" & vbcrlf
End Sub
Sub OutScriptNoBack(str)
	Response.Write "<script language=javascript>" & str & "</script>" & vbcrlf
End Sub
' ��ʼ���ϴ���������
Sub InitUpload()
	sType = UCase(Trim(Request.QueryString("type")))
	sBaseUrl = "1"        '·��ģʽ --- 0=���·��,1=���Ը�·��,2����ȫ·��
	nUploadObject = CInt(enchiasp.UploadClass)   '�ϴ��ļ����� --- 0=������ϴ�,1=�����ϴ����,2=���Ʒ��ϴ����
	
	if  CLng(enchiasp.MaxFileSize)>  CLng(enchiasp.UploadFileSize) then
		nAllowSize =CLng(enchiasp.UploadFileSize)
	else
		nAllowSize = CLng(enchiasp.MaxFileSize)
	end if
	   '�ϴ��ļ�����
	if  enchiasp.UpFileType="" then
		sAllowExt = enchiasp.UploadFileType
	else
		sAllowExt = enchiasp.UpFileType
	end if
 
	If ChannelID <> 0 Then
		if ChannelID=-1 then
			sUploadDir = enchiasp.InstallDir & "fengmian/"
		else
		sUploadDir = enchiasp.InstallDir & enchiasp.ChannelDir    '�ϴ��ļ�·��
		end if
	Else
		sUploadDir = enchiasp.InstallDir    '�ϴ��ļ�·��
	End If

	Select Case sType
		Case "REMOTE"     'Զ���ļ�����
			UploadPath = "UploadPic/"
			sUploadDir = sUploadDir & UploadPath    '�ϴ��ļ�·��
			sAllowExt = "gif|jpg|bmp|png|jpge"           '�ϴ��ļ�����
		Case "FILE"       '�ϴ��ļ�����
			UploadPath = "UploadFile/"
			sUploadDir = sUploadDir & UploadPath    '�ϴ��ļ�·��
		Case "MEDIA"      '�ϴ�ý������
			UploadPath = "UploadFile/"
			sUploadDir = sUploadDir & UploadPath    '�ϴ��ļ�·��
		Case "FLASH"      '�ϴ���������
			UploadPath = "UploadFile/"
			sUploadDir = sUploadDir & UploadPath    '�ϴ��ļ�·��
	Case Else         '�ϴ�ͼƬ����
		UploadPath = "UploadPic/"
		sUploadDir = sUploadDir & UploadPath    '�ϴ��ļ�·��
	End Select
	' �κ�����¶��������ϴ�asp�ű��ļ�
	sAllowExt = Replace(Replace(UCase(sAllowExt), "ASP", ""), "ASPX", "")
End Sub
'================================================
' �õ���ȫ�ַ���,�ڲ�ѯ��ʹ��
'================================================
Function Get_SafeStr(str)
	Get_SafeStr = Replace(Replace(Replace(Trim(str), "'", ""), Chr(34), ""), ";", "")
End Function
'================================================
' ȥ��Html��ʽ�����ڴ����ݿ���ȡ��ֵ���������ʱ
' ע�⣺value="?"���һ��Ҫ��˫����
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