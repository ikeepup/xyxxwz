<!--#include file="config.asp"-->
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
Dim sAction,sType,AutoRename
UploadObject = CInt(enchiasp.UploadClass)   '�ϴ��ļ����� --- 0=������ϴ�,1=Aspupload3.0���,2=SA-FileUp 4.0���
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
<title>�ļ��ϴ�</title>
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
	PathFileName = SaveFileName
	'Call OutScript(PathFileName)
%>
<script language=javascript>
parent.document.myform.filePath.value='<%=PathFileName%>';
</script>
<input type=text name=file1 size=60 value='<%=PathFileName%>'> <input type="button" name="Submit4" onclick="javascript:location.replace('<%=Request.ServerVariables("HTTP_REFERER")%>');" value="�����ϴ��ļ�" class="Button"><br>
<font color=blue>��ѵ�ַ���Ƶ���Ӧ�������</font>
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
<input type="submit" name="Submit" value="��ʼ�ϴ�">
<input type="checkbox" name="Rename" value='1'> ���Զ�����
</td>
</tr><tr vAlign=top><TD colspan=4 class=TableRow1 valign=top>
�����ϴ����ļ����ͣ�<%=AllowFileExt%> ����С��<font color=red><B>

<% 
if  CLng(enchiasp.MaxFileSize)>  CLng(enchiasp.UploadFileSize) then
	response.write Cstr(enchiasp.UploadFileSize)
else
	response.write Cstr(enchiasp.MaxFileSize)
end if


%>
</B></font>&nbsp;KB<br>
�����컹�����ϴ�<font color=red><B><%=CLng(GroupSetting(21)) - CLng(UserToday(1)) %></B></font>���ļ�</td></td>
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
	sFilePath = CreatePath(sUploadDir) '����������Ŀ¼
	FilePath = sUploadDir & sFilePath
	
	Set Upload = New UpFile_Cls
	Upload.UploadType = UploadObject				'�����ϴ��������
	Upload.UploadPath = FilePath					'�����ϴ�·��
	Upload.MaxSize	= AllowFileSize					'��λ KB
	Upload.InceptMaxFile = 10					'ÿ���ϴ��ļ���������
	Upload.InceptFileType	= AllowFileExt				'�����ϴ��ļ�����
	Upload.ChkSessionName	= "uploadfile"
	'ִ���ϴ�
	Upload.SaveUpFile
	If Upload.ErrCodes<>0 Then
		Response.write ("<script>alert('����"& Upload.Description & "');history.go(-1)</script>")
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
		Response.Write ("<script>alert('sorry����ѡ��һ����Ч���ϴ��ļ���');history.go(-1)</script>")
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



