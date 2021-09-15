<!--#include file="savefile.asp"-->
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
Class Download_Cls
	Private sUploadDir
	Private nAllowSize
	Private sAllowExt
	Private sOriginalFileName
	Private sSaveFileName
	Private sPathFileName

	Public Property Get RemoteFileName()
		RemoteFileName = sOriginalFileName
	End Property

	Public Property Get LocalFileName()
		LocalFileName = sSaveFileName
	End Property

	Public Property Get LocalFilePath()
		LocalFilePath = sPathFileName
	End Property

	Public Property Let RemoteDir(ByVal strDir)
		sUploadDir = strDir
	End Property

	Public Property Let AllowMaxSize(ByVal intSize)
		nAllowSize = intSize
	End Property

	Public Property Let AllowExtName(ByVal strExt)
		sAllowExt = strExt
	End Property

	Private Sub Class_Initialize()
		On Error Resume Next
		Script_Object = "Scripting.FileSystemObject"
		sUploadDir = "UploadFile/"
		nAllowSize = 500
		sAllowExt = "gif|jpg|png|bmp"
	End Sub

	Public Function ChangeRemote(sHTML)
		On Error Resume Next
		Dim s_Content
		s_Content = sHTML
		On Error Resume Next
		Dim re, s, RemoteFileUrl, SaveFileName, SaveFileType
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}(([A-Za-z0-9_-])+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(" & sAllowExt & ")))"
		Set s = re.Execute(s_Content)
		Dim a_RemoteUrl(), n, i, bRepeat
		n = 0
		' 转入无重复数据
		For Each RemoteFileUrl In s
			If n = 0 Then
				n = n + 1
				ReDim a_RemoteUrl(n)
				a_RemoteUrl(n) = RemoteFileUrl
			Else
				bRepeat = False
				For i = 1 To UBound(a_RemoteUrl)
					If UCase(RemoteFileUrl) = UCase(a_RemoteUrl(i)) Then
						bRepeat = True
						Exit For
					End If
				Next
				If bRepeat = False Then
					n = n + 1
					ReDim Preserve a_RemoteUrl(n)
					a_RemoteUrl(n) = RemoteFileUrl
				End If
			End If
		Next
		' 开始替换操作
		Dim nFileNum, sContentPath,strFilePath
		sContentPath = RelativePath2RootPath(sUploadDir)
		nFileNum = 0
		For i = 1 To n
			SaveFileType = Mid(a_RemoteUrl(i), InStrRev(a_RemoteUrl(i), ".") + 1)
			SaveFileName = GetRndFileName(SaveFileType)
			strFilePath = sUploadDir & SaveFileName
			If SaveRemoteFile(strFilePath, a_RemoteUrl(i)) = True Then
				nFileNum = nFileNum + 1
				If nFileNum > 0 Then
					sOriginalFileName = sOriginalFileName & "|"
					sSaveFileName = sSaveFileName & "|"
					sPathFileName = sPathFileName & "|"
				End If
				sOriginalFileName = sOriginalFileName & Mid(a_RemoteUrl(i), InStrRev(a_RemoteUrl(i), "/") + 1)
				sSaveFileName = sSaveFileName & SaveFileName
				sPathFileName = sPathFileName & sContentPath & SaveFileName
				s_Content = Replace(s_Content, a_RemoteUrl(i), sContentPath & SaveFileName, 1, -1, 1)
			End If
		Next

		ChangeRemote = s_Content
	End Function
	
	Public Function RelativePath2RootPath(url)
		Dim sTempUrl
		sTempUrl = url
		If Left(sTempUrl, 1) = "/" Then
			RelativePath2RootPath = sTempUrl
			Exit Function
		End If

		Dim sWebEditorPath
		sWebEditorPath = Request.ServerVariables("SCRIPT_NAME")
		sWebEditorPath = Left(sWebEditorPath, InStrRev(sWebEditorPath, "/") - 1)
		Do While Left(sTempUrl, 3) = "../"
			sTempUrl = Mid(sTempUrl, 4)
			sWebEditorPath = Left(sWebEditorPath, InStrRev(sWebEditorPath, "/") - 1)
		Loop
		RelativePath2RootPath = sWebEditorPath & "/" & sTempUrl
	End Function
	
	Public Function GetRndFileName(sExt)
		Dim sRnd
		Randomize
		sRnd = Int(900 * Rnd) + 100
		GetRndFileName = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & sRnd & "." & sExt
	End Function
End Class
%>