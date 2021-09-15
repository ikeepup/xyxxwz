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
Dim url,strUrl,strPath
Dim strInceptFile
strInceptFile = "swf,fla,jpg,jpeg,gif,png,bmp,tif,iff,mp3,wma,rm,wmv,mid,rmi,cda,avi,mpg,mpeg,ra,ram,wov,asf"
url = Replace(Replace(Replace(Request("url"), "'", ""), "%", ""), "\", "/")

If Len(url) > 3 Then
	If Left(url,1) = "/" Then
		Response.Redirect url
	End If
	If Left(url,3) = "../" Then
		Response.Redirect url
	End If
	strUrl = Left(url,10)
	If InStr(strUrl, "://") > 0 Then
		Response.Redirect  url
	End If
	If InStr(url, "/") > 0 Then
		url =Replace(url, "../", "")
		If CheckFileExt(url) Then
			strPath = Server.MapPath(".") & "\" & url
			strPath = Replace(strPath, "/", "\")
			Call downThisFile(strPath)
		End If

	Else
		Response.Redirect url
	End If
End If

Sub downThisFile(thePath)
	Response.Clear
	On Error Resume Next
	Dim stream, fileName, fileContentType
	
	fileName = split(thePath,"\")(UBound(split(thePath,"\")))
	Set stream = Server.CreateObject("adodb.stream")
	stream.Open
	stream.Type = 1
	stream.LoadFromFile(thePath)
	Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName
	Response.AddHeader "Content-Length", stream.Size
	Response.Charset = "UTF-8"
	Response.ContentType = "application/octet-stream"
	Response.BinaryWrite stream.Read 
	Response.Flush
	stream.Close
	Set stream = Nothing
End Sub
Function CheckFileExt(ByVal strFile)
	Dim ArrInceptFile
	Dim i, strFileExt
	
	On Error Resume Next
	
	If Trim(strFile) = "" Or IsEmpty(strFile) Then
		CheckFileExt = False
		Exit Function
	End If
	
	strFileExt = GetFileExtName(strFile)
	strFileExt = LCase(strFileExt)
	strInceptFile = LCase(strInceptFile)
	If Len(strInceptFile) = 0 Then
		CheckFileExt = True
		Exit Function
	End If
	ArrInceptFile = Split(strInceptFile, ",")
	
	For i = 0 To UBound(ArrInceptFile)
		If Trim(strFileExt) = Trim(ArrInceptFile(i)) Then
			CheckFileExt = True
			Exit Function
		Else
			CheckFileExt = False
		End If
	Next
	CheckFileExt = False
End Function
Function GetFileExtName(ByVal strFilePath)
	Dim strExtName
	strExtName = Mid(strFilePath, InStrRev(strFilePath, ".") + 1)
	If InStr(strExtName, "?") > 0 Then
		GetFileExtName = Left(strExtName, InStr(strExtName, "?") - 1)
	Else
		GetFileExtName = strExtName
	End If
End Function

%>