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
'================================================
' 函数名：SaveRemoteFile
' 作  用：保存远程文件到本地
' 参  数：strFileName ----保存文件的名称
'         strRemoteUrl ----远程文件URL
' 返回值：布尔值 True/False
'================================================
Function SaveRemoteFile(ByVal strFileName, ByVal strRemoteUrl)
	Dim oStream, Retrieval, GetRemoteData
	
	SaveRemoteFile = False
	On Error Resume Next
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
	Retrieval.Open "GET", strRemoteUrl, False, "", ""
	Retrieval.Send
	If Retrieval.readyState <> 4 Then Exit Function
	If Retrieval.Status > 300 Then Exit Function
	GetRemoteData = Retrieval.ResponseBody
	Set Retrieval = Nothing

	If LenB(GetRemoteData) > 100 Then
		Set oStream = Server.CreateObject("Adodb.Stream")
		oStream.Type = 1
		oStream.Mode = 3
		oStream.Open
		oStream.Write GetRemoteData
		oStream.SaveToFile Server.MapPath(strFileName), 2
		oStream.Cancel
		oStream.Close
		Set oStream = Nothing
	Else
		Exit Function
	End If

	If Err.Number = 0 Then
		SaveRemoteFile = True
	Else
		Err.Clear
	End If
End Function
%>