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
'================================================
' ��������SaveRemoteFile
' ��  �ã�����Զ���ļ�������
' ��  ����strFileName ----�����ļ�������
'         strRemoteUrl ----Զ���ļ�URL
' ����ֵ������ֵ True/False
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