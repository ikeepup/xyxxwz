<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<LINK href="style.css" type=text/css rel=stylesheet>
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
Call RemoveAllCache()

Sub RemoveAllCache()
	Dim cachelist,i
	Call InnerHtml("UpdateInfo","<b>��ʼִ������ǰվ�㻺��</b>��")
	Cachelist=split(GetallCache(),",")
	If UBound(cachelist)>1 Then
		For i=0 to UBound(cachelist)-1
			DelCahe Cachelist(i)
			Call InnerHtml("UpdateInfo","���� <b>"&cachelist(i)&"</b> ���")
		Next
		Call InnerHtml("UpdateInfo","������"& UBound(cachelist)-1 &"���������<br>")
	Else
		Call InnerHtml("UpdateInfo","<b>��ǰվ��ȫ������������ɡ�</b>��")
	End If
End Sub

Function  GetallCache()
	Dim Cacheobj
	For Each Cacheobj in Application.Contents
		GetallCache = GetallCache & Cacheobj & ","
	Next
End Function

Sub DelCahe(MyCaheName)
	Application.Lock
	Application.Contents.Remove(MyCaheName)
	Application.unLock
End Sub

Sub InnerHtml(obj,msg)
	Response.Write "<li>"&msg&"</li>"
	Response.Flush
End Sub
%>