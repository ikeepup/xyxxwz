<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/cls_public.asp"-->
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
Dim ChannelID,HtmlContent,ChannelRootDir,FoundErr,ErrMsg
FoundErr = False
ChannelID = 4
enchiasp.ReadChannel(ChannelID)
ChannelRootDir = enchiasp.InstallDir & enchiasp.ChannelDir
'=============================================================
'��������GuestStation
'��  �ã�ȡ�����ߵ����
'��  ����userid   ----�û�ID
'=============================================================
Function GuestStation(ByVal userid)
	On Error Resume Next
	If userid = 0 Or Not IsNumeric(userid) Then
		GuestStation = "�ο�"
		Exit Function
	End If
	Dim rsUser,sqlUser
	sqlUser = "SELECT UserGroup FROM ECCMS_User WHERE userid ="& userid
	Set rsUser = enchiasp.Execute(sqlUser)
	If rsUser.BOF And rsUser.EOF Then
		GuestStation = "�ο�"
		Set rsUser = Nothing
		Exit Function
	End If
	GuestStation = rsUser("UserGroup")
	Set rsUser = Nothing
End Function
%>
