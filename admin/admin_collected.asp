<!--#include file="setup.asp" -->
<!--#include file="check.asp"-->
<!--#include file="include/collection.asp"-->
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
Server.ScriptTimeout = 99999
Dim enchicms
Set enchicms = New Collection_Cls
enchicms.DataPath = ChkMapPath(DBPath)
enchicms.Timeout = "1000"
enchicms.ChannelPath = enchiasp.InstallDir & enchiasp.ChannelDir
enchicms.SoftCollect
Set enchicms = Nothing
CloseConn
%>