<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/cls_public.asp"-->
<!--#include file="../inc/FlashChannel.asp"-->
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
Dim flashid,sortid,showid,indexid
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID < 5 Then ChannelID = 5
flashid = enchiasp.ChkNumeric(Request("flashid"))
sortid = enchiasp.ChkNumeric(Request("sortid"))
showid = enchiasp.ChkNumeric(Request("showid"))
indexid = enchiasp.ChkNumeric(Request("indexid"))
enchicms.Channel = ChannelID
enchicms.MainChannel
enchicms.ShowFlush = showid
If indexid > 0 Then
	enchicms.CreateFlashIndex
End If
If flashid > 0 Then
	enchicms.LoadFlashInfo(flashid)
End If
Set enchicms = Nothing
CloseConn
%>