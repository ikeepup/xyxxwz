<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/cls_public.asp"-->
<!--#include file="../inc/ShopChannel.asp"-->
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
Dim shopid,sortid,showid,indexid
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID < 3 Then ChannelID = 3
shopid = enchiasp.ChkNumeric(Request("shopid"))
sortid = enchiasp.ChkNumeric(Request("sortid"))
showid = enchiasp.ChkNumeric(Request("showid"))
indexid = enchiasp.ChkNumeric(Request("indexid"))
enchicms.Channel = ChannelID
enchicms.MainChannel
enchicms.ShowFlush = showid
If indexid > 0 Then
	enchicms.CreateShopIndex
End If
If shopid > 0 Then
	enchicms.LoadShopInfo(shopid)
End If
Set enchicms = Nothing
CloseConn
%>