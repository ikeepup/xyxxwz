<!--#include file="config.asp"-->
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
If enchiasp.CheckPost = False Then
	Call OutAlertScript("<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ��</li>")
	Response.End
End If

If Cint(enchiasp.membergrade) = 999 Or Trim(Session("AdminName")) <> "" Then
	If enchiasp.ChkNumeric(Request("guestid")) > 0 Then
		If enchiasp.ChkNumeric(Request("replyid")) > 0 Then
			Call DelGuestReply
		Else
			Call DelGuestBook
		End If
	Else
		Call OutAlertScript("�����ϵͳ����!")
	End If
Else
	Call OutAlertScript("��ҳ��Ϊ����ר�ã���û��Ȩ�޵�½��ҳ��")
End If
CloseConn
'================================================
'��������DelGuestBook
'��  �ã�ɾ������
'================================================
Sub DelGuestBook()
	Dim guestid
	If Not IsNumeric(Request("guestid")) Then
		Call OutAlertScript("�����ϵͳ����!")
		Exit Sub
	Else
		guestid = CLng(Request("guestid"))
	End If
	enchiasp.Execute("DELETE FROM ECCMS_GuestBook WHERE guestid="& guestid)
	enchiasp.Execute("DELETE FROM ECCMS_GuestReply WHERE guestid="& guestid)
	Call OutputScript("ɾ�����Գɹ���","index.asp")
End Sub
'================================================
'��������DelGuestReply
'��  �ã�ɾ���ظ�����
'================================================
Sub DelGuestReply()
	Dim replyid,guestid
	If Not IsNumeric(Request("replyid")) Or Not IsNumeric(Request("guestid")) Then
		Call OutAlertScript("�����ϵͳ����!")
		Exit Sub
	Else
		replyid = CLng(Request("replyid"))
		guestid = CLng(Request("guestid"))
	End If
	enchiasp.Execute("DELETE FROM ECCMS_GuestReply WHERE id="& replyid)
	enchiasp.Execute ("UPDATE ECCMS_GuestBook SET ReplyNum=ReplyNum-1 WHERE guestid="& guestid)
	Call OutputScript("ɾ���ظ��ɹ���","index.asp")
End Sub
%>