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
Dim SendMail

Sub Jmail(email,topic,mailbody)
	On Error Resume Next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.Message")
	'JMail.silent=true
	JMail.Logging=True
	JMail.Charset="gb2312"
	JMail.MailServerUserName = enchiasp.MailUserName '�����ʼ���������¼��
	JMail.MailServerPassword = enchiasp.MailPassword '��¼����
	JMail.ContentType = "text/html"
	JMail.Priority = 1
	JMail.From = enchiasp.MailFrom  '�ʼ���ַ
	JMail.FromName = enchiasp.SiteName  '��վ����
	JMail.AddRecipient email
	JMail.Subject = topic
	JMail.Body = mailbody
	JMail.Send (enchiasp.MailServer)   '���ʼ���������ַ
	Set JMail=nothing
	SendMail="OK"
	If Err Then SendMail="False"
End Sub
	
Sub Cdonts(email,topic,mailbody)
	On Error Resume Next
	Dim objCDOMail
	Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
	objCDOMail.From = enchiasp.MailFrom  '�ʼ���ַ
	objCDOMail.To = email
	objCDOMail.Subject = topic
	objCDOMail.BodyFormat = 0 
	objCDOMail.MailFormat = 0 
	objCDOMail.Body = mailbody
	objCDOMail.Send
	Set objCDOMail = Nothing
	SendMail="OK"
	If Err Then SendMail="False"
End Sub

Sub aspemail(email,topic,mailbody)
	On Error Resume Next
	Dim Mailer
	Set Mailer=Server.CreateObject("Persits.MailSender") 
	Mailer.Charset = "gb2312"
	Mailer.IsHTML = True
	Mailer.username = enchiasp.MailUserName	'����������Ч���û���
	Mailer.password = enchiasp.MailPassword	'����������Ч������
	Mailer.Priority = 1
	Mailer.Host = enchiasp.setting(9)
	Mailer.Port = 25 ' �����ѡ.�˿�25��Ĭ��ֵ
	Mailer.From = enchiasp.MailFrom   '�ʼ���ַ
	Mailer.FromName = enchiasp.SiteName ' �����ѡ
	Mailer.AddAddress email,email
	Mailer.Subject = topic
	Mailer.Body = mailbody
	Mailer.Send
	SendMail="OK"
	If Err Then SendMail="False"
End Sub
%>
