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
Dim SendMail

Sub Jmail(email,topic,mailbody)
	On Error Resume Next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.Message")
	'JMail.silent=true
	JMail.Logging=True
	JMail.Charset="gb2312"
	JMail.MailServerUserName = enchiasp.MailUserName '您的邮件服务器登录名
	JMail.MailServerPassword = enchiasp.MailPassword '登录密码
	JMail.ContentType = "text/html"
	JMail.Priority = 1
	JMail.From = enchiasp.MailFrom  '邮件地址
	JMail.FromName = enchiasp.SiteName  '网站名称
	JMail.AddRecipient email
	JMail.Subject = topic
	JMail.Body = mailbody
	JMail.Send (enchiasp.MailServer)   '发邮件服务器地址
	Set JMail=nothing
	SendMail="OK"
	If Err Then SendMail="False"
End Sub
	
Sub Cdonts(email,topic,mailbody)
	On Error Resume Next
	Dim objCDOMail
	Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
	objCDOMail.From = enchiasp.MailFrom  '邮件地址
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
	Mailer.username = enchiasp.MailUserName	'服务器上有效的用户名
	Mailer.password = enchiasp.MailPassword	'服务器上有效的密码
	Mailer.Priority = 1
	Mailer.Host = enchiasp.setting(9)
	Mailer.Port = 25 ' 该项可选.端口25是默认值
	Mailer.From = enchiasp.MailFrom   '邮件地址
	Mailer.FromName = enchiasp.SiteName ' 该项可选
	Mailer.AddAddress email,email
	Mailer.Subject = topic
	Mailer.Body = mailbody
	Mailer.Send
	SendMail="OK"
	If Err Then SendMail="False"
End Sub
%>
