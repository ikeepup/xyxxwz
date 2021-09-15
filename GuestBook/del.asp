<!--#include file="config.asp"-->
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
If enchiasp.CheckPost = False Then
	Call OutAlertScript("<li>您提交的数据不合法，请不要从外部提交。</li>")
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
		Call OutAlertScript("错误的系统参数!")
	End If
Else
	Call OutAlertScript("本页面为管理专用，您没有权限登陆本页！")
End If
CloseConn
'================================================
'过程名：DelGuestBook
'作  用：删除留言
'================================================
Sub DelGuestBook()
	Dim guestid
	If Not IsNumeric(Request("guestid")) Then
		Call OutAlertScript("错误的系统参数!")
		Exit Sub
	Else
		guestid = CLng(Request("guestid"))
	End If
	enchiasp.Execute("DELETE FROM ECCMS_GuestBook WHERE guestid="& guestid)
	enchiasp.Execute("DELETE FROM ECCMS_GuestReply WHERE guestid="& guestid)
	Call OutputScript("删除留言成功！","index.asp")
End Sub
'================================================
'过程名：DelGuestReply
'作  用：删除回复留言
'================================================
Sub DelGuestReply()
	Dim replyid,guestid
	If Not IsNumeric(Request("replyid")) Or Not IsNumeric(Request("guestid")) Then
		Call OutAlertScript("错误的系统参数!")
		Exit Sub
	Else
		replyid = CLng(Request("replyid"))
		guestid = CLng(Request("guestid"))
	End If
	enchiasp.Execute("DELETE FROM ECCMS_GuestReply WHERE id="& replyid)
	enchiasp.Execute ("UPDATE ECCMS_GuestBook SET ReplyNum=ReplyNum-1 WHERE guestid="& guestid)
	Call OutputScript("删除回复成功！","index.asp")
End Sub
%>