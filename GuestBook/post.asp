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
Dim Rs, SQL,i,replyid,guestid
Dim strContent,strQuote,strTopic
Dim username,isAdmin
Dim Facestr,FaceOption,FormatInput

enchiasp.LoadTemplates 9999, 3, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$GuestFormContent}", enchiasp.HtmlSetting(11))
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
'--Ƶ��Ŀ¼
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","�ظ�����")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = HTML.ReadAnnounceList(HtmlContent)

HtmlContent = Replace(HtmlContent, "{$MemberName}", enchiasp.membername)
HtmlContent = Replace(HtmlContent,"{$LeastString}", enchiasp.LeastString)
HtmlContent = Replace(HtmlContent, "{$MaxString}", enchiasp.MaxString)


If CInt(enchiasp.PostGrade) > 0 And Trim(Session("AdminName")) = Empty Then
	If CInt(enchiasp.PostGrade) > CInt(enchiasp.membergrade) Then
		Call OutputScript(enchiasp.HtmlSetting(5),"index.asp")
		Response.End
	End If
End If

guestid = enchiasp.ChkNumeric(Request("guestid"))
replyid = enchiasp.ChkNumeric(Request("replyid"))
If guestid = 0 Then
	Response.Write"�����ϵͳ����!��������ȷ������ID��"
	Response.End
Else
	Set Rs = enchiasp.Execute("SELECT title,content,username,isAdmin FROM ECCMS_GuestBook WHERE guestid ="& guestid)
	If Rs.BOF And Rs.EOF Then
		Set Rs = Nothing
		Call OutAlertScript("�����ϵͳ����!û���ҵ����������⡣")
	Else
		strTopic = enchiasp.CheckTopic(Rs("title"))
		strContent = Rs("content")
		username = Rs("username")
		isAdmin = Rs("isAdmin")
	End If
	Rs.Close:Set Rs = Nothing
End If
If replyid > 0 Then
	Set Rs = enchiasp.Execute("SELECT rContent FROM ECCMS_GuestReply WHERE id ="& replyid)
	If Rs.BOF And Rs.EOF Then
		Set Rs = Nothing
		Call OutAlertScript("�����ϵͳ����!û���ҵ��ûظ����ԡ�")
	Else
		strContent = Rs("rContent")
	End If
	Rs.Close:Set Rs = Nothing
End If
If CInt(Request("quote")) = 1 Then
	If isAdmin <> 0 Then
		If username = enchiasp.membername Or enchiasp.membergrade = "999" Or Trim(Session("AdminName")) <> "" Then
			strQuote = "<table class=quote><tr><td>" & strContent & "</td><tr></table>"
		Else
			strQuote =  enchiasp.HtmlSetting(16)
		End If
	Else
		strQuote = "<table class=quote><tr><td>" & strContent & "</td><tr></table>"
	End If
Else
	strQuote = ""
End If

If enchiasp.CheckStr(LCase(Request.Form("action"))) = "save" And Request.Form("action") <> "" Then
	Call SaveGuestReply
Else
	Call ReplyGuestBook
End If

Public Sub ReplyGuestBook()

	HtmlContent = Replace(HtmlContent,"{$Action}","save")
	HtmlContent = Replace(HtmlContent,"{$ReplyContent}",vbNullString)
	HtmlContent = Replace(HtmlContent,"{$SubmitValue}","�ظ�����")
	HtmlContent = Replace(HtmlContent, "{$GuestID}", guestid)
	HtmlContent = Replace(HtmlContent, "{$ReplyID}", replyid)
	HtmlContent = Replace(HtmlContent,"{$GuestTopic}",strTopic)
	HtmlContent = Replace(HtmlContent,"{$UserName}",enchiasp.membername)
	HtmlContent = Replace(HtmlContent,"{$GuestEmail}","mymail@163.com")
	HtmlContent = Replace(HtmlContent,"{$GuestQQ}","123456789")
	HtmlContent = Replace(HtmlContent,"{$RefererUrl}",Request.ServerVariables("HTTP_REFERER"))

	FaceOption = ""
	For i=1 to 20 
		FaceOption = FaceOption & "<option "
		Facestr="images/" & i & ".gif"
		FaceOption = FaceOption & "value='" & Facestr &"'>ͷ��" &i &"</option>"
	Next
	HtmlContent = Replace(HtmlContent, "{$FaceOption}", FaceOption)

	If CInt(enchiasp.membergrade) > 1 Or Trim(Session("AdminName")) <> "" Then
		FormatInput = "<span style=""background-color: #fFfFff"" id=""myt"" onclick=""javascript:formatbt(this);""  style=""cursor:hand; font-size:11pt"">���ñ�����ʽ ABCdef</span>"
		FormatInput = FormatInput & "<input type=""checkbox"" name=""cancel"" value="""" onclick=""Cancelform()""> ȡ����ʽ"
		HtmlContent = Replace(HtmlContent,"{$FormatInput}",FormatInput)
	Else
		HtmlContent = Replace(HtmlContent,"{$FormatInput}","")
	End If
	HtmlContent = Replace(HtmlContent,"{$Topicformat}","")
	HtmlContent = Replace(HtmlContent,"{$GuestContent}",Server.HTMLEncode(strQuote))
	Response.Write HtmlContent
End Sub

Sub SaveGuestReply()
	On Error Resume Next
	Dim ForbidReply
	If CInt(enchiasp.PostGrade) > 0 And Trim(Session("AdminName")) = Empty Then
		If CInt(enchiasp.PostGrade) > CInt(enchiasp.membergrade) Then
			ErrMsg = ErrMsg + enchiasp.HtmlSetting(5)
			FoundErr = True
		End If
	End If
	If enchiasp.CheckPost = False Then
		ErrMsg = ErrMsg + "<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ��</li>"
		FoundErr = True
	End If
	If Trim(Request.Form("username")) = "" Then
		ErrMsg = ErrMsg + "�û�������Ϊ��\n"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request.Form("username")) = False Then
		ErrMsg = ErrMsg + "�û����к��зǷ��ַ�\n"
		Founderr = True
	End If
	If Trim(Request.Form("topic")) = "" Then
		ErrMsg = ErrMsg + "�ظ����ⲻ��Ϊ��\n"
		Founderr = True
	End If
	If Trim(Request.Form("content")) = "" Then
		ErrMsg = ErrMsg + "�ظ����ݲ���Ϊ��\n"
		Founderr = True
	End If
	If Len(Request.Form("content")) < Clng(enchiasp.LeastString) Then
		ErrMsg = ErrMsg + ("�ظ����ݲ���С��" & enchiasp.LeastString & "�ַ���")
		Founderr = True
	End If
	If Len(Request.Form("content")) > Clng(enchiasp.MaxString) Then
		ErrMsg = ErrMsg + ("�ظ����ݲ��ܴ���" & enchiasp.MaxString & "�ַ���")
		Founderr = True
	End If
	If Trim(enchiasp.membergrade) <> "999" And Trim(Session("AdminName")) = "" Then
		ForbidReply =enchiasp.Execute("SELECT ForbidReply FROM ECCMS_GuestBook WHERE guestid=" & enchiasp.ChkNumeric(Request.Form("guestid")))(0)
		If ForbidReply <> 0 Then
			ErrMsg = ErrMsg + enchiasp.HtmlSetting(7)
			Founderr = True
		End If
	End If
	If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		Exit Sub
	End If
	Call PreventRefresh  '��ˢ��
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_GuestReply WHERE (id is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		If enchiasp.membername <> "" And enchiasp.memberid <> "" Then
			Rs("userid") = enchiasp.memberid
			Rs("rusername") = enchiasp.membername
		Else
			Rs("userid") = 0
			Rs("rusername") = Left(Request.Form("username"),50)
		End If
		Rs("guestid") = Trim(Request.Form("guestid"))
		Rs("Topicformat") = Trim(Request.Form("Topicformat"))
		Rs("rTitle") = enchiasp.ChkFormStr(Left(Request.Form("topic"),100))
		Rs("rContent") = Trim(Request.Form("content"))
		Rs("rFace") = Trim(Request.Form("face"))
		Rs("ReplyTime") = Now()
		Rs("ReplyIP") = enchiasp.GetUserIP
	Rs.update
	Rs.Close:Set Rs = Nothing
	Dim GroupSetting
	If enchiasp.membername <> "" And enchiasp.membergrade <> "" Then
		GroupSetting = Split(enchiasp.UserGroupSetting(CInt(enchiasp.membergrade)), "|||")
		enchiasp.Execute ("UPDATE ECCMS_User SET userpoint = userpoint + " & CLng(GroupSetting(27)) & " WHERE userid="& CLng(enchiasp.memberid))
	End If
	enchiasp.Execute ("UPDATE ECCMS_GuestBook SET ReplyNum = ReplyNum + 1,lastime = " & NowString & " WHERE guestid="& guestid)
	Call OutputScript(enchiasp.HtmlSetting(8),Request.Form("url"))
End Sub
Set HTML = Nothing
CloseConn
%>
