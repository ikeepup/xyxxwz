<!--#include file="config.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
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
Dim Rs, SQL,i,replyid
Dim Facestr,FaceOption,FormatInput
enchiasp.LoadTemplates 9999, 3, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$GuestFormContent}", enchiasp.HtmlSetting(11))
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
'--Ƶ��Ŀ¼
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","�༭�ظ�")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = HTML.ReadAnnounceList(HtmlContent)

HtmlContent = Replace(HtmlContent, "{$MemberName}", enchiasp.membername)
HtmlContent = Replace(HtmlContent,"{$LeastString}", enchiasp.LeastString)
HtmlContent = Replace(HtmlContent, "{$MaxString}", enchiasp.MaxString)

replyid = enchiasp.ChkNumeric(Request("replyid"))
If replyid = 0 Then
	Response.Write"�����ϵͳ����!"
	Response.End
End If
If Trim(enchiasp.membergrade) = "999" Or Trim(Session("AdminName")) <> "" Then
	If enchiasp.CheckStr(LCase(Request.Form("action"))) = "save" And Request.Form("action") <> "" Then
		Call SaveGuestReply
	Else
		Call EditGuestReply
	End If
Else
	Call OutAlertScript(enchiasp.HtmlSetting(3))
End If

Sub EditGuestReply()
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_GuestReply WHERE id ="& replyid)
	If Rs.BOF And Rs.EOF Then
		Set Rs = Nothing
		Call OutAlertScript("�����ϵͳ����!")
	End If
	HtmlContent = Replace(HtmlContent,"{$Action}","save")
	HtmlContent = Replace(HtmlContent,"{$ReplyContent}","<b>�������ݣ�</b><br>" & UBBCode(Rs("rContent")))
	HtmlContent = Replace(HtmlContent,"{$SubmitValue}","����༭")
	HtmlContent = Replace(HtmlContent, "{$GuestID}", Rs("guestid"))
	HtmlContent = Replace(HtmlContent, "{$ReplyID}", replyid)
	HtmlContent = Replace(HtmlContent,"{$GuestTopic}",enchiasp.CheckTopic(Rs("rtitle")))
	HtmlContent = Replace(HtmlContent,"{$UserName}",enchiasp.CheckTopic(Rs("rUserName")))
	HtmlContent = Replace(HtmlContent,"{$GuestEmail}","mymail@163.com")
	HtmlContent = Replace(HtmlContent,"{$GuestQQ}","123456789")
	HtmlContent = Replace(HtmlContent,"{$RefererUrl}",Request.ServerVariables("HTTP_REFERER"))

	FaceOption = ""
	For i=1 to 20 
		FaceOption = FaceOption & "<option "
		Facestr="images/" & i & ".gif"
		If LCase(Facestr) = LCase(Rs("rface")) Then FaceOption = FaceOption & "selected "
		FaceOption = FaceOption & "value='" & Facestr &"'>ͷ��" &i &"</option>"
	Next
	HtmlContent = Replace(HtmlContent, "{$FaceOption}", FaceOption)

	FormatInput = "<span style=""background-color: #fFfFff"" id=""myt"" " & enchiasp.CheckTopic(Rs("Topicformat")) & " onclick=""javascript:formatbt(this);""  style=""cursor:hand; font-size:11pt"">���ñ�����ʽ ABCdef</span>"
	FormatInput = FormatInput & "<input type=""checkbox"" name=""cancel"" value="""" onclick=""Cancelform()""> ȡ����ʽ"
	HtmlContent = Replace(HtmlContent,"{$FormatInput}",FormatInput)
	HtmlContent = Replace(HtmlContent,"{$Topicformat}",enchiasp.CheckTopic(Rs("Topicformat")))
	HtmlContent = Replace(HtmlContent,"{$GuestContent}",Server.HTMLEncode(Rs("rContent")))
	Response.Write HtmlContent
	Rs.Close:Set Rs = Nothing
End Sub

Sub SaveGuestReply()
	On Error Resume Next
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
	If Len(Request.Form("content")) < enchiasp.LeastString Then
		ErrMsg = ErrMsg + ("�ظ����ݲ���С��" & enchiasp.LeastString & "�ַ���")
		Founderr = True
	End If
	If Len(Request.Form("content")) > enchiasp.MaxString Then
		ErrMsg = ErrMsg + ("�ظ����ݲ��ܴ���" & enchiasp.MaxString & "�ַ���")
		Founderr = True
	End If
	If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		Exit Sub
	End If
	Call PreventRefresh  '��ˢ��
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_GuestReply WHERE id="& replyid
	Rs.Open SQL,Conn,1,3
		Rs("rusername") = Left(Request.Form("username"),50)
		Rs("Topicformat") = Trim(Request.Form("Topicformat"))
		Rs("rTitle") = enchiasp.ChkFormStr(Left(Request.Form("topic"),100))
		Rs("rContent") = Html2Ubb(Request.Form("content"))
		Rs("rFace") = Trim(Request.Form("face"))

	Rs.update
	Rs.Close:Set Rs = Nothing
	Call OutputScript(enchiasp.HtmlSetting(9),Request.ServerVariables("HTTP_REFERER"))
End Sub
Set HTML = Nothing
CloseConn
%>
