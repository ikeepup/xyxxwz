<!--#include file="config.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
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
Dim Rs, SQL,i,replyid
Dim Facestr,FaceOption,FormatInput
enchiasp.LoadTemplates 9999, 3, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$GuestFormContent}", enchiasp.HtmlSetting(11))
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
'--频道目录
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","编辑回复")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = HTML.ReadAnnounceList(HtmlContent)

HtmlContent = Replace(HtmlContent, "{$MemberName}", enchiasp.membername)
HtmlContent = Replace(HtmlContent,"{$LeastString}", enchiasp.LeastString)
HtmlContent = Replace(HtmlContent, "{$MaxString}", enchiasp.MaxString)

replyid = enchiasp.ChkNumeric(Request("replyid"))
If replyid = 0 Then
	Response.Write"错误的系统参数!"
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
		Call OutAlertScript("错误的系统参数!")
	End If
	HtmlContent = Replace(HtmlContent,"{$Action}","save")
	HtmlContent = Replace(HtmlContent,"{$ReplyContent}","<b>留言内容：</b><br>" & UBBCode(Rs("rContent")))
	HtmlContent = Replace(HtmlContent,"{$SubmitValue}","保存编辑")
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
		FaceOption = FaceOption & "value='" & Facestr &"'>头像" &i &"</option>"
	Next
	HtmlContent = Replace(HtmlContent, "{$FaceOption}", FaceOption)

	FormatInput = "<span style=""background-color: #fFfFff"" id=""myt"" " & enchiasp.CheckTopic(Rs("Topicformat")) & " onclick=""javascript:formatbt(this);""  style=""cursor:hand; font-size:11pt"">设置标题样式 ABCdef</span>"
	FormatInput = FormatInput & "<input type=""checkbox"" name=""cancel"" value="""" onclick=""Cancelform()""> 取消格式"
	HtmlContent = Replace(HtmlContent,"{$FormatInput}",FormatInput)
	HtmlContent = Replace(HtmlContent,"{$Topicformat}",enchiasp.CheckTopic(Rs("Topicformat")))
	HtmlContent = Replace(HtmlContent,"{$GuestContent}",Server.HTMLEncode(Rs("rContent")))
	Response.Write HtmlContent
	Rs.Close:Set Rs = Nothing
End Sub

Sub SaveGuestReply()
	On Error Resume Next
	If enchiasp.CheckPost = False Then
		ErrMsg = ErrMsg + "<li>您提交的数据不合法，请不要从外部提交。</li>"
		FoundErr = True
	End If
	If Trim(Request.Form("username")) = "" Then
		ErrMsg = ErrMsg + "用户名不能为空\n"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request.Form("username")) = False Then
		ErrMsg = ErrMsg + "用户名中含有非法字符\n"
		Founderr = True
	End If
	If Trim(Request.Form("topic")) = "" Then
		ErrMsg = ErrMsg + "回复主题不能为空\n"
		Founderr = True
	End If
	If Trim(Request.Form("content")) = "" Then
		ErrMsg = ErrMsg + "回复内容不能为空\n"
		Founderr = True
	End If
	If Len(Request.Form("content")) < enchiasp.LeastString Then
		ErrMsg = ErrMsg + ("回复内容不能小于" & enchiasp.LeastString & "字符！")
		Founderr = True
	End If
	If Len(Request.Form("content")) > enchiasp.MaxString Then
		ErrMsg = ErrMsg + ("回复内容不能大于" & enchiasp.MaxString & "字符！")
		Founderr = True
	End If
	If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		Exit Sub
	End If
	Call PreventRefresh  '防刷新
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
