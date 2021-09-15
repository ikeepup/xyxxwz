<!--#include file="config.asp"-->
<!--#include file="../inc/chkinput.asp"-->
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
Dim Rs, SQL,i,guestid
Dim Facestr,FaceOption,FormatInput,strEmotion

enchiasp.LoadTemplates 9999, 3, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$GuestFormContent}", enchiasp.HtmlSetting(10))
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
'--频道目录
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","编辑留言")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = HTML.ReadAnnounceList(HtmlContent)

HtmlContent = Replace(HtmlContent, "{$MemberName}", enchiasp.membername)
HtmlContent = Replace(HtmlContent,"{$LeastString}", enchiasp.LeastString)
HtmlContent = Replace(HtmlContent, "{$MaxString}", enchiasp.MaxString)

guestid = enchiasp.ChkNumeric(Request("guestid"))
If guestid = 0 Then
	Response.Write"错误的系统参数!"
	Response.End
End If

HtmlContent = Replace(HtmlContent, "{$GuestID}", guestid)
HtmlContent = Replace(HtmlContent,"{$Action}","save")
HtmlContent = Replace(HtmlContent,"{$SubmitValue}","编辑留言")

If CInt(enchiasp.membergrade) = 999 Or Trim(Session("AdminName")) <> "" Then
	If enchiasp.CheckStr(LCase(Request.Form("action"))) = "save" And Request.Form("action") <> "" Then
		Call SaveGuestBook
	Else
		Call EditGuestBook
	End If
Else
	Call OutAlertScript(enchiasp.HtmlSetting(3))
End If

Sub EditGuestBook()
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_GuestBook WHERE guestid ="& guestid)
	If Rs.BOF And Rs.EOF Then
		Set Rs = Nothing
		Call OutAlertScript("错误的系统参数!")
	End If
	HtmlContent = Replace(HtmlContent,"{$GuestTopic}",enchiasp.CheckTopic(Rs("title")))
	HtmlContent = Replace(HtmlContent,"{$UserName}",enchiasp.CheckTopic(Rs("username")))
	HtmlContent = Replace(HtmlContent,"{$GuestEmail}",enchiasp.CheckTopic(Rs("GuestEmail")))
	HtmlContent = Replace(HtmlContent,"{$GuestQQ}",enchiasp.CheckTopic(Rs("GuestOicq")))
	HtmlContent = Replace(HtmlContent,"{$HomePage}",enchiasp.CheckTopic(Rs("HomePage")))
	HtmlContent = Replace(HtmlContent,"{$SelectOption}","<option value=""" & Rs("ComeFrom") & """>" & Rs("ComeFrom") & "</option>")
	
	FaceOption = ""
	For i=1 to 20 
		FaceOption = FaceOption & "<option "
		Facestr="images/" & i & ".gif"
		If LCase(Facestr) = LCase(Rs("face")) Then FaceOption = FaceOption & "selected "
		FaceOption = FaceOption & "value='" & Facestr &"'>头像" &i &"</option>"
	Next
	HtmlContent = Replace(HtmlContent, "{$FaceOption}", FaceOption)

	FormatInput = "<span style=""background-color: #fFfFff"" id=""myt"" " & enchiasp.CheckTopic(Rs("Topicformat")) & " onclick=""javascript:formatbt(this);""  style=""cursor:hand; font-size:11pt"">设置标题样式 ABCdef</span>"
	FormatInput = FormatInput & "<input type=""checkbox"" name=""cancel"" value="""" onclick=""Cancelform()""> 取消格式"
	HtmlContent = Replace(HtmlContent,"{$FormatInput}",FormatInput)
	HtmlContent = Replace(HtmlContent,"{$Topicformat}",enchiasp.CheckTopic(Rs("Topicformat")))

	strEmotion = "<input type=""radio"" value=""emot/1.gif"" name=""emot"" checked><img src=""emot/1.gif"">&nbsp;"
	For i = 2 To 26
		If i = 14 then strEmotion = strEmotion & "<br>"
		strEmotion = strEmotion & "<input type=radio name=emot  value=emot/" & i & ".gif ><img src=""emot/" & i & ".gif"">&nbsp;"
	Next
	HtmlContent = Replace(HtmlContent,"{$EmotionInput}",strEmotion)
	HtmlContent = Replace(HtmlContent,"{$GuestContent}",Server.HTMLEncode(Rs("content")))
	If Rs("isAdmin") <> 0 Then
		HtmlContent = Replace(HtmlContent,"{$IsAdminChecked}"," checked")
	Else
		HtmlContent = Replace(HtmlContent,"{$IsAdminChecked}","")
	End If
	If Rs("ForbidReply") <> 0 Then
		HtmlContent = Replace(HtmlContent,"{$ForbidChecked}"," checked")
	Else
		HtmlContent = Replace(HtmlContent,"{$ForbidChecked}","")
	End If
	If Rs("isTop") <> 0 Then
		HtmlContent = Replace(HtmlContent,"{$IsTopChecked}"," checked")
	Else
		HtmlContent = Replace(HtmlContent,"{$IsTopChecked}","")
	End If
	If CInt(Rs("isAccept")) = 0 Then
		HtmlContent = Replace(HtmlContent,"{$IsAcceptChecked}","")
	Else
		HtmlContent = Replace(HtmlContent,"{$IsAcceptChecked}"," checked")
	End If
	Response.Write HtmlContent
	Rs.Close:Set Rs = Nothing
End Sub

Sub SaveGuestBook()
	
	On Error Resume Next
	If Trim(Request.Form("username")) = "" Then
		ErrMsg = ErrMsg + "用户名不能为空\n"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request.Form("username")) = False Then
		ErrMsg = ErrMsg + "用户名中含有非法字符\n"
		Founderr = True
	End If
	If Trim(Request.Form("GuestEmail")) = "" Then
		ErrMsg = ErrMsg + "用户邮箱不能为空\n"
		Founderr = True
	End If
	If Not IsValidEmail(Request.Form("GuestEmail")) Then
		ErrMsg = ErrMsg + "请正确填写您的邮箱\n"
		Founderr = True
	End If
	If Trim(Request.Form("topic")) = "" Then
		ErrMsg = ErrMsg + "留言主题不能为空\n"
		Founderr = True
	End If
	If Trim(Request.Form("content")) = "" Then
		ErrMsg = ErrMsg + "留言内容不能为空\n"
		Founderr = True
	End If
	If Len(Request.Form("content")) < enchiasp.LeastString Then
		ErrMsg = ErrMsg + ("留言内容不能小于" & enchiasp.LeastString & "字符！")
		Founderr = True
	End If
	If Len(Request.Form("content")) > enchiasp.MaxString Then
		ErrMsg = ErrMsg + ("留言内容不能大于" & enchiasp.MaxString & "字符！")
		Founderr = True
	End If
	If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		Exit Sub
	End If
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_GuestBook WHERE guestid="& guestid
	Rs.Open SQL,Conn,1,3
		Rs("Topicformat") = enchiasp.CheckStr(Request.Form("Topicformat"))
		Rs("title") = enchiasp.ChkFormStr(Left(Request.Form("topic"),100))
		Rs("content") = Trim(Request.Form("content"))
		Rs("username") = Left(Request.Form("username"),50)
		Rs("face") = Trim(Request.Form("face"))
		Rs("emot") = Trim(Request.Form("emot"))
		Rs("HomePage") = enchiasp.CheckStr(Left(Request.Form("HomePage"),100))
		Rs("GuestEmail") = enchiasp.CheckStr(Trim(Request.Form("GuestEmail")))
		Rs("GuestOicq") = enchiasp.CheckStr(Left(Request.Form("GuestOicq"),30))
		Rs("ComeFrom") = Trim(Request.Form("ComeFrom"))
		Rs("isAdmin") = enchiasp.ChkNumeric(Request.Form("isAdmin"))
		Rs("isTop") = enchiasp.ChkNumeric(Request.Form("isTop"))
		Rs("isAccept") = enchiasp.ChkNumeric(Request.Form("isAccept"))
		Rs("ForbidReply") = enchiasp.ChkNumeric(Request.Form("ForbidReply"))
	Rs.update
	Rs.Close:Set Rs = Nothing
	Call OutputScript(enchiasp.HtmlSetting(4),"index.asp")
End Sub
Set HTML = Nothing
CloseConn
%>