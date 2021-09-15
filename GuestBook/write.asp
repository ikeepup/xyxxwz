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
Dim Rs, SQL,i
Dim Facestr,FaceOption,FormatInput,strEmotion

enchiasp.LoadTemplates 9999, 3, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$GuestFormContent}", enchiasp.HtmlSetting(10))
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
'--频道目录
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","签写留言")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = HTML.ReadAnnounceList(HtmlContent)
HtmlContent = HTML.ReadFriendLink(HtmlContent)
HtmlContent = html.ReadAnnounceContent(HtmlContent, 0)
HtmlContent = Replace(HtmlContent, "{$MemberName}", enchiasp.membername)
HtmlContent = Replace(HtmlContent,"{$LeastString}", enchiasp.LeastString)
HtmlContent = Replace(HtmlContent, "{$MaxString}", enchiasp.MaxString)
'不允许匿名发表
if enchiasp.HtmlSetting(12)="0" then
	If CInt(enchiasp.AppearGrade) > 0 And Trim(Session("AdminName")) = Empty Then
		If CInt(enchiasp.AppearGrade) > CInt(enchiasp.membergrade) Then
			Call OutputScript(enchiasp.HtmlSetting(1),"index.asp")
			Response.End
		End If
	End If
end if
If enchiasp.CheckStr(LCase(Request.Form("action"))) = "save" And Request.Form("action") <> "" Then
	Call SaveGuestBook
Else
	Call WriteGuestBook
End If

Sub WriteGuestBook()
	HtmlContent = Replace(HtmlContent,"{$Action}","save")
	HtmlContent = Replace(HtmlContent,"{$SubmitValue}","发表新留言")
	HtmlContent = Replace(HtmlContent, "{$GuestID}", "")
	HtmlContent = Replace(HtmlContent,"{$GuestTopic}","")
	HtmlContent = Replace(HtmlContent,"{$UserName}",enchiasp.membername)
	HtmlContent = Replace(HtmlContent,"{$GuestEmail}","")
	HtmlContent = Replace(HtmlContent,"{$GuestQQ}","")
	HtmlContent = Replace(HtmlContent,"{$HomePage}","http://")
	HtmlContent = Replace(HtmlContent,"{$SelectOption}","<option value=""未知"">请选择</option>")
	HtmlContent = Replace(HtmlContent,"<!--发表留言表单 @Begin-->",vbNullString)
	HtmlContent = Replace(HtmlContent,"<!--发表留言表单 @End-->",vbNullString)

	FaceOption = ""
	For i=1 to 20 
		FaceOption = FaceOption & "<option "
		Facestr="images/" & i & ".gif"
		FaceOption = FaceOption & "value='" & Facestr &"'>头像" &i &"</option>"
	Next
	HtmlContent = Replace(HtmlContent, "{$FaceOption}", FaceOption)

	If CInt(enchiasp.membergrade) > 1 Or Trim(Session("AdminName")) <> "" Then
		FormatInput = "<span style=""background-color: #fFfFff"" id=""myt"" onclick=""javascript:formatbt(this);""  style=""cursor:hand; font-size:11pt"">设置标题样式 ABCdef</span>"
		FormatInput = FormatInput & "<input type=""checkbox"" name=""cancel"" value="""" onclick=""Cancelform()""> 取消格式"
		HtmlContent = Replace(HtmlContent,"{$FormatInput}",FormatInput)
	Else
		HtmlContent = Replace(HtmlContent,"{$FormatInput}","")
	End If
	HtmlContent = Replace(HtmlContent,"{$Topicformat}","")

	strEmotion = "<input type=""radio"" value=""emot/1.gif"" name=""emot"" checked><img src=""emot/1.gif"">&nbsp;"
	For i = 2 To 26
		If i = 14 then strEmotion = strEmotion & "<br>"
		strEmotion = strEmotion & "<input type=radio name=emot  value=emot/" & i & ".gif ><img src=""emot/" & i & ".gif"">&nbsp;"
	Next
	HtmlContent = Replace(HtmlContent,"{$EmotionInput}",strEmotion)
	HtmlContent = Replace(HtmlContent,"{$GuestContent}","")
	HtmlContent = Replace(HtmlContent,"{$ForbidChecked}","")
	'不允许匿名发表
	if enchiasp.HtmlSetting(12)="0" then
		HtmlContent = Replace(HtmlContent,"{$IsAdminChecked}","")
	else
		HtmlContent = Replace(HtmlContent,"{$IsAdminChecked}"," disabled")
	end if
	
	If CInt(enchiasp.membergrade) = 999 Or Trim(Session("AdminName")) <> "" Then
		HtmlContent = Replace(HtmlContent,"{$IsTopChecked}","")
	Else
		HtmlContent = Replace(HtmlContent,"{$IsTopChecked}"," disabled")
	End If
	If CInt(enchiasp.IsAuditing) = 0 Or CInt(enchiasp.membergrade) = 999 Or Trim(Session("AdminName")) <> "" Then
		HtmlContent = Replace(HtmlContent,"{$IsAcceptChecked}"," checked")
	Else
		HtmlContent = Replace(HtmlContent,"{$IsAcceptChecked}"," disabled")
	End If
	Response.Write HtmlContent
End Sub

Sub SaveGuestBook()
	On Error Resume Next
	'不允许匿名发表
	if enchiasp.HtmlSetting(12)="0" then
		If CInt(enchiasp.AppearGrade) > 0 And Trim(Session("AdminName")) = Empty Then
			If CInt(enchiasp.AppearGrade) > CInt(enchiasp.membergrade) Then
				ErrMsg = ErrMsg + "<li>对不起！你没有发表留言的权限。</li><li>如果你是本站会员, 请先<a href=""../user/"">登陆</a>!</li>"
				FoundErr = True
			End If
		End If
	end if
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
	If Len(Request.Form("content")) < Clng(enchiasp.LeastString) Then
		ErrMsg = ErrMsg + ("留言内容不能小于" & enchiasp.LeastString & "字符！")
		Founderr = True
	End If
	If Len(Request.Form("content")) > Clng(enchiasp.MaxString) Then
		ErrMsg = ErrMsg + ("留言内容不能大于" & enchiasp.MaxString & "字符！")
		Founderr = True
	End If
	If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		Exit Sub
	End If
	Call PreventRefresh  '防刷新

	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_GuestBook WHERE (guestid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		If enchiasp.membername <> "" And enchiasp.memberid <> "" Then
			Rs("userid") = enchiasp.memberid
			Rs("username") = enchiasp.membername
		Else
			Rs("userid") = 0
			Rs("username") = Left(Request.Form("username"),50)
		End If
		Rs("Topicformat") = Trim(Request.Form("Topicformat"))
		Rs("title") = enchiasp.ChkFormStr(Left(Request.Form("topic"),100))
		Rs("content") = Html2Ubb(Request.Form("content"))
		Rs("face") = Trim(Request.Form("face"))
		Rs("emot") = Trim(Request.Form("emot"))
		Rs("HomePage") = enchiasp.CheckStr(Left(Request.Form("HomePage"),100))
		Rs("GuestEmail") = enchiasp.CheckStr(Trim(Request.Form("GuestEmail")))
		Rs("GuestOicq") = enchiasp.CheckStr(Left(Request.Form("GuestOicq"),30))
		Rs("ComeFrom") = Trim(Request.Form("ComeFrom"))
		Rs("WriteTime") = Now()
		Rs("lastime") = Now()
		Rs("GuestIP") = enchiasp.GetUserIP
		Rs("ReplyNum") = 0
		Rs("isAdmin") = enchiasp.ChkNumeric(Request.Form("isAdmin"))
		Rs("isTop") = enchiasp.ChkNumeric(Request.Form("isTop"))
		If CInt(enchiasp.IsAuditing) = 0 	Or CInt(enchiasp.membergrade) = 999 Or Trim(Session("AdminName")) <> "" Then
			Rs("isAccept") = enchiasp.ChkNumeric(Request.Form("isAccept"))
		Else
			Rs("isAccept") = 0
		End If
		Rs("ForbidReply") = enchiasp.ChkNumeric(Request.Form("ForbidReply"))
	Rs.update
	Rs.Close:Set Rs = Nothing
	enchiasp.DelCahe "RenewStatistics"
	enchiasp.DelCahe "TotalStatistics"
	Dim GroupSetting
	If Trim(enchiasp.membername) <> "" And Trim(enchiasp.membergrade) <> "" Then
		GroupSetting = Split(enchiasp.UserGroupSetting(CInt(enchiasp.membergrade)), "|||")
		enchiasp.Execute ("UPDATE ECCMS_User SET userpoint = userpoint + " & CLng(GroupSetting(26)) & " WHERE userid="& CLng(enchiasp.memberid))
	End If
	If CInt(enchiasp.IsAuditing) = 0 Then
		Response.Redirect("index.asp")
	Else
		Call OutputScript(enchiasp.HtmlSetting(2),"index.asp")
	End If
End Sub
Set HTML = Nothing
CloseConn
%>
