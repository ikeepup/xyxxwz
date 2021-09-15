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
Dim TempListContent,ListContent,strClassName
Dim Rs, SQL,i,guestid,j,isAdmin,username,strTopic
Dim maxperpage, totalnumber, TotalPageNum, CurrentPage
Dim strReplyAlt,strTempContent,strPagination
Dim strHomePage,strLockedAlt

enchiasp.LoadTemplates 9999, 2, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
'--频道目录
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
'HtmlContent = Replace(HtmlContent,"{$PageTitle}","回复留言")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = HTML.ReadAnnounceList(HtmlContent)
HtmlContent = HTML.ReadFriendLink(HtmlContent)
HtmlContent = Replace(HtmlContent,"{$LeastString}", enchiasp.LeastString)
HtmlContent = Replace(HtmlContent, "{$MaxString}", enchiasp.MaxString)
HtmlContent = Replace(HtmlContent, "{$MemberName}", enchiasp.membername)
HtmlContent = html.ReadAnnounceContent(HtmlContent, 0)
strClassName = "回复留言"

maxperpage = enchiasp.ChkNumeric(enchiasp.HtmlSetting(1))   '每页显示帖子数

guestid = enchiasp.ChkNumeric(Request("guestid"))
If guestid = 0 Then
	Response.Write"错误的系统参数!"
	Response.End
End If

CurrentPage = enchiasp.ChkNumeric(Request("page"))
If CInt(CurrentPage) = 0 Then CurrentPage = 1
if Session("AdminName")="" then
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_GuestBook WHERE isAccept > 0 And guestid ="& guestid)
else
	
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_GuestBook WHERE guestid ="& guestid)

end if
If Rs.BOF And Rs.EOF Then
	Set Rs = Nothing
	Call OutAlertScript("错误的系统参数!")
Else
	strTopic = enchiasp.HTMLEncode(enchiasp.CheckTopic(Rs("title")))
	HtmlContent = Replace(HtmlContent,"{$GuestID}", guestid)
	HtmlContent = Replace(HtmlContent,"{$GuestTopic}", strTopic)
	If Rs("ReplyNum") <> 0 Then
		strReplyAlt = "<img src='images/collapsed_yes.gif' border=0 alt='有回复（共 " & Rs("ReplyNum") & " 条）'>"
	Else
		strReplyAlt = "<img src='images/collapsed_no.gif' border=0 alt='无回复'>"
	End If
	HtmlContent = Replace(HtmlContent,"{$ReplyAlt}", strReplyAlt)
	
	If CurrentPage = 1 Then
		strTempContent = enchiasp.HtmlSetting(3)
		strTempContent = Replace(strTempContent,"{$GuestID}", guestid)
		strTempContent = Replace(strTempContent,"{$GuestTopic}", strTopic)
		strTempContent = Replace(strTempContent,"{$UserName}", Rs("username"))
		strTempContent = Replace(strTempContent,"{$GuestStatus}", GuestStation(Rs("userid")))
		strTempContent = Replace(strTempContent,"{$UserFace}", enchiasp.CheckTopic(Rs("face")))
		strTempContent = Replace(strTempContent,"{$ComeFrom}", Rs("ComeFrom"))
		strTempContent = Replace(strTempContent,"{$GuestQQ}", Rs("GuestOicq"))
		strTempContent = Replace(strTempContent,"{$Emotion}", Rs("emot"))
		strTempContent = Replace(strTempContent,"{$GuestEmail}", enchiasp.CheckTopic(Rs("GuestEmail")))
		strTempContent = Replace(strTempContent,"{$WriteTime}", Rs("WriteTime"))
		strTempContent = Replace(strTempContent,"{$GuestIP}", Rs("GuestIP"))
		
		If IsNull(Rs("Topicformat")) Then
			strTempContent = Replace(strTempContent, "{$Topicformat}", "")
		Else
			strTempContent = Replace(strTempContent, "{$Topicformat}", " "& Rs("Topicformat"))
		End If
		If Rs("ForbidReply") <> 0 Then
			strLockedAlt = "      <img src='images/a_lock.gif' align=""absmiddle"" alt=""此留言已锁定，禁止回复"">"
		Else
			strLockedAlt = ""
		End If
		strTempContent = Replace(strTempContent,"{$LockedAlt}", strLockedAlt)
		
		If enchiasp.CheckNull(Rs("HomePage")) Then
			strHomePage = Rs("HomePage")
			If LCase(Left(strHomePage,4)) <> "http" Then strHomePage = "http://" & strHomePage
		Else
			strHomePage = "#"
		End If
		strTempContent = Replace(strTempContent,"{$HomePage}", strHomePage)

		If Rs("isAdmin") <> 0 Then
			If Rs("username") = enchiasp.membername Or enchiasp.membergrade = "999" Or Trim(Session("AdminName")) <> "" Then
				strTempContent = Replace(strTempContent, "{$GuestContent}", enchiasp.ChkBadWords(UBBCode(Rs("Content"))))
			Else
				strTempContent = Replace(strTempContent, "{$GuestContent}",enchiasp.HtmlSetting(2))
			End If
		Else
			strTempContent = Replace(strTempContent, "{$GuestContent}", enchiasp.ChkBadWords(UBBCode(Rs("Content"))))
		End If

	End If
	HtmlContent = Replace(HtmlContent,"{$TopContent}", strTempContent)
	HtmlContent = Replace(HtmlContent,"{$PageTitle}",strTopic)
	isAdmin = Rs("isAdmin")
	username = Rs("username")
	Rs.Close:Set Rs = Nothing

	'记录总数
	totalnumber = enchiasp.Execute("SELECT COUNT(id) FROM ECCMS_GuestReply WHERE guestid="& guestid)(0)
	TotalPageNum = CLng(totalnumber / maxperpage)  '得到总页数
	If TotalPageNum < totalnumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_GuestReply WHERE guestid="& guestid &" ORDER BY id ASC"
	Rs.Open SQL, Conn, 1, 1
	If Rs.BOF And Rs.EOF Then
		HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
	Else
		i = 0
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		j = (CurrentPage - 1) * maxperpage + 1
		TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 0)
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			ListContent = ListContent & TempListContent
			ListContent = Replace(ListContent,"{$ReplyID}", Rs("id"))
			ListContent = Replace(ListContent,"{$Several}", j)
			ListContent = Replace(ListContent,"{$ReplyTopic}", enchiasp.HTMLEncode(enchiasp.CheckTopic(Rs("rtitle"))))
			ListContent = Replace(ListContent,"{$ReplyUserName}", Rs("rUserName"))
			ListContent = Replace(ListContent,"{$UserStatus}", GuestStation(Rs("userid")))
			ListContent = Replace(ListContent,"{$ReplyFace}", enchiasp.CheckTopic(Rs("rface")))
			ListContent = Replace(ListContent,"{$ReplyTime}", Rs("ReplyTime"))
			ListContent = Replace(ListContent,"{$ReplyIP}", Rs("ReplyIP"))

			If IsNull(Rs("Topicformat")) Then
				ListContent = Replace(ListContent, "{$ReplyTopicformat}", "")
			Else
				ListContent = Replace(ListContent, "{$ReplyTopicformat}", " "& Rs("Topicformat"))
			End If
			If isAdmin <> 0 Then
				If username = enchiasp.membername Or enchiasp.membergrade = "999" Or Trim(Session("AdminName")) <> "" Then
					ListContent = Replace(ListContent, "{$ReplyContent}", enchiasp.ChkBadWords(UBBCode(Rs("rContent"))))
				Else
					ListContent = Replace(ListContent, "{$ReplyContent}",enchiasp.HtmlSetting(2))
				End If
			Else
				ListContent = Replace(ListContent, "{$ReplyContent}", enchiasp.ChkBadWords(UBBCode(Rs("rContent"))))
			End If

			Rs.movenext
			i = i + 1
			j = j + 1
			If i >= maxperpage Then Exit Do
		Loop
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
	End If
	Rs.Close:Set Rs = Nothing
End If

strPagination = ShowListPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, "", strClassName)
HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strPagination)

Dim Facestr,FaceOption
FaceOption = ""
For i=1 to 20 
	FaceOption = FaceOption & "<option "
	Facestr="images/" & i & ".gif"
	FaceOption = FaceOption & "value='" & Facestr &"'>头像" &i &"</option>"
Next
HtmlContent = Replace(HtmlContent, "{$FaceOption}", FaceOption)
Response.Write HtmlContent

If enchiasp.CheckStr(LCase(Request.Form("action"))) = "save" And Request.Form("action") <> "" Then
	Call SaveGuestReply
End If

Sub SaveGuestReply()
	On Error Resume Next
	Call PreventRefresh
	Dim ForbidReply
	If CInt(enchiasp.PostGrade) > 0 And Trim(Session("AdminName")) = Empty Then
		If CInt(enchiasp.PostGrade) > CInt(enchiasp.membergrade) Then
			ErrMsg = ErrMsg + "对不起！你没有回复留言的权限。\n\n如果你是本站会员, 请先登陆"
			FoundErr = True
		End If
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
	If Trim(Request.Form("content1")) = "" Then
		ErrMsg = ErrMsg + "回复内容不能为空\n"
		Founderr = True
	End If
	If Len(Request.Form("content1")) < Clng(enchiasp.LeastString) Then
		ErrMsg = ErrMsg + ("回复内容不能小于" & enchiasp.LeastString & "字符！")
		Founderr = True
	End If
	If Len(Request.Form("content")) > Clng(enchiasp.MaxString) Then
		ErrMsg = ErrMsg + ("回复内容不能大于" & enchiasp.MaxString & "字符！")
		Founderr = True
	End If
	If enchiasp.membergrade <> "999" And Trim(Session("AdminName")) = "" Then
		ForbidReply =enchiasp.Execute("SELECT ForbidReply from ECCMS_GuestBook WHERE guestid=" & guestid)(0)
		If ForbidReply <> 0 Then
			ErrMsg = ErrMsg + enchiasp.HtmlSetting(4)
			Founderr = True
		End If
	End If
	If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		Exit Sub
	End If
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
		Rs("Topicformat") = ""
		Rs("rTitle") = enchiasp.ChkFormStr(Left(Request.Form("topic"),100))
		Rs("rContent") = Html2Ubb(Request.Form("content1"))
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
	Call OutputScript(enchiasp.HtmlSetting(5),Request.ServerVariables("HTTP_REFERER"))
End Sub
Set HTML = Nothing
CloseConn
%>