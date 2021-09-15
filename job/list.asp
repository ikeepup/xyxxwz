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
dim ClassID
dim ArticleID,ArticleContent
Dim TempListContent,ListContent
Dim Rs, SQL, foundsql, j
dim temptj1,temptj2
Dim maxperpage, totalnumber, TotalPageNum, CurrentPage, i
dim strPagination
Dim strClassName

if Request("classid")="" then
	Call OutputScript("错误的参数，请不要随意输入一些参数！","index.asp")
end if
If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
	CurrentPage = enchiasp.ChkNumeric(Request("page"))
Else
	CurrentPage = 1
End If
ClassID = enchiasp.ChkNumeric(Request("ClassID"))
if checkdanyemian(ClassID) then
	call showdanyemian(ClassID, 1)
else
		response.Redirect "index.asp"
end if



	
	'================================================
	'函数名：showdanyemian
	'作  用：列出单页面内容
	'================================================
	private sub showdanyemian(clsid, n)
		On Error Resume Next
		If Not IsNumeric(clsid) Then
			Exit sub
		Else
			ArticleID = CLng(clsid)
		End If
		
		
		SQL = "SELECT top 1 A.ArticleID,A.ClassID,A.title,A.content,A.Related,A.Author,A.ComeFrom,A.isTop,A.username,A.star,A.isBest,A.WriteTime,A.Allhits,A.HtmlFileDate,A.UserGroup,A.PointNum,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,C.UserGroup As User_Group,C.UseHtml FROM [ECCMS_Article] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.ClassID=" & ArticleID
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			ReadArticleContent = ""
			Set Rs = Nothing
			If CreateHtml = 0 Then
				Response.Write "<meta http-equiv=""refresh"" content=""2;url='/"">" & vbNewLine
				Response.Write "<p align=""center"" style=""font-size: 16px;color: red;"">对不起，该页面没有内容或发生了错误，无法访问! 系统两秒后自动转到网站首页......</p>" & vbNewLine
			End If
			Exit sub
		End If
		CheckUserRead Rs("ArticleID"), Rs("PointNum"), Rs("UserGroup"), Rs("User_Group")
		Call ContentPagination

		enchiasp.LoadTemplates ChannelID, 5, 0
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$dingbu}",enchiasp.HtmlSetting(4))
		HtmlContent = Replace(HtmlContent, "{$dibu}",enchiasp.HtmlSetting(5))
		HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ArticleContent}", ArticleContent)		
		HtmlContent = HTML.ReadArticlePic(HtmlContent)
		HtmlContent = HTML.ReadSoftPic(HtmlContent)
		HtmlContent = HTML.ReadArticleList(HtmlContent)
		HtmlContent = HTML.ReadSoftList(HtmlContent)
		HtmlContent = HTML.ReadShopList(HtmlContent)
		HtmlContent = HTML.ReadShopPic(HtmlContent)
		HtmlContent = HTML.ReadFlashList(HtmlContent)
		HtmlContent = HTML.ReadFlashPic(HtmlContent)
		HtmlContent = HTML.ReadFriendLink(HtmlContent)
		HtmlContent = HTML.ReadNewsPicAndText(HtmlContent)
		HtmlContent = HTML.ReadSoftPicAndText(HtmlContent)
		HtmlContent = HTML.ReadGuestList(HtmlContent)
		HtmlContent = HTML.ReadAnnounceList(HtmlContent)
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
		HtmlContent = HTML.ReadPopularSoft(HtmlContent)
		HtmlContent = HTML.ReadPopularFlash(HtmlContent)
		HtmlContent = HTML.ReadSoftType(HtmlContent)
		HtmlContent = HTML.ReadStatistic(HtmlContent)
		HtmlContent = HTML.ReadUserRank(HtmlContent)

		'--频道目录
		HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
		HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
		HtmlContent = Replace(HtmlContent,"{$PageTitle}",rs("classname"))
		
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		
		
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$ArticleContent}",  Rs("content"))

		HtmlContent = HTML.ReadAnnounceList(HtmlContent)		
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		Response.Write HtmlContent
		Set HTML = Nothing
		CloseConn
	end sub
	
	private function checkdanyemian(classid)
		'数据库重新关闭，打开
		SQL = "SELECT * from [ECCMS_Classify] where isdanyemian=1 and ClassID="& ClassID &""
		Set Rs = enchiasp.Execute(SQL)

		If Rs.BOF And Rs.EOF Then
				checkdanyemian= false
		else
				checkdanyemian = true
		End If

		Rs.Close: Set Rs = Nothing
	end function
		
		Private Function CheckUserRead(ByVal ArticleID, ByVal PointNum, ByVal UserGroup, ByVal User_Group)
		Dim Message, CookiesID
		Dim GroupSetting, GroupName, gradeid
		dim strInstallDir
		strInstallDir = enchiasp.InstallDir
		If CInt(enchiasp.membergrade) = 999 Then Exit Function
		If CInt(enchiasp.membergrade) <> 0 Then
			gradeid = CInt(enchiasp.membergrade)
		Else
			gradeid = 0
		End If
		GroupSetting = Split(enchiasp.UserGroupSetting(gradeid), "|||")
		GroupName = GroupSetting(UBound(GroupSetting))
		If CInt(User_Group) > CInt(gradeid) Or CInt(UserGroup) > CInt(gradeid) Then
			Message = "<li>您没有登录或者你的会员级别不够，不能阅览此文章！</li><li>如果你是本站会员, 请先<a href='"&  strInstallDir &"/user/" &"'>登陆</a></li>"
			Response.Redirect (strInstallDir & "showerr.asp?action=error&Message=" & Server.URLEncode(Message))
			Response.end
		End If

		On Error Resume Next
		Dim rsMember
		If CInt(enchiasp.memberclass) > 0 Then
			Set rsMember = CreateObject("ADODB.Recordset")
			SQL = "SELECT userid,UserGrade,UserClass,ExpireTime FROM ECCMS_User WHERE UserClass>0 And username='" & enchiasp.membername & "' And userid=" & CLng(enchiasp.memberid)
			rsMember.Open SQL, Conn, 1, 3
			If rsMember.BOF And rsMember.EOF Then
				Message = "<li>非法操作~！</li>"
				Response.Redirect (strInstallDir & "showerr.asp?action=error&Message=" & Server.URLEncode(Message))
				Set rsMember = Nothing
				Response.end
			Else
				If DateDiff("D", CDate(rsMember("ExpireTime")), Now()) > 0 Or CInt(rsMember("UserClass")) = 999 Then
					Message = "<li>对不起！您的会员已到期，不能阅览此文章；</li><li>如果你要阅览此文章请联系管理员。</li>"
					Response.Redirect (strInstallDir & "showerr.asp?action=error&Message=" & Server.URLEncode(Message))
					Set rsMember = Nothing
					Response.end
				Else
					Set rsMember = Nothing
					Exit Function
				End If
			End If
			rsMember.Close: Set rsMember = Nothing
			Exit Function
		End If
		CookiesID = "ArticleID_" & ArticleID
		If Trim(Request.Cookies("ReadArticle")) = "" Then
			Response.Cookies("ReadArticle")("userip") = enchiasp.GetUserip
			Response.Cookies("ReadArticle").Expires = Date + 1
		End If
		
		If CLng(Request.Cookies("ReadArticle")(CookiesID)) <> CLng(ArticleID) And CInt(UserGroup) > 0 Then
			Set rsMember = CreateObject("ADODB.Recordset")
			SQL = "SELECT userid,UserGrade,userpoint,ExpireTime FROM ECCMS_User WHERE username='" & enchiasp.membername & "' And userid=" & CLng(enchiasp.memberid)
			rsMember.Open SQL, Conn, 1, 3
			If rsMember.BOF And rsMember.EOF Then
				Message = "<li>非法操作~！</li>"
				Response.Redirect (strInstallDir & "showerr.asp?action=error&Message=" & Server.URLEncode(Message))
				Set rsMember = Nothing
				Response.end
			Else
				If CInt(rsMember("UserGrade")) < CInt(UserGroup) Then
					Message = "<li>您的级别不够，阅览此文章需要<font color=blue>" & GroupName & "</font>以上级别的会员；</li><li>如果你要阅览此文章请联系管理员。</li>"
					Response.Redirect (strInstallDir & "showerr.asp?action=error&Message=" & Server.URLEncode(Message))
					Set rsMember = Nothing
					Response.end
				End If
				
				If CLng(rsMember("userpoint")) < CLng(PointNum) Then
					Message = "<li>对不起!您的点数不足。不能阅览此文章</li><li>阅览此文章所需的点数：" & PointNum & "</li><li>如果你确实要阅览此文章请到<a href=""../user/"">会员中心</a>充值。</li>"
					Response.Redirect (strInstallDir & "showerr.asp?action=error&Message=" & Server.URLEncode(Message))
					Set rsMember = Nothing
					Response.end
				End If
				rsMember("userpoint") = CLng(rsMember("userpoint") - PointNum)
				rsMember.Update
				Response.Cookies("ReadArticle")(CookiesID) = ArticleID
			End If
			rsMember.Close: Set rsMember = Nothing
		End If
	End Function
'=================================================
	'过程名：ContentPagination
	'作  用：以分页方式显示文章具体的内容
	'参  数：无
	'=================================================
	Private Sub ContentPagination()
		Dim ContentLen, maxperpage, Paginate
		Dim arrContent, strContent, i
		
		On Error Resume Next
		strContent = enchiasp.ReadContent(Rs("Content"))
		strContent = Replace(strContent, "[NextPage]", "[page_break]")
		strContent = Replace(strContent, "[Page_Break]", "[page_break]")
		ContentLen = Len(strContent)
		If InStr(strContent, "[page_break]") <= 0 Then
			ArticleContent = strContent
		Else
			arrContent = Split(strContent, "[page_break]")

			Paginate = UBound(arrContent) + 1
			If CurrentPage = 0 Then
				CurrentPage = 1
			Else
				CurrentPage = CLng(CurrentPage)
			End If
			If CurrentPage < 1 Then CurrentPage = 1
			If CurrentPage > Paginate Then CurrentPage = Paginate

			ArticleContent = ArticleContent & arrContent(CurrentPage - 1)

			ArticleContent = ArticleContent & "</p><p align='center'><b>"
			If CurrentPage > 1 Then
					ArticleContent = ArticleContent & "<a href='?classid=" & ArticleID & "&Page=" & CurrentPage - 1 & "'>上一页</a>&nbsp;&nbsp;"
			End If
			For i = 1 To Paginate
				If i = CurrentPage Then
					ArticleContent = ArticleContent & "<font color='red'>[" & CStr(i) & "]</font>&nbsp;"
				Else
					ArticleContent = ArticleContent & "<a href='?classid=" & ArticleID & "&Page=" & i & "'>[" & i & "]</a>&nbsp;"
				End If
			Next
			If CurrentPage < Paginate Then
				ArticleContent = ArticleContent & "&nbsp;<a href='?classid=" & ArticleID & "&Page=" & CurrentPage + 1 & "'>下一页</a>"
			
			End If
			ArticleContent = ArticleContent & "</b></p>"
		End If
	End Sub

	
		
	
%>