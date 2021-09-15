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
Dim enchicms
Set enchicms = New NewsChannel_Cls

Class NewsChannel_Cls
	Private ChannelID, CreateHtml, keyword
	Private Rs, SQL, ChannelRootDir, HtmlContent, strIndexName
	Private ArticleID, ArticleContent, skinid, ClassID
	Private maxperpage, TotalNumber, TotalPageNum, CurrentPage, i, totalrec
	Private strFileDir, ParentID, strParent, strClassName, ChildStr, Child
	Private ListContent, TempListContent, HtmlTemplate, HtmlFilePath
	Private SpecialID, SpecialName, SpecialDir, PageType, ForbidEssay, strInstallDir
	Private IsShowFlush, j
	Private FoundErr,strlen

	Private Sub Class_Initialize()
		On Error Resume Next
		FoundErr = False
		ChannelID = 1
		IsShowFlush = 0
		strlen = 0
	End Sub
	Private Sub Class_Terminate()
		'Set HTML = Nothing
	End Sub
	Public Property Let Channel(chanid)
		ChannelID = chanid
	End Property
	Public Property Let ShowFlush(para)
		IsShowFlush = para
	End Property
	Public Sub ChannelMain()
		enchiasp.ReadChannel (ChannelID)
		CreateHtml = CInt(enchiasp.IsCreateHtml)
		ChannelRootDir = enchiasp.InstallDir & enchiasp.ChannelDir
		strInstallDir = enchiasp.InstallDir
		strIndexName = "<a href='" & ChannelRootDir & "'>" & enchiasp.ChannelName & "</a>"
	End Sub
	'#############################\\执行文章首页开始//#############################
	'=================================================
	'过程名：ShowArticleIndex
	'作  用：显示文章首页
	'=================================================
	Public Sub ShowArticleIndex()
		On Error Resume Next
		LoadArticleIndex
		If CreateHtml <> 0 Then
			Response.Write "<meta http-equiv=refresh content=0;url=index" & enchiasp.HtmlExtName & ">"
		Else
			Response.Write HtmlContent
		End If
	End Sub
	'=================================================
	'过程名：CreateArticleIndex
	'作  用：生成文章首页的HTML
	'=================================================
	Public Sub CreateArticleIndex()
		On Error Resume Next
		LoadArticleIndex
		Dim FilePath
		FilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "index" & enchiasp.HtmlExtName
		enchiasp.CreatedTextFile FilePath, HtmlContent
		If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "首页HTML完成... <a href=" & FilePath & " target=_blank>" & Server.MapPath(FilePath) & "</a></li>" & vbNewLine
		Response.Flush
	End Sub
	'=================================================
	'过程名：LoadArticleIndex
	'作  用：装载文章首页
	'=================================================
	Private Sub LoadArticleIndex()
		On Error Resume Next

		enchiasp.LoadTemplates ChannelID, 1, enchiasp.ChannelSkin
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ChannelName)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ArticleIndex}", strIndexName)
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = HTML.ReadArticlePic(HtmlContent)
		HtmlContent = HTML.ReadSoftPic(HtmlContent)
		HtmlContent = HTML.ReadArticleList(HtmlContent)
		HtmlContent = HTML.ReadSoftList(HtmlContent)
		HtmlContent = HTML.ReadShopList(HtmlContent)
		HtmlContent = HTML.ReadFlashList(HtmlContent)
		HtmlContent = HTML.ReadShopPic(HtmlContent)
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
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = HtmlContent
	End Sub
	'##############################################################################
	'#############################\\执行文章内容开始//#############################
	'=================================================
	'过程名：ShowArticleInfo
	'作  用：显示文章内容页面
	'=================================================
	Public Sub ShowArticleInfo()
		If CreateHtml <> 0 Then
			Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
			Exit Sub
		Else
			enchiasp.PreventInfuse
			ArticleID = enchiasp.ChkNumeric(Request("id"))
			CurrentPage = enchiasp.ChkNumeric(Request("Page"))
			Response.Write ReadArticleContent(ArticleID, CurrentPage)
		End If
	End Sub

	Private Function CheckUserRead(ByVal ArticleID, ByVal PointNum, ByVal UserGroup, ByVal User_Group)
		Dim Message, CookiesID
		Dim GroupSetting, GroupName, gradeid
		
		If CInt(enchiasp.membergrade) = 999 Then Exit Function
		If CInt(enchiasp.membergrade) <> 0 Then
			gradeid = CInt(enchiasp.membergrade)
		Else
			gradeid = 0
		End If
		GroupSetting = Split(enchiasp.UserGroupSetting(gradeid), "|||")
		GroupName = GroupSetting(UBound(GroupSetting))
		If CInt(User_Group) > CInt(gradeid) Or CInt(UserGroup) > CInt(gradeid) Then
			Message = "<li>您没有登录或者你的会员级别不够，不能阅览此文章！</li><li>如果你是本站会员, 请先<a href=""../user/"">登陆</a></li>"
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
	'函数名：ReadArticleContent
	'作  用：读取文章内容
	'参  数：ArticleID ----文章ID
	'=================================================
	Private Function ReadArticleContent(ArticleID, CurrentPage)
		On Error Resume Next
		If Not IsNumeric(ArticleID) Then
			Exit Function
		Else
			ArticleID = CLng(ArticleID)
		End If
		SQL = "SELECT A.ArticleID,A.ClassID,A.title,A.content,A.Related,A.Author,A.ComeFrom,A.isTop,A.username,A.star,A.isBest,A.WriteTime,A.Allhits,A.HtmlFileDate,A.UserGroup,A.PointNum,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,C.UserGroup As User_Group,C.UseHtml FROM [ECCMS_Article] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.ArticleID=" & ArticleID
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			ReadArticleContent = ""
			Set Rs = Nothing
			If CreateHtml = 0 Then
				Response.Write "<meta http-equiv=""refresh"" content=""2;url='/"">" & vbNewLine
				Response.Write "<p align=""center"" style=""font-size: 16px;color: red;"">对不起，该页面发生了错误，无法访问! 系统两秒后自动转到网站首页......</p>" & vbNewLine
			End If
			Exit Function
		End If
		If Rs("skinid") <> 0 Then
			skinid = Rs("skinid")
		Else
			skinid = enchiasp.ChannelSkin
		End If
		enchiasp.LoadTemplates ChannelID, 3, skinid

		If CreateHtml <> 0 Then
			ArticleContent = HtmlPagination(CurrentPage)
		Else
			CheckUserRead Rs("ArticleID"), Rs("PointNum"), Rs("UserGroup"), Rs("User_Group")
			Author=enchiasp.ChkNull(Rs("Author"))
ComeFrom=Rs("ComeFrom")
WriteTime=Rs("WriteTime")
username=Rs("username")
articletitle=Rs("title")
xParentID=Rs("ParentID")
xParentStr=Rs("ParentStr")
xHtmlFileDir=Rs("HtmlFileDir")
			Call ContentPagination

		End If
		HtmlContent = enchiasp.HtmlContent
if enchiasp.HtmlSetting(8)=0 then
	ArticleContent = Replace(ArticleContent, "{$zidongsuofang}", "")
else
	ArticleContent = Replace(ArticleContent, "{$zidongsuofang}", enchiasp.HtmlSetting(9))
end if
ArticleContent = Replace(ArticleContent, "{$zidongsuofang}", "")
				HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ArticleIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)

            	HtmlContent = Replace(HtmlContent, "{$ArticleTitle}", articletitle)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", articletitle)

		HtmlContent = Replace(HtmlContent, "{$ClassID}", ClassID)
		HtmlContent = Replace(HtmlContent, "{$ArticleID}", ArticleID)
	
 
		HtmlContent = Replace(HtmlContent, "{$ArticleContent}", ArticleContent)
		HtmlContent = Replace(HtmlContent, "{$Author}", Author)
		HtmlContent = Replace(HtmlContent, "{$ComeFrom}",ComeFrom)
		HtmlContent = Replace(HtmlContent, "{$WriteTime}", WriteTime)
		HtmlContent = Replace(HtmlContent, "{$UserName}", username)
		
		'HtmlContent = Replace(HtmlContent, "{$Star}", Rs("star"))
		'HtmlContent = Replace(HtmlContent, "{$Best}", Rs("isBest"))
		
		If InStr(HtmlContent, "{$FrontArticle}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$FrontArticle}", FrontArticle(ArticleID))
		End If
		If InStr(HtmlContent, "{$NextArticle}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$NextArticle}", NextArticle(ArticleID))
		End If
		If InStr(HtmlContent, "{$RelatedArticle}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$RelatedArticle}", RelatedArticle(Rs("Related"), Rs("title"), ArticleID))
		End If
		If InStr(HtmlContent, "{$ShowHotArticle}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$ShowHotArticle}", ReadHotArticle(Rs("ClassID")))
		End If
		If InStr(HtmlContent, "{$ArticleComment}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$ArticleComment}", ArticleComment(Rs("ArticleID")))
		End If
		HtmlContent = HTML.ReadCurrentStation(HtmlContent, ChannelID, Rs("ClassID"), Rs("ClassName"), Rs("ParentID"), Rs("ParentStr"), Rs("HtmlFileDir"))
		HtmlContent = HTML.ReadAnnounceList(HtmlContent)		
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = HTML.ReadArticlePic(HtmlContent)
		HtmlContent = HTML.ReadArticleList(HtmlContent)
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
		HtmlContent = HTML.ReadShopList(HtmlContent)
call ReplaceContent
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		ReadArticleContent = HtmlContent
		Rs.Close: Set Rs = Nothing
	End Function
	'判断是否是单页面图文
	Private Function isdanyemian(ChannelID)
		SQL = "SELECT * from [ECCMS_Channel] where modules=6 and ChannelID="& ChannelID &""
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
				isdanyemian = false
		else
				isdanyemian = true
		End If

		Rs.Close: Set Rs = Nothing
	End Function

	'=================================================
	'过程名：CreateArticleContent
	'作  用：生成文章内容
	'参  数：ArticleID ----文章ID
	'=================================================
	Public Function CreateArticleContent(ArticleID)
		Dim arrContent, Paginate, rsCreate, HtmlFileName, strHtmlContent
		Dim sContentText, i
		
		On Error Resume Next
		If CreateHtml = 0 Then Exit Function
		
		SQL = "select A.ArticleID,A.title,A.content,A.HtmlFileDate,C.HtmlFileDir from [ECCMS_Article] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.ArticleID=" & ArticleID
		Set rsCreate = enchiasp.Execute(SQL)
		If rsCreate.BOF And rsCreate.EOF Then
			Set rsCreate = Nothing
			Exit Function
		End If
		
		HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & rsCreate("HtmlFileDir") & enchiasp.ShowDatePath(rsCreate("HtmlFileDate"), enchiasp.HtmlPath)
		enchiasp.CreatPathEx (HtmlFilePath)
		sContentText = Replace(rsCreate("Content"), "[NextPage]", "[page_break]")
		sContentText = Replace(sContentText, "[Page_Break]", "[page_break]")
		arrContent = Split(sContentText, "[page_break]")
		Paginate = UBound(arrContent)
		Response.Flush
		For i = 1 To Paginate + 1
			strHtmlContent = ReadArticleContent(rsCreate("ArticleID"), i)
			'判断单页面
			if isdanyemian(ChannelID) then
				HtmlFileName = HtmlFilePath & "index"& enchiasp.HtmlExtName

			else
				HtmlFileName = HtmlFilePath & enchiasp.ReadFileName(rsCreate("HtmlFileDate"), rsCreate("ArticleID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, i)
			
			end if
			enchiasp.CreatedTextFile HtmlFileName, strHtmlContent
			If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "内容HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
			Response.Flush
		Next
		rsCreate.Close: Set rsCreate = Nothing
	End Function
	'=================================================
	'函数名：FrontArticle
	'作  用：显示上一篇文章
	'参  数：ArticleID ----文章ID
	'=================================================
	Private Function FrontArticle(ArticleID)
		Dim rsContext, SQL, HtmlFileUrl, HtmlFileName
		On Error Resume Next
		SQL = "select Top 1 A.ArticleID,A.ClassID,A.title,A.HtmlFileDate,C.HtmlFileDir from [ECCMS_Article] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.ArticleID < " & ArticleID & " order by A.ArticleID desc"
		Set rsContext = enchiasp.Execute(SQL)
		If rsContext.EOF And rsContext.BOF Then
			FrontArticle = "已经没有了"
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & rsContext("HtmlFileDir") & enchiasp.ShowDatePath(rsContext("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(rsContext("HtmlFileDate"), rsContext("ArticleID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				FrontArticle = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & rsContext("title") & "</a>"
			Else
				FrontArticle = "<a href=?id=" & rsContext("ArticleID") & ">" & rsContext("title") & "</a>"
			End If
		End If
		rsContext.Close
		Set rsContext = Nothing
	End Function
	'=================================================
	'函数名：NextArticle
	'作  用：显示下一篇文章
	'参  数：ArticleID ----文章ID
	'=================================================
	Private Function NextArticle(ArticleID)
		Dim rsContext, SQL, HtmlFileUrl, HtmlFileName
		On Error Resume Next
		SQL = "select Top 1 A.ArticleID,A.ClassID,A.title,A.HtmlFileDate,C.HtmlFileDir from [ECCMS_Article] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.ArticleID > " & ArticleID & " order by A.ArticleID asc"
		Set rsContext = enchiasp.Execute(SQL)
		If rsContext.EOF And rsContext.BOF Then
			NextArticle = "已经没有了"
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & rsContext("HtmlFileDir") & enchiasp.ShowDatePath(rsContext("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(rsContext("HtmlFileDate"), rsContext("ArticleID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				NextArticle = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & rsContext("title") & "</a>"
			Else
				NextArticle = "<a href=?id=" & rsContext("ArticleID") & ">" & rsContext("title") & "</a>"
			End If
		End If
		rsContext.Close
		Set rsContext = Nothing
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
		
strContent = Replace(strContent, "<img", "<img {$zidongsuofang}")


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
				if isdanyemian(ChannelID) then
					ArticleContent = ArticleContent & "<a href='?classid=" & ArticleID & "&Page=" & CurrentPage - 1 & "'>上一页</a>&nbsp;&nbsp;"

				else
					ArticleContent = ArticleContent & "<a href='?id=" & ArticleID & "&Page=" & CurrentPage - 1 & "'>上一页</a>&nbsp;&nbsp;"
				end if
			End If
			For i = 1 To Paginate
				If i = CurrentPage Then
					ArticleContent = ArticleContent & "</font><font color='red'>[" & CStr(i) & "]</font>&nbsp;"
						'ArticleContent = ArticleContent & "<font color='red'>[" & CStr(i) & "]</font>&nbsp;"
				Else
					if isdanyemian(ChannelID) then
						ArticleContent = ArticleContent & "<a href='?classid=" & ArticleID & "&Page=" & i & "'>[" & i & "]</a>&nbsp;"

					else
						ArticleContent = ArticleContent & "<a href='?id=" & ArticleID & "&Page=" & i & "'>[" & i & "]</a>&nbsp;"
					end if
				End If
			Next
			If CurrentPage < Paginate Then
				if isdanyemian(ChannelID) then

					ArticleContent = ArticleContent & "&nbsp;<a href='?classid=" & ArticleID & "&Page=" & CurrentPage + 1 & "'>下一页</a>"
				else
					ArticleContent = ArticleContent & "&nbsp;<a href='?id=" & ArticleID & "&Page=" & CurrentPage + 1 & "'>下一页</a>"

				end if
			End If
			ArticleContent = ArticleContent & "</b></p>"
		End If
	End Sub
	'=================================================
	'函数名：HtmlPagination
	'作  用：以分页方式显示文章具体的内容
	'参  数：无
	'=================================================
	Private Function HtmlPagination(n)
		Dim ContentLen, CurrentPage, maxperpage, Paginate
		Dim arrContent, strContent, TempContent, i
		
		On Error Resume Next
		strContent = enchiasp.ReadContent(Rs("content"))
		ContentLen = Len(strContent)
		CurrentPage = CInt(n)
		If InStr(strContent, "[page_break]") <= 0 Then
			TempContent = strContent
		Else
			arrContent = Split(strContent, "[page_break]")

			Paginate = UBound(arrContent) + 1
			If CurrentPage = 0 Then
				CurrentPage = 1
			Else
				CurrentPage = CInt(CurrentPage)
			End If
			If CurrentPage < 1 Then CurrentPage = 1
			If CurrentPage > Paginate Then CurrentPage = Paginate

			TempContent = TempContent & arrContent(CurrentPage - 1)

			TempContent = TempContent & "</p><p align='center'><b>"
			If CurrentPage > 1 Then
				TempContent = TempContent & "<a href='" & ReadPagination(CurrentPage - 1) & "'>上一页</a>&nbsp;&nbsp;"
			End If
			For i = 1 To Paginate
				If i = CurrentPage Then
					TempContent = TempContent & "<font color='red'>[" & i & "]</font>&nbsp;"
				Else
					TempContent = TempContent & "<a href='" & ReadPagination(i) & "'>[" & i & "]</a>&nbsp;"
				End If
			Next
			If CurrentPage < Paginate Then
				TempContent = TempContent & "&nbsp;<a href='" & ReadPagination(CurrentPage + 1) & "'>下一页</a>"
			End If
			TempContent = TempContent & "</b></p>"
		End If
		HtmlPagination = TempContent
	End Function
	Private Function ReadPagination(n)
		Dim HtmlFileName, CurrentPage
		On Error Resume Next
		CurrentPage = n
		HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("ArticleID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, CurrentPage)
		ReadPagination = HtmlFileName
	End Function
	'=================================================
	'函数名：RelatedArticle
	'作  用：显示相关文章
	'参  数：sRelated ----相关文章
	'=================================================
	Private Function RelatedArticle(sRelated, topic, ArticleID)
		Dim rsRdlated, SQL, HtmlFileUrl, HtmlFileName
		Dim strTitle, strTopic, ArticleTitle, strContent
		Dim strRelated, arrRelated, i, Resize, strRearrange
		Dim strKey
		Dim ArrayTemp()
		
		On Error Resume Next
		strRelated = Replace(Replace(Replace(Replace(Replace(sRelated, "[", ""), "]", ""), "'", ""), "(", ""), ")", "")
		strKey = Left(enchiasp.ChkQueryStr(topic), 5)
		If Not IsNull(sRelated) And sRelated <> Empty Then
			If InStr(strRelated, "|") > 1 Then
				arrRelated = Split(strRelated, "|")
				strRelated = "((A.title like '%" & arrRelated(0) & "%')"
				For i = 1 To UBound(arrRelated)
					strRelated = strRelated & " Or (A.title like '%" & arrRelated(i) & "%')"
				Next
				'strRelated = strRelated & ")"
			Else
				strRelated = "((A.title like '%" & strRelated & "%')"
			End If
			strRelated = strRelated & " Or (A.title like '%" & strKey & "%'))"
		Else
			strRelated = "(A.title like '%" & strKey & "%')"
		End If
		SQL = "SELECT TOP " & CInt(enchiasp.HtmlSetting(1)) & " A.ArticleID,A.ClassID,A.ColorMode,A.FontMode,A.title,A.BriefTopic,A.AllHits,A.WriteTime,A.HtmlFileDate,C.HtmlFileDir FROM [ECCMS_Article] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And " & strRelated & " And A.ArticleID <> " & ArticleID & " ORDER BY A.ArticleID DESC"
		Set rsRdlated = enchiasp.Execute(SQL)
		If rsRdlated.EOF And rsRdlated.BOF Then
			RelatedArticle = ""
			Set rsRdlated = Nothing
			Exit Function
		Else
			i = 0
			Resize = 0
			Do While Not rsRdlated.EOF
				ReDim Preserve ArrayTemp(i + Resize)
				strContent = ArrayTemp(i) & enchiasp.HtmlSetting(4)
				strTitle = enchiasp.GotTopic(rsRdlated("Title"), CInt(enchiasp.HtmlSetting(2)))
				strTitle = enchiasp.ReadFontMode(strTitle, rsRdlated("ColorMode"), rsRdlated("FontMode"))
				strTopic = enchiasp.ReadPicTopic(rsRdlated("BriefTopic"))
				If CreateHtml <> 0 Then
					HtmlFileUrl = ChannelRootDir & rsRdlated("HtmlFileDir") & enchiasp.ShowDatePath(rsRdlated("HtmlFileDate"), enchiasp.HtmlPath)
					HtmlFileName = enchiasp.ReadFileName(rsRdlated("HtmlFileDate"), rsRdlated("ArticleID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
					ArticleTitle = "<a href=" & HtmlFileUrl & HtmlFileName & " title='" & rsRdlated("title") & "'>" & strTitle & "</a>"
				Else
					ArticleTitle = "<a href=show.asp?id=" & rsRdlated("ArticleID") & " title='" & rsRdlated("title") & "'>" & strTitle & "</a>"
				End If
				strContent = Replace(strContent, "{$BriefTopic}", strTopic)
				strContent = Replace(strContent, "{$ArticleTitle}", ArticleTitle)
				strContent = Replace(strContent, "{$AllHits}", rsRdlated("AllHits"))
				strContent = Replace(strContent, "{$WriteTime}", enchiasp.ShowDateTime(rsRdlated("WriteTime"), CInt(enchiasp.HtmlSetting(3))))
				ArrayTemp(i) = strContent
				rsRdlated.MoveNext
				i = i + 1
			Loop
		End If
		rsRdlated.Close
		Set rsRdlated = Nothing
		strRearrange = Join(ArrayTemp, vbCrLf)
		RelatedArticle = strRearrange
	End Function
	'=================================================
	'函数名：ReadHotArticle
	'作  用：显示热门文章
	'参  数：ClassID ----文章分类ID
	'=================================================
	Private Function ReadHotArticle(ClassID)
		Dim rsHot, SQL, HtmlFileUrl, HtmlFileName
		Dim strTitle, strTopic, ArticleTitle, strContent
		Dim i, Resize, strRearrange
		Dim ArrayTemp()
		
		'On Error Resume Next
		SQL = "select Top " & CInt(enchiasp.HtmlSetting(1)) & " A.ArticleID,A.ClassID,A.ColorMode,A.FontMode,A.title,A.BriefTopic,A.AllHits,A.WriteTime,A.HtmlFileDate,C.HtmlFileDir from [ECCMS_Article] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.AllHits >= " & CLng(enchiasp.LeastHotHist) & " order by A.AllHits desc,A.ArticleID desc"
		Set rsHot = enchiasp.Execute(SQL)
		If rsHot.EOF And rsHot.BOF Then
			ReadHotArticle = ""
			Set rsHot = Nothing
			Exit Function
		Else
			i = 0
			Resize = 0
			Do While Not rsHot.EOF
				ReDim Preserve ArrayTemp(i + Resize)
				strContent = ArrayTemp(i) & enchiasp.HtmlSetting(4)
				strTitle = enchiasp.GotTopic(rsHot("Title"), CInt(enchiasp.HtmlSetting(2)))
				strTitle = enchiasp.ReadFontMode(strTitle, rsHot("ColorMode"), rsHot("FontMode"))
				strTopic = enchiasp.ReadPicTopic(rsHot("BriefTopic"))
				If CreateHtml <> 0 Then
					HtmlFileUrl = ChannelRootDir & rsHot("HtmlFileDir") & enchiasp.ShowDatePath(rsHot("HtmlFileDate"), enchiasp.HtmlPath)
					HtmlFileName = enchiasp.ReadFileName(rsHot("HtmlFileDate"), rsHot("ArticleID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
					ArticleTitle = "<a href=" & HtmlFileUrl & HtmlFileName & " title='" & rsHot("title") & "'>" & strTitle & "</a>"
				Else
					ArticleTitle = "<a href=show.asp?id=" & rsHot("ArticleID") & " title='" & rsHot("title") & "'>" & strTitle & "</a>"
				End If
				strContent = Replace(strContent, "{$BriefTopic}", strTopic)
				strContent = Replace(strContent, "{$ArticleTitle}", ArticleTitle)
				strContent = Replace(strContent, "{$AllHits}", rsHot("AllHits"))
				strContent = Replace(strContent, "{$WriteTime}", enchiasp.ShowDateTime(rsHot("WriteTime"), CInt(enchiasp.HtmlSetting(3))))
				ArrayTemp(i) = strContent
				rsHot.MoveNext
				i = i + 1
			Loop
		End If
		rsHot.Close
		Set rsHot = Nothing
		strRearrange = Join(ArrayTemp, vbCrLf)
		ReadHotArticle = strRearrange
	End Function
	'================================================
	'函数名：ArticleComment
	'作  用：文章评论
	'参  数：ArticleID ----文章ID
	'================================================
	Private Function ArticleComment(ArticleID)
		Dim rsComment, SQL, strContent, strComment
		Dim i, Resize, strRearrange
		Dim ArrayTemp()
		
		On Error Resume Next
		Set rsComment = enchiasp.Execute("Select Top " & CInt(enchiasp.HtmlSetting(5)) & " content,Grade,username,postime,postip From ECCMS_Comment where ChannelID=" & ChannelID & " And postid = " & ArticleID & " order by postime desc,CommentID desc")
		If Not (rsComment.EOF And rsComment.BOF) Then
			i = 0
			Resize = 0
			Do While Not rsComment.EOF
				ReDim Preserve ArrayTemp(i + Resize)
				strContent = ArrayTemp(i) & enchiasp.HtmlSetting(7)
				strComment = enchiasp.CutString(rsComment("content"), CInt(enchiasp.HtmlSetting(6)))
				strContent = Replace(strContent, "{$Comment}", enchiasp.HTMLEncode(strComment))
				strContent = Replace(strContent, "{$UserName}", enchiasp.HTMLEncode(rsComment("username")))
				strContent = Replace(strContent, "{$UserGrade}", rsComment("Grade"))
				strContent = Replace(strContent, "{$postime}", rsComment("postime"))
				strContent = Replace(strContent, "{$postip}", rsComment("postip"))
				ArrayTemp(i) = strContent
				rsComment.MoveNext
				i = i + 1
			Loop
		End If
		rsComment.Close
		strRearrange = Join(ArrayTemp, vbCrLf)
		Set rsComment = Nothing
		ArticleComment = strRearrange
	End Function
	'================================================
	'函数名：CurrentStation
	'作  用：当前位置
	'参  数：...
	'================================================
	Private Function CurrentStation(ChannelID, ClassID, ClassName, ParentID, strParent, HtmlFileDir, Compart)
		Dim rsCurrent, SQL, strContent, ChannelDir
		
		On Error Resume Next
		ChannelDir = ChannelRootDir
		If ParentID <> 0 And Len(strParent) <> 0 Then
			SQL = "select ClassID,ClassName,HtmlFileDir from [ECCMS_Classify] where ChannelID = " & ChannelID & " And ClassID in(" & strParent & ")"
			Set rsCurrent = enchiasp.Execute(SQL)
			If Not (rsCurrent.EOF And rsCurrent.BOF) Then
				Do While Not rsCurrent.EOF
					If CInt(enchiasp.IsCreateHtml) <> 0 Then
						strContent = strContent & "<a href='" & ChannelDir & rsCurrent("HtmlFileDir") & "'>" & rsCurrent(1) & "</a>" & Compart & ""
					Else
						strContent = strContent & "<a href='" & ChannelDir & "list.asp?classid=" & rsCurrent("ClassID") & "'>" & rsCurrent("ClassName") & "</a>" & Compart & ""
					End If
					rsCurrent.MoveNext
				Loop
			End If
			rsCurrent.Close
			Set rsCurrent = Nothing
		End If
		If CInt(enchiasp.IsCreateHtml) <> 0 Then
			strContent = strContent & "<a href='" & ChannelDir & HtmlFileDir & "'>" & ClassName & "</a>"
		Else
			strContent = strContent & "<a href='" & ChannelDir & "list.asp?classid=" & ClassID & "'>" & ClassName & "</a>"
		End If
		CurrentStation = strContent
	End Function
	'================================================
	'函数名：ReadCurrentStation
	'作  用：读取当前位置
	'参  数：str ----原字符串
	'================================================
	Private Function ReadCurrentStation(str, ChannelID, ClassID, ClassName, ParentID, strParent, HtmlFileDir)
		Dim strTemp, i, sTempContent, nTempContent
		Dim arrTempContent, arrTempContents
		
		On Error Resume Next
		strTemp = str
		sTempContent = enchiasp.CutMatchContent(strTemp, "{#CurrentStation(", ")}", 1)
		nTempContent = enchiasp.CutMatchContent(strTemp, "{#CurrentStation(", ")}", 0)
		arrTempContents = Split(sTempContent, "|||")
		arrTempContent = Split(nTempContent, "|||")
		For i = 0 To UBound(arrTempContents)
			strTemp = Replace(strTemp, arrTempContents(i), CurrentStation(ChannelID, ClassID, ClassName, ParentID, strParent, HtmlFileDir, arrTempContent(i)))
		Next
		ReadCurrentStation = strTemp
	End Function
	'##############################################################################
	'#############################\\执行文章列表开始//#############################
	Public Sub ShowArticleList()
		On Error Resume Next
		'判断是否是单页面
		if isdanyemian(ChannelID) then
				If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
					CurrentPage = enchiasp.ChkNumeric(Request("page"))
				Else
					CurrentPage = 1
				End If

				ClassID = enchiasp.ChkNumeric(Request("ClassID"))
				Response.Write showdanyemian(ClassID, 1)

		else
			If CreateHtml <> 0 Then
				Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
				Exit Sub
			Else
				enchiasp.PreventInfuse
				If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
					Response.Write ("错误的系统参数!请输入整数")
					Response.end
				End If
				If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
					CurrentPage = enchiasp.ChkNumeric(Request("page"))
				Else
					CurrentPage = 1
				End If
				ClassID = enchiasp.ChkNumeric(Request("ClassID"))
				Response.Write CreateArticleList(ClassID, 1)
			End If
		end if
		
	End Sub
	'================================================
	'函数名：showdanyemian
	'作  用：列出单页面内容
	'================================================
	Public Function showdanyemian(clsid, n)
		On Error Resume Next
		If Not IsNumeric(clsid) Then
			Exit Function
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
			Exit Function
			
		End If
		If rsClass("skinid") <> 0 Then
				skinid = rsClass("skinid")
			Else
				skinid = CLng(enchiasp.ChannelSkin)
			End If

		enchiasp.LoadTemplates ChannelID, 3, skinid

		If CreateHtml <> 0 Then
			ArticleContent = HtmlPagination(CurrentPage)
		Else
			CheckUserRead Rs("ArticleID"), Rs("PointNum"), Rs("UserGroup"), Rs("User_Group")
			Call ContentPagination
		End If
	
		HtmlContent = enchiasp.HtmlContent
		

		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ArticleIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", Rs("title"))
		HtmlContent = Replace(HtmlContent, "{$ClassID}", Rs("ClassID"))
		HtmlContent = Replace(HtmlContent, "{$ArticleID}", ArticleID)
		HtmlContent = Replace(HtmlContent, "{$ArticleTitle}", Rs("title"))
		HtmlContent = Replace(HtmlContent, "{$ArticleContent}", ArticleContent)
		HtmlContent = Replace(HtmlContent, "{$Author}", enchiasp.ChkNull(Rs("Author")))
		HtmlContent = Replace(HtmlContent, "{$ComeFrom}", Rs("ComeFrom"))
		HtmlContent = Replace(HtmlContent, "{$WriteTime}", Rs("WriteTime"))
		HtmlContent = Replace(HtmlContent, "{$UserName}", Rs("username"))
		HtmlContent = Replace(HtmlContent, "{$Star}", Rs("star"))
		HtmlContent = Replace(HtmlContent, "{$Best}", Rs("isBest"))
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ChannelName)
		
		HtmlContent = HTML.ReadAnnounceList(HtmlContent)		
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = HTML.ReadArticlePic(HtmlContent)
		HtmlContent = HTML.ReadSoftPic(HtmlContent)
		HtmlContent = HTML.ReadArticleList(HtmlContent)
		HtmlContent = HTML.ReadSoftList(HtmlContent)
		HtmlContent = HTML.ReadShopList(HtmlContent)
		HtmlContent = HTML.ReadFlashList(HtmlContent)
		HtmlContent = HTML.ReadShopPic(HtmlContent)
		HtmlContent = HTML.ReadFlashPic(HtmlContent)
		HtmlContent = HTML.ReadFriendLink(HtmlContent)
		HtmlContent = HTML.ReadNewsPicAndText(HtmlContent)
		HtmlContent = HTML.ReadSoftPicAndText(HtmlContent)
		HtmlContent = HTML.ReadGuestList(HtmlContent)
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
		HtmlContent = HTML.ReadPopularSoft(HtmlContent)
		HtmlContent = HTML.ReadPopularFlash(HtmlContent)
		HtmlContent = HTML.ReadSoftType(HtmlContent)
		HtmlContent = HTML.ReadStatistic(HtmlContent)
		HtmlContent = HTML.ReadUserRank(HtmlContent)
	
		
		
		If InStr(HtmlContent, "{$FrontArticle}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$FrontArticle}", FrontArticle(ArticleID))
		End If
		If InStr(HtmlContent, "{$NextArticle}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$NextArticle}", NextArticle(ArticleID))
		End If
		If InStr(HtmlContent, "{$RelatedArticle}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$RelatedArticle}", RelatedArticle(Rs("Related"), Rs("title"), ArticleID))
		End If
		If InStr(HtmlContent, "{$ShowHotArticle}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$ShowHotArticle}", ReadHotArticle(Rs("ClassID")))
		End If
		If InStr(HtmlContent, "{$ArticleComment}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$ArticleComment}", ArticleComment(Rs("ArticleID")))
		End If
		HtmlContent = HTML.ReadCurrentStation(HtmlContent, ChannelID, Rs("ClassID"), Rs("ClassName"), Rs("ParentID"), Rs("ParentStr"), Rs("HtmlFileDir"))
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		showdanyemian = HtmlContent
		Rs.Close: Set Rs = Nothing

	end function
	
	
	
	'================================================
	'函数名：CreateArticleList
	'作  用：生成文章列表
	'================================================
	Public Function CreateArticleList(clsid, n)
		On Error Resume Next
		Dim rsClass, TemplateContent, strTemplate, strOrder
		Dim ParentTemplate, ChildTemplate, HtmlFileName
		Dim MaxListnum, strMaxListop, showtree
		
		If Not IsNumeric(clsid) Then Exit Function
		SQL = "select ClassID,ClassName,ChildStr,ParentID,ParentStr,Child,skinid,HtmlFileDir,UseHtml from [ECCMS_Classify] where ChannelID = " & ChannelID & " And ClassID=" & clsid
		Set rsClass = enchiasp.Execute(SQL)
		If rsClass.BOF And rsClass.EOF Then
			If CreateHtml = 0 Then
				Response.Write "<meta http-equiv=""refresh"" content=""2;url='/"">" & vbNewLine
				Response.Write "<p align=""center"" style=""font-size: 16px;color: red;"">对不起，该页面发生了错误，无法访问! 系统两秒后自动转到网站首页......</p>" & vbNewLine
			End If
			Set rsClass = Nothing
			Exit Function
		Else
			strClassName = rsClass("ClassName")
			ClassID = rsClass("ClassID")
			ChildStr = rsClass("ChildStr")
			Child = rsClass("Child")
			strFileDir = rsClass("HtmlFileDir")
			ParentID = rsClass("ParentID")
			strParent = rsClass("ParentStr")
			If rsClass("skinid") <> 0 Then
				skinid = rsClass("skinid")
			Else
				skinid = CLng(enchiasp.ChannelSkin)
			End If
		End If
		rsClass.Close: Set rsClass = Nothing
		
		enchiasp.LoadTemplates ChannelID, 2, skinid

		PageType = 1
		HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & strFileDir
		strTemplate = Split(enchiasp.HtmlContent, "|||@@@|||")
		'-- 大类列表显示方式
		showtree = enchiasp.ChkNumeric(enchiasp.HtmlSetting(4))
		'-- 最多列表数
		MaxListnum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(5))
		
		strlen = enchiasp.ChkNumeric(enchiasp.HtmlSetting(10))
		If CInt(enchiasp.HtmlSetting(0)) <> 0 Then
			ParentTemplate = enchiasp.HtmlTop & strTemplate(1)
			ChildTemplate = strTemplate(0) & enchiasp.HtmlFoot
		Else
			ParentTemplate = strTemplate(1)
			ChildTemplate = strTemplate(0)
		End If
		If Child <> 0 And showtree <> 9 Then
			TemplateContent = ParentTemplate
		Else
			TemplateContent = ChildTemplate
		End If
		enchiasp.HTMLValue = TemplateContent
		HtmlContent = enchiasp.HTMLValue
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
                HtmlContent = Replace(HtmlContent, "{$ClassName}", strClassName)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ClassID}", ClassID)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", strClassName)
		HtmlContent = Replace(HtmlContent, "{$ArticleIndex}", strIndexName)

		If Child <> 0 And showtree <> 9 Then
			Call LoadParentList

			Call ReplaceContent
			If CInt(CreateHtml) <> 0 Then
				'创建分类目录
				enchiasp.CreatPathEx (HtmlFilePath)
				'开始生成父级分类的HTML页
				HtmlFileName = HtmlFilePath & ReadListPageName(ClassID, 0)
				enchiasp.CreatedTextFile HtmlFileName, HtmlContent
				If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "列表HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
				Response.Flush
			End If
		Else

			Call ReplaceContent
			maxperpage = enchiasp.ChkNumeric(enchiasp.HtmlSetting(1))
			If CLng(CurrentPage) = 0 Then CurrentPage = 1
			If enchiasp.CheckStr(LCase(Request("oredr"))) = "hits" Then
				strOrder = "order by A.isTop desc, A.AllHits desc ,A.ArticleID desc"
			ElseIf enchiasp.CheckStr(LCase(Request("oredr"))) = "topic" Then
				strOrder = "order by A.isTop desc, A.title desc ,A.ArticleID desc"
			Else
				strOrder = "order by A.isTop desc, A.WriteTime desc ,A.ArticleID desc"
			End If
			TotalNumber = enchiasp.Execute("Select Count(ArticleID) from ECCMS_Article where ChannelID = " & ChannelID & " And isAccept > 0 And ClassID in (" & ChildStr & ")")(0)
			totalrec = TotalNumber
			'-- 如果开启了父分类显示功能,限制显示数
			If Child > 0 And TotalNumber > MaxListnum And MaxListnum <> 999 Then
				strMaxListop = " TOP " & MaxListnum
				TotalNumber = MaxListnum
			Else
				strMaxListop = vbNullString
			End If
			
			TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
			If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
			If CurrentPage < 1 Then CurrentPage = 1
			If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
			Set Rs = CreateObject("ADODB.Recordset")
			SQL = "select " & strMaxListop & " A.ArticleID,A.ClassID,A.BriefTopic,A.ColorMode,A.FontMode,A.title,A.content,A.Related,A.Author,A.ComeFrom,A.isTop,A.username,A.star,A.isBest,A.WriteTime,A.Allhits,A.HtmlFileDate,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,A.imageurl from [ECCMS_Article] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.ClassID in (" & ChildStr & ") " & strOrder & ""
			Rs.Open SQL, Conn, 1, 1
			If Rs.BOF And Rs.EOF Then
				HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "还没有找到任何" & enchiasp.ModuleName & "")
				HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
				If CreateHtml <> 0 Then
					enchiasp.CreatPathEx (HtmlFilePath)
					HtmlFileName = HtmlFilePath & ReadListPageName(ClassID, CurrentPage)
					enchiasp.CreatedTextFile HtmlFileName, HtmlContent
					If IsShowFlush = 1 Then
						Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "列表HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
						Response.Flush
					End If
				End If
			Else
				TotalNumber = totalrec
				TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
				If CreateHtml <> 0 Then
					Call LoadChildListHtml(n)
				Else
					Call LoadChildListAsp
				End If
			End If
			Rs.Close: Set Rs = Nothing
		End If
		If CreateHtml = 0 Then CreateArticleList = HtmlContent
	End Function
	'================================================
	'过程名：ReplaceContent
	'作  用：替换模板内容
	'================================================
	Private Sub ReplaceContent()
		On Error Resume Next

		HtmlContent = HTML.ReadCurrentStation(HtmlContent, ChannelID, ClassID, strClassName, ParentID, strParent, strFileDir)

		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = ReadClassMenu(HtmlContent)

		HtmlContent = HTML.ReadArticlePic(HtmlContent)
		HtmlContent = HTML.ReadArticleList(HtmlContent)
		HtmlContent = HTML.ReadNewsPicAndText(HtmlContent)
		HtmlContent = HTML.ReadSoftPicAndText(HtmlContent)
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
		HtmlContent = HTML.ReadPopularSoft(HtmlContent)

		HtmlContent = HTML.ReadStatistic(HtmlContent)
                HtmlContent = HTML.ReadFriendLink(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)

	End Sub
	'================================================
	'过程名：LoadParentList
	'作  用：装载父级文章列表
	'================================================
	Private Sub LoadParentList()
		Dim rsClslist, strContent, i, showtree
		Dim ClassUrl, ClassNameStr
		
		showtree = Trim(enchiasp.HtmlSetting(4))
		PageType = 1
		On Error Resume Next
		TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
		If Not IsNull(TempListContent) Then
			SQL = "select Top " & CInt(enchiasp.HtmlSetting(5)) & " ClassID,ClassName,HtmlFileDir from [ECCMS_Classify] where ChannelID = " & ChannelID & " And TurnLink = 0 And ParentID=" & ClassID & " order by rootid asc, ClassID asc"
			Set rsClslist = enchiasp.Execute(SQL)
			If rsClslist.BOF And rsClslist.EOF Then
				Set rsClslist = Nothing
				Exit Sub
			Else
				If showtree <> "1" Then strContent = "<table width=""100%"" align=center border=0 cellpadding=0 cellspacing=0 class=tablist>" & vbCrLf
				Do While Not rsClslist.EOF
					If showtree <> "1" Then strContent = strContent & "<tr valign=""top"">" & vbCrLf
					For i = 1 To 2
						If showtree <> "1" Then strContent = strContent & "<td class=""tdlist"">"
						If Not (rsClslist.EOF) Then
							strContent = strContent & TempListContent
							If CInt(CreateHtml) <> 0 Then
								ClassUrl = ChannelRootDir & rsClslist("HtmlFileDir")
							Else
								ClassUrl = ChannelRootDir & "list.asp?classid=" & rsClslist("ClassID")
							End If
							ClassNameStr = "<a href=""" & ClassUrl & """ class=""showtitle"">" & rsClslist("ClassName") & "</a>"
							strContent = Replace(strContent, "{$ChannelID}", ChannelID)
							strContent = Replace(strContent, "{$ClassifyID}", rsClslist("ClassID"))
							strContent = Replace(strContent, "{$ClassName}", ClassNameStr)
							strContent = Replace(strContent, "{$ClassUrl}", ClassUrl)
							If showtree <> "1" Then strContent = strContent & "</td>" & vbCrLf
							rsClslist.MoveNext
						Else
							If showtree <> "1" Then strContent = strContent & "</td>" & vbCrLf
						End If
					Next
					If showtree <> "1" Then strContent = strContent & "</tr>" & vbCrLf
				Loop
				If showtree <> "1" Then strContent = strContent & "</table>" & vbCrLf
			End If
			HtmlContent = Replace(HtmlContent, TempListContent, strContent)
			HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
			HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
			rsClslist.Close: Set rsClslist = Nothing
		End If
	End Sub
	'================================================
	'过程名：LoadChildListHtml
	'作  用：装载子级文章列表HTML
	'================================================
	Private Sub LoadChildListHtml(n)
		Dim HtmlFileName
		Dim Perownum,ii,w
		Perownum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(8))
		
		If IsNull(TempListContent) Then Exit Sub
		On Error Resume Next
		'创建分类目录
		'Response.Flush
		enchiasp.CreatPathEx (HtmlFilePath)
		For CurrentPage = n To TotalPageNum
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			ListContent = ""
			j = (CurrentPage - 1) * maxperpage + 1

			If Perownum > 1 Then 
				ListContent = enchiasp.HtmlSetting(9)
				w = FormatPercent(100 / Perownum / 100,0)
			End If
			
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.end
				If Perownum > 1 Then
					ListContent = ListContent & "<tr valign=""top"">" & vbCrLf
					For ii = 1 To Perownum
						ListContent = ListContent & "<td width=""" & w & """ class=""softlistrow"">"
						If Not Rs.EOF Then
							Call LoadListDetail
							Rs.movenext
							i = i + 1
							j = j + 1
						End If
						ListContent = ListContent & "</td>" & vbCrLf
					Next
					ListContent = ListContent & "</tr>" & vbCrLf
				Else
					Call LoadListDetail
					Rs.MoveNext
					i = i + 1
					j = j + 1
				End If
				If i >= maxperpage Then Exit Do
			Loop
			Dim strHtmlFront, strHtmlPage
			strHtmlFront = enchiasp.HtmlPrefix & enchiasp.Supplemental(ClassID, 3) & "_"
			strHtmlPage = ShowHtmlPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, strHtmlFront, enchiasp.HtmlExtName, strClassName)
			HtmlTemplate = HtmlContent
			HtmlTemplate = Replace(HtmlTemplate, TempListContent, ListContent)
			HtmlTemplate = Replace(HtmlTemplate, "{$ReadListPage}", strHtmlPage)
			HtmlTemplate = Replace(HtmlTemplate, "[ShowRepetend]", "")
			HtmlTemplate = Replace(HtmlTemplate, "[/ShowRepetend]", "")
			'开始生成子分类的HTML页
			HtmlFileName = HtmlFilePath & ReadListPageName(ClassID, CurrentPage)
			enchiasp.CreatedTextFile HtmlFileName, HtmlTemplate
			If IsShowFlush = 1 Then 
				Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "列表HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
				Response.Flush
			End If
		Next
		Exit Sub
	End Sub
	'================================================
	'过程名：LoadChildListAsp
	'作  用：装载子级文章列表ASP
	'================================================
	Private Sub LoadChildListAsp()
		If IsNull(TempListContent) Then Exit Sub
		Dim Perownum,ii,w
		Perownum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(8))
		
		On Error Resume Next
		i = 0
		Rs.MoveFirst
		If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
		ListContent = ""
		j = (CurrentPage - 1) * maxperpage + 1
		If Perownum > 1 Then 
			ListContent = enchiasp.HtmlSetting(9)
			w = FormatPercent(100 / Perownum / 100,0)
		End If
		
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.end
			If Perownum > 1 Then
				ListContent = ListContent & "<tr valign=""top"">" & vbCrLf
				For ii = 1 To Perownum
					ListContent = ListContent & "<td width=""" & w & """ class=""softlistrow"">"
					If Not Rs.EOF Then
						Call LoadListDetail
						Rs.movenext
						i = i + 1
						j = j + 1
					End If
					ListContent = ListContent & "</td>" & vbCrLf
				Next
				ListContent = ListContent & "</tr>" & vbCrLf
			Else
				Call LoadListDetail
				Rs.MoveNext
				i = i + 1
				j = j + 1
			End If
			If i >= maxperpage Then Exit Do
		Loop
		Dim strPagination
		strPagination = ShowListPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, ASPCurrentPage(PageType), strClassName)
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
		HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strPagination)



	End Sub
	'================================================
	'过程名：LoadArticleList
	'作  用：装载文章列表
	'================================================
	Private Function LoadArticleList(ArticleID, ClassID, title, ColorMode, FontMode, BriefTopic, ClassName, Content, HtmlFileDir, HtmlFileDate, AllHits, UserName, star, isBest)
		On Error Resume Next
	End Function
	'================================================
	'过程名：LoadListDetail
	'作  用：装载子级文章列表细节
	'================================================
	Private Sub LoadListDetail()
		Dim sTitle, sTopic, ArticleTitle, ListStyle
		Dim ArticleContent, ArticleUrl, WriteTime, sClassName,imageurl
		
		On Error Resume Next
		ListContent = ListContent & TempListContent
		If (i Mod 2) = 0 Then
			ListStyle = 1
		Else
			ListStyle = 2
		End If
		If strlen > 0 Then
			sTitle = enchiasp.GotTopic(Rs("title"),strlen)
		Else
			sTitle = Rs("title")
		End If
		sTitle = enchiasp.ReadFontMode(sTitle, Rs("ColorMode"), Rs("FontMode"))
		sTopic = enchiasp.ReadPicTopic(Rs("BriefTopic"))
		If CInt(CreateHtml) <> 0 Then
			ArticleUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath) & ReadPagination(0)
			sClassName = ChannelRootDir & Rs("HtmlFileDir")
		Else
			ArticleUrl = ChannelRootDir & "show.asp?id=" & Rs("ArticleID")
			sClassName = ChannelRootDir & "list.asp?classid=" & Rs("ClassID")
		End If
		sClassName = "<a href='" & sClassName & "' title='" & Rs("ClassName") & "'>" & Rs("ClassName") & "</a>"
		ArticleTitle = "<a href='" & ArticleUrl & "' title='" & Rs("title") & "' class=showtopic>" & sTitle & "</a>"
		imageurl=  Rs("imageurl")
                ArticleContent = enchiasp.CutString(Rs("Content"), CInt(enchiasp.HtmlSetting(3)))
		WriteTime = enchiasp.ShowDateTime(Rs("WriteTime"), CInt(enchiasp.HtmlSetting(2)))
		ListContent = Replace(ListContent, "{$ClassifyName}", sClassName)
		HtmlContent = Replace(HtmlContent, "{$ClassName}", strClassName)
		ListContent = Replace(ListContent, "{$ArticleTitle}", ArticleTitle)
		ListContent = Replace(ListContent, "{$ArticleTopic}", sTitle)
		ListContent = Replace(ListContent, "{$ArticleUrl}", ArticleUrl)
		ListContent = Replace(ListContent, "{$BriefTopic}", sTopic)
		ListContent = Replace(ListContent, "{$ArticleHits}", Rs("AllHits"))
		ListContent = Replace(ListContent, "{$UserName}", Rs("username"))
		ListContent = Replace(ListContent, "{$Star}", Rs("star"))
		ListContent = Replace(ListContent, "{$IsBest}", Rs("isBest"))
		ListContent = Replace(ListContent, "{$IsTop}", Rs("isTop"))
		ListContent = Replace(ListContent, "{$ArticleDateTime}", WriteTime)
		ListContent = Replace(ListContent, "{$ArticleContent}", ArticleContent)
		ListContent = Replace(ListContent, "{$ListStyle}", ListStyle)
		ListContent = Replace(ListContent, "{$Order}", j)
		ListContent = Replace(ListContent, "{$PageID}", CurrentPage)
               ListContent = Replace(ListContent, "{$ArticlePicture}",imageurl)
	End Sub

	Public Function ASPCurrentPage(stype)
		Dim CurrentUrl
		Select Case stype
			Case "1"
				CurrentUrl = "&amp;classid=" & Trim(Request("classid")) & "&amp;order=" & Trim(Request("order"))
			Case "2"
				CurrentUrl = "&amp;sid=" & Trim(Request("sid"))
			Case "3", "4", "5"
				CurrentUrl = ""
			Case Else
				If Trim(Request("word")) <> "" Then
					CurrentUrl = "&amp;word=" & Trim(Request("word"))
				Else
					CurrentUrl = "&amp;act=" & Trim(Request("act")) & "&amp;classid=" & Trim(Request("classid")) & "&amp;keyword=" & Trim(Request("keyword"))
				End If
		End Select
		ASPCurrentPage = CurrentUrl
	End Function

	Private Function ReadListPageName(ClassID, CurrentPage)
		ReadListPageName = enchiasp.ClassFileName(ClassID, enchiasp.HtmlExtName, enchiasp.HtmlPrefix, CurrentPage)
	End Function
	'##############################################################################
	'#############################\\执行专题文章开始//#############################
	Public Sub ShowArticleSpecial()
		On Error Resume Next
		If CreateHtml <> 0 Then
			Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
			Exit Sub
		Else
			enchiasp.PreventInfuse
			If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
				Response.Write ("错误的系统参数!请输入整数")
				Response.end
			End If
			If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
				CurrentPage = CLng(Request("page"))
			Else
				CurrentPage = 1
			End If
			SpecialID = enchiasp.ChkNumeric(Request("sid"))
			Response.Write CreateArticleSpecial(SpecialID, 1)
		End If
	End Sub
	Public Function CreateArticleSpecial(sid, n)
		On Error Resume Next
		Dim rsPecial
		Dim HtmlFileName
		
		PageType = 2
		If Not IsNumeric(SpecialID) Then Exit Function
		Set rsPecial = enchiasp.Execute("select SpecialID,SpecialName,SpecialDir from [ECCMS_Special] where ChannelID = " & ChannelID & " And SpecialID=" & sid)
		If rsPecial.BOF And rsPecial.EOF Then
			Response.Write ("错误的系统参数!")
			Set rsPecial = Nothing
			Exit Function
		Else
			SpecialName = rsPecial("SpecialName")
			SpecialID = rsPecial("SpecialID")
			SpecialDir = rsPecial("SpecialDir")
			skinid = CLng(enchiasp.ChannelSkin)
		End If
		rsPecial.Close: Set rsPecial = Nothing
		enchiasp.LoadTemplates ChannelID, 4, skinid
		If CreateHtml <> 0 Then
			HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "special/" & SpecialDir & "/"
			enchiasp.CreatPathEx (HtmlFilePath)
		End If
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$SpecialID}", SpecialID)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", SpecialName)
		HtmlContent = Replace(HtmlContent, "{$ArticleIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$SpecialName}", SpecialName)
		Call ReplaceString
		
		maxperpage = CInt(enchiasp.HtmlSetting(1))
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		'记录总数
		TotalNumber = enchiasp.Execute("Select Count(ArticleID) from ECCMS_Article where ChannelID = " & ChannelID & " And isAccept > 0 And SpecialID = " & SpecialID)(0)
		TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "select A.ArticleID,A.ClassID,A.BriefTopic,A.ColorMode,A.FontMode,A.title,A.content,A.Related,A.Author,A.ComeFrom,A.isTop,A.username,A.star,A.isBest,A.WriteTime,A.Allhits,A.HtmlFileDate,C.ClassName,C.HtmlFileDir from [ECCMS_Article] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.SpecialID = " & SpecialID & " order by A.isTop desc, A.WriteTime desc ,A.ArticleID desc"
		Rs.Open SQL, Conn, 1, 1
		If Rs.BOF And Rs.EOF Then
			'如果没有找到相关内容,清除掉无用的标签代码
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "还没有找到任何专题" & enchiasp.ModuleName & "")
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
			If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成专题" & enchiasp.ModuleName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
			Response.Flush
			'如果是生成HTML,执行下面的语句
			If CreateHtml <> 0 Then
				HtmlFileName = HtmlFilePath & enchiasp.SpecialFileName(SpecialID, enchiasp.HtmlExtName, 1)
				enchiasp.CreatedTextFile HtmlFileName, HtmlContent
				If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成专题" & enchiasp.ModuleName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
				Response.Flush
			End If
		Else
			'获取模板标签[ShowRepetend][/ReadArticleList]中的字符串
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			If CreateHtml <> 0 Then
				Call LoadArticleListHtml(n)
			Else
				Call LoadChildListAsp
			End If
		End If
		Rs.Close: Set Rs = Nothing
		If CreateHtml = 0 Then CreateArticleSpecial = HtmlContent
		Exit Function
	End Function
	'================================================
	'过程名：LoadArticleListHtml
	'作  用：装载文章列表并生成HTML
	'================================================
	Private Sub LoadArticleListHtml(n)
		Dim HtmlFileName, strFlush
		
		If IsNull(TempListContent) Then Exit Sub
		On Error Resume Next
		For CurrentPage = n To TotalPageNum
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			ListContent = ""
			j = (CurrentPage - 1) * maxperpage + 1
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.end
				Call LoadListDetail
				Rs.MoveNext
				i = i + 1
				j = j + 1
				If i >= maxperpage Then Exit Do
			Loop
			Dim strHtmlFront, strHtmlPage
			strHtmlFront = "Special" & enchiasp.Supplemental(SpecialID, 3) & "_"
			strHtmlPage = ShowHtmlPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, strHtmlFront, enchiasp.HtmlExtName, SpecialName)
			HtmlTemplate = HtmlContent
			HtmlTemplate = Replace(HtmlTemplate, TempListContent, ListContent)
			HtmlTemplate = Replace(HtmlTemplate, "{$ReadListPage}", strHtmlPage)
			HtmlTemplate = Replace(HtmlTemplate, "[ShowRepetend]", "")
			HtmlTemplate = Replace(HtmlTemplate, "[/ShowRepetend]", "")
			'开始生成子分类的HTML页
			HtmlFileName = HtmlFilePath & enchiasp.SpecialFileName(SpecialID, enchiasp.HtmlExtName, CurrentPage)
			enchiasp.CreatedTextFile HtmlFileName, HtmlTemplate
			If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成专题" & enchiasp.ModuleName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
			Response.Flush
		Next
		Exit Sub
	End Sub
	'================================================
	'过程名：ReplaceString
	'作  用：替换模板内容
	'================================================
	Private Sub ReplaceString()
		HtmlContent = HTML.ReadCurrentStation(HtmlContent, ChannelID, ClassID, strClassName, ParentID, strParent, strFileDir)
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		HtmlContent = HTML.ReadAnnounceList(HtmlContent)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = HTML.ReadArticlePic(HtmlContent)
		HtmlContent = HTML.ReadArticleList(HtmlContent)
		HtmlContent = HTML.ReadSoftPic(HtmlContent)
		HtmlContent = HTML.ReadSoftList(HtmlContent)
		HtmlContent = HTML.ReadNewsPicAndText(HtmlContent)
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
                HtmlContent = HTML.ReadStatistic(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
	End Sub
	'##############################################################################
	'#############################\\执行推荐文章开始//#############################
	'================================================
	'过程名：ShowBestArticle
	'作  用：显示推荐文章
	'================================================
	Public Sub ShowBestArticle()
		On Error Resume Next
		If CreateHtml <> 0 Then
			Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
			Exit Sub
		Else
			enchiasp.PreventInfuse
			If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
				Response.Write ("错误的系统参数!请输入整数")
				Response.end
			End If
			If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
				CurrentPage = CLng(Request("page"))
			Else
				CurrentPage = 1
			End If
			Response.Write CreateBestArticle(1)
		End If
	End Sub
	'================================================
	'过程名：ShowNewArticle
	'作  用：显示最新文章
	'================================================
	Public Sub ShowNewArticle()
		On Error Resume Next
		If CreateHtml <> 0 Then
			Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
			Exit Sub
		Else
			enchiasp.PreventInfuse
			If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
				Response.Write ("错误的系统参数!请输入整数")
				Response.end
			End If
			If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
				CurrentPage = CLng(Request("page"))
			Else
				CurrentPage = 1
			End If
			Response.Write CreateBestArticle(0)
		End If
	End Sub
	'================================================
	'过程名：NewBestArticleList
	'作  用：最新推荐文章列表
	'================================================
	Public Function CreateBestArticle(t)
		On Error Resume Next
		Dim HtmlFileName, SQL1, SQL2
		
		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 5, skinid
		HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "special/"
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ArticleIndex}", strIndexName)
		If t = 1 Then
			strClassName = "推荐" & enchiasp.ModuleName
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", "推荐" & enchiasp.ModuleName)
			PageType = 3
			SQL1 = "And IsBest > 0"
			SQL2 = "And A.IsBest > 0"
		Else
			strClassName = "最新" & enchiasp.ModuleName
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", "最新" & enchiasp.ModuleName)
			PageType = 5
			SQL1 = ""
			SQL2 = ""
		End If
		Call ReplaceString
		maxperpage = CInt(enchiasp.HtmlSetting(1))
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		'记录总数
		TotalNumber = enchiasp.Execute("Select Count(ArticleID) from ECCMS_Article where ChannelID = " & ChannelID & " And isAccept > 0 " & SQL1 & "")(0)
		If TotalNumber >= CLng(enchiasp.HtmlSetting(4)) Then TotalNumber = CLng(enchiasp.HtmlSetting(4))
		TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "select top " & CLng(enchiasp.HtmlSetting(4)) & " A.ArticleID,A.ClassID,A.BriefTopic,A.ColorMode,A.FontMode,A.title,A.content,A.Related,A.Author,A.ComeFrom,A.isTop,A.username,A.star,A.isBest,A.WriteTime,A.Allhits,A.HtmlFileDate,C.ClassName,C.HtmlFileDir from [ECCMS_Article] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 " & SQL2 & " order by A.WriteTime desc ,A.ArticleID desc"
		Rs.Open SQL, Conn, 1, 1
		If Rs.BOF And Rs.EOF Then
			'如果没有找到相关内容,清除掉无用的标签代码
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "还没有找到任何推荐" & enchiasp.ModuleName & "")
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
			'如果是生成HTML,执行下面的语句
			If CreateHtml <> 0 Then
				If t = 1 Then
					HtmlFileName = HtmlFilePath & enchiasp.HtmlPrefix & "Best001" & enchiasp.HtmlExtName
				Else
					HtmlFileName = HtmlFilePath & enchiasp.HtmlPrefix & "New001" & enchiasp.HtmlExtName
				End If
				enchiasp.CreatedTextFile HtmlFileName, HtmlContent
				If IsShowFlush = 1 Then 
					Response.Write "<li style=""font-size: 12px;"">生成" & strClassName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
					Response.Flush
				End If
			End If
		Else
			'获取模板标签[ShowRepetend][/ReadArticleList]中的字符串
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			If CreateHtml <> 0 Then
				Call LoadBestArticleListHtml(t)
			Else
				Call LoadChildListAsp
			End If
		End If
		Rs.Close: Set Rs = Nothing
		If CreateHtml = 0 Then Response.Write HtmlContent
		Exit Function
	End Function
	'================================================
	'过程名：LoadBestArticleListHtml
	'作  用：装载文章列表并生成HTML
	'================================================
	Private Sub LoadBestArticleListHtml(t)
		Dim HtmlFileName, sulCurrentPage
		
		If IsNull(TempListContent) Then Exit Sub
		On Error Resume Next
		For CurrentPage = 1 To TotalPageNum
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			ListContent = ""
			j = (CurrentPage - 1) * maxperpage + 1
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.end
				Call LoadListDetail
				Rs.MoveNext
				i = i + 1
				j = j + 1
				If i >= maxperpage Then Exit Do
			Loop
			Dim strHtmlFront, strHtmlPage
			If t = 1 Then
				sulCurrentPage = enchiasp.HtmlPrefix & "Best"
			Else
				sulCurrentPage = enchiasp.HtmlPrefix & "New"
			End If
			strHtmlFront = sulCurrentPage
			strHtmlPage = ShowHtmlPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, strHtmlFront, enchiasp.HtmlExtName, strClassName)
			HtmlTemplate = HtmlContent
			HtmlTemplate = Replace(HtmlTemplate, TempListContent, ListContent)
			HtmlTemplate = Replace(HtmlTemplate, "{$ReadListPage}", strHtmlPage)
			HtmlTemplate = Replace(HtmlTemplate, "[ShowRepetend]", "")
			HtmlTemplate = Replace(HtmlTemplate, "[/ShowRepetend]", "")
			'开始生成子分类的HTML页
			sulCurrentPage = sulCurrentPage & enchiasp.Supplemental(CurrentPage, 3)
			HtmlFileName = HtmlFilePath & sulCurrentPage & enchiasp.HtmlExtName
			enchiasp.CreatedTextFile HtmlFileName, HtmlTemplate
			If IsShowFlush = 1 Then 
				Response.Write "<li style=""font-size: 12px;"">生成" & strClassName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
				Response.Flush
			End If
		Next
		Exit Sub
	End Sub
	'##############################################################################
	'#############################\\执行热门文章开始//#############################
	'================================================
	'过程名：ShowNewArticle
	'作  用：显示最新文章
	'================================================
	Public Sub ShowHotArticle()
		On Error Resume Next
		If CreateHtml <> 0 Then
			Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
			Exit Sub
		Else
			enchiasp.PreventInfuse
			If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
				Response.Write ("错误的系统参数!请输入整数")
				Response.end
			End If
			If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
				CurrentPage = CLng(Request("page"))
			Else
				CurrentPage = 1
			End If
			Response.Write CreateHotArticle()
		End If
	End Sub
	Public Function CreateHotArticle()
		On Error Resume Next
		Dim HtmlFileName
		
		PageType = 4
		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 6, skinid
		HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "special/"
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ArticleIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", "阅览排行")
		Call ReplaceString
		strClassName = "阅览排行"
		maxperpage = CInt(enchiasp.HtmlSetting(1))
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		'记录总数
		TotalNumber = enchiasp.Execute("SELECT COUNT(ArticleID) FROM ECCMS_Article WHERE ChannelID = " & ChannelID & " And isAccept > 0 And AllHits > " & CLng(enchiasp.LeastHotHist) & "")(0)
		If TotalNumber >= CLng(enchiasp.HtmlSetting(4)) Then TotalNumber = CLng(enchiasp.HtmlSetting(4))
		TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT TOP " & CLng(enchiasp.HtmlSetting(4)) & " A.ArticleID,A.ClassID,A.BriefTopic,A.ColorMode,A.FontMode,A.title,A.content,A.Related,A.Author,A.ComeFrom,A.isTop,A.username,A.star,A.isBest,A.WriteTime,A.Allhits,A.HtmlFileDate,C.ClassName,C.HtmlFileDir FROM [ECCMS_Article] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.AllHits > " & CLng(enchiasp.LeastHotHist) & " ORDER BY A.Allhits DESC, A.WriteTime DESC ,A.ArticleID DESC"
		Rs.Open SQL, Conn, 1, 1
		If Rs.BOF And Rs.EOF Then
			'如果没有找到相关内容,清除掉无用的标签代码
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "还没有找到任何热门" & enchiasp.ModuleName & "")
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
			'如果是生成HTML,执行下面的语句
			If CreateHtml <> 0 Then
				HtmlFileName = HtmlFilePath & enchiasp.HtmlPrefix & "Hot001" & enchiasp.HtmlExtName
				enchiasp.CreatedTextFile HtmlFileName, HtmlContent
				If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成热门" & enchiasp.ModuleName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
				Response.Flush
			End If
		Else
			'获取模板标签[ShowRepetend][/ReadArticleList]中的字符串
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			If CreateHtml <> 0 Then
				Call LoadHotArticleListHtml
			Else
				Call LoadChildListAsp
			End If
		End If
		Rs.Close: Set Rs = Nothing
		If CreateHtml = 0 Then Response.Write HtmlContent
		Exit Function
	End Function
	'================================================
	'过程名：LoadHotArticleListHtml
	'作  用：装载文章列表并生成HTML
	'================================================
	Private Sub LoadHotArticleListHtml()
		Dim HtmlFileName, sulCurrentPage
		
		If IsNull(TempListContent) Then Exit Sub
		On Error Resume Next
		For CurrentPage = 1 To TotalPageNum
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			'Dim bookmark:bookmark = Rs.bookmark
			ListContent = ""
			j = (CurrentPage - 1) * maxperpage + 1
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.end
				Call LoadListDetail
				Rs.MoveNext
				i = i + 1
				j = j + 1
				If i >= maxperpage Then Exit Do
			Loop
			Dim strHtmlFront, strHtmlPage
			strHtmlFront = enchiasp.HtmlPrefix & "Hot"
			strHtmlPage = ShowHtmlPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, strHtmlFront, enchiasp.HtmlExtName, strClassName)
			HtmlTemplate = HtmlContent
			HtmlTemplate = Replace(HtmlTemplate, TempListContent, ListContent)
			HtmlTemplate = Replace(HtmlTemplate, "{$ReadListPage}", strHtmlPage)
			HtmlTemplate = Replace(HtmlTemplate, "[ShowRepetend]", "")
			HtmlTemplate = Replace(HtmlTemplate, "[/ShowRepetend]", "")
			'开始生成子分类的HTML页
			sulCurrentPage = enchiasp.HtmlPrefix & "Hot" & enchiasp.Supplemental(CurrentPage, 3)
			HtmlFileName = HtmlFilePath & sulCurrentPage & enchiasp.HtmlExtName
			enchiasp.CreatedTextFile HtmlFileName, HtmlTemplate
			If IsShowFlush = 1 Then 
				Response.Write "<li style=""font-size: 12px;"">生成热门" & enchiasp.ModuleName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
				Response.Flush
			End If
		Next
		Exit Sub
	End Sub
	'##########################################################################
	'#############################\\文章搜索开始//#############################
	Public Sub ShowArticleSearch()
		Dim SearchMaxPageList
		Dim Action, findword
		Dim rsClass, strNoResult
		Dim strWord, s
		
		PageType = 6
		keyword = enchiasp.ChkQueryStr(Trim(Request("keyword")))
		keyword = enchiasp.CheckInfuse(keyword,255)
		strWord = enchiasp.CheckStr(Trim(Request("word")))
		strWord = enchiasp.CheckInfuse(strWord,10)
		s = enchiasp.ChkNumeric(Request.QueryString("s"))
		
		If enchiasp.CheckNull(strWord) Then
			strWord = UCase(Left(strWord, 6))
			keyword = strWord
		Else
			strWord = ""
		End If
		
		If keyword = "" And strWord = "" Then
			Call OutAlertScript("请输入要查询的关键字！")
			Exit Sub
		End If
		
		If strWord = "" Then
			If Not enchiasp.CheckQuery(keyword) Then
				Call OutAlertScript("你查询的关键中有非法字符！\n请返回重新输入关键字查询。")
				Exit Sub
			End If
		End If
		
		
		skinid = CLng(enchiasp.ChannelSkin)
		On Error Resume Next
		enchiasp.LoadTemplates ChannelID, 7, skinid
		If enchiasp.HtmlSetting(4) <> "0" Then
			If IsNumeric(enchiasp.HtmlSetting(4)) Then
				'If CInt(enchiasp.HtmlSetting(4)) Mod CInt(enchiasp.HtmlSetting(1)) = 0 Then
					'SearchMaxPageList = CLng(enchiasp.HtmlSetting(4)) \ CInt(enchiasp.HtmlSetting(1))
				'Else
					'SearchMaxPageList = CLng(enchiasp.HtmlSetting(4)) \ CInt(enchiasp.HtmlSetting(1)) + 1
				'End If
				SearchMaxPageList = CLng(enchiasp.HtmlSetting(4))
			Else
				SearchMaxPageList = 50
			End If
		Else
			SearchMaxPageList = 50
		End If

		strNoResult = Replace(enchiasp.HtmlSetting(8), "{$KeyWord}", keyword)
		Action = enchiasp.CheckStr(Trim(Request("act")))
		Action = enchiasp.CheckStr(Action)

		If strWord = "" And LCase(Action) <> "isweb" Then
			If enchiasp.strLength(keyword) < CLng(enchiasp.HtmlSetting(5)) Or enchiasp.strLength(keyword) > CLng(enchiasp.HtmlSetting(6)) Then
				Call OutAlertScript("查询错误！\n您查询的关键字不能小于 " & enchiasp.HtmlSetting(5) & " 或者大于 " & enchiasp.HtmlSetting(6) & " 个字节。")
				Exit Sub
			End If
		End If
		
		
		If strWord = "" Then
			If LCase(Action) = "topic" Then
				findword = "A.Title like '%" & keyword & "%'"
			ElseIf LCase(Action) = "content" Then
				If CInt(enchiasp.FullContQuery) <> 0 Then
					findword = "A.Content like '%" & keyword & "%'"
				Else
					Call OutAlertScript(Replace(Replace(enchiasp.HtmlSetting(10), Chr(34), "\"""), vbCrLf, ""))
					Exit Sub
				End If
			Else
				findword = "A.Title like '%" & keyword & "%'"
			End If
		Else
			findword = "A.AlphaIndex='" & strWord & "'"
		End If
		If LCase(Action) <> "isweb" Then
			If IsEmpty(Session("QueryLimited")) Then
				Session("QueryLimited") = keyword & "|" & Action & "|" & Now()
			Else
				Dim QueryLimited
				QueryLimited = Split(Session("QueryLimited"), "|")
				If UBound(QueryLimited) = 2 Then
					If CStr(Trim(QueryLimited(0))) = CStr(keyword) And CStr(Trim(QueryLimited(1))) = CStr(Action) Then
						Session("QueryLimited") = keyword & "|" & Action & "|" & Now()
					Else
						If DateDiff("s", QueryLimited(2), Now()) < CLng(enchiasp.HtmlSetting(7)) Then
							Dim strLimited
							strLimited = Replace(enchiasp.HtmlSetting(9), "{$TimeLimited}", enchiasp.HtmlSetting(7))
							Call OutAlertScript(Replace(Replace(strLimited, Chr(34), "\"""), vbCrLf, ""))
							Exit Sub
						Else
							Session("QueryLimited") = keyword & "|" & Action & "|" & Now()
						End If
					End If
				Else
					Session("QueryLimited") = keyword & "|" & Action & "|" & Now()
				End If
			End If
		End If
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ArticleIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$KeyWord}", KeyWord)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ModuleName & "搜索")
		HtmlContent = Replace(HtmlContent, "{$QueryKeyWord}", "<font color=red><strong>" & keyword & "</strong></font>")
		Call ReplaceString
		If LCase(Action) <> "isweb" Then
			If IsNumeric(Request("classid")) And Request("classid") <> "" Then
				Set rsClass = enchiasp.Execute("select ClassID,ChildStr from [ECCMS_Classify] where ChannelID = " & ChannelID & " And ClassID=" & CLng(Request("classid")))
				If rsClass.BOF And rsClass.EOF Then
					HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strNoResult, 1, 1, 1)
					HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "")
					HtmlContent = Replace(HtmlContent, "{$totalrec}", 0)
					HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
					Set rsClass = Nothing
					Response.Write HtmlContent
					Exit Sub
				Else
					findword = "A.ClassID in (" & rsClass("ChildStr") & ") And " & findword
				End If
				rsClass.Close: Set rsClass = Nothing
			End If
			maxperpage = CInt(enchiasp.HtmlSetting(1))
			If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
				Response.Write ("错误的系统参数!请输入整数")
				Response.end
			End If
			If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
				CurrentPage = CLng(Request("page"))
			Else
				CurrentPage = 1
			End If
			If CLng(CurrentPage) = 0 Then CurrentPage = 1
			
			Set Rs = CreateObject("ADODB.Recordset")
			SQL = "select top " & SearchMaxPageList & " A.ArticleID,A.ClassID,A.BriefTopic,A.ColorMode,A.FontMode,A.title,A.content,A.Related,A.Author,A.ComeFrom,A.isTop,A.username,A.star,A.isBest,A.WriteTime,A.Allhits,A.HtmlFileDate,C.ClassName,C.HtmlFileDir from [ECCMS_Article] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And " & findword & " order by A.WriteTime desc,A.ArticleID desc"
			Rs.Open SQL, Conn, 1, 1
			If Rs.BOF And Rs.EOF Then
				'如果没有找到相关内容,清除掉无用的标签代码
				HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strNoResult)
				HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "")
				HtmlContent = Replace(HtmlContent, "{$totalrec}", 0)
				HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
			Else
				TotalNumber = Rs.RecordCount
				If (TotalNumber Mod maxperpage) = 0 Then
					TotalPageNum = TotalNumber \ maxperpage
				Else
					TotalPageNum = TotalNumber \ maxperpage + 1
				End If
				If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
				If CurrentPage < 1 Then CurrentPage = 1
				HtmlContent = Replace(HtmlContent, "{$totalrec}", TotalNumber)
				'获取模板标签[ShowRepetend][/ReadArticleList]中的字符串
				TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
				Call LoadSearchList
			End If
			Rs.Close: Set Rs = Nothing
		Else
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "")
			HtmlContent = Replace(HtmlContent, "{$totalrec}", 0)
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
			If s = 1 Then
				Call isWeb_Query()
				Exit Sub
			End If
			Response.Write HtmlContent & SearchObj
			Exit Sub
		End If
		Response.Write HtmlContent
		Exit Sub
	End Sub
	'================================================
	'过程名：LoadChildListAsp
	'作  用：装载子级文章列表ASP
	'================================================
	Private Sub LoadSearchList()
		If IsNull(TempListContent) Then Exit Sub
		On Error Resume Next
		i = 0
		Rs.MoveFirst
		If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
		j = (CurrentPage - 1) * maxperpage + 1
		ListContent = ""
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.end
			Call SearchResult
			Rs.MoveNext
			i = i + 1
			j = j + 1
			If i >= maxperpage Then Exit Do
		Loop
		Dim strPagination
		strPagination = ShowListPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, ASPCurrentPage(PageType), "搜索结果")
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
		HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strPagination)
	End Sub
	'================================================
	'过程名：SearchResult
	'作  用：装载搜索列表
	'================================================
	Private Sub SearchResult()
		Dim sTitle, sTopic, ArticleTitle, ListStyle, TitleWord
		Dim ArticleContent, ArticleUrl, WriteTime, sClassName
		
		On Error Resume Next
		ListContent = ListContent & TempListContent
		If (i Mod 2) = 0 Then
			ListStyle = 1
		Else
			ListStyle = 2
		End If
		TitleWord = Replace(Rs("title"), "" & keyword & "", "<font color=red>" & keyword & "</font>")
		sTitle = enchiasp.ReadFontMode(TitleWord, Rs("ColorMode"), Rs("FontMode"))
		sTopic = enchiasp.ReadPicTopic(Rs("BriefTopic"))
		If CInt(CreateHtml) <> 0 Then
			ArticleUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath) & ReadPagination(0)
			sClassName = ChannelRootDir & Rs("HtmlFileDir")
		Else
			ArticleUrl = ChannelRootDir & "show.asp?id=" & Rs("ArticleID")
			sClassName = ChannelRootDir & "list.asp?classid=" & Rs("ClassID")
		End If
		sClassName = "<a href='" & sClassName & "' title='" & Rs("ClassName") & "' target=""_blank""><span style=""color:" & enchiasp.MainSetting(3) & """>" & Rs("ClassName") & "</span></a>"
		ArticleTitle = "<a href='" & ArticleUrl & "' title='" & Rs("title") & "' class=showtopic target=""_blank"">" & sTitle & "</a>"
		ArticleContent = enchiasp.CutString(Rs("Content"), CInt(enchiasp.HtmlSetting(3)))
		ArticleContent = Replace(ArticleContent, "" & keyword & "", "<font color=red>" & keyword & "</font>")
		
		WriteTime = enchiasp.ShowDateTime(Rs("WriteTime"), CInt(enchiasp.HtmlSetting(2)))
		ListContent = Replace(ListContent, "{$KeyWord}", keyword)
		ListContent = Replace(ListContent, "{$totalrec}", TotalNumber)
		ListContent = Replace(ListContent, "{$ClassifyName}", sClassName)
		ListContent = Replace(ListContent, "{$ArticleTitle}", ArticleTitle)
		ListContent = Replace(ListContent, "{$ArticleTopic}", sTitle)
		ListContent = Replace(ListContent, "{$ArticleUrl}", ArticleUrl)
		ListContent = Replace(ListContent, "{$BriefTopic}", sTopic)
		ListContent = Replace(ListContent, "{$ArticleHits}", Rs("AllHits"))
		ListContent = Replace(ListContent, "{$UserName}", Rs("username"))
		ListContent = Replace(ListContent, "{$ArticleDateTime}", WriteTime)
		ListContent = Replace(ListContent, "{$ArticleContent}", ArticleContent)
		ListContent = Replace(ListContent, "{$ListStyle}", ListStyle)
		ListContent = Replace(ListContent, "{$Author}", enchiasp.ChkNull(Rs("Author")))
		ListContent = Replace(ListContent, "{$Order}", j)
	End Sub
	'================================================
	'过程名：ShowArticleComment
	'作  用：文章评论
	'================================================
	Public Sub ShowArticleComment()
		Dim ArticleTitle, HtmlFileUrl, HtmlFileName
		Dim AverageGrade, TotalComment, TempListContent
		Dim strComment, strCheckBox, strAdminComment
		enchiasp.PreventInfuse
		strCheckBox = ""
		strAdminComment = ""
		On Error Resume Next

		ArticleID = enchiasp.ChkNumeric(Request("ArticleID"))
		If ArticleID = 0 Then
			Response.Write "<Br><Br><Br>Sorry！错误的系统参数,请选择正确的连接方式。"
			Response.end
		End If
		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 8, skinid
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ArticleIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ModuleName & "评论")
		HtmlContent = Replace(HtmlContent, "{$ArticleID}", ArticleID)
		'获得文章标题
		SQL = "select Top 1 A.ArticleID,A.ClassID,A.title,A.HtmlFileDate,A.ForbidEssay,C.HtmlFileDir from [ECCMS_Article] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.ArticleID = " & ArticleID
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			Response.Write "已经没有了"
			Set Rs = Nothing
			Exit Sub
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("ArticleID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				ArticleTitle = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & Rs("title") & "</a>"
			Else
				ArticleTitle = "<a href=show.asp?id=" & Rs("ArticleID") & ">" & Rs("title") & "</a>"
			End If
			ForbidEssay = Rs("ForbidEssay")
		End If
		Rs.Close
		Set Rs = CreateObject("adodb.recordset")
		SQL = "select Count(CommentID) As TotalComment,AVG(Grade) As avgGrade from ECCMS_Comment where ChannelID=" & ChannelID & " And postid = " & ArticleID
		Set Rs = enchiasp.Execute(SQL)
		TotalComment = Rs("TotalComment")
		AverageGrade = Round(Rs("avgGrade"))
		If IsNull(AverageGrade) Then AverageGrade = 0
		Rs.Close: Set Rs = Nothing
		HtmlContent = Replace(HtmlContent, "{$ArticleTitle}", ArticleTitle)
		HtmlContent = Replace(HtmlContent, "{$TotalComment}", TotalComment)
		HtmlContent = Replace(HtmlContent, "{$AverageGrade}", AverageGrade)
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("错误的系统参数!请输入整数")
			Response.end
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CLng(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		'每页显示评论数
		maxperpage = CInt(enchiasp.PaginalNum)
		'记录总数
		TotalNumber = TotalComment
		TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "select * from ECCMS_Comment where ChannelID=" & ChannelID & " And postid = " & ArticleID & " order by postime desc,CommentID desc"
		Rs.Open SQL, Conn, 1, 1
		If Rs.BOF And Rs.EOF Then
			'如果没有找到相关内容,清除掉无用的标签代码
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "暂时无人参加评论", 1, 1, 1)
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "")
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
		Else
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			ListContent = ""
			'获取模板标签[ShowRepetend][/ReadArticleList]中的字符串
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.end
				ListContent = ListContent & TempListContent
				strComment = enchiasp.HTMLEncode(Rs("Content"))
				ListContent = Replace(ListContent, "{$CommentContent}", strComment)
				ListContent = Replace(ListContent, "{$UserName}", enchiasp.HTMLEncode(Rs("username")))
				ListContent = Replace(ListContent, "{$CommentGrade}", Rs("Grade"))
				ListContent = Replace(ListContent, "{$PostTime}", Rs("postime"))
				ListContent = Replace(ListContent, "{$PostIP}", Rs("postip"))
				If Session("AdminName") <> "" Or enchiasp.membergrade = "999" Then
					strCheckBox = "<input type='checkbox' name='selCommentID' value='" & Rs("CommentID") & "'>"
				End If
				ListContent = Replace(ListContent, "{$SelCheckBox}", strCheckBox)
				Rs.MoveNext
				i = i + 1
				If i >= maxperpage Then Exit Do
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
		HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
		If Session("AdminName") <> "" Or enchiasp.membergrade = "999" Then
			strAdminComment = "<input class=Button type=button name=chkall value='全选' onClick=""CheckAll(this.form)""><input class=Button type=button name=chksel value='反选' onClick=""ContraSel(this.form)"">" & vbNewLine
			strAdminComment = strAdminComment & "<input type=hidden name=ArticleID value='" & ArticleID & "'>" & vbNewLine
			strAdminComment = strAdminComment & "<input type=hidden name=action value='del'>" & vbNewLine
			strAdminComment = strAdminComment & "<input class=Button type=submit name=Submit2 value='删除选中的评论' onclick=""{if(confirm('您确定执行该操作吗?')){this.document.selform.submit();return true;}return false;}"">"
		End If
		HtmlContent = Replace(HtmlContent, "{$AdminComment}", strAdminComment)
		Call ShowCommentPage
		Call ReplaceString
		If enchiasp.CheckStr(LCase(Request.Form("action"))) = "del" Then
			Call CommentDel
		End If
		If enchiasp.CheckStr(LCase(Request.Form("action"))) = "save" Then
			Call CommentSave
		End If
		Response.Write HtmlContent
		Exit Sub
	End Sub
	'================================================
	'过程名：ShowCommentPage
	'作  用：文章评论分页
	'================================================
	Private Sub ShowCommentPage()
		Dim FileName, ii, n, strTemp
		
		FileName = "comment.asp"
		On Error Resume Next
		If TotalNumber Mod maxperpage = 0 Then
			n = TotalNumber \ maxperpage
		Else
			n = TotalNumber \ maxperpage + 1
		End If
		strTemp = "<table cellspacing=1 width='100%' border=0><tr><td align=center> " & vbCrLf
		If CurrentPage < 2 Then
			strTemp = strTemp & " 共有评论 <font COLOR=#FF0000>" & TotalNumber & "</font> 个&nbsp;&nbsp;首 页&nbsp;&nbsp;上一页&nbsp;&nbsp;&nbsp;"
		Else
			strTemp = strTemp & "共有评论 <font COLOR=#FF0000>" & TotalNumber & "</font> 个&nbsp;&nbsp;<a href=" & FileName & "?page=1&ArticleID=" & Request("ArticleID") & ">首 页</a>&nbsp;&nbsp;"
			strTemp = strTemp & "<a href=" & FileName & "?page=" & CurrentPage - 1 & "&ArticleID=" & Request("ArticleID") & ">上一页</a>&nbsp;&nbsp;&nbsp;"
		End If

		If n - CurrentPage < 1 Then
			strTemp = strTemp & "下一页&nbsp;&nbsp;尾 页 " & vbCrLf
		Else
			strTemp = strTemp & "<a href=" & FileName & "?page=" & (CurrentPage + 1) & "&ArticleID=" & Request("ArticleID") & ">下一页</a>"
			strTemp = strTemp & "&nbsp;&nbsp;<a href=" & FileName & "?page=" & n & "&ArticleID=" & Request("ArticleID") & ">尾 页</a>" & vbCrLf
		End If
		strTemp = strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
		strTemp = strTemp & "&nbsp;<b>" & maxperpage & "</b>个/页 " & vbCrLf
		strTemp = strTemp & "</td></tr></table>" & vbCrLf
		HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strTemp)
	End Sub
	'================================================
	'过程名：CommentDel
	'作  用：文章评论删除
	'================================================
	Private Sub CommentDel()
		Dim selCommentID
		If enchiasp.CheckPost = False Then
			Call OutAlertScript("您提交的数据不合法，请不要从外部提交表单。")
			Exit Sub
		End If
		If Not IsEmpty(Request.Form("selCommentID")) Then
			selCommentID = enchiasp.CheckStr(Request("selCommentID"))
			If Session("AdminName") <> "" Or enchiasp.membergrade = "999" Then
				enchiasp.Execute ("delete from ECCMS_Comment where ChannelID=" & ChannelID & " And CommentID in (" & selCommentID & ")")
				Call OutHintScript("评论删除成功！")
			Else
				Call OutAlertScript("非法操作！你没有删除评论的权限。")
				Exit Sub
			End If
		End If
	End Sub
	'================================================
	'过程名：CommentSave
	'作  用：文章评论添加保存
	'================================================
	Public Sub CommentSave()
		If enchiasp.CheckPost = False Then
			FoundErr = True
			Call OutAlertScript("您提交的数据不合法，请不要从外部提交表单。")
			Exit Sub
		End If
		On Error Resume Next
		If CInt(enchiasp.AppearGrade) <> 0 And Session("AdminName") = "" Then
			If CInt(enchiasp.AppearGrade) > CInt(enchiasp.membergrade) Then
				FoundErr = True
				Call OutAlertScript("您没有发表评论的权限，如果您是会员请登陆后再参与评论。")
				Exit Sub
			End If
		End If
		If ForbidEssay <> 0 Then
			FoundErr = True
			Call OutAlertScript("此篇" & enchiasp.ModuleName & "禁止发表评论！")
			Exit Sub
		End If
		If Trim(Request.Form("UserName")) = "" Then
			FoundErr = True
			Call OutAlertScript("用户名不能为空！")
			Exit Sub
		End If
		If Len(Trim(Request.Form("UserName"))) > 15 Then
			FoundErr = True
			Call OutAlertScript("用户名不能大于15个字符！")
			Exit Sub
		End If
		If enchiasp.IsValidStr(Request.Form("UserName")) = False Then
			FoundErr = True
			Call OutAlertScript("用户名中有非法字符！")
			Exit Sub
		End If
		If enchiasp.strLength(Request.Form("content")) < enchiasp.LeastString Then
			FoundErr = True
			Call OutAlertScript("评论内容不能小于" & enchiasp.LeastString & "字符！")
			Exit Sub
		End If
		If enchiasp.strLength(Request.Form("content")) > enchiasp.MaxString Then
			FoundErr = True
			Call OutAlertScript("评论内容不能大于" & enchiasp.MaxString & "字符！")
			Exit Sub
		End If
		Call PreventRefresh
		If FoundErr = True Then Exit Sub
		ArticleID = enchiasp.ChkNumeric(Request.Form("ArticleID"))
		Set Rs = CreateObject("ADODB.RecordSet")
		SQL = "select * from ECCMS_Comment where (CommentID is null)"
		Rs.Open SQL, Conn, 1, 3
		Rs.AddNew
			Rs("ChannelID") = ChannelID
			Rs("postid") = ArticleID
			Rs("UserName") = enchiasp.ChkFormStr(Request.Form("UserName"))
			Rs("Grade") = enchiasp.ChkNumeric(Request.Form("Grade"))
			Rs("content") = Server.HTMLEncode(Request.Form("content"))
			Rs("postime") = Now()
			Rs("postip") = enchiasp.GetUserip
		Rs.Update
		Rs.Close: Set Rs = Nothing
		If CreateHtml <> 0 Then CreateArticleContent (ArticleID)
		Session("UserRefreshTime") = Now()
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
		Exit Sub
	End Sub
	Public Sub PreventRefresh()
		Dim RefreshTime
		RefreshTime = 20
		If DateDiff("s", Session("UserRefreshTime"), Now()) < RefreshTime Then
			FoundErr = True
			Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=" & RefreshTime & "><br>本页面起用了防刷新机制，请不要在" & RefreshTime & "秒内连续刷新本页面<BR>正在打开页面，请稍后……"
			Response.end
		End If
	End Sub

End Class
%>