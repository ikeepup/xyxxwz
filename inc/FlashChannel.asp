<!--#include file="ubbcode.asp"-->
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
Set enchicms = New FlashChannel_Cls

Class FlashChannel_Cls
	
	Private ChannelID, CreateHtml, IsShowFlush
	Private Rs,SQL,ChannelRootDir,HtmlContent,strIndexName,HtmlFilePath
	private flashid,classid,skinid,strInstallDir
	Private strFileDir, ParentID, strParent, strClassName, ChildStr, Child
	Private maxperpage, TotalNumber, TotalPageNum, CurrentPage, i,j
	private ForbidEssay,ListContent,HtmlTemplate,TempListContent
	Private FoundErr,PageType,keyword,strlen

	Public Property Let Channel(ChanID)
		ChannelID = ChanID
	End Property
	Public Property Let ShowFlush(para)
		IsShowFlush = para
	End Property

	Private Sub Class_Initialize()
		On Error Resume Next
		ChannelID = 5
		PageType = 0
		FoundErr = False
		strlen = 0
	End Sub

	Private Sub Class_Terminate()
		Set HTML = Nothing
	End Sub

	Public Sub MainChannel()
		enchiasp.ReadChannel(ChannelID)
		CreateHtml = CInt(enchiasp.IsCreateHtml)
		ChannelRootDir = enchiasp.InstallDir & enchiasp.ChannelDir
		strInstallDir = enchiasp.InstallDir
		strIndexName = "<a href=""" & ChannelRootDir & """>" & enchiasp.ChannelName & "</a>"
		
	End Sub
	'=================================================
	'过程名：BuildFlashIndex
	'作  用：显示FLASH首页
	'=================================================
	Public Sub BuildFlashIndex()
		On Error Resume Next
		LoadFlashIndex
		If CreateHtml <> 0 Then
			Response.Write "<meta http-equiv=refresh content=0;url=index" & enchiasp.HtmlExtName & ">"
		Else
			Response.Write HtmlContent
		End If
	End Sub
	'=================================================
	'过程名：CreateFlashIndex
	'作  用：生成动画首页的HTML
	'=================================================
	Public Sub CreateFlashIndex()
		On Error Resume Next
		LoadFlashIndex
		Dim FilePath
		FilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "index" & enchiasp.HtmlExtName
		enchiasp.CreatedTextFile FilePath, HtmlContent
		If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "首页HTML完成... <a href=" & FilePath & " target=_blank>" & Server.MapPath(FilePath) & "</a></li>" & vbNewLine
		Response.Flush
	End Sub
	Private Sub LoadFlashIndex()
		On Error Resume Next
		enchiasp.LoadTemplates ChannelID, 1, enchiasp.ChannelSkin
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent,"{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent,"{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent,"{$PageTitle}", enchiasp.ChannelName)
		HtmlContent = Replace(HtmlContent,"{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent,"{$FlashIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$IndexTitle}", strIndexName)
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent,ChannelID)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = HTML.ReadArticlePic(HtmlContent)
		HtmlContent = HTML.ReadSoftPic(HtmlContent)
		HtmlContent = HTML.ReadSoftList(HtmlContent)
		HtmlContent = HTML.ReadArticleList(HtmlContent)
		HtmlContent = HTML.ReadFlashList(HtmlContent)
		HtmlContent = HTML.ReadFlashPic(HtmlContent)
		HtmlContent = HTML.ReadFlashList(HtmlContent)
		HtmlContent = HTML.ReadFlashPic(HtmlContent)
		HtmlContent = HTML.ReadFriendLink(HtmlContent)
		HtmlContent = HTML.ReadGuestList(HtmlContent)
		HtmlContent = HTML.ReadAnnounceList(HtmlContent)
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
		HtmlContent = HTML.ReadPopularSoft(HtmlContent)
		HtmlContent = HTML.ReadPopularFlash(HtmlContent)
		HtmlContent = HTML.ReadStatistic(HtmlContent)
		HtmlContent = HTML.ReadUserRank(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
		HtmlContent = HtmlContent
	End Sub
	'#############################\\动画信息开始//#############################
	'=================================================
	'过程名：BuildFlashInfo
	'作  用：显示动画详细页面
	'=================================================
	Public Sub BuildFlashInfo()
		If CreateHtml <> 0 Then
			Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
			Exit Sub
		Else
			enchiasp.PreventInfuse
			flashid = enchiasp.ChkNumeric(Request("id"))
			Response.Write LoadFlashInfo(flashid)
		End If
	End Sub
	
	Public Function LoadFlashInfo(flashid)
		On Error Resume Next
		SQL = "SELECT A.*,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,C.UseHtml FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.flashid=" & flashid
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			LoadFlashInfo = ""
			If CreateHtml = 0 Then
				Response.Write "<meta http-equiv=""refresh"" content=""2;url='/"">" & vbNewLine
				Response.Write "<p align=""center"" style=""font-size: 16px;color: red;"">对不起，该页面发生了错误，无法访问! 系统两秒后自动转到网站首页......</p>" & vbNewLine
			End If
			Set Rs = Nothing
			Exit Function
		End If

		If Rs("skinid") <> 0 Then
			skinid = Rs("skinid")
		Else
			skinid = enchiasp.ChannelSkin
		End If
		
		enchiasp.LoadTemplates ChannelID, 3, skinid
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$Best}", Rs("isBest"))
		HtmlContent = Replace(HtmlContent, "{$Star}", enchiasp.ChkNumeric(Rs("star")))
		HtmlContent = Replace(HtmlContent, "{$DateAndTime}", Rs("addTime"))
		HtmlContent = Replace(HtmlContent, "{$ClassName}", Rs("ClassName"))
		HtmlContent = Replace(HtmlContent, "{$Author}", enchiasp.ChkNull(Rs("Author")))
		HtmlContent = Replace(HtmlContent, "{$Describe}", enchiasp.ChkNull(Rs("Describe")))
		HtmlContent = Replace(HtmlContent, "{$UserName}", Rs("UserName"))
		HtmlContent = Replace(HtmlContent, "{$Grade}", Rs("grade"))
		HtmlContent = Replace(HtmlContent, "{$IsTop}", Rs("IsTop"))
		HtmlContent = Replace(HtmlContent, "{$FileSize}", ReadFilesize(Rs("filesize")))
		HtmlContent = Replace(HtmlContent, "{$ComeFrom}", ReadComeFrom(Rs("ComeFrom")))
		HtmlContent = Replace(HtmlContent, "{$Introduce}", UbbCode(Rs("Introduce")))
		HtmlContent = Replace(HtmlContent, "{$Display}", PreviewMode(Rs("showurl"),Rs("showmode")))
		HtmlContent = Replace(HtmlContent, "{$ShowThisUrl}", enchiasp.ChkNull(Rs("showurl")))
		HtmlContent = Replace(HtmlContent, "{$ShowFullUrl}", FormatShowUrl(Rs("showurl")))
		
		If InStr(HtmlContent, "{$BackFlash}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$BackFlash}", BackFlash(flashid))
		End If
		If InStr(HtmlContent, "{$NextFlash}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$NextFlash}", NextFlash(flashid))
		End If
		If InStr(HtmlContent, "{$FlashComment}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$FlashComment}", FlashComment(Rs("flashid")))
		End If
		If InStr(HtmlContent, "{$RelatedFlash}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$RelatedFlash}", RelatedFlash(enchiasp.ChkNull(Rs("Related")), Rs("title"), Rs("flashid")))
		End If
		
		HtmlContent = Replace(HtmlContent, "{$ShowUrl}", enchiasp.ChkNull(Rs("showurl")))
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", Rs("title"))
		HtmlContent = Replace(HtmlContent, "{$ClassID}", Rs("ClassID"))
		HtmlContent = Replace(HtmlContent, "{$FlashTitle}", Rs("title"))
		HtmlContent = Replace(HtmlContent, "{$FlashID}", Rs("flashid"))
		HtmlContent = HTML.ReadCurrentStation(HtmlContent, ChannelID, Rs("ClassID"), Rs("ClassName"), Rs("ParentID"), Rs("ParentStr"), Rs("HtmlFileDir"))
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = HTML.ReadFlashPic(HtmlContent)
		HtmlContent = HTML.ReadFlashList(HtmlContent)
		HtmlContent = HTML.ReadPopularFlash(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		If CreateHtml <> 0 Then
			Call CreateFlashInfo
		Else
			LoadFlashInfo = HtmlContent
		End If
		Rs.Close: Set Rs = Nothing
	End Function
	'=================================================
	'过程名：CreateFlashInfo
	'作  用：生成FLASH信息HTML
	'=================================================
	Private Sub CreateFlashInfo()
		Dim HtmlFileName
		HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath)
		enchiasp.CreatPathEx (HtmlFilePath)
		HtmlFileName = HtmlFilePath & enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("flashid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, CurrentPage)
		enchiasp.CreatedTextFile HtmlFileName, HtmlContent
		If IsShowFlush = 1 Then 
			Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "信息HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
			Response.Flush
		End If
	End Sub
	'=================================================
	'函数名：BackFlash
	'作  用：显示上一动画
	'=================================================
	Private Function BackFlash(flashid)
		Dim rsContext, SQL, HtmlFileUrl, HtmlFileName
		On Error Resume Next
		SQL = "SELECT TOP 1 A.flashid,A.ClassID,A.title,A.HtmlFileDate,C.HtmlFileDir FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.flashid < " & flashid & " ORDER BY A.flashid DESC"
		Set rsContext = enchiasp.Execute(SQL)
		If rsContext.EOF And rsContext.BOF Then
			HtmlContent = Replace(HtmlContent, "{$BackUrl}", "")
			BackFlash = "已经没有了"
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & rsContext("HtmlFileDir") & enchiasp.ShowDatePath(rsContext("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(rsContext("HtmlFileDate"), rsContext("flashid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				HtmlContent = Replace(HtmlContent, "{$BackUrl}", HtmlFileUrl & HtmlFileName)
				BackFlash = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & rsContext("title") & "</a>"
			Else
				HtmlContent = Replace(HtmlContent, "{$BackUrl}", "?id=" & rsContext("flashid"))
				BackFlash = "<a href=?id=" & rsContext("flashid") & ">" & rsContext("title") & "</a>"
			End If
		End If
		rsContext.Close
		Set rsContext = Nothing
	End Function
	'=================================================
	'函数名：NextFlash
	'作  用：显示下一动画
	'=================================================
	Private Function NextFlash(flashid)
		Dim rsContext, SQL, HtmlFileUrl, HtmlFileName
		On Error Resume Next
		SQL = "SELECT TOP 1 A.flashid,A.ClassID,A.title,A.HtmlFileDate,C.HtmlFileDir FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.flashid > " & flashid & " ORDER BY A.flashid ASC"
		Set rsContext = enchiasp.Execute(SQL)
		If rsContext.EOF And rsContext.BOF Then
			NextFlash = "已经没有了"
			HtmlContent = Replace(HtmlContent, "{$NextUrl}", "")
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & rsContext("HtmlFileDir") & enchiasp.ShowDatePath(rsContext("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(rsContext("HtmlFileDate"), rsContext("flashid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				HtmlContent = Replace(HtmlContent, "{$NextUrl}", HtmlFileUrl & HtmlFileName)
				NextFlash = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & rsContext("title") & "</a>"
			Else
				HtmlContent = Replace(HtmlContent, "{$NextUrl}", "?id=" & rsContext("flashid"))
				NextFlash = "<a href=?id=" & rsContext("flashid") & ">" & rsContext("title") & "</a>"
			End If
		End If
		rsContext.Close
		Set rsContext = Nothing
	End Function
	'=================================================
	'函数名：RelatedFlash
	'作  用：显示相关FLASH
	'参  数：sRelated ----相关FLASH
	'=================================================
	Private Function RelatedFlash(sRelated, topic, flashid)
		Dim rsRdlated, SQL, HtmlFileUrl, HtmlFileName
		Dim strtitle, title, strContent
		Dim strRelated, arrRelated, i, Resize, strRearrange
		Dim strKey,FlashUrl,miniatureUrl,miniature,strminiature
		Dim ArrayTemp()
		
		On Error Resume Next
		strRelated = Replace(Replace(Replace(Replace(sRelated, "[", ""), "]", ""), "'", ""), "%", "")
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
		SQL = "SELECT Top " & CInt(enchiasp.HtmlSetting(1)) & " A.flashid,A.ClassID,A.ColorMode,A.FontMode,A.title,A.AllHits,A.miniature,A.addTime,A.HtmlFileDate,C.HtmlFileDir,C.UseHtml FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.flashid <> " & flashid & " And " & strRelated & " ORDER BY A.flashid DESC"
		Set rsRdlated = enchiasp.Execute(SQL)
		If rsRdlated.EOF And rsRdlated.BOF Then
			RelatedSoft = ""
			Set rsRdlated = Nothing
			Exit Function
		Else
			i = 0
			Resize = 0
			Do While Not rsRdlated.EOF
				ReDim Preserve ArrayTemp(i + Resize)
				strContent = ArrayTemp(i) & enchiasp.HtmlSetting(4)
				strtitle = rsRdlated("title")
				strtitle = enchiasp.GotTopic(strtitle, CInt(enchiasp.HtmlSetting(2)))
				strtitle = enchiasp.ReadFontMode(strtitle, rsRdlated("ColorMode"), rsRdlated("FontMode"))
				If CreateHtml <> 0 Then
					HtmlFileUrl = ChannelRootDir & rsRdlated("HtmlFileDir") & enchiasp.ShowDatePath(rsRdlated("HtmlFileDate"), enchiasp.HtmlPath)
					HtmlFileName = enchiasp.ReadFileName(rsRdlated("HtmlFileDate"), rsRdlated("flashid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
					FlashUrl = HtmlFileUrl & HtmlFileName
					title = "<a href=" & HtmlFileUrl & HtmlFileName & " title='" & rsRdlated("title") & "'>" & strtitle & "</a>"
				Else
					FlashUrl = "show.asp?id=" & rsRdlated("flashid")
					title = "<a href=show.asp?id=" & rsRdlated("flashid") & " title='" & rsRdlated("title") & "'>" & strtitle & "</a>"
				End If
				If Not IsNull(rsRdlated("miniature")) Then
					strminiature = rsRdlated("miniature")
				End If
				miniatureUrl = enchiasp.GetImageUrl(strminiature, enchiasp.ChannelDir)
				miniature = enchiasp.GetFlashAndPic(miniatureUrl, CInt(enchiasp.HtmlSetting(9)), CInt(enchiasp.HtmlSetting(10)))
				miniature = "<a href='" & FlashUrl & "' title='" & Rs("title") & "'>" & miniature & "</a>"
				strContent = Replace(strContent, "{$Miniature}", miniature)
				strContent = Replace(strContent, "{$FlashTopic}", title)
				strContent = Replace(strContent, "{$AllHits}", rsRdlated("AllHits"))
				strContent = Replace(strContent, "{$DateTime}", enchiasp.ShowDateTime(rsRdlated("addTime"), CInt(enchiasp.HtmlSetting(3))))
				ArrayTemp(i) = strContent
				rsRdlated.MoveNext
				i = i + 1
			Loop
		End If
		rsRdlated.Close
		Set rsRdlated = Nothing
		strRearrange = Join(ArrayTemp, vbCrLf)
		RelatedFlash = strRearrange
	End Function
	Private Function PreviewMode(url,modeid)
		PreviewMode = ""
		If Len(url) < 3 Then Exit Function
		Dim strTemp
		Select Case CInt(modeid)
		Case 1
			strTemp = enchiasp.HtmlSetting(11)
		Case 2
			strTemp = enchiasp.HtmlSetting(12)
		Case 3
			strTemp =  enchiasp.HtmlSetting(13)
		Case 4
			strTemp = enchiasp.HtmlSetting(14)
		Case 5
			strTemp = enchiasp.HtmlSetting(15)
		End Select
		strTemp = Replace(strTemp, "{$ShowUrl}", Rs("showurl"))
		PreviewMode = Replace(strTemp, "{$ShowPlayUrl}", FormatShowUrl(url))
	End Function
	Public Function FormatShowUrl(ByVal url)
		FormatShowUrl = ""
		Dim strUrl
		If IsNull(url) Then Exit Function
		If Len(url) < 3 Then Exit Function
		If Left(url,1) = "/" Then
			FormatShowUrl = Trim(url)
			Exit Function
		End If
		strUrl = Left(url,10)
		If InStr(strUrl, "://") > 0 Then
			FormatShowUrl = Trim(url)
			Exit Function
		End If
		If InStr(strUrl, ":\") > 0 Then
			FormatShowUrl = Trim(url)
			Exit Function
		End If
		FormatShowUrl = ChannelRootDir & Trim(url)
	End Function
	
	'================================================
	'过程名：ReplaceString
	'作  用：替换模板内容
	'================================================
	Private Sub ReplaceString()
		HtmlContent = Replace(HtmlContent, "{$SelectedType}", "")
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = HTML.ReadFlashPic(HtmlContent)
		HtmlContent = HTML.ReadFlashList(HtmlContent)
		HtmlContent = HTML.ReadPopularFlash(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent,"{$FlashIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
	End Sub
	'#############################\\FLASH列表开始//#############################
	'=================================================
	'过程名：BuildFlashList
	'作  用：显示FLASH列表页面
	'=================================================
	Public Sub BuildFlashList()
		If CreateHtml <> 0 Then
			Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
			Exit Sub
		Else
			enchiasp.PreventInfuse
			If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
				Response.Write ("错误的系统参数!请输入整数")
				Response.End
			End If
			If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
				CurrentPage = CLng(Request("page"))
			Else
				CurrentPage = 1
			End If
			classid = enchiasp.ChkNumeric(Request("classid"))
			Response.Write LoadFlashList(ClassID, 1)
		End If
	End Sub
	'=================================================
	'过程名：LoadFlashList
	'作  用：载入FLASH列表
	'=================================================
	Public Function LoadFlashList(clsid, n)
		On Error Resume Next
		Dim rsClass
		Dim HtmlFileName,maxparent,strMaxParent

		PageType = 1
		
		If Not IsNumeric(clsid) Then Exit Function
		Set rsClass = enchiasp.Execute("SELECT ClassID,ClassName,ChildStr,ParentID,ParentStr,Child,skinid,HtmlFileDir,UseHtml FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID=" & clsid)
		If rsClass.BOF And rsClass.EOF Then
			If CreateHtml = 0 Then
				Response.Write "<meta http-equiv=""refresh"" content=""2;url='/"">" & vbNewLine
				Response.Write "<p align=""center"" style=""font-size: 12px;color: red;"">对不起，该页面发生了错误，无法访问! 系统两秒后自动转到网站首页......</p>" & vbNewLine
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
		HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & strFileDir
		
		HtmlContent = Replace(enchiasp.HtmlContent, "|||@@@|||", "")
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ClassID}", ClassID)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", strClassName)
		HtmlContent = Replace(HtmlContent, "{$FlashIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$IndexTitle}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$strClassName}", strClassName)
		ReplaceContent
		maxparent = enchiasp.ChkNumeric(enchiasp.HtmlSetting(5))
		maxperpage = CInt(enchiasp.HtmlSetting(1))
		strlen = enchiasp.ChkNumeric(enchiasp.HtmlSetting(9))
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		TotalNumber = enchiasp.Execute("SELECT COUNT(flashid) FROM ECCMS_FlashList WHERE ChannelID = " & ChannelID & " And isAccept > 0 And ClassID in (" & ChildStr & ")")(0)
		If maxparent > 0 And Child > 0 And TotalNumber > maxparent Then
			strMaxParent = " TOP " & maxparent
			TotalNumber = maxparent
		Else
			strMaxParent = ""
		End If
		TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT" & strMaxParent & " A.flashid,A.ClassID,A.title,A.ColorMode,A.FontMode,A.Introduce,A.filesize,A.Author,A.star,A.miniature,A.UserName,A.addTime,A.AllHits,A.grade,A.HtmlFileDate,A.isBest,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,C.UseHtml FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.ClassID in (" & ChildStr & ") ORDER BY A.isTop DESC, A.addTime DESC ,A.flashid DESC"
		Rs.Open SQL, Conn, 1, 1
		If Err.Number <> 0 Then Response.Write "SQL 查询错误"
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
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			If CreateHtml <> 0 Then
				Call LoadFlashHtmlList(n)
			Else
				Call LoadFlashAspList
			End If
		End If
		Rs.Close: Set Rs = Nothing
		LoadFlashList = HtmlContent
	End Function
	'================================================
	'过程名：ReplaceContent
	'作  用：替换模板内容
	'================================================
	Private Sub ReplaceContent()
		HtmlContent = HTML.ReadCurrentStation(HtmlContent, ChannelID, ClassID, strClassName, ParentID, strParent, strFileDir)
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = HTML.ReadNewsPicAndText(HtmlContent)
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
		HtmlContent = HTML.ReadFlashList(HtmlContent)
		HtmlContent = HTML.ReadFlashPic(HtmlContent)
		HtmlContent = HTML.ReadStatistic(HtmlContent)
		HtmlContent = HTML.ReadPopularFlash(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
	End Sub
	'================================================
	'过程名：LoadFlashHtmlList
	'作  用：装载FLASH列表HTML
	'================================================
	Private Sub LoadFlashHtmlList(n)
		Dim HtmlFileName
		Dim Perownum,ii,w
		
		Perownum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(4))
		
		If IsNull(TempListContent) Then Exit Sub
		
		enchiasp.CreatPathEx (HtmlFilePath)
		For CurrentPage = n To TotalPageNum
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			ListContent = ""
			j = (CurrentPage - 1) * maxperpage + 1
			If Perownum > 1 Then 
				ListContent = enchiasp.HtmlSetting(6)
				w = FormatPercent(100 / Perownum / 100,0)
			End If
			
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.end
				
				If Perownum > 1 Then
					ListContent = ListContent & "<tr valign=""top"">" & vbCrLf
					For ii = 1 To Perownum
						ListContent = ListContent & "<td width=""" & w & """ class=""Flashlistrow"">"
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
		
	End Sub
	'================================================
	'过程名：LoadFlashAspList
	'作  用：装载FLASH列表ASP
	'================================================
	Private Sub LoadFlashAspList()
		Dim Perownum,ii,w
		
		If IsNull(TempListContent) Then Exit Sub
		
		Perownum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(4))
		i = 0
		Rs.MoveFirst
		If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
		ListContent = ""
		j = (CurrentPage - 1) * maxperpage + 1
		If Perownum > 1 Then 
			ListContent = enchiasp.HtmlSetting(6)
			w = FormatPercent(100 / Perownum / 100,0)
		End If
		
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.end
			
			If Perownum > 1 Then
				ListContent = ListContent & "<tr valign=""top"">" & vbCrLf
				For ii = 1 To Perownum
					ListContent = ListContent & "<td width=""" & w & """ class=""Flashlistrow"">"
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
		If Perownum > 1 Then ListContent = ListContent & "</table>" & vbCrLf
		Dim strPagination
		strPagination = ShowListPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, ASPCurrentPage(PageType), strClassName)
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
		HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strPagination)
	End Sub
	'================================================
	'过程名：LoadListDetail
	'作  用：装载子级软件列表细节
	'================================================
	Private Sub LoadListDetail()
		Dim sTitle, sTopic, title, ListStyle
		Dim FlashUrl, FlashTime, sClassName,strminiature
		Dim miniatureUrl, miniature,Introduce
		
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
		On Error Resume Next
		If CInt(CreateHtml) <> 0 Then
			FlashUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath) & ReadPagination(0)
			sClassName = ChannelRootDir & Rs("HtmlFileDir")
		Else
			FlashUrl = ChannelRootDir & "show.asp?id=" & Rs("flashid")
			sClassName = ChannelRootDir & "list.asp?classid=" & Rs("ClassID")
		End If
		If Not IsNull(Rs("miniature")) Then
			strminiature = Rs("miniature")
		End If
		miniatureUrl = enchiasp.GetImageUrl(strminiature, enchiasp.ChannelDir)
		miniature = enchiasp.GetFlashAndPic(miniatureUrl, CInt(enchiasp.HtmlSetting(7)), CInt(enchiasp.HtmlSetting(8)))
		miniature = "<a href='" & FlashUrl & "' title='" & Rs("title") & "'>" & miniature & "</a>"
		sClassName = "<a href='" & sClassName & "' title='" & Rs("ClassName") & "'>" & Rs("ClassName") & "</a>"
		title = "<a href='" & FlashUrl & "' title='" & Rs("title") & "' class=""flashtopic"">" & sTitle & "</a>"

		Introduce = enchiasp.CutString(Rs("Introduce"), CInt(enchiasp.HtmlSetting(3)))
		
		FlashTime = enchiasp.ShowDateTime(Rs("addTime"), CInt(enchiasp.HtmlSetting(2)))
		ListContent = Replace(ListContent, "{$ClassifyName}", sClassName)
		ListContent = Replace(ListContent, "{$FlashTitle}", title)
		ListContent = Replace(ListContent, "{$FlashTopic}", sTitle)
		ListContent = Replace(ListContent, "{$FlashUrl}", FlashUrl)
		ListContent = Replace(ListContent, "{$Miniature}", miniature)
		ListContent = Replace(ListContent, "{$FlashID}", Rs("flashid"))
		ListContent = Replace(ListContent, "{$FlashHits}", Rs("AllHits"))
		ListContent = Replace(ListContent, "{$Star}", Rs("star"))
		ListContent = Replace(ListContent, "{$FlashDateTime}", FlashTime)
		ListContent = Replace(ListContent, "{$Introduce}", Introduce)
		ListContent = Replace(ListContent, "{$ListStyle}", ListStyle)
		ListContent = Replace(ListContent, "{$Author}", enchiasp.ChkNull(Rs("Author")))
		ListContent = Replace(ListContent, "{$UserName}", Rs("UserName"))
		ListContent = Replace(ListContent, "{$grade}", Rs("grade"))
		ListContent = Replace(ListContent, "{$IsTop}", Rs("IsTop"))
		ListContent = Replace(ListContent, "{$FileSize}", ReadFilesize(Rs("filesize")))
		ListContent = Replace(ListContent, "{$IsBest}", Rs("IsBest"))
		ListContent = Replace(ListContent, "{$Order}", j)
	End Sub
	'///---FLASH列表结束
	'///---FLASH列表开始,如:最新,推荐,热门FLASH
	'-- 最新FLASH列表
	Public Sub BuildNewFlash()
		CurrentPage = enchiasp.ChkNumeric(Request("page"))
		If CurrentPage = 0 Then CurrentPage = 1
		Response.Write LoadOtherFlshList(3)
	End Sub
	'-- 热门FLASH列表
	Public Sub BuildHotFlash()
		CurrentPage = enchiasp.ChkNumeric(Request("page"))
		If CurrentPage = 0 Then CurrentPage = 1
		Response.Write LoadOtherFlshList(2)
	End Sub
	'-- 推荐FLASH列表
	Public Sub BuildBestFlash()
		CurrentPage = enchiasp.ChkNumeric(Request("page"))
		If CurrentPage = 0 Then CurrentPage = 1
		Response.Write LoadOtherFlshList(1)
	End Sub
	'-- 推荐FLASH列表
	Public Sub BuildFlashSpecial()
		CurrentPage = enchiasp.ChkNumeric(Request("page"))
		If CurrentPage = 0 Then CurrentPage = 1
		Response.Write LoadOtherFlshList(1)
	End Sub
	'=================================================
	'过程名：LoadOtherFlshList
	'作  用：载入其它FLASH列表
	'=================================================
	Public Function LoadOtherFlshList(t)
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
		HtmlContent = Replace(HtmlContent, "{$FlashIndex}", strIndexName)
		PageType = 3
		If CInt(t) = 1 Then
			strClassName = enchiasp.HtmlSetting(10)			
			SQL1 = "And IsBest>0"
			SQL2 = "And A.IsBest>0 ORDER BY A.addTime DESC,A.flashid DESC"
		ElseIf CInt(t) = 2 Then
			strClassName = enchiasp.HtmlSetting(11)
			
			SQL1 = ""
			SQL2 = "ORDER BY A.AllHits DESC,A.addTime DESC,A.flashid DESC"
		Else
			strClassName = enchiasp.HtmlSetting(12)
			SQL1 = ""
			SQL2 = "ORDER BY A.addTime DESC ,A.flashid DESC"
		End If
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", strClassName)
		Call ReplaceString
		maxperpage = CLng(enchiasp.HtmlSetting(1))
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		'记录总数
		TotalNumber = enchiasp.Execute("SELECT COUNT(flashid) FROM ECCMS_FlashList WHERE ChannelID = " & ChannelID & " And isAccept>0  " & SQL1 & "")(0)
		If TotalNumber >= CLng(enchiasp.HtmlSetting(5)) Then TotalNumber = CLng(enchiasp.HtmlSetting(5))
		TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT TOP " & CLng(enchiasp.HtmlSetting(5)) & " A.flashid,A.ClassID,A.title,A.ColorMode,A.FontMode,A.Introduce,A.filesize,A.Author,A.star,A.miniature,A.UserName,A.addTime,A.AllHits,A.grade,A.HtmlFileDate,A.isBest,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,C.UseHtml FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept>0 " & SQL2
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
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			If CreateHtml <> 0 Then
				Call LoadOtherListHtml(t)
			Else
				Call LoadFlashAspList
			End If
		End If
		Rs.Close: Set Rs = Nothing
		
		If CreateHtml = 0 Then LoadOtherFlshList = HtmlContent
	End Function
	'================================================
	'过程名：LoadOtherListHtml
	'作  用：装载其它列表并生成HTML
	'================================================
	Private Sub LoadOtherListHtml(t)
		Dim HtmlFileName, sulCurrentPage
		Dim Perownum,ii,w
		
		If IsNull(TempListContent) Then Exit Sub
		On Error Resume Next

		Perownum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(4))
		
		enchiasp.CreatPathEx (HtmlFilePath)
		For CurrentPage = n To TotalPageNum
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			ListContent = ""
			j = (CurrentPage - 1) * maxperpage + 1
			If Perownum > 1 Then 
				ListContent = enchiasp.HtmlSetting(6)
				w = FormatPercent(100 / Perownum / 100,0)
			End If
			
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.end
				
				If Perownum > 1 Then
					ListContent = ListContent & "<tr valign=""top"">" & vbCrLf
					For ii = 1 To Perownum
						ListContent = ListContent & "<td width=""" & w & """class=""Flashlistrow"">"
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
			If t = 1 Then
				sulCurrentPage = enchiasp.HtmlPrefix & "Best"
			ElseIf t = 2 Then
				sulCurrentPage = enchiasp.HtmlPrefix & "Hot"
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
			HtmlFileName = HtmlFilePath & sulCurrentPage & enchiasp.Supplemental(CurrentPage, 3) & enchiasp.HtmlExtName
			enchiasp.CreatedTextFile HtmlFileName, HtmlTemplate
			If IsShowFlush = 1 Then 
				Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "列表HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
				Response.Flush
			End If
		Next

	End Sub
	'#############################\\FLASH搜索开始//#############################
	Public Sub BuildFlashSearch()
		Dim SearchMaxPageList
		Dim Action, findword
		Dim rsClass, strNoResult
		Dim strWord, s
		
		PageType = 5
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
				findword = "A.title like '%" & keyword & "%'"
			ElseIf LCase(Action) = "content" Then
				If CInt(enchiasp.FullContQuery) <> 0 Then
					findword = "A.Content like '%" & keyword & "%'"
				Else
					Call OutAlertScript(Replace(Replace(enchiasp.HtmlSetting(10), Chr(34), "\"""), vbCrLf, ""))
					Exit Sub
				End If
			Else
				findword = "A.title like '%" & keyword & "%'"
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
		HtmlContent = Replace(HtmlContent, "{$FlashIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$KeyWord}", KeyWord)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ModuleName & "搜索")
		HtmlContent = Replace(HtmlContent, "{$QueryKeyWord}", "<font color=red><strong>" & keyword & "</strong></font>")
		Call ReplaceString
		If LCase(Action) <> "isweb" Then
			If IsNumeric(Request("classid")) And Request("classid") <> "" Then
				Set rsClass = enchiasp.Execute("SELECT ClassID,ChildStr FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID=" & CLng(Request("classid")))
				If rsClass.BOF And rsClass.EOF Then
					HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strNoResult, 1, 1, 1)
					HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "")
					HtmlContent = Replace(HtmlContent, "{$totalrec}", 0)
					HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
					Set rsClass = Nothing
					Response.Write HtmlContent
					Exit Sub
				Else
					findword = "A.ClassID IN (" & rsClass("ChildStr") & ") And " & findword
				End If
				rsClass.Close: Set rsClass = Nothing
			End If
			maxperpage = CInt(enchiasp.HtmlSetting(1))
			If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
				Response.Write ("错误的系统参数!请输入整数")
				Response.End
			End If
			If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
				CurrentPage = CInt(Request("page"))
			Else
				CurrentPage = 1
			End If
			If CInt(CurrentPage) = 0 Then CurrentPage = 1
			
			Set Rs = CreateObject("ADODB.Recordset")
			SQL = "SELECT TOP " & SearchMaxPageList & " A.flashid,A.ClassID,A.title,A.ColorMode,A.FontMode,A.Introduce,A.filesize,A.Author,A.star,A.miniature,A.UserName,A.addTime,A.AllHits,A.grade,A.HtmlFileDate,A.isBest,C.ClassName,C.HtmlFileDir,C.UseHtml FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And " & findword & " ORDER BY A.addTime DESC ,A.flashid DESC"
			Rs.Open SQL, Conn, 1, 1
			If Rs.BOF And Rs.EOF Then
				'如果没有找到相关内容,清除掉无用的标签代码
				HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strNoResult, 1, 1, 1)
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
				'获取模板标签[ShowRepetend][/ReadFlashList]中的字符串
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
	'过程名：LoadSearchList
	'作  用：装载软件搜索列表
	'================================================
	Private Sub LoadSearchList()
		If IsNull(TempListContent) Then Exit Sub
		i = 0
		If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
		ListContent = ""
		j = (CurrentPage - 1) * maxperpage + 1
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
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
	'作  用：装载搜索列表细节
	'================================================
	Private Sub SearchResult()
		Dim sTitle, sTopic, title, ListStyle, TitleWord
		Dim FlashUrl, addTime, sClassName, FlashImage, FlashIntro
		Dim miniatureUrl,miniature,strminiature
		
		ListContent = ListContent & TempListContent
		If (i Mod 2) = 0 Then
			ListStyle = 1
		Else
			ListStyle = 2
		End If
		TitleWord = Replace(Rs("title"), keyword, "<font color=red>" & keyword & "</font>")
		sTitle = enchiasp.ReadFontMode(TitleWord, Rs("ColorMode"), Rs("FontMode"))
		
		If CInt(CreateHtml) <> 0 Then
			FlashUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath) & ReadPagination(0)
			sClassName = ChannelRootDir & Rs("HtmlFileDir")
		Else
			FlashUrl = ChannelRootDir & "show.asp?id=" & Rs("flashid")
			sClassName = ChannelRootDir & "list.asp?classid=" & Rs("ClassID")
		End If
		
		sClassName = "<a href=""" & sClassName & """ title=""" & Rs("ClassName") & """ target=""_blank""><span style=""color:" & enchiasp.MainSetting(3) & """>" & Rs("ClassName") & "</span></a>"
		title = "<a href='" & FlashUrl & "' title='" & Rs("title") & "' class=""showtopic"" target=""_blank"">" & sTitle & "</a>"
		FlashIntro = enchiasp.CutString(Rs("Introduce"), CInt(enchiasp.HtmlSetting(3)))
		FlashIntro = Replace(FlashIntro, keyword, "<font color=red>" & keyword & "</font>")
		If Not IsNull(Rs("miniature")) Then
			strminiature = Rs("miniature")
		End If
		miniatureUrl = enchiasp.GetImageUrl(strminiature, enchiasp.ChannelDir)
		miniature = enchiasp.GetFlashAndPic(miniatureUrl, CInt(enchiasp.HtmlSetting(11)), CInt(enchiasp.HtmlSetting(12)))
		miniature = "<a href='" & FlashUrl & "' title='" & Rs("title") & "'>" & miniature & "</a>"
		
		addTime = enchiasp.ShowDateTime(Rs("addTime"), CInt(enchiasp.HtmlSetting(2)))
		ListContent = Replace(ListContent, "{$KeyWord}", keyword)
		ListContent = Replace(ListContent, "{$totalrec}", TotalNumber)
		ListContent = Replace(ListContent, "{$ClassifyName}", sClassName)
		ListContent = Replace(ListContent, "{$FlashTitle}", title)
		ListContent = Replace(ListContent, "{$FlashTopic}", sTitle)
		ListContent = Replace(ListContent, "{$FlashUrl}", FlashUrl)
		ListContent = Replace(ListContent, "{$Miniature}", miniature)
		ListContent = Replace(ListContent, "{$Star}", Rs("star"))
		ListContent = Replace(ListContent, "{$FlashHits}", Rs("AllHits"))
		ListContent = Replace(ListContent, "{$UserName}", Rs("username"))
		ListContent = Replace(ListContent, "{$DateAndTime}", addTime)
		ListContent = Replace(ListContent, "{$Introduce}", FlashIntro)
		ListContent = Replace(ListContent, "{$ListStyle}", ListStyle)
		ListContent = Replace(ListContent, "{$FlashSize}", ReadFilesize(Rs("filesize")))
		ListContent = Replace(ListContent, "{$Author}", enchiasp.ChkNull(Rs("Author")))
		ListContent = Replace(ListContent, "{$FlashID}", Rs("flashid"))
		ListContent = Replace(ListContent, "{$Order}", j)
	End Sub
	'//--搜索结束
	'================================================
	'函数名：FlashComment
	'作  用：FLASH评论
	'================================================
	Private Function FlashComment(flashid)
		Dim rsComment, SQL, strContent, strComment
		Dim i, Resize, strRearrange
		Dim ArrayTemp()

		On Error Resume Next
		Set rsComment = enchiasp.Execute("SELECT TOP " & CInt(enchiasp.HtmlSetting(5)) & " content,Grade,username,postime,postip FROM ECCMS_Comment WHERE ChannelID=" & ChannelID & " And postid = " & flashid & " ORDER BY postime DESC,CommentID DESC")
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
		FlashComment = strRearrange
	End Function
	'================================================
	'过程名：BuildFlashComment
	'作  用：显示FLASH评论
	'================================================
	Public Sub BuildFlashComment()
		Dim title, HtmlFileUrl, HtmlFileName
		Dim AverageGrade, TotalComment, TempListContent
		Dim strComment, strCheckBox, strAdminComment

		enchiasp.PreventInfuse
		strCheckBox = ""
		strAdminComment = ""
		On Error Resume Next
		flashid = enchiasp.ChkNumeric(Request("flashid"))
		If flashid = 0 Then
			Response.Write "<Br><Br><Br>Sorry！错误的系统参数,请选择正确的连接方式。"
			Response.End
		End If
		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 8, skinid
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$FlashIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$IndexTitle}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ModuleName & "评论")
		HtmlContent = Replace(HtmlContent, "{$flashid}", flashid)
		HtmlContent = Replace(HtmlContent, "{$FlashID}", flashid)
		'获得软件标题
		SQL = "SELECT TOP 1 A.flashid,A.ClassID,A.title,A.HtmlFileDate,A.ForbidEssay,C.HtmlFileDir,C.UseHtml FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.flashid = " & flashid
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			Response.Write "已经没有了"
			Set Rs = Nothing
			Exit Sub
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("flashid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				title = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & Rs("title") & "</a>"
			Else
				title = "<a href=show.asp?id=" & Rs("flashid") & ">" & Rs("title") & "</a>"
			End If
			ForbidEssay = Rs("ForbidEssay")
		End If
		Rs.Close
		Set Rs = CreateObject("adodb.recordset")
		SQL = "SELECT COUNT(CommentID) As TotalComment,AVG(Grade) As avgGrade FROM ECCMS_Comment WHERE ChannelID=" & ChannelID & " And postid = " & flashid
		Set Rs = enchiasp.Execute(SQL)
		TotalComment = Rs("TotalComment")
		AverageGrade = Round(Rs("avgGrade"))
		If IsNull(AverageGrade) Then AverageGrade = 0
		Rs.Close: Set Rs = Nothing
		HtmlContent = Replace(HtmlContent, "{$FlashTitle}", title)
		HtmlContent = Replace(HtmlContent, "{$TotalComment}", TotalComment)
		HtmlContent = Replace(HtmlContent, "{$AverageGrade}", AverageGrade)
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("错误的系统参数!请输入整数")
			Response.End
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		'每页显示评论数
		maxperpage = CInt(enchiasp.PaginalNum)
		'记录总数
		TotalNumber = TotalComment
		TotalPageNum = CInt(TotalNumber / maxperpage)  '得到总页数
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM ECCMS_Comment WHERE ChannelID=" & ChannelID & " And postid = " & flashid & " ORDER BY postime DESC,CommentID DESC"
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
				If Not Response.IsClientConnected Then Response.End
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
			strAdminComment = strAdminComment & "<input type=hidden name=flashid value='" & flashid & "'>" & vbNewLine
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
		
	End Sub
	'================================================
	'过程名：ShowCommentPage
	'作  用：评论分页
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
			strTemp = strTemp & "共有评论 <font COLOR=#FF0000>" & TotalNumber & "</font> 个&nbsp;&nbsp;<a href=" & FileName & "?page=1&flashid=" & Request("flashid") & ">首 页</a>&nbsp;&nbsp;"
			strTemp = strTemp & "<a href=" & FileName & "?page=" & CurrentPage - 1 & "&flashid=" & Request("flashid") & ">上一页</a>&nbsp;&nbsp;&nbsp;"
		End If

		If n - CurrentPage < 1 Then
			strTemp = strTemp & "下一页&nbsp;&nbsp;尾 页 " & vbCrLf
		Else
			strTemp = strTemp & "<a href=" & FileName & "?page=" & (CurrentPage + 1) & "&flashid=" & Request("flashid") & ">下一页</a>"
			strTemp = strTemp & "&nbsp;&nbsp;<a href=" & FileName & "?page=" & n & "&flashid=" & Request("flashid") & ">尾 页</a>" & vbCrLf
		End If
		strTemp = strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
		strTemp = strTemp & "&nbsp;<b>" & maxperpage & "</b>个/页 " & vbCrLf
		strTemp = strTemp & "</td></tr></table>" & vbCrLf
		HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strTemp)
	End Sub
	'================================================
	'过程名：CommentDel
	'作  用：评论删除
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
				enchiasp.Execute ("delete from ECCMS_Comment WHERE ChannelID=" & ChannelID & " And CommentID in (" & selCommentID & ")")
				Call OutHintScript("评论删除成功！")
			Else
				Call OutAlertScript("非法操作！你没有删除评论的权限。")
				Exit Sub
			End If
		End If
	End Sub
	'================================================
	'过程名：CommentSave
	'作  用：软件评论添加保存
	'================================================
	Public Sub CommentSave()
		If enchiasp.CheckPost = False Then
			Call OutAlertScript("您提交的数据不合法，请不要从外部提交表单。")
			Exit Sub
		End If
		On Error Resume Next
		Call PreventRefresh
		If CInt(enchiasp.AppearGrade) <> 0 And Session("AdminName") = "" Then
			If CInt(enchiasp.AppearGrade) > CInt(enchiasp.membergrade) Then
				Call OutAlertScript("您没有发表评论的权限，如果您是会员请登陆后再参与评论。")
				Exit Sub
			End If
		End If
		If ForbidEssay <> 0 Then
			Call OutAlertScript("此" & enchiasp.ModuleName & "禁止发表评论！")
			Exit Sub
		End If
		If Trim(Request.Form("UserName")) = "" Then
			Call OutAlertScript("用户名不能为空！")
			Exit Sub
		End If
		If Len(Trim(Request.Form("UserName"))) > 15 Then
			Call OutAlertScript("用户名不能大于15个字符！")
			Exit Sub
		End If
		If enchiasp.strLength(Request.Form("content")) < enchiasp.LeastString Then
			Call OutAlertScript("评论内容不能小于" & enchiasp.LeastString & "字符！")
			Exit Sub
		End If
		If enchiasp.strLength(Request.Form("content")) > enchiasp.MaxString Then
			Call OutAlertScript("评论内容不能大于" & enchiasp.MaxString & "字符！")
			Exit Sub
		End If
		flashid = enchiasp.ChkNumeric(Request.Form("flashid"))
		Set Rs = CreateObject("ADODB.RecordSet")
		SQL = "SELECT * FROM ECCMS_Comment WHERE (CommentID is null)"
		Rs.Open SQL, Conn, 1, 3
		Rs.AddNew
			Rs("ChannelID") = ChannelID
			Rs("postid") = flashid
			Rs("UserName") = Trim(Request.Form("UserName"))
			Rs("Grade") = Trim(Request.Form("Grade"))
			Rs("content") = Request.Form("content")
			Rs("postime") = Now()
			Rs("postip") = enchiasp.GetUserip
		Rs.Update
		Rs.Close: Set Rs = Nothing
		If CreateHtml <> 0 Then LoadFlashInfo(flashid)
		Session("UserRefreshTime") = Now()
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
		Exit Sub
	End Sub
	Public Sub PreventRefresh()
		Dim RefreshTime

		RefreshTime = 20
		If DateDiff("s", Session("UserRefreshTime"), Now()) < RefreshTime Then
			Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=" & RefreshTime & "><br>本页面起用了防刷新机制，请不要在" & RefreshTime & "秒内连续刷新本页面<BR>正在打开页面，请稍后……"
			Response.End
		End If
	End Sub
	Private Function ReadPagination(n)
		Dim HtmlFileName, CurrentPage
		
		CurrentPage = n
		HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("flashid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, CurrentPage)
		ReadPagination = HtmlFileName
	End Function
	Private Function ReadListPageName(ClassID, CurrentPage)
		ReadListPageName = enchiasp.ClassFileName(ClassID, enchiasp.HtmlExtName, enchiasp.HtmlPrefix, CurrentPage)
	End Function
	Public Function ASPCurrentPage(stype)
		Dim CurrentUrl
		Select Case stype
			Case "1"
				CurrentUrl = "&amp;classid=" & Trim(Request("classid"))
			Case "2"
				CurrentUrl = "&amp;sid=" & Trim(Request("sid"))
			Case "3"
				CurrentUrl = ""
			Case "4"
				CurrentUrl = ""
			Case "6"
				CurrentUrl = "&amp;type=" & enchiasp.CheckStr(Request("type"))
			Case Else
				If Trim(Request("word")) <> "" Then
					CurrentUrl = "&amp;word=" & Trim(Request("word"))
				Else
					CurrentUrl = "&amp;act=" & Trim(Request("act")) & "&amp;classid=" & Trim(Request("classid")) & "&amp;keyword=" & Trim(Request("keyword"))
				End If
		End Select
		ASPCurrentPage = CurrentUrl
	End Function
	'================================================
	'函数名：ReadFilesize
	'作  用：读取文件大小
	'================================================
	Function ReadFilesize(ByVal para)
		On Error Resume Next
		Dim strFileSize, parasize
		
		parasize = Clng(para)
		
		If parasize = 0 Then
			ReadFilesize = "未知"
			Exit Function
		End If

		If parasize > 1024 Then
			strFileSize = Round(parasize / 1024, 2) & " MB"
		Else
			strFileSize = parasize & " KB"
		End If
		ReadFilesize = strFileSize
	End Function
	Public Function ReadComeFrom(ByVal strContent)
		ReadComeFrom = ""
		If IsNull(strContent) Then Exit Function
		If Trim(strContent) = "" Then Exit Function
		strContent = " " & strContent & " "
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""|'])+)"
		strContent = re.Replace(strContent,"<a target=""_blank"" href=$1>$1</a>")
		re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""])+)$([^\[|']*)"
		strContent = re.Replace(strContent,"<a target=""_blank"" href=$1>$1</a>")
		re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""|'])+)"
		strContent = re.Replace(strContent,"$1<a target=""_blank"" href=$2>$2</a>")
		re.Pattern = "([\s])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
		strContent = re.Replace(strContent,"<a target=""_blank"" href=""http://$2"">$2</a>")
		Set re = Nothing
		ReadComeFrom = Trim(strContent)
	End Function
	
End Class
%>