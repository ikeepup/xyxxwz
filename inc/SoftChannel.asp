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
Set enchicms = New SoftChannel_Cls

Class SoftChannel_Cls
	Private ChannelID, CreateHtml, keyword
	Private Rs, SQL, ChannelRootDir, HtmlContent, strIndexName
	Private softid, SoftIntro, skinid, ClassID, SoftType
	Private maxperpage, TotalNumber, TotalPageNum, CurrentPage, i, totalrec
	Private strFileDir, ParentID, strParent, strClassName, ChildStr, Child
	Private ListContent, TempListContent, HtmlTemplate, HtmlFilePath
	Private SpecialID, SpecialName, SpecialDir, PageType, ForbidEssay
	Private IsShowFlush, strInstallDir, j
	Private FoundErr,strlen

	Public Property Let Channel(chanid)
		ChannelID = chanid
	End Property
	Public Property Let ShowFlush(para)
		IsShowFlush = para
	End Property
	Private Sub Class_Initialize()
		On Error Resume Next
		FoundErr = False
		ChannelID = 2
		strlen = 0
	End Sub
	Private Sub Class_Terminate()
		Set HTML = Nothing
	End Sub
	Public Sub ChannelMain()
		enchiasp.ReadChannel (ChannelID)
		CreateHtml = CInt(enchiasp.IsCreateHtml)
		ChannelRootDir = enchiasp.InstallDir & enchiasp.ChannelDir
		strInstallDir = enchiasp.InstallDir
		strIndexName = "<a href='" & ChannelRootDir & "'>" & enchiasp.ChannelName & "</a>"
	End Sub

	'#############################\\执行软件下载首页开始//#############################
	'=================================================
	'过程名：ShowDownIndex
	'作  用：显示下载首页
	'=================================================
	Public Sub ShowDownIndex()
		On Error Resume Next
		LoadDownIndex
		If CreateHtml <> 0 Then
			Response.Write "<meta http-equiv=refresh content=0;url=index" & enchiasp.HtmlExtName & ">"
		Else
			Response.Write HtmlContent
		End If
	End Sub
	'=================================================
	'过程名：CreateDownIndex
	'作  用：生成下载首页的HTML
	'=================================================
	Public Sub CreateDownIndex()
		On Error Resume Next
		LoadDownIndex
		Dim FilePath
		
		FilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "index" & enchiasp.HtmlExtName
		enchiasp.CreatedTextFile FilePath, HtmlContent
		If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "首页HTML完成... <a href=" & FilePath & " target=_blank>" & Server.MapPath(FilePath) & "</a></li>" & vbNewLine
		Response.Flush
	End Sub
	Public Sub LoadDownIndex()
		On Error Resume Next
		Dim FilePath
		
		enchiasp.LoadTemplates ChannelID, 1, enchiasp.ChannelSkin
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ChannelName)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$SoftIndex}", strIndexName)
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
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
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = HtmlContent
	End Sub
	'#############################\\执行软件信息开始//#############################
	'=================================================
	'过程名：ShowArticleInfo
	'作  用：显示下载内容页面
	'=================================================
	Public Sub ShowDownIntro()
		If CreateHtml <> 0 Then
			Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
			Exit Sub
		Else
			enchiasp.PreventInfuse
			softid = enchiasp.ChkNumeric(Request("id"))
			Response.Write ReadSoftIntro(softid)
		End If
	End Sub
	'=================================================
	'函数名：ReadSoftIntro
	'作  用：读取软件内容
	'参  数：SoftID ----软件ID
	'=================================================
	Public Function ReadSoftIntro(softid)
		Dim SoftImageUrl, SoftImage, Previewimg, PreviewUrl, re
		Dim strImageSize, strPreviewSize, SoftReadme, softname, SoftVer
		Dim MemberSoft, HomePage, HomePageUrl, strContact, DownloadAddress
		Dim strDecode, strRegsite, strAuthor
		Dim strRegsites, strPreviewImg

		On Error Resume Next
		SQL = "SELECT A.*,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.SoftID=" & softid
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			ReadSoftIntro = ""
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
		SoftReadme = Rs("content")

		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		If enchiasp.HtmlSetting(18) <> "0" Then
			re.Pattern = "\[br\]"
			SoftReadme = re.Replace(SoftReadme, "<BR>")
			re.Pattern = "\[align=right\]"
			SoftReadme = re.Replace(SoftReadme, "<div align=right>")
			re.Pattern = "\[\/align\]"
			SoftReadme = re.Replace(SoftReadme, "</div>")
		Else
			re.Pattern = "\[br\]"
			SoftReadme = re.Replace(SoftReadme, "")
			re.Pattern = "\[align=right\](.*)\[\/align\]"
			SoftReadme = re.Replace(SoftReadme, "")
		End If
		Set re = Nothing
		DownloadAddress = ShowDownAddress(Rs("softid"))
		SoftIntro = UbbCode(SoftReadme)
		
		strImageSize = Split(enchiasp.HtmlSetting(14), "|")
		strPreviewSize = Split(enchiasp.HtmlSetting(15), "|")
		If enchiasp.CheckNull(Rs("SoftImage")) Then
			SoftImageUrl = enchiasp.GetImageUrl(Rs("SoftImage"), enchiasp.ChannelDir)
			SoftImage = enchiasp.GetFlashAndPic(SoftImageUrl, CInt(strImageSize(0)), CInt(strImageSize(1)))
			SoftImage = "<a href='" & ChannelRootDir & "Previewimg.asp?SoftID=" & softid & "' title='" & Rs("SoftName") & "' target=_blank>" & SoftImage & "</a>"
			Previewimg = enchiasp.GetFlashAndPic(SoftImageUrl, CInt(strPreviewSize(0)), CInt(strPreviewSize(1)))
			PreviewUrl = ChannelRootDir & "Previewimg.asp?SoftID=" & softid
			PreviewUrl = Replace(enchiasp.HtmlSetting(17), "{$PreviewUrl}", PreviewUrl)
		Else
			SoftImage = enchiasp.HtmlSetting(13)
			Previewimg = ""
			PreviewUrl = enchiasp.HtmlSetting(16)
		End If
		If enchiasp.CheckNull(Rs("Homepage")) Then
			HomePageUrl = Rs("Homepage")
			HomePage = Replace(enchiasp.HtmlSetting(10), "{$HomePageUrl}", Rs("Homepage"))
		Else
			HomePage = enchiasp.HtmlSetting(9)
			HomePageUrl = ""
		End If
		If enchiasp.CheckNull(Rs("Contact")) Then
			strContact = Replace(enchiasp.HtmlSetting(12), "{$ContactSite}", Rs("Contact"))
		Else
			strContact = enchiasp.HtmlSetting(11)
		End If
		If enchiasp.CheckNull(Rs("Decode")) Then
			strDecode = Replace(enchiasp.HtmlSetting(20), "{$strDecode}", Rs("Decode"))
		Else
			strDecode = enchiasp.HtmlSetting(19)
		End If
		
		
		If Rs("UserGroup") <> 0 Then
			MemberSoft = enchiasp.HtmlSetting(8)
		End If
		
		If enchiasp.CheckNull(Rs("Regsite")) Then
			strRegsite = Replace(enchiasp.HtmlSetting(24), "{$RegsiteUrl}", Rs("Regsite"))
			strRegsites = Trim(Rs("Regsite"))
		Else
			strRegsite = enchiasp.HtmlSetting(23)
			strRegsites = "#"
		End If
		If strRegsites = "#" Then
			strPreviewImg = ""
		Else
			strPreviewImg = "<img src=""" & strRegsites & """ border=""0"">"
			strPreviewImg = UbbCode(strPreviewImg)
		End If
		If enchiasp.CheckNull(Rs("Author")) Then
			strAuthor = Rs("Author")
		Else
			strAuthor = enchiasp.HtmlSetting(25)
		End If

		HtmlContent = enchiasp.HtmlContent
		If enchiasp.CheckNull(Rs("SoftVer")) Then
			softname = Trim(Rs("SoftName") & " " & Rs("SoftVer"))
			HtmlContent = Replace(HtmlContent, "{$SoftVer}", Rs("SoftVer"))
		Else
			softname = Trim(Rs("SoftName"))
			HtmlContent = Replace(HtmlContent, "{$SoftVer}", "")
		End If
		HtmlContent = Replace(HtmlContent, "{$Soft_Name}", Rs("SoftName"))
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$DownAddress}", DownloadAddress)
		HtmlContent = Replace(HtmlContent, "{$RegsiteUrl}", strRegsite)
		HtmlContent = Replace(HtmlContent, "{$Author}", strAuthor)
		HtmlContent = Replace(HtmlContent, "{$SoftImage}", SoftImage)
		HtmlContent = Replace(HtmlContent, "{$Previewimg}", Previewimg)
		HtmlContent = Replace(HtmlContent, "{$PreviewUrl}", PreviewUrl)
		HtmlContent = Replace(HtmlContent, "{$HomePage}", HomePage)
		HtmlContent = Replace(HtmlContent, "{$HomePageUrl}", HomePageUrl)
		HtmlContent = Replace(HtmlContent, "{$Contact}", strContact)
		HtmlContent = Replace(HtmlContent, "{$Decode}", strDecode)
		HtmlContent = Replace(HtmlContent, "{$MemberSoft}", MemberSoft)
		HtmlContent = Replace(HtmlContent, "{$SoftID}", Rs("SoftID"))
		HtmlContent = Replace(HtmlContent, "{$ClassName}", Rs("ClassName"))
		HtmlContent = Replace(HtmlContent, "{$SoftName}", softname)
		HtmlContent = Replace(HtmlContent, "{$SoftContent}", SoftIntro)
		HtmlContent = Replace(HtmlContent, "{$SoftTime}", Rs("SoftTime"))
		HtmlContent = Replace(HtmlContent, "{$UserName}", Rs("username"))
		HtmlContent = Replace(HtmlContent, "{$Language}", Rs("Languages"))
		HtmlContent = Replace(HtmlContent, "{$SoftType}", Rs("SoftType"))
		HtmlContent = Replace(HtmlContent, "{$RunSystem}", Rs("RunSystem"))
		HtmlContent = Replace(HtmlContent, "{$Impower}", Rs("impower"))
		HtmlContent = Replace(HtmlContent, "{$Star}", Rs("star"))
		HtmlContent = Replace(HtmlContent, "{$IsBest}", Rs("IsBest"))
		HtmlContent = Replace(HtmlContent, "{$IsTop}", Rs("IsTop"))
		HtmlContent = Replace(HtmlContent, "{$Regsite}", Rs("Regsite"))
		HtmlContent = Replace(HtmlContent, "{$PreviewPic}", strPreviewImg)
		HtmlContent = Replace(HtmlContent, "{$showreg}", Rs("showreg"))
		HtmlContent = Replace(HtmlContent, "{$PointNum}", Rs("PointNum"))
		HtmlContent = Replace(HtmlContent, "{$SoftPrice}", Rs("SoftPrice"))
		HtmlContent = Replace(HtmlContent, "{$SoftSize}", ReadSoftsize(Rs("SoftSize")))
		HtmlContent = Replace(HtmlContent, "{$FileSize}", CCur(Rs("SoftSize")))

		If InStr(HtmlContent, "{$FrontSoft}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$FrontSoft}", FrontSoft(softid))
		End If
		If InStr(HtmlContent, "{$NextSoft}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$NextSoft}", NextSoft(softid))
		End If
		If InStr(HtmlContent, "{$RelatedSoft}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$RelatedSoft}", RelatedSoft(Rs("Related"), Rs("SoftName"), Rs("SoftID")))
		End If
		If InStr(HtmlContent, "{$ShowHotSoft}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$ShowHotSoft}", ShowHotSoft(Rs("ClassID")))
		End If
		If InStr(HtmlContent, "{$SoftComment}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$SoftComment}", SoftComment(Rs("SoftID")))
		End If

		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$SoftIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", softname)
		HtmlContent = Replace(HtmlContent, "{$ClassID}", Rs("ClassID"))
		HtmlContent = Replace(HtmlContent, "{$SoftID}", softid)
		HtmlContent = HTML.ReadCurrentStation(HtmlContent, ChannelID, Rs("ClassID"), Rs("ClassName"), Rs("ParentID"), Rs("ParentStr"), Rs("HtmlFileDir"))
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = HTML.ReadSoftPic(HtmlContent)
		HtmlContent = HTML.ReadSoftList(HtmlContent)
		HtmlContent = HTML.ReadPopularSoft(HtmlContent)
	HtmlContent = HTML.ReadFriendLink(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		If CreateHtml <> 0 Then
			Call CreateSoftIntro
		Else
			ReadSoftIntro = HtmlContent
		End If
		Rs.Close: Set Rs = Nothing
	End Function

	'=================================================
	'函数名：CreateSoftIntro
	'作  用：生成软件内容
	'参  数：SoftID ----软件ID
	'=================================================
	Private Sub CreateSoftIntro()
		Dim HtmlFileName
		HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath)
		enchiasp.CreatPathEx (HtmlFilePath)
		HtmlFileName = HtmlFilePath & enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("SoftID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, CurrentPage)
		enchiasp.CreatedTextFile HtmlFileName, HtmlContent
		If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "信息HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
		Response.Flush
	End Sub
	'================================================
	'函数名：ShowDownAddress
	'作  用：显示软件下载地址
	'参  数：SoftID ----软件ID
	'================================================
	Private Function ShowDownAddress(softid)
		Dim rsAddress, sqlAddress, rsDown
		Dim i, AddressNum, ordinal, SoftNameStr, s_DownloadPath
		Dim DownloadName, DownloadPath, sDownloadName, sDownloadPath
		Dim DownAddress, strDownAddress
		On Error Resume Next
		If enchiasp.ChkNumeric(enchiasp.HtmlSetting(26)) = 0 Then
			Set rsDown = enchiasp.Execute("SELECT id,downid,DownFileName FROM [ECCMS_DownAddress] WHERE softid=" & CLng(softid))
			If Not (rsDown.BOF And rsDown.EOF) Then
				Do While Not rsDown.EOF
					'---- 如果使用了下载服务器,就打开下载服务器数据表
					sqlAddress = "SELECT downid,DownloadName,DownloadPath,IsDisp FROM ECCMS_DownServer WHERE ChannelID=" & ChannelID & " And depth=1 And rootid =" & rsDown("downid") & " And isLock=0 ORDER BY orders ASC"
					Set rsAddress = enchiasp.Execute(sqlAddress)
					If rsAddress.EOF And rsAddress.BOF Then
						DownloadPath = ""
						DownloadName = ""
					Else
						ordinal = 1
						Do While Not rsAddress.EOF
							'---- 是否直接显示软件直接的下载地址
							'----将所有的下载服务器地址都连成一个数组
							If rsAddress("IsDisp") = 0 Then
								DownloadPath = DownloadPath & ChannelRootDir & "download.asp?softid=" & softid & "&amp;downid=" & rsAddress("downid") & "&id=" & rsDown(0) & "|"
							Else
								DownloadPath = DownloadPath & rsAddress("DownloadPath") & Replace(rsDown(2), "*", ordinal) & "|"
							End If
							DownloadName = DownloadName & rsAddress("DownLoadName") & "|"
							rsAddress.MoveNext
							ordinal = ordinal + 1
						Loop
					End If
					Set rsAddress = Nothing
					rsDown.MoveNext
				Loop
			End If
			Set rsDown = Nothing
			'---- 读取下载地址字符串
			If enchiasp.CheckNull(Rs("DownAddress")) And Trim(Rs("DownAddress")) <> "|||" Then
				strDownAddress = Rs("DownAddress")
				strDownAddress = Split(strDownAddress, "|||")
				sDownloadName = strDownAddress(0) & "|"
				s_DownloadPath = Split(strDownAddress(1), "|")
				sDownloadPath = ""
				For i = 0 To UBound(s_DownloadPath)
					sDownloadPath = sDownloadPath & ChannelRootDir & "download.asp?softid=" & Rs("softid") & "&amp;down=" & i & "|"
				Next
			Else
				sDownloadName = ""
				sDownloadPath = ""
			End If
			'---- 将下载服务器里面的下载地址和软件下载地址连成一个数组
			DownloadPath = Trim(DownloadPath & sDownloadPath)
			DownloadName = Trim(DownloadName & sDownloadName)
			DownloadPath = Split(DownloadPath, "|")
			DownloadName = Split(DownloadName, "|")
			'---- 得出下载地址数组中的维数
			If UBound(DownloadName) < UBound(DownloadPath) Then
				AddressNum = UBound(DownloadName) - 1
			Else
				AddressNum = UBound(DownloadPath) - 1
			End If
			'--- 开始重新排列下载地址直接显示到页面
			SoftNameStr = Rs("SoftName") & " " & Rs("SoftVer")
			ordinal = 1
			i = 0
			For i = 0 To AddressNum
				DownAddress = DownAddress & enchiasp.HtmlSetting(21)
				DownAddress = Replace(DownAddress, "{$Ordinal}", ordinal)
				DownAddress = Replace(DownAddress, "{$DownLoadUrl}", DownloadPath(i))
				DownAddress = Replace(DownAddress, "{$DownLoadName}", DownloadName(i))
				DownAddress = Replace(DownAddress, "{$SoftName}", SoftNameStr)
				ordinal = ordinal + 1
				If i >= AddressNum Then Exit For
			Next
		Else
			SoftNameStr = Rs("SoftName") & " " & Rs("SoftVer")
			strDownAddress = ChannelRootDir & "softdown.asp?softid=" & softid
			DownAddress = enchiasp.HtmlSetting(27)
			DownAddress = Replace(DownAddress, "{$ChannelRootDir}", ChannelRootDir)
			DownAddress = Replace(DownAddress, "{$InstallDir}", enchiasp.InstallDir)
			DownAddress = Replace(DownAddress, "{$SoftName}", SoftNameStr)
			DownAddress = Replace(DownAddress, "{$SoftID}", softid)
			DownAddress = Replace(DownAddress, "{$DownLoadUrl}", strDownAddress)
		End If
		If enchiasp.CheckNull(DownAddress) Then
			ShowDownAddress = DownAddress
		Else
			ShowDownAddress = enchiasp.HtmlSetting(22)
		End If
		
	End Function
	'=================================================
	'函数名：FrontSoft
	'作  用：显示上一软件
	'参  数：SoftID ----软件ID
	'=================================================
	Private Function FrontSoft(softid)
		Dim rsContext, SQL, HtmlFileUrl, HtmlFileName
		
		SQL = "SELECT TOP 1 A.SoftID,A.ClassID,A.SoftName,A.SoftVer,A.HtmlFileDate,C.HtmlFileDir,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.SoftID < " & softid & " ORDER BY A.SoftID DESC"
		Set rsContext = enchiasp.Execute(SQL)
		If rsContext.EOF And rsContext.BOF Then
			FrontSoft = "已经没有了"
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & rsContext("HtmlFileDir") & enchiasp.ShowDatePath(rsContext("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(rsContext("HtmlFileDate"), rsContext("SoftID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				FrontSoft = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & rsContext("SoftName") & " " & rsContext("SoftVer") & "</a>"
			Else
				FrontSoft = "<a href=?id=" & rsContext("SoftID") & ">" & rsContext("SoftName") & " " & rsContext("SoftVer") & "</a>"
			End If
		End If
		rsContext.Close
		Set rsContext = Nothing
	End Function
	'=================================================
	'函数名：NextSoft
	'作  用：显示下一软件
	'参  数：SoftID ----软件ID
	'=================================================
	Private Function NextSoft(softid)
		Dim rsContext, SQL, HtmlFileUrl, HtmlFileName
		
		SQL = "SELECT TOP 1 A.SoftID,A.ClassID,A.SoftName,A.SoftVer,A.HtmlFileDate,C.HtmlFileDir,C.UseHtml from [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.SoftID > " & softid & " ORDER BY A.SoftID ASC"
		Set rsContext = enchiasp.Execute(SQL)
		If rsContext.EOF And rsContext.BOF Then
			NextSoft = "已经没有了"
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & rsContext("HtmlFileDir") & enchiasp.ShowDatePath(rsContext("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(rsContext("HtmlFileDate"), rsContext("SoftID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				NextSoft = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & rsContext("SoftName") & " " & rsContext("SoftVer") & "</a>"
			Else
				NextSoft = "<a href=?id=" & rsContext("SoftID") & ">" & rsContext("SoftName") & " " & rsContext("SoftVer") & "</a>"
			End If
		End If
		rsContext.Close
		Set rsContext = Nothing
	End Function
	Private Function ReadPagination(n)
		Dim HtmlFileName, CurrentPage
		
		CurrentPage = n
		HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("SoftID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, CurrentPage)
		ReadPagination = HtmlFileName
	End Function
	'=================================================
	'函数名：RelatedSoft
	'作  用：显示相关软件
	'参  数：sRelated ----相关软件
	'=================================================
	Private Function RelatedSoft(sRelated, topic, softid)
		Dim rsRdlated, SQL, HtmlFileUrl, HtmlFileName
		Dim strSoftName, softname, strContent
		Dim strRelated, arrRelated, i, Resize, strRearrange
		Dim strKey
		Dim ArrayTemp()
		
		On Error Resume Next
		strRelated = Replace(Replace(Replace(Replace(sRelated, "[", ""), "]", ""), "'", ""), "%", "")
		strKey = Left(enchiasp.ChkQueryStr(topic), 5)
		If Not IsNull(sRelated) And sRelated <> Empty Then
			If InStr(strRelated, "|") > 1 Then
				arrRelated = Split(strRelated, "|")
				strRelated = "((A.SoftName like '%" & arrRelated(0) & "%')"
				For i = 1 To UBound(arrRelated)
					strRelated = strRelated & " Or (A.SoftName like '%" & arrRelated(i) & "%')"
				Next
				'strRelated = strRelated & ")"
			Else
				strRelated = "((A.SoftName like '%" & strRelated & "%')"
			End If
			strRelated = strRelated & " Or (A.SoftName like '%" & strKey & "%'))"
		Else
			strRelated = "(A.SoftName like '%" & strKey & "%')"
		End If
		SQL = "SELECT Top " & CInt(enchiasp.HtmlSetting(1)) & " A.SoftID,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.AllHits,A.SoftTime,A.HtmlFileDate,C.HtmlFileDir,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.SoftID <> " & softid & " And " & strRelated & " ORDER BY A.SoftID DESC"
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
				strSoftName = rsRdlated("SoftName") & " " & rsRdlated("SoftVer")
				strSoftName = enchiasp.GotTopic(strSoftName, CInt(enchiasp.HtmlSetting(2)))
				strSoftName = enchiasp.ReadFontMode(strSoftName, rsRdlated("ColorMode"), rsRdlated("FontMode"))
				If CreateHtml <> 0 Then
					HtmlFileUrl = ChannelRootDir & rsRdlated("HtmlFileDir") & enchiasp.ShowDatePath(rsRdlated("HtmlFileDate"), enchiasp.HtmlPath)
					HtmlFileName = enchiasp.ReadFileName(rsRdlated("HtmlFileDate"), rsRdlated("SoftID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
					softname = "<a href=" & HtmlFileUrl & HtmlFileName & " title='" & rsRdlated("SoftName") & rsRdlated("SoftVer") & "'>" & strSoftName & "</a>"
				Else
					softname = "<a href=show.asp?id=" & rsRdlated("SoftID") & " title='" & rsRdlated("SoftName") & rsRdlated("SoftVer") & "'>" & strSoftName & "</a>"
				End If
				strContent = Replace(strContent, "{$SoftName}", softname)
				strContent = Replace(strContent, "{$AllHits}", rsRdlated("AllHits"))
				strContent = Replace(strContent, "{$WriteTime}", enchiasp.ShowDateTime(rsRdlated("SoftTime"), CInt(enchiasp.HtmlSetting(3))))
				strContent = Replace(strContent, "{$DateTime}", enchiasp.ShowDateTime(rsRdlated("SoftTime"), CInt(enchiasp.HtmlSetting(3))))
				ArrayTemp(i) = strContent
				rsRdlated.MoveNext
				i = i + 1
			Loop
		End If
		rsRdlated.Close
		Set rsRdlated = Nothing
		strRearrange = Join(ArrayTemp, vbCrLf)
		RelatedSoft = strRearrange
	End Function
	'=================================================
	'函数名：ShowHotSoft
	'作  用：显示热门软件
	'参  数：ClassID ----软件分类ID
	'=================================================
	Private Function ShowHotSoft(ClassID)
		Dim rsHot, SQL, HtmlFileUrl, HtmlFileName
		Dim strSoftName, softname, strContent
		Dim i, Resize, strRearrange
		Dim ArrayTemp()
		
		On Error Resume Next
		SQL = "SELECT Top " & CInt(enchiasp.HtmlSetting(1)) & " A.SoftID,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.AllHits,A.SoftTime,A.HtmlFileDate,C.HtmlFileDir,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.AllHits >= " & CLng(enchiasp.LeastHotHist) & " And A.ClassID  in (" & Rs("ChildStr") & ") ORDER BY A.AllHits DESC,A.SoftID DESC"
		Set rsHot = enchiasp.Execute(SQL)
		If rsHot.EOF And rsHot.BOF Then
			ShowHotSoft = ""
			Set rsHot = Nothing
			Exit Function
		Else
			i = 0
			Resize = 0
			Do While Not rsHot.EOF
				ReDim Preserve ArrayTemp(i + Resize)
				strContent = ArrayTemp(i) & enchiasp.HtmlSetting(4)
				strSoftName = rsHot("SoftName") & " " & rsHot("SoftVer")
				strSoftName = enchiasp.GotTopic(rsHot("SoftName"), CInt(enchiasp.HtmlSetting(2)))
				strSoftName = enchiasp.ReadFontMode(strSoftName, rsHot("ColorMode"), rsHot("FontMode"))
				If CreateHtml <> 0 Then
					HtmlFileUrl = ChannelRootDir & rsHot("HtmlFileDir") & enchiasp.ShowDatePath(rsHot("HtmlFileDate"), enchiasp.HtmlPath)
					HtmlFileName = enchiasp.ReadFileName(rsHot("HtmlFileDate"), rsHot("SoftID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
					softname = "<a href=" & HtmlFileUrl & HtmlFileName & " title='" & rsHot("SoftName") & "'>" & strSoftName & "</a>"
				Else
					softname = "<a href=show.asp?id=" & rsHot("SoftID") & " title='" & rsHot("SoftName") & "'>" & strSoftName & "</a>"
				End If
				strContent = Replace(strContent, "{$SoftName}", softname)
				strContent = Replace(strContent, "{$AllHits}", rsHot("AllHits"))
				strContent = Replace(strContent, "{$WriteTime}", enchiasp.ShowDateTime(rsHot("SoftTime"), CInt(enchiasp.HtmlSetting(3))))
				ArrayTemp(i) = strContent
				rsHot.MoveNext
				i = i + 1
			Loop
		End If
		rsHot.Close
		Set rsHot = Nothing
		strRearrange = Join(ArrayTemp, vbCrLf)
		ShowHotSoft = strRearrange
	End Function
	'================================================
	'函数名：SoftComment
	'作  用：软件评论
	'参  数：SoftID ----软件ID
	'================================================
	Private Function SoftComment(softid)
		Dim rsComment, SQL, strContent, strComment
		Dim i, Resize, strRearrange
		Dim ArrayTemp()
		
		On Error Resume Next
		Set rsComment = enchiasp.Execute("SELECT TOP " & CInt(enchiasp.HtmlSetting(5)) & " content,Grade,username,postime,postip FROM ECCMS_Comment WHERE ChannelID=" & ChannelID & " And postid = " & softid & " ORDER BY postime DESC,CommentID DESC")
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
		SoftComment = strRearrange
	End Function
	'================================================
	'函数名：CurrentStation
	'作  用：当前位置
	'参  数：...
	'================================================
	Public Function CurrentStation(ChannelID, ClassID, ClassName, ParentID, strParent, HtmlFileDir, Compart)
		Dim rsCurrent, SQL, strContent, ChannelDir
		
		ChannelDir = ChannelRootDir
		On Error Resume Next
		If ParentID <> 0 And Len(strParent) <> 0 Then
			SQL = "SELECT ClassID,ClassName,HtmlFileDir,UseHtml FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID in(" & strParent & ")"
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
	Public Function ReadCurrentStation(str, ChannelID, ClassID, ClassName, ParentID, strParent, HtmlFileDir)
		Dim strTemp, i, sTempContent, nTempContent
		Dim arrTempContent, arrTempContents
		
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

	'#############################\\执行软件列表开始//#############################
	Public Sub ShowDownList()
		On Error Resume Next
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
				CurrentPage = enchiasp.ChkNumeric(Request("page"))
			Else
				CurrentPage = 1
			End If
			ClassID = enchiasp.ChkNumeric(Request("ClassID"))
			Response.Write CreateSoftList(ClassID, 1)
		End If
		
	End Sub
	'================================================
	'函数名：ReadSoftList
	'作  用：读取软件列表
	'================================================
	Public Function CreateSoftList(clsid, n)
		On Error Resume Next
		Dim rsClass, TemplateContent, strTemplate, strOrder
		Dim ParentTemplate, ChildTemplate, HtmlFileName
		Dim MaxListnum, strMaxListop, showtree
		
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
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ClassID}", ClassID)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", strClassName)
		HtmlContent = Replace(HtmlContent, "{$SoftIndex}", strIndexName)
		If Child <> 0 And showtree <> 9 Then
			Call LoadParentList
			Call ReplaceContent
			If CInt(CreateHtml) <> 0 Then
				'创建分类目录
				enchiasp.CreatPathEx (HtmlFilePath)
				'开始生成父级分类的HTML页
				HtmlFileName = HtmlFilePath & ReadListPageName(ClassID, 0)
				enchiasp.CreatedTextFile HtmlFileName, HtmlContent
				If IsShowFlush = 1 Then
					Response.Write "<li style=""font-size: 16px;"">生成" & enchiasp.ModuleName & "列表HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
					Response.Flush
				End If
			End If
		Else
			Call ReplaceContent
			'每页显示软件数
			maxperpage = enchiasp.ChkNumeric(enchiasp.HtmlSetting(1))
			If CLng(CurrentPage) = 0 Then CurrentPage = 1
			If enchiasp.CheckStr(LCase(Request("oredr"))) = "hits" Then
				strOrder = "ORDER BY A.isTop DESC, A.AllHits DESC ,A.SoftID DESC"
			ElseIf enchiasp.CheckStr(LCase(Request("oredr"))) = "name" Then
				strOrder = "ORDER BY A.isTop DESC, A.SoftName DESC ,A.SoftID DESC"
			ElseIf enchiasp.CheckStr(LCase(Request("oredr"))) = "size" Then
				strOrder = "ORDER BY A.isTop DESC, A.SoftSize DESC ,A.SoftID DESC"
			Else
				strOrder = "ORDER BY A.isTop DESC, A.SoftTime DESC ,A.SoftID DESC"
			End If
			
			TotalNumber = enchiasp.Execute("SELECT COUNT(SoftID) FROM ECCMS_SoftList WHERE ChannelID = " & ChannelID & " And isAccept > 0 And ClassID in (" & ChildStr & ")")(0)
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
			SQL = "SELECT " & strMaxListop & " A.SoftID,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.content,A.Related,A.SoftType,A.RunSystem,A.impower,A.SoftSize,A.star,A.SoftTime,A.username,A.IsTop,A.IsBest,A.Allhits,A.SoftImage,A.HtmlFileDate,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.ClassID in (" & ChildStr & ") " & strOrder & ""
			Rs.Open SQL, Conn, 1, 1
			If Rs.BOF And Rs.EOF Then
				HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "还没有找到任何" & enchiasp.ModuleName & "")
				HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend", "[/ShowRepetend]", 1), "")
				If CreateHtml <> 0 Then
					enchiasp.CreatPathEx (HtmlFilePath)
					HtmlFileName = HtmlFilePath & ReadListPageName(ClassID, CurrentPage)
					enchiasp.CreatedTextFile HtmlFileName, HtmlContent
					If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "列表HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
					Response.Flush
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
		If CreateHtml = 0 Then CreateSoftList = HtmlContent
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
		HtmlContent = HTML.ReadSoftPic(HtmlContent)
		HtmlContent = HTML.ReadSoftList(HtmlContent)
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
	'作  用：装载父级软件列表
	'================================================
	Private Sub LoadParentList()
		Dim rsClslist, strContent, i, showtree
		Dim ClassUrl, ClassNameStr
		
		showtree = Trim(enchiasp.HtmlSetting(4))
		PageType = 1
		On Error Resume Next
		TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
		If Not IsNull(TempListContent) Then
			SQL = "SELECT TOP " & CInt(enchiasp.HtmlSetting(5)) & " ClassID,ClassName,HtmlFileDir,UseHtml FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And TurnLink = 0 And ParentID=" & ClassID & " ORDER BY rootid ASC, ClassID ASC"
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
	'作  用：装载子级软件列表HTML
	'================================================
	Private Sub LoadChildListHtml(n)
		Dim HtmlFileName
		Dim Perownum,ii,w
		Perownum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(8))
		
		If IsNull(TempListContent) Then Exit Sub
		'创建分类目录
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
				If Not Response.IsClientConnected Then Response.End
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
			If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & enchiasp.ModuleName & "列表HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
			Response.Flush
		Next
		
	End Sub
	'================================================
	'过程名：LoadChildListAsp
	'作  用：装载子级软件列表ASP
	'================================================
	Private Sub LoadChildListAsp()
		If IsNull(TempListContent) Then Exit Sub
		
		Dim Perownum,ii,w
		Perownum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(8))
		
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
			If Not Response.IsClientConnected Then Response.End
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
	'过程名：LoadListDetail
	'作  用：装载子级软件列表细节
	'================================================
	Private Sub LoadListDetail()
		Dim sTitle, sTopic, softname, ListStyle
		Dim SoftUrl, SoftTime, sClassName, SoftImageUrl, SoftImage
		
		ListContent = ListContent & TempListContent
		If (i Mod 2) = 0 Then
			ListStyle = 1
		Else
			ListStyle = 2
		End If
		If strlen > 0 Then
			sTitle = enchiasp.GotTopic(Rs("SoftName") & " " & Rs("SoftVer"),strlen)
		Else
			sTitle = Rs("SoftName") & " " & Rs("SoftVer")
		End If
		sTitle = enchiasp.ReadFontMode(sTitle, Rs("ColorMode"), Rs("FontMode"))
		On Error Resume Next
		If CInt(CreateHtml) <> 0 Then
			SoftUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath) & ReadPagination(0)
			sClassName = ChannelRootDir & Rs("HtmlFileDir")
		Else
			SoftUrl = ChannelRootDir & "show.asp?id=" & Rs("SoftID")
			sClassName = ChannelRootDir & "list.asp?classid=" & Rs("ClassID")
		End If
		SoftImageUrl = enchiasp.GetImageUrl(Rs("SoftImage"), enchiasp.ChannelDir)
		SoftImage = enchiasp.GetFlashAndPic(SoftImageUrl, CInt(enchiasp.HtmlSetting(6)), CInt(enchiasp.HtmlSetting(7)))
		SoftImage = "<a href='" & SoftUrl & "' title='" & Rs("SoftName") & "'>" & SoftImage & "</a>"
		sClassName = "<a href='" & sClassName & "' title='" & Rs("ClassName") & "'>" & Rs("ClassName") & "</a>"
		softname = "<a href='" & SoftUrl & "' title='" & Rs("SoftName") & "' class=showtopic>" & sTitle & "</a>"
		
		SoftIntro = enchiasp.CutString(Rs("Content"), CInt(enchiasp.HtmlSetting(3)))
		
		SoftTime = enchiasp.ShowDateTime(Rs("SoftTime"), CInt(enchiasp.HtmlSetting(2)))
		ListContent = Replace(ListContent, "{$ClassifyName}", sClassName)
		ListContent = Replace(ListContent, "{$SoftName}", softname)
		ListContent = Replace(ListContent, "{$SoftTopic}", sTitle)
		ListContent = Replace(ListContent, "{$SoftUrl}", SoftUrl)
		ListContent = Replace(ListContent, "{$SoftImage}", SoftImage)
		ListContent = Replace(ListContent, "{$SoftHits}", Rs("AllHits"))
		ListContent = Replace(ListContent, "{$UserName}", Rs("username"))
		ListContent = Replace(ListContent, "{$Star}", Rs("star"))
		ListContent = Replace(ListContent, "{$SoftDateTime}", SoftTime)
		ListContent = Replace(ListContent, "{$SoftContent}", SoftIntro)
		ListContent = Replace(ListContent, "{$ListStyle}", ListStyle)
		ListContent = Replace(ListContent, "{$SoftSize}", ReadSoftsize(Rs("SoftSize")))
		ListContent = Replace(ListContent, "{$RunSystem}", Rs("RunSystem"))
		ListContent = Replace(ListContent, "{$Impower}", Rs("impower"))
		ListContent = Replace(ListContent, "{$SoftType}", Rs("SoftType"))
		ListContent = Replace(ListContent, "{$IsTop}", Rs("IsTop"))
		ListContent = Replace(ListContent, "{$IsBest}", Rs("IsBest"))
		ListContent = Replace(ListContent, "{$Order}", j)
		ListContent = Replace(ListContent, "{$PageID}", CurrentPage)
	End Sub

	Public Function ASPCurrentPage(stype)
		Dim CurrentUrl
		Select Case stype
			Case "1"
				CurrentUrl = "&amp;classid=" & Trim(Request("classid")) & "&amp;order=" & Trim(Request("order"))
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
	
	Private Function ReadListPageName(ClassID, CurrentPage)
		ReadListPageName = enchiasp.ClassFileName(ClassID, enchiasp.HtmlExtName, enchiasp.HtmlPrefix, CurrentPage)
	End Function
	'#############################\\执行专题软件开始//#############################
	Public Sub ShowDownSpecial()
		On Error Resume Next
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
			SpecialID = enchiasp.ChkNumeric(Request("sid"))
			Response.Write CreateSoftSpecial(SpecialID, 1)
		End If
	End Sub
	Public Function CreateSoftSpecial(sid, n)
		On Error Resume Next
		Dim rsPecial
		Dim HtmlFileName
		
		PageType = 2
		If Not IsNumeric(SpecialID) Then Exit Function
		Set rsPecial = enchiasp.Execute("SELECT SpecialID,SpecialName,SpecialDir FROM [ECCMS_Special] WHERE ChannelID = " & ChannelID & " And SpecialID=" & sid)
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
		HtmlContent = Replace(HtmlContent, "{$SoftIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$SpecialName}", SpecialName)
		Call ReplaceString
		maxperpage = CInt(enchiasp.HtmlSetting(1))

		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		'记录总数
		TotalNumber = enchiasp.Execute("SELECT COUNT(SoftID) from ECCMS_SoftList WHERE ChannelID = " & ChannelID & " And isAccept > 0 And SpecialID = " & SpecialID)(0)
		TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT A.SoftID,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.content,A.Related,A.SoftType,A.RunSystem,A.impower,A.SoftSize,A.star,A.SoftTime,A.username,A.IsTop,A.IsBest,A.Allhits,A.SoftImage,A.HtmlFileDate,C.ClassName,C.HtmlFileDir,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.SpecialID = " & SpecialID & " ORDER BY A.isTop DESC, A.SoftTime DESC ,A.SoftID DESC"
		Rs.Open SQL, Conn, 1, 1
		If Rs.BOF And Rs.EOF Then
			'如果没有找到相关内容,清除掉无用的标签代码
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "还没有找到任何专题" & enchiasp.ModuleName & "")
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
			'如果是生成HTML,执行下面的语句
			If CreateHtml <> 0 Then
				HtmlFileName = HtmlFilePath & enchiasp.SpecialFileName(SpecialID, enchiasp.HtmlExtName, 1)
				enchiasp.CreatedTextFile HtmlFileName, HtmlContent
				If IsShowFlush = 1 Then 
					Response.Write "<li style=""font-size: 12px;"">生成" & strClassName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
					Response.Flush
				End If
			End If
		Else
			'获取模板标签[ShowRepetend][/ReadSoftList]中的字符串
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			If CreateHtml <> 0 Then
				Call LoadSoftListHtml(n)
			Else
				Call LoadChildListAsp
			End If
		End If
		Rs.Close: Set Rs = Nothing
		If CreateHtml = 0 Then CreateSoftSpecial = HtmlContent
		Exit Function
	End Function
	'================================================
	'过程名：LoadSoftListHtml
	'作  用：装载软件列表并生成HTML
	'================================================
	Private Sub LoadSoftListHtml(n)
		Dim HtmlFileName
		
		If IsNull(TempListContent) Then Exit Sub
		For CurrentPage = n To TotalPageNum
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			'Dim bookmark:bookmark = Rs.bookmark
			ListContent = ""
			j = (CurrentPage - 1) * maxperpage + 1
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.End
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
			If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & strClassName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
			Response.Flush
		Next
		Exit Sub
	End Sub
	'================================================
	'过程名：ReplaceString
	'作  用：替换模板内容
	'================================================
	Private Sub ReplaceString()
		HtmlContent = Replace(HtmlContent, "{$SelectedType}", "")
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = HTML.ReadSoftPic(HtmlContent)
		HtmlContent = HTML.ReadSoftList(HtmlContent)
		HtmlContent = HTML.ReadSoftPicAndText(HtmlContent)
		HtmlContent = HTML.ReadPopularSoft(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
	End Sub
	'#############################\\执行推荐软件开始//#############################
	'================================================
	'过程名：ShowBestDown
	'作  用：显示推荐下载
	'================================================
	Public Sub ShowBestDown()
		On Error Resume Next
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
			Response.Write CreateBestDown(1)
		End If
	End Sub
	'================================================
	'过程名：ShowNewDown
	'作  用：显示最新下载
	'================================================
	Public Sub ShowNewDown()
		On Error Resume Next
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
			Response.Write CreateBestDown(0)
		End If
	End Sub
	'================================================
	'过程名：ShowSoftType
	'作  用：显示软件类型
	'================================================
	Public Sub ShowSoftType()
		On Error Resume Next
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
		SoftType = enchiasp.CheckStr(Request("type"))
		Response.Write CreateBestDown(2)
	End Sub
	'================================================
	'过程名：ShowSoftType
	'作  用：显示软件类型
	'================================================
	Public Sub ShowHotDownload()
		On Error Resume Next
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
		SoftType = enchiasp.CheckStr(Request("type"))
		Response.Write CreateBestDown(3)
	End Sub
	'================================================
	'过程名：CreateBestDown
	'作  用：最新推荐下载列表
	'================================================
	Public Function CreateBestDown(t)
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
		HtmlContent = Replace(HtmlContent, "{$SoftIndex}", strIndexName)
		'HtmlContent = Replace(HtmlContent, "{$PageTitle}", "推荐" & enchiasp.ModuleName)
		If CInt(t) = 1 Then
			strClassName = enchiasp.HtmlSetting(9)
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.HtmlSetting(9))
			PageType = 3
			SQL1 = "And IsBest>0"
			SQL2 = "And A.IsBest>0 ORDER BY A.isTop DESC, A.SoftTime DESC ,A.SoftID DESC"
		ElseIf CInt(t) = 2 Then
			Dim Channel_Setting, strSoftType, SelectedType, i
			Channel_Setting = Split(enchiasp.Channel_Setting, "|||")(2)
			strSoftType = Split(Channel_Setting, ",")
			SelectedType = "<select name=""type"" size=""1"" style=""font-size: 9pt"" onChange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbNewLine
			SelectedType = SelectedType & "<option value=""showtype.asp"">全部" & enchiasp.ModuleName & "</option>" & vbNewLine
			For i = 0 To UBound(strSoftType)
				SelectedType = SelectedType & "<option value=""showtype.asp?type=" & strSoftType(i) & """"
				If Trim(strSoftType(i)) = Trim(SoftType) Then SelectedType = SelectedType & " selected"
				SelectedType = SelectedType & ">" & strSoftType(i) & "</option>" & vbNewLine
			Next
			SelectedType = SelectedType & "<select>" & vbNewLine
			HtmlContent = Replace(HtmlContent, "{$SelectedType}", SelectedType)
			If Trim(SoftType) <> "" Then
				strClassName = SoftType
				HtmlContent = Replace(HtmlContent, "{$PageTitle}", SoftType)
				PageType = 6
				SQL1 = "And SoftType='" & SoftType & "'"
				SQL2 = "And A.SoftType='" & SoftType & "' ORDER BY A.isTop DESC, A.SoftTime DESC ,A.SoftID DESC"
			Else
				strClassName = "全部" & enchiasp.ModuleName & "类型"
				HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ModuleName & "类型")
				PageType = 6
				SQL1 = ""
				SQL2 =  " ORDER BY A.isTop DESC, A.SoftTime DESC ,A.SoftID DESC"
			End If
		ElseIf CInt(t) = 3 Then
			strClassName = enchiasp.HtmlSetting(10)
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.HtmlSetting(10))
			PageType = 3
			SQL1 = "And AllHits > " & CLng(enchiasp.LeastHotHist)
			SQL2 = "And A.AllHits > " & CLng(enchiasp.LeastHotHist) & " ORDER BY A.AllHits DESC, A.SoftTime DESC ,A.SoftID DESC"
		Else
			strClassName = enchiasp.HtmlSetting(8)
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.HtmlSetting(8))
			PageType = 3
			SQL1 = vbNullString
			SQL2 = "ORDER BY A.isTop DESC, A.SoftTime DESC ,A.SoftID DESC"
		End If
		Call ReplaceString
		maxperpage = CLng(enchiasp.HtmlSetting(1))
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		'记录总数
		TotalNumber = enchiasp.Execute("SELECT COUNT(SoftID) FROM ECCMS_SoftList WHERE ChannelID = " & ChannelID & " And isAccept > 0  " & SQL1 & "")(0)
		If TotalNumber >= CLng(enchiasp.HtmlSetting(4)) Then TotalNumber = CLng(enchiasp.HtmlSetting(4))
		TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT TOP " & CLng(enchiasp.HtmlSetting(4)) & " A.SoftID,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.content,A.Related,A.SoftType,A.RunSystem,A.impower,A.SoftSize,A.star,A.SoftTime,A.username,A.IsTop,A.IsBest,A.Allhits,A.SoftImage,A.HtmlFileDate,C.ClassName,C.HtmlFileDir,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 " & SQL2
		Rs.Open SQL, Conn, 1, 1
		
		If Rs.BOF And Rs.EOF Then
			'如果没有找到相关内容,清除掉无用的标签代码
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "还没有找到任何推荐" & enchiasp.ModuleName & "")
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
			'如果是生成HTML,执行下面的语句
			If CreateHtml <> 0 And CInt(t) <> 2 Then
				If CInt(t) = 1 Then
					HtmlFileName = HtmlFilePath & enchiasp.HtmlPrefix & "Best001" & enchiasp.HtmlExtName
				ElseIf CInt(t) = 3 Then
					HtmlFileName = HtmlFilePath & enchiasp.HtmlPrefix & "Hot001" & enchiasp.HtmlExtName
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
			'获取模板标签[ShowRepetend][/ReadSoftList]中的字符串
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			If CreateHtml <> 0 And CInt(t) <> 2 Then
				Call LoadBestSoftListHtml(t)
			Else
				Call LoadChildListAsp
			End If
		End If
		Rs.Close: Set Rs = Nothing
		If CreateHtml = 0 Or CInt(t) = 2 Then CreateBestDown = HtmlContent
	End Function
	'================================================
	'过程名：LoadBestSoftListHtml
	'作  用：装载软件列表并生成HTML
	'================================================
	Private Sub LoadBestSoftListHtml(t)
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
				If Not Response.IsClientConnected Then Response.End
				Call LoadListDetail
				Rs.MoveNext
				i = i + 1
				j = j + 1
				If i >= maxperpage Then Exit Do
			Loop
			Dim strHtmlFront, strHtmlPage
			If t = 1 Then
				sulCurrentPage = enchiasp.HtmlPrefix & "Best"
			ElseIf t = 3 Then
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
			If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & strClassName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
			Response.Flush
		Next
	End Sub
	'#############################\\执行热门软件开始//#############################
	'================================================
	'过程名：ShowHotDown
	'作  用：显示最新下载
	'================================================
	Public Sub ShowHotDown()
		On Error Resume Next

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
			Response.Write CreateHotDown
		End If
	End Sub
	Public Function CreateHotDown()
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
		HtmlContent = Replace(HtmlContent, "{$SoftIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", "下载排行")
		strClassName = "下载排行"
		Call ReplaceString
		maxperpage = CInt(enchiasp.HtmlSetting(1))
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		'记录总数
		TotalNumber = enchiasp.Execute("SELECT COUNT(SoftID) FROM ECCMS_SoftList WHERE ChannelID = " & ChannelID & " And isAccept > 0 And AllHits > " & CLng(enchiasp.LeastHotHist) & "")(0)
		If TotalNumber >= CLng(enchiasp.HtmlSetting(4)) Then TotalNumber = CLng(enchiasp.HtmlSetting(4))
		TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT TOP " & CLng(enchiasp.HtmlSetting(4)) & " A.SoftID,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.content,A.Related,A.SoftType,A.RunSystem,A.impower,A.SoftSize,A.star,A.SoftTime,A.username,A.IsTop,A.IsBest,A.Allhits,A.SoftImage,A.HtmlFileDate,C.ClassName,C.HtmlFileDir,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.AllHits > " & CLng(enchiasp.LeastHotHist) & " ORDER BY A.Allhits DESC, A.SoftTime DESC ,A.SoftID DESC"
		Rs.Open SQL, Conn, 1, 1
		
		If Rs.BOF And Rs.EOF Then
			'如果没有找到相关内容,清除掉无用的标签代码
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "还没有找到任何热门" & enchiasp.ModuleName & "")
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
			'如果是生成HTML,执行下面的语句
			If CreateHtml <> 0 Then
				HtmlFileName = HtmlFilePath & enchiasp.HtmlPrefix & "Hot001" & enchiasp.HtmlExtName
				enchiasp.CreatedTextFile HtmlFileName, HtmlContent
				If IsShowFlush = 1 Then 
					Response.Write "<li style=""font-size: 12px;"">生成" & strClassName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
					Response.Flush
				End If
			End If
		Else
			'获取模板标签[ShowRepetend][/ReadSoftList]中的字符串
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			If CreateHtml <> 0 Then
				Call LoadHotSoftListHtml
			Else
				Call LoadChildListAsp
			End If
		End If
		Rs.Close: Set Rs = Nothing
		If CreateHtml = 0 Then CreateHotDown = HtmlContent
	End Function
	'================================================
	'过程名：LoadHotSoftListHtml
	'作  用：装载软件列表并生成HTML
	'================================================
	Private Sub LoadHotSoftListHtml()
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
				If Not Response.IsClientConnected Then Response.End
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
			If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">生成" & strClassName & "HTML完成... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
			Response.Flush
		Next
		Exit Sub
	End Sub
	'#############################\\软件搜索开始//#############################
	Public Sub ShowDownSearch()
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
				findword = "A.SoftName like '%" & keyword & "%'"
			ElseIf LCase(Action) = "content" Then
				If CInt(enchiasp.FullContQuery) <> 0 Then
					findword = "A.Content like '%" & keyword & "%'"
				Else
					Call OutAlertScript(Replace(Replace(enchiasp.HtmlSetting(10), Chr(34), "\"""), vbCrLf, ""))
					Exit Sub
				End If
			Else
				findword = "A.SoftName like '%" & keyword & "%'"
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
		HtmlContent = Replace(HtmlContent, "{$SoftIndex}", strIndexName)
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
			SQL = "SELECT TOP " & SearchMaxPageList & " A.SoftID,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.content,A.Related,A.Languages,A.SoftType,A.RunSystem,A.impower,A.Contact,A.SoftSize,A.star,A.SoftTime,A.username,A.Allhits,A.HtmlFileDate,C.ClassName,C.HtmlFileDir,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And " & findword & " ORDER BY A.SoftTime DESC ,A.SoftID DESC"
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
				'获取模板标签[ShowRepetend][/ReadSoftList]中的字符串
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
		Dim sTitle, sTopic, softname, ListStyle, TitleWord
		Dim SoftUrl, SoftTime, sClassName, SoftImage, SoftVer
		
		ListContent = ListContent & TempListContent
		If (i Mod 2) = 0 Then
			ListStyle = 1
		Else
			ListStyle = 2
		End If
		TitleWord = Replace(Rs("SoftName"), "" & keyword & "", "<font color=red>" & keyword & "</font>")
		sTitle = enchiasp.ReadFontMode(TitleWord, Rs("ColorMode"), Rs("FontMode")) & " " & Rs("SoftVer")
		
		If CInt(CreateHtml) <> 0 Then
			SoftUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath) & ReadPagination(0)
			sClassName = ChannelRootDir & Rs("HtmlFileDir")
		Else
			SoftUrl = ChannelRootDir & "show.asp?id=" & Rs("SoftID")
			sClassName = ChannelRootDir & "list.asp?classid=" & Rs("ClassID")
		End If
		
		sClassName = "<a href=""" & sClassName & """ title=""" & Rs("ClassName") & """ target=""_blank""><span style=""color:" & enchiasp.MainSetting(3) & """>" & Rs("ClassName") & "</span></a>"
		softname = "<a href='" & SoftUrl & "' title='" & Rs("SoftName") & Rs("SoftVer") & "' class=showtopic target=""_blank"">" & sTitle & "</a>"
		SoftIntro = enchiasp.CutString(Rs("Content"), CInt(enchiasp.HtmlSetting(3)))
		SoftIntro = Replace(SoftIntro, "" & keyword & "", "<font color=red>" & keyword & "</font>")
		
		SoftTime = enchiasp.ShowDateTime(Rs("SoftTime"), CInt(enchiasp.HtmlSetting(2)))
		ListContent = Replace(ListContent, "{$KeyWord}", keyword)
		ListContent = Replace(ListContent, "{$totalrec}", TotalNumber)
		ListContent = Replace(ListContent, "{$ClassifyName}", sClassName)
		ListContent = Replace(ListContent, "{$SoftName}", softname)
		ListContent = Replace(ListContent, "{$SoftTopic}", sTitle)
		ListContent = Replace(ListContent, "{$SoftUrl}", SoftUrl)
		ListContent = Replace(ListContent, "{$Star}", Rs("star"))
		ListContent = Replace(ListContent, "{$SoftHits}", Rs("AllHits"))
		ListContent = Replace(ListContent, "{$UserName}", Rs("username"))
		ListContent = Replace(ListContent, "{$SoftDateTime}", SoftTime)
		ListContent = Replace(ListContent, "{$SoftContent}", SoftIntro)
		ListContent = Replace(ListContent, "{$ListStyle}", ListStyle)
		ListContent = Replace(ListContent, "{$SoftSize}", ReadSoftsize(Rs("SoftSize")))
		ListContent = Replace(ListContent, "{$RunSystem}", Rs("RunSystem"))
		ListContent = Replace(ListContent, "{$Impower}", Rs("impower"))
		ListContent = Replace(ListContent, "{$Language}", Rs("Languages"))
		ListContent = Replace(ListContent, "{$SoftType}", Rs("SoftType"))
		ListContent = Replace(ListContent, "{$SoftID}", Rs("SoftID"))
		ListContent = Replace(ListContent, "{$Order}", j)
	End Sub
	'================================================
	'函数名：ReadSoftsize
	'作  用：读取软件的大小
	'================================================
	Function ReadSoftsize(ByVal para)
		On Error Resume Next
		Dim strFileSize, parasize
		
		parasize = CCur(para)
		
		If parasize = 0 Then
			ReadSoftsize = "未知"
			Exit Function
		End If

		If parasize > 1024 Then
			strFileSize = Round(parasize / 1024, 2) & " MB"
		Else
			strFileSize = parasize & " KB"
		End If
		ReadSoftsize = strFileSize
	End Function
	'================================================
	'过程名：ShowSoftComment
	'作  用：显示软件评论
	'================================================
	Public Sub ShowSoftComment()
		Dim softname, HtmlFileUrl, HtmlFileName
		Dim AverageGrade, TotalComment, TempListContent
		Dim strComment, strCheckBox, strAdminComment
		enchiasp.PreventInfuse
		strCheckBox = ""
		strAdminComment = ""
		On Error Resume Next
		
		SoftID = enchiasp.ChkNumeric(Request("SoftID"))
		If SoftID = 0 Then
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
		HtmlContent = Replace(HtmlContent, "{$SoftIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ModuleName & "评论")
		HtmlContent = Replace(HtmlContent, "{$SoftID}", softid)
		'获得软件标题
		SQL = "SELECT TOP 1 A.SoftID,A.ClassID,A.SoftName,A.SoftVer,A.HtmlFileDate,A.ForbidEssay,C.HtmlFileDir,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.SoftID = " & softid
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			Response.Write "已经没有了"
			Set Rs = Nothing
			Exit Sub
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("SoftID"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				softname = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & Rs("SoftName") & " " & Rs("SoftVer") & "</a>"
			Else
				softname = "<a href=show.asp?id=" & Rs("SoftID") & ">" & Rs("SoftName") & " " & Rs("SoftVer") & "</a>"
			End If
			ForbidEssay = Rs("ForbidEssay")
		End If
		Rs.Close
		Set Rs = CreateObject("adodb.recordset")
		SQL = "SELECT COUNT(CommentID) As TotalComment,AVG(Grade) As avgGrade FROM ECCMS_Comment WHERE ChannelID=" & ChannelID & " And postid = " & softid
		Set Rs = enchiasp.Execute(SQL)
		TotalComment = Rs("TotalComment")
		AverageGrade = Round(Rs("avgGrade"))
		If IsNull(AverageGrade) Then AverageGrade = 0
		Rs.Close: Set Rs = Nothing
		HtmlContent = Replace(HtmlContent, "{$SoftName}", softname)
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
		SQL = "SELECT * FROM ECCMS_Comment WHERE ChannelID=" & ChannelID & " And postid = " & softid & " ORDER BY postime DESC,CommentID DESC"
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
			strAdminComment = strAdminComment & "<input type=hidden name=SoftID value='" & softid & "'>" & vbNewLine
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
	'作  用：软件评论分页
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
			strTemp = strTemp & "共有评论 <font COLOR=#FF0000>" & TotalNumber & "</font> 个&nbsp;&nbsp;<a href=" & FileName & "?page=1&SoftID=" & Request("SoftID") & ">首 页</a>&nbsp;&nbsp;"
			strTemp = strTemp & "<a href=" & FileName & "?page=" & CurrentPage - 1 & "&SoftID=" & Request("SoftID") & ">上一页</a>&nbsp;&nbsp;&nbsp;"
		End If

		If n - CurrentPage < 1 Then
			strTemp = strTemp & "下一页&nbsp;&nbsp;尾 页 " & vbCrLf
		Else
			strTemp = strTemp & "<a href=" & FileName & "?page=" & (CurrentPage + 1) & "&SoftID=" & Request("SoftID") & ">下一页</a>"
			strTemp = strTemp & "&nbsp;&nbsp;<a href=" & FileName & "?page=" & n & "&SoftID=" & Request("SoftID") & ">尾 页</a>" & vbCrLf
		End If
		strTemp = strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
		strTemp = strTemp & "&nbsp;<b>" & maxperpage & "</b>个/页 " & vbCrLf
		strTemp = strTemp & "</td></tr></table>" & vbCrLf
		HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strTemp)
	End Sub
	'================================================
	'过程名：CommentDel
	'作  用：软件评论删除
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
	'作  用：软件评论添加保存
	'================================================
	Public Sub CommentSave()
		If enchiasp.CheckPost = False Then
			FoundErr = True
			Call OutAlertScript("您提交的数据不合法，请不要从外部提交表单。")
			Exit Sub
		End If
		On Error Resume Next
		Call PreventRefresh
		If CInt(enchiasp.AppearGrade) <> 0 And Session("AdminName") = "" Then
			If CInt(enchiasp.AppearGrade) > CInt(enchiasp.membergrade) Then
				FoundErr = True
				Call OutAlertScript("您没有发表评论的权限，如果您是会员请登陆后再参与评论。")
				Exit Sub
			End If
		End If
		If ForbidEssay <> 0 Then
			FoundErr = True
			Call OutAlertScript("此" & enchiasp.ModuleName & "禁止发表评论！")
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
		If FoundErr = True Then Exit Sub
		softid = enchiasp.ChkNumeric(Request.Form("SoftID"))
		Set Rs = CreateObject("ADODB.RecordSet")
		SQL = "SELECT * FROM ECCMS_Comment WHERE (CommentID is null)"
		Rs.Open SQL, Conn, 1, 3
		Rs.AddNew
			Rs("ChannelID") = ChannelID
			Rs("postid") = softid
			Rs("UserName") = enchiasp.ChkFormStr(Request.Form("UserName"))
			Rs("Grade") = enchiasp.ChkNumeric(Request.Form("Grade"))
			Rs("content") = Server.HTMLEncode(Request.Form("content"))
			Rs("postime") = Now()
			Rs("postip") = enchiasp.GetUserip
		Rs.Update
		Rs.Close: Set Rs = Nothing
		If CreateHtml <> 0 Then ReadSoftIntro (softid)
		Session("UserRefreshTime") = Now()
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
		Exit Sub
	End Sub
	Public Sub PreventRefresh()
		Dim RefreshTime
		
		RefreshTime = 20
		If DateDiff("s", Session("UserRefreshTime"), Now()) < RefreshTime Then
			Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=" & RefreshTime & "><br>本页面起用了防刷新机制，请不要在" & RefreshTime & "秒内连续刷新本页面<BR>正在打开页面，请稍后……"
			FoundErr = True
			Response.End
		End If
	End Sub

End Class
%>