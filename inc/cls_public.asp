<!--#include file="classmenu.asp"-->
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
Dim HTML
Set HTML = New enchiaspPublic_Cls
Class enchiaspPublic_Cls
	
	Private Sub Class_Initialize()
		On Error Resume Next
		enchiasp.LoadTemplates 0, 0, 0
	End Sub
	'================================================
	'函数名：LoadArticleList
	'作  用：装载文章列表
	'参  数：ClassID   ----分类ID
	'        ChannelID   ----频道ID
	'        SpecialID  ----专题ID
	'        sType   ----调用文章类型,0=所有最新文章，1=推荐文章，2=热门文章，3=图文文章，4=分类最新文章
	'        TopNum   ----显示文章列表数
	'        strlen   ----显示标题长度
	'        ShowClass   ----是否显示分类
	'        ShowPic   ----是否显示图文标题
	'        ShowDate   ----是否显示日期
	'        DateMode   ----显示日期模式
	'        newindow   ----新窗口打开
	'================================================
	Public Function LoadArticleList(ByVal ChannelID, ByVal ClassID, ByVal SpecialID, _
		ByVal stype, ByVal TopNum, ByVal strLen, _
		ByVal showclass, ByVal showpic, ByVal showdate, _
		ByVal DateMode, ByVal newindow, ByVal styles)
		
		Dim Rs, SQL, i, strContent, foundstr
		Dim sTitle, sTopic, ChildStr, ListStyle, BestCode, BestString
		Dim ArticleTopic, ClassName, HtmlFileUrl, WriteTime, LinkTarget, HtmlFileName

		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		SpecialID = enchiasp.ChkNumeric(SpecialID)
		stype = enchiasp.ChkNumeric(stype)
		On Error Resume Next
		enchiasp.LoadChannel(ChannelID)
		
		If CInt(stype) >= 4 And CLng(ClassID) <> 0 Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID=" & ChannelID & " And ClassID=" & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadArticleList = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
			End If
			Set Rs = Nothing
		Else
			ChildStr = "0"
		End If
		
		Select Case CInt(stype)
			Case 0: foundstr = "Order By A.Writetime Desc ,A.Articleid Desc"
			Case 1: foundstr = "And A.isBest > 0 Order By A.Writetime Desc ,A.Articleid Desc"
			Case 2: foundstr = "Order By A.AllHits Desc ,A.Articleid Desc"
			Case 3: foundstr = "And (A.BriefTopic = 1 Or A.BriefTopic = 2) Order By A.Writetime Desc ,A.Articleid Desc"
			Case 4: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.Writetime Desc ,A.Articleid Desc"
			Case 5: foundstr = "And A.ClassID in (" & ChildStr & ") And A.isBest > 0 Order By A.Writetime Desc ,A.Articleid Desc"
			Case 6: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.AllHits Desc ,A.Articleid Desc"
			Case 7: foundstr = "And A.ClassID in (" & ChildStr & ") And (A.BriefTopic = 1 Or A.BriefTopic = 2) Order By A.Writetime Desc ,A.Articleid Desc"
		Case Else
			foundstr = "Order By A.Writetime Desc ,A.Articleid Desc"
		End Select
		If CInt(stype) >= 4 And CLng(ClassID) = 0 Then
			foundstr = "Order By A.Writetime Desc ,A.Articleid Desc"
		End If
		If CLng(SpecialID) <> 0 Then
			foundstr = "And A.SpecialID =" & CLng(SpecialID) & " " & foundstr
		End If
		SQL = " A.ArticleID,A.ClassID,A.ColorMode,A.FontMode,A.title,A.BriefTopic,A.AllHits,A.WriteTime,A.HtmlFileDate,A.isBest,"
		SQL = "SELECT Top " & CInt(TopNum) & SQL & " C.ClassName,C.ColorModes,C.FontModes,C.HtmlFileDir,C.UseHtml FROM [ECCMS_Article] A INNER JOIn [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ChannelID=" & ChannelID & " " & foundstr & ""
		Set Rs = enchiasp.Execute(SQL)
		i = 0
		If Rs.BOF And Rs.EOF Then
			strContent = "该分类还没有添加任何内容！"
		Else
			strContent = "<table width=""100%"" border=0 cellpadding=2 cellspacing=0>"
			Do While Not Rs.EOF
				If (i Mod 2) = 0 Then
					ListStyle = Trim(styles) & 1
				Else
					ListStyle = Trim(styles) & 2
				End If
				If Rs("isBest") <> 0 Then
					BestCode = 2
					BestString = "<font color='" & enchiasp.MainSetting(3) & "'>推荐</font>"
				Else
					BestCode = 1
					BestString = ""
				End If
				
				strContent = strContent & enchiasp.MainSetting(13)
				
				sTitle = enchiasp.GotTopic(Rs("title"), CInt(strLen))
				sTitle = enchiasp.ReadFontMode(sTitle, Rs("ColorMode"), Rs("FontMode"))
				sTopic = enchiasp.ReadPicTopic(Rs("BriefTopic"))
				
				ClassName = enchiasp.ReadFontMode(Rs("ClassName"), Rs("ColorModes"), Rs("FontModes"))
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("ArticleID"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) <> 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
					ClassName = "<a href='" & enchiasp.ChannelPath & Rs("HtmlFileDir") & "index" & enchiasp.ChannelHtmlExt & "'>" & ClassName & "</a>"
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("ArticleID")
					ClassName = "<a href='" & enchiasp.ChannelPath & "list.asp?classid=" & Rs("ClassID") & "'>" & ClassName & "</a>"
				End If
				
				If CInt(showclass) = 0 Then ClassName = ""
				If CInt(showpic) = 0 Then sTopic = ""
				If CInt(showdate) <> 0 Then
					WriteTime = enchiasp.ShowDateTime(Rs("WriteTime"), CInt(DateMode))
				Else
					WriteTime = ""
				End If
				If CInt(newindow) <> 0 Then
					LinkTarget = " target=""_blank"""
				Else
					LinkTarget = ""
				End If
				ArticleTopic = "<a href='" & HtmlFileUrl & "'" & LinkTarget & " title='" & enchiasp.ChannelModule & "标题：" & Rs("title") & "&#13;&#10;发布时间：" & Rs("WriteTime") & "&#13;&#10;阅览次数：" & Rs("AllHits") & "' class=showlist>" & sTitle & "</a>"
				strContent = Replace(strContent, "{$ArticleTopic}", ArticleTopic)
				strContent = Replace(strContent, "{$ArticleID}", Rs("ArticleID"))
				strContent = Replace(strContent, "{$InstallDir}", enchiasp.InstallDir)
				strContent = Replace(strContent, "{$ArticleTitle}", sTitle)
				strContent = Replace(strContent, "{$Title}", Rs("title"))
				strContent = Replace(strContent, "{$DateAndTitle}", Rs("WriteTime"))
				strContent = Replace(strContent, "{$BriefTopic}", sTopic)
				strContent = Replace(strContent, "{$HtmlFileUrl}", HtmlFileUrl)
				strContent = Replace(strContent, "{$ClassName}", ClassName)
				strContent = Replace(strContent, "[]", "")
				strContent = Replace(strContent, "{$Target}", LinkTarget)
				strContent = Replace(strContent, "{$WriteTime}", WriteTime)
				strContent = Replace(strContent, "{$AticleHits}", Rs("AllHits"))
				strContent = Replace(strContent, "{$ListStyle}", ListStyle)
				strContent = Replace(strContent, "{$BestCode}", BestCode)
				strContent = Replace(strContent, "{$BestString}", BestString)
			Rs.MoveNext
			i = i + 1
			Loop
			strContent = strContent & "</table>"
		End If
		
		Rs.Close: Set Rs = Nothing
		LoadArticleList = strContent
	End Function
	'================================================
	'函数名：ReadArticleList
	'作  用：读取文章列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadArticleList(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent
		Dim arrTempContent, arrTempContents, ArrayList
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$ReadArticleList(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadArticleList(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadArticleList(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadArticleList(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8), ArrayList(9), ArrayList(10), ArrayList(11)))
			Next
		End If
		ReadArticleList = strTemp
	End Function
	
	
	
	'================================================
	'函数名：LoadSoftList
	'作  用：装载软件列表
	'参  数：ClassID   ----分类ID
	'        ChannelID   ----频道ID
	'        sType   ----调用类型
	'        TopNum   ----显示列表数
	'        strlen   ----显示标题长度
	'        ShowClass   ----是否显示分类
	'        ShowDate   ----是否显示日期
	'        DateMode   ----显示日期模式
	'        newindow   ----新窗口打开
	'================================================
	Public Function LoadSoftList(ByVal ChannelID, ByVal ClassID, ByVal SpecialID, _
		ByVal stype, ByVal TopNum, ByVal strLen, ByVal showclass, _
		ByVal showdate, ByVal DateMode, ByVal newindow, ByVal styles)
		
		Dim Rs, SQL, i, strContent, foundstr,j
		Dim strSoftName, ChildStr, ListStyle
		Dim HtmlFileName, BestCode, BestString,ChannelPath
		Dim ClassName, HtmlFileUrl, SoftTime, LinkTarget, SoftTopic
		
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		SpecialID = enchiasp.ChkNumeric(SpecialID)
		stype = enchiasp.ChkNumeric(stype)
		
		On Error Resume Next
		enchiasp.LoadChannel(ChannelID)
		
		If CInt(stype) >= 3 And CLng(ClassID) <> 0 Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID=" & ChannelID & " And ClassID = " & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadSoftList = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
			End If
			Rs.Close
		Else
			ChildStr = 0
		End If
		Select Case CInt(stype)
			Case 0: foundstr = "Order By A.SoftTime Desc ,A.SoftID Desc"
			Case 1: foundstr = "And A.isBest > 0 Order By A.SoftTime Desc ,A.SoftID Desc"
			Case 2: foundstr = "Order By A.AllHits Desc ,A.SoftID Desc"
			Case 3: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.SoftTime Desc ,A.SoftID Desc"
			Case 4: foundstr = "And A.ClassID in (" & ChildStr & ") And A.isBest > 0 Order By A.SoftTime Desc ,A.SoftID Desc"
			Case 5: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.AllHits Desc ,A.SoftID Desc"
		Case Else
			foundstr = "Order By A.SoftTime Desc ,A.SoftID Desc"
		End Select
		If CInt(stype) >= 3 And CLng(ClassID) = 0 Then
			foundstr = "Order By A.SoftTime Desc ,A.SoftID Desc"
		End If
		If CLng(SpecialID) <> 0 Then
			foundstr = "And A.SpecialID =" & CLng(SpecialID) & " " & foundstr
		End If
		SQL = " A.SoftID,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.AllHits,A.SoftTime,A.HtmlFileDate,A.isBest,"
		SQL = "SELECT TOP " & CInt(TopNum) & SQL & " C.ClassName,C.ColorModes,C.FontModes,C.HtmlFileDir,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ChannelID=" & ChannelID & " " & foundstr & ""
		Set Rs = enchiasp.Execute(SQL)
		j = 0
		
		If Rs.BOF And Rs.EOF Then
			strContent = "没有添加任何软件！"
		Else
			SQL=Rs.GetRows(-1)
			strContent = "<table width=""100%"" border=0 cellpadding=2 cellspacing=0>"
			For i=0 To Ubound(SQL,2)
				If (j Mod 2) = 0 Then
					ListStyle = Trim(styles) & 1
				Else
					ListStyle = Trim(styles) & 2
				End If
				If CInt(SQL(9,i)) <> 0 Then
					BestCode = 2
					BestString = "<font color='" & enchiasp.MainSetting(3) & "'>推荐</font>"
				Else
					BestCode = 1
					BestString = ""
				End If
				strContent = strContent & enchiasp.MainSetting(14)
				strSoftName = enchiasp.GotTopic(SQL(4,i) & " " & SQL(5,i), CInt(strLen))
				strSoftName = enchiasp.ReadFontMode(strSoftName, SQL(2,i), SQL(3,i))
				ClassName = enchiasp.ReadFontMode(SQL(10,i), SQL(11,i), SQL(12,i))
				HtmlFileName = enchiasp.ReadFileName(SQL(8,i), SQL(0,i), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) <> 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & SQL(13,i) & enchiasp.ShowDatePath(SQL(8,i), enchiasp.ChannelHtmlPath) & HtmlFileName
					ClassName = "<a href='" & enchiasp.ChannelPath & SQL(13,i) & "index" & enchiasp.ChannelHtmlExt & "'>" & ClassName & "</a>"
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & SQL(0,i)
					ClassName = "<a href='" & enchiasp.ChannelPath & "list.asp?classid=" & SQL(1,i) & "'>" & ClassName & "</a>"
				End If
				If CInt(showclass) = 0 Then ClassName = ""
				If CInt(showdate) <> 0 Then
					SoftTime = enchiasp.ShowDateTime(SQL(7,i), CInt(DateMode))
				Else
					SoftTime = ""
				End If
				If CInt(newindow) <> 0 Then
					LinkTarget = " target=""_blank"""
				Else
					LinkTarget = ""
				End If
				SoftTopic = "<a href='" & HtmlFileUrl & "'" & LinkTarget & " title='" & enchiasp.ChannelModule & "名称：" & Trim(SQL(4,i) & " " & SQL(5,i)) & "&#13;&#10;发布时间：" & SQL(7,i) & "&#13;&#10;下载次数：" & SQL(6,i) & "' class=showlist>" & strSoftName & "</a>"
				strContent = Replace(strContent, "{$SoftTopic}", SoftTopic)
				strContent = Replace(strContent, "{$SoftID}", Rs("softid"))
				strContent = Replace(strContent, "{$InstallDir}", enchiasp.InstallDir)
				strContent = Replace(strContent, "{$SoftName}", strSoftName)
				strContent = Replace(strContent, "{$Title}", SQL(4,i))
				strContent = Replace(strContent, "{$DateAndTitle}", SQL(7,i))
				strContent = Replace(strContent, "{$HtmlFileUrl}", HtmlFileUrl)
				strContent = Replace(strContent, "{$ClassName}", ClassName)
				strContent = Replace(strContent, "[]", "")
				strContent = Replace(strContent, "{$Target}", LinkTarget)
				strContent = Replace(strContent, "{$SoftTime}", SoftTime)
				strContent = Replace(strContent, "{$SoftHits}", SQL(6,i))
				strContent = Replace(strContent, "{$ListStyle}", ListStyle)
				strContent = Replace(strContent, "{$BestCode}", BestCode)
				strContent = Replace(strContent, "{$BestString}", BestString)
			j = j + 1
			Next
			SQL=Null
			strContent = strContent & "</table>"
		End If
		Rs.Close: Set Rs = Nothing
		LoadSoftList = strContent
	End Function
	'================================================
	'函数名：ReadSoftList
	'作  用：读取软件列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadSoftList(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent
		Dim arrTempContent, arrTempContents, ArrayList
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$ReadSoftList(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadSoftList(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadSoftList(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadSoftList(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8), ArrayList(9), ArrayList(10)))
			Next
		End If
		ReadSoftList = strTemp
	End Function

	'================================================
	'函数名：LoadFlashList
	'作  用：装载动画列表
	'参  数：ClassID   ----分类ID
	'        ChannelID   ----频道ID
	'        sType   ----调用类型
	'        TopNum   ----显示列表数
	'        strlen   ----显示标题长度
	'        ShowClass   ----是否显示分类
	'        ShowDate   ----是否显示日期
	'        DateMode   ----显示日期模式
	'        newindow   ----新窗口打开
	'================================================
	Public Function LoadFlashList(ByVal ChannelID, ByVal ClassID, ByVal SpecialID, _
		ByVal stype, ByVal TopNum, ByVal strLen, ByVal showclass, _
		ByVal showdate, ByVal DateMode, ByVal newindow, ByVal styles)
		
		Dim Rs, SQL, i, strContent, foundstr,j
		Dim strTitle, ChildStr, ListStyle
		Dim HtmlFileName, BestCode, BestString,ChannelPath
		Dim ClassName, HtmlFileUrl, addTime, LinkTarget, FlashTopic
		
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		SpecialID = enchiasp.ChkNumeric(SpecialID)
		stype = enchiasp.ChkNumeric(stype)
		
		On Error Resume Next
		enchiasp.LoadChannel(ChannelID)
		
		If CInt(stype) >= 3 And CLng(ClassID) <> 0 Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID=" & ChannelID & " And ClassID = " & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadFlashList = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
			End If
			Rs.Close
		Else
			ChildStr = 0
		End If
		Select Case CInt(stype)
			Case 0: foundstr = "Order By A.addTime Desc ,A.flashid Desc"
			Case 1: foundstr = "And A.isBest > 0 Order By A.addTime Desc ,A.flashid Desc"
			Case 2: foundstr = "Order By A.AllHits Desc ,A.flashid Desc"
			Case 3: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.addTime Desc ,A.flashid Desc"
			Case 4: foundstr = "And A.ClassID in (" & ChildStr & ") And A.isBest > 0 Order By A.addTime Desc ,A.flashid Desc"
			Case 5: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.AllHits Desc ,A.flashid Desc"
		Case Else
			foundstr = "Order By A.addTime Desc ,A.flashid Desc"
		End Select
		If CInt(stype) >= 3 And CLng(ClassID) = 0 Then
			foundstr = "Order By A.addTime Desc ,A.flashid Desc"
		End If
		If CLng(SpecialID) <> 0 Then
			foundstr = "And A.SpecialID =" & CLng(SpecialID) & " " & foundstr
		End If
		SQL = " A.flashid,A.ClassID,A.ColorMode,A.FontMode,A.title,A.Author,A.AllHits,A.addTime,A.HtmlFileDate,A.isBest,"
		SQL = "SELECT TOP " & CInt(TopNum) & SQL & " C.ClassName,C.ColorModes,C.FontModes,C.HtmlFileDir,C.UseHtml FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ChannelID=" & ChannelID & " " & foundstr & ""
		Set Rs = enchiasp.Execute(SQL)
		j = 0
		
		If Rs.BOF And Rs.EOF Then
			strContent = "没有添加任何信息！"
		Else
			SQL=Rs.GetRows(-1)
			strContent = "<table width=""100%"" border=0 cellpadding=2 cellspacing=0>"
			For i=0 To Ubound(SQL,2)
				If (j Mod 2) = 0 Then
					ListStyle = Trim(styles) & 1
				Else
					ListStyle = Trim(styles) & 2
				End If
				If CInt(SQL(9,i)) <> 0 Then
					BestCode = 2
					BestString = "<font color='" & enchiasp.MainSetting(3) & "'>推荐</font>"
				Else
					BestCode = 1
					BestString = ""
				End If
				strContent = strContent & enchiasp.MainSetting(22)
				strTitle = enchiasp.GotTopic(SQL(4,i), CInt(strLen))
				strTitle = enchiasp.ReadFontMode(strTitle, SQL(2,i), SQL(3,i))
				ClassName = enchiasp.ReadFontMode(SQL(10,i), SQL(11,i), SQL(12,i))
				HtmlFileName = enchiasp.ReadFileName(SQL(8,i), SQL(0,i), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) <> 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & SQL(13,i) & enchiasp.ShowDatePath(SQL(8,i), enchiasp.ChannelHtmlPath) & HtmlFileName
					ClassName = "<a href='" & enchiasp.ChannelPath & SQL(13,i) & "index" & enchiasp.ChannelHtmlExt & "'>" & ClassName & "</a>"
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & SQL(0,i)
					ClassName = "<a href='" & enchiasp.ChannelPath & "list.asp?classid=" & SQL(1,i) & "'>" & ClassName & "</a>"
				End If
				If CInt(showclass) = 0 Then ClassName = ""
				If CInt(showdate) <> 0 Then
					addTime = enchiasp.ShowDateTime(SQL(7,i), CInt(DateMode))
				Else
					addTime = ""
				End If
				If CInt(newindow) <> 0 Then
					LinkTarget = " target=""_blank"""
				Else
					LinkTarget = ""
				End If
				FlashTopic = "<a href='" & HtmlFileUrl & "'" & LinkTarget & " title='" & enchiasp.ChannelModule & "名称：" & SQL(4,i) & "&#13;&#10;发布时间：" & SQL(7,i) & "&#13;&#10;下载次数：" & SQL(6,i) & "' class=showlist>" & strTitle & "</a>"
				strContent = Replace(strContent, "{$FlashTopic}", FlashTopic)
				strContent = Replace(strContent, "{$FlashID}", Rs("flashid"))
				strContent = Replace(strContent, "{$InstallDir}", enchiasp.InstallDir)
				strContent = Replace(strContent, "{$FlashTopic}", strTitle)
				strContent = Replace(strContent, "{$Title}", SQL(4,i))
				strContent = Replace(strContent, "{$DateAndTime}", SQL(7,i))
				strContent = Replace(strContent, "{$HtmlFileUrl}", HtmlFileUrl)
				strContent = Replace(strContent, "{$ClassName}", ClassName)
				strContent = Replace(strContent, "[]", "")
				strContent = Replace(strContent, "{$Target}", LinkTarget)
				strContent = Replace(strContent, "{$addTime}", addTime)
				strContent = Replace(strContent, "{$FlashHits}", SQL(6,i))
				strContent = Replace(strContent, "{$ListStyle}", ListStyle)
				strContent = Replace(strContent, "{$BestCode}", BestCode)
				strContent = Replace(strContent, "{$BestString}", BestString)
			j = j + 1
			Next
			SQL=Null
			strContent = strContent & "</table>"
		End If
		Rs.Close: Set Rs = Nothing
		LoadFlashList = strContent
	End Function
	'================================================
	'函数名：ReadFlashList
	'作  用：读取动画列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadFlashList(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent
		Dim arrTempContent, arrTempContents, ArrayList
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$ReadFlashList(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadFlashList(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadFlashList(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadFlashList(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8), ArrayList(9), ArrayList(10)))
			Next
		End If
		ReadFlashList = strTemp
	End Function
	'================================================
	'函数名：LoadAnnounceContent
	'作  用：装载内容公告
	'参  数：str ----原字符串
	'================================================
	Public Function LoadAnnounceContent(ByVal sTopic, ByVal ChannelID)
		Dim SQL, Rs, strTemp
		strTemp = ""
		sTopic = enchiasp.CheckStr(sTopic)
		If sTopic <> "" And sTopic <> "0" Then
			SQL = "Select AnnounceID,Content,PostTime,writer From ECCMS_Announce where AnnounceType=1 And title = '" & sTopic & "' Order By PostTime Desc,AnnounceID Desc"
		Else
			SQL = "Select AnnounceID,Content From ECCMS_Announce where AnnounceType=1 And ChannelID in (" & ChannelID & ",999) Order By PostTime Desc,AnnounceID Desc"
		End If
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.BOF And Rs.EOF) Then
			strTemp = Rs("Content")
		End If
		Rs.Close: Set Rs = Nothing
		LoadAnnounceContent = strTemp
	End Function
	'================================================
	'函数名：ReadAnnounceContent
	'作  用：读取内容公告
	'参  数：str ----原字符串
	'================================================
	Public Function ReadAnnounceContent(ByVal str, ByVal ChannelID)
		Dim strTemp, i, sTempContent, nTempContent, strValue
		Dim arrTempContent, arrTempContents
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$AnnounceContent(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$AnnounceContent(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$AnnounceContent(", ")}", 0)
			If nTempContent = "" Then nTempContent = "0"
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				strValue = arrTempContent(i)
				strTemp = Replace(strTemp, arrTempContents(i), LoadAnnounceContent(strValue, ChannelID))
				strtemp=strtemp&"<br>"
			Next
		End If
	
		ReadAnnounceContent = strTemp
	End Function
	'================================================
	'函数名：LoadAnnounceList
	'作  用：装载公告列表
	'参  数：maxnum ----最多公告数
	'        maxlen ----字符长度
	'        newindow ----是否新窗口打开 1=是，0=否
	'        showdate ----是否显示时间 1=是，0=否
	'        DateMode ----时间模式
	'        showtree ----树型显示
	'================================================
	Public Function LoadAnnounceList(ByVal ChannelID, ByVal maxnum, ByVal maxlen, _
		ByVal newindow, ByVal showdate, ByVal DateMode, ByVal showtree)
		
		Dim Rs, SQL, strContent
		Dim AnnounceTopic, LinkTarget
		Dim PostTime
		
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		maxnum = enchiasp.ChkNumeric(maxnum)
		If maxnum = 0 Then maxnum = 10
		
		On Error Resume Next
		Set Rs = enchiasp.Execute("SELECT TOP " & CInt(maxnum) & " AnnounceID,title,Content,PostTime,writer,hits FROM ECCMS_Announce WHERE (ChannelID=" & ChannelID & " Or ChannelID=999) And AnnounceType<>1 ORDER BY PostTime DESC,AnnounceID DESC")
		If Rs.BOF And Rs.EOF Then
			LoadAnnounceList = ""
			Set Rs = Nothing
			Exit Function
		Else
			Do While Not Rs.EOF
				If CInt(newindow) <> 0 Then
					LinkTarget = " target=""_blank"""
				Else
					LinkTarget = ""
				End If
				If CInt(showdate) <> 0 Then
					PostTime = enchiasp.ShowDateTime(Rs("PostTime"), CInt(DateMode))
				Else
					PostTime = ""
				End If
				AnnounceTopic = enchiasp.GotTopic(Rs("title"), CInt(maxlen))
				AnnounceTopic = "<a href=""" & enchiasp.InstallDir & "Announce.Asp?AnnounceID=" & Rs("AnnounceID") & """ title=""" & Rs("title") & """" & LinkTarget & ">" & AnnounceTopic & "</a>"
				If CInt(showtree) = 1 Then
					strContent = strContent & "<div>· " & AnnounceTopic & "</div><div align=""right"" class=""dottedline"">" & PostTime & "</div>" & vbNewLine
				Else
					strContent = strContent & "· " & AnnounceTopic & "&nbsp;&nbsp;" & PostTime & vbNewLine
				End If
				strContent=strContent&"<br>"
				Rs.MoveNext
			Loop
		End If
		LoadAnnounceList = strContent
	End Function


	'================================================
	'函数名：ReadAnnounceList
	'作  用：读取公告列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadAnnounceList(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$ReadAnnounceList(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadAnnounceList(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadAnnounceList(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadAnnounceList(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6)))
			Next
		End If
		ReadAnnounceList = strTemp
	End Function
	'================================================
	'函数名：LoadArticlePic
	'作  用：装载文章图片列表
	'参  数：ClassID   ----分类ID
	'        ChannelID   ----频道ID
	'        sType   ----调用文章类型,0=所有最新文章，1=推荐文章，2=热门文章，3=图文文章，4=分类最新文章
	'        TopNum   ----显示文章列表数
	'        strlen   ----显示标题长度
	'        ShowClass   ----是否显示分类
	'        ShowPic   ----是否显示图文标题
	'        ShowDate   ----是否显示日期
	'        DateMode   ----显示日期模式
	'        newindow   ----新窗口打开
	'================================================
		Public Function LoadArticlePic(ChannelID, ClassID, SpecialID, stype, TopNum, PerRowNum, strLen, newindow, width, height, showtopic)
		Dim Rs, SQL, i, strContent, foundstr
		Dim sTitle, ChildStr, ImageUrl, HtmlFileName
		Dim HtmlFileUrl, WriteTime, LinkTarget
		
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		SpecialID = enchiasp.ChkNumeric(SpecialID)
		stype = enchiasp.ChkNumeric(stype)
		
		On Error Resume Next
		enchiasp.LoadChannel(ChannelID)
		
		If CInt(stype) >= 3 And CLng(ClassID) <> 0 Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID=" & ChannelID & " And ClassID=" & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadArticlePic = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
			End If
			Set Rs = Nothing
		Else
			ChildStr = 0
		End If
		Select Case CInt(stype)
			Case 0: foundstr = "Order By A.Writetime Desc ,A.Articleid Desc"
			Case 1: foundstr = "And A.isBest > 0 Order By A.Writetime Desc ,A.Articleid Desc"
			Case 2: foundstr = "Order By A.AllHits Desc ,A.Articleid Desc"
			Case 3: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.Writetime Desc ,A.Articleid Desc"
			Case 4: foundstr = "And A.ClassID in (" & ChildStr & ") And A.isBest > 0 Order By A.Writetime Desc ,A.Articleid Desc"
			Case 5: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.AllHits Desc ,A.Articleid Desc"
		Case Else
			foundstr = "Order By A.Writetime Desc ,A.Articleid Desc"
		End Select
		If CInt(stype) >= 4 And CLng(ClassID) = 0 Then
			foundstr = "Order By A.Writetime Desc ,A.Articleid Desc"
		End If
		If CLng(SpecialID) <> 0 Then
			foundstr = "And A.SpecialID =" & CLng(SpecialID) & " " & foundstr
		End If
		SQL = " A.ArticleID,A.ClassID,A.title,A.AllHits,A.WriteTime,A.HtmlFileDate,A.isBest,A.ImageUrl,"
		SQL = "select Top " & CInt(TopNum) & SQL & " C.ClassName,C.HtmlFileDir,C.UseHtml,a.content from [ECCMS_Article] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.isAccept > 0 And A.ImageUrl<>'' And A.ChannelID=" & ChannelID & " " & foundstr & ""
		
		
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			strContent = "<img src='" & enchiasp.InstallDir & "images/no_pic.gif' width=" & width & " height=" & height & " border=0>"
		Else
			strContent = "<table width=""100%"" border=0 cellpadding=0 cellspacing=0>" & vbCrLf
			Do While Not Rs.EOF
			
				strContent = strContent & "<tr>" & vbCrLf
				For i = 1 To CInt(PerRowNum)
					strContent = strContent & "<td align=""center"" class=""imagelist"" style=""padding-top:10px;"">"
					If Not Rs.EOF Then
						'特殊显示时将显示内容，调用仍调用标题标签
						if CInt(showtopic) = 2 then
							sTitle = enchiasp.CutString(Rs("content"), CInt(strLen))
						else
							sTitle = enchiasp.GotTopic(Rs("title"), CInt(strLen))
						end if
						ImageUrl = enchiasp.GetImageUrl(Rs("ImageUrl"), enchiasp.ChannelData(1))
						ImageUrl = enchiasp.GetFlashAndPic(ImageUrl, height, width)
						
						HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("ArticleID"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
						If CInt(enchiasp.ChannelUseHtml) <> 0 Then
							HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
						Else
							HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("ArticleID")
						End If
						
						If CInt(newindow) <> 0 Then
							LinkTarget = " target=""_blank"""
						Else
							LinkTarget = ""
						End If
						if CInt(showtopic) = 2 then
							strContent = strContent & enchiasp.MainSetting(29)
						else
							strContent = strContent & enchiasp.MainSetting(18)
						end if
						strContent = Replace(strContent, "{$ArticleTitle}", Rs("title"))
						
						strContent = Replace(strContent, "{$ArticlePicture}", "<a href='" & HtmlFileUrl & "' title='" & Rs("title") & "'" & LinkTarget & ">" & ImageUrl & "</a>")
						If CInt(showtopic) = 0 Then
							strContent = Replace(strContent, "{$ArticleTopic}", vbNullString)
						Else
							strContent = Replace(strContent, "{$ArticleTopic}", "<a href='" & HtmlFileUrl & "' title='" & Rs("title") & "'" & LinkTarget & ">" & sTitle & "</a>")
						End If
						strContent = strContent & "</td>" & vbCrLf
					Rs.MoveNext
				End If
			Next
			strContent = strContent & "</tr>" & vbCrLf
			Loop
			strContent = strContent & "</table>" & vbCrLf
		End If
		Rs.Close: Set Rs = Nothing
		LoadArticlePic = strContent
	End Function
	'================================================
	'函数名：ReadArticlePic
	'作  用：读取文章图片列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadArticlePic(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$ReadArticlePic(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadArticlePic(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadArticlePic(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadArticlePic(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8), ArrayList(9), ArrayList(10)))
			Next
		End If
		ReadArticlePic = strTemp
	End Function
	'================================================
	'函数名：LoadSoftPic
	'作  用：装载软件图片列表
	'参  数：ClassID   ----分类ID
	'        ChannelID   ----频道ID
	'        sType   ----调用软件类型,0=所有最新软件，1=推荐软件，2=热门软件
	'        TopNum   ----显示软件列表数
	'        strlen   ----显示标题长度
	'        newindow   ----新窗口打开
	'================================================
	Public Function LoadSoftPic(ChannelID, ClassID, SpecialID, stype, TopNum, PerRowNum, strLen, newindow, width, height, showtopic)
		Dim Rs, SQL, i, strContent, foundstr
		Dim strSoftName, ChildStr, SoftImage, HtmlFileName
		Dim HtmlFileUrl, SoftTime, LinkTarget
		
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		SpecialID = enchiasp.ChkNumeric(SpecialID)
		stype = enchiasp.ChkNumeric(stype)
		
		On Error Resume Next
		enchiasp.LoadChannel(ChannelID)
		
		If CInt(stype) >= 3 And CLng(ClassID) <> 0 Then
			SQL = "select ChildStr from [ECCMS_Classify] where ChannelID = " & ChannelID & " And ClassID = " & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadSoftPic = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
			End If
			Rs.Close
		Else
			ChildStr = 0
		End If
		Select Case CInt(stype)
			Case 0: foundstr = "Order By A.SoftTime Desc ,A.SoftID Desc"
			Case 1: foundstr = "And A.isBest > 0 Order By A.SoftTime Desc ,A.SoftID Desc"
			Case 2: foundstr = "Order By A.AllHits Desc ,A.SoftID Desc"
			Case 3: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.SoftTime Desc ,A.SoftID Desc"
			Case 4: foundstr = "And A.ClassID in (" & ChildStr & ") And A.isBest > 0 Order By A.SoftTime Desc ,A.SoftID Desc"
			Case 5: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.AllHits Desc ,A.SoftID Desc"
		Case Else
			foundstr = "Order By A.SoftTime Desc ,A.SoftID Desc"
		End Select
		If CInt(stype) >= 3 And CLng(ClassID) = 0 Then
			foundstr = "Order By A.SoftTime Desc ,A.SoftID Desc"
		End If
		If CLng(SpecialID) <> 0 Then
			foundstr = "And A.SpecialID =" & CLng(SpecialID) & " " & foundstr
		End If
		SQL = " A.SoftID,A.ClassID,A.SoftName,A.SoftVer,A.AllHits,A.SoftTime,A.HtmlFileDate,A.isBest,A.SoftImage,"
		SQL = "select Top " & CInt(TopNum) & SQL & " C.ClassName,C.HtmlFileDir,C.UseHtml from [ECCMS_SoftList] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.isAccept>0 And A.SoftImage<>'' And A.ChannelID=" & ChannelID & " " & foundstr & ""
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			strContent = "<img src='" & enchiasp.InstallDir & "images/no_pic.gif' width=" & width & " height=" & height & " border=0>"
		Else
			strContent = "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""3"">" & vbCrLf
			Do While Not Rs.EOF
				strContent = strContent & "<tr>" & vbCrLf
				For i = 1 To CInt(PerRowNum)
					strContent = strContent & "<td align=""center"" class=""imagelist"">"
					If Not Rs.EOF Then
						strSoftName = enchiasp.GotTopic(Rs("SoftName") & " " & Rs("SoftVer"), CInt(strLen))
						SoftImage = enchiasp.GetImageUrl(Rs("SoftImage"), enchiasp.ChannelData(1))
						SoftImage = enchiasp.GetFlashAndPic(SoftImage, height, width)
						HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("SoftID"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
						If CInt(enchiasp.ChannelUseHtml) <> 0 Then
							HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
						Else
							HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("SoftID")
						End If
						If CInt(newindow) <> 0 Then
							LinkTarget = " target=""_blank"""
						Else
							LinkTarget = ""
						End If
						strContent = strContent & enchiasp.MainSetting(19)
						strContent = Replace(strContent, "{$SoftPicture}", "<a href='" & HtmlFileUrl & "' title='" & Rs("SoftName") & "'" & LinkTarget & ">" & SoftImage & "</a>")
						If CInt(showtopic) = 1 Then
							strContent = Replace(strContent, "{$SoftTopic}", "<a href='" & HtmlFileUrl & "' title='" & Rs("SoftName") & "'" & LinkTarget & ">" & strSoftName & "</a>")
						Else
							strContent = Replace(strContent, "{$SoftTopic}", vbNullString)
						End If
						strContent = strContent & "</td>" & vbCrLf
					Rs.MoveNext
				End If
			Next
			strContent = strContent & "</tr>" & vbCrLf
			Loop
			strContent = strContent & "</table>" & vbCrLf
		End If
		Rs.Close: Set Rs = Nothing
		LoadSoftPic = strContent
	End Function
	'================================================
	'函数名：ReadSoftPic
	'作  用：读取软件图片列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadSoftPic(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$ReadSoftPic(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadSoftPic(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadSoftPic(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadSoftPic(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8), ArrayList(9), ArrayList(10)))
			Next
		End If
		ReadSoftPic = strTemp
	End Function
	
		'================================================
	'函数名：LoadShopPic
	'作  用：装载商品图片列表
	'参  数：ClassID   ----分类ID
	'        ChannelID   ----频道ID
	'        sType   ----调用商品类型,0=所有最新商品，1=推荐商品，2=热门商品
	'        TopNum   ----显示商品列表数
	'        strlen   ----显示标题长度
	'        newindow   ----新窗口打开
	'2007年7月增加标示,是否显示多余内容,最后一个内容是"2"加载带会员价显示的内容
	'================================================
	Public Function LoadShopPic(ChannelID, ClassID, SpecialID, stype, TopNum, PerRowNum, strLen, newindow, width, height, showtopic)
		Dim Rs, SQL, i, strContent, foundstr
		Dim strTradeName, ChildStr, ProductImage, HtmlFileName
		Dim HtmlFileUrl, addTime, LinkTarget,ShopTime
		
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		SpecialID = enchiasp.ChkNumeric(SpecialID)
		stype = enchiasp.ChkNumeric(stype)
		
		On Error Resume Next
		enchiasp.LoadChannel(ChannelID)
		
		If CInt(stype) >= 3 And CLng(ClassID) > 0 Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID=" & ChannelID & " And ClassID=" & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadShopPic = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
			End If
			Rs.Close
		Else
			ChildStr = 0
		End If
		Select Case CInt(stype)
			Case 0: foundstr = "ORDER BY A.addTime DESC ,A.ShopID DESC"
			Case 1: foundstr = "And A.isBest > 0 ORDER BY A.addTime DESC ,A.ShopID DESC"
			Case 2: foundstr = "ORDER BY A.AllHits DESC ,A.ShopID DESC"
			Case 3: foundstr = "And A.ClassID in (" & ChildStr & ") ORDER BY A.addTime DESC ,A.ShopID DESC"
			Case 4: foundstr = "And A.ClassID in (" & ChildStr & ") And A.isBest > 0 ORDER BY A.addTime DESC ,A.ShopID DESC"
			Case 5: foundstr = "And A.ClassID in (" & ChildStr & ") ORDER BY A.AllHits DESC ,A.ShopID DESC"
		Case Else
			foundstr = "Order By A.addTime Desc ,A.ShopID Desc"
		End Select
		If CInt(stype) >= 3 And CLng(ClassID) = 0 Then
			foundstr = "Order By A.addTime Desc ,A.ShopID Desc"
		End If
		If CLng(SpecialID) <> 0 Then
			foundstr = "And A.SpecialID =" & CLng(SpecialID) & " " & foundstr
		End If
		SQL = " A.ShopID,A.ClassID,A.TradeName,A.PastPrice,A.NowPrice,A.AllHits,A.addTime,A.HtmlFileDate,A.isBest,A.ProductImage,A.Star,"
		If CInt(showtopic) = 0 Then
			SQL = "SELECT TOP " & CInt(TopNum) & SQL & " C.ClassName,C.HtmlFileDir,C.UseHtml FROM [ECCMS_ShopList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ProductImage<>'' And A.ChannelID=" & ChannelID & " " & foundstr
		Else
			SQL = "SELECT TOP " & CInt(TopNum) & SQL & " C.ClassName,C.HtmlFileDir,C.UseHtml FROM [ECCMS_ShopList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ChannelID=" & ChannelID & " " & foundstr

		End If
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			strContent = "<img src='" & enchiasp.InstallDir & "images/no_pic.gif' width=" & width & " height=" & height & " border=0>"
		Else
			strContent = "<table width=""100%"" border=0 cellpadding=1 cellspacing=3>" & vbCrLf
			Do While Not Rs.EOF
				strContent = strContent & "<tr>" & vbCrLf
				For i = 1 To CInt(PerRowNum)
					strContent = strContent & "<td align=center class=shopimagelist>"
					If Not Rs.EOF Then
						strTradeName = enchiasp.GotTopic(Rs("TradeName"), CInt(strLen))
						ProductImage = enchiasp.GetImageUrl(Rs("ProductImage"), enchiasp.ChannelData(1))
						ProductImage = enchiasp.GetFlashAndPic(ProductImage, height, width)
						HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("ShopID"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
						If CInt(enchiasp.ChannelUseHtml) <> 0 Then
							HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
						Else
							HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("ShopID")
						End If
						If CInt(newindow) <> 0 Then
							LinkTarget = " target=""_blank"""
						Else
							LinkTarget = ""
						End If
						ShopTime = enchiasp.ShowDateTime(Rs("addTime"), 2)
						if CInt(showtopic) = 2 then
							strContent = strContent & enchiasp.MainSetting(28)
						else
							strContent = strContent & enchiasp.MainSetting(20)
						end if
						strContent = Replace(strContent, "{$ShopID}", Rs("shopid"))
						strContent = Replace(strContent, "{$ShopUrl}", HtmlFileUrl)
						strContent = Replace(strContent, "{$ChannelRootDir}", enchiasp.ChannelPath)
						strContent = Replace(strContent, "{$SkinPath}", enchiasp.SkinPath)
						strContent = Replace(strContent, "{$InstallDir}", enchiasp.InstallDir)
						strContent = Replace(strContent, "{$ShopHits}", Rs("AllHits"))
						strContent = Replace(strContent, "{$Star}", Rs("star"))
						strContent = Replace(strContent, "{$ShopDateTime}", ShopTime)
						strContent = Replace(strContent, "{$PastPrice}", FormatNumber(Rs("PastPrice"),2,-1))
						strContent = Replace(strContent, "{$NowPrice}", FormatNumber(Rs("NowPrice"),2,-1))
						strContent = Replace(strContent, "{$ProductImage}", "<a href='" & HtmlFileUrl & "' title='" & Rs("TradeName") & "'" & LinkTarget & ">" & ProductImage & "</a>")
						strContent = Replace(strContent, "{$TradeName}", "<a href='" & HtmlFileUrl & "' title='" & Rs("TradeName") & "'" & LinkTarget & ">" & strTradeName & "</a>")
						strContent = strContent & "</td>" & vbCrLf
					Rs.MoveNext
				End If
			Next
			strContent = strContent & "</tr>" & vbCrLf
			Loop
			strContent = strContent & "</table>" & vbCrLf
		End If
		Rs.Close: Set Rs = Nothing
		LoadShopPic = strContent
	End Function
	'================================================
	'函数名：ReadShopPic
	'作  用：读取商品图片列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadShopPic(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$ReadShopPic(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadShopPic(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadShopPic(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadShopPic(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8), ArrayList(9), ArrayList(10)))
			Next
		End If
		ReadShopPic = strTemp
	End Function
	
	
	
	
	
	'================================================
	'函数名：LoadFlashPic
	'作  用：装载动画图片列表
	'参  数：ClassID   ----分类ID
	'        ChannelID   ----频道ID
	'        sType   ----调用动画类型,0=所有最新动画，1=推荐动画，2=热门动画
	'        TopNum   ----显示动画列表数
	'        strlen   ----显示标题长度
	'        newindow   ----新窗口打开
	'================================================
	Public Function LoadFlashPic(ByVal ChannelID, ByVal ClassID, ByVal SpecialID, _
		ByVal stype, ByVal TopNum, ByVal PerRowNum, ByVal strLen, ByVal newindow, _
		ByVal width, ByVal height, ByVal showtopic)
		
		Dim Rs, SQL, i, strContent, foundstr
		Dim strtitle, ChildStr, miniature, HtmlFileName
		Dim HtmlFileUrl, addTime, LinkTarget
		
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		SpecialID = enchiasp.ChkNumeric(SpecialID)
		stype = enchiasp.ChkNumeric(stype)
		
		On Error Resume Next
		enchiasp.LoadChannel(ChannelID)
		
		If CInt(stype) >= 3 And CLng(ClassID) <> 0 Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID = " & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadFlashPic = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
			End If
			Rs.Close
		Else
			ChildStr = 0
		End If
		
		Select Case CInt(stype)
			Case 0: foundstr = "ORDER BY A.addTime DESC ,A.flashid DESC"
			Case 1: foundstr = "And A.isBest > 0 ORDER BY A.addTime DESC ,A.flashid DESC"
			Case 2: foundstr = "ORDER BY A.AllHits DESC ,A.flashid DESC"
			Case 3: foundstr = "And A.ClassID in (" & ChildStr & ") ORDER BY A.addTime DESC ,A.flashid DESC"
			Case 4: foundstr = "And A.ClassID in (" & ChildStr & ") And A.isBest > 0 ORDER BY A.addTime DESC ,A.flashid DESC"
			Case 5: foundstr = "And A.ClassID in (" & ChildStr & ") ORDER BY A.AllHits DESC ,A.flashid DESC"
		Case Else
			foundstr = "ORDER BY A.addTime DESC ,A.flashid DESC"
		End Select
		If CInt(stype) >= 3 And CLng(ClassID) = 0 Then
			foundstr = "ORDER BY A.addTime DESC ,A.flashid DESC"
		End If
		If CLng(SpecialID) <> 0 Then
			foundstr = "And A.SpecialID =" & CLng(SpecialID) & " " & foundstr
		End If
		SQL = " A.flashid,A.ClassID,A.title,A.AllHits,A.addTime,A.HtmlFileDate,A.isBest,A.miniature,"
		SQL = "SELECT TOP " & CInt(TopNum) & SQL & " C.ClassName,C.HtmlFileDir,C.UseHtml FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.isAccept>0 And A.miniature<>'' And A.ChannelID=" & ChannelID & " " & foundstr & ""
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			strContent = "<img src='" & enchiasp.InstallDir & "images/no_pic.gif' width=" & width & " height=" & height & " border=0>"
		Else
			strContent = "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""3"">" & vbCrLf
			Do While Not Rs.EOF
				strContent = strContent & "<tr>" & vbCrLf
				For i = 1 To CInt(PerRowNum)
					strContent = strContent & "<td align=""center"" class=""imagelist"">"
					If Not Rs.EOF Then
						strtitle = enchiasp.GotTopic(Rs("title"), CInt(strLen))
						miniature = enchiasp.GetImageUrl(Rs("miniature"), enchiasp.ChannelData(1))
						miniature = enchiasp.GetFlashAndPic(miniature, height, width)
						HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("flashid"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
						If CInt(enchiasp.ChannelUseHtml) <> 0 Then
							HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
						Else
							HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("flashid")
						End If
						If CInt(newindow) <> 0 Then
							LinkTarget = " target=""_blank"""
						Else
							LinkTarget = ""
						End If
						strContent = strContent & enchiasp.MainSetting(21)
						strContent = Replace(strContent, "{$Miniature}", "<a href='" & HtmlFileUrl & "' title='" & Rs("title") & "'" & LinkTarget & ">" & miniature & "</a>")
						If CInt(showtopic) = 1 Then
							strContent = Replace(strContent, "{$FlashTopic}", "<a href='" & HtmlFileUrl & "' title='" & Rs("title") & "'" & LinkTarget & ">" & strtitle & "</a>")
						Else
							strContent = Replace(strContent, "{$FlashTopic}", vbNullString)
						End If
						strContent = strContent & "</td>" & vbCrLf
					Rs.MoveNext
					End If
				Next
			strContent = strContent & "</tr>" & vbCrLf
			Loop
			strContent = strContent & "</table>" & vbCrLf
		End If
		Rs.Close: Set Rs = Nothing
		LoadFlashPic = strContent
	End Function
	'================================================
	'函数名：ReadFlashPic
	'作  用：读取动画图片列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadFlashPic(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$ReadFlashPic(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadFlashPic(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadFlashPic(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadFlashPic(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8), ArrayList(9), ArrayList(10)))
			Next
		End If
		ReadFlashPic = strTemp
	End Function
	'================================================
	'函数名：LoadFriendLink
	'作  用：装载友情连接
	'参  数：str ----原字符串
	'================================================
	Public Function LoadFriendLink(ByVal TopNum, ByVal PerRowNum, ByVal isLogo, ByVal orders)
		Dim Rs, SQL, i, strContent
		Dim strOrder, LinkAddress
	
		strContent = ""
		If Not IsNumeric(TopNum) Then Exit Function
		If Not IsNumeric(PerRowNum) Then Exit Function
		If Not IsNumeric(isLogo) Then Exit Function
		If Not IsNumeric(orders) Then Exit Function
		On Error Resume Next
		If CInt(orders) = 1 Then
			'-- 首页显示按时间升序排列
			strOrder = "And isIndex > 0 Order By LinkTime Desc,LinkID Desc"
		ElseIf CInt(orders) = 2 Then
			'-- 首页显示按点击数升序排列
			strOrder = "And isIndex > 0 Order By LinkHist Desc,LinkID Desc"
		ElseIf CInt(orders) = 3 Then
			'-- 首页显示按点击数降序排列
			strOrder = "And isIndex > 0 Order By LinkHist Desc,LinkID Asc"
		ElseIf CInt(orders) = 4 Then
			'-- 所有按升序排列
			strOrder = "Order By LinkID Desc"
		ElseIf CInt(orders) = 5 Then
			'-- 所有按降序排列
			strOrder = "Order By LinkID Asc"
		ElseIf CInt(orders) = 6 Then
			'-- 所有按点击数升序排列
			strOrder = "Order By LinkHist Desc,LinkID Desc"
		ElseIf CInt(orders) = 7 Then
			'-- 所有按点击数降序排列
			strOrder = "Order By LinkHist Desc,LinkID Asc"
		ElseIf CInt(orders) = 8 Then
			'-- 首页显示按名称排列
			strOrder = "And isIndex > 0 Order By LinkName Desc,LinkID Desc"
		ElseIf CInt(orders) = 9 Then
			'-- 所有按名称排列
			strOrder = "Order By LinkName Desc,LinkID Desc"
		Else
			'-- 首页显示按时间降序排列
			strOrder = "And isIndex > 0 Order By LinkTime Asc,LinkID Asc"
		End If
		If CInt(isLogo) = 1 Or CInt(isLogo) = 3 Then
			SQL = "Select Top " & CInt(TopNum) & " LinkID,LinkName,LinkUrl,LogoUrl,Readme,LinkHist,isLogo from [ECCMS_Link] where isLock = 0 And isLogo > 0 " & strOrder & ""
		Else
			SQL = "Select Top " & CInt(TopNum) & " LinkID,LinkName,LinkUrl,LogoUrl,Readme,LinkHist,isLogo from [ECCMS_Link] where isLock = 0 And isLogo = 0 " & strOrder & ""
		End If
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.BOF And Rs.EOF) Then
			strContent = "<table width=""100%"" border=0 cellpadding=1 cellspacing=3 class=FriendLink1>" & vbCrLf
			Do While Not Rs.EOF
				strContent = strContent & "<tr>" & vbCrLf
				For i = 1 To CInt(PerRowNum)
					strContent = strContent & "<td align=center class=FriendLink2>"
					If Not Rs.EOF Then
						If CInt(isLogo) < 2 Then
							LinkAddress = enchiasp.InstallDir & "link/link.asp?id=" & Rs("LinkID") & "&url=" & Trim(Rs("LinkUrl"))
						Else
							LinkAddress = Trim(Rs("LinkUrl"))
						End If
						If Rs("isLogo") = 1 Or CInt(isLogo) = 3 Then
							strContent = strContent & "<a href='" & LinkAddress & "' target=_blank title='主页名称：" & Rs("LinkName") & "&#13;&#10;点击次数：" & Rs("LinkHist") & "'><img src='" & enchiasp.ReadFileUrl(Rs("LogoUrl")) & "' border=0 width=162 height=48></a>"
						Else
							strContent = strContent & "<a href='" & LinkAddress & "' target=_blank title='主页名称：" & Rs("LinkName") & "&#13;&#10;点击次数：" & Rs("LinkHist") & "'>" & Rs("LinkName") & "</a>"
						End If
						strContent = strContent & "</td>" & vbCrLf
						Rs.MoveNext
					Else
						If CInt(isLogo) = 1 Or CInt(isLogo) = 3 Then
							strContent = strContent & "<a href='" & enchiasp.InstallDir & "link/addlink.asp' target=_blank><img src='" & enchiasp.InstallDir & "images/link.gif'  border=0></a>"
						Else
							strContent = strContent & "<a href='" & enchiasp.InstallDir & "link/' target=_blank>更多连接</a>"
						End If
						strContent = strContent & "</td>" & vbCrLf
					End If
				Next
				strContent = strContent & "</tr>" & vbCrLf
			Loop
			strContent = strContent & "</table>" & vbCrLf
		End If
		LoadFriendLink = strContent
	End Function
	'================================================
	'函数名：ReadFriendLink
	'作  用：读取友情连接
	'参  数：str ----原字符串
	'================================================
	Public Function ReadFriendLink(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		On Error Resume Next
		strTemp = str

		If InStr(strTemp, "{$ReadFriendLink(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadFriendLink(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadFriendLink(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadFriendLink(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3)))
			Next
		End If
		ReadFriendLink = strTemp
	End Function
	'================================================
	'函数名：PageRunTime
	'作  用：页面执行时间
	'================================================
	Public Function ExecutionTime()
		Dim Endtime
		ExecutionTime = ""
		If CInt(enchiasp.IsRunTime) = 1 Then
			Endtime = Timer()
			ExecutionTime = "页面执行时间：" & FormatNumber((((Endtime - startime) * 5000) + 0.5) / 10, 3, -1) & "毫秒"
		Else
			ExecutionTime = ""
		End If
	End Function
	
	'================================================
	'函数名：CurrentStation
	'作  用：当前位置
	'参  数：...
	'================================================
	Public Function CurrentStation(ByVal ChannelID, ByVal ClassID, ByVal ClassName, _
		ByVal ParentID, ByVal strParent, ByVal HtmlFileDir, ByVal Compart)
		
		Dim rsCurrent, SQL, strContent, ChannelDir

		CurrentStation = ""
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		ParentID = enchiasp.ChkNumeric(ParentID)
		
		On Error Resume Next
		enchiasp.LoadChannel(ChannelID)

		ChannelDir = enchiasp.ChannelPath
			
		strContent = "<a href='" & ChannelDir & "'>" & enchiasp.ChannelName & "</a>" & Compart & ""
		If ParentID <> 0 And Len(strParent) <> 0 Then
			SQL = "SELECT ClassID,ClassName,HtmlFileDir,UseHtml FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID in(" & strParent & ")"
			Set rsCurrent = enchiasp.Execute(SQL)
			If Not (rsCurrent.EOF And rsCurrent.BOF) Then
				Do While Not rsCurrent.EOF
					
					If CInt(enchiasp.IsCreateHtml) <> 0 Then
						strContent = strContent & "<a href='" & ChannelDir & rsCurrent("HtmlFileDir") & "'>" & rsCurrent("ClassName") & "</a>" & Compart & ""
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
	Public Function ReadCurrentStation(ByVal str, ByVal ChannelID, ByVal ClassID, _
		ByVal ClassName, ByVal ParentID, ByVal strParent, ByVal HtmlFileDir)
		
		Dim strTemp, i
		Dim sTempContent, nTempContent
		Dim arrTempContent, arrTempContents

		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$CurrentStation(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$CurrentStation(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$CurrentStation(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)

				strTemp = Replace(strTemp, arrTempContents(i), CurrentStation(ChannelID, ClassID, ClassName, ParentID, strParent, HtmlFileDir, arrTempContent(i)))
			Next
		End If
		ReadCurrentStation = strTemp

	End Function
	

	'================================================
	'函数名：NewsPictureAndText
	'作  用：图文混排列表
	'================================================
	Public Function NewsPictureAndText(ByVal chanid, ByVal ClassID, ByVal specid, _
		ByVal stype, ByVal height, ByVal width, ByVal maxlen, _
		ByVal maxline, ByVal hspace, ByVal vspace, ByVal align, _
		ByVal divcss, ByVal target, ByVal start, ByVal showpic, _
		ByVal showclass, ByVal showdate, ByVal dateformat)
		
		Dim Rs, SQL, i, strContent, foundstr
		Dim ChildStr, HtmlFileUrl, HtmlFileName, strPicture
		Dim PicTopic, NewsTitle, ClassName, ArticleTitle, WriteTime
		
		chanid = enchiasp.ChkNumeric(chanid)
		ClassID = enchiasp.ChkNumeric(ClassID)
		specid = enchiasp.ChkNumeric(specid)
		stype = enchiasp.ChkNumeric(stype)
		
		On Error Resume Next
		enchiasp.LoadChannel(chanid)
		
		If CInt(stype) >= 3 And CLng(ClassID) <> 0 Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID = " & chanid & " And ClassID = " & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				NewsPictureAndText = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
			End If
			Rs.Close
		Else
			ChildStr = "0"
		End If
		Select Case CInt(stype)
			Case 0: foundstr = "ORDER BY A.Writetime DESC ,A.Articleid DESC"
			Case 1: foundstr = "And A.isBest > 0 ORDER BY A.Writetime DESC ,A.Articleid DESC"
			Case 2: foundstr = " ORDER BY A.AllHits DESC ,A.Articleid DESC"
			Case 3: foundstr = "And A.ClassID in (" & ChildStr & ") ORDER BY A.Writetime DESC ,A.Articleid DESC"
			Case 4: foundstr = "And A.ClassID in (" & ChildStr & ") And A.isBest > 0 ORDER BY A.Writetime DESC ,A.Articleid DESC"
			Case 5: foundstr = "And A.ClassID in (" & ChildStr & ") And A.AllHits > B.LeastHotHist ORDER BY A.AllHits DESC ,A.Articleid DESC"
			Case 6: foundstr = "And A.ClassID in (" & ChildStr & ") ORDER BY A.Writetime DESC ,A.Articleid DESC"
		Case Else
			foundstr = "ORDER BY A.Writetime DESC ,A.Articleid DESC"
		End Select
		If CInt(stype) >= 3 And CLng(ClassID) = 0 Then
			foundstr = "ORDER BY A.Writetime DESC ,A.Articleid DESC"
		End If
		If CLng(specid) <> 0 Then
			foundstr = "And A.SpecialID =" & CLng(specid) & " " & foundstr
		End If
		SQL = " A.ArticleID,A.ClassID,A.ColorMode,A.FontMode,A.title,A.BriefTopic,A.AllHits,A.WriteTime,A.HtmlFileDate,A.isBest,"
		SQL = "SELECT TOP " & CInt(maxline) & SQL & " C.ClassName,C.ColorModes,C.FontModes,C.HtmlFileDir FROM [ECCMS_Article] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ChannelID=" & chanid & " " & foundstr & ""
		Set Rs = enchiasp.Execute(SQL)
		i = 0
		strContent = ""
		If Rs.BOF And Rs.EOF Then
			strContent = "还没有添加任何内容！"
		Else
			Do While Not Rs.EOF
				NewsTitle = enchiasp.ReadTopic(Rs("title"), CInt(maxlen))
				NewsTitle = enchiasp.ReadFontMode(NewsTitle, Rs("ColorMode"), Rs("FontMode"))
				PicTopic = enchiasp.ReadPicTopic(Rs("BriefTopic"))
				ClassName = enchiasp.ReadFontMode(Rs("ClassName"), Rs("ColorModes"), Rs("FontModes"))
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("ArticleID"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) <> 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
					ClassName = "[<a href='" & enchiasp.ChannelPath & Rs("HtmlFileDir") & "index" & enchiasp.ChannelHtmlExt & "'>" & ClassName & "</a>]"
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("ArticleID")
					ClassName = "[<a href='" & enchiasp.ChannelPath & "list.asp?classid=" & Rs("ClassID") & "'>" & ClassName & "</a>]"
				End If
				If CInt(showclass) = 1 Then
					ClassName = ClassName
				Else
					ClassName = ""
				End If
				If CInt(showdate) = 1 Then
					WriteTime = enchiasp.ShowDateTime(Rs("WriteTime"), CInt(dateformat))
				Else
					WriteTime = ""
				End If
				ArticleTitle = "<div " & divcss & ">" & start & ClassName & " <a href=""" & HtmlFileUrl & """ target=""" & target & """ title=""" & enchiasp.ChannelModule & "标题：" & Rs("title") & "&#13;&#10;发布时间：" & Rs("WriteTime") & "&#13;&#10;阅览次数：" & Rs("AllHits") & """ class=showlist>" & NewsTitle & "</a>  " & WriteTime & "</div>"
				strContent = strContent & ArticleTitle
				Rs.MoveNext
				i = i + 1
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		Dim sExtName, ExtName, ImageUrl
		If CInt(showpic) = 1 Then
			SQL = " A.ArticleID,A.ClassID,A.title,A.AllHits,A.WriteTime,A.HtmlFileDate,A.ImageUrl,"
			SQL = "SELECT " & SQL & " C.HtmlFileDir,B.ChannelDir,B.StopChannel,B.ModuleName,B.BindDomain,B.DomainName,B.IsCreateHtml,B.HtmlExtName,B.HtmlPath,B.HtmlForm,B.HtmlPrefix,B.LeastHotHist FROM ([ECCMS_Article] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID) INNER JOIN [ECCMS_Channel] B On A.ChannelID=B.ChannelID WHERE A.isAccept>0 And A.ChannelID=" & CInt(chanid) & " And A.ImageUrl<>'' " & foundstr & ""
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				strPicture = "<img src='" & enchiasp.SiteUrl & enchiasp.InstallDir & "images/no_pic.gif' width=""" & width & """ height=""" & height & """  hspace=""" & hspace & """ vspace=""" & vspace & """ align=""" & align & """ border=""0"">"
			Else
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("ArticleID"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) <> 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("ArticleID")
				End If
				ImageUrl = enchiasp.GetImageUrl(Rs("ImageUrl"), enchiasp.ChannelData(1))
				sExtName = Split(Rs("ImageUrl"), ".")
				ExtName = sExtName(UBound(sExtName))
				Select Case LCase(ExtName)
				Case "swf", "swi"
					strPicture = "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"" width=""" & width & """ height=""" & height & """ hspace=""" & hspace & """ vspace=""" & vspace & """ align=""" & align & """>" & vbNewLine
					strPicture = strPicture & "     <param name=""movie"" value=""" & ImageUrl & """>" & vbNewLine
					strPicture = strPicture & "     <param name=""quality"" value=""high"">" & vbNewLine
					strPicture = strPicture & "     <embed src=""" & ImageUrl & """ width=""" & width & """ height=""" & height & """ hspace=""" & hspace & """ vspace=""" & vspace & """ align=""" & align & """ quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""></embed>" & vbNewLine
					strPicture = strPicture & "</object>" & vbNewLine
				Case Else
					strPicture = "<a href=""" & HtmlFileUrl & """  target=""" & target & """ title=""" & enchiasp.ChannelModule & "标题：" & Rs("title") & "&#13;&#10;发布时间：" & Rs("WriteTime") & "&#13;&#10;阅览次数：" & Rs("AllHits") & """><img src=""" & ImageUrl & """ width=""" & width & """ height=""" & height & """ hspace=""" & hspace & """ vspace=""" & vspace & """ align=""" & align & """ border=""0""></a>"
				End Select
			End If
			Rs.Close: Set Rs = Nothing
		Else
			strPicture = ""
		End If
		NewsPictureAndText = strPicture & strContent
	End Function
	'================================================
	'函数名：ReadNewsPicAndText
	'作  用：读取图文混排列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadNewsPicAndText(ByVal str)
		Dim strTemp, i, sTempContent
		Dim nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$NewsPictureAndText(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$NewsPictureAndText(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$NewsPictureAndText(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), NewsPictureAndText(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8), ArrayList(9), ArrayList(10), ArrayList(11), ArrayList(12), ArrayList(13), ArrayList(14), ArrayList(15), ArrayList(16), ArrayList(17)))
			Next
		End If
		ReadNewsPicAndText = strTemp
	End Function
	'================================================
	'函数名：SoftPictureAndText
	'作  用：软件图文混排列表
	'================================================
	Public Function SoftPictureAndText(ByVal chanid, ByVal ClassID, ByVal specid, _
		ByVal stype, ByVal height, ByVal width, ByVal maxlen, _
		ByVal maxline, ByVal hspace, ByVal vspace, ByVal align, _
		ByVal divcss, ByVal target, ByVal start, ByVal showpic, _
		ByVal showclass, ByVal showdate, ByVal dateformat)
				
		Dim Rs, SQL, i, strContent, foundstr
		Dim ChildStr, HtmlFileUrl, HtmlFileName, strPicture
		Dim SoftTopic, ClassName, softname, SoftTime
		
		chanid = enchiasp.ChkNumeric(chanid)
		ClassID = enchiasp.ChkNumeric(ClassID)
		specid = enchiasp.ChkNumeric(specid)
		stype = enchiasp.ChkNumeric(stype)
		
		On Error Resume Next
		enchiasp.LoadChannel(chanid)
		
		If CInt(stype) >= 3 And CLng(ClassID) <> 0 Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID = " & chanid & " And ClassID = " & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				SoftPictureAndText = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
			End If
			Rs.Close
		Else
			ChildStr = "0"
		End If
		Select Case CInt(stype)
			Case 0: foundstr = "ORDER BY A.SoftTime DESC ,A.softid DESC"
			Case 1: foundstr = "And A.isBest > 0 ORDER BY A.SoftTime DESC ,A.softid DESC"
			Case 2: foundstr = "ORDER BY A.AllHits DESC ,A.softid DESC"
			Case 3: foundstr = "And A.ClassID in (" & ChildStr & ") ORDER BY A.SoftTime DESC ,A.softid DESC"
			Case 4: foundstr = "And A.ClassID in (" & ChildStr & ") And A.isBest > 0 ORDER BY A.SoftTime DESC ,A.softid DESC"
			Case 5: foundstr = "And A.ClassID in (" & ChildStr & ") ORDER BY A.AllHits DESC ,A.softid DESC"
		Case Else
			foundstr = "ORDER BY A.SoftTime DESC ,A.softid DESC"
		End Select
		If CInt(stype) >= 3 And CLng(ClassID) = 0 Then
			foundstr = "ORDER BY A.SoftTime DESC ,A.softid DESC"
		End If
		If CLng(specid) > 0 Then
			foundstr = "And A.SpecialID =" & CLng(specid) & " " & foundstr
		End If
		SQL = " A.softid,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.AllHits,A.SoftTime,A.HtmlFileDate,A.isBest,"
		SQL = "SELECT TOP " & CInt(maxline) & SQL & " C.ClassName,C.ColorModes,C.FontModes,C.HtmlFileDir FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ChannelID=" & chanid & " " & foundstr & ""
		Set Rs = enchiasp.Execute(SQL)
		i = 0
		strContent = ""
		If Rs.BOF And Rs.EOF Then
			strContent = "还没有添加任何软件！"
		Else
			Do While Not Rs.EOF
				SoftTopic = enchiasp.ReadTopic(Trim(Rs("SoftName") & " " & Rs("SoftVer")), CInt(maxlen))
				SoftTopic = enchiasp.ReadFontMode(SoftTopic, Rs("ColorMode"), Rs("FontMode"))
				ClassName = enchiasp.ReadFontMode(Rs("ClassName"), Rs("ColorModes"), Rs("FontModes"))
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("softid"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) > 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
					ClassName = "[<a href='" & enchiasp.ChannelPath & Rs("HtmlFileDir") & "index" & enchiasp.ChannelHtmlExt & "'>" & ClassName & "</a>]"
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("softid")
					ClassName = "[<a href='" & enchiasp.ChannelPath & "list.asp?classid=" & Rs("ClassID") & "'>" & ClassName & "</a>]"
				End If
				If CInt(showclass) = 1 Then
					ClassName = ClassName
				Else
					ClassName = ""
				End If
				If CInt(showdate) = 1 Then
					SoftTime = enchiasp.ShowDateTime(Rs("SoftTime"), CInt(dateformat))
				Else
					SoftTime = ""
				End If
				softname = "<div " & divcss & ">" & start & ClassName & " <a href=""" & HtmlFileUrl & """ target=""" & target & """ title=""" & enchiasp.ChannelModule & "标题：" & Rs("SoftName") & " " & Rs("SoftVer") & "&#13;&#10;发布时间：" & Rs("SoftTime") & "&#13;&#10;阅览次数：" & Rs("AllHits") & """ class=showlist>" & SoftTopic & "</a>  " & SoftTime & "</div>"
				strContent = strContent & softname
				Rs.MoveNext
				i = i + 1
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		Dim sExtName, ExtName, SoftImage
		If CInt(showpic) = 1 Then
			SQL = " A.softid,A.ClassID,A.SoftName,A.SoftVer,A.AllHits,A.SoftTime,A.HtmlFileDate,A.SoftImage,"
			SQL = "SELECT " & SQL & " C.HtmlFileDir,B.ChannelDir,B.ModuleName,B.BindDomain,B.DomainName,B.IsCreateHtml,B.HtmlExtName,B.HtmlPath,B.HtmlForm,B.HtmlPrefix,B.LeastHotHist FROM ([ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID) INNER JOIN [ECCMS_Channel] B On A.ChannelID=B.ChannelID WHERE A.isAccept>0 And A.ChannelID=" & CInt(chanid) & " And A.SoftImage<>'' " & foundstr & ""
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				strPicture = "<img src='" & enchiasp.SiteUrl & enchiasp.InstallDir & "images/no_pic.gif' width=""" & width & """ height=""" & height & """  hspace=""" & hspace & """ vspace=""" & vspace & """ align=""" & align & """ border=""0"">"
			Else
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("softid"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) <> 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("softid")
				End If
				SoftImage = enchiasp.GetImageUrl(Rs("SoftImage"), enchiasp.ChannelData(1))
				sExtName = Split(Rs("SoftImage"), ".")
				ExtName = sExtName(UBound(sExtName))
				Select Case LCase(ExtName)
				Case "swf", "swi"
					strPicture = "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"" width=""" & width & """ height=""" & height & """ hspace=""" & hspace & """ vspace=""" & vspace & """ align=""" & align & """>" & vbNewLine
					strPicture = strPicture & "     <param name=""movie"" value=""" & SoftImage & """>" & vbNewLine
					strPicture = strPicture & "     <param name=""quality"" value=""high"">" & vbNewLine
					strPicture = strPicture & "     <embed src=""" & SoftImage & """ width=""" & width & """ height=""" & height & """ hspace=""" & hspace & """ vspace=""" & vspace & """ align=""" & align & """ quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""></embed>" & vbNewLine
					strPicture = strPicture & "</object>" & vbNewLine
				Case Else
					strPicture = "<a href=""" & HtmlFileUrl & """  target=""" & target & """ title=""" & enchiasp.ChannelModule & "标题：" & Rs("SoftName") & " " & Rs("SoftVer") & "&#13;&#10;发布时间：" & Rs("SoftTime") & "&#13;&#10;阅览次数：" & Rs("AllHits") & """><img src=""" & SoftImage & """ width=""" & width & """ height=""" & height & """ hspace=""" & hspace & """ vspace=""" & vspace & """ align=""" & align & """ border=""0""></a>"
				End Select
			End If
			Rs.Close: Set Rs = Nothing
		Else
			strPicture = ""
		End If
		SoftPictureAndText = strPicture & strContent
	End Function
	'================================================
	'函数名：ReadSoftPicAndText
	'作  用：读取软件图文混排列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadSoftPicAndText(ByVal str)
		On Error Resume Next
		Dim strTemp, i, sTempContent
		Dim nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		
		strTemp = str
		If InStr(strTemp, "{$SoftPictureAndText(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$SoftPictureAndText(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$SoftPictureAndText(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), SoftPictureAndText(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8), ArrayList(9), ArrayList(10), ArrayList(11), ArrayList(12), ArrayList(13), ArrayList(14), ArrayList(15), ArrayList(16), ArrayList(17)))
			Next
		End If
		ReadSoftPicAndText = strTemp
	End Function
	'================================================
	'函数名：LoadGuestList
	'作  用：装载留言列表
	'参  数：maxnum ----最多留言数
	'        maxlen ----字符长度
	'        newindow ----是否新窗口打开 1=是，0=否
	'        showdate ----是否显示时间 1=是，0=否
	'        DateMode ----时间模式
	'        styles ----风格名称
	'================================================
	Public Function LoadGuestList(ByVal maxnum, ByVal maxlen, ByVal newindow, _
		ByVal showdate, ByVal DateMode, ByVal styles)
		
		Dim Rs, SQL, strContent
		Dim i, ListStyle, GuestTopic, LinkTarget
		Dim WriteTime, lastime, GuestTitle,strChannelDir
		
		On Error Resume Next
		Set Rs = enchiasp.Execute("SELECT TOP " & CInt(maxnum) & " guestid,Topicformat,title,username,WriteTime,lastime,ReplyNum FROM ECCMS_GuestBook WHERE isAccept>0 ORDER BY isTop DESC,lastime DESC,guestid DESC")
		If Rs.BOF And Rs.EOF Then
			LoadGuestList = "没有任何留言!"
			Set Rs = Nothing
			Exit Function
		Else
			i = 0
			strContent = "<table width=""100%"" border=0 cellpadding=2 cellspacing=0>"
			strChannelDir = enchiasp.GetChannelDir(4)
			Do While Not Rs.EOF
				If (i Mod 2) = 0 Then
					ListStyle = Trim(styles) & 1
				Else
					ListStyle = Trim(styles) & 2
				End If
				If CInt(newindow) <> 0 Then
					LinkTarget = " target=""_blank"""
				Else
					LinkTarget = ""
				End If
				If CInt(showdate) <> 0 Then
					WriteTime = enchiasp.ShowDateTime(Rs("WriteTime"), CInt(DateMode))
					lastime = enchiasp.ShowDateTime(Rs("lastime"), CInt(DateMode))
				Else
					WriteTime = ""
					lastime = ""
				End If
				GuestTitle = enchiasp.HTMLEncode(Rs("title"))
				GuestTopic = "<span " & Rs("Topicformat") & ">" & enchiasp.GotTopic(GuestTitle, CInt(maxlen)) & "</span>"
				GuestTopic = "<a href=""" & strChannelDir & "showreply.asp?guestid=" & Rs("guestid") & """ title=""主题：" & GuestTitle & "&#13;&#10;时间：" & Rs("WriteTime") & "&#13;&#10;作者：" & enchiasp.HTMLEncode(Rs("username")) & """" & LinkTarget & ">" & GuestTopic & "</a>"
				strContent = strContent & enchiasp.MainSetting(16)
				strContent = Replace(strContent, "{$GuestID}", Rs("guestid"))
				strContent = Replace(strContent, "{$UserName}", enchiasp.HTMLEncode(Rs("username")))
				strContent = Replace(strContent, "{$GuestTopic}", GuestTopic)
				strContent = Replace(strContent, "{$ListStyle}", ListStyle)
				strContent = Replace(strContent, "{$Number}", i)
				strContent = Replace(strContent, "{$WriteTime}", WriteTime)
				strContent = Replace(strContent, "{$lastime}", lastime)
				Rs.MoveNext
				i = i + 1
			Loop
			strContent = strContent & "</table>"
		End If
		LoadGuestList = strContent
	End Function
	'================================================
	'函数名：ReadGuestList
	'作  用：读取留言列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadGuestList(ByVal str)
		Dim strTemp, i, sTempContent
		Dim nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		
		strTemp = str
		If InStr(strTemp, "{$ReadGuestList(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadGuestList(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadGuestList(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadGuestList(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5)))
			Next
		End If
		ReadGuestList = strTemp
	End Function
	'================================================
	'函数名：LoadPopularSoft
	'作  用：装载排行软件列表
	'参  数：ClassID   ----分类ID
	'        chanid   ----频道ID
	'        stype   ----调用类型
	'        maxline   ----显示列表数
	'        maxlen   ----显示标题长度
	'        showhits   ----是否显示下载数
	'        target   ----连接目标
	'        start   ----标题头标记
	'        styles   ----样式名称
	'================================================
	Public Function LoadPopularSoft(ByVal chanid, ByVal ClassID, ByVal stype, _
		ByVal maxlen, ByVal maxline, ByVal showhits, _
		ByVal target, ByVal start, ByVal styles)
		
		Dim SQL, Rs, foundsql, strHits
		Dim ChildStr, i, strContent
		Dim HtmlFileName, HtmlFileUrl
		Dim NewsTitle, AllHits, strSoftName
		Dim divstyle
		
		chanid = enchiasp.ChkNumeric(chanid)
		ClassID = enchiasp.ChkNumeric(ClassID)
		stype = enchiasp.ChkNumeric(stype)
		If chanid = 0 Then chanid = 1

		On Error Resume Next
		enchiasp.LoadChannel(chanid)
		
		If CLng(ClassID) > 0 And Trim(ClassID) <> "" Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID=" & chanid & " And classid=" & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadPopularSoft = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
				foundsql = "And A.ClassID in (" & ChildStr & ")"
			End If
			Rs.Close
		Else
			ChildStr = "0"
			foundsql = ""
		End If
		
		Select Case CInt(stype)
		Case 1
			foundsql = foundsql & " ORDER BY A.DayHits DESC ,A.softid DESC"
			strHits = "DayHits"
		Case 2
			foundsql = foundsql & " ORDER BY A.WeekHits DESC ,A.softid DESC"
			strHits = "WeekHits"
		Case 3
			foundsql = foundsql & " ORDER BY A.MonthHits DESC ,A.softid DESC"
			strHits = "MonthHits"
		Case 4
			foundsql = foundsql & " And A.isBest>0 ORDER BY A.AllHits DESC ,A.softid DESC"
			strHits = "AllHits"
		Case Else
			foundsql = foundsql & "ORDER BY A.AllHits DESC ,A.softid DESC"
			strHits = "AllHits"
		End Select
		SQL = " A.softid,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.AllHits,A.SoftTime,A.HtmlFileDate,A.isBest,A.DayHits,A.WeekHits,A.MonthHits,"
		SQL = "SELECT TOP " & CInt(maxline) & SQL & " C.ClassName,C.ColorModes,C.FontModes,C.HtmlFileDir FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ChannelID=" & chanid & " " & foundsql
		Set Rs = enchiasp.Execute(SQL)
		i = 0
		strContent = ""
		If Rs.BOF And Rs.EOF Then
			strContent = "还没有找到任何内容！"
		Else
			Do While Not Rs.EOF
				If Trim(styles) <> "" And Trim(styles) <> "0" Then
					If (i Mod 2) = 0 Then
						divstyle = " class=""" & Trim(styles) & "1"""
					Else
						divstyle = " class=""" & Trim(styles) & "2"""
					End If
				End If
				
				NewsTitle = enchiasp.GotTopic(Rs("SoftName") & " " & Rs("SoftVer"), CInt(maxlen))
				NewsTitle = enchiasp.ReadFontMode(NewsTitle, Rs("ColorMode"), Rs("FontMode"))
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("SoftID"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) > 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("softid")
				End If
				If CInt(showhits) > 0 Then
					AllHits = Rs(strHits)
				Else
					AllHits = ""
				End If
				strSoftName = "<div" & divstyle & ">" & start & " <a href=""" & HtmlFileUrl & """ target=""" & target & """ title=""" & enchiasp.ChannelModule & "名称：" & Rs("SoftName") & " " & Rs("SoftVer") & "&#13;&#10;发布时间：" & Rs("SoftTime") & "&#13;&#10;下载总数：" & Rs("AllHits") & """ class=popular>" & NewsTitle & "</a>  " & AllHits & "</div>"
				strContent = strContent & strSoftName
				
				Rs.MoveNext
				i = i + 1
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		LoadPopularSoft = strContent
	End Function
	'================================================
	'函数名：ReadPopularSoft
	'作  用：读取软件排行列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadPopularSoft(ByVal str)
		On Error Resume Next
		Dim strTemp, i, sTempContent
		Dim nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		
		strTemp = str
		If InStr(strTemp, "{$ReadPopularSoft(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadPopularSoft(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadPopularSoft(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadPopularSoft(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8)))
			Next
		End If
		ReadPopularSoft = strTemp
	End Function
	'================================================
	'函数名：LoadPopularArticle
	'作  用：装载排行文章列表
	'参  数：ClassID   ----分类ID
	'        chanid   ----频道ID
	'        stype   ----调用类型
	'        maxline   ----显示列表数
	'        maxlen   ----显示标题长度
	'        showhits   ----是否显示下载数
	'        target   ----连接目标
	'        start   ----标题头标记
	'        styles   ----样式名称
	'================================================
	Public Function LoadPopularArticle(ByVal chanid, ByVal ClassID, ByVal stype, _
		ByVal maxlen, ByVal maxline, ByVal showhits, ByVal target, _
		ByVal start, ByVal styles)

		Dim SQL, Rs, foundsql, strHits
		Dim ChildStr, i, strContent
		Dim HtmlFileName, HtmlFileUrl
		Dim NewsTitle, AllHits, ArticleTitle
		Dim divstyle

		chanid = enchiasp.ChkNumeric(chanid)
		ClassID = enchiasp.ChkNumeric(ClassID)
		stype = enchiasp.ChkNumeric(stype)

		If chanid = 0 Then chanid = 2
		
		On Error Resume Next
		enchiasp.LoadChannel(chanid)
		
		If CLng(ClassID) > 0 And Trim(ClassID) <> "" Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID=" & chanid & " And classid=" & CLng(ClassID)
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadPopularArticle = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
				foundsql = "And A.ClassID in (" & ChildStr & ")"
			End If
			Rs.Close
		Else
			ChildStr = "0"
			foundsql = ""
		End If
		Select Case CInt(stype)
		Case 1
			foundsql = foundsql & " ORDER BY A.DayHits DESC ,A.Articleid DESC"
			strHits = "DayHits"
		Case 2
			foundsql = foundsql & " ORDER BY A.WeekHits DESC ,A.Articleid DESC"
			strHits = "WeekHits"
		Case 3
			foundsql = foundsql & " ORDER BY A.MonthHits DESC ,A.Articleid DESC"
			strHits = "MonthHits"
		Case 4
			foundsql = foundsql & " And A.isBest>0 ORDER BY A.AllHits DESC ,A.Articleid DESC"
			strHits = "AllHits"
		Case Else
			foundsql = foundsql & "ORDER BY A.AllHits DESC ,A.Articleid DESC"
			strHits = "AllHits"
		End Select
		SQL = " A.ArticleID,A.ClassID,A.ColorMode,A.FontMode,A.title,A.BriefTopic,A.AllHits,A.WriteTime,A.HtmlFileDate,A.isBest,A.DayHits,A.WeekHits,A.MonthHits,"
		SQL = "SELECT TOP " & CInt(maxline) & SQL & " C.ClassName,C.ColorModes,C.FontModes,C.HtmlFileDir FROM [ECCMS_Article] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ChannelID=" & chanid & " " & foundsql
		Set Rs = enchiasp.Execute(SQL)
		i = 0
		strContent = ""
		If Rs.BOF And Rs.EOF Then
			strContent = "还没有找到任何软件！"
		Else
			Do While Not Rs.EOF
				If Trim(styles) <> "" And Trim(styles) <> "0" Then
					If (i Mod 2) = 0 Then
						divstyle = " class=""" & Trim(styles) & "1"""
					Else
						divstyle = " class=""" & Trim(styles) & "2"""
					End If
				End If
				NewsTitle = enchiasp.GotTopic(Rs("title"), CInt(maxlen))
				NewsTitle = enchiasp.ReadFontMode(NewsTitle, Rs("ColorMode"), Rs("FontMode"))
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("ArticleID"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) > 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("ArticleID")
				End If
				If CInt(showhits) > 0 Then
					AllHits = Rs(strHits)
				Else
					AllHits = ""
				End If
				ArticleTitle = "<div" & divstyle & ">" & start & " <a href=""" & HtmlFileUrl & """ target=""" & target & """ title=""" & enchiasp.ChannelModule & "标题：" & Rs("title") & "&#13;&#10;发布时间：" & Rs("WriteTime") & "&#13;&#10;阅览次数：" & Rs("AllHits") & """ class=popular>" & NewsTitle & "</a>  " & AllHits & "</div>"
				strContent = strContent & ArticleTitle
				Rs.MoveNext
				i = i + 1
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		LoadPopularArticle = strContent
	End Function
	'================================================
	'函数名：ReadPopularSoft
	'作  用：读取软件排行列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadPopularArticle(ByVal str)
		On Error Resume Next
		Dim strTemp, i, sTempContent
		Dim nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		
		strTemp = str
		If InStr(strTemp, "{$ReadPopularArticle(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadPopularArticle(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadPopularArticle(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadPopularArticle(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8)))
			Next
		End If
		ReadPopularArticle = strTemp
	End Function
	'================================================
	'函数名：LoadPopularFlash
	'作  用：装载排行动画列表
	'参  数：ClassID   ----分类ID
	'        chanid   ----频道ID
	'        stype   ----调用类型
	'        maxline   ----显示列表数
	'        maxlen   ----显示标题长度
	'        showhits   ----是否显示下载数
	'        target   ----连接目标
	'        start   ----标题头标记
	'        styles   ----样式名称
	'================================================
	Public Function LoadPopularFlash(ByVal chanid, ByVal ClassID, ByVal stype, _
		ByVal maxlen, ByVal maxline, ByVal showhits, _
		ByVal target, ByVal start, ByVal styles)
		
		Dim SQL, Rs, foundsql, strHits
		Dim ChildStr, i, strContent
		Dim HtmlFileName, HtmlFileUrl
		Dim NewsTitle, AllHits, strtitle
		Dim divstyle
		
		chanid = enchiasp.ChkNumeric(chanid)
		ClassID = enchiasp.ChkNumeric(ClassID)
		stype = enchiasp.ChkNumeric(stype)
		If chanid = 0 Then chanid = 1

		On Error Resume Next
		enchiasp.LoadChannel(chanid)
		
		If CLng(ClassID) > 0 And Trim(ClassID) <> "" Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID=" & chanid & " And classid=" & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadPopularFlash = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
				foundsql = "And A.ClassID in (" & ChildStr & ")"
			End If
			Rs.Close
		Else
			ChildStr = "0"
			foundsql = ""
		End If
		
		Select Case CInt(stype)
		Case 1
			foundsql = foundsql & " ORDER BY A.DayHits DESC ,A.flashid DESC"
			strHits = "DayHits"
		Case 2
			foundsql = foundsql & " ORDER BY A.WeekHits DESC ,A.flashid DESC"
			strHits = "WeekHits"
		Case 3
			foundsql = foundsql & " ORDER BY A.MonthHits DESC ,A.flashid DESC"
			strHits = "MonthHits"
		Case 4
			foundsql = foundsql & " And A.isBest>0 ORDER BY A.AllHits DESC ,A.flashid DESC"
			strHits = "AllHits"
		Case Else
			foundsql = foundsql & "ORDER BY A.AllHits DESC ,A.flashid DESC"
			strHits = "AllHits"
		End Select
		SQL = " A.flashid,A.ClassID,A.ColorMode,A.FontMode,A.title,A.AllHits,A.addTime,A.HtmlFileDate,A.isBest,A.DayHits,A.WeekHits,A.MonthHits,"
		SQL = "SELECT TOP " & CInt(maxline) & SQL & " C.ClassName,C.ColorModes,C.FontModes,C.HtmlFileDir FROM [ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ChannelID=" & chanid & " " & foundsql
		Set Rs = enchiasp.Execute(SQL)
		i = 0
		strContent = ""
		If Rs.BOF And Rs.EOF Then
			strContent = "还没有找到任何内容！"
		Else
			Do While Not Rs.EOF
				If Trim(styles) <> "" And Trim(styles) <> "0" Then
					If (i Mod 2) = 0 Then
						divstyle = " class=""" & Trim(styles) & "1"""
					Else
						divstyle = " class=""" & Trim(styles) & "2"""
					End If
				End If
				
				NewsTitle = enchiasp.GotTopic(Rs("title"), CInt(maxlen))
				NewsTitle = enchiasp.ReadFontMode(NewsTitle, Rs("ColorMode"), Rs("FontMode"))
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("flashid"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) > 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("flashid")
				End If
				If CInt(showhits) > 0 Then
					AllHits = Rs(strHits)
				Else
					AllHits = ""
				End If
				strtitle = "<div" & divstyle & ">" & start & " <a href=""" & HtmlFileUrl & """ target=""" & target & """ title=""" & enchiasp.ChannelModule & "名称：" & Rs("title") & "&#13;&#10;发布时间：" & Rs("addTime") & "&#13;&#10;下载总数：" & Rs("AllHits") & """ class=popular>" & NewsTitle & "</a>  " & AllHits & "</div>"
				strContent = strContent & strtitle
				
				Rs.MoveNext
				i = i + 1
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		LoadPopularFlash = strContent
	End Function
	
	'================================================
	'函数名：LoadShopList
	'作  用：装载商品列表
	'参  数：ClassID   ----分类ID
	'        ChannelID   ----频道ID
	'        sType   ----调用类型
	'        TopNum   ----显示列表数
	'        strlen   ----显示标题长度
	'        ShowClass   ----是否显示分类
	'        ShowDate   ----是否显示日期
	'        DateMode   ----显示日期模式
	'        newindow   ----新窗口打开
	'================================================
	Public Function LoadShopList(ByVal ChannelID, ByVal ClassID, ByVal SpecialID, _
		ByVal stype, ByVal TopNum, ByVal strLen, ByVal showclass, _
		ByVal showdate, ByVal DateMode, ByVal newindow, ByVal styles)
		
		Dim Rs, SQL, i, strContent, foundstr
		Dim sTradeName, ChildStr, ListStyle, HtmlFileName, BestCode, BestString
		Dim ClassName, HtmlFileUrl, addTime, LinkTarget, TradeTopic, PastPrice, NowPrice
		
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		SpecialID = enchiasp.ChkNumeric(SpecialID)
		stype = enchiasp.ChkNumeric(stype)
		
		On Error Resume Next
		enchiasp.LoadChannel(ChannelID)
		
		If CInt(stype) >= 3 And CLng(ClassID) <> 0 Then
			SQL = "select ChildStr from [ECCMS_Classify] where ChannelID = " & ChannelID & " And ClassID = " & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				LoadShopList = ""
				Exit Function
			Else
				ChildStr = Rs("ChildStr")
			End If
			Rs.Close
		Else
			ChildStr = 0
		End If
		Select Case CInt(stype)
			Case 0: foundstr = "Order By A.addTime Desc ,A.ShopID Desc"
			Case 1: foundstr = "And A.isBest > 0 Order By A.addTime Desc ,A.ShopID Desc"
			Case 2: foundstr = "Order By A.AllHits Desc ,A.ShopID Desc"
			Case 3: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.addTime Desc ,A.ShopID Desc"
			Case 4: foundstr = "And A.ClassID in (" & ChildStr & ") And A.isBest > 0 Order By A.addTime Desc ,A.ShopID Desc"
			Case 5: foundstr = "And A.ClassID in (" & ChildStr & ") Order By A.AllHits Desc ,A.ShopID Desc"
		Case Else
			foundstr = "Order By A.addTime Desc ,A.ShopID Desc"
		End Select
		If CInt(stype) >= 3 And CLng(ClassID) = 0 Then
			foundstr = "Order By A.addTime Desc ,A.ShopID Desc"
		End If
		If CLng(SpecialID) <> 0 Then
			foundstr = "And A.SpecialID =" & CLng(SpecialID) & " " & foundstr
		End If
		SQL = " A.ShopID,A.ClassID,A.TradeName,A.PastPrice,A.NowPrice,A.addTime,A.AllHits,A.HtmlFileDate,A.isBest,"
		SQL = "select Top " & CInt(TopNum) & SQL & " C.ClassName,C.ColorModes,C.FontModes,C.HtmlFileDir,C.UseHtml from [ECCMS_ShopList] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.isAccept>0 And A.ChannelID=" & ChannelID & " " & foundstr & ""
		Set Rs = enchiasp.Execute(SQL)
		i = 0
		If Rs.BOF And Rs.EOF Then
			strContent = "没有添加任何商品！"
		Else
			strContent = "<table width=""100%"" border=0 cellpadding=2 cellspacing=0>"
			Do While Not Rs.EOF
				If (i Mod 2) = 0 Then
					ListStyle = Trim(styles) & 1
				Else
					ListStyle = Trim(styles) & 2
				End If
				If Rs("isBest") <> 0 Then
					BestCode = 2
					BestString = "<font color='" & enchiasp.MainSetting(3) & "'>推荐</font>"
				Else
					BestCode = 1
					BestString = ""
				End If
				strContent = strContent & enchiasp.MainSetting(15)
				sTradeName = enchiasp.GotTopic(Rs("TradeName"), CInt(strLen))
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("ShopID"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				ClassName = enchiasp.ReadFontMode(Rs("ClassName"), Rs("ColorModes"), Rs("FontModes"))
				If CInt(enchiasp.ChannelUseHtml) <> 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
					ClassName = "<a href='" & enchiasp.ChannelPath & Rs("HtmlFileDir") & "index" & enchiasp.ChannelHtmlExt & "'>" & ClassName & "</a>"
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("ShopID")
					ClassName = "<a href='" & enchiasp.ChannelPath & "list.asp?classid=" & Rs("ClassID") & "'>" & ClassName & "</a>"
				End If
				If CInt(showclass) = 0 Then ClassName = ""
				If CInt(showdate) <> 0 Then
					addTime = enchiasp.ShowDateTime(Rs("addTime"), CInt(DateMode))
				Else
					addTime = ""
				End If
				If CInt(newindow) <> 0 Then
					LinkTarget = " target=""_blank"""
				Else
					LinkTarget = ""
				End If
				PastPrice = FormatCurrency(Rs("PastPrice"), , -1)
				NowPrice = FormatCurrency(Rs("NowPrice"), , -1)
				TradeTopic = "<a href='" & HtmlFileUrl & "'" & LinkTarget & " title='" & enchiasp.ChannelModule & "名称：" & Rs("TradeName") & "&#13;&#10;上架时间：" & Rs("addTime") & "&#13;&#10;订购次数：" & Rs("AllHits") & "&#13;&#10;" & enchiasp.ChannelModule & "价格：" & NowPrice & " 元' class=showlist>" & sTradeName & "</a>"
				strContent = Replace(strContent, "{$TradeTopic}", TradeTopic)
				strContent = Replace(strContent, "{$ShopID}", Rs("shopid"))
				strContent = Replace(strContent, "{$InstallDir}", enchiasp.InstallDir)
				strContent = Replace(strContent, "{$TradeName}", sTradeName)
				strContent = Replace(strContent, "{$Title}", Rs("TradeName"))
				strContent = Replace(strContent, "{$DateAndTitle}", Rs("addTime"))
				strContent = Replace(strContent, "{$HtmlFileUrl}", HtmlFileUrl)
				strContent = Replace(strContent, "{$ClassName}", ClassName)
				strContent = Replace(strContent, "[]", "")
				strContent = Replace(strContent, "{$Target}", LinkTarget)
				strContent = Replace(strContent, "{$addTime}", addTime)
				strContent = Replace(strContent, "{$ShopHits}", Rs("AllHits"))
				strContent = Replace(strContent, "{$ListStyle}", ListStyle)
				strContent = Replace(strContent, "{$BestCode}", BestCode)
				strContent = Replace(strContent, "{$BestString}", BestString)
				strContent = Replace(strContent, "{$PastPrice}", PastPrice)
				strContent = Replace(strContent, "{$NowPrice}", NowPrice)
			Rs.MoveNext
			i = i + 1
			Loop
			strContent = strContent & "</table>"
		End If
		Rs.Close: Set Rs = Nothing
		LoadShopList = strContent
	End Function
	'================================================
	'函数名：ReadShopList
	'作  用：读取商品列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadShopList(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent
		Dim arrTempContent, arrTempContents, ArrayList
		On Error Resume Next
		strTemp = str
		If InStr(strTemp, "{$ReadShopList(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadShopList(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadShopList(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadShopList(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8), ArrayList(9), ArrayList(10)))
			Next
		End If
		ReadShopList = strTemp
		Exit Function
	End Function

	
	'================================================
	'函数名：ReadPopularFlash
	'作  用：读取动画排行列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadPopularFlash(ByVal str)
		On Error Resume Next
		Dim strTemp, i, sTempContent
		Dim nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		
		strTemp = str
		If InStr(strTemp, "{$ReadPopularFlash(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadPopularFlash(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadPopularFlash(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadPopularFlash(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8)))
			Next
		End If
		ReadPopularFlash = strTemp
	End Function
	'================================================
	'函数名：LoadSoftType
	'作  用：装载软件类型列表
	'参  数：chanid   ----频道ID
	'        SoftType   ----软件类型
	'        maxline   ----显示列表数
	'        maxlen   ----显示标题长度
	'        showclass   ----是否显示分类
	'        showdate   ----是否显示日期
	'        DateMode   ----显示日期模式
	'        newindow  ----是否新窗口打开连接
	'        styles   ----样式名称
	'================================================
	Public Function LoadSoftType(ByVal chanid, ByVal SoftType, ByVal maxlen, _
		ByVal maxline, ByVal showclass, ByVal showdate, _
		ByVal DateMode, ByVal newindow, ByVal styles)

		Dim SQL, Rs, foundsql, strContent, i
		Dim strSoftName, ChildStr, ListStyle, HtmlFileName, BestCode, BestString
		Dim ClassName, HtmlFileUrl, SoftTime, LinkTarget, SoftTopic
		
		SoftType = enchiasp.CheckStr(SoftType)
		chanid = enchiasp.ChkNumeric(chanid)
		maxline = enchiasp.ChkNumeric(maxline)
		If chanid = 0 Then chanid = 2
		If maxline = 0 Then maxline = 10
		
		On Error Resume Next
		enchiasp.LoadChannel(chanid)
		
		If Trim(SoftType) <> "" Then
			foundsql = "And A.SoftType='" & SoftType & "' Order By A.SoftTime Desc ,A.SoftID Desc"
		Else
			foundsql = "Order By A.SoftTime Desc ,A.SoftID Desc"
		End If
		
		SQL = " A.SoftID,A.ClassID,A.ColorMode,A.FontMode,A.SoftName,A.SoftVer,A.SoftType,A.AllHits,A.SoftTime,A.HtmlFileDate,A.isBest,"
		SQL = "SELECT TOP " & maxline & SQL & " C.ClassName,C.ColorModes,C.FontModes,C.HtmlFileDir,C.UseHtml FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.isAccept>0 And A.ChannelID=" & chanid & " " & foundsql
		Set Rs = enchiasp.Execute(SQL)
		i = 0
		strContent = ""
		If Rs.BOF And Rs.EOF Then
			strContent = "还没有找到任何软件！"
		Else
			strContent = "<table width=""100%"" border=0 cellpadding=2 cellspacing=0>"
			Do While Not Rs.EOF
				If (i Mod 2) = 0 Then
					ListStyle = Trim(styles) & 1
				Else
					ListStyle = Trim(styles) & 2
				End If
				If Rs("isBest") <> 0 Then
					BestCode = 2
					BestString = "<font color='" & enchiasp.MainSetting(3) & "'>推荐</font>"
				Else
					BestCode = 1
					BestString = ""
				End If
				strContent = strContent & enchiasp.MainSetting(14)
				strSoftName = enchiasp.GotTopic(Rs("SoftName") & " " & Rs("SoftVer"), CInt(maxlen))
				strSoftName = enchiasp.ReadFontMode(strSoftName, Rs("ColorMode"), Rs("FontMode"))
				
				ClassName = enchiasp.ReadFontMode(Rs("ClassName"), Rs("ColorModes"), Rs("FontModes"))
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("SoftID"), enchiasp.ChannelHtmlExt, enchiasp.ChannelPrefix, enchiasp.ChannelHtmlForm, "")
				If CInt(enchiasp.ChannelUseHtml) <> 0 Then
					HtmlFileUrl = enchiasp.ChannelPath & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.ChannelHtmlPath) & HtmlFileName
					ClassName = "<a href='" & enchiasp.ChannelPath & Rs("HtmlFileDir") & "index" & enchiasp.ChannelHtmlExt & "'>" & ClassName & "</a>"
				Else
					HtmlFileUrl = enchiasp.ChannelPath & "show.asp?id=" & Rs("SoftID")
					ClassName = "<a href='" & enchiasp.ChannelPath & "list.asp?classid=" & Rs("ClassID") & "'>" & ClassName & "</a>"
				End If
				If CInt(showclass) = 0 Then ClassName = ""
				If CInt(showdate) <> 0 Then
					SoftTime = enchiasp.ShowDateTime(Rs("SoftTime"), CInt(DateMode))
				Else
					SoftTime = ""
				End If
				If CInt(newindow) <> 0 Then
					LinkTarget = " target=""_blank"""
				Else
					LinkTarget = ""
				End If
				
				SoftTopic = "<a href='" & HtmlFileUrl & "'" & LinkTarget & " title='" & enchiasp.ChannelModule & "名称：" & Rs("SoftName") & "&#13;&#10;发布时间：" & Rs("SoftTime") & "&#13;&#10;下载次数：" & Rs("AllHits") & "' class=showlist>" & strSoftName & "</a>"
				strContent = Replace(strContent, "{$SoftTopic}", SoftTopic)
				strContent = Replace(strContent, "{$InstallDir}", enchiasp.InstallDir)
				strContent = Replace(strContent, "{$SoftName}", strSoftName)
				strContent = Replace(strContent, "{$Title}", Rs("SoftName"))
				strContent = Replace(strContent, "{$DateAndTitle}", Rs("SoftTime"))
				strContent = Replace(strContent, "{$HtmlFileUrl}", HtmlFileUrl)
				strContent = Replace(strContent, "{$ClassName}", ClassName)
				strContent = Replace(strContent, "[]", "")
				strContent = Replace(strContent, "{$Target}", LinkTarget)
				strContent = Replace(strContent, "{$SoftTime}", SoftTime)
				strContent = Replace(strContent, "{$SoftHits}", Rs("AllHits"))
				strContent = Replace(strContent, "{$ListStyle}", ListStyle)
				strContent = Replace(strContent, "{$BestCode}", BestCode)
				strContent = Replace(strContent, "{$BestString}", BestString)
				Rs.MoveNext
				i = i + 1
			Loop
			strContent = strContent & "</table>"
		End If
		Set Rs = Nothing
		LoadSoftType = strContent
	End Function
	'================================================
	'函数名：ReadSoftType
	'作  用：读取软件类型列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadSoftType(ByVal str)
		On Error Resume Next
		Dim strTemp, i, sTempContent
		Dim nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		
		strTemp = str
		If InStr(strTemp, "{$ReadSoftType(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadSoftType(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadSoftType(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadSoftType(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5), ArrayList(6), ArrayList(7), ArrayList(8)))
			Next
		End If
		ReadSoftType = strTemp
	End Function
	'================================================
	'函数名：LoadUserRank
	'作  用：装用户排行列表
	'================================================
	Public Function LoadUserRank(ByVal stype,ByVal grade,ByVal maxline,ByVal styles)
		Dim SQL, Rs, foundsql, strContent, i
		Dim ListStyle,username
		
		stype = enchiasp.CheckNumeric(stype)
		grade = enchiasp.CheckNumeric(grade)
		maxline = enchiasp.CheckNumeric(maxline)
		If maxline = 0 Then maxline = 10
		If stype = 1 Then
			foundsql = "ORDER BY JoinTime DESC,userid DESC"
		ElseIf stype = 2 Then
			foundsql = "ORDER BY LastTime DESC,userid DESC"
		ElseIf stype = 3 Then
			foundsql = "ORDER BY userpoint DESC,userid DESC"
		Else
			foundsql = "ORDER BY userlogin DESC,userid DESC"
		End If
		If grade > 0 Then
			SQL = "SELECT TOP " & maxline & " userid,username,userpoint,userlogin FROM [ECCMS_User] WHERE UserGrade=" & grade & " " & foundsql
		Else
			SQL = "SELECT TOP " & maxline & " userid,username,userpoint,userlogin FROM [ECCMS_User] " & foundsql
		End If
		Set Rs = enchiasp.Execute(SQL)
		i = 0
		strContent = ""
		If Not (Rs.BOF And Rs.EOF) Then
			strContent = "<table width=""100%"" border=0 cellpadding=2 cellspacing=0>"
			Do While Not Rs.EOF
				If (i Mod 2) = 0 Then
					ListStyle = Trim(styles) & 1
				Else
					ListStyle = Trim(styles) & 2
				End If
				username = "<a href=""" & enchiasp.InstallDir & "user/userlist.asp?userid=" & Rs("userid") & """ target=""_blank"">" & Rs("username") & "</a>"
				strContent = strContent & enchiasp.MainSetting(23)
				strContent = Replace(strContent, "{$ListStyle}", ListStyle)
				strContent = Replace(strContent, "{$InstallDir}", enchiasp.InstallDir)
				strContent = Replace(strContent, "{$UserName}", username)
				strContent = Replace(strContent, "{$username}", Rs("username"))
				strContent = Replace(strContent, "{$UserID}", Rs("userid"))
				strContent = Replace(strContent, "{$UserLogin}", Rs("userlogin"))
				strContent = Replace(strContent, "{$UserPoint}", Rs("userpoint"))
				Rs.MoveNext
				i = i + 1
				strContent = Replace(strContent, "{$OrderID}", i)
			Loop
			strContent = strContent & "</table>"
		End If
		Rs.Close: Set Rs = Nothing
		
		LoadUserRank = strContent
	End Function
	'================================================
	'函数名：ReadUserRank
	'作  用：读取用户排行列表
	'参  数：str ----原字符串
	'================================================
	Public Function ReadUserRank(ByVal str)
		On Error Resume Next
		Dim strTemp, i, sTempContent
		Dim nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		
		strTemp = str
		If InStr(strTemp, "{$ReadUserRank(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadUserRank(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadUserRank(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadUserRank(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3)))
			Next
		End If
		ReadUserRank = strTemp
	End Function
	'================================================
	'函数名：LoadStatistic
	'作  用：装载频道统计
	'参  数：moduleid ----所属模块
	'        ChannelID ----频道ID
	'        strClass ----所调用的分类ID或者软件类型
	'        stype ----统计类型，0=全部统计，1=今日更新统计，2=点击数统计，3=软件容量统计
	'================================================
	Public Function LoadStatistic(ByVal moduleid, ByVal ChannelID, ByVal strClass, ByVal stype)

		moduleid = enchiasp.CheckNumeric(moduleid)
		ChannelID = enchiasp.CheckNumeric(ChannelID)
		stype = enchiasp.CheckNumeric(stype)
		
		Dim Rs, SQL, StatCount
		Dim foundsql, ClassID, ChildStr
		
		ClassID = enchiasp.CheckNumeric(strClass)
		On Error Resume Next
		LoadStatistic = 0
		If ClassID > 0 Then
			SQL = "SELECT ChildStr FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID=" & ClassID
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				ChildStr = 0
			Else
				ChildStr = Rs("ChildStr")
			End If
			Rs.Close: Set Rs = Nothing
			foundsql = "And ChannelID=" & ChannelID & " And ClassID in (" & ChildStr & ")"
		Else
			foundsql = "And ChannelID=" & ChannelID
		End If
		Select Case moduleid
		Case 1
			If stype = 1 Then
				If isSqlDataBase = 1 Then
					SQL = "SELECT COUNT(ArticleID) FROM ECCMS_Article WHERE isAccept>0 " & foundsql & " And Datediff(d,WriteTime,GetDate())=0"
				Else
					SQL = "SELECT COUNT(ArticleID) FROM ECCMS_Article WHERE isAccept>0 " & foundsql & " And WriteTime>=Date()"
				End If
			ElseIf stype = 2 Then
				SQL = "SELECT SUM(AllHits) FROM ECCMS_Article WHERE isAccept>0 " & foundsql
			ElseIf stype = 4 Then
				SQL = "SELECT SUM(DayHits) FROM ECCMS_Article WHERE isAccept>0 " & foundsql
			Else
				SQL = "SELECT COUNT(ArticleID) FROM ECCMS_Article WHERE isAccept>0 " & foundsql
			End If
		Case 2
			If Not IsNumeric(strClass) Then
				foundsql = foundsql & " And SoftType='" & enchiasp.CheckStr(strClass) & "'"
			End If
			If stype = 1 Then
				If isSqlDataBase = 1 Then
					SQL = "SELECT COUNT(softid) FROM ECCMS_SoftList WHERE isAccept>0 " & foundsql & " And Datediff(d,SoftTime,GetDate())=0"
				Else
					SQL = "SELECT COUNT(softid) FROM ECCMS_SoftList WHERE isAccept>0 " & foundsql & " And SoftTime>=Date()"
				End If
			ElseIf stype = 2 Then
				SQL = "SELECT SUM(AllHits) FROM ECCMS_SoftList WHERE isAccept>0 " & foundsql
			ElseIf stype = 3 Then
				SQL = "SELECT SUM(SoftSize) FROM ECCMS_SoftList WHERE isAccept>0 " & foundsql
			ElseIf stype = 4 Then
				SQL = "SELECT SUM(DayHits) FROM ECCMS_SoftList WHERE isAccept>0 " & foundsql
			Else
				SQL = "SELECT COUNT(softid) FROM ECCMS_SoftList WHERE isAccept>0 " & foundsql
			End If
		Case 4
			If stype = 1 Then
				If isSqlDataBase = 1 Then
					SQL = "SELECT COUNT(GuestID) FROM ECCMS_GuestBook WHERE isAccept>0 And Datediff(d,WriteTime,GetDate())=0"
				Else
					SQL = "SELECT COUNT(GuestID) FROM ECCMS_GuestBook WHERE isAccept>0 And WriteTime>=Date()"
				End If
			Else
				SQL = "SELECT COUNT(GuestID) FROM ECCMS_GuestBook WHERE isAccept>0"
			End If
		Case 5
			If stype = 1 Then
				If isSqlDataBase = 1 Then
					SQL = "SELECT COUNT(flashid) FROM ECCMS_FlashList WHERE isAccept>0 " & foundsql & " And Datediff(d,addTime,GetDate())=0"
				Else
					SQL = "SELECT COUNT(flashid) FROM ECCMS_FlashList WHERE isAccept>0 " & foundsql & " And addTime>=Date()"
				End If
			ElseIf stype = 2 Then
				SQL = "SELECT SUM(AllHits) FROM ECCMS_FlashList WHERE isAccept>0 " & foundsql
			ElseIf stype = 3 Then
				SQL = "SELECT SUM(filesize) FROM ECCMS_FlashList WHERE isAccept>0 " & foundsql
			ElseIf stype = 4 Then
				SQL = "SELECT SUM(DayHits) FROM ECCMS_FlashList WHERE isAccept>0 " & foundsql
			Else
				SQL = "SELECT COUNT(flashid) FROM ECCMS_FlashList WHERE isAccept>0 " & foundsql
			End If
		Case Else
			If stype = 1 Then
				If isSqlDataBase = 1 Then
					SQL = "SELECT COUNT(userid) FROM ECCMS_User WHERE Datediff(d,JoinTime,GetDate())=0"
				Else
					SQL = "SELECT COUNT(userid) FROM ECCMS_User WHERE JoinTime>=Date()"
				End If
			Else
				SQL = "SELECT COUNT(userid) FROM ECCMS_User"
			End If
		End Select
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			StatCount = 0
		Else
			StatCount = CCur(Rs(0))
			If (moduleid = 2 And stype = 3) Or (moduleid = 5 And stype = 3) Then
				StatCount = Round(StatCount / 1024 / 1024, 3)
				StatCount = FormatNumber(StatCount, 3, -1)
			End If
		End If
		Rs.Close: Set Rs = Nothing
		LoadStatistic = StatCount
	End Function
	'================================================
	'函数名：ReadStatistic
	'作  用：读取频道统计
	'参  数：str ----原字符串
	'================================================
	Public Function ReadStatistic(ByVal str)
		On Error Resume Next
		Dim strTemp, i, sTempContent
		Dim nTempContent, ArrayList
		Dim arrTempContent, arrTempContents

		strTemp = str
		'On Error Resume Next
		If InStr(strTemp, "{$ReadStatistic(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadStatistic(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadStatistic(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")

			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")

				strTemp = Replace(strTemp, arrTempContents(i), LoadStatistic(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3)))
			Next
		End If
		ReadStatistic = strTemp
	End Function

	Public Function ShowIndex(ByVal isHtml)
		Dim HtmlContent
		enchiasp.LoadTemplates 0, 1, 0
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", enchiasp.InstallDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", enchiasp.InstallDir)
		If Len(enchiasp.HtmlSetting(1)) < 2 Then
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", "首页")

		Else
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.HtmlSetting(1))
		End If
		
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
		HtmlContent = ReadAnnounceContent(HtmlContent, 0)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = ReadArticlePic(HtmlContent)
		HtmlContent = ReadSoftPic(HtmlContent)
		HtmlContent = ReadArticleList(HtmlContent)
		HtmlContent = ReadSoftList(HtmlContent)
		HtmlContent = ReadShopList(HtmlContent)
		HtmlContent = ReadFlashList(HtmlContent)
		HtmlContent = ReadFlashPic(HtmlContent)
		HtmlContent = ReadShopPic(HtmlContent)
		HtmlContent = ReadFriendLink(HtmlContent)
		HtmlContent = ReadNewsPicAndText(HtmlContent)
		HtmlContent = ReadSoftPicAndText(HtmlContent)
		HtmlContent = ReadGuestList(HtmlContent)
		HtmlContent = ReadAnnounceList(HtmlContent)
		HtmlContent = ReadPopularArticle(HtmlContent)
		HtmlContent = ReadPopularSoft(HtmlContent)
		HtmlContent = ReadPopularFlash(HtmlContent)
		HtmlContent = ReadSoftType(HtmlContent)
		HtmlContent = ReadStatistic(HtmlContent)
		HtmlContent = ReadUserRank(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", enchiasp.InstallDir)
		If isHtml Then
			ShowIndex = HtmlContent
		Else
			Response.Write HtmlContent
		End If
	
	End Function

	Public Function ShowfengmianIndex()	
		Dim Rs, SQL,strContent,i

		On Error Resume Next
		If enchiasp.usefengmian = "1" then
			SQL = "SELECT * FROM [ECCMS_fengmian] WHERE isuse=1"
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				'strContent = "没有设置封面模板！"
				response.Redirect "index_gb.asp"
			Else
			
				strContent=rs("nr")
				'替换关键字
				strContent = Replace(strContent, "{$Keyword}", enchiasp.KeyWords)

				'替换封面图片10个
				for i=1 to 10
					if rs("pic"&i&"")<>"" then
						strContent = Replace(strContent, "{$fengmianpic"&i&"}", rs("pic"& i &""))
					end if
				next 
			
				'替换封面FLASH5个
				for i=1 to 5
					if rs("flash"&i&"")<>"" then
						strContent = Replace(strContent, "{$fengmianflash"&i&"}", rs("flash"& i &""))
					end if
				next 				
				
				'替换版权
				strContent = Replace(strContent, "{$Copyright}", enchiasp.Copyright)
				'替换安装路径
				strContent = Replace(strContent, "{$InstallDir}", enchiasp.InstallDir)
				'替换样式表
				if  rs("css")<>"" then
					strContent = Replace(strContent, "{$fengmiancss}", rs("css"))
				else
					strContent = Replace(strContent, "{$fengmiancss}", "")
				end if
				'替换背景颜色
				if rs("bg")<>"" then
					strContent = Replace(strContent, "{$fengmianbg}", rs("bg"))
				else
					strContent = Replace(strContent, "{$fengmianbg}", "#ffffff")
				end if
				'替换网站名字
				strContent = Replace(strContent, "{$WebSiteName}", enchiasp.SiteName)
				'替换背景图片
				strContent = Replace(strContent, "{$fengmianbgimg}", rs("bgimg"))

				'替换LOGO
				strContent = Replace(strContent, "{$fengmianlogo}", rs("logo"))

				'替换背景音乐
				strContent = Replace(strContent, "{$fengmianbgmidi}", rs("bgmidi"))

				'替换所有的目录
				strContent = Replace(strContent, "{$fengmianinstalldir}", "fengmian/"& rs("usedir")&"/")
				
				Response.Write strContent
			End If
			Rs.Close: Set Rs = Nothing	
		else
			'不使用封面模板
			response.Redirect "index_gb.asp"
		end if


	End Function

'2007-07-15增加识别友情标签,注意特殊使用方法
Public Function Showfrink
		Dim HtmlContent
		enchiasp.LoadTemplates 9999, 6, 0
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", enchiasp.InstallDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", enchiasp.InstallDir)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", "合作伙伴")
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
		HtmlContent = ReadAnnounceContent(HtmlContent, 0)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = ReadArticlePic(HtmlContent)
		HtmlContent = ReadSoftPic(HtmlContent)
		HtmlContent = ReadArticleList(HtmlContent)
		HtmlContent = ReadSoftList(HtmlContent)
		HtmlContent = ReadShopList(HtmlContent)
		HtmlContent = ReadFlashList(HtmlContent)
		HtmlContent = ReadFlashPic(HtmlContent)
		HtmlContent = ReadShopPic(HtmlContent)
		HtmlContent = ReadFriendLink(HtmlContent)
		HtmlContent = ReadNewsPicAndText(HtmlContent)
		HtmlContent = ReadSoftPicAndText(HtmlContent)
		HtmlContent = ReadGuestList(HtmlContent)
		HtmlContent = ReadAnnounceList(HtmlContent)
		HtmlContent = ReadPopularArticle(HtmlContent)
		HtmlContent = ReadPopularSoft(HtmlContent)
		HtmlContent = ReadPopularFlash(HtmlContent)
		HtmlContent = ReadSoftType(HtmlContent)
		HtmlContent = ReadStatistic(HtmlContent)
		HtmlContent = ReadUserRank(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", enchiasp.InstallDir)
		Response.Write HtmlContent
	End Function


End Class
%>