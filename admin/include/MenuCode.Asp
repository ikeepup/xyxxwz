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
'================================================
'函数名：CreationSpecialMenu
'作  用：创建专题栏目菜单
'================================================
Private Sub CreationSpecialMenu()
	Dim strTemp, SpecialPath,strContext
	Dim LinkTarget,IsCreateHtml,Topicformat
	On Error Resume Next
	'此频道是否取胜了生成HTML功能
	IsCreateHtml = CInt(enchiasp.IsCreateHtml)
	SQL = "SELECT SpecialID,SpecialName,Topicformat,Readme,Reopen,SpecialDir,ChangeLink,SpecialUrl FROM ECCMS_Special WHERE ChannelID = "& ChannelID &" ORDER BY orders,SpecialID"
	Set Rs = enchiasp.Execute(SQL)
	Do While Not Rs.EOF
		If Rs("Reopen") <> 0 Then
			LinkTarget = " target=_blank "
		Else
			LinkTarget = " "
		End If
		If Rs("ChangeLink") <> 0 Then
			'如果此专题是外部连接启用此连接URL
			SpecialPath = Rs("SpecialUrl")
		Else
			If IsCreateHtml <> 0 Then
				SpecialPath = ChannelPath & "special/" & Rs("SpecialDir")
			Else
				SpecialPath = ChannelPath & "special.asp?sid=" & Rs("SpecialID")
			End If
		End If
		Topicformat = fixjs(Rs("Topicformat"))
		'显示JS文件的格式
		strTemp = strTemp & "document.write("" · <a href='" & SpecialPath & "'" & LinkTarget & " class=specialmenu><span " & Topicformat & ">" & fixjs(Rs("SpecialName")) & "</span></a><br>"");" & vbCrLf
		Rs.movenext
	Loop
	Rs.Close
	Set Rs = Nothing
	'---------------------------------------------//Begin
	'热门和推荐菜单连接开始,如果不需要请注释掉此段代码
	sModuleName = fixjs(sModuleName)
	If IsCreateHtml <> 0 Then
		strTemp = strTemp & "document.write("" · <a href='" & ChannelPath & "special/" & enchiasp.HtmlPrefix & "Best001" & enchiasp.HtmlExtName & "' class=specialmenu>推荐" & sModuleName & "</a><br>"");" & vbCrLf
		strTemp = strTemp & "document.write("" · <a href='" & ChannelPath & "special/" & enchiasp.HtmlPrefix & "Hot001" & enchiasp.HtmlExtName & "' class=specialmenu>热门" & sModuleName & "</a><br>"");" & vbCrLf
		strTemp = strTemp & "document.write("" · <a href='" & ChannelPath & "special/" & enchiasp.HtmlPrefix & "New001" & enchiasp.HtmlExtName & "' class=specialmenu>最新" & sModuleName & "</a>"");" & vbCrLf
	Else
		strTemp = strTemp & "document.write("" · <a href='" & ChannelPath & "showbest.asp' class=specialmenu>推荐" & sModuleName & "</a><br>"");" & vbCrLf
		strTemp = strTemp & "document.write("" · <a href='" & ChannelPath & "showhot.asp' class=specialmenu>热门" & sModuleName & "</a><br>"");" & vbCrLf
		strTemp = strTemp & "document.write("" · <a href='" & ChannelPath & "shownew.asp' class=specialmenu>最新" & sModuleName & "</a>"");" & vbCrLf
	End If
	If CInt(enchiasp.Modules) = 2 Then
		strTemp = strTemp & "document.write(""<br> · <a href='" & ChannelPath & "showtype.asp' class=specialmenu>最近更新</a>"");" & vbCrLf
	End If
	'菜单结束
	'----------------------------------------------///End
	Dim strFilePath
	'strTemp = Left(strTemp, InstrRev(strTemp, "<br>", -1, 1) - 1)
	strContext = strTemp
	'生成专题菜单的JS文件路径
	strFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "js/specialmenu.js"
	enchiasp.CreatedTextFile strFilePath, strContext
	strContext = Replace(strTemp, "<br>", " ┇ ",1,-1,1)
	strContext = Replace(strContext, "· ", "")
	strFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "js/specmenu.js"
	enchiasp.CreatedTextFile strFilePath, strContext
	If Request("stype") = 2 Then
		OutputScript "创建专题菜单的JS文件成功!" ,"?ChannelID=" & ChannelID
	End If
	Exit Sub
End Sub
'================================================
'函数名：CreationJsMenu
'作  用：创建分类的JS下拉式菜单
'================================================
Private Sub CreationJsMenu()
	Dim strTemp, ChildClass
	Dim LinkTarget,ClassUrl,ClassName,IsCreateHtml
	IsCreateHtml = CInt(enchiasp.IsCreateHtml)
	On Error Resume Next
	SQL = "SELECT * FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And depth = 0 ORDER BY rootid,orders"
	Set Rs = enchiasp.Execute(SQL)
	i = 0
	Do While Not Rs.EOF
		If Rs("LinkTarget") <> 0 Then
			LinkTarget = " target='_blank'"
		Else
			LinkTarget = ""
		End If
		If Rs("TurnLink") <> 0 Then
			ClassUrl = Rs("TurnLinkUrl")
		Else
			If IsCreateHtml <> 0 Then
				ClassUrl = ChannelPath & Rs("HtmlFileDir")
			Else
				ClassUrl = ChannelPath & "list.asp?classid=" & Rs("ClassID")
			End If
			strOption = strOption &"<option value='" & Rs("classid") & "'>" & fixjs(Rs("ClassName")) & "</option>"
		End If
		ChildClass = LoadChildClass(Rs("classid"))
		ClassName = enchiasp.ReadFontMode(Rs("ClassName"), Rs("ColorModes"), Rs("FontModes"))
		strTemp = strTemp & "document.write("" | <a href='" & ClassUrl & "'" & LinkTarget & ChildClass & " class='navbar'>" & fixjs(ClassName) & "</a>"");" & vbCrLf
		Rs.movenext
		i = i + 1
	Loop
	Rs.Close
	Set Rs = Nothing
	Dim strFilePath
	strTemp = Replace(strTemp,"|", "", 1, 1, 1)
	strFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "js/classmenu.js"
	enchiasp.CreatedTextFile strFilePath, strTemp
	Call CreationSearch
	Call CreationSpecialMenu
	enchiasp.LoadTemplates 0, 1, 0
	Dim strChannelMenu
	strChannelMenu = "document.write(""" & fixjs(enchiasp.ChannelMenu) & """);"
	strChannelMenu = Replace(strChannelMenu, "<br>", "", 1, -1, 1)
	enchiasp.CreatedTextFile "../inc/channel.js",strChannelMenu
	If Request("stype") = 1 Then
		OutHintScript "生成导航菜单的JS文件成功!"
	End If
	Exit Sub
End Sub
'================================================
'函数名：LoadChildClass
'作  用：载入子分类
'================================================
private Function LoadChildClass(classid)
	Dim rsChild, strContent
	Dim LinkTarget,strClassName,IsCreateHtml
	IsCreateHtml = CInt(enchiasp.IsCreateHtml)
	On Error Resume Next
	Set rsChild = enchiasp.Execute("SELECT * FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And parentid = " & classid & " ORDER BY orders,classid")
	If rsChild.EOF And rsChild.BOF Then
		Set rsChild = Nothing
		LoadChildClass = ""
		Exit Function
	End If
	Do While Not rsChild.EOF
		If rsChild("LinkTarget") <> 0 Then
			LinkTarget = " target=_blank"
		Else
			LinkTarget = ""
		End If
		strClassName = fixjs(rsChild("ClassName"))
		If rsChild("TurnLink") <> 0 Then
			strClassName = "<a href=" & rsChild("TurnLinkUrl") & LinkTarget & ">" & strClassName & "</a>"
		Else
			If IsCreateHtml <> 0 Then
				strClassName = "<a href=" & ChannelPath & rsChild("HtmlFileDir") & LinkTarget & ">" & strClassName & "</a>"
			Else
				strClassName = "<a href=" & ChannelPath & "list.asp?classid=" & rsChild("ClassID") & LinkTarget & ">" & strClassName & "</a>"
			End If
		End If
		strContent = strContent & "<div class=menuitems>" & strClassName & "<div>"
		rsChild.movenext
	Loop
	rsChild.Close
	Set rsChild = Nothing
	strContent = " onMouseOver=\""showmenu(event,'" & strContent & "')\"""
	LoadChildClass = strContent
	Exit Function
End Function
'================================================
'函数名：CreationSearch
'作  用：创建搜索表单
'================================================
Private Sub CreationSearch()
	Dim strContent,strModules
	If CInt(enchiasp.modules) <> 1 Then
		strModules = sModuleName & "名称"
	Else
		strModules = sModuleName & "标题"
	End If
	strContent = "<table border='0' cellpadding='0' cellspacing='0'>"
	strContent = strContent & "<form method='Get' name='formsearch' action='" & ChannelPath & "search.asp'>" 
	strContent = strContent & "<tr><td height='28' align='center'><select name='act' size='1'>" 
	strContent = strContent & "<option value='Topic' selected>" & strModules & "</option>" 
	strContent = strContent & "<option value='Content'>" & sModuleName & "内容</option>" 
	strContent = strContent & "<option value='isWeb'>网页搜索</option>"
	strContent = strContent & "</select>&nbsp;<select name='classid'><option value=''>所有分类</option>" 
	strContent = strContent & strOption
	strContent = strContent & "</select>&nbsp;<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>&nbsp;<input type='submit' name='btn' class=Button value=' 搜索 '></td></tr></form></table>"
	strContent = "document.write(""" & strContent & """);"
	Dim strFilePath
	strFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "js/search.js"
	enchiasp.CreatedTextFile strFilePath, strContent
	Exit Sub
End Sub
%>