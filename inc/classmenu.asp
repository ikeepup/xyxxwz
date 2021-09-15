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
'函数名：LoadClassMenu
'作  用：装载分类菜单
'参  数：ChannelID ----频道ID
'================================================
Public Function LoadClassMenu(ByVal ChannelID, ByVal ClassID, ByVal TopNum, _
	ByVal PerRowNum, ByVal Compart, ByVal styles)
	
	Dim Rs, SQL, i, strContent, foundsql
	Dim rsClass, ParentID, Child, TotalNumber
	Dim LinkTarget, HtmlFileUrl, ClassName, strClass
	dim temp
	
	LoadClassMenu = ""
	'显示所有的分类
	if ClassID="all" then
		ClassID=999999
	end if
	
	
	'判断是否以树形形式显示
	if ClassID="alltree" then
		'树形
		ClassID=999999
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		If Not IsNumeric(TopNum) Then Exit Function
		If Not IsNumeric(PerRowNum) Then Exit Function
		If styles <> "0" And styles <> "" Then
			strClass = " class=" & Trim(styles)
		Else
			strClass = ""
		End If
		foundsql = "SELECT TOP " & TopNum & " C.ClassID,C.depth,C.ClassName,C.ColorModes,C.FontModes,C.Readme,C.Child,C.LinkTarget,C.TurnLink,C.TurnLinkUrl,C.HtmlFileDir,C.UseHtml,B.ChannelDir,B.StopChannel,B.IsCreateHtml,B.HtmlExtName,B.HtmlPath FROM [ECCMS_Classify] C inner join [ECCMS_Channel] B On C.ChannelID=B.ChannelID WHERE C.ChannelID = " & CLng(ChannelID)
		If CLng(ClassID) <> 0 and CLng(ClassID) <> 999999 Then
			Set rsClass = enchiasp.Execute("SELECT parentid,Child FROM [ECCMS_Classify] WHERE ChannelID = " & CLng(ChannelID) & " And ClassID = " & CLng(ClassID))
			If rsClass.BOF And rsClass.EOF Then
				Exit Function
			Else
				ParentID = rsClass("parentid")
				Child = rsClass("Child")
			End If
			rsClass.Close: Set rsClass = Nothing
			If Child <> 0 Then
				SQL = foundsql & " And C.Parentid = " & CLng(ClassID) & " Order By C.orders,C.ClassID"
			Else
				SQL = foundsql & " And C.Parentid = " & CLng(ParentID) & " Order By C.orders,C.rootid"
			End If
		Else
			if CLng(ClassID)=999999 then
				SQL = foundsql & "  Order By C.rootid,C.ClassID"
			else
				SQL = foundsql & " And C.depth = 0 Order By C.rootid,C.ClassID"
			end if
		End If
		Set Rs = CreateObject("ADODB.Recordset")
		Rs.Open SQL, Conn, 1, 1
		enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
		If Rs.BOF And Rs.EOF Then
			Exit Function
		Else
			If Rs("StopChannel") <> 0 Then
				LoadClassMenu = ""
				Exit Function
			End If
			Do While Not Rs.EOF
				If Rs("LinkTarget") <> 0 Then
					LinkTarget = " target=""_blank"""
				Else
					LinkTarget = ""
				End If
				ClassName = enchiasp.ReadFontMode(Rs("ClassName"), Rs("ColorModes"), Rs("FontModes"))
				If Rs("TurnLink") <> 0 Then
					ClassName = "<a href='" & Rs("TurnLinkUrl") & "'" & LinkTarget & strClass & ">" & ClassName & "</a>"
				Else
					If Rs("IsCreateHtml") <> 0 Then
						ClassName = "<a href='" & enchiasp.InstallDir & Rs("ChannelDir") & Rs("HtmlFileDir") & "'" & LinkTarget & strClass & " title='" & Rs("Readme") & "'>" & ClassName & "</a>"
					Else
						ClassName = "<a href='" & enchiasp.InstallDir & Rs("ChannelDir") & "list.asp?classid=" & Rs("ClassID") & "'" & LinkTarget & strClass & "  title='" & Rs("Readme") & "'>" & ClassName & "</a>"
					End If
				End If
				temp=""
				for i=0 to rs("depth")-1
					temp = Compart &temp
				next
				
				strContent=strContent&temp
				strContent = strContent & ClassName & "<br>" 
				'strContent=enchiasp.LoadSelectClass(ChannelID)
				

			Rs.MoveNext
			Loop
		End If
	else
		'非树形
		ChannelID = enchiasp.ChkNumeric(ChannelID)
		ClassID = enchiasp.ChkNumeric(ClassID)
		If Not IsNumeric(TopNum) Then Exit Function
		If Not IsNumeric(PerRowNum) Then Exit Function
		If styles <> "0" And styles <> "" Then
			strClass = " class=" & Trim(styles)
		Else
			strClass = ""
		End If
		foundsql = "SELECT TOP " & TopNum & " C.ClassID,C.depth,C.ClassName,C.ColorModes,C.FontModes,C.Readme,C.Child,C.LinkTarget,C.TurnLink,C.TurnLinkUrl,C.HtmlFileDir,C.UseHtml,B.ChannelDir,B.StopChannel,B.IsCreateHtml,B.HtmlExtName,B.HtmlPath FROM [ECCMS_Classify] C inner join [ECCMS_Channel] B On C.ChannelID=B.ChannelID WHERE C.ChannelID = " & CLng(ChannelID)
		If CLng(ClassID) <> 0 and CLng(ClassID) <> 999999 Then
			Set rsClass = enchiasp.Execute("SELECT parentid,Child FROM [ECCMS_Classify] WHERE ChannelID = " & CLng(ChannelID) & " And ClassID = " & CLng(ClassID))
			If rsClass.BOF And rsClass.EOF Then
				Exit Function
			Else
				ParentID = rsClass("parentid")
				Child = rsClass("Child")
			End If
			rsClass.Close: Set rsClass = Nothing
			If Child <> 0 Then
				SQL = foundsql & " And C.Parentid = " & CLng(ClassID) & " Order By C.orders,C.ClassID"
			Else
				SQL = foundsql & " And C.Parentid = " & CLng(ParentID) & " Order By C.orders,C.rootid"
			End If
		Else
			if CLng(ClassID)=999999 then
				SQL = foundsql & "  Order By C.rootid,C.ClassID"
			else
				SQL = foundsql & " And C.depth = 0 Order By C.rootid,C.ClassID"
			end if
		End If
		Set Rs = CreateObject("ADODB.Recordset")
		Rs.Open SQL, Conn, 1, 1
		enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
		If Rs.BOF And Rs.EOF Then
			Exit Function
		Else
			If Rs("StopChannel") <> 0 Then
				LoadClassMenu = ""
				Exit Function
			End If
			i = 0
			TotalNumber = Rs.RecordCount
			Do While Not Rs.EOF
				i = i + 1
				If Rs("LinkTarget") <> 0 Then
					LinkTarget = " target=""_blank"""
				Else
					LinkTarget = ""
				End If
				ClassName = enchiasp.ReadFontMode(Rs("ClassName"), Rs("ColorModes"), Rs("FontModes"))
				If Rs("TurnLink") <> 0 Then
					ClassName = "<a href='" & Rs("TurnLinkUrl") & "'" & LinkTarget & strClass & ">" & ClassName & "</a>"
				Else
					If Rs("IsCreateHtml") <> 0 Then
						ClassName = "<a href='" & enchiasp.InstallDir & Rs("ChannelDir") & Rs("HtmlFileDir") & "'" & LinkTarget & strClass & " title='" & Rs("Readme") & "'>" & ClassName & "</a>"
					Else
						ClassName = "<a href='" & enchiasp.InstallDir & Rs("ChannelDir") & "list.asp?classid=" & Rs("ClassID") & "'" & LinkTarget & strClass & "  title='" & Rs("Readme") & "'>" & ClassName & "</a>"
					End If
				End If
				strContent = strContent & ClassName
				If i Mod CInt(PerRowNum) = 0 Or i = TotalNumber Then
					If i = TotalNumber Then
						strContent = strContent
					Else
						strContent = strContent & "<br>"
					End If
				Else
					strContent = strContent & " " & Compart & " "
				End If
			Rs.MoveNext
			Loop
		End If
	end if

	
	
	
	Rs.Close: Set Rs = Nothing
	LoadClassMenu = strContent
End Function



'================================================
'函数名：ReadClassMenu
'作  用：读取分类菜单
'参  数：str ----原字符串
'================================================
Public Function ReadClassMenu(ByVal str)
	Dim strTemp, i
	Dim sTempContent, nTempContent, ArrayList
	Dim arrTempContent, arrTempContents
	On Error Resume Next
	strTemp = str
	
	If InStr(strTemp, "{$ReadClassMenu(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadClassMenu(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadClassMenu(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadClassMenu(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4), ArrayList(5)))
			Next

	End If
	
	ReadClassMenu = strTemp
End Function
'================================================
'函数名：LoadClassMenubar
'作  用：装载分类菜单栏
'参  数：ChannelID ----频道
'================================================
Public Function LoadClassMenubar(ByVal ChannelID, ByVal ClassID, _
	ByVal TopNum, ByVal PerRowNum, ByVal frontstr)
	
	Dim Rs, SQL, i, strContent, foundsql
	Dim rsClass, ParentID, Child
	Dim LinkTarget, HtmlFileUrl, ClassName, strClass
	
	LoadClassMenubar = ""
	ChannelID = enchiasp.ChkNumeric(ChannelID)
	ClassID = enchiasp.ChkNumeric(ClassID)
	If Not IsNumeric(TopNum) Then Exit Function
	If Not IsNumeric(PerRowNum) Then Exit Function
	If frontstr <> "0" And frontstr <> "" Then
		frontstr = frontstr
	Else
		frontstr = ""
	End If
	foundsql = "SELECT TOP " & TopNum & " C.ClassID,C.depth,C.ClassName,C.ColorModes,C.FontModes,C.Readme,C.Child,C.LinkTarget,C.TurnLink,C.TurnLinkUrl,C.HtmlFileDir,C.UseHtml,C.ShowCount,B.ChannelDir,B.StopChannel,B.ModuleName,B.IsCreateHtml,B.HtmlExtName,B.HtmlPath FROM [ECCMS_Classify] C INNER JOIN [ECCMS_Channel] B On C.ChannelID=B.ChannelID where C.ChannelID = " & CInt(ChannelID)
	If CLng(ClassID) <> 0 Then
		Set rsClass = enchiasp.Execute("SELECT parentid,Child FROM [ECCMS_Classify] WHERE ChannelID = " & CInt(ChannelID) & " And ClassID = " & CLng(ClassID))
		If rsClass.BOF And rsClass.EOF Then
			Exit Function
		Else
			ParentID = rsClass("parentid")
			Child = rsClass("Child")
		End If
		rsClass.Close: Set rsClass = Nothing
		If Child <> 0 Then
			SQL = foundsql & " And C.Parentid = " & CLng(ClassID) & " Order By C.orders,C.ClassID"
		Else
			SQL = foundsql & " And C.Parentid = " & CLng(ParentID) & " Order By C.orders,C.rootid"
		End If
	Else
		SQL = foundsql & " And C.depth = 0 Order By C.rootid,C.ClassID"
	End If
	Set Rs = enchiasp.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		Exit Function
	Else
		If Rs("StopChannel") <> 0 Then
			LoadClassMenubar = ""
			Exit Function
		End If
		strContent = "<table border=0 cellpadding=1 cellspacing=3 class=tabmenubar>" & vbCrLf
		Do While Not Rs.EOF
			strContent = strContent & "<tr>" & vbCrLf
			For i = 1 To CInt(PerRowNum)
				strContent = strContent & "<td class=tdmenubar>"
				If Not Rs.EOF Then
					If Rs("LinkTarget") <> 0 Then
						LinkTarget = " target=""_blank"""
					Else
						LinkTarget = ""
					End If
					If Rs("ClassID") = CLng(ClassID) Then
						strClass = " class=distinct"
					Else
						strClass = " class=menubar"
					End If
					ClassName = enchiasp.ReadFontMode(Rs("ClassName"), Rs("ColorModes"), Rs("FontModes"))
					If Rs("TurnLink") <> 0 Then
						ClassName = "<a href='" & Rs("TurnLinkUrl") & "'" & LinkTarget & strClass & ">" & ClassName & "</a>"
					Else
						If Rs("IsCreateHtml") <> 0 Then
							ClassName = "<a href='" & enchiasp.InstallDir & Rs("ChannelDir") & Rs("HtmlFileDir") & "'" & LinkTarget & strClass & " title='" & Rs("Readme") & "&#13;&#10;" & Rs("ModuleName") & "数：" & Rs("ShowCount") & "'>" & ClassName & "</a>"
						Else
							ClassName = "<a href='" & enchiasp.InstallDir & Rs("ChannelDir") & "list.asp?classid=" & Rs("ClassID") & "'" & LinkTarget & strClass & "  title='" & Rs("Readme") & "&#13;&#10;" & Rs("ModuleName") & "数：" & Rs("ShowCount") & "'>" & ClassName & "</a>"
						End If
					End If
					strContent = strContent & frontstr & ClassName
					strContent = strContent & "</td>" & vbCrLf
					Rs.MoveNext
				Else
					strContent = strContent & "</td>" & vbCrLf
				End If
			Next
			strContent = strContent & "</tr>" & vbCrLf
		Loop
		strContent = strContent & "</table>" & vbCrLf
	End If
	Rs.Close: Set Rs = Nothing
	LoadClassMenubar = strContent
End Function
'================================================
'函数名：ReadClassMenubar
'作  用：读取分类菜单栏
'参  数：str ----原字符串
'================================================
Public Function ReadClassMenubar(str)
	Dim strTemp, i
	Dim sTempContent, nTempContent, ArrayList
	Dim arrTempContent, arrTempContents
	On Error Resume Next
	strTemp = str
	If InStr(strTemp, "{$ReadClassMenubar(") > 0 Then
		sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadClassMenubar(", ")}", 1)
		nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadClassMenubar(", ")}", 0)
		arrTempContents = Split(sTempContent, "|||")
		arrTempContent = Split(nTempContent, "|||")
		For i = 0 To UBound(arrTempContents)
			ArrayList = Split(arrTempContent(i), ",")
			strTemp = Replace(strTemp, arrTempContents(i), LoadClassMenubar(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3), ArrayList(4)))
		Next
	End If
	ReadClassMenubar = strTemp
End Function
Public Sub isWeb_Query()
	Dim keyword
	keyword = Replace(Request("keyword"), "'", "")
	
	Response.Write "<div id=""Seardata"" style=""height:500px;"">"
	Response.Write "<iframe name=""WebSearch"" id=""WebSearch"" frameborder=""0"" width=""100%"" height=""100%"" scrolling=""auto"" src=""http://www.enchi.com.cn/search.asp?word="&keyword&"""></iframe>"
	Response.Write "</div>"
	Response.Write "<script language=""JavaScript"">" & vbNewLine
	Response.Write "<!--" & vbNewLine
	Response.Write "var obj=parent.document.getElementById(""searchmain"");" & vbNewLine
	Response.Write "var SearchData = document.getElementById(""Seardata"");" & vbNewLine
	Response.Write "obj.style.height=(parent.document.getElementById(""searchmain"").offsetHeight)+'px';" & vbNewLine
	Response.Write "obj.innerHTML = SearchData.innerHTML;" & vbNewLine
	Response.Write "//-->" & vbNewLine
	Response.Write "</script>" & vbNewLine
End Sub
Public Function SearchObj()
	Dim strTemp,keyword
	keyword = Replace(Request("keyword"), "'", "")
	strTemp = "<script language=""JavaScript"">" & vbNewLine
	strTemp = strTemp & "<!--" & vbNewLine
	strTemp = strTemp & "var ToUrl=""search.asp?act=isweb&keyword=" & keyword & "&s=1"";" & vbNewLine
	strTemp = strTemp & "var HFrame = document.getElementById(""hiddenquery"")" & vbNewLine
	strTemp = strTemp & "var obj = document.getElementById(""searchmain"");" & vbNewLine
	strTemp = strTemp & "if (HFrame){" & vbNewLine
	strTemp = strTemp & "	HFrame.src=ToUrl;" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "if (obj){" & vbNewLine
	strTemp = strTemp & "	obj.style.height=""1024"";" & vbNewLine
	strTemp = strTemp & "	obj.style.display=='none'" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "//-->" & vbNewLine
	strTemp = strTemp & "</script>" & vbNewLine
	SearchObj = strTemp
End Function

'================================================
'函数名：ShowListPage
'作  用：通用分页
'================================================
Public Function ShowListPage(CurrentPage, Pcount, totalrec, PageNum, strLink, ListName)
	Dim strTemp
	On Error Resume Next
	strTemp = vbNewLine & "<script>"
	strTemp = strTemp & "ShowListPage("
	strTemp = strTemp & CurrentPage
	strTemp = strTemp & ","
	strTemp = strTemp & Pcount
	strTemp = strTemp & ","
	strTemp = strTemp & totalrec
	strTemp = strTemp & ","
	strTemp = strTemp & PageNum
	strTemp = strTemp & ",'"
	strTemp = strTemp & strLink
	strTemp = strTemp & "','"
	strTemp = strTemp & ListName
	strTemp = strTemp & "');"
	strTemp = strTemp & "</script>" & vbNewLine
	ShowListPage = strTemp
End Function
'================================================
'函数名：ShowHtmlPage
'作  用：通用HTML分页
'================================================
Public Function ShowHtmlPage(CurrentPage, Pcount, totalrec, PageNum, strLink, ExtName, ListName)
	Dim strTemp
	On Error Resume Next
	strTemp = vbNewLine & "<script>"
	strTemp = strTemp & "ShowHtmlPage("
	strTemp = strTemp & CurrentPage
	strTemp = strTemp & ","
	strTemp = strTemp & Pcount
	strTemp = strTemp & ","
	strTemp = strTemp & totalrec
	strTemp = strTemp & ","
	strTemp = strTemp & PageNum
	strTemp = strTemp & ",'"
	strTemp = strTemp & strLink
	strTemp = strTemp & "','"
	strTemp = strTemp & ExtName
	strTemp = strTemp & "','"
	strTemp = strTemp & ListName
	strTemp = strTemp & "');"
	strTemp = strTemp & "</script>" & vbNewLine
	ShowHtmlPage = strTemp
End Function
%>