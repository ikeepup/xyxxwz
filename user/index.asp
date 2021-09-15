<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/classmenu.asp"-->

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
Dim HtmlContent,ChannelRootDir
ChannelRootDir = enchiasp.InstallDir & "user/"
enchiasp.LoadTemplates 9999, 4, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
'--频道目录
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","用户登录")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = ReadFriendLink(HtmlContent)
Response.Write HtmlContent
CloseConn




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


%>