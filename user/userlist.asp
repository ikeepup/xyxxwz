<!--#include file="config.asp"-->
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
Dim HtmlContent,TempListContent,ChannelRootDir
Dim userid
dim username
username=Request("name")
userid = enchiasp.ChkNumeric(Request("userid"))
ChannelRootDir = enchiasp.InstallDir & "user/"
enchiasp.LoadTemplates 9999, 5, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
'--频道目录
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir)
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent =ReadFriendLink(HtmlContent)
If userid = 0 Then
	if username<>"" then
		Call ShowUserInfo
	else
		Call ShowUserList
	end if
Else
	Call ShowUserInfo
End If

If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
CloseConn
Public Sub ShowUserList()
	Dim Rs,SQL,i,j,forbid
	Dim maxperpage,CurrentPage,Pcount,totalrec,totalnumber
	Dim strList,strName,RowCode,strContent,strUserName
	Dim strHomePage,strUserMail,strShowPage
	
	forbid = enchiasp.ChkNumeric(enchiasp.HtmlSetting(17))
	If forbid = 2 Then
		ErrMsg = enchiasp.HtmlSetting(18)
		Founderr = True
		Exit Sub
	End If
	If forbid = 1 Then
		If CInt(enchiasp.membergrade) = 0 Then
			ErrMsg = enchiasp.HtmlSetting(19)
			Founderr = True
			Exit Sub
		End If
	End If
	maxperpage = enchiasp.ChkNumeric(enchiasp.HtmlSetting(11))
	If maxperpage = 0 Then maxperpage = 20
	CurrentPage = enchiasp.ChkNumeric(Request("page"))
	If CurrentPage = 0 Then CurrentPage = 1
	'If Not IsObject(Conn) Then ConnectionDatabase
	SQL = "SELECT userid,username,nickname,UserGrade,UserGroup,UserClass,UserLock,userpoint,usermoney,TrueName,UserSex,usermail,HomePage,oicq,JoinTime,ExpireTime,LastTime,userlogin FROM [ECCMS_User] ORDER BY JoinTime DESC ,userid DESC"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open SQL,Conn,1,1
	If Not (Rs.BOF And Rs.EOF) Then
		totalrec = Rs.RecordCount
		Pcount = CLng(totalrec / maxperpage)  '得到总页数
		If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > Pcount Then CurrentPage = Pcount
		Rs.PageSize = maxperpage
		Rs.AbsolutePage = CurrentPage
		i = 0
		j = (CurrentPage - 1) * maxperpage + 1
		Do While Not Rs.EOF And i < maxperpage
			If Not Response.IsClientConnected Then ResponseEnd
			If (i mod 2) = 0 Then
				RowCode = 1
			Else
				RowCode = 2
			End If
			strContent = strContent & enchiasp.HtmlSetting(13)
			strUserName = "<a href=""?userid=" & Rs("userid") & """>" & Rs("username") & "</a>"
			strContent = Replace(strContent, "{$UserName}", strUserName)
			strContent = Replace(strContent, "{$UserID}", Rs("userid"))
			If IsNull(Rs("userlogin")) Then
				strContent = Replace(strContent, "{$UserLogin}", 0)
			Else
				strContent = Replace(strContent, "{$UserLogin}", Rs("userlogin"))
			End If
			strContent = Replace(strContent, "{$UserPoint}", Rs("userpoint"))
			strContent = Replace(strContent, "{$UserSex}", Rs("UserSex"))
			strContent = Replace(strContent, "{$UserQQ}", enchiasp.ChkNull(Rs("oicq")))
			strContent = Replace(strContent, "{$LastTime}", Rs("LastTime"))
			strContent = Replace(strContent, "{$DateAndTime}", Rs("JoinTime"))
			strContent = Replace(strContent, "{$OrderID}", j)
			strUserMail = "<a href=""mailto:" & Rs("usermail") & """ target=""_blank"" title=""给此用户发送邮件"">电子信箱</a>"
			strContent = Replace(strContent, "{$UserMail}", strUserMail)
			strContent = Replace(strContent, "{$UserGroup}", Rs("UserGroup"))
			If enchiasp.CheckNull(Rs("HomePage")) Then
				strHomePage = "<a href=""" & Rs("HomePage") & """ target=""_blank"" title=""点击查看用户主页"">用户主页</a>"
				strContent = Replace(strContent, "{$HomePage}", strHomePage)
			Else
				strContent = Replace(strContent, "{$HomePage}", "没有主页")
			End If
			Rs.movenext
			i = i + 1
			j = j + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	strShowPage = ShowPages(CurrentPage,Pcount,totalrec,maxperpage,"")
	TempListContent = enchiasp.HtmlSetting(12) & strContent & enchiasp.HtmlSetting(14)
	HtmlContent = Replace(HtmlContent,"{$UserManageContent}", TempListContent)
	HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strShowPage)
	HtmlContent = Replace(HtmlContent,"{$PageTitle}",enchiasp.HtmlSetting(10))
	Response.Write HtmlContent
End Sub

Sub ShowUserInfo()
	'没有用户姓名，没有用户ID
	If userid = 0 and username="" Then
	 Exit Sub
	end if
	Dim Rs,SQL,forbid
	Dim strContent,strHomePage,strUserMail
	
	forbid = enchiasp.ChkNumeric(enchiasp.HtmlSetting(17))
	If forbid = 2 Then
		ErrMsg = enchiasp.HtmlSetting(18)
		Founderr = True
		Exit Sub
	End If
	If forbid = 1 Then
		If CInt(enchiasp.membergrade) = 0 Then
			ErrMsg = enchiasp.HtmlSetting(19)
			Founderr = True
			Exit Sub
		End If
	End If
	If username<>"" Then
		SQL = "SELECT userid,username,nickname,UserGrade,UserGroup,UserClass,UserLock,userpoint,usermoney,TrueName,UserSex,usermail,HomePage,oicq,JoinTime,ExpireTime,LastTime,userlogin FROM [ECCMS_User] WHERE username='" & username &"'"
	else
		SQL = "SELECT userid,username,nickname,UserGrade,UserGroup,UserClass,UserLock,userpoint,usermoney,TrueName,UserSex,usermail,HomePage,oicq,JoinTime,ExpireTime,LastTime,userlogin FROM [ECCMS_User] WHERE userid=" & userid

	end if
	Set Rs = enchiasp.Execute(SQL)
	strContent = ""
	If Not (Rs.BOF And Rs.EOF) Then
		strContent = enchiasp.HtmlSetting(16)
		strContent = Replace(strContent, "{$UserName}", Rs("username"))
		strContent = Replace(strContent, "{$UserID}", Rs("userid"))
		strContent = Replace(strContent, "{$UserGroup}", Rs("UserGroup"))
		If IsNull(Rs("userlogin")) Then
			strContent = Replace(strContent, "{$UserLogin}", 0)
		Else
			strContent = Replace(strContent, "{$UserLogin}", Rs("userlogin"))
		End If
		strContent = Replace(strContent, "{$UserPoint}", Rs("userpoint"))
		strContent = Replace(strContent, "{$UserSex}", Rs("UserSex"))
		strContent = Replace(strContent, "{$UserQQ}", enchiasp.ChkNull(Rs("oicq")))
		strContent = Replace(strContent, "{$LastTime}", Rs("LastTime"))
		strContent = Replace(strContent, "{$DateAndTime}", Rs("JoinTime"))
		strUserMail = "<a href=""mailto:" & Rs("usermail") & """ target=""_blank"" title=""给此用户发送邮件"">" & Rs("usermail") & "</a>"
		strContent = Replace(strContent, "{$UserMail}", strUserMail)
		If enchiasp.CheckNull(Rs("HomePage")) Then
			strHomePage = "<a href=""" & Rs("HomePage") & """ target=""_blank"" title=""点击查看用户主页"">" & Rs("HomePage") & "</a>"
			strContent = Replace(strContent, "{$HomePage}", strHomePage)
		Else
			strContent = Replace(strContent, "{$HomePage}", "没有主页")
		End If
	End If
	Rs.Close:Set Rs = Nothing
	HtmlContent = Replace(HtmlContent,"{$UserManageContent}", strContent)
	HtmlContent = Replace(HtmlContent,"{$PageTitle}",enchiasp.HtmlSetting(15))
	Response.Write HtmlContent
End Sub


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