<!--#include file="config.asp" -->
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
Dim Rs,SQL,ErrMsg
Dim flashid,downid,ClassID,title
Dim DownFileName,DownAddress,PointNum,UserGroup
Dim DownloadUrl,strDownAddress,strDownloadUrl,User_Group

flashid = enchiasp.ChkNumeric(Request.Querystring("id"))
downid = enchiasp.ChkNumeric(Request.Querystring("downid"))
If flashid = 0 Then
	ErrMsg = ErrMsg & "<li>错误的系统参数!请输入正确的软件ID</li>"
	FoundErr=True
End If
If Not enchiasp.CheckOuterUrl Then
	ErrMsg = ErrMsg & "<li>非法下载，请不要盗链本站资源！</li>"
	FoundErr=True
End If

Call BeginDownload

If FoundErr Then
	Returnerr(ErrMsg)
End If
Set enchicms = Nothing
CloseConn

Sub BeginDownload()
	If FoundErr Then Exit Sub
	Dim GroupSetting,GroupName,gradeid,rootid

	If Trim(enchiasp.membergrade) <> "" Then
		gradeid = CInt(enchiasp.membergrade)
	Else
		gradeid = 0
	End If
	User_Group = 0
	GroupSetting = Split(enchiasp.UserGroupSetting(gradeid), "|||")
	GroupName = GroupSetting(UBound(GroupSetting))
	If CInt(GroupSetting(31)) = 0 Then
		ErrMsg = ErrMsg & "<li>对不起！你是" & GroupName & "；不能下载本站资源。</li>"
		FoundErr=True
		Exit Sub
	End If

	SQL = "SELECT ClassID,title,DownAddress,PointNum,UserGroup FROM ECCMS_FlashList WHERE ChannelID="& ChannelID &" And isAccept > 0 And flashid=" & flashid
	Set Rs = enchiasp.Execute(SQL)
	If Rs.EOF And Rs.BOF Then
		ErrMsg = ErrMsg & "<li>对不起~！没有找到你想下载的软件。</li>"
		FoundErr=True
		Set Rs = Nothing
		Exit Sub
	Else
		ClassID = Rs("ClassID")
		title = Rs("title")
		DownAddress = Rs("DownAddress")
		PointNum = Rs("PointNum")
		UserGroup = Rs("UserGroup")
		
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = enchiasp.Execute("SELECT UserGroup FROM ECCMS_Classify WHERE ChannelID="& ChannelID &" And ClassID="& ClassID)
	If Rs("UserGroup") > gradeid Then
		ErrMsg = ErrMsg & "<li>您没有登录或者你的会员级别不够！</li><li>如果你是本站会员, 请先<a href=""../user/"">登陆</a>后再下载!</li>"
		FoundErr=True
		Set Rs = Nothing
		Exit Sub
	End If
	Set Rs = Nothing
	If downid > 0 Then
		SQL = "SELECT rootid,downid,DownloadPath,UserGroup,DownPoint FROM ECCMS_DownServer WHERE ChannelID="& ChannelID &" And isLock=0 And downid=" & downid
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			ErrMsg = ErrMsg & "<li>注意：您所下载的文件不存在。</li>"
			FoundErr=True
			Set Rs = Nothing
			Exit Sub
		Else
			rootid = Rs("rootid")
			DownloadUrl = Trim(Rs("DownloadPath"))
			User_Group = Rs("UserGroup")
			If User_Group > gradeid Then
				ErrMsg = ErrMsg & "<li>注意：此下载服务器是会员专用；</li><li>如果你是本站会员, 请先<a href=""../user/"">登陆</a>后再下载!</li>"
				FoundErr=True
				Set Rs = Nothing
				Exit Sub
			End If
			If Rs("UserGroup") > 0 Then
				PointNum = Rs("DownPoint")
				CheckUserDownload flashid,PointNum,User_Group,GroupName
			Else
				PointNum = PointNum
			End If
		End If
		Rs.Close:Set Rs = Nothing
		DownloadUrl = Trim(DownloadUrl & DownAddress)
	Else
		DownloadUrl = Trim(DownAddress)
	End If
	If CInt(UserGroup) > 0 And User_Group = 0 Then
		If Trim(enchiasp.memberName) = "" Then
			ErrMsg = ErrMsg & "<li>此文件是会员软件，非会员不能下载。 如果你是本站会员请先<a href=""../user/"">登陆</a>!</li>"
			FoundErr=True
			Exit Sub
		End If
		CheckUserDownload flashid,PointNum,UserGroup,GroupName
	End If
	If FoundErr=True Then Exit Sub
	Response.Redirect (DownloadUrl)
End Sub

Function CheckUserDownload(flashid,PointNum,UserGroup,GroupName)
	If FoundErr Then Exit Function
	If CInt(enchiasp.membergrade) = 999 Then Exit Function
	Dim Rss
	On Error Resume Next
	Dim CookiesID,userpoint,UserGrade,UserToday
	If CInt(enchiasp.memberclass) > 0 Then
		Set Rss = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT userid,UserGrade,UserClass,ExpireTime FROM ECCMS_User WHERE UserClass>0 And username='" & enchiasp.memberName & "' And userid=" & enchiasp.memberid
		Rss.Open SQL,Conn,1,3
		If Rss.BOF And Rss.EOF Then
			ErrMsg = ErrMsg & "<li>非法操作~！</li>"
			FoundErr=True
			Set Rss = Nothing
			Exit Function
		Else
			If DateDiff("D", CDate(Rss("ExpireTime")), Now()) > 0 Or Rss("UserClass") = 999 Then
				ErrMsg = ErrMsg & "<li>对不起！您的会员已到期，不能下载此软件；</li><li>如果你要下载此软件请联系管理员。</li>"
				FoundErr=True
				Set Rss = Nothing
				Exit Function
			Else
				Set Rss = Nothing
				Exit Function
			End If
		End If
		Rss.Close:Set Rss = Nothing
	End If
	CookiesID = "flashid_" & flashid
	If Trim(Request.Cookies("DownLoadFlash")) = "" Then
		Response.Cookies("DownLoadFlash")("userip") = enchiasp.GetUserIP
		Response.Cookies("DownLoadFlash").Expires = Date + 1
	End If
	
	If CLng(Request.Cookies("DownLoadFlash")(CookiesID)) <> CLng(flashid) And CInt(UserGroup) > 0 Then
		Set Rss = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT userid,UserGrade,userpoint,UserToday,ExpireTime FROM ECCMS_User WHERE username='" & enchiasp.memberName & "' And userid=" & enchiasp.memberid
		Rss.Open SQL,Conn,1,3
		If Rss.BOF And Rss.EOF Then
			ErrMsg = ErrMsg & "<li>非法操作~！</li>"
			FoundErr=True
			Set Rss = Nothing
			Exit Function
		Else
			userpoint = Rss("userpoint")
			UserGrade = Rss("UserGrade")
			UserToday = Rss("UserToday")
			UserToday = Split(UserToday, "|")
			If UserGrade < UserGroup  Then
				ErrMsg = ErrMsg & "<li>您的级别不够，下载此软件需要<font color=blue>"& GroupName &"</font>以上级别的会员；</li><li>如果你要下载此软件请联系管理员。</li>"
				FoundErr=True
				Set Rss = Nothing
				Exit Function
			End If
			
			If CInt(enchiasp.memberclass) = 0 Then
				If userpoint < PointNum Then
					ErrMsg = ErrMsg & "<li>对不起!您的点数不足。不能下载此软件</li><li>下载本软件所需的点数："& PointNum &"</li><li>如果你确实要下载此软件请到<a href=""../user/"">会员中心</a>充值。</li>"
					FoundErr=True
					Set Rss = Nothing
					Exit Function
				Else
					Rss("userpoint").Value = CLng(Rss("userpoint") - PointNum)
					Rss.Update
					Response.Cookies("DownLoadFlash")(CookiesID) = flashid
				End If
				
			End If
		End If
		Rss.Close:Set Rss = Nothing
	End If
End Function
%>