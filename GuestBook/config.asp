<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/cls_public.asp"-->
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
Dim ChannelID,HtmlContent,ChannelRootDir,FoundErr,ErrMsg
FoundErr = False
ChannelID = 4
enchiasp.ReadChannel(ChannelID)
ChannelRootDir = enchiasp.InstallDir & enchiasp.ChannelDir
'=============================================================
'函数名：GuestStation
'作  用：取留言者的身份
'参  数：userid   ----用户ID
'=============================================================
Function GuestStation(ByVal userid)
	On Error Resume Next
	If userid = 0 Or Not IsNumeric(userid) Then
		GuestStation = "游客"
		Exit Function
	End If
	Dim rsUser,sqlUser
	sqlUser = "SELECT UserGroup FROM ECCMS_User WHERE userid ="& userid
	Set rsUser = enchiasp.Execute(sqlUser)
	If rsUser.BOF And rsUser.EOF Then
		GuestStation = "游客"
		Set rsUser = Nothing
		Exit Function
	End If
	GuestStation = rsUser("UserGroup")
	Set rsUser = Nothing
End Function
%>
