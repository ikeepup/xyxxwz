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
Dim AdminName, AdminPass, AdminID, ErrorStr
Dim SQLAdmin, RsAdmin, AdminRandomCode
ErrorStr = "<li>确认身份失败！您没有使用当前功能的权限。</li><li>如果有什么问题，请联系管理员。</li>"
If InStr(enchiasp.ScriptName, "editor") > 0 Or InStr(enchiasp.ScriptName, "admin_label") > 0 Or InStr(enchiasp.ScriptName, "admin_collect") > 0 Then AdminPage = True
'If enchiasp.CheckPost = False And AdminPage = False  Then
	'ErrMsg = "<br><li><font color=red>您提交的数据不合法，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</font></li><li>因为你执行了非法操作，<a href=logout.asp target=_top class=showmeun>请您退出本系统！</a></li>"
	'Response.Redirect("showerr.asp?action=error&message=" & server.URLEncode(errmsg) & "")
	'Response.End
'End If
Call AdminCookiesToSession
AdminName = enchiasp.CheckStr(Session("AdminName"))      '管理员名称
AdminPass = enchiasp.CheckStr(Session("AdminPass"))      '管理员密码
AdminID = enchiasp.ChkNumeric(Session("AdminID"))                    '管理员ID
AdminRandomCode = Trim(Session("AdminRandomCode"))     '管理员登陆随机码
If AdminName = "" Then
	ErrMsg = ErrMsg + "<li>您没有进入本页面的权限!本次操作已被记录!<li>可能您还没有登陆或者不具有使用当前功能的权限!请联系管理员.<li>本页面为[<font color=red>管理员</font>]专用,请先<a href=admin_klogin.asp class=showmeun target=_top>登陆</a>后进入。"
	Response.redirect ("showerr.asp?action=error&Message=" & Server.URLEncode(ErrMsg) & "")
	Response.End
End If
SQLAdmin ="select isLock,RandomCode,isAloneLogin,useip from ECCMS_Admin where username='" & AdminName & "' And password='" & AdminPass & "' And id="& AdminID
Set RsAdmin = enchiasp.Execute(SQLAdmin)
If RsAdmin.BOF And RsAdmin.EOF Then
	Session.Abandon
	Response.Cookies(Admin_Cookies_Name) = ""
	RsAdmin.Close:set RsAdmin = Nothing
	Response.Redirect "admin_klogin.asp"
Else
	If RsAdmin("isLock") <> 0 Then
		ErrMsg = "<li>你的用户名已被锁定,你不能登陆！如要开通此帐号，请联系管理员。</li>"
		RsAdmin.Close:set RsAdmin = Nothing
		Response.Redirect("showerr.asp?action=error&message=" & server.URLEncode(errmsg) & "")
		Response.End
	End If
	If RsAdmin("isAloneLogin") <> 0 And Trim(RsAdmin("RandomCode")) <> AdminRandomCode then
		Session.Abandon
		Response.Cookies(Admin_Cookies_Name) = ""
		ErrMsg = "<li><font color='red'>对不起，为了系统安全，本系统不允许两个人使用同一个管理员帐号进行登录！</font></li><li>因为现在有人已经在其他地方使用此管理员帐号进行登录了，所以你将不能继续进行后台管理操作。</li><li>你可以<a href='admin_klogin.asp' target='_top' class=showmeun>点此重新登录</a>。</li>"
		Response.Redirect("showerr.asp?action=error&message=" & server.URLEncode(errmsg) & "")
		RsAdmin.Close:set RsAdmin = Nothing
		Response.End
	End If
	'IP绑定操作
	if RsAdmin("useip")<>"" then
		if 	enchiasp.GetUserip<>RsAdmin("useip") then
			Session.Abandon
			Response.Cookies(Admin_Cookies_Name) = ""
			ErrMsg = "<li><font color='red'>对不起，为了系统安全，本系统不允许您登录！</font></li><li>IP已经绑定！</li><li>你可以<a href='admin_klogin.asp' target='_top' class=showmeun>点此重新登录</a>。</li>"
			Response.Redirect("showerr.asp?action=error&message=" & server.URLEncode(errmsg) & "")
			RsAdmin.Close:set RsAdmin = Nothing
			Response.End
		end if
	end if

	
End If
RsAdmin.Close:Set RsAdmin = Nothing
Dim ChannelID,sChannelName,sChannelDir,sModuleName,rsChannel,ChannelModuleID
If IsNumeric(Request("ChannelID")) Then
	ChannelID = CLng(Request("ChannelID"))
	If ChannelID <> 9999 Then
		Set rsChannel = enchiasp.Execute("Select ChannelID From ECCMS_Channel where ChannelType < 2 And ChannelID = " & ChannelID)
		If Not (rsChannel.BOF And rsChannel.EOF) Then
			enchiasp.ReadChannel(ChannelID)
			sChannelName = enchiasp.ChannelName
			sChannelDir = Replace(enchiasp.ChannelDir, "/", "")
			sModuleName = enchiasp.ModuleName
			ChannelModuleID = CInt(enchiasp.modules)
		End If
		rsChannel.Close:Set rsChannel = Nothing
	End If
Else
	ChannelID = 0
End If
Public Function DeleteHtmlFile(classid,id,HtmlFileDate)
	If CInt(enchiasp.IsCreateHtml)=0 Then Exit Function
	On Error Resume Next
	Dim rsClass,sHtmlFileName,sHtmlFilePath
	SQL = "SELECT HtmlFileDir FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID=" & CLng(classid)
	Set rsClass = enchiasp.Execute(SQL)
	If Not(rsClass.BOF And rsClass.EOF) Then
		sHtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & rsClass("HtmlFileDir") & enchiasp.ShowDatePath(HtmlFileDate,enchiasp.HtmlPath)
		sHtmlFileName = enchiasp.ReadFileName(HtmlFileDate,id,enchiasp.HtmlExtName,enchiasp.HtmlPrefix,enchiasp.HtmlForm,0)
		enchiasp.FileDelete(sHtmlFilePath & sHtmlFileName)
	End If
	rsClass.Close:Set rsClass = Nothing
End Function

Public Function ChkAdmin(para)
	On Error Resume Next
	Dim i, TempAdmin, Adminflag
	ChkAdmin = False
	AdminFlag = Replace(Session("Adminflag"), "'", "''")
	If para = "" Then Exit Function
	If CInt(Session("AdminGrade")) = 999 Then
		ChkAdmin = True
		Exit Function
	Else
		If Adminflag = "" Then
			ChkAdmin = False
			Exit Function
		Else
			tempAdmin = Split(Adminflag, ",")
			For i = 0 To UBound(tempAdmin)
				If Trim(LCase(tempAdmin(i))) = Trim(LCase(para)) Then
					ChkAdmin = True
					Exit For
				End If
			Next
		End If
	End If
End Function


Public Function Chkreg(para)
	On Error Resume Next
	Dim i, TempAdmin, Adminflag
	Chkreg = False
	AdminFlag = enchiasp.urlflag
	If para = "" Then Exit Function
	If Adminflag = "" Then
			Chkreg = False
			Exit Function
	Else
		tempAdmin = Split(Adminflag, ",")
		For i = 0 To UBound(tempAdmin)
			If Trim(LCase(tempAdmin(i))) = Trim(LCase(para)) Then
				Chkreg = True
				Exit For
			End If
		Next
		End If
End Function




Public Function ChkAdminPurview(flag,username)
	On Error Resume Next
	Dim i, TempAdmin, Adminflag, BlnAdminflag
	ChkAdminPurview = False
	BlnAdminflag = False
	If flag = "" Then Exit Function
	Adminflag = Replace(Session("Adminflag"), "'", "''")
	If CInt(Session("AdminGrade")) = 999 Then
		ChkAdminPurview = True
		Exit Function
	Else
		If Trim(Adminflag) = "" Then
			ChkAdminPurview = False
			Exit Function
		Else
			tempAdmin = Split(Adminflag, ",")
			For i = 0 To UBound(tempAdmin)
				If LCase(Trim(tempAdmin(i))) = LCase(Trim(flag)) Then
					BlnAdminflag = True
					Exit For
				End If
			Next
		End If
	End If
	If BlnAdminflag = True Then
		If Trim(username) = Trim(Session("AdminName")) Then
			ChkAdminPurview = True
			Exit Function
		Else
			ChkAdminPurview = False
			Exit Function
		End If
	Else
		ChkAdminPurview = False
		Exit Function
	End If
End Function

Public Sub AdminCookiesToSession()
	If Session("AdminName") = "" Then
		Session("AdminName") = enchiasp.CheckStr(Request.Cookies(Admin_Cookies_Name)("AdminName"))
		Session("AdminPass") = enchiasp.CheckStr(Request.Cookies(Admin_Cookies_Name)("AdminPass"))
		Session("AdminGrade") = enchiasp.CheckStr(Request.Cookies(Admin_Cookies_Name)("AdminLevel"))
		Session("Adminflag") = enchiasp.CheckStr(Request.Cookies(Admin_Cookies_Name)("Adminflag"))
		Session("AdminStatus") = enchiasp.CheckStr(Request.Cookies(Admin_Cookies_Name)("AdminStatus"))
		Session("AdminRandomCode") = enchiasp.CheckStr(Request.Cookies(Admin_Cookies_Name)("RandomCode"))
		Session("AdminID") = enchiasp.CheckStr(Request.Cookies(Admin_Cookies_Name)("AdminID"))
	End If
End Sub
%>