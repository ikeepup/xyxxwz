<%
'=====================================================================
' ������ƣ�������վ����ϵͳ
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж�������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim AdminName, AdminPass, AdminID, ErrorStr
Dim SQLAdmin, RsAdmin, AdminRandomCode
ErrorStr = "<li>ȷ�����ʧ�ܣ���û��ʹ�õ�ǰ���ܵ�Ȩ�ޡ�</li><li>�����ʲô���⣬����ϵ����Ա��</li>"
If InStr(enchiasp.ScriptName, "editor") > 0 Or InStr(enchiasp.ScriptName, "admin_label") > 0 Or InStr(enchiasp.ScriptName, "admin_collect") > 0 Then AdminPage = True
'If enchiasp.CheckPost = False And AdminPage = False  Then
	'ErrMsg = "<br><li><font color=red>���ύ�����ݲ��Ϸ���Ϊ��ϵͳ��ȫ��������ֱ�������ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></li><li>��Ϊ��ִ���˷Ƿ�������<a href=logout.asp target=_top class=showmeun>�����˳���ϵͳ��</a></li>"
	'Response.Redirect("showerr.asp?action=error&message=" & server.URLEncode(errmsg) & "")
	'Response.End
'End If
Call AdminCookiesToSession
AdminName = enchiasp.CheckStr(Session("AdminName"))      '����Ա����
AdminPass = enchiasp.CheckStr(Session("AdminPass"))      '����Ա����
AdminID = enchiasp.ChkNumeric(Session("AdminID"))                    '����ԱID
AdminRandomCode = Trim(Session("AdminRandomCode"))     '����Ա��½�����
If AdminName = "" Then
	ErrMsg = ErrMsg + "<li>��û�н��뱾ҳ���Ȩ��!���β����ѱ���¼!<li>��������û�е�½���߲�����ʹ�õ�ǰ���ܵ�Ȩ��!����ϵ����Ա.<li>��ҳ��Ϊ[<font color=red>����Ա</font>]ר��,����<a href=admin_klogin.asp class=showmeun target=_top>��½</a>����롣"
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
		ErrMsg = "<li>����û����ѱ�����,�㲻�ܵ�½����Ҫ��ͨ���ʺţ�����ϵ����Ա��</li>"
		RsAdmin.Close:set RsAdmin = Nothing
		Response.Redirect("showerr.asp?action=error&message=" & server.URLEncode(errmsg) & "")
		Response.End
	End If
	If RsAdmin("isAloneLogin") <> 0 And Trim(RsAdmin("RandomCode")) <> AdminRandomCode then
		Session.Abandon
		Response.Cookies(Admin_Cookies_Name) = ""
		ErrMsg = "<li><font color='red'>�Բ���Ϊ��ϵͳ��ȫ����ϵͳ������������ʹ��ͬһ������Ա�ʺŽ��е�¼��</font></li><li>��Ϊ���������Ѿ��������ط�ʹ�ô˹���Ա�ʺŽ��е�¼�ˣ������㽫���ܼ������к�̨���������</li><li>�����<a href='admin_klogin.asp' target='_top' class=showmeun>������µ�¼</a>��</li>"
		Response.Redirect("showerr.asp?action=error&message=" & server.URLEncode(errmsg) & "")
		RsAdmin.Close:set RsAdmin = Nothing
		Response.End
	End If
	'IP�󶨲���
	if RsAdmin("useip")<>"" then
		if 	enchiasp.GetUserip<>RsAdmin("useip") then
			Session.Abandon
			Response.Cookies(Admin_Cookies_Name) = ""
			ErrMsg = "<li><font color='red'>�Բ���Ϊ��ϵͳ��ȫ����ϵͳ����������¼��</font></li><li>IP�Ѿ��󶨣�</li><li>�����<a href='admin_klogin.asp' target='_top' class=showmeun>������µ�¼</a>��</li>"
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