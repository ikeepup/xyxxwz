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
Dim GroupSetting,Cookies_Name
Dim rsmember,sqlmember,MemberName,MemberEmail,memberid
MemberName = enchiasp.CheckBadstr(enchiasp.memberName)
memberid = enchiasp.ChkNumeric(enchiasp.memberid)
If Trim(MemberName) = "" Or memberid = 0 Then
	Response.Redirect ("./login.asp")
End If
MemberName = Left(MemberName,45)
If Trim(Request.Cookies(enchiasp.Cookies_Name)) = "" Then
	Response.Redirect ("./login.asp")
End If
sqlmember = "SELECT userid,UserLock,usermail FROM ECCMS_User WHERE username='" & MemberName & "' And UserGrade="& CInt(enchiasp.membergrade) &" And userid=" & CLng(memberid)
Set rsmember = enchiasp.Execute(sqlmember)
If rsmember.BOF And rsmember.EOF Then
	Response.Cookies(enchiasp.Cookies_Name) = ""
	Set rsmember = Nothing
	Response.Redirect "login.asp"
	Response.End
Else
	If rsmember("UserLock") > 0 Then
		Response.Cookies(enchiasp.Cookies_Name) = ""
		Set rsmember = Nothing
		ErrMsg = "<li>����û����ѱ�����,�㲻�ܵ�½����Ҫ��ͨ���ʺţ�����ϵ����Ա��</li>"
		Call Returnerr(ErrMsg)
		Response.End
	End If
	MemberEmail = Trim(rsmember("usermail"))
End If
Set rsmember = Nothing

GroupSetting = Split(enchiasp.UserGroupSetting(CInt(enchiasp.membergrade)), "|||")
Call GetUserTodayInfo
Cookies_Name = "usercookies_" & enchiasp.memberid

If Trim(Request.Cookies(Cookies_Name)) = "" Then
	Response.Cookies(Cookies_Name)("userip") = enchiasp.GetUserIP
	Response.Cookies(Cookies_Name)("dayarticlenum") = 0
	Response.Cookies(Cookies_Name)("daysoftnum") = 0
	Response.Cookies(Cookies_Name).Expires = Date + 1
End If

If CInt(enchiasp.memberclass) > 0 Then
	Dim rsUserClass,SQLUserClass
	Set rsUserClass = Server.CreateObject("ADODB.Recordset")
	SQLUserClass = "SELECT userid,UserClass,UserLock,ExpireTime FROM ECCMS_User WHERE username='" & MemberName & "' And userid=" & CLng(enchiasp.memberid)
	rsUserClass.Open SQLUserClass,Conn,1,3
	If rsUserClass.BOF And rsUserClass.EOF Then
		Response.Cookies(enchiasp.Cookies_Name) = ""
		rsUserClass.Close:Set rsUserClass = Nothing
		Response.Redirect "login.asp"
	Else
		If rsUserClass("UserLock") > 0 Then
			Response.Cookies(enchiasp.Cookies_Name) = ""
			rsUserClass.Close:Set rsUserClass = Nothing
			Response.Redirect "login.asp"
		End If
		If DateDiff("D", CDate(rsUserClass("ExpireTime")), Now()) > 0 And rsUserClass("UserClass") <> 999 Then
			rsUserClass("UserClass") = 999
			rsUserClass.Update
		End If
	End If
	rsUserClass.Close:Set rsUserClass = Nothing
End If
%>