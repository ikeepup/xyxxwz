<!--#include file="../conn.asp"-->
<!-- #include file="../inc/const.asp" -->
<!--#include file="../api/cls_api.asp"-->
<%
'=====================================================================
' ������ƣ�������վ����ϵͳ---���������û�����
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж�������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim username,password
Dim repassword,answer
Dim Rs, SQL
If Trim(Request("username")) = "" Or Trim(Request("pass")) = "" Or Trim(Request("repass")) = "" Or Trim(Request("answer")) = "" Then
	Response.Write "<script>alert('�Ƿ�������');history.go(-1)</script>"
Else
	username = enchiasp.CheckBadstr(Request("username"))
	password = enchiasp.CheckBadstr(Request("pass"))
	repassword = MD5(enchiasp.Checkstr(Request("repass")))
	answer = MD5(Request("answer"))
	SQL = "SELECT password,UserGrade FROM [ECCMS_User] WHERE username='" & username & "' And password='" & password & "' And answer='" & answer & "'"
	Set Rs = Server.CreateObject("adodb.recordset")
	Rs.open SQL, Conn, 1, 3
	If Rs.EOF And Rs.bof Then
		Response.Write "<script>alert('�����ص��û����ϲ���ȷ���Ƿ�������');history.go(-1)</script>"
	Else
		If Rs("UserGrade") = 999 Then
			Response.Write "<script>alert('�Ƿ�����������͹���Ա��ϵȡ�����룡');history.go(-1)</script>"
		Else
			Rs("password") = repassword
			Rs.Update
			'-----------------------------------------------------------------
			'ϵͳ����
			'-----------------------------------------------------------------
			Dim API_enchiasp,API_SaveCookie,SysKey
			If API_Enable Then
				Set API_enchiasp = New API_Conformity
				API_enchiasp.NodeValue "action","reguser",0,False
				API_enchiasp.NodeValue "username",UserName,1,False
				Md5OLD = 1
				SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
				Md5OLD = 0
				API_enchiasp.NodeValue "syskey",SysKey,0,False
				API_enchiasp.NodeValue "password",Request("repass"),0,False
				API_enchiasp.SendHttpData
				Set API_enchiasp = Nothing
			End If
			'-----------------------------------------------------------------
			Response.Write "<script>alert('��������ɹ�������½!');location.replace('./')</script>"
		End If
	End If
	Rs.Close
	Set Rs = Nothing
End If
CloseConn
%>