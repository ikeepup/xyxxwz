<!--#include file="setup.asp"-->

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
Response.CacheControl = "no-cache"
Dim RefreshTime,GetCode
FoundErr = False
RefreshTime = 3 '���÷�ˢ��ʱ��
If DateDiff("s", Session("UserTime"), Now()) < RefreshTime Then
	Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT="&RefreshTime&"><br>��ҳ�������˷�ˢ�»��ƣ��벻Ҫ��"&RefreshTime&"��������ˢ�±�ҳ��<BR>���ڴ�ҳ�棬���Ժ󡭡�"
	Response.End
End If
FoundErr = False
Select Case enchiasp.CheckStr(Request("action"))
	Case "logout" '�˳�ϵͳ
		Call logout()
	Case "login" '��½ϵͳ
		Call chklogin()
	Case Else
		if enchiasp.ercilogin ="1" then
		'ת�������½ҳ��
			session("mypasskey")=enchiasp.mypasskey
	 		response.redirect "admin_loginx.asp" 
		 else
			 Call main()
		end if
		
End Select

If Founderr = True Then
	Session("UserTime") = Now()
	SaveLogInfo("�Ƿ���½��")
	Response.Redirect("showerr.asp?action=error&message=" & Server.URLEncode(ErrMsg) & "")
End If
CloseConn

Sub main()
	
	If Session("AdminName") = "" Then
%>
<html>
<head>
<title>����Ա��½</title>
<meta http-equiv="Content-Type" content="text/html; chaRset=gb2312">
<link rel="stylesheet" href="images/admin.css" type="text/css">
</head>
<script language="javascript">
<!--//
function SetFocus()
{
if (document.myform.AdminName.value=="")
	document.myform.AdminName.focus();
else
	document.myform.AdminName.select();
}
function CheckForm()
{
	if(document.myform.AdminName.value=="")
	{
		alert("�����������û�����");
		document.myform.AdminName.focus();
		return false;
	}
	if(document.myform.PassWord.value == "")
	{
		alert("�������������룡");
		document.myform.PassWord.focus();
		return false;
	}
	if (document.myform.verifycode.value==""){
       alert ("������������֤�룡");
       document.myform.verifycode.focus();
       return(false);
    }
}
function CheckBrowser()
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("ϵͳ������ʾ��\n    ��ʹ�õ���Netscape����������ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  }
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("ϵͳ������ʾ��\n    ����������汾̫�ͣ����ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  }
}
//-->
</script>
<body topmargin="0" leftmargin="0" rightmargin="0">
<script language="JavaScript" src="keyboard.js" type="text/javascript"></script>
<div align="center"><BR>
  <p>��</p>  <p>��</p>
  <form name=myform method="post" action="admin_klogin.asp?action=login" target="_top" onSubmit="return CheckForm();">

<table border="0" cellpadding="0" cellspacing="0" width="642" background="images/adminlogin.jpg" height="425" align="center">
  <tr>
    <td width="100%" height="370">
      <table border="0" cellpadding="0" cellspacing="0" width="100%" height="401">
        <tr>
          <td width="50%" height="401"></td>
          <td width="50%" height="401">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
              <tr>
                <td width="100%"></td>
              </tr>
              <tr>
                <td width="100%"></td>
              </tr>
              <tr>
                <td width="100%"></td>
              </tr>
              <tr>
                <td width="100%">��
                  <p>��</p>
                  <p>��</p>
                  <p>��</p>
                  <p>��</p>
                  <p>��</td>
              </tr>
              <tr>
               <TD height=25 align=center><form name="form1" method="post" action="admin_klogin.asp?action=login" onsubmit="return login()">
�û�����<input type="text" name="AdminName"  style=width:150px autocomplete="off"  class="ycenchi"><input onclick="showkeyboard('Login.AdminName')" type="button" value="����" title="��������������룬��ֹ�ڿ������¼���̵�¼����Ϣ" />
</td></tr>
<tr><td height=25 align=center>
��&nbsp; �룺<input type="password" name="Password"  style=width:150px  class="ycenchi"><input onclick="showkeyboard('Login.AdminName')" type="button" value="����" title="��������������룬��ֹ�ڿ������¼���̵�¼����Ϣ" />     
</td></tr>
<tr><td height=25 align=center>
��֤�룺<input name="verifycode" type="text" size="12" maxlength="9" class="ycenchi">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="../inc/getcode.asp"  id="GetCodePic" align=absmiddle height=16 border=0></td></tr><tr><td height=25 align=center>
<input type="submit" name="Submit" value="�� ¼" class="adminbutton">&nbsp;<input type="reset" name="reset" value="�� ��" class="adminbutton">
</TD>
              </tr>
              <tr>
                <td width="100%"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>


</form>
<script language="JavaScript">
<!--
CheckBrowser();
SetFocus();
-->
</script>
<p align=center></p>
</div>
<%
Else
	Response.Redirect "admin_kindex.asp"
End If
End Sub

Sub logout()
	'���COOKIES�й���Ա��ݵ���֤��Ϣ.
	Session.Abandon
	Session("AdminName") = ""
	Session("AdminPass") = ""
	Session("AdminGrade") = ""
	Session("AdminFlag") = ""
	Session("AdminStatus") = ""
	Session("AdminID") = ""
	Session("AdminRandomCode") = ""
	Response.Cookies(Admin_Cookies_Name) = ""
	Response.Redirect ("../")
End Sub

Sub chklogin()
	Dim adminname, password,RandomCode,mypass
	dim tempmima,ss,tt
	adminname = Trim(Replace(Request("adminname"), "'", ""))
	password = md5(Trim(Replace(Request("password"), "'", "")))
	mypass=Trim(Replace(Request("mypassword"), "'", ""))
	If enchiasp.CheckPost = False Then
		ErrMsg = ErrMsg + "���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ��½��"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request("adminname")) = False Then
		ErrMsg = ErrMsg + "<li>�û����к��зǷ��ַ���</li>"
		Founderr = True
	End If
	If enchiasp.IsValidPassword(Request("password")) = False Then
		ErrMsg = ErrMsg + "<li>�����к��зǷ��ַ���</li>"
		Founderr = True
	End If
	If Request("verifycode") = "" Then
		ErrMsg = ErrMsg + "<br>" + "<li>�뷵������ȷ���롣</li>"
		Founderr = True
	ElseIf Session("getcode") = "9999" Then
		Session("getcode") = ""
		ErrMsg = ErrMsg + "<br>" + "<li>�벻Ҫ�ظ��ύ���������µ�½�뷵�ص�½ҳ�档</li>"
		Founderr = True
	ElseIf CStr(Session("getcode"))<>CStr(Trim(Request("verifycode"))) Then
		ErrMsg = ErrMsg + "<br>" + "<li>���������֤���ϵͳ�����Ĳ�һ�£����������롣</li>"
		Founderr = True
	End If
	Session("getcode") = ""
	If adminname = "" Or password = "" Then
		Founderr = True
		ErrMsg = ErrMsg + "<br>" + "<li>�����������û��������롣</li>"
		Exit Sub
	End If
	
	
	

	
	If Founderr = True Then Exit Sub
	If Not IsObject(Conn) Then ConnectionDatabase
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Admin where password='" & password & "' And username='" & adminname & "'"
	Rs.Open SQL, Conn, 1, 3
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������û��������벻��ȷ����������ϵͳ����Ա����</li>"
		Exit Sub
	Else
		If password <> Rs("password") Then
			FoundErr = True
			ErrMsg = ErrMsg + "<br><li>�û�����������󣡣���</li>"
			Exit Sub
		End If
		If Rs("isLock") <> 0 Or Rs("isLock") = "" Then
			Founderr = True
			ErrMsg = "<li>����û����ѱ�����,�㲻�ܵ�½����Ҫ��ͨ���ʺţ�����ϵ����Ա��</li>"
			Exit Sub
		End If
		'����Ƿ����������뿪��
		if enchiasp.ercilogin="1" then
			if mypass="" then
				ErrMsg = ErrMsg + "<br>" + "<li>�벻Ҫ�Ƿ����Ե�½����ȷ�����Ƿ���ϵͳ����Ա��ϵͳ�Ѿ���¼����Ĳ�����¼��</li>"
				Founderr = True
				exit sub
			else	
				if rs("isuseercima")=1 then
				
					'�ж��Ǽӷ������ǳ˷�����
					if rs("jiafa")=1 then
						'�ӷ�
						tempmima=cstr(cint(mid(Request("verifycode"),rs("weizhi1"),1))+cint(mid(Request("verifycode"),rs("weizhi2"),1)))
						ss=""
						ss=mid(enchiasp.mypass,1,cint(rs("jimaweizhi")))
						tempmima=ss+tempmima+mid(enchiasp.mypass,cint(rs("jimaweizhi"))+1)
						if mypass<>tempmima then
							ErrMsg = ErrMsg + "<br>" +  "<li>�벻Ҫ�Ƿ����Ե�½����ȷ�����Ƿ���ϵͳ����Ա��ϵͳ�Ѿ���¼����Ĳ�����¼��</li>"
							Founderr = True
							exit sub
						end if
						
					elseif rs("jiafa")=0 then
						'�˷�
						tempmima=cstr(cint(mid(Request("verifycode"),rs("weizhi1"),1))*cint(mid(Request("verifycode"),rs("weizhi2"),1)))
						ss=""
						ss=mid(enchiasp.mypass,1,cint(rs("jimaweizhi")))
						tempmima=ss+tempmima+mid(enchiasp.mypass,cint(rs("jimaweizhi"))+1)
						if mypass<>tempmima then
							ErrMsg = ErrMsg + "<br>" + "<li>�벻Ҫ�Ƿ����Ե�½����ȷ�����Ƿ���ϵͳ����Ա��ϵͳ�Ѿ���¼����Ĳ�����¼��</li>"
							Founderr = True
							exit sub
						end if

					end if
					
				else
					'û�п����������
					if mypass<>enchiasp.mypass then
						ErrMsg = ErrMsg + "<br>" + "<li>�벻Ҫ�Ƿ����Ե�½����ȷ�����Ƿ���ϵͳ����Ա��ϵͳ�Ѿ���¼����Ĳ�����¼��</li>"
						Founderr = True
						exit sub
					end if
				end if
			end if
		end if
	
		
		
		
		
	End If
	RandomCode = enchiasp.GetRandomCode
	Rs("LoginTime") = Now()
	Rs("Loginip") = enchiasp.GetUserip
	Rs("RandomCode") = RandomCode
	Rs.Update
	If FoundErr = False Then
		Session("AdminName") = Rs("username")
		Session("AdminPass") = Rs("password")
		Session("AdminGrade") = Rs("AdminGrade")
		Session("Adminflag") = Rs("Adminflag")
		Session("AdminStatus") = Rs("Status")
		Session("AdminRandomCode") = RandomCode
		Session("AdminID") = Rs("id")
		Response.Cookies(Admin_Cookies_Name)("AdminName") = Rs("username")
		Response.Cookies(Admin_Cookies_Name)("AdminPass") = Rs("password")
		Response.Cookies(Admin_Cookies_Name)("AdminGrade") = Rs("AdminGrade")
		Response.Cookies(Admin_Cookies_Name)("Adminflag") = Rs("Adminflag")
		Response.Cookies(Admin_Cookies_Name)("AdminStatus") = Rs("Status")
		Response.Cookies(Admin_Cookies_Name)("RandomCode") = RandomCode
		Response.Cookies(Admin_Cookies_Name)("AdminID") = Rs("id")
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Redirect("admin_kindex.asp")
End Sub

Function GetCode1()
	Dim Test
	On Error Resume Next
	Set Test = Server.CreateObject("Adodb.Stream")
	Set Test = Nothing
	If Err Then
		Dim zNum
		Randomize Timer
		zNum = CInt(8999 * Rnd + 1000)
		Session("GetCode") = zNum
		getcode1 = Session("GetCode")
	Else
		getcode1 = "<img src=""../inc/getcode.asp"">"
	End If
End Function
%>
</body>
</html>
