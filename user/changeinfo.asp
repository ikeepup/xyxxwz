<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/chkinput.asp"-->
<!--#include file="../api/cls_api.asp"-->
<!--#include file="head.inc"-->
<%
'=====================================================================
' ������ƣ�������վ����ϵͳ---�޸Ļ�Ա����
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж�������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================

Call InnerLocation("�޸Ļ�Ա����")

Dim Rs,SQL
If CInt(GroupSetting(1)) = 0 Then
	ErrMsg = ErrMsg + "<li>�Բ�����û���޸��û����ϵ�Ȩ�ޣ�����ʲô��������ϵ����Ա��</li>"
	Founderr = True
ElseIf LCase(Request("action")) = "save" Then
	Call ChangeUserInfo
Else
	Set Rs = enchiasp.Execute("SELECT * FROM [ECCMS_User] WHERE username='" & enchiasp.membername & "' And userid=" & enchiasp.memberid)
	If (Rs.bof And Rs.EOF) Then
		ErrMsg = ErrMsg + "<li>Sorry�������ϵͳ������</li>"
		Founderr = True
	Else
%>
<script language="JavaScript">
<!--
function checkForm() {
	if (document.myform.password.value.length == 0) {
		alert("�����������û�����!");
		document.myform.password.focus();
		return false;
	}
	if (document.myform.nickname.value.length == 0) {
		alert("�����������û��ǳ�!");
		document.myform.nickname.focus();
		return false;
	}
	if (document.myform.codestr.value.length != 4) {
		alert("��֤����������!");
		document.myform.codestr.focus();
		return false;
	}
	if (document.myform.usermail.value.length == 0) {
		alert("����������E-mail");
		document.myform.usermail.focus();
		return false;
	}
		return true;
}
//-->
</script>
<table cellspacing=1 align=center cellpadding=2 bgcolor=#cccccc border=0 class=Usertableborder>
	<tr>
		<th colspan=2>�޸ĸ�������</th>
	</tr>
	<form method="post" name=myform action="?action=save" onsubmit="return checkForm();">
	<tr>
		<td align=right width="25%" class=Usertablerow1 height=20>�û�����</td>
		<td width="75%" class=Usertablerow1> <strong class=userfont1><%=enchiasp.membername%></strong>
			<input type=hidden name=username value="<%=Server.HTMLEncode(Rs("username"))%>"><input type=hidden name=userid value="<%=enchiasp.memberid%>"></td>
	</tr>
	<tr>
		<td align=right class=Usertablerow2 height=20>�û��ǳ�(<span class=userfont1>*</span>)��</td>
		<td class=Usertablerow2> <input class=inputbody type=text size=20 name=nickname value="<%=enchiasp.HTMLEncodes(Rs("nickname"))%>" maxlength="15"></td>
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>��ʵ����(<span class=userfont1>*</span>)��</td>
		<td class=Usertablerow1> <input class=inputbody type=text size=20 name=TrueName value="<%=enchiasp.HTMLEncodes(Rs("TrueName"))%>" maxlength="15"></td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align=right class=Usertablerow2 height=20>&nbsp;�û�����(<span class=userfont1>*</span>)��</td>
		<td class=Usertablerow2> <input type=text class=inputbody size=30 name=usermail value="<%=enchiasp.HTMLEncodes(Rs("usermail"))%>" maxlength="50"> <span class=userfont1>ע�⣺</span><font color=#808080>����д�㳣�õ�����</font></td>
	
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>�Ա�</td>
		<td class=Usertablerow1> <input type=radio name=usersex value="��"<%If Trim(Rs("usersex")) = "��" Then Response.Write " checked"%>> ��&nbsp;&nbsp;    
			<input type=radio name=usersex value="Ů"<%If Trim(Rs("usersex")) = "Ů" Then Response.Write " checked"%>> Ů&nbsp;&nbsp;    
			<input type=radio name=usersex value="Ů"<%If Trim(Rs("usersex")) = "����" Then Response.Write " checked"%>> ����</td>    
	</tr>
	<tr>
		<td align=right class=Usertablerow2 height=20>������ʾ����(<span class=userfont1>*</span>)��</td>
		<td class=Usertablerow2> <input class=inputbody type=text size=30 name=question value="<%=enchiasp.HTMLEncodes(Rs("question"))%>" maxlength="35"> <select onChange="question.value=this.value;")>    
			<option value="" selected>[��ѡ��]</option>
			<option value="��ϲ���ĳ��">��ϲ���ĳ��</option>
			<option value="��ϲ���ĵ�Ӱ��">��ϲ���ĵ�Ӱ��</option>
			<option value="��������� [��/��/��]��">��������� 
            [��/��/��]��</option>
			<option value="���׵����֣�">���׵����֣�</option>
			<option value="��ż�����֣�">��ż�����֣�</option>
			<option value="��һ�����ӵİ��ƣ�">��һ�����ӵİ��ƣ�</option>
			<option value="��ѧ��У����">��ѧ��У����</option>
			<option value="���𾴵���ʦ��">���𾴵���ʦ��</option>
			<option value="��ϲ�������жӣ�">��ϲ�������жӣ�</option>
			</select></td>
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>��������𰸣�</td>
		<td class=Usertablerow1> <input class=inputbody type=text size=30 name=answer maxlength="35"> <font color=#808080>�����������ʾ����𰸣�����ȡ������</font></td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align=right class=Usertablerow2 height=20>��ϵ�绰��</td>
		<td class=Usertablerow2> <input type=text class=inputbody size=20 name=phone value="<%=enchiasp.HTMLEncodes(Rs("phone"))%>" maxlength="20"> <font color=#808080>�磺+86-27-85188888</font></td>    
	
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>���OICQ��</td>
		<td class=Usertablerow1> <input class=inputbody type=text size=20 name=oicq value="<%=enchiasp.HTMLEncodes(Rs("oicq"))%>" maxlength="20"></td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align=right class=Usertablerow2 height=20>�������룺</td>
		<td class=Usertablerow2> <input type=text class=inputbody size=20 name=postcode value="<%=enchiasp.HTMLEncodes(Rs("postcode"))%>" maxlength="20"></td>
	
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>���֤��</td>
		<td class=Usertablerow1> <input class=inputbody type=text size=30 name=UserIDCard value="<%=enchiasp.HTMLEncodes(Rs("UserIDCard"))%>" maxlength="35"></td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align=right class=Usertablerow2 height=20>��ϵ��ַ��</td>
		<td class=Usertablerow2> <input type=text class=inputbody size=50 name=address value="<%=enchiasp.HTMLEncodes(Rs("address"))%>" maxlength="50"></td>
	
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>�������룺</td>
		<td class=Usertablerow1> <input class=inputbody type=text size=30 name=BuyCode maxlength="35"> <font color=#808080>վ��֧�����õĽ�������</font></td>
	</tr>
	<tr>
		<td align=right class=Usertablerow2 height=20>������ҳ��</td>
		<td class=Usertablerow2> <input class=inputbody type=text size=30 name=HomePage value="<%=enchiasp.HTMLEncodes(Rs("HomePage"))%>" maxlength="35"> <font color=#808080>�ԡ�http://����ͷ</font></td>    
	</tr>
	<tr>
		<td align=right class=Usertablerow1 height=20>�û����룺</td>
		<td class=Usertablerow1> <input class=inputbody type=password size=30 name=password value="" maxlength="50"> <span class=userfont1>������ȷ����������޸��û�����</span></td>
	</tr>
	<tr>
		<td align=right class=Usertablerow2 height=20>�� ֤ �룺</td>
		<td class=Usertablerow2> <input class=inputbody type=text size=6 name=codestr maxlength="6">&nbsp;<img src="../inc/getcode.asp" alt="��֤��,�������?����ˢ����֤��" onclick="this.src='../inc/getcode.asp'"> <font color=#808080>��������֤��</font></td>    
	</tr>
	<tr>
		<td align=middle class=Usertablerow2 height=20>&nbsp; </td>
		<td class=Usertablerow2 align=center><input type=submit value=" ȷ �� " name="submit" class="button"></td>
	</tr></form>
</table>
<%
	End If
	Rs.Close:Set Rs = Nothing
End If
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
Sub ChangeUserInfo()
	On Error Resume Next
	Dim username, password,userid
	Dim usersex,sex
	username = enchiasp.CheckBadstr(enchiasp.membername)
	userid = enchiasp.ChkNumeric(memberid)
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If enchiasp.IsValidPassword(Request.Form("answer")) = False And Trim(Request.Form("answer")) <> "" Then
		ErrMsg = ErrMsg + "<li>����������к��зǷ��ַ���</li>"
		Founderr = True
	End If
	If Trim(Request.Form("username")) <> username Then
		ErrMsg = ErrMsg + "<li>�Ƿ�������</li>"
		Founderr = True
	End If
	If Trim(Request.Form("password")) = "" Then
		ErrMsg = ErrMsg + "<li>�������û����룡</li>"
		Founderr = True
	Else
		password = md5(Request.Form("password"))
	End If
	If userid = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	
	If Trim(Request.Form("nickname")) = "" Then
		ErrMsg = ErrMsg + "<li>�û��ǳƲ���Ϊ�գ�</li>"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request.Form("nickname")) = False Then
		ErrMsg = ErrMsg + "<li>�û��ǳ��к��зǷ��ַ���</li>"
		Founderr = True
	End If
	If Trim(Request.Form("TrueName")) = "" Then
		ErrMsg = ErrMsg + "<li>��ʵ��������Ϊ�գ�</li>"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request.Form("TrueName")) = False Then
		ErrMsg = ErrMsg + "<li>��ʵ�����к��зǷ��ַ���</li>"
		Founderr = True
	End If
	If Trim(Request.Form("usermail")) = "" Then
		ErrMsg = ErrMsg + "<li>����Email����Ϊ�գ�</li>"
		Founderr = True
	End If
	If IsValidEmail(Request.Form("usermail")) = False Then
		ErrMsg = ErrMsg + "<li>����Email�д���</li>"
		Founderr = True
	End If
	If Not IsNumeric(Request.Form("oicq")) And Trim(Request.Form("oicq")) <> "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>QQ��������������д��</li>"
	End If
	If Trim(Request.Form("HomePage")) <> "" And Left(Request.Form("HomePage"),7) <> "http://" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������ҳ��ַ�����������ԡ�http://����ͷ��</li>"
	End If
	If Not enchiasp.CodeIsTrue() Then
		ErrMsg = ErrMsg + "<meta http-equiv=""refresh"" content=""2;URL=changeinfo.asp""><li>��֤��У��ʧ�ܣ��뷵��ˢ��ҳ�����ԡ�������Զ�����</li>"
		Session("GetCode") = ""
		Founderr = True
		Exit Sub
	End If
	Session("GetCode") = ""
	If Trim(Request.Form("usersex")) = "" Then
		ErrMsg = ErrMsg + "<li>�����ձ���Ϊ�գ�</li>"
		Founderr = True
	Else
		usersex = enchiasp.CheckBadstr(Request.Form("usersex"))
	End If
	If usersex = "Ů" Then
		sex = 0
	Else
		sex = 1
	End If
	
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	SQL = "SELECT * FROM [ECCMS_user] WHERE username='" & username & "' And userid=" & CLng(userid)
	Rs.Open SQL, Conn, 1, 3
	If Rs.bof And Rs.EOF Then
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ����û���Ϣ��Ϣ��</li>"
		Founderr = True
		Exit Sub
	Else
		If password <> Rs("password") Then
			ErrMsg = ErrMsg + "<li>��������������</li>"
			Founderr = True
			Exit Sub
		End If
		If Founderr = False Then
			'-----------------------------------------------------------------
			'ϵͳ����
			'-----------------------------------------------------------------
			Dim API_enchiasp,API_SaveCookie,SysKey
			If API_Enable Then
				Set API_enchiasp = New API_Conformity
				API_enchiasp.NodeValue "action","update",0,False
				API_enchiasp.NodeValue "username",UserName,1,False
				Md5OLD = 1
				SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
				Md5OLD = 0
				API_enchiasp.NodeValue "syskey",SysKey,0,False
				API_enchiasp.NodeValue "password","",1,False
				API_enchiasp.NodeValue "answer",Request.Form("answer"),1,False
				API_enchiasp.NodeValue "question",Request.Form("question"),1,False
				API_enchiasp.NodeValue "email",Request.Form("usermail"),1,False
				API_enchiasp.NodeValue "gender",sex,0,False
				API_enchiasp.SendHttpData
				If API_enchiasp.Status = "1" Then
					Founderr = True
					ErrMsg = API_enchiasp.Message
					Exit Sub
				End If
				Set API_enchiasp = Nothing
			End If
			'-----------------------------------------------------------------
		End If
		Rs("nickname") = enchiasp.CheckBadstr(Request.Form("nickname"))
		Rs("TrueName") = enchiasp.CheckBadstr(Request.Form("TrueName"))
		Rs("usermail") = Trim(Request.Form("usermail"))
		If Trim(Request.Form("HomePage")) <> "" Then Rs("HomePage") = enchiasp.CheckBadstr(Trim(Request.Form("HomePage")))
		If Trim(Request.Form("usersex")) <> "" Then Rs("usersex") = usersex
		If Trim(Request.Form("question")) <> "" Then Rs("question") =enchiasp.CheckBadstr(Trim(Request.Form("question"))) 
		If Trim(Request.Form("answer")) <> "" Then Rs("answer") = md5(Trim(Request.Form("answer")))
		If Trim(Request.Form("phone")) <> "" Then Rs("phone") = enchiasp.CheckBadstr(Trim(Request.Form("phone")))
		If Trim(Request.Form("oicq")) <> "" Then Rs("oicq") = enchiasp.CheckBadstr(Trim(Request.Form("oicq")))
		If Trim(Request.Form("postcode")) <> "" Then Rs("postcode") = enchiasp.CheckBadstr(Trim(Request.Form("postcode")))
		If Trim(Request.Form("UserIDCard")) <> "" Then Rs("UserIDCard") = enchiasp.CheckBadstr(Trim(Request.Form("UserIDCard")))
		If Trim(Request.Form("address")) <> "" Then Rs("address") = enchiasp.CheckBadstr(Trim(Request.Form("address")))

		If Trim(Request.Form("BuyCode")) <> "" Then Rs("BuyCode") = md5(Trim(Request.Form("BuyCode")))
		Rs.Update
	End If
	Rs.Close
	Set Rs = Nothing
	Call Returnsuc("<li>��ϲ�����û������޸ĳɹ���</li>")
End Sub
%>
<!--#include file="foot.inc"-->