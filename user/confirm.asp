<!--#include file="config.asp"-->
<!--#include file="../inc/chkinput.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
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
Call InnerLocation("����ȷ��")

Dim Action,SQL,Rs
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "save"
	Call SaveConfirm
Case Else
	Call showmain
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
Sub showmain()
%>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr height=20>
		<th colspan=2>����ȷ��</th>
	</tr>
	<tr height=20 align=center>
		<td class=Usertablerow2 colspan=2><font color=red>ע�⣺</font><font color=blue>��һ��Ҫ��ȷ��д���º�*��ѡ��Է������Ǻ˶ԣ�</font></td>
	</tr>
	<form name=form2 method=post action=?action=save>
	<tr height=20>
		<td class=Usertablerow1 width="20%" align=right>������ڣ�</td>
		<td class=Usertablerow1 width="80%"><input type="text" name="PayDate" size=15 value="<%=date()%>"> <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>����</td>
		<td class=Usertablerow1><input type="text" name="PayMoney" size=15 onkeyup=if(isNaN(this.value))this.value=''> Ԫ <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>�� �� �ţ�</td>
		<td class=Usertablerow1><input type="text" name="indent" size=30> <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>��ʽ��</td>
		<td class=Usertablerow1>
		<input type=radio name=paymode value="���л��" checked> ���&nbsp;&nbsp;
		<input type=radio name=paymode value="�ʾֻ��"> �ʻ�&nbsp;&nbsp;
		<input type=radio name=paymode value="����֧��"> ����֧��
		</td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>�û�����</td>
		<td class=Usertablerow1><input type="text" name="username" size=15 value="<%=enchiasp.MemberName%>"> <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>��������ƣ�</td>
		<td class=Usertablerow1><input type="text" name="customer" size=30> <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>��������䣺</td>
		<td class=Usertablerow1><input type="text" name="Email" size=30> <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>����˵����</td>
		<td class=Usertablerow1><textarea name=readme rows=5 cols=50></textarea> <font color=red>*</font></td>
	</tr>
	<tr height=20 align=center>
		<td class=Usertablerow2 colspan=2><input type=submit value=" ȷ���ύ "  class=Button></td>
	</tr>
	</form>
<%
	Response.Write "</table>"
End Sub
Sub SaveConfirm()
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Not IsDate(Request.Form("PayDate")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����������</li>"
	End If
	If Not IsNumeric(Request.Form("PayMoney")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������������</li>"
	End If
	If Trim(Request.Form("indent")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��Ķ�����û�����֣�</li>"
	End If
	If IsValidEmail(Request.Form("Email")) = False Then
		ErrMsg = ErrMsg + "<li>����Email�д���</li>"
		Founderr = True
	End If
	If Trim(Request.Form("customer")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������Ʋ���Ϊ�ա�</li>"
	End If
	If Trim(Request.Form("username")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�û�������Ϊ�գ�</li>"
	End If
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Confirm where (id is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("paymode").Value =enchiasp.CheckBadstr( Trim(Request.Form("paymode")))
		Rs("PayDate").Value = Trim(Request.Form("PayDate"))
		Rs("PayMoney").Value = Trim(Request.Form("PayMoney"))
		Rs("indent").Value = Left(enchiasp.ChkFormStr(Request.Form("indent")),35)
		Rs("Email").Value = Trim(Request.Form("Email"))
		Rs("customer").Value = Left(enchiasp.ChkFormStr(Request.Form("customer")),30)
		Rs("username").Value = Left(enchiasp.ChkFormStr(Request.Form("username")),30)
		Rs("readme").Value = Left(enchiasp.ChkFormStr(Request.Form("readme")),200)
		Rs("isPass").Value = 0
	Rs.Update
	Rs.close:set Rs = Nothing
	Call Returnsuc("<li>��ϲ����ȷ����Ϣ�ύ�ɹ������ǻ���һ���������ڴ�����Ķ�����")
End Sub

%>
<!--#include file="foot.inc"-->











