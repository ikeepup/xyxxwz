<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
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
Dim Action
If Not ChkAdmin("userorder") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Call showpagetop
Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "add"
		Call AddUserMoney
	Case "view"
		Call ViewOrder
	Case "del"
		Call DelOrder
	Case "delfinish"
		Call DelFinishOrder
	Case Else
		Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub showpagetop()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th>��Ա��ֵ��������</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr><form method=Post name=myform action=admin_userorder.asp onSubmit='return JugeQuery(this);'>"
	Response.Write "	  <td class=TableRow1>������ѯ��"
	Response.Write "	  <input name=keyword type=text size=30>"
	Response.Write "	  ������"
	Response.Write "	  <select name=field>"
	Response.Write "		<option value=1 selected>�� �� ��</option>"
	Response.Write "		<option value=2>�� �� ��</option>"
	Response.Write "		<option value=3>��������</option>"
	Response.Write "	  </select> <input type=submit name=Submit value='��ʼ��ѯ' class=Button></td>"
	Response.Write "	  </form>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow2><strong>����ѡ�</strong> <a href='admin_userorder.asp'>������ҳ</a> | "
	Response.Write "	  <a href='admin_userorder.asp?finished=0'>δ������</a> | "
	Response.Write "	  <a href='admin_userorder.asp?finished=1'>�Ѵ�����</a> | "
	Response.Write "	  <a href='admin_userorder.asp?action=delfinish' onClick=""return confirm('ȷ��Ҫ��������Ѵ�������')"">��������Ѵ�����</a></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub
Sub showmain()
	Dim CurrentPage,page_count,totalnumber,Pcount,maxperpage
	Dim keyword,findword,tablebody
	maxperpage = 30
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CLng(CurrentPage)
	End If
	If CLng(CurrentPage) = 0 Then CurrentPage = 1
	If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
		keyword = Replace(Replace(Replace(Replace(Replace(Request("keyword"), "'", "''"), "[", ""), "]", ""), "%", ""), "|", "")
		If CInt(Request("field")) = 1 Then
			findword = "WHERE OrderForm like '%" & keyword & "%'"
		ElseIf CInt(Request("field")) = 2 Then
			findword = "WHERE username like '%" & keyword & "%'"
		Else
			findword = "WHERE OrderForm like '%" & keyword & "%' Or username like '%" & keyword & "%'"
		End If
	Else
		If Trim(Request("finished")) <> "" Then
			If Request("finished") > 0 Then
				findword = "WHERE finished>0"
			Else
				findword = "WHERE finished=0"
			End If
		Else
			findword = ""
		End If
	End If
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th>�� �� ��</th>"
	Response.Write "		<th>�� �� ��</th>"
	Response.Write "		<th>֧�����</th>"
	Response.Write "		<th>֧������</th>"
	Response.Write "		<th>�ύ����</th>"
	Response.Write "		<th>���ʽ</th>"
	Response.Write "		<th>״ ̬</th>"
	Response.Write "		<th>�� ��</th>"
	Response.Write "	</tr>"
	totalnumber = enchiasp.Execute("SELECT COUNT(id) FROM ECCMS_AddMoney " & findword & "")(0)
	Pcount = CLng(totalnumber / maxperpage)  '�õ���ҳ��
	If Pcount < totalnumber / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_AddMoney " & findword & " ORDER BY id DESC"
	If IsSqlDataBase=1 And Trim(Request("keyword"))="" Then
		Set Rs = enchiasp.Execute(SQL)
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=8 class=TableRow1>û�л�Ա��ֵ������</td></tr>"
	Else
		Rs.MoveFirst
		If Pcount > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		page_count = 0
		Do While Not Rs.EOF And page_count < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			If (page_count mod 2) = 0 Then
				tablebody = "class=TableRow1"
			Else
				tablebody = "class=TableRow2"
			End If
			Response.Write "	<tr align=center>"
			Response.Write "		<td " & tablebody & "><a href=""?action=view&id="& Rs("id") &""" title='�鿴������Ϣ'><font color=red>" & Rs("OrderForm") & "</font></a></td>"
			Response.Write "		<td " & tablebody & "><a href=""admin_user.asp?action=edit&userid="& Rs("userid") &""" title='�鿴��Ա��Ϣ'>" & Rs("username") & "</a></td>"
			Response.Write "		<td " & tablebody & ">" & FormatCurrency(Rs("addmoney"),2,-1) & " Ԫ</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("title") & "</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("addtime") & "</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("paytype") & "</td>"
			Response.Write "		<td " & tablebody & ">"
			If Rs("finished") > 0 Then
				Response.Write "<font color=blue>�Ѵ���</font>"
			Else
				Response.Write "<a href=""?action=add&id="& Rs("id") &""" title='�������˶���' onClick=""return confirm('ȷ��Ҫ����˶�����')""><font color=red>δ����</font></a>"
			End If
			Response.Write "</td>"
			Response.Write "		<td " & tablebody & ">"
			Response.Write "<a href=""?action=del&id=" & Rs("id") & """ onClick=""return confirm('ȷ��Ҫɾ���˶�����')"">ɾ ��</a>"
			Response.Write "</td>"
			Response.Write "	</tr>"
			Rs.movenext
			page_count = page_count + 1
			If page_count >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow2 colspan=8>"
	Response.Write ShowPages(CurrentPage,Pcount,totalnumber,maxperpage,"&finished="& Request("finished") &"&keyword="& Request("keyword"))
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
End Sub
Sub DelFinishOrder()
	enchiasp.Execute("DELETE FROM ECCMS_AddMoney WHERE finished>0")
	Response.Redirect  Request.ServerVariables("HTTP_REFERER")
End Sub
Sub DelOrder()
	If Not IsNumeric(Request("id")) Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>����ȷ����ID��</li>"
		Exit Sub
	End If
	enchiasp.Execute("DELETE FROM ECCMS_AddMoney WHERE id="& CLng(Request("id")))
	Response.Redirect  Request.ServerVariables("HTTP_REFERER")
End Sub
Sub AddUserMoney()
	If Not IsNumeric(Request("id")) Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>����ȷ����ID��</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_AddMoney WHERE finished=0 And id="& CLng(Request("id")))
	If Rs.BOF And Rs.EOF Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ���������ߴ˶����Ѿ�����</li>"
		Set Rs = Nothing
		Exit Sub
	Else
		enchiasp.Execute ("UPDATE ECCMS_User SET usermoney=usermoney+"& CCur(Rs("addmoney")) &" WHERE username='"& Rs("username") &"' And userid="& CLng(Rs("userid")))
		enchiasp.Execute ("UPDATE ECCMS_AddMoney SET finished=1 WHERE id="& CLng(Request("id")))
		Dim sqlAccount,rsAccount
		Set rsAccount = Server.CreateObject("ADODB.Recordset")
		sqlAccount = "SELECT * FROM ECCMS_Account WHERE (AccountID is null)"
		rsAccount.Open sqlAccount,Conn,1,3
		rsAccount.addnew
			rsAccount("payer").Value = Rs("username").Value
			rsAccount("payee").Value = enchiasp.SiteName
			rsAccount("product").Value = Rs("title").Value
			rsAccount("Amount").Value = 1
			rsAccount("unit").Value = "��"
			rsAccount("price").Value = Rs("addmoney").Value
			rsAccount("TotalPrices").Value = Rs("addmoney").Value
			rsAccount("DateAndTime").Value = Now()
			rsAccount("Accountype").Value = 0
			rsAccount("Explain").Value = Rs("readme").Value
			rsAccount("Reclaim").Value = 0
		rsAccount.update
		rsAccount.Close:Set rsAccount = Nothing
		Succeed("<li>����������ɡ�</li><li>���ѳɹ�Ϊ�û���<b>" & Rs("username") & "</b> ��ֵ���" & FormatCurrency(Rs("addmoney"),2,-1) & " Ԫ</li>")
	End If
	Set Rs = Nothing
End Sub
Public Sub saveaccount()
	Dim sqlAccount,rsAccount
	Set rsAccount = Server.CreateObject("ADODB.Recordset")
	sqlAccount = "SELECT * FROM ECCMS_Account WHERE (AccountID is null)"
	rsAccount.Open sqlAccount,Conn,1,3
	rsAccount.addnew
		rsAccount("payer").Value = Request.Form("payer")
		rsAccount("payee").Value = Request.Form("payee")
		rsAccount("product").Value = Request.Form("product")
		rsAccount("Amount").Value = Request.Form("Amount")
		rsAccount("unit").Value = Request.Form("unit")
		rsAccount("price").Value = Request.Form("price")
		rsAccount("TotalPrices").Value = Request.Form("TotalPrices")
		rsAccount("DateAndTime").Value = Now()
		rsAccount("Accountype").Value = 0
		rsAccount("Explain").Value = Request.Form("Explain")
		rsAccount("Reclaim").Value = 0
	rsAccount.update
	rsAccount.Close:Set rsAccount = Nothing
End Sub

Sub ViewOrder()
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_AddMoney WHERE id="& CLng(Request("id")))
	If Rs.BOF And Rs.EOF Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������</li>"
		Set Rs = Nothing
		Exit Sub
	Else
		Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
		Response.Write "	<tr>"
		Response.Write "		<th colspan=2>�鿴������Ϣ</th>"
		Response.Write "	</tr>"
		Response.Write "	<form name=addform method=post action=?action=add>"
		Response.Write "	<input type=hidden name=id value="""& Rs("id") &""">"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1 align=right width=""25%""><b>��Ա���ƣ�</b></td>"
		Response.Write "		<td class=tablerow1 width=""75%""><a href=""admin_user.asp?action=edit&userid="& Rs("userid") &""" title='�鿴��Ա��Ϣ'><font color=blue>" & Rs("username") & "</font></a></td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=right><b>��ֵ�����ţ�</b></td>"
		Response.Write "		<td class=tablerow2><font color=red>" & Rs("OrderForm") & "</font></td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1 align=right><b>��ֵ�Ľ�</b></td>"
		Response.Write "		<td class=tablerow1><font color=red>" & FormatCurrency(Rs("addmoney"),2,-1) & "</font> Ԫ</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=right><b>��ֵ���ͣ�</b></td>"
		Response.Write "		<td class=tablerow2>" & Rs("title") & "</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1 align=right><b>֧����ʽ��</b></td>"
		Response.Write "		<td class=tablerow1>" & Rs("paytype") & "</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=right><b>�ύ���ڣ�</b></td>"
		Response.Write "		<td class=tablerow2>" & Rs("addtime") & "</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1 align=right><b>����״̬��</b></td>"
		Response.Write "		<td class=tablerow1>"
		If Rs("finished") > 0 Then
			Response.Write "<font color=blue>�Ѵ���</font>"
		Else
			Response.Write "<font color=red>δ����</font>"
		End If
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=right><b>����˵����</b></td>"
		Response.Write "		<td class=tablerow2>" & Server.HTMLEncode(Rs("readme")) & "</td>"
		Response.Write "	</tr>"
		
		Response.Write "	<tr align=center>"
		Response.Write "		<td class=tablerow1 colspan=2><input type=submit value="" ������ "" class=Button>&nbsp;&nbsp; "
		Response.Write "		<input type=""button"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" name=""B1"" class=Button></td>"
		Response.Write "	</tr>"
		Response.Write "	</form>"
		Response.Write "</table>"
	End If
	Set Rs = Nothing
End Sub

%>