<!--#include file="setup.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
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

Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
Response.Write "	<tr>"
Response.Write "	  <th>" & sModuleName & "����ѡ��</th>"
Response.Write "	</tr>"
Response.Write "	<tr><form method=Post name=myform action='' onSubmit='return JugeQuery(this);'>"
Response.Write "	<td class=TableRow1>������"
Response.Write "	  <input name=keyword type=text size=30>"
Response.Write "	  ������"
Response.Write "	  <select name='field'>"
Response.Write "		<option value='1' selected>������</option>"
Response.Write "		<option value='2'>�� �� ��</option>"
Response.Write "		<option value='3'>�� �� ��</option>"
Response.Write "	  </select> <input type=submit name=Submit value='��ʼ��ѯ' class=Button><br>"
Response.Write "	  <b>˵����</b>��������Ų鿴�ʹ�����</td></form>"
Response.Write "	</tr></form>"
Response.Write "	<tr>"
Response.Write "	  <td colspan=2 class=TableRow2><strong>����ѡ�</strong> <a href='admin_orderform.asp'>������ҳ</a> | "
Response.Write "	  <a href='admin_orderform.asp?finish=1'>�Ѵ�����</a> | "
Response.Write "	  <a href='admin_orderform.asp?finish=0'>δ������</a> | "
Response.Write "	  <a href='admin_orderform.asp?Cancel=1'>����վ����</a></td>"
Response.Write "	</tr>"
Response.Write "</table>"
Response.Write "<br>"
If Not ChkAdmin("adminorder") Then
	Server.Transfer("showerr.asp")
	Response.End
End If

Dim Action,i
If CInt(ChannelID) = 0 Then ChannelID = 3
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "save"
	Call SaveOrderForm
Case "view"
	Call ViewOrderForm
Case "del"
	Call DelOrderForm
Case "cancel"
	Call ReclaimOrder
Case "finish"
	Call FinishOrderForm
Case "pay"
	Call PaymentState
Case Else
	Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Sub showmain()
	Dim finish,Cancel
	Dim keyword,findword,foundsql
	Dim maxperpage,CurrentPage,Pcount,totalrec,totalnumber
	Dim strList,strName,strRowstyle

	maxperpage = 30		'--ÿҳ��ʾ�б���
	
	finish = enchiasp.ChkNumeric(Request("finish"))
	Cancel = enchiasp.ChkNumeric(Request("Cancel"))
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>ѡ��</th>"
	Response.Write "	  <th width='23%'>�� �� ��</th>"
	Response.Write "	  <th width='12%' nowrap>�� �� ��</th>"
	Response.Write "	  <th width='15%' nowrap>�� �� �� ��</th>"
	Response.Write "	  <th width='19%' nowrap>�� �� ʱ ��</th>"
	Response.Write "	  <th width='10%' nowrap>�� �� �� ʽ</th>"
	Response.Write "	  <th width='8%' nowrap>����״̬</th>"
	Response.Write "	  <th width='8%' nowrap>��������</th>"
	Response.Write "	</tr>"
	
	If Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("field")) = 1 Then
			foundsql = " And OrderID like '%" & keyword & "%'"
		ElseIf CInt(Request("field")) = 2 Then
			foundsql = " And Consignee like '%" & keyword & "%'"
		ElseIf CInt(Request("field")) = 3 Then
			foundsql = " And username like '%" & keyword & "%'"
		Else
			foundsql = " And OrderID like '%" & keyword & "%'"
		End If
		strName = "������ѯ"
		strList = "&keyword=" & keyword
	Else
		If Request("finish") <> "" Then
			foundsql = " And finish=" & finish
			strList = "&finish=" & finish
			If finish = 0 Then
				strName = "δ������"
			Else
				strName = "�Ѵ�����"
			End If
		Else
			If Cancel = 0 Then
				strName = "���ж���"
			Else
				strName = "�Ѿ�ɾ������"
			End If
		End If
	End If
	strList = strList & "&Cancel=" & Cancel
	CurrentPage = enchiasp.ChkNumeric(Request("page"))
	If CurrentPage = 0 Then CurrentPage = 1
	totalrec = enchiasp.Execute("SELECT COUNT(id) FROM [ECCMS_OrderForm] WHERE Cancel="& Cancel & foundsql &"")(0)
	Pcount = CLng(totalrec / maxperpage)  '�õ���ҳ��
	If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT id,userid,username,ProductID,OrderID,Surcharge,totalmoney,Consignee,Email,PayMode,addTime,invoice,finish,Cancel,PayDone FROM [ECCMS_OrderForm] WHERE Cancel="& Cancel & foundsql &" ORDER BY id DESC"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=""center"" colspan=""9"" class=""TableRow2"">��û���ҵ��κζ�����</td></tr>"
	Else
		If Pcount > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0

		Response.Write "	<tr>"
		Response.Write "	  <td colspan=""8"" class=""TableRow2"">"
		ShowListPage CurrentPage,Pcount,totalrec,maxperpage,strList,strName
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<form name=selform method=post action=""admin_orderform.asp"">"
		Response.Write "	<input type=hidden name=action value='del'>"
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			If (i mod 2) = 0 Then
				strRowstyle = "class=""TableRow1"""
			Else
				strRowstyle = "class=""TableRow2"""
			End If
			Response.Write "	<tr align=""center"">"
			Response.Write "	  <td align=center " & strRowstyle & "><input type=checkbox name=id value=" & Rs("id") & "></td>"
			Response.Write "	  <td " & strRowstyle & " title=""����˴��鿴������ϸ��Ϣ""><a href='?action=view&id=" & Rs("id") & "' class=""showlink"">"
			Response.Write Rs("OrderID")
			Response.Write "</a></td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & " title=""����û����鿴���û���Ϣ"">"
			If Rs("userid") > 0 Then
				Response.Write "<a href='admin_user.asp?action=edit&userid=" & Rs("userid") & "'>"
				Response.Write Rs("Consignee")
				Response.Write "</a>"
			Else
				Response.Write Rs("Consignee")
			End If
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & " align=""left"">��"
			Response.Write FormatNumber(Rs("totalmoney"))
			Response.Write " Ԫ</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & " align=""left"">"
			If Datediff("d",Rs("addTime"),Now()) = 0 Then
				Response.Write "<font color=""red"">" & Rs("addTime") & "</font>"
			Else
				Response.Write "<font color=""#808080"">" & Rs("addTime") & "</font>"
			End If
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & ">"
			Response.Write Rs("PayMode")
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & ">"
			If Rs("PayDone") > 0 Then
				Response.Write "<a href='?action=pay&sid=0&id=" & Rs("id") & "' title=""����˴��ı�֧��״̬"">"
				Response.Write "<font color=""blue"">��֧��</font>"
				Response.Write "</a>"
			Else
				Response.Write "<a href='?action=pay&sid=1&id=" & Rs("id") & "' title=""����˴��ı�֧��״̬"">"
				Response.Write "<font color=""red"">δ֧��</font>"
				Response.Write "</a>"
			End If
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & ">"
			If Rs("finish") > 0 Then
				'Response.Write "<a href='?action=finish&fid=0&id=" & Rs("id") & "' title=""����˴�ֱ�Ӵ�����"">"
				Response.Write "<font color=""blue"">�Ѵ���</font>"
				'Response.Write "</a>"
			Else
				'Response.Write "<a href='?action=finish&fid=1&id=" & Rs("id") & "' title=""����˴�ȡ������"">"
				Response.Write "<font color=""red"">δ����</font>"
				'Response.Write "</a>"
			End If
			Response.Write "</td>" & vbNewLine
			Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
	<tr>
	  <td colspan="8" class="TableRow1">
	  <input class=Button type="button" name="chkall" value="ȫѡ" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="��ѡ" onClick="ContraSel(this.form)">
	  <input class=Button type="submit" name="Submit2" value="����ɾ��" onclick="return confirm('����ɾ���󽫲��ָܻ�\n��ȷ��ִ�иò�����?');">
	  <%
	  If Cancel = 0 Then
	  %>
	  <input type=hidden name=can value='1'>
	  <input class=Button type="submit" name="Submit3" value="�������վ" onclick="document.selform.action.value='cancel';return confirm('��ȷ��Ҫ����Щ�����������վ��?');">
	  <%
	  Else
	  %>
	  <input type=hidden name=can value='0'>
	  <input class=Button type="submit" name="Submit4" value="��ԭ����վ" onclick="document.selform.action.value='cancel';return confirm('��ȷ����ԭ������?');">
	  <%
	  End If
	  %>
	  </td>
	</tr>
</form>
	<tr>
	  <td colspan="8" align="right" class="TableRow2"><%ShowListPage CurrentPage,Pcount,totalrec,maxperpage,strList,strName %></td>
	</tr>
</table>
<%
End Sub
Sub ReclaimOrder()
	If Request("id") <> "" And Request("can") <> "" Then
		If CInt(Request("can")) = 0 Then
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET Cancel=0 WHERE id in (" & Request("id") & ")")
			OutHintScript("��ѡ��Ķ����ѳɹ���ԭ��")
		Else
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET Cancel=1 WHERE id in (" & Request("id") & ")")
			OutHintScript("��ѡ��Ķ����ѳɹ��������վ��")
		End If
	Else
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
End Sub
Sub DelOrderForm()
	If Request("id") <> "" Then
		enchiasp.Execute ("DELETE FROM [ECCMS_OrderForm] WHERE id in (" & Request("id") & ")")
		enchiasp.Execute ("DELETE FROM [ECCMS_Buy] WHERE orderid in (" & Request("id") & ")")
	Else
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	OutHintScript("��ѡ��Ķ����ѳɹ�ɾ����")
End Sub
Sub FinishOrderForm()
	If Request("id") <> "" And Request("fid") <> "" Then
		If Request("fid") = 0 Then
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET finish=0 WHERE id=" & CLng(Request("id")))
		Else
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET finish=1 WHERE id=" & CLng(Request("id")))
		End If
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
End Sub

Sub PaymentState()
	If Request("id") <> "" And Request("sid") <> "" Then
		If Request("sid") = 0 Then
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET PayDone=0 WHERE id=" & CLng(Request("id")))
		Else
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET PayDone=1 WHERE id=" & CLng(Request("id")))
		End If
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
End Sub

Sub ViewOrderForm()
	Dim id,totalmoney
	id = enchiasp.ChkNumeric(Request("id"))
	If id = 0 Then
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT * FROM [ECCMS_OrderForm] WHERE id=" & id)
	If Rs.BOF And Rs.EOF Then
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Set Rs = Nothing
		Exit Sub
	End If
	totalmoney = FormatNumber(Rs("totalmoney"))
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="tableborder">
<tr>
	<th colspan="4">�����鿴/����</th>
</tr>
<form name="subform" method="post" action="admin_orderform.asp">
<input type="hidden" name="action" value="save">
<input type="hidden" name="id" value="<%=Rs("id")%>">
<tr>
	<td width='15%' class="tablerow1" align="right">�� �� �ţ�</td>
	<td width='42%' class="tablerow1"><font color=red><%=Rs("OrderID")%></font></td>
	<td width='15%' class="tablerow1" align="right">�� �� ����</td>
	<td width='28%' class="tablerow1"><%
	If Rs("userid") > 0 Then
		Response.Write "<a href='admin_user.asp?action=edit&userid=" & Rs("userid") & "'>"
		Response.Write Rs("username")
		Response.Write "</a>"
	Else
		Response.Write "�����û�"
	End If
	%></td>
</tr>
<tr>
	<td class="tablerow2" align="right">�ϼƽ�</td>
	<td class="tablerow2"><font color=blue>��<%=FormatNumber(Rs("totalmoney"))%> Ԫ</font></td>
	<td class="tablerow2" align="right">���ӷ��ã�</td>
	<td class="tablerow2">��<%=FormatNumber(Rs("Surcharge"),,-1)%> Ԫ</td>
</tr>
<tr>
	<td class="tablerow1" align="right">����ʱ�䣺</td>
	<td class="tablerow1"><font color=red><%=Rs("addTime")%></font></td>
	<td class="tablerow1" align="right">���ʽ��</td>
	<td class="tablerow1"><%=Rs("PayMode")%></td>
</tr>
<tr>
	<td class="tablerow2" align="right">�� �� �ˣ�</td>
	<td class="tablerow2"><font color=blue><%=Rs("Consignee")%></font></td>
	<td class="tablerow2" align="right">�ջ���λ��</td>
	<td class="tablerow2"><%=Rs("Company")%></td>
</tr>
<tr>
	<td class="tablerow1" align="right">�ջ��˵绰��</td>
	<td class="tablerow1"><%=Rs("phone")%></td>
	<td class="tablerow1" align="right">�ջ����ʱࣺ</td>
	<td class="tablerow1"><%=Rs("postcode")%></td>
</tr>
<tr>
	<td class="tablerow2" align="right">�ջ������䣺</td>
	<td class="tablerow2"><%=Rs("Email")%></td>
	<td class="tablerow2" align="right">�ջ���QQ�ţ�</td>
	<td class="tablerow2"><%=enchiasp.ChkNull(Rs("oicq"))%></td>
</tr>
<tr>
	<td class="tablerow1" align="right">�ջ��˵�ַ��</td>
	<td class="tablerow1"><%=Rs("Address")%></td>
	<td class="tablerow2" align="right">�Ƿ񿪷�Ʊ��</td>
	<td class="tablerow2"><%
	If Rs("invoice") > 0 Then
		Response.Write "<font color=""red"">��</font>"
	Else
		Response.Write "<font color=""#808080"">��</font>"
	End If
%></td>
</tr>
<tr>
	<td class="tablerow2" align="right">����˵����</td>
	<td class="tablerow2" colspan="3">&nbsp;&nbsp;<%=enchiasp.HTMLEncode(Rs("Readme"))%></td>
</tr>
<tr>
	<td class="tablerow1" align="right">��������</td>
	<td class="tablerow1"><font color=red><%
	If Rs("finish") > 0 Then
		Response.Write "<font color=""blue"">�Ѵ���</font>"
	Else
		Response.Write "<font color=""red"">δ����</font>"
	End If
%></font></td>
	<td class="tablerow1" align="right">����״̬��</td>
	<td class="tablerow1"><%
	If Rs("PayDone") > 0 Then
		Response.Write "��֧�� <input type=radio name=PayDone value=""1"" checked>"
	Else
		Response.Write "δ֧�� <input type=radio name=PayDone value=""0"" checked>&nbsp;&nbsp;"
		Response.Write "��֧�� <input type=radio name=PayDone value=""1"">"
	End If
	Rs.Close:Set Rs = Nothing
%></td>
</tr>
<tr>
	<th>��������</td>
	<th>��Ʒ����</td>
	<th>�� ��</td>
	<th>�� ��</td>
</tr>
<%
	SQL = "SELECT * FROM ECCMS_Buy WHERE orderid=" & id & " ORDER BY ID ASC"
	Set Rs = enchiasp.Execute(SQL)
	If Not (Rs.BOF And Rs.EOF) Then
	Do While Not Rs.EOF
%>
<tr>
	<td class="tablerow1" align="center"><%=Rs("Amount")%></td>
	<td class="tablerow1"><a href="admin_shop.asp?action=view&shopid=<%=Rs("shopid")%>"><%=Rs("TradeName")%></a></td>
	<td class="tablerow1" align="center"><font color=blue>��<%=FormatNumber(Rs("Price"))%> Ԫ</font></td>
	<td class="tablerow1"  align="center"><font color=red>��<%=FormatNumber(Rs("totalmoney"))%> Ԫ</font></td>
</tr>
<%
		Rs.movenext
	Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr align="center">
	<td class="tablerow2" colspan="4"><input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp; 
		<input type=hidden name=can value='1'>
		<input class=Button type="submit" name="Submit3" value="ȡ������" onclick="document.subform.action.value='cancel';return confirm('��ȷ��Ҫ����Щ�����������վ��?');">&nbsp;&nbsp;
		<input class=Button type="submit" name="Submit4" value="ɾ������" onclick="document.subform.action.value='del';return confirm('����ɾ���󽫲��ָܻ�\n��ȷ��ִ�иò�����?');">&nbsp;&nbsp;
		<input class=Button type="submit" name="Submit2" value="����˶���" onclick="return confirm('ȷ������˶�����?');">
	</td>
</tr></form>
<tr>
	<td class="tablerow2" colspan="4"><b>˵����</b><br>
	&nbsp;&nbsp;��ȷ���û��Ѿ����ٴ�����������һ��������֤�����ѷ�����
	</td>
</tr>
</table>
<%
End Sub

Sub SaveOrderForm()
	Dim id,totalmoney,Consignee,Readme,PayDone
	id = enchiasp.ChkNumeric(Request("id"))
	If id = 0 Then
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_OrderForm WHERE finish=0 And id=" & id
	Rs.Open SQL,Conn,1,3
	If Rs.BOF And Rs.EOF Then
		ErrMsg = "<li>�˶����Ѿ������벻Ҫ�ظ���������</li>"
		Founderr = True
		Set Rs = Nothing
		Exit Sub
	Else
		PayDone = Rs("PayDone")
		Rs("finish") = 1
		Rs("PayDone") = enchiasp.ChkNumeric(Request.Form("PayDone"))
		Rs.update
		totalmoney = Rs("totalmoney")
		Consignee = Rs("Consignee")
		Readme = enchiasp.ChkNull(Rs("Readme"))
	End If
	Rs.Close
	'-- ��ʼ��ӽ�����ϸ��
	If PayDone = 0 Then
		SQL = "SELECT * FROM ECCMS_Account WHERE (AccountID is null)"
		Rs.Open SQL,Conn,1,3
		Rs.addnew
			Rs("payer").Value = Consignee
			Rs("payee").Value = enchiasp.CheckRequest(enchiasp.SiteName,20)
			Rs("product").Value = "���Ϲ���"
			Rs("Amount").Value = 1
			Rs("unit").Value = "��"
			Rs("price").Value = totalmoney
			Rs("TotalPrices").Value = totalmoney
			Rs("DateAndTime").Value = Now()
			Rs("Accountype").Value = 0
			Rs("Explain").Value = Readme
			Rs("Reclaim").Value = 0
		Rs.update
		Rs.Close:Set Rs = Nothing
	End If
	Succeed("<li>��ϲ������������ɹ�����Ͽ���û�����ȥ��!</li>")
End Sub
%>