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
If Not ChkAdmin("adminaccount") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Call showpagetop
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "add"
	Call AddAccount
Case "savenew"
	Call SavenewAccount
Case "save"
	Call SaveAccount
Case "view"
	Call ViewAccount
Case "del"
	Call DelAccount
Case "reclaim"
	Call ReclaimAccount
Case "renew"
	Call RenewAccount
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
	Response.Write "	  <th>������ϸ�����</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow1><strong>ѡ�</strong>"
	Response.Write "<a href=""admin_account.asp"">������ϸ</a> | <a href=""admin_account.asp?Accountype=0"">������ϸ</a> | <a href=""admin_account.asp?Accountype=1"">֧����ϸ</a> | <a href=""admin_account.asp?Reclaim=1"">����վ</a> | <a href=""admin_account.asp?action=add""><font color=blue>��ӽ�����ϸ��</font></a>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr align=right>"
	Response.Write "	  <td class=TableRow2>"
	Call AccountCount
	Response.Write "	  </td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub
Sub showmain()
	Dim CurrentPage,page_count,totalnumber,Pcount,maxperpage
	Dim findword,tablebody,BeginDate,LastDate,i,BeginDated,LastDated
	maxperpage = 30
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CLng(CurrentPage)
	End If
	If CLng(CurrentPage) = 0 Then CurrentPage = 1
	If Trim(Request("BeginDate")) <> "" And Trim(Request("LastDate")) <> "" Then
		BeginDate = CDate(Replace(Replace(Request("BeginDate"), ",", "-"), " ", ""))
		LastDate = CDate(Replace(Replace(Request("LastDate"), ",", "-"), " ", ""))
		If IsSqlDataBase=1 Then
			findword = "WHERE Reclaim=0 And Datediff(d,DateAndTime,getdate())<" & DateDiff("d", BeginDate-1, Now()) & " And Datediff(d,DateAndTime,getdate())>" & DateDiff("d", LastDate+1, Now()) & ""
		Else
			findword = "WHERE Reclaim=0 And Datediff('d',DateAndTime,Now())<" & DateDiff("d", BeginDate-1, Now()) & " And Datediff('d',DateAndTime,Now())>" & DateDiff("d", LastDate+1, Now()) & ""
		End If
	Else
		If Not IsNull(Request("Reclaim")) And Request("Reclaim") <> "" Then
			findword = "WHERE Reclaim>0"
		Else
			If Trim(Request("Accountype")) <> "" Then
				If Request("Accountype") > 0 Then
					findword = "WHERE Accountype>0 And Reclaim=0"
				Else
					findword = "WHERE Accountype=0 And Reclaim=0"
				End If
			Else
				findword = "WHERE Reclaim=0"
			End If
		End If
	End If
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th>ѡ��</th>"
	Response.Write "		<th>�� �� ��</th>"
	Response.Write "		<th>�տλ</th>"
	Response.Write "		<th>��Ŀ����</th>"
	Response.Write "		<th>��    ��</th>"
	Response.Write "		<th>��    λ</th>"
	Response.Write "		<th>��    ��</th>"
	Response.Write "		<th>�� �� ��</th>"
	Response.Write "		<th>��������</th>"
	Response.Write "		<th>��������</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=selform method=post action='admin_account.asp'>"
	Response.Write "	<input type=hidden name=action value=""reclaim"">"
	totalnumber = enchiasp.Execute("SELECT COUNT(AccountID) FROM ECCMS_Account " & findword & "")(0)
	Pcount = CLng(totalnumber / maxperpage)  '�õ���ҳ��
	If Pcount < totalnumber / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT AccountID,payer,payee,product,Amount,unit,price,TotalPrices,DateAndTime,Accountype,Explain,Reclaim FROM ECCMS_Account " & findword & " ORDER BY DateAndTime DESC, AccountID DESC"
	Rs.Open SQL, Conn, 1, 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=10 class=TableRow1>û����ϸ��</td></tr>"
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
			Response.Write "		<td " & tablebody & "><input type=checkbox name=AccountID value="""& Rs("AccountID") &"""></td>"
			Response.Write "		<td " & tablebody & ">" & Rs("payer") & "</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("payee") & "</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("product") & "</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("Amount") & "</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("unit") & "</td>"
			Response.Write "		<td " & tablebody & ">" & FormatCurrency(Rs("price"),2,-1) & " Ԫ</td>"
			Response.Write "		<td " & tablebody & ">" & FormatCurrency(Rs("TotalPrices"),2,-1) & " Ԫ</td>"
			Response.Write "		<td " & tablebody & ">" & FormatDateTime(Rs("DateAndTime"),2) & "</td>"
			Response.Write "		<td " & tablebody & ">"
			If Rs("Accountype") > 0 Then
				Response.Write "<font color=red>֧ ��</font>"
			Else
				Response.Write "<font color=blue>�� ��</font>"
			End If
			Response.Write " | <a href=""?action=view&AccountID="& Rs("AccountID") &""" title='�鿴��ϸ��Ϣ'>�� ��</a>"
			Response.Write "</td>"
			Response.Write "	</tr>"
			Rs.movenext
			page_count = page_count + 1
			If page_count >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=5>"
	Response.Write "<input class=Button type=""button"" name=""chkall"" value=""ȫѡ"" onClick=""CheckAll(this.form)""><input class=Button type=""button"" name=""chksel"" value=""��ѡ"" onClick=""ContraSel(this.form)"">"
	Response.Write "<input type=submit name=submit2 value=""�������վ"" onclick=""return confirm('ȷ���������վ��?')"" class=Button>"
	Response.Write "<input type=submit name=submit3 value=""��ԭ����վ"" onclick=""document.selform.action.value='renew';return confirm('ȷ����ԭ��?')"" class=Button>"
	Response.Write "<input type=submit name=submit4 value="" ����ɾ�� "" onclick=""document.selform.action.value='del';return confirm('ȷ��Ҫ����ɾ����?')"" class=Button>"
	Response.Write "</td>"
	Response.Write "		<td class=tablerow1 colspan=5>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow2 colspan=10>"
	Response.Write ShowPages(CurrentPage,Pcount,totalnumber,maxperpage,"&Accountype="& Request("Accountype") &"&Reclaim="& Request("Reclaim") &"&BeginDate="& BeginDate &"&LastDate="& LastDate)
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<form name=queryform method=get action=admin_account.asp>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=10><b>�����ڲ�ѯ��</b>"
	Response.Write " <select size=""1"" name=""BeginDate"">"
	For i = 2001 To Year(Date)
		Response.Write "<option value=""" & i & """"
		If i = Year(Date) Then
			Response.Write " selected"
		End If
		Response.Write ">" & i & "</option>"
	Next
	Response.Write " </select> - "
	Response.Write " <select size=""1"" name=""BeginDate"">"
	For i = 1 To 12
		Response.Write "<option value=""" & i & """"
		If i = Month(Date) Then
			Response.Write " selected"
		End If
		Response.Write ">" & i & "</option>"
	Next
	Response.Write " </select> - "
	Response.Write " <select size=""1"" name=""BeginDate"">"
	For i = 1 To 31
		Response.Write "<option value=""" & i & """"
		If i = Day(Date) Then
			Response.Write " selected"
		End If
		Response.Write ">" & i & "</option>"
	Next
	Response.Write " </select>���� "
	Response.Write " <select size=""1"" name=""LastDate"">"
	For i = 2001 To Year(Date)
		Response.Write "<option value=""" & i & """"
		If i = Year(Date) Then
			Response.Write " selected"
		End If
		Response.Write ">" & i & "</option>"
	Next
	Response.Write " </select> - "
	Response.Write " <select size=""1"" name=""LastDate"">"
	For i = 1 To 12
		Response.Write "<option value=""" & i & """"
		If i = Month(Date) Then
			Response.Write " selected"
		End If
		Response.Write ">" & i & "</option>"
	Next
	Response.Write " </select> - "
	Response.Write " <select size=""1"" name=""LastDate"">"
	For i = 1 To 31
		Response.Write "<option value=""" & i & """"
		If i = Day(Date) Then
			Response.Write " selected"
		End If
		Response.Write ">" & i & "</option>"
	Next
	Response.Write " </select>��"
	Response.Write "		<input type=submit name=submit3 value=""��ʼ��ѯ"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</from>"
	Response.Write "</table>"
End Sub
Sub ViewAccount()
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_Account WHERE AccountID="& CLng(Request("AccountID")))
	If Rs.BOF And Rs.EOF Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������</li>"
		Set Rs = Nothing
		Exit Sub
	Else
		Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
		Response.Write "	<tr>"
		Response.Write "		<th colspan=2>�鿴��ϸ��Ϣ</th>"
		Response.Write "	</tr>"
		Response.Write "	<form name=myform method=post action=?action=save>"
		Response.Write "	<input type=hidden name=AccountID value='"& Rs("AccountID") &"'>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1 align=right width=""25%""><b>�� �� �ˣ�</b></td>"
		Response.Write "		<td class=tablerow1 width=""75%""><input type=""text"" name=""payer"" size=50 value='" & Rs("payer") & "'></td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=right><b>�տλ��</b></td>"
		Response.Write "		<td class=tablerow2><input type=""text"" name=""payee"" size=50 value='" & Rs("payee") & "'></td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1 align=right><b>��Ŀ���ƣ�</b></td>"
		Response.Write "		<td class=tablerow1><input type=""text"" name=""product"" size=50 value='" & Rs("product") & "'></td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=right><b>�� ����</b></td>"
		Response.Write "		<td class=tablerow2><input type=""text"" name=""Amount"" size=5 value='" & Rs("Amount") & "'></td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1 align=right><b>�� λ��</b></td>"
		Response.Write "		<td class=tablerow1><input type=""text"" name=""unit"" size=5 value='" & Rs("unit") & "'></td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=right><b>�� �ۣ�</b></td>"
		Response.Write "		<td class=tablerow2><input type=""text"" name=""price"" size=10 value='" & Rs("price") & "'>  Ԫ</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1 align=right><b>�� �� �</b></td>"
		Response.Write "		<td class=tablerow1><input type=""text"" name=""TotalPrices"" size=10 value='" & Rs("TotalPrices") & "'> Ԫ</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=right><b>����ʱ�䣺</b></td>"
		Response.Write "		<td class=tablerow2><input type=""text"" name=""DateAndTime"" size=30 value='" & Rs("DateAndTime") & "'></td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1 align=right><b>�������ͣ�</b></td>"
		Response.Write "		<td class=tablerow1>"
		Response.Write "<select name=""Accountype"">"
		If Rs("Accountype") > 0 Then
			Response.Write "<option value=1>֧ ��</option>"
			Response.Write "<option value=0>�� ��</option>"
		Else
			Response.Write "<option value=0>�� ��</option>"
			Response.Write "<option value=1>֧ ��</option>"
		End If
		Response.Write "</select>"
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=right><b>����˵����</b></td>"
		Response.Write "		<td class=tablerow2><textarea name=Explain rows=5 cols=50>" & Server.HTMLEncode(Rs("Explain")) & "</textarea></td>"
		Response.Write "	</tr>"
		
		Response.Write "	<tr align=center>"
		Response.Write "		<td class=tablerow1 colspan=2><input type=""button"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" name=""B1"" class=Button>&nbsp;&nbsp;"
		Response.Write "		<input type=submit name=submit2 value=""�޸���ϸ��"" class=Button>"
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	</form>"
		Response.Write "</table>"
	End If
	Set Rs = Nothing
End Sub
Sub ReclaimAccount()
	Dim selAccountID
	If Not IsEmpty(Request("AccountID")) Then
		selAccountID = Request("AccountID")
		enchiasp.Execute ("UPDATE [ECCMS_Account] SET Reclaim=1 WHERE AccountID in (" & selAccountID & ")")
		Response.redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������ID����Ϊ�գ�</li>"
		Exit Sub
	End If
End Sub
Sub RenewAccount()
	Dim selAccountID
	If Not IsEmpty(Request("AccountID")) Then
		selAccountID = Request("AccountID")
		enchiasp.Execute ("UPDATE [ECCMS_Account] SET Reclaim=0 WHERE AccountID in (" & selAccountID & ")")
		Response.redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������ID����Ϊ�գ�</li>"
		Exit Sub
	End If
End Sub
Sub DelAccount()
	Dim selAccountID
	If Not IsEmpty(Request("AccountID")) Then
		selAccountID = Request("AccountID")
		enchiasp.Execute("DELETE FROM [ECCMS_Account] WHERE AccountID in (" & selAccountID & ")")
		Response.redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������ID����Ϊ�գ�</li>"
		Exit Sub
	End If
End Sub
Sub AccountCount()
	Dim Earning,Payout,Balance,Amount
	'---- ������
	Set Rs = enchiasp.Execute("SELECT SUM(TotalPrices) FROM ECCMS_Account WHERE Reclaim=0 And Accountype=0")
	Earning = Rs(0)
	If IsNull(Earning) Then Earning = 0
	Set Rs = Nothing
	'---- ֧�����
	Set Rs = enchiasp.Execute("SELECT SUM(TotalPrices) FROM ECCMS_Account WHERE Reclaim=0 And Accountype>0")
	Payout = Rs(0)
	If IsNull(Payout) Then Payout = 0
	Set Rs = Nothing
	'---- �����ܶ�
	Set Rs = enchiasp.Execute("SELECT SUM(TotalPrices) FROM ECCMS_Account WHERE Reclaim=0")
	Amount = Rs(0)
	If IsNull(Amount) Then Amount = 0
	Set Rs = Nothing
	'---- ������
	Balance = Earning - Payout
	Response.Write "�����ܶ<font color=red><b>"
	Response.Write FormatCurrency(Amount,2,-1)
	Response.Write "</b></font> Ԫ&nbsp;&nbsp;"
	Response.Write "���룺<font color=red><b>"
	Response.Write FormatCurrency(Earning,2,-1)
	Response.Write "</b></font> Ԫ&nbsp;&nbsp;"
	Response.Write "֧����<font color=red><b>"
	Response.Write FormatCurrency(Payout,2,-1)
	Response.Write "</b></font> Ԫ&nbsp;&nbsp;"
	Response.Write "��<font color=red><b>"
	Response.Write FormatCurrency(Balance,2,-1)
	Response.Write "</b></font> Ԫ&nbsp;&nbsp;"
End Sub
Sub AddAccount()
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=2>�����ϸ��</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=savenew>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right width=""25%""><b>�� �� �ˣ�</b></td>"
	Response.Write "		<td class=tablerow1 width=""75%""><input type=""text"" name=""payer"" size=50></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>�տλ��</b></td>"
	Response.Write "		<td class=tablerow2><input type=""text"" name=""payee"" size=50></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>��Ŀ���ƣ�</b></td>"
	Response.Write "		<td class=tablerow1><input type=""text"" name=""product"" size=50></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>�� ����</b></td>"
	Response.Write "		<td class=tablerow2><input type=""text"" name=""Amount"" size=5></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>�� λ��</b></td>"
	Response.Write "		<td class=tablerow1><input type=""text"" name=""unit"" size=5></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>�� �ۣ�</b></td>"
	Response.Write "		<td class=tablerow2><input type=""text"" name=""price"" size=10> Ԫ</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>�� �� �</b></td>"
	Response.Write "		<td class=tablerow1><input type=""text"" name=""TotalPrices"" size=10> Ԫ</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>����ʱ�䣺</b></td>"
	Response.Write "		<td class=tablerow2><input type=""text"" name=""DateAndTime"" size=30 value="""& Now() &"""></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>�������ͣ�</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name=""Accountype"">"
	Response.Write "	<option value=0>�� ��</option>"
	Response.Write "	<option value=1>֧ ��</option>"
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>����˵����</b></td>"
	Response.Write "		<td class=tablerow2><textarea name=Explain rows=5 cols=50></textarea></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow1 colspan=2><input type=""button"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" name=""B1"" class=Button>&nbsp;&nbsp;"
	Response.Write "		<input type=submit name=submit2 value=""������ϸ��"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub
Sub SavenewAccount()
	If Not IsNumeric(Request("price")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����������</li>"
	End If
	If Not IsNumeric(Request("TotalPrices")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�ܽ���������</li>"
	End If
	If Not IsDate(Request("DateAndTime")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ʱ���������</li>"
	End If
	If FoundErr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_Account WHERE (AccountID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.addnew
		Rs("payer").Value = Trim(Request.Form("payer"))
		Rs("payee").Value = Trim(Request.Form("payee"))
		Rs("product").Value = Trim(Request.Form("product"))
		Rs("Amount").Value = Trim(Request.Form("Amount"))
		Rs("unit").Value = Trim(Request.Form("unit"))
		Rs("price").Value = Trim(Request.Form("price"))
		Rs("TotalPrices").Value = Trim(Request.Form("TotalPrices"))
		Rs("DateAndTime").Value = Trim(Request.Form("DateAndTime"))
		Rs("Accountype").Value = Trim(Request.Form("Accountype"))
		Rs("Explain").Value = Trim(Request.Form("Explain"))
		Rs("Reclaim").Value = 0
	Rs.update
	Rs.Close:Set Rs = Nothing
	Succeed("<li>��ϲ���������ϸ�ɹ���</li>")
End Sub
Sub SaveAccount()
	If Trim(Request("AccountID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>AccountID��������Ϊ�գ�</li>"
		Exit Sub
	End If
	If Not IsNumeric(Request("AccountID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������ȷ��ID������</li>"
		Exit Sub
	End If
	If Not IsNumeric(Request("price")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����������</li>"
	End If
	If Not IsNumeric(Request("TotalPrices")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�ܽ���������</li>"
	End If
	If Not IsDate(Request("DateAndTime")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ʱ���������</li>"
	End If
	If FoundErr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_Account WHERE AccountID="& CLng(Request("AccountID"))
	Rs.Open SQL,Conn,1,3
		Rs("payer").Value = Trim(Request.Form("payer"))
		Rs("payee").Value = Trim(Request.Form("payee"))
		Rs("product").Value = Trim(Request.Form("product"))
		Rs("Amount").Value = Trim(Request.Form("Amount"))
		Rs("unit").Value = Trim(Request.Form("unit"))
		Rs("price").Value = Trim(Request.Form("price"))
		Rs("TotalPrices").Value = Trim(Request.Form("TotalPrices"))
		Rs("DateAndTime").Value = Trim(Request.Form("DateAndTime"))
		Rs("Accountype").Value = Trim(Request.Form("Accountype"))
		Rs("Explain").Value = Trim(Request.Form("Explain"))
		Rs("Reclaim").Value = 0
	Rs.update
	Rs.Close:Set Rs = Nothing
	Succeed("<li>��ϲ�����޸���ϸ�ɹ���</li>")
End Sub
%>
