<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
'=====================================================================
' ������ƣ�������վ����ϵͳ----���ѹ���
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
If Not ChkAdmin("adminconfirm") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "del"
	Call DelConfirm
Case "view"
	Call ViewConfirm
Case "pass"
	Call PassConfirm
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
	Dim CurrentPage,page_count,totalnumber,Pcount,maxperpage
	Dim tablebody
	maxperpage = 30
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CLng(CurrentPage)
	End If
	If CLng(CurrentPage) = 0 Then CurrentPage = 1
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th>ѡ��</th>"
	Response.Write "		<th>���ʽ</th>"
	Response.Write "		<th>�û�����</th>"
	Response.Write "		<th>�� �� ��</th>"
	Response.Write "		<th>֧�����</th>"
	Response.Write "		<th>���������</th>"
	Response.Write "		<th>���������</th>"
	Response.Write "		<th>���ʱ��</th>"
	Response.Write "		<th>�鿴˵��</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=selform method=post action='admin_confirm.asp'>"
	Response.Write "	<input type=hidden name=action value=""del"">"
	totalnumber = enchiasp.Execute("SELECT COUNT(id) FROM ECCMS_Confirm")(0)
	Pcount = CLng(totalnumber / maxperpage)  '�õ���ҳ��
	If Pcount < totalnumber / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_Confirm ORDER BY id DESC"
	If IsSqlDataBase=1 Then
		Set Rs = enchiasp.Execute(SQL)
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=9 class=TableRow1>û�н���ȷ�ϣ�</td></tr>"
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
			Response.Write "		<td " & tablebody & "><input type=checkbox name=id value="""& Rs("id") &"""></td>"
			Response.Write "		<td " & tablebody & ">" & Rs("paymode") & "</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("username") & "</td>"
			Response.Write "		<td " & tablebody & "><font color=red>" & Rs("indent") & "</font></td>"
			Response.Write "		<td " & tablebody & "><font color=blue>" & FormatCurrency(Rs("PayMoney"),2,-1) & "</font> Ԫ</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("customer") & "</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("Email") & "</td>"
			Response.Write "		<td " & tablebody & ">" & Rs("PayDate") & "</td>"
			Response.Write "		<td " & tablebody & ">"
			If Rs("ispass") > 0 Then
				Response.Write "<font color=blue>�Ѵ���</font>"
			Else
				Response.Write "<a href='?action=pass&id="&Rs("id")&"' title='���������ȷ��' onclick=""return confirm('ȷ���������Ϣ��?')""><font color=red>δ����</font></a>"
			End If
			Response.Write " | <a href=""?action=view&id="& Rs("id") &""" title='�鿴��ϸ˵��'>�鿴˵��</a>"
			Response.Write "</td>"
			Rs.movenext
			page_count = page_count + 1
			If page_count >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=9>"
	Response.Write "<input class=Button type=""button"" name=""chkall"" value=""ȫѡ"" onClick=""CheckAll(this.form)""><input class=Button type=""button"" name=""chksel"" value=""��ѡ"" onClick=""ContraSel(this.form)"">"
	Response.Write "<input type=submit name=submit2 value="" ɾ �� "" onclick=""return confirm('ȷ��ɾ����?')"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow2 colspan=9>"
	Response.Write ShowPages(CurrentPage,Pcount,totalnumber,maxperpage,"")
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
End Sub
Sub ViewConfirm()
	Set Rs = enchiasp.Execute("SELECT Readme FROM ECCMS_Confirm WHERE id="& CLng(Request("id")))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>û���ҵ�˵����</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th>ȷ��˵��</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=5>"
	Response.Write Rs("Readme")
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Set Rs = Nothing
End Sub
Sub DelConfirm()
	Dim selConfirmID
	If Not IsEmpty(Request("id")) Then
		selConfirmID = Request("id")
		enchiasp.Execute("DELETE FROM [ECCMS_Confirm] WHERE id in (" & selConfirmID & ")")
		Response.redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������ID����Ϊ�գ�</li>"
		Exit Sub
	End If
End Sub
Sub PassConfirm()
	Dim selConfirmID
	If Not IsEmpty(Request("id")) Then
		selConfirmID = Request("id")
		enchiasp.Execute("UPDATE [ECCMS_Confirm] SET isPass=1 WHERE id in (" & selConfirmID & ")")
		Response.redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������ID����Ϊ�գ�</li>"
		Exit Sub
	End If
End Sub
%>