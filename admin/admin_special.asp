<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="include/MenuCode.Asp"-->
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
Dim Action,Flag,i,ChannelPath,strOption
If Request("ChannelID") = 0 Or Request("ChannelID") = "" Then
	ErrMsg = "<li>Sorry�������ϵͳ����,��ѡ����ȷ�����ӷ�ʽ��</li>"
	Response.Redirect("showerr.asp?action=error&message=" & Server.URLEncode(ErrMsg) & "")
	Response.End
Else
	ChannelID = CInt(Request("ChannelID"))
End If
ChannelPath = enchiasp.GetChannelDir(ChannelID)
%>
<script language="javascript">
function formatbt()
{
  var arr = showModalDialog("include/btformat.htm?",null, "dialogWidth:250pt;dialogHeight:166pt;toolbar=no;location=no;directories=no;status=no;menubar=NO;scrollbars=no;resizable=no;help=0; status:0");
  if (arr != null){
     document.myform.Topicformat.value=arr;
     myt.innerHTML="<span style='background-color: #FFFFff;font-size:12px' "+arr+">���ñ�����ʽ ABCdef</span>";
  }
}
function Cancelform()
{
  document.myform.Topicformat.value='';
  myt.innerHTML="<span style='background-color: #FFFFff;font-size:12px'>���ñ�����ʽ ABCdef</span>";
}
//-->
</script>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th><%=sChannelName%>ר����Ŀ����</th>
</tr>
<tr>
	<td class=tablerow1><strong>���������</strong> <a href='?ChannelID=<%=ChannelID%>'>������ҳ</a> | 
	<a href='?action=add&ChannelID=<%=ChannelID%>'>���ר����Ŀ</a> | <a href='?action=orders&ChannelID=<%=ChannelID%>'>ר����Ŀ����</a> | <a href='?action=make&ChannelID=<%=ChannelID%>&stype=2'><font color=blue>����ר����Ŀ�˵�</font></a></td>
</tr>
</table>
<br>
<%
Flag = "Special" & ChannelID
Action = LCase(Request("action"))
If Not ChkAdmin(Flag) Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Select Case Trim(Action)
	Case "save"
		Call SaveSpecial
	Case "modify"
		Call ModifySpecial
	Case "edit"
		Call EditSpecial
	Case "add"
		Call AddSpecial
	Case "del"
		Call DelSpecial
	Case "orders"
		Call SpecialOrder
	Case "saveorder"
		Call SpecialRenewOrder
	Case "make"
		'Call CreationSpecialMenu
		Call CreationJsMenu
	Case Else
		Call MainPage
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub MainPage()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th>ר������</th>
	<th>����Ŀ��</th>
	<th>��������</th>
	<th>�޸�ר����Ŀ</th>
	<th>ɾ��ר����Ŀ</th>
</tr>
<%
Set Rs = enchiasp.Execute("Select SpecialID,ChannelID,SpecialName,Topicformat,Reopen,ChangeLink,SpecialUrl From ECCMS_Special where ChannelID = "& ChannelID &" order by orders,SpecialID")
If Rs.BOF And Rs.EOF Then
	Response.Write "<tr><td align=center colspan=5 class=TableRow2>��û������κ�ר��</td></tr>"
Else
	Do While Not Rs.EOF
		Response.Write "<tr align=center>"
		Response.Write "	<td class=tablerow1>"
		Response.Write "<A href=?action=edit&ChannelID="
		Response.Write Rs("ChannelID")
		Response.Write "&SpecialID="
		Response.Write Rs("SpecialID")
		Response.Write "><span "
		Response.Write Rs("Topicformat")
		Response.Write ">"
		Response.Write Rs("SpecialName")
		Response.Write "</span></A>"
		Response.Write "	</td>"
		Response.Write "	<td class=tablerow1>"
		If Rs("Reopen") <> 0 Then
			Response.Write "<font color=red>�´��ڴ�</font>"
		Else
			Response.Write "<font color=blue>�����ڴ�</font>"
		End If
		Response.Write "	</td>"
		Response.Write "	<td class=tablerow1>"
		If Rs("ChangeLink") <> 0 Then
			Response.Write "<font color=red>ת������</font>"
		Else
			Response.Write "<font color=blue>�ڲ�����</font>"
		End If
		Response.Write "	</td>"
		Response.Write "	<td class=tablerow1>"
		Response.Write "<A href=?action=edit&ChannelID="
		Response.Write Rs("ChannelID")
		Response.Write "&SpecialID="
		Response.Write Rs("SpecialID")
		Response.Write ">�� �� ר ��</A>"
		Response.Write "	</td>"
		Response.Write "	<td class=tablerow1>"
		Response.Write "<A href=?action=del&ChannelID="
		Response.Write Rs("ChannelID")
		Response.Write "&SpecialID="
		Response.Write Rs("SpecialID")
		Response.Write " onclick=""{if(confirm('�˲�����ɾ����ר��\n��ȷ��Ҫɾ����?')){return true;}return false;}"">ɾ �� ר ��</A>"
		Response.Write "	</td>"
		Response.Write "</tr>"
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
End If
%>
<tr align=center>
	<td colspan=5 class=tablerow1><input type=button onclick="javascript:location.href='?action=make&ChannelID=<%=ChannelID%>&stype=1'" value=' ����ר����Ŀ�˵�JS ' class=button></td>
</tr>
</table>

<%
End Sub

Private Sub AddSpecial()
	Dim NewSpecialID
	SQL = "select Max(SpecialID) from ECCMS_Special"
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		NewSpecialID = 1
	Else
		NewSpecialID = Rs(0) + 1
	End If
	If IsNull(NewSpecialID) Then NewSpecialID = 1
	Rs.Close:Set Rs = Nothing
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan=2>���<%=sModuleName%>ר����Ŀ</th>
</tr>
<form name=myform method=post action=?action=save>
<input type=hidden name=ChannelID value='<%=ChannelID%>'>
<input type=hidden name=SpecialID value='<%=NewSpecialID%>'>
<tr>
	<td width="20%" class=tablerow1><strong>ר����Ŀ���ƣ�</strong></td>
	<td width="80%" class=tablerow1><input type=text name=SpecialName size=20 value=''>  
	��ʽ:<input   type="hidden" name="Topicformat" size="1" value="">&nbsp; 
	<span style="background-color: #fFfFff" id="myt" onclick="javascript:formatbt(this);"  style='cursor:hand; font-size:11pt' >���ñ�����ʽ ABCdef</span> 
	<input type=checkbox name=cancel value='' onclick="Cancelform()"> ȡ����ʽ</td>
</tr>
<tr>
	<td class=tablerow2><strong>ר����Ŀ˵����</strong></td>
	<td class=tablerow2><input type=text name=Readme size=50 id=Readme value=''></td>
</tr>
<tr>
	<td class=tablerow1><strong>ר������Ŀ¼��</strong></td>
	<td class=tablerow1><input type=text name=SpecialDir size=20 value=''></td>
</tr>
<tr>
	<td class=tablerow2><strong>�Ƿ��´��ڴ򿪣�</strong></td>
	<td class=tablerow2><input type=radio name=Reopen value='0' checked> ��&nbsp;&nbsp;
	<input type=radio name=Reopen value='1'> ��&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class=tablerow1><strong>�Ƿ�ת�����ӣ�</strong></td>
	<td class=tablerow1><input type=radio name=ChangeLink value='0' checked onClick="ChangeSetting.style.display='none';"> ��&nbsp;&nbsp;
	<input type=radio name=ChangeLink value='1' onClick="ChangeSetting.style.display='';"> ��&nbsp;&nbsp;</td>
</tr>
<tr id=ChangeSetting style="display:none">
	<td class=tablerow2><strong>ת������URL��</strong></td>
	<td class=tablerow2><input type=text name=SpecialUrl size=50 value='http://'></td>
</tr>
<tr align=center>
	<td class=tablerow2></td>
	<td class=tablerow2><input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp;
		<input type="submit" value="���ר��" name="B2" class=Button></td>
</tr>
</form>
</table>

<%
End Sub

Private Sub EditSpecial()
	Set Rs = enchiasp.Execute("Select SpecialID,SpecialName,Topicformat,Readme,Reopen,SpecialDir,ChangeLink,SpecialUrl From ECCMS_Special where ChannelID = "& ChannelID &" And SpecialID = " & Request("SpecialID"))
	If Rs.BOF And Rs.EOF Then
		Response.Write "��������"
		Exit Sub
	End If
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan=2>�޸�<%=sModuleName%>ר����Ŀ</th>
</tr>
<form name=myform method=post action=?action=modify>
<input type=hidden name=ChannelID value='<%=ChannelID%>'>
<input type=hidden name=SpecialID value='<%=Rs("SpecialID")%>'>
<tr>
	<td width="20%" class=tablerow1><strong>ר����Ŀ���ƣ�</strong></td>
	<td width="80%" class=tablerow1><input type=text name=SpecialName size=20 value='<%=Rs("SpecialName")%>'>  
	��ʽ:<input   type="hidden" name="Topicformat" size="1" value="<%=Server.HTMLEncode(Rs("Topicformat"))%>">&nbsp; 
	<span style="background-color: #fFfFff" <%=Rs("Topicformat")%> id="myt" onclick="javascript:formatbt(this);"  style='cursor:hand; font-size:11pt' >���ñ�����ʽ ABCdef</span>
	<input type=checkbox name=cancel value='' onclick="Cancelform()"> ȡ����ʽ</td>
</tr>
<tr>
	<td class=tablerow2><strong>ר����Ŀ˵����</strong></td>
	<td class=tablerow2><input type=text name=Readme size=50 value='<%=Rs("Readme")%>'></td>
</tr>
<tr>
	<td class=tablerow1><strong>ר������Ŀ¼��</strong></td>
	<td class=tablerow1><input type=text name=SpecialDir size=20 value='<%=Rs("SpecialDir")%>'></td>
</tr>
<tr>
	<td class=tablerow2><strong>�Ƿ��´��ڴ򿪣�</strong></td>
	<td class=tablerow2><input type=radio name=Reopen value='0'<%If Rs("Reopen") = 0 Then Response.Write (" checked")%>> ��&nbsp;&nbsp;
	<input type=radio name=Reopen value='1'<%If Rs("Reopen") = 1 Then Response.Write (" checked")%>> ��&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class=tablerow1><strong>�Ƿ�ת�����ӣ�</strong></td>
	<td class=tablerow1><input type=radio name=ChangeLink value='0'<%If Rs("ChangeLink") = 0 Then Response.Write (" checked")%> onClick="ChangeSetting.style.display='none';"> ��&nbsp;&nbsp;
	<input type=radio name=ChangeLink value='1'<%If Rs("ChangeLink") = 1 Then Response.Write (" checked")%> onClick="ChangeSetting.style.display='';"> ��&nbsp;&nbsp;</td>
</tr>
<tr id=ChangeSetting<%If Rs("ChangeLink") = 0 Then Response.Write (" style=""display:none""")%>>
	<td class=tablerow2><strong>ת������URL��</strong></td>
	<td class=tablerow2><input type=text name=SpecialUrl size=50 value='<%=Rs("SpecialUrl")%>'></td>
</tr>
<tr align=center>
	<td class=tablerow2></td>
	<td class=tablerow2><input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp;
		<input type="submit" value="�޸�ר��" name="B2" class=Button></td>
</tr>
</form>
</table>
<%
Rs.Close:Set Rs = Nothing
End Sub

Private Sub CheckSave()
	If Trim(Request.Form("SpecialName")) = "" Or Len(Request.Form("SpecialName")) => 30 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ר�����Ʋ���Ϊ�ջ��߳���30���ַ���</li>"
	End If
	If Trim(Request.Form("Readme")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ר��˵������Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("SpecialDir")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ר��Ŀ¼����Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("SpecialUrl")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ר������URL����Ϊ�գ�</li>"
	End If
	If Not enchiasp.IsValidChar(Trim(Request.Form("SpecialDir"))) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ר��Ŀ¼�к��зǷ��ַ����������ַ���</li>"
	End If
End Sub

Private Sub SaveSpecial()
	Call CheckSave
	Dim neworders,NewSpecialID,SpecialDir
	Set Rs = Conn.Execute("select SpecialID from ECCMS_Special where SpecialID = " & Request("SpecialID"))
	If Not (Rs.EOF And Rs.bof) Then
		ErrMsg = "<li>������ָ���ͱ��Ƶ��һ������š�</li>"
		Founderr = True
		Exit Sub
	Else
		NewSpecialID = Request("SpecialID")
	End If
	SpecialDir = Replace(Replace(Trim(Request.Form("SpecialDir")), "\", ""), "/", "")
	If Founderr = True Then Exit Sub
	Set Rs = enchiasp.Execute ("Select Max(orders) from ECCMS_Special where ChannelID = " & Request("ChannelID"))
	If Not (Rs.EOF And Rs.bof) Then
		neworders = Rs(0)
	End If
	If IsNull(neworders) Then neworders = 0
	Rs.Close
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Special"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("SpecialID") = NewSpecialID
		Rs("ChannelID") = Trim(Request.Form("ChannelID"))
		Rs("SpecialName") = Trim(Request.Form("SpecialName"))
		Rs("Topicformat") = Trim(Request.Form("Topicformat"))
		Rs("Readme") = Trim(Request.Form("Readme"))
		Rs("orders") = neworders + 1
		Rs("Reopen") = Trim(Request.Form("Reopen"))
		Rs("SpecialDir") = Trim(Request.Form("SpecialDir"))
		Rs("ChangeLink") = Trim(Request.Form("ChangeLink"))
		Rs("SpecialUrl") = Trim(Request.Form("SpecialUrl"))
	Rs.update
	Rs.Close:Set Rs = Nothing
	Succeed("<li>����µ�ר����Ŀ�ɹ�</li>")
	Dim FilePath
	FilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "Special/" & SpecialDir
	enchiasp.CreatPathEx(FilePath)
	Call CreationSpecialMenu
End Sub

Private Sub ModifySpecial()
	Call CheckSave
	Dim SpecialDir
	SpecialDir = Replace(Replace(Trim(Request.Form("SpecialDir")), "\", ""), "/", "")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Special where SpecialID = " & Request("SpecialID")
	Rs.Open SQL,Conn,1,3
		Rs("SpecialName") = Trim(Request.Form("SpecialName"))
		Rs("Topicformat") = Trim(Request.Form("Topicformat"))
		Rs("Readme") = Trim(Request.Form("Readme"))
		'Rs("orders") = neworders + 1
		Rs("Reopen") = Trim(Request.Form("Reopen"))
		Rs("SpecialDir") = Trim(Request.Form("SpecialDir"))
		Rs("ChangeLink") = Trim(Request.Form("ChangeLink"))
		Rs("SpecialUrl") = Trim(Request.Form("SpecialUrl"))
	Rs.update
	Rs.Close:Set Rs = Nothing
	Succeed("<li>�޸�" & sChannelName & "��ר����Ŀ�ɹ�</li>")
	Call CreationSpecialMenu
	Dim FilePath
	FilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "Special/" & SpecialDir
	enchiasp.CreatPathEx(FilePath)
	Call CreationSpecialMenu
End Sub

Private Sub DelSpecial()
	Dim FolderPath
	If Trim(Request("SpecialID")) <> "" Then
		Set Rs = enchiasp.Execute("Select SpecialDir From ECCMS_Special where SpecialID = " & Request("SpecialID"))
		FolderPath = enchiasp.InstallDir & enchiasp.ChannelDir & "Special/" & Rs("SpecialDir")
		enchiasp.FolderDelete(FolderPath)
		enchiasp.Execute("Delete From ECCMS_Special where SpecialID = " & Request("SpecialID"))
		Rs.Close:Set Rs = Nothing
		OutHintScript (sChannelName & "ר����Ŀɾ�������ɹ���")
	Else
		OutHintScript ("��ѡ����ȷ��ϵͳ������")
	End If
End Sub

Private Sub SpecialOrder()
	Dim trs
	Dim uporders
	Dim doorders
	Response.Write " <table border=""0"" cellspacing=""1"" cellpadding=""2"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th colspan=2>" & sChannelName & "ר����Ŀ���������޸�"
	Response.Write " </th>"
	Response.Write " </tr>" & vbCrLf
	SQL = "select * from ECCMS_Special where ChannelID = "& Request("ChannelID") &" order by orders"
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		Response.Write "����û�������Ӧ��ר�⡣"
	Else
		Do While Not Rs.EOF
			Response.Write "<form action=?action=saveorder method=post><tr><td width=""50%"" class=TableRow1>" & vbCrLf
			Response.Write "<span " & Rs("Topicformat") & ">" & Rs("SpecialName") & "</span>"
			Response.Write "</td><td width=""50%"" class=TableRow2>" & vbCrLf
			Set trs = enchiasp.Execute("select count(*) from ECCMS_Special where ChannelID = "& Request("ChannelID") &" And orders<" & Rs("orders") & "")
				uporders = trs(0)
				If IsNull(uporders) Then uporders = 0

				Set trs = enchiasp.Execute("select count(*) from ECCMS_Special where ChannelID = "& Request("ChannelID") &" And orders>" & Rs("orders") & "")
				doorders = trs(0)
				If IsNull(doorders) Then doorders = 0
				If uporders > 0 Then
					Response.Write "<select name=uporders size=1><option value=0>��</option>" & vbCrLf
					For i = 1 To uporders
						Response.Write "<option value=" & i & ">��" & i & "</option>" & vbCrLf
					Next
					Response.Write "</select>"
				End If
				If doorders > 0 Then
					If uporders > 0 Then Response.Write "&nbsp;"
					Response.Write "<select name=doorders size=1><option value=0>��</option>" & vbCrLf
					For i = 1 To doorders
						Response.Write "<option value=" & i & ">��" & i & "</option>" & vbCrLf
					Next
					Response.Write "</select>" & vbCrLf
				End If
				If doorders > 0 Or uporders > 0 Then
					Response.Write "<input type=hidden name=""ChannelID"" value=""" & Rs("ChannelID") & """><input type=hidden name=""SpecialID"" value=""" & Rs("SpecialID") & """>&nbsp;<input type=submit name=Submit class=button value='�� ��'>" & vbCrLf
				End If
			Response.Write "</td></tr></form>" & vbCrLf
			Rs.movenext
		Loop
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>"
End Sub

Private Sub SpecialRenewOrder()
	Dim orders
	Dim uporders
	Dim doorders
	Dim oldorders
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	If Not IsNumeric(Request("SpecialID")) Then
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("uporders") <> "" And Not CInt(Request("uporders")) = 0 Then
		If Not IsNumeric(Request("uporders")) Then
			ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
			Founderr = True
			Exit Sub
		ElseIf CInt(Request("uporders")) = 0 Then
			ErrMsg = ErrMsg & "<li>��ѡ��Ҫ���������֣�</li>"
			Founderr = True
			Exit Sub
		End If
		Set Rs = enchiasp.Execute("select SpecialID,orders from ECCMS_Special where ChannelID=" & Request("ChannelID") & " And SpecialID=" & Request("SpecialID"))
		orders = Rs(1)
		i = 0
		oldorders = 0
		Set Rs = enchiasp.Execute("select SpecialID,orders from ECCMS_Special where ChannelID=" & Request("ChannelID") & " And orders<" & orders & " order by orders desc")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("uporders")) >= i Then
				enchiasp.Execute ("update ECCMS_Special set orders=" & orders & "+" & oldorders & " where ChannelID=" & Request("ChannelID") & " And SpecialID=" & Rs(0))
				If CInt(Request("uporders")) = i Then uporders = Rs(1)
			End If
			orders = Rs(1)
			Rs.movenext
		Loop
		enchiasp.Execute ("update ECCMS_Special set orders=" & uporders & " where ChannelID=" & Request("ChannelID") & " And SpecialID=" & Request("SpecialID"))
		Set Rs = Nothing
	ElseIf Request("doorders") <> "" Then
		If Not IsNumeric(Request("doorders")) Then
			ErrMsg = ErrMsg & "<li>�Ƿ��Ĳ�����</li>"
			Founderr = True
			Exit Sub
		ElseIf CInt(Request("doorders")) = 0 Then
			ErrMsg = ErrMsg & "<li>��ѡ��Ҫ�½������֣�</li>"
			Founderr = True
			Exit Sub
		End If
		Set Rs = enchiasp.Execute("select SpecialID,orders from ECCMS_Special where ChannelID=" & Request("ChannelID") & " And SpecialID=" & Request("SpecialID"))
		orders = Rs(1)
		i = 0
		oldorders = 0
		Set Rs = enchiasp.Execute("select SpecialID,orders from ECCMS_Special where ChannelID=" & Request("ChannelID") & " And orders>" & orders & " order by orders desc")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("doorders")) >= i Then
				enchiasp.Execute ("update ECCMS_Special set orders=" & orders & " where ChannelID=" & Request("ChannelID") & " And SpecialID=" & Rs(0))
				If CInt(Request("doorders")) = i Then doorders = Rs(1)
			End If
			orders = Rs(1)
			Rs.movenext
		Loop
		enchiasp.Execute ("update ECCMS_Special set orders=" & doorders & " where ChannelID=" & Request("ChannelID") & " And SpecialID=" & Request("SpecialID"))
		Set Rs = Nothing
	End If
	Response.redirect "admin_special.asp?action=orders&ChannelID=" & ChannelID
End Sub
%>