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
Dim Action,jobID
Dim i,isEdit,TextContent,FoundSQL,oRs,strClass
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Action = LCase(Request("action"))
%>

<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th>��Ƹ����</th>
	
</tr>
<tr>
	<td class=tablerow2><strong>��Ƹ������</strong> <a href='admin_job.asp'>��Ƹ��ҳ</a> 
| <a href='admin_job.asp?action=add'><font color=blue>������Ƹ</font></a> 
	</td>
</tr>
</table>
<br>
<%
'Ȩ���ж�
If Not ChkAdmin("adminjob") Then
			Server.Transfer("showerr.asp")
			Response.End
End If
Select Case Trim(Action)
	Case "save"
		Call Savejob
	Case "modify"
		Call Modifyjob
	Case "add"
		isEdit = False
		Call Editjob(isEdit)
	Case "edit"
		isEdit = True
		Call Editjob(isEdit)
	Case "view"
		Call Viewjob
	Case "del"
		Call Deljob
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
	<th class=tablerow1>ID</th>
	<th class=tablerow1>��Ƹ����</th>
	<th class=tablerow1>��Ƹ����</th>
	<th class=tablerow1>����ʱ��</th>
	<th class=tablerow1>��Ч����</th>
	<th class=tablerow1>����</th>
</tr>
<%
	maxperpage = 20 '###ÿҳ��ʾ��
	
	If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
		Response.Write ("�����ϵͳ����!����������")
		Response.End
	End If
	If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
		CurrentPage = CInt(Request("page"))
	Else
		CurrentPage = 1
	End If
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	FoundSQL = " where isdel=0 "
	TotalNumber = enchiasp.Execute("Select Count(ID) from ECCMS_Job "& FoundSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Job] "& FoundSQL &" order by riqi desc ,ID desc"

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
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=6 class=TableRow2>��û���ҵ��κ���Ƹ��Ϣ��</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td colspan=6 class=tablerow2><%Call showpage()%></td>
</tr>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
		Response.Write "<tr>"
		Response.Write "	<td " & strClass & "><a href='?action=view&ID="
		Response.Write Rs("ID")
		Response.Write "	'>"
		Response.Write Rs("id")
		Response.Write "	</a></td>"
		Response.Write "	<td align=center " & strClass & ">"
		Response.Write "	<a href='?action=view&ID="
		Response.Write Rs("ID")
		Response.Write "	'>"
		Response.Write Rs("duix")
		Response.Write "	</a>"
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("rens")
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		If Rs("riqi") >= Date Then
			Response.Write "<font color=red>"
			Response.Write enchiasp.FormatDate(Rs("riqi"), 2)
			Response.Write "</font>"
		Else
			Response.Write enchiasp.FormatDate(Rs("riqi"), 2)
		End If		
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		response.write enchiasp.FormatDate(Rs("riqi")+Rs("qix"), 2)
		If (Rs("riqi")+Rs("qix")< Date) Then
			Response.Write "<font color=red>"
			Response.Write "�����ڣ�"
			Response.Write "</font>"
		Else
			Response.Write "����Ч��"
		End If		
		Response.Write "	</td>"

		

%>
	
	<td align=center <%=strClass%>><a href='?action=edit&ID=<%=Rs("ID")%>'>�༭</a> | 
	<a href='?action=del&ID=<%=Rs("ID")%>' onclick="{if(confirm('��Ƹ��Ϣɾ���󽫲��ָܻ�����ȷ��Ҫɾ������Ƹ��Ϣ��?')){return true;}return false;}">ɾ��</a></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td colspan=6 class=tablerow2><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub Editjob(isEdit)

	Dim EditTitle
	If isEdit Then
		SQL = "select * from ECCMS_job where ID=" & Request("ID")
		Set Rs = enchiasp.Execute(SQL)
		EditTitle = "�༭��Ƹ"
	Else
		EditTitle = "�����Ƹ"
	End If
%>
<script language=javascript>
    function CheckForm(form1)
{
	if (!validateSubmit()) return (false);
	if (form1.duix.value == "")
	{
		alert("��Ƹְλ����Ϊ�գ�");
		form1.duix.focus();
		return (false);
	}
	form1.content.value=IframeID.document.body.innerHTML; 
	MessageLength=IframeID.document.body.innerHTML.length;
	if(MessageLength<2){alert("���ݲ���С��2���ַ���");return false;}
}
</script>
<div onkeydown=CtrlEnter()>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
  <tr>
    <th colspan="2"><%=EditTitle%></th>
  </tr>
    	<form method=Post name="myform" action="admin_job.asp" onSubmit="return CheckForm(this);">
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""ID"" value="""& Request("ID") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
%>
  <tr>
    <td width="20%" align="right" class="TableRow2"><strong>��Ƹְλ��</strong></td>
    <td width="80%" class="TableRow1"><input name="duix" type="text" id="duix" size="50" value='<%If isEdit Then Response.Write Rs("duix")%>'> 
      <span class="style1">* </span></td>
  </tr>
  
	<tr>
    <td width="20%" align="right" class="TableRow2"><strong>��Ƹ������</strong></td>
    <td width="80%" class="TableRow1"><input name="rens" type="text" id="rens" size="50" value='<%If isEdit Then Response.Write Rs("rens")%>'> 
      <span class="style1">* (ֻ����������)</span></td>
  </tr>
	
	<tr>
    <td width="20%" align="right" class="TableRow2"><strong>�Ա�Ҫ��</strong></td>
    <td width="80%" class="TableRow1"><input name="sex" type="text" id="sex" size="50" value='<%If isEdit Then Response.Write Rs("sex") else response.write "����" end if %>'> 
   </td>
  </tr>

		<tr>
    <td width="20%" align="right" class="TableRow2"><strong>ѧ��Ҫ��</strong></td>
    <td width="80%" class="TableRow1"><input name="xueli" type="text" id="xueli" size="50" value='<%If isEdit Then Response.Write Rs("xueli") else response.write "����" end if %>'> 
   </td>
  </tr>

	<tr>
    <td width="20%" align="right" class="TableRow2"><strong>רҵҪ��</strong></td>
    <td width="80%" class="TableRow1"><input name="zhuanye" type="text" id="zhuanye" size="50" value='<%If isEdit Then Response.Write Rs("zhuanye") else response.write "����" end if %>'> 
   </td>
  </tr>

	
	
	<tr>
    <td width="20%" align="right" class="TableRow2"><strong>�����ص㣺</strong></td>
    <td width="80%" class="TableRow1"><input name="did" type="text" id="did" size="50" value='<%If isEdit Then Response.Write Rs("did") else response.write "�ܲ�" end if %>'> 
     </td>
  </tr>

<tr>
    <td width="20%" align="right" class="TableRow2"><strong>����������</strong></td>
    <td width="80%" class="TableRow1"><input name="daiy" type="text" id="daiy" size="50" value='<%If isEdit Then Response.Write Rs("daiy") else response.write "��̸" end if%>'> 
     </td>
  </tr>
  
  <tr>
    <td align="right" class="TableRow2"><strong>��ƸҪ��</strong></td>
    <td class="TableRow1"><textarea name="content" style="display:none" id="content"><%If isEdit Then Response.Write Server.HTMLEncode(Rs("yaoq"))%></textarea>
    <script src='../editor/edit.js' type=text/javascript></script></td>
  </tr>
  
  
<tr>
    <td width="20%" align="right" class="TableRow2"><strong>����ʱ�䣺</strong></td>
    <td width="80%" class="TableRow1"><input name="riqi" disabled  type="text" id="riqi" size="50" value='<%If isEdit Then Response.Write Rs("riqi") else Response.Write now() end if %>'> 
      <span class="style1">* </span></td>
  </tr>
<tr>
    <td width="20%" align="right" class="TableRow2"><strong>��Ƹ��Ч������</strong></td>
    <td width="80%" class="TableRow1"><input name="qix" type="text" id="qix" size="50" value='<%If isEdit Then Response.Write Rs("qix") else response.write "30" end if %>'> 
      <span class="style1">*(ֻ����������) </span></td>
  </tr>


	


      <%If isEdit Then%>
  <tr>
  <td class="TableRow1"></td>
    <td class="TableRow1">
      <input name="update" type="checkbox" id="update" value="yes">
������Ƹ����ʱ�� 
<%End If%></td>
  </tr>
  
  <tr>
    <td align="right" class="TableRow2">��</td>
    <td align="center" class="TableRow1">
    <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>
    <input name="Submit1" type="submit" class="Button" value="������Ƹ" class=Button></td>
  </tr></form>
</table>
</div>
<%
	If isEdit Then Rs.Close:Set Rs = Nothing
End Sub

Private Sub CheckSave()
	If Trim(Request.Form("duix")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��Ƹְλ����Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("rens")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��Ƹ��������Ϊ�գ�</li>"
	else
		if not IsNumeric(Trim(Request.Form("rens"))) then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��Ƹ����ֻ���������֣�</li>"
		end if
	End If
	
	If Trim(Request.Form("qix")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��Ƹ��Ч�ڲ���Ϊ�գ�</li>"
	else
		if not IsNumeric(Trim(Request.Form("qix"))) then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��Ƹ��Ч��ֻ���������֣�</li>"
		end if
	End If

	
	If Trim(Request.Form("content")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ƸҪ����Ϊ�գ�</li>"
	End If
	'Ҫ��
	TextContent = ""
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next

End Sub
Private Sub Savejob()
	

	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_job where (ID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("duix") = Trim(Request.Form("duix"))
		Rs("rens") = enchiasp.ChkFormStr(Request.Form("rens"))
		Rs("did") = Request.Form("did")
		Rs("daiy") =enchiasp.ChkFormStr(Request.Form("daiy"))
		Rs("yaoq") = TextContent
		Rs("qix") = enchiasp.ChkFormStr(Request.Form("qix"))
		Rs("sex") = Request.Form("sex")
		Rs("xueli") = Request.Form("xueli")
		Rs("zhuanye") = Request.Form("zhuanye")
		Rs("riqi") = Now()
		
	Rs.update
	Rs.Close
	Rs.Open "select top 1 ID from ECCMS_job order by ID desc", Conn, 1, 1
	jobID = Rs("ID")
	Rs.Close:Set Rs = Nothing
	Succeed("<li>��ϲ��������µ���Ƹ��Ϣ�ɹ���</li><li><a href=?action=view&ID=" & jobID & ">����˴��鿴����Ƹ��Ϣ</a></li>")
End Sub
Private Sub Modifyjob()
	

	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_job where ID = " & Request("ID")
	Rs.Open SQL,Conn,1,3
		Rs("duix") = Trim(Request.Form("duix"))
		Rs("rens") = enchiasp.ChkFormStr(Request.Form("rens"))
		Rs("did") = Request.Form("did")
		Rs("daiy") =  enchiasp.ChkFormStr(Request.Form("daiy"))
		Rs("sex") = Request.Form("sex")
		Rs("xueli") = Request.Form("xueli")
		Rs("zhuanye") = Request.Form("zhuanye")
		Rs("yaoq") = TextContent
		Rs("qix") = enchiasp.ChkFormStr(Request.Form("qix"))
		If LCase(Request.Form("Update")) = "yes" Then Rs("riqi") = Now()
	Rs.update
		jobID = Rs("ID")
	Rs.Close:Set Rs = Nothing
	Succeed("<li>��ϲ�����޸���Ƹ�ɹ���</li><li><a href=?action=view&ID=" & jobID & ">����˴��鿴����Ƹ</a></li>")
End Sub
Private Sub Deljob()
	
	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������ƸID��</li>"
		Exit Sub
	End If
	enchiasp.Execute("update [ECCMS_job] set isdel=1 where ID = " & Request("ID"))
	OutHintScript("��Ƹɾ���ɹ���")
End Sub


Private Sub Viewjob()
	

	If Request("ID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	dim n
	n = 1
	SQL = "select * from ECCMS_job where ID=" & Request("ID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κ���Ƹ��������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	Else
%>
<table border="0" align="center" cellpadding="2" cellspacing="1" class="TableBorder">
	<tr>
	  <th ></th>
	  <th >�鿴��Ƹ</th>

	</tr>

	<tr>
	  <td class="TableRow1"><strong>��Ƹְλ��</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("duix"))%></td>
	</tr>
	
	<tr>
	  <td class="TableRow1"><strong>��Ƹ������</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("rens"))%></td>
	</tr>
	
	<tr>
	  <td class="TableRow1"><strong>����ʱ�䣺</strong></td><td class="TableRow1"><%=enchiasp.FormatDate(Rs("riqi"), 2)%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>�Ա�Ҫ��</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("sex"))%></td>
	</tr>
<tr>
	  <td class="TableRow1"><strong>ѧ��Ҫ��</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("xueli"))%></td>
	</tr>
<tr>
	  <td class="TableRow1"><strong>רҵҪ��</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("zhuanye"))%></td>
	</tr>

	
	<tr>
	  <td class="TableRow1"><strong>�����ص㣺</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("did"))%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>��Ч�ڣ�</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("qix"))%></td>
	</tr>
	
	<tr>
	  <td class="TableRow1"><strong>��ֹ���ڣ�</strong></td><td class="TableRow1"><%=enchiasp.FormatDate(Rs("riqi")+enchiasp.ReadContent(Rs("qix")), 2)%></td>
	</tr>

	
	<tr>
	  <td class="TableRow1"><strong>��λ������</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("daiy"))%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>��λҪ��</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("yaoq"))%></td>
	</tr>

	
	<tr>
	  <td class="TableRow2">��һ��Ƹ��<%=Frontjob(Rs("ID"))%>
	  <br>��һ��Ƹ��<%=Nextjob(Rs("ID"))%></td>
	  <td class="TableRow1"></td>
	</tr>
	<tr>
	  <td class="TableRow1"></td>
	  <td align="center" class="TableRow2"><input type="button" onclick="javascript:window.close()" value="�رձ�����" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="���ض���" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&ID=<%=Rs("ID")%>'" value="�༭��Ƹ" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub

Private Function Frontjob(jobID)
	Dim Rss, SQL
	SQL = "select Top 1 ID,duix from ECCMS_job where ID < " & jobID & " order by ID desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		Frontjob = "�Ѿ�û����"
	Else
		Frontjob = "<a href=admin_job.asp?action=view&ID=" & Rss("ID") & ">" & Rss("duix") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function Nextjob(jobID)
	Dim Rss, SQL
	SQL = "select Top 1 ID,duix from ECCMS_job where ID > " & jobID & " order by ID asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		Nextjob = "�Ѿ�û����"
	Else
		Nextjob = "<a href=admin_job.asp?action=view&ID=" & Rss("ID") & ">" & Rss("duix") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Sub showpage()
	Dim n
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action=?ChannelID=" & Request("ChannelID") & "><tr><td align=center> " & vbCrLf
	If CurrentPage < 2 Then
		Response.Write "������Ƹ <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ƪ&nbsp;�� ҳ&nbsp;��һҳ&nbsp;|&nbsp;"
	Else
		Response.Write "������Ƹ <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ƪ&nbsp;<a href=?page=1&ChannelID=" & Request("ChannelID") & ">�� ҳ</a>&nbsp;"
		Response.Write "<a href=?page=" & CurrentPage - 1 & "&ChannelID=" & Request("ChannelID") & ">��һҳ</a>&nbsp;|&nbsp;"
	End If
	If n - CurrentPage < 1 Then
		Response.Write "��һҳ&nbsp;β ҳ" & vbCrLf
	Else
		Response.Write "<a href=?page=" & (CurrentPage + 1) & "&ChannelID=" & Request("ChannelID") & ">��һҳ</a>"
		Response.Write "&nbsp;<a href=?page=" & n & "&ChannelID=" & Request("ChannelID") & ">β ҳ</a>" & vbCrLf
	End If
	Response.Write "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
	Response.Write "&nbsp;ת����"
	Response.Write "<input name=page size=3 value='" & CurrentPage & "'> <input class=Button type=submit name=Submit value='ת��'>"
	Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub

%>