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
Dim Action,jobbookID
Dim i,isEdit,TextContent,FoundSQL,oRs,strClass
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Action = LCase(Request("action"))
If Not ChkAdmin("adminjobbook") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
%>

<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
<%
if LCase(Request("isdel"))="1" then
%>
<th>ӦƸ����վ����</th>
<%
else
%>
	<th>ӦƸ����</th>
<%
end if
%>
</tr>
<tr>
	<td class=tablerow2><strong>��������</strong> 
<a href='admin_jobbook.asp'><font color=blue>ӦƸ����</font></a> 
| <a href='admin_jobbook.asp?isdel=1'><font color=blue>ӦƸ����վ</font></a> 
	</td>
</tr>
</table>
<br>
<%
Select Case Trim(Action)
	Case "view"
		Call Viewjobbook
	Case "del"
		Call Deljobbook
	case "huifu"
		call huifujobbook
	case "luyong"
		call luyongjobbook
	case "realdel"
		call realdel
	case "pinglun"
		call pinglun
	case "savepl"
		call savepl
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
	<th class=tablerow1>ӦƸְλ</th>
	<th class=tablerow1>����</th>
	<th class=tablerow1>�Ա�</th>
	<th class=tablerow1>��ϵ�绰</th>
	<th class=tablerow1>ѧ��</th>
	<th class=tablerow1>��ҵѧУ</th>
	<th class=tablerow1>��ѧרҵ</th>
	<th class=tablerow1>�����ݽ�����</th>
	<th class=tablerow1>��ǰ״̬</th>
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
	
	if LCase(Request("isdel"))="1" then
		FoundSQL = " where isdel=1 "
	else
		FoundSQL = " where isdel=0 "
	end if

	
	TotalNumber = enchiasp.Execute("Select Count(ID) from ECCMS_jobbook "& FoundSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_jobbook] "& FoundSQL &" order by riqi desc ,ID desc"
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
		Response.Write "<tr><td align=center colspan=11 class=TableRow2>��û���ҵ��κ�ӦƸ������</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td colspan=11 class=tablerow2><%Call showpage()%></td>
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
		
		Response.Write "	&jobid="
		Response.Write Rs("jobID")
		Response.Write "	'>"

		
		Response.Write Rs("jobname")
		Response.Write "	</a>"
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write "	<a href='?action=view&ID="
		Response.Write Rs("ID")
		
		Response.Write "	&jobid="
		Response.Write Rs("jobID")
		Response.Write "	'>"
	

		Response.Write Rs("name")
		Response.Write "	</a>"
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("sex")
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("telephone")
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("studydegree")
		Response.Write "	</td>"

		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("school")
		Response.Write "	</td>"
		
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("specialty")
		Response.Write "	</td>"

		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("riqi")
		Response.Write "	</td>"
		
		Response.Write "	<td align=center class=tablerow1>"
		if rs("isuse")=1 then
			Response.Write "<font color=red>¼��</font>"
		else
			Response.Write "<font color=red>δ¼��</font>"
		end if
	
		Response.Write "	</td>"


		
		

%>
	<%
	if LCase(Request("isdel"))="1" then
	%>
	
	
	<td align=center <%=strClass%>><a href='?action=realdel&ID=<%=Rs("ID")%>'>����ɾ��</a> | <a href='?action=pinglun&ID=<%=Rs("ID")%>'>����</a> | 
		<a href='?action=huifu&ID=<%=Rs("ID")%>'>�ָ�</a></td>
	<%
	else
	%>
		<td align=center <%=strClass%>><a href='?action=luyong&ID=<%=Rs("ID")%>'>¼��</a> | <a href='?action=pinglun&ID=<%=Rs("ID")%>'>����</a> | 
		<a href='?action=del&ID=<%=Rs("ID")%>' onclick="{if(confirm('��ȷ��Ҫɾ����ӦƸ��Ϣ��?')){return true;}return false;}">ɾ��</a></td>

	<%	
	end if
	%>

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
	<td colspan=11 class=tablerow2><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub Deljobbook()
	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>����Ĳ�����</li>"
		Exit Sub
	End If
	enchiasp.Execute("update [ECCMS_jobbook] set isdel=1 where ID = " & Request("ID"))
	OutHintScript("ӦƸɾ���ɹ���")
End Sub

Private Sub realdel()
	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>����Ĳ�����</li>"
		Exit Sub
	End If
	if IsSqlDataBase = 1 then
		enchiasp.Execute(" delete from [ECCMS_jobbook] where ID = " & Request("ID"))
	else
		enchiasp.Execute(" delete * from [ECCMS_jobbook] where ID = " & Request("ID"))
	end if
	OutHintScript("ӦƸ����ɾ���ɹ���")
End Sub

private sub pinglun
	response.write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>"
	response.write "<form method=Post name='myform'action='admin_jobbook.asp?action=savepl&id="& Request("ID") &"'>"
	
	response.write "<tr>"
	response.write "	<th class=tablerow1>����</th>"
	response.write "</tr>"
	response.write " <tr>"
	response.write "<td>"
	response.write "<textarea name='pinglun' cols='45' rows='6'  class='face' id='ability' style='font-size: 14px'>"& getpinglun() &"</textarea>"

	response.write " <input type='button' name='Submit4' onclick='javascript:history.go(-1)' value='������һҳ' class=Button>"
	response.write "<input name='Submit' type='submit' class='Button' value='����' class=Button>"
	response.write "</td>"
	response.write "</tr>"
	
	response.write "</form>"
	response.write "</table>"
end sub

private function getpinglun()
	
	If Trim(Request("ID")) <> "" Then
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "select * from [ECCMS_jobbook] where ID = " & Request("ID") 
		If IsSqlDataBase = 1 Then
			If CurrentPage > 100 Then
				Rs.Open SQL, Conn, 1, 1
			Else
				Set Rs = Conn.Execute(SQL)
			End If
		Else
			Rs.Open SQL, Conn, 1, 1
		End If
		if rs.eof then
			getpinglun=""
		else
			getpinglun=rs("pinglun")
		end if
		Rs.Close:Set Rs = Nothing
		
	else
		getpinglun=""
	End If
end function



private sub savepl

	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������ƸID��</li>"
		Exit Sub
	End If
	
		enchiasp.Execute("update [ECCMS_jobbook] set pinglun='"& Trim(Request.Form("pinglun")) &"' where ID = " & Request("ID"))

	
	OutHintScript("���۳ɹ���")

end sub

Private Sub luyongjobbook()
	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������ƸID��</li>"
		Exit Sub
	End If
	enchiasp.Execute("update [ECCMS_jobbook] set isuse=1 where ID = " & Request("ID"))
	OutHintScript("ӦƸ¼�óɹ���������ǰ̨��ȱְλ��")
End Sub

Private Sub huifujobbook()
	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������ƸID��</li>"
		Exit Sub
	End If
	enchiasp.Execute("update [ECCMS_jobbook] set isdel=0 where ID = " & Request("ID"))
	OutHintScript("ӦƸ�ָ��ɹ���")
End Sub


Private Sub Viewjobbook()
	If Request("ID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	dim n
	n = 1
	SQL = "select * from ECCMS_jobbook where ID=" & Request("ID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κ�ӦƸ��������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	Else
%>
<table border="0" align="center" cellpadding="2" cellspacing="1" class="TableBorder">
	<tr>
	  <th ></th>
	  <th >�鿴ӦƸ��Ϣ</th>

	</tr>

	<tr>
	  <td class="TableRow1"><strong>ӦƸְλ��</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("jobname"))%></td>
	</tr>
	
	<tr>
	  <td class="TableRow1"><strong>ӦƸ���ڣ�</strong></td><td class="TableRow1"><%=enchiasp.FormatDate(Rs("riqi"), 2)%></td>
	</tr>

	<tr>
	  <td class="TableRow1"><strong>������</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("name"))%></td>
	</tr>

	<tr>
	  <td class="TableRow1"><strong>�Ա�</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("sex"))%></td>
	</tr>

	
	<tr>
	  <td class="TableRow1"><strong>�������ڣ�</strong></td><td class="TableRow1"><%=enchiasp.FormatDate(Rs("birthday"), 2)%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>����״����</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("marry"))%></td>
	</tr>
	
	<tr>
	  <td class="TableRow1"><strong>��ҵԺУ��</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("school"))%></td>
	</tr>

	
	<tr>
	  <td class="TableRow1"><strong>ѧ����</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("studydegree"))%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>רҵ��</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("specialty"))%></td>
	</tr>

	<tr>
	  <td class="TableRow1"><strong>��ҵʱ�䣺</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("gradyear"))%></td>
	</tr>

<tr>
	  <td class="TableRow1"><strong>��ϵ�绰��</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("telephone"))%></td>
	</tr>
	
<tr>
	  <td class="TableRow1"><strong>EMAIL��</strong></td><td class="TableRow1"><a href=mailto:<%=enchiasp.ReadContent(Rs("email"))%>><%=enchiasp.ReadContent(Rs("email"))%></a><font color=red>�����ţ�</font></td>
	</tr>

<tr>
	  <td class="TableRow1"><strong>��ϵ��ַ��</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("address"))%></td>
	</tr>

<tr>
	  <td class="TableRow1"><strong>ˮƽ��������</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("ability"))%></td>
	</tr>

<tr>
	  <td class="TableRow1"><strong>���˼�����</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("resumes"))%></td>
	</tr>



	<tr>
	  <td class="TableRow2">��һӦƸ��<%=Frontjobbook(Rs("ID"))%>
	  <br>��һӦƸ��<%=Nextjobbook(Rs("ID"))%></td>
	  <td class="TableRow1"></td>
	</tr>
	<tr>
	  <td class="TableRow1"></td>
	  <td align="center" class="TableRow2"><input type="button" onclick="javascript:window.close()" value="�رձ�����" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="���ض���" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&ID=<%=Rs("ID")%>'" value="�༭ӦƸ" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 	
	Viewjob()

End Sub

Private Function Frontjobbook(jobbookID)
	Dim Rss, SQL
	SQL = "select Top 1 ID,name from ECCMS_jobbook where ID < " & jobbookID & " order by ID desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		Frontjobbook = "�Ѿ�û����"
	Else
		Frontjobbook = "<a href=admin_jobbook.asp?action=view&ID=" & Rss("ID") & ">" & Rss("name") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function Nextjobbook(jobbookID)
	Dim Rss, SQL
	SQL = "select Top 1 ID,name from ECCMS_jobbook where ID > " & jobbookID & " order by ID asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		Nextjobbook = "�Ѿ�û����"
	Else
		Nextjobbook = "<a href=admin_jobbook.asp?action=view&ID=" & Rss("ID") & ">" & Rss("name") & "</a>"
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
		Response.Write "����ӦƸ <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ƪ&nbsp;�� ҳ&nbsp;��һҳ&nbsp;|&nbsp;"
	Else
		Response.Write "����ӦƸ <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ƪ&nbsp;<a href=?page=1&ChannelID=" & Request("ChannelID") & ">�� ҳ</a>&nbsp;"
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

Private Sub Viewjob()
	If Request("jobID") = "" Then
		Exit Sub
	End If
	dim n
	n = 1
	SQL = "select * from ECCMS_job where ID=" & Request("jobID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		Response.Write "<table align=center><tr><td align=center><font color=red>û���ҵ��ø�λ����Ƹ��Ϣ��������ѡ���˴����ϵͳ������</font></td></tr></table>"
		Exit Sub
	Else
%>
<table border="0" align="center" cellpadding="2" cellspacing="1" class="TableBorder">
	<tr>
	  <th ></th>
	  <th >�ø�λ��Ӧ��Ƹ��Ϣ</th>

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
	
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub

%>