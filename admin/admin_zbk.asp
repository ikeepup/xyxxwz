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
	<th>�ʱ�������</th>
	
</tr>
<tr>
	<td class=tablerow2><strong>������</strong> <a href='admin_zbk.asp'>�ʱ�����ҳ</a> 
| <a href='admin_zbk.asp?action=add'><font color=blue>�Ǽ��ʱ���</font></a> 
	</td>
</tr>
</table>
<br>
<%
'Ȩ���ж�
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
	<th class=tablerow1>xh</th>
	<th class=tablerow1>�ʱ������</th>
	<th class=tablerow1>���ƺ�</th>
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
	FoundSQL = " "
	TotalNumber = enchiasp.Execute("Select Count(xh) from ECCMS_zb "& FoundSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_zb] "& FoundSQL &" order by xh desc"

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
		Response.Write "<tr><td align=center colspan=4 class=TableRow2>��û���ҵ��κ���Ϣ��</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td colspan=4 class=tablerow2><%Call showpage()%></td>
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
		Response.Write "	<td " & strClass & ">"
		Response.Write Rs("xh")
		Response.Write "	</td>"
		Response.Write "	<td align=center " & strClass & ">"
		Response.Write Rs("bh")
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("cph")
		Response.Write "	</td>"
			

%>
	
	<td align=center <%=strClass%>><a href='?action=edit&xh=<%=Rs("xh")%>'>�༭</a> | 
	<a href='?action=del&xh=<%=Rs("xh")%>' onclick="{if(confirm('��Ϣɾ���󽫲��ָܻ�����ȷ��Ҫɾ������Ϣ��?')){return true;}return false;}">ɾ��</a></td>
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
	<td colspan=4 class=tablerow2><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub Editjob(isEdit)

	Dim EditTitle
	If isEdit Then
		SQL = "select * from ECCMS_zb where xh=" & Request("xh")
		Set Rs = enchiasp.Execute(SQL)
		EditTitle = "�༭"
	Else
		EditTitle = "���"
	End If
%>
<script language=javascript>
    function CheckForm(form1)
{
	if (!validateSubmit()) return (false);
	form1.content.value=IframeID.document.body.innerHTML; 
	MessageLength=IframeID.document.body.innerHTML.length;
}
</script>
<div onkeydown=CtrlEnter()>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
  <tr>
    <th colspan="2"><%=EditTitle%></th>
  </tr>
    	<form method=Post name="myform" action="admin_zbk.asp" onSubmit="return CheckForm(this);">
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""ID"" value="""& Request("ID") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
%>
  <tr>
    <td width="20%" align="right" class="TableRow2"><strong>�ʱ�����ţ�</strong></td>
    <td width="80%" class="TableRow1"><input name="bh" type="text" id="bh" size="50" value='<%If isEdit Then Response.Write Rs("bh")%>'> 
      <span class="style1">* </span></td>
  </tr>
  
	<tr>
    <td width="20%" align="right" class="TableRow2"><strong>���ƺţ�</strong></td>
    <td width="80%" class="TableRow1"><input name="cph" type="text" id="cph" size="50" value='<%If isEdit Then Response.Write Rs("cph") else response.write "��M" end if%>'> 
    </td>
  </tr>

  
  <tr>
    <td align="right" class="TableRow2">��</td>
    <td align="center" class="TableRow1">
    <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>
    <input name="Submit1" type="submit" class="Button" value="����" class=Button></td>
  </tr></form>
</table>
</div>
<%
	If isEdit Then Rs.Close:Set Rs = Nothing
End Sub

Private Sub CheckSave()
	If Trim(Request.Form("bh")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�ʱ�����Ų���Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("cph")) = "" or Trim(Request.Form("cph")) = "��M"  Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���ƺŲ���Ϊ�գ�</li>"
	End If
End Sub
Private Sub Savejob()
	

	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_zb where (xh is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("bh") = Trim(Request.Form("bh"))
		Rs("cph") = enchiasp.ChkFormStr(Request.Form("cph"))
		
	Rs.update
	Rs.Close:Set Rs = Nothing
	Succeed("<li>��ϲ��������µ���Ϣ�ɹ���</li>")
End Sub
Private Sub Modifyjob()
	

	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_zb where xh= " & Request("xh")
	Rs.Open SQL,Conn,1,3
		Rs("bh") = Trim(Request.Form("bh"))
		Rs("cph") = enchiasp.ChkFormStr(Request.Form("cph"))
	Rs.update
	Rs.Close:Set Rs = Nothing
	Succeed("<li>��ϲ�����޸ĳɹ���</li>")
End Sub
Private Sub Deljob()
	
	If Trim(Request("XH")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������XH��</li>"
		Exit Sub
	End If
	enchiasp.Execute("delete from [ECCMS_zb] where xh= " & Request("xh"))
	OutHintScript("ɾ���ɹ���")
End Sub




Private Sub showpage()
	Dim n
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action=?ChannelID=" & Request("ChannelID") & "><tr><td align=center> " & vbCrLf
	If CurrentPage < 2 Then
		Response.Write "���� <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ƪ&nbsp;�� ҳ&nbsp;��һҳ&nbsp;|&nbsp;"
	Else
		Response.Write "���� <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ƪ&nbsp;<a href=?page=1&ChannelID=" & Request("ChannelID") & ">�� ҳ</a>&nbsp;"
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