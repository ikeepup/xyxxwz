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
Dim Action,AnnounceID
Dim i,isEdit,TextContent,FoundSQL,oRs,strClass
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Action = LCase(Request("action"))
If Not ChkAdmin("Announce") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th>�������</th>
</tr>
<tr>
	<td class=tablerow2><strong>���浼����</strong> <a href='admin_announce.asp'>������ҳ</a> 
<%
	Set oRs = enchiasp.Execute("Select ChannelID,ChannelName,ChannelType From ECCMS_Channel where ChannelType < 2  Order By orders")
	Do While Not oRs.EOF
		Response.Write " | <a href='?ChannelID="
		Response.Write oRs("ChannelID")
		Response.Write "'>"
		Response.Write oRs("ChannelName")
		Response.Write "</a>"
	oRs.movenext
	Loop
	oRs.Close:Set oRs = Nothing
%>
| <a href='admin_announce.asp?action=add'><font color=blue>��������</font></a> 
	</td>
</tr>
</table>
<br>
<%
Select Case Trim(Action)
	Case "save"
		Call SaveAnnounce
	Case "modify"
		Call ModifyAnnounce
	Case "add"
		isEdit = False
		Call EditAnnounce(isEdit)
	Case "edit"
		isEdit = True
		Call EditAnnounce(isEdit)
	Case "view"
		Call ViewAnnounce
	Case "del"
		Call DelAnnounce
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
	<th class=tablerow1>�������</th>
	<th class=tablerow1>��ʾλ��</th>
	<th class=tablerow1>��������</th>
	<th class=tablerow1>����ѡ��</th>
	<th class=tablerow1>����ʱ��</th>
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
	If Request("ChannelID") <> "" Then
		FoundSQL = "where ChannelID = " & Request("ChannelID")
	Else
		FoundSQL = ""
	End If
	TotalNumber = enchiasp.Execute("Select Count(AnnounceID) from ECCMS_Announce "& FoundSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Announce] "& FoundSQL &" order by PostTime desc ,AnnounceID desc"
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
		Response.Write "<tr><td align=center colspan=5 class=TableRow2>��û���ҵ��κι��棡</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td colspan=5 class=tablerow2><%Call showpage()%></td>
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
		Response.Write "	<td " & strClass & "><a href='?action=view&AnnounceID="
		Response.Write Rs("AnnounceID")
		Response.Write "	'>"
		Response.Write Rs("title")
		Response.Write "	</a></td>"
		Response.Write "	<td align=center " & strClass & ">"

		If Rs("ChannelID") = 0 Then
			Response.Write "��ҳ����"
		ElseIf Rs("ChannelID") = 1 Then
			Response.Write "<span class=style1>����Ƶ��</span>"
		ElseIf Rs("ChannelID") = 2 Then
			Response.Write "<span class=style2>����Ƶ��</span>"
		ElseIf Rs("ChannelID") = 3 Then
			Response.Write "<span class=style3>�̳�Ƶ��</span>"
		ElseIf Rs("ChannelID") = 4 Then
			Response.Write "<span class=style2>����Ƶ��</span>"
		ElseIf Rs("ChannelID") = 5 Then
			Response.Write "<span class=style3>����Ƶ��</span>"
		ElseIf Rs("ChannelID") = 6 Then
			Response.Write "<span class=style3>��ҳ��ͼ��Ƶ��</span>"
		ElseIf Rs("ChannelID") = 7 Then
			Response.Write "<span class=style3>��ƸƵ��</span>"

		Else
			Response.Write "<span class=style1>����ҳ����ʾ</span>"
		End If
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		If Rs("AnnounceType") = 1 Then
			Response.Write "<span class=style2>���ݹ���</span>"
		ElseIf Rs("AnnounceType") = 2 Then
			Response.Write "<span class=style1>��������</span>"
		Else
			Response.Write "�б���"
		End If
%>
	</td>
	<td align=center <%=strClass%>><a href='?action=edit&AnnounceID=<%=Rs("AnnounceID")%>'>�༭</a> | 
	<a href='?action=del&AnnounceID=<%=Rs("AnnounceID")%>' onclick="{if(confirm('����ɾ���󽫲��ָܻ�����ȷ��Ҫɾ���ù�����?')){return true;}return false;}">ɾ��</a></td>
	<td align=center <%=strClass%>>
<%
		If Rs("PostTime") >= Date Then
			Response.Write "<font color=red>"
			Response.Write enchiasp.FormatDate(Rs("PostTime"), 2)
			Response.Write "</font>"
		Else
			Response.Write enchiasp.FormatDate(Rs("PostTime"), 2)
		End If
%>
	</td>
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
	<td colspan=5 class=tablerow2><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub EditAnnounce(isEdit)
	Dim EditTitle
	If isEdit Then
		SQL = "select * from ECCMS_Announce where AnnounceID=" & Request("AnnounceID")
		Set Rs = enchiasp.Execute(SQL)
		EditTitle = "�༭����"
	Else
		EditTitle = "��ӹ���"
	End If
%>
<script language=javascript>
    function CheckForm(form1)
{
	if (!validateSubmit()) return (false);
	if (form1.title.value == "")
	{
		alert("������ⲻ��Ϊ�գ�");
		form1.title.focus();
		return (false);
	}
	form1.content.value=IframeID.document.body.innerHTML; 
	MessageLength=IframeID.document.body.innerHTML.length;
	if(MessageLength<2){alert("�������ݲ���С��2���ַ���");return false;}
}
</script>
<div onkeydown=CtrlEnter()>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
  <tr>
    <th colspan="2"><%=EditTitle%></th>
  </tr>
    	<form method=Post name="myform" action="admin_announce.asp" onSubmit="return CheckForm(this);">
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""AnnounceID"" value="""& Request("AnnounceID") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
%>
  <tr>
    <td width="20%" align="right" class="TableRow2"><strong>������⣺</strong></td>
    <td width="80%" class="TableRow1"><input name="title" type="text" id="title" size="50" value='<%If isEdit Then Response.Write Rs("title")%>'> 
      <span class="style1">* </span></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>����Ƶ����</strong></td>
    <td class="TableRow1"><select name="ChannelID" id="ChannelID">
      <option value="0"<%If isEdit Then If Rs("ChannelID") = 0 Then Response.Write " selected"%>>��ҳ����</option>
<%
	Set oRs = enchiasp.Execute("Select ChannelID,ChannelName,ChannelType From ECCMS_Channel where ChannelType < 2  Order By orders")
	Do While Not oRs.EOF
		Response.Write "<option value="""& oRs("ChannelID") &""""
		If isEdit Then
			If oRs("ChannelID") = Rs("ChannelID") Then Response.Write " selected"
		End If
		Response.Write ">"
		Response.Write oRs("ChannelName")
		Response.Write "</option>"
	oRs.movenext
	Loop
	oRs.Close:Set oRs = Nothing
%>
      <option value="999"<%If isEdit Then If Rs("ChannelID") = 999 Then Response.Write " selected"%>>����Ƶ����ʾ</option>
    </select></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>�������ͣ�</strong></td>
    <td class="TableRow1"><input name="AnnounceType" type="radio" value="0"<%If isEdit Then If Rs("AnnounceType") = 0 Then Response.Write " checked" End If:Else Response.Write " checked" End If%>>
�б���
<input type="radio" name="AnnounceType" value="1"<%If isEdit Then If Rs("AnnounceType") = 1 Then Response.Write " checked"%>>
���ݹ���</td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>�������ݣ�</strong></td>
    <td class="TableRow1"><textarea name="content" style="display:none" id="content"><%If isEdit Then Response.Write Server.HTMLEncode(Rs("content"))%></textarea>
    <script src='../editor/edit.js' type=text/javascript></script></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>�ϴ��ļ���</strong></td>
    <td class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upfiles.asp></iframe></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>�����ˣ�</strong></td>
    <td class="TableRow1"><input name="writer" type="text" id="writer" size="15" value='<%If isEdit Then Response.Write Rs("title") Else Response.Write AdminName End If%>'> 
      <span class="style1">* </span> 
      <%If isEdit Then%>
      <input name="update" type="checkbox" id="update" value="yes">
���¹���ʱ�� 
<%End If%></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2">��</td>
    <td align="center" class="TableRow1">
    <input type="button" name="Submit2" onclick="CheckLength();" value="�鿴���ݳ���" class=Button>
    <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>
    <input name="Submit1" type="submit" class="Button" value="���湫��" class=Button></td>
  </tr></form>
  <tr>
    <td colspan="2" class="TableRow1"><strong>˵����</strong><br>
      &nbsp;&nbsp;&nbsp;&nbsp;����Ƶ�� ---- ֻ��Ƶ����ҳ����ʾ���棻<br>
      &nbsp;&nbsp;&nbsp;&nbsp;�������� ---- �б���ʾ��ָ�������б����ʽ��ʾ���棬��Ҫ�û�����ſ��Կ����������ݣ����ݹ�����ָ���������ݵķ�ʽ��ʾ�����ڵ�Ƶ����ҳ��ע��ֻ��ʾ���µ�һ�����档</td>
  </tr>
</table>
</div>
<%
	If isEdit Then Rs.Close:Set Rs = Nothing
End Sub

Private Sub CheckSave()
	If Trim(Request.Form("title")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������ⲻ��Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("ChannelID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Ƶ��ID����Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("AnnounceType")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������Ͳ���Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("content")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������ݲ���Ϊ�գ�</li>"
	End If	
	TextContent = ""
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next
	If Len(TextContent) < 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���淢���˲���Ϊ�գ�</li>"
	End If
End Sub
Private Sub SaveAnnounce()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Announce where (AnnounceID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = Trim(Request.Form("ChannelID"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("AnnounceType") = Request.Form("AnnounceType")
		Rs("Content") = TextContent
		Rs("writer") = enchiasp.ChkFormStr(Request.Form("writer"))
		Rs("PostTime") = Now()
		Rs("hits") = 0
	Rs.update
	Rs.Close
	Rs.Open "select top 1 AnnounceID from ECCMS_Announce order by AnnounceID desc", Conn, 1, 1
	AnnounceID = Rs("AnnounceID")
	Rs.Close:Set Rs = Nothing
	Succeed("<li>��ϲ��������µĹ���ɹ���</li><li><a href=?action=view&AnnounceID=" & AnnounceID & ">����˴��鿴�ù���</a></li>")
End Sub
Private Sub ModifyAnnounce()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Announce where AnnounceID = " & Request("AnnounceID")
	Rs.Open SQL,Conn,1,3
		Rs("ChannelID") = Trim(Request.Form("ChannelID"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("AnnounceType") = Request.Form("AnnounceType")
		Rs("Content") = TextContent
		Rs("writer") = enchiasp.ChkFormStr(Request.Form("writer"))
		If LCase(Request.Form("Update")) = "yes" Then Rs("PostTime") = Now()
	Rs.update
		AnnounceID = Rs("AnnounceID")
	Rs.Close:Set Rs = Nothing
	Succeed("<li>��ϲ�����޸Ĺ���ɹ���</li><li><a href=?action=view&AnnounceID=" & AnnounceID & ">����˴��鿴�ù���</a></li>")
End Sub
Private Sub DelAnnounce()
	If Trim(Request("AnnounceID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����빫��ID��</li>"
		Exit Sub
	End If
	enchiasp.Execute("Delete From ECCMS_Announce where AnnounceID = " & Request("AnnounceID"))
	OutHintScript("����ɾ���ɹ���")
End Sub

Private Sub ViewAnnounce()
	If Request("AnnounceID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	dim n
	n = 1
	enchiasp.Execute ("update ECCMS_Announce set hits = hits + "&n&" where AnnounceID=" & Request("AnnounceID"))
	SQL = "select * from ECCMS_Announce where AnnounceID=" & Request("AnnounceID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κι��档������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	Else
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
	<tr>
	  <th colspan="2">�鿴����</th>
	</tr>
	<tr>
	  <td align="center" class="TableRow2" colspan="2"><a href=?action=edit&AnnounceID=<%=Rs("AnnounceID")%>><font size=4><%=Rs("title")%></font></a></td>
	</tr>
	<tr>
	  <td align="center" class="TableRow1"><strong>����ʱ�䣺</strong> <%=Rs("PostTime")%> &nbsp;&nbsp;
	  <strong>�� �� �ˣ�</strong> <%=Rs("writer")%> &nbsp;&nbsp;<strong>���������</strong> <%=Rs("hits")%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>�������ݣ�</strong><br><%=enchiasp.ReadContent(Rs("content"))%></td>
	</tr>
	<tr>
	  <td class="TableRow2">��һ���棺<%=FrontAnnounce(Rs("AnnounceID"))%>
	  <br>��һ���棺<%=NextAnnounce(Rs("AnnounceID"))%></td>
	</tr>
	<tr>
	  <td align="center" colspan="2" class="TableRow2"><input type="button" onclick="javascript:window.close()" value="�رձ�����" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="���ض���" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&AnnounceID=<%=Rs("AnnounceID")%>'" value="�༭����" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub

Private Function FrontAnnounce(AnnounceID)
	Dim Rss, SQL
	SQL = "select Top 1 AnnounceID,title from ECCMS_Announce where AnnounceID < " & AnnounceID & " order by AnnounceID desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		FrontAnnounce = "�Ѿ�û����"
	Else
		FrontAnnounce = "<a href=admin_Announce.asp?action=view&AnnounceID=" & Rss("AnnounceID") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function NextAnnounce(AnnounceID)
	Dim Rss, SQL
	SQL = "select Top 1 AnnounceID,title from ECCMS_Announce where AnnounceID > " & AnnounceID & " order by AnnounceID asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		NextAnnounce = "�Ѿ�û����"
	Else
		NextAnnounce = "<a href=admin_Announce.asp?action=view&AnnounceID=" & Rss("AnnounceID") & ">" & Rss("title") & "</a>"
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
		Response.Write "���й��� <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ƪ&nbsp;�� ҳ&nbsp;��һҳ&nbsp;|&nbsp;"
	Else
		Response.Write "���й��� <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ƪ&nbsp;<a href=?page=1&ChannelID=" & Request("ChannelID") & ">�� ҳ</a>&nbsp;"
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
Public Sub CreateAnnounce()
	Dim rsAnnounce,sqlAnnounce
	sqlAnnounce = "select A.AnnounceID,A.title,A.PostTime,A.AnnounceType,C.ChannelDir,C.ChannelUrl,C.BindDomain,C.DomainName from [ECCMS_Announce] A inner join [ECCMS_Channel] C On A.ChannelID=C.ChannelID where A.ChannelID=" & ChannelID & ""
	Set rsAnnounce = enchiasp.Execute(sqlAnnounce)
End Sub

%>