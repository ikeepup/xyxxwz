<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
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
Call InnerLocation("�ҷ��������")

Dim Action,SQL,Rs,i
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID < 2 Then ChannelID = 2
ChannelID = CLng(ChannelID)

Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveSoft
	Case "edit"
		Call EditSoft
	Case "del"
		Call DeleteSoft
	Case "view"
		Call SoftView
	Case Else
		Call showmain
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If

Sub showmain()
	If Founderr = True Then Exit Sub
%>
<script language="JavaScript">
<!--
function myuser_softlist_top(accept){
	//document.write ('<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>');
	document.write ('<th valign=middle>');
	if (accept==1)
	{
		document.write ('�ҵ�����б�--����˵����');
	}else{
		document.write ('�ҵ�����б�--δ��˵����');
	}
	document.write ('</th>');
	document.write ('<th valign=middle noWrap>���</th>');
	document.write ('<th valign=middle noWrap>��������</th>');
	document.write ('<th valign=middle noWrap>�������</th>');
	document.write ('</tr>');
}
function myuser_softlist_not(){
	document.write ('<tr>');
	document.write ('<td class=Usertablerow1 align=center valign=middle colspan=5>û���ҵ��κ������</td>');
	document.write ('</tr>');
}
function myuser_softlist_loop(channelid,softid,accept,softname,classname,softdate,hits,style){
	var tablebody;
	if (style==1)
	{
		tablebody="Usertablerow1";
	}else{
		tablebody="Usertablerow2";
	}
	document.write ('<tr>');
	document.write ('<td class='+tablebody+' valign=middle>['+classname+'] ');
	document.write ('<a href="softlist.asp?action=view&channelid='+channelid+'&softid='+softid+'">'+softname+'</a></td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	if (accept==1)
	{
		document.write ('<font color=blue>����</font>');
	}else{
		document.write ('<font color=red>δ��</font>');
	}
	document.write ('</td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>'+softdate+'</td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	document.write ('<a href="softlist.asp?action=edit&channelid='+channelid+'&softid='+softid+'">�޸�</a> | ');
	document.write ('<a href="softlist.asp?action=del&channelid='+channelid+'&softid='+softid+'" onClick="return confirm(\'ȷ��Ҫɾ���������������\')">ɾ��</a>');
	document.write ('</td>');
	document.write ('</tr>');
}
-->
</script>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr height=20 align=center>
		<td class=Usertablerow2 colspan=5><a href="?ChannelID=<%=ChannelID%>&Accept=1">����˵����</a> | 
		<a href="?ChannelID=<%=ChannelID%>">δ��˵����</a> | 
		<a href="softpost.asp?ChannelID=<%=ChannelID%>">�����µ����</a> </td>
	</tr>
<%
	Dim CurrentPage,page_count,totalrec,Pcount,maxperpage
	Dim isAccept,s
	maxperpage = 20 '###ÿҳ��ʾ��
	
	If Trim(Request("Accept")) <> "" Then
		isAccept = 1
	Else
		isAccept = 0
	End If
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CInt(CurrentPage)
	End If
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	Response.Write "<script>myuser_softlist_top("& isAccept &")</script>" & vbNewLine
	totalrec = enchiasp.Execute("SELECT COUNT(SoftID) FROM ECCMS_SoftList WHERE ChannelID = " & ChannelID & " And username='" & enchiasp.MemberName & "' And isAccept="& isAccept)(0)
	Pcount = CInt(totalrec / maxperpage)  '�õ���ҳ��
	If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT A.SoftID,A.SoftName,A.SoftVer,A.SoftTime,A.AllHits,A.isAccept,C.ClassName FROM [ECCMS_SoftList] A INNER JOIN [ECCMS_Classify] C on A.ClassID=C.ClassID WHERE A.ChannelID = " & ChannelID & "  And A.username='" & enchiasp.MemberName & "' And isAccept="& isAccept &" ORDER BY A.isTop DESC, A.SoftTime DESC ,A.SoftID DESC"  'And username='" & enchiasp.MemberName & "'
	Rs.Open SQL, Conn, 1, 1
	If Rs.EOF And Rs.BOF Then
		Response.Write "<script>myuser_softlist_not()</script>" & vbNewLine
	Else
		If Pcount > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		page_count = 0
		If Rs.EOf Then Exit Sub
		SQL = Rs.GetRows(maxperpage)
		For i=0 To Ubound(SQL,2)
			If (i mod 2) = 0 Then
				s = 2
			Else
				s = 1
			End If
			Response.Write VbCrLf
			Response.Write "<script>myuser_softlist_loop("
			Response.Write ChannelID
			Response.Write ","
			Response.Write SQL(0,i)
			Response.Write ","
			Response.Write SQL(5,i)
			Response.Write ",'"
			Response.Write EncodeJS(SQL(1,i) &" "& SQL(2,i))
			Response.Write "','"
			Response.Write EncodeJS(SQL(6,i))
			Response.Write "','"
			Response.Write FormatDated(SQL(3,i),4)
			Response.Write "',"
			Response.Write SQL(4,i)
			Response.Write ","
			Response.Write s
			Response.Write ")</script>"
			Response.Write VbCrLf
			page_count = page_count + 1
		Next
		SQL=Null
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "<tr align=right><td class=Usertablerow2 colspan=5>"
	Response.Write ShowPages(CurrentPage,Pcount,totalrec,maxperpage,"&ChannelID="& ChannelID &"&Accept="& Request("Accept"))
	Response.Write "</td></tr>" & vbNewLine
	Response.Write "</table>"

End Sub

Sub DeleteSoft()
	If CInt(GroupSetting(12)) = 0 Then
		ErrMsg = ErrMsg + "<li>�Բ�����û��ɾ�������Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
		Founderr = True
		Exit Sub
	End If
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Trim(Request("SoftID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	SQL = "SELECT isAccept FROM ECCMS_SoftList WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And isAccept=0 And SoftID=" & CLng(Request("SoftID"))
	Set Rs = enchiasp.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��������Ѿ�ͨ�����,��û��Ȩ��ɾ��,����ʲô��������ϵ����Ա��</li>"
		Set Rs = Nothing
		Exit Sub
	Else
		enchiasp.Execute("DELETE FROM ECCMS_SoftList WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And isAccept=0 And SoftID=" & CLng(Request("SoftID")))
	End If
	Set Rs = Nothing
	Response.Redirect  Request.ServerVariables("HTTP_REFERER")
End Sub
Function FormatDownAddress(ByVal str)
	If Trim(str) = ""  Or Trim(str) = "|||" Then
		FormatDownAddress = ""
		Exit Function
	End If
	Dim strDownAddress,sDownAddress,sDownSiteName
	Dim n,AddressNum,strAddress,strDownName,strTemp
	On Error Resume Next
	strDownAddress = Split(str, "|||")
	sDownAddress = Split(strDownAddress(1), "|")
	sDownSiteName = Split(strDownAddress(0), "|")
	If UBound(sDownAddress) < UBound(sDownSiteName) Then
		AddressNum = UBound(sDownAddress)
	Else
		AddressNum = UBound(sDownSiteName)
	End If
	strAddress = ""
	strDownName = ""
	For n = 0 To CInt(AddressNum)
		If Trim(sDownAddress(n)) <> "" And Trim(sDownSiteName(n)) <> "" Then
			strAddress = strAddress & Trim(sDownAddress(n)) & "|"
			strDownName = strDownName & Trim(sDownSiteName(n)) & "|"
		End If
	Next
	If Len(strDownName) > 0 Then strDownName = Left(strDownName, Len(strDownName) - 1)
	If Len(strAddress) > 0 Then strAddress = Left(strAddress, Len(strAddress) - 1)
	strTemp = strDownName & "|||" & strAddress
	FormatDownAddress = Trim(strTemp)
End Function
Sub SaveSoft()
	Dim TextContent,isAccept,ForbidEssay,SoftID
	If CInt(GroupSetting(12)) = 0 Then
		ErrMsg = ErrMsg + "<li>�Բ�����û���޸������Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
		Founderr = True
		Exit Sub
	End If
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If CInt(GroupSetting(2)) = 1 Then
		If Not enchiasp.CodeIsTrue() Then
			ErrMsg = ErrMsg + "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&"""><li>��֤��У��ʧ�ܣ��뷵��ˢ��ҳ�����ԡ�������Զ�����</li>"
			Founderr = True
			Exit Sub
		End If
		Session("GetCode") = ""
	End If
	If Trim(Request.Form("SoftName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������Ʋ���Ϊ�գ�</li>"
	End If
	If Len(Request.Form("SoftName")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������Ʋ��ܳ���200���ַ���</li>"
	End If
	If Len(Request.Form("Related")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���������ܳ���200���ַ���</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>����Ǽ�����Ϊ�ա�</li>"
	End If
	If CLng(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�÷������ⲿ���ӣ�������������</li>"
	End If
	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��һ�������Ѿ����������࣬������������</li>"
	End If
	If Trim(Request.Form("SoftType")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ��������ͣ�</li>"
	End If
	If Trim(Request.Form("impower")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ�������Ȩ��ʽ��</li>"
	End If
	If Trim(Request.Form("Languages")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ��������ԣ�</li>"
	End If
	If Trim(Request.Form("content1")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����鲻��Ϊ�գ�</li>"
	End If
	TextContent = ""
	For i = 1 To Request.Form("content1").Count
		TextContent = TextContent & Request.Form("content1")(i)
	Next
	If Len(Request.Form("RunSystem")) = 0 Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>���л�������Ϊ�գ�</li>"
	End If
	If Not IsNumeric(Request.Form("SoftSize")) Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>�����С������������</li>"
	End If
	If CInt(Request("isAccept")) = 1 Then
		isAccept = 1
	Else
		isAccept = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If
	'---- ��ȡ���ص�ַ���е�����
	Dim TempAddress,TempSiteName,TempDownAddress
	Dim strTempAddress,strTempSiteName,DownAddress
	If Trim(Request.Form("DownAddress")) <> "" And Trim(Request.Form("SiteName")) <> "" Then
		strTempAddress = ""
		For Each TempAddress In Request.Form("DownAddress")
			If LCase(TempAddress) <> "del" And Trim(TempAddress) <> "" Then
				strTempAddress = strTempAddress & Replace(TempAddress, "|", "") & "|"
			End If
		Next
		If Len(strTempAddress) > 0 Then strTempAddress = Left(strTempAddress, Len(strTempAddress) - 1)
		strTempSiteName = ""
		For Each TempSiteName In Request.Form("SiteName")
			If LCase(TempSiteName) <> "del" And Trim(TempSiteName) <> "" Then
				strTempSiteName = strTempSiteName & Replace(TempSiteName, "|", "") & "|"
			End If
		Next
		If Len(strTempSiteName) > 0 Then strTempSiteName = Left(strTempSiteName, Len(strTempSiteName) - 1)
		TempDownAddress = enchiasp.CheckStr(strTempSiteName &"|||"& strTempAddress)
	Else
		TempDownAddress = ""
	End If
	DownAddress = FormatDownAddress(TempDownAddress)
	If Founderr = True Then Exit Sub
	On Error Resume Next
	Call PreventRefresh  '��ˢ��
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_SoftList WHERE username='" & enchiasp.MemberName & "' And SoftID=" & CLng(Request("SoftID"))
	Rs.Open SQL,Conn,1,3
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = enchiasp.ChkNumeric(Request.Form("ClassID"))
		Rs("SoftName") = enchiasp.ChkFormStr(Request.Form("SoftName"))
		Rs("SoftVer") = enchiasp.ChkFormStr(Request.Form("SoftVer"))
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Content") = Html2Ubb(TextContent)
		Rs("Languages") = enchiasp.ChkFormStr(Request.Form("Languages"))
		Rs("SoftType") = enchiasp.ChkFormStr(Request.Form("SoftType"))
		Rs("RunSystem") = enchiasp.ChkFormStr(Request.Form("RunSystem"))
		Rs("impower") = enchiasp.ChkFormStr(Request.Form("impower"))
		If UCase(Request.Form("SizeUnit")) = "MB" Then
			Rs("SoftSize") = enchiasp.CheckNumeric(Request.Form("SoftSize") * 1024)
		Else
			Rs("SoftSize") = enchiasp.CheckNumeric(Request.Form("SoftSize"))
		End If
		Rs("star") = enchiasp.ChkNumeric(Request.Form("star"))
		Rs("Homepage") = enchiasp.ChkFormStr(Request.Form("Homepage"))
		Rs("Contact") = enchiasp.ChkFormStr(Request.Form("Contact"))
		Rs("Author") = enchiasp.ChkFormStr(Request.Form("Author"))
		Rs("Regsite") = enchiasp.ChkFormStr(Request.Form("Regsite"))
		Rs("username") = Trim(enchiasp.MemberName)
		Rs("SoftPrice") = enchiasp.CheckNumeric(Request.Form("SoftPrice"))
		Rs("SoftImage") = enchiasp.ChkFormStr(Request.Form("SoftImage"))
		Rs("Decode") = enchiasp.ChkFormStr(Request.Form("Decode"))
		Rs("DownAddress") = enchiasp.ChkFormStr(DownAddress)
		Rs("isAccept") = isAccept
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	SoftID = Rs("SoftID")
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>��ϲ�����޸�����ɹ���</li><li><a href=?action=view&ChannelID=" & ChannelID & "&SoftID=" & SoftID & ">����˴��鿴�����</a></li>")
End Sub
Function EncodeJS(str)
	str = enchiasp.HtmlEncode(str)
	str = Replace(Replace(Replace(Replace(str,"\","\\"),"'","\'"),VbCrLf,"\n"),Chr(13),"")
	EnCodeJs = str
End Function
Function ShowDownAddress(ByVal Address)
	Dim strDownAddress,sDownAddress,sDownSiteName
	Dim n,strTemp,AddressNum,strAddress,strDownName
	If Not (enchiasp.CheckNull(Address)) Or Trim(Address) = "|||" Then
		ShowDownAddress = "���ص�ַ1|���ص�ַ2|���ص�ַ3|||del|del|del"
		Exit Function
	End If
	On Error Resume Next
	strDownAddress = Split(Address, "|||")
	sDownAddress = Split(strDownAddress(1), "|")
	sDownSiteName = Split(strDownAddress(0), "|")
	If UBound(sDownAddress) < UBound(sDownAddress) Then
		AddressNum = UBound(sDownAddress)
	Else
		AddressNum = UBound(sDownSiteName)
	End If
	For n = 0 To AddressNum
		strAddress = strAddress & sDownAddress(n) & "|"
		strDownName = strDownName & sDownSiteName(n) & "|"
	Next
	strTemp = strDownName &"del|del|del|||"& strAddress &"del|del|del"
	ShowDownAddress = Split(strTemp, "|||")
End Function
Private Sub SoftView()
	If Request("SoftID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	If ChannelID = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ָ��Ƶ����</li>"
		Exit Sub
	End If
	SQL = "SELECT * FROM ECCMS_SoftList WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And SoftID=" & CLng(Request("SoftID"))
	Set Rs = enchiasp.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>û���ҵ��κ������������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	Else
	Dim strDownAddress,sDownAddress
	strDownAddress = Split(Rs("DownAddress"), "|||")
	sDownAddress = Split(strDownAddress(1), "|")
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="UserTableBorder" style="table-layout:fixed;word-break:break-all">
	<tr>
	  <th colspan="2">&gt;&gt;�鿴�����Ϣ&lt;&lt;</th>
	</tr>
	<tr>
	  <td align="center" class="UserTableRow2" colspan="2"><font size=3 color=blue><a href="?action=edit&ChannelID=<%=ChannelID%>&softid=<%=Rs("SoftID")%>"><%=Rs("SoftName")%>&nbsp;<%=Rs("SoftVer")%></a></font></td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>������л�����</strong> <%=Rs("RunSystem")%></td>
	  <td class="UserTableRow1"><strong>������ͣ�</strong> <%=Rs("SoftType")%></td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>�����С��</strong> <%=Rs("SoftSize")%></td>
	  <td class="UserTableRow1"><strong>����Ǽ���</strong> 
<%
	Response.Write "<font color=red>"
	For i = 1 to Rs("star")
		Response.Write "��"
	Next
	Response.Write "</font>"
%>
	  </td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>������ԣ�</strong> <%=Rs("Languages")%></td>
	  <td class="UserTableRow1"><strong>��Ȩ��ʽ��</strong> <%=Rs("impower")%></td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>����ʱ�䣺</strong> <%=Rs("SoftTime")%></td>
	  <td class="UserTableRow1"><strong>������ҳ��</strong> <%=Rs("Homepage")%></td>
	</tr>
	<tr>
	  <td colspan="2" class="UserTableRow1"><strong>�����飺</strong><br><%=UBBCode(Rs("content"))%></td>
	</tr>
	<tr>
	  <td colspan="2" class="UserTableRow1"><strong>���ص�ַ��</strong><br>
<%
	For i = 0 To UBound(sDownAddress)
		Response.Write "<li><a href=""" & sDownAddress(i) & """ target=_blank>" & sDownAddress(i) & "</a></li>" & vbNewLine
	Next

%>
	  </td>
	</tr>
	<tr>
	  <td align="center" colspan="2" class="UserTableRow2"><input type="button" onclick="javascript:window.close()" value="�رձ�����" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="���ض���" class=button>&nbsp;&nbsp;
	  </td>
	</tr>
</table>
<%

	End If
	Rs.Close
	Set Rs = Nothing 
End Sub

Sub EditSoft()
	If Request("SoftID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	If CInt(GroupSetting(12)) = 0 Then
		ErrMsg = ErrMsg + "<li>�Բ�����û���޸������Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
		Founderr = True
		Exit Sub
	End If
	If ChannelID = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ָ��Ƶ����</li>"
		Exit Sub
	End If
	If Founderr = True Then Exit Sub
	SQL = "SELECT * FROM ECCMS_SoftList WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And SoftID=" & CLng(Request("SoftID"))
	Set Rs = enchiasp.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κ������������ѡ���˴����ϵͳ������</li>"
		Set Rs = Nothing 
		Exit Sub
	End If
	Dim Channel_Setting,ClassID,DownAddress,DownSiteName,TempAddress
	Channel_Setting = Split(enchiasp.Channel_Setting, "|||")
	ClassID = Rs("ClassID")
	If Rs("isAccept") <> 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������Ѿ�ͨ�����,��û��Ȩ���޸�,����ʲô��������ϵ����Ա��</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	DownSiteName = Split(ShowDownAddress(Rs("DownAddress"))(0), "|")
	DownAddress = Split(ShowDownAddress(Rs("DownAddress"))(1), "|")
%>
<script language=JavaScript>
var _maxCount = '<%=CLng(GroupSetting(17))%>';
function ToRunsystem(addTitle) {
	var revisedTitle;
	var currentTitle;
	currentTitle = document.myform.RunSystem.value;
	revisedTitle = currentTitle+addTitle;
	document.myform.RunSystem.value=revisedTitle;
	document.myform.RunSystem.focus();
	return; 
}
function doSubmit(){
	if (document.myform.SoftName.value==""){
		alert("������Ʋ���Ϊ�գ�");
		return false;
	}
	if (document.myform.DownAddress.value==""){
		alert("������Ҫ��дһ�����ص�ַ�ɣ�");
		return false;
	}
	if (document.myform.SiteName.value==""){
		alert("�������Ʋ���Ϊ�գ�");
		return false;
	}
	if (document.myform.ClassID.value==""){
		alert("��һ�������Ѿ����������࣬��ѡ�����������࣡");
		return false;
	}
	if (document.myform.ClassID.value=="0"){
		alert("�÷������ⲿ���ӣ�����������ݣ�");
		return false;
	}
	if (document.myform.RunSystem.value==""){
		alert("������л�������Ϊ�գ�");
		return false;
	}
	if (document.myform.SoftType.value==""){
		alert("������Ͳ���Ϊ�գ�");
		return false;
	}
	if (document.myform.SoftSize.value==""){
		alert("�����С��û����д��");
		return false;
	}
	<%If CInt(GroupSetting(2)) = 1 Then%>
	if (document.myform.codestr.value==""){
		alert("����д��֤�룡");
		return false;
	}
	<%End If%>
	myform.content1.value = getHTML(); 
	MessageLength = Composition.document.body.innerHTML.length;
	if(MessageLength < 2){
		alert("�����鲻��С��2���ַ���");
		return false;
	}
	if(MessageLength > _maxCount){
		alert("�����鲻�ܳ���"+_maxCount+"���ַ���");
		return false;
	}
	document.myform.Submit1.disabled = true;
	document.myform.submit();
}
</script>
<div onkeydown=CtrlEnter()>
<table  border="0" align="center" cellpadding="3" cellspacing="1" class="UserTableBorder">
        <tr>
          <th colspan="4">&gt;&gt;�������&lt;&lt;</th>
        </tr>
	<form method=Post name="myform" action="softlist.asp">
	<input type="Hidden" name="action" value="save">
	<input type=hidden name=ChannelID value="<%=ChannelID%>">
	<input type=hidden name=SoftID value="<%=Rs("SoftID")%>">
        <tr>
          <td width="15%" align="right" nowrap class="UserTableRow2"><strong>��������</strong></td>
          <td width="85%" class="UserTableRow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	sClassSelect = Replace(sClassSelect, "{ClassID=" & ClassID & "}", "selected")
	Response.Write sClassSelect
	Response.Write "</select>"
%>
	  </td>
        </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>�������</strong></td>
          <td class="UserTableRow1"><input name="SoftName" type="text" id="SoftName" size="45" value="<%=Rs("SoftName")%>"> 
          <font color=red>*</font> <strong>����汾</strong><input name="SoftVer" type="text" id="SoftVer" size="20" value="<%=Rs("SoftVer")%>"></td>
	  </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>������</strong></td>
          <td class="UserTableRow1"><input name="Related" type="text" id="Related" size="60" value="<%=Rs("Related")%>"> <font color=red>*</font></td>
        </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>���л���</strong></td>
          <td class="UserTableRow1"><input name="RunSystem" type="text" size="60" value="<%=Rs("RunSystem")%>"><br>
<%
	Dim RunSystem
	RunSystem = Split(Channel_Setting(0), "|")
	For i = 0 To UBound(RunSystem)
		Response.Write "<a href='javascript:ToRunsystem(""" & Trim(RunSystem(i)) & """)'><u>" & Trim(RunSystem(i)) & "</u></a>  "
		If i = 10 Then Response.Write "<br>"
	Next
%>
          </td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�������</strong></td>
          <td class="UserTableRow1">
<%
	Dim SoftType
	SoftType = Split(Channel_Setting(2), ",")
	For i = 0 To UBound(SoftType)
		Response.Write "<input type=""radio"" name=""SoftType"" value=""" & Trim(SoftType(i)) & """ "
		If SoftType(i) = Rs("SoftType") Then Response.Write " checked"
		Response.Write ">" & Trim(SoftType(i)) & " "
		If i = 6 Then Response.Write "<br>"
	Next
%>
	  </td>
        </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>Ԥ��ͼƬ</strong></td>
          <td class="UserTableRow1"><input name="SoftImage" id="ImageUrl" type="text" size="60" value="<%=Rs("SoftImage")%>"></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�ϴ�ͼƬ</strong></td>
          <td class="UserTableRow1"><iframe name="image" frameborder=0 width='100%' height=55 scrolling=no src=upload.asp?ChannelID=2></iframe></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�����С</strong></td>
          <td class="UserTableRow1">
	<input type="text" name="SoftSize" size="14" onkeyup="if(isNaN(this.value))this.value=''" value='<%=Rs("SoftSize")%>'> <input name="SizeUnit" type="radio" value="KB" checked> KB <input type="radio" name="SizeUnit" value="MB"> MB <font color="#FF0000">��</font>
	<strong>��ѹ����</strong>
	<input type="text" name="Decode" size="15" maxlength="100" value='<%=Rs("Decode")%>'> <font color="#808080">û��������</font>
          </td>
	</tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�������</strong></td>
          <td class="UserTableRow1">
<%
	Response.Write " <select name=""impower"">"
	Response.Write "<option value=""" & Rs("impower") & """>" & Rs("impower") & "</option>"
	Dim ImpowerStr
	ImpowerStr = Split(Channel_Setting(3), ",")
	For i = 0 To UBound(ImpowerStr)
		Response.Write " <option value=""" & ImpowerStr(i) & """>" & ImpowerStr(i) & "</option>"
	Next
	Response.Write " </select>&nbsp;&nbsp;"
	Response.Write " <select name=""Languages"">"
	Response.Write "<option value=""" & Rs("Languages") & """>" & Rs("Languages") & "</option>"
	Dim LanguagesStr
	LanguagesStr = Split(Channel_Setting(4), ",")
	For i = 0 To UBound(LanguagesStr)
		Response.Write " <option value=""" & LanguagesStr(i) & """>" & LanguagesStr(i) & "</option>"
	Next
	Response.Write " </select>&nbsp;&nbsp;"
%>
		<select name="star">
		<option value=5<%If Rs("star") = 5 Then Response.Write (" selected")%>>������</option>
          	<option value=4<%If Rs("star") = 4 Then Response.Write (" selected")%>>�����</option>
          	<option value=3<%If Rs("star") = 3 Then Response.Write (" selected")%>>����</option>
		<option value=2<%If Rs("star") = 2 Then Response.Write (" selected")%>>���</option>
		<option value=1<%If Rs("star") = 1 Then Response.Write (" selected")%>>��</option>
          </select>&nbsp;&nbsp;
	  <strong><font color=blue>ע������ļ۸�</font></strong>
	  <input name="SoftPrice" type="text" size="10" onkeyup="if(isNaN(this.value))this.value=''" value="<%=Rs("SoftPrice")%>"> Ԫ
	  </td>
	</tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>��ϵ��ʽ</strong></td>
          <td class="UserTableRow1">
		<input name="Contact" type="text" size="33" value="<%=Rs("Contact")%>"> 
		<strong>������ҳ</strong>
		<input name="Homepage" type="text" size="30" value="<%=Rs("Homepage")%>">
	  </td>
	</tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�������</strong></td>
          <td class="UserTableRow1">
		<input name="Author" type="text" size="33" value="<%=Rs("Author")%>">
		<strong>ע����ַ</strong>
		<input name="Regsite" type="text" size="30" value="<%=Rs("Regsite")%>">
	  </td>
        <tr>
          <td align="right" class="UserTableRow2"><strong>������</strong></td>
          <td class="UserTableRow1"><textarea name='content1' id='content1' style='display:none'><%=Server.HTMLEncode(Rs("content"))%></textarea>
		<script Language=Javascript src="../editor/editor1.js"></script></td>
        </tr>
	        </tr>
<%
	If CInt(GroupSetting(2)) = 1 Then
%>
	<tr>
		<td class=UserTableRow2 align="right"><strong>��֤��</strong></td>
		<td class=UserTableRow1><input type="text" name="codestr" maxlength="4" size="4">&nbsp;<img src="../inc/getcode.asp"></td>
	</tr>
<%
	End If
%>
	<tr>
          <td align="right" class="UserTableRow2"><strong>��ֹ����</strong></td>
          <td class="UserTableRow1"><input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"<%If Rs("ForbidEssay") <> 0 Then Response.Write (" checked")%>>&nbsp;&nbsp;&nbsp;&nbsp;
	  <strong>��������</strong>
	  <input name="isAccept" type="checkbox" id="isAccept" value="1" disabled> �ǣ�<font color=blue>���ѡ�еĻ���ֱ�ӷ�����������˺���ܷ�����</font>��</td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>���ص�ַ</strong></td>
          <td class="UserTableRow1"><input name="SiteName" type="text" size="15" value="<%=Replace(DownSiteName(0), "del", "")%>">
	  <input name="DownAddress" type="text" id="filePath" size="60" value="<%=Replace(DownAddress(0), "del", "")%>"> <font color=red>*</font></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>���ص�ַ</strong></td>
          <td class="UserTableRow1"><input name="SiteName" type="text" size="15" value="<%=Replace(DownSiteName(1), "del", "")%>">
	  <input name="DownAddress" type="text" size="60" value="<%=Replace(DownAddress(1), "del", "")%>"></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>���ص�ַ</strong></td>
          <td class="UserTableRow1"><input name="SiteName" type="text" size="15" value="<%=Replace(DownSiteName(2), "del", "")%>">
	  <input name="DownAddress" type="text" size="60" value="<%=Replace(DownAddress(2), "del", "")%>"></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�ļ��ϴ�</strong></td>
          <td class="UserTableRow1"><iframe name="file1" frameborder=0 width='100%' height=45 scrolling=no src=upfile.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
        <tr align="center">
          <td colspan="4" class="UserTableRow2">
	  <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>
	  <input type="button" name="Submit1" value="�޸����" class=Button onclick="doSubmit();"></td>
        </tr></form>
      </table></div>
<%
	Rs.Close:Set Rs = Nothing
End Sub

%>
<!--#include file="foot.inc"-->









