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
Call InnerLocation("�û����ŷ���")

Dim Rs,SQL,i,Action
Dim smsincept,smscontent,smstopic,sid,sendername,Chatloglist

If CInt(GroupSetting(22)) = 0 Then
	ErrMsg = ErrMsg + "<li>�Բ�����û��ʹ�ö��ŷ����Ȩ�ޣ�����ʲô��������ϵ����Ա��</li>"
	Founderr = True
End If
If Trim(Request("touser")) <> "" Then
	sendername = enchiasp.CheckbadStr(Request("touser"))
	smsincept =  enchiasp.CheckbadStr(Request("touser"))
Else
	sendername = enchiasp.CheckbadStr(Request("sender"))
End If
Chatloglist = ""
Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "del"
		Call DelMessage
	Case "alldel"
		Call DelAllMessage
	Case "save"
		Call SaveMessage
	Case "read"
		Call ReadMessage
	Case "outread"
		Call ReadMessage
	Case "new"
		Call SendMessage
	Case "fw"
		Call SendMessage
	Case "ɾ���ռ���"
		Call Delinbox
	Case "����ռ���"
		Call DelAllinbox
	Case "ɾ��������"
		Call DelSendbox
	Case "��շ�����"
		Call DelAllSendbox
	Case Else
		ErrMsg = ErrMsg + "<li>�����ϵͳ����~!</li>"
		Founderr = True
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If

Sub SendMessage()
	Call UserMessage
	If Founderr = True Then Exit Sub
%>
<script language=JavaScript>
var _maxCount = '<%=CLng(GroupSetting(23))%>';
function DoTitle(addTitle) {  
 var revisedTitle;  
 var currentTitle = document.myform.incept.value;
 if(currentTitle=="") revisedTitle = addTitle; 
 else { 
  var arr = currentTitle.split(","); 
  for (var i=0; i < arr.length; i++) { 
   if( addTitle.indexOf(arr[i]) >=0 && arr[i].length==addTitle.length ) return; 
  } 
  revisedTitle = currentTitle+","+addTitle; 
 } 

 document.myform.incept.value=revisedTitle;  
 document.myform.incept.focus(); 
 return; 
} 

function doSubmit(){
	if (document.myform.incept.value==""){
		alert("�ռ��˲���Ϊ�գ�");
		return false;
	}
	if (document.myform.topic.value==""){
		alert("���ű��ⲻ��Ϊ�գ�");
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
		alert("�������ݲ���С��2���ַ���");
		return false;
	}
	if(MessageLength > _maxCount){
		alert("���ŵ����ݲ��ܳ���"+_maxCount+"���ַ���");
		return false;
	}
	document.myform.Submit1.disabled = true;
	document.myform.submit();
}
</script>

<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<form name=myform method=post action="message.asp">
	<input type="hidden" name="action" value="save">
	<tr>
		<th colspan=2>վ�ڶ���Ϣ</th>
	</tr>
<%
	Call MessageTop
%>
	<tr>
		<td class=Usertablerow1>�ռ���</td>
		<td class=Usertablerow1><input type=text name="incept" value="<%=smsincept%>" size=50>
		<select name=friend onchange="DoTitle(this.options[this.selectedIndex].value)">
		<option selected value="">ѡ��</option>
		<%=Option_Friend%> 
		</select></td>
	</tr>
	<tr>
		<td class=Usertablerow1>����</td>
		<td class=Usertablerow1><input type="text" name="topic" maxlength="70" size="70" value="<%=smstopic%>"></td>
	</tr>
<%
	If CInt(GroupSetting(2)) = 1 Then
%>
	<tr>
		<td class=Usertablerow1>��֤��</td>
		<td class=Usertablerow1><input type="text" name="codestr" maxlength="4" size="4">&nbsp;<img src="../inc/getcode.asp"></td>
	</tr>
<%
	End If
%>
	<tr>
		<td class=Usertablerow1 noWrap>��������</td>
		<td class=Usertablerow1><textarea name='content1' id='content1' style='display:none'><%=Server.HTMLEncode(smscontent)%></textarea>
		<script Language=Javascript src="../editor/editor1.js"></script></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 colspan=2><b>˵����</b>�������50���ַ����������<%=CLng(GroupSetting(23))%>���ַ���</td>
	</tr>
	<tr align=center height=20>
		<td class=Usertablerow2 colspan=2><input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>&nbsp;
		<input type="reset" name="submit2" value=" ��� " class=button>&nbsp;
<SCRIPT LANGUAGE="JavaScript">
<!--
var reaction='<%=enchiasp.CheckStr(Request("reaction"))%>';
var action='new';
if (action=='new')
{
if (reaction=='chatlog')
{
document.write ('<input class=button type=button value="�ر������¼" name="chatlog" onclick="location.href=\'?action=new&sid=<%=Request("sid")%>&touser=<%=sendername%>\'">');
}
else{
document.write ('<input class=button type=button value="�鿴�����¼" name="chatlog" onclick="location.href=\'?action=new&sid=<%=Request("sid")%>&touser=<%=sendername%>&reaction=chatlog\'">');
}
}
//-->
</SCRIPT>
		<input type="button" name="Submit1" value=" ���� " onclick="doSubmit();" class=button></td>
	</tr>
<SCRIPT LANGUAGE="JavaScript">
<!--
var reaction='<%=enchiasp.CheckStr(Request("reaction"))%>';
var chatloglist='<%=Chatloglist%>';
var myname='<%=enchiasp.MemberName%>';
var action='new';
if (action=='new')
{
if (reaction=='chatlog')
{
	document.write ('<tr>');
	document.write ('<th colspan=2>����<%=sendername%>�������¼</th>');
	document.write ('</tr>');
	if (myname=='')
	{
		document.write ('<tr>');
		document.write ('<td class=Usertablerow1 colspan=2>�Լ����Լ��������¼ûʲô�ÿ��ģ���</td>');
		document.write ('</tr>');
	}
	else{
		document.write (chatloglist);
	}
}
}
//-->
</SCRIPT>
	</form>
</table>
<%
End Sub

Sub MessageTop()
%>
	<tr align=center height=20>
		<td class=Usertablerow1 colspan=2><a href="message.asp?action=del&sid=<%=Request("sid")%>" onclick=showClick('��ȷ��Ҫɾ���˶�����?')><img src="images/m_delete.gif" border=0 alt="ɾ����Ϣ"></a> &nbsp; 
		<a href="message.asp?action=new"><img src="images/m_write.gif" border=0 alt="������Ϣ"></a> &nbsp;
		<a href="message.asp?action=new&touser=<%=sendername%>&sid=<%=Request("sid")%>"><img src="images/replypm.gif" border=0 alt="�ظ���Ϣ"></a>&nbsp;
		<a href="message.asp?action=fw&sid=<%=Request("sid")%>"><img src="images/m_fw.gif" border=0 alt=ת����Ϣ></a></td>
	</tr>
<%
End Sub

Sub ReadMessage()
	If Founderr = True Then Exit Sub
	If Action = "outread" Then
		Set Rs = enchiasp.Execute("select * from ECCMS_Message where sender='"&enchiasp.MemberName&"' And delSend=0 And id="& CLng(Request("sid")))
	Else
		Set Rs = enchiasp.Execute("select * from ECCMS_Message where (incept='"&enchiasp.MemberName&"' Or flag=1) And id="& CLng(Request("sid")))
	End If
	If Rs.BOF And Rs.EOF Then
		ErrMsg = ErrMsg + "<li>�����ϵͳ����~��</li>"
		Founderr = True
		Set Rs = Nothing
		Exit Sub
	End If
	Dim smsnumber
	If Rs("isRead") = 0 And Action="read" Then
		smsnumber = newincept(enchiasp.membername) - 1
		if smsnumber < 0 Then smsnumber = 0
		SQL = "Update ECCMS_User Set usermsg=" & smsnumber & " where username='"&enchiasp.membername&"'"
		enchiasp.Execute(SQL)
		if Rs("flag") = 0 Then
			SQL = "Update ECCMS_Message Set isRead=1 where id="& CLng(Request("sid"))
			enchiasp.Execute(SQL)
		End If
	End If
%>
<table cellspacing=1 align=center cellpadding=3 bgcolor=#cccccc border=0 class=Usertableborder>
	<tr>
		<th>�Ķ�����Ϣ</th>
	</tr>
<%
	Call MessageTop
%>
	<tr height=20>
		<td class=Usertablerow2>����<b><%=Rs("SendTime")%></b>��
<%
	If Action = "outread" Then
		Response.Write "����<b>" & Server.HTMLEncode(Rs("incept")) & "</b>���͵���Ϣ��"
	Else
		Response.Write "<b>" & Server.HTMLEncode(Request("sender")) & "</b>�������͵���Ϣ��"
	End If
%>
		</td>
	</tr>
	<tr>
		<td class=Usertablerow1><b>���ű��⣺</b><%=Rs("title")%><hr size=1><%=ubbcode(Rs("content"))%></td>
	</tr>
	<tr align=center height=20>
		<td class=Usertablerow2 colspan=2><input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>&nbsp;</td>
	</tr>
</table>
<%
	Set Rs = Nothing
End Sub
Sub UserMessage()
	If Founderr = True Then Exit Sub
	If Not IsNumeric(Request("sid")) And Trim(Request("sid")) <> "" Then
		ErrMsg = ErrMsg + "�����ϵͳ����!ID����������"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request("sid")) <> "" Then
		sid = CLng(Request("sid"))
	End If
	If Action = "fw" And IsNumeric(Request("sid"))  Then
		Set Rs = enchiasp.Execute("select * from ECCMS_Message where (sender='"&enchiasp.MemberName&"' Or incept='"&enchiasp.MemberName&"') And id="& CLng(Request("sid")))
		If Rs.BOF And Rs.EOF Then
			ErrMsg = ErrMsg + "<li>�����ϵͳ����~��</li>"
			Founderr = True
			Set Rs = Nothing
			Exit Sub
		End If
		smsincept = ""
		smscontent = "=================== ������ת����Ϣ =================== <br>" & Rs("content") & "<br>====================================================<br>"
		smstopic = "FW��" & Rs("title")
		sendername = Rs("sender")
		Set Rs = Nothing
	End If
	If Trim(Request("touser")) <> "" And Request("sid") <> "" Then
		Set Rs = enchiasp.Execute("select * from ECCMS_Message where id="& CLng(Request("sid")) &" And incept='"&enchiasp.MemberName&"'")
		If Rs.BOF And Rs.EOF Then
			ErrMsg = ErrMsg + "<li>�����ϵͳ����~��</li>"
			Founderr = True
			Set Rs = Nothing
			Exit Sub
		End If
		smsincept = Rs("incept")
		smscontent = "============�� " & Rs("SendTime") & " ��������д����============<br>" & Rs("content") & "<br>======================================================<br>"
		smstopic = "RW��" & Rs("title")
		sendername = Rs("sender")
		Set Rs = Nothing
	End If
	Dim Touser,temp_chat1,temp_chat2
	If Request("reaction")="chatlog" Then
		Touser=enchiasp.CheckStr(Request("touser"))
		SQL="SELECT top 30 sender,incept,title,content,sendtime FROM ECCMS_Message WHERE ((sender='"&enchiasp.MemberName&"' And incept='"&Touser&"') or (sender='"&Touser&"' And incept='"&enchiasp.MemberName&"')) And delSend=0 ORDER BY sendtime DESC"
		Set Rs=enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			Chatloglist="<tr><td class=Usertablerow1 colspan=2>��û���κ������¼��</td></tr>"
		Else
			SQL=Rs.GetRows(-1)
			Rs.close:Set Rs=nothing

			For i=0 to Ubound(SQL,2)
				chatloglist=chatloglist & "<tr><td class=Usertablerow2 height=25 colspan=2>"
				If Trim(SQL(0,i))=enchiasp.MemberName Then
					temp_chat1 = "��" & SQL(4,i)
					temp_chat1 = temp_chat1 & "�������ʹ���Ϣ��" & enchiasp.HtmlEncode(SQL(1,i))
					chatloglist=chatloglist & temp_chat1
				Else
					temp_chat2 = "��" & SQL(4,i) & "��"
					temp_chat2 = temp_chat2 & enchiasp.HtmlEncode(SQL(0,i)) & "�������͵���Ϣ��"
					chatloglist=chatloglist & temp_chat2
				End If
				chatloglist=chatloglist & "</td></tr><tr><td class=Usertablerow1 valign=top align=left colspan=2><b>��Ϣ���⣺"
				chatloglist=chatloglist & enchiasp.HtmlEncode(SQL(2,i))
				chatloglist=chatloglist & "</b><hr size=1>"
				chatloglist=chatloglist & UbbCode(SQL(3,i))
				chatloglist=chatloglist & "</td></tr>"
			Next
		End If
	End If
End Sub
Sub DelMessage()
	If Founderr = True Then Exit Sub
	If Not IsNumeric(Request("sid")) Then
		ErrMsg = ErrMsg + "<li>�Բ��𣡴����ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	SQL="SELECT incept FROM ECCMS_Message WHERE (sender='"&enchiasp.MemberName&"' Or incept='"&enchiasp.MemberName&"') And id="& Request("sid")
	Set Rs=enchiasp.Execute(SQL)
	If Rs.EOF And Rs.BOF Then
		ErrMsg = ErrMsg + "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
		Set Rs = Nothing
	Else
		If Rs(0) = enchiasp.MemberName Then
			enchiasp.Execute("Delete From ECCMS_Message where flag=0 And incept='"&enchiasp.MemberName&"' And id="& Request("sid"))
		Else
			enchiasp.Execute("Update ECCMS_Message Set delsend=1 where sender='"&enchiasp.MemberName&"' And id="& Request("sid"))
		End If
	End If
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>ɾ������Ϣ��ɣ�</li>")
End Sub
Sub DelAllMessage()
	If Founderr = True Then Exit Sub
	enchiasp.Execute("Delete From ECCMS_Message where flag=0 And incept='"&enchiasp.MemberName&"'")
	enchiasp.Execute("Update ECCMS_Message Set delsend=1 where sender='"&enchiasp.MemberName&"'")
	Call Returnsuc("<li>���Ķ���Ϣ�Ѿ�ȫ�������</li>")
End Sub
'================================================
' ��������Option_Friend
' ��  �ã��û�������������
'================================================
Function Option_Friend()
	DIM i
	SQL = "select friend from ECCMS_Friend where grouping<>2 And userid="& enchiasp.memberid &" order by addtime desc"
	Set Rs = enchiasp.Execute(Sql)
	If Not Rs.EOF Then
		SQL = Rs.GetRows(-1)
		Rs.Close:Set Rs=Nothing
	End if
	If IsArray(SQL) Then
		For i=0 To Ubound(SQL,2)
		Option_Friend = Option_Friend & "<option value="""& SQL(0,i) &""">"& SQL(0,i) &"</option> "
		Next
	Else
		Option_Friend = ""
	End If
End Function
'================================================
' ��������newincept
' ��  �ã�ͳ�ƶ���
'================================================
Function newincept(iusername)
	Dim Rs
	Rs = enchiasp.Execute("Select Count(id) from ECCMS_Message where isRead=0 And flag=0 And incept='"& iusername &"'")
	newincept = Rs(0)
	Set Rs=Nothing
	If IsNull(newincept) Then newincept = 0
End Function
'================================================
' ��������ChkHateName
' ��  �ã���������֤
'================================================
Function ChkHateName(sName)
	DIM SQL,Rs
	ChkHateName = False
	SQL="Select friend From ECCMS_Friend Where (userid="& enchiasp.memberid &" Or username='"& sName &"') And grouping=2"
	Set Rs = enchiasp.Execute(SQL)
	If Not Rs.EOF Then
		SQL=Rs.GetString(,, ",", "", "")
		Rs.Close:Set Rs=Nothing
		If Instr(SQL,sName) Or Instr(SQL,enchiasp.membername) Then ChkHateName = True
	End If
End Function
'================================================
' ��������CheckID
' ��  �ã���֤����ID
'================================================
Function CheckID(CHECK_ID)
	Dim Delid,Fixid
	CheckID=True
	Delid=replace(CHECK_ID,"'","")
	Delid=replace(Delid,";","")
	Delid=replace(Delid,"--","")
	Delid=replace(Delid,")","")
	Fixid=replace(Delid,",","")
	Fixid=Trim(replace(fixid," ",""))
	If Delid="" or isnull(Delid) Then  CheckID=False
	If Not IsNumeric(fixid) Then CheckID=False
End Function
'================================================
' ��������SaveMessage
' ��  �ã��������Ϣ
'================================================
Sub SaveMessage()
	Dim strIncept,strContent,strTitle,InceptName,n
	If CLng(UserToday(4)) => CLng(GroupSetting(29)) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ÿ�����ֻ�ܷ���<font color=red><b>" & GroupSetting(29) & "</b></font>ƪ���£������Ҫ�������������������ɣ�</li>"
	End If
	If Trim(Request.Form("incept")) = "" Then
		ErrMsg = ErrMsg + "<li>����д�ռ���������</li>"
		Founderr = True
	Else
		strIncept = enchiasp.CheckbadStr(Request.Form("incept"))
		strIncept = split(strIncept,",")
	End If
	If Trim(Request.Form("topic")) = "" Then
		ErrMsg = ErrMsg + "<li>����д���ű��⣡</li>"
		Founderr = True
	Else
		strTitle = Left(enchiasp.ChkFormStr(Request.Form("topic")),50)
	End If
	If Trim(Request.Form("content1")) = "" Then
		ErrMsg = ErrMsg + "<li>����д�������ݣ�</li>"
		Founderr = True
	Else
		strContent = Html2Ubb(Request.Form("content1"))
	End If
	If Len(Request.Form("content1")) > CLng(GroupSetting(23)) Then
		ErrMsg = ErrMsg + "<li>�������ݲ��ܴ���" & GroupSetting(23) & "�ַ���</li>"
		Founderr = True
	End If
	If CInt(GroupSetting(2)) = 1 Then
		If Not enchiasp.CodeIsTrue() Then
			ErrMsg = ErrMsg + "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&"""><li>��֤��У��ʧ�ܣ��뷵��ˢ��ҳ�����ԡ�������Զ�����</li>"
			Founderr = True
		End If
		Session("GetCode") = ""
	End If
	If Founderr = True Then Exit Sub
	On Error Resume Next
	Call PreventRefresh  '��ˢ��
	n=0
	For i = 0 To Ubound(strIncept)
		If i >= 5 Then Exit For
		n = n + 1
		InceptName = Trim(strIncept(i))
		SQL = "select username from [ECCMS_User] where username='"&InceptName&"'"
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			ErrMsg = ErrMsg + "<li>û���ҵ�<font color=red>" & InceptName & "</font>����û������ŷ��Ͳ��ɹ�~��</li>"
			Founderr = True
			Rs.Close:Set Rs = Nothing
			Exit Sub
		Else
			InceptName = Rs(0)
		End If
		Rs.Close:Set Rs = Nothing
		If ChkHateName(InceptName) Then
			ErrMsg = ErrMsg + "���ڶԷ�<font color=red>" & InceptName & "</font>�ѽ����������������<font color=red>" & InceptName & "</font>������ĺ������У���˶��ŷ��ͱ���ֹ��"
			Founderr = True
			Exit Sub
		Else
			SQL = "Insert into ECCMS_Message (sender,incept,title,content,flag,SendTime,isRead,delSend) values ('"& enchiasp.membername &"','"& InceptName &"','"& strTitle &"','"& strContent &"',0,"& NowString &",0,0) "
			enchiasp.Execute(SQL)
			SQL = "Update ECCMS_User Set usermsg=usermsg+1 where username='"&InceptName&"'"
			enchiasp.Execute(SQL)
		End If
		
	Next
	Dim strUserToday
	strUserToday = UserToday(0) &","& UserToday(1) &","& UserToday(2) &","& UserToday(3) &","& UserToday(4)+n &","& UserToday(5)
	UpdateUserToday(strUserToday)
	Call Returnsuc("<li>��ϲ�������Ͷ��ųɹ���</li>")
End Sub
'ɾ���ռ���
Sub Delinbox()
	If Not CheckID(Request("id")) Then
		ErrMsg = ErrMsg + "<li>�����ϵͳ������</li>"
		Founderr = True
	End If
	If Founderr = True Then Exit Sub
	enchiasp.Execute("Delete From ECCMS_Message where flag=0 And incept='"&enchiasp.MemberName&"' And id in (" & enchiasp.CheckBadstr(Request("id")) & ")")
	Call Returnsuc("<li>ɾ���ռ����еĶ��ųɹ���</li>")
End Sub
'����ռ���
Sub DelAllinbox()
	If Founderr = True Then Exit Sub
	enchiasp.Execute("Delete From ECCMS_Message where flag=0 And incept='"&enchiasp.MemberName&"'")
	Call Returnsuc("<li>�����ռ����ѳɹ���գ�</li>")
End Sub
'ɾ��������
Sub DelSendbox()
	If Not CheckID(Request("id")) Then
		ErrMsg = ErrMsg + "<li>�����ϵͳ������</li>"
		Founderr = True
	End If
	If Founderr = True Then Exit Sub
	enchiasp.Execute("Update ECCMS_Message Set delsend=1 where sender='"&enchiasp.MemberName&"' And id in (" & enchiasp.CheckBadstr(Request("id")) & ")")
	Call Returnsuc("<li>ɾ���������еĶ��ųɹ���</li>")
End Sub
'��շ�����
Sub DelAllSendbox()
	If Founderr = True Then Exit Sub
	enchiasp.Execute("Update ECCMS_Message Set delsend=1 where sender='"&enchiasp.MemberName&"'")
	Call Returnsuc("<li>���ķ������ѳɹ���գ�</li>")
End Sub
%>
<!--#include file="foot.inc"-->




















