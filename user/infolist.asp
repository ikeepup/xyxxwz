<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
<!--#include file="../inc/chkinput.asp"-->
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
Call InnerLocation("�ҷ�������Ϣ")
Dim Action,SQL,Rs,i

Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveArticle
	Case "edit"
		Call EditArticle
	Case "del"
		Call DeleteArticle
	Case "view"
		Call ArticleView
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
function myuser_articlelist_top(accept){
	document.write ('<th valign=middle>');
	if (accept==1)
	{
		document.write ('�ҵ���Ϣ�б�--����˵���Ϣ');
	}else{
		document.write ('�ҵ���Ϣ�б�--δ��˵���Ϣ');
	}
	document.write ('</th>');
	document.write ('<th valign=middle noWrap>����</th>');
	document.write ('<th valign=middle noWrap>���</th>');
	document.write ('<th valign=middle noWrap>��������</th>');
	document.write ('<th valign=middle noWrap>�������</th>');
	document.write ('</tr>');
}
function myuser_articlelist_not(){
	document.write ('<tr>');
	document.write ('<td class=Usertablerow1 align=center valign=middle colspan=5>û���ҵ��κ���Ϣ��</td>');
	document.write ('</tr>');
}
function myuser_articlelist_loop(ID,isaccept,zhuti,lanmu,writetime,liyou,style){
	var tablebody;
	if (style==1)
	{
		tablebody="Usertablerow1";
	}else{
		tablebody="Usertablerow2";
	}
	document.write ('<tr>');
	document.write ('<td class='+tablebody+' valign=middle>['+lanmu+'] ');
	document.write ('<a href="infolist.asp?action=view&ID='+ID+'">'+zhuti+'</a></td>');
	
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	if (liyou=="1")
	{
		document.write ('<font color=blue>��</font>');
	}else{
		document.write ('<font color=red>��</font>');
	}
	document.write ('</td>');

	
	
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	if (isaccept==1)
	{
		document.write ('<font color=blue>����</font>');
	}else{
		document.write ('<font color=red>δ��</font>');
	}
	document.write ('</td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>'+writetime+'</td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	document.write ('<a href="infolist.asp?action=edit&ID='+ID+'">�޸�</a> | ');
	document.write ('<a href="infolist.asp?action=del&ID='+ID+'" onClick="return confirm(\'ȷ��Ҫɾ������Ϣ��\')">ɾ��</a>');
	document.write ('</td>');
	document.write ('</tr>');
}
-->
</script>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr height=20 align=center>
		<td class=Usertablerow2 colspan=5><a href="?Accept=1">����˵���Ϣ</a> | 
		<a href="?">δ��˵���Ϣ</a> | 
		<a href="infopost.asp">�����µ���Ϣ</a> </td>
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

	Response.Write "<script>myuser_articlelist_top("& isAccept &")</script>" & vbNewLine
	totalrec = enchiasp.Execute("SELECT COUNT(ID) FROM ECCMS_xinxi WHERE xingming='" & enchiasp.MemberName & "' And isAccept="& isAccept)(0)

	Pcount = CInt(totalrec / maxperpage)  '�õ���ҳ��
	If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT id,isaccept,zhuti,lanmu,writetime,liyou FROM [ECCMS_xinxi] WHERE xingming='" & enchiasp.MemberName & "' And isAccept="& isAccept &" ORDER BY WriteTime DESC ,ID DESC"

	Rs.Open SQL, Conn, 1, 1
	If Rs.EOF And Rs.BOF Then
		Response.Write "<script>myuser_articlelist_not()</script>" & vbNewLine
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
			Response.Write "<script>myuser_articlelist_loop("
			Response.Write SQL(0,i)
			Response.Write ","
			Response.Write SQL(1,i)
			Response.Write ",'"
			Response.Write EncodeJS(SQL(2,i))
			Response.Write "','"
			Response.Write EncodeJS(SQL(3,i))
			Response.Write "','"
			Response.Write FormatDated(SQL(4,i),4)
			Response.Write "',"
			if sql(5,i)<>"" then
				response.write "1"
			else
				response.write "0"
			end if
			response.write ","
			Response.Write s
			Response.Write ")</script>"
			Response.Write VbCrLf
			page_count = page_count + 1
		Next
		SQL=Null
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "<tr align=right><td class=Usertablerow2 colspan=5>"
	Response.Write ShowPages(CurrentPage,Pcount,totalrec,maxperpage,"&Accept="& Request("Accept"))
	Response.Write "</td></tr>" & vbNewLine
	Response.Write "</table>"
End Sub
Function EncodeJS(str)
	str = enchiasp.HtmlEncode(str)
	str = Replace(Replace(Replace(Replace(str,"\","\\"),"'","\'"),VbCrLf,"\n"),Chr(13),"")
	EnCodeJs = str
End Function
Sub DeleteArticle()
	If CInt(GroupSetting(8)) = 0 Then
		ErrMsg = ErrMsg + "<li>�Բ�����û��ɾ����Ϣ��Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
		Founderr = True
		Exit Sub
	End If
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Request("ID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	SQL = "SELECT isAccept FROM ECCMS_xinxi WHERE xingming='" & enchiasp.MemberName & "' And isAccept=0 And ID=" & CLng(Request("ID"))
	Set Rs = enchiasp.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry������Ϣ�Ѿ�ͨ�����,��û��Ȩ��ɾ��,����ʲô��������ϵ����Ա��</li>"
		Set Rs = Nothing
		Exit Sub
	Else
		enchiasp.Execute("DELETE FROM ECCMS_xinxi WHERE xingming='" & enchiasp.MemberName & "' And isAccept=0 And ID=" & CLng(Request("ID")))
	End If
	Set Rs = Nothing
	Response.Redirect  Request.ServerVariables("HTTP_REFERER")
End Sub
Sub ArticleView()
	If Request("ID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	SQL = "select * from ECCMS_xinxi where xingming='" & enchiasp.MemberName & "' And ID=" & Request("ID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κ���Ϣ��������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	Else
%>
<script language=javascript>
var enchiasp_fontsize=9;
var enchiasp_lineheight=12;
</script>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="UserTableBorder" style="table-layout:fixed;word-break:break-all">
	<tr>
	  <th>&gt;&gt;�鿴��Ϣ����&lt;&lt;</th>
	</tr>
	<tr>
	  <td align="center" class="usertablerow2"><a href=infoList.Asp?action=edit&ID=<%=Rs("ID")%>></a></td>
	</tr>
	<tr>
	  <td align="center" class="usertablerow1">������ʱ�䣺<%=Rs("WriteTime")%> �������ˣ�<font color=blue><%=Rs("xingming")%></font></td>
	</tr>
	
	<tr>
	  <td class="usertablerow1"><p align="right"><a style="CURSOR: hand; POSITION: relative" onclick="if(enchiasp_fontsize&gt;8){enchiaspContentLabel.style.fontSize=(--enchiasp_fontsize)+&quot;pt&quot;;enchiaspContentLabel.style.lineHeight=(--enchiasp_lineheight)+&quot;pt&quot;;}" title="��С����"><img src="../images/1.gif" border="0" width="15" height="15"><font color="#FF6600">��С����</font></a> 
                    <a style="CURSOR: hand; POSITION: relative" onclick="if(enchiasp_fontsize&lt;64){enchiaspContentLabel.style.fontSize=(++enchiasp_fontsize)+&quot;pt&quot;;enchiaspContentLabel.style.lineHeight=(++enchiasp_lineheight)+&quot;pt&quot;;}" title="��������"><img src="../images/2.gif" border="0" width="15" height="15"><font color="#FF6600">��������</font></a></p>
					<div id="enchiaspContentLabel"><%=Replace(enchiasp.ReadContent(Rs("neirong")), "[page_break]", "", 1, -1, 1)%></div></td>
	</tr>
	
	<%
 if rs("isaccept")=0 then
 if rs("liyou")<>"" then
%>
	<tr>
	  <td align="left" class="usertablerow1">��<font color=red>δ���ͨ��ԭ��</font><font color=blue><%=Rs("liyou")%></font></td>
	</tr>
<%
end if
end if
%>

	
	<tr>
	  <td align="center" class="usertablerow2"><input type="button" onclick="javascript:window.close()" value="�رձ�����" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="���ض���" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub
Sub SaveArticle()
	If CInt(GroupSetting(8)) = 0 Then
		ErrMsg = ErrMsg + "<li>�Բ�����û���޸���Ϣ��Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
		Founderr = True
		Exit Sub
	End If
	Dim TextContent,ForbidEssay,isAccept,i,ArticleID
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Trim(Request.Form("lanmu")) = "��ѡ��" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ����Ŀ��</li>"
		
	End If

	
	If Trim(Request.Form("zhuti")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���ⲻ��Ϊ�գ�</li>"
	End If
	If Len(Request.Form("zhuti")) => 100 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���ⲻ�ܳ���100���ַ���</li>"
	End If
	
	If Trim(Request.Form("xingming")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("dianhua")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�绰����Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("laizi"))="" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���Բ���Ϊ�ա�</li>"
	End If
	If Trim(Request.Form("dizhi"))="" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ַ����Ϊ�ա�</li>"
	End If
	If Trim(Request.Form("email"))="" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>E-MAIL����Ϊ�ա�</li>"
	End If
	
	If IsValidEmail(Request.Form("email")) = False Then
		ErrMsg = ErrMsg + "<li>����Email�д���</li>"
		Founderr = True
	End If

	
	If CInt(GroupSetting(2)) = 1 Then
		If Not enchiasp.CodeIsTrue() Then
			ErrMsg = ErrMsg + "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&"""><li>��֤��У��ʧ�ܣ��뷵��ˢ��ҳ�����ԡ�������Զ�����</li>"
			Founderr = True
		End If
		Session("GetCode") = ""
	End If
	TextContent = ""
	For i = 1 To Request.Form("neirong").Count
		TextContent = TextContent & Request.Form("neirong")(i)
	Next
	If Len(TextContent) < 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���ݲ���Ϊ�գ�</li>"
	End If
	If Len(TextContent) > CLng(GroupSetting(16)) Then
		ErrMsg = ErrMsg + "<li>���ݲ��ܴ���" & GroupSetting(16) & "�ַ���</li>"
		Founderr = True
	End If
	If CInt(Request("isAccept")) = 1 Then
		isAccept = 1
	Else
		isAccept = 0
	End If
	If Founderr = True Then Exit Sub
	On Error Resume Next
	Call PreventRefresh  '��ˢ��
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_xinxi WHERE xingming='" & enchiasp.MemberName & "' And ID=" & CLng(Request("ID"))
	Rs.Open SQL,Conn,1,3
		rs("zhuti")=trim(Request.Form("zhuti"))
		rs("lanmu")=trim(Request.Form("lanmu"))
		rs("xingming")=trim(Request.Form("xingming"))
		rs("dianhua")=trim(Request.Form("dianhua"))
		rs("laizi")=trim(Request.Form("laizi"))
		rs("dizhi")=trim(Request.Form("dizhi"))
		rs("email")=trim(Request.Form("email"))
		rs("neirong")=textcontent
		Rs("isAccept") = isAccept

	Rs.update
	ArticleID = Rs("ID")
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>��ϲ�����޸���Ϣ�ɹ���</li><li><a href=?action=view"&"&ID=" & ArticleID & ">����˴��鿴����Ϣ</a></li>")
End Sub
Sub EditArticle()
	Dim ClassID
	dim rst
	If CInt(GroupSetting(8)) = 0 Then
		ErrMsg = ErrMsg + "<li>�Բ�����û���޸���Ϣ��Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
		Founderr = True
		Exit Sub
	End If
	If Founderr = True Then Exit Sub
	If Request("ID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	SQL = "SELECT * FROM ECCMS_xinxi WHERE xingming='" & enchiasp.MemberName & "' And ID=" & Request("ID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κ���Ϣ��������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	If Rs("isAccept") <> 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>����Ϣ�Ѿ�ͨ�����,��û��Ȩ���޸�,����ʲô��������ϵ����Ա��</li>"
		Set Rs = Nothing
		Exit Sub
	End If
%>
<script language=JavaScript>
var _maxCount = '<%=CLng(GroupSetting(16))%>';
function doChange(objText, objDrop){
	if (!objDrop) return;
	if(document.myform.BriefTopic.selectedIndex<2){
		document.myform.BriefTopic.selectedIndex+=1;
	}
	var str = objText.value;
	var arr = str.split("|");
	var nIndex = objDrop.selectedIndex;
	objDrop.length=1;
	for (var i=0; i<arr.length; i++){
		objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);
	}
	objDrop.selectedIndex = nIndex;
}
function doSubmit(){
	if (document.myform.zhuti.value==""){
		alert("���ⲻ��Ϊ�գ�");
		return false;
	}
	if (document.myform.lanmu.value=="��ѡ��"){
		alert("��ѡ����Ŀ��");
		return false;
	}
	if (document.myform.laizi.value==""){
		alert("��ѡ�����ԣ�");
		return false;
	}

	<%If CInt(GroupSetting(2)) = 1 Then%>
	if (document.myform.codestr.value==""){
		alert("����д��֤�룡");
		return false;
	}
	<%End If%>
	document.myform.Submit1.disabled = true;
	document.myform.submit();
}
</script>
<div onkeydown=CtrlEnter()>
<table  border="0" align="center" cellpadding="3" cellspacing="1" class="Usertableborder">
        <tr>
          <th colspan="2">&gt;&gt;������Ϣ&lt;&lt;</th>
        </tr>
	<form method=Post name="myform" action="infolist.asp">
	<input type="Hidden" name="action" value="save">
	<input type=hidden name=ID value='<%=Rs("ID")%>'>
   <tr>
          <td align="right" noWrap class="usertablerow2"><strong>��Ϣ����</strong></td>
          <td class="usertablerow1">
 <input name="zhuti" type="text" id="zhuti" size="60" value="<%=rs("zhuti")%>">               
          <font color=red>*</font></td>                                                  
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>����</strong></td>
          <td class="usertablerow1">
          <%
          	SQL = "select * from ECCMS_xinxisetup where ID=1"
			Set Rst = enchiasp.Execute(SQL)
			if rst.eof then
				response.write "<select name='lanmu' id='lanmu'><option>��ѡ��</option></select>"
			else
				dim ss
				dim i
				ss=split(rst("lanmu"),"|")
				response.write "<select name='lanmu' id='lanmu'><option>��ѡ��</option>"
				for i=0 to ubound(ss)
					response.write "<option"
					if rs("lanmu")=ss(i) then
						response.write " selected "
					end if
					response.write ">"& ss(i) 
					response.write "</option>"
				next			
				response.write "</select>"
			end if			
			rst.close
			set rst=nothing

			
          %> 
          <font color=red>*</font></td>                     
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>����</strong></td>
          <td class="usertablerow1"><input name="xingming2" type="text" id="xingming2" size="60" value="<%=rs("xingming")%>" disabled><input name="xingming" type="hidden" id="xingming" size="60" value="<%=rs("xingming")%>"> <font color=red>*</font></td>                                                  
        </tr>
 <tr>
          <td align="right" class="usertablerow2"><strong>�绰</strong></td>
          <td class="usertablerow1"><input name="dianhua" type="text" id="dianhua" size="60" value="<%=rs("dianhua")%>"> <font color=red>*</font></td>                                                  
        </tr>
 <tr>
          <td align="right" class="usertablerow2"><strong>����</strong></td>
          <td class="usertablerow1"><input name="laizi" type="text" id="laizi" size="60" value="<%=rs("laizi")%>"> <font color=red>*</font></td>                                                  
        </tr>
 <tr>
          <td align="right" class="usertablerow2"><strong>��ַ</strong></td>
          <td class="usertablerow1"><input name="dizhi" type="text" id="dizhi" size="60" value="<%=rs("dizhi")%>"> <font color=red>*</font></td>                                                  
        </tr>
 <tr>
          <td align="right" class="usertablerow2"><strong>E-MAIL</strong></td>
          <td class="usertablerow1"><input name="email" type="text" id="email" size="60" value="<%=rs("email")%>"> <font color=red>*</font></td>              
        </tr>

       <tr>
	  <td align="right" class="usertablerow2"><strong>��Ϣ����</strong></td>
          <td class="usertablerow1"><textarea name='neirong' id='neirong' rows="5" cols="60"><%=rs("neirong")%></textarea>
		</td>                                        
        </tr>
       <%
	If CInt(GroupSetting(2)) = 1 Then
%>
	<tr>
		<td class=Usertablerow2 align="right"><strong>��֤��</strong></td>
		<td class=Usertablerow1><input type="text" name="codestr" maxlength="4" size="4">&nbsp;<img src="../inc/getcode.asp"></td>
	</tr>
<%
	End If
%>

       
    <tr align="center">
          <td colspan="2" class="usertablerow2">
	  <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>
	  <input type="button" name="Submit1" value="�޸�����" class=Button onclick="doSubmit();"></td>
        </tr></form>
      </table></div>
<%
End Sub
%>
<!--#include file="foot.inc"-->


















































