<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
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
Call InnerLocation("�ҷ���������")
Dim Action,SQL,Rs,i
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID < 1 Then ChannelID = 1
ChannelID = CLng(ChannelID)

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
		document.write ('�ҵ������б�--����˵�����');
	}else{
		document.write ('�ҵ������б�--δ��˵�����');
	}
	document.write ('</th>');
	document.write ('<th valign=middle noWrap>���</th>');
	document.write ('<th valign=middle noWrap>��������</th>');
	document.write ('<th valign=middle noWrap>�������</th>');
	document.write ('</tr>');
}
function myuser_articlelist_not(){
	document.write ('<tr>');
	document.write ('<td class=Usertablerow1 align=center valign=middle colspan=5>û���ҵ��κ����¡�</td>');
	document.write ('</tr>');
}
function myuser_articlelist_loop(channelid,ArticleID,accept,title,classname,dated,hits,style){
	var tablebody;
	if (style==1)
	{
		tablebody="Usertablerow1";
	}else{
		tablebody="Usertablerow2";
	}
	document.write ('<tr>');
	document.write ('<td class='+tablebody+' valign=middle>['+classname+'] ');
	document.write ('<a href="articlelist.asp?action=view&channelid='+channelid+'&ArticleID='+ArticleID+'">'+title+'</a></td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	if (accept==1)
	{
		document.write ('<font color=blue>����</font>');
	}else{
		document.write ('<font color=red>δ��</font>');
	}
	document.write ('</td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>'+dated+'</td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	document.write ('<a href="articlelist.asp?action=edit&channelid='+channelid+'&ArticleID='+ArticleID+'">�޸�</a> | ');
	document.write ('<a href="articlelist.asp?action=del&channelid='+channelid+'&ArticleID='+ArticleID+'" onClick="return confirm(\'ȷ��Ҫɾ����������\')">ɾ��</a>');
	document.write ('</td>');
	document.write ('</tr>');
}
-->
</script>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr height=20 align=center>
		<td class=Usertablerow2 colspan=5><a href="?ChannelID=<%=ChannelID%>&Accept=1">����˵�����</a> | 
		<a href="?ChannelID=<%=ChannelID%>">δ��˵�����</a> | 
		<a href="articlepost.asp?ChannelID=<%=ChannelID%>">�����µ�����</a> </td>
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
	totalrec = enchiasp.Execute("SELECT COUNT(ArticleID) FROM ECCMS_Article WHERE ChannelID = " & ChannelID & " And username='" & enchiasp.MemberName & "' And isAccept="& isAccept)(0)
	Pcount = CInt(totalrec / maxperpage)  '�õ���ҳ��
	If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT A.ArticleID,A.title,A.WriteTime,A.AllHits,A.isAccept,C.ClassName FROM [ECCMS_Article] A INNER JOIN [ECCMS_Classify] C on A.ClassID=C.ClassID WHERE A.ChannelID = " & ChannelID & "  And A.username='" & enchiasp.MemberName & "' And isAccept="& isAccept &" ORDER BY A.isTop DESC, A.WriteTime DESC ,A.ArticleID DESC"
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
			Response.Write ChannelID
			Response.Write ","
			Response.Write SQL(0,i)
			Response.Write ","
			Response.Write SQL(4,i)
			Response.Write ",'"
			Response.Write EncodeJS(SQL(1,i))
			Response.Write "','"
			Response.Write EncodeJS(SQL(5,i))
			Response.Write "','"
			Response.Write FormatDated(SQL(2,i),4)
			Response.Write "',"
			Response.Write SQL(3,i)
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
Function EncodeJS(str)
	str = enchiasp.HtmlEncode(str)
	str = Replace(Replace(Replace(Replace(str,"\","\\"),"'","\'"),VbCrLf,"\n"),Chr(13),"")
	EnCodeJs = str
End Function
Sub DeleteArticle()
	If CInt(GroupSetting(8)) = 0 Then
		ErrMsg = ErrMsg + "<li>�Բ�����û��ɾ�����µ�Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
		Founderr = True
		Exit Sub
	End If
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Request("ArticleID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	SQL = "SELECT isAccept FROM ECCMS_Article WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And isAccept=0 And ArticleID=" & CLng(Request("ArticleID"))
	Set Rs = enchiasp.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry���������Ѿ�ͨ�����,��û��Ȩ��ɾ��,����ʲô��������ϵ����Ա��</li>"
		Set Rs = Nothing
		Exit Sub
	Else
		enchiasp.Execute("DELETE FROM ECCMS_Article WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And isAccept=0 And ArticleID=" & CLng(Request("ArticleID")))
	End If
	Set Rs = Nothing
	Response.Redirect  Request.ServerVariables("HTTP_REFERER")
End Sub
Sub ArticleView()
	If Request("ArticleID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	If ChannelID = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ָ��Ƶ����</li>"
		Exit Sub
	End If
	SQL = "select ArticleID,title,content,ColorMode,FontMode,Author,ComeFrom,WriteTime,username from ECCMS_Article where ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And ArticleID=" & Request("ArticleID")
	'response.write sql
	'response.end
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κ����¡�������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	Else
%>
<script language=javascript>
var enchiasp_fontsize=9;
var enchiasp_lineheight=12;
</script>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="UserTableBorder" style="table-layout:fixed;word-break:break-all">
	<tr>
	  <th>&gt;&gt;�鿴��������&lt;&lt;</th>
	</tr>
	<tr>
	  <td align="center" class="usertablerow2"><a href=ArticleList.Asp?action=edit&ChannelID=<%=ChannelID%>&ArticleID=<%=Rs("ArticleID")%>><font size=4><%=enchiasp.ReadFontMode(Rs("title"),Rs("ColorMode"),Rs("FontMode"))%></font></a></td>
	</tr>
	<tr>
	  <td align="center" class="usertablerow1">���ߣ�<%=Rs("Author")%> ����Դ�ڣ�<%=Rs("ComeFrom")%> ������ʱ�䣺<%=Rs("WriteTime")%> �������ˣ�<font color=blue><%=Rs("username")%></font></td>
	</tr>
	<tr>
	  <td class="usertablerow1"><p align="right"><a style="CURSOR: hand; POSITION: relative" onclick="if(enchiasp_fontsize&gt;8){enchiaspContentLabel.style.fontSize=(--enchiasp_fontsize)+&quot;pt&quot;;enchiaspContentLabel.style.lineHeight=(--enchiasp_lineheight)+&quot;pt&quot;;}" title="��С����"><img src="../images/1.gif" border="0" width="15" height="15"><font color="#FF6600">��С����</font></a> 
                    <a style="CURSOR: hand; POSITION: relative" onclick="if(enchiasp_fontsize&lt;64){enchiaspContentLabel.style.fontSize=(++enchiasp_fontsize)+&quot;pt&quot;;enchiaspContentLabel.style.lineHeight=(++enchiasp_lineheight)+&quot;pt&quot;;}" title="��������"><img src="../images/2.gif" border="0" width="15" height="15"><font color="#FF6600">��������</font></a></p>
					<div id="enchiaspContentLabel"><%=Replace(enchiasp.ReadContent(Rs("content")), "[page_break]", "", 1, -1, 1)%></div></td>
	</tr>
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
		ErrMsg = ErrMsg + "<li>�Բ�����û���޸����µ�Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
		Founderr = True
		Exit Sub
	End If
	Dim TextContent,ForbidEssay,isAccept,i,ArticleID
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Trim(Request.Form("title")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���±��ⲻ��Ϊ�գ�</li>"
	End If
	If Len(Request.Form("title")) => 100 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���±��ⲻ�ܳ���100���ַ���</li>"
	End If
	If Len(Request.Form("Related")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������²��ܳ���200���ַ���</li>"
	End If
	If Trim(Request.Form("Author")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������߲���Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("ComeFrom")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������Դ����Ϊ�գ�</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����Ǽ�����Ϊ�ա�</li>"
	End If

	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��һ�������Ѿ����������࣬����������ݣ�</li>"
	End If
	If Trim(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�÷������ⲿ���ӣ�����������ݣ�</li>"
	End If
	If Trim(Request.Form("content")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������ݲ���Ϊ�գ�</li>"
	End If
	If Len(Request.Form("content")) > CLng(GroupSetting(16)) Then
		ErrMsg = ErrMsg + "<li>�������ݲ��ܴ���" & GroupSetting(16) & "�ַ���</li>"
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
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next
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
	If Founderr = True Then Exit Sub
	On Error Resume Next
	Call PreventRefresh  '��ˢ��
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_Article WHERE username='" & enchiasp.MemberName & "' And ArticleID=" & CLng(Request("ArticleID"))
	Rs.Open SQL,Conn,1,3
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = enchiasp.CheckNumeric(Request.Form("ClassID"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("content") = Html2Ubb(TextContent)
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Author") = enchiasp.ChkFormStr(Request.Form("Author"))
		Rs("ComeFrom") = enchiasp.ChkFormStr(Request.Form("ComeFrom"))
		Rs("star") = enchiasp.ChkNumeric(Request.Form("star"))
		Rs("username") = Trim(enchiasp.MemberName)
		Rs("BriefTopic") = enchiasp.ChkNumeric(Request.Form("BriefTopic"))
		Rs("ImageUrl") = enchiasp.ChkFormStr(Request.Form("ImageUrl"))
		Rs("UploadImage") = enchiasp.ChkFormStr(Request.Form("UploadFileList"))
		Rs("isAccept") = isAccept
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	ArticleID = Rs("ArticleID")
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>��ϲ�����޸����³ɹ���</li><li><a href=?action=view&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & ">����˴��鿴������</a></li>")
End Sub
Function InitSelect(UploadFileList, ImageUrl)
	Dim i
	InitSelect = "<select name='ImageFileList' onChange=""ImageUrl.value=this.value;"">"
	InitSelect = InitSelect & "<option value=''>��ѡ����ҳ�Ƽ�ͼƬ</option>"
	If Not IsNull(UploadFileList) Then
		UploadFileList = Split(UploadFileList, "|")
		For i = 0 To UBound(UploadFileList)
			If UploadFileList(i) <> "" Then
				InitSelect = InitSelect & "<option value=""" & UploadFileList(i) & """"
				If UploadFileList(i) = ImageUrl Then
					InitSelect = InitSelect & " selected"
				End If
				InitSelect = InitSelect & ">" & UploadFileList(i) & "</option>"
			End If
		Next
	End If
	InitSelect = InitSelect & "</select>"
End Function
Sub EditArticle()
	Dim ClassID
	If CInt(GroupSetting(8)) = 0 Then
		ErrMsg = ErrMsg + "<li>�Բ�����û���޸����µ�Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
		Founderr = True
		Exit Sub
	End If
	If Founderr = True Then Exit Sub
	If Request("ArticleID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	If ChannelID = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ָ��Ƶ����</li>"
		Exit Sub
	End If
	SQL = "SELECT * FROM ECCMS_Article WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And ArticleID=" & Request("ArticleID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κ����¡�������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	ClassID = Rs("ClassID")
	If Rs("isAccept") <> 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������Ѿ�ͨ�����,��û��Ȩ���޸�,����ʲô��������ϵ����Ա��</li>"
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
	if (document.myform.title.value==""){
		alert("���±��ⲻ��Ϊ�գ�");
		return false;
	}
	if (document.myform.Author.value==""){
		alert("�������߲���Ϊ�գ�");
		return false;
	}
	if (document.myform.ComeFrom.value==""){
		alert("������Դ����Ϊ�գ�");
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
	<%If CInt(GroupSetting(2)) = 1 Then%>
	if (document.myform.codestr.value==""){
		alert("����д��֤�룡");
		return false;
	}
	<%End If%>
	myform.content.value = getHTML(); 
	MessageLength = Composition.document.body.innerHTML.length;
	if(MessageLength < 2){
		alert("�������ݲ���С��2���ַ���");
		return false;
	}
	if(MessageLength > _maxCount){
		alert("���µ����ݲ��ܳ���"+_maxCount+"���ַ���");
		return false;
	}
	document.myform.Submit1.disabled = true;
	document.myform.submit();
}
</script>
<div onkeydown=CtrlEnter()>
<table  border="0" align="center" cellpadding="3" cellspacing="1" class="Usertableborder">
        <tr>
          <th colspan="2">&gt;&gt;��������&lt;&lt;</th>
        </tr>
	<form method=Post name="myform" action="Articlelist.Asp">
	<input type="Hidden" name="action" value="save">
	<input type=hidden name=ChannelID value='<%=ChannelID%>'>
	<input type=hidden name=ArticleID value='<%=Rs("ArticleID")%>'>
        <tr>
          <td width="15%" align="right" nowrap class="usertablerow2"><strong>��������</strong></td>
          <td width="85%" class="usertablerow1">
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
          <td align="right" noWrap class="usertablerow2"><strong>���±���</strong></td>
          <td class="usertablerow1"><select name="BriefTopic" id="BriefTopic">
            <option value="0">ѡ����</option>
			<option value="0"<%If Rs("BriefTopic") = 0 Then Response.Write (" selected")%>>ѡ����</option>
			<option value="1"<%If Rs("BriefTopic") = 1 Then Response.Write (" selected")%>>[ͼ��]</option>
			<option value="2"<%If Rs("BriefTopic") = 2 Then Response.Write (" selected")%>>[��ͼ]</option>
			<option value="3"<%If Rs("BriefTopic") = 3 Then Response.Write (" selected")%>>[����]</option>
			<option value="4"<%If Rs("BriefTopic") = 4 Then Response.Write (" selected")%>>[�Ƽ�]</option>
			<option value="5"<%If Rs("BriefTopic") = 5 Then Response.Write (" selected")%>>[ע��]</option>
			<option value="6"<%If Rs("BriefTopic") = 6 Then Response.Write (" selected")%>>[ת��]</option>
          </select> <input name="title" type="text" id="title" size="60" value="<%=Rs("title")%>"> 
          <font color=red>*</font></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>�������</strong></td>
          <td class="usertablerow1"><input name="Related" type="text" id="Related" size="60" value="<%=Rs("Related")%>"> <font color=red>*</font></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>��������</strong></td>
          <td class="usertablerow1"><input name="Author" type="text" size="30" value="<%=Rs("Author")%>">
		    <select name=font2 onChange="Author.value=this.value;">
			<option selected value="">ѡ������</OPTION>
			<option value=����>����</option>
			<option value=��վ>��վ</option>
			<option value=����>����</option>
			<option value=δ֪>δ֪</option>
			</select></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>������Դ</strong></td>
          <td class="usertablerow1"><input name="ComeFrom" type="text" size="30" value="<%=Rs("ComeFrom")%>">
		  	<select name=font1 onChange="ComeFrom.value=this.value;">
			<option selected value="">ѡ����Դ</OPTION>
			<option value=��վԭ��>��վԭ��</option>
			<option value=��վ����>��վ����</option>
			<option value=����>����</option>
			<option value=ת��>ת��</option>
			</select></td>
        </tr>
        <tr>
	  <td align="right" class="usertablerow2"><strong>��������</strong></td>
          <td class="usertablerow1"><textarea name='content' id='content' style='display:none'><%=Server.HTMLEncode(Rs("content"))%></textarea>
		<script Language=Javascript src="../editor/post.js"></script></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>��ҳͼƬ</strong></td>
          <td class="usertablerow1"><input name="ImageUrl" type="text" id="ImageUrl" size="60" value="<%=Rs("ImageUrl")%>">
			<input type=hidden name=UploadFileList id=UploadFileList onchange="doChange(this,document.myform.ImageFileList)" value="<%=Rs("UploadImage")%>">
			<br>ֱ�Ӵ��ϴ�ͼƬ��ѡ��
			<%
			Response.Write InitSelect(Rs("UploadImage"),Rs("ImageUrl"))
			%>
			</td>
        </tr>
	<tr>
	  <td align="right" class="usertablerow2"><strong>�ļ��ϴ�</strong></td>
          <td class="usertablerow1"><iframe name="image1" frameborder=0 width='100%' height=55 scrolling=no src=upload.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>�����Ǽ�</strong></td>
          <td class="usertablerow1"><select name="star">
          	<option value=5<%If Rs("star") = 5 Then Response.Write (" selected")%>>������</option>
          	<option value=4<%If Rs("star") = 4 Then Response.Write (" selected")%>>�����</option>
          	<option value=3<%If Rs("star") = 3 Then Response.Write (" selected")%>>����</option>
		<option value=2<%If Rs("star") = 2 Then Response.Write (" selected")%>>���</option>
		<option value=1<%If Rs("star") = 1 Then Response.Write (" selected")%>>��</option>
          </select></td>
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
	<tr>
          <td align="right" class="usertablerow2"><strong>��ֹ����</strong></td>
          <td class="usertablerow1"><input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"<%If Rs("ForbidEssay") <> 0 Then Response.Write (" checked")%>></td>
        </tr>
	<tr>
          <td align="right" class="usertablerow2"><strong>��������</strong></td>
          <td class="usertablerow1"><input name="isAccept" type="checkbox" id="isAccept" value="1" disabled> �ǣ�<font color=blue>���ѡ�еĻ���ֱ�ӷ�����������˺���ܷ�����</font>��</td>
        </tr>
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












