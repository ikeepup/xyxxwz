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
Call InnerLocation("��������")
Dim Rs,SQL
dim temp
dim i
dim tt
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID < 1 Then ChannelID = 1
ChannelID = CLng(ChannelID)

If CInt(GroupSetting(7)) = 0 Then
	ErrMsg = ErrMsg + "<li>�Բ�����û�з������µ�Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
	Founderr = True
End If
tt=false
temp=split(GroupSetting(35),"$$$")
for i=0 to ubound(temp)
	If temp(i) = ChannelID  Then
		tt=true
		exit for
	End If
next

if tt=false then
	ErrMsg = ErrMsg + "<li>�Բ�����û�и�Ƶ���������µ�Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
	Founderr = True
end if


Dim Action:Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveArticle
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
	<form method=Post name="myform" action="ArticlePost.Asp">
	<input type="Hidden" name="action" value="save">
	<input type=hidden name=ChannelID value='<%=ChannelID%>'>
        <tr>
          <td width="15%" align="right" nowrap class="usertablerow2"><strong>��������</strong></td>
          <td width="85%" class="usertablerow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	'sClassSelect = Replace(sClassSelect, "{ClassID=" & ClassID & "}", "selected")
	Response.Write sClassSelect
	Response.Write "</select>"
%>
	  </td>
        </tr>
        <tr>
          <td align="right" noWrap class="usertablerow2"><strong>���±���</strong></td>
          <td class="usertablerow1"><select name="BriefTopic" id="BriefTopic">
            <option value="0">ѡ����</option>
			<option value="1">[ͼ��]</option>
			<option value="2">[��ͼ]</option>
			<option value="3">[����]</option>
			<option value="4">[�Ƽ�]</option>
			<option value="5">[ע��]</option>
			<option value="6">[ת��]</option>
          </select> <input name="title" type="text" id="title" size="60" value="">       
          <font color=red>*</font></td>      
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>�������</strong></td>
          <td class="usertablerow1"><input name="Related" type="text" id="Related" size="60" value=""> <font color=red>*</font></td>      
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>��������</strong></td>
          <td class="usertablerow1"><input name="Author" type="text" size="30" value="">
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
          <td class="usertablerow1"><input name="ComeFrom" type="text" size="30" value="">
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
          <td class="usertablerow1"><textarea name='content' id='content' style='display:none' rows="1" cols="20"></textarea>
		<script Language=Javascript src="../editor/post.js"></script></td>      
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>��ҳͼƬ</strong></td>
          <td class="usertablerow1"><input name="ImageUrl" type="text" id="ImageUrl" size="60" value="">
			<input type=hidden name=UploadFileList id=UploadFileList onchange="doChange(this,document.myform.ImageFileList)" value="">
			<br>ֱ�Ӵ��ϴ�ͼƬ��ѡ��
			<select name="ImageFileList" id="ImageFileList" onChange="ImageUrl.value=this.value;">
			<option value=''>��ѡ����ҳ�Ƽ�ͼƬ</option>
			</select>
			</td>
        </tr>
	<tr>
	  <td align="right" class="usertablerow2"><strong>�ļ��ϴ�</strong></td>
          <td class="usertablerow1"><iframe name="image1" frameborder=0 width='100%' height=55 scrolling=no src=upload.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>�����Ǽ�</strong></td>
          <td class="usertablerow1"><select name="star">
		<option value=5>������</option>
          	<option value=4>�����</option>
          	<option value=3 selected>����</option>
		<option value=2>���</option>
		<option value=1>��</option>
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
          <td class="usertablerow1"><input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"></td>
        </tr>
	<tr>
          <td align="right" class="usertablerow2"><strong>��������</strong></td>
          <td class="usertablerow1"><input name="isAccept" type="checkbox" id="isAccept" value="1" disabled> �ǣ�<font color=blue>���ѡ�еĻ���ֱ�ӷ�����������˺���ܷ�����</font>��</td>
        </tr>
        <tr align="center">
          <td colspan="2" class="usertablerow2">
	  <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>
	  <input type="button" name="Submit1" value="���ڷ���" class=Button onclick="doSubmit();"></td>
        </tr></form>
      </table></div>
<%
End Sub
Private Sub SaveArticle()
	Dim TextContent,ForbidEssay,isAccept,i,ArticleID
	If CLng(UserToday(3)) => CLng(GroupSetting(10)) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ÿ�����ֻ�ܷ���<font color=red><b>" & GroupSetting(10) & "</b></font>ƪ���£������Ҫ�������������������ɣ�</li>"
	End If
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
	If Len(TextContent) < 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������ݲ���Ϊ�գ�</li>"
	End If
	If Len(TextContent) > CLng(GroupSetting(16)) Then
		ErrMsg = ErrMsg + "<li>�������ݲ��ܴ���" & GroupSetting(16) & "�ַ���</li>"
		Founderr = True
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
	If Founderr = True Then Exit Sub
	On Error Resume Next
	Call PreventRefresh  '��ˢ��
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Article where (ArticleID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = enchiasp.ChkNumeric(Request.Form("ClassID"))
		Rs("SpecialID") = 0
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("ColorMode") = 0
		Rs("FontMode") = 0
		Rs("content") = Html2Ubb(TextContent)
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Author") = enchiasp.ChkFormStr(Request.Form("Author"))
		Rs("ComeFrom") = enchiasp.ChkFormStr(Request.Form("ComeFrom"))
		Rs("star") = enchiasp.ChkNumeric(Request.Form("star"))
		Rs("isTop") = 0
		Rs("AllHits") = 0
		Rs("DayHits") = 0
		Rs("WeekHits") = 0
		Rs("MonthHits") = 0
		Rs("HitsTime") = Now()
		Rs("WriteTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("username") = Trim(enchiasp.MemberName)
		Rs("isBest") = 0
		Rs("BriefTopic") = enchiasp.ChkNumeric(Request.Form("BriefTopic"))
		Rs("ImageUrl") = enchiasp.ChkFormStr(Request.Form("ImageUrl"))
		Rs("UploadImage") = enchiasp.ChkFormStr(Request.Form("UploadFileList"))
		Rs("UserGroup") = 0
		Rs("PointNum") = 0
		Rs("isUpdate") = 1
		Rs("isAccept") = isAccept
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	Rs.Close
	Rs.Open "SELECT TOP 1 ArticleID FROM ECCMS_Article WHERE ChannelID=" & ChannelID & " ORDER BY ArticleID DESC", Conn, 1, 1
	ArticleID = Rs("ArticleID")
	Rs.Close:Set Rs = Nothing
	Dim strUserToday
	strUserToday = UserToday(0) &","& UserToday(1) &","& UserToday(2) &","& UserToday(3)+1 &","& UserToday(4) &","& UserToday(5)
	UpdateUserToday(strUserToday)
	Call Returnsuc("<li>��ϲ�����ύ�ɹ�����ȴ�����Ա��֤����ʽ������</li><li><a href=ArticlePost.Asp?action=view&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & ">����˴��鿴������</a></li>")
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
	  <td align="center" class="usertablerow1">���ߣ�<%=Rs("Author")%>     
        ����Դ�ڣ�<%=Rs("ComeFrom")%>  
        ������ʱ�䣺<%=Rs("WriteTime")%>       �������ˣ�<font color=blue><%=Rs("username")%></font></td>
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
%>
<!--#include file="foot.inc"-->