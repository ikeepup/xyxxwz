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
Dim Action
Dim i,ii,isEdit,RsObj
Dim keyword,FindWord,strClass
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Dim s_ClassName,ClassID,ChildStr,FoundSQL,isAccept,selArticleID
Dim TextContent,ArticleTop,ArticleBest,ArticleID,ForbidEssay,ArticleAccept
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If Trim(Request("isAccept")) <> "" Then
	isAccept = 0
Else
	isAccept = 1
End If
If CInt(ChannelID) = 0 Then ChannelID = 1
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "save"
	Call SaveArticle
Case "modify"
	Call ModifyArticle
Case "add"
	isEdit = False
	Call ArticleEdit(isEdit)
Case "edit"
	isEdit = True
	Call ArticleEdit(isEdit)
Case "del"
	Call ArticleDel
Case "batdel"
	Call PageTop
	Call BatcDelete
Case "alldel"
	Call AllDelArticle
Case "view"
	Call ArticleView
Case "renew"
	Call ArticleRenew
Case "setting"
	Call PageTop
	Call BatchSetting
Case "saveset"
	Call SaveSetting
Case "move"
	Call PageTop
	Call BatchMove
Case "savemove"
	Call SaveMove
Case "reset"
	Call ResetDateTime
Case Else
	Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If

Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub PageTop()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th colspan=2>" & sChannelName & "����ѡ��</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr><form method=Post name=myform action=admin_article.asp onSubmit='return JugeQuery(this);'>"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "	  <td class=TableRow1>������"
	Response.Write "	  <input name=keyword type=text size=20>"
	Response.Write "	  ������"
	Response.Write "	  <select name=queryopt>"
	Response.Write "		<option value=1 selected>���±���</option>"
	Response.Write "		<option value=2>¼ �� ��</option>"
	Response.Write "		<option value=3>��������</option>"
	Response.Write "	  </select> <input type=submit name=Submit value='��ʼ����' class=Button></td>"
	Response.Write "	  <td class=TableRow1>" & sModuleName & "������"
	Dim srtClassMenu
	Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Response.Write "<option value='admin_article.asp?ChannelID=" & ChannelID & "'>��ȫ��" & sModuleName & "�б��</option>" & vbCrLf
	srtClassMenu = enchiasp.ClassJumpMenu(ChannelID)
	srtClassMenu = Replace(srtClassMenu, "{ClassID=" & Request("sortid") & "}", "selected")
	Response.Write srtClassMenu
	Response.Write "</select>"
	Response.Write "	  </td></form>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td colspan=2 class=TableRow2><b>����ѡ�</b> <a href='admin_article.asp?ChannelID=" & ChannelID & "'>������ҳ</a> | "
	Response.Write "	  <a href='admin_article.asp?ChannelID=" & ChannelID & "&action=add'>���" & sModuleName & "</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "&action=add'>���" & sModuleName & "����</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "'>" & sModuleName & "�������</a> | "
	Response.Write "	  <a href='?ChannelID=" & ChannelID & "&specialID=0'>" & sModuleName & "ר���б�</a> | "
	Response.Write "	  <a href='?ChannelID=" & ChannelID & "&isAccept=0'>����" & sModuleName & "</a> | "
	Response.Write "	  <a href='admin_createArticle.asp?ChannelID=" & ChannelID & "'>����HTML</a></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub

Private Sub showmain()
	If Not ChkAdmin("AdminArticle" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	If Not IsEmpty(Request("selArticleID")) Then
		selArticleID = Request("selArticleID")
		Select Case enchiasp.CheckStr(Request("act"))
		Case "����ɾ��":Call batdel
		Case "�����Ƽ�":Call isCommend
		Case "ȡ���Ƽ�":Call noCommend
		Case "�����ö�":Call isTop
		Case "ȡ���ö�":Call noTop
		Case "�������":Call BatAccept
		Case "ȡ�����":Call NotAccept
		Case "����HTML"
			Call BatCreateHtml
		Case Else
			Response.Write "��Ч������"
		End Select
	End If
	Call PageTop
	Dim sortid,specialID,Cmd,limitime
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>ѡ��</th>"
	Response.Write "	  <th width='68%'>" & sModuleName & "����</th>"
	Response.Write "	  <th width='9%' nowrap>�������</th>"
	Response.Write "	  <th width='5%' nowrap>���</th>"
	Response.Write "	  <th width='9%' nowrap>¼ �� ��</th>"
	Response.Write "	  <th width='9%' nowrap>��������</th>"
	Response.Write "	</tr>"
	If Request("sortid") <> "" Then
		SQL = "select ClassID,ChannelID,ClassName,ChildStr from [ECCMS_Classify] where ChannelID = " & ChannelID & " And ClassID=" & Request("sortid")
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			Response.Write "Sorry��û���ҵ��κ�" & sModuleName & "���ࡣ������ѡ���˴����ϵͳ����!"
			Response.End
		Else
			s_ClassName = Rs("ClassName")
			ClassID = Rs("ClassID")
			ChildStr = Rs("ChildStr")
			sortid = Request("sortid")
		End If
		Rs.Close
	Else
		s_ClassName = "ȫ��" & sModuleName
		sortid = 0
	End If
	maxperpage = 30 '###ÿҳ��ʾ��
	Dim strListName
	If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
		Response.Write ("�����ϵͳ����!����������")
		Response.End
	End If
	If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
		CurrentPage = CInt(Request("page"))
	Else
		CurrentPage = 1
	End If
	If CLng(CurrentPage) = 0 Then CurrentPage = 1
	If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("queryopt")) = 1 Then
			findword = "A.title like '%" & keyword & "%'"
		ElseIf CInt(Request("queryopt")) = 2 Then
			findword = "A.username like '%" & keyword & "%'"
		Else
			findword = "A.title like '%" & keyword & "%' or A.username like '%" & keyword & "%'"
		End If
		FoundSQL = findword
		s_ClassName = "��ѯ" & sModuleName
	Else
		specialID = 99999
		If Trim(Request("sortid")) <> "" Then
			FoundSQL = "A.isAccept = "& isAccept & " And A.ClassID in (" & ChildStr & ")"
		Else
			If Trim(Request("specialID")) <> "" Then
				specialID = Request("specialID")
				If Request("specialID") <> 0 Then
					FoundSQL = "A.isAccept = " & isAccept & " And specialID =" & Request("specialID")
				Else
					FoundSQL = "A.isAccept = " & isAccept & " And specialID > 0"
				End If
			Else
				FoundSQL = "A.isAccept = " & isAccept
			End If
		End If
	End If
	
	strListName = "&channelid="& ChannelID &"&sortid="& Request("sortid") &"&specialID="& Request("specialID") &"&isAccept="& Request("isAccept") &"&keyword=" & Request("keyword") 
	totalnumber = enchiasp.Execute("Select Count(ArticleID) from ECCMS_Article A where A.ChannelID = " & ChannelID & " And "& FoundSQL &"")(0)
	TotalPageNum = CLng(totalnumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < totalnumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	On Error Resume Next
	SQL = "select A.ArticleID,A.ChannelID,A.ClassID,A.SpecialID,A.title,A.ColorMode,A.FontMode,A.BriefTopic,A.isTop,A.AllHits,A.WriteTime,A.username,A.isBest,A.isAccept,C.ClassName from [ECCMS_Article] A inner join [ECCMS_Classify] C on A.ClassID=C.ClassID where A.ChannelID = " & ChannelID & " And "& FoundSQL &" order by A.isTop DESC, A.WriteTime desc ,A.ArticleID desc"
	Rs.Open SQL, Conn, 1,1
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<form name=selform method=post action="""">"
		Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
		Response.Write "<input type=hidden name=action value=''>"
		Response.Write "<tr><td align=center colspan=6 class=TableRow2>��û���ҵ��κ�" & sModuleName & "��</td></tr>"
	Else
		Rs.MoveFirst
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		If Rs.Eof Then Exit Sub
		SQL=Rs.GetRows(maxperpage)
		i=0
		Response.Write "	<tr>"
		Response.Write "	  <td colspan=6 class=TableRow2>"
		ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<form name=selform method=post action="""">"
		Response.Write "	<input type=hidden name=ChannelID value='" & ChannelID & "'>"
		Response.Write "	<input type=hidden name=action value=''>"
		For i=0 To Ubound(SQL,2)
			If Not Response.IsClientConnected Then Response.End
			If (i mod 2) = 0 Then
				strClass = "class=TableRow1"
			Else
				strClass = "class=TableRow2"
			End If
			Admin_Article_list SQL(0,i),SQL(1,i),SQL(2,i),SQL(4,i),SQL(5,i),SQL(6,i),SQL(7,i),SQL(8,i),SQL(10,i),SQL(11,i),SQL(12,i),SQL(13,i),SQL(14,i),strClass
		Next
		SQL=Null
	End If
	Rs.Close:Set Rs = Nothing
	Set Cmd = Nothing
%>
	<tr>
	  <td colspan="6" class="TableRow1">
	  <input class=Button type="button" name="chkall" value="ȫѡ" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="��ѡ" onClick="ContraSel(this.form)">
	  ����ѡ�
	  <select name="act">
		<option value="0">��ѡ�����ѡ��</option>
		<option value="����ɾ��">����ɾ��</option>
		<option value="�����ö�">�����ö�</option>
		<option value="ȡ���ö�">ȡ���ö�</option>
		<option value="�����Ƽ�">�����Ƽ�</option>
		<option value="ȡ���Ƽ�">ȡ���Ƽ�</option>
		<option value="�������">�������</option>
		<option value="ȡ�����">ȡ�����</option>
		<option value="����HTML">����HTML</option>
	  </select>
	  <input class=Button type="submit" name="Submit2" value="ִ�в���" onclick="return confirm('��ȷ��ִ�иò�����?')">
	  <input class=Button type="submit" name="Submit3" value="��������" onclick="document.selform.action.value='setting';">
	  <input class=Button type="submit" name="Submit4" value="�����ƶ�" onclick="document.selform.action.value='move';">
	  <input class=Button type="submit" name="Submit4" value="����ɾ��" onclick="document.selform.action.value='batdel';"></td>
	</tr>
	</form>
	<tr>
	  <td colspan="6" align="right" class="TableRow2"><%ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName %></td>
	</tr>
</table>
<%

End Sub

Function Admin_Article_list(ArticleID,ChannelID,ClassID,title,ColorMode,FontMode,BriefTopic,isTop,WriteTime,UserName,isBest,isAccept,ClassName,strClass)
	Response.Write "	<tr>"
	Response.Write "	  <td align=center "& strClass &"><input type=checkbox name=selArticleID value=" & ArticleID & "></td>"
	Response.Write "	  <td "& strClass &">"
	If isTop <> 0 Then
		Response.Write "<img src=""images/gotop.gif"" border=0 alt=�ö����� align=absMiddle>"
	End If

	Response.Write "[<a href=?ChannelID=" & ChannelID & "&sortid="
	Response.Write ClassID
	Response.Write ">"
	Response.Write ClassName
	Response.Write "</a>] "
	Response.Write enchiasp.ReadBriefTopic(BriefTopic)
	Response.Write "<u><a href=?action=view&ChannelID=" & ChannelID & "&ArticleID="
	Response.Write  ArticleID
	Response.Write ">"
	Response.Write enchiasp.ReadFontMode(title,ColorMode,FontMode)
	Response.Write "</a></u>" 

	If isBest <> 0 Then
		Response.Write "&nbsp;&nbsp;[<font color=blue>��</font>]"
	End If

	Response.Write "	  </td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &"><a href=?action=edit&ChannelID="& ChannelID &"&ArticleID="& ArticleID &">�༭</a> | <a href=?action=del&ChannelID="& ChannelID &"&ArticleID="& ArticleID &" onclick=""{if(confirm('"& sModuleName &"ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����"& sModuleName &"��?')){return true;}return false;}"">ɾ��</a></td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &">"

	If isAccept = 1 Then
		Response.Write "<font color=blue>��</font>"
	Else
		Response.Write "<font color=red>��</font>"
	End If

	Response.Write "	  </td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &">"& UserName &"</td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &">"

	If WriteTime >= Date Then
		Response.Write "<font color=red>"
		Response.Write enchiasp.FormatDate(WriteTime, 2)
		Response.Write "</font>"
	Else
		Response.Write enchiasp.FormatDate(WriteTime, 2)
	End If

	Response.Write "	  </td>"
	Response.Write "	</tr>"
End Function

Private Sub ArticleEdit(isEdit)
	Dim EditTitle,TitleColor
	If isEdit Then
		SQL = "select * from ECCMS_Article where ArticleID=" & Request("ArticleID")
		Set Rs = enchiasp.Execute(SQL)
		If Not ChkAdmin("AdminArticle" & ChannelID) Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
		ClassID = Rs("ClassID")
		EditTitle = "�༭" & sModuleName & ""
	Else
		EditTitle = "���" & sModuleName & ""
		ClassID = Request("ClassID")
		If Not ChkAdmin("AddArticle" & ChannelID) Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
	End If
%>

<script src='include/ArticleJuge.Js' type=text/javascript></script>
<script language= JavaScript>
function SelectPhoto(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=<%=ChannelID%>&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.ImageUrl.value=ss[0];
    //document.myform.ImageFileList.value=ss[0];
  }
}
</script>
<div onkeydown=CtrlEnter()>
<table  border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
        <tr>
          <th colspan="4"><%=EditTitle%></th>
        </tr>
		<form method=Post name="myform" action="admin_article.asp">
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""ArticleID"" value="""& Request("ArticleID") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
%>
        <tr>
          <td width="15%" align="right" nowrap class="TableRow2"><b>�������ࣺ</b></td>
          <td width="30%" class="TableRow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	sClassSelect = Replace(sClassSelect, "{ClassID=" & ClassID & "}", "selected")
	Response.Write sClassSelect
	Response.Write "</select>"
%>
		  </td>
          <td width="15%" align="right" nowrap class="TableRow2"><b>����ר�⣺</b></td>
          <td width="40%" class="TableRow1"><select name="SpecialID" id="SpecialID">
            <option value="0">��ָ��<%=sModuleName%>ר��</option>
<%
	Set RsObj = enchiasp.Execute("SELECT SpecialID,SpecialName From ECCMS_Special WHERE ChannelID = "& ChannelID &" And ChangeLink=0 ORDER BY orders")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("SpecialID") & """"
		If isEdit Then
			If Rs("SpecialID") = RsObj("SpecialID") Then Response.Write " selected"
		End If
		Response.Write ">"
		Response.Write RsObj("SpecialName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>
          </select></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>���⣺</b></td>
          <td colspan="3" class="TableRow1"><select name="BriefTopic" id="BriefTopic">
			  <%If isEdit Then%>
			<option value="0"<%If Rs("BriefTopic") = 0 Then Response.Write (" selected")%>>ѡ����</option>
			<option value="1"<%If Rs("BriefTopic") = 1 Then Response.Write (" selected")%>>[ͼ��]</option>
			<option value="2"<%If Rs("BriefTopic") = 2 Then Response.Write (" selected")%>>[��ͼ]</option>
			<option value="3"<%If Rs("BriefTopic") = 3 Then Response.Write (" selected")%>>[����]</option>
			<option value="4"<%If Rs("BriefTopic") = 4 Then Response.Write (" selected")%>>[�Ƽ�]</option>
			<option value="5"<%If Rs("BriefTopic") = 5 Then Response.Write (" selected")%>>[ע��]</option>
			<option value="6"<%If Rs("BriefTopic") = 6 Then Response.Write (" selected")%>>[ת��]</option>
			<%Else%>
            <option value="0">ѡ����</option>
			<option value="1">[ͼ��]</option>
			<option value="2">[��ͼ]</option>
			<option value="3">[����]</option>
			<option value="4">[�Ƽ�]</option>
			<option value="5">[ע��]</option>
			<option value="6">[ת��]</option>
			<%End If%>
          </select> <input name="title" type="text" id="title" size="60" value="<%If isEdit Then Response.Write Rs("title")%>"> 
          <font color=red>*</font></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>����ģʽ��</b></td>
          <td colspan="3" class="TableRow1">��ɫ��
            <select size="1" name="ColorMode">
		<option value="0">��ѡ����ɫ</option>
<%
	TitleColor = "," & enchiasp.InitTitleColor
	TitleColor = Split(TitleColor, ",")
	For i = 1 To UBound(TitleColor)
		Response.Write ("<option style=""background-color:"& TitleColor(i) &";color: "& TitleColor(i) &""" value='"& i &"'")
		If isEdit Then
			If Rs("ColorMode") = i Then Response.Write (" selected")
		End If
		Response.Write (">"& TitleColor(i) &"</option>")
	Next
%>
		</select> ���壺
		<select size="1" name="FontMode">
		<%If isEdit Then%>
		<option value="0"<%If Rs("FontMode") = 0 Then Response.Write (" selected")%>>��ѡ������</option>
		<option value="1"<%If Rs("FontMode") = 1 Then Response.Write (" selected")%>>����</option>
		<option value="2"<%If Rs("FontMode") = 2 Then Response.Write (" selected")%>>б��</option>
		<option value="3"<%If Rs("FontMode") = 3 Then Response.Write (" selected")%>>�»���</option>
		<option value="4"<%If Rs("FontMode") = 4 Then Response.Write (" selected")%>>����+б��</option>
		<option value="5"<%If Rs("FontMode") = 5 Then Response.Write (" selected")%>>����+�»���</option>
		<option value="6"<%If Rs("FontMode") = 6 Then Response.Write (" selected")%>>б��+�»���</option>
		<%Else%>
		<option value="0">��ѡ������</option>
		<option value="1">����</option>
		<option value="2">б��</option>
		<option value="3">�»���</option>
		<option value="4">����+б��</option>
		<option value="5">����+�»���</option>
		<option value="6">б��+�»���</option>
		<%End If%>
		</select></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>���<%=sModuleName%>��</b></td>
          <td colspan="3" class="TableRow1"><input name="Related" type="text" id="Related" size="60" value="<%If isEdit Then Response.Write Rs("Related")%>"> <font color=red>*</font></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>���ߣ�</b></td>
          <td colspan="3" class="TableRow1"><input name="Author" type="text" size="30" value="<%If isEdit Then Response.Write Rs("Author")%>">
		    <select name=font2 onChange="Author.value=this.value;">
			<option selected value="">ѡ������</option>
			<option value='����'>����</option>
			<option value='��վ'>��վ</option>
			<option value='����'>����</option>
			<option value='δ֪'>δ֪</option>
			<option value='<%=AdminName%>'><%=AdminName%></option>
			</select></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>��Դ��</b></td>
          <td colspan="3" class="TableRow1"><input name="ComeFrom" type="text" size="30" value="<%If isEdit Then Response.Write Rs("ComeFrom")%>">
		  	<select name=font1 onChange="ComeFrom.value=this.value;">
			<option selected value="">ѡ����Դ</option>
			<option value='��վԭ��'>��վԭ��</option>
			<option value='��վ����'>��վ����</option>
			<option value='����'>����</option>
			<option value='ת��'>ת��</option>
			<option value='<%=Replace(enchiasp.SiteUrl, "http://", "",1,-1,1)%>'><%=Replace(enchiasp.SiteUrl, "http://", "",1,-1,1)%></option>
			</select></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>���ݣ�</b><br><%=sModuleName%>���ݷ�ҳ��ǩ<br><font color=red>[page_break]</font><br>��ע���ǩ��ĸСд</td>
          <td colspan="3" class="TableRow1"><textarea name="content" id='content' style="display:none"><%If isEdit Then Response.Write Server.HTMLEncode(Rs("content"))%></textarea>
			<iframe ID='HtmlEditor1' src='../editor/editor.asp?ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='560' HEIGHT='350'></iframe></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>��ҳͼƬ��</b></td>
          <td colspan="3" class="TableRow1"><input name="ImageUrl" type="text" id="ImageUrl" size="60" value="<%If isEdit Then Response.Write Rs("ImageUrl")%>">
			<input type=hidden name=UploadFileList id=UploadFileList onchange="doChange(this,document.myform.ImageFileList)" value="<%If isEdit Then Response.Write Rs("UploadImage")%>">
			<br>ֱ�Ӵ��ϴ�ͼƬ��ѡ��
			<%
			If IsEdit Then
			Response.Write InitSelect(Rs("UploadImage"),Rs("ImageUrl"))
			Else
			%>
			<select name="ImageFileList" id="ImageFileList" onChange="ImageUrl.value=this.value;"><option value=''>��ѡ����ҳ�Ƽ�ͼƬ</option></select><%End If%>
			<input type='button' name='selectpic' value='�����ϴ�ͼƬ��ѡ��' onclick='SelectPhoto()' class=button></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>����ȼ���</b></td>
          <td class="TableRow1"><select name="UserGroup">
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & """"
		If isEdit Then
			If Rs("UserGroup") = RsObj("Grades") Then Response.Write " selected"
		Else
			If RsObj("Grades") = 0 Then Response.Write " selected"
		End If
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>
          </select></td>
          <td align="right" class="TableRow2"><b>���������</b></td>
          <td class="TableRow1"><input name="PointNum" type="text" size="10" value="<%If isEdit Then Response.Write Rs("PointNum"):Else Response.Write 0:End If%>"> 
            �������û��͹���Ա��Ч </td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>��ʼ�������</b></td>
          <td class="TableRow1"><input name="AllHits" type="text" id="AllHits" size="15" value="<%If isEdit Then Response.Write Rs("AllHits"):Else Response.Write 0%>"> 
          <font color=red>*</font></td>
          <td align="right" class="TableRow2"><b><%=sModuleName%>�Ǽ���</b></td>
          <td class="TableRow1"><select name="star">
		<%If isEdit Then%>
          	<option value=5<%If Rs("star") = 5 Then Response.Write (" selected")%>>������</option>
          	<option value=4<%If Rs("star") = 4 Then Response.Write (" selected")%>>�����</option>
          	<option value=3<%If Rs("star") = 3 Then Response.Write (" selected")%>>����</option>
		<option value=2<%If Rs("star") = 2 Then Response.Write (" selected")%>>���</option>
		<option value=1<%If Rs("star") = 1 Then Response.Write (" selected")%>>��</option>
		<%Else%>
		<option value=5>������</option>
          	<option value=4>�����</option>
          	<option value=3 selected>����</option>
		<option value=2>���</option>
		<option value=1>��</option>
		<%End If%>
          </select></td>
        </tr>
	<tr>
          <td align="right" class="TableRow2"><b>����ѡ�</b></td>
          <td class="TableRow1" colspan="3"><input name="isTop" type="checkbox" id="isTop" value="1"<%If isEdit Then:If Rs("isTop") <> 0 Then Response.Write (" checked")%>>
            <%=sModuleName%>�ö� 
            <input name="isBest" type="checkbox" id="isBest" value="1"<%If isEdit Then:If Rs("isBest") <> 0 Then Response.Write (" checked")%>> 
            <%=sModuleName%>�Ƽ�
	    <input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"<%If isEdit Then:If Rs("ForbidEssay") <> 0 Then Response.Write (" checked")%>> 
            ��ֹ��������
            <%If  ChkAdmin("AdminArticle" & ChannelID) Then

            %>
	    <input name="isAccept" type="checkbox" id="isAccept" value="1" checked> 
            ����������<font color=blue>������˺���ܷ�����</font>��
            <%
            else
            %>
             <input name="isAccept" type="checkbox" id="isAccept" value="0" > 

            <%
            end if
            %>
	    <%If isEdit Then%>
	    <input name="Update" type="checkbox" value="1"> 
            ͬʱ����<%=sModuleName%>����ʱ��<%End If%></td>
        </tr>
        <tr align="center">
          <td colspan="4" class="TableRow2">
	  <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>
	  <input type="button" name="btnSubmit" value="����<%=sModuleName%>" class=Button onclick="doSubmit();"></td>
        </tr></form>
      </table></div>
<%
	If isEdit Then
		Rs.Close
		Set Rs = Nothing
	End If
End Sub

Private Sub CheckSave()
	If Trim(Request.Form("title")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "���ⲻ��Ϊ�գ�</li>"
	End If
	If Len(Request.Form("title")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "���ⲻ�ܳ���200���ַ���</li>"
	End If
	If Trim(Request.Form("ColorMode")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������ɫ��������</li>"
	End If
	If Trim(Request.Form("FontMode")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���������������</li>"
	End If
	If Len(Request.Form("Related")) => 220 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���" & sModuleName & "���ܳ���220���ַ���</li>"
	End If
	If Trim(Request.Form("Author")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "���߲���Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("ComeFrom")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "��Դ����Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("PointNum")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>����" & sModuleName & "����ĵ�������Ϊ�գ�������������������㡣</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "�Ǽ�����Ϊ�ա�</li>"
	End If
	If Not IsNumeric(Request.Form("UserGroup")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>����" & sModuleName & "�ȼ���������</li>"
	End If
	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��һ�������Ѿ����������࣬�������" & sModuleName & "��</li>"
	End If
	If Trim(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�÷������ⲿ���ӣ��������" & sModuleName & "��</li>"
	End If
	If Trim(Request.Form("AllHits")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ʼ���������Ϊ�գ�</li>"
	End If
	If Not IsNumeric(Request("AllHits")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ʼ�����������������</li>"
	End If
	If Not IsNumeric(Request("SpecialID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ר��ID��������</li>"
	End If
	TextContent = ""
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next
	If Len(TextContent) < 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "���ݲ���Ϊ�գ�</li>"
	End If
	If CInt(Request.Form("isTop")) = 1 Then
		ArticleTop = 1
	Else
		ArticleTop = 0
	End If
	If CInt(Request.Form("isBest")) = 1 Then
		ArticleBest = 1
	Else
		ArticleBest = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If
	If CInt(Request("isAccept")) = 1 Then
		ArticleAccept = 1
	Else
		ArticleAccept = 0
	End If
End Sub

Private Sub SaveArticle()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Article where (ArticleID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("SpecialID") = Trim(Request.Form("SpecialID"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("ColorMode") = Trim(Request.Form("ColorMode"))
		Rs("FontMode") = Trim(Request.Form("FontMode"))
		'�ַ�����
		'Rs("content") = enchiasp.HTMLEncode(Html2Ubb(TextContent))
		Rs("content") = Html2Ubb(TextContent)
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Author") = Trim(Request.Form("Author"))
		Rs("ComeFrom") = Trim(Request.Form("ComeFrom"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("isTop") = ArticleTop
		Rs("AllHits") = CLng(Request.Form("AllHits"))
		Rs("DayHits") = 0
		Rs("WeekHits") = 0
		Rs("MonthHits") = 0
		Rs("HitsTime") = Now()
		Rs("WriteTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("username") = Trim(AdminName)
		Rs("isBest") = ArticleBest
		Rs("BriefTopic") = Trim(Request.Form("BriefTopic"))
		Rs("ImageUrl") = Trim(Request.Form("ImageUrl"))
		Rs("UploadImage") = Trim(Request.Form("UploadFileList"))
		Rs("UserGroup") = Trim(Request.Form("UserGroup"))
		Rs("PointNum") = Trim(Request.Form("PointNum"))
		Rs("isUpdate") = 1
		Rs("isAccept") = ArticleAccept
		Rs("ForbidEssay") = ForbidEssay
		Rs("AlphaIndex") = enchiasp.ReadAlpha(Request.Form("title"))
	Rs.update
	Rs.Close
	Rs.Open "select top 1 ArticleID from ECCMS_Article where ChannelID=" & ChannelID & " order by ArticleID desc", Conn, 1, 1
	ArticleID = Rs("ArticleID")
	Rs.Close:Set Rs = Nothing
	ClassUpdateCount Request.Form("ClassID"),1
	Call RemoveCache
	Dim url
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makenews.asp?ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&showid=0"	
		Call ScriptCreation(url,ArticleID)
		SQL = "SELECT TOP 1 ArticleID FROM ECCMS_Article WHERE ChannelID=" & ChannelID & " And isAccept <> 0 And ArticleID < " & ArticleID & " ORDER BY ArticleID DESC"
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.EOF And Rs.BOF) Then
			url = "admin_makenews.asp?ChannelID=" & ChannelID & "&ArticleID=" & Rs("ArticleID") & "&showid=0"	
			Call ScriptCreation(url,Rs("ArticleID"))
		End If
		Rs.Close
		Set Rs = Nothing
	End If
	'Succeed("<li>��ϲ��������µ�" & sModuleName & "�ɹ���</li><li><a href=admin_article.asp?action=view&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & ">����˴��鿴��" & sModuleName & "</a></li><li><a href=?action=add&ChannelID=" & ChannelID & "&classid=" & Request.Form("ClassID") & "><font color=blue>����˴��������" & sModuleName & "</font></a></li>")
Succeed("<li>��ϲ��������µ�" & sModuleName & "�ɹ���</li>")
End Sub

Private Sub ModifyArticle()
	If Not ChkAdmin("AdminArticle" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	CheckSave
	If Founderr = True Then Exit Sub
	Dim Auditing
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Article where ArticleID=" & Request("ArticleID")
	Rs.Open SQL,Conn,1,3
		Auditing = Rs("isAccept")
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("SpecialID") = Trim(Request.Form("SpecialID"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("ColorMode") = Trim(Request.Form("ColorMode"))
		Rs("FontMode") = Trim(Request.Form("FontMode"))
		Rs("content") = Html2Ubb(TextContent)
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Author") = Trim(Request.Form("Author"))
		Rs("ComeFrom") = Trim(Request.Form("ComeFrom"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("isTop") = ArticleTop
		Rs("isBest") = ArticleBest
		If CInt(Request.Form("Update")) = 1 Then Rs("WriteTime") = Now()
		Rs("AllHits") = CLng(Request.Form("AllHits"))
		Rs("BriefTopic") = Trim(Request.Form("BriefTopic"))
		Rs("ImageUrl") = Trim(Request.Form("ImageUrl"))
		Rs("UploadImage") = Trim(Request.Form("UploadFileList"))
		Rs("UserGroup") = Trim(Request.Form("UserGroup"))
		Rs("PointNum") = Trim(Request.Form("PointNum"))
		Rs("isUpdate") = 1
		Rs("isAccept") = ArticleAccept
		Rs("ForbidEssay") = ForbidEssay
		Rs("AlphaIndex") = enchiasp.ReadAlpha(Request.Form("title"))
	Rs.update
	ArticleID = Rs("ArticleID")
	If ArticleAccept = 1 And Auditing = 0 Then
		AddUserPointNum Rs("username"),1
	End If
	If ArticleAccept = 0 And Auditing = 1 Then
		AddUserPointNum Rs("username"),0
	End If
	Rs.Close:Set Rs = Nothing
	Call RemoveCache
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makenews.asp?ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&showid=0"	
		Call ScriptCreation(url,ArticleID)
	End If
	Succeed("<li>��ϲ�����޸�" & sModuleName & "�ɹ���</li><li><a href=admin_article.asp?action=view&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & ">����˴��鿴��" & sModuleName & "</a></li>")
End Sub
Private Sub ArticleView()
	Call PageTop
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
	SQL = "select ArticleID,title,content,ColorMode,FontMode,Author,ComeFrom,WriteTime,username,isAccept from ECCMS_Article where ChannelID=" & ChannelID & " And ArticleID=" & Request("ArticleID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κ�" & sModuleName & "��������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	Else
%>
<script language=javascript>
var enchiasp_fontsize=9;
var enchiasp_lineheight=12;
</script>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
	<tr>
	  <th>�鿴<%=sModuleName%></th>
	</tr>
	<tr>
	  <td align="center" class="TableRow2"><a href=?action=edit&ChannelID=<%=ChannelID%>&ArticleID=<%=Rs("ArticleID")%>><font size=4><%=enchiasp.ReadFontMode(Rs("title"),Rs("ColorMode"),Rs("FontMode"))%></font></a></td>
	</tr>
	<tr>
	  <td align="center" class="TableRow1"><b>���ߣ�</b><%=Rs("Author")%>&nbsp;&nbsp;<b>��Դ��</b><%=Rs("ComeFrom")%>&nbsp;&nbsp;<b>����ʱ�䣺</b><%=Rs("WriteTime")%>&nbsp;&nbsp;
	  <b>�� �� �ˣ�</b> <font color=blue><%=Rs("username")%></font>&nbsp;&nbsp;
	  <b>���״̬��</b>
	  <%If Rs("isAccept") > 0 Then%>
	  <font color=blue>�����</font>
	  <%Else%>
	  <font color=red>δ���</font>
	  <%End If%>
	  </td>
	</tr>
	<tr>
	  <td class="TableRow1"><p align="right"><a style="CURSOR: hand; POSITION: relative" onclick="if(enchiasp_fontsize&gt;8){enchiaspContentLabel.style.fontSize=(--enchiasp_fontsize)+&quot;pt&quot;;enchiaspContentLabel.style.lineHeight=(--enchiasp_lineheight)+&quot;pt&quot;;}" title="��С����"><img src="images/1.gif" border="0" width="15" height="15"><font color="#FF6600">��С����</font></a> 
                    <a style="CURSOR: hand; POSITION: relative" onclick="if(enchiasp_fontsize&lt;64){enchiaspContentLabel.style.fontSize=(++enchiasp_fontsize)+&quot;pt&quot;;enchiaspContentLabel.style.lineHeight=(++enchiasp_lineheight)+&quot;pt&quot;;}" title="��������"><img src="images/2.gif" border="0" width="15" height="15"><font color="#FF6600">��������</font></a></p>
					<div id="enchiaspContentLabel"><%=Replace(enchiasp.ReadContent(Rs("content")), "[page_break]", "", 1, -1, 1)%></div></td>
	</tr>
	<tr>
	  <td class="TableRow1">��һƪ<%=sModuleName%>��<%=FrontArticle(Rs("ArticleID"))%>
	  <br>��һƪ<%=sModuleName%>��<%=NextArticle(Rs("ArticleID"))%></td>
	</tr>
	<tr>
	  <td align="center" class="TableRow2">
	  <input type="button" onclick="{if(confirm('��ȷ��Ҫɾ����������?')){location.href='?action=del&ChannelID=<%=ChannelID%>&ArticleID=<%=Rs("ArticleID")%>';return true;}return false;}" value="ɾ������" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="���ض���" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&ChannelID=<%=ChannelID%>&ArticleID=<%=Rs("ArticleID")%>'" value="�༭<%=sModuleName%>" class=button>&nbsp;&nbsp;
	   [<a href="?act=�������&ChannelID=<%=ChannelID%>&selArticleID=<%=Rs("ArticleID")%>" onclick="return confirm('��ȷ��ִ����˲�����?');" ><font color=blue>ֱ�����</font></a>]</td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub
Private Function FrontArticle(ArticleID)
	Dim Rss, SQL
	SQL = "select Top 1 ArticleID,classid,title from ECCMS_Article where ChannelID=" & ChannelID & " And isAccept <> 0 And ArticleID < " & ArticleID & " order by ArticleID desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		FrontArticle = "�Ѿ�û����"
	Else
		FrontArticle = "<a href=admin_article.asp?action=view&ChannelID=" & ChannelID & "&ArticleID=" & Rss("ArticleID") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function NextArticle(ArticleID)
	Dim Rss, SQL
	SQL = "select Top 1 ArticleID,classid,title from ECCMS_Article where ChannelID=" & ChannelID & " And isAccept <> 0 And ArticleID > " & ArticleID & " order by ArticleID asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		NextArticle = "�Ѿ�û����"
	Else
		NextArticle = "<a href=admin_article.asp?action=view&ChannelID=" & ChannelID & "&ArticleID=" & Rss("ArticleID") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function

Private Sub BatCreateHtml()
	Dim AllArticleID,url
	Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
	Response.Write "<ol>"
	AllArticleID = Split(selArticleID, ",")
	For i = 0 To UBound(AllArticleID)
		ArticleID = CLng(AllArticleID(i))
		url = "admin_makenews.asp?ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&showid=1"	
		Call ScriptCreation(url,ArticleID)
	Next
	Response.Write "</ol>"
	OutHintScript("��ʼ����HTML,����" & i & "��HTMLҳ����Ҫ���ɣ�")
End Sub

Private Function ClassUpdateCount(sortid,stype)
	Dim rscount,Parentstr
	On Error Resume Next
	Set rscount = enchiasp.Execute("SELECT ClassID,Parentstr FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID=" & CLng(sortid))
	If Not (rscount.BOF And rscount.EOF) Then
		Parentstr = rscount("Parentstr") &","& rscount("ClassID")
		If CInt(stype) = 1 Then
			enchiasp.Execute ("UPDATE [ECCMS_Classify] SET ShowCount=ShowCount+1,isUpdate=1 WHERE ChannelID = "& ChannelID &" And ClassID in (" & Parentstr & ")")
		Else
			enchiasp.Execute ("UPDATE [ECCMS_Classify] SET ShowCount=ShowCount-1,isUpdate=1 WHERE ChannelID = "& ChannelID &" And ClassID in (" & Parentstr & ")")
		End If
	End If
	Set rscount = Nothing
End Function

Private Sub ArticleDel()
	If Request("ArticleID") = "" Then
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT ArticleID,classid,username,HtmlFileDate FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ArticleID=" & Request("ArticleID"))
	If Not(Rs.BOF And Rs.EOF) Then
		ClassUpdateCount Rs("classid"),0
		AddUserPointNum Rs("username"),0
		DeleteHtmlFile Rs("classid"),Rs("ArticleID"),Rs("HtmlFileDate")
	End If
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute("Delete From ECCMS_Article Where ChannelID = "& ChannelID &" And ArticleID = " & Request("ArticleID"))
	enchiasp.Execute ("delete from ECCMS_Comment where ChannelID = "& ChannelID &" And PostID = " & Request("ArticleID"))
	Call RemoveCache
	Response.redirect ("admin_article.asp?ChannelID=" & ChannelID)
End Sub

Private Sub batdel()
	Set Rs = enchiasp.Execute("SELECT ArticleID,classid,username,HtmlFileDate FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ArticleID in (" & selArticleID & ")")
	Do While Not Rs.EOF
		ClassUpdateCount Rs("classid"),0
		AddUserPointNum Rs("username"),0
		DeleteHtmlFile Rs("classid"),Rs("ArticleID"),Rs("HtmlFileDate")
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("delete from ECCMS_Article where ArticleID in (" & selArticleID & ")")
	enchiasp.Execute ("delete from ECCMS_Comment where ChannelID = "& ChannelID &" And PostID in (" & selArticleID & ")")
	Call RemoveCache
	OutHintScript ("����ɾ�������ɹ���")
End Sub

Private Sub isCommend()
	enchiasp.Execute ("update ECCMS_Article set isBest=1 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noCommend()
	enchiasp.Execute ("update ECCMS_Article set isBest=0 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub isTop()
	enchiasp.Execute ("update ECCMS_Article set isTop=1 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noTop()
	enchiasp.Execute ("update ECCMS_Article set isTop=0 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub BatAccept()
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ArticleID in (" & selArticleID & ")")
	Do While Not Rs.EOF
		AddUserPointNum Rs("username"),1
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("update ECCMS_Article set isAccept=1 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Sub NotAccept()
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ArticleID in (" & selArticleID & ")")
	Do While Not Rs.EOF
		AddUserPointNum Rs("username"),0
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("update ECCMS_Article set isAccept=0 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Function AddUserPointNum(username,stype)
	On Error Resume Next
	Dim rsuser,GroupSetting,userpoint
	Set rsuser = enchiasp.Execute("SELECT userid,UserGrade,userpoint FROM ECCMS_User WHERE username='"& username &"'")
	If Not(rsuser.BOF And rsuser.EOF) Then
		GroupSetting = Split(enchiasp.UserGroupSetting(rsuser("UserGrade")), "|||")(9)
		If CInt(stype) = 1 Then
			userpoint = CLng(rsuser("userpoint") + GroupSetting)
			enchiasp.Execute ("UPDATE ECCMS_User SET userpoint="& userpoint &",experience=experience+2,charm=charm+1 WHERE userid="& rsuser("userid"))
		Else
			userpoint = CLng(rsuser("userpoint") - GroupSetting)
			enchiasp.Execute ("UPDATE ECCMS_User SET userpoint="& userpoint &",experience=experience-2,charm=charm-1 WHERE userid="& rsuser("userid"))
		End If
	End If
	Set rsuser = Nothing
End Function

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

Private Sub BatchSetting()
	If Not ChkAdmin("AdminArticle" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=4>" & sModuleName & "��������</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=saveset>"
	Response.Write "	<input type=hidden name=ChannelID value='"& ChannelID &"'>"
	Response.Write "	<tr>"
	Response.Write "		<td width=""20%"" rowspan=""14"" class=tablerow2 valign=""top"" id=choose2 style=""display:none""><b>��ѡ��" & sModuleName & "����</b><br>"
	Response.Write "<select name=""ClassID"" size='2' multiple style='height:330px;width:180px;'>"
	Dim strSelectClass,re
	strSelectClass = enchiasp.LoadSelectClass(ChannelID)
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	Re.Pattern="(value=)(.*)("" )"
	strSelectClass = Re.Replace(strSelectClass,"")
	Re.Pattern="({ClassID=)(.*)(}>)"
	strSelectClass = Re.Replace(strSelectClass,"value=""$2"">")
	Response.Write strSelectClass
	Set Re = Nothing
	Response.Write "<option value=""-1"">ָ�����з���</option>"
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "		<td class=tablerow1 colspan=2 align=right><b>����ѡ��</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "		<input type=radio name=choose value='0' checked onClick=""choose1.style.display='';choose2.style.display='none';""> ��" & sModuleName & "ID&nbsp;&nbsp;"
	Response.Write "		<input type=radio name=choose value='1' onClick=""choose2.style.display='';choose1.style.display='none';""> ��" & sModuleName & "����</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=choose1>"
	Response.Write "		<td class=tablerow1 colspan=2 align=right><b>" & sModuleName & "ID��</b>���ID���á�,���ֿ�</td>"
	Response.Write "		<td class=tablerow1><input type=""text"" name=""ArticleID"" size=70 value='"& Request("selArticleID") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 width=""15%"" align=right><b>���" & sModuleName & "��</b></td>"
	Response.Write "		<td class=tablerow1 width=""5%"" align=center><input type=checkbox name=selRelated value='1'></td>"
	Response.Write "		<td class=tablerow1 width=""60%""><input name=Related type=text size=60></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "���ߣ�</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selAuthor value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=Author type=text size=20>"
	Response.Write "		<select name=font2 onChange=""Author.value=this.value;"">"
	Response.Write "		<option selected value=''>ѡ������</option>"
	Response.Write "		<option value='����'>����</option>"
	Response.Write "		<option value='��վ'>��վ</option>"
	Response.Write "		<option value='����'>����</option>"
	Response.Write "		<option value='δ֪'>δ֪</option>"
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "��Դ��</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selComeFrom value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=ComeFrom type=text size=20>"
	Response.Write "		<select name=font1 onChange=""ComeFrom.value=this.value;"">"
	Response.Write "		<option selected value=''>ѡ����Դ</option>"
	Response.Write "		<option value='��վԭ��'>��վԭ��</option>"
	Response.Write "		<option value='��վ����'>��վ����</option>"
	Response.Write "		<option value='����'>����</option>"
	Response.Write "		<option value='ת��'>ת��</option>"
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>���������</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selPointNum value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=PointNum type=text size=10 value=0></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>�������</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selAllHits value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=AllHits type=text size=10 value=0></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>�����ȼ���</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selUserGroup value='1'></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "		<select name=UserGroup size='1'>"
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup order by Groupid")
	Do While Not RsObj.EOF
		Response.Write "	<option value=""" & RsObj("Grades") & """"
		If RsObj("Grades") = 0 Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "�Ǽ���</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selstar value='1'></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "	<select name=star>"
	Response.Write "		<option value=5>������</option>"
	Response.Write "		<option value=4>�����</option>"
	Response.Write "		<option value=3 selected>����</option>"
	Response.Write "		<option value=2>���</option>"
	Response.Write "		<option value=1>��</option>"
	Response.Write "	</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "�ö���</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selTop value='1'></td>"
	Response.Write "		<td class=tablerow1><input type=radio name=istop value='0' checked> ��&nbsp;&nbsp;<input type=radio name=istop value='1'> ��</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "�Ƽ���</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selBest value='1'></td>"
	Response.Write "		<td class=tablerow1><input type=radio name=isbest value='0' checked> ��&nbsp;&nbsp;<input type=radio name=isbest value='1'> ��</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>��ֹ���ۣ�</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selForbidEssay value='1'></td>"
	Response.Write "		<td class=tablerow1><input type=radio name=ForbidEssay value='0' checked> ��&nbsp;&nbsp;<input type=radio name=ForbidEssay value='1'> ��</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=3><b>˵����</b>��Ҫ�����޸�ĳ�����Ե�ֵ������ѡ�������ĸ�ѡ��Ȼ�����趨����ֵ��"
	Response.Write " <a href=?action=reset&ChannelID="& ChannelID & " onclick=""return confirm('��ȷ��Ҫ��������ʱ����?')"">����ʱ��</a></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow1 colspan=3><input type=""button"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" name=""B1"" class=Button>&nbsp;&nbsp;"
	Response.Write "		<input type=submit name=submit2 value=""ȷ������"" class=Button onclick=""return confirm('��ȷ��ִ������������?')"">"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub

Private Sub BatchMove()
	If Not ChkAdmin("AdminArticle" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=3>" & sModuleName & "�����ƶ�</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=savemove>"
	Response.Write "	<input type=hidden name=ChannelID value='"& ChannelID &"'>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=3><input type=radio name=Appointed value='0' checked>"
	Response.Write " <b>ָ��" & sModuleName & "ID��</b> <input type=""text"" name=""ArticleID"" size=80 value='"& Request("selArticleID") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 width=""35%""><input type=radio name=Appointed value='1'> <b>ָ��" & sModuleName & "���ࣺ</b></td>"
	Response.Write "		<td class=tablerow1 width=""10%""></td>"
	Response.Write "		<td class=tablerow1 width=""55%""><b>" & sModuleName & "Ŀ����ࣺ</b><font color=red>������ָ���ⲿ���ࣩ</font></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name='ClassID' size='2' multiple style='height:350px;width:260px;'>"
	Dim strSelectClass,re
	strSelectClass = enchiasp.LoadSelectClass(ChannelID)
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	Re.Pattern="(value=)(.*)("" )"
	strSelectClass = Re.Replace(strSelectClass,"")
	Re.Pattern="({ClassID=)(.*)(}>)"
	strSelectClass = Re.Replace(strSelectClass,"value=""$2"">")
	Response.Write strSelectClass
	Set Re = Nothing
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "		<td class=tablerow1 align=center noWrap>�ƶ�����</td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name='tClassID' size='2' style='height:350px;width:260px;'>"
	Response.Write strSelectClass
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow1 colspan=3><input type=""button"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" name=""B1"" class=Button>&nbsp;&nbsp;"
	Response.Write "		<input type=submit name=submit2 value=""�����ƶ�"" class=Button onclick=""return confirm('��ȷ��ִ�������ƶ���?')"">"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub
Private Sub SaveMove()
	If Trim(Request.Form("tClassID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ��Ŀ����ࡣ</li>"
		Exit Sub
	End If
	If Trim(Request.Form("tClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ƶ����ⲿ���ࡣ</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT child FROM ECCMS_Classify WHERE TurnLink=0 And ChannelID = "& ChannelID &" And ClassID="& CLng(Request.Form("tClassID")))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ�����������ƶ����ⲿ���ࡣ</li>"
		Set Rs = Nothing
		Exit Sub
	Else
		If Rs("child") > 0 Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>�˷��໹���ӷ��࣬��ѡ���ӷ������ƶ���</li>"
			Set Rs = Nothing
			Exit Sub
		End If
	End If
	Set Rs = Nothing
	If CInt(Request.Form("Appointed")) = 0 Then
		If Trim(Request.Form("ArticleID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ��" & sModuleName & "ID��</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_Article SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ArticleID in ("& Request("ArticleID") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ����ࡣ</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_Article SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>��ϲ���������ƶ���ɡ�</li>")
End Sub

Private Sub BatcDelete()
	If Not ChkAdmin("AdminArticle" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th>" & sModuleName & "����ɾ��</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=alldel>"
	Response.Write "	<input type=hidden name=ChannelID value='"& ChannelID &"'>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=3><input type=radio name=Appointed value='0' checked onClick=""choose1.style.display='';choose2.style.display='none';"">"
	Response.Write " <b>ָ��" & sModuleName & "ID��</b> "
	Response.Write "<input type=radio name=Appointed value='1' onClick=""choose2.style.display='';choose1.style.display='none';""> <b>ָ��" & sModuleName & "����</b>"
	Response.Write "<input type=radio name=Appointed value='2'> <b>ɾ������" & sModuleName & "</b>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=choose1>"
	Response.Write "		<td class=tablerow1><b>����ID��</b><input type=""text"" name=""ArticleID"" size=80 value='"& Request("selArticleID") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=choose2 style=""display:none"">"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name='ClassID' size='2' multiple style='height:350px;width:260px;'>"
	Dim strSelectClass,re
	strSelectClass = enchiasp.LoadSelectClass(ChannelID)
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	Re.Pattern="(value=)(.*)("" )"
	strSelectClass = Re.Replace(strSelectClass,"")
	Re.Pattern="({ClassID=)(.*)(}>)"
	strSelectClass = Re.Replace(strSelectClass,"value=""$2"">")
	Response.Write strSelectClass
	Set Re = Nothing
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow1><input type=""button"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" name=""B1"" class=Button>&nbsp;&nbsp;"
	Response.Write "		<input type=submit name=submit2 value=""����ɾ��"" class=Button onclick=""return confirm('��ȷ��ִ������ɾ��������?')"">"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub
Private Sub AllDelArticle()
	On Error Resume Next
	If CInt(Request.Form("Appointed")) = 1 Then
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ����ࡣ</li>"
			Exit Sub
		End If
		Conn.Execute("DELETE ECCMS_Comment FROM ECCMS_Article A INNER JOIN ECCMS_Comment C ON C.PostID=A.ArticleID WHERE A.ChannelID = "& ChannelID &" And A.ClassID IN (" & Request("ClassID") & ")")
		Conn.Execute("DELETE FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ClassID IN (" & Request("ClassID") & ")")
	ElseIf CInt(Request.Form("Appointed")) = 2 Then
		Conn.Execute ("DELETE FROM ECCMS_Article WHERE ChannelID = "& ChannelID)
		Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID)
	Else
		If Trim(Request.Form("ArticleID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ��" & sModuleName & "ID��</li>"
			Exit Sub
		End If
		Conn.Execute ("DELETE FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ArticleID IN (" & Request("ArticleID") & ")")
		Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID IN (" & Request("ArticleID") & ")")
		
	End If
	Call RemoveCache
	Succeed("<li>����ɾ���ɹ���</li>")
End Sub

Private Sub SaveSetting()
	If Founderr = True Then Exit Sub
	Dim strTempValue
	strTempValue = ""
	If Trim(Request.Form("selRelated")) <> "" Then strTempValue = strTempValue & "Related='"& enchiasp.ChkFormStr(Request.Form("Related")) &"',"
	If Trim(Request.Form("selAuthor")) <> "" Then strTempValue = strTempValue & "Author='"& enchiasp.ChkFormStr(Request.Form("Author")) &"',"
	If Trim(Request.Form("selComeFrom")) <> "" Then strTempValue = strTempValue & "ComeFrom='"& enchiasp.ChkFormStr(Request.Form("ComeFrom")) &"',"
	If Trim(Request.Form("selPointNum")) <> "" Then strTempValue = strTempValue & "PointNum="& CLng(Request.Form("PointNum")) &","
	If Trim(Request.Form("selAllHits")) <> "" Then strTempValue = strTempValue & "AllHits="& CLng(Request.Form("AllHits")) &","
	If Trim(Request.Form("selUserGroup")) <> "" Then strTempValue = strTempValue & "UserGroup="& CInt(Request.Form("UserGroup")) &","
	If Trim(Request.Form("selstar")) <> "" Then strTempValue = strTempValue & "star="& CInt(Request.Form("star")) &","
	If Trim(Request.Form("selTop")) <> "" Then strTempValue = strTempValue & "istop="& CInt(Request.Form("istop")) &","
	If Trim(Request.Form("selBest")) <> "" Then strTempValue = strTempValue & "isbest="& CInt(Request.Form("isbest")) &","
	If Trim(Request.Form("selForbidEssay")) <> "" Then strTempValue = strTempValue & "ForbidEssay="& CInt(Request.Form("ForbidEssay")) &","
	If Trim(strTempValue) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ��Ҫ���õĲ�����</li>"
		Exit Sub
	Else
		strTempValue = Replace(Left(strTempValue,Len(strTempValue)-1), " ", "")
	End If
	If CInt(Request.Form("choose")) = 0 Then
		If Trim(Request.Form("ArticleID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ��" & sModuleName & "ID��</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_Article SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ArticleID in ("& Request("ArticleID") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ����ࡣ</li>"
			Exit Sub
		ElseIf Trim(Request.Form("ClassID")) = "-1" Then
			SQL = "UPDATE ECCMS_Article SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID
		Else
			SQL = "UPDATE ECCMS_Article SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>��ϲ��������������ɡ�</li>")
End Sub

Private Sub ResetDateTime()
	Server.ScriptTimeOut = 9999
	Response.Write "<br><table width='400' align=center border=0 cellspacing=1 cellpadding=1>" & vbCrLf
	Response.Write "<tr> " & vbCrLf
	Response.Write "<td style=""border: 1px #384780 solid ;background-color: #FFFFFF;"">" & vbCrLf
	Response.Write "<table width=0 id=table2 name=table2 border=0 cellspacing=0 cellpadding=0>" & vbCrLf
	Response.Write "<tr height=12><td bgcolor=#36D91A>" & vbCrLf
	Response.Write "</td></tr></table></td></tr><tr> " & vbCrLf
	Response.Write "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span> <span style=""font-size:9pt"">%</span></td></tr>" & vbCrLf
	Response.Write "</table><br>" & vbCrLf
	Response.Flush
	Dim totalrec,WriteTime,page_count,pagelist
	i = 0
	page_count = 0
	totalrec = enchiasp.Execute("SELECT COUNT(ArticleID) FROM [ECCMS_Article] WHERE ChannelID = "& ChannelID &" And isAccept>0")(0)
	Set Rs = enchiasp.Execute("SELECT ArticleID,WriteTime FROM [ECCMS_Article] WHERE ChannelID = "& ChannelID &" And isAccept>0 ORDER BY WriteTime DESC")
	If Not (Rs.BOF And Rs.EOF) Then
		SQL=Rs.GetRows(-1)
		For pagelist=0 To Ubound(SQL,2)
			If Not Response.IsClientConnected Then Response.End
			Response.Write "<script>"
			Response.Write "table2.style.width=" & Fix((page_count / totalrec) * 400) & ";"
			Response.Write "txt2.innerHTML=""��ɣ�" & FormatNumber(page_count / totalrec * 100, 2, -1) & """;"
			Response.Write "</script>" & vbCrLf
			Response.Flush
			WriteTime = DateAdd("s", -i, SQL(1,pagelist))
			enchiasp.Execute ("UPDATE ECCMS_Article SET WriteTime='" & WriteTime & "' WHERE ArticleID="& SQL(0,pagelist))
			i = i + 5
			page_count = page_count + 1
		Next
		SQL=Null
	End If
	Set Rs = Nothing
	Response.Write "<script>table2.style.width=400;txt2.innerHTML=""��ɣ�100"";</script>"
End Sub

Private Sub RemoveCache()
	enchiasp.DelCahe "RenewStatistics"
	enchiasp.DelCahe "TotalStatistics"
End Sub
%>

