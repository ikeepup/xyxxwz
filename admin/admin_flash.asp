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
Dim Action,isEdit,Flag,DefaultShowMode
Dim i,ClassID,RsObj,flashid,findword,keyword,strClass
Dim TextContent,FlashTop,FlashBest,ForbidEssay
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Dim s_ClassName,ChildStr,FoundSQL,isAccept,selflashid
Dim FlashAccept,Auditing

ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID = 0 Then ChannelID = 5
If ChannelID = 5 Then
	DefaultShowMode = 1		'-- Ĭ����ʾģʽ
Else
	DefaultShowMode = 2		'-- Ĭ����ʾģʽ
End If

Flag = sChannelDir & ChannelID
If Request("isAccept") <> "" Then
	isAccept = 0
Else
	isAccept = 1
End If
Action = LCase(Request("action"))
If Not ChkAdmin(Flag) Then
	Server.Transfer("showerr.asp")
	Response.End
End If

Select Case Trim(Action)
Case "save"
	Call SaveFlash
Case "modify"
	Call ModifyFlash
Case "add"
	isEdit = False
	Call FlashEdit(isEdit)
Case "edit"
	isEdit = True
	Call FlashEdit(isEdit)
Case "del"
	Call FlashDel
Case "view"
	Call FlashView
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
Case "batdel"
	Call PageTop
	Call BatcDelete
Case "alldel"
	Call AllDelFlash
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
	Response.Write "	  <th colspan=2>" & sModuleName & "����ѡ��</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr><form method=Post name=myform action='admin_flash.asp' onSubmit='return JugeQuery(this);'>"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "<td class=TableRow1>������"
	Response.Write "	  <input name=keyword type=text size=20>"
	Response.Write "	  ������"
	Response.Write "	  <select name=queryopt>"
	Response.Write "		<option value='1' selected>" & sModuleName & "����</option>"
	Response.Write "		<option value='2'>�� �� ��</option>"
	Response.Write "		<option value='3'>��������</option>"
	Response.Write "	  </select> <input type=submit name=Submit value='��ʼ��ѯ' class=Button></td>"
	Response.Write "	  <td class=TableRow1>" & sModuleName & "������"
	Dim srtClassMenu
	Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Response.Write "<option value='admin_flash.asp?ChannelID=" & ChannelID & "'>��ȫ��" & sModuleName & "�б��</option>" & vbCrLf
	srtClassMenu = enchiasp.ClassJumpMenu(ChannelID)
	srtClassMenu = Replace(srtClassMenu, "{ClassID=" & Request("sortid") & "}", "selected")
	Response.Write srtClassMenu
	Response.Write "</select>"
	Response.Write "	  </td></form>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td colspan=2 class=TableRow2><strong>����ѡ�</strong> <a href='admin_flash.asp?ChannelID=" & ChannelID & "'>������ҳ</a> | "
	Response.Write "	  <a href='admin_flash.asp?action=add&ChannelID=" & ChannelID & "'>���" & sModuleName & "</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "&action=add'>���" & sModuleName & "����</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "'>" & sModuleName & "�������</a></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub
Private Sub showmain()
	If Not IsEmpty(Request("selflashid")) Then
		selflashid = Request("selflashid")
		Select Case enchiasp.CheckStr(Request("act"))
		Case "����ɾ��":Call batdel
		Case "�����ƶ�":Call batmove
		Case "����ʱ��":Call upindate
		Case "�����Ƽ�":Call isCommend
		Case "ȡ���Ƽ�":Call noCommend
		Case "�����ö�":Call isTop
		Case "ȡ���ö�":Call noTop
		Case "�������":Call BatAccept
		Case "ȡ�����":Call NotAccept
		Case "����HTML":Call BatCreateHtml
		Case Else
			Response.Write "��Ч������"
		End Select
	End If
	Call PageTop
	Dim strListName
	Dim specialID,sortid,Cmd,child
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>ѡ��</th>"
	Response.Write "	  <th width='68%'>" & sModuleName & "����</th>"
	Response.Write "	  <th width='9%' nowrap>�������</th>"
	Response.Write "	  <th width='9%' nowrap>¼ �� ��</th>"
	Response.Write "	  <th width='9%' nowrap>��������</th>"
	Response.Write "	</tr>"
	strListName = "&channelid="& ChannelID &"&sortid="& Request("sortid") &"&specialID="& Request("specialID") &"&isAccept="& Request("isAccept") &"&keyword=" & Request("keyword") 
	If Request("sortid") <> "" Then
		SQL = "select ClassID,ChannelID,ClassName,child,ChildStr from [ECCMS_Classify] where ChannelID = " & ChannelID & " And ClassID=" & Request("sortid")
		Set Rs = enchiasp.Execute(SQL)
		If Rs.bof And Rs.EOF Then
			Response.Write "Sorry��û���ҵ��κ�" & sModuleName & "���ࡣ������ѡ���˴����ϵͳ����!"
			Response.End
		Else
			s_ClassName = Rs("ClassName")
			ClassID = Rs("ClassID")
			child = Rs("child")
			ChildStr = Rs("ChildStr")
			sortid = CLng(Request("sortid"))
		End If
		Rs.Close
	Else
		s_ClassName = "ȫ��" & sModuleName
		sortid = 0
		child = 0
	End If
	maxperpage = 30 '###ÿҳ��ʾ��
	
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
			FoundSQL = "A.isAccept="& isAccept & " And A.ClassID in (" & ChildStr & ")"
		Else
			If Trim(Request("specialID")) <> "" Then
				specialID = CLng(Request("specialID"))
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
	TotalNumber = enchiasp.Execute("Select Count(flashid) from ECCMS_FlashList A where A.ChannelID = " & ChannelID & " And "& FoundSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select A.*,C.ClassName from [ECCMS_FlashList] A inner join [ECCMS_Classify] C on A.ClassID=C.ClassID where A.ChannelID = " & ChannelID & " And "& FoundSQL &" order by A.isTop desc, A.addTime desc ,A.flashid desc"
	Rs.Open SQL, Conn, 1, 1
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=5 class=TableRow2>��û���ҵ��κ�" & sModuleName & "��</td></tr>"
	Else
		Rs.MoveFirst
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0

	Response.Write "	<tr>"
	Response.Write "	  <td colspan=""5"" class=""TableRow2"">"
	ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<form name=selform method=post action="""">"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "<input type=hidden name=action value=''>"
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
		Response.Write "	<tr>"
		Response.Write "	  <td align=center " & strClass & "><input type=checkbox name=selflashid value=" & Rs("flashid") & "></td>"
		Response.Write "	  <td " & strClass & ">"
		If Rs("isTop") <> 0 Then
			Response.Write "<img src=""images/istop.gif"" width=15 height=17 border=0 alt=�ö�����>"
		End If

		Response.Write "[<a href=?ChannelID=" & Rs("ChannelID") & "&sortid="
		Response.Write Rs("ClassID")
		Response.Write ">"
		Response.Write Rs("ClassName")
		Response.Write "</a>] "
		Response.Write "<a href=?action=view&ChannelID=" & Rs("ChannelID") & "&flashid="
		Response.Write Rs("flashid")
		Response.Write ">"
		Response.Write enchiasp.ReadFontMode(Rs("title"),Rs("ColorMode"),Rs("FontMode"))
		Response.Write "</a>" 

		If Rs("isBest") <> 0 Then
			Response.Write "&nbsp;&nbsp;<font color=blue>��</font>"
		End If
%>
	  </td>
	  <td align="center" nowrap <%=strClass%>><a href=?action=edit&ChannelID=<%=Rs("ChannelID")%>&flashid=<%=Rs("flashid")%>>�༭</a> | <a href=?action=del&ChannelID=<%=Rs("ChannelID")%>&flashid=<%=Rs("flashid")%> onclick="{if(confirm('����ɾ���󽫲��ָܻ�����ȷ��Ҫɾ���ö�����?')){return true;}return false;}">ɾ��</a></td>
	  <td align="center" nowrap <%=strClass%>><%=Rs("UserName")%></td>
	  <td align="center" nowrap <%=strClass%>>
<%
		If Rs("addTime") >= Date Then
			Response.Write "<font color=red>"
			Response.Write enchiasp.FormatDate(Rs("addTime"), 2)
			Response.Write "</font>"
		Else
			Response.Write enchiasp.FormatDate(Rs("addTime"), 2)
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
	Set Cmd = Nothing
	Rs.Close:Set Rs = Nothing
%>
	<tr>
	  <td colspan="5" class="TableRow1">
	  <input class=Button type="button" name="chkall" value="ȫѡ" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="��ѡ" onClick="ContraSel(this.form)">
	  ����ѡ�
	  <select name="act">
		<option value="0">��ѡ�����ѡ��</option>
		<option value="����ɾ��">����ɾ��</option>
		<option value="�����ö�">�����ö�</option>
		<option value="ȡ���ö�">ȡ���ö�</option>
		<option value="�����Ƽ�">�����Ƽ�</option>
		<option value="ȡ���Ƽ�">ȡ���Ƽ�</option>
		<option value="����ʱ��">����ʱ��</option>
		<option value="����HTML">����HTML</option>
	  </select>
	  <input class=Button type="submit" name="Submit2" value="ִ�в���" onclick="return confirm('��ȷ��ִ�иò�����?');">
	  <input class=Button type="submit" name="Submit3" value="��������" onclick="document.selform.action.value='setting';">
	  <input class=Button type="submit" name="Submit4" value="�����ƶ�" onclick="document.selform.action.value='move';">
	  <input class=Button type="submit" name="Submit4" value="����ɾ��" onclick="document.selform.action.value='batdel';"></td>
	</tr>
	</form>
	<tr>
	  <td colspan="5" align="right" class="TableRow2"><%
	  ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName
	  %></td>
	</tr>
</table>
<%
End Sub

Private Sub FlashEdit(isEdit)
	Dim EditTitle,TitleColor,downid
	If isEdit Then
		SQL = "select * from ECCMS_FlashList where flashid=" & Request("flashid")
		Set Rs = enchiasp.Execute(SQL)
		ClassID = Rs("ClassID")
		EditTitle = "�༭" & sModuleName
		downid = Rs("downid")
	Else
		EditTitle = "���" & sModuleName
		ClassID = Request("ClassID")
		downid = 0
	End If
%>
<script src='include/FlashJuge.Js' type=text/javascript></script>
<script language= JavaScript>
function SelectPhoto(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=<%=ChannelID%>&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.miniature.value=ss[0];
  }
}
function SelectFile(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=<%=ChannelID%>&UploadDir=UploadFile', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.showurl.value=ss[0];
    document.myform.filesize.value=ss[1];
  }
}
</script>
<div onkeydown=CtrlEnter()>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
  <tr>
    <th colspan="4"><%=EditTitle%></th>
  </tr>
 	<form method=Post name="myform" action="admin_flash.asp" onSubmit='return CheckForm(this);'>
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""flashid"" value="""& Request("flashid") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
%>
  <tr>
    <td width="12%" align="right" nowrap class="TableRow2"><strong><%=sModuleName%>���ࣺ</strong></td>
    <td width="38%" class="TableRow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	sClassSelect = Replace(sClassSelect, "{ClassID=" & ClassID & "}", "selected")
	Response.Write sClassSelect
	Response.Write "</select>"
%>
    </td>
    <td width="12%" align="right" class="TableRow2"><strong></strong></td>
    <td width="38%" class="TableRow1"></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>���ƣ�</strong></td>
    <td class="TableRow1"><input name="title" type="text" id="title" size="35" value="<%If isEdit Then Response.Write Rs("title")%>"> 
      <span class="style1">* </span></td>
    <td align="right" class="TableRow2"><strong>�������壺</strong></td>
    <td class="TableRow1">
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
		</select>
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
    <td height="130" align="right" class="TableRow2"><strong><%=sModuleName%>��С��</strong></td>
    <td class="TableRow1">
<%
	Response.Write " <input type=""text"" name=""filesize"" id=""filesize"" size=""14"" onkeyup=if(isNaN(this.value))this.value='' value='"
	If isEdit Then
		Response.Write Trim(Rs("filesize"))
	End If
	Response.Write "'> <input name=""SizeUnit"" type=""radio"" value=""KB"" checked>"
	Response.Write " KB"
	Response.Write " <input type=""radio"" name=""SizeUnit"" value=""MB"">"
	Response.Write " MB <font color=""#FF0000"">��</font>"
%>
    </td>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>�Ǽ���</strong></td>
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
    <td align="right" class="TableRow2"><strong><%=sModuleName%>���ߣ�</strong></td>
    <td class="TableRow1"><input name="Author" type="text" id="Author" size="20" value="<%If isEdit Then Response.Write Rs("Author")%>">
	<select name=font2 onChange="Author.value=this.value;">
			<option selected value="">ѡ������</option>
			<option value='����'>����</option>
			<option value='��վԭ��'>��վԭ��</option>
			<option value='����'>����</option>
			<option value='δ֪'>δ֪</option>
			<option value='<%=AdminName%>'><%=AdminName%></option>
		</select></td>
    <td align="right" class="TableRow2"><strong>��Ʒ��Դ��</strong></td>
    <td class="TableRow1"><input name="ComeFrom" type="text" id="ComeFrom" size="25" value="<%If isEdit Then Response.Write Rs("ComeFrom")%>">
    <select name=font1 onChange="ComeFrom.value=this.value;">
			<option selected value="">ѡ����Դ</option>
			<option value='��վԭ��'>��վԭ��</option>
			<option value='��վ����'>��վ����</option>
			<option value='����'>����</option>
			<option value='ת��'>ת��</option>
			</select></td>
  </tr>
  <tr style="display:none">
    <td align="right" class="TableRow2"><strong>���صȼ���</strong></td>
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
    <td align="right" class="TableRow2"><strong>���ص�����</strong></td>
    <td class="TableRow1"><input name="PointNum" type="text" id="PointNum" size="10" value="<%If isEdit Then Response.Write Rs("PointNum") Else Response.Write 0 End If%>"></td>
  </tr>
  <tr>
    <td align="right" nowrap class="TableRow2"><strong><%=sModuleName%>����ͼ��</strong></td>
    <td colspan="3" class="TableRow1"><input name="miniature" type="text" id="ImageUrl" size="60" value="<%If isEdit Then Response.Write enchiasp.ChkNull(Rs("miniature"))%>">
    <input type='button' name='selectpic' value='�����ϴ�ͼƬ��ѡ��' onclick='SelectPhoto()' class=button></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>ͼƬ�ϴ�</strong></td>
    <td colspan="3" class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?ChannelID=<%=ChannelID%>></iframe></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>��飺</strong></td>
    <td colspan="3" class="TableRow1"><textarea name="content" style="display:none" id="content"><%If isEdit Then Response.Write Server.HTMLEncode(Rs("Introduce"))%></textarea>
    <iframe ID='HtmlEditor1' src='../editor/editor.asp?ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='560' HEIGHT='350'></iframe></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>������</strong></td>
    <td colspan="3" class="TableRow1"><input name="Describe" type="text" id="Describe" size="80" value="<%If isEdit Then Response.Write enchiasp.ChkNull(Rs("Describe"))%>"></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>����ѡ�</strong></td>
    <td colspan="3" class="TableRow1"><input name="isTop" type="checkbox" id="isTop" value="1"<%If isEdit Then:If Rs("isTop") <> 0 Then Response.Write (" checked")%>>
            <%=sModuleName%>�ö� 
            <input name="isBest" type="checkbox" id="isBest" value="1"<%If isEdit Then:If Rs("isBest") <> 0 Then Response.Write (" checked")%>> 
            <%=sModuleName%>�Ƽ�
	    <input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"<%If isEdit Then:If Rs("ForbidEssay") <> 0 Then Response.Write (" checked")%>> 
            ��ֹ��������
	    <input name="isAccept" type="checkbox" id="isAccept" value="1" checked> 
            ����������<font color=blue>������˺���ܷ�����</font>��
	    <%If isEdit Then%>
	    <input name="Update" type="checkbox" value="1"> 
            ͬʱ����ʱ��<%End If%></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>��ʾģʽ��</strong></td>
    <td colspan="3" class="TableRow1">
<%
	Dim ShowModeArray
	
	ShowModeArray = Array("����ʾ","FLASH","ͼƬ","Media","Real","DCR")
	For i = 0 To UBound(ShowModeArray)
		Response.Write "<input type=""radio"" name=""showmode"" value=""" & i & """ "
		If isEdit Then
			If i = Rs("showmode") Then Response.Write " checked"
		Else
			If i = DefaultShowMode  Then Response.Write " checked"
		End If
		Response.Write ">" & Trim(ShowModeArray(i)) & " "
		If i = 6 Then Response.Write "<br>"
	Next
%>
    </td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>ҳ����ʾURL��</strong></td>
    <td colspan="3" class="TableRow1"><input name="showurl" type="text" id="filePath" size="60" value="<%If isEdit Then Response.Write enchiasp.ChkNull(Rs("showurl"))%>">
    <input type='button' name='selectfile' value='�����ϴ��ļ���ѡ��' onclick='SelectFile()' class=button></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>�ϴ�<%=sModuleName%>��</strong></td>
    <td colspan="3" class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upflash.asp?ChannelID=<%=ChannelID%>></iframe></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>���ط�������</strong></td>
    <td colspan="3" class="TableRow1"><%=SelDownServer(downid)%> <b>˵����</b><font color=blue>���ط�����·�� + �����ļ����� = �������ص�ַ</font></td>
  </tr>
  <tr>
    <td align="right" nowrap class="TableRow2"><strong>���ص�ַ��</strong></td>
    <td colspan="3" class="TableRow1"><input name="DownAddress" type="text" id="DownAddress" size="80" value="<%If isEdit Then Response.Write enchiasp.ChkNull(Rs("DownAddress"))%>"></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2">��</td>
    <td colspan="3" align="center" class="TableRow1">
    <input type="button" name="Submit2" onclick="CheckLength();" value="�鿴���ݳ���" class=Button>
    <input type="button" name="Submit3" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>
    <input type="Submit" name="Submit1" value="����<%=sModuleName%>" class=Button></td>
  </tr></form>
</table></div>
<%
	If isEdit Then
		Rs.Close
		Set Rs = Nothing
	End If
End Sub

Private Function SelDownServer(intdownid)
	Dim rsobj
	If Not IsNumeric(intdownid) Then intdownid = 0
	Response.Write " <select name=""downid"" size=""1"">"
	Response.Write "<option value=""0"""
	If intdownid = 0 Then Response.Write " selected"
	Response.Write ">����ѡ�����ط�������</option>"
	SQL = "SELECT downid,DownloadName,depth,rootid FROM ECCMS_DownServer WHERE depth=0 And ChannelID="& ChannelID
	Set rsobj = enchiasp.Execute(SQL)
	Do While Not rsobj.EOF
		Response.Write "<option value=""" & rsobj("rootid") & """"
		If intdownid = rsobj("rootid") Then Response.Write " selected"
		Response.Write ">" & rsobj(1) & "</option>"
		rsobj.movenext
	Loop
	rsobj.Close
	Set rsobj = Nothing
	Response.Write "<option value=""0"">��ʹ�����ط�����</option>"
	Response.Write "</select>"
End Function
Private Sub CheckSave()
	If Trim(Request.Form("title")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "���Ʋ���Ϊ�գ�</li>"
	End If
	If Len(Request.Form("title")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "���Ʋ��ܳ���200���ַ���</li>"
	End If
	If Trim(Request.Form("ColorMode")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������ɫ��������</li>"
	End If
	If Trim(Request.Form("FontMode")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���������������</li>"
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
		ErrMsg = ErrMsg + "<li>" & sModuleName & "�ȼ���������</li>"
	End If
	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��һ�������Ѿ����������࣬�������" & sModuleName & "��</li>"
	End If
	If enchiasp.ChkNumeric(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�÷������ⲿ���ӣ��������" & sModuleName & "��</li>"
	End If	
	TextContent = ""
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next
	If Len(TextContent) < 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "��鲻��Ϊ�գ�</li>"
	End If
	If CInt(Request.Form("isTop")) = 1 Then
		FlashTop = 1
	Else
		FlashTop = 0
	End If
	If CInt(Request.Form("isBest")) = 1 Then
		FlashBest = 1
	Else
		FlashBest = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If
	If CInt(Request("isAccept")) = 1 Then
		FlashAccept = 1
	Else
		FlashAccept = 0
	End If
	If trim(Request.Form("filesize")) = "" Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "��С����Ϊ�գ�</li>"
	End If
End Sub

Private Sub SaveFlash()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_FlashList where (flashid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("SpecialID") = 0
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("ColorMode") = Trim(Request.Form("ColorMode"))
		Rs("FontMode") = Trim(Request.Form("FontMode"))
		Rs("Introduce") = TextContent
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Describe") = enchiasp.ChkFormStr(Request.Form("Describe"))
		If UCase(Request.Form("SizeUnit")) = "MB" Then
			Rs("filesize") = CLng(Request.Form("filesize") * 1024)
		Else
			Rs("filesize") = CLng(Request.Form("filesize"))
		End If
		Rs("Author") = Trim(Request.Form("Author"))
		Rs("ComeFrom") = Trim(Request.Form("ComeFrom"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("miniature") = Trim(Request.Form("miniature"))
		Rs("UserGroup") = enchiasp.ChkNumeric(Request.Form("UserGroup"))
		Rs("PointNum") = enchiasp.ChkNumeric(Request.Form("PointNum"))
		Rs("UserName") = Trim(AdminName)
		Rs("addTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("AllHits") = 0
		Rs("DayHits") = 0
		Rs("WeekHits") = 0
		Rs("MonthHits") = 0
		Rs("HitsTime") = Now()
		Rs("grade") = 0
		Rs("showmode") = enchiasp.ChkNumeric(Request.Form("showmode"))
		Rs("isTop") = FlashTop
		Rs("IsBest") = FlashBest
		Rs("downid") = enchiasp.ChkNumeric(Request.Form("downid"))
		Rs("showurl") = Trim(Request.Form("showurl"))
		Rs("DownAddress") = Trim(Request.Form("DownAddress"))
		Rs("isUpdate") = 1
		Rs("isAccept") = FlashAccept
		Rs("ForbidEssay") = ForbidEssay
		Rs("AlphaIndex") = enchiasp.ReadAlpha(Request.Form("title"))
	Rs.update
	Rs.Close
	Rs.Open "select top 1 flashid from ECCMS_FlashList where ChannelID=" & ChannelID & " order by flashid desc", Conn, 1, 1
	flashid = Rs("flashid")
	Rs.Close:Set Rs = Nothing
	Call RemoveCache
	ClassUpdateCount Request.Form("ClassID"),1
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makeflash.asp?ChannelID=" & ChannelID & "&flashid=" & flashid & "&showid=0"	
		Call ScriptCreation(url,flashid)
		SQL = "SELECT TOP 1 flashid FROM ECCMS_FlashList WHERE ChannelID=" & ChannelID & " And isAccept <> 0 And flashid < " & flashid & " ORDER BY flashid DESC"
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.EOF And Rs.BOF) Then
			url = "admin_makeflash.asp?ChannelID=" & ChannelID & "&flashid=" & Rs("flashid") & "&showid=0"	
			Call ScriptCreation(url,Rs("flashid"))
		End If
		Rs.Close
		Set Rs = Nothing
	End If
	Succeed("<li>��ϲ��������µ�" & sModuleName & "�ɹ���</li><li><a href=?action=view&ChannelID=" & ChannelID & "&flashid=" & flashid & ">����˴��鿴��" & sModuleName & "</a></li>")

End Sub

Private Sub ModifyFlash()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_FlashList where flashid=" & Request("flashid")
	Rs.Open SQL,Conn,1,3
		Auditing = Rs("isAccept")
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("SpecialID") = 0
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("ColorMode") = Trim(Request.Form("ColorMode"))
		Rs("FontMode") = Trim(Request.Form("FontMode"))
		Rs("Introduce") = TextContent
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Describe") = enchiasp.ChkFormStr(Request.Form("Describe"))
		If UCase(Request.Form("SizeUnit")) = "MB" Then
			Rs("filesize") = CLng(Request.Form("filesize") * 1024)
		Else
			Rs("filesize") = CLng(Request.Form("filesize"))
		End If
		Rs("Author") = Trim(Request.Form("Author"))
		Rs("ComeFrom") = Trim(Request.Form("ComeFrom"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("miniature") = Trim(Request.Form("miniature"))
		Rs("UserGroup") = Trim(Request.Form("UserGroup"))
		Rs("PointNum") = Trim(Request.Form("PointNum"))
		Rs("UserName") = Trim(AdminName)
		If CInt(Request.Form("Update")) = 1 Then Rs("addTime") = Now()
		Rs("showmode") = enchiasp.ChkNumeric(Request.Form("showmode"))
		Rs("isTop") = FlashTop
		Rs("IsBest") = FlashBest
		Rs("downid") = enchiasp.ChkNumeric(Request.Form("downid"))
		Rs("showurl") = Trim(Request.Form("showurl"))
		Rs("DownAddress") = Trim(Request.Form("DownAddress"))
		Rs("isUpdate") = 1
		Rs("isAccept") = FlashAccept
		Rs("ForbidEssay") = ForbidEssay
		Rs("AlphaIndex") = enchiasp.ReadAlpha(Request.Form("title"))
	Rs.update
	flashid = Rs("flashid")
	If FlashAccept = 1 And Auditing = 0 Then
		AddUserPointNum Rs("username"),1
	End If
	If FlashAccept = 0 And Auditing = 1 Then
		AddUserPointNum Rs("username"),0
	End If
	Rs.Close:Set Rs = Nothing
	Call RemoveCache
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makeflash.asp?ChannelID=" & ChannelID & "&flashid=" & flashid & "&showid=0"	
		Call ScriptCreation(url,flashid)
	End If
	Succeed("<li>��ϲ�����޸�" & sModuleName & "�ɹ���</li><li><a href=?action=view&ChannelID=" & ChannelID & "&flashid=" & flashid & ">����˴��鿴��" & sModuleName & "</a></li>")
End Sub

Private Sub FlashView()
	Call PageTop
	If Request("flashid") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	SQL = "select * from ECCMS_FlashList where ChannelID=" & ChannelID & " And flashid=" & Request("flashid")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κ�" & sModuleName & "��������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	Else
%>
	<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
	<tr>
	  <th colspan="2">�鿴<%=sModuleName%></th>
	</tr>
	<tr>
	  <td align="center" class="TableRow2" colspan="2"><a href=?action=edit&ChannelID=<%=ChannelID%>&flashid=<%=Rs("flashid")%>><font size=4><%=enchiasp.ReadFontMode(Rs("title"),Rs("ColorMode"),Rs("FontMode"))%></font></a></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong><%=sModuleName%>���ƣ�</strong> <%=Rs("title")%></td>
	  <td class="TableRow1"><strong><%=sModuleName%>��С��</strong> <%=Rs("filesize")%> KB</td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong><%=sModuleName%>���ߣ�</strong> <%=Rs("Author")%></td>
	  <td class="TableRow1"><strong>��Ʒ��Դ��</strong> <%=ReadComeFrom(Rs("ComeFrom"))%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>����ʱ�䣺</strong> <%=Rs("addTime")%></td>
	  <td class="TableRow1"><strong><%=sModuleName%>�Ǽ���</strong> 
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
	  <td colspan="2" align="center" class="TableRow1">
<%
	Call PreviewMode(Rs("showurl"),Rs("showmode"))
%>
	  </td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1"><strong><%=sModuleName%>��飺</strong><br>&nbsp;&nbsp;&nbsp;&nbsp;<%=enchiasp.ReadContent(Rs("Introduce"))%></td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1">��һ<%=sModuleName%>��<%=FrontFlash(Rs("flashid"))%>
	  <br>��һ<%=sModuleName%>��<%=NextFlash(Rs("flashid"))%></td>
	</tr>
	<tr>
	  <td align="center" colspan="2" class="TableRow2"><input type="button" onclick="{if(confirm('��ȷ��Ҫɾ����?')){location.href='?action=del&ChannelID=<%=ChannelID%>&flashid=<%=Rs("flashid")%>';return true;}return false;}" value="ɾ��<%=sModuleName%>" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="���ض���" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&ChannelID=<%=ChannelID%>&flashid=<%=Rs("flashid")%>'" value="�༭<%=sModuleName%>" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub
Public Function ReadComeFrom(ByVal strContent)
	ReadComeFrom = ""
	If IsNull(strContent) Then Exit Function
	If Trim(strContent) = "" Then Exit Function
	strContent = " " & strContent & " "
	Dim re
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""|'])+)"
	strContent = re.Replace(strContent,"<a target=""_blank"" href=$1>$1</a>")
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""])+)$([^\[|']*)"
	strContent = re.Replace(strContent,"<a target=""_blank"" href=$1>$1</a>")
	re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""|'])+)"
	strContent = re.Replace(strContent,"$1<a target=""_blank"" href=$2>$2</a>")
	re.Pattern = "([\s])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
	strContent = re.Replace(strContent,"<a target=""_blank"" href=""http://$2"">$2</a>")
	Set re = Nothing
	ReadComeFrom = Trim(strContent)
End Function

Private Sub PreviewMode(url,modeid)
	If Len(url) < 3 Then Exit Sub
	If Left(url,1) <> "/" And InStr(url,"://") = 0 Then
		url = "../" & enchiasp.ChannelDir & url
	End If
	Select Case CInt(modeid)
	Case 1
		Response.Write "<object codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,0,0"" height=""400"" width=""550"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"">"& vbCrLf
		Response.Write "	<param name=""movie"" value=""" & url & """>"& vbCrLf
		Response.Write "	<param name=""quality"" value=""high"">"& vbCrLf
		Response.Write "	<param name=""SCALE"" value=""exactfit"">"& vbCrLf
		Response.Write "	<embed src=""" & url & """ quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""550"" height=""400"">"& vbCrLf
		Response.Write "	</embed>"& vbCrLf
		Response.Write "</object>"& vbCrLf
	Case 2
		Response.Write "<img src=""" & url & """ border=""0"" onload=""return imgzoom(this,550)"">"
	Case 3
		Response.Write "<object classid=""CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95"" class=""OBJECT"" id=""MediaPlayer"" width=""220"" height=""220"">"& vbCrLf
		Response.Write "	<param name= value=""-1"">"& vbCrLf
		Response.Write "	<param name=""CaptioningID"" value>"& vbCrLf
		Response.Write "	<param name=""ClickToPlay"" value=""-1"">"& vbCrLf
		Response.Write "	<param name=""Filename"" value=""" & url & """>"& vbCrLf
		Response.Write "	<embed src=""" & url & """  width= 220 height=""220"" type=""application/x-oleobject"" codebase=""http://activex.microFlash.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=0,1,1,1"" flename=""mp""></embed>"& vbCrLf
		Response.Write "</object>"& vbCrLf
	Case 4
		Response.Write "<object classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" height=""288"" id=""video1"" width=""305"" VIEWASTEXT>"& vbCrLf
		Response.Write "	<param name=""_ExtentX"" value=""5503"">"& vbCrLf
		Response.Write "	<param name=""_ExtentY"" value=""1588"">"& vbCrLf
		Response.Write "	<param name=""AUTOSTART"" value=""-1"">"& vbCrLf
		Response.Write "	<param name=""SHUFFLE"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""PREFETCH"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""NOLABELS"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""SRC"" value=""" & url & """>"& vbCrLf
		Response.Write "	<param name=""CONTROLS"" value=""Imagewindow,StatusBar,ControlPanel"">"& vbCrLf
		Response.Write "	<param name=""CONSOLE"" value=""RAPLAYER"">"& vbCrLf
		Response.Write "	<param name=""LOOP"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""NUMLOOP"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""CENTER"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""MAINTAINASPECT"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""BACKGROUNDCOLOR"" value=""#000000"">"& vbCrLf
		Response.Write "</object>"& vbCrLf
	Case 5
		Response.Write "<object classid=""clsid:166B1BCA-3F9C-11CF-8075-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=8,5,1,0"" width=""100%"" height=""100%"">"& vbCrLf
		Response.Write "	<param name=""src"" value=""" & url & """>"& vbCrLf
		Response.Write "	<param name=""swRemote"" value=""swSaveEnabled='false' swVolume='false' swRestart='false' swPausePlay='false' swFastForward='false' swContextMenu='false' "">"& vbCrLf
		Response.Write "	<param name=""swStretchStyle"" value=""fill"">"& vbCrLf
		Response.Write "	<PARAM name=""bgColor"" value=""#000000"">"& vbCrLf
		Response.Write "	<PARAM name=logo value=""false"">"& vbCrLf
		Response.Write "	<embed src=""" & url & """ bgColor=""#000000"" logo=""FALSE"" width=""550"" height=""400"" swRemote=""swSaveEnabled='false' swVolume='false' swRestart='false' swPausePlay='false' swFastForward='false' swContextMenu='false' "" swStretchStyle=""fill"" type=""application/x-director"" pluginspage=""http://www.macromedia.com/shockwave/download/""></embed>"& vbCrLf
		Response.Write "</object>"& vbCrLf
	

	End Select
End Sub

Private Function FrontFlash(flashid)
	Dim Rss, SQL
	SQL = "select Top 1 flashid,classid,title from ECCMS_FlashList where ChannelID=" & ChannelID & " And isAccept <> 0 And flashid < " & flashid & " order by flashid desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		FrontFlash = "�Ѿ�û����"
	Else
		FrontFlash = "<a href=admin_flash.asp?action=view&ChannelID=" & ChannelID & "&flashid=" & Rss("flashid") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function NextFlash(flashid)
	Dim Rss, SQL
	SQL = "select Top 1 flashid,classid,title from ECCMS_FlashList where ChannelID=" & ChannelID & " And isAccept <> 0 And flashid > " & flashid & " order by flashid asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		NextFlash = "�Ѿ�û����"
	Else
		NextFlash = "<a href=admin_flash.asp?action=view&ChannelID=" & ChannelID & "&flashid=" & Rss("flashid") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Sub BatCreateHtml()
	Dim Allflashid,url
	Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
	Response.Write "<ol>"
	Allflashid = Split(selflashid, ",")
	For i = 0 To UBound(Allflashid)
		flashid = CLng(Allflashid(i))
		url = "admin_makeflash.asp?ChannelID=" & ChannelID & "&flashid=" & flashid & "&showid=1"	
		Call ScriptCreation(url,flashid)
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
Private Sub FlashDel()
	If Request("flashid") = "" Then
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	On Error Resume Next
	Set Rs = enchiasp.Execute("SELECT flashid,classid,username,HtmlFileDate FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And flashid=" & Request("flashid"))
	If Not(Rs.BOF And Rs.EOF) Then
		ClassUpdateCount Rs("classid"),0
		AddUserPointNum Rs("username"),0
		DeleteHtmlFile Rs("classid"),Rs("flashid"),Rs("HtmlFileDate")
	End If
	Rs.Close:Set Rs = Nothing
	Conn.Execute("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And flashid=" & Request("flashid"))
	Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID=" & Request("flashid"))
	Call RemoveCache
	Response.redirect ("admin_flash.asp?ChannelID=" & ChannelID)
End Sub

Private Sub batdel()
	Set Rs = enchiasp.Execute("SELECT flashid,classid,username,HtmlFileDate FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And flashid in (" & selflashid & ")")
	Do While Not Rs.EOF
		ClassUpdateCount Rs("classid"),0
		AddUserPointNum Rs("username"),0
		DeleteHtmlFile Rs("classid"),Rs("flashid"),Rs("HtmlFileDate")
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	Conn.Execute ("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And flashid in (" & selflashid & ")")
	Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID in (" & selflashid & ")")
	Call RemoveCache
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub batmove()
	If Not IsNumeric(Request.Form("ClassID")) Then
		OutHintScript ("��һ�������Ѿ����������࣬���ƶ������������࣡")
		Exit Sub
	End If
	If Trim(Request.Form("classid")) <> "" Then
		enchiasp.Execute ("update ECCMS_FlashList set ClassID = " & Request.Form("ClassID") & ",isUpdate=1 where flashid in (" & selflashid & ")")
		OutHintScript ("�����ƶ������ɹ�")
	Else
		OutHintScript ("�����ƶ����ⲿ���࣡")
	End If
End Sub

Private Sub upindate()
	enchiasp.Execute ("update [ECCMS_FlashList] set addTime = " & NowString & " where flashid in (" & selflashid & ")")
	Call RemoveCache
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub


Private Sub isCommend()
	enchiasp.Execute ("update ECCMS_FlashList set isBest=1 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noCommend()
	enchiasp.Execute ("update ECCMS_FlashList set isBest=0 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub isTop()
	enchiasp.Execute ("update ECCMS_FlashList set isTop=1 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noTop()
	enchiasp.Execute ("update ECCMS_FlashList set isTop=0 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

'----�������
Private Sub BatAccept()
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_FlashList WHERE isAccept=0 And flashid in (" & selflashid & ")")
	Do While Not Rs.EOF
		AddUserPointNum Rs("username"),1
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("update ECCMS_FlashList set isAccept=1 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Sub NotAccept()
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_FlashList WHERE isAccept>0 And flashid in (" & selflashid & ")")
	Do While Not Rs.EOF
		AddUserPointNum Rs("username"),0
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("update ECCMS_FlashList set isAccept=0 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Function AddUserPointNum(username,stype)
	On Error Resume Next
	Dim rsuser,GroupSetting,userpoint
	Set rsuser = enchiasp.Execute("SELECT userid,UserGrade,userpoint FROM ECCMS_User WHERE username='"& username &"'")
	If Not(rsuser.BOF And rsuser.EOF) Then
		GroupSetting = Split(enchiasp.UserGroupSetting(rsuser("UserGrade")), "|||")(13)
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

'--����������ʼ
Private Sub BatchSetting()
	If Not ChkAdmin("AdminFlash" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Dim Channel_Setting
	Channel_Setting = Split(enchiasp.Channel_Setting, "|||")
	Response.Write "<script src=""include/FlashJuge.Js"" type=""text/javascript""></script>" & vbNewLine
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=4>" & sModuleName & "��������</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=saveset>"
	Response.Write "	<input type=hidden name=ChannelID value='"& ChannelID &"'>"
	Response.Write "	<tr>"
	Response.Write "		<td width=""20%"" rowspan=""18"" class=tablerow2 valign=""top"" id=choose2 style=""display:none""><b>��ѡ��" & sModuleName & "����</b><br>"
	Response.Write "<select name=""ClassID"" size='2' multiple style='height:420px;width:180px;'>"
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
	Response.Write "		<td class=tablerow1><input type=""text"" name=""flashid"" size=70 value='"& Request("selflashid") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 width=""15%"" align=right><b>���" & sModuleName & "��</b></td>"
	Response.Write "		<td class=tablerow1 width=""5%"" align=center><input type=checkbox name=selRelated value='1'></td>"
	Response.Write "		<td class=tablerow1 width=""60%""><input name=Related type=text size=60></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "��Դ��</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selComeFrom value='1'></td>"
	Response.Write "		<td class=tablerow1 ><input name=ComeFrom type=text size=35>"
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
	If Not ChkAdmin("AdminFlash" & ChannelID) Then
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
	Response.Write " <b>ָ��" & sModuleName & "ID��</b> <input type=""text"" name=""flashid"" size=80 value='"& Request("selflashid") &"'></td>"
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
		If Trim(Request.Form("flashid")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ��" & sModuleName & "ID��</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_FlashList SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And flashid in ("& Request("flashid") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ����ࡣ</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_FlashList SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>��ϲ���������ƶ���ɡ�</li>")
End Sub

Private Sub BatcDelete()
	If Not ChkAdmin("AdminFlash" & ChannelID) Then
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
	Response.Write "<input type=radio name=Appointed value='1' onClick=""choose2.style.display='';choose1.style.display='none';""> <b>ָ��" & sModuleName & "���ࣺ</b>"
	Response.Write "<input type=radio name=Appointed value='2'> <b>ɾ������" & sModuleName & "</b>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=choose1>"
	Response.Write "		<td class=tablerow1><b>����ID��</b><input type=""text"" name=""flashid"" size=80 value='"& Request("selflashid") &"'></td>"
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
Private Sub AllDelFlash()
	On Error Resume Next
	If CInt(Request.Form("Appointed")) = 1 Then
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ����ࡣ</li>"
			Exit Sub
		End If
		enchiasp.Execute ("DELETE ECCMS_Comment FROM ECCMS_FlashList A INNER JOIN ECCMS_Comment C ON C.PostID=A.flashid WHERE A.ChannelID = "& ChannelID &" And A.ClassID IN (" & Request("ClassID") & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And ClassID IN (" & Request("ClassID") & ")")
	ElseIf CInt(Request.Form("Appointed")) = 2 Then
		enchiasp.Execute ("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID)
		enchiasp.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID)
	Else
		If Trim(Request.Form("flashid")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ��" & sModuleName & "ID��</li>"
			Exit Sub
		End If
		enchiasp.Execute ("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And flashid IN (" & Request("flashid") & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID IN (" & Request("flashid") & ")")
		
	End If
	Call RemoveCache
	Succeed("<li>����ɾ���ɹ���</li>")
End Sub

Private Sub SaveSetting()
	If Founderr = True Then Exit Sub
	Dim strTempValue
	strTempValue = ""
	If Trim(Request.Form("selRelated")) <> "" Then strTempValue = strTempValue & "Related='"& enchiasp.ChkFormStr(Request.Form("Related")) &"',"
	If Trim(Request.Form("selComeFrom")) <> "" Then strTempValue = strTempValue & "ComeFrom='"& enchiasp.ChkFormStr(Request.Form("ComeFrom")) &"',"
	If Trim(Request.Form("selAuthor")) <> "" Then strTempValue = strTempValue & "Author='"& enchiasp.ChkFormStr(Request.Form("Author")) &"',"
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
		If Trim(Request.Form("flashid")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ��" & sModuleName & "ID��</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_FlashList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And flashid in ("& Request("flashid") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ����ࡣ</li>"
			Exit Sub
		ElseIf Trim(Request.Form("ClassID")) = "-1" Then
			SQL = "UPDATE ECCMS_FlashList SET "& strTempValue &" WHERE ChannelID = "& ChannelID &" And isAccept>0"
		Else
			SQL = "UPDATE ECCMS_FlashList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>��ϲ��������������ɡ�</li>")
End Sub

Private Sub RemoveCache()
	enchiasp.DelCahe "RenewStatistics"
	enchiasp.DelCahe "TotalStatistics"
End Sub


%>