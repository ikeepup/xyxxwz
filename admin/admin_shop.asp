<!--#include file="setup.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
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
Dim Action,isEdit
Dim i,ClassID,RsObj,shopid,findword,keyword,strClass
Dim TextContent,ShopTop,ShopBest,ForbidEssay
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Dim s_ClassName,ChildStr,FoundSQL,isAccept,selShopID

If Request("isAccept") <> "" Then
	isAccept = 0
Else
	isAccept = 1
End If
If CInt(ChannelID) = 0 Then ChannelID = 3
Action = LCase(Request("action"))

Select Case Trim(Action)
Case "save"
	Call SaveShop
Case "modify"
	Call ModifyShop
Case "add"
	isEdit = False
	Call ShopEdit(isEdit)
Case "edit"
	isEdit = True
	Call ShopEdit(isEdit)
Case "del"
	Call ShopDel
Case "view"
	Call ShopView
Case "setting"
	Call BatchSetting
Case "saveset"
	Call SaveSetting
Case "move"
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
	Response.Write "	  <th colspan=2>" & sModuleName & "����ѡ��</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr><form method=Post name=myform action='admin_Shop.asp' onSubmit='return JugeQuery(this);'>"
	Response.Write "	<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "	<td class=TableRow1>������"
	Response.Write "	  <input name=keyword type=text size=20>"
	Response.Write "	  ������"
	Response.Write "	  <select name=queryopt>"
	Response.Write "		<option value='1' selected>" & sModuleName & "����</option>"
	Response.Write "		<option value='2'>" & sModuleName & "���</option>"
	Response.Write "		<option value='3'>��������</option>"
	Response.Write "	  </select> <input type=submit name=Submit value='��ʼ��ѯ' class=Button></td>"
	Response.Write "	  <td class=TableRow1>" & sModuleName & "������"
	Dim srtClassMenu
	Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Response.Write "<option value='admin_Shop.asp?ChannelID=" & ChannelID & "'>��ȫ��" & sModuleName & "�б��</option>" & vbCrLf
	srtClassMenu = enchiasp.ClassJumpMenu(ChannelID)
	srtClassMenu = Replace(srtClassMenu, "{ClassID=" & Request("sortid") & "}", "selected")
	Response.Write srtClassMenu
	Response.Write "</select>"
	Response.Write "	  </td></form>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td colspan=2 class=TableRow2><strong>����ѡ�</strong> <a href='admin_Shop.asp?ChannelID=" & ChannelID & "'>������ҳ</a> | "
	Response.Write "	  <a href='admin_Shop.asp?ChannelID=" & ChannelID & "&action=add'>���" & sModuleName & "</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "&action=add'>���" & sModuleName & "����</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "'>" & sModuleName & "�������</a></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub
Private Sub showmain()
	If Not ChkAdmin("AdminShop" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Dim strListName
	If Not IsEmpty(Request("selShopID")) Then
		selShopID = Request("selShopID")
		Select Case enchiasp.CheckStr(Request("act"))
		Case "����ɾ��"
			Call batdel
		Case "�����Ƽ�"
			Call isCommend
		Case "ȡ���Ƽ�"
			Call noCommend
		Case "�����ö�"
			Call isTop
		Case "ȡ���ö�"
			Call noTop
		Case "����HTML"
			Call BatCreateHtml
		Case Else
			Response.Write "��Ч������"
		End Select
	End If
	Call PageTop
	Dim specialID,sortid,Cmd,child
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>ѡ��</th>"
	Response.Write "	  <th width='68%'>" & sModuleName & "����</th>"
	Response.Write "	  <th width='9%' nowrap>�������</th>"
	Response.Write "	  <th width='9%' nowrap>" & sModuleName & "�Ǽ�</th>"
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
		CurrentPage = CLng(Request("page"))
	Else
		CurrentPage = 1
	End If
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("queryopt")) = 1 Then
			findword = "A.TradeName like '%" & keyword & "%'"
		ElseIf CInt(Request("queryopt")) = 2 Then
			findword = "A.Marque like '%" & keyword & "%'"
		Else
			findword = "A.TradeName like '%" & keyword & "%' or A.Marque like '%" & keyword & "%'"
		End If
		FoundSQL = findword
		s_ClassName = "��ѯ" & sModuleName
	Else
		specialID = 99999
		If Request("sortid") <> "" Then
			FoundSQL = "A.isAccept <> "& isAccept & " And A.ClassID in (" & ChildStr & ")"
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
	TotalNumber = enchiasp.Execute("SELECT COUNT(ShopID) FROM ECCMS_ShopList A WHERE A.ChannelID = " & ChannelID & " And "& FoundSQL &"")(0)
	TotalPageNum = CLng(TotalNumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT A.*,C.ClassName FROM [ECCMS_ShopList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.ChannelID = " & ChannelID & " And "& FoundSQL &" ORDER BY A.isTop DESC, A.addTime DESC ,A.ShopID DESC"
	
	If IsSqlDataBase = 1 Then
		'SQL�Ƿ����ô�������
		'If Trim(Request("keyword"))<>"" Or child > 0 Then
			'Set Rs = enchiasp.Execute(SQL)
		'Else
			'Set Cmd = Server.CreateObject("ADODB.Command")
			'Set Cmd.ActiveConnection=conn
			'Cmd.CommandText="ECCMS_ShopAdminList"
			'Cmd.CommandType=4
			'Cmd.Parameters.Append cmd.CreateParameter("@ChannelID",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@sortid",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@specialID",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@isAccept",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@pagenow",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@totalrec",3,2)
			'Cmd("@ChannelID")=ChannelID
			'Cmd("@sortid")=sortid
			'Cmd("@specialID")=specialID
			'Cmd("@isAccept")=isAccept
			'Cmd("@pagenow")=CurrentPage
			'Cmd("@pagesize")=maxperpage
			'Set Rs=Cmd.Execute
			
		'End If
		Rs.Open SQL, Conn, 1, 1
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=5 class=TableRow2>��û���ҵ��κ�" & sModuleName & "��</td></tr>"
	Else
		If IsSqlDataBase<>1 Or Trim(Request("keyword"))<>"" Or child > 0 Then
			Rs.MoveFirst
			If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
			If Rs.Eof Then Exit Sub
		End If
		i = 0

		Response.Write "	<tr>"
		Response.Write "	  <td colspan=5 class=TableRow2>"
		ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<form name=selform method=post action=""admin_shop.asp"">"
		Response.Write "	<input type=hidden name=ChannelID value='" & ChannelID & "'>"
		Response.Write "	<input type=hidden name=action value=''>"
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			If (i mod 2) = 0 Then
				strClass = "class=TableRow1"
			Else
				strClass = "class=TableRow2"
			End If
			Response.Write "	<tr>"
			Response.Write "	  <td align=center " & strClass & "><input type=checkbox name=selShopID value=" & Rs("ShopID") & "></td>"
			Response.Write "	  <td " & strClass & ">"
			If Rs("isTop") <> 0 Then
				Response.Write "<img src=""images/istop.gif"" width=15 height=17 border=0 alt=�ö���Ʒ>"
			End If

			Response.Write "[<a href=?ChannelID=" & Rs("ChannelID") & "&sortid="
			Response.Write Rs("ClassID")
			Response.Write ">"
			Response.Write Rs("ClassName")
			Response.Write "</a>] "
			Response.Write "<a href=?action=view&ChannelID=" & Rs("ChannelID") & "&ShopID="
			Response.Write Rs("ShopID")
			Response.Write ">"
			Response.Write Rs("TradeName")
			Response.Write "</a>" 

			If Rs("isBest") <> 0 Then
				Response.Write "&nbsp;&nbsp;<font color=blue>��</font>"
			End If
			Response.Write "	  </td>"
			Response.Write "	  <td align=center nowrap " & strClass & "><a href=?action=edit&ChannelID=" & Rs("ChannelID") & "&ShopID=" & Rs("ShopID") & ">�༭</a> | <a href=?action=del&ChannelID=" & Rs("ChannelID") & "&ShopID=" & Rs("ShopID") & " onclick=""{if(confirm('��Ʒɾ���󽫲��ָܻ�����ȷ��Ҫɾ������Ʒ��?')){return true;}return false;}"">ɾ��</a></td>"
			Response.Write "	  <td align=center nowrap " & strClass & ">"
			Response.Write "<font color=green>"
			For i = 1 to Rs("star")
				Response.Write "��"
			Next
			Response.Write "</font>"

			Response.Write "	  </td>"
			Response.Write "	  <td align=center nowrap " & strClass & ">"

			If Rs("addTime") >= Date Then
				Response.Write "<font color=red>"
				Response.Write enchiasp.FormatDate(Rs("addTime"), 2)
				Response.Write "</font>"
			Else
				Response.Write enchiasp.FormatDate(Rs("addTime"), 2)
			End If

			Response.Write "	  </td>"
			Response.Write "	</tr>"

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
		<option value="����HTML">����HTML</option>
	  </select>
	  <input class=Button type="submit" name="Submit2" value="ִ�в���" onclick="return confirm('��ȷ��ִ�иò�����?');">
	  <input class=Button type="submit" name="Submit3" value="��������" onclick="document.selform.action.value='setting';">
	  <input class=Button type="submit" name="Submit4" value="�����ƶ�" onclick="document.selform.action.value='move';"></td>
	</tr>
	</form>
	<tr>
	  <td colspan="5" align="right" class="TableRow2"><%ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName %></td>
	</tr>
</table>
<%
End Sub

Private Sub ShopEdit(isEdit)
	Dim EditTitle,TitleColor
	If isEdit Then
		SQL = "SELECT * FROM ECCMS_ShopList WHERE shopid=" & CLng(Request("shopid"))
		Set Rs = enchiasp.Execute(SQL)
		ClassID = Rs("ClassID")
		EditTitle = "�༭" & sModuleName
	Else
		If Not ChkAdmin("AdminShop" & ChannelID) Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
		EditTitle = "���" & sModuleName
		ClassID = Request("ClassID")
	End If
%>
<script src='include/ShopJuge.Js' type=text/javascript></script>
<div onkeydown=CtrlEnter()>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
  <tr>
    <th colspan="4"><%=EditTitle%></th>
  </tr>
  	<form method=Post name="myform" action="admin_shop.asp" onSubmit="return CheckForm(this);">
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""shopid"" value="""& Request("shopid") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
%>
  <tr>
    <td width="15%" align="right" class="TableRow2"><strong><%=sModuleName%>���ࣺ</strong></td>
    <td width="35%" class="TableRow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	sClassSelect = Replace(sClassSelect, "{ClassID=" & ClassID & "}", "selected")
	Response.Write sClassSelect
	Response.Write "</select>"
%>
    </td>
    <td width="15%" align="right" class="TableRow2"><strong></strong></td>
    <td width="35%" class="TableRow1"></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>���ƣ�</strong></td>
    <td class="TableRow1"><input name="TradeName" type="text" id="TradeName" size="30" value="<%If isEdit Then Response.Write Rs("TradeName")%>"> 
      <span class="style1">* </span></td>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>���</strong></td>
    <td class="TableRow1"><input name="Marque" type="text" id="Marque" size="20" value="<%If isEdit Then Response.Write Rs("Marque")%>"></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>��λ��</strong></td>
    <td class="TableRow1"><input name="Unit" type="text" id="Unit2" size="10" value="<%If isEdit Then Response.Write Rs("Unit")%>">
    <select name=font1 onChange="Unit.value=this.value;">
			<option selected value="">��ѡ��λ</OPTION>
			<option value=��>��</option>
			<option value=��>��</option>
			<option value=̨>̨</option>
			<option value=��>��</option>
			<option value=��>��</option>
			<option value=ƿ>ƿ</option>
			<option value=��>��</option>
			<option value=��>��</option>
			</select></td>
    <td align="right" class="TableRow2"><strong>��Դ��</strong></td>
    <td class="TableRow1"><input name="supply" type="text" id="supply" size="10" value="<%If isEdit Then Response.Write Rs("supply")%>">
    <select name=font2 onChange="supply.value=this.value;">
			<option value="">��ѡ��</OPTION>
			<option value=�л�>�л�</option>
			<option value=����>����</option>
			<option value=�޻�>�޻�</option>
			<option value=�ػ�>�ػ�</option>
			<option value=����>����</option>
			<option value=�ؼ�>�ؼ�</option>
			</select></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>�г��ۣ�</strong></td>
    <td class="TableRow1"><input name="PastPrice" type="text" id="PastPrice" size="10" value="<%If isEdit Then Response.Write Rs("PastPrice") else response.write 0 end if%>"> 
      <span class="style2">Ԫ</span> <span class="style1">* </span> ʹ�ñ�ǩ{$PastPrice}</td>
    <td align="right" class="TableRow2"><strong>�𿨼ۣ�</strong></td>
    <td class="TableRow1"><input name="NowPrice" type="text" id="NowPrice" size="10" value="<%If isEdit Then Response.Write Rs("NowPrice") else response.write 0 end if%>"> 
      <span class="style2">Ԫ</span> <span class="style1">* </span>ʹ�ñ�ǩ{$NowPrice}</td>
  </tr>
 <tr>
    <td align="right" class="TableRow2"><strong>�����ۣ�</strong></td>
    <td class="TableRow1"><input name="YinPrice" type="text" id="YinPrice" size="10" value="<%If isEdit Then Response.Write Rs("YinPrice") else response.write 0 end if%>"> 
      <span class="style2">Ԫ</span> <span class="style1">* </span> ʹ�ñ�ǩ{$YinPrice}</td>
    <td align="right" class="TableRow2"><strong>�����ۣ�</strong></td>
    <td class="TableRow1"><input name="OtherPrice" type="text" id="OtherPrice" size="10" value="<%If isEdit Then Response.Write Rs("OtherPrice") else response.write 0 end if%>"> 
      <span class="style2">Ԫ</span> <span class="style1">* </span> ʹ�ñ�ǩ{$OtherPrice}</td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>��Ʒ��˾��</strong></td>
    <td class="TableRow1"><input name="Company" type="text" id="Company" size="30" value="<%If isEdit Then Response.Write Rs("Company")%>"></td>
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
    <td align="right" class="TableRow2"><strong><%=sModuleName%>ͼƬ��</strong></td>
    <td colspan="3" class="TableRow1"><input name="ProductImage" type="text" id="ImageUrl" size="70" value="<%If isEdit Then Response.Write Rs("ProductImage")%>"> 
      <span class="style3">* </span></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>ͼƬ�ϴ���</strong></td>
    <td colspan="3" class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?ChannelID=<%=ChannelID%>></iframe></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>��飺</strong></td>
    <td colspan="3" class="TableRow1"><textarea name="content" style="display:none" id="content"><%If isEdit Then Response.Write Server.HTMLEncode(Rs("Explain"))%></textarea>
    <iframe ID='HtmlEditor1' src='../editor/editor.asp?ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='560' HEIGHT='350'></iframe></td>
  </tr>
  <tr>
          <td align="right" class="TableRow2"><strong>�ϴ��ļ���</strong></td>
          <td colspan="3" class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upfiles.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>����ѡ�</strong></td>
    <td colspan="3" class="TableRow1"><input name="isTop" type="checkbox" id="isTop" value="1"<%If isEdit Then:If Rs("isTop") <> 0 Then Response.Write (" checked")%>>
            <%=sModuleName%>�ö� 
            <input name="isBest" type="checkbox" id="isBest" value="1"<%If isEdit Then:If Rs("isBest") <> 0 Then Response.Write (" checked")%>> 
            <%=sModuleName%>�Ƽ�
	    <input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"<%If isEdit Then:If Rs("ForbidEssay") <> 0 Then Response.Write (" checked")%>> 
            ��ֹ��������
	    <%If isEdit Then%>
	    <input name="Update" type="checkbox" value="1"> 
            ͬʱ�����ϼ�ʱ��<%End If%></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2">��</td>
    <td colspan="3" align="center" class="TableRow1">
    <input type="button" name="Submit2" onclick="CheckLength();" value="�鿴���ݳ���" class=Button>
    <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>
    <input type="submit" name="Submit1" value="����<%=sModuleName%>" class=Button></td>
  </tr></form>
</table></div>
<%
	If isEdit Then
		Rs.Close
		Set Rs = Nothing
	End If
End Sub

Private Sub CheckSave()
	If Trim(Request.Form("TradeName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "���Ʋ���Ϊ�գ�</li>"
	End If
	If Len(Request.Form("TradeName")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "���Ʋ��ܳ���200���ַ���</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "�Ǽ�����Ϊ�ա�</li>"
	End If

	If CLng(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�÷������ⲿ���ӣ��������" & sModuleName & "��</li>"
	End If
	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��һ�������Ѿ����������࣬�������" & sModuleName & "��</li>"
	End If
	If Trim(Request.Form("content")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "��鲻��Ϊ�գ�</li>"
	End If
	TextContent = ""
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next
	If CInt(Request.Form("isTop")) = 1 Then
		ShopTop = 1
	Else
		ShopTop = 0
	End If
	If CInt(Request.Form("isBest")) = 1 Then
		ShopBest = 1
	Else
		ShopBest = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If

End Sub

Private Sub SaveShop()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_ShopList where (shopid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("TradeName") = enchiasp.ChkFormStr(Request.Form("TradeName"))
		Rs("Marque") = enchiasp.ChkFormStr(Request.Form("Marque"))
		'Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Explain") = TextContent
		Rs("Unit") = enchiasp.ChkFormStr(Request.Form("Unit"))
		Rs("supply") = enchiasp.ChkFormStr(Request.Form("supply"))
		Rs("Company") = Trim(Request.Form("Company"))
		Rs("PastPrice") = Trim(Request.Form("PastPrice"))
		Rs("yinPrice") = Trim(Request.Form("yinPrice"))
		Rs("otherPrice") = Trim(Request.Form("otherPrice"))
		Rs("NowPrice") = Trim(Request.Form("NowPrice"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("ProductImage") = Trim(Request.Form("ProductImage"))
		Rs("addTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("BuyCount") = 0
		Rs("IsBest") = ShopBest
		Rs("IsTop") = ShopTop
		Rs("isAccept") = 1
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	Rs.Close
	Rs.Open "select top 1 shopid from ECCMS_ShopList where ChannelID=" & ChannelID & " order by shopid desc", Conn, 1, 1
	shopid = Rs("shopid")
	Rs.Close:Set Rs = Nothing
	ClassUpdateCount Request.Form("ClassID"),1
	Call RemoveCache
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makeshop.asp?ChannelID=" & ChannelID & "&shopid=" & shopid & "&showid=0"	
		Call ScriptCreation(url,shopid)
		SQL = "SELECT TOP 1 shopid FROM ECCMS_ShopList WHERE ChannelID=" & ChannelID & " And isAccept <> 0 And shopid < " & shopid & " ORDER BY shopid DESC"
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.EOF And Rs.BOF) Then
			url = "admin_makeshop.asp?ChannelID=" & ChannelID & "&shopid=" & Rs("shopid") & "&showid=0"	
			Call ScriptCreation(url,Rs("shopid"))
		End If
		Rs.Close
		Set Rs = Nothing
	End If
	Succeed("<li>��ϲ��������µ�" & sModuleName & "�ɹ���</li><li><a href=?action=view&ChannelID=" & ChannelID & "&shopid=" & shopid & ">����˴��鿴��" & sModuleName & "</a></li>")
End Sub

Private Sub ModifyShop()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_ShopList where shopid=" & Request("shopid")
	Rs.Open SQL,Conn,1,3
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("TradeName") = enchiasp.ChkFormStr(Request.Form("TradeName"))
		Rs("Marque") = enchiasp.ChkFormStr(Request.Form("Marque"))
		'Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Explain") = TextContent
		Rs("Unit") = enchiasp.ChkFormStr(Request.Form("Unit"))
		Rs("supply") = enchiasp.ChkFormStr(Request.Form("supply"))
		Rs("Company") = Trim(Request.Form("Company"))
		Rs("PastPrice") = Trim(Request.Form("PastPrice"))
		Rs("NowPrice") = Trim(Request.Form("NowPrice"))
		Rs("yinPrice") = Trim(Request.Form("yinPrice"))
		Rs("otherPrice") = Trim(Request.Form("otherPrice"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("ProductImage") = Trim(Request.Form("ProductImage"))
		If CInt(Request.Form("Update")) = 1 Then Rs("addTime") = Now()
		Rs("IsBest") = ShopBest
		Rs("IsTop") = ShopTop
		Rs("isAccept") = 1
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	shopid = Rs("shopid")
	Rs.Close:Set Rs = Nothing
	Call RemoveCache
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makeshop.asp?ChannelID=" & ChannelID & "&shopid=" & shopid & "&showid=0"	
		Call ScriptCreation(url,shopid)
	End If
	Succeed("<li>��ϲ�����޸�" & sModuleName & "�ɹ���</li><li><a href=?action=view&ChannelID=" & ChannelID & "&shopid=" & shopid & ">����˴��鿴��" & sModuleName & "</a></li>")
End Sub
Private Sub shopView()
	Call PageTop
	If Request("shopid") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	SQL = "select * from ECCMS_ShopList where shopid=" & Request("shopid")
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
	  <td align="center" class="TableRow2" colspan="2"><a href=?action=edit&ChannelID=<%=ChannelID%>&shopid=<%=Rs("shopid")%>><font size=4><%=Rs("TradeName")%></font></a></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong><%=sModuleName%>���ƣ�</strong> <%=Rs("TradeName")%></td>
	  <td class="TableRow1"><strong><%=sModuleName%>�ͺţ�</strong> <%=Rs("Marque")%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong><%=sModuleName%>��λ��</strong> <%=Rs("Unit")%></td>
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
	  <td class="TableRow1"><strong>ԭ�ۣ�</strong> <%=FormatCurrency(Rs("PastPrice"))%> Ԫ/<%=Rs("Unit")%></td>
	  <td class="TableRow1"><strong>�ּۣ�</strong> <%=FormatCurrency(Rs("NowPrice"))%> Ԫ/<%=Rs("Unit")%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>�ϼ�ʱ�䣺</strong> <%=Rs("addTime")%></td>
	  <td class="TableRow1"><strong>��Ʒ��˾��</strong> <%=Rs("Company")%></td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1"><strong><%=sModuleName%>��飺</strong><br>&nbsp;&nbsp;&nbsp;&nbsp;<%=enchiasp.ReadContent(Rs("Explain"))%></td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1">��һ��Ʒ��<%=FrontShop(Rs("ShopID"))%>
	  <br>��һ��Ʒ��<%=NextShop(Rs("ShopID"))%></td>
	</tr>
	<tr>
	  <td align="center" colspan="2" class="TableRow2"><input type="button" onclick="javascript:window.close()" value="�رձ�����" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="���ض���" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&ChannelID=<%=ChannelID%>&ShopID=<%=Rs("ShopID")%>'" value="�༭��Ʒ" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub

Private Function FrontShop(ShopID)
	Dim Rss, SQL
	SQL = "select Top 1 ShopID,classid,TradeName from ECCMS_ShopList where ChannelID=" & ChannelID & " And isAccept <> 0 And ShopID < " & ShopID & " order by ShopID desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		FrontShop = "�Ѿ�û����"
	Else
		FrontShop = "<a href=admin_Shop.asp?action=view&ChannelID=" & ChannelID & "&ShopID=" & Rss("ShopID") & ">" & Rss("TradeName") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function NextShop(ShopID)
	Dim Rss, SQL
	SQL = "select Top 1 ShopID,classid,TradeName from ECCMS_ShopList where ChannelID=" & ChannelID & " And isAccept <> 0 And ShopID > " & ShopID & " order by ShopID asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		NextShop = "�Ѿ�û����"
	Else
		NextShop = "<a href=admin_Shop.asp?action=view&ChannelID=" & ChannelID & "&ShopID=" & Rss("ShopID") & ">" & Rss("TradeName") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Sub BatCreateHtml()
	Dim Allshopid,url
	Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
	Response.Write "<ol>"
	Allshopid = Split(selshopid, ",")
	For i = 0 To UBound(Allshopid)
		shopid = CLng(Allshopid(i))
		url = "admin_makeshop.asp?ChannelID=" & ChannelID & "&shopid=" & shopid & "&showid=1"	
		Call ScriptCreation(url,shopid)
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

Private Sub ShopDel()
	If Request("ShopID") = "" Then
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT shopid,classid,HtmlFileDate FROM ECCMS_ShopList WHERE ChannelID = "& ChannelID &" And shopid=" & Request("shopid"))
	If Not(Rs.BOF And Rs.EOF) Then
		ClassUpdateCount Rs("classid"),0
		DeleteHtmlFile Rs("classid"),Rs("shopid"),Rs("HtmlFileDate")
	End If
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute("DELETE FROM ECCMS_ShopList WHERE ShopID = " & Request("ShopID"))
	Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID=" & Request("ShopID"))
	Call RemoveCache
	OutHintScript("" & sModuleName & "ɾ���ɹ���")
End Sub

Private Sub batdel()
	Set Rs = enchiasp.Execute("SELECT shopid,classid,HtmlFileDate FROM ECCMS_ShopList WHERE ChannelID = "& ChannelID &" And shopid in (" & selShopID & ")")
	Do While Not Rs.EOF
		ClassUpdateCount Rs("classid"),0
		DeleteHtmlFile Rs("classid"),Rs("shopid"),Rs("HtmlFileDate")
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	Conn.Execute ("DELETE FROM ECCMS_ShopList WHERE ShopID in (" & selShopID & ")")
	Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID in (" & selShopID & ")")
	Call RemoveCache
	OutHintScript ("����ɾ�������ɹ���")
End Sub

Private Sub isCommend()
	enchiasp.Execute ("update ECCMS_ShopList set isBest=1 where ShopID in (" & selShopID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noCommend()
	enchiasp.Execute ("update ECCMS_ShopList set isBest=0 where ShopID in (" & selShopID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub isTop()
	enchiasp.Execute ("update ECCMS_ShopList set isTop=1 where ShopID in (" & selShopID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noTop()
	enchiasp.Execute ("update ECCMS_ShopList set isTop=0 where ShopID in (" & selShopID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Sub BatchSetting()
	If Not ChkAdmin("AdminShop" & ChannelID) Then
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
	Response.Write "		<td class=tablerow1><input type=""text"" name=""shopid"" size=70 value='"& Request("selshopid") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 width=""15%"" align=right><b>" & sModuleName & "��˾��</b></td>"
	Response.Write "		<td class=tablerow1 width=""5%"" align=center><input type=checkbox name=selCompany value='1'></td>"
	Response.Write "		<td class=tablerow1 width=""60%""><input name=Company type=text size=60></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "��λ��</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selUnit value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=Unit type=text size=20>"
	Response.Write "		<select name=font2 onChange=""Unit.value=this.value;"">"
	Response.Write "		<option selected value=''>��ѡ��λ</OPTION>"
	Response.Write "		<option value=��>��</option>"
	Response.Write "		<option value=��>��</option>"
	Response.Write "		<option value=̨>̨</option>"
	Response.Write "		<option value=��>��</option>"
	Response.Write "		<option value=��>��</option>"
	Response.Write "		<option value=ƿ>ƿ</option>"
	Response.Write "		<option value=��>��</option>"
	Response.Write "		<option value=��>��</option>"
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "��Դ��</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selsupply value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=supply type=text size=20>"
	Response.Write "		<select name=font1 onChange=""supply.value=this.value;"">"
	Response.Write "		<option selected value=''>��ѡ��</option>"
	Response.Write "		<option value=�л�>�л�</option>"
	Response.Write "		<option value=����>����</option>"
	Response.Write "		<option value=�޻�>�޻�</option>"
	Response.Write "		<option value=�ػ�>�ػ�</option>"
	Response.Write "		<option value=����>����</option>"
	Response.Write "		<option value=�ؼ�>�ؼ�</option>"
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "ԭ�ۣ�</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selPastPrice value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=PastPrice type=text size=10 value=0> Ԫ</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "�ּۣ�</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selNowPrice value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=NowPrice type=text size=10 value=0> Ԫ</td>"
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
	If Not ChkAdmin("AdminShop" & ChannelID) Then
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
	Response.Write " <b>ָ��" & sModuleName & "ID��</b> <input type=""text"" name=""shopid"" size=80 value='"& Request("selshopid") &"'></td>"
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
		If Trim(Request.Form("shopid")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ��" & sModuleName & "ID��</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_ShopList SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And shopid in ("& Request("shopid") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ����ࡣ</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_ShopList SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>��ϲ���������ƶ���ɡ�</li>")
End Sub

Private Sub SaveSetting()
	If Founderr = True Then Exit Sub
	Dim strTempValue
	strTempValue = ""
	If Trim(Request.Form("selCompany")) <> "" Then strTempValue = strTempValue & "Company='"& enchiasp.ChkFormStr(Request.Form("Company")) &"',"
	If Trim(Request.Form("selUnit")) <> "" Then strTempValue = strTempValue & "Unit='"& enchiasp.ChkFormStr(Request.Form("Unit")) &"',"
	If Trim(Request.Form("selsupply")) <> "" Then strTempValue = strTempValue & "supply='"& enchiasp.ChkFormStr(Request.Form("supply")) &"',"
	If Trim(Request.Form("selPastPrice")) <> "" Then strTempValue = strTempValue & "PastPrice="& CLng(Request.Form("PastPrice")) &","
	If Trim(Request.Form("selNowPrice")) <> "" Then strTempValue = strTempValue & "NowPrice="& CLng(Request.Form("NowPrice")) &","
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
		If Trim(Request.Form("shopid")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ��" & sModuleName & "ID��</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_ShopList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And shopid in ("& Request("shopid") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ����ࡣ</li>"
			Exit Sub
		ElseIf Trim(Request.Form("ClassID")) = "-1" Then
			SQL = "UPDATE ECCMS_ShopList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID
		Else
			SQL = "UPDATE ECCMS_ShopList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>��ϲ��������������ɡ�</li>")
End Sub

Private Sub ResetDateTime()
	Response.Write "<br><table width='400' align=center border=0 cellspacing=1 cellpadding=1>" & vbCrLf
	Response.Write "<tr> " & vbCrLf
	Response.Write "<td style=""border: 1px #384780 solid ;background-color: #FFFFFF;"">" & vbCrLf
	Response.Write "<table width=0 id=table2 name=table2 border=0 cellspacing=0 cellpadding=0>" & vbCrLf
	Response.Write "<tr height=12><td bgcolor=#36D91A>" & vbCrLf
	Response.Write "</td></tr></table></td></tr><tr> " & vbCrLf
	Response.Write "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span> <span style=""font-size:9pt"">%</span></td></tr>" & vbCrLf
	Response.Write "</table><br>" & vbCrLf
	Response.Flush
	Dim totalrec,addTime,page_count
	i = 0
	page_count = 0
	totalrec = enchiasp.Execute("SELECT COUNT(shopid) FROM [ECCMS_ShopList] WHERE ChannelID = "& ChannelID &" And isAccept>0")(0)
	Set Rs = enchiasp.Execute("SELECT shopid,addTime FROM [ECCMS_ShopList] WHERE ChannelID = "& ChannelID &" And isAccept>0 ORDER BY addTime DESC")
	If Not (Rs.BOF And Rs.EOF) Then
		Do While Not Rs.EOF
			Response.Write "<script>"
			Response.Write "table2.style.width=" & Fix((page_count / totalrec) * 400) & ";"
			Response.Write "txt2.innerHTML=""��ɣ�" & FormatNumber(page_count / totalrec * 100, 2, -1) & """;"
			Response.Write "</script>" & vbCrLf
			Response.Flush
			addTime = DateAdd("s", -i, Rs("addTime"))
			enchiasp.Execute ("UPDATE ECCMS_ShopList SET addTime='" & addTime & "' WHERE shopid="& Rs("shopid"))
			Rs.movenext
			i = i + 5
			page_count = page_count + 1
		Loop
	End If
	Set Rs = Nothing
	Response.Write "<script>table2.style.width=400;txt2.innerHTML=""��ɣ�100"";</script>"
End Sub

Private Sub RemoveCache()
	enchiasp.DelCahe "RenewStatistics"
	enchiasp.DelCahe "TotalStatistics"
End Sub
%>


