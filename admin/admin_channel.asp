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
Response.Write "<script language = JavaScript>" & vbCrLf
Response.Write "function ChannelSetting(n){" & vbCrLf
Response.Write "	if (n == 1){" & vbCrLf
Response.Write "		ChannelSetting1.style.display='none';" & vbCrLf
Response.Write "		ChannelSetting2.style.display='';" & vbCrLf
Response.Write "	}" & vbCrLf
Response.Write "	else{" & vbCrLf
Response.Write "		ChannelSetting1.style.display='';" & vbCrLf
Response.Write "		ChannelSetting2.style.display='none';" & vbCrLf
Response.Write "	}" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
Response.Write "	<tr>"
Response.Write "		<th colspan=""2"">վ��Ƶ������</th>"
Response.Write "	</tr>"
Response.Write "	<tr>"
Response.Write "		<td width=""100%"" class=TableRow2 colspan=2><b>����ѡ�</b><a href=admin_channel.asp>������ҳ</a>" 
Response.Write "		| <a href=?action=add>���Ƶ��</a> | "
Dim Rsm,ModuleName,strModuleName,sChannelID,NewChannelID
Set Rsm = enchiasp.Execute("SELECT ChannelID,ModuleName From ECCMS_Channel WHERE ChannelType < 2  ORDER BY orders ASC")
Do While Not Rsm.EOF
	Response.Write "<a href=?action=edit&ChannelID="
	Response.Write Rsm("ChannelID")
	Response.Write ">"
	Response.Write Rsm("ModuleName")
	Response.Write "����</a> | "
	strModuleName = strModuleName & Rsm("ModuleName") & "|||"
	sChannelID = sChannelID & Rsm("ChannelID") & "|||"
	Rsm.movenext
Loop
Set Rsm = Nothing
Response.Write "<a href=?action=orders>Ƶ������</a>"
Response.Write "		</td>"
Response.Write "	</tr>"
Response.Write "</table>"
Response.Write "<br>"
Dim Action,ChannelDir,TitleColor,mChannelDir,mChannelID
Dim i,RsObj
Action = LCase(enchiasp.RemoveBadCharacters(Request("action")))
If Not ChkAdmin("Channel") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Select Case Trim(Action)
Case "savenew"
	Call SavenewChannel
Case "savedit"
	Call SaveditChannel
Case "add"
	Call ChannelAdd
Case "edit"
	Call ChannelEdit
Case "del"
	Call ChannelDel
Case "orders"
	Call ChannelOrders
Case "saveorder"
	Call SaveOrder
Case "stopchannel"
	Call UpdateStop
Case "ishidden"
	Call UpdateHidden
Case "linktarget"
	Call UpdateLinkTarget
Case "createhtml"
	Call UpdateCreateHtml
Case "reload"
	Call ReloadChannelCache
Case Else
	Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If

Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub showmain()
	Response.Write "<table border=""0"" align=""center"" cellspacing=""1"" cellpadding=""3"" class=""TableBorder"">"
	Response.Write "	<tr>"
	Response.Write "		<th>Ƶ������</th>"
	Response.Write "		<th>Ƶ������</th>"
	Response.Write "		<th>Ƶ��״̬</th>"
	Response.Write "		<th>�Ƿ�HTML</th>"
	Response.Write "		<th>����״̬</th>"
	Response.Write "		<th>����Ŀ��</th>"
	Response.Write "		<th>����ѡ��</th>"
	Response.Write "	</tr>"

	SQL = "SELECT * FROM ECCMS_Channel ORDER BY orders"
	Set Rs = enchiasp.Execute(SQL)
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	Do While Not Rs.EOF
		Response.Write "	<tr>"
		Response.Write "		<td class=""TableRow2"">"
		Response.Write ("<a href=?action=edit&ChannelID=" & Rs("ChannelID") & " title=�޸Ĵ�Ƶ������>")
		Response.Write (enchiasp.ReadFontMode(Rs("ChannelName"),Rs("ColorModes"),Rs("FontModes")))
		Response.Write ("</a>")
		Response.Write "		</td>"
		Response.Write "		<td class=""TableRow1"" align=""center"">"
		If Rs("ChannelType") = 0 Then
			Response.Write ("<font color=blue>ϵͳƵ��")
		Elseif Rs("ChannelType") = 1 Then
			Response.Write ("<font color=green>�ڲ�Ƶ��")
		Else
			Response.Write ("<font color=red>�ⲿƵ��")
		End If
		Response.Write ("<font>")
		Response.Write ("</td>")
		If Rs("ChannelType") < 2 Then
			Response.Write ("<td class=""TableRow2"" align=""center"">")
			If Rs("StopChannel") <> 0 Then
				Response.Write ("<a href=?action=StopChannel&ChannelID=" & Rs("ChannelID") & "&EditID=0 title=""�л������򿪴�Ƶ��""><font color=red>�ر�<font></a>")
			Else
				Response.Write ("<a href=?action=StopChannel&ChannelID=" & Rs("ChannelID") & "&EditID=1 title=""�л������رմ�Ƶ��"">��</a>")
			End If
			Response.Write "		</td>"
			Response.Write "		<td class=""TableRow1"" align=""center"">"
			If Rs("IsCreateHtml") = 0 Then
				If Rs("ChannelID") = 4 Then
					Response.Write ("��")
				Else
					Response.Write ("<a href=?action=createhtml&ChannelID=" & Rs("ChannelID") & "&EditID=1 title=""�л���������HTML"">��</a>")
				End If
			Else
				Response.Write ("<a href=?action=createhtml&ChannelID=" & Rs("ChannelID") & "&EditID=0 title=""�л�����������HTML""><font color=blue>��</font></a>")
			End If
		Else
			Response.Write ("<td colspan=""2"" class=""TableRow2"" align=""center"">")
			Response.Write ("<a href=" & Rs("ChannelUrl") & " target=_blank><font color=blue>" & Rs("ChannelUrl") & "</font></a>")
		End If
		Response.Write "		</td>"
		Response.Write "				<td class=""TableRow2"" align=""center"">"
		If Rs("IsHidden") <> 0 Then
			Response.Write ("<a href=?action=ishidden&ChannelID=" & Rs("ChannelID") & "&EditID=0 title=""�л���������Ƶ������""><font color=green>����<font></a>")
		Else
			Response.Write ("<a href=?action=ishidden&ChannelID=" & Rs("ChannelID") & "&EditID=1 title=""�л���������Ƶ������"">��ʾ</a>")
		End If
		Response.Write "		</td>"
		Response.Write "		<td class=""TableRow1"" align=""center"">"
		If Rs("LinkTarget") = 0 Then
			Response.Write ("<a href=?action=linktarget&ChannelID=" & Rs("ChannelID") & "&EditID=1 title=""�л������´��ڴ�"">�����ڴ�</a>")
		Else
			Response.Write ("<a href=?action=linktarget&ChannelID=" & Rs("ChannelID") & "&EditID=0 title=""�л����������ڴ�""><font color=blue>�´��ڴ�<font></a>")
		End If
		Response.Write "		</td>"
		Response.Write "		<td class=""TableRow2"" align=""center""><A HREF=?action=edit&ChannelID="
		Response.Write Rs("ChannelID")
		Response.Write ">�� ��</A>"
		If Rs("ChannelID") => 10 Then
			Response.Write " | <A HREF=?action=del&ChannelID="
			Response.Write Rs("ChannelID")
			Response.Write " onclick=""{if(confirm('�˲�����ɾ����Ƶ��\n��ȷ��Ҫɾ����?')){return true;}return false;}"">ɾ ��</A>"
		End If
		If Rs("ChannelType") < 2 Then
			'Response.Write " | <A HREF=?action=reload&ChannelID="
			'Response.Write Rs("ChannelID")
			'Response.Write "><font color=blue>���»���</font></a>"
			If Rs("ChannelID") <> 4 Then
				Response.Write " | <A HREF=admin_classify.asp?action=jsmenu&ChannelID="
				Response.Write Rs("ChannelID")
				Response.Write "&stype=1><font color=green>����JS�˵�</font></a>"
			End If
		End If
		Response.Write "		</td>"
		Response.Write "	</tr>"

	Rs.movenext
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write "	<tr>"
	Response.Write "		<td colspan=""7"" class=""TableRow1""><b>˵����</b> <br>�١������Ӧ��״̬�����Խ�����ؿ���л�������<br>"
	Response.Write "�ڡ����л�HTML���ɹ��ܺ���<font color=red>��������JS</font>�˵���"
	Response.Write "</td>	</tr>"
	Response.Write "</table>"
End Sub

Private Sub ChannelAdd()
	
	Set Rs = enchiasp.Execute("select Max(ChannelID) from ECCMS_Channel")
	If Rs.bof And Rs.EOF Then
		NewChannelID = 1
	Else
		NewChannelID = Rs(0) + 1
	End If
	If IsNull(NewChannelID) Then NewChannelID = 1
	Rs.Close
	If NewChannelID < 10 Then NewChannelID = 10
%>
	<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
		<tr>
			<th colspan="2" align="left"><img src="images/welcome.gif" width="16" height="17" align="absMiddle"> ���վ��Ƶ��</th>
		</tr>
		<form method="POST" action="?action=savenew">
		<input type="hidden" name="NewChannelID" value="<%=NewChannelID%>">
		<tr>
			<td width="20%" class="TableRow2"><div class="divbody">Ƶ������</td>
			<td width="80%" class="TableRow1">
			<input type="text" name="ChannelName" size="20"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">Ƶ������ģʽ</div></td>
			<td class="TableRow1"> ��ɫ��
			<select size="1" name="ColorModes">
			<option value="0">��ѡ�������ɫ</option>
<%
	TitleColor = "," & enchiasp.InitTitleColor
	TitleColor = Split(TitleColor, ",")
	For i = 1 To UBound(TitleColor)
		Response.Write ("<option style=""background-color:"& TitleColor(i) &";color: "& TitleColor(i) &""" value='"& i &"'>"& TitleColor(i) &"</option>")
	Next
%>
			</select> ���壺
			<select size="1" name="FontModes">
		<option value="0">��ѡ������</option>
		<option value="1">����</option>
		<option value="2">б��</option>
		<option value="3">�»���</option>
		<option value="4">����+б��</option>
		<option value="5">����+�»���</option>
		<option value="6">б��+�»���</option>
		</select></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">Ƶ��ע��</div></td>
			<td class="TableRow1">
			<input type="text" name="Caption" size="60"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">Ƶ������</div></td>
			<td class="TableRow1">
			<input type="radio" value="2" checked name="ChannelType" onClick="ChannelType1.style.display='';ChannelType2.style.display='none';ChannelType3.style.display='none';"> �ⲿƵ��&nbsp;&nbsp; 
			<input type="radio" name="ChannelType" value="1" onClick="ChannelType1.style.display='none';ChannelType2.style.display='';ChannelType3.style.display='';"> �ڲ�Ƶ��</td>
		</tr>
		<tr id=ChannelType1>
			<td class="TableRow2"><div class="divbody">Ƶ������URL</div></td>
			<td class="TableRow1">
			<input type="text" name="ChannelUrl" size="45" value="<%=enchiasp.SiteUrl%>"> <font color="#FF0000">
			* ������������URL</font></td>
		</tr>
		<tr id=ChannelType2 style="display:none">
			<td class="TableRow2"><div class="divbody">����ģ��</div></td>
			<td class="TableRow1">
			<select name="modules" szie=1>
				<option value='1'>����</option>
				<option value='2'>���</option>
				<option value='3'>�̳�</option>
				<option value='5'>����</option>
				<option value='6'>��ҳͼ��</option>

			</select></td>
		</tr>
		<tr id=ChannelType3 style="display:none">
			<td class="TableRow2"><div class="divbody">Ƶ��Ŀ¼</div></td>
			<td class="TableRow1"><input type="text" name="ChannelDir" size=20 value='dir'></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">����Ŀ��</div></td>
			<td class="TableRow1">
			<input type="radio" value="0" checked name="LinkTarget"> �����ڴ�&nbsp;&nbsp; 
			<input type="radio" name="LinkTarget" value="1"> �´��ڴ�</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">Ƶ���˵�״̬</div></td>
			<td class="TableRow1">
			<input type="radio" name="IsHidden" value="0" checked> ����&nbsp;&nbsp; 
			<input type="radio" name="IsHidden" value="1"> ����</td>
		</tr>
		<tr>
			<td class="TableRow2">��</td>
			<td class="TableRow1">
			<p align="center"><input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp;
			<input type="submit" value="��������" name="B2" class=Button></td>
		</tr>
		</form>
	</table>
<%
End Sub

Private Sub ChannelEdit()
	Dim Rs_c,tempstr
	Dim Channel_Setting
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_Channel WHERE ChannelID=" & CLng(Request("ChannelID")))
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = "���ݿ���ִ���,û�д�վ��Ƶ��!"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
	Channel_Setting = Split(Rs("Channel_Setting"), "|||")
	tempstr = enchiasp.HtmlRndFileName
%>
	<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
		<tr>
			<th colspan="2" align="left"><img src="images/welcome.gif" width="16" height="17" align="absMiddle"> �༭վ��Ƶ��</th>
		</tr>
		<form method="POST" action="?action=savedit">
		<input type="hidden" name="ChannelID" value="<%=Rs("ChannelID")%>">
		<tr>
			<td width="28%" class="TableRow2"><div class="divbody">Ƶ�����ƣ�</div></td>
			<td width="72%" class="TableRow1">
			<input type="text" name="ChannelName" size="20" value="<%=Rs("ChannelName")%>"></td>
		</tr>
				<tr>
			<td class="TableRow2"><div class="divbody">Ƶ������ģʽ��</div></td>
			<td class="TableRow1">��ɫ�� 
			<select size="1" name="ColorModes">
			<option value="0">��ѡ�������ɫ</option>
<%
	TitleColor = "," & enchiasp.InitTitleColor
	TitleColor = Split(TitleColor, ",")
	For i = 1 To UBound(TitleColor)
		Response.Write ("<option style=""background-color:"& TitleColor(i) &";color: "& TitleColor(i) &""" value='"& i &"'")
		If Rs("ColorModes") = i Then Response.Write (" selected")
		Response.Write (">"& TitleColor(i) &"</option>")
	Next
%>
			</select> ���壺
		<select size="1" name="FontModes">
		<option value="0"<%If Rs("FontModes") = 0 Then Response.Write (" selected")%>>��ѡ������</option>
		<option value="1"<%If Rs("FontModes") = 1 Then Response.Write (" selected")%>>����</option>
		<option value="2"<%If Rs("FontModes") = 2 Then Response.Write (" selected")%>>б��</option>
		<option value="3"<%If Rs("FontModes") = 3 Then Response.Write (" selected")%>>�»���</option>
		<option value="4"<%If Rs("FontModes") = 4 Then Response.Write (" selected")%>>����+б��</option>
		<option value="5"<%If Rs("FontModes") = 5 Then Response.Write (" selected")%>>����+�»���</option>
		<option value="6"<%If Rs("FontModes") = 6 Then Response.Write (" selected")%>>б��+�»���</option>
		
		</select></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">Ƶ��ע�ͣ�</div></td>
			<td class="TableRow1">
			<input type="text" name="Caption" size="60" value="<%=Rs("Caption")%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">����Ŀ�꣺</div></td>
			<td class="TableRow1">
			<input type="radio" name="LinkTarget" value="0"<%If Rs("LinkTarget") = 0 Then Response.Write (" checked")%>> �����ڴ�&nbsp;&nbsp; 
			<input type="radio" name="LinkTarget" value="1"<%If Rs("LinkTarget") = 1 Then Response.Write (" checked")%>> �´��ڴ�</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">Ƶ���˵�״̬��</div></td>
			<td class="TableRow1">
			<input type="radio" name="IsHidden" value="0"<%If Rs("IsHidden") = 0 Then Response.Write (" checked")%>> ����&nbsp;&nbsp; 
			<input type="radio" name="IsHidden" value="1"<%If Rs("IsHidden") = 1 Then Response.Write (" checked")%>> ����</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">�������ͣ�</div></td>
			<td class="TableRow1">
			<%If Rs("ChannelType") = 0 Then%>
			<input type="radio" name="ChannelType" value="0" checked> ϵͳƵ��
			<%ElseIf Rs("ChannelType") = 1 Then%>
			<input type="radio" name="ChannelType" value="1"<%If Rs("ChannelType") = 1 Then Response.Write (" checked")%>> �ڲ�Ƶ��&nbsp;&nbsp; 
			<%Else%>
			<input type="radio" name="ChannelType" value="2"<%If Rs("ChannelType") = 2 Then Response.Write (" checked")%>> �ⲿƵ��
			<%End IF%></td>
		</tr>
		<tr id=ChannelSetting1<%If Rs("ChannelType") = 0 Or Rs("ChannelType") = 1 Then Response.Write (" style=""display:'none'""")%>>
			<td class="TableRow2"><div class="divbody">Ƶ������URL��</div></td>
			<td class="TableRow1">
			<input type="text" name="ChannelUrl" size="45" value="<%=Rs("ChannelUrl")%>"> <font color="#FF0000">
			* �ⲿ����URL�ԡ�http://����ͷ</font></td>
		</tr>
		<tr id=ChannelSetting2<%If Rs("ChannelType") => 2 Then Response.Write (" style=""display:'none'""")%>>
		<td class="TableRow1" colspan="2"><fieldset style="cursor: default"><legend>&nbsp;ϵͳƵ������</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
			<tr>
				<td class="TableRow2"><div class="divbody">�Ƿ�رձ�Ƶ����</div></td>
				<td class="TableRow1">
				<input type="radio" name="StopChannel" value="0"<%If Rs("StopChannel") = 0 Then Response.Write (" checked")%>> ��&nbsp;&nbsp; 
				<input type="radio" name="StopChannel" value="1"<%If Rs("StopChannel") = 1 Then Response.Write (" checked")%>> �ر�&nbsp;&nbsp; </td>
			</tr>
<%If (Rs("modules") = 6 or Rs("modules") =7) Then %>	
<input  type=hidden type="text"  name="ModuleName" size="10" value="<%=Rs("ModuleName")%>">
<input type=hidden type="text"   name="modules" size="10" value="<%=Rs("modules")%>"> 
<input type=hidden type="text"   name="ChannelSkin" size="10" value="<%=Rs("ChannelSkin")%>"> 

<% else %>		
<tr>
				<td class="TableRow2"><div class="divbody">Ƶ��ģ�����ƣ�</div></td>
				<td class="TableRow1">
				<input type="text" name="ModuleName" size="10" value="<%=Rs("ModuleName")%>"></td>
			</tr>

			<tr>
				<td width="28%" class="TableRow1"><div class="divbody">Ƶ������ģ�飺</div></td>
				<td width="72%" class="TableRow1">
					<select size="1" name="modules" disabled>
<%
		Response.Write "	<option value=0"
		If Rs("modules") = 0 Then Response.Write (" selected")
		Response.Write ">�ⲿ</option>"
		strModuleName = Split(strModuleName,"|||")
		sChannelID = Split(sChannelID,"|||")
		For i = 0 To UBound(strModuleName) - 1
			Response.Write "	<option value="
			Response.Write sChannelID(i)
			If Rs("modules") = Clng(sChannelID(i)) Then Response.Write (" selected")
			Response.Write ">"
			Response.Write strModuleName(i)
			Response.Write "</option>"
		Next

	Response.Write "					</select></td>"
	Response.Write "			</tr>"
	Response.Write "			<tr>"
	Response.Write "				<td class=""TableRow1""><div class=""divbody"">Ƶ��Ĭ��ģ�壺</div></td>"
	Response.Write "				<td class=""TableRow1"">"
	Response.Write "				<select size=""1"" name=""ChannelSkin"">"

	Response.Write "		<option value=""0"""
	If Rs("ChannelSkin") = 0 Then Response.Write " selected"
	Response.Write ">ʹ��Ĭ��ģ��</option>" & vbCrLf
	SQL = "Select skinid,page_name,isDefault From ECCMS_Template Where pageid = 0 order by TemplateID"
	Set RsObj = enchiasp.Execute(SQL)
	If RsObj.bof And RsObj.EOF Then
		Response.Write "		<option value=""0"">����û������κ�ģ���ļ�</option>" & vbCrLf
	Else
		Do While Not RsObj.EOF
			Response.Write "		<option value=""" & RsObj("skinid") & """"
			If Rs("ChannelSkin") = RsObj("skinid") Then Response.Write " selected"
			Response.Write ">"
			Response.Write RsObj("page_name")
			Response.Write "</option>" & vbCrLf
			RsObj.movenext
		Loop
	End IF
	Set RsObj = Nothing
%>		</select></td>
			</tr>
<%end if%>
			<tr>
				<td class="TableRow2"><div class="divbody">Ƶ������Ŀ¼��</div></td>
				<td class="TableRow1">
				<input type="text" name="ChannelDir" size="20" value="<%=Rs("ChannelDir")%>"> <font color="#FF0000">
				* ���Ҫ�޸�Ƶ������Ŀ¼�����ֹ��޸���Ӧ��Ŀ¼����</font></td>
			</tr>
			<tr style="display:none">
				<td class="TableRow1"><div class="divbody">�Ƿ����������󶨹��ܣ�</div></td>
				<td class="TableRow1">
				<input type="radio" name="BindDomain" value="0"<%If Rs("BindDomain") = 0 Then Response.Write (" checked")%> onClick="setBindDomain.style.display='none';"> ��&nbsp;&nbsp; 
				<input type="radio" name="BindDomain" value="1"<%If Rs("BindDomain") = 1 Then Response.Write (" checked")%> onClick="setBindDomain.style.display='';"<%If Rs("ChannelID") = 5 Then Response.Write (" disabled")%>> ��&nbsp;&nbsp;
				<font color=blue>* ������������󶨹��ܣ���Ƶ�����������õ��������ʱ�Ƶ��</font></td>
			</tr>
			<tr id="setBindDomain"<%If Rs("BindDomain") = 0 Then Response.Write (" style=""display:none""")%>>
				<td class="TableRow2"><div class="divbody">Ƶ�����󶨵�������</div></td>
				<td class="TableRow1">
				<input type="text" name="DomainName" size="40" value="<%=Rs("DomainName")%>"> 
				<br><font color="#FF0000">* ��������Ҫ�󶨵��������磺http://www.enchi.com.cn/</font></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">�Ƿ�����HTML��</div></td>
				<td class="TableRow1">
				<input type="radio" name="IsCreateHtml" value="0"<%If Rs("IsCreateHtml") = 0 Then Response.Write (" checked")%>> ��&nbsp;&nbsp; 
				<input type="radio" name="IsCreateHtml" value="1"<%If Rs("IsCreateHtml") = 1 Then Response.Write (" checked")%><%If Rs("ChannelID") = 4 Then Response.Write (" disabled")%>> ��</td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">����HTML�ļ�����չ����</div></td>
				<td class="TableRow1"><input type="text" name="HtmlExtName" size="10" value="<%=Rs("HtmlExtName")%>"> <font color=blue>* �磺��.html������.htm������.shtml������.asp��</font></td>
			</tr>
<%If (Rs("modules") = 6 or Rs("modules") = 7) Then %>
<input type=hidden type="text"   name="HtmlPrefix" size="10" value="<%=Rs("HtmlPrefix")%>"> 
<input type=hidden type="text"   name="HtmlPath" size="10" value="<%=Rs("HtmlPath")%>"> 
<input type=hidden type="text"   name="HtmlForm" size="10" value="<%=Rs("HtmlForm")%>"> 
<% else %>
			<tr>
				<td class="TableRow1"><div class="divbody">����HTML�ļ���ǰ׺��</div></td>
				<td class="TableRow1"><input type="text" name="HtmlPrefix" size="10" value="<%=Rs("HtmlPrefix")%>"> <font color=blue>* ��ʽ�磺��<%=Rs("HtmlPrefix")%>12345.html������<%=Rs("HtmlPrefix")%>list123_1.html��</font></td>
			</tr>
			
			<tr>
				<td class="TableRow1"><div class="divbody">�����ڱ���HTML�ļ���·����ʽ��</div></td>
				<td class="TableRow1">
				<select  size="1" name="HtmlPath" onChange="chkselect(options[selectedIndex].value,'know2');">
				<option value="0"<%If Rs("HtmlPath") = 0 Then Response.Write (" selected")%>>��ʹ������Ŀ¼</option>
				
				<option value="1"<%If Rs("HtmlPath") = 1 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,1)%></option>
				<option value="2"<%If Rs("HtmlPath") = 2 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,2)%></option>
				<option value="3"<%If Rs("HtmlPath") = 3 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,3)%></option>
				<option value="4"<%If Rs("HtmlPath") = 4 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,4)%></option>
				<option value="5"<%If Rs("HtmlPath") = 5 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,5)%></option>
				<option value="6"<%If Rs("HtmlPath") = 6 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,6)%></option>
				<option value="7"<%If Rs("HtmlPath") = 7 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,7)%></option>
				<option value="8"<%If Rs("HtmlPath") = 8 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,8)%></option>
				
				
				</select> <font color=blue>��Ŀ¼�Ǹ���������ݵ��������ɣ�����ڸ�����Ŀ¼����,��ҳ��ͼ��Ƶ����Ч</font><div id="know2" style="color: red;font-weight:bold;"></div></td>
			</tr>
		
			<tr>
				<td class="TableRow1"><div class="divbody">����HTML�ļ��ĸ�ʽ��</div></td>
				<td class="TableRow1">
				<select size="1" name="HtmlForm" onChange="chkselect(options[selectedIndex].value,'know1');">
				<option value="0"<%If Rs("HtmlForm") = 0 Then Response.Write (" selected")%>>���ں�ʱ��</option>
				<option value="1"<%If Rs("HtmlForm") = 1 Then Response.Write (" selected")%>><%=sModuleName%>ID</option>
				<option value="2"<%If Rs("HtmlForm") = 2 Then Response.Write (" selected")%>>�ļ�ǰ׺+<%=sModuleName%>ID</option>
				<option value="3"<%If Rs("HtmlForm") = 3 Then Response.Write (" selected")%>>����+<%=sModuleName%>ID</option>
				<option value="4"<%If Rs("HtmlForm") = 4 Then Response.Write (" selected")%>>�����+<%=sModuleName%>ID</option>
				</select><div id="know1" style="color: red;font-weight:bold;"></div></td>
			</tr>
<%end if%>
			<tr>
				<td class="TableRow1"><div class="divbody">�Ƿ������û��ϴ��ļ���</div></td>
				<td class="TableRow1">
				<input type="radio" name="StopUpload" value="1"<%If Rs("StopUpload") = 1 Then Response.Write (" checked")%>> ��&nbsp;&nbsp; 
				<input type="radio" name="StopUpload" value="0"<%If Rs("StopUpload") = 0 Then Response.Write (" checked")%>> ��</td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">�����ϴ��ļ��Ĵ�С��</div></td>
				<td class="TableRow1"><input type="text" name="MaxFileSize" size="10" value="<%=Rs("MaxFileSize")%>"> <b>KB</b><font color=red>���벻Ҫ����<%=Cstr(enchiasp.UploadFileSize)%>KB�����Ҫ���������[��������]���޸��ϴ��ļ���С�����ޣ�</font></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">�����ϴ��ļ������ͣ�<br>�����ļ�����֮���á�|���ָ�</div></td>
				<td class="TableRow1"><input type="text" name="UpFileType" size="60" value="<%=Rs("UpFileType")%>"></td>
			</tr>
			
			
<%If (Rs("modules") = 6 or Rs("modules") = 7) Then %>

<input type=hidden type="text"  name="AppearGrade" size="10" value="<%=Rs("AppearGrade")%>">
<input type=hidden type="text"  name="PostGrade" size="10" value="<%=Rs("PostGrade")%>">

<input type=hidden name="IsAuditing" value="<%=Rs("IsAuditing")%>">



<% else %>	
<tr>
				<td class="TableRow1"><div class="divbody">�Ƿ�����˹��ܣ�</div></td>
				<td class="TableRow1">
				<input type="radio" name="IsAuditing" value="0"<%If Rs("IsAuditing") = 0 Then Response.Write (" checked")%>> �ر�&nbsp;&nbsp; 
				<input type="radio" name="IsAuditing" value="1"<%If Rs("IsAuditing") = 1 Then Response.Write (" checked")%>> ��</td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">
<%
				If Rs("ChannelID") = 4 Then
					Response.Write "��������"
				Else
					Response.Write "��������"
				End If
%>
���û��ȼ���</div></td>
				<td class="TableRow1"><select size="1" name="AppearGrade">
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & """"
		If Rs("AppearGrade") = RsObj("Grades") Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>		</select><font color=red>��������HTML�ļ�����Ч��</font></td>
			</tr>
				
			<tr>
				<td class="TableRow1"><div class="divbody">
				<%
				If Rs("ChannelID") = 4 Then
					Response.Write "�ظ�����"
				Else
					Response.Write "����" & sModuleName
				End If
				%>���û��ȼ���</div></td>
				<td class="TableRow1"><select size="1" name="PostGrade">
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & """"
		If Rs("PostGrade") = RsObj("Grades") Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>		</select></td>
			</tr>
<%end if%>
<%If (Rs("modules") = 6 or Rs("modules") = 7) Then %>
<input type=hidden type="text"  name="LeastString" size="10" value="<%=Rs("LeastString")%>">
<input type=hidden type="text"   name="MaxString" size="10" value="<%=Rs("MaxString")%>">
	<%If Rs("modules") = 7 Then %>
	<tr>
				<td class="TableRow1"><div class="divbody">ÿҳ��ʾ�б�����</div></td>
				<td class="TableRow1"><input type="text" name="PaginalNum" size="10" value="<%=Rs("PaginalNum")%>"></td>
			</tr>
	<% else %>
		<input type=hidden name="PaginalNum" size="10" value="<%=Rs("PaginalNum")%>">
	<%end if%>			
			

<input type=hidden type="text"   name="LeastHotHist" size="10" value="<%=Rs("LeastHotHist")%>">
<% else %>
			
			<tr>
				<td class="TableRow1"><div class="divbody">��С���������ַ���</div></td>
				<td class="TableRow1"><input type="text" name="LeastString" size="10" value="<%=Rs("LeastString")%>"></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">������������ַ���</div></td>
				<td class="TableRow1"><input type="text" name="MaxString" size="10" value="<%=Rs("MaxString")%>"></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">ÿҳ��ʾ�б�����</div></td>
				<td class="TableRow1"><input type="text" name="PaginalNum" size="10" value="<%=Rs("PaginalNum")%>"></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">��С���ŵ������</div></td>
				<td class="TableRow1"><input type="text" name="LeastHotHist" size="10" value="<%=Rs("LeastHotHist")%>"></td>
			</tr>
<%end if%>	
<%
If Rs("modules") = 2 Then
%>
			<tr>
				<td class="TableRow1"><div class="divbody">����������л�����</div><br>ÿ�����л������á�|���ֿ�</td>
				<td class="TableRow1"><textarea name="ChannelSetting" cols="60" rows="3"><%=Channel_Setting(0)%></textarea></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">�������Ĭ�����л�����</div></td>
				<td class="TableRow1"><input type="text" name="ChannelSetting" size="60" value="<%=Channel_Setting(1)%>"></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">����������ͣ�</div><br>ÿ������������á�,���ֿ�</td>
				<td class="TableRow1"><textarea name="ChannelSetting" cols="60" rows="3"><%=Channel_Setting(2)%></textarea></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">���������Ȩ��ʽ��</div><br>ÿ����Ȩ��ʽ���á�,���ֿ�</td>
				<td class="TableRow1"><input type="text" name="ChannelSetting" size="60" value="<%=Channel_Setting(3)%>"></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">����������ԣ�</div><br>ÿ������������á�,���ֿ�</td>
				<td class="TableRow1"><input type="text" name="ChannelSetting" size="60" value="<%=Channel_Setting(4)%>"></td>
			</tr>
<%
	Else
		Response.Write "<input type=""hidden"" name=""ChannelSetting"" value=""|||"">"
		Response.Write "<input type=""hidden"" name=""ChannelSetting"" value=""@@@"">"
		Response.Write "<input type=""hidden"" name=""ChannelSetting"" value=""@@@"">"
	End If
%>
		</table></fieldset></td>
		</tr>
		<tr>
			<td class="TableRow2">��</td>
			<td class="TableRow1" align="center"><input type="button" onclick="javascript:history.go(-1)" value="������һҳ" name="B1" class=Button>&nbsp;&nbsp;
			<input type="submit" value="��������" name="B2" class=Button></td>
		</tr>
		</form>
	</table>
<div id="Issubport0" style="display:none"><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),0,"")%></div>
<div id="Issubport1" style="display:none"><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),1,"")%></div>
<div id="Issubport2" style="display:none"><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),2,"")%></div>
<div id="Issubport3" style="display:none"><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),3,"")%></div>
<div id="Issubport4" style="display:none"><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),4,"")%></div>
<div id="Issubport5" style="display:none">��ʹ������Ŀ¼,HTML�ļ������浽����Ŀ¼����<br><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>����Ŀ¼/<%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport6" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>����Ŀ¼/<%=enchiasp.ShowDatePath(tempstr,1)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport7" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>����Ŀ¼/<%=enchiasp.ShowDatePath(tempstr,2)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport8" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>����Ŀ¼/<%=enchiasp.ShowDatePath(tempstr,3)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport9" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>����Ŀ¼/<%=enchiasp.ShowDatePath(tempstr,4)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport10" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>����Ŀ¼/<%=enchiasp.ShowDatePath(tempstr,5)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport11" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>����Ŀ¼/<%=enchiasp.ShowDatePath(tempstr,6)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport12" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>����Ŀ¼/<%=enchiasp.ShowDatePath(tempstr,7)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport13" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>����Ŀ¼/<%=enchiasp.ShowDatePath(tempstr,8)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<SCRIPT LANGUAGE="JavaScript">
<!--
function chkselect(s,divid)
{
	var divname='Issubport';
	var chkreport;
	s=Number(s)
	if (divid=="know1")
	{
		divname=divname+s;
	}
	if (divid=="know2")
	{
		s+=5;
		divname=divname+s;
	}
	document.getElementById(divid).innerHTML=divname;
	chkreport=document.getElementById(divname).innerHTML;
	document.getElementById(divid).innerHTML=chkreport;
}
//-->
</SCRIPT>
<%
Rs.Close:Set Rs = Nothing
End Sub

Private Sub CheckSave()
	If Len(Request.Form("ChannelName")) = 0 Or Len(Request.Form("ChannelName")) => 25 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��վƵ�����Ʋ���Ϊ�ջ��߳���20���ַ���</li>"
	End If
	If Len(Request.Form("ColorModes")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������ɫ��������</li>"
	End If
	If Len(Request.Form("Caption")) = 0 Or Len(Request.Form("Caption")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Ƶ��ע�Ͳ���Ϊ�ջ��߳���200���ַ���</li>"
	End If
	If Len(Request.Form("ChannelUrl")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Ƶ������URL����Ϊ�գ�</li>"
	End If
	
End Sub

Private Sub SavenewChannel()
	CheckSave
	Dim neworders
	If Len(Request.Form("ChannelDir")) = 0 And Request.Form("ChannelType") <> 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Ƶ������Ŀ¼����Ϊ�գ�</li>"
	End If
	ChannelDir = Replace(Replace(Replace(Request.Form("ChannelDir"), "\","/"), " ",""), "'","")
	If Right(ChannelDir, 1) <> "/" Then
		ChannelDir = ChannelDir & "/"
	Else
		ChannelDir = ChannelDir
	End If
	If Request.Form("ChannelType") = 1 Then
		If Request.Form("modules") = 0 Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ѡ����ȷ��ģ�飡</li>"
			Exit Sub
		End If
		Set Rs = Conn.Execute("SELECT ChannelID,ChannelDir FROM ECCMS_Channel WHERE ChannelType=0 And ChannelID=" & CLng(Request.Form("modules")))
		If Rs.EOF And Rs.BOF Then
			ErrMsg = "<li>�Ҳ���ָ��ģ�顣</li>"
			Founderr = True
			Exit Sub
		Else
			mChannelID = Rs("ChannelID")
			mChannelDir = Rs("ChannelDir")
			If LCase(ChannelDir) = LCase(mChannelDir) Then
				ErrMsg = "<li>����ָ����ϵͳƵ����ͬ��Ŀ¼��</li>"
				Founderr = True
				Exit Sub
			End If
		End If
		Set Rs = Nothing
	End If
	
	Set Rs = Conn.Execute("SELECT ChannelID FROM ECCMS_Channel WHERE ChannelID=" & CLng(Request("NewChannelID")))
	If Not (Rs.EOF And Rs.BOF) Then
		ErrMsg = "<li>������ָ���ͱ��Ƶ��һ������š�</li>"
		Founderr = True
		Exit Sub
	Else
		NewChannelID = CLng(Request("NewChannelID"))
	End If
	Set Rs = Nothing
	If NewChannelID = 999 Then NewChannelID = NewChannelID + 1
	If NewChannelID = 9999 Then NewChannelID = NewChannelID + 1
	If Founderr = True Then Exit Sub
	Set Rs = enchiasp.Execute ("SELECT MAX(orders) FROM ECCMS_Channel")
	If Not (Rs.EOF And Rs.bof) Then
		neworders = Rs(0)
	End If
	If IsNull(neworders) Then neworders = 0
	Set Rs = Nothing
	'Call ChannelCopy
	'Succeed("<li>����µ�Ƶ���ɹ�</li>"):exit sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_Channel"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = NewChannelID
		Rs("orders") = neworders + 1
		Rs("ColorModes") = Trim(Request.Form("ColorModes"))
		Rs("FontModes") = Trim(Request.Form("FontModes"))
		Rs("ChannelName") = enchiasp.ChkFormStr(Request.Form("ChannelName"))
		Rs("Caption") = enchiasp.ChkFormStr(Request.Form("Caption"))
		Rs("ChannelDir") = ChannelDir
		Rs("StopChannel") = 0
		Rs("IsHidden") = Trim(Request.Form("IsHidden"))
		Rs("LinkTarget") = Trim(Request.Form("LinkTarget"))
		Rs("ChannelType") = CInt(Request.Form("ChannelType"))
		Rs("ChannelUrl") = Trim(Request.Form("ChannelUrl"))
		Rs("modules") = CInt(Request.Form("modules"))
		Rs("BindDomain") = 0
		Rs("DomainName") = "http://"
		If CInt(Request.Form("ChannelType")) = 1 Then
			Rs("ModuleName") = "��Ƶ��"
		Else
			Rs("ModuleName") = "�ⲿ"
		End If
		
		Rs("ChannelSkin") = 0
		Rs("HtmlPath") = 0
		Rs("HtmlForm") = 3
		Rs("IsCreateHtml") = 0
		Rs("HtmlExtName") = ".html"
		Rs("HtmlPrefix") = "HTML_"
		Rs("StopUpload") = 1
		Rs("MaxFileSize") = 500
		Rs("UpFileType") = "rar|zip|exe|gif|jpg|png|bmp|swf"
		Rs("IsAuditing") = 1
		Rs("AppearGrade") = 0
		Rs("PostGrade") = 0
		Rs("LeastString") = 10
		Rs("MaxString") = 500
		Rs("PaginalNum") = 15
		Rs("LeastHotHist") = 50
		If CInt(Request.Form("modules")) = 2 Then
			Rs("Channel_Setting") = "Win2003/|WinNet/|WinXP/|Win2000/|NT/|WinME/|Win9X/|Linux/|Unix/|Mac/|||Win9X/Win2000/WinXP/Win2003/|||�������,�������,��������,��������|||�������,������,�������,�������,��ʾ���,��ҵ���|||��������,��������,Ӣ��|||"
		Else
			Rs("Channel_Setting") = "|||@@@|||@@@|||"
		End If
	Rs.update
	Rs.Close:Set Rs = Nothing
	enchiasp.DelCahe "ChannelMenu"
	Succeed("<li>����µ�Ƶ���ɹ�</li>")
	If CInt(Request.Form("modules")) > 0 And CInt(Request.Form("ChannelType")) = 1 Then
		Call ChannelCopy
	End If
	
End Sub
Private Sub ChannelCopy()
	Dim newChannelDir,oldChannelDir
	Dim tmpChannel,tmpChannelArray
	oldChannelDir = enchiasp.InstallDir & mChannelDir
	newChannelDir = enchiasp.InstallDir & ChannelDir
	enchiasp.CreatPathEx(newChannelDir & "js")
	enchiasp.CreatPathEx(newChannelDir & "special")
	enchiasp.CreatPathEx(newChannelDir & "UploadPic")
	enchiasp.CreatPathEx(newChannelDir & "UploadFile")
	enchiasp.CopyToFile oldChannelDir & "index.asp",newChannelDir & "index.asp"
	enchiasp.CopyToFile oldChannelDir & "list.asp",newChannelDir & "list.asp"
	enchiasp.CopyToFile oldChannelDir & "show.asp",newChannelDir & "show.asp"
	enchiasp.CopyToFile oldChannelDir & "special.asp",newChannelDir & "special.asp"
	enchiasp.CopyToFile oldChannelDir & "search.asp",newChannelDir & "search.asp"
	enchiasp.CopyToFile oldChannelDir & "showbest.asp",newChannelDir & "showbest.asp"
	enchiasp.CopyToFile oldChannelDir & "showhot.asp",newChannelDir & "showhot.asp"
	enchiasp.CopyToFile oldChannelDir & "shownew.asp",newChannelDir & "shownew.asp"
	enchiasp.CopyToFile oldChannelDir & "comment.asp",newChannelDir & "comment.asp"
	enchiasp.CopyToFile oldChannelDir & "Hits.Asp",newChannelDir & "Hits.Asp"
	enchiasp.CopyToFile oldChannelDir & "RemoveCache.Asp",newChannelDir & "RemoveCache.Asp"
	enchiasp.CopyToFile oldChannelDir & "rssfeed.asp",newChannelDir & "rssfeed.asp"
	enchiasp.CopyToFile oldChannelDir & "js/ShowPage.JS",newChannelDir & "js/ShowPage.JS"
	enchiasp.CopyToFile oldChannelDir & "js/Show_Page.JS",newChannelDir & "js/Show_Page.JS"
	tmpChannel = enchiasp.ReadFile("include/Channel.dat")
	tmpChannel = Replace(tmpChannel, "$ChannelID$", NewChannelID,1,-1,1)
	tmpChannelArray = Split(tmpChannel, "@@@")
	If CInt(Request.Form("modules")) = 1 Then
		enchiasp.CopyToFile oldChannelDir & "sendmail.asp",newChannelDir & "sendmail.asp"
		enchiasp.CreatedTextFile newChannelDir & "config.asp",tmpChannelArray(0)
	ElseIf CInt(Request.Form("modules")) = 2 Then
		enchiasp.CopyToFile oldChannelDir & "showtype.asp",newChannelDir & "showtype.asp"
		enchiasp.CopyToFile oldChannelDir & "error.asp",newChannelDir & "error.asp"
		enchiasp.CopyToFile oldChannelDir & "download.asp",newChannelDir & "download.asp"
		enchiasp.CopyToFile oldChannelDir & "softdown.asp",newChannelDir & "softdown.asp"
		enchiasp.CopyToFile oldChannelDir & "previewimg.asp",newChannelDir & "previewimg.asp"
		enchiasp.CreatedTextFile newChannelDir & "config.asp",tmpChannelArray(1)
	'��ҳ��ͼ��
	Elseif CInt(Request.Form("modules")) = 6 then
		enchiasp.CreatedTextFile newChannelDir & "config.asp",tmpChannelArray(0)

	else
		enchiasp.CopyToFile oldChannelDir & "download.asp",newChannelDir & "download.asp"
		enchiasp.CopyToFile oldChannelDir & "down.asp",newChannelDir & "down.asp"
		enchiasp.CopyToFile oldChannelDir & "downfile.asp",newChannelDir & "downfile.asp"
		enchiasp.CopyToFile oldChannelDir & "play.html",newChannelDir & "play.html"
		enchiasp.CreatedTextFile newChannelDir & "config.asp",tmpChannelArray(2)
	End If
	Dim rstmp,i
	Dim TemplateDir,TemplateFields,TemplateValues
	Set rstmp = enchiasp.Execute("SELECT * FROM ECCMS_Template WHERE ChannelID=" & CLng(Request.Form("modules")))
	SQL=rstmp.GetRows(-1)
	Set rstmp = Nothing
	For i=0 To Ubound(SQL,2)
		TemplateDir = ""
		TemplateFields = "ChannelID,skinid,pageid,TemplateDir,page_name,page_content,page_setting,Template_Help,isDefault"
		TemplateValues = "" & NewChannelID & ","& SQL(2,i) &"," & SQL(3,i) & ",'" & TemplateDir & "','" & enchiasp.CheckStr(SQL(5,i)) & "','" & enchiasp.CheckStr(SQL(6,i)) & "','" & enchiasp.CheckStr(SQL(7,i)) & "','" & enchiasp.CheckStr(SQL(8,i)) & "'," & SQL(9,i) & ""
		Conn.Execute ("INSERT INTO ECCMS_Template (" & TemplateFields & ") VALUES (" & TemplateValues & ")")
	Next
	SQL=Null
End Sub

Private Sub SaveditChannel()
	CheckSave
	Dim HtmlExtName,sDomainName
	If Len(Request.Form("ChannelDir")) = 0 And Request.Form("ChannelType") <> 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Ƶ������Ŀ¼����Ϊ�գ�</li>"
	End If
	ChannelDir = Replace(Replace(Replace(Request.Form("ChannelDir"), "\","/"), " ",""), "'","")
	If Right(ChannelDir, 1) <> "/" Then
		ChannelDir = ChannelDir & "/"
	Else
		ChannelDir = ChannelDir
	End If
	If Trim(Request.Form("IsCreateHtml")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ���Ƿ�����HTML�ļ���</li>"
	End If
	If Left(Trim(Request.Form("HtmlExtName")),1) <> "." Then
		HtmlExtName = "." & Trim(Request.Form("HtmlExtName"))
	Else
		HtmlExtName = Trim(Request.Form("HtmlExtName"))
	End If
	If Not enchiasp.IsValidChar(Request.Form("HtmlExtName")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�ļ���չ���к��зǷ��ַ����������ַ���</li>"
	End If
	If Not enchiasp.IsValidChar(ChannelDir) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Ƶ��Ŀ¼�к��зǷ��ַ����������ַ���</li>"
	End If
	If Not IsNumeric(Request("MaxFileSize")) Then
		ErrMsg = ErrMsg & "<li>�ϴ��ļ���С��ʹ��������</li>"
		Founderr = True
	End If
	if  CLng(Request("MaxFileSize"))>  CLng(enchiasp.UploadFileSize) then
		ErrMsg = ErrMsg & "<li>�ϴ��ļ���С����ϵͳ���õ�"&CLng(enchiasp.UploadFileSize)&"KB������б�Ҫ���޸�[��������]�е�[ϵͳ��������]��</li>"
		Founderr = True

	end if

	If Not IsNumeric(Request("LeastString")) Then
		ErrMsg = ErrMsg & "<li>��С�ַ���ʹ��������</li>"
		Founderr = True
	End If
	If Not IsNumeric(Request("MaxString")) Then
		ErrMsg = ErrMsg & "<li>����ַ���ʹ��������</li>"
		Founderr = True
	End If
	If Not IsNumeric(Request("PaginalNum")) Then
		ErrMsg = ErrMsg & "<li>ÿҳ��ʾ�б�����ʹ��������</li>"
		Founderr = True
	End If
	If Not IsNumeric(Request("LeastHotHist")) Then
		ErrMsg = ErrMsg & "<li>��С���ŵ������ʹ��������</li>"
		Founderr = True
	End If
	sDomainName = Replace(Replace(Replace(Request.Form("DomainName"), "\","/"), " ",""), "'","")
	If Right(sDomainName, 1) <> "/" Then
		sDomainName = sDomainName & "/"
	Else
		sDomainName = sDomainName
	End If
	Dim TempStr, ChannelSetting
	For Each TempStr In Request.Form("ChannelSetting")
			ChannelSetting = ChannelSetting & Replace(TempStr, "|||", "") & "|||"
	Next
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Channel where ChannelID = " & Request("ChannelID")
	Rs.Open SQL,Conn,1,3
		Rs("ColorModes") = Trim(Request.Form("ColorModes"))
		Rs("FontModes") = Trim(Request.Form("FontModes"))
		Rs("ChannelName") = enchiasp.ChkFormStr(Request.Form("ChannelName"))
		Rs("Caption") = enchiasp.ChkFormStr(Request.Form("Caption"))
		Rs("ChannelDir") = Trim(ChannelDir)
		Rs("StopChannel") = Trim(Request.Form("StopChannel"))
		Rs("IsHidden") = Trim(Request.Form("IsHidden"))
		Rs("LinkTarget") = Trim(Request.Form("LinkTarget"))
		Rs("ChannelType") = Trim(Request.Form("ChannelType"))
		Rs("ChannelUrl") = Trim(Request.Form("ChannelUrl"))
		Rs("ModuleName") = Trim(Request.Form("ModuleName"))
		Rs("BindDomain") = Trim(Request.Form("BindDomain"))
		Rs("DomainName") = Trim(sDomainName)
		Rs("ChannelSkin") = Trim(Request.Form("ChannelSkin"))
		Rs("HtmlPath") = Trim(Request.Form("HtmlPath"))
		Rs("HtmlForm") = Trim(Request.Form("HtmlForm"))
		Rs("IsCreateHtml") = Trim(Request.Form("IsCreateHtml"))
		Rs("HtmlExtName") = HtmlExtName
		Rs("HtmlPrefix") = Trim(Request.Form("HtmlPrefix"))
		Rs("StopUpload") = Trim(Request.Form("StopUpload"))
		Rs("MaxFileSize") = CLng(Request.Form("MaxFileSize"))
		Rs("UpFileType") = Trim(Request.Form("UpFileType"))
		Rs("IsAuditing") = Trim(Request.Form("IsAuditing"))
		Rs("AppearGrade") = Trim(Request.Form("AppearGrade"))
		Rs("PostGrade") = Trim(Request.Form("PostGrade"))
		Rs("LeastString") = CLng(Request.Form("LeastString"))
		Rs("MaxString") = CLng(Request.Form("MaxString"))
		Rs("PaginalNum") = CInt(Request.Form("PaginalNum"))
		Rs("LeastHotHist") = CLng(Request.Form("LeastHotHist"))
		Rs("Channel_Setting") = Trim(ChannelSetting)
	Rs.update
	Rs.Close
	Set Rs = Nothing
	Call RemoveCache
	Succeed("<li>�޸�Ƶ�����óɹ���</li>")
End Sub

Private Sub ChannelDel()
	If Request("ChannelID") = "" Then
		ErrMsg = "<li>��ѡ����ȷ��Ƶ��ID��</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") < 10 Then
		ErrMsg = "<li>��Ƶ��Ϊϵͳ��ʼƵ������ɾ������ѡ������Ƶ��ɾ����</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT ClassID FROM [ECCMS_Classify] WHERE ChannelID=" & CLng(Request("ChannelID")))
	If Not (Rs.BOF And Rs.EOF) Then
		Set Rs = Nothing
		ErrMsg = "<li>��Ƶ������ʹ���в���ɾ�������Ҫɾ����Ƶ��������ɾ�����з��ࡣ</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = Nothing
	Set Rs = enchiasp.Execute("SELECT ChannelDir,ChannelType FROM [ECCMS_Channel] WHERE ChannelID=" & CLng(Request("ChannelID")))
	If Not (Rs.BOF And Rs.EOF) Then
		If Rs("ChannelType") = 0 Then
			Set Rs = Nothing
			ErrMsg = "<li>��Ƶ����ϵͳƵ������ɾ����</li>"
			Founderr = True
			Exit Sub
		Else
			enchiasp.FolderDelete(enchiasp.InstallDir & Rs("ChannelDir"))
			Conn.Execute("DELETE FROM ECCMS_Template WHERE ChannelID=" & CLng(Request("ChannelID")))
		End If
	End If
	Set Rs = Nothing
	Call RemoveCache
	
	Conn.Execute("DELETE FROM ECCMS_Channel WHERE ChannelID=" & CLng(Request("ChannelID")))
	Succeed("<li>Ƶ��ɾ���ɹ���</li>")
End Sub
Private Sub ChannelOrders()
	Dim trs
	Dim uporders
	Dim doorders
	Response.Write " <table border=""0"" cellspacing=""1"" cellpadding=""2"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th colspan=2>Ƶ�����������޸�"
	Response.Write " </th>"
	Response.Write " </tr>" & vbCrLf
	SQL = "select * from ECCMS_Channel order by orders"
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		Response.Write "����û�������Ӧ��Ƶ����"
	Else
		Do While Not Rs.EOF
			Response.Write "<form action=?action=saveorder method=post><tr><td width=""50%"" class=TableRow1>" & vbCrLf
			Response.Write enchiasp.ReadFontMode(Rs("ChannelName"),Rs("ColorModes"),Rs("FontModes"))
			Response.Write "</td><td width=""50%"" class=TableRow2>" & vbCrLf
			Set trs = enchiasp.Execute("select count(*) from ECCMS_Channel where orders<" & Rs("orders") & "")
				uporders = trs(0)
				If IsNull(uporders) Then uporders = 0

				Set trs = enchiasp.Execute("select count(*) from ECCMS_Channel where orders>" & Rs("orders") & "")
				doorders = trs(0)
				If IsNull(doorders) Then doorders = 0
				If uporders > 0 Then
					Response.Write "<select name=uporders size=1><option value=0>��</option>" & vbCrLf
					For i = 1 To uporders
						Response.Write "<option value=" & i & ">��" & i & "</option>" & vbCrLf
					Next
					Response.Write "</select>"
				End If
				If doorders > 0 Then
					If uporders > 0 Then Response.Write "&nbsp;"
					Response.Write "<select name=doorders size=1><option value=0>��</option>" & vbCrLf
					For i = 1 To doorders
						Response.Write "<option value=" & i & ">��" & i & "</option>" & vbCrLf
					Next
					Response.Write "</select>" & vbCrLf
				End If
				If doorders > 0 Or uporders > 0 Then
					Response.Write "<input type=hidden name=""ChannelID"" value=""" & Rs("ChannelID") & """>&nbsp;<input type=submit name=Submit class=button value='�� ��'>" & vbCrLf
				End If
			Response.Write "</td></tr></form>" & vbCrLf
			Rs.movenext
		Loop
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>"
End Sub

Private Sub SaveOrder()
	Dim orders
	Dim uporders
	Dim doorders
	Dim oldorders
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("uporders") <> "" And Not CInt(Request("uporders")) = 0 Then
		If Not IsNumeric(Request("uporders")) Then
			ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
			Founderr = True
			Exit Sub
		ElseIf CInt(Request("uporders")) = 0 Then
			ErrMsg = ErrMsg & "<li>��ѡ��Ҫ���������֣�</li>"
			Founderr = True
			Exit Sub
		End If
		Set Rs = enchiasp.Execute("select ChannelID,orders from ECCMS_Channel where ChannelID=" & Request("ChannelID"))
		orders = Rs(1)
		i = 0
		oldorders = 0
		Set Rs = enchiasp.Execute("select ChannelID,orders from ECCMS_Channel where orders<" & orders & " order by orders desc")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("uporders")) >= i Then
				enchiasp.Execute ("update ECCMS_Channel set orders=" & orders & "+" & oldorders & " where ChannelID=" & Rs(0))
				If CInt(Request("uporders")) = i Then uporders = Rs(1)
			End If
			orders = Rs(1)
			Rs.movenext
		Loop
		enchiasp.Execute ("update ECCMS_Channel set orders=" & uporders & " where ChannelID=" & Request("ChannelID"))
		Set Rs = Nothing
	ElseIf Request("doorders") <> "" Then
		If Not IsNumeric(Request("doorders")) Then
			ErrMsg = ErrMsg & "<li>�Ƿ��Ĳ�����</li>"
			Founderr = True
			Exit Sub
		ElseIf CInt(Request("doorders")) = 0 Then
			ErrMsg = ErrMsg & "<li>��ѡ��Ҫ�½������֣�</li>"
			Founderr = True
			Exit Sub
		End If
		Set Rs = enchiasp.Execute("select ChannelID,orders from ECCMS_Channel where ChannelID=" & Request("ChannelID"))
		orders = Rs(1)
		i = 0
		oldorders = 0
		Set Rs = enchiasp.Execute("select ChannelID,orders from ECCMS_Channel where orders>" & orders & " order by orders desc")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("doorders")) >= i Then
				enchiasp.Execute ("update ECCMS_Channel set orders=" & orders & " where ChannelID=" & Rs(0))
				If CInt(Request("doorders")) = i Then doorders = Rs(1)
			End If
			orders = Rs(1)
			Rs.movenext
		Loop
		enchiasp.Execute ("update ECCMS_Channel set orders=" & doorders & " where ChannelID=" & Request("ChannelID"))
		Set Rs = Nothing
	End If
	Call RemoveCache
	Response.redirect "admin_channel.asp?action=orders"
End Sub

Private Sub UpdateStop()
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") <> "" And Request("EditID") <> ""  Then
		enchiasp.Execute ("update ECCMS_Channel set StopChannel=" & CInt(Request("EditID")) & " where ChannelID=" & Request("ChannelID"))
		Call RemoveCache
	Else
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("EditID") <> 0  Then
		OutHintScript("��ϲ������Ƶ���ѳɹ��رա�")
	Else
		OutHintScript("��ϲ������Ƶ���ѳɹ��򿪡�")
	End If
End Sub

Private Sub UpdateHidden()
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") <> "" And Request("EditID") <> ""  Then
		enchiasp.Execute ("update ECCMS_Channel set IsHidden=" & CInt(Request("EditID")) & " where ChannelID=" & Request("ChannelID"))
		Call RemoveCache
	Else
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("EditID") <> 0  Then
		OutHintScript("��ϲ��������Ƶ���˵��ɹ���")
	Else
		OutHintScript("��ϲ������ʾƵ���˵��ɹ���")
	End If
End Sub

Private Sub UpdateLinkTarget()
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") <> "" And Request("EditID") <> ""  Then
		enchiasp.Execute ("update ECCMS_Channel set LinkTarget=" & CInt(Request("EditID")) & " where ChannelID=" & Request("ChannelID"))
		Call RemoveCache
		OutHintScript("��ϲ������������Ŀ��ɹ���")
	Else
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
End Sub

Private Sub UpdateCreateHtml()
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") <> "" And Request("EditID") <> ""  Then
		enchiasp.Execute ("update ECCMS_Channel set IsCreateHtml=" & CInt(Request("EditID")) & " where ChannelID=" & Request("ChannelID"))
		Call RemoveCache
	Else
		ErrMsg = ErrMsg & "<li>�Ƿ���ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("EditID") <> 0  Then
		OutHintScript("��ϲ�����򿪴�Ƶ������HTML���ܳɹ���")
	Else
		OutHintScript("��ϲ�����رմ�Ƶ������HTML���ܳɹ���")
	End If
End Sub
Private Sub ReloadChannelCache()
	enchiasp.DelCahe "Channel" & Request("ChannelID")
	enchiasp.DelCahe "MyChannel" & Request("ChannelID")
	enchiasp.DelCahe "ChannelMenu"
	enchiasp.DelCahe "SiteClassMap"
	Response.Write "<script>alert('���»���ɹ���');javascript:history.back(1)</script>"
End Sub
Private Sub RemoveCache()
	enchiasp.DelCahe "Channel" & Request("ChannelID")
	enchiasp.DelCahe "MyChannel" & Request("ChannelID")
	enchiasp.DelCahe "ChannelMenu"
	enchiasp.DelCahe "SiteClassMap"
End Sub

%>