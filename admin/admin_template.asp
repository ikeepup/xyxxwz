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
Dim Rsm,Action,i
Dim ModuleName,MouseStyle,sChannelID

Response.Write "<script language=JavaScript>" & vbCrLf
Response.Write "function Juge(form1)" & vbCrLf
Response.Write "{" & vbCrLf
Response.Write " if (form1.page_name.value == """")" & vbCrLf
Response.Write " {" & vbCrLf
Response.Write "  alert(""������ģ������!"");" & vbCrLf
Response.Write "  form1.page_name.focus();" & vbCrLf
Response.Write "  return (false);" & vbCrLf
Response.Write " }" & vbCrLf
Response.Write " if (form1.TemplateDir.value == """")" & vbCrLf
Response.Write " {" & vbCrLf
Response.Write "  alert(""������ģ��Ŀ¼!"");" & vbCrLf
Response.Write "  form1.TemplateDir.focus();" & vbCrLf
Response.Write "  return (false);" & vbCrLf
Response.Write " }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
Response.Write " <tr>"
Response.Write "   <th colspan=""2""></th>"
Response.Write " </tr>"
Response.Write " <tr>"
Response.Write "   <td colspan=""2"" class=""TableRow1""><strong>ע�⣺</strong><br>"
Response.Write " ��������������½����޸�ģ�壬���Ա༭CSS��ʽ�������½�ģ��ҳ�棻<br>"
Response.Write " �ڵ�ǰ����ʹ�õ�Ĭ��ģ�岻��ɾ����<br>"
Response.Write " ���������Ϊÿ��ҳ����Ʋ�ͬ��ģ�壬������Ӧ��<span class=style2>ģ���������</span>ȡ��ʹ��վ��ͨ����</td>"
Response.Write " </tr>"
Response.Write " <tr>"
Response.Write "   <td width=""10%"" nowrap class=""TableRow2"">����ѡ�</td>"
Response.Write "   <td width=""90%"" class=""TableRow2"">"
Response.Write "<a href=admin_template.asp class=showmeun>ģ�������ҳ</a> | "
Set Rsm = enchiasp.Execute("Select ChannelID,ModuleName From ECCMS_Channel where ChannelType < 2 And ChannelID <> 4 And stopChannel=0 Order By ChannelID Asc")
Do While Not Rsm.EOF
	Response.Write "<a href=?action=manage&ChannelID="
	Response.Write Rsm("ChannelID")
	Response.Write " class=showmeun>"
	Response.Write Rsm("ModuleName")
	Response.Write "ģ�����</a> | "
	sModuleName = sModuleName & Rsm("ModuleName") & "|||"
	sChannelID = sChannelID & Rsm("ChannelID") & "|||"
	Rsm.MoveNext
Loop
Set Rsm = Nothing
Response.Write "<a href=?action=manage&ChannelID=9999 class=showmeun>����ģ�����</a> | "
Response.Write "<a href=admin_loadskin.asp class=showmeun>ģ�嵼��</a> | "
Response.Write "<a href=admin_loadskin.asp?action=load class=showmeun>ģ�嵼��</a>"
Response.Write "</td>"
Response.Write " </tr>"
Response.Write "</table>"
Response.Write "<br>"
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))

If ChannelID > 0 Then
	Set Rsm = enchiasp.Execute("SELECT ChannelID,ModuleName FROM ECCMS_Channel WHERE ChannelType=0 And ChannelID<>9999 And ChannelID=" & ChannelID)
	If Rsm.BOF And Rsm.EOF Then
		ModuleName = "ȫ��"
	Else
		ModuleName = Rsm("ModuleName")
	End If
	Set Rsm = Nothing
Else
	ModuleName = "ȫ��"
End If
MouseStyle = " bgcolor=""#EEEEE6"" onmouseover=""this.style.backgroundColor='#FFFF00';this.style.color='red'"" onmouseout=""this.style.backgroundColor='';this.style.color=''"""
Action = LCase(Request("action"))
If Not ChkAdmin("Template") Then
	Server.Transfer ("showerr.asp")
	Response.End
End If
Select Case Trim(Action)
	Case "del"
		Call DelTemplate
	Case "newtemplate"
		Call NewTemplate
	Case "default"
		Call DefaultTemplate
	Case "editstyle"
		Call EditStyle
	Case "savestyle"
		Call SaveStyle
	Case "set"
		Call SettingTemplate
	Case "saveset"
		Call SaveTemplateSet
	Case "help"
		Call EditTemplateHelp
	Case "savehelp"
		Call SaveTemplateHelp
	Case "manage"
		Call ChannelTemplate
	Case "edit"
		Call EditTemplatePage
	Case "save"
		Call SaveTemplatePage
	Case "newpage"
		Call NewTemplatePage
	Case Else
		Call MainPage
End Select
If FoundErr = True Then
	ReturnError (ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Sub MainPage()
	SQL = "select * from [ECCMS_Template] where ChannelID = 0 And pageid = 0 order by skinid asc"
	Set Rs = enchiasp.Execute(SQL)
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write " <tr>"
	Response.Write "   <th>ģ������</th>"
	Response.Write "   <th>�༭CSS��ʽ</th>"
	Response.Write "   <th>ģ�峣������</th>"
	Response.Write "   <th>�༭ͨ��ģ��</th>"
	Response.Write "   <th>����ѡ��</th>"
	Response.Write " </tr>"
	Do While Not Rs.EOF
		Response.Write " <tr "
		Response.Write MouseStyle
		Response.Write ">"
		Response.Write "   <td align=""center"">"
		If Rs("IsDefault") = 1 Then
			Response.Write "<img src=images/arrow.gif> "
			Response.Write "<a href=?action=manage&skinid="
			Response.Write Rs("skinid")
			Response.Write " class=showmeun>"
		Else
			Response.Write "<a href=?action=manage&skinid="
			Response.Write Rs("skinid")
			Response.Write ">"
		End If
		Response.Write Rs("page_name")
		Response.Write "<a/></td>"
		Response.Write "   <td align=""center""><a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=1>�༭CSS��ʽ</a></td>"
		Response.Write "   <td align=""center""><a href=?action=set&TemplateID=" & Rs("TemplateID") & ">ģ�峣������</a></td>"
		Response.Write "   <td align=""center""><a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=0>�༭ͨ��ģ��</a></td>"
		Response.Write "   <td align=""center"">"
		Response.Write "   <a href=?action=default&skinid=" & Rs("skinid") & " onclick=""{if(confirm('��ȷ��Ҫ����ģ����ΪĬ��ģ����?')){return true;}return false;}"">��ΪĬ��ģ��</a> |"
		Response.Write "   <a href=?action=del&skinid=" & Rs("skinid") & "&TemplateID=" & Rs("TemplateID") & " onclick=""{if(confirm('ģ��ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����ģ����?')){return true;}return false;}"">ɾ��ģ��</a></td>"
		Response.Write " </tr>"
		Rs.MoveNext
	Loop
	Set Rs = Nothing
	Response.Write "<form method=Post name=""myform"" action=""?action=newtemplate"" onSubmit=""return Juge(this)"">"
	Response.Write " <tr>"
	Response.Write "   <td colspan=""5"" align=""center"" class=""TableRow2"">ģ�����ƣ�<input name=""page_name"" type=""text"" size=""20"">"
	Response.Write "   ģ��Ŀ¼��<input name=""TemplateDir"" type=""text"" size=""20"" value=""skin/default/"">"
	Response.Write "   <input type=""submit"" name=""Submit"" value=""�½�ģ��"" class=Button><br>"
	Response.Write "   <strong>ע�⣺</strong>ģ��Ŀ¼�����ϵͳ��Ŀ¼�£�ģ���½��ɹ����뵽��Ӧ��Ƶ��ģ���½���ҳģ��</td>"
	Response.Write " </tr></form>"
	Response.Write "</table>"

End Sub

Sub EditStyle()
	Dim StyleTitle
	Dim PageContent

	If CInt(Request("StyleID")) = 1 Then
		StyleTitle = "�༭CSS��ʽ"
	Else
		StyleTitle = "�༭ģ��ͨ��"
	End If
	Set Rs = enchiasp.Execute("select * From [ECCMS_Template] where pageid = 0 And TemplateID = " & Request("TemplateID"))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	PageContent = Split(Rs("page_content"), "|||")
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write " <tr>"
	Response.Write "   <th colspan=""2"">" & StyleTitle & "���޸��������ñ���߱�һ����ҳ֪ʶ��</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <td Colspan=2 align=right class=TableRow1>"
	Call TemplateJumpList
	Response.Write "</td>"
	Response.Write " </tr><form method=Post name=""myform"" action=""?action=savestyle"" onSubmit=""return Juge(this)"">"
	Response.Write "  <input type=hidden name=TemplateID value=""" & Rs("TemplateID") & """>"
	Response.Write "  <input type=hidden name=StyleID value=""" & Request("StyleID") & """>"
	Response.Write " <tr>"
	Response.Write "   <td width=""10%"" nowrap class=""TableRow2""><strong>ģ������</strong></td>"
	Response.Write "   <td width=""90%"" class=""TableRow1""><input name=""page_name"" type=""text"" size=""20"" value=""" & Rs("page_name") & """>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "   <a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=1&ChannelID=" & ChannelID & " class=showmeun>�༭CSS��ʽ</a> | " & vbCrLf
	Response.Write "   <a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=0&ChannelID=" & ChannelID & " class=showmeun>�༭ͨ��ģ��</a> | " & vbCrLf
	Response.Write "   <a href=?action=set&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & " class=showmeun>ģ���������</a></td>" & vbCrLf
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write "   <td class=""TableRow2""><strong>ģ��Ŀ¼</strong></td>"
	Response.Write "   <td class=""TableRow1""><input name=""TemplateDir"" type=""text"" size=""20"" value=""" & Rs("TemplateDir") & """></td>" & vbCrLf
	Response.Write " </tr>"
	Response.Write " <tr"
	If CInt(Request("StyleID")) <> 1 Then
		Response.Write " style=""display:none"""
	End If
	Response.Write ">"
	Response.Write "   <td nowrap class=""TableRow2""><strong>CSS��ʽ����</strong><br>��ر�ǩ˵��<br><br>{$InstallDir}<br>ϵͳ��Ŀ¼<br><br>{$SkinPath}<br>Ƥ��ͼƬ·��</td>"
	Response.Write "   <td class=""TableRow1""><textarea name=""page_style"" style=""width:100%;"" rows=""30"" wrap=""OFF"" id=page_style>" & Server.HTMLEncode(PageContent(0)) & "</textarea>"
	Response.Write "   <div align=right><a href=""javascript:admin_Size(-15,'page_style')""><img src=""images/minus.gif"" unselectable=on border=0></a> <a href=""javascript:admin_Size(15,'page_style')""><img src=""images/plus.gif"" unselectable=on border=0></div></td>"
	Response.Write " </tr>"
	Response.Write " <tr"
	If CInt(Request("StyleID")) = 1 Then
		Response.Write " style=""display:none"""
	End If
	Response.Write ">"
	Response.Write "   <td nowrap class=""TableRow2""><strong>ģ�嶥��ͨ��</strong></td>"
	Response.Write "   <td class=""TableRow1""><textarea name=""page_content1"" style=""width:100%;"" rows=""20"" wrap=""OFF"" id=content1>" & Server.HTMLEncode(PageContent(1)) & "</textarea>"
	Response.Write "   <div align=right><a href=""javascript:admin_Size(-10,'page_content1')""><img src=""images/minus.gif"" unselectable=on border=0></a> <a href=""javascript:admin_Size(10,'page_content1')""><img src=""images/plus.gif"" unselectable=on border=0></div></td>"
	Response.Write " </tr>"
	Response.Write " <tr"
	If CInt(Request("StyleID")) = 1 Then
		Response.Write " style=""display:none"""
	End If
	Response.Write ">"
	Response.Write "   <td nowrap class=""TableRow2""><strong>ģ��ײ�ͨ��</strong></td>"
	Response.Write "   <td class=""TableRow1""><textarea name=""page_content2"" style=""width:100%;"" rows=""20"" wrap=""OFF"" id=page_content2>" & Server.HTMLEncode(PageContent(2)) & "</textarea>"
	Response.Write "   <div align=right><a href=""javascript:admin_Size(-10,'page_content2')""><img src=""images/minus.gif"" unselectable=on border=0></a> <a href=""javascript:admin_Size(10,'page_content2')""><img src=""images/plus.gif"" unselectable=on border=0></div></td>"
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write "   <td class=""TableRow2""></td>"
	Response.Write "   <td class=""TableRow1"" align=""center""><input type=""button"" name=""Submit4"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" class=Button>        <input type=""submit"" name=""btnSubmit"" value=""��������"" class=Button></td>"
	Response.Write " </tr></form>"
	Response.Write "</table>"
	Set Rs = Nothing

End Sub

Sub SettingTemplate()
	Dim TemplateStr
	Dim TemplateHelpStr
	Dim TempHelpStr
	Dim TempTitleStr
	Dim TempHelpValue
	Dim TempTitleValue

	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ��ID����Ϊ�գ�</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("Select * From ECCMS_Template where TemplateID = " & Request("TemplateID"))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ģ�������</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	
	TemplateStr = Split(Rs("page_setting"), "|||")
	TemplateHelpStr = Split(Rs("Template_Help"), "@@@")
	TempTitleStr = Split(TemplateHelpStr(0), "|||")
	TempHelpStr = Split(TemplateHelpStr(1), "|||")
	
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write "<tr>"
	Response.Write " <th Colspan=2>��ǰģ�� (" & Rs("page_name") & ") ��������</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <td width=""30%"" Class=BodyTitle align=""center"">"
	Response.Write Rs("page_name")
	Response.Write "</td>" & vbCrLf
	Response.Write " <td width=""70%"" Class=BodyTitle align=""center"">"
	If Rs("pageid") <> 0 Then
		Response.Write "<a href=?action=edit&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & " class=showmeun>�༭��ģ�������</a> | "
	Else
		Response.Write "<a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=0&ChannelID=" & ChannelID & " class=showmeun>�༭��ģ��ͨ��</a> | "
	End If
	Response.Write "<a href=?action=manage&ChannelID=" & Rs("ChannelID") & " class=showmeun>����ģ����ҳ</a>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <td Colspan=2 align=right class=TableRow1>"
	Call TemplateJumpList
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<form name=myform method=""post"" action=""?action=saveset&ChannelID=" & ChannelID & """>"
	Response.Write "<input type=""hidden"" name=""TemplateID"" value=""" & Rs("TemplateID") & """>"
	If TemplateStr(UBound(TemplateStr)) = "" Then TemplateStr(UBound(TemplateStr)) = "del"
	For i = 0 To UBound(TemplateStr)
		If i < UBound(TempHelpStr) Then
			TempHelpValue = TempHelpStr(i)
		Else
			TempHelpValue = "//"
		End If
		If i < UBound(TempTitleStr) Then
			TempTitleValue = TempTitleStr(i)
		Else
			TempTitleValue = "��������˵��"
		End If
		Response.Write "<tr>"
		Response.Write " <td class=""TableRow2"">"
		Response.Write "<font color=blue style=""font-family:tahoma"">"
		Response.Write i
		Response.Write "��</font>"
		Response.Write TempTitleValue
		Response.Write " </td>"
		Response.Write " <td class=""TableRow1"">"
		If Rs("pageid") = 0 And i <= 6 And LCase(TemplateStr(i)) <> "del" Then
			Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t" & i & """ value="""
			Response.Write Server.HTMLEncode(TemplateStr(i))
			Response.Write """ size=10> "
			If i <> 0 Then
				Response.Write "<font size=3 color=" & TemplateStr(i) & "><b>��</b></font>"
			End If
		ElseIf LenB(TemplateStr(i)) > 90 Then
			Response.Write "<textarea name=""TemplateStr"" id=""t" & i & """  cols=""80"" rows=""3"">"
			Response.Write Server.HTMLEncode(TemplateStr(i))
			Response.Write "</textarea><br>"
			Response.Write "<a href=""javascript:admin_Size(-10,'t" & i & "')""><img src=""images/minus.gif"" unselectable=""on"" border='0'></a> <a href=""javascript:admin_Size(10,'t" & i & "')""><img src=""images/plus.gif"" unselectable=""on"" border='0'></a> "
		ElseIf LenB(TemplateStr(i)) <= 21 And LCase(TemplateStr(i)) <> "del" Then
			Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t" & i & """ value="""
			Response.Write Server.HTMLEncode(TemplateStr(i))
			Response.Write """ size=20> "
		Else
			Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t" & i & """ value="""
			Response.Write Server.HTMLEncode(TemplateStr(i))
			Response.Write """ size=60> "
		End If
		Response.Write "<INPUT TYPE=""hidden"" NAME=""ReadME"" id=""r" & i & """ value=""" & TempHelpValue & """>"
		Response.Write "<a href=# onclick=""helpscript(r" & i & ");return false;"" class=""helplink""><img src=""images/help.gif"" border=0 title=""������Ĺ��������""></a>"
		Response.Write " </td>"
		Response.Write "</tr>"
	Next
	Response.Write "<tr>"
	Response.Write " <td class=""TableRow2"" align=""center""><a href=""?action=help&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & """><font color=blue>��ģ���������</font></a></td>"
	Response.Write " <td class=""TableRow1"" align=""center""><input type=""button"" name=""Submit4"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" class=Button>          <input type=""submit"" name=""btnSubmit"" value=""��������"" class=Button></td>"
	Response.Write "</tr>"
	Response.Write "</form>"
	Response.Write "<tr>"
	Response.Write " <td Colspan=2 class=""TableRow2""><font color=red><b>���棺</b></font><li><font color=blue>�벻Ҫ���ı��������롰del����������ɾ����Ӧ���������ݣ���ôģ�����ִ��󣬵�����վ�����������ʡ�</font></li></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Set Rs = Nothing

End Sub

Sub SaveTemplateSet()
	Dim TempStr
	Dim TemplateStr

	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ��ID����Ϊ�գ�</li>"
		Exit Sub
	End If
	'��ȡ���е�����

	TemplateStr = ""
	For Each TempStr In Request.Form("TemplateStr")
		If LCase(TempStr) <> "del" Then
			TemplateStr = TemplateStr & Replace(TempStr, "|||", "") & "|||"
		End If
	Next
	TemplateStr = enchiasp.CheckStr(TemplateStr)
	enchiasp.Execute ("update [ECCMS_Template] set page_setting ='" & TemplateStr & "' Where TemplateID =" & Request("TemplateID"))
	Call RemoveCache
	Succeed ("<li>��ϲ�����޸�ģ��������óɹ���</li>")

End Sub

Sub EditTemplateHelp()
	Dim TemplateHelpStr
	Dim TempTitleStr
	Dim TempHelpStr
	'-----------ģ��������ÿ�ʼ----------------
	'�༭ģ�����

	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ��ID����Ϊ�գ�</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("Select TemplateID,page_name,Template_Help From ECCMS_Template where TemplateID = " & Request("TemplateID"))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ģ�������</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	TemplateHelpStr = Split(Rs("Template_Help"), "@@@")
	TempTitleStr = Split(TemplateHelpStr(0), "|||")
	TempHelpStr = Split(TemplateHelpStr(1), "|||")
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write "<tr>"
	Response.Write " <th Colspan=2>��ǰģ�� (" & Rs("page_name") & ") ��������</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <td width=""40%"" Class=TableTitle align=""center"">ģ�����ñ���˵��</td>"
	Response.Write " <td width=""60%"" Class=TableTitle align=""center"">ģ�����ð�����ϸ˵��</td>"
	Response.Write "</tr>"
	Response.Write "<form name=myform method=""post"" action=""?action=savehelp&ChannelID=" & ChannelID & """>"
	Response.Write "<input type=""hidden"" name=""TemplateID"" value=""" & Rs("TemplateID") & """>"
	If TempTitleStr(UBound(TempTitleStr)) = "" Then
		TempTitleStr(UBound(TempTitleStr)) = "del"
	End If
	For i = 0 To UBound(TempTitleStr)
		Response.Write "<tr>" & vbCrLf
		Response.Write Chr(9) & "<td class=""TableRow1"">"
		Response.Write "<input Type=""text"" name=""TempTitleStr"" value="""
		Response.Write Server.HTMLEncode(TempTitleStr(i))
		Response.Write """ size=50> "
		Response.Write "</td>" & vbCrLf
		Response.Write Chr(9) & "<td class=""TableRow1"">"
		If LenB(TempHelpStr(i)) > 70 Then
			Response.Write "<textarea name=""TempHelpStr""  cols=""70"" rows=""3"">"
			Response.Write Server.HTMLEncode(TempHelpStr(i))
			Response.Write "</textarea>"
		Else
			Response.Write "<input Type=""text"" name=""TempHelpStr"" value="""
			Response.Write Server.HTMLEncode(TempHelpStr(i))
			Response.Write """ size=50> "
		End If
		Response.Write "</td>" & vbCrLf
		Response.Write "</tr>" & vbCrLf
	Next
	Response.Write "<tr>"
	Response.Write " <td class=""TableRow2"" align=""center""><a href=""?action=set&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & """><font color=blue>��ģ���������</font></a></td>"
	Response.Write " <td class=""TableRow1"" align=""center""><input type=""button"" name=""Submit4"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" class=Button>          <input type=""submit"" name=""btnSubmit"" value=""�������"" class=Button></td>"
	Response.Write "</tr></form><tr>"
	Response.Write " <td Colspan=2 class=""TableRow2""><font color=blue><b>ע�⣺</b> ���������������Ӧ��ģ��������á�</font><li>�����༭�������������ð��������ڶ�Ӧ���ı��������롰del������ô�������ݵ���žͻ�ǰ�ơ�</li>"
	Response.Write " <li>�������ı�������ݵ����,���Ѹ���Ŀ���������,��ֻ��Ҫ��������ա�</li></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Set Rs = Nothing
End Sub

Sub SaveTemplateHelp()
	Dim TempStr
	Dim HelpStr
	Dim TemplateHelpStr
	Dim TempHelpStr
	Dim TempTitleStr
	'����ģ�����

	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ��ID����Ϊ�գ�</li>"
		Exit Sub
	End If
	'��ȡ���е�����

	TempTitleStr = ""
	For Each TempStr In Request.Form("TempTitleStr")
		If LCase(TempStr) <> "del" Then
			TempTitleStr = TempTitleStr & Replace(TempStr, "|||", "") & "|||"
		End If
	Next
	TempHelpStr = ""
	For Each HelpStr In Request.Form("TempHelpStr")
		TempHelpStr = TempHelpStr & Replace(HelpStr, "|||", "") & "|||"
	Next
	TemplateHelpStr = enchiasp.CheckStr(TempTitleStr & "@@@" & TempHelpStr)
	enchiasp.Execute ("update [ECCMS_Template] set Template_Help ='" & TemplateHelpStr & "' Where TemplateID =" & Request("TemplateID"))
	Call RemoveCache
	OutHintScript ("��ϲ��������ģ������ɹ���")
	'-----------ģ��������ý���----------------
End Sub

Sub SaveStyle()
	Dim TemplateDir
	Dim page_content
	Dim FileName
	Dim FileContent
	

	If Trim(Request.Form("page_name")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ�����Ʋ���Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("TemplateDir")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ��Ŀ¼����Ϊ�գ�</li>"
	End If
	If Not enchiasp.IsValidChar(Request.Form("TemplateDir")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ��Ŀ¼�к��зǷ��ַ����������ַ���</li>"
	End If
	If Right(Request.Form("TemplateDir"), 1) <> "/" Then
		TemplateDir = Trim(Request.Form("TemplateDir")) & "/"
	Else
		TemplateDir = Trim(Request.Form("TemplateDir"))
	End If
	If FoundErr Then Exit Sub
	page_content = enchiasp.CheckStr(Request.Form("page_style") & "|||" & Request.Form("page_content1") & "|||" & Request.Form("page_content2") & "|||")
	enchiasp.Execute ("update [ECCMS_Template] set TemplateDir='" & TemplateDir & "',page_name='" & Trim(Request.Form("page_name")) & "',page_content='" & page_content & "' where TemplateID = " & Request("TemplateID"))
	
	enchiasp.CreatPathEx (enchiasp.InstallDir & TemplateDir)
	FileName = enchiasp.InstallDir & TemplateDir & "style.css"
	FileContent = Request.Form("page_style")
	FileContent = Replace(FileContent, "{$InstallDir}", enchiasp.InstallDir)
	FileContent = Replace(FileContent, "{$SkinPath}", TemplateDir)
	enchiasp.CreatedTextFile FileName, FileContent
	Call RemoveCache
	If CInt(Request("StyleID")) = 1 Then
		SucMsg = ("<li>��ϲ�����༭CSS��ʽ�ɹ���</li>")
	Else
		SucMsg = ("<li>��ϲ�����༭ģ��ͨ���ɹ���</li>")
	End If
	Succeed (SucMsg)
	
End Sub

Sub NewTemplate()
	Dim TemplateDir
	Dim TemplateName
	
	Dim TemplateFields
	Dim TemplateValues
	Dim newskinid

	If Trim(Request.Form("page_name")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ�����Ʋ���Ϊ�գ�</li>"
	End If
	If Trim(Request.Form("TemplateDir")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ��Ŀ¼����Ϊ�գ�</li>"
	End If
	If Not enchiasp.IsValidChar(Request.Form("TemplateDir")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ��Ŀ¼�к��зǷ��ַ����������ַ���</li>"
	End If
	If Right(Request.Form("TemplateDir"), 1) <> "/" Then
		TemplateDir = Trim(Request.Form("TemplateDir")) & "/"
	Else
		TemplateDir = Trim(Request.Form("TemplateDir"))
	End If
	If FoundErr Then Exit Sub
	
	enchiasp.CreatPathEx (enchiasp.InstallDir & TemplateDir)
	'Response.Write "<li>�����½�ģ�塭�� ���Ժ򡭡� �����벻Ҫˢ��ҳ�档</li>"
	TemplateName = enchiasp.CheckStr(Trim(Request("page_name")))
	Set Rs = enchiasp.Execute("Select Max(skinid) from [ECCMS_Template] where pageid = 0")
	If Not (Rs.EOF And Rs.BOF) Then
		newskinid = Rs(0)
	End If
	If IsNull(newskinid) Then newskinid = 0
	Rs.Close
	newskinid = newskinid + 1
	Set Rs = enchiasp.Execute("select * From [ECCMS_Template] where IsDefault = 1")
	If Not (Rs.BOF And Rs.EOF) Then
		Do While Not Rs.EOF
			If Rs("pageid") <> 0 Then
				TemplateName = Rs("page_name")
			End If
			TemplateFields = "ChannelID,skinid,pageid,TemplateDir,page_name,page_content,page_setting,Template_Help,isDefault"
			TemplateValues = "" & Rs("ChannelID") & "," & newskinid & "," & Rs("pageid") & ",'" & TemplateDir & "','" & TemplateName & "','" & enchiasp.CheckStr(Rs("page_content")) & "','" & enchiasp.CheckStr(Rs("page_setting")) & "','" & enchiasp.CheckStr(Rs("Template_Help")) & "',0"
			SQL = "insert into [ECCMS_Template](" & TemplateFields & ")values(" & TemplateValues & ")"
			enchiasp.Execute (SQL)
			Rs.MoveNext
		Loop
	Else
		TemplateValues = "0," & newskinid & ",0,'" & TemplateDir & "','" & TemplateName & "','|||||||||','|||','|||@@@|||',0"
		SQL = "insert into [ECCMS_Template](ChannelID,skinid,pageid,TemplateDir,page_name,page_content,page_setting,Template_Help,isDefault)values(" & TemplateValues & ")"
		enchiasp.Execute (SQL)
	End If
	Set Rs = Nothing
	OutHintScript ("�½�ģ�塰" & Request.Form("page_name") & "���ɹ���")
	
End Sub

Sub DelTemplate()
	Set Rs = enchiasp.Execute("Select IsDefault From ECCMS_Template where TemplateID = " & Request("TemplateID"))
	If Rs(0) = 1 Then
		ErrMsg = ErrMsg + "<li>��ģ����Ĭ��ģ�棬������ɾ����"
		FoundErr = True
		Exit Sub
	Else
		enchiasp.Execute ("Delete From ECCMS_Template where skinid = " & Request("skinid"))
		Application.Contents.RemoveAll
		OutHintScript ("ģ��ɾ���ɹ���")
	End If
	Set Rs = Nothing
End Sub

Sub DefaultTemplate()
	If Trim(Request("skinid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ��ID����Ϊ�գ�</li>"
		Exit Sub
	End If
	enchiasp.Execute ("update ECCMS_Template set isDefault = 0 where isDefault = 1")
	enchiasp.Execute ("update ECCMS_Template set isDefault = 1 where skinid = " & Request("skinid"))
	enchiasp.DelCahe "MainStyle" & Request("skinid")
	enchiasp.DelCahe "DefaultSkinID"
	OutHintScript ("��ϲ��������Ĭ��ģ��ɹ���")
End Sub

Sub ChannelTemplate()
	Dim skinid
	Dim Rss
	skinid = CLng(Request("skinid"))
	If Request("skinid") <> "" And ChannelID = 0 Then
		SQL = "skinid = " & skinid
	ElseIf Request("skinid") <> "" And ChannelID <> 0 Then
		SQL = "(skinid = " & skinid & " And ChannelID = " & ChannelID & ") Or (skinid = " & skinid & " And ChannelID = 0)"
	Else
		Set Rs = enchiasp.Execute("Select * From ECCMS_Template where isDefault = 1 And pageid = 0")
		If ChannelID <> 0 Then
			SQL = "(skinid = " & Rs("skinid") & " And ChannelID = " & ChannelID & ") Or (skinid = " & Rs("skinid") & " And ChannelID = 0)"
		Else
			SQL = "skinid = " & Rs("skinid")
		End If
		skinid = Rs("skinid")
		Set Rs = Nothing
	End If
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write "<tr>"
	Response.Write " <th colspan=2>"
	Response.Write ModuleName
	Response.Write "ģ������б�</th>"
	Response.Write "</tr>"
	Response.Write " <form name=myform method=""post"" action=""?action=manage"">"
	Response.Write " <input type=""hidden"" name=""ChannelID"" value=""" & ChannelID & """>"
	Response.Write " <td class=""TableRow2"">��ѡ��ģ�壺"
	Set Rss = enchiasp.Execute("Select * From ECCMS_Template where pageid = 0")
	Response.Write "<select name=skinid>"
	Do While Not Rss.EOF
		Response.Write " <option value="
		Response.Write Rss("skinid")
		If Rss("skinid") = skinid Then Response.Write " selected"
		Response.Write ">"
		Response.Write Rss("page_name")
		Response.Write "</option>"
		Rss.MoveNext
	Loop
	Set Rss = Nothing
	Response.Write "</select>"
	Response.Write "&nbsp;<input type=submit value=""�� ��"" name=""B1"" class=button>"
	Response.Write " </td>"
	Response.Write " </form>"
	Response.Write " <form name=myform method=""post"" action=""?action=newpage"">"
	Response.Write " <input type=""hidden"" name=""skinid"" value=""" & skinid & """>"
	Response.Write " <td class=""TableRow2"">�½�"
	Response.Write ModuleName
	Response.Write "��ģ��ҳ�棺"
	Response.Write "<select name=pageid>"
	Response.Write " <option value=''>����ѡ��ģ�����͡�</option>"
	Response.Write " <option value=0>��վ��ҳģ��</option>"
	Response.Write " <option value=1>��"
	Response.Write ModuleName
	Response.Write "��ҳ��</option>"
	Response.Write " <option value=2>�����б�ҳ��</option>"
	Response.Write " <option value=3>��������ҳ��</option>"
	Response.Write " <option value=4>����ר��ҳ��</option>"
	Response.Write " <option value=5>�����Ƽ�ҳ��</option>"
	Response.Write " <option value=6>��������ҳ��</option>"
	Response.Write " <option value=7>��������ҳ��</option>"
	Response.Write " <option value=8>��������ҳ��</option>"
	Response.Write "</select> "
	If ChannelID = 0 Then
		Response.Write "��ѡ��Ƶ����"
		Response.Write "<select name=ChannelID>"
		sModuleName = Split(sModuleName, "|||")
		sChannelID = Split(sChannelID, "|||")
		For i = 0 To UBound(sModuleName) - 1
			Response.Write " <option value="
			Response.Write sChannelID(i)
			Response.Write ">"
			Response.Write sModuleName(i)
			Response.Write "</option>"
		Next
		Response.Write "</select>"
	Else
		Response.Write " <input type=""hidden"" name=""ChannelID"" value=""" & ChannelID & """>"
	End If
	Response.Write "<br>ģ�����ƣ�"
	Response.Write "<input type=""text"" name=""pagename"" size=35>"
	'Response.Write "ģ��Ψһ��ʶ��(����Ӣ��)<input type=""text"" name=""pagemark"" size=20>"
	Response.Write "&nbsp;<input type=submit value=""�½���ģ��"" name=""B2"" class=button>&nbsp;"
	Response.Write " </td>"
	Response.Write " </form>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <th width=""40%"">ģ������</th>"
	Response.Write " <th width=""60%"">ģ���������</th>"
	Response.Write "</tr>"
	Set Rs = enchiasp.Execute("Select * From ECCMS_Template where " & SQL & " Order By TemplateID")
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td colspan=2 align=center>û���ҵ����ģ�壡</td></tr>"
	Else
		Response.Write "<tr>"
		Response.Write " <td colspan=2 Class=BodyTitle>��ǰģ�壺<font color=blue>"
		If Rs("pageid") = 0 Then
			Response.Write Rs("page_name")
		End If
		Response.Write "</font></td>"
		Response.Write "</tr>"
		Do While Not Rs.EOF
			Response.Write "<tr "
			Response.Write MouseStyle
			Response.Write ">"
			Response.Write " <td><li>"
			If Rs("ChannelID") = 0 Then
				Response.Write "<font color=blue>"
				Response.Write Rs("page_name")
				Response.Write "</font>"
			ElseIf ChannelID = 0 And Rs("pageid") = 1 Then
				Response.Write "<font color=red>"
				Response.Write Rs("page_name")
				Response.Write "</font>"
			Else
				Response.Write Rs("page_name")
			End If
			Response.Write "</li></td>"
			Response.Write " <td>�༭��ģ�飺 "
			If Rs("pageid") = 0 Then
				Response.Write "<a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=1&ChannelID=" & ChannelID & ">�༭CSS��ʽ</a> | "
				Response.Write "<a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=0&ChannelID=" & ChannelID & ">�����͵ײ�ͨ��</a> | "
				Response.Write "<a href=?action=set&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & ">ģ�峣������</a>"
			Else
				Response.Write "<a href=?action=edit&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & ">�༭ģ�������</a> | "
				Response.Write "<a href=?action=set&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & ">�޸�ģ���������</a>"
			End If
			'If Rs("pageid") = 8 Then
				'Response.Write " | <a href=?action=del&TemplateID=" & Rs("TemplateID") & ">ɾ����ҳ��ģ��</a>"
			'End If
			Response.Write "</td>"
			Response.Write "</tr>"
			Rs.MoveNext
		Loop
		Response.Write "<tr>"
		Response.Write " <td class=""TableRow1""></td>"
		Response.Write " <td class=""TableRow1""></td>"
		Response.Write "</tr>"
	End If
	Set Rs = Nothing
	Response.Write "<form method=Post name=""myform"" action=""?action=newtemplate"" onSubmit=""return Juge(this)"">"
	Response.Write " <tr>"
	Response.Write "   <td colspan=""5"" align=""center"" class=""TableRow2"">ģ�����ƣ�<input name=""page_name"" type=""text"" size=""20"">"
	Response.Write "   ģ��Ŀ¼��<input name=""TemplateDir"" type=""text"" size=""20"" value=""skin/default/"">"
	Response.Write "   <input type=""submit"" name=""Submit"" value=""�½�ģ��"" class=Button><br>"
	Response.Write "   <strong>ע�⣺</strong>ģ��Ŀ¼�����ϵͳ��Ŀ¼�£�ģ���½��ɹ����뵽��Ӧ��Ƶ��ģ���½���ҳģ��</td>"
	Response.Write " </tr></form>"
	Response.Write "</table>"
End Sub

Sub EditTemplatePage()
	Dim page_content
	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("select * From [ECCMS_Template] where TemplateID = " & Request("TemplateID"))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	If Not IsNull(Rs("page_content")) Then
		page_content = Split(Rs("page_content") & "|||@@@|||", "|||@@@|||")
	End If
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write " <tr>"
	Response.Write "   <th colspan=""2"">�༭��ǰģ�壺" & Rs("page_name") & " ���޸��������ñ���߱�һ����ҳ֪ʶ��</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <td Colspan=2 align=right class=TableRow1>"
	Call TemplateJumpList
	Response.Write "</td>"
	Response.Write " </tr><form method=Post name=""myform"" action=""?action=save&ChannelID=" & ChannelID & """>"
	Response.Write "  <input type=hidden name=TemplateID value=""" & Rs("TemplateID") & """>"
	Response.Write "  <input type=hidden name=pageid value=""" & Rs("pageid") & """>"
	Response.Write " <tr>"
	Response.Write "   <td width=""10%"" nowrap class=""TableRow2""><strong>��ǰģ������</strong></td>"
	Response.Write "   <td width=""90%"" class=""TableRow1"">"
	Response.Write "<input type=""text"" name=""pagename"" value="""
	Response.Write Rs("page_name")
	Response.Write """ size=35>"
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "   <a href=?action=set&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & " class=showmeun>�޸ĸ�ģ���������</a> | "
	Response.Write "<a href=?action=manage&ChannelID=" & Rs("ChannelID") & " class=showmeun>����ģ����ҳ</a></td>" & vbCrLf
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write "   <td class=""TableRow2""><strong>���ɱ�ǩ</strong></td>"
	Response.Write "   <td class=""TableRow1"">"
	Response.Write "<a href=javascript:openDialog('admin_label.asp?action=list&ChannelID=1',550,490)>���±�ǩ</a> | "
	Response.Write "<a href=javascript:openDialog('admin_label.asp?action=list&ChannelID=2',550,460)>�����ǩ</a> | "
	Response.Write "<a href=javascript:openDialog('admin_label.asp?action=list&ChannelID=3',550,460)>�̳Ǳ�ǩ</a> | "
	Response.Write "<a href=javascript:openDialog('admin_label.asp?action=list&ChannelID=5',550,460)>������ǩ</a> | "
	Response.Write "<a href=javascript:openDialog('admin_label.asp?action=image&ChannelID=" & ChannelID & "',550,460)>ͼƬ��ǩ</a> | "
	Response.Write "<a href=javascript:openDialog('admin_label.asp',550,460)>ģ���ǩ����</a>"
	Response.Write "</td>" & vbCrLf
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write "   <td width=""10%"" class=""TableRow2""><strong>ģ������</strong><br>��ر�ǩ˵��<br><br>{$InstallDir}<br>ϵͳ��Ŀ¼<br><br>{$SkinPath}<br>Ƥ��ͼƬ·��</td>"
	Response.Write "   <td class=""TableRow1""><textarea name=""content"" style=""width:100%;"" rows=""30"" wrap=""OFF"" id=PageContent>" & Server.HTMLEncode(page_content(0)) & "</textarea>"
	Response.Write "   <div align=right><a href=""javascript:admin_Size(-10,'PageContent')""><img src=""images/minus.gif"" unselectable=on border=0></a> <a href=""javascript:admin_Size(10,'PageContent')""><img src=""images/plus.gif"" unselectable=on border=0></div></td>"
	Response.Write " </tr>"
	If Rs("pageid") = 2 And ChannelID <> 9999 Then
		Response.Write " <tr"
		If ChannelID = 3 Or ChannelModuleID = 5 Then
			Response.Write " style=""display:none"""
		End If
		Response.Write ">"
		Response.Write "   <td width=""10%"" class=""TableRow2""><strong>ģ������</strong><br>˵����<br>��ģ���Ǵ����б�ҳ��ģ�壬�����ֻ��һ�������ģ����ܲ��ñ༭��<br>�磺��ķ�����������ӷ��ࡣ���÷��ʸ��������ʱ�����ʾ��ģ������</td>"
		If ChannelID = 3 Or ChannelModuleID = 5 Then
			Response.Write "   <td class=""TableRow1""><textarea name=""content1"" id=PageContent1></textarea>"
		Else
			Response.Write "   <td class=""TableRow1""><textarea name=""content1"" style=""width:100%;"" rows=""30"" wrap=""OFF"" id=PageContent1>" & Server.HTMLEncode(page_content(1)) & "</textarea>"
		End If
		Response.Write "   <div align=right><a href=""javascript:admin_Size(-10,'PageContent1')""><img src=""images/minus.gif"" unselectable=on border=0></a> <a href=""javascript:admin_Size(10,'PageContent1')""><img src=""images/plus.gif"" unselectable=on border=0></div></td>"
		Response.Write " </tr>"
	End If
	Response.Write " <tr>"
	Response.Write "   <td class=""TableRow2""></td>"
	Response.Write "   <td class=""TableRow1"" align=""center""><input type=""button"" name=""Submit4"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" class=Button>        <input type=""submit"" name=""btnSubmit"" value=""����ģ��"" class=Button></td>"
	Response.Write " </tr></form>"
	Response.Write "</table>"
	Set Rs = Nothing
End Sub

Sub SaveTemplatePage()
	Dim TemplateContent
	Dim page_name

	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ��ID����Ϊ�գ�</li>"
		Exit Sub
	End If
	If Trim(Request("pagename")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ�����ݲ���Ϊ�գ�</li>"
		Exit Sub
	End If
	If Trim(Request("content")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ�����ݲ���Ϊ�գ�</li>"
		Exit Sub
	End If
	TemplateContent = Request.Form("content")
	If Request.Form("pageid") = 2 And ChannelID <> 9999 And ChannelID <> 3 And ChannelModuleID <> 5 Then
		TemplateContent = TemplateContent & "|||@@@|||" & Request.Form("content1")
	End If
	TemplateContent = enchiasp.CheckStr(TemplateContent)
	page_name = enchiasp.CheckStr(Request.Form("pagename"))
	enchiasp.Execute ("update [ECCMS_Template] set page_name = '" & page_name & "', page_content ='" & TemplateContent & "' Where TemplateID =" & Request("TemplateID"))
	Call RemoveCache
	Succeed ("<li>��ϲ�����޸�ģ��������óɹ���</li>")
End Sub

Sub NewTemplatePage()
	Dim Rss
	Dim pageid
	Dim skinid
	Dim TemplateName
	Dim TemplateFields
	Dim TemplateValues
	If Trim(Request("pageid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ��ģ���������ͣ�</li>"
		Exit Sub
	End If
	If Trim(Request("pagename")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ģ�����Ʋ���Ϊ�գ�</li>"
		Exit Sub
	End If
	If Trim(Request("skinid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������</li>"
		Exit Sub
	End If
	If Trim(Request("ChannelID")) = "" Or Request("ChannelID") = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������</li>"
		Exit Sub
	End If
	If CInt(Request("pageid")) = 0 Then
		ChannelID = 0
		pageid = 1
	Else
		ChannelID = CInt(Request("ChannelID"))
		pageid = CInt(Request("pageid"))
	End If
	skinid = CLng(Request("skinid"))
	TemplateName = enchiasp.CheckStr(Trim(Request("pagename")))
	'If pageid <> 8 Then
		Set Rss = enchiasp.Execute("select pageid From [ECCMS_Template] where skinid = " & skinid & " And ChannelID = " & ChannelID & " And pageid = " & pageid)
		If Not (Rss.BOF And Rss.EOF) Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>��ģ�������Ѿ�����,��ѡ����������ģ�壡</li>"
			Exit Sub
		End If
		Set Rss = Nothing
	'End If
	Set Rss = enchiasp.Execute("select * From [ECCMS_Template] where pageid = 0 And IsDefault = 1")
	Set Rs = enchiasp.Execute("select * From [ECCMS_Template] where skinid = " & Rss("skinid") & " And ChannelID = " & ChannelID & " And pageid = " & pageid)
	If Not (Rs.BOF And Rs.EOF) Then
		TemplateFields = "ChannelID,skinid,pageid,page_name,page_content,page_setting,Template_Help,isDefault"
		TemplateValues = "" & ChannelID & "," & skinid & "," & pageid & ",'" & TemplateName & "','" & enchiasp.CheckStr(Rs("page_content")) & "','" & enchiasp.CheckStr(Rs("page_setting")) & "','" & enchiasp.CheckStr(Rs("Template_Help")) & "',0"
		SQL = "insert into [ECCMS_Template](" & TemplateFields & ")values(" & TemplateValues & ")"
	Else
		TemplateValues = "" & ChannelID & "," & skinid & "," & pageid & ",'" & TemplateName & "','|||','1|||','|||@@@|||',0"
		SQL = "insert into [ECCMS_Template](ChannelID,skinid,pageid,page_name,page_content,page_setting,Template_Help,isDefault)values(" & TemplateValues & ")"
	End If
	Set Rs = Nothing
	Set Rss = Nothing
	enchiasp.Execute (SQL)
	OutHintScript ("�½���ģ�塰" & Request.Form("pagename") & "���ɹ���")
End Sub
Sub TemplateJumpList()
	Dim rstmp, tmpsql, tmpname, sel
	Dim strTemp, strContent, strStetting
	strTemp = ""
	On Error Resume Next
	If ChannelID > 0 Then
		If Trim(Request("skinid")) <> "" And Trim(Request("skinid")) <> "0" Then
			tmpsql = "And skinid=" & Trim(Request("skinid"))
		Else
			tmpsql = "And isDefault=1"
		End If
		tmpsql = "SELECT TemplateID,pageid,page_name FROM ECCMS_Template WHERE (ChannelID=0 Or ChannelID=" & ChannelID & ") " & tmpsql & " ORDER BY TemplateID"
		Set rstmp = enchiasp.Execute(tmpsql)
		If rstmp.BOF And rstmp.EOF Then
			Set rstmp = Nothing
			Exit Sub
		End If
		Do While Not rstmp.EOF
			If rstmp("TemplateID") = CLng(Request("TemplateID")) Then
				sel = " selected"
			Else
				sel = ""
			End If
			If rstmp("pageid") = 0 Then
				strContent = strContent & "<option>��" & rstmp("page_name") & "-�������</option>" & vbCrLf
				strContent = strContent & "<option value='?action=editstyle&TemplateID=" & rstmp("TemplateID") & "&StyleID=0&ChannelID=" & ChannelID & "'" & sel & ">�����͵ײ�ͨ��</option>" & vbCrLf
				strStetting = strStetting & "<option>��" & rstmp("page_name") & "-�������á�</option>" & vbCrLf
				strStetting = strStetting & "<option value='?action=set&TemplateID=" & rstmp("TemplateID") & "&ChannelID=" & ChannelID & "'" & sel & ">��ģ�峣������</option>" & vbCrLf
			Else
				strContent = strContent & "<option value='?action=edit&TemplateID=" & rstmp("TemplateID") & "&ChannelID=" & ChannelID & "'" & sel & ">" & rstmp("page_name") & "-����</option>" & vbCrLf
				strStetting = strStetting & "<option value='?action=set&TemplateID=" & rstmp("TemplateID") & "&ChannelID=" & ChannelID & "'" & sel & ">" & rstmp("page_name") & "-����</option>" & vbCrLf
			End If
			rstmp.MoveNext
		Loop
		rstmp.Close: Set rstmp = Nothing
		Response.Write "ѡ���ҳģ�壺"
		Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
		Response.Write strContent
		Response.Write "</select>"
		Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
		Response.Write strStetting
		Response.Write "</select>"
	End If
End Sub
Sub RemoveCache()
	If Not IsNumeric(Request("TemplateID")) Then
		Exit Sub
	End If
	Dim rsCache
	Set rsCache = enchiasp.Execute("SELECT TemplateID,ChannelID,skinid,pageid FROM ECCMS_Template WHERE TemplateID=" & CLng(Request("TemplateID")))
	enchiasp.DelCahe "MainStyle" & rsCache("skinid")
	enchiasp.DelCahe "Templates" & rsCache("ChannelID") & rsCache("skinid") & rsCache("pageid")
	enchiasp.DelCahe "DefaultSkinID"
	enchiasp.DelCahe "ChannelMenu"
	enchiasp.DelCahe "SiteClassMap"
	rsCache.Close: Set rsCache = Nothing
End Sub
%>