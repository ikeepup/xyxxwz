<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/cls_public.asp"-->
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
Dim Action,i
If Not ChkAdmin("AdminJsFile" & ChannelID) Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "del"
	Call DeleteJsFile
Case "add"
	Call AddJsFile
Case "edit"
	Call EditJsFile
Case "savenew"
	Call SaveNewJsFile
Case "save"
	Call SaveJsFile
Case "make"
	Call MakeJsFile
Case "demo"
	Call DemoJsFile
Case Else
	Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Sub showmain()
	Dim page_count,tablebody,JsFileName
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th  width=""20%"">˵������</th>"
	Response.Write "		<th  width=""40%"">���÷�ʽ</th>"
	Response.Write "		<th  width=""15%"">JS�ļ�����</th>"
	Response.Write "		<th  width=""25%"">����ѡ��</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=selform method=post action='admin_jsfile.asp'>"
	Response.Write "	<input type=hidden name=action value=""make"">"
	Response.Write "	<input type=hidden name=ChannelID value="""& ChannelID &""">"
	Set Rs = enchiasp.Execute("SELECT id,ChannelID,sTitle,stype,sFileName FROM ECCMS_ScriptFile WHERE ChannelID="& ChannelID &" ORDER BY id DESC")
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=4 class=TableRow1>û�����" & sModuleName & "JS�ļ���</td></tr>"
	Else
		page_count = 0
		Do While Not Rs.EOF
			If (page_count mod 2) = 0 Then
				tablebody = "class=TableRow1"
			Else
				tablebody = "class=TableRow2"
			End If
			JsFileName = "<script src="""& enchiasp.SiteUrl & enchiasp.InstallDir & enchiasp.ChannelDir &"js/"& Rs("sFileName") &"""></script>"
			Response.Write "	<input type=hidden name=id value="""& Rs("id") &""">"
			Response.Write "	<tr align=center>"
			Response.Write "		<td " & tablebody & ">"& Rs("sTitle") &"</td>"
			Response.Write "		<td " & tablebody & "><input type=text name=jsfile size=50 value='" & Server.HTMLEncode(JsFileName) & "'></td>"
			Response.Write "		<td " & tablebody & " noWrap>" & Rs("sFileName") & "</td>"
			Response.Write "		<td " & tablebody & " noWrap><a href='?action=edit&ChannelID="& ChannelID &"&id="& Rs("id") &"'>�� ��</a> | "
			Response.Write "<a href='?action=del&ChannelID="& ChannelID &"&id="& Rs("id") &"' onclick=""return confirm('��ȷ��Ҫɾ����JS�ļ���?')"">ɾ ��</a> | "
			Response.Write "<a href='?action=make&ChannelID="& ChannelID &"&id="& Rs("id") &"'>�� ��</a> | "
			Response.Write "<a href='?action=demo&ChannelID="& ChannelID &"&id="& Rs("id") &"'>�� ʾ</a>"
			Response.Write "</td>"
			Response.Write "	</tr>"
			Rs.movenext
			page_count = page_count + 1
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 colspan=4 align=center>"
	Response.Write "<input type=submit name=submit3 value="" ��������JS�ļ� "" class=Button>&nbsp;&nbsp;"
	Response.Write "<input type=submit name=submit4 value="" ����µ�JS�ļ� "" onclick=""document.selform.action.value='add';"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=4>"
	Response.Write "<b>˵����</b><br>&nbsp;&nbsp;&nbsp;&nbsp;�뽫�����JS���ô��븴�Ƶ�ģ����Ӧ��λ�ã�"
	Response.Write "����JS�ļ���ϵͳ���ɵľ�̬�ļ�������Ҫ�����ڵ���������JS�ļ���"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
End Sub

Sub AddJsFile()
	Response.Write "<table cellspacing=1 align=center cellpadding=0 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=4>���" & sModuleName & "JS�ļ�</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action='admin_jsfile.asp'>"
	Response.Write "	<input type=hidden name=action value=""savenew"">"
	Response.Write "	<input type=hidden name=ChannelID value="""& ChannelID &""">"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right width=""20%""><b>�ļ�˵����</b></td>"
	Response.Write "		<td class=tablerow1 width=""30%""><input type=text name=sTitle size=35 value=''></td>"
	Response.Write "		<td class=tablerow1 align=right width=""20%""><b>JS�ļ����ƣ�</b></td>"
	Response.Write "		<td class=tablerow1 width=""30%""><input type=text name=sFileName size=20 value='file.js'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>��ʾ���ͣ�</b></td>"
	Response.Write "		<td class=tablerow2><input type=radio name=stype value=0 checked onClick=""stype1.style.display='';stype2.style.display='none';""> �б�&nbsp;&nbsp;"
	Response.Write "<input type=radio name=stype value=1 onClick=""stype2.style.display='';stype1.style.display='none';""> ͼƬ</td>"
	Response.Write "		<td class=tablerow2 align=right><b>ѡ����ࣺ</b></td>"
	Response.Write "		<td class=tablerow2>"
	Response.Write "<select name=""setting"" size='1'>"
	Response.Write "<option value=0>ָ�����з���</option>"
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
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>����ר�⣺</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name=""setting"" size='1'>"
	Response.Write "<option value=0>��ָ��ר��</option>"
	Dim RsObj
	Set RsObj = enchiasp.Execute("SELECT SpecialID,SpecialName FROM ECCMS_Special WHERE ChannelID="& ChannelID &" And ChangeLink=0")
	Do While Not RsObj.EOF
		Response.Write "<option value='" & RsObj("SpecialID") & "'>" & RsObj("SpecialName") & "</option>"
		RsObj.movenext
	Loop
	Set RsObj = Nothing
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "		<td class=tablerow1 align=right><b>�������ͣ�</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name=""setting"" size='1'>"
        Response.Write "  <option value=0>��������" & sModuleName & "</option>"
        Response.Write "  <option value='1'>�����Ƽ�" & sModuleName & "</option>"
	Response.Write "  <option value='2'>��������" & sModuleName & "</option>"
	Response.Write "  <option value='3'>��������" & sModuleName & "</option>"
	Response.Write "  <option value='4'>�����Ƽ�" & sModuleName & "</option>"
	Response.Write "  <option value='5'>��������" & sModuleName & "</option>"
        Response.Write "</select><font color=""#0066CC""></font>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=stype1>"
	Response.Write "		<td  bgcolor=""#FFFFFF"" colspan=4>"
	Response.Write "<table width=""100%"" cellspacing=1 align=center cellpadding=3 border=0>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>ͼƬ��ȣ�</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='85'></td>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>ͼƬ�߶ȣ�</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='85'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>����ַ�����</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='22'></td>"
	Response.Write "		<td class=tablerow1 align=right><b>����б�����</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='10'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>��ֱ�߾ࣺ</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='10'></td>"
	Response.Write "		<td class=tablerow2 align=right><b>ˮƽ�߾ࣺ</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='10'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>ͼ�Ķ��뷽ʽ��</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='left' selected>�����</option>"
	Response.Write "		<option value='right'>�Ҷ���</option>"
	Response.Write "		<option value='middle'>���ж���</option>"
	Response.Write "		<option value='texttop'>�ı��Ϸ�</option>"
	Response.Write "		<option value='baseline'>����</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b>������ʽ��</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=20 value='class=dottedline'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>����Ŀ�꣺</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	Response.Write "		<option value='_blank' selected>_blank</option>"
	Response.Write "		<option value='_self'>_self</option>"
	Response.Write "		<option value='_top'>_top</option>"
	Response.Write "		<option value='_parent'>_parent</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow2 align=right><b>��ʶ����</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=20 value='�� '></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>�Ƿ���ʾͼƬ��</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0' selected>����ʾ</option>"
	Response.Write "		<option value='1'>��ʾ</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b>�Ƿ���ʾ���ࣺ</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0' selected>����ʾ</option>"
	Response.Write "		<option value='1'>��ʾ</option>"
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>�Ƿ���ʾʱ�䣺</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	Response.Write "		<option value='0' selected>����ʾ</option>"
	Response.Write "		<option value='1'>��ʾ</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow2 align=right><b>��ʾʱ���ʽ��</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	For i = 1 To 9
		Response.Write "<option value='" & i & "'"
		If i = 5 Then Response.Write " selected"
		Response.Write ">"
		Response.Write enchiasp.FormatDate(Now(),i)
		Response.Write "</option>" & vbCrLf
	Next
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=stype2 style=""display:none"">"
	Response.Write "		<td  bgcolor=""#FFFFFF"" colspan=4>"
	Response.Write "<table width=""100%"" cellspacing=1 align=center cellpadding=3 border=0>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>�����ʾ����ͼƬ��</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='5'></td>"
	Response.Write "		<td class=tablerow2 align=right><b>ÿ����ʾ����ͼƬ��</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='5'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>��ʾ����ַ�����</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='22'></td>"
	Response.Write "		<td class=tablerow1 align=right><b>�Ƿ��´��ڴ򿪣�</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0' selected>�����ڴ�</option>"
	Response.Write "		<option value='1'>�´��ڴ�</option>"
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>ͼƬ��ȣ�</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='120'></td>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>ͼƬ�߶ȣ�</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='100'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>�Ƿ���ʾ�������ƣ�</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0' selected>����ʾ</option>"
	Response.Write "		<option value='1'>��ʾ</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b></b></td>"
	Response.Write "		<td class=tablerow1></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 colspan=4 align=center>"
	Response.Write "	<input type=button name=Submit4 onclick=""javascript:history.go(-1)"" value=""������һҳ"" class=Button>&nbsp;&nbsp;"
	Response.Write "	<input type=submit name=submit3 value="" ����µ�JS�ļ� "" onclick=""document.selform.action.value='add';"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub

Sub EditJsFile()
	On Error Resume Next
	Dim JsSetting
	Set Rs = enchiasp.Execute("SELECT id,sTitle,stype,sFileName,setting FROM ECCMS_ScriptFile WHERE ChannelID = "& ChannelID &" And id ="& CLng(Request("id")))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ������</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	JsSetting = Split(Rs("setting"), ",")
	Response.Write "<table cellspacing=1 align=center cellpadding=0 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=4>���" & sModuleName & "JS�ļ�</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action='admin_jsfile.asp'>"
	Response.Write "	<input type=hidden name=action value=""save"">"
	Response.Write "	<input type=hidden name=ChannelID value="""& ChannelID &""">"
	Response.Write "	<input type=hidden name=id value="""& Rs("id") &""">"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right width=""20%""><b>�ļ�˵����</b></td>"
	Response.Write "		<td class=tablerow1 width=""30%""><input type=text name=sTitle size=35 value='"& Rs("sTitle") &"'></td>"
	Response.Write "		<td class=tablerow1 align=right width=""20%""><b>JS�ļ����ƣ�</b></td>"
	Response.Write "		<td class=tablerow1 width=""30%""><input type=text name=sFileName size=20 value='"& Rs("sFileName") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>��ʾ���ͣ�</b></td>"
	Response.Write "		<td class=tablerow2><input type=radio name=stype value=0 onClick=""stype1.style.display='';stype2.style.display='none';"""
	If Rs("stype") = 0 Then Response.Write " checked"
	Response.Write "> �б�&nbsp;&nbsp;"
	Response.Write "<input type=radio name=stype value=1 onClick=""stype2.style.display='';stype1.style.display='none';"""
	If Rs("stype") = 1 Then Response.Write " checked"
	Response.Write "> ͼƬ</td>"
	Response.Write "		<td class=tablerow2 align=right><b>ѡ����ࣺ</b></td>"
	Response.Write "		<td class=tablerow2>"
	Response.Write "<select name=""setting"" size='1'>"
	Response.Write "<option value=0>ָ�����з���</option>"
	Dim strSelectClass,re
	strSelectClass = enchiasp.LoadSelectClass(ChannelID)
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	Re.Pattern="(value=)(.*)("" )"
	strSelectClass = Re.Replace(strSelectClass,"")
	Re.Pattern="({ClassID=)(.*)(}>)"
	strSelectClass = Re.Replace(strSelectClass,"value=""$2"">")
	Set Re = Nothing
	strSelectClass = Replace(strSelectClass, "value=""" & Trim(JsSetting(0)) & """", "value=""" & Trim(JsSetting(0)) & """ selected")
	Response.Write strSelectClass
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>����ר�⣺</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name=""setting"" size='1'>"
	Response.Write "<option value=0>��ָ��ר��</option>"
	Dim RsObj
	Set RsObj = enchiasp.Execute("SELECT SpecialID,SpecialName FROM ECCMS_Special WHERE ChannelID="& ChannelID &" And ChangeLink=0")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("SpecialID") & """"
		If CLng(JsSetting(1)) = RsObj("SpecialID") Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("SpecialName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "		<td class=tablerow1 align=right><b>�������ͣ�</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name=""setting"" size='1'>"
        Response.Write "  <option value='0'"
	If CLng(JsSetting(2)) = 0 Then Response.Write " selected"
	Response.Write ">��������" & sModuleName & "</option>"
        Response.Write "  <option value='1'"
	If CLng(JsSetting(2)) = 1 Then Response.Write " selected"
	Response.Write ">�����Ƽ�" & sModuleName & "</option>"
	Response.Write "  <option value='2'"
	If CLng(JsSetting(2)) = 2 Then Response.Write " selected"
	Response.Write ">��������" & sModuleName & "</option>"
	Response.Write "  <option value='3'"
	If CLng(JsSetting(2)) = 3 Then Response.Write " selected"
	Response.Write ">��������" & sModuleName & "</option>"
	Response.Write "  <option value='4'"
	If CLng(JsSetting(2)) = 4 Then Response.Write " selected"
	Response.Write ">�����Ƽ�" & sModuleName & "</option>"
	Response.Write "  <option value='5'"
	If CLng(JsSetting(2)) = 5 Then Response.Write " selected"
	Response.Write ">��������" & sModuleName & "</option>"
        Response.Write "</select><font color=""#0066CC""></font>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=stype1"
	If Rs("stype") = 1 Then Response.Write " style=""display:none"""
	Response.Write ">"
	Response.Write "		<td  bgcolor=""#FFFFFF"" colspan=4>"
	Response.Write "<table width=""100%"" cellspacing=1 align=center cellpadding=3 border=0>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>ͼƬ��ȣ�</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='" & Trim(JsSetting(3)) & "'></td>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>ͼƬ�߶ȣ�</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='" & Trim(JsSetting(4)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>����ַ�����</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='" & Trim(JsSetting(5)) & "'></td>"
	Response.Write "		<td class=tablerow1 align=right><b>����б�����</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='" & Trim(JsSetting(6)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>��ֱ�߾ࣺ</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='" & Trim(JsSetting(7)) & "'></td>"
	Response.Write "		<td class=tablerow2 align=right><b>ˮƽ�߾ࣺ</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='" & Trim(JsSetting(8)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>ͼ�Ķ��뷽ʽ��</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='left'"
	If LCase(Trim(JsSetting(9))) = "left" Then Response.Write " selected"
	Response.Write ">�����</option>"
	Response.Write "		<option value='right'"
	If LCase(Trim(JsSetting(9))) = "right" Then Response.Write " selected"
	Response.Write ">�Ҷ���</option>"
	Response.Write "		<option value='middle'"
	If LCase(Trim(JsSetting(9))) = "middle" Then Response.Write " selected"
	Response.Write ">���ж���</option>"
	Response.Write "		<option value='texttop'"
	If LCase(Trim(JsSetting(9))) = "texttop" Then Response.Write " selected"
	Response.Write ">�ı��Ϸ�</option>"
	Response.Write "		<option value='baseline'"
	If LCase(Trim(JsSetting(9))) = "baseline" Then Response.Write " selected"
	Response.Write ">����</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b>������ʽ��</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=20 value='" & Trim(JsSetting(10)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>����Ŀ�꣺</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	Response.Write "		<option value='" & Trim(JsSetting(11)) & "'>" & Trim(JsSetting(11)) & "</option>"
	Response.Write "		<option value='_blank'>_blank</option>"
	Response.Write "		<option value='_self'>_self</option>"
	Response.Write "		<option value='_top'>_top</option>"
	Response.Write "		<option value='_parent'>_parent</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow2 align=right><b>��ʶ����</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=20 value='" & JsSetting(12) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>�Ƿ���ʾͼƬ��</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0'"
	If CInt(JsSetting(13)) = 0 Then Response.Write " selected"
	Response.Write ">����ʾ</option>"
	Response.Write "		<option value='1'"
	If CInt(JsSetting(13)) = 1 Then Response.Write " selected"
	Response.Write ">��ʾ</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b>�Ƿ���ʾ���ࣺ</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0'"
	If CInt(JsSetting(14)) = 0 Then Response.Write " selected"
	Response.Write ">����ʾ</option>"
	Response.Write "		<option value='1'"
	If CInt(JsSetting(14)) = 1 Then Response.Write " selected"
	Response.Write ">��ʾ</option>"
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>�Ƿ���ʾʱ�䣺</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	Response.Write "		<option value='0'"
	If CInt(JsSetting(15)) = 0 Then Response.Write " selected"
	Response.Write ">����ʾ</option>"
	Response.Write "		<option value='1'"
	If CInt(JsSetting(15)) = 1 Then Response.Write " selected"
	Response.Write ">��ʾ</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow2 align=right><b>��ʾʱ���ʽ��</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	For i = 1 To 9
		Response.Write "<option value='" & i & "'"
		If CLng(JsSetting(16)) = i Then Response.Write " selected"
		Response.Write ">"
		Response.Write enchiasp.FormatDate(Now(),i)
		Response.Write "</option>" & vbCrLf
	Next
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=stype2"
	If Rs("stype") = 0 Then Response.Write " style=""display:none"""
	Response.Write ">"
	Response.Write "		<td  bgcolor=""#FFFFFF"" colspan=4>"
	Response.Write "<table width=""100%"" cellspacing=1 align=center cellpadding=3 border=0>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>�����ʾ����ͼƬ��</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='" & Trim(JsSetting(17)) & "'></td>"
	Response.Write "		<td class=tablerow2 align=right><b>ÿ����ʾ����ͼƬ��</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='" & Trim(JsSetting(18)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>��ʾ����ַ�����</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='" & Trim(JsSetting(19)) & "'></td>"
	Response.Write "		<td class=tablerow1 align=right><b>�Ƿ��´��ڴ򿪣�</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0'"
	If CInt(JsSetting(20)) = 0 Then Response.Write " selected"
	Response.Write ">�����ڴ�</option>"
	Response.Write "		<option value='1'"
	If CInt(JsSetting(20)) = 1 Then Response.Write " selected"
	Response.Write ">�´��ڴ�</option>"
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>ͼƬ��ȣ�</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='" & Trim(JsSetting(21)) & "'></td>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>ͼƬ�߶ȣ�</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='" & Trim(JsSetting(22)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>�Ƿ���ʾ�������ƣ�</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0'"
	If CInt(JsSetting(23)) = 0 Then Response.Write " selected"
	Response.Write ">����ʾ</option>"
	Response.Write "		<option value='1'"
	If CInt(JsSetting(23)) = 1 Then Response.Write " selected"
	Response.Write ">��ʾ</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b></b></td>"
	Response.Write "		<td class=tablerow1></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 colspan=4 align=center>"
	Response.Write "	<input type=button name=Submit4 onclick=""javascript:history.go(-1)"" value=""������һҳ"" class=Button>&nbsp;&nbsp;"
	Response.Write "	<input type=submit name=submit3 value="" ��������JS�ļ� "" onclick=""document.selform.action.value='add';"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
	Set Rs = Nothing
End Sub

Sub SaveNewJsFile()
	If Trim(Request.Form("sTitle")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>JS�ļ�˵������Ϊ�գ�</li>"
		Exit Sub
	End If
	If Trim(Request.Form("sFileName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>JS�ļ�������Ϊ�գ�</li>"
		Exit Sub
	End If
	If LCase(Right(Trim(Request.Form("sFileName")),3)) <> ".js" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������ȷ��JS�ļ�������չ��һ��Ҫ��*.js��</li>"
		Exit Sub
	End If
	SQL = "INSERT INTO ECCMS_ScriptFile (ChannelID,sTitle,stype,sFileName,setting) VALUES ("& ChannelID &",'"& enchiasp.CheckStr(Request("sTitle")) &"',"& Request("stype") &",'"& enchiasp.CheckStr(Request("sFileName")) &"','"& Request("setting") &"')"
	enchiasp.Execute(SQL)
	Response.Redirect("admin_jsfile.asp?ChannelID="& ChannelID)
End Sub

Sub SaveJsFile()
	If Trim(Request.Form("sTitle")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>JS�ļ�˵������Ϊ�գ�</li>"
		Exit Sub
	End If
	If Trim(Request.Form("sFileName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>JS�ļ�������Ϊ�գ�</li>"
		Exit Sub
	End If
	If LCase(Right(Trim(Request.Form("sFileName")),3)) <> ".js" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������ȷ��JS�ļ�������չ��һ��Ҫ��*.js��</li>"
		Exit Sub
	End If
	SQL = "UPDATE ECCMS_ScriptFile SET sTitle='"& enchiasp.CheckStr(Request("sTitle")) &"',stype="& Request("stype") &",sFileName='"& enchiasp.CheckStr(Request("sFileName")) &"',setting='"& Request("setting") &"' WHERE ChannelID = "& ChannelID &" And id="& Request("id")
	enchiasp.Execute(SQL)
	Response.Redirect("admin_jsfile.asp?ChannelID="& ChannelID)
End Sub

Sub MakeJsFile()
	If Trim(Request("id")) = "" Then
		ErrMsg = "<li>�����ϵͳ����,��ѡ���ļ�ID</li>"
		FoundErr = True
		Exit Sub
	End If
	Dim FileName,strJsContent,JsSetting
	
	On Error Resume Next
	SQL = "SELECT stype,sFileName,setting FROM ECCMS_ScriptFile WHERE ChannelID="& ChannelID &" And id in("& Request("id") &")"
	Set Rs = enchiasp.Execute(SQL)
	If Not(Rs.BOF And Rs.EOF) Then
		Do While Not Rs.EOF
			JsSetting = Split(Rs("setting"), ",")
			enchiasp.InstallDir = enchiasp.SiteUrl & "/"
			If Rs("stype") = 1 Then
				Select Case CInt(enchiasp.modules)
				Case 1
					strJsContent = HTML.LoadArticlePic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				Case 2
					strJsContent = HTML.LoadSoftPic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				Case 3
					strJsContent = HTML.LoadShopPic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				Case 5
					strJsContent = HTML.LoadFlashPic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				End Select
			Else
				Select Case CInt(enchiasp.modules)
				Case 1
					strJsContent = HTML.NewsPictureAndText(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(3)),Trim(JsSetting(4)),Trim(JsSetting(5)),Trim(JsSetting(6)),Trim(JsSetting(7)),Trim(JsSetting(8)),Trim(JsSetting(9)),Trim(JsSetting(10)),Trim(JsSetting(11)),JsSetting(12),Trim(JsSetting(13)),Trim(JsSetting(14)),Trim(JsSetting(15)),Trim(JsSetting(16)))
				Case 2
					strJsContent = HTML.SoftPictureAndText(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(3)),Trim(JsSetting(4)),Trim(JsSetting(5)),Trim(JsSetting(6)),Trim(JsSetting(7)),Trim(JsSetting(8)),Trim(JsSetting(9)),Trim(JsSetting(10)),Trim(JsSetting(11)),JsSetting(12),Trim(JsSetting(13)),Trim(JsSetting(14)),Trim(JsSetting(15)),Trim(JsSetting(16)))
				Case 3
					strJsContent = HTML.LoadShopPic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				Case 5
					strJsContent = HTML.LoadFlashPic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				End Select
			End If
			
			strJsContent = "document.write ('"& fixjs(strJsContent) &"');"
			FileName = "../"& enchiasp.ChannelDir &"js/"& Rs("sFileName")
			enchiasp.CreatedTextFile FileName,strJsContent
			Rs.movenext
		Loop
	End If
	Set Rs = Nothing
	Succeed("<li>��ϲ��������JS�ļ��ɹ���</li>")
End Sub

Sub DeleteJsFile()
	If Trim(Request("id")) <> "" Then
		On Error Resume Next
		Set Rs = enchiasp.Execute("SELECT sFileName FROM ECCMS_ScriptFile WHERE ChannelID = "& ChannelID &" And id=" & CLng(Request("id")))
		If Not(Rs.BOF And Rs.EOF) Then
			enchiasp.FileDelete("../"& enchiasp.ChannelDir &"js/"& Rs("sFileName"))
		End If
		Set Rs = Nothing
		enchiasp.Execute ("DELETE FROM ECCMS_ScriptFile WHERE ChannelID = "& ChannelID &" And id=" & CLng(Request("id")))
		Response.Redirect  Request.ServerVariables("HTTP_REFERER")
	Else
		ErrMsg = "<li>�����ϵͳ����,��ѡ��Ҫɾ�����ļ�ID</li>"
		FoundErr = True
		Exit Sub
	End If
End Sub
Sub DemoJsFile()
	Dim JsFileName
	Set Rs = enchiasp.Execute("SELECT sFileName FROM ECCMS_ScriptFile WHERE ChannelID = "& ChannelID &" And id=" & CLng(Request("id")))
	If Not(Rs.BOF And Rs.EOF) Then
		Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
		Response.Write "	<tr>"
		Response.Write "		<th>" & sModuleName & "JS�ļ�������ʾ</th>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 colspan=4 align=center>"
		JsFileName = "<script src="""& enchiasp.InstallDir & enchiasp.ChannelDir &"js/"& Rs("sFileName") &""" type=""text/javascript""></script>"
		Response.Write Server.HTMLEncode(JsFileName)
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1>"
		Response.Write JsFileName
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=center>"
		Response.Write "	<input type=button name=Submit4 onclick=""javascript:history.go(-1)"" value=""������һҳ"" class=Button>&nbsp;&nbsp;"
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "</table>"
	End If
End Sub

%>