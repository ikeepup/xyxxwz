<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
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
Dim mdbname, sConn, i
Dim Action
Set Rs = Server.CreateObject("ADODB.Recordset")
Admin_header
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""5"" align=center width=""95%"" class=""tableBorder"">" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<th><a href=? Class=showtitle>ģ�浼������</a> | <a href=?action=load Class=showtitle>ģ�浼�빦��</a></th>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td Class=TableRow1>" & vbCrLf
Response.Write "ע��<br>" & vbCrLf
Response.Write "����ȷ��ģ�����ݿ�����ȷ��<br>" & vbCrLf
Response.Write "������ģ�����ݿ����skinĿ¼�£�����д��"& enchiasp.InstallDir &"skin/ECCMS_Skin.mdb��<br>" & vbCrLf
Response.Write "����ģ�����ݿ��ڱ��ݵı���ΪECCMS_Template,�벻Ҫ���ģ�<br>" & vbCrLf
Response.Write "����ģ�����ݰ����ãӣ����ã��뼰���н������ã�" & vbCrLf
Response.Write "</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "</table><br>" & vbCrLf
If Not ChkAdmin("TemplateLoad") Then
	Server.Transfer ("showerr.asp")
	Response.End
End If
mdbname = enchiasp.CheckStr(Request("mdbname"))
Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "del"
		Call DelTemplate
	Case "input"
		Call InputSkin
	Case "load"
		Call loadTemplate
	Case "loadskin"
		Call LoadSkin
	Case "skin"
		Call SkinsTemplate
	Case "rename"
		Call rename
	Case "savenm"
		Call savenm
	Case Else
		Call PageMain
End Select
If Founderr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
If IsObject(sConn) Then
	sConn.Close
	Set sConn = Nothing
End If

Private Sub PageMain()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>"
	Response.Write "<tr>"
	Response.Write "	<th Colspan=3>ģ�嵼��</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "	<td width=""70%"" Class=TableTitle align=center>ģ������</td>"
	Response.Write "	<td width=""20%"" Class=TableTitle align=center>�� ��</td>"
	Response.Write "	<td width=""10%"" Class=TableTitle align=center>ѡ ��</td>"
	Response.Write "<form name=selform method=post action=?action=input>"
	Response.Write "</tr>"
	SQL = "SELECT TemplateID,skinid,page_name FROM ECCMS_Template WHERE pageid = 0"
	Set Rs = enchiasp.Execute(SQL)
	Do While Not Rs.EOF
		Response.Write "<tr>"
		Response.Write "	<td class=tablerow1>"
		Response.Write Rs("page_name")
		Response.Write "	</td>"
		Response.Write "	<td class=tablerow1 align=center><a href=admin_template.asp?action=manage&skinid="
		Response.Write Rs("skinid")
		Response.Write ">�� ��</a> | "
		Response.Write "<a href=?action=rename&act=loadskin&skid="
		Response.Write Rs("TemplateID")
		Response.Write "&mdbname="
		Response.Write mdbname
		Response.Write ">�� ��</a></td>"
		Response.Write "	<td class=tablerow1 align=center><input type=radio name=skinid value=""" & Rs("skinid") & """></td>"
		Response.Write "</tr>"
		Rs.movenext
	Loop
	Set Rs = Nothing
	Response.Write "<tr>"
	Response.Write "	<td Colspan=3 Class=TableRow1 align=center>ģ�����ݿ�·����<input type=text name=mdbname size=40 value="""
	Response.Write enchiasp.InstallDir
	Response.Write "skin/ECCMS_Skins.Mdb"">"
	Response.Write "	<input type=submit class=Button value=""����ģ��"" onclick=""{if(confirm('��ȷ��Ҫ������ģ����?')){this.document.selform.submit();return true;}return false;}""></td>"
	Response.Write "</tr>"
	Response.Write "</form>"
	Response.Write "</table>"

End Sub

Private Sub LoadTemplate()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>"
	Response.Write "<tr>"
	Response.Write "	<th Colspan=3>ģ�嵼��</th>"
	Response.Write "</tr>"
	Response.Write "<form name=myform method=post action=?action=skin>"
	Response.Write "<tr>"
	Response.Write "	<td Colspan=3 Class=TableRow1 align=center>ģ�����ݿ�·����<input type=text name=mdbname size=40 value="""
	Response.Write enchiasp.InstallDir
	Response.Write "skin/ECCMS_Skins.Mdb"">"
	Response.Write "	<input class=Button type=submit name=act value=""����ģ��""> "
	Response.Write "	<input class=Button type=submit name=act value=""ѹ�����ݿ�""> "
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</form>"
	Response.Write "</table>"
End Sub

Private Sub SkinsTemplate()
	If Trim(Request.Form("mdbname")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������� ��ѡ��ģ�����ݿ⣡</li>"
		Exit Sub
	End If
	If Trim(Request.Form("act")) = "ѹ�����ݿ�" Then
		If CompressMDB(mdbname) Then OutHintScript("��ϲ��ģ�����ݿ�ѹ���ɹ���")
		Exit Sub
	End If
	SkinConnection(mdbname)
	If FoundErr Then Exit Sub
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>"
	Response.Write "<tr>"
	Response.Write "	<th Colspan=3>ģ�嵼��</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "	<td width=""70%"" Class=TableTitle align=center>ģ������</td>"
	Response.Write "	<td width=""20%"" Class=TableTitle align=center>�� ��</td>"
	Response.Write "	<td width=""10%"" Class=TableTitle align=center>ѡ ��</td>"
	Response.Write "<form name=myform method=post action=?action=loadskin>"
	Response.Write "</tr>"
	SQL = "SELECT TemplateID,skinid,page_name FROM ECCMS_Template WHERE pageid = 0"
	Set Rs = sConn.Execute(SQL)
	Do While Not Rs.EOF
		Response.Write "<tr>"
		Response.Write "	<td class=tablerow1>"
		Response.Write Rs("page_name")
		Response.Write "	</td>"
		Response.Write "	<td class=tablerow1 align=center><a href=?action=rename&act=loadskin&skid="
		Response.Write Rs("TemplateID")
		Response.Write "&mdbname="
		Response.Write mdbname
		Response.Write ">�� ��</a> | "
		Response.Write "<a href=?action=del&skinid="
		Response.Write Rs("skinid")
		Response.Write "&mdbname="
		Response.Write mdbname
		Response.Write " onclick=""{if(confirm('ģ��ɾ�����ָܻ�����ȷ��Ҫִ�иò�����?')){return true;}return false;}"">ɾ ��</a></td>"
		Response.Write "	<td class=tablerow1 align=center><input type=radio name=skinid value=""" & Rs("skinid") & """></td>"
		Response.Write "</tr>"
		Rs.movenext
	Loop
	Set Rs = Nothing
	Response.Write "<tr>"
	Response.Write "	<td Colspan=3 align=center>ģ�����ݿ�·����<input type=text name=mdbname size=40 value="""
	Response.Write mdbname
	Response.Write """>"
	Response.Write "	<input class=Button type=submit name=B1 value=""����ģ��""  onclick=""{if(confirm('��ȷ��Ҫ�����ģ����?')){this.document.myform.submit();return true;}return false;}""></td>"
	Response.Write "</tr>"
	Response.Write "</form>"
	Response.Write "</table>"

End Sub

Private Sub SkinConnection(mdbname)
	On Error Resume Next
	Set sConn = CreateObject("ADODB.Connection")
	sConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	If Err.Number = "-2147467259" Then
		ErrMsg = ErrMsg + "<li>" & mdbname & "���ݿⲻ���ڡ�"
		Founderr = True
		Exit Sub
	End If
End Sub

Private Sub LoadSkin()
	If Trim(Request.Form("skinid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������� ��ѡ����Ҫ�����ģ�壡</li>"
		Exit Sub
	End If
	If Trim(Request.Form("mdbname")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������� ��ѡ��ģ�����ݿ⣡</li>"
		Exit Sub
	End If
	SkinConnection(mdbname)
	If FoundErr Then Exit Sub
	Dim SkinRs,newskinid,TemplateDir
	Dim TemplateFields,TemplateValues
	Set SkinRs = Conn.Execute("SELECT Max(skinid) FROM [ECCMS_Template] WHERE pageid = 0")
	If Not (SkinRs.EOF And SkinRs.BOF) Then
		newskinid = SkinRs(0)
	End If
	If IsNull(newskinid) Then newskinid = 0
	SkinRs.Close
	newskinid = newskinid + 1
	SQL = "SELECT * FROM ECCMS_Template WHERE skinid = " & CLng(Request("skinid")) & " ORDER BY ChannelID ASC,TemplateID ASC"
	Set Rs = sConn.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������� û���ҵ���Ҫ������ģ�壡</li>"
		Exit Sub
	End If
	Do While Not Rs.EOF
		If Not IsNull(Rs("TemplateDir")) Then
			TemplateDir = enchiasp.CheckStr(Rs("TemplateDir"))
		Else
			TemplateDir = ""
		End If
		TemplateFields = "ChannelID,skinid,pageid,TemplateDir,page_name,page_content,page_setting,Template_Help,isDefault"
		TemplateValues = "" & Rs("ChannelID") & ","& newskinid &"," & Rs("pageid") & ",'" & TemplateDir & "','" & enchiasp.CheckStr(Rs("page_name")) & "','" & enchiasp.CheckStr(Rs("page_content")) & "','" & enchiasp.CheckStr(Rs("page_setting")) & "','" & enchiasp.CheckStr(Rs("Template_Help")) & "',0"
		SQL = "INSERT INTO [ECCMS_Template](" & TemplateFields & ")VALUES(" & TemplateValues & ")"
		Conn.Execute (SQL)
		Rs.movenext
	Loop
	Set Rs = Nothing
	Response.Write "<script language=JavaScript>" & vbCrLf
	Response.Write "alert('��ϲ�� ģ�嵼��ɹ�����');"
	Response.Write "location.replace('admin_loadskin.asp')" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub

Private Sub InputSkin()
	If Trim(Request.Form("skinid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>����������ѡ����Ҫ������ģ�壡</li>"
		Exit Sub
	End If
	If Trim(Request.Form("mdbname")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������� ��ѡ����Ҫ������ģ�����ݿ⣡</li>"
		Exit Sub
	End If
	SkinConnection(mdbname)
	If FoundErr Then Exit Sub
	Dim SkinRs,newskinid,TemplateDir
	Dim TemplateFields,TemplateValues
	Set SkinRs = sConn.Execute("SELECT MAX(skinid) FROM [ECCMS_Template] WHERE pageid = 0")
	If Not (SkinRs.EOF And SkinRs.BOF) Then
		newskinid = SkinRs(0)
	End If
	If IsNull(newskinid) Then newskinid = 0
	SkinRs.Close
	newskinid = newskinid + 1
	SQL = "SELECT * FROM ECCMS_Template where skinid = " & CLng(Request("skinid"))
	Set Rs = Conn.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������û���ҵ���Ҫ������ģ�壡</li>"
		Exit Sub
	End If
	Do While Not Rs.EOF
		If Not IsNull(Rs("TemplateDir")) Then
			TemplateDir = enchiasp.CheckStr(Rs("TemplateDir"))
		Else
			TemplateDir = ""
		End If
		TemplateFields = "ChannelID,skinid,pageid,TemplateDir,page_name,page_content,page_setting,Template_Help,isDefault"
		TemplateValues = "" & Rs("ChannelID") & ","& newskinid &"," & Rs("pageid") & ",'" & TemplateDir & "','" & enchiasp.CheckStr(Rs("page_name")) & "','" & enchiasp.CheckStr(Rs("page_content")) & "','" & enchiasp.CheckStr(Rs("page_setting")) & "','" & enchiasp.CheckStr(Rs("Template_Help")) & "',0"
		SQL = "INSERT INTO [ECCMS_Template](" & TemplateFields & ")VALUES(" & TemplateValues & ")"
		sConn.Execute (SQL)
		Rs.movenext
	Loop
	Set Rs = Nothing
	Response.Write "<script language=JavaScript>" & vbCrLf
	Response.Write "alert('��ϲ��!ģ�嵼���ɹ���');"
	Response.Write "location.replace('admin_loadskin.asp')" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub

Private Sub rename()
	Dim sRs,skid
	'ģ�����
	skid = enchiasp.checkStr(Request("skid"))
	mdbname = enchiasp.checkStr(Trim(Request("mdbname")))
	If skid <> "" And IsNumeric(skid) Then skid = CLng(skid) Else skid = 1
	If Request("act") = "loadskin" And mdbname <> "" Then
		SkinConnection (mdbname)
		Set sRs = sConn.Execute("SELECT TemplateID,page_name,skinid FROM ECCMS_Template WHERE TemplateID=" & skid)
	Else
		Set sRs = enchiasp.Execute("SELECT TemplateID,page_name,skinid FROM ECCMS_Template WHERE TemplateID=" & skid)
	End If
	Response.Write "<form action=""?action=savenm"" method=post >" & vbCrLf
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""5"" align=center class=""tableBorder"">" & vbCrLf
	Response.Write "<tr><th colspan=""2"">����ģ������ ID="
	Response.Write sRs(2)
	Response.Write "</td></tr>" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write Chr(9) & "<td width=""20%"" class=""TableRow1"">ģ��ԭ����</td>" & vbCrLf
	Response.Write Chr(9) & "<td width=""80%"" class=""TableRow1"">"
	Response.Write sRs(1)
	Response.Write "</td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write Chr(9) & "<td class=""TableRow1"">ģ��������</td>" & vbCrLf
	Response.Write Chr(9) & "<td class=""TableRow1""><input type=""text"" name=""skinNAME"" size=""30"" value=""""></td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "<tr><td align=center class=TableRow2 colspan=""2""><input class=button type=""submit"" name=""submit"" value=""�� ��""></td></tr>" & vbCrLf
	If Request("act") = "loadskin" Then
		Response.Write "<input TYPE=""hidden"" NAME=""mdbname"" VALUE="""
		Response.Write mdbname
		Response.Write """>" & vbCrLf
	End If
	Response.Write "<input TYPE=""hidden"" NAME=""skid"" VALUE="""
	Response.Write sRs(0)
	Response.Write """>" & vbCrLf
	Response.Write "<input TYPE=""hidden"" NAME=""act"" VALUE="""
	Response.Write Request("act")
	Response.Write """>" & vbCrLf
	Response.Write "</table></form>" & vbCrLf
	sRs.Close
	Set sRs = Nothing
End Sub

Private Sub savenm()
	Dim skinNAME,skid
	'ģ���������
	skid = enchiasp.checkStr(Request.Form("skid"))
	mdbname = enchiasp.checkStr(Trim(Request.Form("mdbname")))
	skinNAME = enchiasp.checkStr(Trim(Request.Form("skinname")))
	If skid = "" Or Not IsNumeric(skid) Then
		ErrMsg = ErrMsg + "<BR><li>��ѡ����ȷ�Ĳ���</li>"
		Exit Sub
	End If
	If skinNAME = "" Then
		ErrMsg = ErrMsg + "<li>��ģ�����Ʋ���Ϊ�գ�</li>"
		Exit Sub
	End If
	If Request("act") = "loadskin" And mdbname <> "" Then
		SkinConnection(mdbname)
		sConn.Execute ("UPDATE ECCMS_Template SET page_name='" & skinNAME & "'  WHERE TemplateID=" & skid)
	Else
		enchiasp.Execute ("UPDATE ECCMS_Template SET page_name='" & skinNAME & "'  WHERE TemplateID=" & skid)
	End If
	Succeed ("<li>��ϲ����ģ������ɹ���")
End Sub
Private Sub DelTemplate()
	If Trim(Request("skinid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������^_^ ��ѡ����Ҫɾ����ģ�壡</li>"
		Exit Sub
	End If
	If Trim(Request("mdbname")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�������� ��ѡ����Ҫɾ����ģ�����ݿ⣡</li>"
		Exit Sub
	End If
	SkinConnection(mdbname)
	If FoundErr Then Exit Sub
	sConn.Execute("DELETE FROM ECCMS_Template WHERE skinid = " & CLng(Request("skinid")))
	Response.Write "<script language=JavaScript>" & vbCrLf
	Response.Write "alert('ģ��ɾ���ɹ�����');"
	Response.Write "location.replace('admin_loadskin.asp?action=load')" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub
'================================================
' ��������CompressMDB
' ��  �ã�ѹ��ACCESS���ݿ�
' ��  ����dbPath ----���ݿ�·��
' ����ֵ��True  ----  False
'================================================
Public Function CompressMDB(DBPath)
        Dim fso, Engine, strDBPath
        CompressMDB = False
        If DBPath = "" Then Exit Function
        If InStr(DBPath, ":") = 0 Then DBPath = Server.MapPath(DBPath)
        strDBPath = Left(DBPath, InStrRev(DBPath, "\"))
        Set fso = CreateObject(enchiasp.FSO_ScriptName)

        If fso.FileExists(DBPath) Then
                fso.CopyFile DBPath, strDBPath & "temp.mdb"
                Set Engine = CreateObject("JRO.JetEngine")

                Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
                "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"

                fso.CopyFile strDBPath & "temp1.mdb", DBPath
                fso.DeleteFile (strDBPath & "temp.mdb")
                fso.DeleteFile (strDBPath & "temp1.mdb")
                Set fso = Nothing
                Set Engine = Nothing
                CompressMDB = True
        Else
                CompressMDB = False
        End If
End Function
%>
