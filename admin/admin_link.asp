<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Dim keyword,readme,Tlink,strurl
Dim totalPut,totalnumber,CurrentPage,maxpagecount,maxperpage
Dim TotalPages,PageName,pagestart,pageend,pubUserName
Dim j, ii, n, face, i
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
If Not ChkAdmin("FriendLink") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
Response.Write " <tr> "
Response.Write " <th height=""22"" colspan=6><a href=""admin_link.asp""><font color=#FFFFFF>����������ҳ</font></a> | <a href=""admin_link.asp?action=add""><font color=#FFFFFF>�����µ���������</font></a></th>"
Response.Write " </tr>"
Response.Write " <tr> "
Response.Write " <td height=""22"" colspan=6 class=TableRow1><form name=""searchsoft"" method=""POST"" action=""admin_link.asp"" target=""main"">"
Response.Write "������������<input class=smallInput type=""text"" name=""keyword"" size=""35""> "
Response.Write "	  ������"
Response.Write "	  <select name=field>"
Response.Write "		<option value=1 selected>��վ����</option>"
Response.Write "		<option value=2>��վ URL</option>"
Response.Write "		<option value=0>��������</option>"
Response.Write "	  </select> "
Response.Write "<input type=""submit"" value=""��������"" name=""submit"" class=""Button"">"
Response.Write " </td></form>"
Response.Write " </tr>"
Response.Write " </table><br>"
If Request("action") = "add" Then
	Call addlink
ElseIf Request("action") = "edit" Then
	Call editlink
ElseIf Request("action") = "savenew" Then
	Call savenew
ElseIf Request("action") = "savedit" Then
	Call savedit
ElseIf Request("action") = "del" Then
	Call del
ElseIf Request("action") = "lock" Then
	Call locklink
ElseIf Request("action") = "free" Then
	Call freelink
Else
	Call linkinfo
End If
If Founderr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Private Sub addlink()
	Response.Write "<form name=myform action=""?action=savenew"" method = post>"
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write " <tr> "
	Response.Write " <th colspan=2>����������� </th>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td width=""30%"" class=TableRow1>��ҳ���� </td>"
	Response.Write " <td width=""70%"" class=TableRow1> "
	Response.Write " <input type=""text"" name=""name"" size=40>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>����URL </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <input type=""text"" name=""url"" value=""http://"" size=60>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>����LOGO��ַ </td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <input type=""text"" name=""logo"" id=""ImageUrl"" value=""http://"" size=60>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>�ϴ�ͼƬ </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <iframe name=image frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?stype=link></iframe>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>��� </td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <textarea name=readme rows=5 cols=50></textarea>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>��������</td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " ��������<input type=""radio"" name=""islogo"" value=0 checked>&nbsp;&nbsp;LOGO����<input type=""radio"" name=""islogo"" value=1>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>�Ƿ�����ҳ��ʾ</td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <input type=""radio"" name=""isIndex"" value=0 checked> ��&nbsp;&nbsp;<input type=""radio"" name=""isIndex"" value=1> ��"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>ǰ̨�޸��������õ����� </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <input type=""text"" name=""password"" value=""" & RndPassWord & """ size=20> "
	Response.Write "<input type=checkbox name=AutoLoad value='yes'> ����Զ��ͼƬ"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td height=""15"" align=center colspan=""2"" class=TableRow1> "
	Response.Write " <input type=""button"" name=""Submit1"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" class=button>��"
	Response.Write " <input type=""submit"" name=""Submit"" class=""button"" value=""�� ��"">"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write "</table>"
	Response.Write "</form>"
End Sub


Private Sub editlink()
	Set Rs = CreateObject("adodb.recordset")
	SQL = "select * from [ECCMS_Link] where linkid=" & Request("id")
	Rs.Open SQL, Conn, 1, 1
	Response.Write "<form name=myform action=""?action=savedit"" method=post>"
	Response.Write "<input type=hidden name=id value="
	Response.Write Request("id")
	Response.Write ">"
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write " <tr> "
	Response.Write " <th colspan=2>�༭��������</th>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td width=""30%"" class=TableRow1>��ҳ���ƣ�</td>"
	Response.Write " <td width=""70%"" class=TableRow1> "
	Response.Write " <input type=""text"" name=""name"" size=40 value="""
	Response.Write Rs("Linkname")
	Response.Write """>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>����URL�� </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <input type=""text"" name=""url"" size=60 value="""
	Response.Write Rs("Linkurl")
	Response.Write """>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>����LOGO��ַ�� </td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <input type=""text"" name=""logo"" id=""ImageUrl"" size=60 value="""
	Response.Write Rs("logourl")
	Response.Write """>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>�ϴ�ͼƬ </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <iframe name=image frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?stype=link></iframe>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>��飺</td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <textarea name=readme rows=5 cols=50>"
	Response.Write Server.HTMLEncode(Rs("readme"))
	Response.Write "</textarea>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>�������� </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " ��������<input type=""radio"" name=""islogo"" value=0"
	If Rs("islogo") = 0 Then
		Response.Write " checked"
	End If
	Response.Write ">&nbsp;&nbsp;LOGO����<input type=""radio"" name=""islogo"" value=1"
	If Rs("islogo") = 1 Then
		Response.Write " checked"
	End If
	Response.Write ">"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>�Ƿ�����ҳ��ʾ </td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <input type=""radio"" name=""isIndex"" value=0"
	If Rs("isIndex") = 0 Then
		Response.Write " checked"
	End If
	Response.Write "> ��&nbsp;&nbsp;<input type=""radio"" name=""isIndex"" value=1"
	If Rs("isIndex") = 1 Then
		Response.Write " checked"
	End If
	Response.Write "> ��"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>������������ </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <input type=""text"" name=""password"" size=20> <font color=blue>���޸�������</font>"
	Response.Write "<input type=checkbox name=AutoLoad value='yes'> ����Զ��ͼƬ"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td height=""15"" align=center colspan=""2"" class=TableRow1> "
	Response.Write " <div align=""center"">"
	Response.Write " <input type=""button"" name=""Submit1"" onclick=""javascript:history.go(-1)"" value=""������һҳ"" class=button>��"
	Response.Write " <input type=""submit"" name=""Submit"" class=""button"" value=""�� ��"">"
	Response.Write " </div>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write "</table>"
	Response.Write "</form>"
	Rs.Close
	Set Rs = Nothing
End Sub


Private Sub linkinfo()
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write " <tr align=center>"
	Response.Write " <th width=""10%"">�� ��</td>"
	Response.Write " <th width=""30%""><B>�� ��</th>"
	Response.Write " <th width=""12%""><B>��������</th>"
	Response.Write " <th width=""30%""><B>�� ��</th>"
	Response.Write " <th width=""10%""><B>״ ̬</th>"
	Response.Write " <th width=""8%""><B>��ҳ</th>"
	Response.Write " </tr>"
	keyword = Trim(Request("keyword"))
	If Not IsEmpty(Request("page")) Then
		CurrentPage = CInt(Request("page"))
	Else
		CurrentPage = 1
	End If
	maxperpage = 15 '###ÿҳ��ʾ��
	PageName = "admin_link.asp"
	Set Rs = Server.CreateObject("adodb.recordset")
	If Not IsNull(keyword) And keyword <> "" Then
		keyword = Replace(Replace(Replace(keyword, "'", "��"), "<", "&lt;"), ">", "&gt;")
		If CInt(Request("field")) = 1 Then
			SQL = "SELECT * FROM [ECCMS_Link] WHERE LinkName LIKE '%" & keyword & "%'"
		ElseIf CInt(Request("field")) = 2 Then
			SQL = "SELECT * FROM [ECCMS_Link] WHERE Linkurl LIKE '%" & keyword & "%'"
		Else
			SQL = "SELECT * FROM [ECCMS_Link] WHERE LinkName LIKE '%" & keyword & "%' Or Linkurl LIKE '%" & keyword & "%'"
		End If
		SQL = SQL & " ORDER BY linkid DESC"
	Else
		SQL = " SELECT * FROM [ECCMS_Link] ORDER BY linkid DESC"
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	Rs.Open SQL, Conn, 1, 1
	If Not (Rs.bof Or Rs.EOF) Then
		Rs.pagesize = maxperpage
		maxpagecount = Rs.pagecount '###��¼��ҳ��
		totalnumber = CInt(Rs.recordcount) '###��¼����
		Rs.absolutepage = CurrentPage '###��ǰҳ��
		ii = 0
		Rem #######��ʾ����ҳ##########
		pagestart = CurrentPage - 3
		pageend = CurrentPage + 3
		Rem ##########################
		n = CurrentPage
		If pagestart < 1 Then
			pagestart = 1
		End If
		If pageend > maxpagecount Then
			pageend = maxpagecount
		End If
		If n < maxpagecount Then
			n = maxpagecount
		End If
		j = (CurrentPage - 1) * maxperpage + 1
		Do While Not Rs.EOF And ii < Rs.pagesize
			Response.Write " <tr align=center>"
			Response.Write " <td height=25 class=TableRow1><font color=red>"
			Response.Write j
			Response.Write "</font></td>"
			Response.Write " <td class=TableRow1><a href="
			Response.Write Rs("Linkurl")
			Response.Write " target=_blank>"
			Response.Write Rs("Linkname")
			Response.Write "</a></td>"
			Response.Write " <td class=TableRow1>"
			If Rs("islogo") = 1 Then
				Response.Write "LOGO����"
			Else
				Response.Write "��������"
			End If
			Response.Write "</td>"
			Response.Write " <td class=TableRow1> <a href=""admin_link.asp?action=edit&id="
			Response.Write Rs("Linkid")
			Response.Write """><u>�༭</u></a> | <a href=""admin_link.asp?action=lock&id="
			Response.Write Rs("linkid")
			Response.Write """><u>����</u></a> | <a href=""admin_link.asp?action=free&id="
			Response.Write Rs("linkid")
			Response.Write """><u>����</u></a> | <a href=""admin_link.asp?action=del&id="
			Response.Write Rs("linkid")
			Response.Write """ onclick=""{if(confirm('�˲�����ɾ������������\n ��ȷ��ִ�д˲�����?')){this.document.myform.submit();return true;}return false;}""><u>ɾ��</u></a></td>"
			Response.Write " <td class=TableRow1>"
			If Rs("isLock") = 0 Then
				Response.Write "����"
			Else
				Response.Write "<font color=red>����</font>"
			End If
			Response.Write "</td>"
			Response.Write " <td class=TableRow1>"
			If Rs("isIndex") = 0 Then
				Response.Write "<font color=red>��</font>"
			Else
				Response.Write "<font color=blue>��</font>"
			End If
			Response.Write "</td>"
			Response.Write " </tr>"
			Rs.movenext
			j = j + 1
			ii = ii + 1
		Loop
		Rs.Close
		Set Rs = Nothing
	Else
		Response.Write ("<tr><td colspan=5 class=TableRow2>��ʱ��û���κ���������</td></tr>")
	End If
	Response.Write "<tr><td colspan=6 class=TableRow2>"
	Call showpage
	Response.Write "</td></tr>"
	Response.Write "</table>"
End Sub

Private Sub savenew()
	Dim sUploadDir,strUploadDir,SaveFileType,SaveFilesName
	Dim password,strLogo
	password = md5(Request("password"))
	strLogo = Trim(Request.Form("logo"))
	If Trim(Request("url")) <> "" And Trim(Request("readme")) <> "" And Trim(Request("name")) <> "" Then
		If Trim(Request("AutoLoad")) = "yes" Then
			sUploadDir = "../link/UploadPic/"
			strUploadDir = CreatePath(sUploadDir)
			SaveFileType = Mid(strLogo, InStrRev(strLogo, ".") + 1)
			SaveFilesName = GetRndFileName(SaveFileType)
			If SaveRemotePic(sUploadDir & strUploadDir & SaveFilesName, strLogo) = True Then
				strLogo = "link/UploadPic/" & strUploadDir & SaveFilesName
			Else
				strLogo = strLogo
			End If
		End If
		Set Rs = CreateObject("adodb.recordset")
		SQL = "select * from [ECCMS_Link] where (Linkid is null)"
		Rs.Open SQL, Conn, 1, 3
		Rs.addnew
			Rs("Linkname").Value =  enchiasp.CheckStr(Request.Form("name"))
			Rs("readme").Value = enchiasp.CheckStr(Request.Form("readme"))
			Rs("logourl").Value = Trim(Request.Form("logo"))
			Rs("Linkurl").Value = Request.Form("url")
			Rs("password").Value = password
			Rs("islogo").Value = Request.Form("islogo")
			Rs("isLock").Value = 0
			Rs("isIndex").Value = Request.Form("isIndex")
			Rs.Update
		Rs.Close
		Set Rs = Nothing
		Succeed("��ӳɹ������������������")
	Else
		ErrMsg = ErrMsg + "<br>" + "��������������������Ϣ��"
		Founderr = True
		Exit Sub
	End If
End Sub

Private Sub savedit()
	Dim sUploadDir,strUploadDir,SaveFileType,SaveFilesName
	Dim strLogo
	strLogo = Trim(Request.Form("logo"))
	If Trim(Request("AutoLoad")) = "yes" Then
		sUploadDir = "../link/UploadPic/"
		strUploadDir = CreatePath(sUploadDir)
		SaveFileType = Mid(strLogo, InStrRev(strLogo, ".") + 1)
		SaveFilesName = GetRndFileName(SaveFileType)
		If SaveRemotePic(sUploadDir & strUploadDir & SaveFilesName, strLogo) = True Then
			strLogo = "link/UploadPic/" & strUploadDir & SaveFilesName
		Else
			strLogo = strLogo
		End If
	End If
	Set Rs = CreateObject("adodb.recordset")
	SQL = "select * from [ECCMS_Link] where Linkid=" & Request("id")
	Rs.Open SQL, Conn, 1, 3
		Rs("Linkname").Value = Trim(Request.Form("name"))
		Rs("readme").Value = Trim(Request.Form("readme"))
		Rs("logourl").Value = strLogo
		Rs("Linkurl").Value = Trim(Request.Form("url"))
		If Trim(Request("password")) <> "" Then Rs("password").Value = Request.Form("password")
		Rs("islogo").Value = Request.Form("islogo")
		Rs("isIndex").Value = Request.Form("isIndex")
		Succeed ("���³ɹ������������������")
		Rs.Update
	Rs.Close
	Set Rs = Nothing
End Sub
Private Sub del()
	Dim id
	id = Request("id")
	SQL = "delete from [ECCMS_Link] where Linkid=" + id
	Conn.Execute (SQL)
	Succeed ("ɾ���ɹ������������������")
End Sub
Private Sub locklink()
	Dim id
	id = Request("id")
	Conn.Execute ("update [ECCMS_Link] set islock=1 where Linkid=" + id)
	Succeed ("���������ɹ������������������")
End Sub
Private Sub freelink()
	Dim id
	id = Request("id")
	Conn.Execute ("update [ECCMS_Link] set islock=0 where Linkid=" + id)
	Succeed ("������������ɹ������������������")
End Sub
Private Function SaveRemotePic(s_LocalFileName, s_RemoteFileUrl)
	Dim Ads
	Dim Retrieval
	Dim GetRemoteData
	Dim bError
	bError = False
	SaveRemotePic = False
	On Error Resume Next
	Set Retrieval = CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", s_RemoteFileUrl, False
		.Send
		If .readyState <> 4 Then Exit Function
		If .Status > 300 Then Exit Function
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	
	Set Ads = CreateObject("Adodb.Stream")
	With Ads
		.type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile Server.MapPath(s_LocalFileName), 2
		.Cancel
		.Close
	End With
	Set Ads = Nothing
	If Err.Number = 0 And bError = False Then
		SaveRemotePic = True
	Else
		Err.Clear
	End If
End Function
Private Function GetRndFileName(ByVal sExt)
	Dim sRnd
	Randomize
	sRnd = Int(900 * Rnd) + 100
	GetRndFileName = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & sRnd & "." & sExt
End Function
Private Sub showpage()
	Response.Write "<table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
	Response.Write " <tr><form method=""POST"" action="""
	Response.Write PageName
	Response.Write """ >"
	Response.Write " <td class=""td1"" align=""center"">����"
	Response.Write totalnumber
	Response.Write "�� <a href="
	Response.Write PageName
	Response.Write "?page=1 title=���ص�һҳ><font face=""Webdings"">97</font></a> "
	For i = pagestart To pageend
		If i = 0 Then
			i = 1
		End If
		strurl = "<a href=" & PageName & "?page=" & i & " title=��" & i & "ҳ>[" & i & "]</a>"
		Response.Write strurl
		Response.Write " "
	Next
	Response.Write "<a href="
	Response.Write PageName
	Response.Write "?page="
	Response.Write maxpagecount
	Response.Write " title=βҳ><font face=""Webdings"">8:</font></a> ҳ��:<font color=red>"
	Response.Write CurrentPage
	Response.Write "</font> / "
	Response.Write maxpagecount
	Response.Write "ҳ ÿҳ:"
	Response.Write maxperpage
	Response.Write " ת��:<select name='page' align=""absmiddle"" size='1' style=""font-size: 9pt"" onChange='javascript:submit()'>"
	Response.Write " "
	For i = 1 To n
		Response.Write " <option value='"
		Response.Write i
		Response.Write "' "
		If CurrentPage = CInt(i) Then
			Response.Write " selected "
		End If
		Response.Write ">��"
		Response.Write i
		Response.Write "ҳ</option>"
		Response.Write " "
	Next
	Response.Write " </select>"
	Response.Write " </td></form>"
	Response.Write " </tr>"
	Response.Write " </table>"
End Sub
Public Function RndPassWord()
	Dim num1,rndnum
	Randomize
	Do While Len(rndnum) < 8
		num1 = CStr(Chr((57 - 48) * rnd + 48))
		rndnum = rndnum & num1
	loop
	RndPassWord = rndnum
End Function
%>
