<!--#include file="setup.asp" -->
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
Admin_header
Dim Str
If Not ChkAdmin("MailList") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
If Not IsObject(Conn) Then ConnectionDatabase
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=tableBorder>"& vbCrLf
Response.Write "<tr>"& vbCrLf
Response.Write "<th>�ʼ��б�������</th>"& vbCrLf
Response.Write "</tr>"& vbCrLf
Response.Write "<tr>"& vbCrLf
Response.Write "<td class=tablerow2><B>˵��</B>��<BR>1�����������ݿ�ʱ��ȷ��maillist.mdb��databaseĿ¼�У���<BR>2��ʹ�õ������ı��Ĺ�����Ҫ�������˱���֧��FSO������FSO���ѯ΢�����վ��<BR>3�������ʼ��б�ǳ��ķѷ�������Դ���뾡���ڱ��ػ������粻��æ��ʱ��ִ��<br></font></td>"& vbCrLf
Response.Write "</tr>"& vbCrLf
Response.Write "</table>"& vbCrLf
Response.Write "<P></P>"& vbCrLf
Response.Write "<table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3""  align=center class=""tableBorder"">"& vbCrLf
Response.Write "<form name=""maildbout"" method=""post"" action=""admin_mailout.asp?action=maildb"">"& vbCrLf
Response.Write "<tr>"& vbCrLf
Response.Write "<th width=""100%"" colspan=2 align=center>�ʼ��б��������������ݿ�</th>"& vbCrLf
Response.Write "</tr>"& vbCrLf
Response.Write "        <tr>"& vbCrLf
Response.Write "    <td class=tablerow1>�����ʼ��б����ݿ⣺"& vbCrLf
Response.Write "      <input type=""text"" name=""maildb"" value="""& enchiasp.InstallDir &"database/maillist.mdb"" size=35>"& vbCrLf
Response.Write "      <input type=""submit"" name=""Submit"" value=""�����ʼ�"" class=""button"">"& vbCrLf
Response.Write "    </td>"& vbCrLf
Response.Write "  </tr>"& vbCrLf
Response.Write "  </form>"& vbCrLf
Response.Write "</table>"& vbCrLf
Response.Write "<BR>"& vbCrLf
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3""  align=center class=""tableBorder"">"& vbCrLf
Response.Write "<form name=""mailtxtout"" method=""post"" action=""admin_mailout.asp?action=mailtxt"">"& vbCrLf
Response.Write "<tr>"& vbCrLf
Response.Write "<th width=""100%"" colspan=2 align=center>�ʼ��б������������ı���ע�⣺ʹ�øù��ܷ������˱���֧��FSO��</th>"& vbCrLf
Response.Write "</tr>"& vbCrLf
Response.Write "  <tr>"& vbCrLf
Response.Write "    <td class=tablerow1>�����ʼ��б� �� ����"& vbCrLf
Response.Write "      <input type=""text"" name=""mailtxt"" value=""maillist.txt"" size=35>"& vbCrLf
Response.Write "      <input type=""submit"" name=""Submit2"" value=""�����ʼ�"" class=""button"">"& vbCrLf
Response.Write "    </td>"& vbCrLf
Response.Write "  </tr>"& vbCrLf
Response.Write "  </form>"& vbCrLf
Response.Write "</table>"& vbCrLf
Dim temp_count
Set Rs = conn.Execute("select count(*) from [ECCMS_User] where usermail like '%@%'")
temp_count = Rs(0)
Set Rs = server.CreateObject("adodb.recordset")
If temp_count > 0 Then
	sql = "select top "&temp_count&" usermail from [ECCMS_User] where usermail like '%@%'"
	Set Rs = conn.Execute(sql)
End If
Select Case Request("action")
	Case "maildb"
		Call mailoutdb()
	Case "mailtxt"
		Call mailouttxt()
End Select

Sub mailoutdb
        Dim tconn, tconnstr, trs, tsql, tdb, temp_count

        tdb = Request("maildb")
        Set tconn = Server.CreateObject("ADODB.Connection")
        tconnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(tdb)
        tconn.Open tconnstr
		tconn.Execute("delete from [ECCMS_User]")

        Do While Not Rs.EOF
                Set trs = tconn.Execute("insert into [ECCMS_User](usermail) values ('"&rs(0)&"')")
                Rs.movenext
        Loop
        Set trs = tconn.Execute("select count(*) from [ECCMS_User]")
        Response.Write "<table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3""  align=center class=""tableBorder"">"& vbCrLf
        Response.Write "<form name=""maildbout"" method=""post"" action=""admin_mailout.asp?action=maildb"">"& vbCrLf
        Response.Write "<tr>"& vbCrLf
        Response.Write "<th width=""100%"" colspan=2 align=left>"& vbCrLf
        Response.Write "�����ɹ��������� "&trs(0)&" ���û�Email��ַ�����ݿ� "&tdb&" (<a href="&tdb&"><font color=ffffff>����������ػر���</font></a>)"
        Response.Write "</th>"& vbCrLf
        Response.Write "</tr>"& vbCrLf
        Response.Write "</table>"& vbCrLf
        Response.Write "<BR>"& vbCrLf
        Rs.Close
        Set Rs = Nothing
        tConn.Close
        Set tconn = Nothing
End Sub

Sub mailouttxt
        Dim ttxt, File, filepath, writefile

        ttxt = Request("mailtxt")
        Set File = CreateObject("Scripting.FileSystemObject")
        Application.Lock
        filepath = Server.MapPath(""&ttxt&"")
        Set Writefile = File.CreateTextFile(filepath, true)
        Do While Not Rs.EOF
                Writefile.WriteLine Rs(0)
                Rs.movenext
        Loop
        Response.Write "<table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3""  align=center class=""tableBorder"">"& vbCrLf
        Response.Write "<form name=""maildbout"" method=""post"" action=""admin_mailout.asp?action=maildb"">"& vbCrLf
        Response.Write "<tr>"& vbCrLf
        Response.Write "<th width=""100%"" colspan=2 align=left>"& vbCrLf
        Response.Write "�������ı�"&ttxt&"�ɹ���(<a href="&ttxt&" class=TableTitleLink>�������鿴�ʼ��б�</a>)"
        Response.Write "</th>"& vbCrLf
        Response.Write "</tr>"& vbCrLf
        Response.Write "</table>"& vbCrLf
        Response.Write "<BR>"& vbCrLf
        Rs.Close
        Set Rs = Nothing
        Writefile.Close
        Application.unlock
End Sub
Admin_footer
SaveLogInfo(AdminName)
CloseConn
%>
