<!--#include file="setup.asp" -->
<!--#include file="check.asp"-->
<%
'=====================================================================
' 软件名称：恩池网站管理系统
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
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
Response.Write "<th>邮件列表导出管理</th>"& vbCrLf
Response.Write "</tr>"& vbCrLf
Response.Write "<tr>"& vbCrLf
Response.Write "<td class=tablerow2><B>说明</B>：<BR>1、导出到数据库时请确认maillist.mdb在database目录中）。<BR>2、使用导出到文本的功能需要服务器端必须支持FSO，关于FSO请查询微软的网站或<BR>3、导出邮件列表非常耗费服务器资源，请尽量在本地或在网络不繁忙的时候执行<br></font></td>"& vbCrLf
Response.Write "</tr>"& vbCrLf
Response.Write "</table>"& vbCrLf
Response.Write "<P></P>"& vbCrLf
Response.Write "<table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3""  align=center class=""tableBorder"">"& vbCrLf
Response.Write "<form name=""maildbout"" method=""post"" action=""admin_mailout.asp?action=maildb"">"& vbCrLf
Response.Write "<tr>"& vbCrLf
Response.Write "<th width=""100%"" colspan=2 align=center>邮件列表批量导出到数据库</th>"& vbCrLf
Response.Write "</tr>"& vbCrLf
Response.Write "        <tr>"& vbCrLf
Response.Write "    <td class=tablerow1>导出邮件列表到数据库："& vbCrLf
Response.Write "      <input type=""text"" name=""maildb"" value="""& enchiasp.InstallDir &"database/maillist.mdb"" size=35>"& vbCrLf
Response.Write "      <input type=""submit"" name=""Submit"" value=""导出邮件"" class=""button"">"& vbCrLf
Response.Write "    </td>"& vbCrLf
Response.Write "  </tr>"& vbCrLf
Response.Write "  </form>"& vbCrLf
Response.Write "</table>"& vbCrLf
Response.Write "<BR>"& vbCrLf
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3""  align=center class=""tableBorder"">"& vbCrLf
Response.Write "<form name=""mailtxtout"" method=""post"" action=""admin_mailout.asp?action=mailtxt"">"& vbCrLf
Response.Write "<tr>"& vbCrLf
Response.Write "<th width=""100%"" colspan=2 align=center>邮件列表批量导出到文本（注意：使用该功能服务器端必须支持FSO）</th>"& vbCrLf
Response.Write "</tr>"& vbCrLf
Response.Write "  <tr>"& vbCrLf
Response.Write "    <td class=tablerow1>导出邮件列表到 文 本："& vbCrLf
Response.Write "      <input type=""text"" name=""mailtxt"" value=""maillist.txt"" size=35>"& vbCrLf
Response.Write "      <input type=""submit"" name=""Submit2"" value=""导出邮件"" class=""button"">"& vbCrLf
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
        Response.Write "操作成功，共导出 "&trs(0)&" 个用户Email地址到数据库 "&tdb&" (<a href="&tdb&"><font color=ffffff>点击这里下载回本地</font></a>)"
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
        Response.Write "导出到文本"&ttxt&"成功，(<a href="&ttxt&" class=TableTitleLink>点击这里查看邮件列表</a>)"
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
