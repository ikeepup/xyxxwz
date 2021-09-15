<!--#include file="setup.asp"-->
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
Dim mdbname, sConn, i
Dim Action
Set Rs = Server.CreateObject("ADODB.Recordset")
Admin_header
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""5"" align=center width=""95%"" class=""tableBorder"">" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<th><a href=? Class=showtitle>模版导出功能</a> | <a href=?action=load Class=showtitle>模版导入功能</a></th>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td Class=TableRow1>" & vbCrLf
Response.Write "注意<br>" & vbCrLf
Response.Write "１，确认模版数据库名正确；<br>" & vbCrLf
Response.Write "２，如模版数据库放在skin目录下，即填写："& enchiasp.InstallDir &"skin/ECCMS_Skin.mdb；<br>" & vbCrLf
Response.Write "３，模版数据库内备份的表名为ECCMS_Template,请不要更改；<br>" & vbCrLf
Response.Write "４，模版数据包括ＣＳＳ设置，与及所有界面设置．" & vbCrLf
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
	Response.Write "	<th Colspan=3>模板导出</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "	<td width=""70%"" Class=TableTitle align=center>模板名称</td>"
	Response.Write "	<td width=""20%"" Class=TableTitle align=center>操 作</td>"
	Response.Write "	<td width=""10%"" Class=TableTitle align=center>选 择</td>"
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
		Response.Write ">编 辑</a> | "
		Response.Write "<a href=?action=rename&act=loadskin&skid="
		Response.Write Rs("TemplateID")
		Response.Write "&mdbname="
		Response.Write mdbname
		Response.Write ">改 名</a></td>"
		Response.Write "	<td class=tablerow1 align=center><input type=radio name=skinid value=""" & Rs("skinid") & """></td>"
		Response.Write "</tr>"
		Rs.movenext
	Loop
	Set Rs = Nothing
	Response.Write "<tr>"
	Response.Write "	<td Colspan=3 Class=TableRow1 align=center>模板数据库路径：<input type=text name=mdbname size=40 value="""
	Response.Write enchiasp.InstallDir
	Response.Write "skin/ECCMS_Skins.Mdb"">"
	Response.Write "	<input type=submit class=Button value=""导出模板"" onclick=""{if(confirm('您确定要导出该模板吗?')){this.document.selform.submit();return true;}return false;}""></td>"
	Response.Write "</tr>"
	Response.Write "</form>"
	Response.Write "</table>"

End Sub

Private Sub LoadTemplate()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>"
	Response.Write "<tr>"
	Response.Write "	<th Colspan=3>模板导入</th>"
	Response.Write "</tr>"
	Response.Write "<form name=myform method=post action=?action=skin>"
	Response.Write "<tr>"
	Response.Write "	<td Colspan=3 Class=TableRow1 align=center>模板数据库路径：<input type=text name=mdbname size=40 value="""
	Response.Write enchiasp.InstallDir
	Response.Write "skin/ECCMS_Skins.Mdb"">"
	Response.Write "	<input class=Button type=submit name=act value=""导入模板""> "
	Response.Write "	<input class=Button type=submit name=act value=""压缩数据库""> "
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</form>"
	Response.Write "</table>"
End Sub

Private Sub SkinsTemplate()
	If Trim(Request.Form("mdbname")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>出错啦！ 请选择模板数据库！</li>"
		Exit Sub
	End If
	If Trim(Request.Form("act")) = "压缩数据库" Then
		If CompressMDB(mdbname) Then OutHintScript("恭喜您模板数据库压缩成功！")
		Exit Sub
	End If
	SkinConnection(mdbname)
	If FoundErr Then Exit Sub
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>"
	Response.Write "<tr>"
	Response.Write "	<th Colspan=3>模板导入</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "	<td width=""70%"" Class=TableTitle align=center>模板名称</td>"
	Response.Write "	<td width=""20%"" Class=TableTitle align=center>操 作</td>"
	Response.Write "	<td width=""10%"" Class=TableTitle align=center>选 择</td>"
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
		Response.Write ">改 名</a> | "
		Response.Write "<a href=?action=del&skinid="
		Response.Write Rs("skinid")
		Response.Write "&mdbname="
		Response.Write mdbname
		Response.Write " onclick=""{if(confirm('模板删除后不能恢复，您确定要执行该操作吗?')){return true;}return false;}"">删 除</a></td>"
		Response.Write "	<td class=tablerow1 align=center><input type=radio name=skinid value=""" & Rs("skinid") & """></td>"
		Response.Write "</tr>"
		Rs.movenext
	Loop
	Set Rs = Nothing
	Response.Write "<tr>"
	Response.Write "	<td Colspan=3 align=center>模板数据库路径：<input type=text name=mdbname size=40 value="""
	Response.Write mdbname
	Response.Write """>"
	Response.Write "	<input class=Button type=submit name=B1 value=""导入模板""  onclick=""{if(confirm('您确定要导入该模板吗?')){this.document.myform.submit();return true;}return false;}""></td>"
	Response.Write "</tr>"
	Response.Write "</form>"
	Response.Write "</table>"

End Sub

Private Sub SkinConnection(mdbname)
	On Error Resume Next
	Set sConn = CreateObject("ADODB.Connection")
	sConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	If Err.Number = "-2147467259" Then
		ErrMsg = ErrMsg + "<li>" & mdbname & "数据库不存在。"
		Founderr = True
		Exit Sub
	End If
End Sub

Private Sub LoadSkin()
	If Trim(Request.Form("skinid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>出错啦！ 请选择你要导入的模板！</li>"
		Exit Sub
	End If
	If Trim(Request.Form("mdbname")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>出错啦！ 请选择模板数据库！</li>"
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
		ErrMsg = ErrMsg + "<li>出错啦！ 没有找到你要导出的模板！</li>"
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
	Response.Write "alert('恭喜您 模板导入成功啦！');"
	Response.Write "location.replace('admin_loadskin.asp')" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub

Private Sub InputSkin()
	If Trim(Request.Form("skinid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>出错啦！请选择你要导出的模板！</li>"
		Exit Sub
	End If
	If Trim(Request.Form("mdbname")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>出错啦！ 请选择你要导出的模板数据库！</li>"
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
		ErrMsg = ErrMsg + "<li>出错啦！没有找到你要导出的模板！</li>"
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
	Response.Write "alert('恭喜您!模板导出成功！');"
	Response.Write "location.replace('admin_loadskin.asp')" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub

Private Sub rename()
	Dim sRs,skid
	'模板改名
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
	Response.Write "<tr><th colspan=""2"">更改模版名称 ID="
	Response.Write sRs(2)
	Response.Write "</td></tr>" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write Chr(9) & "<td width=""20%"" class=""TableRow1"">模版原名：</td>" & vbCrLf
	Response.Write Chr(9) & "<td width=""80%"" class=""TableRow1"">"
	Response.Write sRs(1)
	Response.Write "</td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write Chr(9) & "<td class=""TableRow1"">模版新名：</td>" & vbCrLf
	Response.Write Chr(9) & "<td class=""TableRow1""><input type=""text"" name=""skinNAME"" size=""30"" value=""""></td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "<tr><td align=center class=TableRow2 colspan=""2""><input class=button type=""submit"" name=""submit"" value=""更 新""></td></tr>" & vbCrLf
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
	'模板改名保存
	skid = enchiasp.checkStr(Request.Form("skid"))
	mdbname = enchiasp.checkStr(Trim(Request.Form("mdbname")))
	skinNAME = enchiasp.checkStr(Trim(Request.Form("skinname")))
	If skid = "" Or Not IsNumeric(skid) Then
		ErrMsg = ErrMsg + "<BR><li>请选择正确的参数</li>"
		Exit Sub
	End If
	If skinNAME = "" Then
		ErrMsg = ErrMsg + "<li>新模板名称不能为空！</li>"
		Exit Sub
	End If
	If Request("act") = "loadskin" And mdbname <> "" Then
		SkinConnection(mdbname)
		sConn.Execute ("UPDATE ECCMS_Template SET page_name='" & skinNAME & "'  WHERE TemplateID=" & skid)
	Else
		enchiasp.Execute ("UPDATE ECCMS_Template SET page_name='" & skinNAME & "'  WHERE TemplateID=" & skid)
	End If
	Succeed ("<li>恭喜您，模板更名成功！")
End Sub
Private Sub DelTemplate()
	If Trim(Request("skinid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>出错啦！^_^ 请选择你要删除的模板！</li>"
		Exit Sub
	End If
	If Trim(Request("mdbname")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>出错啦！ 请选择你要删除的模板数据库！</li>"
		Exit Sub
	End If
	SkinConnection(mdbname)
	If FoundErr Then Exit Sub
	sConn.Execute("DELETE FROM ECCMS_Template WHERE skinid = " & CLng(Request("skinid")))
	Response.Write "<script language=JavaScript>" & vbCrLf
	Response.Write "alert('模板删除成功啦！');"
	Response.Write "location.replace('admin_loadskin.asp?action=load')" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub
'================================================
' 函数名：CompressMDB
' 作  用：压缩ACCESS数据库
' 参  数：dbPath ----数据库路径
' 返回值：True  ----  False
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
