<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Dim i, j
Dim haveid
Set Rs = Server.CreateObject("ADODB.Recordset")
Server.ScriptTimeout = 9999999
Admin_header
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
If Not ChkAdmin("BatchReplace") Then
	Server.Transfer("showerr.asp")
	Request.End
End If
Select Case Trim(Request("Action"))
	Case "replace"
		Call ReplaceString
	Case "search"
		Call TableColumn
	Case "table"
		Call Tabletop
	Case Else
		Call ReplaceMain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Private Sub ReplaceMain()
	Response.Write "<form name=""myform"" action="""" method=""post"">" & vbNewLine
	Response.Write " <table cellpadding=""3"" cellspacing=""1"" border=""0"" width=""100%"" class=""tableBorder"" align=center>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <th height=""22"">数据库批量替换管理——选择数据表名</th>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <td height=""25"" align=""center"" class=TableRow1>请选择要替换的数据表名： " & vbNewLine
	Response.Write " <input name=""action"" type=""hidden"" value=""table"">" & vbNewLine
	Response.Write " <select name=""TableName"">" & vbNewLine
	Response.Write " "
	Set Rs = Conn.openSchema(28)
	While Not Rs.EOF
		Response.Write ("<option value='" & Rs(2) & "'>" & Rs(2) & "</option>")
		Rs.movenext
	Wend
	Response.Write " </select>" & vbNewLine
	Response.Write " <input type=""submit"" name=""Submit"" value=""下一步"" class=button>" & vbNewLine
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td class=TableRow1 align=""center"">请选择要替换的数据表</td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " </table>" & vbNewLine
	Response.Write " </form>" & vbNewLine
End Sub


Private Sub Tabletop()
	Session("TableName") = enchiasp.checkStr(Trim(Request.Form("TableName")))
	Response.Write "<form name=""myform"" action="""" method=""post"">" & vbNewLine
	Response.Write " <table cellpadding=""3"" cellspacing=""1"" border=""0"" width=""100%"" class=""tableBorder"" align=center>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <th height=""22"">数据库批量替换管理——选择字段名输入查找内容</th>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <td height=""25"" align=""center"" class=TableRow1>要替换的数据表名： " & vbNewLine
	Response.Write " <select name=""TableName"">" & vbNewLine
	Response.Write " <option value="""
	Response.Write Session("TableName")
	Response.Write """ selected>"
	Response.Write Session("TableName")
	Response.Write "</option>" & vbNewLine
	Response.Write " </select>" & vbNewLine
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <td height=""25"" align=""center"" class=TableRow1>要替换的字段名： " & vbNewLine
	Response.Write " <select name=""ColumnName"">" & vbNewLine
	Response.Write " "
	haveid = 0
	Set Rs = CreateObject("adodb.recordset")
	SQL = "SELECT * FROM [" & Session("TableName") & "] WHERE 1<>1"
	Rs.Open SQL, Conn, 1, 1
	j = Rs.Fields.Count
	Session("ECCMS_PRIMARY") = Rs.Fields(0).Name
	For i = 0 To (j - 1)
		Response.Write ("<option value='" & Rs.Fields(i).Name & "'>" & Rs.Fields(i).Name & "</option>")
	Next
	Rs.Close
	Response.Write " </select>" & vbNewLine
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <td height=""25"" align=""center"" class=TableRow1>字段名中包含的字符： " & vbNewLine
	Response.Write " <input name=""action"" type=""hidden"" value=""search"">" & vbNewLine
	Response.Write " <input name=""oldString"" type=""text"" size=""45"">" & vbNewLine
	Response.Write " <input type=""submit"" name=""Submit"" value=""开始查找"" class=button>" & vbNewLine
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td class=TableRow1 align=""center""><a href=""javascript:history.go(-1)"" >&lt;&lt; 返回上一页</a></td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <td height=""25"" align=""center"" class=TableRow1>注意：单引号“'”将被自动过滤掉</td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " </table>" & vbNewLine
	Response.Write " </form>" & vbNewLine
End Sub


Private Sub TableColumn()
	Response.Write " <form name=""myform"" action="""" method=""post"">" & vbNewLine
	Response.Write " <table cellpadding=""3"" cellspacing=""1"" border=""0"" width=""100%"" class=""tableBorder"" align=center>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <th height=""22"">数据库批量替换管理——替换" & vbNewLine
	Response.Write "<input name=""action"" type=""hidden"" value=""replace""></th>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <td height=""25"" align=""center"" class=TableRow1> " & vbNewLine
	Session("ColumnName") = enchiasp.checkStr(Trim(Request.Form("ColumnName")))
	Set Rs = CreateObject("adodb.recordset")
	SQL = "SELECT COUNT(" & Session("ECCMS_PRIMARY") & ") FROM " & Session("TableName") & " WHERE " & Session("ColumnName") & " like '%" & enchiasp.checkStr(Trim(Request.Form("oldString"))) & "%'"
	Rs.Open SQL, Conn, 1, 1
	If Not (Rs.EOF And Rs.bof) Then
		Session("oldString") = enchiasp.checkStr(Trim(Request.Form("oldString")))
		Response.Write ("本次搜索找到了 <b>" & Rs(0) & "</b> 个相关字符串。")
		Response.Write ("<a href=""javascript:history.go(-1)"">返回重新查找</a>")
	Else
		Response.Write ("没有找到相关字符串，<a href=""javascript:history.go(-1)"">返回重新查找</a>")
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <td height=""25"" align=""center"" class=TableRow1>将字符 " & vbNewLine
	Response.Write " <input disabled name=""oldString"" type=""text"" value="""
	Response.Write Trim(Request.Form("oldString"))
	Response.Write """ size=""45""> " & vbNewLine
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <td height=""25"" align=""center"" class=TableRow1>替换成 " & vbNewLine
	Response.Write " <input name=""newString"" type=""text"" value="""" size=""45""> " & vbNewLine
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr> " & vbNewLine
	Response.Write " <td height=""25"" align=""center"" class=TableRow1>" & vbNewLine
	Response.Write "<input type=""submit"" name=""Submit2"" value=""开始替换"" class=button></td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td class=TableRow1 align=""center""><a href=""javascript:history.go(-1)"" >&lt;&lt; 返回上一页</a></td>" & vbNewLine
	Response.Write " </tr>" & vbNewLine
	Response.Write " </table>" & vbNewLine
	Response.Write " </form>" & vbNewLine
End Sub


Private Sub ReplaceString()
	Dim oldString
	Dim newString
	Dim TableName
	Dim ColumnName
	Dim RepString
	Dim uprs
	Dim id
	oldString = enchiasp.checkStr(Trim(Session("oldString")))
	newString = enchiasp.checkStr(Trim(Request.Form("newString")))
	TableName = enchiasp.checkStr(Trim(Session("TableName")))
	ColumnName = enchiasp.checkStr(Trim(Session("ColumnName")))
	id = Trim(Session("ECCMS_PRIMARY"))
	Set Rs = CreateObject("adodb.recordset")
	Set uprs = CreateObject("adodb.recordset")
	i = 0
	SQL = "SELECT " & id & "," & Session("ColumnName") & " FROM " & Session("TableName") & " WHERE " & Session("ColumnName") & " like '%" & Trim(Session("oldString")) & "%'"
	Rs.Open SQL, Conn, 1, 1
	If Not (Rs.EOF And Rs.BOF) Then
		'i = Rs.recordcount
		Do While Not Rs.EOF
			RepString = Replace(Rs(1), "" & oldString & "", "" & newString & "")
			SQL = "SELECT * FROM " & TableName & " WHERE " & id & "=" & Rs(0)
			uprs.Open SQL, Conn, 1, 3
			uprs(ColumnName) = RepString
			uprs.Update
			uprs.Close
			Rs.movenext
			i = i + 1
		Loop
	End If
	Rs.Close
	Set uprs = Nothing
	Set Rs = Nothing
	Succeed("<li>批量替换操作成功，共更新了 " & i & " 条信息！</li>")
End Sub
%>
