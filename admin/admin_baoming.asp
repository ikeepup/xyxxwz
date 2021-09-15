<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
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
Dim Action,jobID
Dim i,isEdit,TextContent,FoundSQL,oRs,strClass
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Action = LCase(Request("action"))
%>

<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th>报名申请管理</th>
	
</tr>

</table>
<br>
<%
'权限判断
Select Case Trim(Action)
	Case "del"
		Call Deljiameng
	Case Else
		call MainPage
End Select

Admin_footer
SaveLogInfo(AdminName)
CloseConn


Private Sub Deljiameng()
	
	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ID有错误</li>"
		Exit Sub
	End If
	if IsSqlDataBase = 1 then
		enchiasp.Execute(" delete from [ECCMS_baoming] where ID = " & Request("ID"))
	else
		enchiasp.Execute(" delete * from [ECCMS_baoming] where ID = " & Request("ID"))
	end if
	'enchiasp.Execute(" delete [ECCMS_baoming] where ID = " & Request("ID"))
	OutHintScript("删除成功！")
End Sub
Private Sub MainPage()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th class=tablerow1>ID</th>
	<th class=tablerow1>类型</th>
	<th class=tablerow1>姓名</th>
	<th class=tablerow1>年龄</th>
	<th class=tablerow1>性别</th>
<th class=tablerow1>手机号码</th>
<th class=tablerow1>电话号码</th>
<th class=tablerow1>EMAIL</th>
<th class=tablerow1>联系地址</th>
<th class=tablerow1>邮编</th>
<th class=tablerow1>个人介绍</th>
<th class=tablerow1>申请时间</th>
<th class=tablerow1>操作</th>
</tr>
<%
	maxperpage = 20 '###每页显示数
	
	If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
		Response.Write ("错误的系统参数!请输入整数")
		Response.End
	End If
	If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
		CurrentPage = CInt(Request("page"))
	Else
		CurrentPage = 1
	End If
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	FoundSQL = " where baomingtypeid='"& request("baomingtypeid") &"'"
	TotalNumber = enchiasp.Execute("Select Count(ID) from ECCMS_baoming "& FoundSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '得到总页数
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_baoming] "& FoundSQL &" order by id desc "

	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=21 class=TableRow2>还没有找到任何信息！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td colspan=21 class=tablerow2><%Call showpage()%></td>
</tr>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
		Response.Write "<tr>"
		Response.Write "	<td " & strClass & ">"
		Response.Write Rs("id")
		Response.Write "	</td>"
		

Response.Write "	<td " & strClass & ">"
		Response.Write Rs("baomingtype")
		Response.Write "	</td>"

Response.Write "	<td " & strClass & ">"
		Response.Write Rs("xingming")
		Response.Write "	</td>"

Response.Write "	<td " & strClass & ">"
		Response.Write Rs("nianling")
		Response.Write "	</td>"

Response.Write "	<td " & strClass & ">"
		Response.Write Rs("xingbie")
		Response.Write "	</td>"


Response.Write "	<td " & strClass & ">"
		Response.Write Rs("shouji")
		Response.Write "	</td>"

Response.Write "	<td " & strClass & ">"
		Response.Write Rs("dianhua")
		Response.Write "	</td>"

Response.Write "	<td " & strClass & ">"
		Response.Write Rs("email")
		Response.Write "	</td>"

Response.Write "	<td " & strClass & ">"
		Response.Write Rs("dizhi")
		Response.Write "	</td>"

Response.Write "	<td " & strClass & ">"
		Response.Write Rs("youbian")
		Response.Write "	</td>"



Response.Write "	<td " & strClass & ">"
		Response.Write Rs("jianli")
		Response.Write "	</td>"



Response.Write "	<td " & strClass & ">"
		Response.Write Rs("writetime")
		Response.Write "	</td>"



Response.Write "	<td " & strClass & "><a href='?action=del&baomingtypeid="& rs("baomingtypeid") &"&id="& rs("id") &"'>"
		Response.Write "删除"
		Response.Write "	</a></td>"



		

%>

</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td colspan=21 class=tablerow2><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub showpage()
	Dim n
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action=?baomingtypeid=" & Request("baomingtypeid") & "><tr><td align=center> " & vbCrLf
	If CurrentPage < 2 Then
		Response.Write "共有 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 篇&nbsp;首 页&nbsp;上一页&nbsp;|&nbsp;"
	Else
		Response.Write "共有 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 篇&nbsp;<a href=?page=1&baomingtypeid=" & Request("baomingtypeid") & ">首 页</a>&nbsp;"
		Response.Write "<a href=?page=" & CurrentPage - 1 & "&baomingtypeid=" & Request("baomingtypeid") & ">上一页</a>&nbsp;|&nbsp;"
	End If
	If n - CurrentPage < 1 Then
		Response.Write "下一页&nbsp;尾 页" & vbCrLf
	Else
		Response.Write "<a href=?page=" & (CurrentPage + 1) & "&baomingtypeid=" & Request("baomingtypeid") & ">下一页</a>"
		Response.Write "&nbsp;<a href=?page=" & n & "&baomingtypeid=" & Request("baomingtypeid") & ">尾 页</a>" & vbCrLf
	End If
	Response.Write "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
	Response.Write "&nbsp;转到："
	Response.Write "<input name=page size=3 value='" & CurrentPage & "'> <input class=Button type=submit name=Submit value='转到'>"
	Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub

%>