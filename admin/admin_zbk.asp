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
	<th>质保卡管理</th>
	
</tr>
<tr>
	<td class=tablerow2><strong>导航：</strong> <a href='admin_zbk.asp'>质保卡首页</a> 
| <a href='admin_zbk.asp?action=add'><font color=blue>登记质保卡</font></a> 
	</td>
</tr>
</table>
<br>
<%
'权限判断
Select Case Trim(Action)
	Case "save"
		Call Savejob
	Case "modify"
		Call Modifyjob
	Case "add"
		isEdit = False
		Call Editjob(isEdit)
	Case "edit"
		isEdit = True
		Call Editjob(isEdit)
	Case "view"
		Call Viewjob
	Case "del"
		Call Deljob
	Case Else
		Call MainPage
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn



Private Sub MainPage()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th class=tablerow1>xh</th>
	<th class=tablerow1>质保卡编号</th>
	<th class=tablerow1>车牌号</th>
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
	FoundSQL = " "
	TotalNumber = enchiasp.Execute("Select Count(xh) from ECCMS_zb "& FoundSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '得到总页数
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_zb] "& FoundSQL &" order by xh desc"

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
		Response.Write "<tr><td align=center colspan=4 class=TableRow2>还没有找到任何信息！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td colspan=4 class=tablerow2><%Call showpage()%></td>
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
		Response.Write Rs("xh")
		Response.Write "	</td>"
		Response.Write "	<td align=center " & strClass & ">"
		Response.Write Rs("bh")
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("cph")
		Response.Write "	</td>"
			

%>
	
	<td align=center <%=strClass%>><a href='?action=edit&xh=<%=Rs("xh")%>'>编辑</a> | 
	<a href='?action=del&xh=<%=Rs("xh")%>' onclick="{if(confirm('信息删除后将不能恢复，您确定要删除该信息吗?')){return true;}return false;}">删除</a></td>
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
	<td colspan=4 class=tablerow2><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub Editjob(isEdit)

	Dim EditTitle
	If isEdit Then
		SQL = "select * from ECCMS_zb where xh=" & Request("xh")
		Set Rs = enchiasp.Execute(SQL)
		EditTitle = "编辑"
	Else
		EditTitle = "添加"
	End If
%>
<script language=javascript>
    function CheckForm(form1)
{
	if (!validateSubmit()) return (false);
	form1.content.value=IframeID.document.body.innerHTML; 
	MessageLength=IframeID.document.body.innerHTML.length;
}
</script>
<div onkeydown=CtrlEnter()>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
  <tr>
    <th colspan="2"><%=EditTitle%></th>
  </tr>
    	<form method=Post name="myform" action="admin_zbk.asp" onSubmit="return CheckForm(this);">
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""ID"" value="""& Request("ID") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
%>
  <tr>
    <td width="20%" align="right" class="TableRow2"><strong>质保卡编号：</strong></td>
    <td width="80%" class="TableRow1"><input name="bh" type="text" id="bh" size="50" value='<%If isEdit Then Response.Write Rs("bh")%>'> 
      <span class="style1">* </span></td>
  </tr>
  
	<tr>
    <td width="20%" align="right" class="TableRow2"><strong>车牌号：</strong></td>
    <td width="80%" class="TableRow1"><input name="cph" type="text" id="cph" size="50" value='<%If isEdit Then Response.Write Rs("cph") else response.write "晋M" end if%>'> 
    </td>
  </tr>

  
  <tr>
    <td align="right" class="TableRow2">　</td>
    <td align="center" class="TableRow1">
    <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>
    <input name="Submit1" type="submit" class="Button" value="保存" class=Button></td>
  </tr></form>
</table>
</div>
<%
	If isEdit Then Rs.Close:Set Rs = Nothing
End Sub

Private Sub CheckSave()
	If Trim(Request.Form("bh")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>质保卡编号不能为空！</li>"
	End If
	If Trim(Request.Form("cph")) = "" or Trim(Request.Form("cph")) = "晋M"  Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>车牌号不能为空！</li>"
	End If
End Sub
Private Sub Savejob()
	

	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_zb where (xh is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("bh") = Trim(Request.Form("bh"))
		Rs("cph") = enchiasp.ChkFormStr(Request.Form("cph"))
		
	Rs.update
	Rs.Close:Set Rs = Nothing
	Succeed("<li>恭喜您！添加新的信息成功。</li>")
End Sub
Private Sub Modifyjob()
	

	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_zb where xh= " & Request("xh")
	Rs.Open SQL,Conn,1,3
		Rs("bh") = Trim(Request.Form("bh"))
		Rs("cph") = enchiasp.ChkFormStr(Request.Form("cph"))
	Rs.update
	Rs.Close:Set Rs = Nothing
	Succeed("<li>恭喜您！修改成功。</li>")
End Sub
Private Sub Deljob()
	
	If Trim(Request("XH")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请输入XH！</li>"
		Exit Sub
	End If
	enchiasp.Execute("delete from [ECCMS_zb] where xh= " & Request("xh"))
	OutHintScript("删除成功！")
End Sub




Private Sub showpage()
	Dim n
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action=?ChannelID=" & Request("ChannelID") & "><tr><td align=center> " & vbCrLf
	If CurrentPage < 2 Then
		Response.Write "共有 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 篇&nbsp;首 页&nbsp;上一页&nbsp;|&nbsp;"
	Else
		Response.Write "共有 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 篇&nbsp;<a href=?page=1&ChannelID=" & Request("ChannelID") & ">首 页</a>&nbsp;"
		Response.Write "<a href=?page=" & CurrentPage - 1 & "&ChannelID=" & Request("ChannelID") & ">上一页</a>&nbsp;|&nbsp;"
	End If
	If n - CurrentPage < 1 Then
		Response.Write "下一页&nbsp;尾 页" & vbCrLf
	Else
		Response.Write "<a href=?page=" & (CurrentPage + 1) & "&ChannelID=" & Request("ChannelID") & ">下一页</a>"
		Response.Write "&nbsp;<a href=?page=" & n & "&ChannelID=" & Request("ChannelID") & ">尾 页</a>" & vbCrLf
	End If
	Response.Write "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
	Response.Write "&nbsp;转到："
	Response.Write "<input name=page size=3 value='" & CurrentPage & "'> <input class=Button type=submit name=Submit value='转到'>"
	Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub

%>