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
Dim Action,AnnounceID
Dim i,isEdit,TextContent,FoundSQL,oRs,strClass
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Action = LCase(Request("action"))
If Not ChkAdmin("Announce") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th>公告管理</th>
</tr>
<tr>
	<td class=tablerow2><strong>公告导航：</strong> <a href='admin_announce.asp'>管理首页</a> 
<%
	Set oRs = enchiasp.Execute("Select ChannelID,ChannelName,ChannelType From ECCMS_Channel where ChannelType < 2  Order By orders")
	Do While Not oRs.EOF
		Response.Write " | <a href='?ChannelID="
		Response.Write oRs("ChannelID")
		Response.Write "'>"
		Response.Write oRs("ChannelName")
		Response.Write "</a>"
	oRs.movenext
	Loop
	oRs.Close:Set oRs = Nothing
%>
| <a href='admin_announce.asp?action=add'><font color=blue>发布公告</font></a> 
	</td>
</tr>
</table>
<br>
<%
Select Case Trim(Action)
	Case "save"
		Call SaveAnnounce
	Case "modify"
		Call ModifyAnnounce
	Case "add"
		isEdit = False
		Call EditAnnounce(isEdit)
	Case "edit"
		isEdit = True
		Call EditAnnounce(isEdit)
	Case "view"
		Call ViewAnnounce
	Case "del"
		Call DelAnnounce
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
	<th class=tablerow1>公告标题</th>
	<th class=tablerow1>显示位置</th>
	<th class=tablerow1>公告类型</th>
	<th class=tablerow1>操作选项</th>
	<th class=tablerow1>发布时间</th>
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
	If Request("ChannelID") <> "" Then
		FoundSQL = "where ChannelID = " & Request("ChannelID")
	Else
		FoundSQL = ""
	End If
	TotalNumber = enchiasp.Execute("Select Count(AnnounceID) from ECCMS_Announce "& FoundSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '得到总页数
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Announce] "& FoundSQL &" order by PostTime desc ,AnnounceID desc"
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
		Response.Write "<tr><td align=center colspan=5 class=TableRow2>还没有找到任何公告！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td colspan=5 class=tablerow2><%Call showpage()%></td>
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
		Response.Write "	<td " & strClass & "><a href='?action=view&AnnounceID="
		Response.Write Rs("AnnounceID")
		Response.Write "	'>"
		Response.Write Rs("title")
		Response.Write "	</a></td>"
		Response.Write "	<td align=center " & strClass & ">"

		If Rs("ChannelID") = 0 Then
			Response.Write "首页公告"
		ElseIf Rs("ChannelID") = 1 Then
			Response.Write "<span class=style1>文章频道</span>"
		ElseIf Rs("ChannelID") = 2 Then
			Response.Write "<span class=style2>下载频道</span>"
		ElseIf Rs("ChannelID") = 3 Then
			Response.Write "<span class=style3>商城频道</span>"
		ElseIf Rs("ChannelID") = 4 Then
			Response.Write "<span class=style2>动画频道</span>"
		ElseIf Rs("ChannelID") = 5 Then
			Response.Write "<span class=style3>留言频道</span>"
		ElseIf Rs("ChannelID") = 6 Then
			Response.Write "<span class=style3>单页面图文频道</span>"
		ElseIf Rs("ChannelID") = 7 Then
			Response.Write "<span class=style3>招聘频道</span>"

		Else
			Response.Write "<span class=style1>所有页面显示</span>"
		End If
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		If Rs("AnnounceType") = 1 Then
			Response.Write "<span class=style2>内容公告</span>"
		ElseIf Rs("AnnounceType") = 2 Then
			Response.Write "<span class=style1>弹出公告</span>"
		Else
			Response.Write "列表公告"
		End If
%>
	</td>
	<td align=center <%=strClass%>><a href='?action=edit&AnnounceID=<%=Rs("AnnounceID")%>'>编辑</a> | 
	<a href='?action=del&AnnounceID=<%=Rs("AnnounceID")%>' onclick="{if(confirm('公告删除后将不能恢复，您确定要删除该公告吗?')){return true;}return false;}">删除</a></td>
	<td align=center <%=strClass%>>
<%
		If Rs("PostTime") >= Date Then
			Response.Write "<font color=red>"
			Response.Write enchiasp.FormatDate(Rs("PostTime"), 2)
			Response.Write "</font>"
		Else
			Response.Write enchiasp.FormatDate(Rs("PostTime"), 2)
		End If
%>
	</td>
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
	<td colspan=5 class=tablerow2><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub EditAnnounce(isEdit)
	Dim EditTitle
	If isEdit Then
		SQL = "select * from ECCMS_Announce where AnnounceID=" & Request("AnnounceID")
		Set Rs = enchiasp.Execute(SQL)
		EditTitle = "编辑公告"
	Else
		EditTitle = "添加公告"
	End If
%>
<script language=javascript>
    function CheckForm(form1)
{
	if (!validateSubmit()) return (false);
	if (form1.title.value == "")
	{
		alert("公告标题不能为空！");
		form1.title.focus();
		return (false);
	}
	form1.content.value=IframeID.document.body.innerHTML; 
	MessageLength=IframeID.document.body.innerHTML.length;
	if(MessageLength<2){alert("公告内容不能小于2个字符！");return false;}
}
</script>
<div onkeydown=CtrlEnter()>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
  <tr>
    <th colspan="2"><%=EditTitle%></th>
  </tr>
    	<form method=Post name="myform" action="admin_announce.asp" onSubmit="return CheckForm(this);">
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""AnnounceID"" value="""& Request("AnnounceID") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
%>
  <tr>
    <td width="20%" align="right" class="TableRow2"><strong>公告标题：</strong></td>
    <td width="80%" class="TableRow1"><input name="title" type="text" id="title" size="50" value='<%If isEdit Then Response.Write Rs("title")%>'> 
      <span class="style1">* </span></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>所属频道：</strong></td>
    <td class="TableRow1"><select name="ChannelID" id="ChannelID">
      <option value="0"<%If isEdit Then If Rs("ChannelID") = 0 Then Response.Write " selected"%>>首页公告</option>
<%
	Set oRs = enchiasp.Execute("Select ChannelID,ChannelName,ChannelType From ECCMS_Channel where ChannelType < 2  Order By orders")
	Do While Not oRs.EOF
		Response.Write "<option value="""& oRs("ChannelID") &""""
		If isEdit Then
			If oRs("ChannelID") = Rs("ChannelID") Then Response.Write " selected"
		End If
		Response.Write ">"
		Response.Write oRs("ChannelName")
		Response.Write "</option>"
	oRs.movenext
	Loop
	oRs.Close:Set oRs = Nothing
%>
      <option value="999"<%If isEdit Then If Rs("ChannelID") = 999 Then Response.Write " selected"%>>所有频道显示</option>
    </select></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>公告类型：</strong></td>
    <td class="TableRow1"><input name="AnnounceType" type="radio" value="0"<%If isEdit Then If Rs("AnnounceType") = 0 Then Response.Write " checked" End If:Else Response.Write " checked" End If%>>
列表公告
<input type="radio" name="AnnounceType" value="1"<%If isEdit Then If Rs("AnnounceType") = 1 Then Response.Write " checked"%>>
内容公告</td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>公告内容：</strong></td>
    <td class="TableRow1"><textarea name="content" style="display:none" id="content"><%If isEdit Then Response.Write Server.HTMLEncode(Rs("content"))%></textarea>
    <script src='../editor/edit.js' type=text/javascript></script></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>上传文件：</strong></td>
    <td class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upfiles.asp></iframe></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>发布人：</strong></td>
    <td class="TableRow1"><input name="writer" type="text" id="writer" size="15" value='<%If isEdit Then Response.Write Rs("title") Else Response.Write AdminName End If%>'> 
      <span class="style1">* </span> 
      <%If isEdit Then%>
      <input name="update" type="checkbox" id="update" value="yes">
更新公告时间 
<%End If%></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2">　</td>
    <td align="center" class="TableRow1">
    <input type="button" name="Submit2" onclick="CheckLength();" value="查看内容长度" class=Button>
    <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>
    <input name="Submit1" type="submit" class="Button" value="保存公告" class=Button></td>
  </tr></form>
  <tr>
    <td colspan="2" class="TableRow1"><strong>说明：</strong><br>
      &nbsp;&nbsp;&nbsp;&nbsp;所属频道 ---- 只有频道首页才显示公告；<br>
      &nbsp;&nbsp;&nbsp;&nbsp;公告类型 ---- 列表显示是指公告以列表的形式显示公告，需要用户点击才可以看到公告内容；内容公告是指公告以内容的方式显示在所在的频道首页，注意只显示最新的一条公告。</td>
  </tr>
</table>
</div>
<%
	If isEdit Then Rs.Close:Set Rs = Nothing
End Sub

Private Sub CheckSave()
	If Trim(Request.Form("title")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>公告标题不能为空！</li>"
	End If
	If Trim(Request.Form("ChannelID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>频道ID不能为空！</li>"
	End If
	If Trim(Request.Form("AnnounceType")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>公告类型不能为空！</li>"
	End If
	If Trim(Request.Form("content")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>公告内容不能为空！</li>"
	End If	
	TextContent = ""
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next
	If Len(TextContent) < 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>公告发布人不能为空！</li>"
	End If
End Sub
Private Sub SaveAnnounce()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Announce where (AnnounceID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = Trim(Request.Form("ChannelID"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("AnnounceType") = Request.Form("AnnounceType")
		Rs("Content") = TextContent
		Rs("writer") = enchiasp.ChkFormStr(Request.Form("writer"))
		Rs("PostTime") = Now()
		Rs("hits") = 0
	Rs.update
	Rs.Close
	Rs.Open "select top 1 AnnounceID from ECCMS_Announce order by AnnounceID desc", Conn, 1, 1
	AnnounceID = Rs("AnnounceID")
	Rs.Close:Set Rs = Nothing
	Succeed("<li>恭喜您！添加新的公告成功。</li><li><a href=?action=view&AnnounceID=" & AnnounceID & ">点击此处查看该公告</a></li>")
End Sub
Private Sub ModifyAnnounce()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Announce where AnnounceID = " & Request("AnnounceID")
	Rs.Open SQL,Conn,1,3
		Rs("ChannelID") = Trim(Request.Form("ChannelID"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("AnnounceType") = Request.Form("AnnounceType")
		Rs("Content") = TextContent
		Rs("writer") = enchiasp.ChkFormStr(Request.Form("writer"))
		If LCase(Request.Form("Update")) = "yes" Then Rs("PostTime") = Now()
	Rs.update
		AnnounceID = Rs("AnnounceID")
	Rs.Close:Set Rs = Nothing
	Succeed("<li>恭喜您！修改公告成功。</li><li><a href=?action=view&AnnounceID=" & AnnounceID & ">点击此处查看该公告</a></li>")
End Sub
Private Sub DelAnnounce()
	If Trim(Request("AnnounceID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请输入公告ID！</li>"
		Exit Sub
	End If
	enchiasp.Execute("Delete From ECCMS_Announce where AnnounceID = " & Request("AnnounceID"))
	OutHintScript("公告删除成功！")
End Sub

Private Sub ViewAnnounce()
	If Request("AnnounceID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	dim n
	n = 1
	enchiasp.Execute ("update ECCMS_Announce set hits = hits + "&n&" where AnnounceID=" & Request("AnnounceID"))
	SQL = "select * from ECCMS_Announce where AnnounceID=" & Request("AnnounceID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！没有找到任何公告。或者您选择了错误的系统参数！</li>"
		Exit Sub
	Else
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
	<tr>
	  <th colspan="2">查看公告</th>
	</tr>
	<tr>
	  <td align="center" class="TableRow2" colspan="2"><a href=?action=edit&AnnounceID=<%=Rs("AnnounceID")%>><font size=4><%=Rs("title")%></font></a></td>
	</tr>
	<tr>
	  <td align="center" class="TableRow1"><strong>发布时间：</strong> <%=Rs("PostTime")%> &nbsp;&nbsp;
	  <strong>发 布 人：</strong> <%=Rs("writer")%> &nbsp;&nbsp;<strong>浏览次数：</strong> <%=Rs("hits")%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>公告内容：</strong><br><%=enchiasp.ReadContent(Rs("content"))%></td>
	</tr>
	<tr>
	  <td class="TableRow2">上一公告：<%=FrontAnnounce(Rs("AnnounceID"))%>
	  <br>下一公告：<%=NextAnnounce(Rs("AnnounceID"))%></td>
	</tr>
	<tr>
	  <td align="center" colspan="2" class="TableRow2"><input type="button" onclick="javascript:window.close()" value="关闭本窗口" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="返回上一页" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="返回顶部" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&AnnounceID=<%=Rs("AnnounceID")%>'" value="编辑公告" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub

Private Function FrontAnnounce(AnnounceID)
	Dim Rss, SQL
	SQL = "select Top 1 AnnounceID,title from ECCMS_Announce where AnnounceID < " & AnnounceID & " order by AnnounceID desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		FrontAnnounce = "已经没有了"
	Else
		FrontAnnounce = "<a href=admin_Announce.asp?action=view&AnnounceID=" & Rss("AnnounceID") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function NextAnnounce(AnnounceID)
	Dim Rss, SQL
	SQL = "select Top 1 AnnounceID,title from ECCMS_Announce where AnnounceID > " & AnnounceID & " order by AnnounceID asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		NextAnnounce = "已经没有了"
	Else
		NextAnnounce = "<a href=admin_Announce.asp?action=view&AnnounceID=" & Rss("AnnounceID") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Sub showpage()
	Dim n
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action=?ChannelID=" & Request("ChannelID") & "><tr><td align=center> " & vbCrLf
	If CurrentPage < 2 Then
		Response.Write "共有公告 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 篇&nbsp;首 页&nbsp;上一页&nbsp;|&nbsp;"
	Else
		Response.Write "共有公告 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 篇&nbsp;<a href=?page=1&ChannelID=" & Request("ChannelID") & ">首 页</a>&nbsp;"
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
Public Sub CreateAnnounce()
	Dim rsAnnounce,sqlAnnounce
	sqlAnnounce = "select A.AnnounceID,A.title,A.PostTime,A.AnnounceType,C.ChannelDir,C.ChannelUrl,C.BindDomain,C.DomainName from [ECCMS_Announce] A inner join [ECCMS_Channel] C On A.ChannelID=C.ChannelID where A.ChannelID=" & ChannelID & ""
	Set rsAnnounce = enchiasp.Execute(sqlAnnounce)
End Sub

%>