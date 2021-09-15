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
Dim Action,jobbookID
Dim i,isEdit,TextContent,FoundSQL,oRs,strClass
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Action = LCase(Request("action"))
If Not ChkAdmin("adminjobbook") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
%>

<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
<%
if LCase(Request("isdel"))="1" then
%>
<th>应聘回收站管理</th>
<%
else
%>
	<th>应聘管理</th>
<%
end if
%>
</tr>
<tr>
	<td class=tablerow2><strong>管理导航：</strong> 
<a href='admin_jobbook.asp'><font color=blue>应聘管理</font></a> 
| <a href='admin_jobbook.asp?isdel=1'><font color=blue>应聘回收站</font></a> 
	</td>
</tr>
</table>
<br>
<%
Select Case Trim(Action)
	Case "view"
		Call Viewjobbook
	Case "del"
		Call Deljobbook
	case "huifu"
		call huifujobbook
	case "luyong"
		call luyongjobbook
	case "realdel"
		call realdel
	case "pinglun"
		call pinglun
	case "savepl"
		call savepl
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
	<th class=tablerow1>ID</th>
	<th class=tablerow1>应聘职位</th>
	<th class=tablerow1>姓名</th>
	<th class=tablerow1>性别</th>
	<th class=tablerow1>联系电话</th>
	<th class=tablerow1>学历</th>
	<th class=tablerow1>毕业学校</th>
	<th class=tablerow1>所学专业</th>
	<th class=tablerow1>简历递交日期</th>
	<th class=tablerow1>当前状态</th>
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
	
	if LCase(Request("isdel"))="1" then
		FoundSQL = " where isdel=1 "
	else
		FoundSQL = " where isdel=0 "
	end if

	
	TotalNumber = enchiasp.Execute("Select Count(ID) from ECCMS_jobbook "& FoundSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '得到总页数
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_jobbook] "& FoundSQL &" order by riqi desc ,ID desc"
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
		Response.Write "<tr><td align=center colspan=11 class=TableRow2>还没有找到任何应聘简历！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td colspan=11 class=tablerow2><%Call showpage()%></td>
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
		Response.Write "	<td " & strClass & "><a href='?action=view&ID="
		Response.Write Rs("ID")
		Response.Write "	'>"
		Response.Write Rs("id")
		Response.Write "	</a></td>"
		Response.Write "	<td align=center " & strClass & ">"
		Response.Write "	<a href='?action=view&ID="
		Response.Write Rs("ID")
		
		Response.Write "	&jobid="
		Response.Write Rs("jobID")
		Response.Write "	'>"

		
		Response.Write Rs("jobname")
		Response.Write "	</a>"
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write "	<a href='?action=view&ID="
		Response.Write Rs("ID")
		
		Response.Write "	&jobid="
		Response.Write Rs("jobID")
		Response.Write "	'>"
	

		Response.Write Rs("name")
		Response.Write "	</a>"
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("sex")
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("telephone")
		Response.Write "	</td>"
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("studydegree")
		Response.Write "	</td>"

		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("school")
		Response.Write "	</td>"
		
		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("specialty")
		Response.Write "	</td>"

		Response.Write "	<td align=center class=tablerow1>"
		Response.Write Rs("riqi")
		Response.Write "	</td>"
		
		Response.Write "	<td align=center class=tablerow1>"
		if rs("isuse")=1 then
			Response.Write "<font color=red>录用</font>"
		else
			Response.Write "<font color=red>未录用</font>"
		end if
	
		Response.Write "	</td>"


		
		

%>
	<%
	if LCase(Request("isdel"))="1" then
	%>
	
	
	<td align=center <%=strClass%>><a href='?action=realdel&ID=<%=Rs("ID")%>'>彻底删除</a> | <a href='?action=pinglun&ID=<%=Rs("ID")%>'>评价</a> | 
		<a href='?action=huifu&ID=<%=Rs("ID")%>'>恢复</a></td>
	<%
	else
	%>
		<td align=center <%=strClass%>><a href='?action=luyong&ID=<%=Rs("ID")%>'>录用</a> | <a href='?action=pinglun&ID=<%=Rs("ID")%>'>评价</a> | 
		<a href='?action=del&ID=<%=Rs("ID")%>' onclick="{if(confirm('您确定要删除该应聘信息吗?')){return true;}return false;}">删除</a></td>

	<%	
	end if
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
	<td colspan=11 class=tablerow2><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub Deljobbook()
	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的参数！</li>"
		Exit Sub
	End If
	enchiasp.Execute("update [ECCMS_jobbook] set isdel=1 where ID = " & Request("ID"))
	OutHintScript("应聘删除成功！")
End Sub

Private Sub realdel()
	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的参数！</li>"
		Exit Sub
	End If
	if IsSqlDataBase = 1 then
		enchiasp.Execute(" delete from [ECCMS_jobbook] where ID = " & Request("ID"))
	else
		enchiasp.Execute(" delete * from [ECCMS_jobbook] where ID = " & Request("ID"))
	end if
	OutHintScript("应聘档案删除成功！")
End Sub

private sub pinglun
	response.write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>"
	response.write "<form method=Post name='myform'action='admin_jobbook.asp?action=savepl&id="& Request("ID") &"'>"
	
	response.write "<tr>"
	response.write "	<th class=tablerow1>评论</th>"
	response.write "</tr>"
	response.write " <tr>"
	response.write "<td>"
	response.write "<textarea name='pinglun' cols='45' rows='6'  class='face' id='ability' style='font-size: 14px'>"& getpinglun() &"</textarea>"

	response.write " <input type='button' name='Submit4' onclick='javascript:history.go(-1)' value='返回上一页' class=Button>"
	response.write "<input name='Submit' type='submit' class='Button' value='评论' class=Button>"
	response.write "</td>"
	response.write "</tr>"
	
	response.write "</form>"
	response.write "</table>"
end sub

private function getpinglun()
	
	If Trim(Request("ID")) <> "" Then
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "select * from [ECCMS_jobbook] where ID = " & Request("ID") 
		If IsSqlDataBase = 1 Then
			If CurrentPage > 100 Then
				Rs.Open SQL, Conn, 1, 1
			Else
				Set Rs = Conn.Execute(SQL)
			End If
		Else
			Rs.Open SQL, Conn, 1, 1
		End If
		if rs.eof then
			getpinglun=""
		else
			getpinglun=rs("pinglun")
		end if
		Rs.Close:Set Rs = Nothing
		
	else
		getpinglun=""
	End If
end function



private sub savepl

	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请输入招聘ID！</li>"
		Exit Sub
	End If
	
		enchiasp.Execute("update [ECCMS_jobbook] set pinglun='"& Trim(Request.Form("pinglun")) &"' where ID = " & Request("ID"))

	
	OutHintScript("评论成功！")

end sub

Private Sub luyongjobbook()
	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请输入招聘ID！</li>"
		Exit Sub
	End If
	enchiasp.Execute("update [ECCMS_jobbook] set isuse=1 where ID = " & Request("ID"))
	OutHintScript("应聘录用成功，将更新前台空缺职位！")
End Sub

Private Sub huifujobbook()
	If Trim(Request("ID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请输入招聘ID！</li>"
		Exit Sub
	End If
	enchiasp.Execute("update [ECCMS_jobbook] set isdel=0 where ID = " & Request("ID"))
	OutHintScript("应聘恢复成功！")
End Sub


Private Sub Viewjobbook()
	If Request("ID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	dim n
	n = 1
	SQL = "select * from ECCMS_jobbook where ID=" & Request("ID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！没有找到任何应聘。或者您选择了错误的系统参数！</li>"
		Exit Sub
	Else
%>
<table border="0" align="center" cellpadding="2" cellspacing="1" class="TableBorder">
	<tr>
	  <th ></th>
	  <th >查看应聘信息</th>

	</tr>

	<tr>
	  <td class="TableRow1"><strong>应聘职位：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("jobname"))%></td>
	</tr>
	
	<tr>
	  <td class="TableRow1"><strong>应聘日期：</strong></td><td class="TableRow1"><%=enchiasp.FormatDate(Rs("riqi"), 2)%></td>
	</tr>

	<tr>
	  <td class="TableRow1"><strong>姓名：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("name"))%></td>
	</tr>

	<tr>
	  <td class="TableRow1"><strong>性别：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("sex"))%></td>
	</tr>

	
	<tr>
	  <td class="TableRow1"><strong>出生日期：</strong></td><td class="TableRow1"><%=enchiasp.FormatDate(Rs("birthday"), 2)%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>婚姻状况：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("marry"))%></td>
	</tr>
	
	<tr>
	  <td class="TableRow1"><strong>毕业院校：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("school"))%></td>
	</tr>

	
	<tr>
	  <td class="TableRow1"><strong>学历：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("studydegree"))%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>专业：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("specialty"))%></td>
	</tr>

	<tr>
	  <td class="TableRow1"><strong>毕业时间：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("gradyear"))%></td>
	</tr>

<tr>
	  <td class="TableRow1"><strong>联系电话：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("telephone"))%></td>
	</tr>
	
<tr>
	  <td class="TableRow1"><strong>EMAIL：</strong></td><td class="TableRow1"><a href=mailto:<%=enchiasp.ReadContent(Rs("email"))%>><%=enchiasp.ReadContent(Rs("email"))%></a><font color=red>（发信）</font></td>
	</tr>

<tr>
	  <td class="TableRow1"><strong>联系地址：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("address"))%></td>
	</tr>

<tr>
	  <td class="TableRow1"><strong>水平与能力：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("ability"))%></td>
	</tr>

<tr>
	  <td class="TableRow1"><strong>个人简历：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("resumes"))%></td>
	</tr>



	<tr>
	  <td class="TableRow2">上一应聘：<%=Frontjobbook(Rs("ID"))%>
	  <br>下一应聘：<%=Nextjobbook(Rs("ID"))%></td>
	  <td class="TableRow1"></td>
	</tr>
	<tr>
	  <td class="TableRow1"></td>
	  <td align="center" class="TableRow2"><input type="button" onclick="javascript:window.close()" value="关闭本窗口" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="返回上一页" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="返回顶部" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&ID=<%=Rs("ID")%>'" value="编辑应聘" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 	
	Viewjob()

End Sub

Private Function Frontjobbook(jobbookID)
	Dim Rss, SQL
	SQL = "select Top 1 ID,name from ECCMS_jobbook where ID < " & jobbookID & " order by ID desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		Frontjobbook = "已经没有了"
	Else
		Frontjobbook = "<a href=admin_jobbook.asp?action=view&ID=" & Rss("ID") & ">" & Rss("name") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function Nextjobbook(jobbookID)
	Dim Rss, SQL
	SQL = "select Top 1 ID,name from ECCMS_jobbook where ID > " & jobbookID & " order by ID asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		Nextjobbook = "已经没有了"
	Else
		Nextjobbook = "<a href=admin_jobbook.asp?action=view&ID=" & Rss("ID") & ">" & Rss("name") & "</a>"
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
		Response.Write "共有应聘 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 篇&nbsp;首 页&nbsp;上一页&nbsp;|&nbsp;"
	Else
		Response.Write "共有应聘 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 篇&nbsp;<a href=?page=1&ChannelID=" & Request("ChannelID") & ">首 页</a>&nbsp;"
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

Private Sub Viewjob()
	If Request("jobID") = "" Then
		Exit Sub
	End If
	dim n
	n = 1
	SQL = "select * from ECCMS_job where ID=" & Request("jobID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		Response.Write "<table align=center><tr><td align=center><font color=red>没有找到该岗位的招聘信息。或者您选择了错误的系统参数！</font></td></tr></table>"
		Exit Sub
	Else
%>
<table border="0" align="center" cellpadding="2" cellspacing="1" class="TableBorder">
	<tr>
	  <th ></th>
	  <th >该岗位相应招聘信息</th>

	</tr>

	<tr>
	  <td class="TableRow1"><strong>招聘职位：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("duix"))%></td>
	</tr>
	
	<tr>
	  <td class="TableRow1"><strong>招聘人数：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("rens"))%></td>
	</tr>
	
	<tr>
	  <td class="TableRow1"><strong>发布时间：</strong></td><td class="TableRow1"><%=enchiasp.FormatDate(Rs("riqi"), 2)%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>性别要求：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("sex"))%></td>
	</tr>
<tr>
	  <td class="TableRow1"><strong>学历要求：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("xueli"))%></td>
	</tr>
<tr>
	  <td class="TableRow1"><strong>专业要求：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("zhuanye"))%></td>
	</tr>

	
	<tr>
	  <td class="TableRow1"><strong>工作地点：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("did"))%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>有效期：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("qix"))%></td>
	</tr>
	
	<tr>
	  <td class="TableRow1"><strong>截止日期：</strong></td><td class="TableRow1"><%=enchiasp.FormatDate(Rs("riqi")+enchiasp.ReadContent(Rs("qix")), 2)%></td>
	</tr>

	
	<tr>
	  <td class="TableRow1"><strong>岗位待遇：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("daiy"))%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>岗位要求：</strong></td><td class="TableRow1"><%=enchiasp.ReadContent(Rs("yaoq"))%></td>
	</tr>
	
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub

%>