<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/cls_public.asp"-->
<!--#include file="../inc/NewsChannel.asp"-->
<%

Server.ScriptTimeout = 99999
Admin_header
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID = 0 Then ChannelID = 6
'=====================================================================
' 软件名称：恩池网站管理系统----单页面图文生成HTML管理界面
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim Action,flag,i
Dim TableTop,TableEnd
TableTop = "<p align=center>"
TableEnd = "</p><p align=center><a href='" & LinkedStr("") & "'><b class=style2>恭喜您！生成HTML完成...总费时<font color=""#FF0000"">" & FormatNumber((Timer() - startime), 2,-1) & "</font>秒,完成时间" & Now() & " 请返回进行其它操作</b></a></p>"

Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
Response.Write "<tr>" & vbNewLine
Response.Write "	<th colspan=2>" & sModuleName & "HTML管理导航</th>" & vbNewLine
Response.Write "</tr>" & vbNewLine
Response.Write "<tr>" & vbNewLine
Response.Write "	<td colspan=2 class=tablerow1><b>选项：</b>" & vbNewLine
Response.Write "	<a href=" & LinkedStr("") & ">管理首页</a> | " & vbNewLine
Response.Write "	<a href=" & LinkedStr("index") & ">生成首页HTML</a> | " & vbNewLine
Response.Write "	<a href=" & LinkedStr("info") & ">生成全部文章内容</a>" & vbNewLine
Response.Write "	</td>" & vbNewLine
Response.Write "</tr>" & vbNewLine
Response.Write "</table><br>" & vbNewLine
Flag = "CreateArticle" & ChannelID
Action = LCase(Request("action"))
If Not ChkAdmin(Flag) Then
	Server.Transfer("showerr.asp")
	Response.End
End If
If CInt(enchiasp.IsCreateHtml) = 0 Then
	ErrMsg = "对不起~！此频道不支持生成HTML功能，如果你要使用生成HTML功能，请在此频道管理中打开HTML生成功能。"
	Response.Redirect ("showerr.asp?action=error&Message="& ErrMsg)
	Response.End
End If

enchicms.Channel = ChannelID
enchicms.ChannelMain

Select Case Trim(Action)
	Case "index"
		Response.Write "<ol>"
		enchicms.ShowFlush = 1
		enchicms.CreateArticleIndex
		Call CreateSiteIndex
		Response.Write "</ol>"
		Response.Write "<p align=center><a href='" & LinkedStr("") & "'><b class=style2>恭喜您！生成HTML完成...总费时<font color=""#FF0000"">" & FormatNumber((Timer() - startime), 2,-1) & "</font>秒,完成时间" & Now() & " 请返回进行其它操作</b></a></p>"
	Case "list"
		Call Create_Article_List
	Case "cont"
		Call Create_Article_Content
	Case "info"
		Call Create_Article_Info
	Case "renewall"
		Call RenewAll_Article
	Case Else
		Call ShowMain
End Select
If Action <> "info" Then Admin_footer
If FoundErr = True Then
	ReturnError(ErrMsg)
End If

If IsObject(enchicms) Then
	Set enchicms = Nothing
End If
CloseConn
Sub ShowMain()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<th colspan=2>批量生成" & sModuleName & "HTML页管理</th>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write "<tr height=""32"">"
	Response.Write "	<td colspan=2 class=tablerow1>"
	If CInt(Request("stype")) = 1 Then
		Response.Write "<b class=style2>恭喜您！生成" & sModuleName & "分类列表的HTML页完成..."
		Response.Write "总费时<font color=""#FF0000"">" & FormatNumber((Timer() - Request("d")), 2,-1) & "</font>秒,完成时间" & Now() & "</b>"
	End If
	Session("totalnumber") = ""
	Session("PageCount") = ""
	If CInt(Request("stype")) = 2 Then
		Response.Write "<b class=style2>恭喜您！生成" & sModuleName & "内容的HTML页完成..."
		Response.Write "共生成页面<font color=""#FF0000"">" & Request("num") & "</font>个,总费时<font color=""#FF0000"">" & FormatNumber((Timer() - Request("D")), 2,-1) & "</font>秒,完成时间" & Now() & "</b>"
		Session("BuildArticleID")=""
	End If
	Response.Write "</td></tr>"
	Response.Write "<tr>"
	Response.Write "	<td colspan=2 class=tablerow2>①、<font color=blue>生成" & sModuleName & "内容HTML页</font></td>"
	Response.Write "</tr>"
	Response.Write "</tr>"
	Response.Write " <tr><form name=""Createform2"" method=""post"" action=""?action=cont&type=date"">"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write " <td class=tablerow1>按日期生成内容HTML页</td>"
	Response.Write " <td class=tablerow1>"
	Response.Write " <select size=""1"" name=""datYear"">"
	For i = 2001 To Year(Date)
		Response.Write "<option value=""" & i & """"
		If i = Year(Date) Then
			Response.Write " selected"
		End If
		Response.Write ">" & i & "</option>"
	Next
	Response.Write " </select> - "
	Response.Write " <select size=""1"" name=""datMonth"">"
	For i = 1 To 12
		Response.Write "<option value=""" & i & """"
		If i = Month(Date) Then
			Response.Write " selected"
		End If
		Response.Write ">" & i & "</option>"
	Next
	Response.Write " </select> - "
	Response.Write " <select size=""1"" name=""datDay"">"
	For i = 1 To 31
		Response.Write "<option value=""" & i & """"
		If i = Day(Date) Then
			Response.Write " selected"
		End If
		Response.Write ">" & i & "</option>"
	Next
	Response.Write " </select>　"
	Response.Write "<input type=""submit"" class=button name=""Submit"" value=""生成HTML页""></td>"
	Response.Write " </tr></form>"
	Response.Write " <tr><form name=Createform3 method=""post"" action=""?action=cont&type=new"">"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write " <td class=tablerow1 height=""22"">按个数生成内容HTML页</td>"
	Response.Write " <td class=tablerow1><input size=10 name=""newnum"" value=""1"">　<input type=""submit"" name=""Submit"" class=button value=""生成HTML页""> <b>说明：</b>将按最近更新的顺序生成</td>"
	Response.Write " </tr></form>"
	Response.Write " <tr><form name=Createform4 method=""post"" action=""?action=cont&type=seid"">"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write " <td class=tablerow1 height=""22"">按ID生成内容HTML页</td>"
	Response.Write " <td class=tablerow1><input size=45 name=""id"" value=""1""> 　<input type=""submit"" name=""Submit"" class=button value=""生成HTML页""> <b>说明：</b>多个ID请用“,”分开</td>"
	Response.Write " </tr></form>"
	Response.Write " <tr><form name=""Createform5"" method=""post"" action=""?action=cont"">"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write " <td class=tablerow1 height=""22"">生成全部内容HTML页（方法一）</td>"
	Response.Write " <td class=tablerow1><input type=""submit"" class=button name=""Submit"" value=""生成HTML页""> "
	Response.Write "　<input type=radio name=type checked value=""renew""> 生成已更新"
	Response.Write "　<input type=radio name=type value=""all""> 生成全部"
	Response.Write "</td></form>"
	Response.Write " </tr>"
	Response.Write " <tr><form name=""Createform6"" method=""post"" action=""?action=info"">"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write " <td class=tablerow1 height=""22"">生成全部内容HTML页（方法二）</td>"
	Response.Write " <td class=tablerow1><input type=""submit"" class=button name=""Submit"" value=""生成HTML页""> "
	Response.Write "　<input type=radio name=type checked value=""renew""> 生成已更新"
	Response.Write "　<input type=radio name=type value=""all""> 生成全部"
	Response.Write "</td></form>"
	Response.Write " </tr>"
	Response.Write " <tr><form name=""Createform7"" method=""post"" action=""?action=renewall"">"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write " <td class=tablerow1 height=""22"">更新全部文章</td>"
	Response.Write " <td class=tablerow1> <input type=""submit"" class=button name=""Submit"" value=""更新全部文章""> "
	Response.Write "</td></form>"
	Response.Write " </tr>"
	Response.Write "<tr>"
	Response.Write "	<td colspan=2 class=tablerow2>②、<font color=blue>按分类生成" & sModuleName & "相关HTML页</font></td>"
	Response.Write "</tr>"
	Response.Write "<form name=Createform8 method=post action=Admin_Createyemian.Asp>" & vbNewLine
	Response.Write "<input type=hidden name=action value='list'>"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "<input type=hidden name=field value='2'>"
	Response.Write "<tr>"
	Response.Write " <td class=tablerow1>"
	Response.Write "<select name='classid' size='2' multiple style='height:260px;width:200px;'>"
	Dim strSelectClass,re
	strSelectClass = enchiasp.LoadSelectClass(ChannelID)
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	Re.Pattern="(value=)(.*)("" )"
	strSelectClass = Re.Replace(strSelectClass,"")
	Re.Pattern="({ClassID=)(.*)(}>)"
	strSelectClass = Re.Replace(strSelectClass,"value=""$2"">")
	Response.Write strSelectClass
	Set Re = Nothing
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write " <td class=tablerow1 valign=top>"
	Response.Write "<input class=Button type=submit name=Submit3 value=' 生成文章内容的HTML' onclick=""document.Createform8.action.value='cont';""><br><br>"
	Response.Write "请按 CTRL 键多选"
	Response.Write "</td></form>"
	Response.Write " </tr>"
	Response.Write "</table>"
End Sub
Function LinkedStr(str)
	LinkedStr = "Admin_Createyemian.Asp?action=" & str & "&amp;ChannelID="& ChannelID
End Function
Function ShowInnerHTML(totalnumber,n)
	Response.Write "<script>table2.style.width=" & Fix((n / totalnumber) * 400) & ";"
	Response.Write "txt2.innerHTML=""完成：" & FormatNumber(n / totalnumber * 100, 2, -1) & """;"
	Response.Write "table2.title=""(" & n & ")"";</script>" & vbNewLine
End Function
Sub ShowGuage()
	Response.Write "<br><table width='400' align=center border=0 cellspacing=1 cellpadding=1>" & vbCrLf
	Response.Write "<tr> " & vbCrLf
	Response.Write "<td style=""border: 1px #384780 solid ;background-color: #FFFFFF;"">" & vbCrLf
	Response.Write "<table width=0 id=table2 name=table2 border=0 cellspacing=0 cellpadding=0>" & vbCrLf
	Response.Write "<tr height=12><td bgcolor=#36D91A>" & vbCrLf
	Response.Write "</td></tr></table></td></tr><tr> " & vbCrLf
	Response.Write "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span> <span style=""font-size:9pt"">%</span>   <span id=txt3 name=txt3 style=""font-size:9pt""></span></td></tr>" & vbCrLf
	Response.Write "</table><br>" & vbCrLf
End Sub

Sub CreateSiteIndex()
	Dim HtmlContent,FilePath
	On Error Resume Next
	HtmlContent = HTML.ShowIndex(True)
	FilePath = "../" & enchiasp.IndexName
	enchiasp.CreatedTextFile FilePath,HtmlContent
	Response.Write "<li>生成首页HTML完成... <a href=" & FilePath & " target=_blank>"
	Response.Write Server.MapPath(FilePath)
	Response.Write "</a></li>"
End Sub
Sub Create_Article_List()
	Dim totalnumber
	If CInt(Request("field")) = 1 Then
		SQL = " And TurnLink=0 And isUpdate>0"
	ElseIf CInt(Request("field")) = 2 Then
		If Trim(Request("classid")) = "" Then
			ErrMsg = "<li>请选择分类~！</li>"
			FoundErr = True 
			Exit Sub
		End If
		SQL = " And TurnLink=0 And ClassID in (" & Request("classid") & ")"
	Else
		SQL = " And TurnLink=0"
	End If
	totalnumber = enchiasp.Execute("SELECT COUNT(ClassID) FROM ECCMS_Classify WHERE ChannelID=" & ChannelID & SQL & "")(0)
	Set Rs = enchiasp.Execute("SELECT ClassID FROM [ECCMS_Classify] WHERE ChannelID=" & ChannelID & SQL & " ORDER BY classid ASC")
	If Rs.BOF And Rs.EOF Then
		Response.Write "<meta http-equiv=""refresh"" content=""1;url='Admin_Createyemian.Asp?ChannelID=" & ChannelID & "&stype=1&D=" & startime & "'"">"
		Set Rs = Nothing
		Exit Sub
	End If
	Response.Flush
	Call ShowGuage
	SQL=Rs.GetRows(-1)
	enchicms.ShowFlush = 0
	Response.Write "<ol>"
	For i=0 To Ubound(SQL,2)
		If Not Response.IsClientConnected Then ResponseEnd
		enchicms.CreateArticleList SQL(0,i),1
		ShowInnerHTML totalnumber,i
		enchiasp.Execute ("UPDATE ECCMS_Classify SET isUpdate=0 WHERE Classid="& CLng(SQL(0,i)))
		Response.Flush
	Next
	Response.Write "</ol>"
	If i>=totalnumber Then
		Response.Write "<script>table2.style.width=400;txt2.innerHTML=""完成：100"";</script>"
	End If
	Rs.Close:Set Rs = Nothing
	
	Response.Write "<meta http-equiv=""refresh"" content=""2;url='Admin_Createyemian.Asp?ChannelID=" & ChannelID & "&D=" & startime & "&stype=1'"">"
End Sub
Sub Create_Article_Content()
	Dim strDate,d,strLink,j
	Dim CurrentPage,page_count,totalnumber,Pcount,maxperpage
	If Trim(Request("type")) = "date" Then
		strDate = Request("datYear") & "-" & Request("datMonth") & "-" & Request("datDay")
		strLink = "type=date&datYear=" & Request("datYear") & "&datMonth=" & Request("datMonth") & "&datDay=" & Request("datDay")
		If isSqlDataBase = 1 Then
			SQL = "WHERE ChannelID=" & ChannelID & " And isAccept>0 And Datediff(d,WriteTime,getdate())=" & DateDiff("d", strDate, Now()) & ""
		Else
			SQL = "WHERE ChannelID=" & ChannelID & " And isAccept>0 And WriteTime like '%" & strDate & "%'"
		End If
	End If
	If Trim(Request("type")) = "new" Then
		If Not IsNull(Session("PageCount")) Then Session("PageCount") = 0
		strLink = "type=new&newnum="& Request("newnum")
		If CLng(Request("newnum")) > 0 Then
			SQL = "WHERE ChannelID=" & ChannelID & " And isAccept>0 "
		Else
			Response.Write "<meta http-equiv=""refresh"" content=""0;url='" & Request.ServerVariables("HTTP_REFERER") & "'"">"
		End If
	End If
	If Trim(Request("type")) = "all" Then
		strLink = "type=all"
		SQL = "WHERE ChannelID=" & ChannelID & " And isAccept>0"
	End If
	If Trim(Request("type")) = "renew" Then
		strLink = "type=renew"
		SQL = "WHERE ChannelID=" & ChannelID & " And isAccept>0 And isUpdate=1"
	End If
	If Trim(Request("type")) = "seid" Then
		strLink = "type=seid&id=" & Request("id")
		SQL = "WHERE ChannelID=" & ChannelID & " And isAccept>0 And ArticleID in (" & Request("id") & ")"
	End If
	If CInt(Request("field")) = 2 Then
		If Trim(Request("classid")) = "" Then
			ErrMsg = "<li>请选择分类~！</li>"
			FoundErr = True 
			Exit Sub
		End If
		strLink = "field=2&classid=" & Request("classid")
		SQL = "WHERE ChannelID=" & ChannelID & " And isAccept>0 And ClassID in (" & Request("classid") & ")"
	End If
	If Trim(Request("p")) <> "" Then
		CurrentPage = CLng(Request("p"))
		d = Trim(Request("D"))
	Else
		CurrentPage = 1
		d = startime
	End If
	maxperpage = 100   '每页显示列表数
	If CurrentPage = 0 Then CurrentPage = 1
	'得到总数
	If Trim(Request("type")) = "new" Then
		totalnumber = CLng(Request("newnum"))
		SQL = "SELECT TOP " & CLng(Request("newnum")) & " ArticleID FROM ECCMS_Article "& SQL &" ORDER BY ArticleID DESC"
	Else
		If Trim(Session("totalnumber")) = "" Then
			totalnumber = enchiasp.Execute("SELECT COUNT(ArticleID) FROM [ECCMS_Article] "& SQL &"")(0)
			Session("totalnumber") = totalnumber
			Session("PageCount") = 0
		Else
			totalnumber = Session("totalnumber")
		End If
		SQL = "SELECT ArticleID FROM ECCMS_Article "& SQL &" ORDER BY ArticleID DESC"
	End If
	Pcount = CLng(totalnumber / maxperpage)  '得到总页数
	If Pcount < totalnumber / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("adodb.recordset")
	If IsSqlDataBase = 1 Then
		Set Rs = Conn.Execute(SQL)
	Else
		Rs.Open SQL, Conn, 1,1
	End If
	page_count = 0
	Call ShowGuage
	Response.Flush
	If Trim(Request("type")) = "renew" Then CurrentPage=1
	If Not (Rs.EOF And Rs.BOF) Then
		Rs.MoveFirst
		If Pcount > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		j = (CurrentPage - 1) * maxperpage + 1
		If Rs.Eof Then Exit Sub
		SQL=Rs.GetRows(maxperpage)
		Set Rs = Nothing
		enchicms.ShowFlush = 0
		Response.Write "<ol>"
		For i=0 To Ubound(SQL,2)
			If Not Response.IsClientConnected Then ResponseEnd
			Session("PageCount") = CLng(Session("PageCount")) + 1
			enchicms.CreateArticleContent(CLng(SQL(0,i)))
			Response.Flush
			Response.Write "<script>"
			Response.Write "table2.style.width=" & Fix((i / maxperpage) * 400) & ";"
			Response.Write "txt2.innerHTML=""完成：" & FormatNumber(i / maxperpage * 100, 2, -1) & """;"
			Response.Write "txt3.innerHTML=""　　共有文章 <B><font color=RED>" & totalnumber & "</font></B> 篇 　正在生成第 " & Session("PageCount") & " 篇 "";" & vbCrLf
			Response.Write "table2.title=""(" & i & ")"";</script>" & vbNewLine
			enchiasp.Execute ("UPDATE ECCMS_Article SET isUpdate=0 WHERE ArticleID="& CLng(SQL(0,i)))
			Response.Flush
			j = j + 1
		Next
		Response.Write "</ol>"
		SQL=Null
	Else
		Response.Write "<script>table2.style.width=400;txt2.innerHTML=""完成：100"";</script>"
		Response.Write "<meta http-equiv=""refresh"" content=""0;url='Admin_Createyemian.Asp?ChannelID=" & ChannelID & "&stype=2&num=" & totalnumber & "&D=" & d & "&p=" & CurrentPage & "'"">"
		Exit Sub
	End If
	If IsObject(Rs) Then Set Rs = Nothing
	If CurrentPage => Pcount Then
		Response.Write "<script>table2.style.width=400;txt2.innerHTML=""完成：100"";</script>"
		Response.Write "<meta http-equiv=""refresh"" content=""1;url='Admin_Createyemian.Asp?ChannelID=" & ChannelID & "&stype=2&num=" & totalnumber & "&D=" & d & "&p=" & CurrentPage & "'"">"
	Else
		Response.Write "<script>table2.style.width=400;txt2.innerHTML=""完成：100"";</script>"
		Response.Write "<script language='JavaScript'>" & vbNewLine
		Response.Write "function build(){window.location.href='Admin_Createyemian.Asp?action=cont&ChannelID=" & ChannelID & "&" & strLink & "&D=" & d & "&p=" & CurrentPage + 1 & "';}" & vbNewLine
		Response.Write "setTimeout('build()',1000);" & vbNewLine
		Response.Write "</script>" & vbNewLine
	End If
End Sub
Sub RenewAll_Article()
	enchiasp.Execute ("UPDATE ECCMS_Article SET isUpdate=1 WHERE ChannelID=" & ChannelID)
	Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Sub Create_Article_Info()
	Dim totalnumber,ArticleID
	dim a(),s,d,num
	If Trim(Request("num")) = "" Then
		d = startime
		If Trim(Request("type")) = "renew" Then
			SQL = "WHERE ChannelID=" & ChannelID & " And isAccept>0 And isUpdate=1"
		Else
			SQL = "WHERE ChannelID=" & ChannelID & " And isAccept>0 "
		End If
		totalnumber = enchiasp.Execute("SELECT COUNT(ArticleID) FROM [ECCMS_Article] " & SQL & "")(0)
		SQL = "SELECT ArticleID FROM ECCMS_Article " & SQL & " ORDER BY ArticleID"
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			Session("BuildArticleID")=""
			Set Rs = Nothing
			Response.Write "<meta http-equiv=""refresh"" content=""1;url='Admin_Createyemian.Asp?ChannelID=" & ChannelID & "&stype=2&D=" & d & "&num=0'"">"
			Exit Sub
		End If
		SQL = Rs.GetRows(-1)
		For i=0 To Ubound(SQL,2)
			ReDim Preserve a(i)
			a(i)=SQL(0,i)
		Next
		Set Rs = Nothing
		s=Join(a,"|")
		Session("BuildArticleID") = Split(s, "|")
		Response.Write "<script language='JavaScript'>" & vbNewLine
		Response.Write "function build(){window.location.href='Admin_Createyemian.Asp?action=info&ChannelID=" & ChannelID & "&total=" & totalnumber & "&D=" & d & "&num=0';}" & vbNewLine
		Response.Write "setTimeout('build()',1000);" & vbNewLine
		Response.Write "</script>" & vbNewLine
	Else
		totalnumber = CLng(Request("total"))
		num = CLng(Request("num"))
		d = Request("d")

		Response.Write "<br><br>" & vbNewLine
		Response.Write "<table width='400' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbNewLine
		Response.Write "  <tr>" & vbNewLine
		Response.Write "    <td height='50'>总共需要生成 <font color='blue'><b>" & totalnumber & "</b></font> 个页面，正在生成第 <font color='red'><b>" & num & "</b></font>  个页面……</td>" & vbNewLine
		Response.Write "  </tr>" & vbNewLine
		Response.Write "  <tr>" & vbNewLine
		Response.Write "    <td><table width='100%' border='0' cellpadding='1' cellspacing='1'>" & vbNewLine
		Response.Write "      <tr>" & vbNewLine
		Response.Write "        <td style=""border: 1px #384780 solid ;background-color: #FFFFFF;""><table width='" & Fix((num / totalnumber) * 400) & "' height='12' border='0' cellpadding='0' cellspacing='0' bgcolor=#36D91A><tr><td></td></tr></table></td>" & vbNewLine
		Response.Write "      </tr>" & vbNewLine
		Response.Write "    </table></td>" & vbNewLine
		Response.Write "  </tr>" & vbNewLine
		Response.Write "  <tr>" & vbNewLine
		Response.Write "    <td align='center'>" & FormatNumber(num / totalnumber * 100, 2, -1) & " %</td>" & vbNewLine
		Response.Write "  </tr>" & vbNewLine
		Response.Write "</table>" & vbNewLine
		Response.Write "<table width='400' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbNewLine
		Response.Write "   <tr><td height='30' align='center'><input type='button' name='stop' value=' 停止生成HTML ' onclick=""window.location.href='Admin_Createyemian.Asp?ChannelID=" & ChannelID & "&stype=2&D=" & d & "&num=" & num & "';"" class=button></td></tr>" & vbNewLine
		Response.Write "</table>" & vbNewLine

		If num > totalnumber Or num > CLng(UBound(Session("BuildArticleID"))) Then
			Session("BuildArticleID")=""
			Response.Write "<meta http-equiv=""refresh"" content=""1;url='Admin_Createyemian.Asp?ChannelID=" & ChannelID & "&stype=2&D=" & d & "&num=" & num & "'"">"
			Exit Sub
		Else
			ArticleID = Session("BuildArticleID")(num)
			enchicms.CreateArticleContent(CLng(ArticleID))
			enchiasp.Execute ("UPDATE ECCMS_Article SET isUpdate=0 WHERE ArticleID="& CLng(ArticleID))
			Response.Write "<script language='JavaScript'>" & vbNewLine
			Response.Write "function build(){window.location.href='Admin_Createyemian.Asp?action=info&ChannelID=" & ChannelID & "&total=" & totalnumber & "&D=" & d & "&num=" & num+1 & "';}" & vbNewLine
			Response.Write "setTimeout('build()',1000);" & vbNewLine
			Response.Write "</script>" & vbNewLine
		End If
	End If
End Sub
%>
