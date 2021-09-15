<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
'=====================================================================
' 软件名称：恩池网站管理系统--评论管理
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim Action
If Not ChkAdmin("Comment" & ChannelID) Then
	Server.Transfer("showerr.asp")
	Response.End
End If

Action = LCase(Request("action"))
Select Case Trim(Action)
Case "del"
	Call DeleteComment
Case "delall"
	Call DelAllComment
Case Else
	Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Sub showmain()
	Dim CurrentPage,page_count,totalnumber,Pcount,maxperpage
	Dim strTopic
	maxperpage = 20
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CLng(CurrentPage)
	End If
	If CLng(CurrentPage) = 0 Then CurrentPage = 1
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th width='5%'>选择</th>"
	Response.Write "		<th width='50%'>评论主题</th>"
	Response.Write "		<th width='16%'>用户名称</th>"
	Response.Write "		<th width='5%'>打分</th>"
	Response.Write "		<th width='12%'>评论时间</th>"
	Response.Write "		<th width='12%'>用户IP</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=selform method=post action='admin_comment.asp'>"
	Response.Write "	<input type=hidden name=action value=""del"">"
	Response.Write "	<input type=hidden name=ChannelID value="""& ChannelID &""">"
	totalnumber = enchiasp.Execute("SELECT COUNT(commentid) FROM ECCMS_Comment WHERE ChannelID=" & ChannelID)(0)
	Pcount = CLng(totalnumber / maxperpage)  '得到总页数
	If Pcount < totalnumber / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Select Case CInt(enchiasp.modules)
	Case 1
		SQL = "SELECT C.commentid,C.postid,C.content,C.Grade,C.username,C.postime,C.postip,A.ArticleID,A.title FROM [ECCMS_Comment] C INNER JOIN [ECCMS_Article] A ON C.Postid=A.ArticleID WHERE C.ChannelID=" & ChannelID & " ORDER BY C.Postime DESC"
	Case 2
		SQL = "SELECT C.commentid,C.postid,C.content,C.Grade,C.username,C.postime,C.postip,A.softid,A.SoftName,A.SoftVer FROM [ECCMS_Comment] C INNER JOIN [ECCMS_SoftList] A ON C.Postid=A.softid WHERE C.ChannelID=" & ChannelID & " ORDER BY C.Postime DESC"
	Case 3
		SQL = "SELECT C.commentid,C.postid,C.content,C.Grade,C.username,C.postime,C.postip,A.shopid,A.TradeName FROM [ECCMS_Comment] C INNER JOIN [ECCMS_ShopList] A ON C.Postid=A.shopid WHERE C.ChannelID=" & ChannelID & " ORDER BY C.Postime DESC"
	Case 5
		SQL = "SELECT C.commentid,C.postid,C.content,C.Grade,C.username,C.postime,C.postip,A.flashid,A.title FROM [ECCMS_Comment] C INNER JOIN [ECCMS_FlashList] A ON C.Postid=A.flashid WHERE C.ChannelID=" & ChannelID & " ORDER BY C.Postime DESC"
	Case Else
		ErrMsg = "<li>错误的系统参数~!</li>"
		FoundErr = True
		Exit Sub
	End Select
	Set Rs = Server.CreateObject("ADODB.Recordset")
	If IsSqlDataBase=1 Then
		Set Rs = enchiasp.Execute(SQL)
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=6 class=TableRow1>没有" & sModuleName & "评论！</td></tr>"
	Else
		Rs.MoveFirst
		If Pcount > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		page_count = 0
		Do While Not Rs.EOF And page_count < CInt(maxperpage)
			If Not Response.IsClientConnected Then ResponseEnd
			Select Case CInt(enchiasp.modules)
			Case 1
				strTopic = "<a href=""../" & enchiasp.ChannelDir & "Comment.Asp?ArticleID="& Rs(7) &""" title='点击查看此" & sModuleName & "评论' target=_blank>"& Rs(8) &"</a>"
			Case 2
				strTopic = "<a href=""../" & enchiasp.ChannelDir & "Comment.Asp?softid="& Rs(7) &""" title='点击查看此" & sModuleName & "评论' target=_blank>"& Rs(8) &" "& Rs(9) &"</a>"
			Case 3
				strTopic = "<a href=""../" & enchiasp.ChannelDir & "Comment.Asp?shopid="& Rs(7) &""" title='点击查看此" & sModuleName & "评论' target=_blank>"& Rs(8) &"</a>"
			Case 5
				strTopic = "<a href=""../" & enchiasp.ChannelDir & "Comment.Asp?flashid="& Rs(7) &""" title='点击查看此" & sModuleName & "评论' target=_blank>"& Rs(8) &"</a>"
			End Select
			
			Admin_Comment_list Rs(0),strTopic,Rs(1),Rs(2),Rs(3),Rs(4),Rs(5),Rs(6)
			Rs.movenext
			page_count = page_count + 1
			If page_count >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=6>"
	Response.Write "<input class=Button type=""button"" name=""chkall"" value=""全选"" onClick=""CheckAll(this.form)""><input class=Button type=""button"" name=""chksel"" value=""反选"" onClick=""ContraSel(this.form)"">"
	Response.Write "<input type=submit name=submit2 value=""删除评论"" onclick=""return confirm('您确定要删除此评论吗?')"" class=Button>"
	Response.Write "<input type=submit name=submit3 value="" 全部删除 "" onclick=""document.selform.action.value='delall';return confirm('您确定要删除所以评论吗?')"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow2 colspan=6>"
	Response.Write showpages(CurrentPage,Pcount,totalnumber,maxperpage,"&ChannelID="& ChannelID)
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
End Sub
Function Admin_Comment_list(commentid,topic,postid,content,Grade,username,postime,postip)
	Response.Write "	<tr>"
	Response.Write "		<td class=TableRow2 align=center><input type=checkbox name=commentid value="""& commentid &"""></td>"
	Response.Write "		<td class=TableRow2>" & topic & "</td>"
	Response.Write "		<td class=TableRow2 align=center><font color=blue>" & username & "</font></td>"
	Response.Write "		<td class=TableRow2 align=center><font color=red>" & Grade & "</font></td>"
	Response.Write "		<td class=TableRow2 align=center>" & enchiasp.FormatDate(postime,2) & "</td>"
	Response.Write "		<td class=TableRow2 align=center>" & postip & "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=TableRow1 colspan=6>"& enchiasp.CutString(content,100) &"</td>"
	Response.Write "	</tr>"
End Function

Sub DeleteComment()
	If Trim(Request("commentid")) <> "" Then
		enchiasp.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And CommentID in (" & Request("CommentID") & ")")
		Response.Redirect  Request.ServerVariables("HTTP_REFERER")
	Else
		ErrMsg = "<li>错误的系统参数,请选择要删除的评论ID</li>"
		FoundErr = True
		Exit Sub
	End If
End Sub
Sub DelAllComment()
	enchiasp.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID)
	Response.Redirect  Request.ServerVariables("HTTP_REFERER")
End Sub
%>