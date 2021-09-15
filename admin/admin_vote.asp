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
Dim CurrentPage,maxperpage,totalnumber,Pcount,totalPut
Dim isEdit,selVoteid,VoteTitle,i,Action
Action = LCase(Request("action"))
If Not ChkAdmin("Vote") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Select Case Trim(Action)
	Case "save"
		Call SaveVote
	Case "modify"
		Call ModifyVote
	Case "add"
		isEdit = False
		EditVote
	Case "edit"
		isEdit = True
		EditVote
	Case Else
		Call VoteMain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub VoteMain()
	Dim bookmark
	If Not IsEmpty(Request("selVoteid")) Then
		selVoteid = Request("selVoteid")
		Select Case Request("type")
			Case "del"
				Call DelVote
			Case "isLock"
				Call isLock
			Case "noLock"
				Call noLock
			Case "radio"
				Call setRadio
			Case "checkbox"
				Call setCheckbox
			Case Else
				Response.Write "无效参数！"
				Response.End
		End Select
	End If
	Response.Write "<TABLE width=""99%"" border=0 cellpadding=3 cellspacing=1 align=center class=tableBorder>" & vbNewLine
	Response.Write "<TR>" & vbNewLine
	Response.Write " <TH colspan=6>投票管理</TH>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""28"">" & vbNewLine
	Response.Write " <TD colspan=6 class=TableRow1>投票调用方法:<BR>①&lt;script src=""vote/showvote.js""&gt;&lt;/script&gt;<BR>" & vbNewLine
	Response.Write " ②&lt;IFRAME name=vote src=""vote/vote.htm"" frameBorder=no scrolling=no width=180 height=220&gt;&lt;/IFRAME&gt;<BR>" & vbNewLine
	Response.Write " </TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR>" & vbNewLine
	Response.Write " <TH noWrap>选 择</TH>" & vbNewLine
	Response.Write " <TH noWrap>投票主题　[<a href=admin_vote.asp?action=add Class=showtitle>添加投票</a>]</TH>" & vbNewLine
	Response.Write " <TH noWrap>投 票 数</TH>" & vbNewLine
	Response.Write " <TH noWrap>编辑投票</TH>" & vbNewLine
	Response.Write " <TH noWrap>投票类型</TH>" & vbNewLine
	Response.Write " <TH noWrap>状 态</TH>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	If Not IsEmpty(Request("page")) Then
		CurrentPage = CLng(Request("page"))
	Else
		CurrentPage = 1
	End If
	maxperpage = 20 '###每页显示数
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Vote order by id desc"
	Rs.Open SQL, Conn, 1, 1
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td colspan=10 class=TableRow1>没有投票主题！</td></tr>"
	Else
		totalnumber = Rs.recordcount
		If (totalnumber Mod maxperpage) = 0 Then
			Pcount = totalnumber \ maxperpage
		Else
			Pcount = totalnumber \ maxperpage + 1
		End If
		Rs.MoveFiRst
		If CurrentPage > Pcount Then CurrentPage = Pcount
		If CurrentPage < 1 Then CurrentPage = 1
		Rs.Move (CurrentPage - 1) * maxperpage
		bookmark = Rs.bookmark
		i = 0
		Response.Write "<TR height=""28"">" & vbNewLine
		Response.Write " <TD colspan=6 class=TableRow2 align=center>"
		Call showpage
		Response.Write "</TD>" & vbNewLine
		Response.Write "</TR>" & vbNewLine
		Response.Write "<form name=""selform"" method=""post"" action="""">" & vbNewLine
		Do While Not Rs.EOF And i < CLng(maxperpage)
			Response.Write "<TR>" & vbNewLine
			Response.Write " <TD noWrap class=TableRow1 align=center><input type=checkbox name=selVoteid value="""
			Response.Write Rs("id")
			Response.Write """></TD>" & vbNewLine
			Response.Write " <TD noWrap class=TableRow1 width=""70%""><a href=../vote/vote.htm title=""查看投票"" target=_blank>"
			Response.Write Rs("topic")
			Response.Write "</a></TD>" & vbNewLine
			Response.Write " <TD noWrap class=TableRow1 align=center><FONT COLOR=RED><B>"
			Response.Write Rs("VoteNum")
			Response.Write "</B></FONT></TD>" & vbNewLine
			Response.Write " <TD noWrap class=TableRow1 align=center><a href=admin_vote.asp?action=edit&id="
			Response.Write Rs("id")
			Response.Write " title=""查看编辑:"
			Response.Write Rs("topic")
			Response.Write """>编辑投票</a></TD>" & vbNewLine
			Response.Write " <TD noWrap class=TableRow1 align=center>"
			If CInt(Rs("VoteType")) = 0 Then
				Response.Write "单选"
			Else
				Response.Write "多选"
			End If
			Response.Write "</TD>" & vbNewLine
			Response.Write " <TD noWrap class=TableRow1 align=center>"
			If CInt(Rs("isLock")) = 0 Then
				Response.Write "正常"
			Else
				Response.Write "<FONT COLOR=RED>锁定</FONT>"
			End If
			Response.Write "</TD>" & vbNewLine
			Response.Write "</TR>" & vbNewLine
			Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Write "<TR height=""30"">" & vbNewLine
	Response.Write " <TD class=TableRow1>管理</TD>" & vbNewLine
	Response.Write " <TD colspan=5 class=TableRow1><input class=Button type=button name=chkall value='全选' onClick=""CheckAll(this.form)""><input class=Button type=button name=chksel value='反选' onClick=""ContraSel(this.form)""> " & vbNewLine
	Response.Write " <input type=""radio"" name=""type"" value=""del"" title=""管理选项：批量删除选中的新闻"">批量删除 " & vbNewLine
	Response.Write " <input type=""radio"" name=""type"" value=""isLock"" title=""管理选项：批量锁定投票主题"">锁定 " & vbNewLine
	Response.Write " <input type=""radio"" name=""type"" value=""noLock"" title=""管理选项：批量解除锁定"">解锁 " & vbNewLine
	Response.Write " <input type=""radio"" name=""type"" value=""radio"" title=""管理选项：批量设置单选投票"">单选 " & vbNewLine
	Response.Write " <input type=""radio"" name=""type"" value=""checkbox"" title=""管理选项：批量设置多选投票"">多选 " & vbNewLine
	Response.Write " <input type=submit name=Submit value=""执行操作"" class=button onclick=""{if(confirm('您确定执行此操作吗?')){this.document.selform.submit();return true;}return false;}""></TD>" & vbNewLine
	Response.Write "</TR></form>" & vbNewLine
	Response.Write "<TR height=""28"">" & vbNewLine
	Response.Write " <TD colspan=6 class=TableRow2 align=center>"
	Call showpage
	Response.Write "</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "</TABLE>" & vbNewLine
End Sub


Private Sub EditVote()
	If isEdit Then
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "select * from ECCMS_Vote where id=" & Request("id")
		Rs.Open SQL, Conn, 1, 1
		enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
		VoteTitle = "编辑投票"
	Else
		VoteTitle = "添加新的投票"
	End If
	Response.Write " <TABLE width=""99%"" border=0 cellpadding=3 cellspacing=1 align=center class=tableBorder>" & vbNewLine
	Response.Write "<TR><form name=""myform"" method=""post"" action=""admin_vote.asp"">" & vbNewLine
	Response.Write " <input type=""Hidden"" name=""action"" value='"
	If isEdit Then
		Response.Write "modify"
	Else
		Response.Write "save"
	End If
	Response.Write "'>" & vbNewLine
	Response.Write " "
	If isEdit Then
		Response.Write " <input type=""Hidden"" name=""id"" value='"
		Response.Write CStr(Request("id"))
		Response.Write "'>" & vbNewLine
		Response.Write " "
	End If
	Response.Write " <TH colspan=2>"
	Response.Write VoteTitle
	Response.Write "</TH>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>投票类型：</TD>" & vbNewLine
	Response.Write " <TD width=""85%"" class=TableRow1>" & vbNewLine
	Response.Write " <input type=""radio"" name=""VoteType"" value=""0"" title=""设置单选投票"" "
	If isEdit Then
		If CInt(Rs("VoteType")) = 0 Then
			Response.Write "checked"
		End If
	Else
		Response.Write "checked"
	End If
	Response.Write ">单选 " & vbNewLine
	Response.Write " <input type=""radio"" name=""VoteType"" value=""1"" title=""设置多选投票"" "
	If isEdit Then
		If CInt(Rs("VoteType")) = 1 Then
			Response.Write "checked"
		End If
	End If
	Response.Write ">多选</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>投票主题：</TD>" & vbNewLine
	Response.Write " <TD class=TableRow1><input type=text name=topic size=50 value="""
	If isEdit Then
		Response.Write Rs("topic")
	End If
	Response.Write """></TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>投票选项1：</TD>" & vbNewLine
	Response.Write " <TD class=TableRow1><input type=text name=Choose_1 size=20 value="""
	If isEdit Then
		Response.Write Rs("Choose_1")
	End If
	Response.Write """>&nbsp;&nbsp;&nbsp;" & vbNewLine
	Response.Write " "
	If isEdit Then
		Response.Write "投票数：<input type=text name=ChooseNum_1 size=10 value="""
		Response.Write Rs("ChooseNum_1")
		Response.Write """>"
	End If
	Response.Write "</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>投票选项2：</TD>" & vbNewLine
	Response.Write " <TD class=TableRow1><input type=text name=Choose_2 size=20 value="""
	If isEdit Then
		Response.Write Rs("Choose_2")
	End If
	Response.Write """>&nbsp;&nbsp;&nbsp;" & vbNewLine
	Response.Write " "
	If isEdit Then
		Response.Write "投票数：<input type=text name=ChooseNum_2 size=10 value="""
		Response.Write Rs("ChooseNum_2")
		Response.Write """>"
	End If
	Response.Write "</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>投票选项3：</TD>" & vbNewLine
	Response.Write " <TD class=TableRow1><input type=text name=Choose_3 size=20 value="""
	If isEdit Then
		Response.Write Rs("Choose_3")
	End If
	Response.Write """>&nbsp;&nbsp;&nbsp;" & vbNewLine
	Response.Write " "
	If isEdit Then
		Response.Write "投票数：<input type=text name=ChooseNum_3 size=10 value="""
		Response.Write Rs("ChooseNum_3")
		Response.Write """>"
	End If
	Response.Write "</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>投票选项4：</TD>" & vbNewLine
	Response.Write " <TD class=TableRow1><input type=text name=Choose_4 size=20 value="""
	If isEdit Then
		Response.Write Rs("Choose_4")
	End If
	Response.Write """>&nbsp;&nbsp;&nbsp;" & vbNewLine
	Response.Write " "
	If isEdit Then
		Response.Write "投票数：<input type=text name=ChooseNum_4 size=10 value="""
		Response.Write Rs("ChooseNum_4")
		Response.Write """>"
	End If
	Response.Write "</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>投票选项5：</TD>" & vbNewLine
	Response.Write " <TD class=TableRow1><input type=text name=Choose_5 size=20 value="""
	If isEdit Then
		Response.Write Rs("Choose_5")
	End If
	Response.Write """>&nbsp;&nbsp;&nbsp;" & vbNewLine
	Response.Write " "
	If isEdit Then
		Response.Write "投票数：<input type=text name=ChooseNum_5 size=10 value="""
		Response.Write Rs("ChooseNum_5")
		Response.Write """>"
	End If
	Response.Write "</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>背景颜色：</TD>" & vbNewLine
	Response.Write " <TD class=TableRow1><input type=text name=bgcolor size=10 value="""
	If isEdit Then
		Response.Write Rs("bgcolor")
	Else
		Response.Write "FFFFFF"
	End If
	Response.Write """>&nbsp;" & vbNewLine
	Response.Write " 如：FFFFFF 不用加&quot;<font color=""#FF3300"">#</font>&quot;</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>字体颜色：</TD>" & vbNewLine
	Response.Write " <TD class=TableRow1><input type=text name=FontColor size=10 value="""
	If isEdit Then
		Response.Write Rs("FontColor")
	Else
		Response.Write "000000"
	End If
	Response.Write """>&nbsp;" & vbNewLine
	Response.Write " 如：000000 不用加&quot;<font color=""#FF3300"">#</font>&quot;</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>字体大小：</TD>" & vbNewLine
	Response.Write " <TD class=TableRow1><input type=text name=FontSize size=10 value="""
	If isEdit Then
		Response.Write Rs("FontSize")
	Else
		Response.Write "12"
	End If
	Response.Write """>&nbsp;" & vbNewLine
	Response.Write " 单位px 如：这是12px <span style=""font-size:14px"">这是14px</span>,只输入数字</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>总投票数：</TD>" & vbNewLine
	Response.Write " <TD class=TableRow1><input type=text name=VoteNum size=10 value="""
	If isEdit Then
		Response.Write Rs("VoteNum")
	Else
		Response.Write "0"
	End If
	Response.Write """></TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR height=""22"">" & vbNewLine
	Response.Write " <TD noWrap align=""right"" class=TableRow2>是否锁定：</TD>" & vbNewLine
	Response.Write " <TD width=""85%"" class=TableRow1>" & vbNewLine
	Response.Write " <input type=""radio"" name=""isLock"" value=""0"" title=""设置单选投票"" "
	If isEdit Then
		If CInt(Rs("isLock")) = 0 Then
			Response.Write "checked"
		End If
	Else
		Response.Write "checked"
	End If
	Response.Write ">否 " & vbNewLine
	Response.Write " <input type=""radio"" name=""isLock"" value=""1"" title=""设置多选投票"" "
	If isEdit Then
		If CInt(Rs("isLock")) = 1 Then
			Response.Write "checked"
		End If
	End If
	Response.Write ">是</TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write " <tr>" & vbNewLine
	Response.Write " <td height=""22"" align=""right"" class=""TableRow2"">&nbsp;</td>" & vbNewLine
	Response.Write " <td align=""center"" class=""TableRow1"">" & vbNewLine
	Response.Write " <input type=""button"" name=""Submit1"" onclick=""javascript:history.go(-1)"" value=""返回上一页"" class=button>　" & vbNewLine
	Response.Write " <input type=reset name=Submit2 class=button value=""清 除"">　" & vbNewLine
	Response.Write "　<input type=Submit class=button value=""保存投票"" name=Submit>　" & vbNewLine
	Response.Write " </td>" & vbNewLine
	Response.Write " </tr></form>" & vbNewLine
	Response.Write "</table>" & vbNewLine
	If isEdit Then
		Rs.Close
		Set Rs = Nothing
	End If
End Sub


Private Sub SaveVote()
	'保存新的投票
	If Trim(Request.Form("topic")) = "" Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>投票主题不能为空！</li>"
		Exit Sub
	End If
	If Founderr = False Then
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "select * from ECCMS_Vote where (id is null)"
		Rs.Open SQL, Conn, 1, 3
		Rs.addnew
		Rs("Topic") = Request.Form("Topic")
		Rs("Choose_1") = Request.Form("Choose_1")
		Rs("Choose_2") = Request.Form("Choose_2")
		Rs("Choose_3") = Request.Form("Choose_3")
		Rs("Choose_4") = Request.Form("Choose_4")
		Rs("Choose_5") = Request.Form("Choose_5")
		Rs("ChooseNum_1") = 0
		Rs("ChooseNum_2") = 0
		Rs("ChooseNum_3") = 0
		Rs("ChooseNum_4") = 0
		Rs("ChooseNum_5") = 0
		Rs("isLock") = Request.Form("isLock")
		Rs("bgcolor") = Request.Form("bgcolor")
		Rs("FontColor") = Request.Form("FontColor")
		Rs("FontSize") = Request.Form("FontSize")
		Rs("VoteTime") = Now
		Rs("VoteNum") = 0
		Rs("VoteType") = Request.Form("VoteType")
		Rs("ChannelID") = 0
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Succeed ("<li>添加" & Request.Form("topic") & "成功!</li>")
	End If
End Sub


Private Sub ModifyVote()
	'修改投票
	If Trim(Request.Form("topic")) = "" Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>投票主题不能为空！</li>"
		Exit Sub
	End If
	If Founderr = False Then
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "select * from ECCMS_Vote where id=" & Request.Form("id")
		Rs.Open SQL, Conn, 1, 3
		Rs("Topic") = Request.Form("Topic")
		Rs("Choose_1") = Request.Form("Choose_1")
		Rs("Choose_2") = Request.Form("Choose_2")
		Rs("Choose_3") = Request.Form("Choose_3")
		Rs("Choose_4") = Request.Form("Choose_4")
		Rs("Choose_5") = Request.Form("Choose_5")
		Rs("ChooseNum_1") = Request.Form("ChooseNum_1")
		Rs("ChooseNum_2") = Request.Form("ChooseNum_2")
		Rs("ChooseNum_3") = Request.Form("ChooseNum_3")
		Rs("ChooseNum_4") = Request.Form("ChooseNum_4")
		Rs("ChooseNum_5") = Request.Form("ChooseNum_5")
		Rs("isLock") = Request.Form("isLock")
		Rs("bgcolor") = Request.Form("bgcolor")
		Rs("FontColor") = Request.Form("FontColor")
		Rs("FontSize") = Request.Form("FontSize")
		Rs("VoteNum") = Request.Form("VoteNum")
		Rs("VoteType") = Request.Form("VoteType")
		'Rs("ChannelID") = 0
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Succeed ("<li>修改" & Request.Form("topic") & "成功!</li>")
	End If
End Sub


Private Sub DelVote()
	'删除投票
	enchiasp.Execute ("delete from ECCMS_Vote where id in (" & selVoteid & ")")
End Sub


Private Sub isLock()
	'锁定投票
	enchiasp.Execute ("update ECCMS_Vote set isLock=1 where id in (" & selVoteid & ")")
End Sub ' islock


Private Sub noLock()
	'解除锁定
	enchiasp.Execute ("update ECCMS_Vote set isLock=0 where id in (" & selVoteid & ")")
End Sub


Private Sub setRadio()
	'设置单选投票
	enchiasp.Execute ("update ECCMS_Vote set VoteType=0 where id in (" & selVoteid & ")")
End Sub


Private Sub setCheckbox()
	'设置多选投票
	enchiasp.Execute ("update ECCMS_Vote set VoteType=1 where id in (" & selVoteid & ")")
End Sub


Private Sub showpage()
	Dim FileName
	Dim n
	Dim ii
	' 分页
	FileName = "admin_vote.asp"
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action=" & FileName & "><tr><td align=center> " & vbCrLf
	If CurrentPage < 2 Then
		Response.Write "投票主题 <font COLOR=#FF0000><B>" & totalnumber & "</B></font>&nbsp;首 页&nbsp;上一页&nbsp;"
	Else
		Response.Write "投票主题 <font COLOR=#FF0000><B>" & totalnumber & "</B></font>&nbsp;<a href=" & FileName & "?page=1>首 页</a>&nbsp;"
		Response.Write "<a href=" & FileName & "?page=" & CurrentPage - 1 & ">上一页</a>&nbsp;"
	End If
	If n - CurrentPage < 1 Then
		Response.Write "下一页&nbsp;尾 页 " & vbCrLf
	Else
		Response.Write "<a href=" & FileName & "?page=" & (CurrentPage + 1) & ">下一页</a>"
		Response.Write "&nbsp;<a href=" & FileName & "?page=" & n & ">尾 页</a>" & vbCrLf
	End If
	Response.Write "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
	Response.Write "&nbsp;转到："
	Response.Write "&nbsp;<select name='page' size='1' style=""font-size: 9pt"" onChange='javascript:submit()'>" & vbCrLf
	For ii = 1 To n
		Response.Write "<option value='" & ii & "' "
		If CurrentPage = Int(ii) Then
			Response.Write "selected "
		End If
		Response.Write ">第" & ii & "页</option>"
	Next
	Response.Write "&nbsp;</select> " & vbCrLf
	Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub
%>
