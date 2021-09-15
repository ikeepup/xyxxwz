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
Dim Rsm,Action,i
Dim ModuleName,MouseStyle,sChannelID

Response.Write "<script language=JavaScript>" & vbCrLf
Response.Write "function Juge(form1)" & vbCrLf
Response.Write "{" & vbCrLf
Response.Write " if (form1.page_name.value == """")" & vbCrLf
Response.Write " {" & vbCrLf
Response.Write "  alert(""请输入模板名称!"");" & vbCrLf
Response.Write "  form1.page_name.focus();" & vbCrLf
Response.Write "  return (false);" & vbCrLf
Response.Write " }" & vbCrLf
Response.Write " if (form1.TemplateDir.value == """")" & vbCrLf
Response.Write " {" & vbCrLf
Response.Write "  alert(""请输入模板目录!"");" & vbCrLf
Response.Write "  form1.TemplateDir.focus();" & vbCrLf
Response.Write "  return (false);" & vbCrLf
Response.Write " }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
Response.Write " <tr>"
Response.Write "   <th colspan=""2""></th>"
Response.Write " </tr>"
Response.Write " <tr>"
Response.Write "   <td colspan=""2"" class=""TableRow1""><strong>注意：</strong><br>"
Response.Write " ①在这里，您可以新建和修改模板，可以编辑CSS样式，可以新建模板页面；<br>"
Response.Write " ②当前正在使用的默认模板不能删除；<br>"
Response.Write " ③如果你想为每个页面设计不同的模板，请在相应的<span class=style2>模板基本设置</span>取消使用站点通栏。</td>"
Response.Write " </tr>"
Response.Write " <tr>"
Response.Write "   <td width=""10%"" nowrap class=""TableRow2"">管理选项：</td>"
Response.Write "   <td width=""90%"" class=""TableRow2"">"
Response.Write "<a href=admin_template.asp class=showmeun>模板管理首页</a> | "
Set Rsm = enchiasp.Execute("Select ChannelID,ModuleName From ECCMS_Channel where ChannelType < 2 And ChannelID <> 4 And stopChannel=0 Order By ChannelID Asc")
Do While Not Rsm.EOF
	Response.Write "<a href=?action=manage&ChannelID="
	Response.Write Rsm("ChannelID")
	Response.Write " class=showmeun>"
	Response.Write Rsm("ModuleName")
	Response.Write "模板管理</a> | "
	sModuleName = sModuleName & Rsm("ModuleName") & "|||"
	sChannelID = sChannelID & Rsm("ChannelID") & "|||"
	Rsm.MoveNext
Loop
Set Rsm = Nothing
Response.Write "<a href=?action=manage&ChannelID=9999 class=showmeun>公共模板管理</a> | "
Response.Write "<a href=admin_loadskin.asp class=showmeun>模板导出</a> | "
Response.Write "<a href=admin_loadskin.asp?action=load class=showmeun>模板导入</a>"
Response.Write "</td>"
Response.Write " </tr>"
Response.Write "</table>"
Response.Write "<br>"
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))

If ChannelID > 0 Then
	Set Rsm = enchiasp.Execute("SELECT ChannelID,ModuleName FROM ECCMS_Channel WHERE ChannelType=0 And ChannelID<>9999 And ChannelID=" & ChannelID)
	If Rsm.BOF And Rsm.EOF Then
		ModuleName = "全部"
	Else
		ModuleName = Rsm("ModuleName")
	End If
	Set Rsm = Nothing
Else
	ModuleName = "全部"
End If
MouseStyle = " bgcolor=""#EEEEE6"" onmouseover=""this.style.backgroundColor='#FFFF00';this.style.color='red'"" onmouseout=""this.style.backgroundColor='';this.style.color=''"""
Action = LCase(Request("action"))
If Not ChkAdmin("Template") Then
	Server.Transfer ("showerr.asp")
	Response.End
End If
Select Case Trim(Action)
	Case "del"
		Call DelTemplate
	Case "newtemplate"
		Call NewTemplate
	Case "default"
		Call DefaultTemplate
	Case "editstyle"
		Call EditStyle
	Case "savestyle"
		Call SaveStyle
	Case "set"
		Call SettingTemplate
	Case "saveset"
		Call SaveTemplateSet
	Case "help"
		Call EditTemplateHelp
	Case "savehelp"
		Call SaveTemplateHelp
	Case "manage"
		Call ChannelTemplate
	Case "edit"
		Call EditTemplatePage
	Case "save"
		Call SaveTemplatePage
	Case "newpage"
		Call NewTemplatePage
	Case Else
		Call MainPage
End Select
If FoundErr = True Then
	ReturnError (ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Sub MainPage()
	SQL = "select * from [ECCMS_Template] where ChannelID = 0 And pageid = 0 order by skinid asc"
	Set Rs = enchiasp.Execute(SQL)
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write " <tr>"
	Response.Write "   <th>模板名称</th>"
	Response.Write "   <th>编辑CSS样式</th>"
	Response.Write "   <th>模板常规设置</th>"
	Response.Write "   <th>编辑通栏模板</th>"
	Response.Write "   <th>操作选项</th>"
	Response.Write " </tr>"
	Do While Not Rs.EOF
		Response.Write " <tr "
		Response.Write MouseStyle
		Response.Write ">"
		Response.Write "   <td align=""center"">"
		If Rs("IsDefault") = 1 Then
			Response.Write "<img src=images/arrow.gif> "
			Response.Write "<a href=?action=manage&skinid="
			Response.Write Rs("skinid")
			Response.Write " class=showmeun>"
		Else
			Response.Write "<a href=?action=manage&skinid="
			Response.Write Rs("skinid")
			Response.Write ">"
		End If
		Response.Write Rs("page_name")
		Response.Write "<a/></td>"
		Response.Write "   <td align=""center""><a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=1>编辑CSS样式</a></td>"
		Response.Write "   <td align=""center""><a href=?action=set&TemplateID=" & Rs("TemplateID") & ">模板常规设置</a></td>"
		Response.Write "   <td align=""center""><a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=0>编辑通栏模板</a></td>"
		Response.Write "   <td align=""center"">"
		Response.Write "   <a href=?action=default&skinid=" & Rs("skinid") & " onclick=""{if(confirm('您确定要将该模板设为默认模板吗?')){return true;}return false;}"">设为默认模板</a> |"
		Response.Write "   <a href=?action=del&skinid=" & Rs("skinid") & "&TemplateID=" & Rs("TemplateID") & " onclick=""{if(confirm('模板删除后将不能恢复，您确定要删除该模板吗?')){return true;}return false;}"">删除模板</a></td>"
		Response.Write " </tr>"
		Rs.MoveNext
	Loop
	Set Rs = Nothing
	Response.Write "<form method=Post name=""myform"" action=""?action=newtemplate"" onSubmit=""return Juge(this)"">"
	Response.Write " <tr>"
	Response.Write "   <td colspan=""5"" align=""center"" class=""TableRow2"">模板名称：<input name=""page_name"" type=""text"" size=""20"">"
	Response.Write "   模板目录：<input name=""TemplateDir"" type=""text"" size=""20"" value=""skin/default/"">"
	Response.Write "   <input type=""submit"" name=""Submit"" value=""新建模板"" class=Button><br>"
	Response.Write "   <strong>注意：</strong>模板目录相对于系统根目录下，模板新建成功后，请到相应的频道模板新建分页模板</td>"
	Response.Write " </tr></form>"
	Response.Write "</table>"

End Sub

Sub EditStyle()
	Dim StyleTitle
	Dim PageContent

	If CInt(Request("StyleID")) = 1 Then
		StyleTitle = "编辑CSS样式"
	Else
		StyleTitle = "编辑模板通栏"
	End If
	Set Rs = enchiasp.Execute("select * From [ECCMS_Template] where pageid = 0 And TemplateID = " & Request("TemplateID"))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的系统参数！</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	PageContent = Split(Rs("page_content"), "|||")
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write " <tr>"
	Response.Write "   <th colspan=""2"">" & StyleTitle & "（修改以下设置必须具备一定网页知识）</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <td Colspan=2 align=right class=TableRow1>"
	Call TemplateJumpList
	Response.Write "</td>"
	Response.Write " </tr><form method=Post name=""myform"" action=""?action=savestyle"" onSubmit=""return Juge(this)"">"
	Response.Write "  <input type=hidden name=TemplateID value=""" & Rs("TemplateID") & """>"
	Response.Write "  <input type=hidden name=StyleID value=""" & Request("StyleID") & """>"
	Response.Write " <tr>"
	Response.Write "   <td width=""10%"" nowrap class=""TableRow2""><strong>模板名称</strong></td>"
	Response.Write "   <td width=""90%"" class=""TableRow1""><input name=""page_name"" type=""text"" size=""20"" value=""" & Rs("page_name") & """>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "   <a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=1&ChannelID=" & ChannelID & " class=showmeun>编辑CSS样式</a> | " & vbCrLf
	Response.Write "   <a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=0&ChannelID=" & ChannelID & " class=showmeun>编辑通栏模板</a> | " & vbCrLf
	Response.Write "   <a href=?action=set&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & " class=showmeun>模板基本设置</a></td>" & vbCrLf
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write "   <td class=""TableRow2""><strong>模板目录</strong></td>"
	Response.Write "   <td class=""TableRow1""><input name=""TemplateDir"" type=""text"" size=""20"" value=""" & Rs("TemplateDir") & """></td>" & vbCrLf
	Response.Write " </tr>"
	Response.Write " <tr"
	If CInt(Request("StyleID")) <> 1 Then
		Response.Write " style=""display:none"""
	End If
	Response.Write ">"
	Response.Write "   <td nowrap class=""TableRow2""><strong>CSS样式内容</strong><br>相关标签说明<br><br>{$InstallDir}<br>系统根目录<br><br>{$SkinPath}<br>皮肤图片路径</td>"
	Response.Write "   <td class=""TableRow1""><textarea name=""page_style"" style=""width:100%;"" rows=""30"" wrap=""OFF"" id=page_style>" & Server.HTMLEncode(PageContent(0)) & "</textarea>"
	Response.Write "   <div align=right><a href=""javascript:admin_Size(-15,'page_style')""><img src=""images/minus.gif"" unselectable=on border=0></a> <a href=""javascript:admin_Size(15,'page_style')""><img src=""images/plus.gif"" unselectable=on border=0></div></td>"
	Response.Write " </tr>"
	Response.Write " <tr"
	If CInt(Request("StyleID")) = 1 Then
		Response.Write " style=""display:none"""
	End If
	Response.Write ">"
	Response.Write "   <td nowrap class=""TableRow2""><strong>模板顶部通栏</strong></td>"
	Response.Write "   <td class=""TableRow1""><textarea name=""page_content1"" style=""width:100%;"" rows=""20"" wrap=""OFF"" id=content1>" & Server.HTMLEncode(PageContent(1)) & "</textarea>"
	Response.Write "   <div align=right><a href=""javascript:admin_Size(-10,'page_content1')""><img src=""images/minus.gif"" unselectable=on border=0></a> <a href=""javascript:admin_Size(10,'page_content1')""><img src=""images/plus.gif"" unselectable=on border=0></div></td>"
	Response.Write " </tr>"
	Response.Write " <tr"
	If CInt(Request("StyleID")) = 1 Then
		Response.Write " style=""display:none"""
	End If
	Response.Write ">"
	Response.Write "   <td nowrap class=""TableRow2""><strong>模板底部通栏</strong></td>"
	Response.Write "   <td class=""TableRow1""><textarea name=""page_content2"" style=""width:100%;"" rows=""20"" wrap=""OFF"" id=page_content2>" & Server.HTMLEncode(PageContent(2)) & "</textarea>"
	Response.Write "   <div align=right><a href=""javascript:admin_Size(-10,'page_content2')""><img src=""images/minus.gif"" unselectable=on border=0></a> <a href=""javascript:admin_Size(10,'page_content2')""><img src=""images/plus.gif"" unselectable=on border=0></div></td>"
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write "   <td class=""TableRow2""></td>"
	Response.Write "   <td class=""TableRow1"" align=""center""><input type=""button"" name=""Submit4"" onclick=""javascript:history.go(-1)"" value=""返回上一页"" class=Button>        <input type=""submit"" name=""btnSubmit"" value=""保存设置"" class=Button></td>"
	Response.Write " </tr></form>"
	Response.Write "</table>"
	Set Rs = Nothing

End Sub

Sub SettingTemplate()
	Dim TemplateStr
	Dim TemplateHelpStr
	Dim TempHelpStr
	Dim TempTitleStr
	Dim TempHelpValue
	Dim TempTitleValue

	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板ID不能为空！</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("Select * From ECCMS_Template where TemplateID = " & Request("TemplateID"))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的模板参数！</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	
	TemplateStr = Split(Rs("page_setting"), "|||")
	TemplateHelpStr = Split(Rs("Template_Help"), "@@@")
	TempTitleStr = Split(TemplateHelpStr(0), "|||")
	TempHelpStr = Split(TemplateHelpStr(1), "|||")
	
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write "<tr>"
	Response.Write " <th Colspan=2>当前模板 (" & Rs("page_name") & ") 基本设置</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <td width=""30%"" Class=BodyTitle align=""center"">"
	Response.Write Rs("page_name")
	Response.Write "</td>" & vbCrLf
	Response.Write " <td width=""70%"" Class=BodyTitle align=""center"">"
	If Rs("pageid") <> 0 Then
		Response.Write "<a href=?action=edit&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & " class=showmeun>编辑该模板界面风格</a> | "
	Else
		Response.Write "<a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=0&ChannelID=" & ChannelID & " class=showmeun>编辑该模板通栏</a> | "
	End If
	Response.Write "<a href=?action=manage&ChannelID=" & Rs("ChannelID") & " class=showmeun>返回模板首页</a>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <td Colspan=2 align=right class=TableRow1>"
	Call TemplateJumpList
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<form name=myform method=""post"" action=""?action=saveset&ChannelID=" & ChannelID & """>"
	Response.Write "<input type=""hidden"" name=""TemplateID"" value=""" & Rs("TemplateID") & """>"
	If TemplateStr(UBound(TemplateStr)) = "" Then TemplateStr(UBound(TemplateStr)) = "del"
	For i = 0 To UBound(TemplateStr)
		If i < UBound(TempHelpStr) Then
			TempHelpValue = TempHelpStr(i)
		Else
			TempHelpValue = "//"
		End If
		If i < UBound(TempTitleStr) Then
			TempTitleValue = TempTitleStr(i)
		Else
			TempTitleValue = "基本设置说明"
		End If
		Response.Write "<tr>"
		Response.Write " <td class=""TableRow2"">"
		Response.Write "<font color=blue style=""font-family:tahoma"">"
		Response.Write i
		Response.Write "、</font>"
		Response.Write TempTitleValue
		Response.Write " </td>"
		Response.Write " <td class=""TableRow1"">"
		If Rs("pageid") = 0 And i <= 6 And LCase(TemplateStr(i)) <> "del" Then
			Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t" & i & """ value="""
			Response.Write Server.HTMLEncode(TemplateStr(i))
			Response.Write """ size=10> "
			If i <> 0 Then
				Response.Write "<font size=3 color=" & TemplateStr(i) & "><b>■</b></font>"
			End If
		ElseIf LenB(TemplateStr(i)) > 90 Then
			Response.Write "<textarea name=""TemplateStr"" id=""t" & i & """  cols=""80"" rows=""3"">"
			Response.Write Server.HTMLEncode(TemplateStr(i))
			Response.Write "</textarea><br>"
			Response.Write "<a href=""javascript:admin_Size(-10,'t" & i & "')""><img src=""images/minus.gif"" unselectable=""on"" border='0'></a> <a href=""javascript:admin_Size(10,'t" & i & "')""><img src=""images/plus.gif"" unselectable=""on"" border='0'></a> "
		ElseIf LenB(TemplateStr(i)) <= 21 And LCase(TemplateStr(i)) <> "del" Then
			Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t" & i & """ value="""
			Response.Write Server.HTMLEncode(TemplateStr(i))
			Response.Write """ size=20> "
		Else
			Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t" & i & """ value="""
			Response.Write Server.HTMLEncode(TemplateStr(i))
			Response.Write """ size=60> "
		End If
		Response.Write "<INPUT TYPE=""hidden"" NAME=""ReadME"" id=""r" & i & """ value=""" & TempHelpValue & """>"
		Response.Write "<a href=# onclick=""helpscript(r" & i & ");return false;"" class=""helplink""><img src=""images/help.gif"" border=0 title=""点击查阅管理帮助！""></a>"
		Response.Write " </td>"
		Response.Write "</tr>"
	Next
	Response.Write "<tr>"
	Response.Write " <td class=""TableRow2"" align=""center""><a href=""?action=help&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & """><font color=blue>该模板帮助设置</font></a></td>"
	Response.Write " <td class=""TableRow1"" align=""center""><input type=""button"" name=""Submit4"" onclick=""javascript:history.go(-1)"" value=""返回上一页"" class=Button>          <input type=""submit"" name=""btnSubmit"" value=""保存设置"" class=Button></td>"
	Response.Write "</tr>"
	Response.Write "</form>"
	Response.Write "<tr>"
	Response.Write " <td Colspan=2 class=""TableRow2""><font color=red><b>警告：</b></font><li><font color=blue>请不要在文本框中输入“del”，这样会删除相应的设置数据，那么模板会出现错误，导致网站不能正常访问。</font></li></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Set Rs = Nothing

End Sub

Sub SaveTemplateSet()
	Dim TempStr
	Dim TemplateStr

	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板ID不能为空！</li>"
		Exit Sub
	End If
	'提取表单中的数据

	TemplateStr = ""
	For Each TempStr In Request.Form("TemplateStr")
		If LCase(TempStr) <> "del" Then
			TemplateStr = TemplateStr & Replace(TempStr, "|||", "") & "|||"
		End If
	Next
	TemplateStr = enchiasp.CheckStr(TemplateStr)
	enchiasp.Execute ("update [ECCMS_Template] set page_setting ='" & TemplateStr & "' Where TemplateID =" & Request("TemplateID"))
	Call RemoveCache
	Succeed ("<li>恭喜您！修改模板基本设置成功。</li>")

End Sub

Sub EditTemplateHelp()
	Dim TemplateHelpStr
	Dim TempTitleStr
	Dim TempHelpStr
	'-----------模板帮助设置开始----------------
	'编辑模板帮助

	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板ID不能为空！</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("Select TemplateID,page_name,Template_Help From ECCMS_Template where TemplateID = " & Request("TemplateID"))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的模板参数！</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	TemplateHelpStr = Split(Rs("Template_Help"), "@@@")
	TempTitleStr = Split(TemplateHelpStr(0), "|||")
	TempHelpStr = Split(TemplateHelpStr(1), "|||")
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write "<tr>"
	Response.Write " <th Colspan=2>当前模板 (" & Rs("page_name") & ") 帮助管理</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <td width=""40%"" Class=TableTitle align=""center"">模板设置标题说明</td>"
	Response.Write " <td width=""60%"" Class=TableTitle align=""center"">模板设置帮助详细说明</td>"
	Response.Write "</tr>"
	Response.Write "<form name=myform method=""post"" action=""?action=savehelp&ChannelID=" & ChannelID & """>"
	Response.Write "<input type=""hidden"" name=""TemplateID"" value=""" & Rs("TemplateID") & """>"
	If TempTitleStr(UBound(TempTitleStr)) = "" Then
		TempTitleStr(UBound(TempTitleStr)) = "del"
	End If
	For i = 0 To UBound(TempTitleStr)
		Response.Write "<tr>" & vbCrLf
		Response.Write Chr(9) & "<td class=""TableRow1"">"
		Response.Write "<input Type=""text"" name=""TempTitleStr"" value="""
		Response.Write Server.HTMLEncode(TempTitleStr(i))
		Response.Write """ size=50> "
		Response.Write "</td>" & vbCrLf
		Response.Write Chr(9) & "<td class=""TableRow1"">"
		If LenB(TempHelpStr(i)) > 70 Then
			Response.Write "<textarea name=""TempHelpStr""  cols=""70"" rows=""3"">"
			Response.Write Server.HTMLEncode(TempHelpStr(i))
			Response.Write "</textarea>"
		Else
			Response.Write "<input Type=""text"" name=""TempHelpStr"" value="""
			Response.Write Server.HTMLEncode(TempHelpStr(i))
			Response.Write """ size=50> "
		End If
		Response.Write "</td>" & vbCrLf
		Response.Write "</tr>" & vbCrLf
	Next
	Response.Write "<tr>"
	Response.Write " <td class=""TableRow2"" align=""center""><a href=""?action=set&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & """><font color=blue>该模板基本设置</font></a></td>"
	Response.Write " <td class=""TableRow1"" align=""center""><input type=""button"" name=""Submit4"" onclick=""javascript:history.go(-1)"" value=""返回上一页"" class=Button>          <input type=""submit"" name=""btnSubmit"" value=""保存帮助"" class=Button></td>"
	Response.Write "</tr></form><tr>"
	Response.Write " <td Colspan=2 class=""TableRow2""><font color=blue><b>注意：</b> 帮助内容是针对相应的模板基本设置。</font><li>帮助编辑规则：如果想清除该帮助，请在对应的文本框中输入“del”，那么帮助数据的序号就会前移。</li>"
	Response.Write " <li>如果不想改变帮助数据的序号,仅把该项目的数据清空,则只需要把内容清空。</li></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Set Rs = Nothing
End Sub

Sub SaveTemplateHelp()
	Dim TempStr
	Dim HelpStr
	Dim TemplateHelpStr
	Dim TempHelpStr
	Dim TempTitleStr
	'保存模板帮助

	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板ID不能为空！</li>"
		Exit Sub
	End If
	'提取表单中的数据

	TempTitleStr = ""
	For Each TempStr In Request.Form("TempTitleStr")
		If LCase(TempStr) <> "del" Then
			TempTitleStr = TempTitleStr & Replace(TempStr, "|||", "") & "|||"
		End If
	Next
	TempHelpStr = ""
	For Each HelpStr In Request.Form("TempHelpStr")
		TempHelpStr = TempHelpStr & Replace(HelpStr, "|||", "") & "|||"
	Next
	TemplateHelpStr = enchiasp.CheckStr(TempTitleStr & "@@@" & TempHelpStr)
	enchiasp.Execute ("update [ECCMS_Template] set Template_Help ='" & TemplateHelpStr & "' Where TemplateID =" & Request("TemplateID"))
	Call RemoveCache
	OutHintScript ("恭喜您！设置模板帮助成功。")
	'-----------模板帮助设置结束----------------
End Sub

Sub SaveStyle()
	Dim TemplateDir
	Dim page_content
	Dim FileName
	Dim FileContent
	

	If Trim(Request.Form("page_name")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板名称不能为空！</li>"
	End If
	If Trim(Request.Form("TemplateDir")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板目录不能为空！</li>"
	End If
	If Not enchiasp.IsValidChar(Request.Form("TemplateDir")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板目录中含有非法字符或者中文字符！</li>"
	End If
	If Right(Request.Form("TemplateDir"), 1) <> "/" Then
		TemplateDir = Trim(Request.Form("TemplateDir")) & "/"
	Else
		TemplateDir = Trim(Request.Form("TemplateDir"))
	End If
	If FoundErr Then Exit Sub
	page_content = enchiasp.CheckStr(Request.Form("page_style") & "|||" & Request.Form("page_content1") & "|||" & Request.Form("page_content2") & "|||")
	enchiasp.Execute ("update [ECCMS_Template] set TemplateDir='" & TemplateDir & "',page_name='" & Trim(Request.Form("page_name")) & "',page_content='" & page_content & "' where TemplateID = " & Request("TemplateID"))
	
	enchiasp.CreatPathEx (enchiasp.InstallDir & TemplateDir)
	FileName = enchiasp.InstallDir & TemplateDir & "style.css"
	FileContent = Request.Form("page_style")
	FileContent = Replace(FileContent, "{$InstallDir}", enchiasp.InstallDir)
	FileContent = Replace(FileContent, "{$SkinPath}", TemplateDir)
	enchiasp.CreatedTextFile FileName, FileContent
	Call RemoveCache
	If CInt(Request("StyleID")) = 1 Then
		SucMsg = ("<li>恭喜您！编辑CSS样式成功。</li>")
	Else
		SucMsg = ("<li>恭喜您！编辑模板通栏成功。</li>")
	End If
	Succeed (SucMsg)
	
End Sub

Sub NewTemplate()
	Dim TemplateDir
	Dim TemplateName
	
	Dim TemplateFields
	Dim TemplateValues
	Dim newskinid

	If Trim(Request.Form("page_name")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板名称不能为空！</li>"
	End If
	If Trim(Request.Form("TemplateDir")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板目录不能为空！</li>"
	End If
	If Not enchiasp.IsValidChar(Request.Form("TemplateDir")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板目录中含有非法字符或者中文字符！</li>"
	End If
	If Right(Request.Form("TemplateDir"), 1) <> "/" Then
		TemplateDir = Trim(Request.Form("TemplateDir")) & "/"
	Else
		TemplateDir = Trim(Request.Form("TemplateDir"))
	End If
	If FoundErr Then Exit Sub
	
	enchiasp.CreatPathEx (enchiasp.InstallDir & TemplateDir)
	'Response.Write "<li>正在新建模板…… 请稍候…… 现在请不要刷新页面。</li>"
	TemplateName = enchiasp.CheckStr(Trim(Request("page_name")))
	Set Rs = enchiasp.Execute("Select Max(skinid) from [ECCMS_Template] where pageid = 0")
	If Not (Rs.EOF And Rs.BOF) Then
		newskinid = Rs(0)
	End If
	If IsNull(newskinid) Then newskinid = 0
	Rs.Close
	newskinid = newskinid + 1
	Set Rs = enchiasp.Execute("select * From [ECCMS_Template] where IsDefault = 1")
	If Not (Rs.BOF And Rs.EOF) Then
		Do While Not Rs.EOF
			If Rs("pageid") <> 0 Then
				TemplateName = Rs("page_name")
			End If
			TemplateFields = "ChannelID,skinid,pageid,TemplateDir,page_name,page_content,page_setting,Template_Help,isDefault"
			TemplateValues = "" & Rs("ChannelID") & "," & newskinid & "," & Rs("pageid") & ",'" & TemplateDir & "','" & TemplateName & "','" & enchiasp.CheckStr(Rs("page_content")) & "','" & enchiasp.CheckStr(Rs("page_setting")) & "','" & enchiasp.CheckStr(Rs("Template_Help")) & "',0"
			SQL = "insert into [ECCMS_Template](" & TemplateFields & ")values(" & TemplateValues & ")"
			enchiasp.Execute (SQL)
			Rs.MoveNext
		Loop
	Else
		TemplateValues = "0," & newskinid & ",0,'" & TemplateDir & "','" & TemplateName & "','|||||||||','|||','|||@@@|||',0"
		SQL = "insert into [ECCMS_Template](ChannelID,skinid,pageid,TemplateDir,page_name,page_content,page_setting,Template_Help,isDefault)values(" & TemplateValues & ")"
		enchiasp.Execute (SQL)
	End If
	Set Rs = Nothing
	OutHintScript ("新建模板“" & Request.Form("page_name") & "”成功！")
	
End Sub

Sub DelTemplate()
	Set Rs = enchiasp.Execute("Select IsDefault From ECCMS_Template where TemplateID = " & Request("TemplateID"))
	If Rs(0) = 1 Then
		ErrMsg = ErrMsg + "<li>此模板是默认模版，不允许删除。"
		FoundErr = True
		Exit Sub
	Else
		enchiasp.Execute ("Delete From ECCMS_Template where skinid = " & Request("skinid"))
		Application.Contents.RemoveAll
		OutHintScript ("模板删除成功！")
	End If
	Set Rs = Nothing
End Sub

Sub DefaultTemplate()
	If Trim(Request("skinid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板ID不能为空！</li>"
		Exit Sub
	End If
	enchiasp.Execute ("update ECCMS_Template set isDefault = 0 where isDefault = 1")
	enchiasp.Execute ("update ECCMS_Template set isDefault = 1 where skinid = " & Request("skinid"))
	enchiasp.DelCahe "MainStyle" & Request("skinid")
	enchiasp.DelCahe "DefaultSkinID"
	OutHintScript ("恭喜您，设置默认模板成功！")
End Sub

Sub ChannelTemplate()
	Dim skinid
	Dim Rss
	skinid = CLng(Request("skinid"))
	If Request("skinid") <> "" And ChannelID = 0 Then
		SQL = "skinid = " & skinid
	ElseIf Request("skinid") <> "" And ChannelID <> 0 Then
		SQL = "(skinid = " & skinid & " And ChannelID = " & ChannelID & ") Or (skinid = " & skinid & " And ChannelID = 0)"
	Else
		Set Rs = enchiasp.Execute("Select * From ECCMS_Template where isDefault = 1 And pageid = 0")
		If ChannelID <> 0 Then
			SQL = "(skinid = " & Rs("skinid") & " And ChannelID = " & ChannelID & ") Or (skinid = " & Rs("skinid") & " And ChannelID = 0)"
		Else
			SQL = "skinid = " & Rs("skinid")
		End If
		skinid = Rs("skinid")
		Set Rs = Nothing
	End If
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write "<tr>"
	Response.Write " <th colspan=2>"
	Response.Write ModuleName
	Response.Write "模板管理列表</th>"
	Response.Write "</tr>"
	Response.Write " <form name=myform method=""post"" action=""?action=manage"">"
	Response.Write " <input type=""hidden"" name=""ChannelID"" value=""" & ChannelID & """>"
	Response.Write " <td class=""TableRow2"">请选择模板："
	Set Rss = enchiasp.Execute("Select * From ECCMS_Template where pageid = 0")
	Response.Write "<select name=skinid>"
	Do While Not Rss.EOF
		Response.Write " <option value="
		Response.Write Rss("skinid")
		If Rss("skinid") = skinid Then Response.Write " selected"
		Response.Write ">"
		Response.Write Rss("page_name")
		Response.Write "</option>"
		Rss.MoveNext
	Loop
	Set Rss = Nothing
	Response.Write "</select>"
	Response.Write "&nbsp;<input type=submit value=""提 交"" name=""B1"" class=button>"
	Response.Write " </td>"
	Response.Write " </form>"
	Response.Write " <form name=myform method=""post"" action=""?action=newpage"">"
	Response.Write " <input type=""hidden"" name=""skinid"" value=""" & skinid & """>"
	Response.Write " <td class=""TableRow2"">新建"
	Response.Write ModuleName
	Response.Write "分模板页面："
	Response.Write "<select name=pageid>"
	Response.Write " <option value=''>↓请选择模板类型↓</option>"
	Response.Write " <option value=0>网站首页模板</option>"
	Response.Write " <option value=1>≡"
	Response.Write ModuleName
	Response.Write "首页≡</option>"
	Response.Write " <option value=2>　├列表页面</option>"
	Response.Write " <option value=3>　├内容页面</option>"
	Response.Write " <option value=4>　├专题页面</option>"
	Response.Write " <option value=5>　├推荐页面</option>"
	Response.Write " <option value=6>　├热门页面</option>"
	Response.Write " <option value=7>　├搜索页面</option>"
	Response.Write " <option value=8>　├其它页面</option>"
	Response.Write "</select> "
	If ChannelID = 0 Then
		Response.Write "请选择频道："
		Response.Write "<select name=ChannelID>"
		sModuleName = Split(sModuleName, "|||")
		sChannelID = Split(sChannelID, "|||")
		For i = 0 To UBound(sModuleName) - 1
			Response.Write " <option value="
			Response.Write sChannelID(i)
			Response.Write ">"
			Response.Write sModuleName(i)
			Response.Write "</option>"
		Next
		Response.Write "</select>"
	Else
		Response.Write " <input type=""hidden"" name=""ChannelID"" value=""" & ChannelID & """>"
	End If
	Response.Write "<br>模板名称："
	Response.Write "<input type=""text"" name=""pagename"" size=35>"
	'Response.Write "模板唯一标识：(请用英文)<input type=""text"" name=""pagemark"" size=20>"
	Response.Write "&nbsp;<input type=submit value=""新建分模板"" name=""B2"" class=button>&nbsp;"
	Response.Write " </td>"
	Response.Write " </form>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <th width=""40%"">模板名称</th>"
	Response.Write " <th width=""60%"">模板相关设置</th>"
	Response.Write "</tr>"
	Set Rs = enchiasp.Execute("Select * From ECCMS_Template where " & SQL & " Order By TemplateID")
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td colspan=2 align=center>没有找到相关模板！</td></tr>"
	Else
		Response.Write "<tr>"
		Response.Write " <td colspan=2 Class=BodyTitle>当前模板：<font color=blue>"
		If Rs("pageid") = 0 Then
			Response.Write Rs("page_name")
		End If
		Response.Write "</font></td>"
		Response.Write "</tr>"
		Do While Not Rs.EOF
			Response.Write "<tr "
			Response.Write MouseStyle
			Response.Write ">"
			Response.Write " <td><li>"
			If Rs("ChannelID") = 0 Then
				Response.Write "<font color=blue>"
				Response.Write Rs("page_name")
				Response.Write "</font>"
			ElseIf ChannelID = 0 And Rs("pageid") = 1 Then
				Response.Write "<font color=red>"
				Response.Write Rs("page_name")
				Response.Write "</font>"
			Else
				Response.Write Rs("page_name")
			End If
			Response.Write "</li></td>"
			Response.Write " <td>编辑该模块： "
			If Rs("pageid") = 0 Then
				Response.Write "<a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=1&ChannelID=" & ChannelID & ">编辑CSS样式</a> | "
				Response.Write "<a href=?action=editstyle&TemplateID=" & Rs("TemplateID") & "&StyleID=0&ChannelID=" & ChannelID & ">顶部和底部通栏</a> | "
				Response.Write "<a href=?action=set&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & ">模板常规设置</a>"
			Else
				Response.Write "<a href=?action=edit&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & ">编辑模板界面风格</a> | "
				Response.Write "<a href=?action=set&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & ">修改模板基本设置</a>"
			End If
			'If Rs("pageid") = 8 Then
				'Response.Write " | <a href=?action=del&TemplateID=" & Rs("TemplateID") & ">删除分页面模板</a>"
			'End If
			Response.Write "</td>"
			Response.Write "</tr>"
			Rs.MoveNext
		Loop
		Response.Write "<tr>"
		Response.Write " <td class=""TableRow1""></td>"
		Response.Write " <td class=""TableRow1""></td>"
		Response.Write "</tr>"
	End If
	Set Rs = Nothing
	Response.Write "<form method=Post name=""myform"" action=""?action=newtemplate"" onSubmit=""return Juge(this)"">"
	Response.Write " <tr>"
	Response.Write "   <td colspan=""5"" align=""center"" class=""TableRow2"">模板名称：<input name=""page_name"" type=""text"" size=""20"">"
	Response.Write "   模板目录：<input name=""TemplateDir"" type=""text"" size=""20"" value=""skin/default/"">"
	Response.Write "   <input type=""submit"" name=""Submit"" value=""新建模板"" class=Button><br>"
	Response.Write "   <strong>注意：</strong>模板目录相对于系统根目录下，模板新建成功后，请到相应的频道模板新建分页模板</td>"
	Response.Write " </tr></form>"
	Response.Write "</table>"
End Sub

Sub EditTemplatePage()
	Dim page_content
	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的系统参数！</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("select * From [ECCMS_Template] where TemplateID = " & Request("TemplateID"))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的系统参数！</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	If Not IsNull(Rs("page_content")) Then
		page_content = Split(Rs("page_content") & "|||@@@|||", "|||@@@|||")
	End If
	Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
	Response.Write " <tr>"
	Response.Write "   <th colspan=""2"">编辑当前模板：" & Rs("page_name") & " （修改以下设置必须具备一定网页知识）</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write " <td Colspan=2 align=right class=TableRow1>"
	Call TemplateJumpList
	Response.Write "</td>"
	Response.Write " </tr><form method=Post name=""myform"" action=""?action=save&ChannelID=" & ChannelID & """>"
	Response.Write "  <input type=hidden name=TemplateID value=""" & Rs("TemplateID") & """>"
	Response.Write "  <input type=hidden name=pageid value=""" & Rs("pageid") & """>"
	Response.Write " <tr>"
	Response.Write "   <td width=""10%"" nowrap class=""TableRow2""><strong>当前模板名称</strong></td>"
	Response.Write "   <td width=""90%"" class=""TableRow1"">"
	Response.Write "<input type=""text"" name=""pagename"" value="""
	Response.Write Rs("page_name")
	Response.Write """ size=35>"
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "   <a href=?action=set&TemplateID=" & Rs("TemplateID") & "&ChannelID=" & ChannelID & " class=showmeun>修改该模板基本设置</a> | "
	Response.Write "<a href=?action=manage&ChannelID=" & Rs("ChannelID") & " class=showmeun>返回模板首页</a></td>" & vbCrLf
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write "   <td class=""TableRow2""><strong>生成标签</strong></td>"
	Response.Write "   <td class=""TableRow1"">"
	Response.Write "<a href=javascript:openDialog('admin_label.asp?action=list&ChannelID=1',550,490)>文章标签</a> | "
	Response.Write "<a href=javascript:openDialog('admin_label.asp?action=list&ChannelID=2',550,460)>软件标签</a> | "
	Response.Write "<a href=javascript:openDialog('admin_label.asp?action=list&ChannelID=3',550,460)>商城标签</a> | "
	Response.Write "<a href=javascript:openDialog('admin_label.asp?action=list&ChannelID=5',550,460)>动画标签</a> | "
	Response.Write "<a href=javascript:openDialog('admin_label.asp?action=image&ChannelID=" & ChannelID & "',550,460)>图片标签</a> | "
	Response.Write "<a href=javascript:openDialog('admin_label.asp',550,460)>模板标签管理</a>"
	Response.Write "</td>" & vbCrLf
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write "   <td width=""10%"" class=""TableRow2""><strong>模板内容</strong><br>相关标签说明<br><br>{$InstallDir}<br>系统根目录<br><br>{$SkinPath}<br>皮肤图片路径</td>"
	Response.Write "   <td class=""TableRow1""><textarea name=""content"" style=""width:100%;"" rows=""30"" wrap=""OFF"" id=PageContent>" & Server.HTMLEncode(page_content(0)) & "</textarea>"
	Response.Write "   <div align=right><a href=""javascript:admin_Size(-10,'PageContent')""><img src=""images/minus.gif"" unselectable=on border=0></a> <a href=""javascript:admin_Size(10,'PageContent')""><img src=""images/plus.gif"" unselectable=on border=0></div></td>"
	Response.Write " </tr>"
	If Rs("pageid") = 2 And ChannelID <> 9999 Then
		Response.Write " <tr"
		If ChannelID = 3 Or ChannelModuleID = 5 Then
			Response.Write " style=""display:none"""
		End If
		Response.Write ">"
		Response.Write "   <td width=""10%"" class=""TableRow2""><strong>模板内容</strong><br>说明：<br>此模板是大类列表页面模板，如果你只有一级分类此模板可能不用编辑。<br>如：你的分类下面包含子分类。当用访问父级分类的时候就显示此模板内容</td>"
		If ChannelID = 3 Or ChannelModuleID = 5 Then
			Response.Write "   <td class=""TableRow1""><textarea name=""content1"" id=PageContent1></textarea>"
		Else
			Response.Write "   <td class=""TableRow1""><textarea name=""content1"" style=""width:100%;"" rows=""30"" wrap=""OFF"" id=PageContent1>" & Server.HTMLEncode(page_content(1)) & "</textarea>"
		End If
		Response.Write "   <div align=right><a href=""javascript:admin_Size(-10,'PageContent1')""><img src=""images/minus.gif"" unselectable=on border=0></a> <a href=""javascript:admin_Size(10,'PageContent1')""><img src=""images/plus.gif"" unselectable=on border=0></div></td>"
		Response.Write " </tr>"
	End If
	Response.Write " <tr>"
	Response.Write "   <td class=""TableRow2""></td>"
	Response.Write "   <td class=""TableRow1"" align=""center""><input type=""button"" name=""Submit4"" onclick=""javascript:history.go(-1)"" value=""返回上一页"" class=Button>        <input type=""submit"" name=""btnSubmit"" value=""保存模板"" class=Button></td>"
	Response.Write " </tr></form>"
	Response.Write "</table>"
	Set Rs = Nothing
End Sub

Sub SaveTemplatePage()
	Dim TemplateContent
	Dim page_name

	If Trim(Request("TemplateID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板ID不能为空！</li>"
		Exit Sub
	End If
	If Trim(Request("pagename")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板内容不能为空！</li>"
		Exit Sub
	End If
	If Trim(Request("content")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板内容不能为空！</li>"
		Exit Sub
	End If
	TemplateContent = Request.Form("content")
	If Request.Form("pageid") = 2 And ChannelID <> 9999 And ChannelID <> 3 And ChannelModuleID <> 5 Then
		TemplateContent = TemplateContent & "|||@@@|||" & Request.Form("content1")
	End If
	TemplateContent = enchiasp.CheckStr(TemplateContent)
	page_name = enchiasp.CheckStr(Request.Form("pagename"))
	enchiasp.Execute ("update [ECCMS_Template] set page_name = '" & page_name & "', page_content ='" & TemplateContent & "' Where TemplateID =" & Request("TemplateID"))
	Call RemoveCache
	Succeed ("<li>恭喜您！修改模板基本设置成功。</li>")
End Sub

Sub NewTemplatePage()
	Dim Rss
	Dim pageid
	Dim skinid
	Dim TemplateName
	Dim TemplateFields
	Dim TemplateValues
	If Trim(Request("pageid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择模板所属类型！</li>"
		Exit Sub
	End If
	If Trim(Request("pagename")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>模板名称不能为空！</li>"
		Exit Sub
	End If
	If Trim(Request("skinid")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的系统参数！</li>"
		Exit Sub
	End If
	If Trim(Request("ChannelID")) = "" Or Request("ChannelID") = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的系统参数！</li>"
		Exit Sub
	End If
	If CInt(Request("pageid")) = 0 Then
		ChannelID = 0
		pageid = 1
	Else
		ChannelID = CInt(Request("ChannelID"))
		pageid = CInt(Request("pageid"))
	End If
	skinid = CLng(Request("skinid"))
	TemplateName = enchiasp.CheckStr(Trim(Request("pagename")))
	'If pageid <> 8 Then
		Set Rss = enchiasp.Execute("select pageid From [ECCMS_Template] where skinid = " & skinid & " And ChannelID = " & ChannelID & " And pageid = " & pageid)
		If Not (Rss.BOF And Rss.EOF) Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>此模板类型已经存在,请选择其它类型模板！</li>"
			Exit Sub
		End If
		Set Rss = Nothing
	'End If
	Set Rss = enchiasp.Execute("select * From [ECCMS_Template] where pageid = 0 And IsDefault = 1")
	Set Rs = enchiasp.Execute("select * From [ECCMS_Template] where skinid = " & Rss("skinid") & " And ChannelID = " & ChannelID & " And pageid = " & pageid)
	If Not (Rs.BOF And Rs.EOF) Then
		TemplateFields = "ChannelID,skinid,pageid,page_name,page_content,page_setting,Template_Help,isDefault"
		TemplateValues = "" & ChannelID & "," & skinid & "," & pageid & ",'" & TemplateName & "','" & enchiasp.CheckStr(Rs("page_content")) & "','" & enchiasp.CheckStr(Rs("page_setting")) & "','" & enchiasp.CheckStr(Rs("Template_Help")) & "',0"
		SQL = "insert into [ECCMS_Template](" & TemplateFields & ")values(" & TemplateValues & ")"
	Else
		TemplateValues = "" & ChannelID & "," & skinid & "," & pageid & ",'" & TemplateName & "','|||','1|||','|||@@@|||',0"
		SQL = "insert into [ECCMS_Template](ChannelID,skinid,pageid,page_name,page_content,page_setting,Template_Help,isDefault)values(" & TemplateValues & ")"
	End If
	Set Rs = Nothing
	Set Rss = Nothing
	enchiasp.Execute (SQL)
	OutHintScript ("新建分模板“" & Request.Form("pagename") & "”成功！")
End Sub
Sub TemplateJumpList()
	Dim rstmp, tmpsql, tmpname, sel
	Dim strTemp, strContent, strStetting
	strTemp = ""
	On Error Resume Next
	If ChannelID > 0 Then
		If Trim(Request("skinid")) <> "" And Trim(Request("skinid")) <> "0" Then
			tmpsql = "And skinid=" & Trim(Request("skinid"))
		Else
			tmpsql = "And isDefault=1"
		End If
		tmpsql = "SELECT TemplateID,pageid,page_name FROM ECCMS_Template WHERE (ChannelID=0 Or ChannelID=" & ChannelID & ") " & tmpsql & " ORDER BY TemplateID"
		Set rstmp = enchiasp.Execute(tmpsql)
		If rstmp.BOF And rstmp.EOF Then
			Set rstmp = Nothing
			Exit Sub
		End If
		Do While Not rstmp.EOF
			If rstmp("TemplateID") = CLng(Request("TemplateID")) Then
				sel = " selected"
			Else
				sel = ""
			End If
			If rstmp("pageid") = 0 Then
				strContent = strContent & "<option>↓" & rstmp("page_name") & "-界面风格↓</option>" & vbCrLf
				strContent = strContent & "<option value='?action=editstyle&TemplateID=" & rstmp("TemplateID") & "&StyleID=0&ChannelID=" & ChannelID & "'" & sel & ">顶部和底部通栏</option>" & vbCrLf
				strStetting = strStetting & "<option>↓" & rstmp("page_name") & "-基本设置↓</option>" & vbCrLf
				strStetting = strStetting & "<option value='?action=set&TemplateID=" & rstmp("TemplateID") & "&ChannelID=" & ChannelID & "'" & sel & ">主模板常规设置</option>" & vbCrLf
			Else
				strContent = strContent & "<option value='?action=edit&TemplateID=" & rstmp("TemplateID") & "&ChannelID=" & ChannelID & "'" & sel & ">" & rstmp("page_name") & "-界面</option>" & vbCrLf
				strStetting = strStetting & "<option value='?action=set&TemplateID=" & rstmp("TemplateID") & "&ChannelID=" & ChannelID & "'" & sel & ">" & rstmp("page_name") & "-设置</option>" & vbCrLf
			End If
			rstmp.MoveNext
		Loop
		rstmp.Close: Set rstmp = Nothing
		Response.Write "选择分页模板："
		Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
		Response.Write strContent
		Response.Write "</select>"
		Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
		Response.Write strStetting
		Response.Write "</select>"
	End If
End Sub
Sub RemoveCache()
	If Not IsNumeric(Request("TemplateID")) Then
		Exit Sub
	End If
	Dim rsCache
	Set rsCache = enchiasp.Execute("SELECT TemplateID,ChannelID,skinid,pageid FROM ECCMS_Template WHERE TemplateID=" & CLng(Request("TemplateID")))
	enchiasp.DelCahe "MainStyle" & rsCache("skinid")
	enchiasp.DelCahe "Templates" & rsCache("ChannelID") & rsCache("skinid") & rsCache("pageid")
	enchiasp.DelCahe "DefaultSkinID"
	enchiasp.DelCahe "ChannelMenu"
	enchiasp.DelCahe "SiteClassMap"
	rsCache.Close: Set rsCache = Nothing
End Sub
%>