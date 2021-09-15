<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/cls_public.asp"-->
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
Dim Action,i
If Not ChkAdmin("AdminJsFile" & ChannelID) Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "del"
	Call DeleteJsFile
Case "add"
	Call AddJsFile
Case "edit"
	Call EditJsFile
Case "savenew"
	Call SaveNewJsFile
Case "save"
	Call SaveJsFile
Case "make"
	Call MakeJsFile
Case "demo"
	Call DemoJsFile
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
	Dim page_count,tablebody,JsFileName
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th  width=""20%"">说明标题</th>"
	Response.Write "		<th  width=""40%"">调用方式</th>"
	Response.Write "		<th  width=""15%"">JS文件名称</th>"
	Response.Write "		<th  width=""25%"">管理选项</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=selform method=post action='admin_jsfile.asp'>"
	Response.Write "	<input type=hidden name=action value=""make"">"
	Response.Write "	<input type=hidden name=ChannelID value="""& ChannelID &""">"
	Set Rs = enchiasp.Execute("SELECT id,ChannelID,sTitle,stype,sFileName FROM ECCMS_ScriptFile WHERE ChannelID="& ChannelID &" ORDER BY id DESC")
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=4 class=TableRow1>没有添加" & sModuleName & "JS文件！</td></tr>"
	Else
		page_count = 0
		Do While Not Rs.EOF
			If (page_count mod 2) = 0 Then
				tablebody = "class=TableRow1"
			Else
				tablebody = "class=TableRow2"
			End If
			JsFileName = "<script src="""& enchiasp.SiteUrl & enchiasp.InstallDir & enchiasp.ChannelDir &"js/"& Rs("sFileName") &"""></script>"
			Response.Write "	<input type=hidden name=id value="""& Rs("id") &""">"
			Response.Write "	<tr align=center>"
			Response.Write "		<td " & tablebody & ">"& Rs("sTitle") &"</td>"
			Response.Write "		<td " & tablebody & "><input type=text name=jsfile size=50 value='" & Server.HTMLEncode(JsFileName) & "'></td>"
			Response.Write "		<td " & tablebody & " noWrap>" & Rs("sFileName") & "</td>"
			Response.Write "		<td " & tablebody & " noWrap><a href='?action=edit&ChannelID="& ChannelID &"&id="& Rs("id") &"'>设 置</a> | "
			Response.Write "<a href='?action=del&ChannelID="& ChannelID &"&id="& Rs("id") &"' onclick=""return confirm('您确定要删除此JS文件吗?')"">删 除</a> | "
			Response.Write "<a href='?action=make&ChannelID="& ChannelID &"&id="& Rs("id") &"'>生 成</a> | "
			Response.Write "<a href='?action=demo&ChannelID="& ChannelID &"&id="& Rs("id") &"'>演 示</a>"
			Response.Write "</td>"
			Response.Write "	</tr>"
			Rs.movenext
			page_count = page_count + 1
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 colspan=4 align=center>"
	Response.Write "<input type=submit name=submit3 value="" 生成所有JS文件 "" class=Button>&nbsp;&nbsp;"
	Response.Write "<input type=submit name=submit4 value="" 添加新的JS文件 "" onclick=""document.selform.action.value='add';"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=4>"
	Response.Write "<b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;请将上面的JS调用代码复制到模板相应的位置；"
	Response.Write "由于JS文件是系统生成的静态文件，所以要不定期的生成所有JS文件。"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
End Sub

Sub AddJsFile()
	Response.Write "<table cellspacing=1 align=center cellpadding=0 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=4>添加" & sModuleName & "JS文件</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action='admin_jsfile.asp'>"
	Response.Write "	<input type=hidden name=action value=""savenew"">"
	Response.Write "	<input type=hidden name=ChannelID value="""& ChannelID &""">"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right width=""20%""><b>文件说明：</b></td>"
	Response.Write "		<td class=tablerow1 width=""30%""><input type=text name=sTitle size=35 value=''></td>"
	Response.Write "		<td class=tablerow1 align=right width=""20%""><b>JS文件名称：</b></td>"
	Response.Write "		<td class=tablerow1 width=""30%""><input type=text name=sFileName size=20 value='file.js'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>显示类型：</b></td>"
	Response.Write "		<td class=tablerow2><input type=radio name=stype value=0 checked onClick=""stype1.style.display='';stype2.style.display='none';""> 列表&nbsp;&nbsp;"
	Response.Write "<input type=radio name=stype value=1 onClick=""stype2.style.display='';stype1.style.display='none';""> 图片</td>"
	Response.Write "		<td class=tablerow2 align=right><b>选择分类：</b></td>"
	Response.Write "		<td class=tablerow2>"
	Response.Write "<select name=""setting"" size='1'>"
	Response.Write "<option value=0>指定所有分类</option>"
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
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>所属专题：</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name=""setting"" size='1'>"
	Response.Write "<option value=0>不指定专题</option>"
	Dim RsObj
	Set RsObj = enchiasp.Execute("SELECT SpecialID,SpecialName FROM ECCMS_Special WHERE ChannelID="& ChannelID &" And ChangeLink=0")
	Do While Not RsObj.EOF
		Response.Write "<option value='" & RsObj("SpecialID") & "'>" & RsObj("SpecialName") & "</option>"
		RsObj.movenext
	Loop
	Set RsObj = Nothing
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "		<td class=tablerow1 align=right><b>调用类型：</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name=""setting"" size='1'>"
        Response.Write "  <option value=0>所有最新" & sModuleName & "</option>"
        Response.Write "  <option value='1'>所有推荐" & sModuleName & "</option>"
	Response.Write "  <option value='2'>所有热门" & sModuleName & "</option>"
	Response.Write "  <option value='3'>分类最新" & sModuleName & "</option>"
	Response.Write "  <option value='4'>分类推荐" & sModuleName & "</option>"
	Response.Write "  <option value='5'>分类热门" & sModuleName & "</option>"
        Response.Write "</select><font color=""#0066CC""></font>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=stype1>"
	Response.Write "		<td  bgcolor=""#FFFFFF"" colspan=4>"
	Response.Write "<table width=""100%"" cellspacing=1 align=center cellpadding=3 border=0>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>图片宽度：</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='85'></td>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>图片高度：</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='85'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>最多字符数：</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='22'></td>"
	Response.Write "		<td class=tablerow1 align=right><b>最多列表数：</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='10'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>垂直边距：</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='10'></td>"
	Response.Write "		<td class=tablerow2 align=right><b>水平边距：</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='10'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>图文对齐方式：</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='left' selected>左对齐</option>"
	Response.Write "		<option value='right'>右对齐</option>"
	Response.Write "		<option value='middle'>居中对齐</option>"
	Response.Write "		<option value='texttop'>文本上方</option>"
	Response.Write "		<option value='baseline'>基线</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b>调用样式：</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=20 value='class=dottedline'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>连接目标：</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	Response.Write "		<option value='_blank' selected>_blank</option>"
	Response.Write "		<option value='_self'>_self</option>"
	Response.Write "		<option value='_top'>_top</option>"
	Response.Write "		<option value='_parent'>_parent</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow2 align=right><b>标识符：</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=20 value='· '></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>是否显示图片：</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0' selected>不显示</option>"
	Response.Write "		<option value='1'>显示</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b>是否显示分类：</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0' selected>不显示</option>"
	Response.Write "		<option value='1'>显示</option>"
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>是否显示时间：</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	Response.Write "		<option value='0' selected>不显示</option>"
	Response.Write "		<option value='1'>显示</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow2 align=right><b>显示时间格式：</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	For i = 1 To 9
		Response.Write "<option value='" & i & "'"
		If i = 5 Then Response.Write " selected"
		Response.Write ">"
		Response.Write enchiasp.FormatDate(Now(),i)
		Response.Write "</option>" & vbCrLf
	Next
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=stype2 style=""display:none"">"
	Response.Write "		<td  bgcolor=""#FFFFFF"" colspan=4>"
	Response.Write "<table width=""100%"" cellspacing=1 align=center cellpadding=3 border=0>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>最多显示多少图片：</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='5'></td>"
	Response.Write "		<td class=tablerow2 align=right><b>每行显示多少图片：</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='5'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>显示最多字符数：</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='22'></td>"
	Response.Write "		<td class=tablerow1 align=right><b>是否新窗口打开：</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0' selected>本窗口打开</option>"
	Response.Write "		<option value='1'>新窗口打开</option>"
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>图片宽度：</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='120'></td>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>图片高度：</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='100'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>是否显示标题名称：</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0' selected>不显示</option>"
	Response.Write "		<option value='1'>显示</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b></b></td>"
	Response.Write "		<td class=tablerow1></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 colspan=4 align=center>"
	Response.Write "	<input type=button name=Submit4 onclick=""javascript:history.go(-1)"" value=""返回上一页"" class=Button>&nbsp;&nbsp;"
	Response.Write "	<input type=submit name=submit3 value="" 添加新的JS文件 "" onclick=""document.selform.action.value='add';"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub

Sub EditJsFile()
	On Error Resume Next
	Dim JsSetting
	Set Rs = enchiasp.Execute("SELECT id,sTitle,stype,sFileName,setting FROM ECCMS_ScriptFile WHERE ChannelID = "& ChannelID &" And id ="& CLng(Request("id")))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的系统参数！</li>"
		Set Rs = Nothing
		Exit Sub
	End If
	JsSetting = Split(Rs("setting"), ",")
	Response.Write "<table cellspacing=1 align=center cellpadding=0 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=4>添加" & sModuleName & "JS文件</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action='admin_jsfile.asp'>"
	Response.Write "	<input type=hidden name=action value=""save"">"
	Response.Write "	<input type=hidden name=ChannelID value="""& ChannelID &""">"
	Response.Write "	<input type=hidden name=id value="""& Rs("id") &""">"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right width=""20%""><b>文件说明：</b></td>"
	Response.Write "		<td class=tablerow1 width=""30%""><input type=text name=sTitle size=35 value='"& Rs("sTitle") &"'></td>"
	Response.Write "		<td class=tablerow1 align=right width=""20%""><b>JS文件名称：</b></td>"
	Response.Write "		<td class=tablerow1 width=""30%""><input type=text name=sFileName size=20 value='"& Rs("sFileName") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>显示类型：</b></td>"
	Response.Write "		<td class=tablerow2><input type=radio name=stype value=0 onClick=""stype1.style.display='';stype2.style.display='none';"""
	If Rs("stype") = 0 Then Response.Write " checked"
	Response.Write "> 列表&nbsp;&nbsp;"
	Response.Write "<input type=radio name=stype value=1 onClick=""stype2.style.display='';stype1.style.display='none';"""
	If Rs("stype") = 1 Then Response.Write " checked"
	Response.Write "> 图片</td>"
	Response.Write "		<td class=tablerow2 align=right><b>选择分类：</b></td>"
	Response.Write "		<td class=tablerow2>"
	Response.Write "<select name=""setting"" size='1'>"
	Response.Write "<option value=0>指定所有分类</option>"
	Dim strSelectClass,re
	strSelectClass = enchiasp.LoadSelectClass(ChannelID)
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	Re.Pattern="(value=)(.*)("" )"
	strSelectClass = Re.Replace(strSelectClass,"")
	Re.Pattern="({ClassID=)(.*)(}>)"
	strSelectClass = Re.Replace(strSelectClass,"value=""$2"">")
	Set Re = Nothing
	strSelectClass = Replace(strSelectClass, "value=""" & Trim(JsSetting(0)) & """", "value=""" & Trim(JsSetting(0)) & """ selected")
	Response.Write strSelectClass
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>所属专题：</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name=""setting"" size='1'>"
	Response.Write "<option value=0>不指定专题</option>"
	Dim RsObj
	Set RsObj = enchiasp.Execute("SELECT SpecialID,SpecialName FROM ECCMS_Special WHERE ChannelID="& ChannelID &" And ChangeLink=0")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("SpecialID") & """"
		If CLng(JsSetting(1)) = RsObj("SpecialID") Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("SpecialName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "		<td class=tablerow1 align=right><b>调用类型：</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name=""setting"" size='1'>"
        Response.Write "  <option value='0'"
	If CLng(JsSetting(2)) = 0 Then Response.Write " selected"
	Response.Write ">所有最新" & sModuleName & "</option>"
        Response.Write "  <option value='1'"
	If CLng(JsSetting(2)) = 1 Then Response.Write " selected"
	Response.Write ">所有推荐" & sModuleName & "</option>"
	Response.Write "  <option value='2'"
	If CLng(JsSetting(2)) = 2 Then Response.Write " selected"
	Response.Write ">所有热门" & sModuleName & "</option>"
	Response.Write "  <option value='3'"
	If CLng(JsSetting(2)) = 3 Then Response.Write " selected"
	Response.Write ">分类最新" & sModuleName & "</option>"
	Response.Write "  <option value='4'"
	If CLng(JsSetting(2)) = 4 Then Response.Write " selected"
	Response.Write ">分类推荐" & sModuleName & "</option>"
	Response.Write "  <option value='5'"
	If CLng(JsSetting(2)) = 5 Then Response.Write " selected"
	Response.Write ">分类热门" & sModuleName & "</option>"
        Response.Write "</select><font color=""#0066CC""></font>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=stype1"
	If Rs("stype") = 1 Then Response.Write " style=""display:none"""
	Response.Write ">"
	Response.Write "		<td  bgcolor=""#FFFFFF"" colspan=4>"
	Response.Write "<table width=""100%"" cellspacing=1 align=center cellpadding=3 border=0>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>图片宽度：</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='" & Trim(JsSetting(3)) & "'></td>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>图片高度：</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='" & Trim(JsSetting(4)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>最多字符数：</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='" & Trim(JsSetting(5)) & "'></td>"
	Response.Write "		<td class=tablerow1 align=right><b>最多列表数：</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='" & Trim(JsSetting(6)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>垂直边距：</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='" & Trim(JsSetting(7)) & "'></td>"
	Response.Write "		<td class=tablerow2 align=right><b>水平边距：</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='" & Trim(JsSetting(8)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>图文对齐方式：</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='left'"
	If LCase(Trim(JsSetting(9))) = "left" Then Response.Write " selected"
	Response.Write ">左对齐</option>"
	Response.Write "		<option value='right'"
	If LCase(Trim(JsSetting(9))) = "right" Then Response.Write " selected"
	Response.Write ">右对齐</option>"
	Response.Write "		<option value='middle'"
	If LCase(Trim(JsSetting(9))) = "middle" Then Response.Write " selected"
	Response.Write ">居中对齐</option>"
	Response.Write "		<option value='texttop'"
	If LCase(Trim(JsSetting(9))) = "texttop" Then Response.Write " selected"
	Response.Write ">文本上方</option>"
	Response.Write "		<option value='baseline'"
	If LCase(Trim(JsSetting(9))) = "baseline" Then Response.Write " selected"
	Response.Write ">基线</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b>调用样式：</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=20 value='" & Trim(JsSetting(10)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>连接目标：</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	Response.Write "		<option value='" & Trim(JsSetting(11)) & "'>" & Trim(JsSetting(11)) & "</option>"
	Response.Write "		<option value='_blank'>_blank</option>"
	Response.Write "		<option value='_self'>_self</option>"
	Response.Write "		<option value='_top'>_top</option>"
	Response.Write "		<option value='_parent'>_parent</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow2 align=right><b>标识符：</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=20 value='" & JsSetting(12) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>是否显示图片：</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0'"
	If CInt(JsSetting(13)) = 0 Then Response.Write " selected"
	Response.Write ">不显示</option>"
	Response.Write "		<option value='1'"
	If CInt(JsSetting(13)) = 1 Then Response.Write " selected"
	Response.Write ">显示</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b>是否显示分类：</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0'"
	If CInt(JsSetting(14)) = 0 Then Response.Write " selected"
	Response.Write ">不显示</option>"
	Response.Write "		<option value='1'"
	If CInt(JsSetting(14)) = 1 Then Response.Write " selected"
	Response.Write ">显示</option>"
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>是否显示时间：</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	Response.Write "		<option value='0'"
	If CInt(JsSetting(15)) = 0 Then Response.Write " selected"
	Response.Write ">不显示</option>"
	Response.Write "		<option value='1'"
	If CInt(JsSetting(15)) = 1 Then Response.Write " selected"
	Response.Write ">显示</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow2 align=right><b>显示时间格式：</b></td>"
	Response.Write "		<td class=tablerow2><select name=setting size=1>"
	For i = 1 To 9
		Response.Write "<option value='" & i & "'"
		If CLng(JsSetting(16)) = i Then Response.Write " selected"
		Response.Write ">"
		Response.Write enchiasp.FormatDate(Now(),i)
		Response.Write "</option>" & vbCrLf
	Next
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=stype2"
	If Rs("stype") = 0 Then Response.Write " style=""display:none"""
	Response.Write ">"
	Response.Write "		<td  bgcolor=""#FFFFFF"" colspan=4>"
	Response.Write "<table width=""100%"" cellspacing=1 align=center cellpadding=3 border=0>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 align=right><b>最多显示多少图片：</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='" & Trim(JsSetting(17)) & "'></td>"
	Response.Write "		<td class=tablerow2 align=right><b>每行显示多少图片：</b></td>"
	Response.Write "		<td class=tablerow2><input type=text name=setting size=10 value='" & Trim(JsSetting(18)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>显示最多字符数：</b></td>"
	Response.Write "		<td class=tablerow1><input type=text name=setting size=10 value='" & Trim(JsSetting(19)) & "'></td>"
	Response.Write "		<td class=tablerow1 align=right><b>是否新窗口打开：</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0'"
	If CInt(JsSetting(20)) = 0 Then Response.Write " selected"
	Response.Write ">本窗口打开</option>"
	Response.Write "		<option value='1'"
	If CInt(JsSetting(20)) = 1 Then Response.Write " selected"
	Response.Write ">新窗口打开</option>"
	Response.Write "	</select></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>图片宽度：</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='" & Trim(JsSetting(21)) & "'></td>"
	Response.Write "		<td class=tablerow2 width=""20%"" align=right><b>图片高度：</b></td>"
	Response.Write "		<td class=tablerow2 width=""30%""><input type=text name=setting size=10 value='" & Trim(JsSetting(22)) & "'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>是否显示标题名称：</b></td>"
	Response.Write "		<td class=tablerow1><select name=setting size=1>"
	Response.Write "		<option value='0'"
	If CInt(JsSetting(23)) = 0 Then Response.Write " selected"
	Response.Write ">不显示</option>"
	Response.Write "		<option value='1'"
	If CInt(JsSetting(23)) = 1 Then Response.Write " selected"
	Response.Write ">显示</option>"
	Response.Write "	</select></td>"
	Response.Write "		<td class=tablerow1 align=right><b></b></td>"
	Response.Write "		<td class=tablerow1></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow2 colspan=4 align=center>"
	Response.Write "	<input type=button name=Submit4 onclick=""javascript:history.go(-1)"" value=""返回上一页"" class=Button>&nbsp;&nbsp;"
	Response.Write "	<input type=submit name=submit3 value="" 重新设置JS文件 "" onclick=""document.selform.action.value='add';"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
	Set Rs = Nothing
End Sub

Sub SaveNewJsFile()
	If Trim(Request.Form("sTitle")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>JS文件说明不能为空！</li>"
		Exit Sub
	End If
	If Trim(Request.Form("sFileName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>JS文件名不能为空！</li>"
		Exit Sub
	End If
	If LCase(Right(Trim(Request.Form("sFileName")),3)) <> ".js" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请输入正确的JS文件名，扩展名一定要是*.js！</li>"
		Exit Sub
	End If
	SQL = "INSERT INTO ECCMS_ScriptFile (ChannelID,sTitle,stype,sFileName,setting) VALUES ("& ChannelID &",'"& enchiasp.CheckStr(Request("sTitle")) &"',"& Request("stype") &",'"& enchiasp.CheckStr(Request("sFileName")) &"','"& Request("setting") &"')"
	enchiasp.Execute(SQL)
	Response.Redirect("admin_jsfile.asp?ChannelID="& ChannelID)
End Sub

Sub SaveJsFile()
	If Trim(Request.Form("sTitle")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>JS文件说明不能为空！</li>"
		Exit Sub
	End If
	If Trim(Request.Form("sFileName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>JS文件名不能为空！</li>"
		Exit Sub
	End If
	If LCase(Right(Trim(Request.Form("sFileName")),3)) <> ".js" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请输入正确的JS文件名，扩展名一定要是*.js！</li>"
		Exit Sub
	End If
	SQL = "UPDATE ECCMS_ScriptFile SET sTitle='"& enchiasp.CheckStr(Request("sTitle")) &"',stype="& Request("stype") &",sFileName='"& enchiasp.CheckStr(Request("sFileName")) &"',setting='"& Request("setting") &"' WHERE ChannelID = "& ChannelID &" And id="& Request("id")
	enchiasp.Execute(SQL)
	Response.Redirect("admin_jsfile.asp?ChannelID="& ChannelID)
End Sub

Sub MakeJsFile()
	If Trim(Request("id")) = "" Then
		ErrMsg = "<li>错误的系统参数,请选择文件ID</li>"
		FoundErr = True
		Exit Sub
	End If
	Dim FileName,strJsContent,JsSetting
	
	On Error Resume Next
	SQL = "SELECT stype,sFileName,setting FROM ECCMS_ScriptFile WHERE ChannelID="& ChannelID &" And id in("& Request("id") &")"
	Set Rs = enchiasp.Execute(SQL)
	If Not(Rs.BOF And Rs.EOF) Then
		Do While Not Rs.EOF
			JsSetting = Split(Rs("setting"), ",")
			enchiasp.InstallDir = enchiasp.SiteUrl & "/"
			If Rs("stype") = 1 Then
				Select Case CInt(enchiasp.modules)
				Case 1
					strJsContent = HTML.LoadArticlePic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				Case 2
					strJsContent = HTML.LoadSoftPic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				Case 3
					strJsContent = HTML.LoadShopPic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				Case 5
					strJsContent = HTML.LoadFlashPic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				End Select
			Else
				Select Case CInt(enchiasp.modules)
				Case 1
					strJsContent = HTML.NewsPictureAndText(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(3)),Trim(JsSetting(4)),Trim(JsSetting(5)),Trim(JsSetting(6)),Trim(JsSetting(7)),Trim(JsSetting(8)),Trim(JsSetting(9)),Trim(JsSetting(10)),Trim(JsSetting(11)),JsSetting(12),Trim(JsSetting(13)),Trim(JsSetting(14)),Trim(JsSetting(15)),Trim(JsSetting(16)))
				Case 2
					strJsContent = HTML.SoftPictureAndText(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(3)),Trim(JsSetting(4)),Trim(JsSetting(5)),Trim(JsSetting(6)),Trim(JsSetting(7)),Trim(JsSetting(8)),Trim(JsSetting(9)),Trim(JsSetting(10)),Trim(JsSetting(11)),JsSetting(12),Trim(JsSetting(13)),Trim(JsSetting(14)),Trim(JsSetting(15)),Trim(JsSetting(16)))
				Case 3
					strJsContent = HTML.LoadShopPic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				Case 5
					strJsContent = HTML.LoadFlashPic(ChannelID,Trim(JsSetting(0)),Trim(JsSetting(1)),Trim(JsSetting(2)),Trim(JsSetting(17)),Trim(JsSetting(18)),Trim(JsSetting(19)),Trim(JsSetting(20)),Trim(JsSetting(21)),Trim(JsSetting(22)),Trim(JsSetting(23)))
				End Select
			End If
			
			strJsContent = "document.write ('"& fixjs(strJsContent) &"');"
			FileName = "../"& enchiasp.ChannelDir &"js/"& Rs("sFileName")
			enchiasp.CreatedTextFile FileName,strJsContent
			Rs.movenext
		Loop
	End If
	Set Rs = Nothing
	Succeed("<li>恭喜您！生成JS文件成功。</li>")
End Sub

Sub DeleteJsFile()
	If Trim(Request("id")) <> "" Then
		On Error Resume Next
		Set Rs = enchiasp.Execute("SELECT sFileName FROM ECCMS_ScriptFile WHERE ChannelID = "& ChannelID &" And id=" & CLng(Request("id")))
		If Not(Rs.BOF And Rs.EOF) Then
			enchiasp.FileDelete("../"& enchiasp.ChannelDir &"js/"& Rs("sFileName"))
		End If
		Set Rs = Nothing
		enchiasp.Execute ("DELETE FROM ECCMS_ScriptFile WHERE ChannelID = "& ChannelID &" And id=" & CLng(Request("id")))
		Response.Redirect  Request.ServerVariables("HTTP_REFERER")
	Else
		ErrMsg = "<li>错误的系统参数,请选择要删除的文件ID</li>"
		FoundErr = True
		Exit Sub
	End If
End Sub
Sub DemoJsFile()
	Dim JsFileName
	Set Rs = enchiasp.Execute("SELECT sFileName FROM ECCMS_ScriptFile WHERE ChannelID = "& ChannelID &" And id=" & CLng(Request("id")))
	If Not(Rs.BOF And Rs.EOF) Then
		Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
		Response.Write "	<tr>"
		Response.Write "		<th>" & sModuleName & "JS文件调用演示</th>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 colspan=4 align=center>"
		JsFileName = "<script src="""& enchiasp.InstallDir & enchiasp.ChannelDir &"js/"& Rs("sFileName") &""" type=""text/javascript""></script>"
		Response.Write Server.HTMLEncode(JsFileName)
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow1>"
		Response.Write JsFileName
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<tr>"
		Response.Write "		<td class=tablerow2 align=center>"
		Response.Write "	<input type=button name=Submit4 onclick=""javascript:history.go(-1)"" value=""返回上一页"" class=Button>&nbsp;&nbsp;"
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "</table>"
	End If
End Sub

%>