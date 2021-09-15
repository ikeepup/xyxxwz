<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="include/MenuCode.Asp"-->
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
Dim Action,TitleColor,ChannelDir,strModules,strOption,ChannelPath
Dim RsObj,i,Flag,HtmlFileDir,AddContentLink,ClassDir,strClassDir
Dim moduleid,UseHtml,IsCreateHtml,strClass
dim addyemianlink,edityemianlink
If Request("ChannelID") = 0 Or Request("ChannelID") = "" Then
	ErrMsg = "<li>Sorry！错误的系统参数,请选择正确的连接方式。</li>"
	Response.Redirect("showerr.asp?action=error&message=" & Server.URLEncode(ErrMsg) & "")
	Response.End
Else
	ChannelID = CInt(Request("ChannelID"))
End If
Set Rs = enchiasp.Execute("select ChannelDir,modules,IsCreateHtml from ECCMS_Channel Where ChannelID = "& ChannelID)
ChannelDir = Rs(0)
moduleid = Rs("modules")
IsCreateHtml = Rs("IsCreateHtml")
Select Case Rs("modules")
	Case 1:strModules = "article"
	Case 2:strModules = "soft"
	Case 3:strModules = "shop"
	Case 4:strModules = "flash"
	Case 6:strModules = "yemian"
	case 7:strModules = "job"

Case Else
	strModules = "article"
End Select

if rs("modules")<>7 then
	ErrMsg = "<li>Sorry！错误的系统参数,请选择正确的连接方式。</li>"
	Response.Redirect("showerr.asp?action=error&message=" & Server.URLEncode(ErrMsg) & "")
	Response.End

end if

Set Rs = Nothing


ChannelPath = enchiasp.InstallDir & ChannelDir
Flag = "AdminClass" & ChannelID
AddContentLink = "admin_" & strModules & ".asp?action=add&ChannelID=" & ChannelID & "&ClassID="
addyemianlink="admin_yemian.asp?action=add&ChannelID=" & ChannelID & "&ClassID="
edityemianlink="admin_yemian.asp?action=edit&ChannelID=" & ChannelID & "&ClassID="
%>
<script language = "JavaScript">
function BatchAddClass(){
	if(document.myform.BatchID.checked==true){
		document.myform.BatchClassName.disabled=false;
		document.myform.ClassName.disabled=true;
		BatchClass.style.display='';
	}
	else{
		document.myform.BatchClassName.disabled=true;
		document.myform.ClassName.disabled=false;
		BatchClass.style.display='none';
	}
}

function ClassSetting(n){
	if (n == 1){
		ClassSetting1.style.display='none';
		ClassSetting2.style.display='';
		ClassSetting3.style.display='';
	}
	else{
		ClassSetting1.style.display='';
		ClassSetting2.style.display='none';
		ClassSetting3.style.display='none';
	}
}
</script>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
	<tr>
		<th colspan="2"><%=sModuleName%>分类管理</th>
	</tr>
	<tr>
		<td width="100%" class="TableRow2" colspan="2"><b>频道设置选项：</b><a href="admin_channel.asp">频道设置首页</a> 
		| <a href="admin_channel.asp?action=add">添加频道</a> |
<%
Dim Rsm
Set Rsm = enchiasp.Execute("Select ChannelID,ModuleName From ECCMS_Channel where ChannelType < 2 Order By orders Asc")
Do While Not Rsm.EOF
	Response.Write "<a href=admin_channel.asp?action=edit&ChannelID="
	Response.Write Rsm("ChannelID")
	Response.Write ">"
	Response.Write Rsm("ModuleName")
	Response.Write "设置</a> | "
	Rsm.movenext
Loop
Set Rsm = Nothing
%>
		</td>
	</tr>
</table>
<br>
<%

If Not ChkAdmin("Adminclassjob" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
Action = enchiasp.RemoveBadCharacters(Request("action"))
Select Case LCase(Action)
Case "savenew"
	Call savenew
Case "savedit"
	Call savedit
Case "add"
	Call ClassAdd
Case "edit"
	Call ClassEdit
Case "del"
	Call DelClass
Case "deldir"
	Call DelClassDir
Case "orders"
	Call orders
Case "neworders"
	Call updateorders
Case "restore"
	Call RestoreClass
Case "classorders"
	Call classorders
Case "newclassorders"
	Call updateclassorders
Case "jsmenu"
	Call CreationJsMenu
Case "alljs"
	Call CreationAllJsFile
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
	Response.Write " <table align=center class=""tableBorder"" cellspacing=""1"" cellpadding=""2"">"
	Response.Write " <tr>"
	Response.Write " <th width=""3%"">选择</th>"
	Response.Write " <th width=""35%"">"& sModuleName &"分类 </th>"
	Response.Write " <th width=""43%"">管理选项</th>"
	Response.Write " <th noWrap width=""9%"">连接性质</th>"
	Response.Write "</tr>"
	SQL = "SELECT * FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" ORDER BY rootid,orders"
	Set Rs = Server.CreateObject("adodb.recordset")
	Rs.Open SQL, Conn, 1, 1
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.bof And Rs.EOF Then
		Response.Write " <tr> <td align=center colspan=4 class=""TableRow1"">您还没有添加任何分类！</td></tr>"
	End If
	Response.Write "	<form name=selform method=post action=""admin_create" & strModules & ".asp"">"
	Response.Write "<input type=hidden name=action value='list'>"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "<input type=hidden name=Field value='2'>"
	Response.Write "<input type=hidden name=stype value='1'>"
	i = 0
	Do While Not Rs.EOF
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
		Response.Write " <tr>"
		Response.Write " <td align=center " & strClass & ">"
		Response.Write "<input type=checkbox name=""classid"" value=""" & Rs("ClassID") & """>"
		Response.Write " </td>"
		Response.Write " <td " & strClass & ">"
		Response.Write " "
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font>"
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;<font color=""#666666"">│</font>"
			Next
			Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font> "
		End If
		If Rs("parentid") = 0 Then Response.Write ("<b>")
		Response.Write enchiasp.ReadFontMode(Rs("classname"),Rs("ColorModes"),Rs("FontModes"))
		If Rs("child") > 0 Then Response.Write "(" & Rs("child") & ")"
		Response.Write " </td>"
		Response.Write " <td class=""TableRow2"" align=center>"
		if rs("isdanyemian")=1 then
			Response.Write "<a href="""
			Response.Write AddyemianLink
			Response.Write Rs("classid")
			Response.Write """>添加内容</a> | "
			
			Response.Write "<a href="""
			Response.Write edityemianLink
			Response.Write Rs("classid")
			Response.Write """>编辑内容</a> | "

		else
			Response.Write "<a href="""
			Response.Write AddContentLink
			Response.Write Rs("classid")
			Response.Write """>添加内容</a> | "
		end if
		
		
		Response.Write "<a href=""?action=add&ChannelID="&ChannelID&"&editid="
		Response.Write Rs("classid")
		Response.Write """>添加分类</a>"
		Response.Write " | <a href=""?action=edit&ChannelID="&ChannelID&"&editid="
		Response.Write Rs("classid")
		Response.Write """>分类设置</a>"
		Response.Write " |"
		Response.Write " "
		If Rs("child") = 0 Then
			Response.Write " <a href=""?action=del&ChannelID="&ChannelID&"&editid="
			Response.Write Rs("classid")
			Response.Write """ onclick=""{if(confirm('删除将包括该分类的所有文章，确定删除吗?')){return true;}return false;}"">删除分类 "
			Response.Write " "
		Else
			Response.Write "<a href=""#"" onclick=""{if(confirm('该分类含有下属分类，必须先删除其下属分类方能删除本分类！')){return true;}return false;}"">"
			Response.Write " 删除分类</a>  "
			Response.Write " "
		End If
		
		Response.Write " </td>"
		Response.Write " <td align=center " & strClass & ">"
		If Rs("TurnLink") <> 0 Then
			Response.Write "<font color=red>转向连接</font>"
		Else
			Response.Write "<font color=blue>系统连接</font>"
		End If
		Response.Write " </td>"
		Response.Write "</tr>"
		Rs.movenext
		i = i + 1
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write " <tr>"
	Response.Write "<td colspan=4 class=TableRow2>"
	Response.Write "<input class=""Button"" type=""button"" name=""chkall"" value=""全选"" onClick=""CheckAll(this.form)"">"
	Response.Write "<input class=Button type=""button"" name=""chksel"" value=""反选"" onClick=""ContraSel(this.form)"">" & vbNewLine
	Response.Write " </td>"
	Response.Write "</tr></form>"
	Response.Write "</table>"
End Sub

Private Sub ClassAdd()
	Dim NewClassID
	SQL = "SELECT MAX(ClassID) FROM ECCMS_Classify"
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		NewClassID = 1
	Else
		NewClassID = Rs(0) + 1
	End If
	If IsNull(NewClassID) Then NewClassID = 1
	Rs.Close
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
	<tr>
		<th colspan="2">添加<%=sModuleName%>分类</th>
	</tr>
	<form name=myform method="POST" action="?action=savenew">
	<input type="hidden" name="NewClassID" value="<%=NewClassID%>">
	<input type="hidden" name="ChannelID" value="<%=ChannelID%>">
	<tr>
		<td width="20%" class="TableRow2"><strong><%=sModuleName%>分类名称：</strong></td>
		<td width="80%" class="TableRow1">
		<input type="text" name="ClassName" id="ClassName" size="35">
		</td>
	</tr>
	<tr>
		<td class="TableRow2"><strong><%=sModuleName%>分类标题模式：</strong></td>
		<td class="TableRow1">颜色：
		<select size="1" name="ColorModes">
		<option value="0">请选择颜色</option>
<%
	TitleColor = "," & enchiasp.InitTitleColor
	TitleColor = Split(TitleColor, ",")
	For i = 1 To UBound(TitleColor)
		Response.Write ("<option style=""background-color:"& TitleColor(i) &";color: "& TitleColor(i) &""" value='"& i &"'>"& TitleColor(i) &"</option>")
	Next
%>
		</select> 字体：
		<select size="1" name="FontModes">
		<option value="0">请选择字体</option>
		<option value="1">粗体</option>
		<option value="2">斜体</option>
		<option value="3">下划线</option>
		<option value="4">粗体+斜体</option>
		<option value="5">粗体+下划线</option>
		<option value="6">斜体+下划线</option>
		
		</select></td>
	</tr>
	<tr>
		<td class="TableRow2"><strong><%=sModuleName%>分类注释：</strong></td>
		<td class="TableRow1">
		<input type="text" name="Readme" size="60"> </td>
	</tr>
	<tr>
		<td class="TableRow2"><strong>所属分类：</strong></td>
		<td class="TableRow1">
<%
	Response.Write " <select name=""class"">"
	Response.Write "<option value=""0"">做为一级分类</option>"
	SQL = "SELECT classid,depth,ClassName FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" ORDER BY rootid,orders"
	Set Rs = enchiasp.Execute(SQL)
	Do While Not Rs.EOF
		Response.Write "<option value=""" & Rs("classid") & """ "
		If Request("editid") <> "" And CLng(Request("editid")) = Rs("classid") Then Response.Write "selected"
		Response.Write ">"
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├ "
		End If
		Response.Write Rs("ClassName") & "</option>" & vbCrLf
		Rs.movenext
	Loop
	Rs.Close
	Response.Write "</select>"
	Set Rs = Nothing
%>
		</td>
	</tr>
	<tr>
		<td class="TableRow2"><strong>连接目标：</strong></td>
		<td class="TableRow1">
		<input type="radio" value="0" checked name="LinkTarget"> 本窗口打开&nbsp;&nbsp; 
		<input type="radio" name="LinkTarget" value="1"> 新窗口打开</td>
	</tr>
	<tr>
		<td class="TableRow2"><b>是否为单页面图文</b></td>
		<td class="TableRow1">
		<input type="radio" name="isdanyemian" value="0"   checked> 否&nbsp;&nbsp; 
		<input type="radio" name="isdanyemian" value="1"  > 是</td>
	</tr>
	<tr>
		<td class="TableRow2"><strong>是否转向连接：</strong></td>
		<td class="TableRow1">
		<input type="radio" name="TurnLink" value="0"  onClick="ClassSetting(1)" checked> 否&nbsp;&nbsp; 
		<input type="radio" name="TurnLink" value="1"  onClick="ClassSetting(2)"> 是</td>
	</tr>
	<tr>
		<td class="TableRow2"><strong>分类目录：</strong></td>
		<td class="TableRow1"><input type="text" name="ClassDir" size="15" value="<%=NewClassID%>"> <br><font color=blue>一级分类相对于此频道目录，N级分类相对于上级分类目录，可以是多级目录,如：html/asp请认真填写。</font></td>
	</tr>
	<tr id=ClassSetting1 style="display:none">
		<td class="TableRow2"><strong>转向连接URL：</strong></td>
		<td class="TableRow1"><input type="text" name="TurnLinkUrl" size="45" value="<%=enchiasp.SiteUrl%>"></td>
	</tr>
	<tr >
		<td class="TableRow2"><strong>可以添加内容：</strong></td>
		<td class="TableRow1">
		<input type="radio" name="isallow" value="0" > 否&nbsp;&nbsp; 
		<input type="radio" name="isallow" value="1" checked> 是</td>
	</tr>
	<tr id=ClassSetting2 style="display:">
		<td class="TableRow2"><strong>用户组：</strong></td>
		<td class="TableRow1"><select size="1" name="UserGroup">
<%
	Set Rs = enchiasp.Execute("SELECT GroupName,Grades FROM ECCMS_UserGroup ORDER BY Groupid")
	Do While Not Rs.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & Rs("Grades") & """"
		If Rs("Grades") = 0 Then Response.Write " selected"
		Response.Write ">"
		Response.Write Rs("GroupName")
		Response.Write "</option>" & vbCrLf
		Rs.movenext
	Loop
	Set Rs = Nothing
%>		</select></td>
	</tr>
	<tr id=ClassSetting3 style="display:">
		<td class="TableRow2"><strong>使用模板：</strong></td>
		<td class="TableRow1"><select size="1" name="skinid">
<%
	Response.Write "		<option value=""0"" selected>使用默认模板</option>" & vbCrLf
	SQL = "SELECT skinid,page_name,isDefault FROM ECCMS_Template WHERE pageid = 0 ORDER BY TemplateID"
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		Response.Write "		<option value=""0"">您还没有添加任何模板文件</option>" & vbCrLf
	Else
		Do While Not Rs.EOF
			Response.Write "		<option value=""" & Rs("skinid") & """"
			'If Rs("isDefault") = 1 Then Response.Write " selected"
			Response.Write ">"
			Response.Write Rs("page_name")
			Response.Write "</option>" & vbCrLf
			Rs.movenext
		Loop
	End IF
	Set Rs = Nothing
%>		</select></td>
	</tr>
	<tr>
		<td class="TableRow2">　</td>
		<td class="TableRow1">
		<p align="center"><input type="button" onclick="javascript:history.go(-1)" value="返回上一页" name="B1" class=Button>&nbsp;&nbsp;
		<input type="submit" value="保存设置" name="B2" class=Button></td>
	</tr>
	</form>
</table>
<%
End Sub

Private Sub ClassEdit()
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_Classify WHERE ChannelID = " & ChannelID & " And ClassID = " & Request("editid"))
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = "数据库出现错误,没有此站点栏目!"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
	<tr>
		<th colspan="2">添加<%=sModuleName%>分类</th>
	</tr>
	<form name=myform method="POST" action="?action=savedit">
	<input type="hidden" name="editid" value="<%=Request("editid")%>">
	<input type="hidden" name="ChannelID" value="<%=ChannelID%>">
	<tr>
		<td width="20%" class="TableRow2"><strong><%=sModuleName%>分类名称：</strong></td>
		<td width="80%" class="TableRow1">
		<input type="text" name="ClassName" id="ClassName" size="35" value="<%=Rs("ClassName")%>">
		</td>
	</tr>
	<tr>
		<td class="TableRow2"><strong><%=sModuleName%>分类标题模式：</strong></td>
		<td class="TableRow1">颜色：
		<select size="1" name="ColorModes">
		<option value="0"<%If Rs("ColorModes") = 0 Then Response.Write (" selected")%>>请选择颜色</option>
<%
	TitleColor = "," & enchiasp.InitTitleColor
	TitleColor = Split(TitleColor, ",")
	For i = 1 To UBound(TitleColor)
		Response.Write ("<option style=""background-color:"& TitleColor(i) &";color: "& TitleColor(i) &""" value='"& i &"'")
		If Rs("ColorModes") = i Then Response.Write (" selected")
		Response.Write (">"& TitleColor(i) &"</option>")
	Next
%>
		</select> 字体：
		<select size="1" name="FontModes">
		<option value="0"<%If Rs("FontModes") = 0 Then Response.Write (" selected")%>>请选择字体</option>
		<option value="1"<%If Rs("FontModes") = 1 Then Response.Write (" selected")%>>粗体</option>
		<option value="2"<%If Rs("FontModes") = 2 Then Response.Write (" selected")%>>斜体</option>
		<option value="3"<%If Rs("FontModes") = 3 Then Response.Write (" selected")%>>下划线</option>
		<option value="4"<%If Rs("FontModes") = 4 Then Response.Write (" selected")%>>粗体+斜体</option>
		<option value="5"<%If Rs("FontModes") = 5 Then Response.Write (" selected")%>>粗体+下划线</option>
		<option value="6"<%If Rs("FontModes") = 6 Then Response.Write (" selected")%>>斜体+下划线</option>
		
		</select></td>
	</tr>
	<tr>
		<td class="TableRow2"><strong><%=sModuleName%>分类注释：</strong></td>
		<td class="TableRow1">
		<input type="text" name="Readme" size="60" value="<%=Rs("Readme")%>"> </td>
	</tr>
	<tr>
		<td class="TableRow2"><strong>所属分类：</strong></td>
		<td class="TableRow1">
<%
	Response.Write " <select name=""class"">"
	Response.Write "<option value=""0"">做为一级分类</option>"
	SQL = "SELECT classid,depth,ClassName FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" ORDER BY rootid,orders"
	Set RsObj = enchiasp.Execute(SQL)
	Do While Not RsObj.EOF
		Response.Write "<option value=""" & RsObj("classid") & """ "
		If CLng(Rs("parentid")) = RsObj("classid") Then Response.Write "selected"
		Response.Write ">"
		If RsObj("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
		If RsObj("depth") > 1 Then
			For i = 2 To RsObj("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├ "
		End If
		Response.Write RsObj("ClassName") & "</option>" & vbCrLf
		RsObj.movenext
	Loop
	RsObj.Close
	Response.Write "</select>"
	Set RsObj = Nothing
%>
		</td>
	</tr>
	<tr>
		<td class="TableRow2"><strong>连接目标：</strong></td>
		<td class="TableRow1">
		<input type="radio" value="0" name="LinkTarget"<%If Rs("LinkTarget") = 0 Then Response.Write " checked"%>> 本窗口打开&nbsp;&nbsp; 
		<input type="radio" name="LinkTarget" value="1"<%If Rs("LinkTarget") = 1 Then Response.Write " checked"%>> 新窗口打开</td>
	</tr>
	<tr>
		<td class="TableRow2"><b>是否为单页面图文</b></td>
		<td class="TableRow1">
		<input type="radio" name="isdanyemian" value="0" <%If Rs("isdanyemian") = 0 Then Response.Write " checked"%> > 否&nbsp;&nbsp; 
		<input type="radio" name="isdanyemian" value="1" <%If Rs("isdanyemian") = 1 Then Response.Write " checked"%> > 是</td>
	</tr>
	<tr>
		<td class="TableRow2"><strong>是否转向连接：</strong></td>
		<td class="TableRow1">
		<input type="radio" name="TurnLink" value="0" onClick="ClassSetting(1)"<%If Rs("TurnLink") = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp; 
		<input type="radio" name="TurnLink" value="1" onClick="ClassSetting(2)"<%If Rs("TurnLink") = 1 Then Response.Write " checked"%>> 是</td>
	</tr>
	<tr>
		<td class="TableRow2"><strong>分类目录：</strong></td>
		<td class="TableRow1"><input type="text" name="ClassDir" size="15" value="<%=Rs("ClassDir")%>"> <font color=red>相对于此频道目录，请不要随意修改，一但修改需要生成所有的HTML文件；谨用！</font></td>
	</tr>
	<tr id=ClassSetting1<%If Rs("TurnLink") = 0 Then Response.Write " style=""display:none"""%>>
		<td class="TableRow2"><strong>转向连接URL：</strong></td>
		<td class="TableRow1"><input type="text" name="TurnLinkUrl" size="45" value="<%=Rs("TurnLinkUrl")%>"></td>
	</tr>
	<tr>
		<td class="TableRow2"><strong>可以添加内容：</strong></td>
		<td class="TableRow1">
		<input type="radio" name="isallow" value="0" <%If Rs("isallow") = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp; 
		<input type="radio" name="isallow" value="1" <%If Rs("isallow") = 1 Then Response.Write " checked"%>> 是</td>
	</tr>
	<tr id=ClassSetting2<%If Rs("TurnLink") <> 0 Then Response.Write " style=""display:none"""%>>
		<td class="TableRow2"><strong>用户组：</strong></td>
		<td class="TableRow1"><select size="1" name="UserGroup">
<%
	Set RsObj = enchiasp.Execute("SELECT GroupName,Grades FROm ECCMS_UserGroup ORDER BY Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & """"
		If Rs("UserGroup") = RsObj("Grades") Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>		</select></td>
	</tr>
	<tr id=ClassSetting3<%If Rs("TurnLink") <> 0 Then Response.Write " style=""display:none"""%>>
		<td class="TableRow2"><strong>使用模板：</strong></td>
		<td class="TableRow1"><select size="1" name="skinid">
<%
	Response.Write "		<option value=""0"""
	If Rs("skinid") = 0 Then Response.Write " selected"
	Response.Write ">使用默认模板</option>" & vbCrLf
	SQL = "SELECT skinid,page_name,isDefault FROM ECCMS_Template WHERE pageid = 0 ORDER BY TemplateID"
	Set RsObj = enchiasp.Execute(SQL)
	If RsObj.bof And RsObj.EOF Then
		Response.Write "		<option value=""0"">您还没有添加任何模板文件</option>" & vbCrLf
	Else
		Do While Not RsObj.EOF
			Response.Write "		<option value=""" & RsObj("skinid") & """"
			If Rs("skinid") = RsObj("skinid") Then Response.Write " selected"
			Response.Write ">"
			Response.Write RsObj("page_name")
			Response.Write "</option>" & vbCrLf
			RsObj.movenext
		Loop
	End IF
	Set RsObj = Nothing
%>		</select></td>
	</tr>
	<tr>
		<td class="TableRow2">　</td>
		<td class="TableRow1">
		<p align="center"><input type="button" onclick="javascript:history.go(-1)" value="返回上一页" name="B1" class=Button>&nbsp;&nbsp;
		<input type="submit" value="保存设置" name="B2" class=Button></td>
	</tr>
	</form>
</table>
<%
Set Rs = Nothing
End Sub

Private Sub CheckSave()
	If Trim(Request("classname")) = "" Then
		ErrMsg = ErrMsg + "<li>请输入分类名称。</li>"
		Founderr = True
	End If
	If Not IsNumeric(Request("class")) Then
		ErrMsg = ErrMsg + "<li>请选择所属分类。</li>"
		Founderr = True
	End If
	If Trim(Request("Readme")) = "" Then
		ErrMsg = ErrMsg + "<li>请输入分类说明。</li>"
		Founderr = True
	End If
	If Trim(Request.Form("TurnLink")) = "" Then
		ErrMsg = ErrMsg + "<li>转向连接的URL不能为空。</li>"
		Founderr = True
	End If
	If Trim(Request.Form("LinkTarget")) = "" Then
		ErrMsg = ErrMsg + "<li>请选择连接目标。</li>"
		Founderr = True
	End If
	If Trim(Request.Form("ColorModes")) = "" Then
		ErrMsg = ErrMsg + "<li>请选择标题颜色。</li>"
		Founderr = True
	End If
	If Trim(Request.Form("FontModes")) = "" Then
		ErrMsg = ErrMsg + "<li>请选择标题字体。</li>"
		Founderr = True
	End If
	If CInt(Request.Form("TurnLink")) = 1 Then
		If Request("TurnLinkUrl") = "" Then
			ErrMsg = ErrMsg + "<li>转向连接的URL不能为空。</li>"
			Founderr = True
		End If
	Else
		If Request("UserGroup") = "" Then
			ErrMsg = ErrMsg + "<li>请选择用户组。</li>"
			Founderr = True
		End If
		If Request("skinid") = "" Then
			ErrMsg = ErrMsg + "<li>请选择模板。</li>"
			Founderr = True
		End If
	End If
	If Len(Request.Form("ChannelName")) => 25 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>分类名称名称不能超过50个字符！</li>"
	End If
	If Len(Request.Form("Readme")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>栏目注释不能超过200个字符！</li>"
	End If
	If Len(Request.Form("ClassDir")) = 0 And Request.Form("TurnLink") = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>分类目录不能为空！</li>"
	End If
	If Not enchiasp.IsValidChar(Request.Form("ClassDir")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>目录名中含有非法字符或者中文字符！</li>"
	End If
	strClassDir = Replace(Replace(Replace(Request.Form("ClassDir"), "\","/"), " ",""), "'","")
	If Right(strClassDir, 1) <> "/" Then
		strClassDir = strClassDir
	Else
		strClassDir = Left(strClassDir,Len(strClassDir)-1)
	End If
	If Left(strClassDir, 1) = "/" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>目录前面不能有“/”，请认真填写分类目录！</li>"
	End If
	
End Sub

Private Sub savenew()
	Dim classid
	Dim rootid
	Dim ParentID
	Dim depth
	Dim orders
	Dim Maxrootid
	Dim ParentStr
	Dim ChildStr
	Dim neworders
	dim isallow
	dim isdanyemian
	'保存添加分类信息
	CheckSave
	If Founderr = True Then Exit Sub
	If Request("class") <> "0" Then
		SQL = "SELECT rootid,classid,depth,orders,ParentStr,TurnLink,HtmlFileDir FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & Request("class")
		Set Rs = enchiasp.Execute (SQL)
		rootid = Rs(0)
		ParentID = Rs(1)
		depth = Rs(2)
		orders = Rs(3)
		If depth + 1 > 20 Then
			ErrMsg = "<li>本系统限制最多只能有20级子分类</li>"
			Founderr = True
			Exit Sub
		End If
		If Rs("TurnLink") = 1 Then
			ErrMsg = "<li>该分类是外部连接，您不能指定该分类作为所属分类</li>"
			Founderr = True
			Exit Sub
		End If
		ParentStr = Rs(4)
		HtmlFileDir = Rs("HtmlFileDir")
		Rs.Close
		neworders = orders
		SQL = "SELECT MAX(orders) FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And ParentID=" & Request("class")
		Set Rs = enchiasp.Execute (SQL)
		If Not (Rs.EOF And Rs.bof) Then
			neworders = Rs(0)
		End If
		If IsNull(neworders) Then neworders = orders
		Rs.Close
		enchiasp.Execute ("UPDATE ECCMS_Classify SET orders=orders+1 WHERE ChannelID = "& ChannelID &" And orders>" & CInt(neworders) & "")
	Else
		SQL = "SELECT MAX(rootid) FROM ECCMS_Classify WHERE ChannelID = "& ChannelID
		Set Rs = enchiasp.Execute (SQL)
		Maxrootid = Rs(0) + 1
		If IsNull(Maxrootid) Then Maxrootid = 1
		Rs.Close
	End If
	SQL = "SELECT classid FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & Request("newclassid")
	Set Rs = enchiasp.Execute (SQL)
	If Not (Rs.EOF And Rs.bof) Then
		ErrMsg = "<li>您不能指定和别的分类一样的序号。</li>"
		Founderr = True
		Exit Sub
	Else
		classid = Request("newclassid")
	End If
	Rs.Close
	Set Rs = Server.CreateObject("adodb.recordset")
	SQL = "SELECT * FROM ECCMS_Classify"
	Rs.Open SQL, Conn, 1, 3
	Rs.addnew
	If Request("class") <> "0" Then
		Rs("depth") = depth + 1
		Rs("rootid") = rootid
		Rs("orders") = neworders + 1
		Rs("parentid") = Request.Form("class")
		HtmlFileDir = HtmlFileDir & strClassDir & "/"
		If ParentStr = "0" Then
			Rs("ParentStr") = Request.Form("class")
		Else
			Rs("ParentStr") = ParentStr & "," & Request.Form("class")
		End If
	Else
		Rs("depth") = 0
		Rs("rootid") = Maxrootid
		Rs("orders") = 0
		Rs("parentid") = 0
		Rs("ParentStr") = 0
		HtmlFileDir = strClassDir & "/"
	End If
	Rs("ChannelID") = ChannelID
	Rs("ColorModes") = Trim(Request.Form("ColorModes"))
	Rs("FontModes") = Trim(Request.Form("FontModes"))
	Rs("child") = 0
	Rs("ChildStr") = Trim(Request.Form("newclassid"))
	Rs("LinkTarget") = Trim(Request.Form("LinkTarget"))
	Rs("TurnLink") = Trim(Request.Form("TurnLink"))
	Rs("TurnLinkUrl") = Trim(Request.Form("TurnLinkUrl"))
	Rs("UserGroup") = Trim(Request.Form("UserGroup"))
	Rs("HtmlFileDir") = Trim(HtmlFileDir)
	Rs("ClassDir") = Trim(strClassDir)
	Rs("classid") = Trim(Request.Form("newclassid"))
	Rs("classname") = enchiasp.ChkFormStr(Request.Form("classname"))
	Rs("readme") = Trim(Request.Form("readme"))
	Rs("skinid") = Trim(Request.Form("skinid"))
	Rs("UseHtml") = 1
	Rs("ShowCount") = 0
	Rs("isUpdate") = 1
	Rs("isallow") = Trim(Request.Form("isallow"))
	Rs("isdanyemian") = Trim(Request.Form("isdanyemian"))
	Rs.Update
	Rs.Close
	If Request("class") <> "0" Then
		Dim nClassID
		ParentStr = ParentStr & "," & Request.Form("class")
		nClassID = Trim(Request.Form("newclassid"))
		SQL = "SELECT classid,ParentStr,ChildStr FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid in (" & ParentStr & ")"
		Set Rs = enchiasp.Execute (SQL)
		Do While Not Rs.EOF
			ChildStr = Rs("ChildStr") & "," & nClassID
			enchiasp.Execute ("UPDATE ECCMS_Classify SET ChildStr='"&ChildStr&"' WHERE ChannelID = "& ChannelID &" And classid = " & Rs("classid"))
		Rs.movenext
		Loop
		Rs.Close
		If depth > 0 Then
			enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child+1 where ChannelID = "& ChannelID &" And classid in (" & ParentStr & ")")
		Else
			enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child+1 where ChannelID = "& ChannelID &" And classid=" & Request("class"))
		End If
	End If
	Dim LocalPath
	If CInt(enchiasp.IsCreateHtml) <> 0 And CInt(Request.Form("TurnLink")) = 0 Then
		LocalPath = enchiasp.InstallDir & ChannelDir & HtmlFileDir
		enchiasp.CreatPathEx(LocalPath)
	End If
	Call RemoveCache
	SucMsg = "<li>恭喜您！分类添加成功。</li>"
	Set Rs = Nothing
	Succeed(SucMsg)
End Sub

Private Sub savedit()
	Dim newclassid
	Dim Maxrootid
	Dim ParentID
	Dim depth
	Dim Child
	Dim ParentStr
	Dim rootid
	Dim iparentid
	Dim iParentStr
	Dim trs
	Dim brs
	Dim mrs
	Dim Rsc
	Dim Rss
	Dim k
	Dim nParentStr
	Dim mParentStr
	Dim ParentSql
	Dim ChildStr
	Dim nChildStr
	Dim ArrChildStr
	Dim ii
	Dim ClassCount
	dim isallow
	dim isdanyemian
	'保存编辑分类信息
	If CLng(Request("editid")) = CLng(Request("class")) Then
		ErrMsg = "<li>所属分类不能指定自己</li>"
		Founderr = True
		Exit Sub
	End If
	CheckSave
	If Founderr = True Then Exit Sub
	If CLng(Request("class")) <> 0 Then
		HtmlFileDir = enchiasp.Execute("SELECT HtmlFileDir FROM [ECCMS_Classify] WHERE ChannelID = "& ChannelID &" And classid=" & Request("class"))(0)
		HtmlFileDir = HtmlFileDir & strClassDir & "/"
	End If
	Set Rs = Server.CreateObject("adodb.recordset")
	SQL = "SELECT * FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & Request("editid")
	Rs.Open SQL, Conn, 1, 3
	newclassid = Rs("classid")
	ParentID = Rs("parentid")
	iparentid = Rs("parentid")
	ParentStr = Rs("ParentStr")
	ChildStr = Rs("ChildStr")
	ClassDir = Rs("ClassDir")
	depth = Rs("depth")
	Child = Rs("child")
	rootid = Rs("rootid")
	If CLng(Request("class")) = 0 Then
		HtmlFileDir = strClassDir & "/"
	End If
	If Child <> 0 And LCase(ClassDir) <> LCase(strClassDir) Then
		ErrMsg = "<li>对不起！该分类中有下属分类不能修改分类目录！</li>"
		Founderr = True
		Exit Sub
	End If
	If Child <> 0 And ParentID <> Clng(Request("class")) Then
		ErrMsg = "<li>对不起！该分类中有下属分类不能移动，请先移动其下属分类。</li>"
		Founderr = True
		Exit Sub
	End If
	
	If ParentID = 0 Then
		If CLng(Request("class")) <> 0 Then
			Set trs = enchiasp.Execute("SELECT rootid,TurnLink FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & Request("class"))
			If rootid = trs(0) Then
				ErrMsg = "<li>您不能指定该分类的下属分类作为所属分类</li>"
				Founderr = True
				Exit Sub
			End If
			If trs(1) = 1 Then
				ErrMsg = "<li>该分类是外部连接，您不能指定该分类作为所属分类</li>"
				Founderr = True
				Exit Sub
			End If
		End If
	Else
		Set trs = enchiasp.Execute("SELECT classid FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And ParentStr like '%" & ParentStr & "%' and classid=" & Request("class"))
		If Not (trs.EOF And trs.bof) Then
			ErrMsg = "<li>您不能指定该分类的下属分类作为所属分类</li>"
			Founderr = True
			Exit Sub
		End If
	End If
	If ParentID = 0 Then
		ParentID = Rs("classid")
		iparentid = 0
	End If
	Rs("classname") = enchiasp.ChkFormStr(Request.Form("classname"))
	Rs("readme") = Trim(Request.Form("readme"))
	Rs("ColorModes") = Trim(Request.Form("ColorModes"))
	Rs("FontModes") = Trim(Request.Form("FontModes"))
	Rs("LinkTarget") = Trim(Request.Form("LinkTarget"))
	Rs("TurnLink") = Trim(Request.Form("TurnLink"))
	Rs("TurnLinkUrl") = Trim(Request.Form("TurnLinkUrl"))
	Rs("UserGroup") = Trim(Request.Form("UserGroup"))
	Rs("ClassDir") = Trim(strClassDir)
	Rs("HtmlFileDir") = Trim(HtmlFileDir)
	Rs("UseHtml") = 1
	Rs("skinid") = Request.Form("skinid")
	Rs("isUpdate") = 1
	Rs("isallow") = Trim(Request.Form("isallow"))
	Rs("isdanyemian") = Trim(Request.Form("isdanyemian"))

	Rs.Update
	Rs.Close
	Set Rs = Nothing
	Set mrs = enchiasp.Execute("SELECT MAX(rootid) FROM ECCMS_Classify WHERE ChannelID = "& ChannelID)
	Maxrootid = mrs(0) + 1
	'假如更改了所属分类 
	'需要更新其原来所属分类信息，包括深度、父级ID、分类数、排序
	'需要更新当前所属分类信息
	If CLng(ParentID) <> CLng(Request("class")) And Not (iparentid = 0 And CInt(Request("class")) = 0) Then
		'如果原来不是一级分类改成一级分类
		If iparentid > 0 And CInt(Request("class")) = 0 Then
			'如果不是一级分类改成一级分类,更新子分类数据
			'开始更新子分类
			'ChildStr = "," & ChildStr
			Set Rsc = enchiasp.Execute ("SELECT classid,ChildStr FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid in (" & ParentStr & ")")
			Do While Not Rsc.EOF
				ArrChildStr = Split(Rsc("ChildStr"), ",")
				nChildStr = ""
				For ii = 0 to Ubound(ArrChildStr)
					If ArrChildStr(ii) <> ChildStr Then
						nChildStr = nChildStr & ArrChildStr(ii) & Chr(32)
					End If
				Next
				nChildStr = Replace(Trim(nChildStr), Chr(32), ",")
				'nChildStr = Replace(Rsc("ChildStr"), ChildStr, "")
				enchiasp.Execute ("UPDATE ECCMS_Classify SET ChildStr='" & nChildStr & "' WHERE ChannelID = "& ChannelID &" And classid = " & Rsc("classid"))
			Rsc.movenext
			Loop
			Rsc.Close
			Set Rsc = Nothing
			'更新子分类结束
			'更新当前分类数据
			enchiasp.Execute ("update ECCMS_Classify set depth=0,orders=0,rootid=" & Maxrootid & ",parentid=0,ParentStr='0' where classid=" & newclassid)
			ParentStr = ParentStr & ","
			Set Rs = enchiasp.Execute("SELECT COUNT(ClassID) FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And ParentStr like '%" & ParentStr & "%'")
			ClassCount = Rs(0)
			If IsNull(ClassCount) Then
				ClassCount = 1
			Else
				ClassCount = ClassCount + 1
			End If
			'更新其原来所属分类数
			enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child-" & ClassCount & " WHERE ChannelID = "& ChannelID &" And classid=" & iparentid)
			'更新其原来所属分类数据，排序相当于剪枝而不需考虑
			For i = 1 To depth
				'得到其父类的父类的ID
				Set Rs = enchiasp.Execute("SELECT parentid FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & iparentid)
				If Not (Rs.EOF And Rs.bof) Then
					iparentid = Rs(0)
					enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child-" & ClassCount & " WHERE ChannelID = "& ChannelID &" And classid=" & iparentid)
				End If
			Next
			If Child > 0 Then
				'更新其下属分类数据
				'有下属分类，排序不需考虑，更新下属分类深度和一级排序ID(rootid)数据
				'更新当前分类数据
				i = 0
				Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_Classify where ChannelID = "& ChannelID &" And ParentStr LIKE '%" & ParentStr & "%'")
				Do While Not Rs.EOF
					i = i + 1
					mParentStr = Replace(Rs("ParentStr"), ParentStr, "")
					enchiasp.Execute ("UPDATE ECCMS_Classify SET depth=depth-" & depth & ",rootid=" & Maxrootid & ",ParentStr='" & mParentStr & "' WHERE ChannelID = "& ChannelID &" And classid=" & Rs("classid"))
					Rs.movenext
				Loop
			End If
		ElseIf iparentid > 0 And CInt(Request("class")) > 0 Then
			'将一个分类移动到其他分类下
			'开始更新子分类
			'ChildStr = "," & ChildStr
			Set Rsc = enchiasp.Execute ("SELECT classid,ChildStr FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid in (" & ParentStr & ")")
			Do While Not Rsc.EOF
				ArrChildStr = Split(Rsc("ChildStr"), ",")
				nChildStr = ""
				For ii = 0 to Ubound(ArrChildStr)
					If ArrChildStr(ii) <> ChildStr Then
						nChildStr = nChildStr & ArrChildStr(ii) & Chr(32)
					End If
				Next
				nChildStr = Replace(Trim(nChildStr), Chr(32), ",")
				'nChildStr = Replace(Rsc("ChildStr"), ChildStr, "")
				enchiasp.Execute ("UPDATE ECCMS_Classify SET ChildStr='" & nChildStr & "' WHERE ChannelID = "& ChannelID &" And classid = " & Rsc("classid"))
			Rsc.movenext
			Loop
			Rsc.Close
			Set Rsc = Nothing
			'更新子分类结束
			'获得所指定的分类的相关信息
			Set trs = enchiasp.Execute("SELECT * FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & Request("class"))
			'得到其下属分类数 
			ParentStr = ParentStr & ","
			Set Rs = enchiasp.Execute("SELECT COUNT(ClassID) FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And ParentStr like '%" & ParentStr & "%'")
			ClassCount = Rs(0)
			If IsNull(ClassCount) Then ClassCount = 1
			'在获得移动过来的分类数后更新排序在指定分类之后的分类排序数据
			enchiasp.Execute ("UPDATE ECCMS_Classify SET orders=orders + " & ClassCount & " + 1 WHERE rootid=" & trs("rootid") & " and orders>" & trs("orders") & "")
			'更新当前分类数据
			enchiasp.Execute ("UPDATE ECCMS_Classify SET depth=" & trs("depth") & "+1,orders=" & trs("orders") & "+1,rootid=" & trs("rootid") & ",ParentID=" & Request("class") & ",ParentStr='" & trs("ParentStr") & "," & trs("classid") & "' WHERE ChannelID = "& ChannelID &" And classid=" & newclassid)
			i = 1
			'如果有则更新下属分类数据
			'深度为原有深度加上当前所属分类的深度
			Set Rs = enchiasp.Execute("select * from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentStr like '%" & ParentStr & "%' order by orders")
			Do While Not Rs.EOF
				i = i + 1
				iParentStr = trs("ParentStr") & "," & trs("classid") & "," & Replace(Rs("ParentStr"), ParentStr, "")
				enchiasp.Execute ("UPDATE ECCMS_Classify SET depth=depth+" & trs("depth") & "-" & depth & "+1,orders=" & trs("orders") & "+" & i & ",rootid=" & trs("rootid") & ",ParentStr='" & iParentStr & "' WHERE ChannelID = "& ChannelID &" And classid=" & Rs("classid"))
				Rs.movenext
			Loop
			ParentID = Request("class")
			If rootid = trs("rootid") Then
				'在同一分类下移动
				'更新所指向的上级分类数，i为本次移动过来的分类数
				'更新其父类分类数
				enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child+" & i & " WHERE ChannelID = "& ChannelID &" And (not ParentID=0) and classid=" & ParentID)
				For k = 1 To trs("depth")
					'得到其父类的父类的分类ID
					Set Rs = enchiasp.Execute("SELECT parentid FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And (not ParentID=0) and classid=" & ParentID)
					If Not (Rs.EOF And Rs.bof) Then
						ParentID = Rs(0)
						'更新其父类的父类分类数
						enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child+" & i & " WHERE ChannelID = "& ChannelID &" And (not ParentID=0) and  classid=" & ParentID)
					End If
				Next
				'更新其原父类分类数
				enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child-" & i & " WHERE ChannelID = "& ChannelID &" And (not ParentID=0) and classid=" & iparentid)
				'更新其原来所属分类数据
				For k = 1 To depth
					'得到其原父类的父类的分类ID
					Set Rs = enchiasp.Execute("SELECT parentid FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And (not ParentID=0) and classid=" & iparentid)
					If Not (Rs.EOF And Rs.bof) Then
						iparentid = Rs(0)
						'更新其原父类的父类分类数
						enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child-" & i & " WHERE ChannelID = "& ChannelID &" And (not ParentID=0) and  classid=" & iparentid)
					End If
				Next
			Else
				'更新所指向的上级分类数，i为本次移动过来的分类数
				'更新其父类分类数
				enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child+" & i & " WHERE ChannelID = "& ChannelID &" And classid=" & ParentID)
				For k = 1 To trs("depth")
					'得到其父类的父类的分类ID
					Set Rs = enchiasp.Execute("SELECT parentid FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & ParentID)
					If Not (Rs.EOF And Rs.bof) Then
						ParentID = Rs(0)
						'更新其父类的父类分类数
						enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child+" & i & " WHERE ChannelID = "& ChannelID &" And classid=" & ParentID)
					End If
				Next
				'更新其原父类分类数
				enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child-" & i & " where ChannelID = "& ChannelID &" And classid=" & iparentid)
				For k = 1 To depth
					'得到其原父类的父类的分类ID
					Set Rs = enchiasp.Execute("SELECT parentid FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & iparentid)
					If Not (Rs.EOF And Rs.bof) Then
						iparentid = Rs(0)
						'更新其原父类的父类分类数
						enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child-" & i & " WHERE ChannelID = "& ChannelID &" And classid=" & iparentid)
					End If
				Next
			End If
			'开始更新子分类
			SQL = "SELECT classid,parentid,ParentStr,ChildStr FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & Request("class")
			Set Rss = enchiasp.Execute (SQL)
			If Rss("parentid") <> 0 Then
				'如果是一级分类移动到其它一级分类的子分类
				nChildStr = Rss("ChildStr") & "," & Request("editid")
				enchiasp.Execute ("UPDATE ECCMS_Classify SET ChildStr='"&nChildStr&"' WHERE ChannelID = "& ChannelID &" And classid = " & Rss("classid"))
				SQL = "SELECT classid,ParentStr,ChildStr FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid in (" & Rss("ParentStr") & ")"
				Set Rsc = enchiasp.Execute (SQL)
				Do While Not Rsc.EOF
					nChildStr = Rsc("ChildStr") & "," & Request("editid")
					enchiasp.Execute ("UPDATE ECCMS_Classify SET ChildStr='"&nChildStr&"' WHERE ChannelID = "& ChannelID &" And classid = " & Rsc("classid"))
				Rsc.movenext
				Loop
				Rsc.Close
				Set Rsc = Nothing
			Else
				'如果是一级分类移动到其它一级分类，执行以下更新
				nChildStr = Rss("ChildStr") & "," & Request("editid")
				enchiasp.Execute ("UPDATE ECCMS_Classify SET ChildStr='"&nChildStr&"' WHERE ChannelID = "& ChannelID &" And classid = " & Rss("classid"))
			End If
			Rss.Close
			Set Rss = Nothing
			'更新子分类结束
		Else
			'如果原来是一级分类改成其他分类的下属分类
			'更新一级分类的子分类
			'开始更新子分类
			SQL = "SELECT classid,parentid,ParentStr,ChildStr FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & Request("class")
			Set Rss = enchiasp.Execute (SQL)
			If Rss("parentid") <> 0 Then
				'如果是一级分类移动到其它一级分类的子分类
				nChildStr = Rss("ChildStr") & "," & ChildStr
				enchiasp.Execute ("UPDATE ECCMS_Classify SET ChildStr='"&nChildStr&"' WHERE ChannelID = "& ChannelID &" And classid = " & Rss("classid"))
				SQL = "SELECT classid,ParentStr,ChildStr FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid in (" & Rss("ParentStr") & ")"
				Set Rsc = enchiasp.Execute (SQL)
				Do While Not Rsc.EOF
					nChildStr = Rsc("ChildStr") & "," & ChildStr
					enchiasp.Execute ("UPDATE ECCMS_Classify SET ChildStr='"&nChildStr&"' WHERE ChannelID = "& ChannelID &" And classid = " & Rsc("classid"))
				Rsc.movenext
				Loop
				Rsc.Close
				Set Rsc = Nothing
			Else
				'如果是一级分类移动到其它一级分类，执行以下更新
				nChildStr = Rss("ChildStr") & "," & ChildStr
				enchiasp.Execute ("UPDATE ECCMS_Classify SET ChildStr='"&nChildStr&"' WHERE ChannelID = "& ChannelID &" And classid = " & Rss("classid"))
			End If
			Rss.Close
			Set Rss = Nothing
			'更新子分类结束
			'得到所指定的分类的相关信息
			Set trs = enchiasp.Execute("SELECT * FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & Request("class"))
			Set Rs = enchiasp.Execute("SELECT COUNT(ClassID) FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And rootid=" & rootid)
			ClassCount = Rs(0)
			Rs.Close
			'更新所指向的上级分类数，i为本次移动过来的分类数
			ParentID = Request("class")
			'更新其父类分类数
			enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child+" & ClassCount & " WHERE ChannelID = "& ChannelID &" And classid=" & ParentID)
			For k = 1 To trs("depth")
				'得到其父类的父类的分类ID
				Set Rs = enchiasp.Execute("SELECT parentid FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And classid=" & ParentID)
				If Not (Rs.EOF And Rs.bof) Then
					ParentID = Rs(0)
					'更新其父类的父类分类数
					enchiasp.Execute ("UPDATE ECCMS_Classify SET child=child+" & ClassCount & " where ChannelID = "& ChannelID &" And classid=" & ParentID)
				End If

			Next
			'在获得移动过来的分类数后更新排序在指定分类之后的分类排序数据
			enchiasp.Execute ("UPDATE ECCMS_Classify SET orders=orders + " & ClassCount & " + 1 WHERE ChannelID = "& ChannelID &" And rootid=" & trs("rootid") & " and orders>" & trs("orders") & "")
			i = 0
			Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" And rootid=" & rootid & " order by orders")
			Do While Not Rs.EOF
				i = i + 1
				If Rs("parentid") = 0 Then
					If trs("ParentStr") = "0" Then
						ParentStr = trs("classid")
					Else
						ParentStr = trs("ParentStr") & "," & trs("classid")
					End If
					enchiasp.Execute ("UPDATE ECCMS_Classify SET depth=depth+" & trs("depth") & "+1,orders=" & trs("orders") & "+" & i & ",rootid=" & trs("rootid") & ",ParentStr='" & ParentStr & "',parentid=" & Request("class") & " WHERE ChannelID = "& ChannelID &" And classid=" & Rs("classid"))
				Else
					If trs("ParentStr") = "0" Then
						ParentStr = trs("classid") & "," & Rs("ParentStr")
					Else
						ParentStr = trs("ParentStr") & "," & trs("classid") & "," & Rs("ParentStr")
					End If
					enchiasp.Execute ("UPDATE ECCMS_Classify SET depth=depth+" & trs("depth") & "+1,orders=" & trs("orders") & "+" & i & ",rootid=" & trs("rootid") & ",ParentStr='" & ParentStr & "' WHERE ChannelID = "& ChannelID &" And classid=" & Rs("classid"))
				End If
				Rs.movenext
			Loop
		End If
	End If
	Set Rs = Nothing
	Set mrs = Nothing
	Set trs = Nothing
	Dim LocalPath
	If CInt(enchiasp.IsCreateHtml) > 0 And CInt(Request.Form("TurnLink")) = 0 Then
		LocalPath = enchiasp.InstallDir & ChannelDir & HtmlFileDir
		enchiasp.CreatPathEx(LocalPath)
	End If
	Call RemoveCache
	SucMsg = "<li>恭喜您！分类修改成功。</li>"
	Succeed(SucMsg)
End Sub

Private Sub DelClass()
	Dim ChildStr,nChildStr
	Dim Rss,Rsc
	On Error Resume Next
	Set Rs = enchiasp.Execute("select ParentStr,child,depth,parentid,HtmlFileDir,UseHtml from ECCMS_Classify where ChannelID = "& ChannelID &" And classid=" & Request("editid"))
	If Not (Rs.EOF And Rs.bof) Then
		If Rs(1) > 0 Then
			ErrMsg = "<li>该分类含有下属分类，请删除其下属分类后再进行删除本分类的操作</li>"
			Founderr = True
			Exit Sub
		End If
		HtmlFileDir = Rs(4)
		UseHtml = Rs(5)
		If Rs(3) > 0 Then
			ChildStr = "," & Request("editid")
			SQL = "select classid,ParentStr from ECCMS_Classify where ChannelID = "& ChannelID &" And classid=" & Request("editid")
			Set Rss = enchiasp.Execute (SQL)
			SQL = "select classid,ChildStr from ECCMS_Classify where ChannelID = "& ChannelID &" And classid in (" & Rss("ParentStr") & ")"
			Set Rsc = enchiasp.Execute (SQL)
			Do While Not Rsc.EOF
				nChildStr = Replace(Rsc("ChildStr"), ChildStr, "")
				enchiasp.Execute ("update ECCMS_Classify set ChildStr='"&nChildStr&"' where ChannelID = "& ChannelID &" And classid = " & Rsc("classid"))
			Rsc.movenext
			Loop
			Rsc.Close
			Set Rsc = Nothing
			Set Rss = Nothing
		End If
		If Rs(2) > 0 Then
			enchiasp.Execute ("update ECCMS_Classify set child=child-1 where ChannelID = "& ChannelID &" And classid in (" & Rs(0) & ")")
		End If
		SQL = "delete from ECCMS_Classify where ChannelID = "& ChannelID &" And classid=" & Request("editid")
		enchiasp.Execute (SQL)
		Call DelRelated
	End If
	Set Rs = Nothing
	Call RemoveCache
	Succeed ("恭喜您！分类删除成功。")
End Sub

Private Sub DelRelated()
	On Error Resume Next
	Select Case moduleid
	Case 1
		enchiasp.Execute("DELETE ECCMS_Comment FROM ECCMS_Article A INNER JOIN ECCMS_Comment C ON C.PostID=A.ArticleID WHERE A.ChannelID = "& ChannelID &" And A.classid=" & Request("editid"))
		enchiasp.Execute("DELETE FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And classid=" & Request("editid"))
	Case 2
		enchiasp.Execute ("DELETE ECCMS_DownAddress FROM ECCMS_SoftList A INNER JOIN ECCMS_DownAddress D ON D.SoftID=A.SoftID WHERE A.ChannelID = "& ChannelID &" And A.classid=" & Request("editid"))
		enchiasp.Execute ("DELETE ECCMS_Comment FROM ECCMS_SoftList A INNER JOIN ECCMS_Comment C ON C.PostID=A.SoftID WHERE A.ChannelID = "& ChannelID &" And A.classid=" & Request("editid"))
		enchiasp.Execute ("DELETE FROM ECCMS_SoftList WHERE ChannelID = "& ChannelID &" And classid=" & Request("editid"))
	Case 3
		enchiasp.Execute("DELETE ECCMS_Comment FROM ECCMS_ShopList A INNER JOIN ECCMS_Comment C ON C.PostID=A.ShopID WHERE A.ChannelID = "& ChannelID &" And A.classid=" & Request("editid"))
		enchiasp.Execute("DELETE FROM ECCMS_ShopList WHERE ChannelID = "& ChannelID &" And classid=" & Request("editid"))
	Case 5
		enchiasp.Execute("DELETE ECCMS_Comment FROM ECCMS_FlashList A INNER JOIN ECCMS_Comment C ON C.PostID=A.flashid WHERE A.ChannelID = "& ChannelID &" And A.classid=" & Request("editid"))
		enchiasp.Execute("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And classid=" & Request("editid"))
	case 6
		enchiasp.Execute("DELETE ECCMS_Comment FROM ECCMS_Article A INNER JOIN ECCMS_Comment C ON C.PostID=A.ArticleID WHERE A.ChannelID = "& ChannelID &" And A.classid=" & Request("editid"))
		enchiasp.Execute("DELETE FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And classid=" & Request("editid"))
	case 7
		enchiasp.Execute("DELETE ECCMS_Comment FROM ECCMS_Article A INNER JOIN ECCMS_Comment C ON C.PostID=A.ArticleID WHERE A.ChannelID = "& ChannelID &" And A.classid=" & Request("editid"))
		enchiasp.Execute("DELETE FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And classid=" & Request("editid"))
	
	End Select
	enchiasp.FolderDelete(enchiasp.InstallDir & ChannelDir & HtmlFileDir)
End Sub

Private Sub DelClassDir()
	On Error Resume Next
	Set Rs = enchiasp.Execute("select HtmlFileDir from ECCMS_Classify where ChannelID = "& ChannelID &" And classid=" & Request("editid"))
	If Not (Rs.EOF And Rs.BOF) Then
		enchiasp.FolderDelete(enchiasp.InstallDir & ChannelDir & Rs("HtmlFileDir"))
	End If
	Succeed ("恭喜您！分类目录删除成功。")
End Sub

Private Sub orders()
	Response.Write " <table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th>分类一级分类重新排序修改(请在相应分类的排序表单内输入相应的排列序号) </th>"
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td class=TableRow1>"
	Response.Write "<table width=""50%"">" & vbCrLf
	SQL = "select * from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentID=0 order by RootID"
	Set Rs = enchiasp.Execute (SQL)
	If Rs.bof And Rs.EOF Then
		ErrMsg = "<li>还没有相应的" & sModuleName & "分类。</li>"
		Founderr = True
		Exit Sub
	Else
		Do While Not Rs.EOF
			Response.Write "<form action=?action=neworders method=post><tr><td width=""50%"">" & enchiasp.ReadFontMode(Rs("classname"),Rs("ColorModes"),Rs("FontModes")) & "</td>" & vbCrLf
			Response.Write "<td width=""50%""><input type=""hidden"" name=""ChannelID"" value=""" & ChannelID & """><input type=text name=""OrderID"" size=4 value=""" & Rs("rootid") & """><input type=hidden name=""cID"" value=""" & Rs("rootid") & """>&nbsp;&nbsp;<input type=submit name=Submit class=button value='修 改'></td></tr></form>" & vbCrLf
			Rs.movenext
		Loop
		Response.Write "</table>" & vbCrLf
		Response.Write "<BR>&nbsp;<font color=red>请注意，这里一定<B>不能填写相同的序号</B>，否则非常难修复！</font>"
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write "</table>" & vbCrLf
End Sub

Private Sub updateorders()
	Dim cID
	Dim OrderID
	Dim ClassName
	cID = Replace(Request.Form("cID"), "'", "")
	OrderID = Replace(Request.Form("OrderID"), "'", "")
	Set Rs = enchiasp.Execute("select classid from ECCMS_Classify where ChannelID = "& ChannelID &" And rootid=" & OrderID)
	If Rs.bof And Rs.EOF Then
		Succeed ("恭喜您！设置成功，请返回。")
		enchiasp.Execute ("update ECCMS_Classify set rootid=" & OrderID & " where ChannelID = "& ChannelID &" And rootid=" & cID)
	Else
		ErrMsg = "<li>请不要和其他分类设置相同的序号</li>"
		Founderr = True
		Exit Sub
	End If
	Call RemoveCache
	Set Rs = Nothing
End Sub

Private Sub classorders()
	Dim trs
	Dim uporders
	Dim doorders
	Response.Write " <table border=""0"" cellspacing=""1"" cellpadding=""2"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th colspan=2 class=""TableRow1"">分类N级分类重新排序修改(请在相应分类的排序表单内输入相应的排列序号)"
	Response.Write " </th>"
	Response.Write " </tr>" & vbCrLf
	Set Rs = Server.CreateObject("Adodb.recordset")
	SQL = "select * from ECCMS_Classify where ChannelID = "& ChannelID &" order by RootID,orders"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "还没有相应的分类。"
	Else
		Do While Not Rs.EOF
			Response.Write "<form action=admin_classify.asp?action=newclassorders&ChannelID=" & ChannelID & " method=post><tr><td width=""50%"" class=TableRow1>" & vbCrLf
			If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font>"
			If Rs("depth") > 1 Then
				For i = 2 To Rs("depth")
					Response.Write "&nbsp;&nbsp;<font color=""#666666"">│</font>"
				Next
				Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font> "
			End If
			If Rs("parentid") = 0 Then Response.Write ("<b>")
			Response.Write enchiasp.ReadFontMode(Rs("classname"),Rs("ColorModes"),Rs("FontModes"))
			If Rs("child") > 0 Then Response.Write "(" & Rs("child") & ")"
			Response.Write "</td><td width=""50%"" class=TableRow2>" & vbCrLf
			If Rs("ParentID") > 0 Then
				Set trs = enchiasp.Execute("select count(*) from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentID=" & Rs("ParentID") & " And orders<" & Rs("orders") & "")
				uporders = trs(0)
				If IsNull(uporders) Then uporders = 0

				Set trs = enchiasp.Execute("select count(*) from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentID=" & Rs("ParentID") & " And orders>" & Rs("orders") & "")
				doorders = trs(0)
				If IsNull(doorders) Then doorders = 0
				If uporders > 0 Then
					Response.Write "<select name=uporders size=1><option value=0>↑</option>" & vbCrLf
					For i = 1 To uporders
						Response.Write "<option value=" & i & ">↑" & i & "</option>" & vbCrLf
					Next
					Response.Write "</select>"
				End If
				If doorders > 0 Then
					If uporders > 0 Then Response.Write "&nbsp;"
					Response.Write "<select name=doorders size=1><option value=0>↓</option>" & vbCrLf
					For i = 1 To doorders
						Response.Write "<option value=" & i & ">↓" & i & "</option>" & vbCrLf
					Next
					Response.Write "</select>" & vbCrLf
				End If
				If doorders > 0 Or uporders > 0 Then
					Response.Write "<input type=hidden name=""editID"" value=""" & Rs("classid") & """>&nbsp;<input type=submit name=Submit class=button value='修 改'>" & vbCrLf
				End If
			End If
			Response.Write "</td></tr></form>" & vbCrLf
			uporders = 0
			doorders = 0
			Rs.movenext
		Loop
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>"
End Sub

Private Sub updateclassorders()
	Dim ParentID
	Dim orders
	Dim ParentStr
	Dim Child
	Dim uporders
	Dim doorders
	Dim oldorders
	Dim trs
	Dim ii
	If Not IsNumeric(Request("editID")) Then
		ErrMsg = ErrMsg & "<li>非法的参数！</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") = "" Then
		ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	
	If Request("uporders") <> "" And Not CInt(Request("uporders")) = 0 Then
		If Not IsNumeric(Request("uporders")) Then
			ErrMsg = ErrMsg & "<li>非法的参数！</li>"
			Founderr = True
			Exit Sub
		ElseIf CInt(Request("uporders")) = 0 Then
			ErrMsg = ErrMsg & "<li>请选择要提升的数字！</li>"
			Founderr = True
			Exit Sub
		End If
		Set Rs = enchiasp.Execute("select ParentID,orders,ParentStr,child from ECCMS_Classify where ChannelID = "& ChannelID &" And classid=" & Request("editID"))
		ParentID = Rs(0)
		orders = Rs(1)
		ParentStr = Rs(2) & "," & Request("editID")
		Child = Rs(3)
		i = 0
		If Child > 0 Then
			Set Rs = enchiasp.Execute("select count(*) from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentStr like '%" & ParentStr & "%'")
			oldorders = Rs(0)
		Else
			oldorders = 0
		End If
		Set Rs = enchiasp.Execute("select classid,orders,child,ParentStr from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentID=" & ParentID & " and orders<" & orders & " order by orders desc")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("uporders")) >= i Then
				If Rs(2) > 0 Then
					ii = 0
					Set trs = enchiasp.Execute("select classid,orders from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentStr like '%" & Rs(3) & "," & Rs(0) & "%' order by orders")
					If Not (trs.EOF And trs.bof) Then
						Do While Not trs.EOF
							ii = ii + 1
							enchiasp.Execute ("update ECCMS_Classify set orders=" & orders & "+" & oldorders & "+" & ii & " where ChannelID = "& ChannelID &" And classid=" & trs(0))
							trs.movenext
						Loop
					End If
				End If
				enchiasp.Execute ("update ECCMS_Classify set orders=" & orders & "+" & oldorders & " where ChannelID = "& ChannelID &" And classid=" & Rs(0))
				If CInt(Request("uporders")) = i Then uporders = Rs(1)
			End If
			orders = Rs(1)
			Rs.movenext
		Loop
		enchiasp.Execute ("update ECCMS_Classify set orders=" & uporders & " where ChannelID = "& ChannelID &" And classid=" & Request("editID"))
		If Child > 0 Then
			i = uporders
			Set Rs = enchiasp.Execute("select classid from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentStr like '%" & ParentStr & "%' order by orders")
			Do While Not Rs.EOF
				i = i + 1
				enchiasp.Execute ("update ECCMS_Classify set orders=" & i & " where ChannelID = "& ChannelID &" And classid=" & Rs(0))
				Rs.movenext
			Loop
		End If
		Set Rs = Nothing
		Set trs = Nothing
	ElseIf Request("doorders") <> "" Then
		If Not IsNumeric(Request("doorders")) Then
			ErrMsg = ErrMsg & "<li>非法的参数！</li>"
			Founderr = True
			Exit Sub
		ElseIf CInt(Request("doorders")) = 0 Then
			ErrMsg = ErrMsg & "<li>请选择要下降的数字！</li>"
			Founderr = True
			Exit Sub
		End If
		Set Rs = enchiasp.Execute("select ParentID,orders,ParentStr,child from ECCMS_Classify where ChannelID = "& ChannelID &" And classid=" & Request("editID"))
		ParentID = Rs(0)
		orders = Rs(1)
		ParentStr = Rs(2) & "," & Request("editID")
		Child = Rs(3)
		i = 0
		If Child > 0 Then
			Set Rs = enchiasp.Execute("select count(*) from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentStr like '%" & ParentStr & "%'")
			oldorders = Rs(0)
		Else
			oldorders = 0
		End If
		Set Rs = enchiasp.Execute("select classid,orders,child,ParentStr from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentID=" & ParentID & " and orders>" & orders & " order by orders")
		Response.Write "<li>"&ChannelID&" 错误参数！</li>"
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("doorders")) >= i Then
				If Rs(2) > 0 Then
					ii = 0
					Set trs = enchiasp.Execute("select classid,orders from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentStr like '%" & Rs(3) & "," & Rs(0) & "%' order by orders")
					If Not (trs.EOF And trs.bof) Then
						Do While Not trs.EOF
							ii = ii + 1
							enchiasp.Execute ("update ECCMS_Classify set orders=" & orders & "+" & ii & " where ChannelID = "& ChannelID &" And classid=" & trs(0))
							trs.movenext
						Loop
					End If
				End If
				enchiasp.Execute ("update ECCMS_Classify set orders=" & orders & " where ChannelID = "& ChannelID &" And classid=" & Rs(0))
				If CInt(Request("doorders")) = i Then doorders = Rs(1)
			End If
			orders = Rs(1)
			Rs.movenext
		Loop
		enchiasp.Execute ("update ECCMS_Classify set orders=" & doorders & " where ChannelID = "& ChannelID &" And classid=" & Request("editID"))
		If Child > 0 Then
			i = doorders
			Set Rs = enchiasp.Execute("select classid from ECCMS_Classify where ChannelID = "& ChannelID &" And ParentStr like '%" & ParentStr & "%' order by orders")
			Do While Not Rs.EOF
				i = i + 1
				enchiasp.Execute ("update ECCMS_Classify set orders=" & i & " where ChannelID = "& ChannelID &" And classid=" & Rs(0))
				Rs.movenext
			Loop
		End If
	End If
	Set Rs = Nothing
	Set trs = Nothing
	Call RemoveCache
	Response.redirect "admin_classify.asp?action=classorders&ChannelID=" & ChannelID
	Response.End
End Sub

Private Sub RestoreClass()
	i = 0
	Set Rs = enchiasp.Execute("SELECT classid FROM ECCMS_Classify WHERE ChannelID = "& ChannelID &" order by rootid,orders")
	Do While Not Rs.EOF
		i = i + 1
		enchiasp.Execute ("UPDATE ECCMS_Classify SET rootid=" & i & ",depth=0,orders=0,ParentID=0,ParentStr='0',child=0, ChildStr='"&Rs(0)&"' WHERE ChannelID = "& ChannelID &" And classid=" & Rs(0))
		Rs.movenext
	Loop
	Set Rs = Nothing
	Call RemoveCache
	Succeed("<li>复位成功，请返回做分类归属设置。</li>")
End Sub
Sub RemoveCache()
	enchiasp.DelCahe "SelectClass" & ChannelID
	enchiasp.DelCahe "ClassJumpMenu" & ChannelID
	enchiasp.DelCahe "SiteClassMap"
	enchiasp.DelCache "SelectClass" & ChannelID
	enchiasp.DelCache "ClassJumpMenu" & ChannelID
End Sub
%>