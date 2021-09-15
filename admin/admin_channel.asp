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
Response.Write "<script language = JavaScript>" & vbCrLf
Response.Write "function ChannelSetting(n){" & vbCrLf
Response.Write "	if (n == 1){" & vbCrLf
Response.Write "		ChannelSetting1.style.display='none';" & vbCrLf
Response.Write "		ChannelSetting2.style.display='';" & vbCrLf
Response.Write "	}" & vbCrLf
Response.Write "	else{" & vbCrLf
Response.Write "		ChannelSetting1.style.display='';" & vbCrLf
Response.Write "		ChannelSetting2.style.display='none';" & vbCrLf
Response.Write "	}" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder"">"
Response.Write "	<tr>"
Response.Write "		<th colspan=""2"">站点频道管理</th>"
Response.Write "	</tr>"
Response.Write "	<tr>"
Response.Write "		<td width=""100%"" class=TableRow2 colspan=2><b>管理选项：</b><a href=admin_channel.asp>管理首页</a>" 
Response.Write "		| <a href=?action=add>添加频道</a> | "
Dim Rsm,ModuleName,strModuleName,sChannelID,NewChannelID
Set Rsm = enchiasp.Execute("SELECT ChannelID,ModuleName From ECCMS_Channel WHERE ChannelType < 2  ORDER BY orders ASC")
Do While Not Rsm.EOF
	Response.Write "<a href=?action=edit&ChannelID="
	Response.Write Rsm("ChannelID")
	Response.Write ">"
	Response.Write Rsm("ModuleName")
	Response.Write "设置</a> | "
	strModuleName = strModuleName & Rsm("ModuleName") & "|||"
	sChannelID = sChannelID & Rsm("ChannelID") & "|||"
	Rsm.movenext
Loop
Set Rsm = Nothing
Response.Write "<a href=?action=orders>频道排序</a>"
Response.Write "		</td>"
Response.Write "	</tr>"
Response.Write "</table>"
Response.Write "<br>"
Dim Action,ChannelDir,TitleColor,mChannelDir,mChannelID
Dim i,RsObj
Action = LCase(enchiasp.RemoveBadCharacters(Request("action")))
If Not ChkAdmin("Channel") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Select Case Trim(Action)
Case "savenew"
	Call SavenewChannel
Case "savedit"
	Call SaveditChannel
Case "add"
	Call ChannelAdd
Case "edit"
	Call ChannelEdit
Case "del"
	Call ChannelDel
Case "orders"
	Call ChannelOrders
Case "saveorder"
	Call SaveOrder
Case "stopchannel"
	Call UpdateStop
Case "ishidden"
	Call UpdateHidden
Case "linktarget"
	Call UpdateLinkTarget
Case "createhtml"
	Call UpdateCreateHtml
Case "reload"
	Call ReloadChannelCache
Case Else
	Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If

Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub showmain()
	Response.Write "<table border=""0"" align=""center"" cellspacing=""1"" cellpadding=""3"" class=""TableBorder"">"
	Response.Write "	<tr>"
	Response.Write "		<th>频道名称</th>"
	Response.Write "		<th>频道类型</th>"
	Response.Write "		<th>频道状态</th>"
	Response.Write "		<th>是否HTML</th>"
	Response.Write "		<th>名称状态</th>"
	Response.Write "		<th>连接目标</th>"
	Response.Write "		<th>管理选项</th>"
	Response.Write "	</tr>"

	SQL = "SELECT * FROM ECCMS_Channel ORDER BY orders"
	Set Rs = enchiasp.Execute(SQL)
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	Do While Not Rs.EOF
		Response.Write "	<tr>"
		Response.Write "		<td class=""TableRow2"">"
		Response.Write ("<a href=?action=edit&ChannelID=" & Rs("ChannelID") & " title=修改此频道设置>")
		Response.Write (enchiasp.ReadFontMode(Rs("ChannelName"),Rs("ColorModes"),Rs("FontModes")))
		Response.Write ("</a>")
		Response.Write "		</td>"
		Response.Write "		<td class=""TableRow1"" align=""center"">"
		If Rs("ChannelType") = 0 Then
			Response.Write ("<font color=blue>系统频道")
		Elseif Rs("ChannelType") = 1 Then
			Response.Write ("<font color=green>内部频道")
		Else
			Response.Write ("<font color=red>外部频道")
		End If
		Response.Write ("<font>")
		Response.Write ("</td>")
		If Rs("ChannelType") < 2 Then
			Response.Write ("<td class=""TableRow2"" align=""center"">")
			If Rs("StopChannel") <> 0 Then
				Response.Write ("<a href=?action=StopChannel&ChannelID=" & Rs("ChannelID") & "&EditID=0 title=""切换到：打开此频道""><font color=red>关闭<font></a>")
			Else
				Response.Write ("<a href=?action=StopChannel&ChannelID=" & Rs("ChannelID") & "&EditID=1 title=""切换到：关闭此频道"">打开</a>")
			End If
			Response.Write "		</td>"
			Response.Write "		<td class=""TableRow1"" align=""center"">"
			If Rs("IsCreateHtml") = 0 Then
				If Rs("ChannelID") = 4 Then
					Response.Write ("否")
				Else
					Response.Write ("<a href=?action=createhtml&ChannelID=" & Rs("ChannelID") & "&EditID=1 title=""切换到：生成HTML"">否</a>")
				End If
			Else
				Response.Write ("<a href=?action=createhtml&ChannelID=" & Rs("ChannelID") & "&EditID=0 title=""切换到：不生成HTML""><font color=blue>是</font></a>")
			End If
		Else
			Response.Write ("<td colspan=""2"" class=""TableRow2"" align=""center"">")
			Response.Write ("<a href=" & Rs("ChannelUrl") & " target=_blank><font color=blue>" & Rs("ChannelUrl") & "</font></a>")
		End If
		Response.Write "		</td>"
		Response.Write "				<td class=""TableRow2"" align=""center"">"
		If Rs("IsHidden") <> 0 Then
			Response.Write ("<a href=?action=ishidden&ChannelID=" & Rs("ChannelID") & "&EditID=0 title=""切换到：隐藏频道名称""><font color=green>隐藏<font></a>")
		Else
			Response.Write ("<a href=?action=ishidden&ChannelID=" & Rs("ChannelID") & "&EditID=1 title=""切换到：隐藏频道名称"">显示</a>")
		End If
		Response.Write "		</td>"
		Response.Write "		<td class=""TableRow1"" align=""center"">"
		If Rs("LinkTarget") = 0 Then
			Response.Write ("<a href=?action=linktarget&ChannelID=" & Rs("ChannelID") & "&EditID=1 title=""切换到：新窗口打开"">本窗口打开</a>")
		Else
			Response.Write ("<a href=?action=linktarget&ChannelID=" & Rs("ChannelID") & "&EditID=0 title=""切换到：本窗口打开""><font color=blue>新窗口打开<font></a>")
		End If
		Response.Write "		</td>"
		Response.Write "		<td class=""TableRow2"" align=""center""><A HREF=?action=edit&ChannelID="
		Response.Write Rs("ChannelID")
		Response.Write ">编 辑</A>"
		If Rs("ChannelID") => 10 Then
			Response.Write " | <A HREF=?action=del&ChannelID="
			Response.Write Rs("ChannelID")
			Response.Write " onclick=""{if(confirm('此操作将删除此频道\n您确定要删除吗?')){return true;}return false;}"">删 除</A>"
		End If
		If Rs("ChannelType") < 2 Then
			'Response.Write " | <A HREF=?action=reload&ChannelID="
			'Response.Write Rs("ChannelID")
			'Response.Write "><font color=blue>更新缓存</font></a>"
			If Rs("ChannelID") <> 4 Then
				Response.Write " | <A HREF=admin_classify.asp?action=jsmenu&ChannelID="
				Response.Write Rs("ChannelID")
				Response.Write "&stype=1><font color=green>生成JS菜单</font></a>"
			End If
		End If
		Response.Write "		</td>"
		Response.Write "	</tr>"

	Rs.movenext
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write "	<tr>"
	Response.Write "		<td colspan=""7"" class=""TableRow1""><b>说明：</b> <br>①、点击相应的状态名可以进行相关快捷切换操作；<br>"
	Response.Write "②、在切换HTML生成功能后，请<font color=red>重新生成JS</font>菜单。"
	Response.Write "</td>	</tr>"
	Response.Write "</table>"
End Sub

Private Sub ChannelAdd()
	
	Set Rs = enchiasp.Execute("select Max(ChannelID) from ECCMS_Channel")
	If Rs.bof And Rs.EOF Then
		NewChannelID = 1
	Else
		NewChannelID = Rs(0) + 1
	End If
	If IsNull(NewChannelID) Then NewChannelID = 1
	Rs.Close
	If NewChannelID < 10 Then NewChannelID = 10
%>
	<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
		<tr>
			<th colspan="2" align="left"><img src="images/welcome.gif" width="16" height="17" align="absMiddle"> 添加站点频道</th>
		</tr>
		<form method="POST" action="?action=savenew">
		<input type="hidden" name="NewChannelID" value="<%=NewChannelID%>">
		<tr>
			<td width="20%" class="TableRow2"><div class="divbody">频道名称</td>
			<td width="80%" class="TableRow1">
			<input type="text" name="ChannelName" size="20"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">频道名称模式</div></td>
			<td class="TableRow1"> 颜色：
			<select size="1" name="ColorModes">
			<option value="0">请选择标题颜色</option>
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
			<td class="TableRow2"><div class="divbody">频道注释</div></td>
			<td class="TableRow1">
			<input type="text" name="Caption" size="60"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">频道类型</div></td>
			<td class="TableRow1">
			<input type="radio" value="2" checked name="ChannelType" onClick="ChannelType1.style.display='';ChannelType2.style.display='none';ChannelType3.style.display='none';"> 外部频道&nbsp;&nbsp; 
			<input type="radio" name="ChannelType" value="1" onClick="ChannelType1.style.display='none';ChannelType2.style.display='';ChannelType3.style.display='';"> 内部频道</td>
		</tr>
		<tr id=ChannelType1>
			<td class="TableRow2"><div class="divbody">频道连接URL</div></td>
			<td class="TableRow1">
			<input type="text" name="ChannelUrl" size="45" value="<%=enchiasp.SiteUrl%>"> <font color="#FF0000">
			* 请输入完整的URL</font></td>
		</tr>
		<tr id=ChannelType2 style="display:none">
			<td class="TableRow2"><div class="divbody">所属模块</div></td>
			<td class="TableRow1">
			<select name="modules" szie=1>
				<option value='1'>文章</option>
				<option value='2'>软件</option>
				<option value='3'>商城</option>
				<option value='5'>动画</option>
				<option value='6'>单页图文</option>

			</select></td>
		</tr>
		<tr id=ChannelType3 style="display:none">
			<td class="TableRow2"><div class="divbody">频道目录</div></td>
			<td class="TableRow1"><input type="text" name="ChannelDir" size=20 value='dir'></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">连接目标</div></td>
			<td class="TableRow1">
			<input type="radio" value="0" checked name="LinkTarget"> 本窗口打开&nbsp;&nbsp; 
			<input type="radio" name="LinkTarget" value="1"> 新窗口打开</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">频道菜单状态</div></td>
			<td class="TableRow1">
			<input type="radio" name="IsHidden" value="0" checked> 正常&nbsp;&nbsp; 
			<input type="radio" name="IsHidden" value="1"> 隐藏</td>
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

Private Sub ChannelEdit()
	Dim Rs_c,tempstr
	Dim Channel_Setting
	Set Rs = enchiasp.Execute("SELECT * FROM ECCMS_Channel WHERE ChannelID=" & CLng(Request("ChannelID")))
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = "数据库出现错误,没有此站点频道!"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
	Channel_Setting = Split(Rs("Channel_Setting"), "|||")
	tempstr = enchiasp.HtmlRndFileName
%>
	<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
		<tr>
			<th colspan="2" align="left"><img src="images/welcome.gif" width="16" height="17" align="absMiddle"> 编辑站点频道</th>
		</tr>
		<form method="POST" action="?action=savedit">
		<input type="hidden" name="ChannelID" value="<%=Rs("ChannelID")%>">
		<tr>
			<td width="28%" class="TableRow2"><div class="divbody">频道名称：</div></td>
			<td width="72%" class="TableRow1">
			<input type="text" name="ChannelName" size="20" value="<%=Rs("ChannelName")%>"></td>
		</tr>
				<tr>
			<td class="TableRow2"><div class="divbody">频道名称模式：</div></td>
			<td class="TableRow1">颜色： 
			<select size="1" name="ColorModes">
			<option value="0">请选择标题颜色</option>
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
			<td class="TableRow2"><div class="divbody">频道注释：</div></td>
			<td class="TableRow1">
			<input type="text" name="Caption" size="60" value="<%=Rs("Caption")%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">连接目标：</div></td>
			<td class="TableRow1">
			<input type="radio" name="LinkTarget" value="0"<%If Rs("LinkTarget") = 0 Then Response.Write (" checked")%>> 本窗口打开&nbsp;&nbsp; 
			<input type="radio" name="LinkTarget" value="1"<%If Rs("LinkTarget") = 1 Then Response.Write (" checked")%>> 新窗口打开</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">频道菜单状态：</div></td>
			<td class="TableRow1">
			<input type="radio" name="IsHidden" value="0"<%If Rs("IsHidden") = 0 Then Response.Write (" checked")%>> 正常&nbsp;&nbsp; 
			<input type="radio" name="IsHidden" value="1"<%If Rs("IsHidden") = 1 Then Response.Write (" checked")%>> 隐藏</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">连接类型：</div></td>
			<td class="TableRow1">
			<%If Rs("ChannelType") = 0 Then%>
			<input type="radio" name="ChannelType" value="0" checked> 系统频道
			<%ElseIf Rs("ChannelType") = 1 Then%>
			<input type="radio" name="ChannelType" value="1"<%If Rs("ChannelType") = 1 Then Response.Write (" checked")%>> 内部频道&nbsp;&nbsp; 
			<%Else%>
			<input type="radio" name="ChannelType" value="2"<%If Rs("ChannelType") = 2 Then Response.Write (" checked")%>> 外部频道
			<%End IF%></td>
		</tr>
		<tr id=ChannelSetting1<%If Rs("ChannelType") = 0 Or Rs("ChannelType") = 1 Then Response.Write (" style=""display:'none'""")%>>
			<td class="TableRow2"><div class="divbody">频道连接URL：</div></td>
			<td class="TableRow1">
			<input type="text" name="ChannelUrl" size="45" value="<%=Rs("ChannelUrl")%>"> <font color="#FF0000">
			* 外部连接URL以“http://”开头</font></td>
		</tr>
		<tr id=ChannelSetting2<%If Rs("ChannelType") => 2 Then Response.Write (" style=""display:'none'""")%>>
		<td class="TableRow1" colspan="2"><fieldset style="cursor: default"><legend>&nbsp;系统频道设置</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
			<tr>
				<td class="TableRow2"><div class="divbody">是否关闭本频道：</div></td>
				<td class="TableRow1">
				<input type="radio" name="StopChannel" value="0"<%If Rs("StopChannel") = 0 Then Response.Write (" checked")%>> 打开&nbsp;&nbsp; 
				<input type="radio" name="StopChannel" value="1"<%If Rs("StopChannel") = 1 Then Response.Write (" checked")%>> 关闭&nbsp;&nbsp; </td>
			</tr>
<%If (Rs("modules") = 6 or Rs("modules") =7) Then %>	
<input  type=hidden type="text"  name="ModuleName" size="10" value="<%=Rs("ModuleName")%>">
<input type=hidden type="text"   name="modules" size="10" value="<%=Rs("modules")%>"> 
<input type=hidden type="text"   name="ChannelSkin" size="10" value="<%=Rs("ChannelSkin")%>"> 

<% else %>		
<tr>
				<td class="TableRow2"><div class="divbody">频道模块名称：</div></td>
				<td class="TableRow1">
				<input type="text" name="ModuleName" size="10" value="<%=Rs("ModuleName")%>"></td>
			</tr>

			<tr>
				<td width="28%" class="TableRow1"><div class="divbody">频道功能模块：</div></td>
				<td width="72%" class="TableRow1">
					<select size="1" name="modules" disabled>
<%
		Response.Write "	<option value=0"
		If Rs("modules") = 0 Then Response.Write (" selected")
		Response.Write ">外部</option>"
		strModuleName = Split(strModuleName,"|||")
		sChannelID = Split(sChannelID,"|||")
		For i = 0 To UBound(strModuleName) - 1
			Response.Write "	<option value="
			Response.Write sChannelID(i)
			If Rs("modules") = Clng(sChannelID(i)) Then Response.Write (" selected")
			Response.Write ">"
			Response.Write strModuleName(i)
			Response.Write "</option>"
		Next

	Response.Write "					</select></td>"
	Response.Write "			</tr>"
	Response.Write "			<tr>"
	Response.Write "				<td class=""TableRow1""><div class=""divbody"">频道默认模板：</div></td>"
	Response.Write "				<td class=""TableRow1"">"
	Response.Write "				<select size=""1"" name=""ChannelSkin"">"

	Response.Write "		<option value=""0"""
	If Rs("ChannelSkin") = 0 Then Response.Write " selected"
	Response.Write ">使用默认模板</option>" & vbCrLf
	SQL = "Select skinid,page_name,isDefault From ECCMS_Template Where pageid = 0 order by TemplateID"
	Set RsObj = enchiasp.Execute(SQL)
	If RsObj.bof And RsObj.EOF Then
		Response.Write "		<option value=""0"">您还没有添加任何模板文件</option>" & vbCrLf
	Else
		Do While Not RsObj.EOF
			Response.Write "		<option value=""" & RsObj("skinid") & """"
			If Rs("ChannelSkin") = RsObj("skinid") Then Response.Write " selected"
			Response.Write ">"
			Response.Write RsObj("page_name")
			Response.Write "</option>" & vbCrLf
			RsObj.movenext
		Loop
	End IF
	Set RsObj = Nothing
%>		</select></td>
			</tr>
<%end if%>
			<tr>
				<td class="TableRow2"><div class="divbody">频道所在目录：</div></td>
				<td class="TableRow1">
				<input type="text" name="ChannelDir" size="20" value="<%=Rs("ChannelDir")%>"> <font color="#FF0000">
				* 如果要修改频道所在目录，请手工修改相应的目录名称</font></td>
			</tr>
			<tr style="display:none">
				<td class="TableRow1"><div class="divbody">是否启用域名绑定功能：</div></td>
				<td class="TableRow1">
				<input type="radio" name="BindDomain" value="0"<%If Rs("BindDomain") = 0 Then Response.Write (" checked")%> onClick="setBindDomain.style.display='none';"> 否&nbsp;&nbsp; 
				<input type="radio" name="BindDomain" value="1"<%If Rs("BindDomain") = 1 Then Response.Write (" checked")%> onClick="setBindDomain.style.display='';"<%If Rs("ChannelID") = 5 Then Response.Write (" disabled")%>> 是&nbsp;&nbsp;
				<font color=blue>* 如果启用域名绑定功能，此频道将用你设置的域名访问本频道</font></td>
			</tr>
			<tr id="setBindDomain"<%If Rs("BindDomain") = 0 Then Response.Write (" style=""display:none""")%>>
				<td class="TableRow2"><div class="divbody">频道所绑定的域名：</div></td>
				<td class="TableRow1">
				<input type="text" name="DomainName" size="40" value="<%=Rs("DomainName")%>"> 
				<br><font color="#FF0000">* 请输入你要绑定的域名，如：http://www.enchi.com.cn/</font></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">是否生成HTML：</div></td>
				<td class="TableRow1">
				<input type="radio" name="IsCreateHtml" value="0"<%If Rs("IsCreateHtml") = 0 Then Response.Write (" checked")%>> 否&nbsp;&nbsp; 
				<input type="radio" name="IsCreateHtml" value="1"<%If Rs("IsCreateHtml") = 1 Then Response.Write (" checked")%><%If Rs("ChannelID") = 4 Then Response.Write (" disabled")%>> 是</td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">生成HTML文件的扩展名：</div></td>
				<td class="TableRow1"><input type="text" name="HtmlExtName" size="10" value="<%=Rs("HtmlExtName")%>"> <font color=blue>* 如：“.html”，“.htm”，“.shtml”，“.asp”</font></td>
			</tr>
<%If (Rs("modules") = 6 or Rs("modules") = 7) Then %>
<input type=hidden type="text"   name="HtmlPrefix" size="10" value="<%=Rs("HtmlPrefix")%>"> 
<input type=hidden type="text"   name="HtmlPath" size="10" value="<%=Rs("HtmlPath")%>"> 
<input type=hidden type="text"   name="HtmlForm" size="10" value="<%=Rs("HtmlForm")%>"> 
<% else %>
			<tr>
				<td class="TableRow1"><div class="divbody">生成HTML文件的前缀：</div></td>
				<td class="TableRow1"><input type="text" name="HtmlPrefix" size="10" value="<%=Rs("HtmlPrefix")%>"> <font color=blue>* 格式如：“<%=Rs("HtmlPrefix")%>12345.html”，“<%=Rs("HtmlPrefix")%>list123_1.html”</font></td>
			</tr>
			
			<tr>
				<td class="TableRow1"><div class="divbody">按日期保存HTML文件的路径格式：</div></td>
				<td class="TableRow1">
				<select  size="1" name="HtmlPath" onChange="chkselect(options[selectedIndex].value,'know2');">
				<option value="0"<%If Rs("HtmlPath") = 0 Then Response.Write (" selected")%>>不使用日期目录</option>
				
				<option value="1"<%If Rs("HtmlPath") = 1 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,1)%></option>
				<option value="2"<%If Rs("HtmlPath") = 2 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,2)%></option>
				<option value="3"<%If Rs("HtmlPath") = 3 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,3)%></option>
				<option value="4"<%If Rs("HtmlPath") = 4 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,4)%></option>
				<option value="5"<%If Rs("HtmlPath") = 5 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,5)%></option>
				<option value="6"<%If Rs("HtmlPath") = 6 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,6)%></option>
				<option value="7"<%If Rs("HtmlPath") = 7 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,7)%></option>
				<option value="8"<%If Rs("HtmlPath") = 8 Then Response.Write (" selected")%>><%=enchiasp.ShowDatePath(tempstr,8)%></option>
				
				
				</select> <font color=blue>此目录是根据添加内容的日期生成，相对于各分类目录下面,单页面图文频道无效</font><div id="know2" style="color: red;font-weight:bold;"></div></td>
			</tr>
		
			<tr>
				<td class="TableRow1"><div class="divbody">保存HTML文件的格式：</div></td>
				<td class="TableRow1">
				<select size="1" name="HtmlForm" onChange="chkselect(options[selectedIndex].value,'know1');">
				<option value="0"<%If Rs("HtmlForm") = 0 Then Response.Write (" selected")%>>日期和时间</option>
				<option value="1"<%If Rs("HtmlForm") = 1 Then Response.Write (" selected")%>><%=sModuleName%>ID</option>
				<option value="2"<%If Rs("HtmlForm") = 2 Then Response.Write (" selected")%>>文件前缀+<%=sModuleName%>ID</option>
				<option value="3"<%If Rs("HtmlForm") = 3 Then Response.Write (" selected")%>>日期+<%=sModuleName%>ID</option>
				<option value="4"<%If Rs("HtmlForm") = 4 Then Response.Write (" selected")%>>随机数+<%=sModuleName%>ID</option>
				</select><div id="know1" style="color: red;font-weight:bold;"></div></td>
			</tr>
<%end if%>
			<tr>
				<td class="TableRow1"><div class="divbody">是否允许用户上传文件：</div></td>
				<td class="TableRow1">
				<input type="radio" name="StopUpload" value="1"<%If Rs("StopUpload") = 1 Then Response.Write (" checked")%>> 否&nbsp;&nbsp; 
				<input type="radio" name="StopUpload" value="0"<%If Rs("StopUpload") = 0 Then Response.Write (" checked")%>> 是</td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">允许上传文件的大小：</div></td>
				<td class="TableRow1"><input type="text" name="MaxFileSize" size="10" value="<%=Rs("MaxFileSize")%>"> <b>KB</b><font color=red>（请不要超过<%=Cstr(enchiasp.UploadFileSize)%>KB，如果要超过，请从[基本设置]中修改上传文件大小的上限）</font></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">允许上传文件的类型：<br>多种文件类型之间用“|”分隔</div></td>
				<td class="TableRow1"><input type="text" name="UpFileType" size="60" value="<%=Rs("UpFileType")%>"></td>
			</tr>
			
			
<%If (Rs("modules") = 6 or Rs("modules") = 7) Then %>

<input type=hidden type="text"  name="AppearGrade" size="10" value="<%=Rs("AppearGrade")%>">
<input type=hidden type="text"  name="PostGrade" size="10" value="<%=Rs("PostGrade")%>">

<input type=hidden name="IsAuditing" value="<%=Rs("IsAuditing")%>">



<% else %>	
<tr>
				<td class="TableRow1"><div class="divbody">是否开启审核功能：</div></td>
				<td class="TableRow1">
				<input type="radio" name="IsAuditing" value="0"<%If Rs("IsAuditing") = 0 Then Response.Write (" checked")%>> 关闭&nbsp;&nbsp; 
				<input type="radio" name="IsAuditing" value="1"<%If Rs("IsAuditing") = 1 Then Response.Write (" checked")%>> 打开</td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">
<%
				If Rs("ChannelID") = 4 Then
					Response.Write "发表留言"
				Else
					Response.Write "发表评论"
				End If
%>
的用户等级：</div></td>
				<td class="TableRow1"><select size="1" name="AppearGrade">
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & """"
		If Rs("AppearGrade") = RsObj("Grades") Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>		</select><font color=red>（当生成HTML文件后无效）</font></td>
			</tr>
				
			<tr>
				<td class="TableRow1"><div class="divbody">
				<%
				If Rs("ChannelID") = 4 Then
					Response.Write "回复留言"
				Else
					Response.Write "发布" & sModuleName
				End If
				%>的用户等级：</div></td>
				<td class="TableRow1"><select size="1" name="PostGrade">
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & """"
		If Rs("PostGrade") = RsObj("Grades") Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>		</select></td>
			</tr>
<%end if%>
<%If (Rs("modules") = 6 or Rs("modules") = 7) Then %>
<input type=hidden type="text"  name="LeastString" size="10" value="<%=Rs("LeastString")%>">
<input type=hidden type="text"   name="MaxString" size="10" value="<%=Rs("MaxString")%>">
	<%If Rs("modules") = 7 Then %>
	<tr>
				<td class="TableRow1"><div class="divbody">每页显示列表数：</div></td>
				<td class="TableRow1"><input type="text" name="PaginalNum" size="10" value="<%=Rs("PaginalNum")%>"></td>
			</tr>
	<% else %>
		<input type=hidden name="PaginalNum" size="10" value="<%=Rs("PaginalNum")%>">
	<%end if%>			
			

<input type=hidden type="text"   name="LeastHotHist" size="10" value="<%=Rs("LeastHotHist")%>">
<% else %>
			
			<tr>
				<td class="TableRow1"><div class="divbody">最小评论留言字符：</div></td>
				<td class="TableRow1"><input type="text" name="LeastString" size="10" value="<%=Rs("LeastString")%>"></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">最大评论留言字符：</div></td>
				<td class="TableRow1"><input type="text" name="MaxString" size="10" value="<%=Rs("MaxString")%>"></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">每页显示列表数：</div></td>
				<td class="TableRow1"><input type="text" name="PaginalNum" size="10" value="<%=Rs("PaginalNum")%>"></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">最小热门点击数：</div></td>
				<td class="TableRow1"><input type="text" name="LeastHotHist" size="10" value="<%=Rs("LeastHotHist")%>"></td>
			</tr>
<%end if%>	
<%
If Rs("modules") = 2 Then
%>
			<tr>
				<td class="TableRow1"><div class="divbody">设置软件运行环境：</div><br>每个运行环境请用“|”分开</td>
				<td class="TableRow1"><textarea name="ChannelSetting" cols="60" rows="3"><%=Channel_Setting(0)%></textarea></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">设置软件默认运行环境：</div></td>
				<td class="TableRow1"><input type="text" name="ChannelSetting" size="60" value="<%=Channel_Setting(1)%>"></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">设置软件类型：</div><br>每个软件类型请用“,”分开</td>
				<td class="TableRow1"><textarea name="ChannelSetting" cols="60" rows="3"><%=Channel_Setting(2)%></textarea></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">设置软件授权方式：</div><br>每种授权方式请用“,”分开</td>
				<td class="TableRow1"><input type="text" name="ChannelSetting" size="60" value="<%=Channel_Setting(3)%>"></td>
			</tr>
			<tr>
				<td class="TableRow1"><div class="divbody">设置软件语言：</div><br>每种软件语言请用“,”分开</td>
				<td class="TableRow1"><input type="text" name="ChannelSetting" size="60" value="<%=Channel_Setting(4)%>"></td>
			</tr>
<%
	Else
		Response.Write "<input type=""hidden"" name=""ChannelSetting"" value=""|||"">"
		Response.Write "<input type=""hidden"" name=""ChannelSetting"" value=""@@@"">"
		Response.Write "<input type=""hidden"" name=""ChannelSetting"" value=""@@@"">"
	End If
%>
		</table></fieldset></td>
		</tr>
		<tr>
			<td class="TableRow2">　</td>
			<td class="TableRow1" align="center"><input type="button" onclick="javascript:history.go(-1)" value="返回上一页" name="B1" class=Button>&nbsp;&nbsp;
			<input type="submit" value="保存设置" name="B2" class=Button></td>
		</tr>
		</form>
	</table>
<div id="Issubport0" style="display:none"><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),0,"")%></div>
<div id="Issubport1" style="display:none"><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),1,"")%></div>
<div id="Issubport2" style="display:none"><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),2,"")%></div>
<div id="Issubport3" style="display:none"><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),3,"")%></div>
<div id="Issubport4" style="display:none"><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),4,"")%></div>
<div id="Issubport5" style="display:none">不使用日期目录,HTML文件将保存到分类目录下面<br><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>分类目录/<%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport6" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>分类目录/<%=enchiasp.ShowDatePath(tempstr,1)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport7" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>分类目录/<%=enchiasp.ShowDatePath(tempstr,2)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport8" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>分类目录/<%=enchiasp.ShowDatePath(tempstr,3)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport9" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>分类目录/<%=enchiasp.ShowDatePath(tempstr,4)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport10" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>分类目录/<%=enchiasp.ShowDatePath(tempstr,5)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport11" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>分类目录/<%=enchiasp.ShowDatePath(tempstr,6)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport12" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>分类目录/<%=enchiasp.ShowDatePath(tempstr,7)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<div id="Issubport13" style="display:none"><%=enchiasp.GetChannelDir(Rs("ChannelID"))%>分类目录/<%=enchiasp.ShowDatePath(tempstr,8)%><%=enchiasp.ReadFileName(tempstr,9988,Rs("HtmlExtName"),Rs("HtmlPrefix"),Rs("HtmlForm"),"")%></div>
<SCRIPT LANGUAGE="JavaScript">
<!--
function chkselect(s,divid)
{
	var divname='Issubport';
	var chkreport;
	s=Number(s)
	if (divid=="know1")
	{
		divname=divname+s;
	}
	if (divid=="know2")
	{
		s+=5;
		divname=divname+s;
	}
	document.getElementById(divid).innerHTML=divname;
	chkreport=document.getElementById(divname).innerHTML;
	document.getElementById(divid).innerHTML=chkreport;
}
//-->
</SCRIPT>
<%
Rs.Close:Set Rs = Nothing
End Sub

Private Sub CheckSave()
	If Len(Request.Form("ChannelName")) = 0 Or Len(Request.Form("ChannelName")) => 25 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>网站频道名称不能为空或者超过20个字符！</li>"
	End If
	If Len(Request.Form("ColorModes")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>标题颜色参数错误！</li>"
	End If
	If Len(Request.Form("Caption")) = 0 Or Len(Request.Form("Caption")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>频道注释不能为空或者超过200个字符！</li>"
	End If
	If Len(Request.Form("ChannelUrl")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>频道连接URL不能为空！</li>"
	End If
	
End Sub

Private Sub SavenewChannel()
	CheckSave
	Dim neworders
	If Len(Request.Form("ChannelDir")) = 0 And Request.Form("ChannelType") <> 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>频道所在目录不能为空！</li>"
	End If
	ChannelDir = Replace(Replace(Replace(Request.Form("ChannelDir"), "\","/"), " ",""), "'","")
	If Right(ChannelDir, 1) <> "/" Then
		ChannelDir = ChannelDir & "/"
	Else
		ChannelDir = ChannelDir
	End If
	If Request.Form("ChannelType") = 1 Then
		If Request.Form("modules") = 0 Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择正确的模块！</li>"
			Exit Sub
		End If
		Set Rs = Conn.Execute("SELECT ChannelID,ChannelDir FROM ECCMS_Channel WHERE ChannelType=0 And ChannelID=" & CLng(Request.Form("modules")))
		If Rs.EOF And Rs.BOF Then
			ErrMsg = "<li>找不到指定模块。</li>"
			Founderr = True
			Exit Sub
		Else
			mChannelID = Rs("ChannelID")
			mChannelDir = Rs("ChannelDir")
			If LCase(ChannelDir) = LCase(mChannelDir) Then
				ErrMsg = "<li>不能指定和系统频道相同的目录。</li>"
				Founderr = True
				Exit Sub
			End If
		End If
		Set Rs = Nothing
	End If
	
	Set Rs = Conn.Execute("SELECT ChannelID FROM ECCMS_Channel WHERE ChannelID=" & CLng(Request("NewChannelID")))
	If Not (Rs.EOF And Rs.BOF) Then
		ErrMsg = "<li>您不能指定和别的频道一样的序号。</li>"
		Founderr = True
		Exit Sub
	Else
		NewChannelID = CLng(Request("NewChannelID"))
	End If
	Set Rs = Nothing
	If NewChannelID = 999 Then NewChannelID = NewChannelID + 1
	If NewChannelID = 9999 Then NewChannelID = NewChannelID + 1
	If Founderr = True Then Exit Sub
	Set Rs = enchiasp.Execute ("SELECT MAX(orders) FROM ECCMS_Channel")
	If Not (Rs.EOF And Rs.bof) Then
		neworders = Rs(0)
	End If
	If IsNull(neworders) Then neworders = 0
	Set Rs = Nothing
	'Call ChannelCopy
	'Succeed("<li>添加新的频道成功</li>"):exit sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_Channel"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = NewChannelID
		Rs("orders") = neworders + 1
		Rs("ColorModes") = Trim(Request.Form("ColorModes"))
		Rs("FontModes") = Trim(Request.Form("FontModes"))
		Rs("ChannelName") = enchiasp.ChkFormStr(Request.Form("ChannelName"))
		Rs("Caption") = enchiasp.ChkFormStr(Request.Form("Caption"))
		Rs("ChannelDir") = ChannelDir
		Rs("StopChannel") = 0
		Rs("IsHidden") = Trim(Request.Form("IsHidden"))
		Rs("LinkTarget") = Trim(Request.Form("LinkTarget"))
		Rs("ChannelType") = CInt(Request.Form("ChannelType"))
		Rs("ChannelUrl") = Trim(Request.Form("ChannelUrl"))
		Rs("modules") = CInt(Request.Form("modules"))
		Rs("BindDomain") = 0
		Rs("DomainName") = "http://"
		If CInt(Request.Form("ChannelType")) = 1 Then
			Rs("ModuleName") = "新频道"
		Else
			Rs("ModuleName") = "外部"
		End If
		
		Rs("ChannelSkin") = 0
		Rs("HtmlPath") = 0
		Rs("HtmlForm") = 3
		Rs("IsCreateHtml") = 0
		Rs("HtmlExtName") = ".html"
		Rs("HtmlPrefix") = "HTML_"
		Rs("StopUpload") = 1
		Rs("MaxFileSize") = 500
		Rs("UpFileType") = "rar|zip|exe|gif|jpg|png|bmp|swf"
		Rs("IsAuditing") = 1
		Rs("AppearGrade") = 0
		Rs("PostGrade") = 0
		Rs("LeastString") = 10
		Rs("MaxString") = 500
		Rs("PaginalNum") = 15
		Rs("LeastHotHist") = 50
		If CInt(Request.Form("modules")) = 2 Then
			Rs("Channel_Setting") = "Win2003/|WinNet/|WinXP/|Win2000/|NT/|WinME/|Win9X/|Linux/|Unix/|Mac/|||Win9X/Win2000/WinXP/Win2003/|||国产软件,国外软件,汉化补丁,病毒防治|||共享软件,免费软件,自由软件,试用软件,演示软件,商业软件|||简体中文,繁体中文,英文|||"
		Else
			Rs("Channel_Setting") = "|||@@@|||@@@|||"
		End If
	Rs.update
	Rs.Close:Set Rs = Nothing
	enchiasp.DelCahe "ChannelMenu"
	Succeed("<li>添加新的频道成功</li>")
	If CInt(Request.Form("modules")) > 0 And CInt(Request.Form("ChannelType")) = 1 Then
		Call ChannelCopy
	End If
	
End Sub
Private Sub ChannelCopy()
	Dim newChannelDir,oldChannelDir
	Dim tmpChannel,tmpChannelArray
	oldChannelDir = enchiasp.InstallDir & mChannelDir
	newChannelDir = enchiasp.InstallDir & ChannelDir
	enchiasp.CreatPathEx(newChannelDir & "js")
	enchiasp.CreatPathEx(newChannelDir & "special")
	enchiasp.CreatPathEx(newChannelDir & "UploadPic")
	enchiasp.CreatPathEx(newChannelDir & "UploadFile")
	enchiasp.CopyToFile oldChannelDir & "index.asp",newChannelDir & "index.asp"
	enchiasp.CopyToFile oldChannelDir & "list.asp",newChannelDir & "list.asp"
	enchiasp.CopyToFile oldChannelDir & "show.asp",newChannelDir & "show.asp"
	enchiasp.CopyToFile oldChannelDir & "special.asp",newChannelDir & "special.asp"
	enchiasp.CopyToFile oldChannelDir & "search.asp",newChannelDir & "search.asp"
	enchiasp.CopyToFile oldChannelDir & "showbest.asp",newChannelDir & "showbest.asp"
	enchiasp.CopyToFile oldChannelDir & "showhot.asp",newChannelDir & "showhot.asp"
	enchiasp.CopyToFile oldChannelDir & "shownew.asp",newChannelDir & "shownew.asp"
	enchiasp.CopyToFile oldChannelDir & "comment.asp",newChannelDir & "comment.asp"
	enchiasp.CopyToFile oldChannelDir & "Hits.Asp",newChannelDir & "Hits.Asp"
	enchiasp.CopyToFile oldChannelDir & "RemoveCache.Asp",newChannelDir & "RemoveCache.Asp"
	enchiasp.CopyToFile oldChannelDir & "rssfeed.asp",newChannelDir & "rssfeed.asp"
	enchiasp.CopyToFile oldChannelDir & "js/ShowPage.JS",newChannelDir & "js/ShowPage.JS"
	enchiasp.CopyToFile oldChannelDir & "js/Show_Page.JS",newChannelDir & "js/Show_Page.JS"
	tmpChannel = enchiasp.ReadFile("include/Channel.dat")
	tmpChannel = Replace(tmpChannel, "$ChannelID$", NewChannelID,1,-1,1)
	tmpChannelArray = Split(tmpChannel, "@@@")
	If CInt(Request.Form("modules")) = 1 Then
		enchiasp.CopyToFile oldChannelDir & "sendmail.asp",newChannelDir & "sendmail.asp"
		enchiasp.CreatedTextFile newChannelDir & "config.asp",tmpChannelArray(0)
	ElseIf CInt(Request.Form("modules")) = 2 Then
		enchiasp.CopyToFile oldChannelDir & "showtype.asp",newChannelDir & "showtype.asp"
		enchiasp.CopyToFile oldChannelDir & "error.asp",newChannelDir & "error.asp"
		enchiasp.CopyToFile oldChannelDir & "download.asp",newChannelDir & "download.asp"
		enchiasp.CopyToFile oldChannelDir & "softdown.asp",newChannelDir & "softdown.asp"
		enchiasp.CopyToFile oldChannelDir & "previewimg.asp",newChannelDir & "previewimg.asp"
		enchiasp.CreatedTextFile newChannelDir & "config.asp",tmpChannelArray(1)
	'单页面图文
	Elseif CInt(Request.Form("modules")) = 6 then
		enchiasp.CreatedTextFile newChannelDir & "config.asp",tmpChannelArray(0)

	else
		enchiasp.CopyToFile oldChannelDir & "download.asp",newChannelDir & "download.asp"
		enchiasp.CopyToFile oldChannelDir & "down.asp",newChannelDir & "down.asp"
		enchiasp.CopyToFile oldChannelDir & "downfile.asp",newChannelDir & "downfile.asp"
		enchiasp.CopyToFile oldChannelDir & "play.html",newChannelDir & "play.html"
		enchiasp.CreatedTextFile newChannelDir & "config.asp",tmpChannelArray(2)
	End If
	Dim rstmp,i
	Dim TemplateDir,TemplateFields,TemplateValues
	Set rstmp = enchiasp.Execute("SELECT * FROM ECCMS_Template WHERE ChannelID=" & CLng(Request.Form("modules")))
	SQL=rstmp.GetRows(-1)
	Set rstmp = Nothing
	For i=0 To Ubound(SQL,2)
		TemplateDir = ""
		TemplateFields = "ChannelID,skinid,pageid,TemplateDir,page_name,page_content,page_setting,Template_Help,isDefault"
		TemplateValues = "" & NewChannelID & ","& SQL(2,i) &"," & SQL(3,i) & ",'" & TemplateDir & "','" & enchiasp.CheckStr(SQL(5,i)) & "','" & enchiasp.CheckStr(SQL(6,i)) & "','" & enchiasp.CheckStr(SQL(7,i)) & "','" & enchiasp.CheckStr(SQL(8,i)) & "'," & SQL(9,i) & ""
		Conn.Execute ("INSERT INTO ECCMS_Template (" & TemplateFields & ") VALUES (" & TemplateValues & ")")
	Next
	SQL=Null
End Sub

Private Sub SaveditChannel()
	CheckSave
	Dim HtmlExtName,sDomainName
	If Len(Request.Form("ChannelDir")) = 0 And Request.Form("ChannelType") <> 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>频道所在目录不能为空！</li>"
	End If
	ChannelDir = Replace(Replace(Replace(Request.Form("ChannelDir"), "\","/"), " ",""), "'","")
	If Right(ChannelDir, 1) <> "/" Then
		ChannelDir = ChannelDir & "/"
	Else
		ChannelDir = ChannelDir
	End If
	If Trim(Request.Form("IsCreateHtml")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择是否生成HTML文件！</li>"
	End If
	If Left(Trim(Request.Form("HtmlExtName")),1) <> "." Then
		HtmlExtName = "." & Trim(Request.Form("HtmlExtName"))
	Else
		HtmlExtName = Trim(Request.Form("HtmlExtName"))
	End If
	If Not enchiasp.IsValidChar(Request.Form("HtmlExtName")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>文件扩展名中含有非法字符或者中文字符！</li>"
	End If
	If Not enchiasp.IsValidChar(ChannelDir) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>频道目录中含有非法字符或者中文字符！</li>"
	End If
	If Not IsNumeric(Request("MaxFileSize")) Then
		ErrMsg = ErrMsg & "<li>上传文件大小请使用整数！</li>"
		Founderr = True
	End If
	if  CLng(Request("MaxFileSize"))>  CLng(enchiasp.UploadFileSize) then
		ErrMsg = ErrMsg & "<li>上传文件大小超过系统设置的"&CLng(enchiasp.UploadFileSize)&"KB，如果有必要请修改[基本设置]中的[系统基本设置]！</li>"
		Founderr = True

	end if

	If Not IsNumeric(Request("LeastString")) Then
		ErrMsg = ErrMsg & "<li>最小字符请使用整数！</li>"
		Founderr = True
	End If
	If Not IsNumeric(Request("MaxString")) Then
		ErrMsg = ErrMsg & "<li>最大字符请使用整数！</li>"
		Founderr = True
	End If
	If Not IsNumeric(Request("PaginalNum")) Then
		ErrMsg = ErrMsg & "<li>每页显示列表数请使用整数！</li>"
		Founderr = True
	End If
	If Not IsNumeric(Request("LeastHotHist")) Then
		ErrMsg = ErrMsg & "<li>最小热门点击数请使用整数！</li>"
		Founderr = True
	End If
	sDomainName = Replace(Replace(Replace(Request.Form("DomainName"), "\","/"), " ",""), "'","")
	If Right(sDomainName, 1) <> "/" Then
		sDomainName = sDomainName & "/"
	Else
		sDomainName = sDomainName
	End If
	Dim TempStr, ChannelSetting
	For Each TempStr In Request.Form("ChannelSetting")
			ChannelSetting = ChannelSetting & Replace(TempStr, "|||", "") & "|||"
	Next
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Channel where ChannelID = " & Request("ChannelID")
	Rs.Open SQL,Conn,1,3
		Rs("ColorModes") = Trim(Request.Form("ColorModes"))
		Rs("FontModes") = Trim(Request.Form("FontModes"))
		Rs("ChannelName") = enchiasp.ChkFormStr(Request.Form("ChannelName"))
		Rs("Caption") = enchiasp.ChkFormStr(Request.Form("Caption"))
		Rs("ChannelDir") = Trim(ChannelDir)
		Rs("StopChannel") = Trim(Request.Form("StopChannel"))
		Rs("IsHidden") = Trim(Request.Form("IsHidden"))
		Rs("LinkTarget") = Trim(Request.Form("LinkTarget"))
		Rs("ChannelType") = Trim(Request.Form("ChannelType"))
		Rs("ChannelUrl") = Trim(Request.Form("ChannelUrl"))
		Rs("ModuleName") = Trim(Request.Form("ModuleName"))
		Rs("BindDomain") = Trim(Request.Form("BindDomain"))
		Rs("DomainName") = Trim(sDomainName)
		Rs("ChannelSkin") = Trim(Request.Form("ChannelSkin"))
		Rs("HtmlPath") = Trim(Request.Form("HtmlPath"))
		Rs("HtmlForm") = Trim(Request.Form("HtmlForm"))
		Rs("IsCreateHtml") = Trim(Request.Form("IsCreateHtml"))
		Rs("HtmlExtName") = HtmlExtName
		Rs("HtmlPrefix") = Trim(Request.Form("HtmlPrefix"))
		Rs("StopUpload") = Trim(Request.Form("StopUpload"))
		Rs("MaxFileSize") = CLng(Request.Form("MaxFileSize"))
		Rs("UpFileType") = Trim(Request.Form("UpFileType"))
		Rs("IsAuditing") = Trim(Request.Form("IsAuditing"))
		Rs("AppearGrade") = Trim(Request.Form("AppearGrade"))
		Rs("PostGrade") = Trim(Request.Form("PostGrade"))
		Rs("LeastString") = CLng(Request.Form("LeastString"))
		Rs("MaxString") = CLng(Request.Form("MaxString"))
		Rs("PaginalNum") = CInt(Request.Form("PaginalNum"))
		Rs("LeastHotHist") = CLng(Request.Form("LeastHotHist"))
		Rs("Channel_Setting") = Trim(ChannelSetting)
	Rs.update
	Rs.Close
	Set Rs = Nothing
	Call RemoveCache
	Succeed("<li>修改频道设置成功！</li>")
End Sub

Private Sub ChannelDel()
	If Request("ChannelID") = "" Then
		ErrMsg = "<li>请选择正确的频道ID！</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") < 10 Then
		ErrMsg = "<li>此频道为系统初始频道不能删除，请选择其它频道删除！</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT ClassID FROM [ECCMS_Classify] WHERE ChannelID=" & CLng(Request("ChannelID")))
	If Not (Rs.BOF And Rs.EOF) Then
		Set Rs = Nothing
		ErrMsg = "<li>此频道正在使用中不能删除！如果要删除此频道，请先删除所有分类。</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = Nothing
	Set Rs = enchiasp.Execute("SELECT ChannelDir,ChannelType FROM [ECCMS_Channel] WHERE ChannelID=" & CLng(Request("ChannelID")))
	If Not (Rs.BOF And Rs.EOF) Then
		If Rs("ChannelType") = 0 Then
			Set Rs = Nothing
			ErrMsg = "<li>此频道是系统频道不能删除。</li>"
			Founderr = True
			Exit Sub
		Else
			enchiasp.FolderDelete(enchiasp.InstallDir & Rs("ChannelDir"))
			Conn.Execute("DELETE FROM ECCMS_Template WHERE ChannelID=" & CLng(Request("ChannelID")))
		End If
	End If
	Set Rs = Nothing
	Call RemoveCache
	
	Conn.Execute("DELETE FROM ECCMS_Channel WHERE ChannelID=" & CLng(Request("ChannelID")))
	Succeed("<li>频道删除成功！</li>")
End Sub
Private Sub ChannelOrders()
	Dim trs
	Dim uporders
	Dim doorders
	Response.Write " <table border=""0"" cellspacing=""1"" cellpadding=""2"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th colspan=2>频道重新排序修改"
	Response.Write " </th>"
	Response.Write " </tr>" & vbCrLf
	SQL = "select * from ECCMS_Channel order by orders"
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		Response.Write "您还没有添加相应的频道。"
	Else
		Do While Not Rs.EOF
			Response.Write "<form action=?action=saveorder method=post><tr><td width=""50%"" class=TableRow1>" & vbCrLf
			Response.Write enchiasp.ReadFontMode(Rs("ChannelName"),Rs("ColorModes"),Rs("FontModes"))
			Response.Write "</td><td width=""50%"" class=TableRow2>" & vbCrLf
			Set trs = enchiasp.Execute("select count(*) from ECCMS_Channel where orders<" & Rs("orders") & "")
				uporders = trs(0)
				If IsNull(uporders) Then uporders = 0

				Set trs = enchiasp.Execute("select count(*) from ECCMS_Channel where orders>" & Rs("orders") & "")
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
					Response.Write "<input type=hidden name=""ChannelID"" value=""" & Rs("ChannelID") & """>&nbsp;<input type=submit name=Submit class=button value='修 改'>" & vbCrLf
				End If
			Response.Write "</td></tr></form>" & vbCrLf
			Rs.movenext
		Loop
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>"
End Sub

Private Sub SaveOrder()
	Dim orders
	Dim uporders
	Dim doorders
	Dim oldorders
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("uporders") <> "" And Not CInt(Request("uporders")) = 0 Then
		If Not IsNumeric(Request("uporders")) Then
			ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
			Founderr = True
			Exit Sub
		ElseIf CInt(Request("uporders")) = 0 Then
			ErrMsg = ErrMsg & "<li>请选择要提升的数字！</li>"
			Founderr = True
			Exit Sub
		End If
		Set Rs = enchiasp.Execute("select ChannelID,orders from ECCMS_Channel where ChannelID=" & Request("ChannelID"))
		orders = Rs(1)
		i = 0
		oldorders = 0
		Set Rs = enchiasp.Execute("select ChannelID,orders from ECCMS_Channel where orders<" & orders & " order by orders desc")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("uporders")) >= i Then
				enchiasp.Execute ("update ECCMS_Channel set orders=" & orders & "+" & oldorders & " where ChannelID=" & Rs(0))
				If CInt(Request("uporders")) = i Then uporders = Rs(1)
			End If
			orders = Rs(1)
			Rs.movenext
		Loop
		enchiasp.Execute ("update ECCMS_Channel set orders=" & uporders & " where ChannelID=" & Request("ChannelID"))
		Set Rs = Nothing
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
		Set Rs = enchiasp.Execute("select ChannelID,orders from ECCMS_Channel where ChannelID=" & Request("ChannelID"))
		orders = Rs(1)
		i = 0
		oldorders = 0
		Set Rs = enchiasp.Execute("select ChannelID,orders from ECCMS_Channel where orders>" & orders & " order by orders desc")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("doorders")) >= i Then
				enchiasp.Execute ("update ECCMS_Channel set orders=" & orders & " where ChannelID=" & Rs(0))
				If CInt(Request("doorders")) = i Then doorders = Rs(1)
			End If
			orders = Rs(1)
			Rs.movenext
		Loop
		enchiasp.Execute ("update ECCMS_Channel set orders=" & doorders & " where ChannelID=" & Request("ChannelID"))
		Set Rs = Nothing
	End If
	Call RemoveCache
	Response.redirect "admin_channel.asp?action=orders"
End Sub

Private Sub UpdateStop()
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") <> "" And Request("EditID") <> ""  Then
		enchiasp.Execute ("update ECCMS_Channel set StopChannel=" & CInt(Request("EditID")) & " where ChannelID=" & Request("ChannelID"))
		Call RemoveCache
	Else
		ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("EditID") <> 0  Then
		OutHintScript("恭喜您！本频道已成功关闭。")
	Else
		OutHintScript("恭喜您！本频道已成功打开。")
	End If
End Sub

Private Sub UpdateHidden()
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") <> "" And Request("EditID") <> ""  Then
		enchiasp.Execute ("update ECCMS_Channel set IsHidden=" & CInt(Request("EditID")) & " where ChannelID=" & Request("ChannelID"))
		Call RemoveCache
	Else
		ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("EditID") <> 0  Then
		OutHintScript("恭喜您！隐藏频道菜单成功。")
	Else
		OutHintScript("恭喜您！显示频道菜单成功。")
	End If
End Sub

Private Sub UpdateLinkTarget()
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") <> "" And Request("EditID") <> ""  Then
		enchiasp.Execute ("update ECCMS_Channel set LinkTarget=" & CInt(Request("EditID")) & " where ChannelID=" & Request("ChannelID"))
		Call RemoveCache
		OutHintScript("恭喜您！更新连接目标成功。")
	Else
		ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
End Sub

Private Sub UpdateCreateHtml()
	If Not IsNumeric(Request("ChannelID")) Then
		ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("ChannelID") <> "" And Request("EditID") <> ""  Then
		enchiasp.Execute ("update ECCMS_Channel set IsCreateHtml=" & CInt(Request("EditID")) & " where ChannelID=" & Request("ChannelID"))
		Call RemoveCache
	Else
		ErrMsg = ErrMsg & "<li>非法的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	If Request("EditID") <> 0  Then
		OutHintScript("恭喜您！打开此频道生成HTML功能成功。")
	Else
		OutHintScript("恭喜您！关闭此频道生成HTML功能成功。")
	End If
End Sub
Private Sub ReloadChannelCache()
	enchiasp.DelCahe "Channel" & Request("ChannelID")
	enchiasp.DelCahe "MyChannel" & Request("ChannelID")
	enchiasp.DelCahe "ChannelMenu"
	enchiasp.DelCahe "SiteClassMap"
	Response.Write "<script>alert('更新缓存成功！');javascript:history.back(1)</script>"
End Sub
Private Sub RemoveCache()
	enchiasp.DelCahe "Channel" & Request("ChannelID")
	enchiasp.DelCahe "MyChannel" & Request("ChannelID")
	enchiasp.DelCahe "ChannelMenu"
	enchiasp.DelCahe "SiteClassMap"
End Sub

%>