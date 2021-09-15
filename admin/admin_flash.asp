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
Dim Action,isEdit,Flag,DefaultShowMode
Dim i,ClassID,RsObj,flashid,findword,keyword,strClass
Dim TextContent,FlashTop,FlashBest,ForbidEssay
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Dim s_ClassName,ChildStr,FoundSQL,isAccept,selflashid
Dim FlashAccept,Auditing

ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID = 0 Then ChannelID = 5
If ChannelID = 5 Then
	DefaultShowMode = 1		'-- 默认显示模式
Else
	DefaultShowMode = 2		'-- 默认显示模式
End If

Flag = sChannelDir & ChannelID
If Request("isAccept") <> "" Then
	isAccept = 0
Else
	isAccept = 1
End If
Action = LCase(Request("action"))
If Not ChkAdmin(Flag) Then
	Server.Transfer("showerr.asp")
	Response.End
End If

Select Case Trim(Action)
Case "save"
	Call SaveFlash
Case "modify"
	Call ModifyFlash
Case "add"
	isEdit = False
	Call FlashEdit(isEdit)
Case "edit"
	isEdit = True
	Call FlashEdit(isEdit)
Case "del"
	Call FlashDel
Case "view"
	Call FlashView
Case "setting"
	Call PageTop
	Call BatchSetting
Case "saveset"
	Call SaveSetting
Case "move"
	Call PageTop
	Call BatchMove
Case "savemove"
	Call SaveMove
Case "batdel"
	Call PageTop
	Call BatcDelete
Case "alldel"
	Call AllDelFlash
Case Else
	Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub PageTop()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th colspan=2>" & sModuleName & "管理选项</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr><form method=Post name=myform action='admin_flash.asp' onSubmit='return JugeQuery(this);'>"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "<td class=TableRow1>搜索："
	Response.Write "	  <input name=keyword type=text size=20>"
	Response.Write "	  条件："
	Response.Write "	  <select name=queryopt>"
	Response.Write "		<option value='1' selected>" & sModuleName & "名称</option>"
	Response.Write "		<option value='2'>添 加 人</option>"
	Response.Write "		<option value='3'>不限条件</option>"
	Response.Write "	  </select> <input type=submit name=Submit value='开始查询' class=Button></td>"
	Response.Write "	  <td class=TableRow1>" & sModuleName & "导航："
	Dim srtClassMenu
	Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Response.Write "<option value='admin_flash.asp?ChannelID=" & ChannelID & "'>≡全部" & sModuleName & "列表≡</option>" & vbCrLf
	srtClassMenu = enchiasp.ClassJumpMenu(ChannelID)
	srtClassMenu = Replace(srtClassMenu, "{ClassID=" & Request("sortid") & "}", "selected")
	Response.Write srtClassMenu
	Response.Write "</select>"
	Response.Write "	  </td></form>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td colspan=2 class=TableRow2><strong>操作选项：</strong> <a href='admin_flash.asp?ChannelID=" & ChannelID & "'>管理首页</a> | "
	Response.Write "	  <a href='admin_flash.asp?action=add&ChannelID=" & ChannelID & "'>添加" & sModuleName & "</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "&action=add'>添加" & sModuleName & "分类</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "'>" & sModuleName & "分类管理</a></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub
Private Sub showmain()
	If Not IsEmpty(Request("selflashid")) Then
		selflashid = Request("selflashid")
		Select Case enchiasp.CheckStr(Request("act"))
		Case "批量删除":Call batdel
		Case "批量移动":Call batmove
		Case "更新时间":Call upindate
		Case "批量推荐":Call isCommend
		Case "取消推荐":Call noCommend
		Case "批量置顶":Call isTop
		Case "取消置顶":Call noTop
		Case "批量审核":Call BatAccept
		Case "取消审核":Call NotAccept
		Case "生成HTML":Call BatCreateHtml
		Case Else
			Response.Write "无效参数！"
		End Select
	End If
	Call PageTop
	Dim strListName
	Dim specialID,sortid,Cmd,child
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>选择</th>"
	Response.Write "	  <th width='68%'>" & sModuleName & "名称</th>"
	Response.Write "	  <th width='9%' nowrap>管理操作</th>"
	Response.Write "	  <th width='9%' nowrap>录 入 者</th>"
	Response.Write "	  <th width='9%' nowrap>整理日期</th>"
	Response.Write "	</tr>"
	strListName = "&channelid="& ChannelID &"&sortid="& Request("sortid") &"&specialID="& Request("specialID") &"&isAccept="& Request("isAccept") &"&keyword=" & Request("keyword") 
	If Request("sortid") <> "" Then
		SQL = "select ClassID,ChannelID,ClassName,child,ChildStr from [ECCMS_Classify] where ChannelID = " & ChannelID & " And ClassID=" & Request("sortid")
		Set Rs = enchiasp.Execute(SQL)
		If Rs.bof And Rs.EOF Then
			Response.Write "Sorry！没有找到任何" & sModuleName & "分类。或者您选择了错误的系统参数!"
			Response.End
		Else
			s_ClassName = Rs("ClassName")
			ClassID = Rs("ClassID")
			child = Rs("child")
			ChildStr = Rs("ChildStr")
			sortid = CLng(Request("sortid"))
		End If
		Rs.Close
	Else
		s_ClassName = "全部" & sModuleName
		sortid = 0
		child = 0
	End If
	maxperpage = 30 '###每页显示数
	
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
	If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("queryopt")) = 1 Then
			findword = "A.title like '%" & keyword & "%'"
		ElseIf CInt(Request("queryopt")) = 2 Then
			findword = "A.username like '%" & keyword & "%'"
		Else
			findword = "A.title like '%" & keyword & "%' or A.username like '%" & keyword & "%'"
		End If
		FoundSQL = findword
		s_ClassName = "查询" & sModuleName
	Else
		specialID = 99999
		If Trim(Request("sortid")) <> "" Then
			FoundSQL = "A.isAccept="& isAccept & " And A.ClassID in (" & ChildStr & ")"
		Else
			If Trim(Request("specialID")) <> "" Then
				specialID = CLng(Request("specialID"))
				If Request("specialID") <> 0 Then
					FoundSQL = "A.isAccept = " & isAccept & " And specialID =" & Request("specialID")
				Else
					FoundSQL = "A.isAccept = " & isAccept & " And specialID > 0"
				End If
			Else
				FoundSQL = "A.isAccept = " & isAccept
			End If
		End If
	End If
	TotalNumber = enchiasp.Execute("Select Count(flashid) from ECCMS_FlashList A where A.ChannelID = " & ChannelID & " And "& FoundSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '得到总页数
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select A.*,C.ClassName from [ECCMS_FlashList] A inner join [ECCMS_Classify] C on A.ClassID=C.ClassID where A.ChannelID = " & ChannelID & " And "& FoundSQL &" order by A.isTop desc, A.addTime desc ,A.flashid desc"
	Rs.Open SQL, Conn, 1, 1
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=5 class=TableRow2>还没有找到任何" & sModuleName & "！</td></tr>"
	Else
		Rs.MoveFirst
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0

	Response.Write "	<tr>"
	Response.Write "	  <td colspan=""5"" class=""TableRow2"">"
	ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<form name=selform method=post action="""">"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "<input type=hidden name=action value=''>"
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
		Response.Write "	<tr>"
		Response.Write "	  <td align=center " & strClass & "><input type=checkbox name=selflashid value=" & Rs("flashid") & "></td>"
		Response.Write "	  <td " & strClass & ">"
		If Rs("isTop") <> 0 Then
			Response.Write "<img src=""images/istop.gif"" width=15 height=17 border=0 alt=置顶动画>"
		End If

		Response.Write "[<a href=?ChannelID=" & Rs("ChannelID") & "&sortid="
		Response.Write Rs("ClassID")
		Response.Write ">"
		Response.Write Rs("ClassName")
		Response.Write "</a>] "
		Response.Write "<a href=?action=view&ChannelID=" & Rs("ChannelID") & "&flashid="
		Response.Write Rs("flashid")
		Response.Write ">"
		Response.Write enchiasp.ReadFontMode(Rs("title"),Rs("ColorMode"),Rs("FontMode"))
		Response.Write "</a>" 

		If Rs("isBest") <> 0 Then
			Response.Write "&nbsp;&nbsp;<font color=blue>荐</font>"
		End If
%>
	  </td>
	  <td align="center" nowrap <%=strClass%>><a href=?action=edit&ChannelID=<%=Rs("ChannelID")%>&flashid=<%=Rs("flashid")%>>编辑</a> | <a href=?action=del&ChannelID=<%=Rs("ChannelID")%>&flashid=<%=Rs("flashid")%> onclick="{if(confirm('动画删除后将不能恢复，您确定要删除该动画吗?')){return true;}return false;}">删除</a></td>
	  <td align="center" nowrap <%=strClass%>><%=Rs("UserName")%></td>
	  <td align="center" nowrap <%=strClass%>>
<%
		If Rs("addTime") >= Date Then
			Response.Write "<font color=red>"
			Response.Write enchiasp.FormatDate(Rs("addTime"), 2)
			Response.Write "</font>"
		Else
			Response.Write enchiasp.FormatDate(Rs("addTime"), 2)
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
	Set Cmd = Nothing
	Rs.Close:Set Rs = Nothing
%>
	<tr>
	  <td colspan="5" class="TableRow1">
	  <input class=Button type="button" name="chkall" value="全选" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="反选" onClick="ContraSel(this.form)">
	  管理选项：
	  <select name="act">
		<option value="0">请选择操作选项</option>
		<option value="批量删除">批量删除</option>
		<option value="批量置顶">批量置顶</option>
		<option value="取消置顶">取消置顶</option>
		<option value="批量推荐">批量推荐</option>
		<option value="取消推荐">取消推荐</option>
		<option value="更新时间">更新时间</option>
		<option value="生成HTML">生成HTML</option>
	  </select>
	  <input class=Button type="submit" name="Submit2" value="执行操作" onclick="return confirm('您确定执行该操作吗?');">
	  <input class=Button type="submit" name="Submit3" value="批量设置" onclick="document.selform.action.value='setting';">
	  <input class=Button type="submit" name="Submit4" value="批量移动" onclick="document.selform.action.value='move';">
	  <input class=Button type="submit" name="Submit4" value="批量删除" onclick="document.selform.action.value='batdel';"></td>
	</tr>
	</form>
	<tr>
	  <td colspan="5" align="right" class="TableRow2"><%
	  ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName
	  %></td>
	</tr>
</table>
<%
End Sub

Private Sub FlashEdit(isEdit)
	Dim EditTitle,TitleColor,downid
	If isEdit Then
		SQL = "select * from ECCMS_FlashList where flashid=" & Request("flashid")
		Set Rs = enchiasp.Execute(SQL)
		ClassID = Rs("ClassID")
		EditTitle = "编辑" & sModuleName
		downid = Rs("downid")
	Else
		EditTitle = "添加" & sModuleName
		ClassID = Request("ClassID")
		downid = 0
	End If
%>
<script src='include/FlashJuge.Js' type=text/javascript></script>
<script language= JavaScript>
function SelectPhoto(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=<%=ChannelID%>&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.miniature.value=ss[0];
  }
}
function SelectFile(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=<%=ChannelID%>&UploadDir=UploadFile', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.showurl.value=ss[0];
    document.myform.filesize.value=ss[1];
  }
}
</script>
<div onkeydown=CtrlEnter()>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
  <tr>
    <th colspan="4"><%=EditTitle%></th>
  </tr>
 	<form method=Post name="myform" action="admin_flash.asp" onSubmit='return CheckForm(this);'>
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""flashid"" value="""& Request("flashid") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
%>
  <tr>
    <td width="12%" align="right" nowrap class="TableRow2"><strong><%=sModuleName%>分类：</strong></td>
    <td width="38%" class="TableRow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	sClassSelect = Replace(sClassSelect, "{ClassID=" & ClassID & "}", "selected")
	Response.Write sClassSelect
	Response.Write "</select>"
%>
    </td>
    <td width="12%" align="right" class="TableRow2"><strong></strong></td>
    <td width="38%" class="TableRow1"></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>名称：</strong></td>
    <td class="TableRow1"><input name="title" type="text" id="title" size="35" value="<%If isEdit Then Response.Write Rs("title")%>"> 
      <span class="style1">* </span></td>
    <td align="right" class="TableRow2"><strong>名称字体：</strong></td>
    <td class="TableRow1">
            <select size="1" name="ColorMode">
		<option value="0">请选择颜色</option>
<%
	TitleColor = "," & enchiasp.InitTitleColor
	TitleColor = Split(TitleColor, ",")
	For i = 1 To UBound(TitleColor)
		Response.Write ("<option style=""background-color:"& TitleColor(i) &";color: "& TitleColor(i) &""" value='"& i &"'")
		If isEdit Then
			If Rs("ColorMode") = i Then Response.Write (" selected")
		End If
		Response.Write (">"& TitleColor(i) &"</option>")
	Next
%>
		</select>
		<select size="1" name="FontMode">
		<%If isEdit Then%>
		<option value="0"<%If Rs("FontMode") = 0 Then Response.Write (" selected")%>>请选择字体</option>
		<option value="1"<%If Rs("FontMode") = 1 Then Response.Write (" selected")%>>粗体</option>
		<option value="2"<%If Rs("FontMode") = 2 Then Response.Write (" selected")%>>斜体</option>
		<option value="3"<%If Rs("FontMode") = 3 Then Response.Write (" selected")%>>下划线</option>
		<option value="4"<%If Rs("FontMode") = 4 Then Response.Write (" selected")%>>粗体+斜体</option>
		<option value="5"<%If Rs("FontMode") = 5 Then Response.Write (" selected")%>>粗体+下划线</option>
		<option value="6"<%If Rs("FontMode") = 6 Then Response.Write (" selected")%>>斜体+下划线</option>
		<%Else%>
		<option value="0">请选择字体</option>
		<option value="1">粗体</option>
		<option value="2">斜体</option>
		<option value="3">下划线</option>
		<option value="4">粗体+斜体</option>
		<option value="5">粗体+下划线</option>
		<option value="6">斜体+下划线</option>
		<%End If%>
		</select></td>
  </tr>
  <tr>
          <td align="right" class="TableRow2"><b>相关<%=sModuleName%>：</b></td>
          <td colspan="3" class="TableRow1"><input name="Related" type="text" id="Related" size="60" value="<%If isEdit Then Response.Write Rs("Related")%>"> <font color=red>*</font></td>
  </tr>
  <tr>
    <td height="130" align="right" class="TableRow2"><strong><%=sModuleName%>大小：</strong></td>
    <td class="TableRow1">
<%
	Response.Write " <input type=""text"" name=""filesize"" id=""filesize"" size=""14"" onkeyup=if(isNaN(this.value))this.value='' value='"
	If isEdit Then
		Response.Write Trim(Rs("filesize"))
	End If
	Response.Write "'> <input name=""SizeUnit"" type=""radio"" value=""KB"" checked>"
	Response.Write " KB"
	Response.Write " <input type=""radio"" name=""SizeUnit"" value=""MB"">"
	Response.Write " MB <font color=""#FF0000"">！</font>"
%>
    </td>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>星级：</strong></td>
    <td class="TableRow1"><select name="star">
		<%If isEdit Then%>
          	<option value=5<%If Rs("star") = 5 Then Response.Write (" selected")%>>★★★★★</option>
          	<option value=4<%If Rs("star") = 4 Then Response.Write (" selected")%>>★★★★</option>
          	<option value=3<%If Rs("star") = 3 Then Response.Write (" selected")%>>★★★</option>
		<option value=2<%If Rs("star") = 2 Then Response.Write (" selected")%>>★★</option>
		<option value=1<%If Rs("star") = 1 Then Response.Write (" selected")%>>★</option>
		<%Else%>
		<option value=5>★★★★★</option>
          	<option value=4>★★★★</option>
          	<option value=3 selected>★★★</option>
		<option value=2>★★</option>
		<option value=1>★</option>
		<%End If%>
          </select></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>作者：</strong></td>
    <td class="TableRow1"><input name="Author" type="text" id="Author" size="20" value="<%If isEdit Then Response.Write Rs("Author")%>">
	<select name=font2 onChange="Author.value=this.value;">
			<option selected value="">选择作者</option>
			<option value='佚名'>佚名</option>
			<option value='本站原创'>本站原创</option>
			<option value='不详'>不详</option>
			<option value='未知'>未知</option>
			<option value='<%=AdminName%>'><%=AdminName%></option>
		</select></td>
    <td align="right" class="TableRow2"><strong>作品来源：</strong></td>
    <td class="TableRow1"><input name="ComeFrom" type="text" id="ComeFrom" size="25" value="<%If isEdit Then Response.Write Rs("ComeFrom")%>">
    <select name=font1 onChange="ComeFrom.value=this.value;">
			<option selected value="">选择来源</option>
			<option value='本站原创'>本站原创</option>
			<option value='本站整理'>本站整理</option>
			<option value='不详'>不详</option>
			<option value='转载'>转载</option>
			</select></td>
  </tr>
  <tr style="display:none">
    <td align="right" class="TableRow2"><strong>下载等级：</strong></td>
    <td class="TableRow1"><select name="UserGroup">
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & """"
		If isEdit Then
			If Rs("UserGroup") = RsObj("Grades") Then Response.Write " selected"
		Else
			If RsObj("Grades") = 0 Then Response.Write " selected"
		End If
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>
          </select></td>
    <td align="right" class="TableRow2"><strong>下载点数：</strong></td>
    <td class="TableRow1"><input name="PointNum" type="text" id="PointNum" size="10" value="<%If isEdit Then Response.Write Rs("PointNum") Else Response.Write 0 End If%>"></td>
  </tr>
  <tr>
    <td align="right" nowrap class="TableRow2"><strong><%=sModuleName%>缩略图：</strong></td>
    <td colspan="3" class="TableRow1"><input name="miniature" type="text" id="ImageUrl" size="60" value="<%If isEdit Then Response.Write enchiasp.ChkNull(Rs("miniature"))%>">
    <input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto()' class=button></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>图片上传</strong></td>
    <td colspan="3" class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?ChannelID=<%=ChannelID%>></iframe></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>简介：</strong></td>
    <td colspan="3" class="TableRow1"><textarea name="content" style="display:none" id="content"><%If isEdit Then Response.Write Server.HTMLEncode(Rs("Introduce"))%></textarea>
    <iframe ID='HtmlEditor1' src='../editor/editor.asp?ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='560' HEIGHT='350'></iframe></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>描述：</strong></td>
    <td colspan="3" class="TableRow1"><input name="Describe" type="text" id="Describe" size="80" value="<%If isEdit Then Response.Write enchiasp.ChkNull(Rs("Describe"))%>"></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>其它选项：</strong></td>
    <td colspan="3" class="TableRow1"><input name="isTop" type="checkbox" id="isTop" value="1"<%If isEdit Then:If Rs("isTop") <> 0 Then Response.Write (" checked")%>>
            <%=sModuleName%>置顶 
            <input name="isBest" type="checkbox" id="isBest" value="1"<%If isEdit Then:If Rs("isBest") <> 0 Then Response.Write (" checked")%>> 
            <%=sModuleName%>推荐
	    <input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"<%If isEdit Then:If Rs("ForbidEssay") <> 0 Then Response.Write (" checked")%>> 
            禁止发表评论
	    <input name="isAccept" type="checkbox" id="isAccept" value="1" checked> 
            立即发布（<font color=blue>否则审核后才能发布。</font>）
	    <%If isEdit Then%>
	    <input name="Update" type="checkbox" value="1"> 
            同时更新时间<%End If%></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>显示模式：</strong></td>
    <td colspan="3" class="TableRow1">
<%
	Dim ShowModeArray
	
	ShowModeArray = Array("不显示","FLASH","图片","Media","Real","DCR")
	For i = 0 To UBound(ShowModeArray)
		Response.Write "<input type=""radio"" name=""showmode"" value=""" & i & """ "
		If isEdit Then
			If i = Rs("showmode") Then Response.Write " checked"
		Else
			If i = DefaultShowMode  Then Response.Write " checked"
		End If
		Response.Write ">" & Trim(ShowModeArray(i)) & " "
		If i = 6 Then Response.Write "<br>"
	Next
%>
    </td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>页面显示URL：</strong></td>
    <td colspan="3" class="TableRow1"><input name="showurl" type="text" id="filePath" size="60" value="<%If isEdit Then Response.Write enchiasp.ChkNull(Rs("showurl"))%>">
    <input type='button' name='selectfile' value='从已上传文件中选择' onclick='SelectFile()' class=button></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>上传<%=sModuleName%>：</strong></td>
    <td colspan="3" class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upflash.asp?ChannelID=<%=ChannelID%>></iframe></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>下载服务器：</strong></td>
    <td colspan="3" class="TableRow1"><%=SelDownServer(downid)%> <b>说明：</b><font color=blue>下载服务器路径 + 下载文件名称 = 完整下载地址</font></td>
  </tr>
  <tr>
    <td align="right" nowrap class="TableRow2"><strong>下载地址：</strong></td>
    <td colspan="3" class="TableRow1"><input name="DownAddress" type="text" id="DownAddress" size="80" value="<%If isEdit Then Response.Write enchiasp.ChkNull(Rs("DownAddress"))%>"></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2">　</td>
    <td colspan="3" align="center" class="TableRow1">
    <input type="button" name="Submit2" onclick="CheckLength();" value="查看内容长度" class=Button>
    <input type="button" name="Submit3" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>
    <input type="Submit" name="Submit1" value="保存<%=sModuleName%>" class=Button></td>
  </tr></form>
</table></div>
<%
	If isEdit Then
		Rs.Close
		Set Rs = Nothing
	End If
End Sub

Private Function SelDownServer(intdownid)
	Dim rsobj
	If Not IsNumeric(intdownid) Then intdownid = 0
	Response.Write " <select name=""downid"" size=""1"">"
	Response.Write "<option value=""0"""
	If intdownid = 0 Then Response.Write " selected"
	Response.Write ">↓请选择下载服务器↓</option>"
	SQL = "SELECT downid,DownloadName,depth,rootid FROM ECCMS_DownServer WHERE depth=0 And ChannelID="& ChannelID
	Set rsobj = enchiasp.Execute(SQL)
	Do While Not rsobj.EOF
		Response.Write "<option value=""" & rsobj("rootid") & """"
		If intdownid = rsobj("rootid") Then Response.Write " selected"
		Response.Write ">" & rsobj(1) & "</option>"
		rsobj.movenext
	Loop
	rsobj.Close
	Set rsobj = Nothing
	Response.Write "<option value=""0"">不使用下载服务器</option>"
	Response.Write "</select>"
End Function
Private Sub CheckSave()
	If Trim(Request.Form("title")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "名称不能为空！</li>"
	End If
	If Len(Request.Form("title")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "名称不能超过200个字符！</li>"
	End If
	If Trim(Request.Form("ColorMode")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>标题颜色参数错误！</li>"
	End If
	If Trim(Request.Form("FontMode")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>标题字体参数错误！</li>"
	End If
	If Trim(Request.Form("PointNum")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>下载" & sModuleName & "所需的点数不能为空！如果不想设置请输入零。</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "星级不能为空。</li>"
	End If
	If Not IsNumeric(Request.Form("UserGroup")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "等级参数错误！</li>"
	End If
	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该一级分类已经有下属分类，不能添加" & sModuleName & "！</li>"
	End If
	If enchiasp.ChkNumeric(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该分类是外部连接，不能添加" & sModuleName & "！</li>"
	End If	
	TextContent = ""
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next
	If Len(TextContent) < 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "简介不能为空！</li>"
	End If
	If CInt(Request.Form("isTop")) = 1 Then
		FlashTop = 1
	Else
		FlashTop = 0
	End If
	If CInt(Request.Form("isBest")) = 1 Then
		FlashBest = 1
	Else
		FlashBest = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If
	If CInt(Request("isAccept")) = 1 Then
		FlashAccept = 1
	Else
		FlashAccept = 0
	End If
	If trim(Request.Form("filesize")) = "" Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "大小不能为空！</li>"
	End If
End Sub

Private Sub SaveFlash()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_FlashList where (flashid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("SpecialID") = 0
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("ColorMode") = Trim(Request.Form("ColorMode"))
		Rs("FontMode") = Trim(Request.Form("FontMode"))
		Rs("Introduce") = TextContent
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Describe") = enchiasp.ChkFormStr(Request.Form("Describe"))
		If UCase(Request.Form("SizeUnit")) = "MB" Then
			Rs("filesize") = CLng(Request.Form("filesize") * 1024)
		Else
			Rs("filesize") = CLng(Request.Form("filesize"))
		End If
		Rs("Author") = Trim(Request.Form("Author"))
		Rs("ComeFrom") = Trim(Request.Form("ComeFrom"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("miniature") = Trim(Request.Form("miniature"))
		Rs("UserGroup") = enchiasp.ChkNumeric(Request.Form("UserGroup"))
		Rs("PointNum") = enchiasp.ChkNumeric(Request.Form("PointNum"))
		Rs("UserName") = Trim(AdminName)
		Rs("addTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("AllHits") = 0
		Rs("DayHits") = 0
		Rs("WeekHits") = 0
		Rs("MonthHits") = 0
		Rs("HitsTime") = Now()
		Rs("grade") = 0
		Rs("showmode") = enchiasp.ChkNumeric(Request.Form("showmode"))
		Rs("isTop") = FlashTop
		Rs("IsBest") = FlashBest
		Rs("downid") = enchiasp.ChkNumeric(Request.Form("downid"))
		Rs("showurl") = Trim(Request.Form("showurl"))
		Rs("DownAddress") = Trim(Request.Form("DownAddress"))
		Rs("isUpdate") = 1
		Rs("isAccept") = FlashAccept
		Rs("ForbidEssay") = ForbidEssay
		Rs("AlphaIndex") = enchiasp.ReadAlpha(Request.Form("title"))
	Rs.update
	Rs.Close
	Rs.Open "select top 1 flashid from ECCMS_FlashList where ChannelID=" & ChannelID & " order by flashid desc", Conn, 1, 1
	flashid = Rs("flashid")
	Rs.Close:Set Rs = Nothing
	Call RemoveCache
	ClassUpdateCount Request.Form("ClassID"),1
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makeflash.asp?ChannelID=" & ChannelID & "&flashid=" & flashid & "&showid=0"	
		Call ScriptCreation(url,flashid)
		SQL = "SELECT TOP 1 flashid FROM ECCMS_FlashList WHERE ChannelID=" & ChannelID & " And isAccept <> 0 And flashid < " & flashid & " ORDER BY flashid DESC"
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.EOF And Rs.BOF) Then
			url = "admin_makeflash.asp?ChannelID=" & ChannelID & "&flashid=" & Rs("flashid") & "&showid=0"	
			Call ScriptCreation(url,Rs("flashid"))
		End If
		Rs.Close
		Set Rs = Nothing
	End If
	Succeed("<li>恭喜您！添加新的" & sModuleName & "成功。</li><li><a href=?action=view&ChannelID=" & ChannelID & "&flashid=" & flashid & ">点击此处查看该" & sModuleName & "</a></li>")

End Sub

Private Sub ModifyFlash()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_FlashList where flashid=" & Request("flashid")
	Rs.Open SQL,Conn,1,3
		Auditing = Rs("isAccept")
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("SpecialID") = 0
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("ColorMode") = Trim(Request.Form("ColorMode"))
		Rs("FontMode") = Trim(Request.Form("FontMode"))
		Rs("Introduce") = TextContent
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Describe") = enchiasp.ChkFormStr(Request.Form("Describe"))
		If UCase(Request.Form("SizeUnit")) = "MB" Then
			Rs("filesize") = CLng(Request.Form("filesize") * 1024)
		Else
			Rs("filesize") = CLng(Request.Form("filesize"))
		End If
		Rs("Author") = Trim(Request.Form("Author"))
		Rs("ComeFrom") = Trim(Request.Form("ComeFrom"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("miniature") = Trim(Request.Form("miniature"))
		Rs("UserGroup") = Trim(Request.Form("UserGroup"))
		Rs("PointNum") = Trim(Request.Form("PointNum"))
		Rs("UserName") = Trim(AdminName)
		If CInt(Request.Form("Update")) = 1 Then Rs("addTime") = Now()
		Rs("showmode") = enchiasp.ChkNumeric(Request.Form("showmode"))
		Rs("isTop") = FlashTop
		Rs("IsBest") = FlashBest
		Rs("downid") = enchiasp.ChkNumeric(Request.Form("downid"))
		Rs("showurl") = Trim(Request.Form("showurl"))
		Rs("DownAddress") = Trim(Request.Form("DownAddress"))
		Rs("isUpdate") = 1
		Rs("isAccept") = FlashAccept
		Rs("ForbidEssay") = ForbidEssay
		Rs("AlphaIndex") = enchiasp.ReadAlpha(Request.Form("title"))
	Rs.update
	flashid = Rs("flashid")
	If FlashAccept = 1 And Auditing = 0 Then
		AddUserPointNum Rs("username"),1
	End If
	If FlashAccept = 0 And Auditing = 1 Then
		AddUserPointNum Rs("username"),0
	End If
	Rs.Close:Set Rs = Nothing
	Call RemoveCache
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makeflash.asp?ChannelID=" & ChannelID & "&flashid=" & flashid & "&showid=0"	
		Call ScriptCreation(url,flashid)
	End If
	Succeed("<li>恭喜您！修改" & sModuleName & "成功。</li><li><a href=?action=view&ChannelID=" & ChannelID & "&flashid=" & flashid & ">点击此处查看该" & sModuleName & "</a></li>")
End Sub

Private Sub FlashView()
	Call PageTop
	If Request("flashid") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	SQL = "select * from ECCMS_FlashList where ChannelID=" & ChannelID & " And flashid=" & Request("flashid")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！没有找到任何" & sModuleName & "。或者您选择了错误的系统参数！</li>"
		Exit Sub
	Else
%>
	<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
	<tr>
	  <th colspan="2">查看<%=sModuleName%></th>
	</tr>
	<tr>
	  <td align="center" class="TableRow2" colspan="2"><a href=?action=edit&ChannelID=<%=ChannelID%>&flashid=<%=Rs("flashid")%>><font size=4><%=enchiasp.ReadFontMode(Rs("title"),Rs("ColorMode"),Rs("FontMode"))%></font></a></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong><%=sModuleName%>名称：</strong> <%=Rs("title")%></td>
	  <td class="TableRow1"><strong><%=sModuleName%>大小：</strong> <%=Rs("filesize")%> KB</td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong><%=sModuleName%>作者：</strong> <%=Rs("Author")%></td>
	  <td class="TableRow1"><strong>作品来源：</strong> <%=ReadComeFrom(Rs("ComeFrom"))%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>更新时间：</strong> <%=Rs("addTime")%></td>
	  <td class="TableRow1"><strong><%=sModuleName%>星级：</strong> 
<%
Response.Write "<font color=red>"
For i = 1 to Rs("star")
	Response.Write "★"
Next
Response.Write "</font>"
%>
	  </td>
	</tr>
	<tr>
	  <td colspan="2" align="center" class="TableRow1">
<%
	Call PreviewMode(Rs("showurl"),Rs("showmode"))
%>
	  </td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1"><strong><%=sModuleName%>简介：</strong><br>&nbsp;&nbsp;&nbsp;&nbsp;<%=enchiasp.ReadContent(Rs("Introduce"))%></td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1">上一<%=sModuleName%>：<%=FrontFlash(Rs("flashid"))%>
	  <br>下一<%=sModuleName%>：<%=NextFlash(Rs("flashid"))%></td>
	</tr>
	<tr>
	  <td align="center" colspan="2" class="TableRow2"><input type="button" onclick="{if(confirm('您确定要删除吗?')){location.href='?action=del&ChannelID=<%=ChannelID%>&flashid=<%=Rs("flashid")%>';return true;}return false;}" value="删除<%=sModuleName%>" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="返回上一页" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="返回顶部" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&ChannelID=<%=ChannelID%>&flashid=<%=Rs("flashid")%>'" value="编辑<%=sModuleName%>" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub
Public Function ReadComeFrom(ByVal strContent)
	ReadComeFrom = ""
	If IsNull(strContent) Then Exit Function
	If Trim(strContent) = "" Then Exit Function
	strContent = " " & strContent & " "
	Dim re
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""|'])+)"
	strContent = re.Replace(strContent,"<a target=""_blank"" href=$1>$1</a>")
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""])+)$([^\[|']*)"
	strContent = re.Replace(strContent,"<a target=""_blank"" href=$1>$1</a>")
	re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""|'])+)"
	strContent = re.Replace(strContent,"$1<a target=""_blank"" href=$2>$2</a>")
	re.Pattern = "([\s])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
	strContent = re.Replace(strContent,"<a target=""_blank"" href=""http://$2"">$2</a>")
	Set re = Nothing
	ReadComeFrom = Trim(strContent)
End Function

Private Sub PreviewMode(url,modeid)
	If Len(url) < 3 Then Exit Sub
	If Left(url,1) <> "/" And InStr(url,"://") = 0 Then
		url = "../" & enchiasp.ChannelDir & url
	End If
	Select Case CInt(modeid)
	Case 1
		Response.Write "<object codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,0,0"" height=""400"" width=""550"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"">"& vbCrLf
		Response.Write "	<param name=""movie"" value=""" & url & """>"& vbCrLf
		Response.Write "	<param name=""quality"" value=""high"">"& vbCrLf
		Response.Write "	<param name=""SCALE"" value=""exactfit"">"& vbCrLf
		Response.Write "	<embed src=""" & url & """ quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""550"" height=""400"">"& vbCrLf
		Response.Write "	</embed>"& vbCrLf
		Response.Write "</object>"& vbCrLf
	Case 2
		Response.Write "<img src=""" & url & """ border=""0"" onload=""return imgzoom(this,550)"">"
	Case 3
		Response.Write "<object classid=""CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95"" class=""OBJECT"" id=""MediaPlayer"" width=""220"" height=""220"">"& vbCrLf
		Response.Write "	<param name= value=""-1"">"& vbCrLf
		Response.Write "	<param name=""CaptioningID"" value>"& vbCrLf
		Response.Write "	<param name=""ClickToPlay"" value=""-1"">"& vbCrLf
		Response.Write "	<param name=""Filename"" value=""" & url & """>"& vbCrLf
		Response.Write "	<embed src=""" & url & """  width= 220 height=""220"" type=""application/x-oleobject"" codebase=""http://activex.microFlash.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=0,1,1,1"" flename=""mp""></embed>"& vbCrLf
		Response.Write "</object>"& vbCrLf
	Case 4
		Response.Write "<object classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" height=""288"" id=""video1"" width=""305"" VIEWASTEXT>"& vbCrLf
		Response.Write "	<param name=""_ExtentX"" value=""5503"">"& vbCrLf
		Response.Write "	<param name=""_ExtentY"" value=""1588"">"& vbCrLf
		Response.Write "	<param name=""AUTOSTART"" value=""-1"">"& vbCrLf
		Response.Write "	<param name=""SHUFFLE"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""PREFETCH"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""NOLABELS"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""SRC"" value=""" & url & """>"& vbCrLf
		Response.Write "	<param name=""CONTROLS"" value=""Imagewindow,StatusBar,ControlPanel"">"& vbCrLf
		Response.Write "	<param name=""CONSOLE"" value=""RAPLAYER"">"& vbCrLf
		Response.Write "	<param name=""LOOP"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""NUMLOOP"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""CENTER"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""MAINTAINASPECT"" value=""0"">"& vbCrLf
		Response.Write "	<param name=""BACKGROUNDCOLOR"" value=""#000000"">"& vbCrLf
		Response.Write "</object>"& vbCrLf
	Case 5
		Response.Write "<object classid=""clsid:166B1BCA-3F9C-11CF-8075-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=8,5,1,0"" width=""100%"" height=""100%"">"& vbCrLf
		Response.Write "	<param name=""src"" value=""" & url & """>"& vbCrLf
		Response.Write "	<param name=""swRemote"" value=""swSaveEnabled='false' swVolume='false' swRestart='false' swPausePlay='false' swFastForward='false' swContextMenu='false' "">"& vbCrLf
		Response.Write "	<param name=""swStretchStyle"" value=""fill"">"& vbCrLf
		Response.Write "	<PARAM name=""bgColor"" value=""#000000"">"& vbCrLf
		Response.Write "	<PARAM name=logo value=""false"">"& vbCrLf
		Response.Write "	<embed src=""" & url & """ bgColor=""#000000"" logo=""FALSE"" width=""550"" height=""400"" swRemote=""swSaveEnabled='false' swVolume='false' swRestart='false' swPausePlay='false' swFastForward='false' swContextMenu='false' "" swStretchStyle=""fill"" type=""application/x-director"" pluginspage=""http://www.macromedia.com/shockwave/download/""></embed>"& vbCrLf
		Response.Write "</object>"& vbCrLf
	

	End Select
End Sub

Private Function FrontFlash(flashid)
	Dim Rss, SQL
	SQL = "select Top 1 flashid,classid,title from ECCMS_FlashList where ChannelID=" & ChannelID & " And isAccept <> 0 And flashid < " & flashid & " order by flashid desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		FrontFlash = "已经没有了"
	Else
		FrontFlash = "<a href=admin_flash.asp?action=view&ChannelID=" & ChannelID & "&flashid=" & Rss("flashid") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function NextFlash(flashid)
	Dim Rss, SQL
	SQL = "select Top 1 flashid,classid,title from ECCMS_FlashList where ChannelID=" & ChannelID & " And isAccept <> 0 And flashid > " & flashid & " order by flashid asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		NextFlash = "已经没有了"
	Else
		NextFlash = "<a href=admin_flash.asp?action=view&ChannelID=" & ChannelID & "&flashid=" & Rss("flashid") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Sub BatCreateHtml()
	Dim Allflashid,url
	Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
	Response.Write "<ol>"
	Allflashid = Split(selflashid, ",")
	For i = 0 To UBound(Allflashid)
		flashid = CLng(Allflashid(i))
		url = "admin_makeflash.asp?ChannelID=" & ChannelID & "&flashid=" & flashid & "&showid=1"	
		Call ScriptCreation(url,flashid)
	Next
	Response.Write "</ol>"
	OutHintScript("开始生成HTML,共有" & i & "个HTML页面需要生成！")
End Sub
Private Function ClassUpdateCount(sortid,stype)
	Dim rscount,Parentstr
	On Error Resume Next
	Set rscount = enchiasp.Execute("SELECT ClassID,Parentstr FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID=" & CLng(sortid))
	If Not (rscount.BOF And rscount.EOF) Then
		Parentstr = rscount("Parentstr") &","& rscount("ClassID")
		If CInt(stype) = 1 Then
			enchiasp.Execute ("UPDATE [ECCMS_Classify] SET ShowCount=ShowCount+1,isUpdate=1 WHERE ChannelID = "& ChannelID &" And ClassID in (" & Parentstr & ")")
		Else
			enchiasp.Execute ("UPDATE [ECCMS_Classify] SET ShowCount=ShowCount-1,isUpdate=1 WHERE ChannelID = "& ChannelID &" And ClassID in (" & Parentstr & ")")
		End If
	End If
	Set rscount = Nothing
End Function
Private Sub FlashDel()
	If Request("flashid") = "" Then
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	On Error Resume Next
	Set Rs = enchiasp.Execute("SELECT flashid,classid,username,HtmlFileDate FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And flashid=" & Request("flashid"))
	If Not(Rs.BOF And Rs.EOF) Then
		ClassUpdateCount Rs("classid"),0
		AddUserPointNum Rs("username"),0
		DeleteHtmlFile Rs("classid"),Rs("flashid"),Rs("HtmlFileDate")
	End If
	Rs.Close:Set Rs = Nothing
	Conn.Execute("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And flashid=" & Request("flashid"))
	Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID=" & Request("flashid"))
	Call RemoveCache
	Response.redirect ("admin_flash.asp?ChannelID=" & ChannelID)
End Sub

Private Sub batdel()
	Set Rs = enchiasp.Execute("SELECT flashid,classid,username,HtmlFileDate FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And flashid in (" & selflashid & ")")
	Do While Not Rs.EOF
		ClassUpdateCount Rs("classid"),0
		AddUserPointNum Rs("username"),0
		DeleteHtmlFile Rs("classid"),Rs("flashid"),Rs("HtmlFileDate")
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	Conn.Execute ("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And flashid in (" & selflashid & ")")
	Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID in (" & selflashid & ")")
	Call RemoveCache
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub batmove()
	If Not IsNumeric(Request.Form("ClassID")) Then
		OutHintScript ("该一级分类已经有下属分类，请移动到其下属分类！")
		Exit Sub
	End If
	If Trim(Request.Form("classid")) <> "" Then
		enchiasp.Execute ("update ECCMS_FlashList set ClassID = " & Request.Form("ClassID") & ",isUpdate=1 where flashid in (" & selflashid & ")")
		OutHintScript ("批量移动操作成功")
	Else
		OutHintScript ("不能移动到外部分类！")
	End If
End Sub

Private Sub upindate()
	enchiasp.Execute ("update [ECCMS_FlashList] set addTime = " & NowString & " where flashid in (" & selflashid & ")")
	Call RemoveCache
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub


Private Sub isCommend()
	enchiasp.Execute ("update ECCMS_FlashList set isBest=1 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noCommend()
	enchiasp.Execute ("update ECCMS_FlashList set isBest=0 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub isTop()
	enchiasp.Execute ("update ECCMS_FlashList set isTop=1 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noTop()
	enchiasp.Execute ("update ECCMS_FlashList set isTop=0 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

'----批量审核
Private Sub BatAccept()
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_FlashList WHERE isAccept=0 And flashid in (" & selflashid & ")")
	Do While Not Rs.EOF
		AddUserPointNum Rs("username"),1
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("update ECCMS_FlashList set isAccept=1 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Sub NotAccept()
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_FlashList WHERE isAccept>0 And flashid in (" & selflashid & ")")
	Do While Not Rs.EOF
		AddUserPointNum Rs("username"),0
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("update ECCMS_FlashList set isAccept=0 where flashid in (" & selflashid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Function AddUserPointNum(username,stype)
	On Error Resume Next
	Dim rsuser,GroupSetting,userpoint
	Set rsuser = enchiasp.Execute("SELECT userid,UserGrade,userpoint FROM ECCMS_User WHERE username='"& username &"'")
	If Not(rsuser.BOF And rsuser.EOF) Then
		GroupSetting = Split(enchiasp.UserGroupSetting(rsuser("UserGrade")), "|||")(13)
		If CInt(stype) = 1 Then
			userpoint = CLng(rsuser("userpoint") + GroupSetting)
			enchiasp.Execute ("UPDATE ECCMS_User SET userpoint="& userpoint &",experience=experience+2,charm=charm+1 WHERE userid="& rsuser("userid"))
		Else
			userpoint = CLng(rsuser("userpoint") - GroupSetting)
			enchiasp.Execute ("UPDATE ECCMS_User SET userpoint="& userpoint &",experience=experience-2,charm=charm-1 WHERE userid="& rsuser("userid"))
		End If
	End If
	Set rsuser = Nothing
End Function

'--批量操作开始
Private Sub BatchSetting()
	If Not ChkAdmin("AdminFlash" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Dim Channel_Setting
	Channel_Setting = Split(enchiasp.Channel_Setting, "|||")
	Response.Write "<script src=""include/FlashJuge.Js"" type=""text/javascript""></script>" & vbNewLine
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=4>" & sModuleName & "批量设置</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=saveset>"
	Response.Write "	<input type=hidden name=ChannelID value='"& ChannelID &"'>"
	Response.Write "	<tr>"
	Response.Write "		<td width=""20%"" rowspan=""18"" class=tablerow2 valign=""top"" id=choose2 style=""display:none""><b>请选择" & sModuleName & "分类</b><br>"
	Response.Write "<select name=""ClassID"" size='2' multiple style='height:420px;width:180px;'>"
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
	Response.Write "<option value=""-1"">指定所有分类</option>"
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "		<td class=tablerow1 colspan=2 align=right><b>设置选择：</b></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "		<input type=radio name=choose value='0' checked onClick=""choose1.style.display='';choose2.style.display='none';""> 按" & sModuleName & "ID&nbsp;&nbsp;"
	Response.Write "		<input type=radio name=choose value='1' onClick=""choose2.style.display='';choose1.style.display='none';""> 按" & sModuleName & "分类</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=choose1>"
	Response.Write "		<td class=tablerow1 colspan=2 align=right><b>" & sModuleName & "ID：</b>多个ID请用“,”分开</td>"
	Response.Write "		<td class=tablerow1><input type=""text"" name=""flashid"" size=70 value='"& Request("selflashid") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 width=""15%"" align=right><b>相关" & sModuleName & "：</b></td>"
	Response.Write "		<td class=tablerow1 width=""5%"" align=center><input type=checkbox name=selRelated value='1'></td>"
	Response.Write "		<td class=tablerow1 width=""60%""><input name=Related type=text size=60></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "来源：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selComeFrom value='1'></td>"
	Response.Write "		<td class=tablerow1 ><input name=ComeFrom type=text size=35>"
	Response.Write "		<select name=font1 onChange=""ComeFrom.value=this.value;"">"
	Response.Write "		<option selected value=''>选择来源</option>"
	Response.Write "		<option value='本站原创'>本站原创</option>"
	Response.Write "		<option value='本站整理'>本站整理</option>"
	Response.Write "		<option value='不详'>不详</option>"
	Response.Write "		<option value='转载'>转载</option>"
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "作者：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selAuthor value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=Author type=text size=20>"
	Response.Write "		<select name=font2 onChange=""Author.value=this.value;"">"
	Response.Write "		<option selected value=''>选择作者</option>"
	Response.Write "		<option value='佚名'>佚名</option>"
	Response.Write "		<option value='本站'>本站</option>"
	Response.Write "		<option value='不详'>不详</option>"
	Response.Write "		<option value='未知'>未知</option>"
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>所需点数：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selPointNum value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=PointNum type=text size=10 value=0></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>点击数：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selAllHits value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=AllHits type=text size=10 value=0></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>阅览等级：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selUserGroup value='1'></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "		<select name=UserGroup size='1'>"
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup order by Groupid")
	Do While Not RsObj.EOF
		Response.Write "	<option value=""" & RsObj("Grades") & """"
		If RsObj("Grades") = 0 Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "星级：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selstar value='1'></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "	<select name=star>"
	Response.Write "		<option value=5>★★★★★</option>"
	Response.Write "		<option value=4>★★★★</option>"
	Response.Write "		<option value=3 selected>★★★</option>"
	Response.Write "		<option value=2>★★</option>"
	Response.Write "		<option value=1>★</option>"
	Response.Write "	</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "置顶：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selTop value='1'></td>"
	Response.Write "		<td class=tablerow1><input type=radio name=istop value='0' checked> 否&nbsp;&nbsp;<input type=radio name=istop value='1'> 是</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "推荐：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selBest value='1'></td>"
	Response.Write "		<td class=tablerow1><input type=radio name=isbest value='0' checked> 否&nbsp;&nbsp;<input type=radio name=isbest value='1'> 是</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>禁止评论：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selForbidEssay value='1'></td>"
	Response.Write "		<td class=tablerow1><input type=radio name=ForbidEssay value='0' checked> 否&nbsp;&nbsp;<input type=radio name=ForbidEssay value='1'> 是</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=3><b>说明：</b>若要批量修改某个属性的值，请先选中其左侧的复选框，然后再设定属性值。"
	Response.Write " <a href=?action=reset&ChannelID="& ChannelID & " onclick=""return confirm('您确定要重置所有时间吗?')"">重置时间</a></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow1 colspan=3><input type=""button"" onclick=""javascript:history.go(-1)"" value=""返回上一页"" name=""B1"" class=Button>&nbsp;&nbsp;"
	Response.Write "		<input type=submit name=submit2 value=""确定设置"" class=Button onclick=""return confirm('您确定执行批量设置吗?')"">"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub
Private Sub BatchMove()
	If Not ChkAdmin("AdminFlash" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=3>" & sModuleName & "批量移动</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=savemove>"
	Response.Write "	<input type=hidden name=ChannelID value='"& ChannelID &"'>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=3><input type=radio name=Appointed value='0' checked>"
	Response.Write " <b>指定" & sModuleName & "ID：</b> <input type=""text"" name=""flashid"" size=80 value='"& Request("selflashid") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 width=""35%""><input type=radio name=Appointed value='1'> <b>指定" & sModuleName & "分类：</b></td>"
	Response.Write "		<td class=tablerow1 width=""10%""></td>"
	Response.Write "		<td class=tablerow1 width=""55%""><b>" & sModuleName & "目标分类：</b><font color=red>（不能指定外部分类）</font></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name='ClassID' size='2' multiple style='height:350px;width:260px;'>"
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
	Response.Write "		<td class=tablerow1 align=center noWrap>移动到→</td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name='tClassID' size='2' style='height:350px;width:260px;'>"
	Response.Write strSelectClass
	Response.Write "</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow1 colspan=3><input type=""button"" onclick=""javascript:history.go(-1)"" value=""返回上一页"" name=""B1"" class=Button>&nbsp;&nbsp;"
	Response.Write "		<input type=submit name=submit2 value=""批量移动"" class=Button onclick=""return confirm('您确定执行批量移动吗?')"">"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub
Private Sub SaveMove()
	If Trim(Request.Form("tClassID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择目标分类。</li>"
		Exit Sub
	End If
	If Trim(Request.Form("tClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>不能移动到外部分类。</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT child FROM ECCMS_Classify WHERE TurnLink=0 And ChannelID = "& ChannelID &" And ClassID="& CLng(Request.Form("tClassID")))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的系统参数！不能移动到外部分类。</li>"
		Set Rs = Nothing
		Exit Sub
	Else
		If Rs("child") > 0 Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>此分类还有子分类，请选择子分类再移动。</li>"
			Set Rs = Nothing
			Exit Sub
		End If
	End If
	Set Rs = Nothing
	If CInt(Request.Form("Appointed")) = 0 Then
		If Trim(Request.Form("flashid")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_FlashList SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And flashid in ("& Request("flashid") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_FlashList SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>恭喜您！批量移动完成。</li>")
End Sub

Private Sub BatcDelete()
	If Not ChkAdmin("AdminFlash" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th>" & sModuleName & "批量删除</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=alldel>"
	Response.Write "	<input type=hidden name=ChannelID value='"& ChannelID &"'>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=3><input type=radio name=Appointed value='0' checked onClick=""choose1.style.display='';choose2.style.display='none';"">"
	Response.Write " <b>指定" & sModuleName & "ID：</b> "
	Response.Write "<input type=radio name=Appointed value='1' onClick=""choose2.style.display='';choose1.style.display='none';""> <b>指定" & sModuleName & "分类：</b>"
	Response.Write "<input type=radio name=Appointed value='2'> <b>删除所有" & sModuleName & "</b>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=choose1>"
	Response.Write "		<td class=tablerow1><b>分类ID：</b><input type=""text"" name=""flashid"" size=80 value='"& Request("selflashid") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=choose2 style=""display:none"">"
	Response.Write "		<td class=tablerow1>"
	Response.Write "<select name='ClassID' size='2' multiple style='height:350px;width:260px;'>"
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
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow1><input type=""button"" onclick=""javascript:history.go(-1)"" value=""返回上一页"" name=""B1"" class=Button>&nbsp;&nbsp;"
	Response.Write "		<input type=submit name=submit2 value=""批量删除"" class=Button onclick=""return confirm('您确定执行批量删除操作吗?')"">"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub
Private Sub AllDelFlash()
	On Error Resume Next
	If CInt(Request.Form("Appointed")) = 1 Then
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		End If
		enchiasp.Execute ("DELETE ECCMS_Comment FROM ECCMS_FlashList A INNER JOIN ECCMS_Comment C ON C.PostID=A.flashid WHERE A.ChannelID = "& ChannelID &" And A.ClassID IN (" & Request("ClassID") & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And ClassID IN (" & Request("ClassID") & ")")
	ElseIf CInt(Request.Form("Appointed")) = 2 Then
		enchiasp.Execute ("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID)
		enchiasp.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID)
	Else
		If Trim(Request.Form("flashid")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		End If
		enchiasp.Execute ("DELETE FROM ECCMS_FlashList WHERE ChannelID = "& ChannelID &" And flashid IN (" & Request("flashid") & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID IN (" & Request("flashid") & ")")
		
	End If
	Call RemoveCache
	Succeed("<li>批量删除成功！</li>")
End Sub

Private Sub SaveSetting()
	If Founderr = True Then Exit Sub
	Dim strTempValue
	strTempValue = ""
	If Trim(Request.Form("selRelated")) <> "" Then strTempValue = strTempValue & "Related='"& enchiasp.ChkFormStr(Request.Form("Related")) &"',"
	If Trim(Request.Form("selComeFrom")) <> "" Then strTempValue = strTempValue & "ComeFrom='"& enchiasp.ChkFormStr(Request.Form("ComeFrom")) &"',"
	If Trim(Request.Form("selAuthor")) <> "" Then strTempValue = strTempValue & "Author='"& enchiasp.ChkFormStr(Request.Form("Author")) &"',"
	If Trim(Request.Form("selPointNum")) <> "" Then strTempValue = strTempValue & "PointNum="& CLng(Request.Form("PointNum")) &","
	If Trim(Request.Form("selAllHits")) <> "" Then strTempValue = strTempValue & "AllHits="& CLng(Request.Form("AllHits")) &","
	If Trim(Request.Form("selUserGroup")) <> "" Then strTempValue = strTempValue & "UserGroup="& CInt(Request.Form("UserGroup")) &","
	If Trim(Request.Form("selstar")) <> "" Then strTempValue = strTempValue & "star="& CInt(Request.Form("star")) &","
	If Trim(Request.Form("selTop")) <> "" Then strTempValue = strTempValue & "istop="& CInt(Request.Form("istop")) &","
	If Trim(Request.Form("selBest")) <> "" Then strTempValue = strTempValue & "isbest="& CInt(Request.Form("isbest")) &","
	If Trim(Request.Form("selForbidEssay")) <> "" Then strTempValue = strTempValue & "ForbidEssay="& CInt(Request.Form("ForbidEssay")) &","
	If Trim(strTempValue) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择要设置的参数。</li>"
		Exit Sub
	Else
		strTempValue = Replace(Left(strTempValue,Len(strTempValue)-1), " ", "")
	End If
	If CInt(Request.Form("choose")) = 0 Then
		If Trim(Request.Form("flashid")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_FlashList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And flashid in ("& Request("flashid") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		ElseIf Trim(Request.Form("ClassID")) = "-1" Then
			SQL = "UPDATE ECCMS_FlashList SET "& strTempValue &" WHERE ChannelID = "& ChannelID &" And isAccept>0"
		Else
			SQL = "UPDATE ECCMS_FlashList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>恭喜您！批量设置完成。</li>")
End Sub

Private Sub RemoveCache()
	enchiasp.DelCahe "RenewStatistics"
	enchiasp.DelCahe "TotalStatistics"
End Sub


%>