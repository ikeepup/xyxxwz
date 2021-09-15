<!--#include file="setup.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
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
Dim Action
Dim i,ii,isEdit,RsObj
Dim keyword,FindWord,strClass
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Dim s_ClassName,ClassID,ChildStr,FoundSQL,isAccept,selSoftID
Dim TextContent,SoftTop,SoftBest,SoftID,ForbidEssay,showreg,SoftAccept
Dim DownAddress
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If Trim(Request("isAccept")) <> "" Then
	isAccept = 0
Else
	isAccept = 1
End If
If CInt(ChannelID) = 0 Then ChannelID = 2
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "save"
	Call SaveSoft
Case "modify"
	Call ModifySoft
Case "add"
	isEdit = False
	Call SoftEdit(isEdit)
Case "edit"
	isEdit = True
	Call SoftEdit(isEdit)
Case "del"
	Call SoftDel
Case "batdel"
	Call PageTop
	Call BatcDelete
Case "alldel"
	Call AllDelSoft
Case "view"
	Call SoftView
Case "renew"
	Call SoftRenew
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
Case "reset"
	Call ResetDateTime
Case "sdel"
	Call DelDownAddress
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
	Response.Write "	<tr><form method=Post name=myform action='admin_Soft.asp' onSubmit='return JugeQuery(this);'>"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "<td class=TableRow1>搜索："
	Response.Write "	  <input name=keyword type=text size=20>"
	Response.Write "	  条件："
	Response.Write "	  <select name=queryopt>"
	Response.Write "		<option value='1' selected>" & sModuleName & "名称</option>"
	Response.Write "		<option value='2'>录 入 者</option>"
	Response.Write "		<option value='3'>不限条件</option>"
	Response.Write "	  </select> <input type=submit name=Submit value='开始搜索' class=Button></td>"
	Response.Write "	  <td class=TableRow1>软件导航："
	Dim srtClassMenu
	Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Response.Write "<option value='admin_Soft.asp?ChannelID=" & ChannelID & "'>≡全部" & sModuleName & "列表≡</option>" & vbCrLf
	srtClassMenu = enchiasp.ClassJumpMenu(ChannelID)
	srtClassMenu = Replace(srtClassMenu, "{ClassID=" & Request("sortid") & "}", "selected")
	Response.Write srtClassMenu
	Response.Write "</select>"
	Response.Write "	  </td></form>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td colspan=2 class=TableRow2><b>操作选项：</b> <a href='admin_Soft.asp?ChannelID=" & ChannelID & "'>管理首页</a> | "
	Response.Write "	  <a href='admin_Soft.asp?ChannelID=" & ChannelID & "&action=add'>添加软件</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "&action=add'>添加软件分类</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "'>软件分类管理</a> | "
	Response.Write "	  <a href='?ChannelID=" & ChannelID & "&specialID=0'>" & sModuleName & "专题列表</a> | "
	Response.Write "	  <a href='?ChannelID=" & ChannelID & "&isAccept=0'>待审" & sModuleName & "</a> | "
	Response.Write "	  <a href='Admin_CreateSoft.Asp?ChannelID=" & ChannelID & "'>生成HTML</a></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub

Private Sub showmain()
	If Not ChkAdmin("AdminSoft" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Dim strListName
	If Not IsEmpty(Request("selSoftID")) Then
		selSoftID = Request("selSoftID")
		Select Case enchiasp.CheckStr(Request("act"))
		Case "批量删除":Call batdel
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
	Dim specialID,sortid,Cmd,child
	Response.Write chr(0)
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>选择</th>"
	Response.Write "	  <th width='68%'>" & sModuleName & "名称</th>"
	Response.Write "	  <th width='9%' nowrap>管理操作</th>"
	Response.Write "	  <th width='5%' nowrap>审核</th>"
	Response.Write "	  <th width='9%' nowrap>录 入 者</th>"
	Response.Write "	  <th width='9%' nowrap>整理日期</th>"
	Response.Write "	</tr>"
	strListName = "&channelid="& ChannelID &"&sortid="& Request("sortid") &"&specialID="& Request("specialID") &"&isAccept="& Request("isAccept") &"&keyword=" & Request("keyword") 
	If Trim(Request("sortid")) <> "" Then
		SQL = "select ClassID,ChannelID,ClassName,child,ChildStr from [ECCMS_Classify] where ChannelID = " & ChannelID & " And ClassID=" & Request("sortid")
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			Response.Write "Sorry！没有找到任何软件分类。或者您选择了错误的系统参数!"
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
		CurrentPage = CLng(Request("page"))
	Else
		CurrentPage = 1
	End If
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("queryopt")) = 1 Then
			findword = "A.SoftName like '%" & keyword & "%'"
		ElseIf CInt(Request("queryopt")) = 2 Then
			findword = "A.username like '%" & keyword & "%'"
		Else
			findword = "A.SoftName like '%" & keyword & "%' or A.username like '%" & keyword & "%'"
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
	On Error Resume Next
	TotalNumber = enchiasp.Execute("Select Count(SoftID) from ECCMS_SoftList A where A.ChannelID = " & ChannelID & " And "& FoundSQL &"")(0)
	TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select A.SoftID,A.ChannelID,A.ClassID,A.SpecialID,A.SoftName,A.SoftVer,A.ColorMode,A.FontMode,A.username,A.SoftTime,A.isTop,A.isBest,A.isAccept,C.ClassName from [ECCMS_SoftList] A inner join [ECCMS_Classify] C on A.ClassID=C.ClassID where A.ChannelID = " & ChannelID & " And "& FoundSQL &" order by A.isTop desc, A.SoftTime desc ,A.SoftID desc"
	Rs.Open SQL, Conn, 1, 1
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<form name=selform method=post action="""">"
		Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
		Response.Write "<input type=hidden name=action value=''>"
		Response.Write "<tr><td align=center colspan=6 class=TableRow2>还没有找到任何" & sModuleName & "！</td></tr>"
	Else
		Rs.MoveFirst
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		If Rs.Eof Then Exit Sub
		SQL=Rs.GetRows(maxperpage)

		Response.Write "	<tr>"
		Response.Write "	  <td colspan=6 class=TableRow2>"
		ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<form name=selform method=post action="""">"
		Response.Write "	<input type=hidden name=ChannelID value='" & ChannelID & "'>"
		Response.Write "	<input type=hidden name=action value=''>"
		For i=0 To Ubound(SQL,2)
			If Not Response.IsClientConnected Then Response.End
			If (i mod 2) = 0 Then
				strClass = "class=TableRow1"
			Else
				strClass = "class=TableRow2"
			End If
			Admin_Soft_list SQL(0,i),SQL(1,i),SQL(2,i),SQL(4,i),SQL(5,i),SQL(6,i),SQL(7,i),SQL(8,i),SQL(9,i),SQL(10,i),SQL(11,i),SQL(12,i),SQL(13,i),strClass
		Next
		SQL=Null
	End If
	Rs.Close:Set Rs = Nothing
	Set Cmd = Nothing
%>
	<tr>
	  <td colspan="6" class="TableRow1">
	  <input class=Button type="button" name="chkall" value="全选" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="反选" onClick="ContraSel(this.form)">
	  管理选项：
	  <select name="act">
		<option value="0">请选择操作选项</option>
		<option value="批量删除">批量删除</option>
		<option value="批量置顶">批量置顶</option>
		<option value="取消置顶">取消置顶</option>
		<option value="批量推荐">批量推荐</option>
		<option value="取消推荐">取消推荐</option>
		<option value="批量审核">批量审核</option>
		<option value="取消审核">取消审核</option>
		<option value="生成HTML">生成HTML</option>
	  </select>
	  <input class=Button type="submit" name="Submit2" value="执行操作" onclick="return confirm('您确定执行该操作吗?');">
	  <input class=Button type="submit" name="Submit3" value="批量设置" onclick="document.selform.action.value='setting';">
	  <input class=Button type="submit" name="Submit4" value="批量移动" onclick="document.selform.action.value='move';">
	  <input class=Button type="submit" name="Submit4" value="批量删除" onclick="document.selform.action.value='batdel';"></td>
	</tr>
	</form>
	<tr>
	  <td colspan="6" align="right" class="TableRow2"><%ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName %></td>
	</tr>
</table>
<%

End Sub

Function Admin_Soft_list(SoftID,ChannelID,ClassID,SoftName,SoftVer,ColorMode,FontMode,username,SoftTime,isTop,isBest,isAccept,ClassName,strClass)
	Response.Write "	<tr>"
	Response.Write "	  <td align=center " & strClass & "><input type=checkbox name=selSoftID value=" & SoftID & "></td>"
	Response.Write "	  <td " & strClass & ">"
	
	If isTop <> 0 Then
		Response.Write "<img src=""images/istop.gif"" width=15 height=17 border=0 alt=置顶软件>"
	End If

	Response.Write "[<a href=?ChannelID=" & ChannelID & "&sortid="
	Response.Write ClassID
	Response.Write ">"
	Response.Write ClassName
	Response.Write "</a>] "
	Response.Write "<a href=?action=view&ChannelID=" & ChannelID & "&SoftID="
	Response.Write SoftID
	Response.Write ">"
	Response.Write enchiasp.ReadFontMode(SoftName &" "& SoftVer,ColorMode,FontMode)
	Response.Write "</a>" 
	
	If isBest <> 0 Then
		Response.Write "&nbsp;&nbsp;<font color=blue>荐</font>"
	End If

	Response.Write "	  </td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &"><a href=?action=edit&ChannelID="& ChannelID &"&SoftID="& SoftID &">编辑</a> | <a href=?action=del&ChannelID="& ChannelID &"&SoftID="& SoftID &" onclick=""{if(confirm('软件删除后将不能恢复，您确定要删除该软件吗?')){return true;}return false;}"">删除</a></td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &">"
	
	If isAccept = 1 Then
		Response.Write "<font color=blue>√</font>"
	Else
		Response.Write "<font color=red>×</font>"
	End If

	Response.Write "	  </td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &">"& UserName &"</td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &">"

	If SoftTime >= Date Then
		Response.Write "<font color=red>"
		Response.Write enchiasp.FormatDate(SoftTime, 2)
		Response.Write "</font>"
	Else
		Response.Write enchiasp.FormatDate(SoftTime, 2)
	End If

	Response.Write "	  </td>"
	Response.Write "	</tr>"
End Function

Private Sub SoftEdit(isEdit)
	Dim EditTitle,SoftNameColor,Channel_Setting,downid
	If isEdit Then
		If Not ChkAdmin("AddSoft" & ChannelID) Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
		SQL = "select * from ECCMS_SoftList where SoftID=" & CLng(Request("SoftID"))
		Set Rs = enchiasp.Execute(SQL)
		ClassID = Rs("ClassID")
		softid = Rs("softid")
		EditTitle = "编辑" & sModuleName
	Else
		If Not ChkAdmin("AddSoft" & ChannelID) Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
		EditTitle = "添加" & sModuleName
		ClassID = Request("ClassID")
		softid = 0
	End If
	Channel_Setting = Split(enchiasp.Channel_Setting, "|||")
%>
<script src="include/SoftJuge.Js" type="text/javascript"></script>
<script language= JavaScript>
function SelectPhoto(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=<%=ChannelID%>&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.SoftImage.value=ss[0];
  }
}
function SelectFile(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=<%=ChannelID%>&UploadDir=UploadFile', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DownFilePath.value=ss[0];
    document.myform.SoftSize.value=ss[1];
  }
}
</script>
<div onkeydown=CtrlEnter()>
<table  border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
        <tr>
          <th colspan="4"><%=EditTitle%></th>
        </tr>
		<form method=Post name="myform" action="admin_Soft.asp" onSubmit="return CheckForm(this);">
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""SoftID"" value="""& Request("SoftID") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
%>
        <tr>
          <td width="15%" align="right" nowrap class="TableRow2"><b>所属分类：</b></td>
          <td width="35%" class="TableRow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	sClassSelect = Replace(sClassSelect, "{ClassID=" & ClassID & "}", "selected")
	Response.Write sClassSelect
	Response.Write "</select>"
%>
		  </td>
          <td width="15%" align="right" nowrap class="TableRow2"><b>所属专题：</b></td>
          <td width="35%" class="TableRow1"><select name="SpecialID" id="SpecialID">
            <option value="0">不指定<%=sModuleName%>专题</option>
<%
	Set RsObj = enchiasp.Execute("SELECT SpecialID,SpecialName FROM ECCMS_Special WHERE ChannelID="& ChannelID &" And ChangeLink=0 ORDER BY orders")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("SpecialID") & """"
		If isEdit Then
			If Rs("SpecialID") = RsObj("SpecialID") Then Response.Write " selected"
		End If
		Response.Write ">"
		Response.Write RsObj("SpecialName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>
          </select></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>名称：</b></td>
          <td class="TableRow1"><input name="SoftName" type="text" id="SoftName" size="35" value="<%If isEdit Then Response.Write Rs("SoftName")%>"> 
          <font color=red>*</font></td>
	  <td align="right" class="TableRow2"><b><%=sModuleName%>版本：</b></td>
          <td class="TableRow1"><input name="SoftVer" type="text" id="SoftVer" size="25" value="<%If isEdit Then Response.Write Rs("SoftVer")%>"></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>名称模式：</b></td>
          <td colspan="3" class="TableRow1">颜色：
            <select size="1" name="ColorMode">
		<option value="0">请选择颜色</option>
<%
	SoftNameColor = "," & enchiasp.InitTitleColor
	SoftNameColor = Split(SoftNameColor, ",")
	For i = 1 To UBound(SoftNameColor)
		Response.Write ("<option style=""background-color:"& SoftNameColor(i) &";color: "& SoftNameColor(i) &""" value='"& i &"'")
		If isEdit Then
			If Rs("ColorMode") = i Then Response.Write (" selected")
		End If
		Response.Write (">"& SoftNameColor(i) &"</option>")
	Next
%>
		</select> 字体：
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
          <td align="right" class="TableRow2"><b>软件运行环境：</b></td>
          <td colspan="3" class="TableRow1"><input name="RunSystem" type="text" size="60" value="<%If isEdit Then Response.Write Rs("RunSystem") Else Response.Write Channel_Setting(1) End If%>"><br>
<%
	Dim RunSystem
	RunSystem = Split(Channel_Setting(0), "|")
	For i = 0 To UBound(RunSystem)
		Response.Write "<a href='javascript:ToRunsystem(""" & Trim(RunSystem(i)) & """)'><u>" & Trim(RunSystem(i)) & "</u></a> | "
		If i = 10 Then Response.Write "<br>"
	Next
%>
		    </td>
        </tr>
	<tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>类型：</b></td>
          <td colspan="3" class="TableRow1">
<%
	Dim SoftType
	SoftType = Split(Channel_Setting(2), ",")
	For i = 0 To UBound(SoftType)
		Response.Write "<input type=""radio"" name=""SoftType"" value=""" & Trim(SoftType(i)) & """ "
		If isEdit Then
			If SoftType(i) = Rs("SoftType") Then Response.Write " checked"
		Else
			If i = 0 Then Response.Write " checked"
		End If
		Response.Write ">" & Trim(SoftType(i)) & " "
		If i = 6 Then Response.Write "<br>"
	Next
%>
		    </td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>预览图片：</b></td>
          <td colspan="3" class="TableRow1"><input name="SoftImage" id="ImageUrl" type="text" size="60" value="<%If isEdit Then Response.Write Rs("SoftImage")%>">
	  <input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto()' class=button></td>
        </tr>
	<tr>
          <td align="right" class="TableRow2"><b>上传图片：</b></td>
          <td colspan="3" class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
	<tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>大小：</b></td>
          <td class="TableRow1">
<%
	Response.Write " <input type=""text"" name=""SoftSize"" id=""filesize"" size=""14"" onkeyup=if(isNaN(this.value))this.value='' value='"
	If isEdit Then
		Response.Write Trim(Rs("SoftSize"))
	End If
	Response.Write "'> <input name=""SizeUnit"" type=""radio"" value=""KB"" checked>"
	Response.Write " KB"
	Response.Write " <input type=""radio"" name=""SizeUnit"" value=""MB"">"
	Response.Write " MB <font color=""#FF0000"">！</font>"
%>
          </td>
          <td align="right" class="TableRow2"><b><%=sModuleName%>星级：</b></td>
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
          <td align="right" class="TableRow2"><b><%=sModuleName%>性质：</b></td>
          <td class="TableRow1">
<%
	Response.Write " <select name=""impower"">"
	If isEdit Then
		Response.Write "<option value=""" & Rs("impower") & """>" & Rs("impower") & "</option>"
	End If
	Dim ImpowerStr
	ImpowerStr = Split(Channel_Setting(3), ",")
	For i = 0 To UBound(ImpowerStr)
		Response.Write " <option value=""" & ImpowerStr(i) & """>" & ImpowerStr(i) & "</option>"
	Next
	Response.Write " </select>&nbsp;&nbsp;"
	Response.Write " <select name=""Languages"">"
	Response.Write " "
	If isEdit Then
		Response.Write "<option value=""" & Rs("Languages") & """>" & Rs("Languages") & "</option>"
	End If
	Dim LanguagesStr
	LanguagesStr = Split(Channel_Setting(4), ",")
	For i = 0 To UBound(LanguagesStr)
		Response.Write " <option value=""" & LanguagesStr(i) & """>" & LanguagesStr(i) & "</option>"
	Next
	Response.Write " </select>&nbsp;&nbsp;"
%>
	  </td>
          <td align="right" class="TableRow2"><b>解压密码：</b></td>
          <td class="TableRow1"><%Response.Write "<input type=""text"" name=""Decode"" size=""15"" maxlength=""100"" value='"
	If isEdit Then
		Response.Write Trim(Rs("Decode"))
	End If
	Response.Write "'> <font color=""#808080"">没有请留空</font>"%></td>
	</tr>
	<tr>
          <td align="right" class="TableRow2"><b>联系方式：</b></td>
          <td class="TableRow1">
<%
	Response.Write "<input name=""Contact"" type=""text"" size=""33"" "
	If isEdit Then
		Response.Write "value="""
		Response.Write Rs("Contact")
		Response.Write """"
	Else
		Response.Write "onblur=""if (value ==''){value='"
		Response.Write enchiasp.MasterMail
		Response.Write "'}"" onmouseover=this.focus() onfocus=this.select() onclick=""if(this.value=='"
		Response.Write enchiasp.MasterMail
		Response.Write "')this.value=''"" value="""
		Response.Write enchiasp.MasterMail
		Response.Write """"
	End If
	Response.Write ">"
%>
	  </td>
          <td align="right" class="TableRow2"><b>程序主页：</b></td>
          <td class="TableRow1">
<%
Response.Write "<input name=""Homepage"" type=""text"" size=""30"" "
	If isEdit Then
		Response.Write "value="""
		Response.Write Rs("Homepage")
		Response.Write """"
	Else
		Response.Write "onblur=""if (value ==''){value='"
		Response.Write enchiasp.SiteUrl
		Response.Write "'}"" onmouseover=this.focus() onfocus=this.select() onclick=""if(this.value=='"
		Response.Write enchiasp.SiteUrl
		Response.Write "')this.value=''"" value="""
		Response.Write enchiasp.SiteUrl
		Response.Write """"
	End If
	Response.Write ">"
%>
	  </td>
	</tr>
	<tr>
          <td align="right" class="TableRow2"><b>软件作者：</b></td>
          <td class="TableRow1">
<%
	Response.Write "<input name=""Author"" type=""text"" size=""20"" "
	If isEdit Then
		Response.Write "value="""
		Response.Write Rs("Author")
		Response.Write """"
	End If
	Response.Write ">"
%>
		<select name=font2 onChange="Author.value=this.value;">
			<option selected value="">选择作者</option>
			<option value='佚名'>佚名</option>
			<option value='本站原创'>本站原创</option>
			<option value='不详'>不详</option>
			<option value='未知'>未知</option>
			<option value='<%=AdminName%>'><%=AdminName%></option>
		</select>
	  </td>
          <td align="right" class="TableRow2"><b>注册网址：</b></td>
          <td class="TableRow1">
<%
	Response.Write "<input name=""Regsite"" type=""text"" size=""30"" "
	If isEdit Then
		Response.Write "value="""
		Response.Write Rs("Regsite")
		Response.Write """"
	End If
	Response.Write ">"
%>
	  </td>
	</tr>
        <tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>简介：</b></td>
          <td colspan="3" class="TableRow1"><textarea name="content" style="display:none"><%If isEdit Then Response.Write Server.HTMLEncode(Rs("content"))%></textarea>
		<iframe ID='HtmlEditor1' src='../editor/editor.asp?ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='560' HEIGHT='350'></iframe></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>下载等级：</b></td>
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
          <td align="right" class="TableRow2"><b>所需点数：</b></td>
          <td class="TableRow1"><input name="PointNum" type="text" size="10" value="<%If isEdit Then Response.Write Rs("PointNum"):Else Response.Write 0:End If%>"> 
            对匿名用户和管理员无效 </td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>初始点击数：</b></td>
          <td class="TableRow1"><input name="AllHits" type="text" id="AllHits" size="15" value="<%If isEdit Then Response.Write Rs("AllHits"):Else Response.Write 0%>"> 
          <font color=red>*</font></td>
          <td align="right" class="TableRow2"><b><%=sModuleName%>价格：</b></td>
          <td class="TableRow1"><input name="SoftPrice" type="text" size="10" value="<%If isEdit Then Response.Write Rs("SoftPrice"):Else Response.Write 0:End If%>"> 元</td>
        </tr>
	<tr>
          <td align="right" class="TableRow2"><b>其它选项：</b></td>
          <td class="TableRow1" colspan="3"><input name="isTop" type="checkbox" id="isTop" value="1"<%If isEdit Then:If Rs("isTop") <> 0 Then Response.Write (" checked")%>>
            <%=sModuleName%>置顶 
            <input name="isBest" type="checkbox" id="isBest" value="1"<%If isEdit Then:If Rs("isBest") <> 0 Then Response.Write (" checked")%>> 
            <%=sModuleName%>推荐
	    <input name="showreg" type="checkbox" id="showreg" value="1"<%If isEdit Then:If Rs("showreg") <> 0 Then Response.Write (" checked")%>> 
            注册<%=sModuleName%>
	    <input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"<%If isEdit Then:If Rs("ForbidEssay") <> 0 Then Response.Write (" checked")%>> 
            禁止发表评论
	    <input name="isAccept" type="checkbox" id="isAccept" value="1" checked> 
            立即发布（<font color=blue>否则审核后才能发布。</font>）
	    <%If isEdit Then%>
	    <input name="Update" type="checkbox" value="1"> 
            同时更新<%=sModuleName%>发布时间<%End If%></td>
        </tr>
	<tr>
          <td align="right" class="TableRow2"><b>下载文件名称：</b></td>
          <td colspan="3" class="TableRow1"> <b>说明：</b><font color=blue>下载服务器路径 + 下载文件名称 = 完整下载地址</font><br>
	  <font color=red>如果选择了使用下载服务器，请在下面↓输入软件名称，前面请不要添加“/”。</font><br>
	  <%
	  Response.Write ReadDownAddress(softid)
	  %></td>
        </tr>
 <script language="javascript">
 function setid() {
	 str='';
	 if(!window.myform.no.value)
	 window.myform.no.value=1;
	 for(i=1;i<=window.myform.no.value;i++)
	 str+=''+'<b>下载名称：</b><input type="text" name="SiteName" value="下载地址'+i+'" size=12>&nbsp;<b>下载地址：</b><input type="text" name="DownAddress" size=70 value="">&nbsp;<BR>';
	 window.upid.innerHTML=str;
	 if (i==1) {
		downsite.style.display='none';
	 }else{
		downsite.style.display='';
	 }
 }
 </script>
	<tr>
          <td align="right" class="TableRow2"><b>设置下载数量：</b></td>
          <td colspan="3" class="TableRow1"><input type="text" name="no" value="1" size=2>&nbsp;&nbsp;<input type="button" name="Button" class=button onclick="setid();" value="添加下载地址数"> 
	  <input type='button' name='selectfile' value='从已上传文件中选择' onclick='SelectFile()' class=button></td>
        </tr>
<%
	If isEdit Then
		Response.Write ShowDownAddress(Rs("DownAddress"))
	End If
%>
	<tr id=downsite style="display='none';">
          <td colspan="4" id="upid" class="TableRow2"></td>
        </tr>
	<tr>
          <td align="right" class="TableRow2"><b>文件上传：</b></td>
          <td colspan="3" class="TableRow1"><iframe name="file1" frameborder=0 width='100%' height=45 scrolling=no src=upfile.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
        <tr align="center">
          <td colspan="4" class="TableRow2">
	  <input type="button" name="Submit2" onclick="CheckLength();" value="查看内容长度" class=Button>
	  <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>
	  <input type="Submit" name="Submit1" value="保存<%=sModuleName%>" class=Button></td>
        </tr></form>
      </table></div>
<%
	If isEdit Then
		Rs.Close
		Set Rs = Nothing
	End If
End Sub
Private Function ReadDownAddress(ByVal softid)
	Response.Write "<input name=""id"" type=""hidden"" value='0'>"  & vbCrLf
	Response.Write "<input name=""DownFileName"" id=DownFilePath type=""text"" size=""45""> " & vbCrLf
	Response.Write SelDownServer(0)
	Response.Write "&nbsp;&nbsp;<br>"
	If softid = 0 Then Exit Function
	Dim rsDown
	If CLng(softid) <> 0 Then
		Set rsDown = enchiasp.Execute("SELECT id,downid,DownFileName FROM ECCMS_DownAddress WHERE softid="& CLng(softid))
		Do While Not rsDown.EOF
			Response.Write "<input name=""id"" type=""hidden"" value='"
			Response.Write rsDown("id")
			Response.Write "'>" & vbCrLf
			Response.Write "<input name=""DownFileName"" type=""text"" size=""45"" value='"
			Response.Write rsDown("DownFileName")
			Response.Write "'> "
			Response.Write SelDownServer(rsDown("downid"))
			Response.Write " <a href='?action=sdel&ChannelID="
			Response.Write ChannelID
			Response.Write "&id="
			Response.Write rsDown("id")
			Response.Write "' class=showmenu onclick=""return confirm('您确定要删除吗?');"">删 除</a><br>"  & vbCrLf
			rsDown.movenext
		Loop
		Set rsDown = Nothing
	End If
End Function

Private Function ShowDownAddress(Address)
	Dim strDownAddress,sDownAddress,sDownSiteName
	Dim n,strTemp,AddressNum
	If IsNull(Address) Or Trim(Address) = "|||" Then
		ShowDownAddress = ""
		Exit Function
	End If
	On Error Resume Next
	strTemp = "<tr>	<td colspan=4 class=TableRow2><b>说明：</b><font color=blue>如果要删除下载地址，请清空下面的↓输入框或者输入“del”。</font><br>"
	strDownAddress = Split(Address, "|||")
	sDownAddress = Split(strDownAddress(1), "|")
	sDownSiteName = Split(strDownAddress(0), "|")
	If UBound(sDownAddress) < UBound(sDownAddress) Then
		AddressNum = UBound(sDownAddress)
	Else
		AddressNum = UBound(sDownSiteName)
	End If
	For n = 0 To AddressNum
		strTemp = strTemp & "<b>下载名称：</b><input type=text name=SiteName value='" & sDownSiteName(n) & "' size=12>&nbsp;<b>下载地址：</b><input type=text name=DownAddress size=70 value='" & sDownAddress(n) & "'>&nbsp;<BR>"
	Next
	strTemp = strTemp & "	</td></tr>"
	ShowDownAddress = strTemp
End Function
Function FormatDownAddress(ByVal str)
	If Trim(str) = ""  Or Trim(str) = "|||" Then
		FormatDownAddress = ""
		Exit Function
	End If
	Dim strDownAddress,sDownAddress,sDownSiteName
	Dim n,AddressNum,strAddress,strDownName,strTemp
	On Error Resume Next
	strDownAddress = Split(str, "|||")
	sDownAddress = Split(strDownAddress(1), "|")
	sDownSiteName = Split(strDownAddress(0), "|")
	If UBound(sDownAddress) < UBound(sDownSiteName) Then
		AddressNum = UBound(sDownAddress)
	Else
		AddressNum = UBound(sDownSiteName)
	End If
	strAddress = ""
	strDownName = ""
	For n = 0 To CInt(AddressNum)
		If Trim(sDownAddress(n)) <> "" And Trim(sDownSiteName(n)) <> "" Then
			strAddress = strAddress & Trim(sDownAddress(n)) & "|"
			strDownName = strDownName & Trim(sDownSiteName(n)) & "|"
		End If
	Next
	If Len(strDownName) > 0 Then strDownName = Left(strDownName, Len(strDownName) - 1)
	If Len(strAddress) > 0 Then strAddress = Left(strAddress, Len(strAddress) - 1)
	strTemp = strDownName & "|||" & strAddress
	FormatDownAddress = Trim(strTemp)
End Function
Private Sub CheckSave()

	If Trim(Request.Form("SoftName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "名称不能为空！</li>"
	End If
	If Len(Request.Form("SoftName")) => 200 Then
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
	If Len(Request.Form("Related")) => 220 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>相关" & sModuleName & "不能超过220个字符！</li>"
	End If
	If Trim(Request.Form("PointNum")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>下载所需的点数不能为空！如果不想设置请输入零。</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "星级不能为空。</li>"
	End If
	If Not IsNumeric(Request.Form("UserGroup")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "等级参数错误！</li>"
	End If
	If CLng(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该分类是外部连接，不能添加" & sModuleName & "！</li>"
	End If
	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该一级分类已经有下属分类，不能添加" & sModuleName & "！</li>"
	End If
	If Trim(Request.Form("SoftType")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "类型！</li>"
	End If
	If Trim(Request.Form("impower")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "授权方式！</li>"
	End If
	If Trim(Request.Form("Languages")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "语言！</li>"
	End If
	If Trim(Request.Form("AllHits")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>初始点击数不能为空！</li>"
	End If
	If Not IsNumeric(Request("AllHits")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>初始点击数请输入整数！</li>"
	End If
	If Not IsNumeric(Request("SpecialID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>专题ID参数错误！</li>"
	End If
	TextContent = ""
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next
	If Len(TextContent) < 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>软件简介不能为空！</li>"
	End If
	If CInt(Request.Form("isTop")) = 1 Then
		SoftTop = 1
	Else
		SoftTop = 0
	End If
	If CInt(Request.Form("isBest")) = 1 Then
		SoftBest = 1
	Else
		SoftBest = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If
	If CInt(Request.Form("showreg")) = 1 Then
		showreg = 1
	Else
		showreg = 0
	End If
	If CInt(Request("isAccept")) = 1 Then
		SoftAccept = 1
	Else
		SoftAccept = 0
	End If
	If Len(Request.Form("RunSystem")) = 0 Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>运行环境不能为空！</li>"
	End If
	If Not IsNumeric(Request.Form("SoftSize")) Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "大小请输入整数！</li>"
	End If
	'---- 提取下载地址表单中的数据
	Dim TempAddress,TempSiteName
	Dim strTempAddress,strTempSiteName
	If Trim(Request.Form("DownAddress")) <> "" And Trim(Request.Form("SiteName")) <> "" Then
		strTempAddress = ""
		For Each TempAddress In Request.Form("DownAddress")
			If LCase(TempAddress) <> "del" And Trim(TempAddress) <> "" Then
				strTempAddress = strTempAddress & Replace(TempAddress, "|", "") & "|"
			End If
		Next
		If Len(strTempAddress) > 0 Then strTempAddress = Left(strTempAddress, Len(strTempAddress) - 1)
		strTempSiteName = ""
		For Each TempSiteName In Request.Form("SiteName")
			If LCase(TempSiteName) <> "del" And Trim(TempSiteName) <> "" Then
				strTempSiteName = strTempSiteName & Replace(TempSiteName, "|", "") & "|"
			End If
		Next
		If Len(strTempSiteName) > 0 Then strTempSiteName = Left(strTempSiteName, Len(strTempSiteName) - 1)
		DownAddress = enchiasp.CheckStr(strTempSiteName &"|||"& strTempAddress)
	Else
		DownAddress = ""
	End If
	DownAddress = FormatDownAddress(DownAddress)
	'---- 提取下载地址表单数据完成
End Sub

Private Sub SaveSoft()
	CheckSave
	If Trim(Request.Form("DownFileName")) <> "" And Trim(Request.Form("downid")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择下载服务器！</li>"
	End If
	If Trim(Request.Form("DownFileName")) = "" And Trim(Request.Form("downid")) <> 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>您已经选择了下载服务器，请填写下载文件名称！</li>"
	End If
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_SoftList WHERE (SoftID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("SpecialID") = Trim(Request.Form("SpecialID"))
		Rs("SoftName") = enchiasp.ChkFormStr(Request.Form("SoftName"))
		Rs("SoftVer") = enchiasp.ChkFormStr(Request.Form("SoftVer"))
		Rs("ColorMode") = Trim(Request.Form("ColorMode"))
		Rs("FontMode") = Trim(Request.Form("FontMode"))
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Content") = enchiasp.JAPEncode(TextContent)
		Rs("Languages") = enchiasp.ChkFormStr(Request.Form("Languages"))
		Rs("SoftType") = enchiasp.ChkFormStr(Request.Form("SoftType"))
		Rs("RunSystem") = enchiasp.ChkFormStr(Replace(Replace(Request.Form("RunSystem"), ",", "/"), " ", ""))
		Rs("impower") = enchiasp.ChkFormStr(Request.Form("impower"))
		If UCase(Request.Form("SizeUnit")) = "MB" Then
			Rs("SoftSize") = CLng(Request.Form("SoftSize") * 1024)
		Else
			Rs("SoftSize") = CLng(Request.Form("SoftSize"))
		End If
		Rs("star") = Trim(Request.Form("star"))
		Rs("Homepage") = Trim(Request.Form("Homepage"))
		Rs("Contact") = Trim(Request.Form("Contact"))
		Rs("Author") = enchiasp.ChkFormStr(Request.Form("Author"))
		Rs("Regsite") = enchiasp.ChkFormStr(Request.Form("Regsite"))
		Rs("showreg") = CInt(showreg)
		Rs("username") = Trim(AdminName)
		Rs("PointNum") = Trim(Request.Form("PointNum"))
		Rs("SoftPrice") = Trim(Request.Form("SoftPrice"))
		Rs("SoftTime") = Now()
		Rs("isTop") = SoftTop
		Rs("AllHits") = Trim(Request.Form("AllHits"))
		Rs("DayHits") = 0
		Rs("WeekHits") = 0
		Rs("MonthHits") = 0
		Rs("HitsTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("SoftImage") = Trim(Request.Form("SoftImage"))
		Rs("Decode") = Trim(Request.Form("Decode"))
		Rs("isBest") = SoftBest
		Rs("UserGroup") = Trim(Request.Form("UserGroup"))
		Rs("isUpdate") = 1
		Rs("ErrCode") = 0
		Rs("DownAddress") = Trim(DownAddress)
		Rs("isAccept") = SoftAccept
		Rs("ForbidEssay") = ForbidEssay
		Rs("AlphaIndex") = enchiasp.ReadAlpha(Request.Form("SoftName"))
	Rs.update
	Rs.Close
	Rs.Open "select top 1 softid from ECCMS_SoftList where ChannelID=" & ChannelID & " order by softid desc", Conn, 1, 1
	SoftID = Rs("SoftID")
	Rs.Close:Set Rs = Nothing
	AddDownAddress CLng(Request.Form("downid")),softid,Trim(Request.Form("DownFileName"))
	ClassUpdateCount Request.Form("ClassID"),1
	Call RemoveCache
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makesoft.asp?ChannelID=" & ChannelID & "&softid=" & softid & "&showid=0"	
		Call ScriptCreation(url,softid)
		SQL = "SELECT TOP 1 SoftID FROM ECCMS_SoftList WHERE ChannelID=" & ChannelID & " And isAccept <> 0 And SoftID < " & SoftID & " ORDER BY SoftID DESC"
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.EOF And Rs.BOF) Then
			url = "admin_makesoft.asp?ChannelID=" & ChannelID & "&softid=" & Rs("SoftID") & "&showid=0"	
			Call ScriptCreation(url,Rs("SoftID"))
		End If
		Rs.Close
		Set Rs = Nothing
	End If
	Succeed("<li>恭喜您！添加新的" & sModuleName & "成功。</li><li><a href=?action=view&ChannelID=" & ChannelID & "&SoftID=" & SoftID & ">点击此处查看该" & sModuleName & "</a></li><li><a href=?action=add&ChannelID=" & ChannelID & "&classid=" & Request.Form("ClassID") & "><font color=blue>点击此处继续添加" & sModuleName & "</font></a></li>")
End Sub

Private Function AddDownAddress(downid,softid,DownFileName)
	If CLng(downid) <> 0 And Trim(DownFileName) <> "" Then
		SQL = "Insert Into ECCMS_DownAddress (ChannelID,softid,downid,DownFileName) values (" & ChannelID & "," & softid & "," & downid & ",'" & DownFileName & "')"
		enchiasp.Execute(SQL)
	End If
End Function

Private Function UpdateDownAddress(softid)
	Dim valDownID,valDownFileName,valDownAddressID
	valDownID = Split(Request("downid"), ",")
	valDownFileName = Split(Request("DownFileName"), ",")
	valDownAddressID = Split(Request("id"), ",")
	For i = 0 To UBound(valDownFileName)
		If i = 0 Then
			If Trim(valDownFileName(0)) <> "" And Trim(valDownID(0)) <> "0" Then
				enchiasp.Execute("Insert Into ECCMS_DownAddress (ChannelID,softid,downid,DownFileName) values (" & ChannelID & "," & softid & "," & Trim(valDownID(0)) & ",'" & Trim(valDownFileName(0)) & "')")
			End If
		Else
			If Trim(valDownFileName(i)) <> "" And Trim(valDownID(i)) <> "0" Then
				enchiasp.Execute ("UPDATE ECCMS_DownAddress SET downid=" & Trim(valDownID(i)) & ",DownFileName='" & Trim(valDownFileName(i)) & "' WHERE id="& CLng(valDownAddressID(i)))
			End If
		End If
	Next
End Function

Private Sub ModifySoft()
	CheckSave
	If Founderr = True Then Exit Sub
	Dim Auditing
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_SoftList WHERE SoftID=" & Request("SoftID")
	Rs.Open SQL,Conn,1,3
		Auditing = Rs("isAccept")
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("SpecialID") = Trim(Request.Form("SpecialID"))
		Rs("SoftName") = enchiasp.ChkFormStr(Request.Form("SoftName"))
		Rs("SoftVer") = enchiasp.ChkFormStr(Request.Form("SoftVer"))
		Rs("ColorMode") = Trim(Request.Form("ColorMode"))
		Rs("FontMode") = Trim(Request.Form("FontMode"))
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Content") = enchiasp.JAPEncode(TextContent)
		Rs("Languages") = enchiasp.ChkFormStr(Request.Form("Languages"))
		Rs("SoftType") = enchiasp.ChkFormStr(Request.Form("SoftType"))
		Rs("RunSystem") = enchiasp.ChkFormStr(Replace(Replace(Request.Form("RunSystem"), ", ", "/"), " ", ""))
		Rs("impower") = enchiasp.ChkFormStr(Request.Form("impower"))
		If UCase(Request.Form("SizeUnit")) = "MB" Then
			Rs("SoftSize") = CLng(Request.Form("SoftSize") * 1024)
		Else
			Rs("SoftSize") = CLng(Request.Form("SoftSize"))
		End If
		Rs("star") = Trim(Request.Form("star"))
		Rs("Homepage") = Trim(Request.Form("Homepage"))
		Rs("Contact") = Trim(Request.Form("Contact"))
		Rs("Author") = enchiasp.ChkFormStr(Request.Form("Author"))
		Rs("Regsite") = enchiasp.ChkFormStr(Request.Form("Regsite"))
		Rs("showreg") = CInt(showreg)
		Rs("PointNum") = Trim(Request.Form("PointNum"))
		Rs("SoftPrice") = Trim(Request.Form("SoftPrice"))
		If CInt(Request.Form("Update")) = 1 Then Rs("SoftTime") = Now()
		Rs("isTop") = SoftTop
		Rs("AllHits") = Trim(Request.Form("AllHits"))
		Rs("SoftImage") = Trim(Request.Form("SoftImage"))
		Rs("Decode") = Trim(Request.Form("Decode"))
		Rs("isBest") = SoftBest
		Rs("UserGroup") = Trim(Request.Form("UserGroup"))
		Rs("isUpdate") = 1
		Rs("ErrCode") = 0
		Rs("DownAddress") = Trim(DownAddress)
		Rs("isAccept") = SoftAccept
		Rs("ForbidEssay") = ForbidEssay
		Rs("AlphaIndex") = enchiasp.ReadAlpha(Request.Form("SoftName"))
	Rs.update
	SoftID = Rs("SoftID")
	If SoftAccept = 1 And Auditing = 0 Then
		AddUserPointNum Rs("username"),1
	End If
	If SoftAccept = 0 And Auditing = 1 Then
		AddUserPointNum Rs("username"),0
	End If
	Rs.Close:Set Rs = Nothing
	UpdateDownAddress(softid)
	Call RemoveCache
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makesoft.asp?ChannelID=" & ChannelID & "&softid=" & softid & "&showid=0"	
		Call ScriptCreation(url,softid)
	End If
	Succeed("<li>恭喜您！修改" & sModuleName & "成功。</li><li><a href=admin_Soft.asp?action=view&ChannelID=" & ChannelID & "&SoftID=" & SoftID & ">点击此处查看该" & sModuleName & "</a></li>")
End Sub
Private Sub SoftView()
	Call PageTop
	If Request("SoftID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	On Error Resume Next
	Dim blnDownAddress:blnDownAddress = False
	SQL = "select * from ECCMS_SoftList where ChannelID=" & ChannelID & " And SoftID=" & Request("SoftID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！没有找到任何" & sModuleName & "。或者您选择了错误的系统参数！</li>"
		Exit Sub
	Else
	Dim strDownAddress,sDownAddress
	If Not IsNull(Rs("DownAddress")) And Rs("DownAddress") <> "" Then
		strDownAddress = Split(Rs("DownAddress"), "|||")
		sDownAddress = Split(strDownAddress(1), "|")
		blnDownAddress = True
	Else
		blnDownAddress = False
	End If
%>

<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
	<tr>
	  <th colspan="2">查看<%=sModuleName%></th>
	</tr>
	<tr>
	  <td align="center" class="TableRow2" colspan="2"><a href=?action=edit&ChannelID=<%=ChannelID%>&SoftID=<%=Rs("SoftID")%>><font size=4><%=enchiasp.ReadFontMode(Rs("SoftName"),Rs("ColorMode"),Rs("FontMode"))%></font></a></td>
	</tr>
	<tr>
	  <td class="TableRow1"><b><%=sModuleName%>运行环境：</b> <%=Rs("RunSystem")%></td>
	  <td class="TableRow1"><b><%=sModuleName%>类型：</b> <%=Rs("SoftType")%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><b><%=sModuleName%>大小：</b> <%=Rs("SoftSize")%> <b>KB</b></td>
	  <td class="TableRow1"><b><%=sModuleName%>星级：</b> 
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
	  <td class="TableRow1"><b><%=sModuleName%>语言：</b> <%=Rs("Languages")%></td>
	  <td class="TableRow1"><b>授权方式：</b> <%=Rs("impower")%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><b>更新时间：</b> <%=Rs("SoftTime")%></td>
	  <td class="TableRow1"><b>程序主页：</b> <%=Rs("Homepage")%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><b>发 布 人：</b> <font color=blue><%=Rs("username")%></font></td>
	  <td class="TableRow1"><b>审核状态：</b> <%If Rs("isAccept") > 0 Then%>
	  <font color=blue>已审核</font>
	  <%Else%>
	  <font color=red>未审核</font>
	  <%End If%>
	  </td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1"><b><%=sModuleName%>简介：</b><br><%=UBBCode(Rs("content"))%></td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1"><b>下载地址：</b><br>
<%
	If blnDownAddress Then
		For i = 0 To UBound(sDownAddress)
			Response.Write "<li><a href=""" & sDownAddress(i) & """ target=_blank>" & sDownAddress(i) & "</a></li>" & vbNewLine
		Next
	End If
	Response.Write SoftDownAddress(Rs("SoftID"))
%>
	  </td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1">上一<%=sModuleName%>：<%=FrontSoft(Rs("SoftID"))%>
	  <br>下一<%=sModuleName%>：<%=NextSoft(Rs("SoftID"))%></td>
	</tr>
	<tr>
	  <td align="center" colspan="2" class="TableRow2"><input type="button" onclick="{if(confirm('您确定要删除此软件吗?')){location.href='?action=del&ChannelID=<%=ChannelID%>&softid=<%=Rs("softid")%>';return true;}return false;}" value="删除软件" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="返回上一页" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="返回顶部" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit2" onclick="javascript:location.href='?action=edit&ChannelID=<%=ChannelID%>&SoftID=<%=Rs("SoftID")%>'" value="编辑软件" class=button>&nbsp;&nbsp;
	  [<a href="?act=批量审核&ChannelID=<%=ChannelID%>&selSoftID=<%=Rs("SoftID")%>" onclick="return confirm('您确定执行审核操作吗?');" ><font color=blue>直接审核</font></a>]</td>
	</tr>
</table>

<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub
Private Function SoftDownAddress(softid)
	Dim rsAddress, sqlAddress, rsDown
	Dim strDownAddress,sDownAddress
	strDownAddress = ""
	On Error Resume Next
	Set rsDown = enchiasp.Execute("SELECT downid,DownFileName FROM [ECCMS_DownAddress] WHERE softid=" & CLng(softid))
	If Not (rsDown.BOF And rsDown.EOF) Then
		Do While Not rsDown.EOF
			sqlAddress = "SELECT downid,DownloadName,DownloadPath FROM ECCMS_DownServer WHERE ChannelID=" & ChannelID & " And depth=1 And rootid =" & rsDown("downid") & " And isLock=0 ORDER BY orders ASC"
			Set rsAddress = enchiasp.Execute(sqlAddress)
			If Not(rsAddress.EOF And rsAddress.BOF) Then
				Do While Not rsAddress.EOF
					strDownAddress = rsAddress("DownloadPath") & rsDown("DownFileName")
					sDownAddress = sDownAddress & "<li><a href=""" & strDownAddress & """ target=_blank>" & strDownAddress & "</a></li>" & vbNewLine
					rsAddress.movenext
				Loop
			End If
			Set rsAddress = Nothing
			rsDown.movenext
		Loop
	End If
	Set rsDown = Nothing
	SoftDownAddress = sDownAddress
End Function
Private Function FrontSoft(SoftID)
	Dim Rss, SQL
	SQL = "select Top 1 SoftID,classid,SoftName from ECCMS_SoftList where ChannelID=" & ChannelID & " And isAccept <> 0 And SoftID < " & SoftID & " order by SoftID desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		FrontSoft = "已经没有了"
	Else
		FrontSoft = "<a href=admin_Soft.asp?action=view&ChannelID=" & ChannelID & "&SoftID=" & Rss("SoftID") & ">" & Rss("SoftName") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function NextSoft(SoftID)
	Dim Rss, SQL
	SQL = "select Top 1 SoftID,classid,SoftName from ECCMS_SoftList where ChannelID=" & ChannelID & " And isAccept <> 0 And SoftID > " & SoftID & " order by SoftID asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		NextSoft = "已经没有了"
	Else
		NextSoft = "<a href=admin_Soft.asp?action=view&ChannelID=" & ChannelID & "&SoftID=" & Rss("SoftID") & ">" & Rss("SoftName") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Sub BatCreateHtml()
	Dim AllSoftID,url
	Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
	Response.Write "<ol>"
	AllSoftID = Split(selSoftID, ",")
	For i = 0 To UBound(AllSoftID)
		softid = CLng(AllSoftID(i))
		url = "admin_makesoft.asp?ChannelID=" & ChannelID & "&softid=" & softid & "&showid=1"	
		Call ScriptCreation(url,softid)
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
Private Sub DelDownAddress()
	If Not IsNumeric(Request("id")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ID参数错误！</li>"
		Exit Sub
	End If
	enchiasp.Execute("Delete From ECCMS_DownAddress Where ChannelID = "& ChannelID &" And id=" & Request("id"))
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Sub SoftDel()
	If Trim(Request("SoftID")) = "" Then
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	On Error Resume Next
	Set Rs = enchiasp.Execute("SELECT softid,classid,username,HtmlFileDate FROM ECCMS_SoftList WHERE ChannelID = "& ChannelID &" And SoftID=" & Request("SoftID"))
	If Not(Rs.BOF And Rs.EOF) Then
		ClassUpdateCount Rs("classid"),0
		AddUserPointNum Rs("username"),0
		DeleteHtmlFile Rs("classid"),Rs("softid"),Rs("HtmlFileDate")
	End If
	Rs.Close:Set Rs = Nothing
	Conn.Execute("Delete From ECCMS_SoftList Where ChannelID = "& ChannelID &" And SoftID=" & Request("SoftID"))
	Conn.Execute("Delete From ECCMS_DownAddress Where ChannelID = "& ChannelID &" And SoftID=" & Request("SoftID"))
	Conn.Execute ("delete from ECCMS_Comment where ChannelID = "& ChannelID &" And PostID=" & Request("SoftID"))
	Call RemoveCache
	Response.redirect ("admin_soft.asp?ChannelID=" & ChannelID)
End Sub

Private Sub batdel()
	Set Rs = enchiasp.Execute("SELECT softid,classid,username,HtmlFileDate FROM ECCMS_SoftList WHERE ChannelID = "& ChannelID &" And SoftID in (" & selSoftID & ")")
	Do While Not Rs.EOF
		ClassUpdateCount Rs("classid"),0
		AddUserPointNum Rs("username"),0
		DeleteHtmlFile Rs("classid"),Rs("softid"),Rs("HtmlFileDate")
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	Conn.Execute ("delete from ECCMS_SoftList where ChannelID = "& ChannelID &" And SoftID in (" & selSoftID & ")")
	Conn.Execute ("delete from ECCMS_DownAddress where ChannelID = "& ChannelID &" And SoftID in (" & selSoftID & ")")
	Conn.Execute ("delete from ECCMS_Comment where ChannelID = "& ChannelID &" And PostID in (" & selSoftID & ")")
	Call RemoveCache
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub isCommend()
	enchiasp.Execute ("update ECCMS_SoftList set isBest=1 where SoftID in (" & selSoftID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noCommend()
	enchiasp.Execute ("update ECCMS_SoftList set isBest=0 where SoftID in (" & selSoftID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub isTop()
	enchiasp.Execute ("update ECCMS_SoftList set isTop=1 where SoftID in (" & selSoftID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noTop()
	enchiasp.Execute ("update ECCMS_SoftList set isTop=0 where SoftID in (" & selSoftID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
'----批量审核
Private Sub BatAccept()
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_SoftList WHERE isAccept=0 And SoftID in (" & selSoftID & ")")
	Do While Not Rs.EOF
		AddUserPointNum Rs("username"),1
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("update ECCMS_SoftList set isAccept=1 where SoftID in (" & selSoftID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Sub NotAccept()
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_SoftList WHERE isAccept>0 And SoftID in (" & selSoftID & ")")
	Do While Not Rs.EOF
		AddUserPointNum Rs("username"),0
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("update ECCMS_SoftList set isAccept=0 where SoftID in (" & selSoftID & ")")
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
Private Sub BatchSetting()
	If Not ChkAdmin("AdminSoft" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Dim Channel_Setting
	Channel_Setting = Split(enchiasp.Channel_Setting, "|||")
	Response.Write "<script src=""include/SoftJuge.Js"" type=""text/javascript""></script>" & vbNewLine
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
	Response.Write "		<td class=tablerow1><input type=""text"" name=""SoftID"" size=70 value='"& Request("selSoftID") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 width=""15%"" align=right><b>相关" & sModuleName & "：</b></td>"
	Response.Write "		<td class=tablerow1 width=""5%"" align=center><input type=checkbox name=selRelated value='1'></td>"
	Response.Write "		<td class=tablerow1 width=""60%""><input name=Related type=text size=60></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>运行环境：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selRunSystem value='1'></td>"
	Response.Write "		<td class=tablerow1 ><input name=RunSystem type=text size=60><br>"
	Dim RunSystem
	RunSystem = Split(Channel_Setting(0), "|")
	For i = 0 To UBound(RunSystem)
		Response.Write "<a href='javascript:ToRunsystem(""" & Trim(RunSystem(i)) & """)'><u>" & Trim(RunSystem(i)) & "</u></a> | "
		If i = 10 Then Response.Write "<br>"
	Next
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "类型：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selSoftType value='1'></td>"
	Response.Write "		<td class=tablerow1 >"
	Dim SoftType
	SoftType = Split(Channel_Setting(2), ",")
	For i = 0 To UBound(SoftType)
		Response.Write "	<input type=""radio"" name=""SoftType"" value=""" & Trim(SoftType(i)) & """ "
		If i = 0 Then Response.Write " checked"
		Response.Write ">" & Trim(SoftType(i)) & " "
		If i = 6 Then Response.Write "<br>"
	Next
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
	Response.Write "		<td class=tablerow1 align=right><b>授权方式：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selImpower value='1'></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "		<select name=Impower>"
	Dim ImpowerStr
	ImpowerStr = Split(Channel_Setting(3), ",")
	For i = 0 To UBound(ImpowerStr)
		Response.Write "	<option value=""" & ImpowerStr(i) & """>" & ImpowerStr(i) & "</option>"
	Next
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "语言：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selLanguages value='1'></td>"
	Response.Write "		<td class=tablerow1>"
	Response.Write "		<select name=Languages>"
	Dim LanguagesStr
	LanguagesStr = Split(Channel_Setting(4), ",")
	For i = 0 To UBound(LanguagesStr)
		Response.Write "	<option value=""" & LanguagesStr(i) & """>" & LanguagesStr(i) & "</option>"
	Next
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
	If Not ChkAdmin("AdminSoft" & ChannelID) Then
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
	Response.Write " <b>指定" & sModuleName & "ID：</b> <input type=""text"" name=""SoftID"" size=80 value='"& Request("selSoftID") &"'></td>"
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
		If Trim(Request.Form("SoftID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_SoftList SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And SoftID in ("& Request("SoftID") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_SoftList SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>恭喜您！批量移动完成。</li>")
End Sub

Private Sub BatcDelete()
	If Not ChkAdmin("AdminSoft" & ChannelID) Then
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
	Response.Write "		<td class=tablerow1><b>分类ID：</b><input type=""text"" name=""SoftID"" size=80 value='"& Request("selSoftID") &"'></td>"
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
Private Sub AllDelSoft()
	On Error Resume Next
	If CInt(Request.Form("Appointed")) = 1 Then
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		End If
		enchiasp.Execute ("DELETE ECCMS_DownAddress FROM ECCMS_SoftList A INNER JOIN ECCMS_DownAddress D ON D.SoftID=A.SoftID WHERE A.ChannelID = "& ChannelID &" And A.ClassID IN (" & Request("ClassID") & ")")
		enchiasp.Execute ("DELETE ECCMS_Comment FROM ECCMS_SoftList A INNER JOIN ECCMS_Comment C ON C.PostID=A.SoftID WHERE A.ChannelID = "& ChannelID &" And A.ClassID IN (" & Request("ClassID") & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_SoftList WHERE ChannelID = "& ChannelID &" And ClassID IN (" & Request("ClassID") & ")")
	ElseIf CInt(Request.Form("Appointed")) = 2 Then
		enchiasp.Execute ("DELETE FROM ECCMS_SoftList WHERE ChannelID = "& ChannelID)
		enchiasp.Execute ("DELETE FROM ECCMS_DownAddress WHERE ChannelID = "& ChannelID)
		enchiasp.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID)
	Else
		If Trim(Request.Form("SoftID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		End If
		enchiasp.Execute ("DELETE FROM ECCMS_SoftList WHERE ChannelID = "& ChannelID &" And SoftID IN (" & Request("SoftID") & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_DownAddress WHERE ChannelID = "& ChannelID &" And SoftID IN (" & Request("SoftID") & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID IN (" & Request("SoftID") & ")")
		
	End If
	Call RemoveCache
	Succeed("<li>批量删除成功！</li>")
End Sub

Private Sub SaveSetting()
	If Founderr = True Then Exit Sub
	Dim strTempValue
	strTempValue = ""
	If Trim(Request.Form("selRelated")) <> "" Then strTempValue = strTempValue & "Related='"& enchiasp.ChkFormStr(Request.Form("Related")) &"',"
	If Trim(Request.Form("selRunSystem")) <> "" Then strTempValue = strTempValue & "RunSystem='"& enchiasp.ChkFormStr(Request.Form("RunSystem")) &"',"
	If Trim(Request.Form("selSoftType")) <> "" Then strTempValue = strTempValue & "SoftType='"& enchiasp.ChkFormStr(Request.Form("SoftType")) &"',"
	If Trim(Request.Form("selAuthor")) <> "" Then strTempValue = strTempValue & "Author='"& enchiasp.ChkFormStr(Request.Form("Author")) &"',"
	If Trim(Request.Form("selImpower")) <> "" Then strTempValue = strTempValue & "Impower='"& enchiasp.ChkFormStr(Request.Form("Impower")) &"',"
	If Trim(Request.Form("selLanguages")) <> "" Then strTempValue = strTempValue & "Languages='"& enchiasp.ChkFormStr(Request.Form("Languages")) &"',"
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
		If Trim(Request.Form("SoftID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_SoftList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And SoftID in ("& Request("SoftID") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		ElseIf Trim(Request.Form("ClassID")) = "-1" Then
			SQL = "UPDATE ECCMS_SoftList SET "& strTempValue &" WHERE ChannelID = "& ChannelID &" And isAccept>0"
		Else
			SQL = "UPDATE ECCMS_SoftList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>恭喜您！批量设置完成。</li>")
End Sub
Private Sub ResetDateTime()
	Server.ScriptTimeOut = 9999
	Response.Write "<br><table width='400' align=center border=0 cellspacing=1 cellpadding=1>" & vbCrLf
	Response.Write "<tr> " & vbCrLf
	Response.Write "<td style=""border: 1px #384780 solid ;background-color: #FFFFFF;"">" & vbCrLf
	Response.Write "<table width=0 id=table2 name=table2 border=0 cellspacing=0 cellpadding=0>" & vbCrLf
	Response.Write "<tr height=12><td bgcolor=#36D91A>" & vbCrLf
	Response.Write "</td></tr></table></td></tr><tr> " & vbCrLf
	Response.Write "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span> <span style=""font-size:9pt"">%</span></td></tr>" & vbCrLf
	Response.Write "</table><br>" & vbCrLf
	Response.Flush
	Dim totalrec,SoftTime,page_count,pagelist
	i = 0
	page_count = 0
	totalrec = enchiasp.Execute("SELECT COUNT(SoftID) FROM [ECCMS_SoftList] WHERE ChannelID = "& ChannelID &" And isAccept>0")(0)
	Set Rs = enchiasp.Execute("SELECT SoftID,SoftTime FROM [ECCMS_SoftList] WHERE ChannelID = "& ChannelID &" And isAccept>0 ORDER BY SoftTime DESC")
	If Not (Rs.BOF And Rs.EOF) Then
		SQL=Rs.GetRows(-1)
		For pagelist=0 To Ubound(SQL,2)
			If Not Response.IsClientConnected Then Response.End
			Response.Write "<script>"
			Response.Write "table2.style.width=" & Fix((page_count / totalrec) * 400) & ";"
			Response.Write "txt2.innerHTML=""完成：" & FormatNumber(page_count / totalrec * 100, 2, -1) & """;"
			Response.Write "</script>" & vbCrLf
			Response.Flush
			SoftTime = DateAdd("s", -i, SQL(1,pagelist))
			enchiasp.Execute ("UPDATE [ECCMS_SoftList] SET SoftTime='" & SoftTime & "' WHERE SoftID="& SQL(0,pagelist))
			i = i + 5
			page_count = page_count + 1
		Next
		SQL=Null
	End If
	Set Rs = Nothing
	Response.Write "<script>table2.style.width=400;txt2.innerHTML=""完成：100"";</script>"
End Sub

Private Sub RemoveCache()
	enchiasp.DelCahe "RenewStatistics"
	enchiasp.DelCahe "TotalStatistics"
End Sub
%>