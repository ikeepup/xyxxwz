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
Dim Action
Dim i,ii,isEdit,RsObj
Dim keyword,FindWord,strClass
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Dim s_ClassName,ClassID,ChildStr,FoundSQL,isAccept,selArticleID
Dim TextContent,ArticleTop,ArticleBest,ArticleID,ForbidEssay,ArticleAccept
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If Trim(Request("isAccept")) <> "" Then
	isAccept = 0
Else
	isAccept = 1
End If
If CInt(ChannelID) = 0 Then ChannelID = 1
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "save"
	Call SaveArticle
Case "modify"
	Call ModifyArticle
Case "add"
	isEdit = False
	Call ArticleEdit(isEdit)
Case "edit"
	isEdit = True
	Call ArticleEdit(isEdit)
Case "del"
	Call ArticleDel
Case "batdel"
	Call PageTop
	Call BatcDelete
Case "alldel"
	Call AllDelArticle
Case "view"
	Call ArticleView
Case "renew"
	Call ArticleRenew
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
	Response.Write "	  <th colspan=2>" & sChannelName & "管理选项</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr><form method=Post name=myform action=admin_article.asp onSubmit='return JugeQuery(this);'>"
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "	  <td class=TableRow1>搜索："
	Response.Write "	  <input name=keyword type=text size=20>"
	Response.Write "	  条件："
	Response.Write "	  <select name=queryopt>"
	Response.Write "		<option value=1 selected>文章标题</option>"
	Response.Write "		<option value=2>录 入 者</option>"
	Response.Write "		<option value=3>不限条件</option>"
	Response.Write "	  </select> <input type=submit name=Submit value='开始搜索' class=Button></td>"
	Response.Write "	  <td class=TableRow1>" & sModuleName & "导航："
	Dim srtClassMenu
	Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Response.Write "<option value='admin_article.asp?ChannelID=" & ChannelID & "'>≡全部" & sModuleName & "列表≡</option>" & vbCrLf
	srtClassMenu = enchiasp.ClassJumpMenu(ChannelID)
	srtClassMenu = Replace(srtClassMenu, "{ClassID=" & Request("sortid") & "}", "selected")
	Response.Write srtClassMenu
	Response.Write "</select>"
	Response.Write "	  </td></form>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td colspan=2 class=TableRow2><b>操作选项：</b> <a href='admin_article.asp?ChannelID=" & ChannelID & "'>管理首页</a> | "
	Response.Write "	  <a href='admin_article.asp?ChannelID=" & ChannelID & "&action=add'>添加" & sModuleName & "</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "&action=add'>添加" & sModuleName & "分类</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "'>" & sModuleName & "分类管理</a> | "
	Response.Write "	  <a href='?ChannelID=" & ChannelID & "&specialID=0'>" & sModuleName & "专题列表</a> | "
	Response.Write "	  <a href='?ChannelID=" & ChannelID & "&isAccept=0'>待审" & sModuleName & "</a> | "
	Response.Write "	  <a href='admin_createArticle.asp?ChannelID=" & ChannelID & "'>生成HTML</a></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub

Private Sub showmain()
	If Not ChkAdmin("AdminArticle" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	If Not IsEmpty(Request("selArticleID")) Then
		selArticleID = Request("selArticleID")
		Select Case enchiasp.CheckStr(Request("act"))
		Case "批量删除":Call batdel
		Case "批量推荐":Call isCommend
		Case "取消推荐":Call noCommend
		Case "批量置顶":Call isTop
		Case "取消置顶":Call noTop
		Case "批量审核":Call BatAccept
		Case "取消审核":Call NotAccept
		Case "生成HTML"
			Call BatCreateHtml
		Case Else
			Response.Write "无效参数！"
		End Select
	End If
	Call PageTop
	Dim sortid,specialID,Cmd,limitime
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>选择</th>"
	Response.Write "	  <th width='68%'>" & sModuleName & "标题</th>"
	Response.Write "	  <th width='9%' nowrap>管理操作</th>"
	Response.Write "	  <th width='5%' nowrap>审核</th>"
	Response.Write "	  <th width='9%' nowrap>录 入 者</th>"
	Response.Write "	  <th width='9%' nowrap>整理日期</th>"
	Response.Write "	</tr>"
	If Request("sortid") <> "" Then
		SQL = "select ClassID,ChannelID,ClassName,ChildStr from [ECCMS_Classify] where ChannelID = " & ChannelID & " And ClassID=" & Request("sortid")
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			Response.Write "Sorry！没有找到任何" & sModuleName & "分类。或者您选择了错误的系统参数!"
			Response.End
		Else
			s_ClassName = Rs("ClassName")
			ClassID = Rs("ClassID")
			ChildStr = Rs("ChildStr")
			sortid = Request("sortid")
		End If
		Rs.Close
	Else
		s_ClassName = "全部" & sModuleName
		sortid = 0
	End If
	maxperpage = 30 '###每页显示数
	Dim strListName
	If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
		Response.Write ("错误的系统参数!请输入整数")
		Response.End
	End If
	If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
		CurrentPage = CInt(Request("page"))
	Else
		CurrentPage = 1
	End If
	If CLng(CurrentPage) = 0 Then CurrentPage = 1
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
			FoundSQL = "A.isAccept = "& isAccept & " And A.ClassID in (" & ChildStr & ")"
		Else
			If Trim(Request("specialID")) <> "" Then
				specialID = Request("specialID")
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
	
	strListName = "&channelid="& ChannelID &"&sortid="& Request("sortid") &"&specialID="& Request("specialID") &"&isAccept="& Request("isAccept") &"&keyword=" & Request("keyword") 
	totalnumber = enchiasp.Execute("Select Count(ArticleID) from ECCMS_Article A where A.ChannelID = " & ChannelID & " And "& FoundSQL &"")(0)
	TotalPageNum = CLng(totalnumber / maxperpage)  '得到总页数
	If TotalPageNum < totalnumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	On Error Resume Next
	SQL = "select A.ArticleID,A.ChannelID,A.ClassID,A.SpecialID,A.title,A.ColorMode,A.FontMode,A.BriefTopic,A.isTop,A.AllHits,A.WriteTime,A.username,A.isBest,A.isAccept,C.ClassName from [ECCMS_Article] A inner join [ECCMS_Classify] C on A.ClassID=C.ClassID where A.ChannelID = " & ChannelID & " And "& FoundSQL &" order by A.isTop DESC, A.WriteTime desc ,A.ArticleID desc"
	Rs.Open SQL, Conn, 1,1
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
		i=0
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
			Admin_Article_list SQL(0,i),SQL(1,i),SQL(2,i),SQL(4,i),SQL(5,i),SQL(6,i),SQL(7,i),SQL(8,i),SQL(10,i),SQL(11,i),SQL(12,i),SQL(13,i),SQL(14,i),strClass
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
	  <input class=Button type="submit" name="Submit2" value="执行操作" onclick="return confirm('您确定执行该操作吗?')">
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

Function Admin_Article_list(ArticleID,ChannelID,ClassID,title,ColorMode,FontMode,BriefTopic,isTop,WriteTime,UserName,isBest,isAccept,ClassName,strClass)
	Response.Write "	<tr>"
	Response.Write "	  <td align=center "& strClass &"><input type=checkbox name=selArticleID value=" & ArticleID & "></td>"
	Response.Write "	  <td "& strClass &">"
	If isTop <> 0 Then
		Response.Write "<img src=""images/gotop.gif"" border=0 alt=置顶文章 align=absMiddle>"
	End If

	Response.Write "[<a href=?ChannelID=" & ChannelID & "&sortid="
	Response.Write ClassID
	Response.Write ">"
	Response.Write ClassName
	Response.Write "</a>] "
	Response.Write enchiasp.ReadBriefTopic(BriefTopic)
	Response.Write "<u><a href=?action=view&ChannelID=" & ChannelID & "&ArticleID="
	Response.Write  ArticleID
	Response.Write ">"
	Response.Write enchiasp.ReadFontMode(title,ColorMode,FontMode)
	Response.Write "</a></u>" 

	If isBest <> 0 Then
		Response.Write "&nbsp;&nbsp;[<font color=blue>荐</font>]"
	End If

	Response.Write "	  </td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &"><a href=?action=edit&ChannelID="& ChannelID &"&ArticleID="& ArticleID &">编辑</a> | <a href=?action=del&ChannelID="& ChannelID &"&ArticleID="& ArticleID &" onclick=""{if(confirm('"& sModuleName &"删除后将不能恢复，您确定要删除该"& sModuleName &"吗?')){return true;}return false;}"">删除</a></td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &">"

	If isAccept = 1 Then
		Response.Write "<font color=blue>√</font>"
	Else
		Response.Write "<font color=red>×</font>"
	End If

	Response.Write "	  </td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &">"& UserName &"</td>"
	Response.Write "	  <td align=""center"" nowrap "& strClass &">"

	If WriteTime >= Date Then
		Response.Write "<font color=red>"
		Response.Write enchiasp.FormatDate(WriteTime, 2)
		Response.Write "</font>"
	Else
		Response.Write enchiasp.FormatDate(WriteTime, 2)
	End If

	Response.Write "	  </td>"
	Response.Write "	</tr>"
End Function

Private Sub ArticleEdit(isEdit)
	Dim EditTitle,TitleColor
	If isEdit Then
		SQL = "select * from ECCMS_Article where ArticleID=" & Request("ArticleID")
		Set Rs = enchiasp.Execute(SQL)
		If Not ChkAdmin("AdminArticle" & ChannelID) Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
		ClassID = Rs("ClassID")
		EditTitle = "编辑" & sModuleName & ""
	Else
		EditTitle = "添加" & sModuleName & ""
		ClassID = Request("ClassID")
		If Not ChkAdmin("AddArticle" & ChannelID) Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
	End If
%>

<script src='include/ArticleJuge.Js' type=text/javascript></script>
<script language= JavaScript>
function SelectPhoto(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=<%=ChannelID%>&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.ImageUrl.value=ss[0];
    //document.myform.ImageFileList.value=ss[0];
  }
}
</script>
<div onkeydown=CtrlEnter()>
<table  border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
        <tr>
          <th colspan="4"><%=EditTitle%></th>
        </tr>
		<form method=Post name="myform" action="admin_article.asp">
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""ArticleID"" value="""& Request("ArticleID") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
%>
        <tr>
          <td width="15%" align="right" nowrap class="TableRow2"><b>所属分类：</b></td>
          <td width="30%" class="TableRow1">
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
          <td width="40%" class="TableRow1"><select name="SpecialID" id="SpecialID">
            <option value="0">不指定<%=sModuleName%>专题</option>
<%
	Set RsObj = enchiasp.Execute("SELECT SpecialID,SpecialName From ECCMS_Special WHERE ChannelID = "& ChannelID &" And ChangeLink=0 ORDER BY orders")
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
          <td align="right" class="TableRow2"><b><%=sModuleName%>标题：</b></td>
          <td colspan="3" class="TableRow1"><select name="BriefTopic" id="BriefTopic">
			  <%If isEdit Then%>
			<option value="0"<%If Rs("BriefTopic") = 0 Then Response.Write (" selected")%>>选择话题</option>
			<option value="1"<%If Rs("BriefTopic") = 1 Then Response.Write (" selected")%>>[图文]</option>
			<option value="2"<%If Rs("BriefTopic") = 2 Then Response.Write (" selected")%>>[组图]</option>
			<option value="3"<%If Rs("BriefTopic") = 3 Then Response.Write (" selected")%>>[新闻]</option>
			<option value="4"<%If Rs("BriefTopic") = 4 Then Response.Write (" selected")%>>[推荐]</option>
			<option value="5"<%If Rs("BriefTopic") = 5 Then Response.Write (" selected")%>>[注意]</option>
			<option value="6"<%If Rs("BriefTopic") = 6 Then Response.Write (" selected")%>>[转载]</option>
			<%Else%>
            <option value="0">选择话题</option>
			<option value="1">[图文]</option>
			<option value="2">[组图]</option>
			<option value="3">[新闻]</option>
			<option value="4">[推荐]</option>
			<option value="5">[注意]</option>
			<option value="6">[转载]</option>
			<%End If%>
          </select> <input name="title" type="text" id="title" size="60" value="<%If isEdit Then Response.Write Rs("title")%>"> 
          <font color=red>*</font></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>标题模式：</b></td>
          <td colspan="3" class="TableRow1">颜色：
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
          <td align="right" class="TableRow2"><b><%=sModuleName%>作者：</b></td>
          <td colspan="3" class="TableRow1"><input name="Author" type="text" size="30" value="<%If isEdit Then Response.Write Rs("Author")%>">
		    <select name=font2 onChange="Author.value=this.value;">
			<option selected value="">选择作者</option>
			<option value='佚名'>佚名</option>
			<option value='本站'>本站</option>
			<option value='不详'>不详</option>
			<option value='未知'>未知</option>
			<option value='<%=AdminName%>'><%=AdminName%></option>
			</select></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>来源：</b></td>
          <td colspan="3" class="TableRow1"><input name="ComeFrom" type="text" size="30" value="<%If isEdit Then Response.Write Rs("ComeFrom")%>">
		  	<select name=font1 onChange="ComeFrom.value=this.value;">
			<option selected value="">选择来源</option>
			<option value='本站原创'>本站原创</option>
			<option value='本站整理'>本站整理</option>
			<option value='不详'>不详</option>
			<option value='转载'>转载</option>
			<option value='<%=Replace(enchiasp.SiteUrl, "http://", "",1,-1,1)%>'><%=Replace(enchiasp.SiteUrl, "http://", "",1,-1,1)%></option>
			</select></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b><%=sModuleName%>内容：</b><br><%=sModuleName%>内容分页标签<br><font color=red>[page_break]</font><br>请注意标签字母小写</td>
          <td colspan="3" class="TableRow1"><textarea name="content" id='content' style="display:none"><%If isEdit Then Response.Write Server.HTMLEncode(Rs("content"))%></textarea>
			<iframe ID='HtmlEditor1' src='../editor/editor.asp?ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='560' HEIGHT='350'></iframe></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>首页图片：</b></td>
          <td colspan="3" class="TableRow1"><input name="ImageUrl" type="text" id="ImageUrl" size="60" value="<%If isEdit Then Response.Write Rs("ImageUrl")%>">
			<input type=hidden name=UploadFileList id=UploadFileList onchange="doChange(this,document.myform.ImageFileList)" value="<%If isEdit Then Response.Write Rs("UploadImage")%>">
			<br>直接从上传图片中选择：
			<%
			If IsEdit Then
			Response.Write InitSelect(Rs("UploadImage"),Rs("ImageUrl"))
			Else
			%>
			<select name="ImageFileList" id="ImageFileList" onChange="ImageUrl.value=this.value;"><option value=''>不选择首页推荐图片</option></select><%End If%>
			<input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto()' class=button></td>
        </tr>
        <tr>
          <td align="right" class="TableRow2"><b>浏览等级：</b></td>
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
          <td align="right" class="TableRow2"><b>其它选项：</b></td>
          <td class="TableRow1" colspan="3"><input name="isTop" type="checkbox" id="isTop" value="1"<%If isEdit Then:If Rs("isTop") <> 0 Then Response.Write (" checked")%>>
            <%=sModuleName%>置顶 
            <input name="isBest" type="checkbox" id="isBest" value="1"<%If isEdit Then:If Rs("isBest") <> 0 Then Response.Write (" checked")%>> 
            <%=sModuleName%>推荐
	    <input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"<%If isEdit Then:If Rs("ForbidEssay") <> 0 Then Response.Write (" checked")%>> 
            禁止发表评论
            <%If  ChkAdmin("AdminArticle" & ChannelID) Then

            %>
	    <input name="isAccept" type="checkbox" id="isAccept" value="1" checked> 
            立即发布（<font color=blue>否则审核后才能发布。</font>）
            <%
            else
            %>
             <input name="isAccept" type="checkbox" id="isAccept" value="0" > 

            <%
            end if
            %>
	    <%If isEdit Then%>
	    <input name="Update" type="checkbox" value="1"> 
            同时更新<%=sModuleName%>发布时间<%End If%></td>
        </tr>
        <tr align="center">
          <td colspan="4" class="TableRow2">
	  <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>
	  <input type="button" name="btnSubmit" value="保存<%=sModuleName%>" class=Button onclick="doSubmit();"></td>
        </tr></form>
      </table></div>
<%
	If isEdit Then
		Rs.Close
		Set Rs = Nothing
	End If
End Sub

Private Sub CheckSave()
	If Trim(Request.Form("title")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "标题不能为空！</li>"
	End If
	If Len(Request.Form("title")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "标题不能超过200个字符！</li>"
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
	If Trim(Request.Form("Author")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "作者不能为空！</li>"
	End If
	If Trim(Request.Form("ComeFrom")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "来源不能为空！</li>"
	End If
	If Trim(Request.Form("PointNum")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>阅览" & sModuleName & "所需的点数不能为空！如果不想设置请输入零。</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "星级不能为空。</li>"
	End If
	If Not IsNumeric(Request.Form("UserGroup")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>阅览" & sModuleName & "等级参数错误！</li>"
	End If
	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该一级分类已经有下属分类，不能添加" & sModuleName & "！</li>"
	End If
	If Trim(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该分类是外部连接，不能添加" & sModuleName & "！</li>"
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
		ErrMsg = ErrMsg + "<li>" & sModuleName & "内容不能为空！</li>"
	End If
	If CInt(Request.Form("isTop")) = 1 Then
		ArticleTop = 1
	Else
		ArticleTop = 0
	End If
	If CInt(Request.Form("isBest")) = 1 Then
		ArticleBest = 1
	Else
		ArticleBest = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If
	If CInt(Request("isAccept")) = 1 Then
		ArticleAccept = 1
	Else
		ArticleAccept = 0
	End If
End Sub

Private Sub SaveArticle()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Article where (ArticleID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("SpecialID") = Trim(Request.Form("SpecialID"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("ColorMode") = Trim(Request.Form("ColorMode"))
		Rs("FontMode") = Trim(Request.Form("FontMode"))
		'字符过率
		'Rs("content") = enchiasp.HTMLEncode(Html2Ubb(TextContent))
		Rs("content") = Html2Ubb(TextContent)
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Author") = Trim(Request.Form("Author"))
		Rs("ComeFrom") = Trim(Request.Form("ComeFrom"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("isTop") = ArticleTop
		Rs("AllHits") = CLng(Request.Form("AllHits"))
		Rs("DayHits") = 0
		Rs("WeekHits") = 0
		Rs("MonthHits") = 0
		Rs("HitsTime") = Now()
		Rs("WriteTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("username") = Trim(AdminName)
		Rs("isBest") = ArticleBest
		Rs("BriefTopic") = Trim(Request.Form("BriefTopic"))
		Rs("ImageUrl") = Trim(Request.Form("ImageUrl"))
		Rs("UploadImage") = Trim(Request.Form("UploadFileList"))
		Rs("UserGroup") = Trim(Request.Form("UserGroup"))
		Rs("PointNum") = Trim(Request.Form("PointNum"))
		Rs("isUpdate") = 1
		Rs("isAccept") = ArticleAccept
		Rs("ForbidEssay") = ForbidEssay
		Rs("AlphaIndex") = enchiasp.ReadAlpha(Request.Form("title"))
	Rs.update
	Rs.Close
	Rs.Open "select top 1 ArticleID from ECCMS_Article where ChannelID=" & ChannelID & " order by ArticleID desc", Conn, 1, 1
	ArticleID = Rs("ArticleID")
	Rs.Close:Set Rs = Nothing
	ClassUpdateCount Request.Form("ClassID"),1
	Call RemoveCache
	Dim url
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makenews.asp?ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&showid=0"	
		Call ScriptCreation(url,ArticleID)
		SQL = "SELECT TOP 1 ArticleID FROM ECCMS_Article WHERE ChannelID=" & ChannelID & " And isAccept <> 0 And ArticleID < " & ArticleID & " ORDER BY ArticleID DESC"
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.EOF And Rs.BOF) Then
			url = "admin_makenews.asp?ChannelID=" & ChannelID & "&ArticleID=" & Rs("ArticleID") & "&showid=0"	
			Call ScriptCreation(url,Rs("ArticleID"))
		End If
		Rs.Close
		Set Rs = Nothing
	End If
	'Succeed("<li>恭喜您！添加新的" & sModuleName & "成功。</li><li><a href=admin_article.asp?action=view&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & ">点击此处查看该" & sModuleName & "</a></li><li><a href=?action=add&ChannelID=" & ChannelID & "&classid=" & Request.Form("ClassID") & "><font color=blue>点击此处继续添加" & sModuleName & "</font></a></li>")
Succeed("<li>恭喜您！添加新的" & sModuleName & "成功。</li>")
End Sub

Private Sub ModifyArticle()
	If Not ChkAdmin("AdminArticle" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	CheckSave
	If Founderr = True Then Exit Sub
	Dim Auditing
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Article where ArticleID=" & Request("ArticleID")
	Rs.Open SQL,Conn,1,3
		Auditing = Rs("isAccept")
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("SpecialID") = Trim(Request.Form("SpecialID"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("ColorMode") = Trim(Request.Form("ColorMode"))
		Rs("FontMode") = Trim(Request.Form("FontMode"))
		Rs("content") = Html2Ubb(TextContent)
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Author") = Trim(Request.Form("Author"))
		Rs("ComeFrom") = Trim(Request.Form("ComeFrom"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("isTop") = ArticleTop
		Rs("isBest") = ArticleBest
		If CInt(Request.Form("Update")) = 1 Then Rs("WriteTime") = Now()
		Rs("AllHits") = CLng(Request.Form("AllHits"))
		Rs("BriefTopic") = Trim(Request.Form("BriefTopic"))
		Rs("ImageUrl") = Trim(Request.Form("ImageUrl"))
		Rs("UploadImage") = Trim(Request.Form("UploadFileList"))
		Rs("UserGroup") = Trim(Request.Form("UserGroup"))
		Rs("PointNum") = Trim(Request.Form("PointNum"))
		Rs("isUpdate") = 1
		Rs("isAccept") = ArticleAccept
		Rs("ForbidEssay") = ForbidEssay
		Rs("AlphaIndex") = enchiasp.ReadAlpha(Request.Form("title"))
	Rs.update
	ArticleID = Rs("ArticleID")
	If ArticleAccept = 1 And Auditing = 0 Then
		AddUserPointNum Rs("username"),1
	End If
	If ArticleAccept = 0 And Auditing = 1 Then
		AddUserPointNum Rs("username"),0
	End If
	Rs.Close:Set Rs = Nothing
	Call RemoveCache
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makenews.asp?ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&showid=0"	
		Call ScriptCreation(url,ArticleID)
	End If
	Succeed("<li>恭喜您！修改" & sModuleName & "成功。</li><li><a href=admin_article.asp?action=view&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & ">点击此处查看该" & sModuleName & "</a></li>")
End Sub
Private Sub ArticleView()
	Call PageTop
	If Request("ArticleID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	If ChannelID = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请指定频道。</li>"
		Exit Sub
	End If
	SQL = "select ArticleID,title,content,ColorMode,FontMode,Author,ComeFrom,WriteTime,username,isAccept from ECCMS_Article where ChannelID=" & ChannelID & " And ArticleID=" & Request("ArticleID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！没有找到任何" & sModuleName & "。或者您选择了错误的系统参数！</li>"
		Exit Sub
	Else
%>
<script language=javascript>
var enchiasp_fontsize=9;
var enchiasp_lineheight=12;
</script>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
	<tr>
	  <th>查看<%=sModuleName%></th>
	</tr>
	<tr>
	  <td align="center" class="TableRow2"><a href=?action=edit&ChannelID=<%=ChannelID%>&ArticleID=<%=Rs("ArticleID")%>><font size=4><%=enchiasp.ReadFontMode(Rs("title"),Rs("ColorMode"),Rs("FontMode"))%></font></a></td>
	</tr>
	<tr>
	  <td align="center" class="TableRow1"><b>作者：</b><%=Rs("Author")%>&nbsp;&nbsp;<b>来源：</b><%=Rs("ComeFrom")%>&nbsp;&nbsp;<b>发布时间：</b><%=Rs("WriteTime")%>&nbsp;&nbsp;
	  <b>发 布 人：</b> <font color=blue><%=Rs("username")%></font>&nbsp;&nbsp;
	  <b>审核状态：</b>
	  <%If Rs("isAccept") > 0 Then%>
	  <font color=blue>已审核</font>
	  <%Else%>
	  <font color=red>未审核</font>
	  <%End If%>
	  </td>
	</tr>
	<tr>
	  <td class="TableRow1"><p align="right"><a style="CURSOR: hand; POSITION: relative" onclick="if(enchiasp_fontsize&gt;8){enchiaspContentLabel.style.fontSize=(--enchiasp_fontsize)+&quot;pt&quot;;enchiaspContentLabel.style.lineHeight=(--enchiasp_lineheight)+&quot;pt&quot;;}" title="减小字体"><img src="images/1.gif" border="0" width="15" height="15"><font color="#FF6600">减小字体</font></a> 
                    <a style="CURSOR: hand; POSITION: relative" onclick="if(enchiasp_fontsize&lt;64){enchiaspContentLabel.style.fontSize=(++enchiasp_fontsize)+&quot;pt&quot;;enchiaspContentLabel.style.lineHeight=(++enchiasp_lineheight)+&quot;pt&quot;;}" title="增大字体"><img src="images/2.gif" border="0" width="15" height="15"><font color="#FF6600">增大字体</font></a></p>
					<div id="enchiaspContentLabel"><%=Replace(enchiasp.ReadContent(Rs("content")), "[page_break]", "", 1, -1, 1)%></div></td>
	</tr>
	<tr>
	  <td class="TableRow1">上一篇<%=sModuleName%>：<%=FrontArticle(Rs("ArticleID"))%>
	  <br>下一篇<%=sModuleName%>：<%=NextArticle(Rs("ArticleID"))%></td>
	</tr>
	<tr>
	  <td align="center" class="TableRow2">
	  <input type="button" onclick="{if(confirm('您确定要删除此文章吗?')){location.href='?action=del&ChannelID=<%=ChannelID%>&ArticleID=<%=Rs("ArticleID")%>';return true;}return false;}" value="删除文章" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="返回上一页" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="返回顶部" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&ChannelID=<%=ChannelID%>&ArticleID=<%=Rs("ArticleID")%>'" value="编辑<%=sModuleName%>" class=button>&nbsp;&nbsp;
	   [<a href="?act=批量审核&ChannelID=<%=ChannelID%>&selArticleID=<%=Rs("ArticleID")%>" onclick="return confirm('您确定执行审核操作吗?');" ><font color=blue>直接审核</font></a>]</td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub
Private Function FrontArticle(ArticleID)
	Dim Rss, SQL
	SQL = "select Top 1 ArticleID,classid,title from ECCMS_Article where ChannelID=" & ChannelID & " And isAccept <> 0 And ArticleID < " & ArticleID & " order by ArticleID desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		FrontArticle = "已经没有了"
	Else
		FrontArticle = "<a href=admin_article.asp?action=view&ChannelID=" & ChannelID & "&ArticleID=" & Rss("ArticleID") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function NextArticle(ArticleID)
	Dim Rss, SQL
	SQL = "select Top 1 ArticleID,classid,title from ECCMS_Article where ChannelID=" & ChannelID & " And isAccept <> 0 And ArticleID > " & ArticleID & " order by ArticleID asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		NextArticle = "已经没有了"
	Else
		NextArticle = "<a href=admin_article.asp?action=view&ChannelID=" & ChannelID & "&ArticleID=" & Rss("ArticleID") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function

Private Sub BatCreateHtml()
	Dim AllArticleID,url
	Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
	Response.Write "<ol>"
	AllArticleID = Split(selArticleID, ",")
	For i = 0 To UBound(AllArticleID)
		ArticleID = CLng(AllArticleID(i))
		url = "admin_makenews.asp?ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&showid=1"	
		Call ScriptCreation(url,ArticleID)
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

Private Sub ArticleDel()
	If Request("ArticleID") = "" Then
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT ArticleID,classid,username,HtmlFileDate FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ArticleID=" & Request("ArticleID"))
	If Not(Rs.BOF And Rs.EOF) Then
		ClassUpdateCount Rs("classid"),0
		AddUserPointNum Rs("username"),0
		DeleteHtmlFile Rs("classid"),Rs("ArticleID"),Rs("HtmlFileDate")
	End If
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute("Delete From ECCMS_Article Where ChannelID = "& ChannelID &" And ArticleID = " & Request("ArticleID"))
	enchiasp.Execute ("delete from ECCMS_Comment where ChannelID = "& ChannelID &" And PostID = " & Request("ArticleID"))
	Call RemoveCache
	Response.redirect ("admin_article.asp?ChannelID=" & ChannelID)
End Sub

Private Sub batdel()
	Set Rs = enchiasp.Execute("SELECT ArticleID,classid,username,HtmlFileDate FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ArticleID in (" & selArticleID & ")")
	Do While Not Rs.EOF
		ClassUpdateCount Rs("classid"),0
		AddUserPointNum Rs("username"),0
		DeleteHtmlFile Rs("classid"),Rs("ArticleID"),Rs("HtmlFileDate")
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("delete from ECCMS_Article where ArticleID in (" & selArticleID & ")")
	enchiasp.Execute ("delete from ECCMS_Comment where ChannelID = "& ChannelID &" And PostID in (" & selArticleID & ")")
	Call RemoveCache
	OutHintScript ("批量删除操作成功！")
End Sub

Private Sub isCommend()
	enchiasp.Execute ("update ECCMS_Article set isBest=1 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noCommend()
	enchiasp.Execute ("update ECCMS_Article set isBest=0 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub isTop()
	enchiasp.Execute ("update ECCMS_Article set isTop=1 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noTop()
	enchiasp.Execute ("update ECCMS_Article set isTop=0 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub BatAccept()
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ArticleID in (" & selArticleID & ")")
	Do While Not Rs.EOF
		AddUserPointNum Rs("username"),1
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("update ECCMS_Article set isAccept=1 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Sub NotAccept()
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ArticleID in (" & selArticleID & ")")
	Do While Not Rs.EOF
		AddUserPointNum Rs("username"),0
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute ("update ECCMS_Article set isAccept=0 where ArticleID in (" & selArticleID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Function AddUserPointNum(username,stype)
	On Error Resume Next
	Dim rsuser,GroupSetting,userpoint
	Set rsuser = enchiasp.Execute("SELECT userid,UserGrade,userpoint FROM ECCMS_User WHERE username='"& username &"'")
	If Not(rsuser.BOF And rsuser.EOF) Then
		GroupSetting = Split(enchiasp.UserGroupSetting(rsuser("UserGrade")), "|||")(9)
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

Function InitSelect(UploadFileList, ImageUrl)
	Dim i
	InitSelect = "<select name='ImageFileList' onChange=""ImageUrl.value=this.value;"">"
	InitSelect = InitSelect & "<option value=''>不选择首页推荐图片</option>"
	If Not IsNull(UploadFileList) Then
		UploadFileList = Split(UploadFileList, "|")
		For i = 0 To UBound(UploadFileList)
			If UploadFileList(i) <> "" Then
				InitSelect = InitSelect & "<option value=""" & UploadFileList(i) & """"
				If UploadFileList(i) = ImageUrl Then
					InitSelect = InitSelect & " selected"
				End If
				InitSelect = InitSelect & ">" & UploadFileList(i) & "</option>"
			End If
		Next
	End If
	InitSelect = InitSelect & "</select>"
End Function

Private Sub BatchSetting()
	If Not ChkAdmin("AdminArticle" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=4>" & sModuleName & "批量设置</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=saveset>"
	Response.Write "	<input type=hidden name=ChannelID value='"& ChannelID &"'>"
	Response.Write "	<tr>"
	Response.Write "		<td width=""20%"" rowspan=""14"" class=tablerow2 valign=""top"" id=choose2 style=""display:none""><b>请选择" & sModuleName & "分类</b><br>"
	Response.Write "<select name=""ClassID"" size='2' multiple style='height:330px;width:180px;'>"
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
	Response.Write "		<td class=tablerow1><input type=""text"" name=""ArticleID"" size=70 value='"& Request("selArticleID") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 width=""15%"" align=right><b>相关" & sModuleName & "：</b></td>"
	Response.Write "		<td class=tablerow1 width=""5%"" align=center><input type=checkbox name=selRelated value='1'></td>"
	Response.Write "		<td class=tablerow1 width=""60%""><input name=Related type=text size=60></td>"
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
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "来源：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selComeFrom value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=ComeFrom type=text size=20>"
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
	If Not ChkAdmin("AdminArticle" & ChannelID) Then
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
	Response.Write " <b>指定" & sModuleName & "ID：</b> <input type=""text"" name=""ArticleID"" size=80 value='"& Request("selArticleID") &"'></td>"
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
		If Trim(Request.Form("ArticleID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_Article SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ArticleID in ("& Request("ArticleID") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_Article SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>恭喜您！批量移动完成。</li>")
End Sub

Private Sub BatcDelete()
	If Not ChkAdmin("AdminArticle" & ChannelID) Then
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
	Response.Write "<input type=radio name=Appointed value='1' onClick=""choose2.style.display='';choose1.style.display='none';""> <b>指定" & sModuleName & "分类</b>"
	Response.Write "<input type=radio name=Appointed value='2'> <b>删除所有" & sModuleName & "</b>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr id=choose1>"
	Response.Write "		<td class=tablerow1><b>分类ID：</b><input type=""text"" name=""ArticleID"" size=80 value='"& Request("selArticleID") &"'></td>"
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
Private Sub AllDelArticle()
	On Error Resume Next
	If CInt(Request.Form("Appointed")) = 1 Then
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		End If
		Conn.Execute("DELETE ECCMS_Comment FROM ECCMS_Article A INNER JOIN ECCMS_Comment C ON C.PostID=A.ArticleID WHERE A.ChannelID = "& ChannelID &" And A.ClassID IN (" & Request("ClassID") & ")")
		Conn.Execute("DELETE FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ClassID IN (" & Request("ClassID") & ")")
	ElseIf CInt(Request.Form("Appointed")) = 2 Then
		Conn.Execute ("DELETE FROM ECCMS_Article WHERE ChannelID = "& ChannelID)
		Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID)
	Else
		If Trim(Request.Form("ArticleID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		End If
		Conn.Execute ("DELETE FROM ECCMS_Article WHERE ChannelID = "& ChannelID &" And ArticleID IN (" & Request("ArticleID") & ")")
		Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID IN (" & Request("ArticleID") & ")")
		
	End If
	Call RemoveCache
	Succeed("<li>批量删除成功！</li>")
End Sub

Private Sub SaveSetting()
	If Founderr = True Then Exit Sub
	Dim strTempValue
	strTempValue = ""
	If Trim(Request.Form("selRelated")) <> "" Then strTempValue = strTempValue & "Related='"& enchiasp.ChkFormStr(Request.Form("Related")) &"',"
	If Trim(Request.Form("selAuthor")) <> "" Then strTempValue = strTempValue & "Author='"& enchiasp.ChkFormStr(Request.Form("Author")) &"',"
	If Trim(Request.Form("selComeFrom")) <> "" Then strTempValue = strTempValue & "ComeFrom='"& enchiasp.ChkFormStr(Request.Form("ComeFrom")) &"',"
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
		If Trim(Request.Form("ArticleID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_Article SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ArticleID in ("& Request("ArticleID") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		ElseIf Trim(Request.Form("ClassID")) = "-1" Then
			SQL = "UPDATE ECCMS_Article SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID
		Else
			SQL = "UPDATE ECCMS_Article SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
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
	Dim totalrec,WriteTime,page_count,pagelist
	i = 0
	page_count = 0
	totalrec = enchiasp.Execute("SELECT COUNT(ArticleID) FROM [ECCMS_Article] WHERE ChannelID = "& ChannelID &" And isAccept>0")(0)
	Set Rs = enchiasp.Execute("SELECT ArticleID,WriteTime FROM [ECCMS_Article] WHERE ChannelID = "& ChannelID &" And isAccept>0 ORDER BY WriteTime DESC")
	If Not (Rs.BOF And Rs.EOF) Then
		SQL=Rs.GetRows(-1)
		For pagelist=0 To Ubound(SQL,2)
			If Not Response.IsClientConnected Then Response.End
			Response.Write "<script>"
			Response.Write "table2.style.width=" & Fix((page_count / totalrec) * 400) & ";"
			Response.Write "txt2.innerHTML=""完成：" & FormatNumber(page_count / totalrec * 100, 2, -1) & """;"
			Response.Write "</script>" & vbCrLf
			Response.Flush
			WriteTime = DateAdd("s", -i, SQL(1,pagelist))
			enchiasp.Execute ("UPDATE ECCMS_Article SET WriteTime='" & WriteTime & "' WHERE ArticleID="& SQL(0,pagelist))
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

