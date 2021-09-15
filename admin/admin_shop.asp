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
Dim Action,isEdit
Dim i,ClassID,RsObj,shopid,findword,keyword,strClass
Dim TextContent,ShopTop,ShopBest,ForbidEssay
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
Dim s_ClassName,ChildStr,FoundSQL,isAccept,selShopID

If Request("isAccept") <> "" Then
	isAccept = 0
Else
	isAccept = 1
End If
If CInt(ChannelID) = 0 Then ChannelID = 3
Action = LCase(Request("action"))

Select Case Trim(Action)
Case "save"
	Call SaveShop
Case "modify"
	Call ModifyShop
Case "add"
	isEdit = False
	Call ShopEdit(isEdit)
Case "edit"
	isEdit = True
	Call ShopEdit(isEdit)
Case "del"
	Call ShopDel
Case "view"
	Call ShopView
Case "setting"
	Call BatchSetting
Case "saveset"
	Call SaveSetting
Case "move"
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
	Response.Write "	  <th colspan=2>" & sModuleName & "管理选项</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr><form method=Post name=myform action='admin_Shop.asp' onSubmit='return JugeQuery(this);'>"
	Response.Write "	<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "	<td class=TableRow1>搜索："
	Response.Write "	  <input name=keyword type=text size=20>"
	Response.Write "	  条件："
	Response.Write "	  <select name=queryopt>"
	Response.Write "		<option value='1' selected>" & sModuleName & "名称</option>"
	Response.Write "		<option value='2'>" & sModuleName & "规格</option>"
	Response.Write "		<option value='3'>不限条件</option>"
	Response.Write "	  </select> <input type=submit name=Submit value='开始查询' class=Button></td>"
	Response.Write "	  <td class=TableRow1>" & sModuleName & "导航："
	Dim srtClassMenu
	Response.Write "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Response.Write "<option value='admin_Shop.asp?ChannelID=" & ChannelID & "'>≡全部" & sModuleName & "列表≡</option>" & vbCrLf
	srtClassMenu = enchiasp.ClassJumpMenu(ChannelID)
	srtClassMenu = Replace(srtClassMenu, "{ClassID=" & Request("sortid") & "}", "selected")
	Response.Write srtClassMenu
	Response.Write "</select>"
	Response.Write "	  </td></form>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td colspan=2 class=TableRow2><strong>操作选项：</strong> <a href='admin_Shop.asp?ChannelID=" & ChannelID & "'>管理首页</a> | "
	Response.Write "	  <a href='admin_Shop.asp?ChannelID=" & ChannelID & "&action=add'>添加" & sModuleName & "</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "&action=add'>添加" & sModuleName & "分类</a> | "
	Response.Write "	  <a href='admin_classify.asp?ChannelID=" & ChannelID & "'>" & sModuleName & "分类管理</a></td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub
Private Sub showmain()
	If Not ChkAdmin("AdminShop" & ChannelID) Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	Dim strListName
	If Not IsEmpty(Request("selShopID")) Then
		selShopID = Request("selShopID")
		Select Case enchiasp.CheckStr(Request("act"))
		Case "批量删除"
			Call batdel
		Case "批量推荐"
			Call isCommend
		Case "取消推荐"
			Call noCommend
		Case "批量置顶"
			Call isTop
		Case "取消置顶"
			Call noTop
		Case "生成HTML"
			Call BatCreateHtml
		Case Else
			Response.Write "无效参数！"
		End Select
	End If
	Call PageTop
	Dim specialID,sortid,Cmd,child
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>选择</th>"
	Response.Write "	  <th width='68%'>" & sModuleName & "名称</th>"
	Response.Write "	  <th width='9%' nowrap>管理操作</th>"
	Response.Write "	  <th width='9%' nowrap>" & sModuleName & "星级</th>"
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
		CurrentPage = CLng(Request("page"))
	Else
		CurrentPage = 1
	End If
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("queryopt")) = 1 Then
			findword = "A.TradeName like '%" & keyword & "%'"
		ElseIf CInt(Request("queryopt")) = 2 Then
			findword = "A.Marque like '%" & keyword & "%'"
		Else
			findword = "A.TradeName like '%" & keyword & "%' or A.Marque like '%" & keyword & "%'"
		End If
		FoundSQL = findword
		s_ClassName = "查询" & sModuleName
	Else
		specialID = 99999
		If Request("sortid") <> "" Then
			FoundSQL = "A.isAccept <> "& isAccept & " And A.ClassID in (" & ChildStr & ")"
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
	TotalNumber = enchiasp.Execute("SELECT COUNT(ShopID) FROM ECCMS_ShopList A WHERE A.ChannelID = " & ChannelID & " And "& FoundSQL &"")(0)
	TotalPageNum = CLng(TotalNumber / maxperpage)  '得到总页数
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT A.*,C.ClassName FROM [ECCMS_ShopList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.ChannelID = " & ChannelID & " And "& FoundSQL &" ORDER BY A.isTop DESC, A.addTime DESC ,A.ShopID DESC"
	
	If IsSqlDataBase = 1 Then
		'SQL是否起用存贮过程
		'If Trim(Request("keyword"))<>"" Or child > 0 Then
			'Set Rs = enchiasp.Execute(SQL)
		'Else
			'Set Cmd = Server.CreateObject("ADODB.Command")
			'Set Cmd.ActiveConnection=conn
			'Cmd.CommandText="ECCMS_ShopAdminList"
			'Cmd.CommandType=4
			'Cmd.Parameters.Append cmd.CreateParameter("@ChannelID",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@sortid",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@specialID",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@isAccept",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@pagenow",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
			'Cmd.Parameters.Append cmd.CreateParameter("@totalrec",3,2)
			'Cmd("@ChannelID")=ChannelID
			'Cmd("@sortid")=sortid
			'Cmd("@specialID")=specialID
			'Cmd("@isAccept")=isAccept
			'Cmd("@pagenow")=CurrentPage
			'Cmd("@pagesize")=maxperpage
			'Set Rs=Cmd.Execute
			
		'End If
		Rs.Open SQL, Conn, 1, 1
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=5 class=TableRow2>还没有找到任何" & sModuleName & "！</td></tr>"
	Else
		If IsSqlDataBase<>1 Or Trim(Request("keyword"))<>"" Or child > 0 Then
			Rs.MoveFirst
			If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
			If Rs.Eof Then Exit Sub
		End If
		i = 0

		Response.Write "	<tr>"
		Response.Write "	  <td colspan=5 class=TableRow2>"
		ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<form name=selform method=post action=""admin_shop.asp"">"
		Response.Write "	<input type=hidden name=ChannelID value='" & ChannelID & "'>"
		Response.Write "	<input type=hidden name=action value=''>"
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			If (i mod 2) = 0 Then
				strClass = "class=TableRow1"
			Else
				strClass = "class=TableRow2"
			End If
			Response.Write "	<tr>"
			Response.Write "	  <td align=center " & strClass & "><input type=checkbox name=selShopID value=" & Rs("ShopID") & "></td>"
			Response.Write "	  <td " & strClass & ">"
			If Rs("isTop") <> 0 Then
				Response.Write "<img src=""images/istop.gif"" width=15 height=17 border=0 alt=置顶商品>"
			End If

			Response.Write "[<a href=?ChannelID=" & Rs("ChannelID") & "&sortid="
			Response.Write Rs("ClassID")
			Response.Write ">"
			Response.Write Rs("ClassName")
			Response.Write "</a>] "
			Response.Write "<a href=?action=view&ChannelID=" & Rs("ChannelID") & "&ShopID="
			Response.Write Rs("ShopID")
			Response.Write ">"
			Response.Write Rs("TradeName")
			Response.Write "</a>" 

			If Rs("isBest") <> 0 Then
				Response.Write "&nbsp;&nbsp;<font color=blue>荐</font>"
			End If
			Response.Write "	  </td>"
			Response.Write "	  <td align=center nowrap " & strClass & "><a href=?action=edit&ChannelID=" & Rs("ChannelID") & "&ShopID=" & Rs("ShopID") & ">编辑</a> | <a href=?action=del&ChannelID=" & Rs("ChannelID") & "&ShopID=" & Rs("ShopID") & " onclick=""{if(confirm('商品删除后将不能恢复，您确定要删除该商品吗?')){return true;}return false;}"">删除</a></td>"
			Response.Write "	  <td align=center nowrap " & strClass & ">"
			Response.Write "<font color=green>"
			For i = 1 to Rs("star")
				Response.Write "★"
			Next
			Response.Write "</font>"

			Response.Write "	  </td>"
			Response.Write "	  <td align=center nowrap " & strClass & ">"

			If Rs("addTime") >= Date Then
				Response.Write "<font color=red>"
				Response.Write enchiasp.FormatDate(Rs("addTime"), 2)
				Response.Write "</font>"
			Else
				Response.Write enchiasp.FormatDate(Rs("addTime"), 2)
			End If

			Response.Write "	  </td>"
			Response.Write "	</tr>"

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
		<option value="生成HTML">生成HTML</option>
	  </select>
	  <input class=Button type="submit" name="Submit2" value="执行操作" onclick="return confirm('您确定执行该操作吗?');">
	  <input class=Button type="submit" name="Submit3" value="批量设置" onclick="document.selform.action.value='setting';">
	  <input class=Button type="submit" name="Submit4" value="批量移动" onclick="document.selform.action.value='move';"></td>
	</tr>
	</form>
	<tr>
	  <td colspan="5" align="right" class="TableRow2"><%ShowListPage CurrentPage,TotalPageNum,totalnumber,maxperpage,strListName,s_ClassName %></td>
	</tr>
</table>
<%
End Sub

Private Sub ShopEdit(isEdit)
	Dim EditTitle,TitleColor
	If isEdit Then
		SQL = "SELECT * FROM ECCMS_ShopList WHERE shopid=" & CLng(Request("shopid"))
		Set Rs = enchiasp.Execute(SQL)
		ClassID = Rs("ClassID")
		EditTitle = "编辑" & sModuleName
	Else
		If Not ChkAdmin("AdminShop" & ChannelID) Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
		EditTitle = "添加" & sModuleName
		ClassID = Request("ClassID")
	End If
%>
<script src='include/ShopJuge.Js' type=text/javascript></script>
<div onkeydown=CtrlEnter()>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
  <tr>
    <th colspan="4"><%=EditTitle%></th>
  </tr>
  	<form method=Post name="myform" action="admin_shop.asp" onSubmit="return CheckForm(this);">
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""shopid"" value="""& Request("shopid") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
%>
  <tr>
    <td width="15%" align="right" class="TableRow2"><strong><%=sModuleName%>分类：</strong></td>
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
    <td width="15%" align="right" class="TableRow2"><strong></strong></td>
    <td width="35%" class="TableRow1"></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>名称：</strong></td>
    <td class="TableRow1"><input name="TradeName" type="text" id="TradeName" size="30" value="<%If isEdit Then Response.Write Rs("TradeName")%>"> 
      <span class="style1">* </span></td>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>规格：</strong></td>
    <td class="TableRow1"><input name="Marque" type="text" id="Marque" size="20" value="<%If isEdit Then Response.Write Rs("Marque")%>"></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>单位：</strong></td>
    <td class="TableRow1"><input name="Unit" type="text" id="Unit2" size="10" value="<%If isEdit Then Response.Write Rs("Unit")%>">
    <select name=font1 onChange="Unit.value=this.value;">
			<option selected value="">请选择单位</OPTION>
			<option value=套>套</option>
			<option value=件>件</option>
			<option value=台>台</option>
			<option value=盒>盒</option>
			<option value=部>部</option>
			<option value=瓶>瓶</option>
			<option value=个>个</option>
			<option value=本>本</option>
			</select></td>
    <td align="right" class="TableRow2"><strong>货源：</strong></td>
    <td class="TableRow1"><input name="supply" type="text" id="supply" size="10" value="<%If isEdit Then Response.Write Rs("supply")%>">
    <select name=font2 onChange="supply.value=this.value;">
			<option value="">请选择</OPTION>
			<option value=有货>有货</option>
			<option value=限量>限量</option>
			<option value=无货>无货</option>
			<option value=特惠>特惠</option>
			<option value=打折>打折</option>
			<option value=特价>特价</option>
			</select></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>市场价：</strong></td>
    <td class="TableRow1"><input name="PastPrice" type="text" id="PastPrice" size="10" value="<%If isEdit Then Response.Write Rs("PastPrice") else response.write 0 end if%>"> 
      <span class="style2">元</span> <span class="style1">* </span> 使用标签{$PastPrice}</td>
    <td align="right" class="TableRow2"><strong>金卡价：</strong></td>
    <td class="TableRow1"><input name="NowPrice" type="text" id="NowPrice" size="10" value="<%If isEdit Then Response.Write Rs("NowPrice") else response.write 0 end if%>"> 
      <span class="style2">元</span> <span class="style1">* </span>使用标签{$NowPrice}</td>
  </tr>
 <tr>
    <td align="right" class="TableRow2"><strong>银卡价：</strong></td>
    <td class="TableRow1"><input name="YinPrice" type="text" id="YinPrice" size="10" value="<%If isEdit Then Response.Write Rs("YinPrice") else response.write 0 end if%>"> 
      <span class="style2">元</span> <span class="style1">* </span> 使用标签{$YinPrice}</td>
    <td align="right" class="TableRow2"><strong>其它价：</strong></td>
    <td class="TableRow1"><input name="OtherPrice" type="text" id="OtherPrice" size="10" value="<%If isEdit Then Response.Write Rs("OtherPrice") else response.write 0 end if%>"> 
      <span class="style2">元</span> <span class="style1">* </span> 使用标签{$OtherPrice}</td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>产品公司：</strong></td>
    <td class="TableRow1"><input name="Company" type="text" id="Company" size="30" value="<%If isEdit Then Response.Write Rs("Company")%>"></td>
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
    <td align="right" class="TableRow2"><strong><%=sModuleName%>图片：</strong></td>
    <td colspan="3" class="TableRow1"><input name="ProductImage" type="text" id="ImageUrl" size="70" value="<%If isEdit Then Response.Write Rs("ProductImage")%>"> 
      <span class="style3">* </span></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>图片上传：</strong></td>
    <td colspan="3" class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?ChannelID=<%=ChannelID%>></iframe></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2"><strong><%=sModuleName%>简介：</strong></td>
    <td colspan="3" class="TableRow1"><textarea name="content" style="display:none" id="content"><%If isEdit Then Response.Write Server.HTMLEncode(Rs("Explain"))%></textarea>
    <iframe ID='HtmlEditor1' src='../editor/editor.asp?ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='560' HEIGHT='350'></iframe></td>
  </tr>
  <tr>
          <td align="right" class="TableRow2"><strong>上传文件：</strong></td>
          <td colspan="3" class="TableRow1"><iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upfiles.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
  <tr>
    <td align="right" class="TableRow2"><strong>其它选项：</strong></td>
    <td colspan="3" class="TableRow1"><input name="isTop" type="checkbox" id="isTop" value="1"<%If isEdit Then:If Rs("isTop") <> 0 Then Response.Write (" checked")%>>
            <%=sModuleName%>置顶 
            <input name="isBest" type="checkbox" id="isBest" value="1"<%If isEdit Then:If Rs("isBest") <> 0 Then Response.Write (" checked")%>> 
            <%=sModuleName%>推荐
	    <input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"<%If isEdit Then:If Rs("ForbidEssay") <> 0 Then Response.Write (" checked")%>> 
            禁止发表评论
	    <%If isEdit Then%>
	    <input name="Update" type="checkbox" value="1"> 
            同时更新上架时间<%End If%></td>
  </tr>
  <tr>
    <td align="right" class="TableRow2">　</td>
    <td colspan="3" align="center" class="TableRow1">
    <input type="button" name="Submit2" onclick="CheckLength();" value="查看内容长度" class=Button>
    <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>
    <input type="submit" name="Submit1" value="保存<%=sModuleName%>" class=Button></td>
  </tr></form>
</table></div>
<%
	If isEdit Then
		Rs.Close
		Set Rs = Nothing
	End If
End Sub

Private Sub CheckSave()
	If Trim(Request.Form("TradeName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "名称不能为空！</li>"
	End If
	If Len(Request.Form("TradeName")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "名称不能超过200个字符！</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "星级不能为空。</li>"
	End If

	If CLng(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该分类是外部连接，不能添加" & sModuleName & "！</li>"
	End If
	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该一级分类已经有下属分类，不能添加" & sModuleName & "！</li>"
	End If
	If Trim(Request.Form("content")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>" & sModuleName & "简介不能为空！</li>"
	End If
	TextContent = ""
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next
	If CInt(Request.Form("isTop")) = 1 Then
		ShopTop = 1
	Else
		ShopTop = 0
	End If
	If CInt(Request.Form("isBest")) = 1 Then
		ShopBest = 1
	Else
		ShopBest = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If

End Sub

Private Sub SaveShop()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_ShopList where (shopid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("TradeName") = enchiasp.ChkFormStr(Request.Form("TradeName"))
		Rs("Marque") = enchiasp.ChkFormStr(Request.Form("Marque"))
		'Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Explain") = TextContent
		Rs("Unit") = enchiasp.ChkFormStr(Request.Form("Unit"))
		Rs("supply") = enchiasp.ChkFormStr(Request.Form("supply"))
		Rs("Company") = Trim(Request.Form("Company"))
		Rs("PastPrice") = Trim(Request.Form("PastPrice"))
		Rs("yinPrice") = Trim(Request.Form("yinPrice"))
		Rs("otherPrice") = Trim(Request.Form("otherPrice"))
		Rs("NowPrice") = Trim(Request.Form("NowPrice"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("ProductImage") = Trim(Request.Form("ProductImage"))
		Rs("addTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("BuyCount") = 0
		Rs("IsBest") = ShopBest
		Rs("IsTop") = ShopTop
		Rs("isAccept") = 1
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	Rs.Close
	Rs.Open "select top 1 shopid from ECCMS_ShopList where ChannelID=" & ChannelID & " order by shopid desc", Conn, 1, 1
	shopid = Rs("shopid")
	Rs.Close:Set Rs = Nothing
	ClassUpdateCount Request.Form("ClassID"),1
	Call RemoveCache
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makeshop.asp?ChannelID=" & ChannelID & "&shopid=" & shopid & "&showid=0"	
		Call ScriptCreation(url,shopid)
		SQL = "SELECT TOP 1 shopid FROM ECCMS_ShopList WHERE ChannelID=" & ChannelID & " And isAccept <> 0 And shopid < " & shopid & " ORDER BY shopid DESC"
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.EOF And Rs.BOF) Then
			url = "admin_makeshop.asp?ChannelID=" & ChannelID & "&shopid=" & Rs("shopid") & "&showid=0"	
			Call ScriptCreation(url,Rs("shopid"))
		End If
		Rs.Close
		Set Rs = Nothing
	End If
	Succeed("<li>恭喜您！添加新的" & sModuleName & "成功。</li><li><a href=?action=view&ChannelID=" & ChannelID & "&shopid=" & shopid & ">点击此处查看该" & sModuleName & "</a></li>")
End Sub

Private Sub ModifyShop()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_ShopList where shopid=" & Request("shopid")
	Rs.Open SQL,Conn,1,3
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = Trim(Request.Form("ClassID"))
		Rs("TradeName") = enchiasp.ChkFormStr(Request.Form("TradeName"))
		Rs("Marque") = enchiasp.ChkFormStr(Request.Form("Marque"))
		'Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Explain") = TextContent
		Rs("Unit") = enchiasp.ChkFormStr(Request.Form("Unit"))
		Rs("supply") = enchiasp.ChkFormStr(Request.Form("supply"))
		Rs("Company") = Trim(Request.Form("Company"))
		Rs("PastPrice") = Trim(Request.Form("PastPrice"))
		Rs("NowPrice") = Trim(Request.Form("NowPrice"))
		Rs("yinPrice") = Trim(Request.Form("yinPrice"))
		Rs("otherPrice") = Trim(Request.Form("otherPrice"))
		Rs("star") = Trim(Request.Form("star"))
		Rs("ProductImage") = Trim(Request.Form("ProductImage"))
		If CInt(Request.Form("Update")) = 1 Then Rs("addTime") = Now()
		Rs("IsBest") = ShopBest
		Rs("IsTop") = ShopTop
		Rs("isAccept") = 1
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	shopid = Rs("shopid")
	Rs.Close:Set Rs = Nothing
	Call RemoveCache
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		Dim url
		Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
		url = "admin_makeshop.asp?ChannelID=" & ChannelID & "&shopid=" & shopid & "&showid=0"	
		Call ScriptCreation(url,shopid)
	End If
	Succeed("<li>恭喜您！修改" & sModuleName & "成功。</li><li><a href=?action=view&ChannelID=" & ChannelID & "&shopid=" & shopid & ">点击此处查看该" & sModuleName & "</a></li>")
End Sub
Private Sub shopView()
	Call PageTop
	If Request("shopid") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	SQL = "select * from ECCMS_ShopList where shopid=" & Request("shopid")
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
	  <td align="center" class="TableRow2" colspan="2"><a href=?action=edit&ChannelID=<%=ChannelID%>&shopid=<%=Rs("shopid")%>><font size=4><%=Rs("TradeName")%></font></a></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong><%=sModuleName%>名称：</strong> <%=Rs("TradeName")%></td>
	  <td class="TableRow1"><strong><%=sModuleName%>型号：</strong> <%=Rs("Marque")%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong><%=sModuleName%>单位：</strong> <%=Rs("Unit")%></td>
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
	  <td class="TableRow1"><strong>原价：</strong> <%=FormatCurrency(Rs("PastPrice"))%> 元/<%=Rs("Unit")%></td>
	  <td class="TableRow1"><strong>现价：</strong> <%=FormatCurrency(Rs("NowPrice"))%> 元/<%=Rs("Unit")%></td>
	</tr>
	<tr>
	  <td class="TableRow1"><strong>上架时间：</strong> <%=Rs("addTime")%></td>
	  <td class="TableRow1"><strong>产品公司：</strong> <%=Rs("Company")%></td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1"><strong><%=sModuleName%>简介：</strong><br>&nbsp;&nbsp;&nbsp;&nbsp;<%=enchiasp.ReadContent(Rs("Explain"))%></td>
	</tr>
	<tr>
	  <td colspan="2" class="TableRow1">上一商品：<%=FrontShop(Rs("ShopID"))%>
	  <br>下一商品：<%=NextShop(Rs("ShopID"))%></td>
	</tr>
	<tr>
	  <td align="center" colspan="2" class="TableRow2"><input type="button" onclick="javascript:window.close()" value="关闭本窗口" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" onclick="javascript:history.go(-1)" value="返回上一页" name="B1" class=Button>&nbsp;&nbsp; 
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="返回顶部" class=button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='?action=edit&ChannelID=<%=ChannelID%>&ShopID=<%=Rs("ShopID")%>'" value="编辑商品" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub

Private Function FrontShop(ShopID)
	Dim Rss, SQL
	SQL = "select Top 1 ShopID,classid,TradeName from ECCMS_ShopList where ChannelID=" & ChannelID & " And isAccept <> 0 And ShopID < " & ShopID & " order by ShopID desc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		FrontShop = "已经没有了"
	Else
		FrontShop = "<a href=admin_Shop.asp?action=view&ChannelID=" & ChannelID & "&ShopID=" & Rss("ShopID") & ">" & Rss("TradeName") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Function NextShop(ShopID)
	Dim Rss, SQL
	SQL = "select Top 1 ShopID,classid,TradeName from ECCMS_ShopList where ChannelID=" & ChannelID & " And isAccept <> 0 And ShopID > " & ShopID & " order by ShopID asc"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		NextShop = "已经没有了"
	Else
		NextShop = "<a href=admin_Shop.asp?action=view&ChannelID=" & ChannelID & "&ShopID=" & Rss("ShopID") & ">" & Rss("TradeName") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Sub BatCreateHtml()
	Dim Allshopid,url
	Response.Write "<IE:Download ID=CreationID STYLE=""behavior:url(#default#download)"" />" & vbCrLf
	Response.Write "<ol>"
	Allshopid = Split(selshopid, ",")
	For i = 0 To UBound(Allshopid)
		shopid = CLng(Allshopid(i))
		url = "admin_makeshop.asp?ChannelID=" & ChannelID & "&shopid=" & shopid & "&showid=1"	
		Call ScriptCreation(url,shopid)
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

Private Sub ShopDel()
	If Request("ShopID") = "" Then
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT shopid,classid,HtmlFileDate FROM ECCMS_ShopList WHERE ChannelID = "& ChannelID &" And shopid=" & Request("shopid"))
	If Not(Rs.BOF And Rs.EOF) Then
		ClassUpdateCount Rs("classid"),0
		DeleteHtmlFile Rs("classid"),Rs("shopid"),Rs("HtmlFileDate")
	End If
	Rs.Close:Set Rs = Nothing
	enchiasp.Execute("DELETE FROM ECCMS_ShopList WHERE ShopID = " & Request("ShopID"))
	Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID=" & Request("ShopID"))
	Call RemoveCache
	OutHintScript("" & sModuleName & "删除成功！")
End Sub

Private Sub batdel()
	Set Rs = enchiasp.Execute("SELECT shopid,classid,HtmlFileDate FROM ECCMS_ShopList WHERE ChannelID = "& ChannelID &" And shopid in (" & selShopID & ")")
	Do While Not Rs.EOF
		ClassUpdateCount Rs("classid"),0
		DeleteHtmlFile Rs("classid"),Rs("shopid"),Rs("HtmlFileDate")
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	Conn.Execute ("DELETE FROM ECCMS_ShopList WHERE ShopID in (" & selShopID & ")")
	Conn.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID in (" & selShopID & ")")
	Call RemoveCache
	OutHintScript ("批量删除操作成功！")
End Sub

Private Sub isCommend()
	enchiasp.Execute ("update ECCMS_ShopList set isBest=1 where ShopID in (" & selShopID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noCommend()
	enchiasp.Execute ("update ECCMS_ShopList set isBest=0 where ShopID in (" & selShopID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub isTop()
	enchiasp.Execute ("update ECCMS_ShopList set isTop=1 where ShopID in (" & selShopID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub noTop()
	enchiasp.Execute ("update ECCMS_ShopList set isTop=0 where ShopID in (" & selShopID & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
Private Sub BatchSetting()
	If Not ChkAdmin("AdminShop" & ChannelID) Then
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
	Response.Write "		<td class=tablerow1><input type=""text"" name=""shopid"" size=70 value='"& Request("selshopid") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 width=""15%"" align=right><b>" & sModuleName & "公司：</b></td>"
	Response.Write "		<td class=tablerow1 width=""5%"" align=center><input type=checkbox name=selCompany value='1'></td>"
	Response.Write "		<td class=tablerow1 width=""60%""><input name=Company type=text size=60></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "单位：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selUnit value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=Unit type=text size=20>"
	Response.Write "		<select name=font2 onChange=""Unit.value=this.value;"">"
	Response.Write "		<option selected value=''>请选择单位</OPTION>"
	Response.Write "		<option value=套>套</option>"
	Response.Write "		<option value=件>件</option>"
	Response.Write "		<option value=台>台</option>"
	Response.Write "		<option value=盒>盒</option>"
	Response.Write "		<option value=部>部</option>"
	Response.Write "		<option value=瓶>瓶</option>"
	Response.Write "		<option value=个>个</option>"
	Response.Write "		<option value=本>本</option>"
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "货源：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selsupply value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=supply type=text size=20>"
	Response.Write "		<select name=font1 onChange=""supply.value=this.value;"">"
	Response.Write "		<option selected value=''>请选择</option>"
	Response.Write "		<option value=有货>有货</option>"
	Response.Write "		<option value=限量>限量</option>"
	Response.Write "		<option value=无货>无货</option>"
	Response.Write "		<option value=特惠>特惠</option>"
	Response.Write "		<option value=打折>打折</option>"
	Response.Write "		<option value=特价>特价</option>"
	Response.Write "		</select>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "原价：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selPastPrice value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=PastPrice type=text size=10 value=0> 元</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 align=right><b>" & sModuleName & "现价：</b></td>"
	Response.Write "		<td class=tablerow1 align=center><input type=checkbox name=selNowPrice value='1'></td>"
	Response.Write "		<td class=tablerow1><input name=NowPrice type=text size=10 value=0> 元</td>"
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
	If Not ChkAdmin("AdminShop" & ChannelID) Then
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
	Response.Write " <b>指定" & sModuleName & "ID：</b> <input type=""text"" name=""shopid"" size=80 value='"& Request("selshopid") &"'></td>"
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
		If Trim(Request.Form("shopid")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_ShopList SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And shopid in ("& Request("shopid") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_ShopList SET ClassID="& CLng(Request.Form("tClassID")) &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>恭喜您！批量移动完成。</li>")
End Sub

Private Sub SaveSetting()
	If Founderr = True Then Exit Sub
	Dim strTempValue
	strTempValue = ""
	If Trim(Request.Form("selCompany")) <> "" Then strTempValue = strTempValue & "Company='"& enchiasp.ChkFormStr(Request.Form("Company")) &"',"
	If Trim(Request.Form("selUnit")) <> "" Then strTempValue = strTempValue & "Unit='"& enchiasp.ChkFormStr(Request.Form("Unit")) &"',"
	If Trim(Request.Form("selsupply")) <> "" Then strTempValue = strTempValue & "supply='"& enchiasp.ChkFormStr(Request.Form("supply")) &"',"
	If Trim(Request.Form("selPastPrice")) <> "" Then strTempValue = strTempValue & "PastPrice="& CLng(Request.Form("PastPrice")) &","
	If Trim(Request.Form("selNowPrice")) <> "" Then strTempValue = strTempValue & "NowPrice="& CLng(Request.Form("NowPrice")) &","
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
		If Trim(Request.Form("shopid")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择" & sModuleName & "ID。</li>"
			Exit Sub
		Else
			SQL = "UPDATE ECCMS_ShopList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And shopid in ("& Request("shopid") &")"
		End If
	Else
		If Trim(Request.Form("ClassID")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请选择分类。</li>"
			Exit Sub
		ElseIf Trim(Request.Form("ClassID")) = "-1" Then
			SQL = "UPDATE ECCMS_ShopList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID
		Else
			SQL = "UPDATE ECCMS_ShopList SET "& strTempValue &" WHERE isAccept>0 And ChannelID = "& ChannelID &" And ClassID in ("& Request("ClassID") &")"
		End If
	End If
	enchiasp.Execute(SQL)
	Succeed("<li>恭喜您！批量设置完成。</li>")
End Sub

Private Sub ResetDateTime()
	Response.Write "<br><table width='400' align=center border=0 cellspacing=1 cellpadding=1>" & vbCrLf
	Response.Write "<tr> " & vbCrLf
	Response.Write "<td style=""border: 1px #384780 solid ;background-color: #FFFFFF;"">" & vbCrLf
	Response.Write "<table width=0 id=table2 name=table2 border=0 cellspacing=0 cellpadding=0>" & vbCrLf
	Response.Write "<tr height=12><td bgcolor=#36D91A>" & vbCrLf
	Response.Write "</td></tr></table></td></tr><tr> " & vbCrLf
	Response.Write "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span> <span style=""font-size:9pt"">%</span></td></tr>" & vbCrLf
	Response.Write "</table><br>" & vbCrLf
	Response.Flush
	Dim totalrec,addTime,page_count
	i = 0
	page_count = 0
	totalrec = enchiasp.Execute("SELECT COUNT(shopid) FROM [ECCMS_ShopList] WHERE ChannelID = "& ChannelID &" And isAccept>0")(0)
	Set Rs = enchiasp.Execute("SELECT shopid,addTime FROM [ECCMS_ShopList] WHERE ChannelID = "& ChannelID &" And isAccept>0 ORDER BY addTime DESC")
	If Not (Rs.BOF And Rs.EOF) Then
		Do While Not Rs.EOF
			Response.Write "<script>"
			Response.Write "table2.style.width=" & Fix((page_count / totalrec) * 400) & ";"
			Response.Write "txt2.innerHTML=""完成：" & FormatNumber(page_count / totalrec * 100, 2, -1) & """;"
			Response.Write "</script>" & vbCrLf
			Response.Flush
			addTime = DateAdd("s", -i, Rs("addTime"))
			enchiasp.Execute ("UPDATE ECCMS_ShopList SET addTime='" & addTime & "' WHERE shopid="& Rs("shopid"))
			Rs.movenext
			i = i + 5
			page_count = page_count + 1
		Loop
	End If
	Set Rs = Nothing
	Response.Write "<script>table2.style.width=400;txt2.innerHTML=""完成：100"";</script>"
End Sub

Private Sub RemoveCache()
	enchiasp.DelCahe "RenewStatistics"
	enchiasp.DelCahe "TotalStatistics"
End Sub
%>


