<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="include/collection.asp"-->
<%
Server.ScriptTimeOut = 18000
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
If LCase(Request("Action")) <> "savenew" Then
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write " <tr>"
	Response.Write "   <th>" & sModuleName & "HTTP采集管理</th>"
	Response.Write " </tr>"
	Response.Write " <tr>"
	Response.Write "   <td class=TableRow1><b>说明：</b><br>"
	Response.Write "&nbsp;&nbsp;①、第一次使用本功能，请修改<a href='?action=config&ChannelID=" & ChannelID & "' class='showlink'>采集基本设置</a>；<br>"
	Response.Write "&nbsp;&nbsp;②、采集前请<font color=blue>编辑</font>采集项目，选择正确的分类，然后<font color=blue>演示</font>项目确定无误后再进行采集。<br>"
	Response.Write "	</td> </tr>"
	Response.Write " <tr>"
	Response.Write "   <td class=TableRow2><strong>操作选项：</strong> <a href='?ChannelID=" & ChannelID & "'>管理首页</a> | "
	Response.Write "   <a href='?action=add&ChannelID=" & ChannelID & "'>添加采集项目</a> | "
	Response.Write "   <a href='?action=config&ChannelID=" & ChannelID & "' class='showmenu'>采集基本设置</a> | "
	Response.Write "   <a href='?action=remove&ChannelID=" & ChannelID & "'>系统缓存清理</a></td> "
	Response.Write " </tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End If

If Not CheckAdmin("ArticleCollect") Then
	Server.Transfer ("showerr.asp")
	Response.End
End If

Dim Myenchicms

On Error Resume Next
Set Myenchicms = New Cls_NewsCollection
Myenchicms.ChannelPath = enchiasp.InstallDir & enchiasp.ChannelDir
Myenchicms.ModuleName = sModuleName
Myenchicms.ReadNewsConfig
Myenchicms.ShowCollection
If LCase(Request("Action")) <> "savenew" Then Admin_footer
Set Myenchicms = Nothing
Set Myenchiasp = Nothing
CloseConn

Class Cls_NewsCollection

	Private ScriptName, ChannelID, ChannelDir, sModuleName
	Private maxperpage, Action, isEdit, Rs, SQL, CacheData, CacheItemData

	Private AdminName, ItemID, HTTPHtmlCode, TableMarquee

	'--项目基本设置变量
	Private stopGather, RepeatDeal, MaxPicSize, AllowPicExt, setInterval
	'--采集项目变量
	Private ClassID, SpecialID, StopItem, Encoding, IsDown, AutoClass, PathForm
	Private IsNowTime, AllHits, star, RemoveCode, RemoteListUrl
	Private PaginalList, IsPagination, startid, lastid, FindListCode
	Private FindInfoCode, RetuneClass, IsNextPage, strReplace


	'-- 频道目录
	Public Property Let PageListNum(ByVal NewValue)
		maxperpage = NewValue
	End Property
	'-- 频道模块名称
	Public Property Let ModuleName(ByVal NewValue)
		sModuleName = NewValue
	End Property
	'-- 频道目录
	Public Property Let ChannelPath(ByVal NewValue)
		ChannelDir = NewValue
	End Property

	Private Sub Class_Initialize()
		On Error Resume Next
		
		ChannelID = 1
		maxperpage = 30
		ScriptName = "Admin_ArticleGather.Asp"
		sModuleName = "文章"
		ChannelDir = "/article/"
	End Sub

	Private Sub Class_Terminate()
		If IsObject(MyConn) Then
			MyConn.Close
			Set MyConn = Nothing
		End If
	End Sub

	Public Sub ReloadNewsItem(ItemID)
		If Not IsConnection Then DatabaseConnection
		Dim rsItem
		SQL = "SELECT * FROM [ECCMS_NewsItem] WHERE ItemID=" & ItemID
		Set rsItem = MyConn.Execute(SQL)
		Myenchiasp.Value = rsItem.GetRows(1)
		Set rsItem = Nothing
	End Sub
	Public Sub ReloadNewsConfig()
		If Not IsConnection Then DatabaseConnection
		SQL = "SELECT * FROM [ECCMS_NewsConfig] "
		Set Rs = MyConn.Execute(SQL)
		Myenchiasp.Value = Rs.GetRows(1)
		Set Rs = Nothing
	End Sub
	Public Sub ReadNewsConfig()
		On Error Resume Next
		
		Myenchiasp.Name = "NewsConfig"
		If Myenchiasp.ObjIsEmpty() Then ReloadNewsConfig
		CacheData = Myenchiasp.Value
		'第一次起用系统或者重启IIS的时候加载缓存
		Myenchiasp.Name = "Date"
		If Myenchiasp.ObjIsEmpty() Then
			Myenchiasp.Value = date
		Else
			If CStr(Myenchiasp.Value) <> CStr(date) Then
				Myenchiasp.Name = "NewsConfig"
				Call ReloadNewsConfig
				CacheData = Myenchiasp.Value
			End If
		End If
		
		stopGather = CacheData(1, 0): RepeatDeal = CacheData(2, 0): MaxPicSize = CacheData(3, 0)
		AllowPicExt = CacheData(4, 0): setInterval = CacheData(5, 0)
	End Sub
	'--读取项目设置
	Public Sub ReadNewsItem(ItemID)
		On Error Resume Next
		
		Myenchiasp.Name = "NewsItem" & ItemID
		If Myenchiasp.ObjIsEmpty() Then ReloadNewsItem (ItemID)
		CacheItemData = Myenchiasp.Value
		
		ClassID = CacheItemData(4, 0): SpecialID = CacheItemData(5, 0): StopItem = CacheItemData(6, 0)
		Encoding = CacheItemData(7, 0): IsDown = CacheItemData(8, 0)
		AutoClass = CacheItemData(9, 0): PathForm = CacheItemData(10, 0): IsNowTime = CacheItemData(11, 0)
		AllHits = CacheItemData(12, 0): star = CacheItemData(13, 0): RemoveCode = CacheItemData(14, 0)
		RemoteListUrl = CacheItemData(16, 0): PaginalList = CacheItemData(17, 0)
		IsPagination = CacheItemData(18, 0): startid = CacheItemData(19, 0): lastid = CacheItemData(20, 0)
		FindListCode = CacheItemData(21, 0): FindInfoCode = CacheItemData(22, 0)
		
		If Not IsNull(CacheItemData(23, 0)) Then
			RetuneClass = CacheItemData(23, 0)
		End If
		
		IsNextPage = CacheItemData(24, 0)
		
		If Not IsNull(CacheItemData(26, 0)) Then
			strReplace = CacheItemData(26, 0)
		End If
	End Sub
	Public Sub ShowCollection()
		TableMarquee = "<p align=center><div style=""width:200px;height:30px;position:absolute;"">"
		TableMarquee = TableMarquee & "<table align=center border=0 cellpadding=0 cellspacing=1 bgcolor=#000000 width='200' height='30'><tr><td bgcolor=#0650D2><marquee align=middle behavior=alternate scrollamount=5 style=""font-size:9pt""><font color=#FFFFFF>...正在收集数据...请稍候...</font></marquee></td></tr></table>"
		TableMarquee = TableMarquee & "</div></p>"

		On Error Resume Next
		If Not IsConnection Then DatabaseConnection
		ChannelID = Myenchiasp.ChkNumeric(Request("ChannelID"))
		If ChannelID = 0 Then ChannelID = 1
		ChannelID = CLng(ChannelID)
		AdminName = enchiasp.CheckStr(Session("AdminName"))
		Action = LCase(Request("action"))
		Select Case Trim(Action)
		Case "copy"
			Call CopyNewItem
		Case "del"
			Call DeleteItem
		Case "config"
			Call BasalConfig
		Case "save"
			Call SaveConfig
		Case "edit"
			ItemID = Myenchiasp.ChkNumeric(Request("ItemID"))
			If ItemID = 0 Then
				OutErrors ("请选择正确的项目ID!")
				Exit Sub
			End If
			isEdit = True
			Call CollectionItem(isEdit)
		Case "add"
			isEdit = False
			Call CollectionItem(isEdit)
		Case "step2"
			Call ItemStep2
		Case "step3"
			Call ItemStep3
		Case "step4"
			Call ItemStep4
		Case "demo"
			Call ItemStep4
		Case "begin"
			BeginCollection
		Case "savenew"
			DataCollection
		Case "remove"
			RemoveAllCache
		Case Else
			Call showmain
		End Select
	End Sub
	Private Sub showmain()
		Dim totalnumber, Pcount, CurrentPage
		Dim i, stylestr
		
		With Response
		.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
		.Write "<table border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""tableborder"">"
		.Write "<tr>"
		.Write " <th>项目名称</th>"
		.Write " <th>所属分类</th>"
		.Write " <th>所属专题</th>"
		.Write " <th>状态</th>"
		.Write " <th>上次采集时间</th>"
		.Write " <th>管理操作</th>"
		.Write "</tr>"
		
		totalnumber = MyConn.Execute("SELECT COUNT(ItemID) FROM ECCMS_NewsItem WHERE ChannelID=" & ChannelID)(0)
		
		CurrentPage = Myenchiasp.ChkNumeric(Request("page"))
		CurrentPage = CLng(CurrentPage)
		If CurrentPage = 0 Then CurrentPage = 1
		Pcount = CLng(totalnumber / maxperpage) '得到总页数
		If Pcount < totalnumber / maxperpage Then Pcount = Pcount + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > Pcount Then CurrentPage = Pcount
		
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT ItemID,ItemName,SiteUrl,ChannelID,ClassID,SpecialID,StopItem,lastime,RemoteListUrl FROM [ECCMS_NewsItem] WHERE ChannelID=" & ChannelID & " ORDER BY lastime DESC,ItemID DESC"
		Rs.Open SQL, MyConn, 1, 1
		
		If Rs.BOF And Rs.EOF Then
			.Write "<tr><td align=center colspan=9 class=TableRow2>还没有添加任何采集项目！</td></tr>"
		Else
			If Pcount > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			i = 0
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.End
				If (i Mod 2) = 0 Then
					stylestr = "class=TableRow1"
				Else
					stylestr = "class=TableRow2"
				End If
				.Write "<tr align=center>"
				.Write " <td " & stylestr & " title='点击打开目标网站'><a href='" & Rs("SiteUrl") & "' target=_blank>" & Rs("ItemName") & "</a></td>"
				.Write " <td " & stylestr & " title='点击查看目标网站列表'><a href='" & Rs("RemoteListUrl") & "' target=_blank>" & Read_Class_Name(Rs("ClassID")) & "</a></td>"
				.Write " <td " & stylestr & ">" & Read_Special_Name(Rs("SpecialID")) & "</td>"
				.Write " <td " & stylestr & ">"
				If Rs("StopItem") = 0 Then
					.Write "<font color=blue>√</font>"
				Else
					.Write "<font color=red>×</font>"
				End If
				.Write "</td>"
				.Write " <td " & stylestr & ">"
				
				If DateDiff("D", Rs("lastime"), Now()) = 0 Then
					.Write "<font color=red>"
					.Write Rs("lastime")
					.Write "</font>"
				Else
					.Write Rs("lastime")
				End If
				.Write "</td>"
				.Write " <td " & stylestr & "><a href='?action=edit&ItemID=" & Rs("ItemID") & "&ChannelID=" & ChannelID & "'>编辑</a> | "
				.Write "<a href='?action=begin&ItemID=" & Rs("ItemID") & "&ChannelID=" & ChannelID & "'>采集</a> | "
				.Write "<a href='?action=demo&ItemID=" & Rs("ItemID") & "&ChannelID=" & ChannelID & "'>演示</a> | "
				.Write "<a href='?action=copy&ItemID=" & Rs("ItemID") & "&ChannelID=" & ChannelID & "'>克隆</a> | "
				.Write "<a href='?action=del&ItemID=" & Rs("ItemID") & "&ChannelID=" & ChannelID & "' onclick=""{if(confirm('您确定要删除此项目吗?')){return true;}return false;}"">删除</a>"
				.Write "</td>"
				.Write "</tr>"
				Rs.MoveNext
				i = i + 1
				If i >= maxperpage Then Exit Do
			Loop
		End If
		Rs.Close
		Set Rs = Nothing
		.Write "<tr>"
		.Write " <td colspan=""9"" class=""tablerow2"" align=""right"">"
		ShowListPage CurrentPage, Pcount, totalnumber, maxperpage, "&ChannelID=" & ChannelID & "", sModuleName & "采集"
		.Write "</td>"
		.Write "</tr>"
		If LCase(Request("action")) = "yes" Then
			.Write "<tr>"
			.Write " <td colspan=9 class=tablerow2>"
			.Write "<b class=style2>恭喜您！采集" & sModuleName & "全部完成..."
			.Write "成功采集" & sModuleName & " <font color=""#FF0000"">" & Session("SucceedCount") & "</font> 个，总费时 <font color=""#FF0000"">" & FormatNumber((Timer() - Request("D")), 2, -1) & "</font> 秒,完成时间" & Now() & "</b>"
			.Write "</td>"
			.Write "</tr>"
			Session("SucceedCount") = 0
		End If
		.Write "</table>"
		End With
	End Sub
	'=================================================
	'函数名：Read_Class_Name
	'作  用：读取分类名称
	'=================================================
	Private Function Read_Class_Name(ByVal ClassID)
		Dim rsClass

		On Error Resume Next
		Set rsClass = enchiasp.Execute("SELECT ClassName FROM ECCMS_Classify WHERE ClassID=" & ClassID)
		If rsClass.BOF And rsClass.EOF Then
			Read_Class_Name = "没有分类"
			Set rsClass = Nothing
			Exit Function
		End If
		Read_Class_Name = rsClass(0)
		Set rsClass = Nothing
	End Function
	'=================================================
	'函数名：Read_Special_Name
	'作  用：读取专题名称
	'=================================================
	Private Function Read_Special_Name(ByVal SpecialID)
		Dim rsSpecial
		On Error Resume Next
		Set rsSpecial = enchiasp.Execute("SELECT SpecialName FROM ECCMS_Special WHERE SpecialID=" & SpecialID)
		If rsSpecial.BOF And rsSpecial.EOF Then
			Read_Special_Name = "没有指定专题"
			Set rsSpecial = Nothing
			Exit Function
		End If
		Read_Special_Name = rsSpecial(0)
		Set rsSpecial = Nothing
	End Function
	'=================================================
	'函数名：GetClassID
	'作  用：读取分类ID
	'=================================================
	Public Function GetClassID(ByVal chanid, ByVal superior, ByVal inferior)
		superior = Replace(Trim(superior), "'", "")
		inferior = Replace(Trim(inferior), "'", "")
		chanid = Myenchiasp.ChkNumeric(chanid)
		If superior = "" Or chanid = 0 Then
			GetClassID = 0
			Exit Function
		End If
		On Error Resume Next
		Dim oRs, SQL, clsid, iRs
		clsid = 0
		SQL = "SELECT ClassID,ClassName,child FROM [ECCMS_Classify] WHERE ChannelID=" & chanid & " And TurnLink=0 And ClassName='" & superior & "'"
		Set oRs = enchiasp.Execute(SQL)
		If Not (oRs.BOF And oRs.EOF) Then
			If oRs("child") = 0 Then
				clsid = oRs("ClassID")
			Else
				If inferior <> "" Then
					Set iRs = enchiasp.Execute("SELECT ClassID,ClassName,child FROM [ECCMS_Classify] WHERE ChannelID=" & chanid & " And parentid=" & oRs("classid") & " And child=0 And TurnLink=0 And ClassName='" & inferior & "'")
					If Not (iRs.BOF And iRs.EOF) Then
						clsid = iRs("ClassID")
					End If
					Set iRs = Nothing
				End If
			End If
		Else
			clsid = 0
		End If
		Set oRs = Nothing
		GetClassID = clsid
	End Function
	Public Function ClassUpdateCount(ChannelID, sortid)
		Dim rscount, Parentstr
		On Error Resume Next
		Set rscount = enchiasp.Execute("SELECT ClassID,Parentstr FROM [ECCMS_Classify] WHERE ChannelID = " & CLng(ChannelID) & " And ClassID=" & CLng(sortid))
		If Not (rscount.BOF And rscount.EOF) Then
			Parentstr = rscount("Parentstr") & "," & rscount("ClassID")
			enchiasp.Execute ("UPDATE [ECCMS_Classify] SET ShowCount=ShowCount+1,isUpdate=1 WHERE ChannelID = " & CLng(ChannelID) & " And ClassID in (" & Parentstr & ")")
		End If
		Set rscount = Nothing
	End Function
	'--采集基本设置
	Private Sub BasalConfig()
		With Response
			.Write "<form name=myform method=post action='?action=save'>" & vbCrLf
			.Write "<input type=hidden name='ChannelID' value='" & ChannelID & "'>" & vbCrLf
			.Write "<table  border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder""> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <th colspan=""2"">" & sModuleName & "采集基本设置</th> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td width=""23%"" align=""right"" nowrap class=""TableRow1""><strong>采集功能开关：</strong></td> " & vbCrLf
			.Write "    <td width=""77%"" class=""TableRow1""><input name=""stopGather"" type=""radio"" value=""1"""
			If CInt(stopGather) = 1 Then .Write " checked"
			.Write ">" & vbCrLf
			.Write "      关闭　　" & vbCrLf
			.Write "      <input type=""radio"" name=""stopGather"" value=""0"""
			If CInt(stopGather) = 0 Then .Write " checked"
			.Write ">" & vbCrLf
			.Write "      打开　　" & vbCrLf
			 .Write "      <input type=""radio"" name=""stopGather"" value=""9"""
			If CInt(stopGather) = 9 Then .Write " checked"
			.Write ">" & vbCrLf
			.Write "      采集测试<font color='red'>(供测试程序用，不写数据库)</font></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>重复" & sModuleName & "处理：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><input name=""RepeatDeal"" type=""radio"" value=""0"""
			If CInt(RepeatDeal) = 0 Then .Write " checked"
			.Write ">" & vbCrLf
			.Write "      跳过　　" & vbCrLf
			.Write "      <input type=""radio"" name=""RepeatDeal"" value=""1"""
			If CInt(RepeatDeal) > 0 Then .Write " checked"
			.Write ">" & vbCrLf
			.Write "      更新 </td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>允许下载的图片大小：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><input name=""MaxPicSize"" type=""text"" id=""MaxPicSize"" size=""12"" value=""" & MaxPicSize & """ maxlength=""10""> " & vbCrLf
			.Write "      <strong><font color=""blue"">KB </font></strong>&nbsp;&nbsp;<font color=""red"">* 不限制请输入“0”</font></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>允许下载的文件类型：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><input name=""AllowPicExt"" type=""text"" id=""AllowPicExt"" size=""50"" value=""" & AllowPicExt & """ maxlength=""255""> " & vbCrLf
			.Write "      <font color=""blue"">* 每个文件类型请用“|”分开</font></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>采集过程间隔时间：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""> <input name=""setInterval"" type=""text"" id=""setInterval"" size=""12"" value=""" & setInterval & """ maxlength=""10""> " & vbCrLf
			.Write "      <font color=""blue"">毫秒 </font></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2"">&nbsp;</td> " & vbCrLf
			.Write "    <td class=""TableRow2""><div align=""center""> " & vbCrLf
			.Write "        <input name=""B12"" type=""button"" class=""Button"" onclick=""javascript:history.go(-1)"" value=""返回上一页""> " & vbCrLf
			.Write "&nbsp;&nbsp; " & vbCrLf
			.Write "<input name=""B22"" type=""submit"" class=""Button"" value=""保存设置"">" & vbCrLf
			.Write "</div></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "</table></form> " & vbCrLf
		End With
	End Sub
	Private Sub SaveConfig()
		If Len(Request.Form("AllowPicExt")) = 0 Then
			OutErrors ("请输入允许下载的文件类型!")
			Exit Sub
		End If
		Myenchiasp.DelCahe ("NewsConfig")
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM ECCMS_NewsConfig WHERE id=1"
		Rs.Open SQL, MyConn, 1, 3
			Rs("stopGather") = Myenchiasp.ChkNumeric(Request.Form("stopGather"))
			Rs("RepeatDeal") = Myenchiasp.ChkNumeric(Request.Form("RepeatDeal"))
			Rs("setInterval") = Myenchiasp.ChkNumeric(Request.Form("setInterval"))
			Rs("MaxPicSize") = Myenchiasp.ChkNumeric(Request.Form("MaxPicSize"))
			Rs("AllowPicExt") = Trim(Request.Form("AllowPicExt"))
		Rs.Update
		Rs.Close: Set Rs = Nothing
		OutScript ("保存采集基本设置成功!")
	End Sub
	'--项目设置步骤
	Private Sub SettingStep(ItemID)
		With Response
			.Write "<tr>" & vbNewLine
			.Write " <td colspan=2 align=center class=tablerow2>"
			.Write "<a href='?ChannelID=" & ChannelID & "' style=""color: green;"">管理首页</a> | "
			.Write "<a href='?action=edit&ChannelID=" & ChannelID & "&ItemID=" & ItemID & "' class=showmenu>设置第一步</a> | "
			.Write "<a href='?action=step2&ChannelID=" & ChannelID & "&ItemID=" & ItemID & "' class=showmenu>设置第二步</a> | "
			.Write "<a href='?action=step3&ChannelID=" & ChannelID & "&ItemID=" & ItemID & "' class=showmenu>设置第三步</a> | "
			.Write "<a href='?action=demo&ChannelID=" & ChannelID & "&ItemID=" & ItemID & "' class=showmenu>项目演示</a> | "
			.Write "<a href='?action=begin&ChannelID=" & ChannelID & "&ItemID=" & ItemID & "' style=""color: red;"">开始采集</a>"
			.Write "</td>" & vbNewLine
			.Write "</tr>" & vbNewLine
		End With
	End Sub
	'--编辑采集项目设置
	Private Sub CollectionItem(isEdit)
		Dim sClassSelect, RsObj, ItemTitle
		Dim i, ArrayRetuneClass
		Dim ArrayRemoveCode
		
		If isEdit Then
			Set Rs = MyConn.Execute("SELECT * FROM ECCMS_NewsItem WHERE ChannelID=" & ChannelID & " And ItemID=" & ItemID)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				OutErrors ("错误的系统参数!")
				Exit Sub
			End If
			ItemTitle = "编辑采集项目 第一步"
		Else
			ItemID = 0
			ItemTitle = "添加新的采集项目"
		End If
		With Response
			.Write "<script language=""javascript"" src=""include/Gatherer.js""></script>" & vbCrLf
			.Write "<form name=myform method=post action=""" & ScriptName & """ onSubmit='return CheckForm();'>" & vbCrLf
			.Write "<input type=""hidden"" name=""action"" value=""step2"">" & vbCrLf
			.Write "<input type=""hidden"" name=""ChannelID"" value=""" & ChannelID & """>" & vbCrLf
			.Write "<input type=""hidden"" name=""ItemID"" value=""" & ItemID & """>" & vbCrLf
			.Write "<input type=hidden name='change' value='yes'>" & vbNewLine
			.Write "<table  border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder""> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <th colspan=""2"">" & ItemTitle & "</th> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			If ItemID > 0 Then
				SettingStep (ItemID)
			End If
			.Write "  <tr> " & vbCrLf
			.Write "    <td width=""23%"" align=""right"" nowrap class=""TableRow1""><strong>项目名称：</strong></td> " & vbCrLf
			.Write "    <td width=""77%"" class=""TableRow1""><input name=""ItemName"" type=""text"" id=""ItemName"" size=""30"""
			If isEdit Then .Write " value=""" & Rs("ItemName") & """"
			.Write "></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>目标站点URL：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><input name=""SiteUrl"" type=""text"" id=""SiteUrl"" size=""30"""
			If isEdit Then
				.Write " value=""" & Rs("SiteUrl") & """"
			Else
				.Write " value=""http://"""
			End If
			.Write "></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>所属分类：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><select name=""ClassID"" size=""1"" id=""ClassID"">" & vbCrLf
			sClassSelect = enchiasp.LoadSelectClass(ChannelID)
			If isEdit Then
				sClassSelect = Replace(sClassSelect, "{ClassID=" & Rs("ClassID") & "}", "selected")
			End If
			.Write sClassSelect
			.Write "    </select></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>所属专题：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><select name=""SpecialID"" size=""1"" id=""SpecialID"">" & vbCrLf
			.Write "      <option value=""0"">不指定专题</option>" & vbCrLf
			
			Set RsObj = enchiasp.Execute("SELECT SpecialID,SpecialName FROM ECCMS_Special Where ChannelID = " & ChannelID & " ORDER BY orders")
			Do While Not RsObj.EOF
				.Write "        <option value=""" & RsObj("SpecialID") & """"
				If isEdit Then
					If Rs("SpecialID") = RsObj("SpecialID") Then .Write " selected"
				End If
				.Write ">"
				.Write RsObj("SpecialName")
				.Write "</option>" & vbCrLf
				RsObj.MoveNext
			Loop
			Set RsObj = Nothing
			
			.Write "    </select></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>关闭采集项目：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><input name=""StopItem"" type=""radio"" value=""0"""
			If isEdit Then
				If Rs("StopItem") = 0 Then .Write " checked"
			Else
				.Write " checked"
			End If
			.Write "> 打开&nbsp;&nbsp;　" & vbCrLf
			.Write "      <input type=""radio"" name=""StopItem"" value=""1"""
			If isEdit Then
				If Rs("StopItem") > 0 Then .Write " checked"
			End If
			.Write "> 关闭</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>目标文档编码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><input name=""Encoding"" type=""text"" id=""Encoding"" size=""15"""
			If isEdit Then
				.Write " value=""" & Rs("Encoding") & """"
			Else
				.Write " value=""GB2312"""
			End If
			.Write "> " & vbCrLf
			.Write "      <span class=""style2"">请选择编码→</span>      <select name=""selEncoding"" size=""1"" onChange=""Encoding.value=this.value;"">" & vbCrLf
			.Write "        <option>请选择编码</option>" & vbCrLf
			.Write "        <option value=""GB2312"">GB2312</option>" & vbCrLf
			.Write "        <option value=""UTF-8"">UTF-8</option>" & vbCrLf
			.Write "        <option value=""BIG5"">BIG5</option>" & vbCrLf
			.Write "          </select>" & vbCrLf
			.Write "      </td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>是否下载图片到本地：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><input name=""IsDown"" type=""radio"" value=""0"""
			If isEdit Then
				If Rs("IsDown") = 0 Then .Write " checked"
			Else
				.Write " checked"
			End If
			.Write "> 否&nbsp;&nbsp;" & vbCrLf
			.Write "      <input type=""radio"" name=""IsDown"" value=""1"""
			If isEdit Then
				If Rs("IsDown") > 0 Then .Write " checked"
			End If
			.Write "> 是 </td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr>" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>是否自动归类：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow2""><input name=""AutoClass"" type=""radio"" value=""0"""
			If isEdit Then
				If Rs("AutoClass") = 0 Then .Write " checked"
			Else
				.Write " checked"
			End If
			.Write " onClick=""RetuneClassID.style.display='none';""> 否&nbsp;&nbsp;" & vbCrLf
			.Write "      <input type=""radio"" name=""AutoClass"" value=""1"""
			If isEdit Then
				If Rs("AutoClass") > 0 Then .Write " checked"
			End If
			.Write " onClick=""RetuneClassID.style.display='';""> 是</td>" & vbCrLf ' disabled
			.Write "  </tr>" & vbCrLf
			.Write "  <tr id=""RetuneClassID"""
			If isEdit Then
				If Rs("AutoClass") = 0 Then .Write " style=""display:none"""
			Else
				.Write " style=""display:none"""
			End If
			.Write ">" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>分类替换操作：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow2""><table border=""0"" cellpadding=""3""><tr><td><select name=""RetuneClass"" id=name=""RetuneClass"" style=""width:300;height:100"" size=""2"" ondblclick=""return ModifyCalss();"">" & vbCrLf
			If isEdit Then
				If Not IsNull(Rs("RetuneClass")) Then
					ArrayRetuneClass = Split(Rs("RetuneClass"), "$$$")
					For i = 0 To UBound(ArrayRetuneClass)
						If Len(ArrayRetuneClass(i)) > 3 Then
							.Write "      <option value=""" & ArrayRetuneClass(i) & """>" & ArrayRetuneClass(i) & "</option>" & vbCrLf
						End If
					Next
					
				End If
			End If
			.Write "        " & vbCrLf
			.Write "      </select></td><td>" & vbCrLf
			.Write "      <input type=""button"" name=""addclass"" value=""添加替换分类"" class=""button"" onclick=""AddClass();""><br><br style=""overflow: hidden; line-height: 5px"">" & vbCrLf
			.Write "      <input type=""button"" name=""modifyclass"" value=""修改当前分类"" class=""button"" onclick=""return ModifyClass();""><br><br style=""overflow: hidden; line-height: 5px"">" & vbCrLf
			.Write "      <input type=""button"" name=""delclass"" value=""删除当前分类"" class=""button"" onclick=""DelClass();""><br>" & vbCrLf
			.Write "      <input type=""hidden"" name=""ClassList"" value="""">" & vbCrLf
			.Write "        </td><tr></table>" & vbCrLf
			.Write "      </td>" & vbCrLf
			.Write "  </tr>" & vbCrLf
			.Write "  <tr>" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>日期路径格式：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow1""><select name=""PathForm"" size=""1"" id=""PathForm"">" & vbCrLf
			.Write "      <option value=""0"""
			If isEdit Then
				If Rs("PathForm") = 0 Then .Write " selected"
			End If
			.Write ">不使用日期目录</option>" & vbCrLf
			.Write "      <option value=""1"""
			If isEdit Then
				If Rs("PathForm") = 1 Then .Write " selected"
			Else
				.Write " selected"
			End If
			.Write ">2005-8</option>" & vbCrLf
			.Write "      <option value=""2"""
			If isEdit Then
				If Rs("PathForm") = 2 Then .Write " selected"
			End If
			.Write ">2005_8</option>" & vbCrLf
			.Write "      <option value=""3"""
			If isEdit Then
				If Rs("PathForm") = 3 Then .Write " selected"
			End If
			.Write ">20058</option>" & vbCrLf
			.Write "      <option value=""4"""
			If isEdit Then
				If Rs("PathForm") = 4 Then .Write " selected"
			End If
			.Write ">2005</option>" & vbCrLf
			.Write "      <option value=""5"""
			If isEdit Then
				If Rs("PathForm") = 5 Then .Write " selected"
			End If
			.Write ">2005/8</option>" & vbCrLf
			.Write "      <option value=""6"""
			If isEdit Then
				If Rs("PathForm") = 6 Then .Write " selected"
			End If
			.Write ">2005/8/8</option>" & vbCrLf
			.Write "      <option value=""7"""
			If isEdit Then
				If Rs("PathForm") = 7 Then .Write " selected"
			End If
			.Write ">200588</option>" & vbCrLf
			.Write "    </select></td>" & vbCrLf
			.Write "  </tr>" & vbCrLf
			.Write "  <tr>" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>是否显示为最新时间：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow2""><input name=""IsNowTime"" type=""radio"" value=""0"""
			If isEdit Then
				If Rs("IsNowTime") = 0 Then .Write " checked"
			Else
				.Write " checked"
			End If
			.Write "> 否&nbsp;&nbsp;" & vbCrLf
			.Write "      <input type=""radio"" name=""IsNowTime"" value=""1"""
			If isEdit Then
				If Rs("IsNowTime") > 0 Then .Write " checked"
			End If
			.Write "> 是</td>" & vbCrLf
			.Write "  </tr>" & vbCrLf
			.Write "  <tr>" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>初始点击数：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow1""><input name=""AllHits"" type=""text"" id=""AllHits"" size=""10"""
			If isEdit Then
				.Write " value=""" & Rs("AllHits") & """"
			Else
				.Write " value=""0"""
			End If
			.Write ">" & vbCrLf
			.Write "  </tr>" & vbCrLf
			.Write "  <tr>" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>软件星级：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow2""><select name=""star"" size=""1"" id=""star"">" & vbCrLf
			.Write "      <option value=""5"""
			If isEdit Then
				If Rs("star") = 5 Then .Write " selected"
			End If
			.Write ">★★★★★</option>" & vbCrLf
			.Write "      <option value=""4"""
			If isEdit Then
				If Rs("star") = 4 Then .Write " selected"
			End If
			.Write ">★★★★</option>" & vbCrLf
			.Write "      <option value=""3"""
			If isEdit Then
				If Rs("star") = 3 Then .Write " selected"
			Else
				.Write " selected"
			End If
			.Write ">★★★</option>" & vbCrLf
			.Write "      <option value=""2"""
			If isEdit Then
				If Rs("star") = 2 Then .Write " selected"
			End If
			.Write ">★★</option>" & vbCrLf
			.Write "      <option value=""1"""
			If isEdit Then
				If Rs("star") = 1 Then .Write " selected"
				ArrayRemoveCode = Split(Rs("RemoveCode"), "|")
			End If
			.Write ">★</option>" & vbCrLf
			.Write "    </select></td>" & vbCrLf
			.Write "  </tr>" & vbCrLf
			.Write "  <tr>" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong class=""TableRow1"">内容过滤设置：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow1""><input name=""RemoveCode0"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(0)) = 1 Then .Write " checked"
			Else
				.Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      SCRIPT " & vbCrLf
			.Write "      <input name=""RemoveCode1"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(1)) = 1 Then .Write " checked"
			Else
				.Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      IFARME " & vbCrLf
			.Write "      <input name=""RemoveCode2"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(2)) = 1 Then .Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      OBJECT " & vbCrLf
			.Write "      <input name=""RemoveCode3"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(3)) = 1 Then .Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      APPLET " & vbCrLf
			.Write "      <input name=""RemoveCode4"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(4)) = 1 Then .Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      DIV " & vbCrLf
			.Write "      <br>" & vbCrLf
			.Write "      <input name=""RemoveCode5"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(5)) = 1 Then .Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      FONT " & vbCrLf
			.Write "      <input name=""RemoveCode6"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(6)) = 1 Then .Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      SPAN " & vbCrLf
			.Write "      <input name=""RemoveCode7"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(7)) = 1 Then .Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      A " & vbCrLf
			.Write "      <input name=""RemoveCode8"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(8)) = 1 Then .Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      IMG " & vbCrLf
			.Write "      <input name=""RemoveCode9"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(9)) = 1 Then .Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      FORM " & vbCrLf
			.Write "      <input name=""RemoveCode10"" type=""checkbox"" value=""1"""
			If isEdit Then
				If Myenchiasp.ChkNumeric(ArrayRemoveCode(10)) = 1 Then .Write " checked"
			End If
			.Write "> " & vbCrLf
			.Write "      HTML </td>" & vbCrLf
			.Write "  </tr>" & vbCrLf
			.Write "  <tr>" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong class=""TableRow2"">远程列表URL：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow2""><span class=""TableRow1"">" & vbCrLf
			.Write "      <input name=""RemoteListUrl"" type=""text"" id=""RemoteListUrl"" size=""70"""
			If isEdit Then
				.Write " value=""" & Rs("RemoteListUrl") & """"
			End If
			.Write ">" & vbCrLf
			.Write "    </span></td>" & vbCrLf
			.Write "  </tr>" & vbCrLf
			.Write "  <tr>" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong class=""TableRow1"">是否列表分页采集：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow1""><input name=""IsPagination"" type=""radio"" value=""0"""
			If isEdit Then
				If Rs("IsPagination") = 0 Then .Write " checked"
			Else
				.Write " checked"
			End If
			.Write " onClick=""Pageinate1.style.display='none';Pageinate2.style.display='none';""> 否&nbsp;&nbsp;" & vbCrLf
			.Write "      <input type=""radio"" name=""IsPagination"" value=""1"""
			If isEdit Then
				If Rs("IsPagination") > 0 Then .Write " checked"
			End If
			.Write " onClick=""Pageinate1.style.display='';Pageinate2.style.display='';""> 是</td>" & vbCrLf
			.Write "  </tr>" & vbCrLf
			.Write "  <tr id=""Pageinate1"""
			If isEdit Then
				If Rs("IsPagination") = 0 Then .Write " style=""display:'none';"""
			Else
				.Write " style=""display:'none';"""
			End If
			.Write ">" & vbCrLf
			
			.Write "    <td align=""right"" class=""TableRow2""><strong class=""TableRow2"">远程列表分页URL：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow2""><input name=""PaginalList"" type=""text"" id=""PaginalList"" size=""70"""
			If isEdit Then
				.Write " value=""" & Rs("PaginalList") & """"
			End If
			.Write ">" & vbCrLf
			.Write "      <span class=""style2"">      * 分页代码 <font color=""red"">{$pageid}</font></span></td>" & vbCrLf
			.Write "  </tr>" & vbCrLf
			.Write "  <tr id=""Pageinate2"""
			If isEdit Then
				If Rs("IsPagination") = 0 Then .Write " style=""display:'none';"""
			Else
				.Write " style=""display:'none';"""
			End If
			.Write ">" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong class=""TableRow1"">远程列表起始页：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow1"">开始页：" & vbCrLf
			.Write "    <input name=""startid"" type=""text"" id=""startid"" size=""6"""
			If isEdit Then
				.Write " value=""" & Rs("startid") & """"
			Else
				.Write " value=""1"""
			End If
			.Write ">&nbsp;-" & vbCrLf
			.Write "    结束页：" & vbCrLf
			.Write "    <input name=""lastid"" type=""text"" id=""lastid"" size=""6"""
			If isEdit Then
				.Write " value=""" & Rs("lastid") & """"
			Else
				.Write " value=""2"""
			End If
			.Write ">&nbsp;&nbsp;<span class=""style2"">* 例如：1 - 9 或者 9 - 1</span></td>" & vbCrLf
			.Write "  </tr>" & vbCrLf
			 '--内容字符替换操作
			.Write "  <tr>" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>内容字符替换操作：</strong></td>" & vbCrLf
			.Write "    <td class=""TableRow2""><table border=""0"" cellpadding=""3""><tr><td><select name=""strReplace"" id=""strReplace"" style=""width:380;height:100"" size=""2"" ondblclick=""return ModifyReplace();"">" & vbCrLf
			
			Dim strReplaceArray
			
			If isEdit Then
				If Not IsNull(Rs("strReplace")) Then
					strReplaceArray = Split(Rs("strReplace"), "$$$")
					For i = 0 To UBound(strReplaceArray)
						If Len(strReplaceArray(i)) > 1 Then
							.Write "      <option value=""" & strReplaceArray(i) & """>" & strReplaceArray(i) & "</option>" & vbCrLf
						End If
					Next
					
				End If
			End If
			.Write "        " & vbCrLf
			.Write "      </select></td><td>" & vbCrLf
			.Write "      <input type=""button"" name=""addreplace"" value=""添加替换字符"" class=""button"" onclick=""AddReplace();""><br><br style=""overflow: hidden; line-height: 5px"">" & vbCrLf
			.Write "      <input type=""button"" name=""modifyreplace"" value=""修改当前字符"" class=""button"" onclick=""return ModifyReplace();""><br><br style=""overflow: hidden; line-height: 5px"">" & vbCrLf
			.Write "      <input type=""button"" name=""delreplace"" value=""删除当前字符"" class=""button"" onclick=""DelReplace();""><br>" & vbCrLf
			.Write "      <input type=""hidden"" name=""ReplaceList"" value="""">" & vbCrLf
			.Write "        </td><tr></table>" & vbCrLf
			.Write "      </td>" & vbCrLf
			.Write "  </tr>" & vbCrLf
			.Write "  <tr>" & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1"">&nbsp;</td>" & vbCrLf
			.Write "    <td class=""TableRow1""><div align=""center"">" & vbCrLf
			.Write "      <input name=""B12"" type=""button"" class=""Button"" onclick=""javascript:history.go(-1)"" value=""返回上一页"">&nbsp;&nbsp; " & vbCrLf
			.Write "      <input name=""B22"" type=""submit"" class=""Button"" value="" 下一步 "">&nbsp;&nbsp;" & vbCrLf
			.Write "      <input name=""ShowCode"" type=""checkbox"" value=""1""> 显示源码" & vbCrLf
			.Write "        </div></td>" & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "</table> " & vbCrLf
			.Write "</form>" & vbCrLf
			
			If isEdit Then Rs.Close: Set Rs = Nothing
		End With
	End Sub
	Private Sub ItemStep2()
		Dim tmpRemoveCode, i, showcode
		Dim NewItemID, strFindListCode
		
		ItemID = Myenchiasp.ChkNumeric(Request("ItemID"))
		showcode = Myenchiasp.ChkNumeric(Request("showcode"))
		If Trim(Request("change")) = "yes" Then
			If Len(Trim(Request.Form("ItemName"))) = 0 Then
				OutErrors ("请填写项目名称！")
				Exit Sub
			End If
			If Len(Trim(Request.Form("SiteUrl"))) = 0 Then
				OutErrors ("请填写目标站点URL！")
				Exit Sub
			End If
			If Left(LCase(Request.Form("SiteUrl")), 4) <> "http" Then
				OutErrors ("目标站点URL输入错误，请在URL前面加上“http://”！")
				Exit Sub
			End If
			If Len(Trim(Request.Form("Encoding"))) < 3 Then
				OutErrors ("请选择目标站点的文件编码！")
				Exit Sub
			End If
			If Myenchiasp.ChkNumeric(Request.Form("AutoClass")) = 0 Then
				If Myenchiasp.ChkNumeric(Request.Form("ClassID")) = 0 Then
					OutErrors ("该一级分类已经有下属分类，不能采集；请重新选择分类！")
					Exit Sub
				End If
			End If
			If Len(Trim(Request.Form("RemoteListUrl"))) = 0 Then
				OutErrors ("请填写远程列表URL！")
				Exit Sub
			End If
			If Myenchiasp.ChkNumeric(Request.Form("IsPagination")) > 0 Then
				If Len(Trim(Request.Form("PaginalList"))) = 0 Then
					OutErrors ("请填写远程分页列表URL！")
					Exit Sub
				End If
			End If
			
			Myenchiasp.DelCahe "NewsItem" & ItemID
			
			For i = 0 To 10
				tmpRemoveCode = tmpRemoveCode & Myenchiasp.ChkNumeric(Request.Form("RemoveCode" & i & "")) & "|"
			Next
			tmpRemoveCode = tmpRemoveCode & "0|0|0|0|0|0|0|0|0"
			
			If ItemID = 0 Then
				SQL = "SELECT * FROM ECCMS_NewsItem WHERE (ItemID is null)"
			Else
				SQL = "SELECT * FROM ECCMS_NewsItem WHERE ItemID=" & ItemID
			End If
			
			Set Rs = CreateObject("ADODB.Recordset")
			Rs.Open SQL, MyConn, 1, 3
				If ItemID = 0 Then Rs.AddNew
				Rs("ItemName") = Trim(Request.Form("ItemName"))
				Rs("SiteUrl") = Trim(Request.Form("SiteUrl"))
				Rs("ChannelID") = ChannelID
				Rs("ClassID") = Myenchiasp.ChkNumeric(Request.Form("ClassID"))
				Rs("SpecialID") = Myenchiasp.ChkNumeric(Request.Form("SpecialID"))
				Rs("StopItem") = Myenchiasp.ChkNumeric(Request.Form("StopItem"))
				Rs("Encoding") = Trim(Request.Form("Encoding"))
				Rs("IsDown") = Myenchiasp.ChkNumeric(Request.Form("IsDown"))
				Rs("AutoClass") = Myenchiasp.ChkNumeric(Request.Form("AutoClass"))
				Rs("PathForm") = Myenchiasp.ChkNumeric(Request.Form("PathForm"))
				Rs("IsNowTime") = Myenchiasp.ChkNumeric(Request.Form("IsNowTime"))
				Rs("AllHits") = Myenchiasp.ChkNumeric(Request.Form("AllHits"))
				Rs("star") = Myenchiasp.ChkNumeric(Request.Form("star"))
				Rs("RemoveCode") = Trim(tmpRemoveCode)
				
				Rs("RemoteListUrl") = Trim(Request.Form("RemoteListUrl"))
				Rs("PaginalList") = Trim(Request.Form("PaginalList"))
				Rs("IsPagination") = Myenchiasp.ChkNumeric(Request.Form("IsPagination"))
				Rs("startid") = Myenchiasp.ChkNumeric(Request.Form("startid"))
				Rs("lastid") = Myenchiasp.ChkNumeric(Request.Form("lastid"))
				
				If ItemID = 0 Then
					Rs("lastime") = Now()
					Rs("FindListCode") = "0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0"
					Rs("FindInfoCode") = "0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0"
					Rs("IsNextPage") = 0
					Rs("NamedDemourl") = ""
				End If
				Rs("RetuneClass") = Trim(Request.Form("ClassList"))
				Rs("strReplace") = Trim(Request.Form("ReplaceList"))
			Rs.Update
			Rs.Close: Set Rs = Nothing
		End If
		
		Set Rs = CreateObject("ADODB.Recordset")
		If ItemID = 0 Then
			Rs.Open "SELECT TOP 1 ItemID,FindListCode FROM ECCMS_NewsItem WHERE ChannelID=" & ChannelID & " ORDER BY ItemID DESC", MyConn, 1, 1
		Else
			Rs.Open "SELECT ItemID,FindListCode FROM ECCMS_NewsItem WHERE ChannelID=" & ChannelID & " And ItemID=" & ItemID & "", MyConn, 1, 1
		End If
		
		NewItemID = Rs("ItemID")
		strFindListCode = Split(Rs("FindListCode"), "$$$")
		Rs.Close: Set Rs = Nothing
		
		With Response
			.Write "<form name=myform method=post action=""" & ScriptName & """>" & vbCrLf
			.Write "<input type=""hidden"" name=""action"" value=""step3"">" & vbCrLf
			.Write "<input type=""hidden"" name=""ChannelID"" value=""" & ChannelID & """>" & vbCrLf
			.Write "<input type=""hidden"" name=""ItemID"" value=""" & NewItemID & """>" & vbCrLf
			.Write "<input type=hidden name='change' value='yes'>" & vbNewLine
			.Write "<table  border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder""> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <th colspan=""2"">采集项目第二步</th> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			If ItemID > 0 Then
				SettingStep (ItemID)
			End If
			'--如果选择了显示源码
			If showcode > 0 Then
				HTTPHtmlCode = Myenchiasp.GetRemoteData(Trim(Request.Form("RemoteListUrl")), Trim(Request.Form("Encoding")))
				If HTTPHtmlCode = "" Then
					.Write "<script language=""javascript"">" & vbCrLf
					.Write "alert('获取远程信息出错！请确定你的远程列表URL输入无误。');"
					.Write "location.replace('?action=edit&" & ChannelID & "=1&ItemID=" & NewItemID & "');" & vbCrLf
					.Write "</script>" & vbCrLf
					Exit Sub
				End If
				.Write "  <tr> " & vbCrLf
				.Write "    <td class=""TableTitle"" align=""center"" colspan=""2"">项 目 编 辑 -- 采集目标网站源代码&nbsp;&nbsp;&nbsp;&nbsp;"
				.Write "<Input type=""radio"" value=""0"" name=""soucode"" onClick=""soucodeid.style.display='none';""> 关闭源代码窗口&nbsp;&nbsp;<Input type=""radio"" value=""1"" name=""soucode"" onClick=""soucodeid.style.display='';"" checked> 查看源代码"
				.Write "        </td> " & vbCrLf
				.Write "  </tr> " & vbCrLf
				.Write "  <tr> " & vbCrLf
				.Write "    <td class=""TableRow1"" colspan=""2"" id='soucodeid'><textarea name='content' id='content' wrap='OFF' style='width:100%;' rows='20'>"
				.Write Server.HTMLEncode(HTTPHtmlCode)
				.Write "</textarea><div align='right'><a href=""javascript:admin_Size(-20,'content')""><img src='images/minus.gif' unselectable=on border=0></a> <a href=""javascript:admin_Size(20,'content')""><img src='images/plus.gif' unselectable=on border=0></div></td> " & vbCrLf
				.Write "  </tr> " & vbCrLf
				.Write "  <tr> " & vbCrLf
				.Write "    <td class=""TableRow2"" colspan=""2"">"
				.Write "采集的目标地址 → <a href='" & Trim(Request.Form("RemoteListUrl")) & "' target='_blank'><font color='red'>" & Trim(Request.Form("RemoteListUrl")) & "</font></a>"
				.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='view-source:" & Trim(Request.Form("RemoteListUrl")) & "' target='_blank'><font color='blue'>点击查看目标源代码</font></a></td> " & vbCrLf
				.Write "  </tr> " & vbCrLf
			End If
			.Write "  <tr> " & vbCrLf
			.Write "    <td class=""TableTitle"" align=""center"" colspan=""2"">项 目 编 辑 -- 列表连接设置</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td width='25%' align=""right"" class=""TableRow1""><strong>获取列表开始代码：</strong></td> " & vbCrLf
			.Write "    <td width='75%' class=""TableRow1""><textarea name=FindListCode0 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindListCode(0))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>获取列表结束代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindListCode1 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindListCode(1))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>获取连接开始代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindListCode2 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindListCode(2))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>获取连接结束代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindListCode3 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindListCode(3))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			'--特殊设置开始
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>特殊设置：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><div><Input type=""radio"" value=""0"" name=""FindListCode4"" onClick=""especial.style.display='none';"""
			If Myenchiasp.ChkNumeric(strFindListCode(4)) = 0 Then .Write " checked"
			.Write "> 不作处理&nbsp;&nbsp;<Input type=""radio"" value=""1"" name=""FindListCode4"" onClick=""especial.style.display='';"""
			If Myenchiasp.ChkNumeric(strFindListCode(4)) > 0 Then .Write " checked"
			.Write " disabled> 重新定位"
			.Write "</div><div id='especial' style=""display:none""><input type=""text"" name=""FindListCode5"" size=60 value='"
			.Write Server.HTMLEncode(strFindListCode(5))
			.Write "'></div>"
			.Write "<div></div></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			'--特殊设置结束
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""></td> " & vbCrLf
			.Write "    <td class=""TableRow2"" align=""center"">"
			.Write "      <input name=""B12"" type=""button"" class=""Button"" onclick=""javascript:history.go(-1)"" value=""返回上一页"">&nbsp;&nbsp; " & vbCrLf
			.Write "      <input name=""B22"" type=""submit"" class=""Button"" value="" 下一步 "">&nbsp;&nbsp;" & vbCrLf
			.Write "      <input name=""ShowCode"" type=""checkbox"" value=""1""> 显示源码" & vbCrLf
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "</table> " & vbCrLf
			.Write "</form>" & vbCrLf
		End With
	End Sub
	'--采集项目第三步
	Private Sub ItemStep3()
		Dim i, showcode
		Dim tmpFindListCode
		Dim strEncoding, NamedDemourl
		Dim strRemoteLisCode, strRemoteListUrl
		Dim strFindListCode, strFindInfoCode
		
		ItemID = Myenchiasp.ChkNumeric(Request("ItemID"))
		showcode = Myenchiasp.ChkNumeric(Request("showcode"))
		
		If Trim(Request("change")) = "yes" Then
			'--如果是更新项目，执行下面的操作
			Myenchiasp.DelCahe "NewsItem" & ItemID
			For i = 0 To 5
				tmpFindListCode = tmpFindListCode & Request.Form("FindListCode" & i & "") & "$$$"
			Next
			tmpFindListCode = tmpFindListCode & "0$$$0$$$0$$$0$$$0$$$0"
			SQL = "SELECT ItemID,FindListCode FROM ECCMS_NewsItem WHERE ChannelID=" & ChannelID & " And ItemID=" & ItemID
			Set Rs = CreateObject("ADODB.Recordset")
			Rs.Open SQL, MyConn, 1, 3
			If Rs.BOF And Rs.EOF Then
				OutErrors ("错误的系统参数！")
				Set Rs = Nothing
				Exit Sub
			Else
				Rs("FindListCode") = tmpFindListCode
				Rs.Update
			End If
			Rs.Close: Set Rs = Nothing
		End If
		'--获取项目设置
		SQL = "SELECT ItemID,Encoding,RemoteListUrl,FindListCode,FindInfoCode,IsNextPage,NamedDemourl FROM ECCMS_NewsItem WHERE ChannelID=" & ChannelID & " And ItemID=" & ItemID
		Set Rs = MyConn.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			OutErrors ("错误的系统参数！")
			Set Rs = Nothing
			Exit Sub
		Else
			strEncoding = Trim(Rs("Encoding"))
			RemoteListUrl = Trim(Rs("RemoteListUrl"))
			strFindListCode = Split(Myenchiasp.ReplaceTrim(Rs("FindListCode")), "$$$")
			strFindInfoCode = Split(Rs("FindInfoCode"), "$$$")
			IsNextPage = Rs("IsNextPage")
			If Not IsNull(Rs("NamedDemourl")) Then
				NamedDemourl = Rs("NamedDemourl")
			End If
		End If
		Rs.Close: Set Rs = Nothing

		With Response
			.Write "<form name=myform method=post action=""" & ScriptName & """>" & vbCrLf
			.Write "<input type=""hidden"" name=""action"" value=""step4"">" & vbCrLf
			.Write "<input type=""hidden"" name=""ChannelID"" value=""" & ChannelID & """>" & vbCrLf
			.Write "<input type=""hidden"" name=""ItemID"" value=""" & ItemID & """>" & vbCrLf
			.Write "<input type=hidden name='change' value='yes'>" & vbNewLine
			.Write "<table  border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder""> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <th colspan=""2"">采集项目第三步</th> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			If ItemID > 0 Then
				SettingStep (ItemID)
			End If
			'--如果选择了显示源码,开始获取远程信息
			If showcode > 0 Then
				'--获取远程列表网页源代码Myenchiasp.ReplaceTrim(
				
				HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(RemoteListUrl, strEncoding))
				If HTTPHtmlCode = "" Then
					OutErrors ("获取远程信息出错！请确定你的远程列表URL输入无误。")
					Exit Sub
				End If
				
				'--获取远程列表代码
				strRemoteLisCode = Myenchiasp.CutFixed(HTTPHtmlCode, strFindListCode(0), strFindListCode(1))
				strRemoteLisCode = Myenchiasp.ReplacedTrim(strRemoteLisCode)
				If strRemoteLisCode = "" Then
					OutErrors ("获取远程列表出错！请确定你的远程列表开始和结束代码输入无误。")
					Exit Sub
				End If
				'--获取列表URL
				strRemoteListUrl = Myenchiasp.CutFixed(strRemoteLisCode, strFindListCode(2), strFindListCode(3))
				strRemoteListUrl = Myenchiasp.FormatRemoteUrl(RemoteListUrl, strRemoteListUrl)
				If strRemoteListUrl = "" Then
					OutErrors ("获取远程连接出错！请确定你的连接开始和结束代码输入无误。")
					Exit Sub
				End If
				HTTPHtmlCode = Myenchiasp.GetRemoteData(strRemoteListUrl, strEncoding)
				If HTTPHtmlCode = "" Then
					OutErrors ("获取远程信息出错！请确定你的远程连接代码输入无误。")
					Exit Sub
				End If
				.Write "  <tr> " & vbCrLf
				.Write "    <td class=""TableTitle"" align=""center"" colspan=""2"">项 目 编 辑 -- 采集目标网站源代码&nbsp;&nbsp;&nbsp;&nbsp;"
				.Write "<Input type=""radio"" value=""0"" name=""soucode"" onClick=""soucodeid.style.display='none';""> 关闭源代码窗口&nbsp;&nbsp;<Input type=""radio"" value=""1"" name=""soucode"" onClick=""soucodeid.style.display='';"" checked> 查看源代码"
				.Write "        </td> " & vbCrLf
				.Write "  </tr> " & vbCrLf
				.Write "  <tr> " & vbCrLf
				.Write "    <td class=""TableRow1"" colspan=""2"" id='soucodeid'><textarea name='content' id='content' wrap='OFF' style='width:100%;' rows='20'>"
				.Write Server.HTMLEncode(HTTPHtmlCode)
				.Write "</textarea><div align='right'><a href=""javascript:admin_Size(-20,'content')""><img src='images/minus.gif' unselectable=on border=0></a> <a href=""javascript:admin_Size(20,'content')""><img src='images/plus.gif' unselectable=on border=0></div></td> " & vbCrLf
				.Write "  </tr> " & vbCrLf
				.Write "  <tr> " & vbCrLf
				.Write "    <td class=""TableRow2"" colspan=""2"">"
				.Write "采集的目标地址 → <a href='" & strRemoteListUrl & "' target='_blank'><font color='red'>" & strRemoteListUrl & "</font></a>"
				.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='view-source:" & strRemoteListUrl & "' target='_blank'><font color='blue'>点击查看目标源代码</font></a></td> " & vbCrLf
				.Write "  </tr> " & vbCrLf
			End If
			.Write "  <tr> " & vbCrLf
			.Write "    <td class=""TableTitle"" align=""center"" colspan=""2"">项 目 编 辑 -- " & sModuleName & "信息设置</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf

			.Write "    <td width='25%' align=""right"" class=""TableRow1""><strong>获取" & sModuleName & "标题开始代码：</strong></td> " & vbCrLf
			.Write "    <td width='75%' class=""TableRow1""><textarea name=FindInfoCode0 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(0))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>获取" & sModuleName & "标题结束代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode1 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(1))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>获取" & sModuleName & "内容开始代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode2 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(2))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>获取" & sModuleName & "内容结束代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode3 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(3))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			'--分类设置 可选项
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>" & sModuleName & "分类设置(可选项)：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2"">"
			.Write "<Input type=""radio"" value=""0"" name=""selClass"" onClick=""InfoCode4.style.display='none';InfoCode5.style.display='none';InfoCode6.style.display='none';InfoCode7.style.display='none';"" checked> 分类设置窗口&nbsp;&nbsp;"
			.Write "<Input type=""radio"" value=""1"" name=""selClass"" onClick=""InfoCode4.style.display='';InfoCode5.style.display='';InfoCode6.style.display='';InfoCode7.style.display='';""> 打开设置窗口&nbsp;&nbsp;"
			.Write "<font color='red'>* 如果第一步设置了自动归类，请设置此项</font>"
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode4"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>获取父分类名称开始代码：</strong><br><font color='blue'>不获取分类请输入“0”</font></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode4 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(4))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode5"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>获取父分类名称结束代码：</strong><br><font color='blue'>手动设置，请直接输入分类名称</font></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode5 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(5))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode6"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>获取子分类名称开始代码：</strong><br><font color='blue'>不获取分类请输入“0”</font></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode6 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(6))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode7"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>获取子分类名称结束代码：</strong><br><font color='blue'>手动设置，请直接输入分类名称</font></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode7 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(7))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			
			'--文章作者设置
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>" & sModuleName & "作者设置：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2"">"
			.Write "<Input type=""radio"" value=""0"" name=""selfont8"" onClick=""InfoCode8.style.display='none';InfoCode9.style.display='none';"" checked> 隐藏设置窗口&nbsp;&nbsp;"
			.Write "<Input type=""radio"" value=""1"" name=""selfont8"" onClick=""InfoCode8.style.display='';InfoCode9.style.display='';"">打开设置窗口&nbsp;&nbsp;"
			.Write "<font color='blue'>* 如果指定作者,开始代码填“0”，结束代码填作者名称</font>"
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode8"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong><font color=""blue"">获取" & sModuleName & "作者开始代码：</font></strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode8 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(8))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode9"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong><font color=""blue"">获取" & sModuleName & "作者结束代码：</font></strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode9 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(9))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			'--文章来源设置
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>" & sModuleName & "来源设置：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2"">"
			.Write "<Input type=""radio"" value=""0"" name=""selfont10"" onClick=""InfoCode10.style.display='none';InfoCode11.style.display='none';"" checked> 隐藏设置窗口&nbsp;&nbsp;"
			.Write "<Input type=""radio"" value=""1"" name=""selfont10"" onClick=""InfoCode10.style.display='';InfoCode11.style.display='';"">打开设置窗口&nbsp;&nbsp;"
			.Write "<font color='blue'>* 如果要指定来源,开始代码填“0”，结束代码填来源</font>"
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode10"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>获取" & sModuleName & "来源开始代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode10 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(10))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode11"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>获取" & sModuleName & "来源结束代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode11 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(11))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			'--更新时间设置
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>" & sModuleName & "更新时间设置：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2"">"
			.Write "<Input type=""radio"" value=""0"" name=""selfont12"" onClick=""InfoCode12.style.display='none';InfoCode13.style.display='none';"" checked> 隐藏设置窗口&nbsp;&nbsp;"
			.Write "<Input type=""radio"" value=""1"" name=""selfont12"" onClick=""InfoCode12.style.display='';InfoCode13.style.display='';"">打开设置窗口&nbsp;&nbsp;"
			.Write "<font color='blue'>* 如果第一步设置显示为最新时间，此设置无效</font>"
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode12"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>获取更新时间开始代码：</strong><br><font color='blue'>不设置请输入“0”</font></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode12 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(12))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode13"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>获取更新时间结束代码：</strong><br><font color='blue'>不设置请输入“0”</font></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode13 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(13))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			'--获取内容分页代码
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>是否内容分页采集：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2"">"
			.Write "<Input type=""radio"" value=""0"" name=""IsNextPage"" onClick=""InfoCode14.style.display='none';InfoCode15.style.display='none';InfoCode16.style.display='none';InfoCode17.style.display='none';"""
			If IsNextPage = 0 Then .Write " checked"
			.Write "> 不作设置&nbsp;&nbsp;"
			.Write "<Input type=""radio"" value=""1"" name=""IsNextPage"" onClick=""InfoCode14.style.display='';InfoCode15.style.display='';InfoCode16.style.display='';InfoCode17.style.display='';"""
			If IsNextPage = 1 Then .Write " checked"
			.Write "> 内容分页设置&nbsp;&nbsp;"
			.Write "<Input type=""radio"" value=""2"" name=""IsNextPage"" onClick=""InfoCode14.style.display='';InfoCode15.style.display='';InfoCode16.style.display='';InfoCode17.style.display='';"""
			If IsNextPage > 1 Then .Write " checked"
			.Write "> 分页设置&nbsp;&nbsp;"
			.Write "<font color='red'>* 如果内容有分页，请设置此项</font>"
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode14"""
			If IsNextPage = 0 Then .Write " style=""display:'none';"""
			.Write "> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>内容分页列表开始代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode14 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(14))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode15"""
			If IsNextPage = 0 Then .Write " style=""display:'none';"""
			.Write "> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>内容分页列表结束代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode15 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(15))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode16"""
			If IsNextPage = 0 Then .Write " style=""display:'none';"""
			.Write "> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>获取分页连接开始代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode16 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(16))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode17"""
			If IsNextPage = 0 Then .Write " style=""display:'none';"""
			.Write "> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>获取分页连接结束代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode17 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(17))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			'--内容过滤设置
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>内容过滤设置：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2"">"
			.Write "<Input type=""radio"" value=""0"" name=""selfont18"" onClick=""InfoCode18.style.display='none';InfoCode19.style.display='none';InfoCode20.style.display='none';InfoCode21.style.display='none';"" checked> 隐藏设置窗口&nbsp;&nbsp;"
			.Write "<Input type=""radio"" value=""1"" name=""selfont18"" onClick=""InfoCode18.style.display='';InfoCode19.style.display='';InfoCode20.style.display='';InfoCode21.style.display='';"">打开设置窗口&nbsp;&nbsp;"
			.Write "<font color='red'></font>"
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode18"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>内容过虑字符一：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode18 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(18))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode19"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>内容过虑字符二：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode19 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(19))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode20"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong>内容过虑字符三：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode20 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(20))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode21"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>内容过虑字符四：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode21 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(21))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			'--内容匹配字符过滤设置
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong>匹配字符设置：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2"">"
			.Write "<Input type=""radio"" value=""0"" name=""selfont22"" onClick=""InfoCode22.style.display='none';InfoCode23.style.display='none';InfoCode24.style.display='none';InfoCode25.style.display='none';"" checked> 隐藏设置窗口&nbsp;&nbsp;"
			.Write "<Input type=""radio"" value=""1"" name=""selfont22"" onClick=""InfoCode22.style.display='';InfoCode23.style.display='';InfoCode24.style.display='';InfoCode25.style.display='';"">打开设置窗口&nbsp;&nbsp;"
			.Write "<font color='red'></font>"
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode22"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong class='style1'>匹配字符过滤一开始代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode22 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(22))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode23"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong class='style1'>匹配字符过滤一结束代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode23 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(23))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode24"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong class='style2'>匹配字符过滤二开始代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><textarea name=FindInfoCode24 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(24))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr id=""InfoCode25"" style=""display:'none';""> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""><strong class='style2'>匹配字符过滤二开始代码：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow2""><textarea name=FindInfoCode25 rows=5 cols=80>"
			.Write Server.HTMLEncode(strFindInfoCode(25))
			.Write "</textarea></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			'--指定页面演示
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow1""><strong class='style1'>指定演示URL：</strong></td> " & vbCrLf
			.Write "    <td class=""TableRow1""><input type=""text"" name=NamedDemourl size=80 value='"
			If Len(NamedDemourl) > 0 Then
				.Write Trim(Replace(Replace(NamedDemourl, "'", ""), """", ""))
			End If
			.Write "'></td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			
			.Write "  <tr> " & vbCrLf
			.Write "    <td align=""right"" class=""TableRow2""></td> " & vbCrLf
			.Write "    <td class=""TableRow2""align=""center"">"
			.Write "      <input name=""B12"" type=""button"" class=""Button"" onclick=""javascript:history.go(-1)"" value=""返回上一页"">&nbsp;&nbsp; " & vbCrLf
			.Write "      <input name=""B22"" type=""submit"" class=""Button"" value="" 下一步 "">&nbsp;&nbsp;" & vbCrLf
			.Write "      <input name=""ShowCode"" type=""checkbox"" value=""1""> 显示源码" & vbCrLf
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <td class=""TableRow1"" colspan=""2""><b>说明：</b><br>"
			.Write "“标题”和“内容”必需截取；其它的如果不截取，请在开始代码输入“0”或者留空，在结束代码输入初始值。截取后的代码将自动清除HTML格式(内容除外)<br>"
			.Write "<b>相关标签说明：</b><br>" & sModuleName & "标题 <font style='font-family:tahoma;color:red;'>{@NewsTitle}</font>&nbsp;"
			.Write "父分类名称 <font style='font-family:tahoma;color:red;'>{@ParentName}</font>&nbsp;"
			.Write "子分类名称 <font style='font-family:tahoma;color:red;'>{@ChildName}</font>&nbsp;<br>"
			.Write "<font color='blue'>注意：开始代码必需是网页源代码中唯一的字符</font>"
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "</table> " & vbCrLf
			.Write "</form>" & vbCrLf
		End With
	End Sub
	'--采集项目第四步
	Private Sub ItemStep4()
		Dim i, showcode, NamedDemourl
		Dim tmpFindInfoCode, strEncoding
		Dim strRemoteLisCode, strRemoteListUrl
		Dim strFindListCode, strFindInfoCode
		Dim RemoveCode, startcode, lastcode
		
		Dim strNewsTitle, NewsContent, TextContent
		Dim TempHtmlCode, TempContent, strTempContent, PaginationUrl
		Dim datNewsTime, strAuthor, strComeFrom
		Dim strParent, strChild

		Dim strAddedCode, strAddedlist, AddedlistArray
		
		ItemID = Myenchiasp.ChkNumeric(Request("ItemID"))
		showcode = Myenchiasp.ChkNumeric(Request("showcode"))
		
		If Trim(Request("change")) = "yes" Then
			'--如果是更新项目，执行下面的操作
			For i = 0 To 25
				tmpFindInfoCode = tmpFindInfoCode & Request.Form("FindInfoCode" & i & "") & "$$$"
			Next
			tmpFindInfoCode = tmpFindInfoCode & "0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0"
			SQL = "SELECT ItemID,FindInfoCode,IsNextPage,NamedDemourl FROM ECCMS_NewsItem WHERE ChannelID=" & ChannelID & " And ItemID=" & ItemID
			Set Rs = CreateObject("ADODB.Recordset")
			Rs.Open SQL, MyConn, 1, 3
			If Rs.BOF And Rs.EOF Then
				OutErrors ("错误的系统参数！")
				Set Rs = Nothing
				Exit Sub
			Else
				Rs("FindInfoCode") = tmpFindInfoCode
				Rs("IsNextPage") = Myenchiasp.ChkNumeric(Request("IsNextPage"))
				Rs("NamedDemourl") = Trim(Replace(Request("NamedDemourl"), "'", ""))
				Rs.Update
			End If
			Rs.Close: Set Rs = Nothing
		End If

		'--获取项目设置
		SQL = "SELECT ItemID,AutoClass,Encoding,RemoteListUrl,RemoveCode,FindListCode,FindInfoCode,IsNextPage,RetuneClass,NamedDemourl,strReplace FROM ECCMS_NewsItem WHERE ChannelID=" & ChannelID & " And ItemID=" & ItemID
		Set Rs = MyConn.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			OutErrors ("错误的系统参数！")
			Set Rs = Nothing
			Exit Sub
		Else
			AutoClass = Rs("AutoClass")
			strEncoding = Trim(Rs("Encoding"))
			RemoteListUrl = Trim(Rs("RemoteListUrl"))
			RemoveCode = Rs("RemoveCode")
			strFindListCode = Split(Myenchiasp.ReplaceTrim(Rs("FindListCode")), "$$$")
			strFindInfoCode = Split(Myenchiasp.ReplaceTrim(Rs("FindInfoCode")), "$$$")
			IsNextPage = Rs("IsNextPage")
			RetuneClass = Rs("RetuneClass")
			If Not IsNull(Rs("NamedDemourl")) Then
				NamedDemourl = Trim(Rs("NamedDemourl"))
			End If
			If Not IsNull(Rs("strReplace")) Then
				strReplace = Rs("strReplace")
			End If
		End If
		Rs.Close: Set Rs = Nothing
		
		With Response
			.Write "<table  border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""TableBorder""> " & vbCrLf
			.Write "  <tr> " & vbCrLf
			.Write "    <th colspan=""2"">采集项目第三步</th> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			If ItemID > 0 Then
				SettingStep (ItemID)
			End If
			'--如果选择了显示源码,开始获取远程信息
			If showcode > 0 Or LCase(Trim(Request("action"))) = "demo" Then
				If Len(NamedDemourl) < 10 Then
					'--获取远程列表网页源代码
					HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(RemoteListUrl, strEncoding))
					If HTTPHtmlCode = "" Then
						OutErrors ("获取远程信息出错！请确定你的远程列表URL输入无误。")
						Exit Sub
					End If
					
					'--获取远程列表代码
					strRemoteLisCode = Myenchiasp.CutFixed(HTTPHtmlCode, strFindListCode(0), strFindListCode(1))
					strRemoteLisCode = Myenchiasp.ReplacedTrim(strRemoteLisCode)
					If strRemoteLisCode = "" Then
						OutErrors ("获取远程列表出错！请确定你的远程列表开始和结束代码输入无误。")
						Exit Sub
					End If
					'--获取列表URL
					strRemoteListUrl = Myenchiasp.CutFixed(strRemoteLisCode, strFindListCode(2), strFindListCode(3))
					strRemoteListUrl = Myenchiasp.FormatRemoteUrl(RemoteListUrl, strRemoteListUrl)
					If strRemoteListUrl = "" Then
						OutErrors ("获取远程连接出错！请确定你的连接开始和结束代码输入无误。")
						Exit Sub
					End If
				Else
					strRemoteListUrl = Trim(Replace(NamedDemourl, """", ""))
				End If
				HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(strRemoteListUrl, strEncoding))
				If HTTPHtmlCode = "" Then
					OutErrors ("获取远程信息出错！请确定你的远程连接代码输入无误。")
					Exit Sub
				End If
				
				'--获取文章标题
				strNewsTitle = Myenchiasp.CutFixed(HTTPHtmlCode, strFindInfoCode(0), strFindInfoCode(1))
				strNewsTitle = Trim(Myenchiasp.CheckHTML(strNewsTitle))
				If Len(strNewsTitle) = 0 Then
					OutErrors ("获取标题代码出错！请确定你的代码输入正确。")
					Exit Sub
				End If
				
				'--获取文章内容
				NewsContent = Myenchiasp.CutFixed(HTTPHtmlCode, strFindInfoCode(2), strFindInfoCode(3))
				If Len(NewsContent) = 0 Then
					OutErrors ("获取文章内容代码出错！请确定你的代码输入正确。")
					Exit Sub
				End If
				
				
				'--开始获取分类名称
				'--获取父分类名称
				If strFindInfoCode(4) <> "" And strFindInfoCode(4) <> "0" Then
					startcode = Replace(strFindInfoCode(4), "{@NewsTitle}", strNewsTitle)
					lastcode = Replace(strFindInfoCode(5), "{@NewsTitle}", strNewsTitle)
					strParent = Myenchiasp.CutFixed(HTTPHtmlCode, startcode, lastcode)
					strParent = Myenchiasp.CheckHTML(strParent)
				Else
					If strFindInfoCode(5) <> "" And strFindInfoCode(5) <> "0" Then
						strParent = Trim(strFindInfoCode(5))
					Else
						strParent = ""
					End If
				End If
				'strParent = Myenchiasp.CheckNostr(strParent)
				'--获取子分类名称
				If strFindInfoCode(6) <> "" And strFindInfoCode(6) <> "0" Then
					startcode = Replace(Replace(strFindInfoCode(6), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent)
					lastcode = Replace(Replace(strFindInfoCode(7), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent)
					strChild = Myenchiasp.CutFixed(HTTPHtmlCode, startcode, lastcode)
					strChild = Myenchiasp.CheckHTML(strChild)
				Else
					If strFindInfoCode(7) <> "" And strFindInfoCode(7) <> "0" Then
						strChild = Trim(strFindInfoCode(7))
					Else
						strChild = ""
					End If
				End If
				'strChild = Myenchiasp.CheckNostr(strChild)
				'--获取分类结束
				
				'--获取文章作者
				If strFindInfoCode(8) <> "" And strFindInfoCode(8) <> "0" Then
					startcode = Replace(Replace(Replace(strFindInfoCode(8), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild)
					lastcode = Replace(Replace(Replace(strFindInfoCode(9), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild)
					strAuthor = Myenchiasp.CutFixed(HTTPHtmlCode, startcode, lastcode)
					strAuthor = Myenchiasp.CheckHTML(Trim(strAuthor))
				Else
					If strFindInfoCode(9) <> "" And strFindInfoCode(9) <> "0" Then
						strAuthor = Trim(strFindInfoCode(9))
					Else
						strAuthor = "佚名"
					End If
				End If
				strAuthor = Myenchiasp.CheckNostr(strAuthor)
				
				'--获取文章来源
				If strFindInfoCode(10) <> "" And strFindInfoCode(10) <> "0" Then
					startcode = Replace(Replace(Replace(Replace(strFindInfoCode(10), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild), "{@NewsAuthor}", strAuthor)
					lastcode = Replace(Replace(Replace(Replace(strFindInfoCode(11), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild), "{@NewsAuthor}", strAuthor)
					strComeFrom = Myenchiasp.CutFixed(HTTPHtmlCode, startcode, lastcode)
					strComeFrom = Myenchiasp.CheckHTML(Trim(strComeFrom))
				Else
					If strFindInfoCode(11) <> "" And strFindInfoCode(11) <> "0" Then
						strComeFrom = Trim(strFindInfoCode(11))
					Else
						strComeFrom = "不详"
					End If
				End If
				strComeFrom = Myenchiasp.CheckNostr(strComeFrom)
				
				'--获取软件更新时间
				If strFindInfoCode(12) <> "" And strFindInfoCode(12) <> "0" Then
					startcode = Replace(Replace(Replace(Replace(Replace(strFindInfoCode(12), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild), "{@NewsAuthor}", strAuthor), "{@NewsComeFrom}", strComeFrom)
					lastcode = Replace(Replace(Replace(Replace(Replace(strFindInfoCode(13), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild), "{@NewsAuthor}", strAuthor), "{@NewsComeFrom}", strComeFrom)
					datNewsTime = Myenchiasp.CutFixed(HTTPHtmlCode, startcode, lastcode)
					datNewsTime = Myenchiasp.CheckHTML(datNewsTime)
					datNewsTime = Myenchiasp.CheckNostr(datNewsTime)
					datNewsTime = Myenchiasp.Formatime(Trim(datNewsTime))
				Else
					datNewsTime = Now
				End If
				
				'--------------获取分页内容部分开始-----------------
				Dim n, strTempArray
				If CInt(IsNextPage) > 0 And strFindInfoCode(14) <> "" And strFindInfoCode(14) <> "0" And strFindInfoCode(15) <> "" And strFindInfoCode(15) <> "0" Then
					'-- 分页方法
					
					If strFindInfoCode(16) <> "" And strFindInfoCode(16) <> "0" And strFindInfoCode(17) <> "" And strFindInfoCode(17) <> "0" Then
						If CInt(IsNextPage) = 1 Then
							'--从内容中读取分页获取列表
							strAddedCode = Myenchiasp.CutFixate(NewsContent, strFindInfoCode(14), strFindInfoCode(15))
						Else
							'--从整个HTML代码中获取列表
							strAddedCode = Myenchiasp.CutFixate(HTTPHtmlCode, strFindInfoCode(14), strFindInfoCode(15))
						End If
						
						strAddedCode = Myenchiasp.ReplaceTrim(strAddedCode)
						'--如果可以获取分页列表,开始获取分页URL
						If Len(strAddedCode) > 0 Then
							strAddedlist = Myenchiasp.FindMatch(strAddedCode, strFindInfoCode(16), strFindInfoCode(17))
							'--判断是否获取到URL
							If Len(strAddedlist) > 0 Then
								strTempContent = ""
								'--把所有URL分割成数组
								AddedlistArray = Split(strAddedlist, "|||")
								For i = 0 To UBound(AddedlistArray)
									'--格式化URL成绝对路径
									PaginationUrl = Myenchiasp.FormatRemoteUrl(strRemoteListUrl, AddedlistArray(i))
									'--只有URL和当前URL不一样的时候才采集分页信息
									If Len(PaginationUrl) > 8 And LCase(PaginationUrl) <> LCase(strRemoteListUrl) Then
										TempHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(PaginationUrl, strEncoding))
										If Len(TempHtmlCode) > 10 Then
											TempContent = Myenchiasp.CutFixed(TempHtmlCode, strFindInfoCode(2), strFindInfoCode(3))
											If Len(TempContent) > 0 Then
												'--将获取到分页内容写入到一个临时变量
												strTempContent = strTempContent & "[page_break]" & TempContent
											End If
										End If
									End If
								Next
								NewsContent = NewsContent & strTempContent
								NewsContent = Myenchiasp.CheckMatch(NewsContent, strFindInfoCode(14), strFindInfoCode(15))
								NewsContent = Replace(NewsContent, "[page_break]", "<br /><span style=""color:red;font-size:12px;font-family:tahoma;font-weight:bold;"">此处是内容分页标签：[page_break]</span><br />")
							End If
						End If
					End If
				End If
				'----------------获取分页内容结束-------------------
				'--内容过滤
				TextContent = Myenchiasp.Html2Ubb(NewsContent, RemoveCode)
				If strFindInfoCode(18) <> "" And strFindInfoCode(18) <> "0" Then
					TextContent = Replace(TextContent, strFindInfoCode(18), "")
				End If
				If strFindInfoCode(19) <> "" And strFindInfoCode(19) <> "0" Then
					TextContent = Replace(TextContent, strFindInfoCode(19), "")
				End If
				If strFindInfoCode(20) <> "" And strFindInfoCode(20) <> "0" Then
					TextContent = Replace(TextContent, strFindInfoCode(20), "")
				End If
				If strFindInfoCode(21) <> "" And strFindInfoCode(21) <> "0" Then
					TextContent = Replace(TextContent, strFindInfoCode(21), "")
				End If
				If strFindInfoCode(22) <> "" And strFindInfoCode(22) <> "0" Then
					If strFindInfoCode(23) <> "" And strFindInfoCode(23) <> "0" Then
						TextContent = Myenchiasp.CheckMatch(TextContent, strFindInfoCode(22), strFindInfoCode(23))
					End If
				End If
				If strFindInfoCode(24) <> "" And strFindInfoCode(24) <> "0" Then
					If strFindInfoCode(25) <> "" And strFindInfoCode(25) <> "0" Then
						TextContent = Myenchiasp.CheckMatch(TextContent, strFindInfoCode(24), strFindInfoCode(25))
					End If
				End If
				TextContent = Myenchiasp.FormatContentUrl(TextContent, strRemoteListUrl)
				'--软件简介替换操作
				If Len(strReplace) > 0 Then
					TextContent = Myenchiasp.ReplaceClass(TextContent, strReplace)
					strComeFrom = Myenchiasp.ReplaceClass(strComeFrom, strReplace)
				End If
				
				strNewsTitle = Myenchiasp.CheckNostr(strNewsTitle)
				
				.Write "  <tr> " & vbCrLf
				.Write "    <td class=""TableRow1"">"
				.Write "<b>" & sModuleName & "标题：</b><span class='style1'>"
				.Write strNewsTitle
				.Write "</span><br><b>更新时间：</b>"
				.Write datNewsTime
				.Write "<br><b>" & sModuleName & "作者：</b>"
				.Write strAuthor
				.Write "<br><b>" & sModuleName & "来源：</b>"
				.Write strComeFrom
				If CInt(AutoClass) > 0 Then
					.Write "<br><b>" & sModuleName & "类别：</b>"
					.Write strParent & " / " & strChild
				End If
				
				.Write "<br><b>目标地址：</b>"
				.Write "<a href='" & strRemoteListUrl & "' target='_blank'>" & strRemoteListUrl & "</a>"
				.Write "<hr style='height: 1;width: 65%;color: red;text-align:left;'>"
				.Write "<br><b  class='style3'>" & sModuleName & "内容：</b><hr style='height: 1;width: 65%;color: red;text-align:left;'><div class='style2'>"
				.Write TextContent
				.Write "</div></td> " & vbCrLf
				.Write "  </tr> " & vbCrLf
			Else
				.Write "  <tr> " & vbCrLf
				.Write "    <td class=""TableRow1"">"
				.Write "<li>恭喜您！采集项目设置全部完成。</li>"
				.Write "<li>如果要查看项目设置是否正确，请点击<a href='?action=demo&ChannelID=" & ChannelID & "&ItemID=" & ItemID & "' class=showmenu>项目演示</a> </li></td> " & vbCrLf
				.Write "  </tr> " & vbCrLf
			End If
			.Write "  <tr> " & vbCrLf
			.Write "    <td class=""TableRow2""align=""center"">"
			.Write "      <input name=""B12"" type=""button"" class=""Button"" onclick=""javascript:history.go(-1)"" value=""返回上一页"">&nbsp;&nbsp; " & vbCrLf
			.Write "      <input name=""B22"" type=""button"" class=""Button"" onclick=""window.location.href='?Channel=" & ChannelID & "';"" value=""全部设置完成"">&nbsp;&nbsp;" & vbCrLf
			.Write "      <input name=""B32"" type=""button"" class=""Button"" onclick=""window.location.href='?action=begin&ChannelID=" & ChannelID & "&ItemID=" & ItemID & "';"" value=""开始采集"">&nbsp;&nbsp; " & vbCrLf
			.Write "</td> " & vbCrLf
			.Write "  </tr> " & vbCrLf
			.Write "</table> " & vbCrLf
		End With
	End Sub
	'--数据采集
	Private Sub DataCollection()
		ItemID = Myenchiasp.ChkNumeric(Request("ItemID"))
		
		Dim ObjStream
		Dim strTemp, fromPath
		Dim RemoteListArray
		Dim d, RemoteUrl
		Dim totalnumber, CurrentPage
		
		fromPath = "tmpNewslist" & ItemID & ".dat"
		fromPath = Server.MapPath(fromPath)
		
		Set ObjStream = CreateObject("ADODB.Stream")
		ObjStream.Type = 1
		ObjStream.Mode = 3
		ObjStream.Open
		ObjStream.Position = 0
		ObjStream.LoadFromFile fromPath
		ObjStream.Position = 0
		ObjStream.Type = 2
		ObjStream.Charset = "GB2312"
		strTemp = ObjStream.ReadText()
		ObjStream.Close
		Set ObjStream = Nothing
		
		If Len(strTemp) < 10 Then
			ReturnError ("获取软件列表错误！")
			Exit Sub
		End If
		RemoteListArray = Split(strTemp, vbNewLine)
		
		totalnumber = CLng(UBound(RemoteListArray) + 1)
		
		If Not IsEmpty(Request("page")) And Trim(Request("page")) <> "" Then
			CurrentPage = CLng(Request("page"))
			d = Request("d")
		Else
			CurrentPage = 0
			d = Timer()
		End If
		
		Response.Write "<br><br>" & vbNewLine
		Response.Write "<table width='400' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbNewLine
		Response.Write "  <tr>" & vbNewLine
		Response.Write "    <td height='50'>总共需要采集 <font color='blue'><b>" & totalnumber & "</b></font> 个页面，正在采集第 <font color='red'><b>" & CurrentPage & "</b></font>  个页面…… 成功采集：<font color='blue'><b>" & Session("SucceedCount") & "</b></font></td>" & vbNewLine
		Response.Write "  </tr>" & vbNewLine
		Response.Write "  <tr>" & vbNewLine
		Response.Write "    <td><table width='100%' border='0' cellpadding='1' cellspacing='1'>" & vbNewLine
		Response.Write "      <tr>" & vbNewLine
		Response.Write "        <td style=""border: 1px #384780 solid ;background-color: #FFFFFF;""><table width='" & Fix((CurrentPage / totalnumber) * 400) & "' height='12' border='0' cellpadding='0' cellspacing='0' bgcolor=#36D91A><tr><td></td></tr></table></td>" & vbNewLine
		Response.Write "      </tr>" & vbNewLine
		Response.Write "    </table></td>" & vbNewLine
		Response.Write "  </tr>" & vbNewLine
		Response.Write "  <tr>" & vbNewLine
		Response.Write "    <td align='center'>" & FormatNumber(CurrentPage / totalnumber * 100, 2, -1) & " %</td>" & vbNewLine
		Response.Write "  </tr>" & vbNewLine
		Response.Write "</table>" & vbNewLine
		Response.Write "<table width='400' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbNewLine
		Response.Write "   <tr><td height='30' align='center'><input type='button' name='stop' value=' 立即停止采集 ' onclick=""window.location.href='" & ScriptName & "?action=yes&ChannelID=" & ChannelID & "&D=" & d & "&page=" & CurrentPage & "';"" class=button></td></tr>" & vbNewLine
		Response.Write "</table>" & vbNewLine
		Response.Flush
		
		If CurrentPage >= totalnumber Then
			Myenchiasp.DeleteFiles fromPath
			Response.Write "<meta http-equiv=""refresh"" content=""1;url='" & ScriptName & "?action=yes&ChannelID=" & ChannelID & "&page=" & CurrentPage + 1 & "&D=" & d & "'"">"
			Response.Flush
			Exit Sub
		End If
		
		RemoteUrl = RemoteListArray(CurrentPage)
		Call SaveNewsData(RemoteUrl)
		
		Response.Write "<script language='JavaScript'>" & vbNewLine
		Response.Write "function buildRefresh(){window.location.href='" & ScriptName & "?action=savenew&ChannelID=" & ChannelID & "&page=" & CurrentPage + 1 & "&ItemID=" & ItemID & "&D=" & d & "';}" & vbNewLine
		Response.Write "setTimeout('buildRefresh()'," & setInterval & ");" & vbNewLine
		Response.Write "</script>" & vbNewLine
		Response.Flush

	End Sub
	'--开始采集
	Private Sub BeginCollection()
		ItemID = Myenchiasp.ChkNumeric(Request("ItemID"))
		If ItemID = 0 Then
			OutErrors ("错误的系统参数，请输入项目ID！")
			Exit Sub
		End If
		
		ReadNewsItem (ItemID)
		
		If CInt(stopGather) = 1 Then
			OutErrors ("采集功能已经关闭，如果要进行采集，请在采集基本设置中打采集功能。\n如果还有此提示，请联系我们：www.enchiasp.cn")
			Exit Sub
		End If
		
		If CInt(StopItem) > 0 Then
			OutErrors ("此项目已经关闭，不能采集！")
			Exit Sub
		End If
		
		Response.Write TableMarquee
		Response.Flush
		
		Dim strRemoteLisCode, strRemoteListUrl
		Dim strFindListCode
		Dim i, n, strUrl
		Dim TempArray, RemoteListArray
		
		strUrl = Trim(RemoteListUrl)
		strFindListCode = Split(Myenchiasp.ReplaceTrim(FindListCode), "$$$")
		'--获取远程列表网页源代码
		If CInt(IsPagination) = 0 Then
			HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(strUrl, Encoding))
			If HTTPHtmlCode = "" Then
				OutErrors ("获取远程信息出错！请确定你的远程列表URL输入无误。")
				Exit Sub
			End If
			'--获取远程列表代码
			strRemoteLisCode = Myenchiasp.CutFixed(HTTPHtmlCode, strFindListCode(0), strFindListCode(1))
			strRemoteLisCode = Myenchiasp.ReplacedTrim(strRemoteLisCode)
			'--获取列表URL
			strRemoteListUrl = Myenchiasp.FindMatch(strRemoteLisCode, strFindListCode(2), strFindListCode(3))
		Else
			startid = Myenchiasp.ChkNumeric(startid)
			lastid = Myenchiasp.ChkNumeric(lastid)
			
			If startid = lastid Then
				strUrl = Replace(Replace(PaginalList, "*", startid), "{$pageid}", startid, 1, -1, 1)
				If Myenchiasp.CheckHTTP(strUrl) Then
					HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(strUrl, Encoding))
				Else
					HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(Trim(RemoteListUrl), Encoding))
				End If
				If HTTPHtmlCode = "" Then
					OutErrors ("获取远程信息出错！请确定你的远程列表URL输入无误。")
					Exit Sub
				End If
				'--获取远程列表代码
				strRemoteLisCode = Myenchiasp.CutFixed(HTTPHtmlCode, strFindListCode(0), strFindListCode(1))
				strRemoteLisCode = Myenchiasp.ReplacedTrim(strRemoteLisCode)
				'--获取列表URL
				strRemoteListUrl = Myenchiasp.FindMatch(strRemoteLisCode, strFindListCode(2), strFindListCode(3))
			ElseIf startid < lastid Then
				For i = startid To lastid
					If Not Response.IsClientConnected Then Response.End
					strUrl = Replace(Replace(PaginalList, "*", i), "{$pageid}", i, 1, -1, 1)
					If i < 2 Then
						If Myenchiasp.CheckHTTP(strUrl) Then
							HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(strUrl, Encoding))
						Else
							HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(Trim(RemoteListUrl), Encoding))
						End If
					Else
						HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(strUrl, Encoding))
					End If
					'--获取远程列表代码
					strRemoteLisCode = Myenchiasp.CutFixed(HTTPHtmlCode, strFindListCode(0), strFindListCode(1))
					strRemoteLisCode = Myenchiasp.ReplacedTrim(strRemoteLisCode)
					'--获取列表URL
					strRemoteListUrl = strRemoteListUrl & "|||" & Myenchiasp.FindMatch(strRemoteLisCode, strFindListCode(2), strFindListCode(3))
				Next
			Else
				For i = lastid To startid
					If Not Response.IsClientConnected Then Response.End
					strUrl = Replace(Replace(PaginalList, "*", i), "{$pageid}", i, 1, -1, 1)
					If i < 2 Then
						If Myenchiasp.CheckHTTP(strUrl) Then
							HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(strUrl, Encoding))
						Else
							HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(Trim(RemoteListUrl), Encoding))
						End If
					Else
						HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(strUrl, Encoding))
					End If
					'--获取远程列表代码
					strRemoteLisCode = Myenchiasp.CutFixed(HTTPHtmlCode, strFindListCode(0), strFindListCode(1))
					strRemoteLisCode = Myenchiasp.ReplacedTrim(strRemoteLisCode)
					'--获取列表URL
					strRemoteListUrl = Myenchiasp.FindMatch(strRemoteLisCode, strFindListCode(2), strFindListCode(3)) & "|||" & strRemoteListUrl
				Next
			End If
		End If
		Session("SucceedCount") = 0
		Dim TmpFilePath
		Dim oSteram
		Set oSteram = CreateObject("ADODB.Stream")
		TmpFilePath = "tmpNewslist" & ItemID & ".dat"
		TmpFilePath = Server.MapPath(TmpFilePath)
		
		'Set oSteram = CreateObject("ADODB.Stream")
		'---- 设置为可读可写 ---- 内容为文本
		oSteram.Mode = 3
		oSteram.Type = 2
		oSteram.Open
		oSteram.Charset = "GB2312"
		
		RemoteListArray = Split(strRemoteListUrl, "|||")
		n = UBound(RemoteListArray)
		For i = 0 To n
			If Len(RemoteListArray(i)) > 5 Then
				If Not Response.IsClientConnected Then Response.End
				If i = n Then
					oSteram.WriteText Myenchiasp.FormatRemoteUrl(strUrl, RemoteListArray(i))
				Else
					oSteram.WriteText Myenchiasp.FormatRemoteUrl(strUrl, RemoteListArray(i)) & vbNewLine
				End If
			End If
		Next
		oSteram.SaveToFile TmpFilePath, 2
		'Response.Write oSteram.ReadText()'读出全部内容，写入传送流
		oSteram.Close
		Set oSteram = Nothing
		
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM ECCMS_NewsItem WHERE ItemID= " & ItemID
		Rs.Open SQL, MyConn, 1, 3
			Rs("lastime").Value = Now()
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		
		Response.Write "<script language='JavaScript'>" & vbNewLine
		Response.Write "function reFresh(){window.location.href='" & ScriptName & "?action=savenew&ChannelID=" & ChannelID & "&ItemID=" & ItemID & "';}" & vbNewLine
		Response.Write "setTimeout('reFresh()',1000);" & vbNewLine
		Response.Write "</script>" & vbNewLine
		
	End Sub

	'--保存数据库
	Public Sub SaveNewsData(URL)
		Dim i, FileNameArray
		Dim strEncoding, strFileExt
		Dim strRemoteLisCode, strRemoteListUrl
		Dim strFindListCode, strFindInfoCode
		Dim RemoveCode, startcode, lastcode
		
		Dim strNewsTitle, NewsContent, TextContent
		Dim TempHtmlCode, TempContent, strTempContent, PaginationUrl
		Dim datNewsTime, strAuthor, strComeFrom
		Dim NewsBriefTopic, NewsRelated
		Dim NewsUploadFileList, NewsImageUrl
		Dim strParent, strChild, strParentName, strChildName

		Dim strAddedCode, strAddedlist, AddedlistArray
		Dim strFilePath, FullFilePath
		
		ItemID = Myenchiasp.ChkNumeric(Request("ItemID"))
		If ItemID = 0 Then Exit Sub
		NewsBriefTopic = 0
		ReadNewsItem (ItemID)
		
		If CInt(stopGather) = 1 Then
			ReturnError ("采集功能已经关闭，如果要进行采集，请在采集基本设置中打采集功能")
			Exit Sub
		End If
		
		strFindInfoCode = Split(Myenchiasp.ReplaceTrim(FindInfoCode), "$$$")
		strEncoding = Trim(Encoding)
		strRemoteListUrl = Trim(URL)
		
		If Len(strRemoteListUrl) < 10 Then Exit Sub
		
		HTTPHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(strRemoteListUrl, strEncoding))
		If HTTPHtmlCode = "" Then
			ReturnError ("获取远程信息出错！请确定你的远程连接代码输入无误。")
			Exit Sub
		End If
		
		'--获取文章标题
		strNewsTitle = Myenchiasp.CutFixed(HTTPHtmlCode, strFindInfoCode(0), strFindInfoCode(1))
		strNewsTitle = Trim(Myenchiasp.CheckHTML(strNewsTitle))
		If Len(strNewsTitle) = 0 Then
			ReturnError ("获取标题代码出错！请确定你的代码输入正确。")
			Exit Sub
		End If
		
		'--获取文章内容
		NewsContent = Myenchiasp.CutFixed(HTTPHtmlCode, strFindInfoCode(2), strFindInfoCode(3))
		If Len(NewsContent) = 0 Then
			ReturnError ("获取文章内容代码出错！请确定你的代码输入正确。")
			Exit Sub
		End If
		
		'--开始获取分类名称
		If CInt(AutoClass) > 0 Then
			'--获取父分类名称
			If strFindInfoCode(4) <> "" And strFindInfoCode(4) <> "0" Then
				startcode = Replace(strFindInfoCode(4), "{@NewsTitle}", strNewsTitle)
				lastcode = Replace(strFindInfoCode(5), "{@NewsTitle}", strNewsTitle)
				strParent = Myenchiasp.CutFixed(HTTPHtmlCode, startcode, lastcode)
				strParent = Myenchiasp.CheckHTML(strParent)
			Else
				If strFindInfoCode(5) <> "" And strFindInfoCode(5) <> "0" Then
					strParent = Trim(strFindInfoCode(5))
				Else
					strParent = ""
				End If
			End If
			If Len(strParent) > 22 Then strParent = ""
			'--获取子分类名称
			If strFindInfoCode(6) <> "" And strFindInfoCode(6) <> "0" Then
				startcode = Replace(Replace(strFindInfoCode(6), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent)
				lastcode = Replace(Replace(strFindInfoCode(7), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent)
				strChild = Myenchiasp.CutFixed(HTTPHtmlCode, startcode, lastcode)
				strChild = Myenchiasp.CheckHTML(strChild)
			Else
				If strFindInfoCode(7) <> "" And strFindInfoCode(7) <> "0" Then
					strChild = Trim(strFindInfoCode(7))
				Else
					strChild = ""
				End If
			End If
			If Len(strChild) > 22 Then strChild = ""
			
			strParentName = Myenchiasp.CheckNostr(strParent)
			strChildName = Myenchiasp.CheckNostr(strChild)
			If Len(RetuneClass) > 0 Then
				strParentName = Myenchiasp.ReplaceClass(strParentName, RetuneClass)
				strChildName = Myenchiasp.ReplaceClass(strChildName, RetuneClass)
			End If
			ClassID = GetClassID(ChannelID, Trim(strParentName), Trim(strChildName))
			If ClassID = 0 Then
				ReturnError ("<li>自动获取分类错误！可能这个分类是外部连接。</li><li>目标类别：" & strParent & " / " & strChild & " </li><li>当前类别：" & strParentName & " / " & strChildName & " </li>")
				Exit Sub
			End If
		Else
			Dim iRs
			Set iRs = enchiasp.Execute("SELECT ClassID FROM [ECCMS_Classify] WHERE ChannelID=" & ChannelID & " And ClassID=" & ClassID & " And child=0 And TurnLink=0")
			If iRs.BOF And iRs.EOF Then
				ReturnError ("<li>分类ID错误！可能这个分类是外部连接。</li><li>请编辑采集项目重新选择分类。</li>")
				Exit Sub
			End If
			Set iRs = Nothing
		End If
		If ClassID = 0 Then
			ReturnError ("<li>分类ID错误！可能这个分类是外部连接。</li><li>请编辑采集项目重新选择分类。</li>")
			Exit Sub
		End If
		'--获取分类结束
		
		'--获取文章作者
		If strFindInfoCode(8) <> "" And strFindInfoCode(8) <> "0" Then
			startcode = Replace(Replace(Replace(strFindInfoCode(8), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild)
			lastcode = Replace(Replace(Replace(strFindInfoCode(9), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild)
			strAuthor = Myenchiasp.CutFixed(HTTPHtmlCode, startcode, lastcode)
			strAuthor = Myenchiasp.CheckHTML(Trim(strAuthor))
		Else
			If strFindInfoCode(9) <> "" And strFindInfoCode(9) <> "0" Then
				strAuthor = Trim(strFindInfoCode(9))
			Else
				strAuthor = "佚名"
			End If
		End If
		strAuthor = Myenchiasp.CheckNostr(strAuthor)
		If Len(strAuthor) = 0 Then strAuthor = "佚名"
		
		'--获取文章来源
		If strFindInfoCode(10) <> "" And strFindInfoCode(10) <> "0" Then
			startcode = Replace(Replace(Replace(Replace(strFindInfoCode(10), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild), "{@NewsAuthor}", strAuthor)
			lastcode = Replace(Replace(Replace(Replace(strFindInfoCode(11), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild), "{@NewsAuthor}", strAuthor)
			strComeFrom = Myenchiasp.CutFixed(HTTPHtmlCode, startcode, lastcode)
			strComeFrom = Myenchiasp.CheckHTML(Trim(strComeFrom))
		Else
			If strFindInfoCode(11) <> "" And strFindInfoCode(11) <> "0" Then
				strComeFrom = Trim(strFindInfoCode(11))
			Else
				strComeFrom = "不详"
			End If
		End If
		strComeFrom = Myenchiasp.CheckNostr(strComeFrom)
		If Len(strComeFrom) = 0 Then strComeFrom = "不详"
		
		If CInt(IsNowTime) = 0 Then
			'--获取更新时间
			If strFindInfoCode(12) <> "" And strFindInfoCode(12) <> "0" Then
				startcode = Replace(Replace(Replace(Replace(Replace(strFindInfoCode(12), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild), "{@NewsAuthor}", strAuthor), "{@NewsComeFrom}", strComeFrom)
				lastcode = Replace(Replace(Replace(Replace(Replace(strFindInfoCode(13), "{@NewsTitle}", strNewsTitle), "{@ParentName}", strParent), "{@ChildName}", strChild), "{@NewsAuthor}", strAuthor), "{@NewsComeFrom}", strComeFrom)
				datNewsTime = Myenchiasp.CutFixed(HTTPHtmlCode, startcode, lastcode)
				datNewsTime = Myenchiasp.CheckHTML(datNewsTime)
				datNewsTime = Myenchiasp.CheckNostr(datNewsTime)
				datNewsTime = Myenchiasp.Formatime(Trim(datNewsTime))
			Else
				datNewsTime = Now
			End If
		Else
			datNewsTime = Now
		End If
		
		'--------------获取分页内容部分开始-----------------
		Dim n, strTempArray
		
		If CInt(IsNextPage) > 0 And strFindInfoCode(14) <> "" And strFindInfoCode(14) <> "0" And strFindInfoCode(15) <> "" And strFindInfoCode(15) <> "0" Then
			'-- 分页方法
			If strFindInfoCode(16) <> "" And strFindInfoCode(16) <> "0" And strFindInfoCode(17) <> "" And strFindInfoCode(17) <> "0" Then
				If CInt(IsNextPage) = 1 Then
					'--从内容中读取分页获取列表
					strAddedCode = Myenchiasp.CutFixate(NewsContent, strFindInfoCode(14), strFindInfoCode(15))
				Else
					'--从整个HTML代码中获取列表
					strAddedCode = Myenchiasp.CutFixate(HTTPHtmlCode, strFindInfoCode(14), strFindInfoCode(15))
				End If
				
				strAddedCode = Myenchiasp.ReplaceTrim(strAddedCode)
				'--如果可以获取分页列表,开始获取分页URL
				If Len(strAddedCode) > 0 Then
					strAddedlist = Myenchiasp.FindMatch(strAddedCode, strFindInfoCode(16), strFindInfoCode(17))
					'--判断是否获取到URL
					If Len(strAddedlist) > 0 Then
						strTempContent = ""
						'--把所有URL分割成数组
						AddedlistArray = Split(strAddedlist, "|||")
						For i = 0 To UBound(AddedlistArray)
							'--格式化URL成绝对路径
							PaginationUrl = Myenchiasp.FormatRemoteUrl(strRemoteListUrl, AddedlistArray(i))
							'--只有URL和当前URL不一样的时候才采集分页信息
							If Len(PaginationUrl) > 8 And LCase(PaginationUrl) <> LCase(strRemoteListUrl) Then
								TempHtmlCode = Myenchiasp.ReplaceTrim(Myenchiasp.GetRemoteData(PaginationUrl, strEncoding))
								If Len(TempHtmlCode) > 10 Then
									TempContent = Myenchiasp.CutFixed(TempHtmlCode, strFindInfoCode(2), strFindInfoCode(3))
									If Len(TempContent) > 0 Then
										'--将获取到分页内容写入到一个临时变量
										strTempContent = strTempContent & "[page_break]" & TempContent
									End If
								End If
							End If
						Next
						NewsContent = NewsContent & strTempContent
						NewsContent = Myenchiasp.CheckMatch(NewsContent, strFindInfoCode(14), strFindInfoCode(15))
						'NewsContent = Replace(NewsContent, "[page_break]", "")
					End If
				End If
			End If
		End If
		
		'-----------------获取分页内容结束--------------------
		'------------ 内容替换操作 -----------------------
		TextContent = Myenchiasp.Html2Ubb(NewsContent, RemoveCode)
		If strFindInfoCode(18) <> "" And strFindInfoCode(18) <> "0" Then
			TextContent = Replace(TextContent, strFindInfoCode(18), "")
		End If
		If strFindInfoCode(19) <> "" And strFindInfoCode(19) <> "0" Then
			TextContent = Replace(TextContent, strFindInfoCode(19), "")
		End If
		If strFindInfoCode(20) <> "" And strFindInfoCode(20) <> "0" Then
			TextContent = Replace(TextContent, strFindInfoCode(20), "")
		End If
		If strFindInfoCode(21) <> "" And strFindInfoCode(21) <> "0" Then
			TextContent = Replace(TextContent, strFindInfoCode(21), "")
		End If
		If strFindInfoCode(22) <> "" And strFindInfoCode(22) <> "0" Then
			If strFindInfoCode(23) <> "" And strFindInfoCode(23) <> "0" Then
				TextContent = Myenchiasp.CheckMatch(TextContent, strFindInfoCode(22), strFindInfoCode(23))
			End If
		End If
		If strFindInfoCode(24) <> "" And strFindInfoCode(24) <> "0" Then
			If strFindInfoCode(25) <> "" And strFindInfoCode(25) <> "0" Then
				TextContent = Myenchiasp.CheckMatch(TextContent, strFindInfoCode(24), strFindInfoCode(25))
				strComeFrom = Myenchiasp.ReplaceClass(strComeFrom, strReplace)
			End If
		End If
		'--软件简介替换操作
		If Len(strReplace) > 0 Then
			TextContent = Myenchiasp.ReplaceClass(TextContent, strReplace)
		End If
		'---------- 内容字符替换完成 ---------------------------------
		
		'--重新格式化文章标题
		strNewsTitle = Myenchiasp.CheckNostr(strNewsTitle)
		strNewsTitle = Myenchiasp.FormatStr(strNewsTitle)
		If CLng(AllHits) = 999 Then AllHits = Myenchiasp.GetRndHits
		'--文章关键字
		NewsRelated = strNewsTitle
		NewsRelated = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(NewsRelated, "|", ""), "[", ""), "]", ""), "<", ""), ">", ""), "'", ""), """", ""), "$", "")
		NewsRelated = Left(NewsRelated, 4) & "|" & Right(NewsRelated, 4)
		
		Response.Flush
		Response.Write "<p></p><br><table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
		Response.Write " <tr>"
		Response.Write "   <th><span id=txt1>正在采集，请稍候....</span></th>"
		Response.Write " </tr>"
		Response.Write " <tr>"
		Response.Write "   <td class=TableRow1><strong><font color=blue>" & sModuleName & "标题：</font></strong>"
		Response.Write "<font color=red>" & strNewsTitle & "</font> &nbsp;&nbsp;<br>"
		Response.Write "<strong><font color=blue>" & sModuleName & "作者：</font></strong>"
		Response.Write strAuthor
		Response.Write "<br><strong><font color=blue>" & sModuleName & "来源：</font></strong>"
		Response.Write strComeFrom
		If CInt(AutoClass) > 0 Then
			Response.Write "<br><strong><font color=blue>" & sModuleName & "类别：</font></strong>"
			Response.Write strParentName & " / " & strChildName
		End If
		Response.Write "<br><strong><font color=blue>采集时间：</font></strong>"
		Response.Write Now()
		Response.Write "<br><strong><font color=blue>目标地址：</font></strong>"
		Response.Write "<a href='" & URL & "' target=_blank>" & URL & "</a>"
		Response.Write "<div><li><span id=txt2 name=txt2 style=""font-size:9pt;color:red;"">正在采集，请稍候....</span></div>"
		Response.Write "<br><div align=center>"
		Response.Write "[<a href='?ChannelID=" & ChannelID & "'><font color=blue>停止采集</font></a>]</div>"
		Response.Write "   </td>"
		Response.Write " </tr>"
		Response.Write "</table>"
		Response.Flush
		
		'---------- 格式化内容图片URL 供下载使用----------------------
		TextContent = Myenchiasp.FormatContentUrl(TextContent, strRemoteListUrl)
		'--如果内容中有图片就保存
		If Myenchiasp.PictureEx Then
			NewsBriefTopic = 1
			If CInt(IsDown) > 0 Then
				strFilePath = ChannelDir & "UploadPic/" & Myenchiasp.BuildDatePath(PathForm)
				FullFilePath = Myenchiasp.CheckMapPath(strFilePath)
				Myenchiasp.CreatedPathEx (FullFilePath)
				Myenchiasp.MaxSize = MaxPicSize
				Myenchiasp.AllowExt = AllowPicExt
				TextContent = Myenchiasp.RemoteToLocal(TextContent, strFilePath)
				NewsUploadFileList = Myenchiasp.AllFileName
				FileNameArray = Split(NewsUploadFileList, "|")
				If UBound(FileNameArray) < 3 Then
					NewsBriefTopic = 1
				Else
					NewsBriefTopic = 2
				End If
				For i = 0 To UBound(FileNameArray)
					If Len(FileNameArray(i)) > 0 Then
						strFileExt = LCase(Myenchiasp.GetFileExtName(FileNameArray(i)))
						If strFileExt = "gif" Then
							NewsImageUrl = FileNameArray(i)
							Exit For
						End If
						If strFileExt = "jpg" Then
							NewsImageUrl = FileNameArray(i)
							Exit For
						End If
						If strFileExt = "png" Then
							NewsImageUrl = FileNameArray(i)
							Exit For
						End If
						If strFileExt = "bmp" Then
							NewsImageUrl = FileNameArray(i)
							Exit For
						End If
					End If
				Next
			End If
		Else
			NewsBriefTopic = 0
		End If
		'------------图片下载完成------------------
		
		Dim IsUpdates, blnUpdates
		Dim strInfo, strMessage
		'--开始入库
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM ECCMS_Article WHERE title='" & strNewsTitle & "'"
		Rs.Open SQL, Conn, 1, 3
		If Rs.BOF And Rs.EOF Then
			IsUpdates = True
			blnUpdates = False
			If CInt(stopGather) <> 9 Then
				ClassUpdateCount CLng(ChannelID), CLng(ClassID)
			End If
		Else
			If CInt(RepeatDeal) = 1 Then
				IsUpdates = True
			Else
				IsUpdates = False
			End If
			blnUpdates = True
		End If
		If IsUpdates Then
			If CInt(stopGather) <> 9 Then
				If Not blnUpdates Then Rs.AddNew
				Rs("ChannelID") = Myenchiasp.ChkNumeric(ChannelID)
				Rs("SpecialID") = Myenchiasp.ChkNumeric(SpecialID)
				Rs("ClassID") = Myenchiasp.ChkNumeric(ClassID)
				Rs("title") = strNewsTitle
				Rs("ColorMode") = 0
				Rs("FontMode") = 0
				Rs("content") = TextContent
				Rs("Related") = enchiasp.ChkFormStr(Left(NewsRelated, 200))
				Rs("Author") = Left(strAuthor, 100)
				Rs("ComeFrom") = Left(strComeFrom, 100)
				Rs("star") = Myenchiasp.ChkNumeric(star)
				Rs("isTop") = 0
				Rs("AllHits") = Myenchiasp.ChkNumeric(AllHits)
				Rs("DayHits") = 0
				Rs("WeekHits") = 0
				Rs("MonthHits") = 0
				Rs("HitsTime") = Now()
				Rs("WriteTime") = datNewsTime
				Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
				Rs("username") = Trim(AdminName)
				Rs("isBest") = 0
				Rs("BriefTopic") = NewsBriefTopic
				Rs("ImageUrl") = Trim(NewsImageUrl)
				Rs("UploadImage") = Trim(NewsUploadFileList)
				Rs("UserGroup") = 0
				Rs("PointNum") = 0
				Rs("isUpdate") = 1
				Rs("isAccept") = 1
				Rs("ForbidEssay") = 0
				Rs("AlphaIndex") = enchiasp.ReadAlpha(strNewsTitle)
				Rs.Update
			End If
			strMessage = "采集成功"
			strInfo = "恭喜您！采集成功"
			Session("SucceedCount") = Myenchiasp.ChkNumeric(Session("SucceedCount")) + 1
		Else
			strMessage = "采集失败"
			strInfo = "目标文章已存在，不予采集"
		End If
		Rs.Close
		Set Rs = Nothing
		
		'-- 返回提示信息
		Response.Write "<script>"
		Response.Write "txt1.innerHTML='" & strMessage & "';"
		Response.Write "txt2.innerHTML='" & strInfo & "';"
		Response.Write "</script>" & vbCrLf
		Response.Flush
	End Sub


	Private Sub DeleteItem()
		If Trim(Request("ItemID")) <> "" Then
			MyConn.Execute ("DELETE FROM ECCMS_NewsItem WHERE ItemID in (" & Request("ItemID") & ")")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		Else
			OutErrors ("请选择正确的系统参数！")
		End If
	End Sub
	Private Sub CopyNewItem()
		Dim rsCollect
		ItemID = Myenchiasp.ChkNumeric(Request("ItemID"))
		If ItemID = 0 Then
			OutErrors ("请选择正确的系统参数！")
			Exit Sub
		End If
		Set rsCollect = MyConn.Execute("SELECT * FROM ECCMS_NewsItem WHERE ChannelID=" & ChannelID & " And ItemID=" & ItemID)
		If rsCollect.BOF And rsCollect.EOF Then
			Set rsCollect = Nothing
			OutErrors ("请选择正确的系统参数！")
			Exit Sub
		Else
			Set Rs = CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM ECCMS_NewsItem WHERE (ItemID is null)"
			Rs.Open SQL, MyConn, 1, 3
			Rs.AddNew
				Rs("ItemName").Value = rsCollect("ItemName").Value
				Rs("SiteUrl").Value = rsCollect("SiteUrl").Value
				Rs("ChannelID").Value = rsCollect("ChannelID").Value
				Rs("ClassID").Value = rsCollect("ClassID").Value
				Rs("SpecialID").Value = rsCollect("SpecialID").Value
				Rs("StopItem").Value = rsCollect("StopItem").Value
				Rs("Encoding").Value = rsCollect("Encoding").Value
				Rs("IsDown").Value = rsCollect("IsDown").Value
				Rs("AutoClass").Value = rsCollect("AutoClass").Value
				Rs("PathForm").Value = rsCollect("PathForm").Value
				Rs("IsNowTime").Value = rsCollect("IsNowTime").Value
				Rs("AllHits").Value = rsCollect("AllHits").Value
				Rs("star").Value = rsCollect("star").Value
				Rs("RemoveCode").Value = rsCollect("RemoveCode").Value
				Rs("lastime").Value = Now
				Rs("RemoteListUrl").Value = rsCollect("RemoteListUrl").Value
				Rs("PaginalList").Value = rsCollect("PaginalList").Value
				Rs("IsPagination").Value = rsCollect("IsPagination").Value
				Rs("startid").Value = rsCollect("startid").Value
				Rs("lastid").Value = rsCollect("lastid").Value
				Rs("FindListCode").Value = rsCollect("FindListCode").Value
				Rs("FindInfoCode").Value = rsCollect("FindInfoCode").Value
				Rs("RetuneClass").Value = rsCollect("RetuneClass").Value
				Rs("IsNextPage").Value = rsCollect("IsNextPage").Value
				Rs("NamedDemourl").Value = rsCollect("NamedDemourl").Value
				Rs("strReplace").Value = rsCollect("strReplace").Value
			Rs.Update
			Rs.Close
			Set Rs = Nothing
		End If
		Set rsCollect = Nothing
		OutScript ("恭喜您！采集项目克隆成功。")
	End Sub

	'================================================
	'函数名：RemoveAllCache
	'作  用：删除全部缓存
	'================================================
	Public Sub RemoveAllCache()
		Dim Cacheobj
		For Each Cacheobj In Application.Contents
			Myenchiasp.DelCahe Cacheobj
			Call InnerHtml("更新 <b>" & Cacheobj & "</b> 完成")
		Next
		Call InnerHtml("当前站点全部缓存清理完成。")
	End Sub

	Public Sub InnerHtml(msg)
		Response.Write "<li>" & msg & "</li>"
		Response.Flush
	End Sub
	
End Class
%>

