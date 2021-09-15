<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Dim keyword,readme,Tlink,strurl
Dim totalPut,totalnumber,CurrentPage,maxpagecount,maxperpage
Dim TotalPages,PageName,pagestart,pageend,pubUserName
Dim j, ii, n, face, i
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
If Not ChkAdmin("FriendLink") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
Response.Write " <tr> "
Response.Write " <th height=""22"" colspan=6><a href=""admin_link.asp""><font color=#FFFFFF>友情链接首页</font></a> | <a href=""admin_link.asp?action=add""><font color=#FFFFFF>增加新的友情链接</font></a></th>"
Response.Write " </tr>"
Response.Write " <tr> "
Response.Write " <td height=""22"" colspan=6 class=TableRow1><form name=""searchsoft"" method=""POST"" action=""admin_link.asp"" target=""main"">"
Response.Write "按名称搜索：<input class=smallInput type=""text"" name=""keyword"" size=""35""> "
Response.Write "	  条件："
Response.Write "	  <select name=field>"
Response.Write "		<option value=1 selected>网站名称</option>"
Response.Write "		<option value=2>网站 URL</option>"
Response.Write "		<option value=0>不限条件</option>"
Response.Write "	  </select> "
Response.Write "<input type=""submit"" value=""搜索链接"" name=""submit"" class=""Button"">"
Response.Write " </td></form>"
Response.Write " </tr>"
Response.Write " </table><br>"
If Request("action") = "add" Then
	Call addlink
ElseIf Request("action") = "edit" Then
	Call editlink
ElseIf Request("action") = "savenew" Then
	Call savenew
ElseIf Request("action") = "savedit" Then
	Call savedit
ElseIf Request("action") = "del" Then
	Call del
ElseIf Request("action") = "lock" Then
	Call locklink
ElseIf Request("action") = "free" Then
	Call freelink
Else
	Call linkinfo
End If
If Founderr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Private Sub addlink()
	Response.Write "<form name=myform action=""?action=savenew"" method = post>"
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write " <tr> "
	Response.Write " <th colspan=2>添加友情链接 </th>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td width=""30%"" class=TableRow1>主页名称 </td>"
	Response.Write " <td width=""70%"" class=TableRow1> "
	Response.Write " <input type=""text"" name=""name"" size=40>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>连接URL </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <input type=""text"" name=""url"" value=""http://"" size=60>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>连接LOGO地址 </td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <input type=""text"" name=""logo"" id=""ImageUrl"" value=""http://"" size=60>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>上传图片 </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <iframe name=image frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?stype=link></iframe>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>简介 </td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <textarea name=readme rows=5 cols=50></textarea>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>连接类型</td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " 文字连接<input type=""radio"" name=""islogo"" value=0 checked>&nbsp;&nbsp;LOGO连接<input type=""radio"" name=""islogo"" value=1>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>是否在首页显示</td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <input type=""radio"" name=""isIndex"" value=0 checked> 否&nbsp;&nbsp;<input type=""radio"" name=""isIndex"" value=1> 是"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>前台修改连接所用的密码 </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <input type=""text"" name=""password"" value=""" & RndPassWord & """ size=20> "
	Response.Write "<input type=checkbox name=AutoLoad value='yes'> 保存远程图片"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td height=""15"" align=center colspan=""2"" class=TableRow1> "
	Response.Write " <input type=""button"" name=""Submit1"" onclick=""javascript:history.go(-1)"" value=""返回上一页"" class=button>　"
	Response.Write " <input type=""submit"" name=""Submit"" class=""button"" value=""添 加"">"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write "</table>"
	Response.Write "</form>"
End Sub


Private Sub editlink()
	Set Rs = CreateObject("adodb.recordset")
	SQL = "select * from [ECCMS_Link] where linkid=" & Request("id")
	Rs.Open SQL, Conn, 1, 1
	Response.Write "<form name=myform action=""?action=savedit"" method=post>"
	Response.Write "<input type=hidden name=id value="
	Response.Write Request("id")
	Response.Write ">"
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write " <tr> "
	Response.Write " <th colspan=2>编辑友情链接</th>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td width=""30%"" class=TableRow1>主页名称：</td>"
	Response.Write " <td width=""70%"" class=TableRow1> "
	Response.Write " <input type=""text"" name=""name"" size=40 value="""
	Response.Write Rs("Linkname")
	Response.Write """>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>连接URL： </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <input type=""text"" name=""url"" size=60 value="""
	Response.Write Rs("Linkurl")
	Response.Write """>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>连接LOGO地址： </td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <input type=""text"" name=""logo"" id=""ImageUrl"" size=60 value="""
	Response.Write Rs("logourl")
	Response.Write """>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>上传图片 </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <iframe name=image frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?stype=link></iframe>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>简介：</td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <textarea name=readme rows=5 cols=50>"
	Response.Write Server.HTMLEncode(Rs("readme"))
	Response.Write "</textarea>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>连接类型 </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " 文字连接<input type=""radio"" name=""islogo"" value=0"
	If Rs("islogo") = 0 Then
		Response.Write " checked"
	End If
	Response.Write ">&nbsp;&nbsp;LOGO连接<input type=""radio"" name=""islogo"" value=1"
	If Rs("islogo") = 1 Then
		Response.Write " checked"
	End If
	Response.Write ">"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow1>是否在首页显示 </td>"
	Response.Write " <td class=TableRow1> "
	Response.Write " <input type=""radio"" name=""isIndex"" value=0"
	If Rs("isIndex") = 0 Then
		Response.Write " checked"
	End If
	Response.Write "> 否&nbsp;&nbsp;<input type=""radio"" name=""isIndex"" value=1"
	If Rs("isIndex") = 1 Then
		Response.Write " checked"
	End If
	Response.Write "> 是"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td class=TableRow2>友情连接密码 </td>"
	Response.Write " <td class=TableRow2> "
	Response.Write " <input type=""text"" name=""password"" size=20> <font color=blue>不修改请留空</font>"
	Response.Write "<input type=checkbox name=AutoLoad value='yes'> 保存远程图片"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td height=""15"" align=center colspan=""2"" class=TableRow1> "
	Response.Write " <div align=""center"">"
	Response.Write " <input type=""button"" name=""Submit1"" onclick=""javascript:history.go(-1)"" value=""返回上一页"" class=button>　"
	Response.Write " <input type=""submit"" name=""Submit"" class=""button"" value=""修 改"">"
	Response.Write " </div>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write "</table>"
	Response.Write "</form>"
	Rs.Close
	Set Rs = Nothing
End Sub


Private Sub linkinfo()
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write " <tr align=center>"
	Response.Write " <th width=""10%"">序 号</td>"
	Response.Write " <th width=""30%""><B>名 称</th>"
	Response.Write " <th width=""12%""><B>链接类型</th>"
	Response.Write " <th width=""30%""><B>操 作</th>"
	Response.Write " <th width=""10%""><B>状 态</th>"
	Response.Write " <th width=""8%""><B>首页</th>"
	Response.Write " </tr>"
	keyword = Trim(Request("keyword"))
	If Not IsEmpty(Request("page")) Then
		CurrentPage = CInt(Request("page"))
	Else
		CurrentPage = 1
	End If
	maxperpage = 15 '###每页显示数
	PageName = "admin_link.asp"
	Set Rs = Server.CreateObject("adodb.recordset")
	If Not IsNull(keyword) And keyword <> "" Then
		keyword = Replace(Replace(Replace(keyword, "'", "‘"), "<", "&lt;"), ">", "&gt;")
		If CInt(Request("field")) = 1 Then
			SQL = "SELECT * FROM [ECCMS_Link] WHERE LinkName LIKE '%" & keyword & "%'"
		ElseIf CInt(Request("field")) = 2 Then
			SQL = "SELECT * FROM [ECCMS_Link] WHERE Linkurl LIKE '%" & keyword & "%'"
		Else
			SQL = "SELECT * FROM [ECCMS_Link] WHERE LinkName LIKE '%" & keyword & "%' Or Linkurl LIKE '%" & keyword & "%'"
		End If
		SQL = SQL & " ORDER BY linkid DESC"
	Else
		SQL = " SELECT * FROM [ECCMS_Link] ORDER BY linkid DESC"
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	Rs.Open SQL, Conn, 1, 1
	If Not (Rs.bof Or Rs.EOF) Then
		Rs.pagesize = maxperpage
		maxpagecount = Rs.pagecount '###记录总页数
		totalnumber = CInt(Rs.recordcount) '###记录总数
		Rs.absolutepage = CurrentPage '###当前页数
		ii = 0
		Rem #######显示多少页##########
		pagestart = CurrentPage - 3
		pageend = CurrentPage + 3
		Rem ##########################
		n = CurrentPage
		If pagestart < 1 Then
			pagestart = 1
		End If
		If pageend > maxpagecount Then
			pageend = maxpagecount
		End If
		If n < maxpagecount Then
			n = maxpagecount
		End If
		j = (CurrentPage - 1) * maxperpage + 1
		Do While Not Rs.EOF And ii < Rs.pagesize
			Response.Write " <tr align=center>"
			Response.Write " <td height=25 class=TableRow1><font color=red>"
			Response.Write j
			Response.Write "</font></td>"
			Response.Write " <td class=TableRow1><a href="
			Response.Write Rs("Linkurl")
			Response.Write " target=_blank>"
			Response.Write Rs("Linkname")
			Response.Write "</a></td>"
			Response.Write " <td class=TableRow1>"
			If Rs("islogo") = 1 Then
				Response.Write "LOGO链接"
			Else
				Response.Write "文字链接"
			End If
			Response.Write "</td>"
			Response.Write " <td class=TableRow1> <a href=""admin_link.asp?action=edit&id="
			Response.Write Rs("Linkid")
			Response.Write """><u>编辑</u></a> | <a href=""admin_link.asp?action=lock&id="
			Response.Write Rs("linkid")
			Response.Write """><u>锁定</u></a> | <a href=""admin_link.asp?action=free&id="
			Response.Write Rs("linkid")
			Response.Write """><u>解锁</u></a> | <a href=""admin_link.asp?action=del&id="
			Response.Write Rs("linkid")
			Response.Write """ onclick=""{if(confirm('此操作将删除本友情连接\n 您确定执行此操作吗?')){this.document.myform.submit();return true;}return false;}""><u>删除</u></a></td>"
			Response.Write " <td class=TableRow1>"
			If Rs("isLock") = 0 Then
				Response.Write "正常"
			Else
				Response.Write "<font color=red>锁定</font>"
			End If
			Response.Write "</td>"
			Response.Write " <td class=TableRow1>"
			If Rs("isIndex") = 0 Then
				Response.Write "<font color=red>×</font>"
			Else
				Response.Write "<font color=blue>√</font>"
			End If
			Response.Write "</td>"
			Response.Write " </tr>"
			Rs.movenext
			j = j + 1
			ii = ii + 1
		Loop
		Rs.Close
		Set Rs = Nothing
	Else
		Response.Write ("<tr><td colspan=5 class=TableRow2>暂时还没有任何友情连接</td></tr>")
	End If
	Response.Write "<tr><td colspan=6 class=TableRow2>"
	Call showpage
	Response.Write "</td></tr>"
	Response.Write "</table>"
End Sub

Private Sub savenew()
	Dim sUploadDir,strUploadDir,SaveFileType,SaveFilesName
	Dim password,strLogo
	password = md5(Request("password"))
	strLogo = Trim(Request.Form("logo"))
	If Trim(Request("url")) <> "" And Trim(Request("readme")) <> "" And Trim(Request("name")) <> "" Then
		If Trim(Request("AutoLoad")) = "yes" Then
			sUploadDir = "../link/UploadPic/"
			strUploadDir = CreatePath(sUploadDir)
			SaveFileType = Mid(strLogo, InStrRev(strLogo, ".") + 1)
			SaveFilesName = GetRndFileName(SaveFileType)
			If SaveRemotePic(sUploadDir & strUploadDir & SaveFilesName, strLogo) = True Then
				strLogo = "link/UploadPic/" & strUploadDir & SaveFilesName
			Else
				strLogo = strLogo
			End If
		End If
		Set Rs = CreateObject("adodb.recordset")
		SQL = "select * from [ECCMS_Link] where (Linkid is null)"
		Rs.Open SQL, Conn, 1, 3
		Rs.addnew
			Rs("Linkname").Value =  enchiasp.CheckStr(Request.Form("name"))
			Rs("readme").Value = enchiasp.CheckStr(Request.Form("readme"))
			Rs("logourl").Value = Trim(Request.Form("logo"))
			Rs("Linkurl").Value = Request.Form("url")
			Rs("password").Value = password
			Rs("islogo").Value = Request.Form("islogo")
			Rs("isLock").Value = 0
			Rs("isIndex").Value = Request.Form("isIndex")
			Rs.Update
		Rs.Close
		Set Rs = Nothing
		Succeed("添加成功，请继续其他操作。")
	Else
		ErrMsg = ErrMsg + "<br>" + "请输入完整友情链接信息。"
		Founderr = True
		Exit Sub
	End If
End Sub

Private Sub savedit()
	Dim sUploadDir,strUploadDir,SaveFileType,SaveFilesName
	Dim strLogo
	strLogo = Trim(Request.Form("logo"))
	If Trim(Request("AutoLoad")) = "yes" Then
		sUploadDir = "../link/UploadPic/"
		strUploadDir = CreatePath(sUploadDir)
		SaveFileType = Mid(strLogo, InStrRev(strLogo, ".") + 1)
		SaveFilesName = GetRndFileName(SaveFileType)
		If SaveRemotePic(sUploadDir & strUploadDir & SaveFilesName, strLogo) = True Then
			strLogo = "link/UploadPic/" & strUploadDir & SaveFilesName
		Else
			strLogo = strLogo
		End If
	End If
	Set Rs = CreateObject("adodb.recordset")
	SQL = "select * from [ECCMS_Link] where Linkid=" & Request("id")
	Rs.Open SQL, Conn, 1, 3
		Rs("Linkname").Value = Trim(Request.Form("name"))
		Rs("readme").Value = Trim(Request.Form("readme"))
		Rs("logourl").Value = strLogo
		Rs("Linkurl").Value = Trim(Request.Form("url"))
		If Trim(Request("password")) <> "" Then Rs("password").Value = Request.Form("password")
		Rs("islogo").Value = Request.Form("islogo")
		Rs("isIndex").Value = Request.Form("isIndex")
		Succeed ("更新成功，请继续其他操作。")
		Rs.Update
	Rs.Close
	Set Rs = Nothing
End Sub
Private Sub del()
	Dim id
	id = Request("id")
	SQL = "delete from [ECCMS_Link] where Linkid=" + id
	Conn.Execute (SQL)
	Succeed ("删除成功，请继续其他操作。")
End Sub
Private Sub locklink()
	Dim id
	id = Request("id")
	Conn.Execute ("update [ECCMS_Link] set islock=1 where Linkid=" + id)
	Succeed ("锁定操作成功，请继续其他操作。")
End Sub
Private Sub freelink()
	Dim id
	id = Request("id")
	Conn.Execute ("update [ECCMS_Link] set islock=0 where Linkid=" + id)
	Succeed ("解除锁定操作成功，请继续其他操作。")
End Sub
Private Function SaveRemotePic(s_LocalFileName, s_RemoteFileUrl)
	Dim Ads
	Dim Retrieval
	Dim GetRemoteData
	Dim bError
	bError = False
	SaveRemotePic = False
	On Error Resume Next
	Set Retrieval = CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", s_RemoteFileUrl, False
		.Send
		If .readyState <> 4 Then Exit Function
		If .Status > 300 Then Exit Function
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	
	Set Ads = CreateObject("Adodb.Stream")
	With Ads
		.type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile Server.MapPath(s_LocalFileName), 2
		.Cancel
		.Close
	End With
	Set Ads = Nothing
	If Err.Number = 0 And bError = False Then
		SaveRemotePic = True
	Else
		Err.Clear
	End If
End Function
Private Function GetRndFileName(ByVal sExt)
	Dim sRnd
	Randomize
	sRnd = Int(900 * Rnd) + 100
	GetRndFileName = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & sRnd & "." & sExt
End Function
Private Sub showpage()
	Response.Write "<table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
	Response.Write " <tr><form method=""POST"" action="""
	Response.Write PageName
	Response.Write """ >"
	Response.Write " <td class=""td1"" align=""center"">共有"
	Response.Write totalnumber
	Response.Write "个 <a href="
	Response.Write PageName
	Response.Write "?page=1 title=返回第一页><font face=""Webdings"">97</font></a> "
	For i = pagestart To pageend
		If i = 0 Then
			i = 1
		End If
		strurl = "<a href=" & PageName & "?page=" & i & " title=第" & i & "页>[" & i & "]</a>"
		Response.Write strurl
		Response.Write " "
	Next
	Response.Write "<a href="
	Response.Write PageName
	Response.Write "?page="
	Response.Write maxpagecount
	Response.Write " title=尾页><font face=""Webdings"">8:</font></a> 页次:<font color=red>"
	Response.Write CurrentPage
	Response.Write "</font> / "
	Response.Write maxpagecount
	Response.Write "页 每页:"
	Response.Write maxperpage
	Response.Write " 转到:<select name='page' align=""absmiddle"" size='1' style=""font-size: 9pt"" onChange='javascript:submit()'>"
	Response.Write " "
	For i = 1 To n
		Response.Write " <option value='"
		Response.Write i
		Response.Write "' "
		If CurrentPage = CInt(i) Then
			Response.Write " selected "
		End If
		Response.Write ">第"
		Response.Write i
		Response.Write "页</option>"
		Response.Write " "
	Next
	Response.Write " </select>"
	Response.Write " </td></form>"
	Response.Write " </tr>"
	Response.Write " </table>"
End Sub
Public Function RndPassWord()
	Dim num1,rndnum
	Randomize
	Do While Len(rndnum) < 8
		num1 = CStr(Chr((57 - 48) * rnd + 48))
		rndnum = rndnum & num1
	loop
	RndPassWord = rndnum
End Function
%>
