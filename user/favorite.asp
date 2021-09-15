<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
<%
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
Dim Rs,SQL,i,Action
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum

Call InnerLocation("我的收藏夹")

If CInt(GroupSetting(3)) = 0 Then
	ErrMsg = ErrMsg + "<li>对不起！您没有使用收藏夹的权限，如有什么问题请联系管理员。</li>"
	Founderr = True
End If
Action = enchiasp.CheckStr(LCase(Trim(Request("action"))))
Select Case Trim(Action)
	Case "save","添加"
		Call SaveFavorite
	Case "add"
		Call AddFavorite
	Case "del"
		Call DelFavorite
	Case "清空收藏夹"
		Call DelAllFavorite
	Case Else
		Call showmain
End Select

If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
Sub showmain()
	If Founderr = True Then Exit Sub
	maxperpage = 20 '###每页显示数
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
%>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr>
		<th colspan=3>>> 我的收藏夹 <<</th>
	</tr>
	<tr>
		<td width="65%" align=center class=Usertablerow2><b class=userfont2>标题</b></td>
		<td width="23%" align=center class=Usertablerow2><b class=userfont2>收藏时间</b></td>
		<td width="12%" align=center class=Usertablerow2><b class=userfont2>操作</b></td>
	</tr>
<%
	TotalNumber = enchiasp.Execute("Select Count(FavoriteID) from ECCMS_Favorite where username='"& enchiasp.CheckStr(enchiasp.membername) &"'")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '得到总页数
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Favorite] where username='"& enchiasp.CheckStr(enchiasp.membername) &"' order by FavoriteID desc"
	Rs.Open SQL, Conn, 1, 1
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Not (Rs.bof And Rs.EOF) Then
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
		Do While Not Rs.EOF And i < CInt(maxperpage)
%>
	<tr>
		<td class=Usertablerow1><a href="<%=Rs("fondurl")%>" target=_blank><%=Server.HTMLEncode(Rs("fondtopic"))%></a></td>
		<td align=center class=Usertablerow1><%=Rs("addTime")%></td>
		<td align=center class=Usertablerow1><a href="?action=del&favid=<%=Rs("FavoriteID")%>" onclick="showClick('删除后将不能恢复，您确定要删除吗?')"><img src="images/delete.gif" width="52" height="16" border=0 alt="删除"></a></td>
	</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
	<tr>
		<td colspan=3 align=center class=Usertablerow1><%Response.Write ShowPages (CurrentPage,TotalPageNum,TotalNumber,maxperpage,"")%></td>
	</tr>
	<tr>
		<th colspan=3>>> 添加收藏 <<</th>
	</tr>
	<form name=myform method=post action="">
	<tr>
		<td colspan=3 align=center class=Usertablerow1><b class=userfont2>标题：</b><input type="text" name="fondtopic" size=20>
		<b class=userfont2>URL：</b><input type="text" name="fondurl" size=30 value="http://">
		<input type=submit name="action" value="添加" class=button> <input type=submit name="action" value="清空收藏夹" onclick="{if(confirm('清空后将不能恢复，确定清除所有的纪录吗?')){this.document.myform.submit();return true;}return false;}" class=button><br>
		<div><b>注意：</b><%If CLng(GroupSetting(5)) <> 0 Then%>你最多只能收藏 <b class=userfont1><%=GroupSetting(5)%></b> 条信息，<%End If%>请定时删除无用的信息。</div></td>
	</tr>
	</form>
</table>
<%
End Sub
'================================================
' 过程名：DelFavorite
' 作  用：删除收藏信息
'================================================
Sub DelFavorite()
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	If Not IsNumeric(Request("favid")) Then
		ErrMsg = ErrMsg + "<li>对不起！您没有使用收藏夹的权限，如有什么问题请联系管理员。</li>"
		Founderr = True
		Exit Sub
	End If
	enchiasp.Execute("Delete From ECCMS_Favorite where username='"& enchiasp.membername &"' And FavoriteID="& CLng(Request("favid")))
	Call Returnsuc("<li>记录删除成功！</li>")
End Sub
'================================================
' 过程名：DelAllFavorite
' 作  用：清空用户收藏夹
'================================================
Sub DelAllFavorite()
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	enchiasp.Execute("Delete From ECCMS_Favorite where username='"& enchiasp.membername &"'")
	Call Returnsuc("<li>收藏夹清空完成！</li>")
End Sub
'================================================
' 过程名：SaveFavorite
' 作  用：保存收藏
'================================================
Sub SaveFavorite()
	Call PreventRefresh
	If Trim(Request.Form("fondtopic")) = "" Then
		ErrMsg = ErrMsg + "<li>收藏的标题不能为空！</li>"
		Founderr = True
	End If
	If Trim(Request.Form("fondurl")) = "" Then
		ErrMsg = ErrMsg + "<li>收藏的URL不能为空！</li>"
		Founderr = True
	End If
	If CLng(GroupSetting(5)) <> 0 Then
		TotalNumber = enchiasp.Execute("Select Count(FavoriteID) from ECCMS_Favorite where username='"& enchiasp.CheckStr(enchiasp.membername) &"'")(0)
		If CLng(TotalNumber) >= CLng(GroupSetting(5)) Then
			ErrMsg = ErrMsg + "<li>对不起！你最多只能收藏" & GroupSetting(5) & "条信息。</li>"
			Founderr = True
		End If
	End  If
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Favorite] where (FavoriteID is null)"
	Rs.Open SQL, Conn, 1, 3
	Rs.Addnew
		Rs("userid") = enchiasp.memberid
		Rs("username") = enchiasp.membername
		Rs("fondtopic") = Left(enchiasp.ChkFormStr(Request.Form("fondtopic")),80)
		Rs("fondurl") = Left(enchiasp.ChkFormStr(Request.Form("fondurl")),220)
		Rs("addTime") = Now()
	Rs.update
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>恭喜您！添加收藏成功。</li>")
End Sub
'================================================
' 过程名：AddFavorite
' 作  用：添加收藏
'================================================
Sub AddFavorite()
	Dim fondtopic,fondurl
	If Trim(Request("topic")) = "" Then
		ErrMsg = ErrMsg + "<li>收藏的标题不能为空！</li>"
		Founderr = True
	Else
		fondtopic = Trim(Request("topic"))
	End If
	If CLng(GroupSetting(5)) <> 0 Then
		TotalNumber = enchiasp.Execute("Select Count(FavoriteID) from ECCMS_Favorite where username='"& enchiasp.CheckStr(enchiasp.membername) &"'")(0)
		If CLng(TotalNumber) >= CLng(GroupSetting(5)) Then
			ErrMsg = ErrMsg + "<li>对不起！你最多只能收藏" & GroupSetting(5) & "条信息。</li>"
			Founderr = True
		End If
	End  If
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Favorite] where (FavoriteID is null)"
	Rs.Open SQL, Conn, 1, 3
	Rs.Addnew
		Rs("userid") = enchiasp.memberid
		Rs("username") = enchiasp.membername
		Rs("fondtopic") = Left(enchiasp.ChkFormStr(Trim(fondtopic)),80)
		Rs("fondurl") = Left(Request.ServerVariables("HTTP_REFERER"),220)
		Rs("addTime") = Now()
	Rs.update
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>恭喜您！添加收藏成功。</li>")
End Sub
%>
<!--#include file="foot.inc"-->











