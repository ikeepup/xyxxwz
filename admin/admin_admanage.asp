<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="include/cls_admanage.asp"-->
<%
Admin_header
'=====================================================================
' 软件名称：恩池网站管理系统--广告管理
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Response.Write "<table border=0 align=center cellpadding=2 cellspacing=1 class=tableBorder>" & vbNewLine
Response.Write "  <tr>" & vbNewLine
Response.Write "    <th><a href='admin_admanage.asp' Class=showtitle><strong>广告管理</strong></a></th></tr>" & vbNewLine
Response.Write "  <tr height=25>" & vbNewLine
Response.Write "    <td class=TableRow1><B>说明：</B><br> " & vbNewLine 
Response.Write "①、本系统所有广告代码都将生成JS文件，文件位于/adfile/目录下面，请在隐藏删除操作后<font color=red>重新生成相关JS</font>文件；<br>" & vbNewLine
Response.Write "②、你可以任意增加广告位（商业版）,然后在模板的相应位置调用此JS文件即可。" & vbNewLine
Response.Write "    </td>" & vbNewLine
Response.Write "  </tr>" & vbNewLine
Response.Write "  <tr height=25>" & vbNewLine
Response.Write "    <td class=TableRow2><B>广告导航：</B> <A HREF='admin_admanage.asp'>广告管理首页</A> |" & vbNewLine 
Response.Write "    <a href='admin_admanage.asp?action=add' class=showmeun>添加广告</a> |" & vbNewLine
Response.Write "    <a href='admin_admanage.asp?action=board' class=showmeun>添加广告位</a> |" & vbNewLine
Response.Write "    <a href='admin_admanage.asp?action=create&stype=all&boardid=0'><span style=""color: red;"">生成所有广告的JS文件</span></a> |" & vbNewLine
Response.Write "    <a href='Admin_UploadFile.Asp?ChannelID=0&UploadDir=UploadPic'>上传文件管理</a>" & vbNewLine
Response.Write "    </td>" & vbNewLine
Response.Write "  </tr>" & vbNewLine
Response.Write "</table>" & vbNewLine
Response.Write "<br>" & vbNewLine



Dim Action,isEdit,AdvertiseID
Action = LCase(Request("action"))
If Not ChkAdmin("Advertise") Then
	Server.Transfer("showerr.asp")
	Response.End
End If

Select Case Trim(Action)
Case "save"
	Call saveAdvertise
Case "modify"
	Call modifyAdvertise
Case "add"
	isEdit = False
	Call EditAdvertise(isEdit)
Case "edit"
	isEdit = True
	Call EditAdvertise(isEdit)
Case "del"
	Call DelAdvertise
Case "board"
	Call boardlist
Case "saveboard"
	Call saveboard
Case "delboard"
	Call delboard
Case "create"
	Call CreateBoardJs
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
	If LCase(Request("act")) = "lock" Then
		Call isLock
	End If
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<th width='20%' nowrap>网站名称</th>" & vbNewLine
	Response.Write "	<th width='50%'>广告图片代码</th>" & vbNewLine
	Response.Write "	<th width='10%' nowrap> 广告性质 </th>" & vbNewLine
	Response.Write "	<th width='10%' nowrap> 操作选项 </th>" & vbNewLine
	Response.Write "	<th width='10%' nowrap>状 态</th>" & vbNewLine
	Response.Write "</tr>" & vbNewLine

	Dim intWidth,intHeight
	Dim CurrentPage,page_count,totalrec,Pcount,maxperpage
	Dim strClass
	maxperpage = 20 '###每页显示数
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CLng(CurrentPage)
	End If
	SQL = "SELECT * FROM ECCMS_Adlist ORDER BY id DESC"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open SQL,conn,1,1
	If Not (Rs.EOF And Rs.BOF) Then
		Rs.PageSize = maxperpage
		Rs.AbsolutePage = CurrentPage
		page_count = 0
		totalrec = Rs.recordcount
		Do While Not Rs.EOF And (Not page_count = Rs.PageSize)
			page_count = page_count + 1
			If Not Response.IsClientConnected Then Response.End
			If (page_count mod 2) = 0 Then
				strClass = "class=TableRow1"
			Else
				strClass = "class=TableRow2"
			End If
			If Rs("width") > 468 Then
				intWidth = 486
			Else
				intWidth = Rs("width")
			End If
			If Rs("height") > 60 Then
				intHeight = 60
			Else
				intHeight = Rs("height")
			End If
			Response.Write "<tr>"
			Response.Write "	<td " & strClass & "><a href='?action=edit&id="
			Response.Write Rs("id")
			Response.Write "' title='点击此处修改该广告'>"
			Response.Write Rs("title")
			Response.Write "	</a></td>"
			Response.Write "	<td align=center " & strClass & ">"
			If Rs("flag") = 5 Then
				Response.Write Left(Server.HTMLEncode(Rs("AdCode")),200)
			Else
				If Rs("isFlash") = 1 Then
					Response.Write "<embed src=" & enchiasp.ReadFileUrl(Rs("picurl")) & " quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='" & intwidth & "' height='" & intHeight & "'></embed>"
				Else
					Response.Write "<a href='" & Rs("url") & "' target=_blank><img src='" & enchiasp.ReadFileUrl(Rs("picurl")) & "' width='" & intwidth & "' height='" & intHeight & "' border=0 alt='" & Rs("Readme") & "'></a>"
				End If
			End If
			Response.Write "	</td>"
			Response.Write "	<td align=center nowrap " & strClass & "><a href='admin_admanage.asp?action=create&boardid=" & Rs("boardid") & "' title='点击生成广告JS文件'>"
			Response.Write ReadBoardName(Rs("boardid"))
			Response.Write "</a><br><br style=""OVERFLOW: hidden; LINE-HEIGHT: 5px""><font color=blue>"
			Select Case Cint(Rs("flag"))
			Case 1
				Response.Write "漂浮广告"
			Case 2
				Response.Write "左边固定广告"
			Case 3
				Response.Write "右边固定广告"
			Case 4
				Response.Write "弹出广告"
			Case 5
				Response.Write "代码广告"
			Case Else
				Response.Write "普通广告"		
			End Select
			Response.Write "	</font></td>" & vbNewLine
			Response.Write "	<td align=center " & strClass & "><a href='?action=edit&id=" & Rs("id") & "'>编辑广告</a><br><br style=""OVERFLOW: hidden; LINE-HEIGHT: 5px"">" & vbNewLine
			Response.Write "	<a href='?action=del&id=" & Rs("id") & "' onclick=""{if(confirm('广告删除后将不能恢复，您确定要删除该广告吗?')){return true;}return false;}"">删除广告</a></td>" & vbNewLine
			Response.Write "	<td align=center " & strClass & ">"
			If Rs("IsLock") <> 0 Then
				Response.Write "<a href='?act=lock&isLock=0&id="& Rs("id") &"' title='点击此处解除隐藏' onclick=""{if(confirm('您确定要解除隐藏吗?')){return true;}return false;}""><font color=red>"
				Response.Write "隐藏"
				Response.Write "</font></a>"
			Else
				Response.Write "<a href='?act=lock&isLock=1&id="& Rs("id") &"' title='点击此处隐藏广告' onclick=""{if(confirm('您确定要隐藏该广告吗?')){return true;}return false;}"">正常</a>"
			End If
			Response.Write "	</td>" & vbNewLine
			Response.Write "</tr>" & vbNewLine
			Rs.movenext
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	If totalrec Mod maxperpage = 0 Then
		Pcount =  totalrec \ maxperpage
	Else
		Pcount =  totalrec \ maxperpage+1
	End If
	If page_count = 0 Then CurrentPage = 0
	Response.Write "<tr height=20>" & vbNewLine
	Response.Write "	<td colspan=6 class=tablerow2>"
	Response.Write showpages(CurrentPage,Pcount,totalrec,maxperpage,"")
	Response.Write "</td>"
	Response.Write "</tr>" & vbNewLine
	Response.Write "</table>"
End Sub

Function ReadBoardName(Byval boardid)
	Dim rsBoard
	Set rsBoard = enchiasp.Execute("SELECT BoardName FROM ECCMS_Adboard WHERE boardid="& boardid)
	If rsBoard.BOF And rsBoard.EOF Then
		Set rsBoard = Nothing
		ReadBoardName = ""
		Exit Function
	End If
	ReadBoardName = rsBoard("BoardName")
	Set rsBoard = Nothing
End Function

Public Sub CreateBoardJs()
	Dim rsBoard,sqlBoard,adenchiasp
	If LCase(Request("stype")) = "all" Then
		sqlBoard = " ORDER BY boardid DESC"
	Else
		sqlBoard = " WHERE boardid=" & Request("boardid") & " ORDER BY boardid DESC"
	End If
	If Not IsNumeric(Request.Form("boardid")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>广告位ID请输入整数！</li>"
		Exit Sub
	End If
	Set rsBoard = enchiasp.Execute("SELECT boardid FROM ECCMS_Adboard " & sqlBoard & "")
	If rsBoard.BOF And rsBoard.EOF Then
		Set rsBoard = Nothing
		Exit Sub
	End If
	Set adenchiasp = New Admanage_Cls
	Do While Not rsBoard.EOF
		adenchiasp.adboardid = rsBoard("boardid")
		adenchiasp.CreateJsFile
		rsBoard.movenext
	Loop
	Set adenchiasp = Nothing
	rsBoard.Close:Set rsBoard = Nothing
	Succeed("<li>恭喜您！生成广告JS文件完成。</li>")
End Sub

Sub boardlist()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<th>广告位名称</th>" & vbNewLine
	Response.Write "	<th>JS文件名</th>" & vbNewLine
	Response.Write "	<th>最多广告数</th>" & vbNewLine
	Response.Write "	<th>广告位价格</th>" & vbNewLine
	Response.Write "	<th>管理操作</th>" & vbNewLine
	Response.Write "</tr>" & vbNewLine

	Dim CurrentPage,page_count,totalrec,Pcount,maxperpage
	Dim newboardid
	maxperpage = 20 '###每页显示数
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CLng(CurrentPage)
	End If
	SQL = "SELECT boardid,BoardName,AdRate,FileName,Maxads FROM ECCMS_Adboard ORDER BY boardid ASC"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open SQL,conn,1,1
	If Not (Rs.EOF And Rs.BOF) Then
		Rs.PageSize = maxperpage
		Rs.AbsolutePage = CurrentPage
		page_count = 0
		totalrec = Rs.recordcount
		Do While Not Rs.EOF And (Not page_count = Rs.PageSize)
			page_count = page_count + 1
			Response.Write "<form name=form" & Rs("boardid") & " method=post action=admin_admanage.asp>" & vbNewLine
			Response.Write "<tr align=center>" & vbNewLine
			Response.Write "<input type=hidden name=action value='saveboard'>" & vbNewLine
			Response.Write "<input type=hidden name=boardid value='" & Rs("boardid") & "'>" & vbNewLine
			Response.Write "	<td class=tablerow1><input type=text name=BoardName size=30 value='" & Rs("BoardName") & "'></td>" & vbNewLine
			Response.Write "	<td class=tablerow1><input type=text name=FileName size=15 value='" & Rs("FileName") & "'></td>" & vbNewLine
			Response.Write "	<td class=tablerow1><input type=text name=Maxads size=8 value='" & Rs("Maxads") & "'> 条</td>" & vbNewLine
			Response.Write "	<td class=tablerow1><input type=text name=AdRate size=8 value='" & Rs("AdRate") & "'> 元/月</td>" & vbNewLine
			Response.Write "	<td class=tablerow1><input class=Button type=submit name=act value='修改'>" & vbNewLine
			Response.Write "	<input class=Button type=submit name=act value='生成广告JS' onclick=""document.form" & Rs("boardid") & ".action.value='create';"">" & vbNewLine
			Response.Write "	<input class=Button type=submit name=submit2 value='删除' "
			If Rs("boardid") < 4 Then Response.Write " disabled "
			Response.Write "onclick=""document.form" & Rs("boardid") & ".action.value='delboard';return confirm('本次操作将删除此广告位下所有的广告信息。\n\n确定要删除当前广告位吗？')""></td>" & vbNewLine
			Response.Write "</tr>" & vbNewLine
			Response.Write "</form>" & vbNewLine
			Rs.movenext
		Loop
	End If
	Rs.close:Set Rs = nothing
	If totalrec Mod maxperpage = 0 Then
		Pcount =  totalrec \ maxperpage
	Else
		Pcount =  totalrec \ maxperpage+1
	End If
	If page_count = 0 Then CurrentPage = 0
	Response.Write "	<tr height=20>" & vbNewLine
	Response.Write "		<td colspan=6 class=tablerow2>"
	Response.Write showpages(CurrentPage,Pcount,totalrec,maxperpage,"&action=board")
	Response.Write "</td>"
	Response.Write "	</tr>" & vbNewLine
	
	Set Rs = enchiasp.Execute("SELECT MAX(boardid) FROM ECCMS_Adboard")
	If Rs.BOF And Rs.EOF Then
		newboardid = 1
	Else
		newboardid = Rs(0) + 1
	End If
	If IsNull(newboardid) Then newboardid = 1
	Rs.close:Set Rs = nothing

	Response.Write "<form name=addform method=post action=admin_admanage.asp>" & vbNewLine
	Response.Write "<input type=hidden name=action value='saveboard'>" & vbNewLine
	Response.Write "<input type=hidden name=boardid value='" & newboardid & "'>" & vbNewLine
	Response.Write "<tr align=center>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=text name=BoardName size=30></td>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=text name=FileName size=15></td>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=text name=Maxads size=8> 条</td>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=text name=AdRate size=8> 元/月</td>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=submit name=act value='添加广告位'  class=Button></td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write "</form>" & vbNewLine
	Response.Write "	<tr height=20>" & vbNewLine
	Response.Write "		<td colspan=6 class=tablerow2>"
	Response.Write "<b>说明：</b><br>①、广告位名称任意填写；<br>②、JS文件名的扩展名一定要是*.js,文件路径相对于站点根目录下在adfile目录；<br>"
	Response.Write "③、最多广告数，是指在此广告位显示多少条广告；<br>④、广告位价格，供做广告者参考；<br>"
	Response.Write "⑤、JS文件的调用方法：&lt;script src=/adfile/ad.js&gt;&lt;/script&gt;"
	Response.Write "</td>" & vbNewLine
	Response.Write "	</tr>" & vbNewLine
	Response.Write "</table>" & vbNewLine
End Sub

Sub saveboard()
	If Trim(Request.Form("BoardName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>广告位名称不能为空！</li>"
	End If
	If Not IsNumeric(Request.Form("boardid")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>广告位ID请输入整数！</li>"
	End If
	If Not IsNumeric(Request.Form("AdRate")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>广告位价格请输入整数！</li>"
	End If
	If Not IsNumeric(Request.Form("Maxads")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>显示最多广告数请输入整数！</li>"
	End If
	If Trim(Request.Form("FileName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>JS文件名不能为空！</li>"
	End If
	If LCase(Right(Trim(Request.Form("FileName")),3)) <> ".js" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请输入正确的JS文件名，扩展名一定要是*.js！</li>"
	End If

	If Founderr = True Then Exit Sub
	If Trim(Request.Form("act")) = "修改" Then
		enchiasp.Execute ("update ECCMS_Adboard set BoardName='"& Request.Form("BoardName") & "',FileName='"& Request.Form("FileName") & "',Maxads="& Request.Form("Maxads") & ",AdRate="& Request.Form("AdRate") & " where boardid="& CLng(Request.Form("boardid")))
		Succeed("<li>恭喜您！修改广告位成功</li>")
	Else
		SQL = "Insert into ECCMS_Adboard (boardid,BoardName,Readme,AdRate,FileName,Maxads,useup) values (" &_
		""& Request.Form("boardid") & "," &_
		"'"& Request.Form("BoardName") & "'," &_
		"''," &_
		Request.Form("AdRate") & "," &_
		"'"& Request.Form("FileName") & "'," &_
		Request.Form("Maxads") & "," &_
		"0)"
		enchiasp.Execute(SQL)
		Succeed("<li>恭喜您！添加新的广告位成功</li>")
	End If
End Sub

Sub delboard()
	If Not IsNumeric(Request.Form("boardid")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>广告位ID请输入整数！</li>"
		Exit Sub
	End If
	On Error Resume Next
	Set Rs = enchiasp.Execute("SELECT FileName FROM ECCMS_Adboard WHERE boardid=" & Request("boardid"))
	If Not (Rs.BOF And Rs.EOF) Then
		enchiasp.FileDelete("../adfile/" & Rs("FileName"))
	End If
	Set Rs = Nothing
	enchiasp.Execute("DELETE FROM ECCMS_Adboard WHERE boardid="& CLng(Request.Form("boardid")))
	enchiasp.Execute("DELETE FROM ECCMS_Adlist WHERE boardid="& CLng(Request.Form("boardid")))
	Response.Redirect  Request.ServerVariables("HTTP_REFERER")
End Sub

Sub isLock()
	If Trim(Request("id")) = "" Or Trim(Request("isLock")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	enchiasp.Execute ("update ECCMS_Adlist set isLock = "& CInt(Request("isLock")) &" where id=" & CLng(Request("id")))
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Sub DelAdvertise()
	If Not IsNumeric(Request("id")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ID请输入整数！</li>"
		Exit Sub
	End If

	enchiasp.Execute("DELETE FROM ECCMS_Adlist WHERE id="& CLng(Request("id")))
	Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Function FrontAdvertise(id)
	Dim Rss, SQL
	SQL = "SELECT TOP 1 id,title FROM ECCMS_Adlist WHERE id < " & id & " ORDER BY id DESC"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		FrontAdvertise = "已经没有了"
	Else
		FrontAdvertise = "<a href=?action=view&id=" & Rss("id") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function

Function NextAdvertise(id)
	Dim Rss, SQL
	SQL = "SELECT TOP 1 id,title FROM ECCMS_Adlist WHERE id > " & id & " ORDER BY id ASC"
	Set Rss = enchiasp.Execute(SQL)
	If Rss.EOF And Rss.bof Then
		NextAdvertise = "已经没有了"
	Else
		NextAdvertise = "<a href=?action=view&id=" & Rss("id") & ">" & Rss("title") & "</a>"
	End If
	Rss.Close
	Set Rss = Nothing
End Function
Private Sub EditAdvertise(isEdit)
	Dim EditTitle
	If isEdit Then
		SQL = "select * from ECCMS_Adlist where id=" & Request("id")
		Set Rs = enchiasp.Execute(SQL)
		EditTitle = "编辑广告"
	Else
		EditTitle = "添加新的广告"
	End If
%>
<script language = JavaScript>
function showsetting(myform){
	var tab = myform.flag.selectedIndex;
	if(tab==5)  {
		flagsetting1.style.display='none';
		flagsetting2.style.display='none';
		flagsetting3.style.display='none';
		flagsetting4.style.display='none';
		flagsetting5.style.display='none';
		flagsetting6.style.display='none';
		flagsetting7.style.display='';
	}
	if(tab==1||tab==3||tab==4)    {
		flagsetting1.style.display='';
		flagsetting2.style.display = '';
		flagsetting3.style.display='';
		flagsetting4.style.display='';
		flagsetting5.style.display='';
		flagsetting6.style.display='';
		flagsetting7.style.display='none';
	}
	if(tab==2)    {
		flagsetting1.style.display='';
		flagsetting2.style.display = '';
		flagsetting3.style.display='';
		flagsetting4.style.display='';
		flagsetting5.style.display='';
		flagsetting6.style.display='';
		flagsetting7.style.display='none';
	}

	if(tab==0){
		flagsetting1.style.display='none';
		flagsetting2.style.display = '';
		flagsetting3.style.display='';
		flagsetting4.style.display='';
		flagsetting5.style.display='';
		flagsetting6.style.display='';
		flagsetting7.style.display='none';
	}
}

function flagsetting(n){
	if (n == 1){
		flagsetting3.style.display='';
		flagsetting4.style.display='';
		flagsetting5.style.display='';
		flagsetting7.style.display='none';
		}
	if (n == 2){
		flagsetting3.style.display='none';
		flagsetting4.style.display='none';
		flagsetting5.style.display='';
		flagsetting7.style.display='none';
	}

}
function SelectPhoto(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.picurl.value=ss[0];
  }
}
</script>
<table border=0 align=center cellpadding=2 cellspacing=1 class=tableBorder>
    <tr>
      <th colspan=2><%=EditTitle%></th>
 </tr>
<form name=myform method=post action='admin_admanage.asp'>
<%
	If isEdit Then
		Response.Write "<input type=""Hidden"" name=""action"" value=""modify"">"
		Response.Write "<input type=""Hidden"" name=""id"" value="""& Request("id") &""">"
	Else
		Response.Write "<input type=""Hidden"" name=""action"" value=""save"">"
	End If
%>
 <tr>
 <td width='20%' class=TableRow1 align=right><strong>广告位置:</strong></td>
 <td width='80%' class=TableRow1><select name='boardid' id='boardid'>
<%
	Dim oRs
	Set oRs = enchiasp.Execute("SELECT boardid,BoardName FROM ECCMS_Adboard")
	Do While Not oRs.EOF
		Response.Write "<option value="""& oRs("boardid") &""""
		If isEdit Then
			If oRs("boardid") = Rs("boardid") Then Response.Write " selected"
		End If
		Response.Write ">"
		Response.Write oRs("BoardName")
		Response.Write "</option>"
	oRs.movenext
	Loop
	oRs.Close:Set oRs = Nothing
%>
</select> </td>
</tr>
<tr>
 <td class=TableRow2 align=right><strong>广告类型:</strong></td>
 <td class=TableRow2><select name='flag' id='flag' onchange=showsetting(this.form)>
	<option value='0'<%If isEdit Then If Rs("flag") = 0 Then Response.Write " selected"%>>普通广告</option>
	<option value='1'<%If isEdit Then If Rs("flag") = 1 Then Response.Write " selected"%>>漂浮广告</option>
	<option value='2'<%If isEdit Then If Rs("flag") = 2 Then Response.Write " selected"%>>左边固定广告</option>
	<option value='3'<%If isEdit Then If Rs("flag") = 3 Then Response.Write " selected"%>>右边固定广告</option>
	<option value='4'<%If isEdit Then If Rs("flag") = 4 Then Response.Write " selected"%>>弹出广告</option>
	<option value='5'<%If isEdit Then If Rs("flag") = 5 Then Response.Write " selected"%>>代码广告</option>
</select></td>
</tr>
<tr id='flagsetting1'<%If isEdit Then If Rs("flag")<>5 And Rs("flag")<>0 Then Response.Write (" style=""display:''"""):Else:Response.Write (" style=""display:'none'"""): End If:Else Response.Write " style=""display:none"""%>>
 <td class=TableRow1 align=right><strong>广告设置:</strong></td>
 <td class=TableRow1>左边距：<input name='sidemargin' type='text' id='sidemargin' value='<%If isEdit Then Response.Write Rs("sidemargin") Else Response.Write "10" End If%>' size='6' maxlength='5'>
 上边距：<input name='topmargin' type='text' id='topmargin' value='<%If isEdit Then Response.Write Rs("topmargin") Else Response.Write "100" End If%>' size='6' maxlength='5'></td>            
</tr>
<tr id=flagsetting2<%If isEdit Then If Rs("flag")=5 Then Response.Write " style=""display:none"""%>>
 <td class=TableRow2 align=right><strong>图片或FLASH:</strong></td>
 <td class=TableRow2><input type='radio' name='isFlash' value='0' onClick="flagsetting(1)"<%If isEdit Then If Rs("isFlash") = 0 Then Response.Write " checked" End If:Else Response.Write " checked" End If%>>图片&nbsp;&nbsp;            
 <input type='radio' name='isFlash' value='1' onClick="flagsetting(2)"<%If isEdit Then If Rs("isFlash") = 1 Then Response.Write " checked"%>>FLASH(系统默认以透明方式显示)&nbsp;&nbsp;            
<%If isEdit Then%><input type=checkbox name=UpdateTime value='yes'> 更新成最新广告<%End If%></td>            
</tr>
<tr>
 <td class=TableRow1 align=right><strong>网站名称:</strong></td>
 <td class=TableRow1><input name='title' type='text' id='title' size=30 value='<%If isEdit Then Response.Write Rs("title")%>'></td>
</tr>

<tr id=flagsetting3<%If isEdit Then If Rs("isFlash")=1 Or Rs("flag")=5 Then Response.Write " style=""display:none"""%>>
 <td class=TableRow2 align=right><strong>网站连接URL:</strong></td>
 <td class=TableRow2><input name='url' type='text' id='url' size=60 value='<%If isEdit Then Response.Write Rs("url") Else Response.Write "http://" End If%>'></td>
</tr>
<tr id=flagsetting4<%If isEdit Then If Rs("isFlash")=1 Or Rs("flag")=5 Then Response.Write " style=""display:none"""%>>
 <td class=TableRow1 align=right><strong>连接注释:</strong></td>
 <td class=TableRow1><input name='Readme' type='text' id='Readme' size=60 value='<%If isEdit Then Response.Write Rs("Readme")%>'></td>
</tr>
<tr id=flagsetting5<%If isEdit Then If Rs("flag")=5  Then Response.Write " style=""display:none"""%>>
 <td width='20%' class=TableRow2 align=right><strong>图片或FLASH URL:</strong></td>            
 <td width='80%' class=TableRow2><input name='picurl' id=ImageUrl type='text' size=60 value='<%If isEdit Then Response.Write Rs("picurl")%>'>
 <input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto()' class=button></td>
</tr>
<tr id=flagsetting6<%If isEdit Then If Rs("flag")=5  Then Response.Write " style=""display:none"""%>>
 <td class=TableRow1 align=right><strong>上传文件:</strong></td>
 <td class=TableRow1><iframe name="image" frameborder=0 width=100% height=42 scrolling=no src=Upload.asp?sType=AD></iframe></td>
</tr>
<tr id=flagsetting7<%If isEdit Then If Rs("flag") <> 5 Then Response.Write " style=""display:none""" End If:Else Response.Write " style=""display:none""" End If%>>
 <td class=TableRow2 align=right><strong>广告代码:</strong><br>支持HTML方法</td>
 <td class=TableRow2><textarea name=AdCode rows=10 cols=70><%If isEdit Then Response.Write Server.HTMLEncode(Rs("AdCode"))%></textarea></td>
</tr>
<tr>
  <td class=TableRow1 align=right><strong>文件尺寸:</strong></td>
  <td class=TableRow1>宽度：<input name='width' type='text' id='width' size='6' maxlength='5' value='<%If isEdit Then Response.Write Rs("width") Else Response.Write 100%>'> 像素 * 
  高度：<input name='height' type='text' id='height' size='6' maxlength='5' value='<%If isEdit Then Response.Write Rs("height") Else Response.Write 100%>'> 像素&nbsp;&nbsp;
  <font color=blue>* 设置图片、FLASH、弹出窗口的大小</font></td>
</tr>
<tr>
 <td class=TableRow2 align=right><strong>是否隐藏广告:</strong></td>
 <td class=TableRow2><input type='radio' name='isLock' value='0' <%If isEdit Then If Rs("isLock") = 0 Then Response.Write " checked" End If:Else Response.Write " checked" End If%>> 否&nbsp;&nbsp;
 <input type='radio' name='isLock' value='1'<%If isEdit Then If Rs("isLock") = 1 Then Response.Write " checked"%>> 是&nbsp;&nbsp;
 </td>
</tr>
<tr>
 <td class=TableRow1 colspan=2 align=center>
 <input type="button" name="Submit1" onclick="javascript:history.go(-1)" value="返回上一页" class=button>&nbsp;&nbsp;&nbsp;&nbsp;
 <input type='submit' name='Submit' value='保存广告' class=button>
 </td>
</tr></form>
</table>
<%
End Sub
Sub CheckSave()
	If Trim(Request.Form("title")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>网站名称不能为空！</li>"
	End If
	If Trim(Request.Form("flag")) = 5 Then
		If Trim(Request.Form("AdCode")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>请输入广告代码！</li>"
		End If
	Else
		If Trim(Request.Form("picurl")) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>URL不能为空！</li>"
		End If
	End IF
	If Trim(Request.Form("height")) = "" Or Trim(Request.Form("width")) = ""  Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>文件尺寸不能为空！</li>"
	End If
	If Trim(Request.Form("sidemargin")) = "" Or Trim(Request.Form("topmargin")) = ""  Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请设置边距！</li>"
	End If
	If Trim(Request.Form("flag")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请设置广告类型！</li>"
	End If
End Sub

Sub SaveAdvertise()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Adlist where (id is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("boardid") = Trim(Request.Form("boardid"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("url") = Trim(Request.Form("url"))
		Rs("picurl") = Trim(Request.Form("picurl"))
		Rs("Readme") = enchiasp.ChkFormStr(Request.Form("Readme"))
		Rs("AdCode") = Request.Form("AdCode")
		Rs("height") = Trim(Request.Form("height"))
		Rs("width") = Trim(Request.Form("width"))
		Rs("topmargin") = Trim(Request.Form("topmargin"))
		Rs("sidemargin") = Trim(Request.Form("sidemargin"))
		Rs("startime") = Now()
		Rs("flag") = Trim(Request.Form("flag"))
		Rs("isFlash") = Trim(Request.Form("isFlash"))
		Rs("IsLock") = CInt(Request.Form("IsLock"))
	Rs.update
	Rs.Close:Set Rs = Nothing
	Dim adenchiasp
	Set adenchiasp = New Admanage_Cls
	adenchiasp.adboardid = CLng(Request.Form("boardid"))
	adenchiasp.CreateJsFile
	Set adenchiasp = Nothing
	Succeed("<li>恭喜您！添加新的广告成功。</li>")
End Sub

Sub ModifyAdvertise()
	CheckSave
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Adlist where id = " & Request("id")
	Rs.Open SQL,Conn,1,3
		Rs("boardid") = Trim(Request.Form("boardid"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("url") = Trim(Request.Form("url"))
		Rs("picurl") = Trim(Request.Form("picurl"))
		Rs("Readme") = enchiasp.ChkFormStr(Request.Form("Readme"))
		Rs("AdCode") = Request.Form("AdCode")
		Rs("height") = Trim(Request.Form("height"))
		Rs("width") = Trim(Request.Form("width"))
		Rs("topmargin") = Trim(Request.Form("topmargin"))
		Rs("sidemargin") = Trim(Request.Form("sidemargin"))
		If LCase(Request.Form("UpdateTime")) = "yes" Then Rs("startime") = Now()
		Rs("flag") = Trim(Request.Form("flag"))
		Rs("isFlash") = Trim(Request.Form("isFlash"))
		Rs("IsLock") = CInt(Request.Form("IsLock"))
	Rs.update
	AdvertiseID = Rs("id")
	Rs.Close:Set Rs = Nothing
	Dim adenchiasp
	Set adenchiasp = New Admanage_Cls
	adenchiasp.adboardid = CLng(Request.Form("boardid"))
	adenchiasp.CreateJsFile
	Set adenchiasp = Nothing
	Succeed("<li>恭喜您！修改广告成功。</li>")
End Sub
%>