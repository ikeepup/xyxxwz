<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
<%
'=====================================================================
' 软件名称：恩池网站管理系统---用户短信服务
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Call InnerLocation("用户短信服务")

Dim Rs,SQL,i,Action
Dim Maxsms,boxname,smstype,readaction

If CInt(GroupSetting(22)) = 0 Then
	ErrMsg = ErrMsg + "<li>对不起！您没有使用短信服务的权限，如有什么问题请联系管理员。</li>"
	Founderr = True
End If
Maxsms = CLng(GroupSetting(24))
Call showmain
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If

Sub showmain()
	If Founderr = True Then Exit Sub
	Dim smsCount,DelCount
	smsCount=0
	Set Rs = enchiasp.Execute("select Count(id) from ECCMS_Message Where flag=0 And incept='"& enchiasp.membername &"'")
	smsCount = CLng(Rs(0))
	'以下判断为自动删除多出来的短消息
	If smsCount > Maxsms And Maxsms <> 0 Then
		i = smsCount-Maxsms
		Set Rs=enchiasp.Execute("select top "& i &" id from ECCMS_Message Where incept='"& enchiasp.membername &"' Order by id,isRead Desc")
		While Not Rs.EOF
			enchiasp.Execute("Delete from ECCMS_Message Where id="& rs(0))
			Rs.movenext
		Wend
		smsCount = Maxsms
	End if
	Rs.Close:Set Rs = Nothing
%>
<script language="JavaScript">
<!--
function enchiasp_usersms_smsbox_top(smstype){
	document.write ('<th valign=middle width=30 height=25 noWrap>已读</th>');
	document.write ('<th valign=middle width=100>');
	if (smstype=='inbox')
	{
		document.write ('发件人');
	}else{
		document.write ('收件人');
	}
	document.write ('</th>');
	document.write ('<th valign=middle width=300>主题</th>');
	document.write ('<th valign=middle width=150>日期</th>');
	document.write ('<th valign=middle width=50>大小</th>');
	document.write ('<th valign=middle width=30 noWrap>操作</th>');
	document.write ('</tr>');
}
function enchiasp_usersms_smsbox_emp(boxname){
	document.write ('<tr>');
	document.write ('<td class=Usertablerow1 align=center valign=middle colspan=6>您的'+boxname+'中没有任何内容。</td>');
	document.write ('</tr>');
}
function enchiasp_usersms_smsbox_loop(flag,isread,sms_type,sender,incept,title,sendtime,clength,id,readaction){
	var tablebody,newstyle;
	if (isread==0)
	{
		tablebody="Usertablerow2";
		newstyle="font-weight:bold";
	}else{
		tablebody="Usertablerow1";
		newstyle="font-weight:normal";
	}
	document.write ('<tr>');
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	if (flag==0){
		if (isread==0){
			document.write ('<img src="images/m_news.gif" border=0 alt="新短信">');
			}else{
			document.write ('<img src="images/m_olds.gif" border=0 alt="旧短信">');
		}
	}else{
		document.write ('<img src="images/m_issend_2.gif" border=0 alt="系统短信">');
	}
	document.write ('</td>');
	document.write ('<td noWrap class='+tablebody+' align=center valign=middle style="'+newstyle+'">');
	if (sms_type=='inbox')
	{
		document.write ('<a href="userlist.asp?name='+sender+'" target=_blank>'+sender+'</a>');
	}else
	{
		document.write ('<a href="userlist.asp?name='+incept+'" target=_blank>'+incept+'</a>');
	}
	document.write ('</td>');
	document.write ('<td class='+tablebody+' align=left style="'+newstyle+'"><a href="message.asp?action='+readaction+'&sid='+id+'&sender='+sender+'">'+title+'</a>	</td>');
	document.write ('<td noWrap class='+tablebody+' style="'+newstyle+'">'+sendtime+'</td>');
	document.write ('<td noWrap class='+tablebody+' style="'+newstyle+'">'+clength+'Byte</td>');
	document.write ('<td align=center valign=middle width=30 class='+tablebody+'><input type=checkbox name=id value='+id+'></td>');
	document.write ('</tr>');
}
function enchiasp_usersms_smsbox_footer(boxname){
	document.write ('<tr>');
	document.write ('<td align=right valign=middle colspan=6 class=Usertablerow2>节省每一分空间，请及时删除无用信息&nbsp;<input type=checkbox name=chkall value=on onclick="CheckAll2(this.form)">选中所有显示记录&nbsp;<input type=submit name=action onclick="{if(confirm(\'确定删除选定的纪录吗?\')){return true;}return false;}" value="删除'+boxname+'" class=button>&nbsp;<input type=submit name=action onclick="{if(confirm(\'确定清除'+boxname+'所有的纪录吗?\')){this.document.inbox.submit();return true;}return false;}" value="清空'+boxname+'" class=button></td>');
	document.write ('</tr>');
	document.write ('</form></table>');
}
//-->
</script>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr>
		<th>>> 短信服务 <<</th>
	</tr>
	<tr>
		<td align=center class=Usertablerow1><a href="usersms.asp?action=inbox"><img src="images/m_inbox.gif" border="0" alt="收件箱"></a>&nbsp;
		<a href="usersms.asp?action=sendbox"><img src="images/M_issend.gif" border="0" alt="已发送邮件"></a>&nbsp;
		<a href="message.asp?action=alldel" onclick=showClick('您确定要清空所有短消息吗?')><img src="images/recycle.gif" border="0" alt="清空所有短消息"></a>&nbsp;
		<a href="friend.asp"><img src="images/M_address.gif" border="0" alt="地址簿"></a>&nbsp;
		<a href="message.asp?action=new"><img src="images/m_write.gif" border="0" alt="发送讯息"></a></td>
	</tr>
</table>
<br style="overflow: hidden; line-height: 10px">
<table cellspacing=1 align=center cellpadding=3 bgcolor=#cccccc border=0 class=Usertableborder>
	<tr height=20>
		<td colspan=6 class=Usertablerow1><table Width="100%" cellpadding=2 cellspacing=1 border=0 align=center style="display:nowrap"><TR>
<td Width="100" align=right>您的邮箱容量：</td>
<td Width="*"><img src="images/bar1.gif" width="0" height="16" id="Sms_bar" align=absmiddle></td>
<td Width="150" align=center id="Sms_txt">0%</td>
</tr></table></td>
	</tr>
	<form action="message.asp" method=post name=inbox>
<%
	SQL = "select * from ECCMS_Message "
	Action = LCase(Request("action"))
	Select Case Trim(Action)
		Case "inbox"
			SQL = SQL + " where incept = '"& enchiasp.membername &"' Or flag = 1 order by id desc"
			boxname = "收件箱"
			smstype = "inbox"
			readaction = "read"
		Case "sendbox"
			SQL = SQL + " where sender = '"& enchiasp.membername &"' And delSend = 0 order by id desc"
			boxname = "发件箱"
			smstype = "sendbox"
			readaction = "outread"
		Case Else
			SQL = SQL + " where incept = '"& enchiasp.membername &"' Or flag = 1 order by id desc"
			boxname = "收件箱"
			smstype = "inbox"
			readaction = "read"
	End Select
	Call usersmsbox
	Response.Write ShowTable("Sms_bar","Sms_txt",smsCount,Maxsms)
End Sub
'================================================
' 过程名：usersmsbox
' 作  用：用户信箱列表
'================================================
Sub usersmsbox()
	Dim newstyle
	Dim CurrentPage,page_count,totalrec,Pcount,PageListNum
	PageListNum = 20
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CLng(CurrentPage)
	End If
	Response.Write "<script>enchiasp_usersms_smsbox_top('"& smstype &"')</script>"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open SQL,conn,1,1
	If Rs.EOF And Rs.BOF Then
		Response.Write "<script>enchiasp_usersms_smsbox_emp('"& boxname &"')</script>"
	Else
		Rs.PageSize = PageListNum
		Rs.AbsolutePage = CurrentPage
		page_count = 0
		totalrec = Rs.recordcount
		Do While Not Rs.EOF And (Not page_count = Rs.PageSize)
			Response.Write VbCrLf
			Response.Write "<script>enchiasp_usersms_smsbox_loop("
			Response.Write Rs("flag")
			Response.Write ","
			Response.Write Rs("isRead")
			Response.Write ",'"
			Response.Write smstype
			Response.Write "','"
			Response.Write EncodeJS(Rs("sender"))
			Response.Write "','"
			Response.Write EncodeJS(Rs("incept"))
			Response.Write "','"
			Response.Write EncodeJS(Rs("title"))
			Response.Write "','"
			Response.Write Rs("sendtime")
			Response.Write "',"
			Response.Write Len(Rs("content"))
			Response.Write ","
			Response.Write Rs("id")
			Response.Write ",'"
			Response.Write readaction
			Response.Write "')</script>"
			Response.Write VbCrLf
			page_count = page_count + 1
		Rs.movenext
		Loop
	End If
	Rs.close:Set Rs = nothing
	If totalrec Mod PageListNum = 0 Then
		Pcount =  totalrec \ PageListNum
	Else
		Pcount =  totalrec \ PageListNum+1
	End If
	If page_count = 0 Then CurrentPage = 0
	Response.Write "	<tr height=20>" & vbNewLine
	Response.Write "		<td colspan=6 class=Usertablerow1>"
	Response.Write ShowPages(CurrentPage,Pcount,totalrec,PageListNum,"action="& Request("action"))
	Response.Write "</td>"
	Response.Write "	</tr>" & vbNewLine
	Response.Write VbCrLf
	Response.Write "<script>enchiasp_usersms_smsbox_footer('"& boxname &"')</script>"
End Sub
'================================================
' 函数名：ShowTable
' 作  用：表示信箱容量
' 参  数：（图片对象名称，标题对象名称，更新数，总数）
'================================================
Function ShowTable(SrcName,TxtName,str,c)
	Dim Tempstr,Src_js,Txt_js,TempPercent
	Tempstr = str/C
	TempPercent = FormatPercent(tempstr,0,-1)
	Src_js = "document.getElementById(""" + SrcName + """)"
	Txt_js = "document.getElementById(""" + TxtName + """)"
	ShowTable = VbCrLf + "<script>"
	ShowTable = ShowTable + Src_js + ".width=""" & FormatNumber(tempstr*300,0,-1) & """;"
	ShowTable = ShowTable + Src_js + ".title=""容量上限为："&c&"条，总共已储存（"&str&"）条短信！"";"
	ShowTable = ShowTable + Txt_js + ".innerHTML="""
	If FormatNumber(tempstr*100,0,-1) < 80 Then
		ShowTable = ShowTable + "已使用:" & TempPercent & """;"
	Else
		ShowTable = ShowTable + "<font color=\""red\"">已使用:" & TempPercent & ",请赶快清理！</font>"";"
	End If
	ShowTable = ShowTable + "</script>"
End Function
Function EncodeJS(str)
	str = enchiasp.HtmlEncode(str)
	str = Replace(Replace(Replace(Replace(str,"\","\\"),"'","\'"),VbCrLf,"\n"),Chr(13),"")
	EnCodeJs = str
End Function
%><!--#include file="foot.inc"-->





