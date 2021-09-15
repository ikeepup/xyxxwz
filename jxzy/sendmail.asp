<!--#include file="config.asp" -->
<!--#include file="../inc/Email.asp" -->
<!--#include file="../inc/chkinput.asp" -->
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
Dim enchicms_Ads
Dim Rs
Dim SQL
Dim NowStats
Dim HtmlTitle
Dim Style_CSS
Dim HtmlTempStr
Dim TempTopStr
Dim TempFootStr
Dim boardtype
Dim FoundErr
Dim ErrMsg
Dim ID
Dim SucMsg
Dim topic
Dim mailbody
Dim announce
Dim useremail
ArticlEndMail
CloseConn
Public Sub ArticlEndMail()
	On Error Resume Next
	enchiasp.LoadTemplates ("")
	Set enchicms_Ads = Server.CreateObject("enchicmsAsp.Admin_Adcolumn")
	Set Rs = Server.CreateObject("adodb.recordset")
	NowStats = "邮件打包发送"
	HtmlTitle = "邮件打包发送"
	TempTopStr = enchiasp.mainhtml(0) & enchiasp.mainhtml(1) & enchiasp.mainhtml(2) & enchiasp.mainhtml(3)
	TempFootStr = enchiasp.mainhtml(4)
	Style_CSS = Replace(Replace(enchiasp.Style_CSS, "{$SetupDir}", enchiasp.SetupDir), "{$PicUrl}", enchiasp.TempDir)
	HtmlTempStr = TempTopStr
	HtmlTempStr = Replace(HtmlTempStr, "{$NavMenu}", enchiasp.SortingMenu)
	HtmlTempStr = Replace(HtmlTempStr, "{$Width}", enchiasp.mainset(0))
	HtmlTempStr = Replace(HtmlTempStr, "{$Style_CSS}", Style_CSS)
	If CInt(enchiasp.Setting(5)) = 0 Then
		HtmlTempStr = Replace(HtmlTempStr, "{$TopMeun}", enchiasp.mainset(9))
	Else
		HtmlTempStr = Replace(HtmlTempStr, "{$TopMeun}", enchiasp.mainset(10))
	End If
	HtmlTempStr = Replace(HtmlTempStr, "{$NowStats}", NowStats)
	HtmlTempStr = Replace(HtmlTempStr, "{$Title}", HtmlTitle)
	HtmlTempStr = Replace(HtmlTempStr, "{$Adcolumn(0)}", enchicms_Ads.RunScriptAds(7))
	HtmlTempStr = Replace(HtmlTempStr, "{$Adcolumn(1)}", enchicms_Ads.BannerAds(7))
	HtmlTempStr = Replace(HtmlTempStr, "{$Adcolumn(2)}", enchicms_Ads.AdsColumn(7, 2))
	HtmlTempStr = Replace(HtmlTempStr, "{$Adcolumn(3)}", enchicms_Ads.AdsColumn(7, 3))
	Response.Write HtmlTempStr
	TempFootStr = Replace(TempFootStr, "{$FootMeun}", enchiasp.mainset(11))
	TempFootStr = Replace(TempFootStr, "{$Width}", enchiasp.mainset(0))
	TempFootStr = Replace(TempFootStr, "{$Adcolumn(4)}", enchicms_Ads.ScriptFloatAds(7))
	TempFootStr = Replace(TempFootStr, "{$Adcolumn(5)}", enchicms_Ads.ScriptFixedAds(7))
	Response.Write "<table width="""
	Response.Write enchiasp.mainset(0)
	Response.Write """ class=TableBorder border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""Border"">"
	Response.Write " <tr> "
	Response.Write " <th>邮件打包发送</th>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td valign=""top"" class=Border2><BR>"
	FoundErr = False
	ID = CLng(Request("ID"))
	If Not IsNumeric(ID) And ID<>"" then
  		Response.write"错误的系统参数!ID必须是数字"
		Exit Sub
		Response.End
	End If
	If ID = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<br>" + "<li>请指定相关文章</li>"
		Exit Sub
	End If
	Set Rs = Server.CreateObject("adodb.recordset")
	If FoundErr Then
		Call errormsg
	Else
		Call showPage
	End If
	Response.Write "<BR>"
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write "</table>"
	Response.Write TempFootStr
	Set enchicms_Ads = Nothing
End Sub
Private Sub showPage()
	On Error Resume Next
	If FoundErr Then
		Call errormsg
	Else
		If Request("action") = "sendmail" Then
			If IsValidEmail(Trim(Request.Form("mail"))) = False Then
				ErrMsg = ErrMsg + "<br>" + "<li>您的Email有错误!</li>"
				FoundErr = True
			Else
				useremail = Trim(Request.Form("mail"))
			End If
			If SendMail = "OK" Then
				Call success
			End If
			Call announceinfo
			If FoundErr Then
				Call errormsg
			Else
				Call success
			End If
		Else
			Call pag
		End If
	End If
	If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub announceinfo()
	topic = "您从" & enchiasp.Setting(0) & "发来的文章资料"
	mailbody = mailbody & "<style>A:visited { TEXT-DECORATION: none }"
	mailbody = mailbody & "A:active  { TEXT-DECORATION: none }"
	mailbody = mailbody & "A:hover   { TEXT-DECORATION: underline overline }"
	mailbody = mailbody & "A:link    { text-decoration: none;}"
	mailbody = mailbody & "A:visited { text-decoration: none;}"
	mailbody = mailbody & "A:active  { TEXT-DECORATION: none;}"
	mailbody = mailbody & "A:hover   { TEXT-DECORATION: underline overline}"
	mailbody = mailbody & "BODY   { FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	mailbody = mailbody & "TD    { FONT-FAMILY: 宋体; FONT-SIZE: 9pt }</style>"
	Rs.Open "Select title,content,infotime,writer from ECCMS_Article where id=" & ID & "", conn, 1, 1
	If Rs.bof And Rs.Eof Then
		FoundErr = True
		ErrMsg = ErrMsg + "<br>" + "<li>没有有找到相关文章</li>"
	Else
		announce = announce & "<TABLE border=0 width='95%' align=center><TBODY><TR>"
		announce = announce & "<TD valign=middle align=top>"
		announce = announce & "--&nbsp;&nbsp;作者：" & Rs("writer") & "<br>"
		announce = announce & "--&nbsp;&nbsp;发布时间：" & Rs("infotime") & "<br><br>"
		announce = announce & "--&nbsp;&nbsp;" & Rs("title") & "<br>"
		announce = announce & "" & Rs("content") & ""
		announce = announce & "<hr></TD></TR></TBODY></TABLE>"
	End If
	Rs.Close
	mailbody = mailbody + announce
	mailbody = mailbody & "<center><a href=http://www.enchi.com.cn>恩池软件</a>"
	Select Case CInt(enchiasp.Setting(10))
		Case 0
			SucMsg = SucMsg + "对不起!系统未开启邮件功能。"
		Case 1
			Call Jmail(useremail, topic, mailbody)
		Case 2
			Call Cdonts(useremail, topic, mailbody)
		Case 3
			Call aspemail(useremail, topic, mailbody)
		Case Else
			SucMsg = SucMsg + "系统未开启邮件功能，请记住您的注册信息。"
	End Select
	If SendMail = "OK" Then
		SucMsg = SucMsg + "恭喜您，您的打包邮递发送成功。"
	Else
		SucMsg = SucMsg + "由于系统错误，您的打包邮递发送未成功。"
	End If
End Sub
Private Sub pag()
	Response.Write "<table cellpadding=0 cellspacing=0 border=0 width=460 class=""Border"" align=center>"
	Response.Write " <tr>"
	Response.Write " <td class=Border2>"
	Response.Write " <table cellpadding=6 cellspacing=1 bgColor=#CECECE border=0 width=""100%"">"
	Response.Write " <form action=""sendmail.asp?action=sendmail&id="
	Response.Write ID
	Response.Write """ method=post>"
	Response.Write " <tr>"
	Response.Write " <th valign=middle colspan=2 align=center>"
	Response.Write " <b>打包邮递</b></th></tr>"
	Response.Write " <tr>"
	Response.Write " <td class=Border2 valign=middle colspan=2>"
	Response.Write " <b>把本文打包邮递。</b><br>请正确输入你要邮递的邮件地址！"
	Response.Write " </td></tr><tr>"
	Response.Write " <td class=Border2><b>邮递的 Email 地址：</b></td>"
	Response.Write " <td class=Border2><input type=text size=40 name=""mail""></td>"
	Response.Write " </tr><tr>"
	Response.Write " <td colspan=2 class=Border2 align=center><input type=submit value=""发 送"" name=""Submit""></table></td></form></tr></table>"
End Sub
Private Sub success()
	Response.Write " <table cellpadding=0 cellspacing=1 border=0 bgColor=#CECECE width=460 align=center>"
	Response.Write " <tr>"
	Response.Write " <td class=""Border2"">"
	Response.Write " <table cellpadding=3 cellspacing=1 border=0 width=""100%"">"
	Response.Write " <tr align=""center""> "
	Response.Write " <th width=""100%"">成功：打包邮递</th>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td width=""100%"" class=""Border2"">"
	Response.Write SucMsg
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr align=""center""> "
	Response.Write " <td width=""100%"" class=""Border1"">"
	Response.Write "<a href=""javascript:history.go(-1)""> << 返回上一页</a>"
	Response.Write " </td>"
	Response.Write " </tr> "
	Response.Write " </table> </td></tr></table>"
End Sub
Private Sub errormsg()
	Response.Write "<br>"
	Response.Write " <table cellpadding=0 cellspacing=1 border=0 width=65% bgColor=#CECECE align=center>"
	Response.Write " <tr>"
	Response.Write " <td class=""Border2"">"
	Response.Write " <table cellpadding=3 cellspacing=1 border=0 width=""100%"">"
	Response.Write " <tr align=""center""> "
	Response.Write " <th>错误信息</th>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td width=""100%"" class=""Border2""><b>产生错误的可能原因：</b><br><br>"
	Response.Write ErrMsg
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr align=""center""> "
	Response.Write " <td width=""100%"" class=""Border1"">"
	Response.Write "<a href=""javascript:history.go(-1)""> << 返回上一页</a>"
	Response.Write " </td>"
	Response.Write " </tr> "
	Response.Write " </table> </td></tr></table>"
End Sub
%>
