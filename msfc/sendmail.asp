<!--#include file="config.asp" -->
<!--#include file="../inc/Email.asp" -->
<!--#include file="../inc/chkinput.asp" -->
<%
'=====================================================================
' ������ƣ�������վ����ϵͳ
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж�������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
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
	NowStats = "�ʼ��������"
	HtmlTitle = "�ʼ��������"
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
	Response.Write " <th>�ʼ��������</th>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td valign=""top"" class=Border2><BR>"
	FoundErr = False
	ID = CLng(Request("ID"))
	If Not IsNumeric(ID) And ID<>"" then
  		Response.write"�����ϵͳ����!ID����������"
		Exit Sub
		Response.End
	End If
	If ID = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<br>" + "<li>��ָ���������</li>"
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
				ErrMsg = ErrMsg + "<br>" + "<li>����Email�д���!</li>"
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
	topic = "����" & enchiasp.Setting(0) & "��������������"
	mailbody = mailbody & "<style>A:visited { TEXT-DECORATION: none }"
	mailbody = mailbody & "A:active  { TEXT-DECORATION: none }"
	mailbody = mailbody & "A:hover   { TEXT-DECORATION: underline overline }"
	mailbody = mailbody & "A:link    { text-decoration: none;}"
	mailbody = mailbody & "A:visited { text-decoration: none;}"
	mailbody = mailbody & "A:active  { TEXT-DECORATION: none;}"
	mailbody = mailbody & "A:hover   { TEXT-DECORATION: underline overline}"
	mailbody = mailbody & "BODY   { FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
	mailbody = mailbody & "TD    { FONT-FAMILY: ����; FONT-SIZE: 9pt }</style>"
	Rs.Open "Select title,content,infotime,writer from ECCMS_Article where id=" & ID & "", conn, 1, 1
	If Rs.bof And Rs.Eof Then
		FoundErr = True
		ErrMsg = ErrMsg + "<br>" + "<li>û�����ҵ��������</li>"
	Else
		announce = announce & "<TABLE border=0 width='95%' align=center><TBODY><TR>"
		announce = announce & "<TD valign=middle align=top>"
		announce = announce & "--&nbsp;&nbsp;���ߣ�" & Rs("writer") & "<br>"
		announce = announce & "--&nbsp;&nbsp;����ʱ�䣺" & Rs("infotime") & "<br><br>"
		announce = announce & "--&nbsp;&nbsp;" & Rs("title") & "<br>"
		announce = announce & "" & Rs("content") & ""
		announce = announce & "<hr></TD></TR></TBODY></TABLE>"
	End If
	Rs.Close
	mailbody = mailbody + announce
	mailbody = mailbody & "<center><a href=http://www.enchi.com.cn>�������</a>"
	Select Case CInt(enchiasp.Setting(10))
		Case 0
			SucMsg = SucMsg + "�Բ���!ϵͳδ�����ʼ����ܡ�"
		Case 1
			Call Jmail(useremail, topic, mailbody)
		Case 2
			Call Cdonts(useremail, topic, mailbody)
		Case 3
			Call aspemail(useremail, topic, mailbody)
		Case Else
			SucMsg = SucMsg + "ϵͳδ�����ʼ����ܣ����ס����ע����Ϣ��"
	End Select
	If SendMail = "OK" Then
		SucMsg = SucMsg + "��ϲ�������Ĵ���ʵݷ��ͳɹ���"
	Else
		SucMsg = SucMsg + "����ϵͳ�������Ĵ���ʵݷ���δ�ɹ���"
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
	Response.Write " <b>����ʵ�</b></th></tr>"
	Response.Write " <tr>"
	Response.Write " <td class=Border2 valign=middle colspan=2>"
	Response.Write " <b>�ѱ��Ĵ���ʵݡ�</b><br>����ȷ������Ҫ�ʵݵ��ʼ���ַ��"
	Response.Write " </td></tr><tr>"
	Response.Write " <td class=Border2><b>�ʵݵ� Email ��ַ��</b></td>"
	Response.Write " <td class=Border2><input type=text size=40 name=""mail""></td>"
	Response.Write " </tr><tr>"
	Response.Write " <td colspan=2 class=Border2 align=center><input type=submit value=""�� ��"" name=""Submit""></table></td></form></tr></table>"
End Sub
Private Sub success()
	Response.Write " <table cellpadding=0 cellspacing=1 border=0 bgColor=#CECECE width=460 align=center>"
	Response.Write " <tr>"
	Response.Write " <td class=""Border2"">"
	Response.Write " <table cellpadding=3 cellspacing=1 border=0 width=""100%"">"
	Response.Write " <tr align=""center""> "
	Response.Write " <th width=""100%"">�ɹ�������ʵ�</th>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td width=""100%"" class=""Border2"">"
	Response.Write SucMsg
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr align=""center""> "
	Response.Write " <td width=""100%"" class=""Border1"">"
	Response.Write "<a href=""javascript:history.go(-1)""> << ������һҳ</a>"
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
	Response.Write " <th>������Ϣ</th>"
	Response.Write " </tr>"
	Response.Write " <tr> "
	Response.Write " <td width=""100%"" class=""Border2""><b>��������Ŀ���ԭ��</b><br><br>"
	Response.Write ErrMsg
	Response.Write " </td>"
	Response.Write " </tr>"
	Response.Write " <tr align=""center""> "
	Response.Write " <td width=""100%"" class=""Border1"">"
	Response.Write "<a href=""javascript:history.go(-1)""> << ������һҳ</a>"
	Response.Write " </td>"
	Response.Write " </tr> "
	Response.Write " </table> </td></tr></table>"
End Sub
%>
