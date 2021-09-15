<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
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
Dim FoundErr,ErrMsg,Position,Postmsg,sucmsg
Dim ChannelID,rsChannel
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
ChannelID = CLng(ChannelID)
If ChannelID > 0 Then
	Set rsChannel = enchiasp.Execute("Select ChannelID From ECCMS_Channel where ChannelType < 2 And ChannelID = " & ChannelID)
	If Not (rsChannel.bof And rsChannel.EOF) Then
		enchiasp.ReadChannel(ChannelID)
	End If
	rsChannel.Close:Set rsChannel = Nothing
Else
	ChannelID = 0
End If
FoundErr = False
Postmsg = "<li>非法操作，请不要从外部提交数据！</li>"

Function GetVerifyCode()
	Dim Test
	On Error Resume Next
	Set Test = Server.CreateObject("Adodb.Stream")
	Set Test = Nothing
	If Err Then
		Dim zNum
		Randomize Timer
		zNum = CInt(8999 * Rnd + 1000)
		Session("GetCode") = zNum
		GetVerifyCode = Session("GetCode")
	Else
		GetVerifyCode = "<img src=""../inc/getcode.asp"">"
	End If
End Function
'================================================
'函数名：ShowPages
'作  用：通用分页
'================================================
Function ShowPages(CurrentPage,Pcount,totalrec,PageNum,str)
	Dim strTemp,strRequest
	strRequest = str
	strTemp = "<table border=0 cellpadding=0 cellspacing=3 width=""100%"" align=center>" & vbNewLine
	strTemp = strTemp & "<tr><td valign=middle nowrap>" & vbNewLine
	strTemp = strTemp & "页次：<b><font color=red>" & CurrentPage & "</font></b>/<b>" & Pcount & "</b>页&nbsp;" & vbNewLine
	strTemp = strTemp & "每页<b>" & PageNum & "</b> 总数<b>" & totalrec & "</b></td>" & vbNewLine
	strTemp = strTemp & "<td valign=middle nowrap align=right>分页：" & vbNewLine
	strTemp = strTemp & "<script language=""JavaScript"">" & vbNewLine
	strTemp = strTemp & "<!--" & vbNewLine
	strTemp = strTemp & "var CurrentPage=" & CurrentPage & ";" & vbNewLine
	strTemp = strTemp & "var Pcount=" & Pcount & ";" & vbNewLine
	strTemp = strTemp & "var Endpage=0;" & vbNewLine
	strTemp = strTemp & "if (CurrentPage > 4){" & vbNewLine
	strTemp = strTemp & "	document.write ('<a href=""?page=1" & strRequest & """>[1]</a> ...');" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "if (Pcount>CurrentPage+3)" & vbNewLine
	strTemp = strTemp & "{" & vbNewLine
	strTemp = strTemp & "	Endpage=CurrentPage+3" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "else{" & vbNewLine
	strTemp = strTemp & "	Endpage=Pcount" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "for (var i=CurrentPage-3;i<=Endpage;i++)" & vbNewLine
	strTemp = strTemp & "{" & vbNewLine
	strTemp = strTemp & "	if (i>=1){" & vbNewLine
	strTemp = strTemp & "		if (i == CurrentPage)" & vbNewLine
	strTemp = strTemp & "		{" & vbNewLine
	strTemp = strTemp & "			document.write ('<font color=""#FF0000"">['+i+']</font>');" & vbNewLine
	strTemp = strTemp & "			}" & vbNewLine
	strTemp = strTemp & "		else{" & vbNewLine
	strTemp = strTemp & "			document.write ('<a href=""?page='+i+'" & strRequest & """>['+i+']</a>');" & vbNewLine
	strTemp = strTemp & "		}" & vbNewLine
	strTemp = strTemp & "	}" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "if (CurrentPage+3 < Pcount){" & vbNewLine 
	strTemp = strTemp & "	document.write ('...<a href=""?page='+Pcount+'" & strRequest & """>['+Pcount+']</a>');" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "if (Endpage == 0){ " & vbNewLine
	strTemp = strTemp & "	document.write ('...');" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "//-->" & vbNewLine
	strTemp = strTemp & "</script>" & vbNewLine
	strTemp = strTemp & "</td></tr></table>"
	ShowPages = strTemp
End Function
'================================================
'过程名：Returnerr
'作  用：返回错误信息
'================================================
Sub Returnerr(message)
	Response.Write "<html><head><title>错误提示信息!</title><meta http-equiv=Content-Type content=text/html; charset=gb2312>" & vbCrLf
	Response.Write "<link href=user_style.css rel=stylesheet type=text/css></head><body><br /><br />" & vbCrLf
	Response.Write "<table width=460 border=0 align=center cellpadding=0 cellspacing=0>"
	Response.Write "<tr bgcolor='#3795d2'>"
	Response.Write "  <td height='25' valign='top' bgcolor='#3795d2'> <img src='images/user_msg.gif' width=69 height=20></td>"
	Response.Write "  <td align='right' valign='top'> <img src='images/user_login_02.gif' width=4 height=4></td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "  <td width=526 height=1 colspan=2 bgcolor=#f8f6f5></td>"
	Response.Write "</tr>"
	Response.Write "<tr bgcolor=#f8f6f5>"
	Response.Write "  <td width=355 valign='top' style='padding-left: 10px;padding-top: 5px;'><font color=#3795D2><b>产生错误的可能原因：</b></font><br>" & message & "</td>"
	Response.Write "  <td> <img src='images/user_err.gif' width=95 height=97></td>"
	Response.Write "</tr>"
	Response.Write "<tr bgcolor=#f8f6f5><td align=center colspan=2><a href=javascript:history.go(-1)>返回上一页...</a></td></tr>"
	Response.Write "<tr bgcolor='#3795d2'>"
	Response.Write "  <td height='8' valign='bottom'> <img src='images/user_login_04.gif' width=4 height=4></td>"
	Response.Write "  <td align='right' valign='bottom'> <img src='images/user_login_05.gif' width=4 height=4></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<br /><br /></body></html>"
End Sub
'================================================
'过程名：Returnsuc
'作  用：返回成功信息
'================================================
Sub Returnsuc(message)
	Response.Write "<html><head><title>成功提示信息!</title><meta http-equiv=Content-Type content=text/html; charset=gb2312>" & vbCrLf
	Response.Write "<link href=user_style.css rel=stylesheet type=text/css></head><body><br /><br />" & vbCrLf
	Response.Write "<table width=460 border=0 align=center cellpadding=0 cellspacing=0>"
	Response.Write "<tr bgcolor='#3795d2'>"
	Response.Write "  <td height='25' valign='top' bgcolor='#3795d2'> <img src='images/user_msg.gif' width=69 height=20></td>"
	Response.Write "  <td align='right' valign='top'> <img src='images/user_login_02.gif' width=4 height=4></td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "  <td width=526 height=1 colspan=2 bgcolor=#f8f6f5></td>"
	Response.Write "</tr>"
	Response.Write "<tr bgcolor=#f8f6f5>"
	Response.Write "  <td width=355 style='padding-left: 10px;padding-top: 5px;'><br>" & message & "</td>"
	Response.Write "  <td> <img src='images/user_suc.gif' width=95 height=97></td>"
	Response.Write "</tr>"
	Response.Write "<tr bgcolor=#f8f6f5><td align=center colspan=2><a href=" & Request.ServerVariables("HTTP_REFERER") & ">返回上一页...</a></td></tr>"
	Response.Write "<tr bgcolor='#3795d2'>"
	Response.Write "  <td height='8' valign='bottom'> <img src='images/user_login_04.gif' width=4 height=4></td>"
	Response.Write "  <td align='right' valign='bottom'> <img src='images/user_login_05.gif' width=4 height=4></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<br /><br /></body></html>"
End Sub
Function FormatDated(DateAndTime, para)
	FormatDated = ""
	Dim strDate
	If Not IsDate(DateAndTime) Then Exit Function
	If DateAndTime >= Date Then
		strDate = "<font color=red>"
		strDate = strDate & enchiasp.FormatDate(DateAndTime, para)
		strDate = strDate & "</font>"
	Else
		strDate = "<font color=#808080>"
		strDate = strDate & enchiasp.FormatDate(DateAndTime, para)
		strDate = strDate & "</font>"
	End If
	FormatDated = strDate
End Function

Sub InnerLocation(msg)
	Response.Write "<script language=""JavaScript"">locationid.innerHTML = """ & msg & """;</script>"
End Sub

Function AddUserPointNum(username,stype)
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

Function CheckLogin()
	CheckLogin = False
	Dim Rs,C_UserName,C_UserID
	C_UserName = enchiasp.CheckBadstr(enchiasp.memberName)
	C_UserID = enchiasp.ChkNumeric(enchiasp.memberid)
	Set Rs = enchiasp.Execute("SELECT userid FROM [ECCMS_User] WHERE username='" & C_UserName & "' And userid=" & CLng(C_UserID))
	If Rs.BOF And Rs.EOF Then
		Response.Cookies(enchiasp.Cookies_Name) = ""
		CheckLogin = False
	Else
		CheckLogin = True
	End If
	Set Rs = Nothing
End Function

%>
