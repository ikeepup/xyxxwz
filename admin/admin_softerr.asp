<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
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
Dim Action
If Not ChkAdmin("ErrorSoft" & ChannelID) Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "del"
	Call DeleteErrSoft
Case "amend"
	Call AmendErrSoft
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
	Dim CurrentPage,page_count,totalnumber,Pcount,maxperpage,tablebody
	maxperpage = 30
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CLng(CurrentPage)
	End If
	If CLng(CurrentPage) = 0 Then CurrentPage = 1
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th width='5%'>ѡ��</th>"
	Response.Write "		<th width='60%'>�������</th>"
	Response.Write "		<th width='20%'>����ʱ��</th>"
	Response.Write "		<th width='15%'>�������</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=selform method=post action='admin_softerr.asp'>"
	Response.Write "	<input type=hidden name=action value=""amend"">"
	Response.Write "	<input type=hidden name=ChannelID value="""& ChannelID &""">"
	totalnumber = enchiasp.Execute("SELECT COUNT(softid) FROM ECCMS_SoftList WHERE ChannelID=" & ChannelID & " And ErrCode>0 And isAccept>0")(0)
	Pcount = CLng(totalnumber / maxperpage)  '�õ���ҳ��
	If Pcount < totalnumber / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT softid,ChannelID,Classid,SoftName,SoftVer,SoftTime FROM ECCMS_SoftList WHERE ChannelID=" & ChannelID & " And ErrCode>0 And isAccept>0 ORDER BY SoftTime DESC"
	If IsSqlDataBase=1 Then
		Set Rs = enchiasp.Execute(SQL)
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=4 class=TableRow1>û�д��������</td></tr>"
	Else
		Rs.MoveFirst
		If Pcount > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		page_count = 0
		Do While Not Rs.EOF And page_count < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			If (page_count mod 2) = 0 Then
				tablebody = "class=TableRow1"
			Else
				tablebody = "class=TableRow2"
			End If
			Response.Write "	<tr>"
			Response.Write "		<td " & tablebody & " align=center><input type=checkbox name=SoftID value="""& Rs("SoftID") &"""></td>"
			Response.Write "		<td " & tablebody & "><a href=""admin_soft.asp?action=view&ChannelID="&ChannelID&"&softid="& Rs("softid") &""" title='����鿴�����'>" & Rs("SoftName") & " " & Rs("SoftVer") & "</a></td>"
			Response.Write "		<td " & tablebody & ">" & Rs("SoftTime") & "</td>"
			Response.Write "		<td " & tablebody & " align=center><a href=""admin_soft.asp?action=edit&ChannelID="&ChannelID&"&softid="& Rs("softid") &""">�༭</a> | <a href=""?action=del&ChannelID="&ChannelID&"&softid="& Rs("softid") &""" onclick=""return confirm('��ȷ��Ҫɾ���������?')"">ɾ��</a></td>"
			Response.Write "	</tr>"
			Rs.movenext
			page_count = page_count + 1
			If page_count >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "	<tr>"
	Response.Write "		<td class=tablerow1 colspan=4>"
	Response.Write "<input class=Button type=""button"" name=""chkall"" value=""ȫѡ"" onClick=""CheckAll(this.form)""><input class=Button type=""button"" name=""chksel"" value=""��ѡ"" onClick=""ContraSel(this.form)"">"
	Response.Write "<input type=submit name=submit2 value=""ֱ������"" onclick=""return confirm('ȷ��ֱ�����������?')"" class=Button>"
	Response.Write "<input type=submit name=submit3 value="" ֱ��ɾ�� "" onclick=""document.selform.action.value='del';return confirm('ȷ��Ҫɾ����?')"" class=Button>"
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=tablerow2 colspan=4>"
	Response.Write showpages(CurrentPage,Pcount,totalnumber,maxperpage,"&ChannelID="& ChannelID)
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
End Sub
Sub DeleteErrSoft()
	If Trim(Request("SoftID")) <> "" Then
		Set Rs = enchiasp.Execute("SELECT softid,classid,username FROM ECCMS_SoftList WHERE ChannelID = "& ChannelID &" And SoftID in (" & Request("SoftID") & ")")
		Do While Not Rs.EOF
			ClassUpdateCount (Rs("classid"))
			AddUserPointNum (Rs("username"))
			Rs.movenext
		Loop
		Rs.Close:Set Rs = Nothing
		enchiasp.Execute ("DELETE FROM ECCMS_SoftList WHERE ChannelID = "& ChannelID &" And SoftID in (" & Request("SoftID") & ")")
		enchiasp.Execute("DELETE FROM ECCMS_DownAddress WHERE ChannelID = "& ChannelID &" And SoftID in (" & Request("SoftID") & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_Comment WHERE ChannelID = "& ChannelID &" And PostID in (" & Request("SoftID") & ")")
		Response.Redirect  Request.ServerVariables("HTTP_REFERER")
	Else
		ErrMsg = "<li>�����ϵͳ����,��ѡ��Ҫɾ�������ID</li>"
		FoundErr = True
		Exit Sub
	End If
End Sub
Sub AmendErrSoft()
	If Trim(Request("SoftID")) <> "" Then
		enchiasp.Execute ("UPDATE ECCMS_SoftList SET ErrCode=0 WHERE ChannelID = "& ChannelID &" And SoftID in (" & Request("SoftID") & ")")
		Response.Redirect  Request.ServerVariables("HTTP_REFERER")
	Else
		ErrMsg = "<li>�����ϵͳ����,��ѡ��Ҫ���������ID</li>"
		FoundErr = True
		Exit Sub
	End If
End Sub
Private Function AddUserPointNum(username)
	On Error Resume Next
	Dim rsuser,GroupSetting,userpoint
	Set rsuser = enchiasp.Execute("SELECT userid,UserGrade,userpoint FROM ECCMS_User WHERE username='"& username &"'")
	If Not(rsuser.BOF And rsuser.EOF) Then
		GroupSetting = Split(enchiasp.UserGroupSetting(rsuser("UserGrade")), "|||")(13)
		userpoint = CLng(rsuser("userpoint") - GroupSetting)
		enchiasp.Execute ("UPDATE ECCMS_User SET userpoint="& userpoint &",experience=experience-2,charm=charm-1 WHERE userid="& rsuser("userid"))
	End If
	Set rsuser = Nothing
End Function
Private Function ClassUpdateCount(sortid)
	Dim rscount,Parentstr
	On Error Resume Next
	Set rscount = enchiasp.Execute("SELECT ClassID,Parentstr FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID=" & CLng(sortid))
	If Not (rscount.BOF And rscount.EOF) Then
		Parentstr = rscount("Parentstr") &","& rscount("ClassID")
		enchiasp.Execute ("UPDATE [ECCMS_Classify] SET ShowCount=ShowCount-1,isUpdate=1 WHERE ChannelID = "& ChannelID &" And ClassID in (" & Parentstr & ")")
	End If
	Set rscount = Nothing
End Function
%>