<!--#include file="config.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
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
Dim TempListContent,ListContent
Dim Rs, SQL, foundsql, j, keyword,rsGuest
Dim maxperpage, totalnumber, TotalPageNum, CurrentPage, i
Dim strReplyAlt,strHomePage,strLockedAlt,strPagination,strClassName
Dim maxstrlen,IsReply,GuestContent,ReplyContent

maxperpage = 4	'--ÿҳ��ʾ������

enchiasp.LoadTemplates 9999, 1, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
'--Ƶ��Ŀ¼
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","�����б�")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
	HtmlContent = HTML.ReadFriendLink(HtmlContent)

HtmlContent = Replace(HtmlContent, "{$MemberName}", enchiasp.membername)

maxstrlen = enchiasp.ChkNumeric(enchiasp.HtmlSetting(5))	'--������������ַ�����
IsReply = enchiasp.ChkNumeric(enchiasp.HtmlSetting(6))	'--�Ƿ���ʾ�ظ�
strClassName = enchiasp.ChannelName
CurrentPage = enchiasp.ChkNumeric(Request("page"))
If CInt(CurrentPage) = 0 Then CurrentPage = 1

If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
	keyword = enchiasp.ChkQueryStr(Request("keyword"))
	If LCase(Request("act")) = "topic" Then
		foundsql = "WHERE isAccept > 0 And title like '%" & keyword & "%'"
	ElseIf LCase(Request("act")) = "username" Then
		foundsql = "WHERE isAccept > 0 And username like '%" & keyword & "%'"
	Else
		foundsql = "WHERE isAccept > 0 And title like '%" & keyword & "%'"
	End If
Else
	foundsql = "WHERE isAccept > 0"
End If
'��¼����
totalnumber = enchiasp.Execute("SELECT COUNT(guestid) FROM ECCMS_GuestBook " & foundsql & "")(0)
TotalPageNum = CLng(totalnumber / maxperpage)  '�õ���ҳ��
If TotalPageNum < totalnumber / maxperpage Then TotalPageNum = TotalPageNum + 1
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
Set Rs = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM ECCMS_GuestBook " & foundsql & " ORDER BY isTop DESC,lastime DESC,guestid DESC"
Rs.Open SQL, Conn, 1, 1
If Rs.BOF And Rs.EOF Then
	HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), enchiasp.HtmlSetting(1))
Else
	i = 0
	If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
	j = totalnumber - ((CurrentPage - 1) * maxperpage)
	TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 0)
	Do While Not Rs.EOF And i < CLng(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		ListContent = ListContent & TempListContent
		ListContent = Replace(ListContent,"{$GuestID}", Rs("guestid"))
		ListContent = Replace(ListContent,"{$Several}", j)
		ListContent = Replace(ListContent,"{$GuestTopic}", enchiasp.HTMLEncode(enchiasp.CheckTopic(Rs("title"))))
		ListContent = Replace(ListContent,"{$UserName}", Rs("username"))
		ListContent = Replace(ListContent,"{$UserStatus}", GuestStation(Rs("userid")))
		ListContent = Replace(ListContent,"{$UserFace}", enchiasp.CheckTopic(Rs("face")))
		ListContent = Replace(ListContent,"{$ComeFrom}", Rs("ComeFrom"))
		ListContent = Replace(ListContent,"{$GuestQQ}", Rs("GuestOicq"))
		ListContent = Replace(ListContent,"{$Emotion}", Rs("emot"))
		ListContent = Replace(ListContent,"{$GuestEmail}", enchiasp.CheckTopic(Rs("GuestEmail")))
		ListContent = Replace(ListContent,"{$WriteTime}", Rs("WriteTime"))
		ListContent = Replace(ListContent,"{$GuestIP}", Rs("GuestIP"))
		
		If IsNull(Rs("Topicformat")) Then
			ListContent = Replace(ListContent, "{$Topicformat}", "")
		Else
			ListContent = Replace(ListContent, "{$Topicformat}", " "& Rs("Topicformat"))
		End If
		If Rs("ForbidReply") <> 0 Then
			strLockedAlt = "      <img src='images/a_lock.gif' align=""absmiddle"" alt=""����������������ֹ�ظ�"">"
		Else
			strLockedAlt = ""
		End If
		ListContent = Replace(ListContent,"{$LockedAlt}", strLockedAlt)
		If Rs("ReplyNum") <> 0 Then
			strReplyAlt = "<img src='images/collapsed_yes.gif' border=0 alt='�лظ����� " & Rs("ReplyNum") & " ����'>"
		Else
			strReplyAlt = "<img src='images/collapsed_no.gif' border=0 alt='�޻ظ�'>"
		End If
		ListContent = Replace(ListContent,"{$ReplyAlt}", strReplyAlt)
		If enchiasp.CheckNull(Rs("HomePage")) Then
			strHomePage = Rs("HomePage")
			If LCase(Left(strHomePage,4)) <> "http" Then strHomePage = "http://" & strHomePage
		Else
			strHomePage = "#"
		End If
		ListContent = Replace(ListContent,"{$HomePage}", strHomePage)
		GuestContent = enchiasp.ChkBadWords(UBBCode(Rs("Content")))
		If maxstrlen > 0 Then GuestContent = enchiasp.CutString(GuestContent,maxstrlen)
		
		If Rs("isAdmin") <> 0 Then
			If Rs("username") = enchiasp.membername Or enchiasp.membergrade = "999" Or Trim(Session("AdminName")) <> "" Then
				ListContent = Replace(ListContent, "{$GuestContent}", GuestContent)
			Else
				ListContent = Replace(ListContent, "{$GuestContent}",enchiasp.HtmlSetting(2))
			End If
		Else
			ListContent = Replace(ListContent, "{$GuestContent}", GuestContent)
		End If
		
		If IsReply > 0 And InStr(ListContent, "{$ReplyMessage}") > 0 Then
			ListContent = Replace(ListContent,"{$ReplyMessage}", ReplyMessage)
		Else
			ListContent = Replace(ListContent,"{$ReplyMessage}", vbNullString)
		End If
		
		Rs.movenext
		i = i + 1
		j = j - 1
		If i >= maxperpage Then Exit Do
	Loop
	HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
End If
Rs.Close:Set Rs = Nothing

strPagination = ShowListPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, "", strClassName)
HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strPagination)
HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
HtmlContent = HTML.ReadAnnounceList(HtmlContent)
Response.Write HtmlContent
Set HTML = Nothing
CloseConn

Public Function ReplyMessage()
	Dim strTempContent
	On Error Resume Next
	
	Set rsGuest = enchiasp.Execute("SELECT rUserName,rTitle,rContent,ReplyTime,ReplyIP FROM ECCMS_GuestReply WHERE guestid="& Rs("guestid") & " ORDER BY id DESC")
	If Not (rsGuest.BOF And rsGuest.EOF) Then
		strTempContent = enchiasp.HtmlSetting(3)
		If IsNull(rsGuest("Topicformat")) Then
			strTempContent = Replace(strTempContent, "{$Topicformat}", "")
		Else
			strTempContent = Replace(strTempContent, "{$Topicformat}", " "& rsGuest("Topicformat"))
		End If
		strTempContent = Replace(strTempContent, "{$ReplyTopic}", enchiasp.HTMLEncode(rsGuest("rTitle")))
		ReplyContent = enchiasp.ChkBadWords(UBBCode(rsGuest("rContent")))
		If maxstrlen > 0 Then ReplyContent = enchiasp.CutString(ReplyContent,maxstrlen)
		If Rs("isAdmin") <> 0 Then
			If Rs("username") = enchiasp.membername Or enchiasp.membergrade = "999" Or Trim(Session("AdminName")) <> "" Then
				strTempContent = Replace(strTempContent, "{$ReplyContent}", ReplyContent)
			Else
				strTempContent = Replace(strTempContent, "{$ReplyContent}",enchiasp.HtmlSetting(2))
			End If
		Else
			strTempContent = Replace(strTempContent, "{$ReplyContent}", ReplyContent)
		End If
		strTempContent = Replace(strTempContent, "{$ReplyUserName}", enchiasp.HTMLEncode(rsGuest("rUserName")))
		strTempContent = Replace(strTempContent, "{$ReplyTime}", rsGuest("ReplyTime"))
	End If
	Set rsGuest = Nothing
	ReplyMessage = strTempContent
End Function

%>