<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/classmenu.asp"-->

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
Dim HtmlContent,ChannelRootDir
ChannelRootDir = enchiasp.InstallDir & "user/"
enchiasp.LoadTemplates 9999, 4, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
'--Ƶ��Ŀ¼
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","�û���¼")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = ReadFriendLink(HtmlContent)
Response.Write HtmlContent
CloseConn




	'================================================
	'��������LoadFriendLink
	'��  �ã�װ����������
	'��  ����str ----ԭ�ַ���
	'================================================
	Public Function LoadFriendLink(ByVal TopNum, ByVal PerRowNum, ByVal isLogo, ByVal orders)
		Dim Rs, SQL, i, strContent
		Dim strOrder, LinkAddress
	
		strContent = ""
		If Not IsNumeric(TopNum) Then Exit Function
		If Not IsNumeric(PerRowNum) Then Exit Function
		If Not IsNumeric(isLogo) Then Exit Function
		If Not IsNumeric(orders) Then Exit Function
		On Error Resume Next
		If CInt(orders) = 1 Then
			'-- ��ҳ��ʾ��ʱ����������
			strOrder = "And isIndex > 0 Order By LinkTime Desc,LinkID Desc"
		ElseIf CInt(orders) = 2 Then
			'-- ��ҳ��ʾ���������������
			strOrder = "And isIndex > 0 Order By LinkHist Desc,LinkID Desc"
		ElseIf CInt(orders) = 3 Then
			'-- ��ҳ��ʾ���������������
			strOrder = "And isIndex > 0 Order By LinkHist Desc,LinkID Asc"
		ElseIf CInt(orders) = 4 Then
			'-- ���а���������
			strOrder = "Order By LinkID Desc"
		ElseIf CInt(orders) = 5 Then
			'-- ���а���������
			strOrder = "Order By LinkID Asc"
		ElseIf CInt(orders) = 6 Then
			'-- ���а��������������
			strOrder = "Order By LinkHist Desc,LinkID Desc"
		ElseIf CInt(orders) = 7 Then
			'-- ���а��������������
			strOrder = "Order By LinkHist Desc,LinkID Asc"
		ElseIf CInt(orders) = 8 Then
			'-- ��ҳ��ʾ����������
			strOrder = "And isIndex > 0 Order By LinkName Desc,LinkID Desc"
		ElseIf CInt(orders) = 9 Then
			'-- ���а���������
			strOrder = "Order By LinkName Desc,LinkID Desc"
		Else
			'-- ��ҳ��ʾ��ʱ�併������
			strOrder = "And isIndex > 0 Order By LinkTime Asc,LinkID Asc"
		End If
		If CInt(isLogo) = 1 Or CInt(isLogo) = 3 Then
			SQL = "Select Top " & CInt(TopNum) & " LinkID,LinkName,LinkUrl,LogoUrl,Readme,LinkHist,isLogo from [ECCMS_Link] where isLock = 0 And isLogo > 0 " & strOrder & ""
		Else
			SQL = "Select Top " & CInt(TopNum) & " LinkID,LinkName,LinkUrl,LogoUrl,Readme,LinkHist,isLogo from [ECCMS_Link] where isLock = 0 And isLogo = 0 " & strOrder & ""
		End If
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.BOF And Rs.EOF) Then
			strContent = "<table width=""100%"" border=0 cellpadding=1 cellspacing=3 class=FriendLink1>" & vbCrLf
			Do While Not Rs.EOF
				strContent = strContent & "<tr>" & vbCrLf
				For i = 1 To CInt(PerRowNum)
					strContent = strContent & "<td align=center class=FriendLink2>"
					If Not Rs.EOF Then
						If CInt(isLogo) < 2 Then
							LinkAddress = enchiasp.InstallDir & "link/link.asp?id=" & Rs("LinkID") & "&url=" & Trim(Rs("LinkUrl"))
						Else
							LinkAddress = Trim(Rs("LinkUrl"))
						End If
						If Rs("isLogo") = 1 Or CInt(isLogo) = 3 Then
							strContent = strContent & "<a href='" & LinkAddress & "' target=_blank title='��ҳ���ƣ�" & Rs("LinkName") & "&#13;&#10;���������" & Rs("LinkHist") & "'><img src='" & enchiasp.ReadFileUrl(Rs("LogoUrl")) & "' border=0 width=162 height=48></a>"
						Else
							strContent = strContent & "<a href='" & LinkAddress & "' target=_blank title='��ҳ���ƣ�" & Rs("LinkName") & "&#13;&#10;���������" & Rs("LinkHist") & "'>" & Rs("LinkName") & "</a>"
						End If
						strContent = strContent & "</td>" & vbCrLf
						Rs.MoveNext
					Else
						If CInt(isLogo) = 1 Or CInt(isLogo) = 3 Then
							strContent = strContent & "<a href='" & enchiasp.InstallDir & "link/addlink.asp' target=_blank><img src='" & enchiasp.InstallDir & "images/link.gif'  border=0></a>"
						Else
							strContent = strContent & "<a href='" & enchiasp.InstallDir & "link/' target=_blank>��������</a>"
						End If
						strContent = strContent & "</td>" & vbCrLf
					End If
				Next
				strContent = strContent & "</tr>" & vbCrLf
			Loop
			strContent = strContent & "</table>" & vbCrLf
		End If
		LoadFriendLink = strContent
	End Function
	'================================================
	'��������ReadFriendLink
	'��  �ã���ȡ��������
	'��  ����str ----ԭ�ַ���
	'================================================
	Public Function ReadFriendLink(ByVal str)
		Dim strTemp, i
		Dim sTempContent, nTempContent, ArrayList
		Dim arrTempContent, arrTempContents
		On Error Resume Next
		strTemp = str

		If InStr(strTemp, "{$ReadFriendLink(") > 0 Then
			sTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadFriendLink(", ")}", 1)
			nTempContent = enchiasp.CutMatchContent(strTemp, "{$ReadFriendLink(", ")}", 0)
			arrTempContents = Split(sTempContent, "|||")
			arrTempContent = Split(nTempContent, "|||")
			For i = 0 To UBound(arrTempContents)
				ArrayList = Split(arrTempContent(i), ",")
				strTemp = Replace(strTemp, arrTempContents(i), LoadFriendLink(ArrayList(0), ArrayList(1), ArrayList(2), ArrayList(3)))
			Next
		End If
		ReadFriendLink = strTemp
	End Function


%>