<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/classmenu.asp"-->
<!--#include file="../inc/cls_public.asp"-->
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



Dim maxperpage, totalnumber, TotalPageNum, CurrentPage, i
Dim Rs, SQL, sqlLink,HtmlContent
Dim FlushAddress,LinkAddress
Dim TempListContent,ListContent
Dim strLinkName,LinkName,strLinkLogo,strLinkPage

enchiasp.PreventInfuse

enchiasp.LoadTemplates 9999, 6, 0

HtmlContent = enchiasp.HtmlContent
if cint(enchiasp.HtmlSetting(15))=1 then
	HTML.Showfrink
	Set HTML = Nothing
	CloseConn
else
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","��������")

maxperpage = enchiasp.ChkNumeric(enchiasp.HtmlSetting(1))  'ÿҳ��ʾ������
FlushAddress = enchiasp.ChkNumeric(enchiasp.HtmlSetting(2))  '�Ƿ�ֱ����ʾ���ӵ�ַ


CurrentPage = enchiasp.ChkNumeric(Request("page"))
If CInt(CurrentPage) = 0 Then CurrentPage = 1

If Request("type") <> "" Then
	sqlLink = "where isLock <> 1 And isLogo=" & enchiasp.ChkNumeric(Request("type"))
Else
	sqlLink = "where isLock <> 1"
End If
'��¼����
totalnumber = enchiasp.Execute("SELECT Count(LinkID) FROM ECCMS_Link " & sqlLink & "")(0)
TotalPageNum = CInt(totalnumber / maxperpage)  '�õ���ҳ��
If TotalPageNum < totalnumber / maxperpage Then TotalPageNum = TotalPageNum + 1
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
Set Rs = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * from ECCMS_Link " & sqlLink & " ORDER BY LinkTime DESC,LinkID DESC"
If IsSqlDataBase = 1 Then
	If CurrentPage > 100 Then
		Rs.Open SQL, Conn, 1, 1
	Else
		Set Rs = enchiasp.Execute(SQL)
	End If
Else
	Rs.Open SQL, Conn, 1, 1
End If
If Rs.BOF And Rs.EOF Then
	'HtmlContent = HtmlContent & enchiasp.HtmlSetting(4)
	HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), enchiasp.HtmlSetting(4))
Else
	i = 0
	If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
	TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
	'ListContent = TempListContent
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If FlushAddress = 1 Then
			LinkAddress = enchiasp.CheckTopic(Rs("LinkUrl"))
		Else
			LinkAddress = "link.asp?id=" & Rs("LinkID") & "&url=" & enchiasp.CheckTopic(Rs("LinkUrl"))
		End If
		LinkName = enchiasp.HTMLEncode(Rs("LinkName"))
		strLinkName = "<a href=""" & LinkAddress & """ title=""" & LinkName & """ target=""_blank"">" & LinkName & "</a>"
		
		If Rs("isLogo") <> 0 Then
			If Not IsNull(Rs("LogoUrl")) And Trim(Rs("LogoUrl")) <> "" Then
				strLinkLogo = "<a href=""" & LinkAddress & """ title=""" & LinkName & """ target=""_blank""><img src='" & enchiasp.ReadFileUrl(Rs("LogoUrl")) & "' width=88 height=31 border=0></a>"
			Else
				strLinkLogo = "<a href=""" & LinkAddress & """ title=""" & LinkName & """ target=""_blank"">����LOGO</a>"
			End If
		Else
			strLinkLogo = "<a href=""" & LinkAddress & """ title=""" & LinkName & """ target=""_blank"">��������</a>"
		End If
		ListContent = ListContent & TempListContent
		ListContent = Replace(ListContent,"{$LinkID}", Rs("LinkID"))
		ListContent = Replace(ListContent,"{$LinkUrl}", LinkAddress)
		ListContent = Replace(ListContent,"{$LinkName}", strLinkName)
		ListContent = Replace(ListContent,"{$LinkLogo}", strLinkLogo)
		ListContent = Replace(ListContent,"{$LinkHist}", Rs("LinkHist"))
		ListContent = Replace(ListContent,"{$Readme}", enchiasp.HTMLEncode(Rs("Readme")))

		Rs.movenext
		i = i + 1
		If i >= maxperpage Then Exit Do
	Loop
	HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
	HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
	HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")

End If
Rs.Close:Set Rs = Nothing
strLinkPage = ShowLinkPage
HtmlContent = Replace(HtmlContent, "{$�������ӷ�ҳ����}", strLinkPage)
HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strLinkPage)
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
Response.Write HtmlContent
end if
'================================================
'��������ShowLinkPage
'��  �ã��������ӷ�ҳ
'================================================
Function ShowLinkPage()
	Dim filename, ii, n,strTemp
	filename = "index.asp"
	On Error Resume Next
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	strTemp = "<table cellspacing=1 width='100%' border=0><tr><td align=center> " & vbCrLf
	If CurrentPage < 2 Then
		strTemp = strTemp & " ������������ <font COLOR=#FF0000>" & totalnumber & "</font> ��&nbsp;&nbsp;�� ҳ&nbsp;&nbsp;��һҳ&nbsp;&nbsp;&nbsp;"
	Else
		strTemp = strTemp & "������������ <font COLOR=#FF0000>" & totalnumber & "</font> ��&nbsp;&nbsp;<a href=" & filename & "?page=1&type=" & Request("type") & ">�� ҳ</a>&nbsp;&nbsp;"
		strTemp = strTemp & "<a href=" & filename & "?page=" & CurrentPage - 1 & "&type=" & Request("type") & ">��һҳ</a>&nbsp;&nbsp;&nbsp;"
	End If

	If n - CurrentPage < 1 Then
		strTemp = strTemp & "��һҳ&nbsp;&nbsp;β ҳ " & vbCrLf
	Else
		strTemp = strTemp & "<a href=" & filename & "?page=" & (CurrentPage + 1) & "&type=" & Request("type") & ">��һҳ</a>"
		strTemp = strTemp & "&nbsp;&nbsp;<a href=" & filename & "?page=" & n & "&type=" & Request("type") & ">β ҳ</a>" & vbCrLf
	End If
	strTemp = strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
	strTemp = strTemp & "&nbsp;<b>" & maxperpage & "</b>��/ҳ " & vbCrLf
	strTemp = strTemp & "</td></tr></table>" & vbCrLf
	ShowLinkPage = strTemp
End Function
CloseConn
%>