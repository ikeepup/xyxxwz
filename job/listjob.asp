<!--#include file="config.asp"-->
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
Dim Rs, SQL, foundsql, j
dim temptj1,temptj2
Dim maxperpage, totalnumber, TotalPageNum, CurrentPage, i
dim strPagination
Dim strClassName
if Request("jobid")="" then
	Call OutputScript("����Ĳ������벻Ҫ��������һЩ������","index.asp")
end if

strClassName = enchiasp.ChannelName
enchiasp.LoadTemplates ChannelID, 4, 0

HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent, "{$dingbu}",enchiasp.HtmlSetting(4))
HtmlContent = Replace(HtmlContent, "{$dibu}",enchiasp.HtmlSetting(5))
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
HtmlContent = HTML.ReadArticlePic(HtmlContent)
		HtmlContent = HTML.ReadSoftPic(HtmlContent)
		HtmlContent = HTML.ReadArticleList(HtmlContent)
		HtmlContent = HTML.ReadSoftList(HtmlContent)
		HtmlContent = HTML.ReadShopList(HtmlContent)
		HtmlContent = HTML.ReadShopPic(HtmlContent)
		HtmlContent = HTML.ReadFlashList(HtmlContent)
		HtmlContent = HTML.ReadFlashPic(HtmlContent)
		HtmlContent = HTML.ReadFriendLink(HtmlContent)
		HtmlContent = HTML.ReadNewsPicAndText(HtmlContent)
		HtmlContent = HTML.ReadSoftPicAndText(HtmlContent)
		HtmlContent = HTML.ReadGuestList(HtmlContent)
		HtmlContent = HTML.ReadAnnounceList(HtmlContent)
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
		HtmlContent = HTML.ReadPopularSoft(HtmlContent)
		HtmlContent = HTML.ReadPopularFlash(HtmlContent)
		HtmlContent = HTML.ReadSoftType(HtmlContent)
		HtmlContent = HTML.ReadStatistic(HtmlContent)
		HtmlContent = HTML.ReadUserRank(HtmlContent)

'--Ƶ��Ŀ¼
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","��Ƹְλ�б�")

HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
if IsSqlDataBase = 1 then
foundsql = "WHERE isdel=0 and id="& enchiasp.ChkNumeric(Request("jobid")) &" and getdate()<=dateadd(d,cast(qix as int),riqi) "
else
foundsql = "WHERE isdel=0 and id="& enchiasp.ChkNumeric(Request("jobid")) &" and date()<=riqi+qix"
end if
Set Rs = Server.CreateObject("ADODB.Recordset")

SQL = "SELECT * FROM ECCMS_job " & foundsql & " ORDER BY id DESC,riqi DESC"
Rs.Open SQL, Conn, 1, 1
If Rs.BOF And Rs.EOF Then
	HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "û�и���Ƹְλ��Ϣ����ȷ�����������ȷ")
Else
		TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 0)
		If Not Response.IsClientConnected Then Response.End
		ListContent = ListContent & TempListContent
		ListContent = Replace(ListContent,"{$zhiwei}", Rs("duix"))
		if enchiasp.HtmlSetting(2)="1" then
			temptj1 = enchiasp.Execute("SELECT COUNT(jobid) FROM ECCMS_jobbook where jobid="& rs("id") &" and isdel=0")(0)
			if enchiasp.HtmlSetting(3)="1" then
				temptj2 = enchiasp.Execute("SELECT COUNT(jobid) FROM ECCMS_jobbook where jobid="&rs("id")&" and isdel=0 and isuse=1")(0)
				ListContent = Replace(ListContent,"{$renshu}", Rs("rens")&"<font color=red>���Ѿ��ݽ�����"& temptj1 &"��,����¼��"& temptj2 &"�ˣ�</font>")
			else
				ListContent = Replace(ListContent,"{$renshu}", Rs("rens")&"<font color=red>���Ѿ��ݽ�����"& temptj1 &"�ݣ�</font>")

			end if
		else
			ListContent = Replace(ListContent,"{$renshu}", Rs("rens"))
		end if
		ListContent = Replace(ListContent,"{$didian}", Rs("did"))
		ListContent = Replace(ListContent,"{$daiyu}", Rs("Daiy"))
		ListContent = Replace(ListContent,"{$shijian}", Rs("riqi"))
		ListContent = Replace(ListContent,"{$youxiaoqi}", Rs("Qix")&"���죩")
		ListContent = Replace(ListContent,"{$xingbieyaoqiu}", Rs("sex"))
		ListContent = Replace(ListContent,"{$xueliyaoqiu}", Rs("xueli"))
		ListContent = Replace(ListContent,"{$zhuanyeyaoqiu}", Rs("zhuanye"))
		ListContent = Replace(ListContent,"{$zhaopinyaoqiu}", Rs("Yaoq"))
		ListContent = Replace(ListContent,"{$jobid}", Rs("id"))
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
End If
Rs.Close:Set Rs = Nothing


'���м�ı�ʾ���˵�
HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strPagination)
HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
HtmlContent = HTML.ReadAnnounceList(HtmlContent)

Response.Write HtmlContent
Set HTML = Nothing
CloseConn
%>