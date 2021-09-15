<!--#include file="config.asp"-->
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
Dim TempListContent,ListContent
Dim Rs, SQL, foundsql, j
dim temptj1,temptj2
Dim maxperpage, totalnumber, TotalPageNum, CurrentPage, i
dim strPagination
Dim strClassName
if Request("jobid")="" then
	Call OutputScript("错误的参数，请不要随意输入一些参数！","index.asp")
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

'--频道目录
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","招聘职位列表")

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
	HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "没有该招聘职位信息，请确认输入参数正确")
Else
		TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 0)
		If Not Response.IsClientConnected Then Response.End
		ListContent = ListContent & TempListContent
		ListContent = Replace(ListContent,"{$zhiwei}", Rs("duix"))
		if enchiasp.HtmlSetting(2)="1" then
			temptj1 = enchiasp.Execute("SELECT COUNT(jobid) FROM ECCMS_jobbook where jobid="& rs("id") &" and isdel=0")(0)
			if enchiasp.HtmlSetting(3)="1" then
				temptj2 = enchiasp.Execute("SELECT COUNT(jobid) FROM ECCMS_jobbook where jobid="&rs("id")&" and isdel=0 and isuse=1")(0)
				ListContent = Replace(ListContent,"{$renshu}", Rs("rens")&"<font color=red>（已经递交简历"& temptj1 &"份,其中录用"& temptj2 &"人）</font>")
			else
				ListContent = Replace(ListContent,"{$renshu}", Rs("rens")&"<font color=red>（已经递交简历"& temptj1 &"份）</font>")

			end if
		else
			ListContent = Replace(ListContent,"{$renshu}", Rs("rens"))
		end if
		ListContent = Replace(ListContent,"{$didian}", Rs("did"))
		ListContent = Replace(ListContent,"{$daiyu}", Rs("Daiy"))
		ListContent = Replace(ListContent,"{$shijian}", Rs("riqi"))
		ListContent = Replace(ListContent,"{$youxiaoqi}", Rs("Qix")&"（天）")
		ListContent = Replace(ListContent,"{$xingbieyaoqiu}", Rs("sex"))
		ListContent = Replace(ListContent,"{$xueliyaoqiu}", Rs("xueli"))
		ListContent = Replace(ListContent,"{$zhuanyeyaoqiu}", Rs("zhuanye"))
		ListContent = Replace(ListContent,"{$zhaopinyaoqiu}", Rs("Yaoq"))
		ListContent = Replace(ListContent,"{$jobid}", Rs("id"))
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
End If
Rs.Close:Set Rs = Nothing


'将中间的标示过滤掉
HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strPagination)
HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
HtmlContent = HTML.ReadAnnounceList(HtmlContent)

Response.Write HtmlContent
Set HTML = Nothing
CloseConn
%>