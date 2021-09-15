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
Dim Rs,SQL,foundstr
Dim classid,ChildStr
Dim RssBody,RssTitle,RssHomePageUrl
Dim XMLDOM,node,Cnode,Cnode1,msginfo
Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
XMLDOM.appendChild(XMLDOM.createElement("rss"))
XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"version","")).text="2.0"
Set node = XMLDOM.documentElement.appendChild(XMLDOM.createNode(1,"channel",""))
RssHomePageUrl = enchiasp.SiteUrl
RssTitle = "获取信息列表"
classid = enchiasp.CheckNumeric(Request("classid"))
Sub XMLFlashList()
	If classid > 0 Then
		SQL = "SELECT ClassName,ChildStr FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID=" & CLng(classid)
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			Set Cnode=node.appendChild(XMLDOM.createNode(1,"item",""))
			Cnode.appendChild(XMLDOM.createNode(1,"title","")).text="没有找到分类"
			Cnode.appendChild(XMLDOM.createNode(1,"link","")).text=RssHomePageUrl
			Cnode.appendChild(XMLDOM.createNode(1,"author","")).text=enchiasp.SiteName
			Cnode.appendChild(XMLDOM.createNode(1,"pubDate","")).text=Now()
			Set Cnode1=Cnode.appendChild(XMLDOM.createNode(1,"description",""))
			msginfo= "没有找到软件分类！"
			Cnode1.appendChild(XMLDOM.createCDATASection(msginfo))
			Rs.Close: Set Rs = Nothing
			Exit Sub
		Else
			RssTitle = Rs("ClassName")
			ChildStr = Rs("ChildStr")
		End If
		Rs.Close:Set Rs = Nothing
		foundstr = "And A.ClassID in (" & ChildStr & ") ORDER BY A.addTime DESC ,A.flashid DESC"
	Else
		RssTitle = "全部软件列表"
		foundstr = "ORDER BY A.addTime DESC ,A.flashid DESC"
	End If

	node.appendChild(XMLDOM.createNode(1,"title","")).text=enchiasp.SiteName&"--"&RssTitle
	node.appendChild(XMLDOM.createNode(1,"link","")).text=enchiasp.SiteUrl
	node.appendChild(XMLDOM.createNode(1,"language","")).text="zh-cn"
	node.appendChild(XMLDOM.createNode(1,"description","")).text=enchiasp.SiteName
	node.appendChild(XMLDOM.createNode(1,"copyright","")).text=enchiasp.SiteUrl
	node.appendChild(XMLDOM.createNode(1,"generator","")).text="Rss Generator By enchi.com"

	Dim HtmlFileName,HtmlFileUrl
	SQL = " A.flashid,A.ClassID,A.title,A.addTime,A.HtmlFileDate,A.username,"
	SQL = "SELECT TOP 100 " & SQL & " C.ClassName,C.HtmlFileDir,C.UseHtml,B.ChannelDir,B.StopChannel,B.ModuleName,B.IsCreateHtml,B.HtmlExtName,B.HtmlPath,B.HtmlForm,B.HtmlPrefix FROM ([ECCMS_FlashList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID) INNER JOIN [ECCMS_Channel] B On A.ChannelID=B.ChannelID WHERE A.isAccept>0 And A.ChannelID=" & CLng(ChannelID) & " " & foundstr & ""
	Set Rs = enchiasp.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		Set Cnode=node.appendChild(XMLDOM.createNode(1,"item",""))
		Cnode.appendChild(XMLDOM.createNode(1,"title","")).text="没有找到你想要的信息"
		Cnode.appendChild(XMLDOM.createNode(1,"link","")).text=RssHomePageUrl
		Cnode.appendChild(XMLDOM.createNode(1,"author","")).text=enchiasp.SiteName
		Cnode.appendChild(XMLDOM.createNode(1,"pubDate","")).text=Now()
		Set Cnode1=Cnode.appendChild(XMLDOM.createNode(1,"description",""))
		msginfo= "没有找到你想下载的软件！"
		Cnode1.appendChild(XMLDOM.createCDATASection(msginfo))
		Rs.Close: Set Rs = Nothing
		Exit Sub
	Else
		Do While Not Rs.EOF
			HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("flashid"), Rs("HtmlExtName"), Rs("HtmlPrefix"), Rs("HtmlForm"), "")
			If Rs("IsCreateHtml") <> 0 Then
				HtmlFileUrl = enchiasp.GetChannelDir(ChannelID) & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), Rs("HtmlPath")) & HtmlFileName
			Else
				HtmlFileUrl = enchiasp.GetChannelDir(ChannelID) & "show.asp?id=" & Rs("flashid")
			End If
			If LCase(Left(HtmlFileUrl,7)) <> "http://" Then HtmlFileUrl = RssHomePageUrl & HtmlFileUrl
			Set Cnode=node.appendChild(XMLDOM.createNode(1,"item",""))
			Cnode.appendChild(XMLDOM.createNode(1,"title","")).text=Rs("title")
			Cnode.appendChild(XMLDOM.createNode(1,"link","")).text=HtmlFileUrl
			Cnode.appendChild(XMLDOM.createNode(1,"category","")).text=Rs("ClassName")
			Cnode.appendChild(XMLDOM.createNode(1,"author","")).text=Rs("username")
			Cnode.appendChild(XMLDOM.createNode(1,"pubDate","")).text=Rs("addTime")
			Set Cnode1=Cnode.appendChild(XMLDOM.createNode(1,"description",""))
			msginfo=  "要浏览本条信息请点击标题名称。"
			Cnode1.appendChild(XMLDOM.createCDATASection(msginfo))
			Rs.MoveNext
		Loop
	End If
	Rs.Close: Set Rs = Nothing
End Sub

Sub ShowXML()
	Response.Clear
	Response.CharSet="gb2312"
	Response.ContentType="text/xml"
	Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
	Response.Write XMLDOM.xml
	Set XMLDOM=Nothing
End Sub

XMLFlashList()
ShowXML()
CloseConn
%>