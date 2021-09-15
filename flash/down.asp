<!--#include file="config.asp" -->
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
Dim Rs,SQL,HtmlContent
Dim ChannelRootDir,strInstallDir,strIndexName
Dim flashid,downid,showurl,ErrMsg
Dim strTitle,strAddress,addTime
Dim AllHits,Introduce,filesize
Dim HtmlFileUrl,HtmlFileName,strUrl

enchiasp.ReadChannel(ChannelID)
ChannelRootDir = enchiasp.InstallDir & enchiasp.ChannelDir
strInstallDir = enchiasp.InstallDir
strIndexName = "<a href='" & ChannelRootDir & "'>" & enchiasp.ChannelName & "</a>"
flashid = enchiasp.ChkNumeric(Request.Querystring("id"))

If flashid = 0 Then
	OutAlertScript("错误的系统参数!请输入正确的软件ID")
	Response.End
End If

enchiasp.LoadTemplates ChannelID, 6, enchiasp.ChannelSkin
HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
HtmlContent = Replace(HtmlContent, "{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
HtmlContent = Replace(HtmlContent, "{$FlashIndex}", strIndexName)
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)

SQL = "SELECT A.flashid,A.title,A.Introduce,A.filesize,A.downid,A.showurl,A.addTime,A.AllHits,A.HtmlFileDate,A.DownAddress,C.HtmlFileDir FROM ECCMS_FlashList A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID="& ChannelID &" And A.isAccept > 0 And A.flashid=" & flashid
Set Rs = enchiasp.Execute(SQL)
If Rs.EOF And Rs.BOF Then
	ErrMsg = ErrMsg & "<li>对不起~！没有找到你想下载的软件。</li>"
	Returnerr(ErrMsg)
	Set Rs = Nothing
	Response.End
Else
	strTitle = Rs("title")
	strAddress = enchiasp.ChkNull(Rs("DownAddress"))
	showurl = enchiasp.ChkNull(Rs("showurl"))
	addTime = Rs("addTime")
	AllHits = Rs("AllHits")
	downid = Rs("downid")
	Introduce = Ubbcode(Rs("Introduce"))
	If CLng(Rs("filesize")) > 0 Then
		filesize = enchicms.Readfilesize(Rs("fileSize"))
	Else
		filesize = "未知大小"
	End If
	If CInt(enchiasp.IsCreateHtml) <> 0 Then
		HtmlFileUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath)
		HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("flashid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
		strUrl = HtmlFileUrl & HtmlFileName
	Else
		strUrl = ChannelRootDir & "show.asp?id="& Rs("flashid")
	End If
End If
Rs.Close:Set Rs = Nothing

HtmlContent = Replace(HtmlContent, "{$PageTitle}", strTitle)
HtmlContent = Replace(HtmlContent, "{$FlashTitle}", strTitle)
HtmlContent = Replace(HtmlContent, "{$strUrl}", strUrl)
HtmlContent = Replace(HtmlContent, "{$DateAndTime}", addTime)
HtmlContent = Replace(HtmlContent, "{$FleshSize}", filesize)
HtmlContent = Replace(HtmlContent, "{$AllHits}", AllHits)
HtmlContent = Replace(HtmlContent, "{$Introduce}", Introduce)
HtmlContent = Replace(HtmlContent, "{$ShowDownAddress}", ShowDownAddress())
HtmlContent = Replace(HtmlContent, "{$FlashTitle}", strTitle)
HtmlContent = Replace(HtmlContent, "{$FlashID}", flashid)
HtmlContent = Replace(HtmlContent, "{$flashid}", flashid)
HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)

Response.Write HtmlContent
Set enchicms = Nothing

Public Function ShowDownAddress()
	On Error Resume Next
	
	Dim rsDown,strDownAddress
	Dim i,DownloadPath
	strDownAddress = ""
	If Len(showurl) > 3 Then
		strDownAddress = enchiasp.HtmlSetting(3)
		strDownAddress = Replace(strDownAddress, "{$DownLoadName}", "点击立即下载")
		If CInt(enchiasp.HtmlSetting(1)) > 0 Then
			strDownAddress = Replace(strDownAddress, "{$DownLoadUrl}", enchicms.FormatShowUrl(showurl))
		Else
			strDownAddress = Replace(strDownAddress, "{$DownLoadUrl}", "downfile.asp?url=" & showurl)
		End If
	End If
	If Len(strAddress) > 3 Then
		Set rsDown = enchiasp.Execute("SELECT downid,DownloadName,DownloadPath,IsDisp FROM ECCMS_DownServer WHERE ChannelID=" & ChannelID & " And depth=1 And rootid =" & downid & " And isLock=0 ORDER BY orders ASC")
		If Not (rsDown.BOF And rsDown.EOF) Then
			i = 0
			Do While Not rsDown.EOF
				If rsDown("IsDisp") > 0 Then
					DownloadPath = rsDown("DownloadPath") & strAddress
				Else
					DownloadPath = "download.asp?id=" & flashid & "&amp;downid=" & rsDown("downid")
				End If
				strDownAddress = strDownAddress & enchiasp.HtmlSetting(3)
				strDownAddress = Replace(strDownAddress, "{$DownLoadUrl}", DownloadPath)
				strDownAddress = Replace(strDownAddress, "{$DownLoadName}", rsDown("DownloadName"))
			rsDown.MoveNext
			i = i + 1
			Loop
		Else
			strDownAddress = strDownAddress & enchiasp.HtmlSetting(3)
			strDownAddress = Replace(strDownAddress, "{$DownLoadName}", "点击立即下载")
			If CInt(enchiasp.HtmlSetting(1)) > 0 Then
				strDownAddress = Replace(strDownAddress, "{$DownLoadUrl}", strAddress)
			Else
				strDownAddress = Replace(strDownAddress, "{$DownLoadUrl}", "download.asp?id=" & flashid & "&amp;downid=0")
			End If
		End If
		Set rsDown = Nothing
	End If
	ShowDownAddress = strDownAddress
End Function

CloseConn
%>