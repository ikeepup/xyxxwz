<!--#include file="cls_main.asp"-->
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
Const IsUseClassDir = 1
Const IsUseRemark = 1
Const MsxmlVersion = ".3.0"
Dim enchiasp,UserToday
Set enchiasp = New enchiaspMain_Cls
enchiasp.ReadConfig
'================================================
' 函数名：RelativePath2RootPath
' 作  用：转为根路径格式
' 参  数：url ----原URL
' 返回值：转换后的URL
'================================================
Function RelativePath2RootPath(url)
	Dim sTempUrl
	sTempUrl = url
	If Left(sTempUrl, 1) = "/" Then
		RelativePath2RootPath = sTempUrl
		Exit Function
	End If

	Dim sFilePath
	sFilePath = Request.ServerVariables("SCRIPT_NAME")
	sFilePath = Left(sFilePath, InstrRev(sFilePath, "/") - 1)
	Do While Left(sTempUrl, 3) = "../"
		sTempUrl = Mid(sTempUrl, 4)
		sFilePath = Left(sFilePath, InstrRev(sFilePath, "/") - 1)
	Loop
	RelativePath2RootPath = sFilePath & "/" & sTempUrl
End Function
'================================================
' 函数名：RootPath2DomainPath
' 作  用：根路径转为带域名全路径格式
' 参  数：url ----原URL
' 返回值：转换后的URL
'================================================
Function RootPath2DomainPath(url)
	Dim sHost, sPort
	sHost = Split(LCase(Request.ServerVariables("SERVER_PROTOCOL")), "/")(0) & "://" & Request.ServerVariables("HTTP_HOST")
	sPort = Request.ServerVariables("SERVER_PORT")
	If sPort <> "80" Then
		sHost = sHost & ":" & sPort
	End If
	RootPath2DomainPath = sHost & url
End Function
'================================================
' 函数名：ChkMapPath
' 作  用：相对路径转换为绝对路径
' 参  数：strPath ----原路径
' 返回值：绝对路径
'================================================
Public Function ChkMapPath(ByVal strPath)
	On Error Resume Next
	Dim fullPath
	strPath = Replace(Replace(Trim(strPath), "//", "/"), "\\", "\")

	If strPath = "" Then strPath = "."
	If InStr(strPath,":\") = 0 Then 
		fullPath = Server.MapPath(strPath)
	Else
		strPath = Replace(strPath,"/","\")
		fullPath = Trim(strPath)
		If Right(fullPath, 1) = "\" Then
			fullPath = Left(fullPath, Len(fullPath) - 1)
		End If
	End If
	ChkMapPath = fullPath
End Function
'================================================
' 函数名：CreatePath
' 作  用：按月份自动创建文件夹
' 参  数：fromPath ----原文件夹路径
'================================================
Function CreatePath(fromPath)
	Dim objFSO, uploadpath
	uploadpath = Year(Now) & "-" & Month(Now) '以年月创建上传文件夹，格式：2003－8
	On Error Resume Next
	Set objFSO = CreateObject(enchiasp.FSO_ScriptName)
	If objFSO.FolderExists(Server.MapPath(fromPath & uploadpath)) = False Then
		objFSO.CreateFolder Server.MapPath(fromPath & uploadpath)
	End If
	If Err.Number = 0 Then
		CreatePath = uploadpath & "/"
	Else
		CreatePath = ""
	End If
	Set objFSO = Nothing
End Function
'================================================
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'        False ----没有安装
'================================================
Function IsObjInstalled(ByVal strClassString)
	Dim xTestObj,ClsString
	On Error Resume Next
	IsObjInstalled = False
	ClsString = strClassString
	Err = 0
	Set xTestObj = Server.CreateObject(ClsString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
	Exit Function
End Function
Public Sub GetUserTodayInfo()
	Dim Lastlogin,UserDayInfo
	Lastlogin = Request.Cookies("enchiasp_net")("LastTime")
	UserDayInfo = Request.Cookies("enchiasp_net")("UserToday")
	If Not IsDate(LastLogin) Then LastLogin = Now()
	On Error Resume Next
	If DateDiff("d",LastLogin,Now())<>0 Then
		enchiasp.Execute("UPDATE [ECCMS_User] SET UserToday='0,0,0,0,0,0',LastTime=" & NowString & " WHERE username='"& enchiasp.membername &"' And userid=" & enchiasp.memberid)
		UserDayInfo = "0,0,0,0,0,0"
		Response.Cookies("enchiasp_net")("UserToday") = UserDayInfo
		Response.Cookies("enchiasp_net")("LastTime") = Now()
	End If
	UserToday = Split(UserDayInfo, ",")
	If Ubound(UserToday) <> 5 Then
		enchiasp.Execute("UPDATE [ECCMS_User] SET UserToday='0,0,0,0,0,0',LastTime=" & NowString & " WHERE username='"& enchiasp.membername &"' And userid=" & enchiasp.memberid)
		UserDayInfo = "0,0,0,0,0,0"
		Response.Cookies("enchiasp_net")("UserToday") = UserDayInfo
		Response.Cookies("enchiasp_net")("LastTime") = Now()
		UserToday = Split(UserDayInfo, ",")
	End If
End Sub
Public Function UpdateUserToday(ByVal str)
	On Error Resume Next
	If Trim(str) <> "" Then
		enchiasp.Execute("UPDATE [ECCMS_User] SET UserToday='" & str & "' WHERE username='"& enchiasp.membername &"' And userid=" & enchiasp.memberid)
		Response.Cookies("enchiasp_net")("UserToday") = str
	End If
End Function
'================================================
'作  用：输出错误警告脚本
'参  数：str ----参数入口
'返回值：警告信息
'================================================
Sub OutAlertScript(str)
	Response.Write "<script language=javascript>" & vbcrlf
	Response.Write "alert('" & str & "');"
	Response.Write "history.back()" & vbcrlf
	Response.Write "</script>" & vbcrlf
	Response.End
End Sub
Sub OutHintScript(str)
	Response.Write "<script language=JavaScript>" & vbCrLf
	Response.Write "alert('" & str & "');"
	Response.Write "location.replace('" & Request.ServerVariables("HTTP_REFERER") & "')" & vbCrLf
	Response.Write "</script>" & vbCrLf
	Response.End
End Sub
Sub OutputScript(str,url)
	Response.Write "<script language=JavaScript>" & vbCrLf
	Response.Write "alert('" & str & "');"
	Response.Write "location.replace('" & url & "')" & vbCrLf
	Response.Write "</script>" & vbCrLf
	Response.End
End Sub
'================================================
'函数名：URLDecode
'作  用：URL解码
'================================================
Function URLDecode(str)
	If IsNull(str) Then
		URLDecode = ""
		Exit Function
	End If
	str = Replace(str, "%7F", vbNullString, 1, -1, 1)
	str = Replace(str, "%B0l", "發", 1, -1, 1)
	Dim i, strReturn, strSpecial
	Dim thischr, intasc
	strSpecial = "!""#$%&'()*+,/:;<=>?@[\]^`{|}~%._-"
	strReturn = ""
	For i = 1 To Len(str)
		thischr = Mid(str, i, 1)
		If thischr = "%" Then
			intasc = eval("&h" + Mid(str, i + 1, 2))
			If InStr(strSpecial, Chr(intasc)) > 0 Then
				strReturn = strReturn & Chr(intasc)
				i = i + 2
			Else
				intasc = eval("&h" + Mid(str, i + 1, 2) + Mid(str, i + 4, 2))
				strReturn = strReturn & Chr(intasc)
				i = i + 5
			End If
		Else
			If thischr = "+" Then
				strReturn = strReturn & " "
			Else
				strReturn = strReturn & thischr
			End If
		End If
	Next
	URLDecode = strReturn
End Function
'================================================
'过程名：PreventRefresh
'作  用：防止刷新页面
'================================================
Sub PreventRefresh()
	Dim RefreshTime,isRefresh
	RefreshTime = 10   '防止刷新时间,单位（秒）
	isRefresh = 1    '是否使用防刷新功能，0=否，1=是
	If isRefresh = 1 Then
		If (Not IsEmpty(Session("RefreshTime"))) And RefreshTime > 0 Then
			If DateDiff("s", Session("RefreshTime"), Now()) < RefreshTime Then
				Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT="&RefreshTime&"><br>本页面起用了防刷新机制，请不要在"&RefreshTime&"秒内连续刷新本页面<BR>正在打开页面，请稍后……"
				Response.End
			Else
				Session("RefreshTime") = Now()
			End If
		Else
			Session("RefreshTime") = Now()
		End If
	End If
End Sub

Public Sub GetUserTodayInfo()
	Dim Lastlogin,UserDayInfo
	Lastlogin = Request.Cookies("enchiasp_net")("LastTime")
	UserDayInfo = Request.Cookies("enchiasp_net")("UserToday")
	If Not IsDate(LastLogin) Then LastLogin = Now()
	On Error Resume Next
	If DateDiff("d",LastLogin,Now())<>0 Then
		enchiasp.Execute("UPDATE [ECCMS_User] SET UserToday='0,0,0,0,0,0',LastTime=" & NowString & " WHERE username='"& enchiasp.membername &"' And userid=" & enchiasp.memberid)
		UserDayInfo = "0,0,0,0,0,0"
		Response.Cookies("enchiasp_net")("UserToday") = UserDayInfo
		Response.Cookies("enchiasp_net")("LastTime") = Now()
	End If
	UserToday = Split(UserDayInfo, ",")
	If Ubound(UserToday) <> 5 Then
		enchiasp.Execute("UPDATE [ECCMS_User] SET UserToday='0,0,0,0,0,0',LastTime=" & NowString & " WHERE username='"& enchiasp.membername &"' And userid=" & enchiasp.memberid)
		UserDayInfo = "0,0,0,0,0,0"
		Response.Cookies("enchiasp_net")("UserToday") = UserDayInfo
		Response.Cookies("enchiasp_net")("LastTime") = Now()
		UserToday = Split(UserDayInfo, ",")
	End If
End Sub
Public Function UpdateUserToday(ByVal str)
	On Error Resume Next
	If Trim(str) <> "" Then
		enchiasp.Execute("UPDATE [ECCMS_User] SET UserToday='" & str & "' WHERE username='"& enchiasp.membername &"' And userid=" & enchiasp.memberid)
		Response.Cookies("enchiasp_net")("UserToday") = str
	End If
End Function
'================================================
'作  用：获取频道URL
'参  数：ChannelDir ----频道目录
'        BindDomain ----绑定域名
'        DomainName ----完整域名URL
'返回值：频道URL
'================================================
Function GetChannelUrl(ChannelDir,BindDomain,DomainName)
	Dim strTempUrl
	If CInt(BindDomain) <> 0 Then
		strTempUrl = DomainName
	Else
		strTempUrl = enchiasp.InstallDir & ChannelDir
	End If
	GetChannelUrl = strTempUrl
End Function

'================================================
'作  用：内容路径转换
'参  数：ChannelDir ----频道目录
'        BindDomain ----绑定域名
'        DomainName ----完整域名URL
'返回值：频道URL
'================================================
Function ContentPathConvert(url,BindDomain)
	Dim strTempUrl
	If CInt(BindDomain) <> 0 Then
		strTempUrl = enchiasp.SiteUrl & url
	Else
		strTempUrl = url
	End If
	ContentPathConvert = strTempUrl
End Function
'================================================
'作  用：读取图片或者FLASH
'参  数：url ----文件URL
'        height ----高度
'        width ----宽度
'================================================
Function GetFlashAndPic(url,height,width)
	Dim sExtName,ExtName,strTemp
	sExtName = Split(url, ".")
	ExtName = sExtName(UBound(sExtName))
	If LCase(ExtName) = "swf" Then
		strTemp = "<embed src=""" & url & """ width=" & width & " height=" & height & ">"
	Else
		strTemp = "<img src=""" & url & """ width=" & width & " height=" & height & " border=0>"
	End If
	GetFlashAndPic = strTemp
End Function
Function Html2Ubb(str)
	If Str<>"" And Not IsNull(Str) Then
		Dim re,tmpstr
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern = "(<STRONG>)":Str = re.Replace(Str,"<b>")
		re.Pattern = "(<\/STRONG>)":Str = re.Replace(Str,"</b>")
		re.Pattern ="(<TBODY>)":Str = re.Replace(Str,"")
		re.Pattern ="(<\/TBODY>)":Str = re.Replace(Str,"")
		re.Pattern ="(<TABLE)":Str = re.Replace(Str,"<table")
		re.Pattern ="(TABLE>)":Str = re.Replace(Str,"table>")
		re.Pattern ="(<TR)":Str = re.Replace(Str,"<tr")
		re.Pattern ="(TR>)":Str = re.Replace(Str,"tr>")
		re.Pattern ="(<TD)":Str = re.Replace(Str,"<td")
		re.Pattern ="(TD>)":Str = re.Replace(Str,"td>")
		re.Pattern ="(<DIV)":Str = re.Replace(Str,"<div")
		re.Pattern ="(Div>)":Str = re.Replace(Str,"div>")
		re.Pattern ="(<IMG )":Str = re.Replace(Str,"<img ")
		re.Pattern ="(<BR)":Str = re.Replace(Str,"<br")
		re.Pattern ="(<A )":Str = re.Replace(Str,"<a ")
		re.Pattern ="(<\/A>)":Str = re.Replace(Str,"</a>")
		re.Pattern ="(<FONT )":Str = re.Replace(Str,"<font ")
		re.Pattern ="(<\/FONT>)":Str = re.Replace(Str,"</font>")
		re.Pattern = "(<s+cript(.+?)<\/s+cript>)":Str = re.Replace(Str, "")
		re.Pattern ="(\{)":Str = re.Replace(Str,"&#123;")
		re.Pattern ="(\})":Str = re.Replace(Str,"&#125;")
		re.Pattern ="(\$)":Str = re.Replace(Str,"&#36;")
		re.Pattern = "(<div(.+?)>)":Str = re.replace(Str,"<div>")
		re.Pattern = "(<span(.+?)>)":Str = re.replace(Str,"<span>")
		Set Re=Nothing
		Html2Ubb = Str
	Else
		Html2Ubb = ""
	End If
End Function
%>