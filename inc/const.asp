<!--#include file="cls_main.asp"-->
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
Const IsUseClassDir = 1
Const IsUseRemark = 1
Const MsxmlVersion = ".3.0"
Dim enchiasp,UserToday
Set enchiasp = New enchiaspMain_Cls
enchiasp.ReadConfig
'================================================
' ��������RelativePath2RootPath
' ��  �ã�תΪ��·����ʽ
' ��  ����url ----ԭURL
' ����ֵ��ת�����URL
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
' ��������RootPath2DomainPath
' ��  �ã���·��תΪ������ȫ·����ʽ
' ��  ����url ----ԭURL
' ����ֵ��ת�����URL
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
' ��������ChkMapPath
' ��  �ã����·��ת��Ϊ����·��
' ��  ����strPath ----ԭ·��
' ����ֵ������·��
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
' ��������CreatePath
' ��  �ã����·��Զ������ļ���
' ��  ����fromPath ----ԭ�ļ���·��
'================================================
Function CreatePath(fromPath)
	Dim objFSO, uploadpath
	uploadpath = Year(Now) & "-" & Month(Now) '�����´����ϴ��ļ��У���ʽ��2003��8
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
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'        False ----û�а�װ
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
'��  �ã�������󾯸�ű�
'��  ����str ----�������
'����ֵ��������Ϣ
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
'��������URLDecode
'��  �ã�URL����
'================================================
Function URLDecode(str)
	If IsNull(str) Then
		URLDecode = ""
		Exit Function
	End If
	str = Replace(str, "%7F", vbNullString, 1, -1, 1)
	str = Replace(str, "%B0l", "�l", 1, -1, 1)
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
'��������PreventRefresh
'��  �ã���ֹˢ��ҳ��
'================================================
Sub PreventRefresh()
	Dim RefreshTime,isRefresh
	RefreshTime = 10   '��ֹˢ��ʱ��,��λ���룩
	isRefresh = 1    '�Ƿ�ʹ�÷�ˢ�¹��ܣ�0=��1=��
	If isRefresh = 1 Then
		If (Not IsEmpty(Session("RefreshTime"))) And RefreshTime > 0 Then
			If DateDiff("s", Session("RefreshTime"), Now()) < RefreshTime Then
				Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT="&RefreshTime&"><br>��ҳ�������˷�ˢ�»��ƣ��벻Ҫ��"&RefreshTime&"��������ˢ�±�ҳ��<BR>���ڴ�ҳ�棬���Ժ󡭡�"
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
'��  �ã���ȡƵ��URL
'��  ����ChannelDir ----Ƶ��Ŀ¼
'        BindDomain ----������
'        DomainName ----��������URL
'����ֵ��Ƶ��URL
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
'��  �ã�����·��ת��
'��  ����ChannelDir ----Ƶ��Ŀ¼
'        BindDomain ----������
'        DomainName ----��������URL
'����ֵ��Ƶ��URL
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
'��  �ã���ȡͼƬ����FLASH
'��  ����url ----�ļ�URL
'        height ----�߶�
'        width ----���
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