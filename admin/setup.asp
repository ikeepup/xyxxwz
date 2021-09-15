<!--#include file="../conn.asp" -->
<!--#include file="../inc/const.asp"-->
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
'Response.Expires = -1
'Response.ExpiresAbsolute = Now() - 1
'Response.Expires = 0
Dim Rs,SQL,lconn
Dim FoundErr,ErrMsg,SucMsg,AdminPage
FoundErr = False
AdminPage = False
Const Admin_Cookies_Name = "admin_enchiasp"
'Session.TimeOut = SessionTimeout
Sub ConnectionLogDatabase()
	On Error Resume Next
	Dim lconnstr
	lconnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Logdata.Asa")
	Set lconn = Server.CreateObject("ADODB.Connection")
	lconn.open lconnstr
End Sub

Sub SaveLogInfo(lname)
	Dim RequestStr
	Dim lsql,istoplog
	istoplog = 0      '是否停止日志,1=停止,0=启用
	If istoplog = 1 Then Exit Sub
	On Error Resume Next
	ConnectionLogDatabase
	If InStr(enchiasp.ScriptName, "_index") > 0 Or InStr(enchiasp.ScriptName, "admin_log") > 0 Then Exit Sub
	lname = enchiasp.CheckStr(lname) 
	RequestStr = lcase(Request.ServerVariables("Query_String"))
	If RequestStr <> "" Then 
		RequestStr=checkStr(RequestStr)
		RequestStr=Left(RequestStr,250)
		lsql = "insert into [ECCMS_LogInfo] (UserName,UserIP,ScriptName,ActContent,LogAddTime,LogType) values ('"& lname &"','"& enchiasp.GetUserip &"','"& enchiasp.ScriptName &"','"& RequestStr &"','"& Now() &"',0)"		
		lconn.Execute(lsql)
	End If
	If Request.form <> "" Then
		RequestStr = checkStr(request.form)
		RequestStr = Left(RequestStr,250)
		lsql = "insert into [ECCMS_LogInfo] (UserName,UserIP,ScriptName,ActContent,LogAddTime,LogType) values ('"& lname &"','"& enchiasp.GetUserip &"','"& enchiasp.ScriptName &"','"& RequestStr &"','"& Now() &"',1)"		
		lconn.Execute(lsql)
	End If
	If IsObject(lconn) Then
		lconn.Close
		Set lconn = Nothing
	End If
End Sub

Function fixjs(str)
        If str <> "" Then
                str = Replace(str, "\", "\\")
                str = Replace(str, Chr(34), "\""")
                str = Replace(str, Chr(39), "\'")
                str = Replace(str, Chr(13), "")
                str = Replace(str, Chr(10), "")
                'str = replace(str,"'", "&#39;")
        End If
        fixjs = str
        Exit Function
End Function
'================================================
'函数名：ShowListPage
'作  用：通用分页
'================================================
Function ShowListPage(CurrentPage,Pcount,totalrec,PageNum,strLink,ListName)
	With Response
		.Write "<script>"
		.Write "ShowListPage("
		.Write CurrentPage
		.Write ","
		.Write Pcount
		.Write ","
		.Write totalrec
		.Write ","
		.Write PageNum
		.Write ",'"
		.Write strLink
		.Write "','"
		.Write ListName
		.Write "');"
		.Write "</script>" & vbNewLine
	End With
End Function
'================================================
'函数名：showpages
'作  用：通用分页
'================================================
Function showpages(CurrentPage,Pcount,totalrec,PageNum,str)
	Dim strTemp,strRequest
	strRequest = str
	strTemp = "<table border=0 cellpadding=0 cellspacing=3 width=""100%"" align=center>" & vbNewLine
	strTemp = strTemp & "<tr><td valign=middle nowrap>" & vbNewLine
	strTemp = strTemp & "页次：<b><font color=red>" & CurrentPage & "</font></b>/<b>" & Pcount & "</b>页&nbsp;" & vbNewLine
	strTemp = strTemp & "每页<b>" & PageNum & "</b> 总数<b>" & totalrec & "</b></td>" & vbNewLine
	strTemp = strTemp & "<td valign=middle nowrap align=right>分页：" & vbNewLine
	strTemp = strTemp & "<script language=""JavaScript"">" & vbNewLine
	strTemp = strTemp & "<!--" & vbNewLine
	strTemp = strTemp & "var CurrentPage=" & CurrentPage & ";" & vbNewLine
	strTemp = strTemp & "var Pcount=" & Pcount & ";" & vbNewLine
	strTemp = strTemp & "var Endpage=0;" & vbNewLine
	strTemp = strTemp & "if (CurrentPage > 4){" & vbNewLine
	strTemp = strTemp & "	document.write ('<a href=""?page=1" & strRequest & """>[1]</a> ...');" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "if (Pcount>CurrentPage+3)" & vbNewLine
	strTemp = strTemp & "{" & vbNewLine
	strTemp = strTemp & "	Endpage=CurrentPage+3" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "else{" & vbNewLine
	strTemp = strTemp & "	Endpage=Pcount" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "for (var i=CurrentPage-3;i<=Endpage;i++)" & vbNewLine
	strTemp = strTemp & "{" & vbNewLine
	strTemp = strTemp & "	if (i>=1){" & vbNewLine
	strTemp = strTemp & "		if (i == CurrentPage)" & vbNewLine
	strTemp = strTemp & "		{" & vbNewLine
	strTemp = strTemp & "			document.write ('<font color=""#FF0000"">['+i+']</font>');" & vbNewLine
	strTemp = strTemp & "			}" & vbNewLine
	strTemp = strTemp & "		else{" & vbNewLine
	strTemp = strTemp & "			document.write ('<a href=""?page='+i+'" & strRequest & """>['+i+']</a>');" & vbNewLine
	strTemp = strTemp & "		}" & vbNewLine
	strTemp = strTemp & "	}" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "if (CurrentPage+3 < Pcount){" & vbNewLine 
	strTemp = strTemp & "	document.write ('...<a href=""?page='+Pcount+'" & strRequest & """>['+Pcount+']</a>');" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "if (Endpage == 0){ " & vbNewLine
	strTemp = strTemp & "	document.write ('...');" & vbNewLine
	strTemp = strTemp & "}" & vbNewLine
	strTemp = strTemp & "//-->" & vbNewLine
	strTemp = strTemp & "</script>" & vbNewLine
	strTemp = strTemp & "</td></tr></table>"
	ShowPages = strTemp
End Function

Public Sub ReturnError(ErrMsg)
	Response.Write "<html><head><title>错误提示信息!</title><meta http-equiv=Content-Type content=text/html; charset=gb2312>" & vbCrLf
	Response.Write "<meta http-equiv=refresh content=3;url=javascript:history.go(-1)>"
	Response.Write "<link href=style.css rel=stylesheet type=text/css></head><body><p>&nbsp;</p>" & vbCrLf
	Response.Write "<table cellpadding=5 cellspacing=0 border=0 align=center class=tableBorder1>" & vbCrLf
	Response.Write "  <tr><th colspan=2 align=""left""><img src=""images/welcome.gif"" width=""16"" height=""17"" align=""absMiddle""> 错误提示信息!</th></tr>" & vbCrLf
	Response.Write "  <tr><td align=center width=""20%"" class=TableRow1><img src=""images/err.gif"" width=95 height=97 border=0></td><td width=""80%"" class=TableRow1><b style=color:blue><span id=jump>3</span> 秒钟后系统将自动返回</b><br><b>产生错误的可能原因：</b><BR>" & ErrMsg & "</td></tr>" & vbCrLf
	Response.Write "  <tr><td colspan=2 align=center height=25 class=TableRow2><a href=javascript:history.go(-1)>返回上一页...</a></td></tr>" & vbCrLf
	Response.Write "</table><p>&nbsp;</p>" & vbCrLf
	Response.Write "</body></html>" & vbCrLf
	Response.Write "<script>function countDown(secs){jump.innerText=secs;if(--secs>0)setTimeout(""countDown(""+secs+"")"",1000);}countDown(3);</script>"
End Sub

Public Sub Succeed(SucMsg)
	Response.Write "<html><head><title>错误提示信息!</title><meta http-equiv=Content-Type content=text/html; charset=gb2312>" & vbCrLf
	Response.Write "<meta http-equiv=refresh content=5;url=" & Request.ServerVariables("HTTP_REFERER") & ">"
	Response.Write "<link href=style.css rel=stylesheet type=text/css></head><body><p>&nbsp;</p>" & vbCrLf
	Response.Write "<table align=""center"" border=""0"" cellpadding=""5"" cellspacing=""0"" class=""tableBorder1"">" & vbCrLf
	Response.Write "    <tr> " & vbCrLf
	Response.Write "      <th colspan=2 align=""left""><img src=""images/welcome.gif"" width=""16"" height=""17"" align=""absMiddle""> 成功提示信息!</th>" & vbCrLf
	Response.Write "    </tr>" & vbCrLf
	Response.Write "  <tr><td align=center width=""20%"" class=TableRow1><img src=""images/succ.gif"" width=95 height=97 border=0></td><td width=""80%"" class=TableRow1>"
	Response.Write " <b style=color:blue><span id=jump>5</span> 秒钟后系统将自动返回</b><br>"
	Response.Write SucMsg & "</td></tr>" & vbCrLf
	Response.Write "  <tr><td colspan=2 align=center height=25 class=TableRow2><a href='" & Request.ServerVariables("HTTP_REFERER") & "'>返回上一页...</a></td></tr>" & vbCrLf
	Response.Write " </table><p>&nbsp;</p>" & vbCrLf
	Response.Write "</body></html>" & vbCrLf
	Response.Write "<script>function countDown(secs){jump.innerText=secs;if(--secs>0)setTimeout(""countDown(""+secs+"")"",1000);}countDown(3);</script>"
End Sub

Public Function ErrAlert(thistr)
	Response.Write "<script language=JavaScript>" & vbCrLf
	Response.Write "alert('" & thistr & "');"
	Response.Write "javascript:history.back(1)" & vbCrLf
	Response.Write "</script>" & vbCrLf
	Response.End
End Function

Public Function SucInform(thistr)
	Response.Write "<script language=JavaScript>" & vbCrLf
	Response.Write "alert('" & thistr & "');"
	Response.Write "location.replace('" & Request.ServerVariables("HTTP_REFERER") & "')" & vbCrLf
	Response.Write "</script>" & vbCrLf
	Response.End
End Function

Public Function AlertInform(this_str,this_url)
	Response.Write "<script language=JavaScript>" & vbCrLf
	Response.Write "alert('" & this_str & "');"
	Response.Write "location.replace('" & this_url & "')" & vbCrLf
	Response.Write "</script>" & vbCrLf
	Response.End
End Function

Public Function CheckAdmin(flag)
	Dim Rs, SQL
	Dim i, TempAdmin, Adminflag,AdminGrade
	CheckAdmin = False
	On Error Resume Next
	SQL ="SELECT id,AdminGrade,Adminflag FROM [ECCMS_Admin] WHERE username='"& Replace(Session("AdminName"), "'", "''") &"' And password='"& Replace(Session("AdminPass"), "'", "''") &"' And isLock=0 And id="& CLng(Session("AdminID"))
	Set Rs = enchiasp.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		CheckAdmin = False
		Set Rs = Nothing
		Exit Function
	Else
		Adminflag = Rs("Adminflag")
		AdminGrade = Rs("AdminGrade")
	End If
	Rs.Close:Set Rs = Nothing
	If CInt(AdminGrade) = 999 Then
		CheckAdmin = True
		Exit Function
	Else
		If Trim(flag) = "" Then Exit Function
		If Adminflag = "" Then
			CheckAdmin = False
			Exit Function
		Else
			tempAdmin = Split(AdminFlag, ",")
			For i = 0 To UBound(tempAdmin)
				If LCase(tempAdmin(i)) = LCase(flag) Then
					CheckAdmin = True
					Exit For
				End If
			Next
		End If
	End If
End Function

Sub Admin_footer()
        Response.Write "<br /><table align=center>" & vbCrLf
        Response.Write "<tr align=center><td width=""100%"" style=""LINE-HEIGHT: 150%"" class=copyright>" & vbCrLf        
        If CInt(isSqlDataBase) = 1 Then
                Response.Write " Powered by：<a href=http://www.enchi.com.cn target=_blank>enchi cms ver 3.0.0</a> （MS SQL 版）<br>" & vbCrLf
        Else
                Response.Write " Powered by：<a href=http://www.enchi.com.cn target=_blank>enchi cms ver 3.0.0</a> （ACCESS 版）<br>" & vbCrLf
        End If
        Response.Write enchiasp.Copyright & vbCrLf
        
        If CInt(enchiasp.IsRunTime) = 1 Then
                Dim Endtime
                Endtime = Timer()
                Response.Write "<BR>执行时间：" & FormatNumber(Endtime - startime,5, -1) & "毫秒。查询数据库" & enchiasp.SqlQueryNum & "次。" & vbCrLf
                'Response.Write "<li>共使用了" & Application.Contents.Count & "个缓存对象。</li>"
        End If
        Response.Write "</td>" & vbCrLf
        Response.Write "</tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
        Response.Write "</body></html>"
End Sub

Sub Admin_header()
        Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" & vbCrLf
        Response.Write enchiasp.CopyrightStr
        Response.Write "<html>" & vbCrLf
        Response.Write "<head>" & vbCrLf
        Response.Write "<meta http-equiv=""X-UA-Compatible"" content=""IE=6"" />" &  vbCrLf
 Response.Write "<base target=""_self"">" &  vbCrLf

        Response.Write "<title>" & enchiasp.SiteName & "-管理页面</title>" & vbCrLf
        Response.Write "<LINK href=""style.css"" type=text/css rel=stylesheet>" & vbCrLf
        Response.Write "<script src=""include/admin.js"" type=""text/javascript""></script>" & vbCrLf
        Response.Write "</head>" & vbCrLf
        Response.Write "<body leftmargin=0 bottommargin=0 rightmargin=0 topmargin=0>" & vbCrLf
        Response.Write "<BR style=""OVERFLOW: hidden; LINE-HEIGHT: 3px"">" & vbCrLf
End Sub
Public Sub ScriptCreation(url,id)
	Response.Write "<span id='showimport" & id & "'></span>"
	Response.Write "<script>"
	Response.Write "function CreationDone(str){"
	Response.Write "	showimport" & id & ".innerHTML = str;"
	Response.Write "}"
	Response.Write "CreationID.startDownload('" & url & "',CreationDone)"
	Response.Write "</script>" & vbCrLf
End Sub
%>
