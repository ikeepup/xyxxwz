
<%
'
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
'Option Explicit
Dim startime,Conn,db,comDomain 
Response.Buffer = True
startime = Timer()
Dim NowString
Dim ConnStr
'域名精确定向

'comDomain = "www.carnurse.com" '定义COM域名 
'If Request.ServerVariables("SERVER_NAME") <> comDomain Then '如果请求的域名不是com的域名
            'Response.Status = 301 '表示状态切换成301
            'Response.AddHeader "Location","http://" & comDomain & "/"&Request.ServerVariables("HTTP_X_REWRITE_URL") '在头中添加Location字段，使用当前的求情的路径加上COM的域名组合成新的地址。
'End If


'定义数据库类别，1为SQL数据库，0为Access数据库
const isSqlDataBase = 0

If IsSqlDataBase = 1 Then
	'-----------------------SQL数据库连接参数---------------------------------------
	Dim SqlDatabaseName,SqlPassword,SqlUsername,SqlLocalName
	NowString = "GetDate()"
	SqlDatabaseName = "Ec_cms_tongda"     '数据库名
	SqlUsername = "sa"          '用户名
	SqlPassword = "avdswx"          '用户密码
	SqlLocalName = "."        '连接名（本地用local，外地用IP）
	ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " &  SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	'-------------------------------------------------------------------------------
Else
	'-----------------------ACCESS数据库连接----------------------------------------
	NowString = "Now()"
	'ACCESS数据库连接,请使用根路径或者绝对路径
	db = "/database/#ECCMS.mdb"
	Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ChkMapPath(db)
	'-------------------------------------------------------------------------------
End If

Sub ConnectionDatabase()
	On Error Resume Next
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Connstr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "数据库连接出错。"
		Response.End
	End If
	
	
	
End Sub

Dim DBPath
'-- 采集数据库连接路径
DBPath = "/database/#EC_Collection.mdb"		'-- 请用根相对路径

Sub CloseConn()
	On Error Resume Next
	If IsObject(Conn) Then
		Conn.Close
		Set Conn = Nothing
	End If
	Set enchiasp = Nothing
End Sub
%>