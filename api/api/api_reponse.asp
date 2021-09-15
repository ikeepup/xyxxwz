<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!--#include file="../inc/chkinput.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="cls_api.asp"-->
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
Dim XMLDom,XmlDoc,Node,Status,Messenge
Dim UserName,Act,appid
Status = 1
Messenge = ""

If Request.QueryString<>"" And API_Enable Then
	SaveUserCookie()
Else
	Set XmlDoc = Server.CreateObject("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	XmlDoc.ASYNC = False
	If API_Enable Then
		If Not XmlDoc.LOAD(Request) Then
			Status = 1
			Messenge = "数据非法，操作中止！"
			appid = "未知"
		Else
			If CheckPost() Then
				Select Case Act
					Case "checkname"
						Checkname()
					Case "reguser"
						UserReguser()
					Case "login"
						UesrLogin()
					Case "logout"
						LogoutUser()
					Case "update"
						UpdateUser()
					Case "delete"
						Deleteuser()
					Case "lock"
						Lockuser()
					Case "getinfo"
						GetUserinfo()
				End Select
			End If
		End If
	Else
		Status = 0
		Messenge = "API接口关闭，操作中止！"
		appid = "enchiasp"
	End If
	ReponseData()
	Set XmlDoc = Nothing
End If

Sub ReponseData()
	If Act <> "getinfo" Then
		XmlDoc.loadxml "<root><appid>dvbbs</appid><status>0</status><body><message/></body></root>"
	End If
	XmlDoc.documentElement.selectSingleNode("appid").text = "enchiasp"
	If API_Debug And Act <> "reguser" Then
		XmlDoc.documentElement.selectSingleNode("status").text = 0
		Messenge = ""
	Else
		XmlDoc.documentElement.selectSingleNode("status").text = status
	End If
	XmlDoc.documentElement.selectSingleNode("body/message").text = ""
	Set Node = XmlDoc.createCDATASection(Replace(Messenge,"]]>","]]&gt;"))
	XmlDoc.documentElement.selectSingleNode("body/message").appendChild(Node)
	Response.Clear
	Response.ContentType="text/xml"
	Response.CharSet="gb2312"
	Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
	Response.Write XmlDoc.documentElement.XML
End Sub

Function CheckPost()
	CheckPost = False
	Dim Syskey
	If XmlDoc.documentElement.selectSingleNode("action") is Nothing or XmlDoc.documentElement.selectSingleNode("syskey") is Nothing or XmlDoc.documentElement.selectSingleNode("username")  is Nothing Then
		Status = 1
		Messenge = Messenge & "<li>非法请求。</li>"
		Exit Function
	End If
	UserName = enchiasp.CheckBadstr(XmlDoc.documentElement.selectSingleNode("username").text)
	Syskey = XmlDoc.documentElement.selectSingleNode("syskey").text
	Act = XmlDoc.documentElement.selectSingleNode("action").text
	Appid = XmlDoc.documentElement.selectSingleNode("appid").text
	
	Dim NewMd5,OldMd5
	NewMd5 = Md5(UserName & API_ConformKey)
	Md5OLD = 1
	OldMd5 = Md5(UserName & API_ConformKey)
	Md5OLD = 0

	If Syskey=NewMd5 or Syskey=OldMd5 Then
		CheckPost = True
	Else
		Status = 1
		Messenge = Messenge & "<li>请求数据验证不通过，请与管理员联系。</li>"
	End If
End Function

Sub GetUserinfo()
	Dim Rs,Sql
	XmlDoc.loadxml "<root><appid>dvbbs</appid><status>0</status><body><message/><email/><question/><answer/><savecookie/><truename/><gender/><birthday/><qq/><msn/><mobile/><telephone/><address/><zipcode/><homepage/><userip/><jointime/><experience/><ticket/><valuation/><balance/><posts/><userstatus/></body></root>"
	
	Sql = "SELECT TOP 1 * FROM ECCMS_User WHERE UserName='" & enchiasp.CheckBadstr(UserName) & "'"
	Set Rs = enchiasp.Execute(Sql)
	If Not Rs.Eof And Not Rs.Bof Then
		XmlDoc.documentElement.selectSingleNode("body/email").text = Rs("usermail") & ""
		XmlDoc.documentElement.selectSingleNode("body/question").text = Rs("question") & ""
		XmlDoc.documentElement.selectSingleNode("body/answer").text = Rs("answer") & ""
		XmlDoc.documentElement.selectSingleNode("body/gender").text = Rs("Usersex") & ""
		XmlDoc.documentElement.selectSingleNode("body/birthday").text = ""
		XmlDoc.documentElement.selectSingleNode("body/mobile").text = ""
		XmlDoc.documentElement.selectSingleNode("body/userip").text = Rs("userlastip") & ""
		XmlDoc.documentElement.selectSingleNode("body/jointime").text = Rs("JoinTime") & ""
		XmlDoc.documentElement.selectSingleNode("body/experience").text = Rs("experience") & ""
		XmlDoc.documentElement.selectSingleNode("body/ticket").text = ""
		XmlDoc.documentElement.selectSingleNode("body/valuation").text = Rs("userpoint") & ""
		XmlDoc.documentElement.selectSingleNode("body/balance").text = Rs("usermoney") & ""
		XmlDoc.documentElement.selectSingleNode("body/posts").text = Rs("postcode") & ""
		XmlDoc.documentElement.selectSingleNode("body/userstatus").text = Rs("UserLock")
		XmlDoc.documentElement.selectSingleNode("body/homepage").text = Rs("HomePage") & ""
		XmlDoc.documentElement.selectSingleNode("body/qq").text = Rs("oicq")
		XmlDoc.documentElement.selectSingleNode("body/msn").text = ""
		XmlDoc.documentElement.selectSingleNode("body/truename").text = Rs("TrueName") & ""
		XmlDoc.documentElement.selectSingleNode("body/telephone").text = Rs("phone") & ""
		XmlDoc.documentElement.selectSingleNode("body/address").text = Rs("address") & ""
		Status = 0
		Messenge = Messenge & "<li>读取用户资料成功。</li>"
	Else
		Status = 1
		Messenge = Messenge & "<li>该用户不存在。</li>"
	End If
	Rs.Close
	Set Rs = Nothing
End Sub

Sub Checkname()
	Dim Rs,SQL,UserEmail
	UserEmail = enchiasp.checkstr(Trim(XmlDoc.documentElement.selectSingleNode("email").text))
	If IsValidEmail(UserEmail) = False Then
		Messenge = "<li>您的Email有错误！</li>"
		Status = 1
		Exit Sub
	End If
	If CInt(enchiasp.ChkSameMail) = 1 Then
		Set Rs = enchiasp.Execute("SELECT userid FROM ECCMS_User WHERE usermail='" & UserEmail & "'")
		If Not Rs.EOF Then
			Status = 1
			Messenge = "<li>此邮箱["&UserEmail&"]已经占用，请您换一个邮箱再注册吧。</li>"
			Exit Sub
		End If
		Rs.Close:Set Rs = Nothing
	End If
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_User WHERE username = '" & UserName & "'")
	If Not (Rs.bof And Rs.EOF) Then
		Status = 1
		Messenge =  "<li>Sorry！此用户已经存在,请换一个用户名再试！</li>"
	Else
		Status = 0
		Messenge =  "<li><font color=red><b>" & UserName & "</b></font> 尚未被人使用，赶紧注册吧！</li>"
	End If
	Rs.Close:Set Rs = Nothing
End Sub

Sub UserReguser()
	Dim nickname,UserPass,UserEmail,Question,Answer,usercookies
	Dim strGroupName,Password,usersex,sex
	Dim Rs,SQL
	UserPass = enchiasp.checkstr(XmlDoc.documentElement.selectSingleNode("password").text)
	UserEmail = enchiasp.checkstr(Trim(XmlDoc.documentElement.selectSingleNode("email").text))
	Question = enchiasp.checkstr(XmlDoc.documentElement.selectSingleNode("question").text)
	Answer = enchiasp.checkstr(XmlDoc.documentElement.selectSingleNode("answer").text)
	sex = enchiasp.ChkNumeric(XmlDoc.documentElement.selectSingleNode("gender").text)
	If sex = 0 Then
		usersex = "女"
	Else
		usersex = "男"
	End If
	usercookies = 1
	If UserName = "" Or UserPass = "" Then
		Status = 1
		Messenge = Messenge & "<li>请填写用户名或密码。"
		Exit Sub
	End If
	If Question = "" Then Question = enchiasp.GetRandomCode
	If Answer = "" Then Answer = enchiasp.GetRandomCode
	nickname = UserName
	Password = md5(UserPass)
	Answer = md5(Answer)
	If enchiasp.IsValidStr(UserName) = False Then
		Messenge = Messenge & "<li>登录账号中含有非法字符！</li>"
		Status = 1
		Exit Sub
	End If
	If IsValidEmail(UserEmail) = False Then
		Messenge = Messenge & "<li>您的Email有错误！</li>"
		Status = 1
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_User WHERE username='" & UserName & "'")
	If Not (Rs.BOF And Rs.EOF) Then
		Status = 1
		Messenge = Messenge & "<li>Sorry！此用户已经存在,请换一个用户名再试！</li>"
		Exit Sub
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_Admin WHERE username='" & UserName & "'")
	If Not (Rs.BOF And Rs.EOF) Then
		Status = 1
		Messenge = Messenge & "<li>Sorry！此用户已经存在,请换一个用户名再试！</li>"
		Exit Sub
	End If
	Rs.Close:Set Rs = Nothing
	If CInt(enchiasp.ChkSameMail) = 1 Then
		Set Rs = enchiasp.Execute("SELECT userid FROM ECCMS_User WHERE usermail='" & UserEmail & "'")
		If Not Rs.EOF Then
			Status = 1
			Messenge = Messenge & "<li>对不起！本系统已经限制一个邮箱只能注册一个账号。</li><li>此邮箱["&UserEmail&"]已经占用，请您换一个邮箱再注册吧。</li>"
			Exit Sub
		End If
		Rs.Close:Set Rs = Nothing
	End If
	'---
	Set Rs = enchiasp.Execute("SELECT GroupName FROM ECCMS_UserGroup WHERE Groupid=3")
	If Rs.BOF And Rs.EOF Then
		strGroupName = "普通会员"
	Else
		strGroupName = enchiasp.CheckBadstr(Rs(0))
		If Len(strGroupName) = 0 Then strGroupName = "普通会员"
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_User WHERE (userid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("username") = UserName
		Rs("password") = Password
		Rs("nickname") = UserName
		Rs("UserGrade") = 1
		Rs("UserGroup") = strGroupName
		Rs("UserClass") = 0
		If CInt(enchiasp.AdminCheckReg) = 1 Then
			Rs("UserLock") = 1
		Else
			Rs("UserLock") = 0
		End If
		Rs("UserFace") = "face/1.gif"
		Rs("userpoint") = CLng(enchiasp.AddUserPoint)
		Rs("usermoney") = 0
		Rs("savemoney") = 0
		Rs("prepaid") = 0
		Rs("experience") = 10
		Rs("charm") = 10
		Rs("TrueName") = UserName
		Rs("usersex") = usersex
		Rs("usermail") = UserEmail
		Rs("oicq") = ""
		Rs("question") = Question
		Rs("answer") = Answer
		Rs("JoinTime") = Now()
		Rs("ExpireTime") = Now()
		Rs("LastTime") = Now()
		Rs("Protect") = 0
		Rs("usermsg") = 0
		Rs("userlastip") = enchiasp.GetUserip
		Rs("userlogin") = 0
		Rs("usersetting") = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
	Rs.update
	Rs.Close
	Set Rs = Nothing
	Status = 0
	Messenge = "用户注册成功。"
End Sub

Sub UesrLogin()
	Dim UserPass
	
	UserPass = enchiasp.checkstr(XmlDoc.documentElement.selectSingleNode("password").text)
	If UserName="" or UserPass="" Then
		Status = 1
		Messenge = Messenge & "<li>请填写用户名或密码。</li>"
		Exit Sub
	End If
	UserPass = Md5(UserPass)
	
	If ChkUserLogin(username,UserPass,1) Then
		Status = 0
		Messenge = Messenge & "<li>登陆成功。</li>"
	Else
		Status = 1
		Messenge = Messenge & "<li>登陆失败。</li>"
	End If
End Sub

Sub LogoutUser()
	Response.Cookies(enchiasp.Cookies_Name).path = "/"
	Response.Cookies(enchiasp.Cookies_Name)("userid") = ""
	Response.Cookies(enchiasp.Cookies_Name)("username") = ""
	Response.Cookies(enchiasp.Cookies_Name)("password") = ""
	Response.Cookies(enchiasp.Cookies_Name)("nickname") = ""
	Response.Cookies(enchiasp.Cookies_Name)("UserGrade") = ""
	Response.Cookies(enchiasp.Cookies_Name)("UserGroup") = ""
	Response.Cookies(enchiasp.Cookies_Name)("UserClass") = ""
	Response.Cookies(enchiasp.Cookies_Name)("UserToday") = ""
	Response.Cookies(enchiasp.Cookies_Name)("usercookies") = ""
	Response.Cookies(enchiasp.Cookies_Name)("LastTimeDate") = ""
	Response.Cookies(enchiasp.Cookies_Name)("LastTimeIP") = ""
	Response.Cookies(enchiasp.Cookies_Name)("LastTime") = ""
	Response.Cookies(enchiasp.Cookies_Name) = ""
End Sub

Sub UpdateUser()
	Dim Rs,SQL
	Dim UserPass,UserEmail,Question,Answer
	UserPass = enchiasp.checkstr(XmlDoc.documentElement.selectSingleNode("password").text)
	UserEmail = enchiasp.checkstr(Trim(XmlDoc.documentElement.selectSingleNode("email").text))
	Question = enchiasp.checkstr(XmlDoc.documentElement.selectSingleNode("question").text)
	Answer = enchiasp.checkstr(XmlDoc.documentElement.selectSingleNode("answer").text)
	If UserPass <> "" Then
		UserPass = Md5(UserPass)
	End If
	If Answer <> "" THen
		Answer = Md5(Answer)
	End If
	If IsValidEmail(UserEmail) = False Then
		UserEmail = ""
	End If
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	SQL = "SELECT TOP 1 * FROM [ECCMS_User] WHERE Username='" & UserName & "'"
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open SQL,Conn,1,3
	If Not Rs.Eof And Not Rs.Bof Then
		If UserPass <> "" Then Rs("password") = UserPass
		If Answer <> "" THen Rs("answer") = Answer
		If UserEmail <> "" Then Rs("usermail") = UserEmail
		If Question <> "" Then Rs("question") = Question
		Rs.update
		Status = 0
		Messenge = "<li>基本资料修改成功。</li>"
	Else
		Status = 1
		Messenge = "<li>该用户不存在，修改资料失败。</li>"
	End If
	Rs.Close
	Set Rs = Nothing
	If UserPass <> "" And Status = 0 Then
		Response.Cookies(enchiasp.Cookies_Name)("password") = UserPass
	End If
End Sub

Sub Deleteuser()
	Dim Del_Users,i,AllUserID,Del_UserName
	Dim Rs
	Del_Users = Split(UserName,",")
	For i = 0 To UBound(Del_Users)
		Del_UserName = enchiasp.CheckBadstr(Del_Users(i))
		Set Rs = enchiasp.Execute("SELECT userid,username FROM [ECCMS_User] WHERE UserName='" & Del_UserName & "'")
		If Not (Rs.Eof And Rs.Bof) Then
			AllUserID = AllUserID & Rs(0) & ","
			enchiasp.Execute("UPDATE ECCMS_Message SET delsend=1 WHERE sender='"& enchiasp.CheckStr(Rs(1)) &"'")
			enchiasp.Execute("DELETE FROM ECCMS_Message WHERE flag=0 And incept='"& enchiasp.CheckStr(Rs(1)) &"'")
			Messenge = Messenge & "<li>用户（" & Del_UserName & "）删除成功。</li>"
		End If
	Next
	Set Rs = Nothing
	If AllUserID <> "" Then
		If Right(AllUserID,1) = "," Then AllUserID = Left(AllUserID,Len(AllUserID)-1)
		enchiasp.Execute ("DELETE FROM ECCMS_User WHERE userid in (" & AllUserID & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_Favorite WHERE userid in (" & AllUserID & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_Friend WHERE userid in (" & AllUserID & ")")
	End If
	Status = 0
End Sub

Sub Lockuser()
	Dim UserStatus
	If XmlDoc.documentElement.selectSingleNode("userstatus") is Nothing Then
		Messenge = "<li>参数非法，中止请求。</li>"
		Status = 1
		Exit Sub
	ElseIf Not IsNumeric(XmlDoc.documentElement.selectSingleNode("userstatus").text) Then
		Messenge = "<li>参数非法，中止请求。</li>"
		Status = 1
		Exit Sub
	Else
		UserStatus = Clng(XmlDoc.documentElement.selectSingleNode("userstatus").text)
	End If
	If UserStatus = 0 Then
		enchiasp.Execute ("UPDATE ECCMS_User SET UserLock=0 WHERE Username='" & UserName & "'")
	Else
		enchiasp.Execute ("UPDATE ECCMS_User SET UserLock=1 WHERE Username='" & UserName & "'")
	End If
	Status = 0
End Sub

Sub SaveUserCookie()
	Dim S_syskey,Password,usercookies,TruePassWord,userclass,Userhidden
	
	S_syskey = Request.QueryString("syskey")
	UserName = enchiasp.CheckBadstr(Request.QueryString("UserName"))
	Password = Request.QueryString("Password")
	usercookies = Request.QueryString("savecookie")
	If UserName="" or S_syskey="" Then Exit Sub
	Dim NewMd5,OldMd5
	NewMd5 = Md5(UserName & API_ConformKey)
	Md5OLD = 1
	OldMd5 = Md5(UserName & API_ConformKey)
	Md5OLD = 0
	If Not (S_syskey=NewMd5 or S_syskey=OldMd5) Then
		Exit Sub
	End If
	If usercookies="" or Not IsNumeric(usercookies) Then usercookies = 0
	
	'用户退出
	If Password = "" Then
		Response.Cookies(enchiasp.Cookies_Name).path = "/"
		Response.Cookies(enchiasp.Cookies_Name)("userid") = ""
		Response.Cookies(enchiasp.Cookies_Name)("username") = ""
		Response.Cookies(enchiasp.Cookies_Name)("password") = ""
		Response.Cookies(enchiasp.Cookies_Name)("nickname") = ""
		Response.Cookies(enchiasp.Cookies_Name)("UserGrade") = ""
		Response.Cookies(enchiasp.Cookies_Name)("UserGroup") = ""
		Response.Cookies(enchiasp.Cookies_Name)("UserClass") = ""
		Response.Cookies(enchiasp.Cookies_Name)("UserToday") = ""
		Response.Cookies(enchiasp.Cookies_Name)("usercookies") = ""
		Response.Cookies(enchiasp.Cookies_Name)("LastTimeDate") = ""
		Response.Cookies(enchiasp.Cookies_Name)("LastTimeIP") = ""
		Response.Cookies(enchiasp.Cookies_Name)("LastTime") = ""
		Response.Cookies(enchiasp.Cookies_Name) = ""
		Exit Sub
	End If
	ChkUserLogin username,password,usercookies
End Sub

Function ChkUserLogin(username,password,usercookies)
	ChkUserLogin = False
	Dim Rs,SQL,Group_Setting
	
	If Not IsObject(Conn) Then ConnectionDatabase
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM [ECCMS_User] WHERE username='" & UserName & "'"
	Rs.Open SQL, Conn, 1, 3

	If Not (Rs.BOF And Rs.EOF) Then
		If password <> Rs("password") Then
			ChkUserLogin = False
			Exit Function
		End If
		If Rs("UserLock") <> 0 Then
			ChkUserLogin = False
			Exit Function
		End If
		Response.Cookies(enchiasp.Cookies_Name)("LastTimeDate") = Rs("LastTime")
		Response.Cookies(enchiasp.Cookies_Name)("LastTimeIP") = Rs("userlastip")
		Response.Cookies(enchiasp.Cookies_Name)("LastTime") = Rs("LastTime")
		Group_Setting=Split(enchiasp.UserGroupSetting(Rs("UserGrade")), "|||")
		If Rs("userpoint") < 0 Then
			Rs("userpoint") = CLng(Group_Setting(25))
		Else
			Rs("userpoint") = Rs("userpoint") + CLng(Group_Setting(25))
		End If
		If Rs("experience") < 0 Then
			Rs("experience") = CLng(Group_Setting(32))
		Else
			Rs("experience") = Rs("experience") + CLng(Group_Setting(32))
		End If
		If Rs("charm") < 0 Then
			Rs("charm") = CLng(Group_Setting(33))
		Else
			Rs("charm") = Rs("charm") + CLng(Group_Setting(33))
		End If
		Rs("LastTime") = Now()
		Rs("userlastip") = enchiasp.GetUserip
		Rs("UserLogin") = Rs("UserLogin") + 1
		Rs.Update
		
		Select Case usercookies
		Case 0
			Response.Cookies(enchiasp.Cookies_Name)("usercookies") = usercookies
		Case 1
			Response.Cookies(enchiasp.Cookies_Name).Expires=Date+1
			Response.Cookies(enchiasp.Cookies_Name)("usercookies") = usercookies
		Case 2
			Response.Cookies(enchiasp.Cookies_Name).Expires=Date+31
			Response.Cookies(enchiasp.Cookies_Name)("usercookies") = usercookies
		Case 3
			Response.Cookies(enchiasp.Cookies_Name).Expires=Date+365
			Response.Cookies(enchiasp.Cookies_Name)("usercookies") = usercookies
		End Select
		Response.Cookies(enchiasp.Cookies_Name).path = "/"
		Response.Cookies(enchiasp.Cookies_Name)("userid") = Rs("userid")
		Response.Cookies(enchiasp.Cookies_Name)("username") = Rs("username")
		Response.Cookies(enchiasp.Cookies_Name)("password") = Rs("password")
		Response.Cookies(enchiasp.Cookies_Name)("nickname") = Rs("nickname")
		Response.Cookies(enchiasp.Cookies_Name)("UserGrade") = Rs("UserGrade")
		Response.Cookies(enchiasp.Cookies_Name)("UserGroup") = Rs("UserGroup")
		Response.Cookies(enchiasp.Cookies_Name)("UserClass") = Rs("UserClass")
		Response.Cookies(enchiasp.Cookies_Name)("UserToday") = Rs("UserToday")
		ChkUserLogin = True
	End If
	Rs.Close
	Set Rs = Nothing
End Function

%>